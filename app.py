"""
Intrastat Factuur Import App
Auteur: Lander Smits

Extraheert artikelregels uit aankoopfacturen (PDF), aggregeert per product
over alle pagina's, en genereert een Intrastat Excel-bestand.
"""

import io
import os
import re
import uuid
import logging
from collections import defaultdict
from datetime import datetime

import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from flask import Flask, request, jsonify, send_file, render_template, abort
from werkzeug.utils import secure_filename

# ---------------------------------------------------------------------------
# App setup
# ---------------------------------------------------------------------------
app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024  # 20 MB

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

RESULT_STORE: dict[str, tuple[bytes, str]] = {}

# ---------------------------------------------------------------------------
# Constants & patterns
# ---------------------------------------------------------------------------

# Units we recognise when splitting "250 kg" into (250, "kg")
UNITS = [
    "kg", "g", "gram", "liter", "l", "ml", "stuks", "stuk", "st",
    "dozen", "doos", "pallet", "pallets", "pak", "pakken",
    "ton", "t", "m", "m2", "m3", "meter",
]
UNIT_RE = re.compile(
    r"^(\d[\d\s.,]*)[\s\xa0]+(" + "|".join(UNITS) + r")\.?$",
    re.IGNORECASE,
)

# Rows to skip (subtotals, transport, VAT, headers, page markers, …)
SKIP_RE = re.compile(
    r"(subtotaal|sub-totaal|totaal|total|transport|levering|verzend|"
    r"btw|vat|tva|korting|discount|intracommunautair|betalings|"
    r"bladzijde|pagina|page\s*\d|iban|bank|valuta|factuurdatum|"
    r"leveringsdatum|factuurnummer|klant|leverancier|btw-nummer|"
    r"^artikel\s*$|^omschrijving\s*$|hoeveelheid\s+eenheid)",
    re.IGNORECASE,
)

AMOUNT_RE = re.compile(r"€?\s*(\d[\d.,\s]*\d|\d+)[,.](\d{2})(?!\d)")
NUMBER_RE = re.compile(r"^\d[\d.,\s]*$")


# ---------------------------------------------------------------------------
# Number helpers
# ---------------------------------------------------------------------------

def clean_number(val: str | None) -> float | None:
    """Parse European / US number strings into a Python float."""
    if val is None:
        return None
    val = str(val).strip().replace("\xa0", "").replace(" ", "")
    val = re.sub(r"[€$£]", "", val)
    if not val:
        return None
    # European thousands separator: "1.234,56"
    if re.search(r"\d\.\d{3},", val):
        val = val.replace(".", "").replace(",", ".")
    # Only comma as decimal: "234,56"
    elif val.count(",") == 1 and val.count(".") == 0:
        val = val.replace(",", ".")
    # Comma as thousands, dot as decimal: "1,234.56"
    elif val.count(",") >= 1 and val.count(".") == 1:
        val = val.replace(",", "")
    # Dot as thousands, no decimal: "1.234" → treat as integer
    elif val.count(".") == 1 and len(val.split(".")[1]) == 3:
        val = val.replace(".", "")
    try:
        return float(val)
    except ValueError:
        return None


def split_qty_unit(text: str) -> tuple[float | None, str]:
    """
    Split a combined quantity-unit string like '250 kg' or '180 stuks'.
    Returns (quantity_float, unit_string).
    Falls back to (quantity_float, '') if no unit found.
    """
    text = text.strip()
    m = UNIT_RE.match(text)
    if m:
        qty = clean_number(m.group(1))
        unit = m.group(2).lower().rstrip(".")
        return qty, unit

    # Maybe it's just a plain number
    num = clean_number(text)
    if num is not None:
        return num, ""

    return None, ""


# ---------------------------------------------------------------------------
# Raw line-item extraction
# ---------------------------------------------------------------------------

def _parse_table_page(table: list[list]) -> list[dict]:
    """Parse one pdfplumber table into raw line items."""
    if not table or len(table) < 2:
        return []

    # Normalise cells
    rows = [[str(c).strip() if c else "" for c in row] for row in table]

    # Find header row (first row with recognisable column labels)
    header_kw = re.compile(
        r"(artikel|product|omschr|descri|ref|code|"
        r"hoeveelh|aantal|qty|quantit|"
        r"prijs|price|bedrag|amount|totaal|total|"
        r"eenheid|unit)",
        re.IGNORECASE,
    )
    header_idx = 0
    for i, row in enumerate(rows[:5]):
        if sum(1 for c in row if header_kw.search(c)) >= 2:
            header_idx = i
            break

    headers = [h.lower() for h in rows[header_idx]]
    data_rows = rows[header_idx + 1:]

    # Map columns
    def col(patterns, exclude=None):
        """Return first column index whose header contains any of 'patterns',
        but NOT any string from 'exclude'."""
        for i, h in enumerate(headers):
            if exclude and any(ex in h for ex in exclude):
                continue
            for p in patterns:
                if p in h:
                    return i
        return None

    i_art  = col(["artikel", "product", "omschr", "descri", "goed"])
    i_qty  = col(["hoeveelh", "aantal", "qty", "quantit"])
    # "eenheid" must NOT match "eenheidsprijs" → exclude price-related headers
    i_unit = col(["eenheid", "unit"], exclude=["prijs", "price", "waarde", "bedrag"])
    i_up   = col(["eenheidsprijs", "stuksprijs", "unit price", "prijs p"])
    i_tot  = col(["totaal", "total", "bedrag", "amount"])

    items = []
    for row in data_rows:
        if not any(row):
            continue
        combined = " ".join(row)
        if SKIP_RE.search(combined):
            continue
        if not re.search(r"\d", combined):
            continue

        def get(idx):
            return row[idx] if idx is not None and idx < len(row) else ""

        art_val  = get(i_art)
        qty_val  = get(i_qty)
        unit_val = get(i_unit)
        up_val   = get(i_up)
        tot_val  = get(i_tot)

        # Always attempt to split quantity column for embedded unit ("250 kg")
        qty_num, embedded_unit = split_qty_unit(qty_val) if qty_val else (None, "")

        # Prefer dedicated unit column; fall back to embedded unit from qty cell
        # Reject unit_val that looks like a price (contains digits + comma/dot pattern)
        if unit_val and not re.search(r"[€$£]|\d+[.,]\d{2}", unit_val):
            final_unit = unit_val.lower().strip()
        else:
            final_unit = embedded_unit  # e.g. "kg" extracted from "250 kg"

        item = {
            "artikel":        art_val.strip(),
            "hoeveelheid":    qty_num,
            "eenheid":        final_unit,
            "eenheidsprijs":  clean_number(up_val),
            "bedrag":         clean_number(tot_val),
        }

        if item["artikel"] or item["bedrag"] is not None:
            items.append(item)

    return items


def _parse_text_page(text: str) -> list[dict]:
    """
    Regex fallback: scan plain-text lines for invoice article lines.

    Expected formats (any of):
      Verse tomaten  250 kg  €2,10  €525,00
      Verse tomaten  250  kg  2,10  525,00
      Verse tomaten  250kg  525,00
    """
    items = []

    # Build unit alternation for inline regex
    unit_alt = "|".join(UNITS)

    # Pattern: description  number [unit]  [price]  total
    LINE_RE = re.compile(
        r"^(.+?)\s{2,}"                            # description (≥2 spaces after)
        r"(\d[\d.,]*)[\s\xa0]*"                    # quantity
        r"(?:(" + unit_alt + r")\.?\s*)?"          # optional unit
        r"(?:€?\s*\d[\d.,]*\s+)?"                 # optional unit price
        r"€?\s*(\d[\d.,]+)$",                      # total at end
        re.IGNORECASE,
    )

    for raw_line in text.split("\n"):
        line = raw_line.strip()
        if not line or len(line) < 6:
            continue
        if SKIP_RE.search(line):
            continue

        m = LINE_RE.match(line)
        if m:
            art  = m.group(1).strip()
            qty  = clean_number(m.group(2))
            unit = (m.group(3) or "").lower().strip()
            tot  = clean_number(m.group(4))

            if art and (qty is not None or tot is not None):
                items.append({
                    "artikel":       art,
                    "hoeveelheid":   qty,
                    "eenheid":       unit,
                    "eenheidsprijs": None,
                    "bedrag":        tot,
                })
            continue

        # Simpler fallback: line with at least 2 amounts and some text
        amounts = re.findall(r"€?\s*(\d[\d.,]+)", line)
        if len(amounts) >= 2:
            tot = clean_number(amounts[-1])
            # Remove amounts from line to get description
            desc = re.sub(r"€?\s*\d[\d.,]+", "", line).strip()
            desc = re.sub(r"\s{2,}", " ", desc).strip()
            if len(desc) >= 3 and tot and tot > 0:
                items.append({
                    "artikel":       desc,
                    "hoeveelheid":   None,
                    "eenheid":       "",
                    "eenheidsprijs": None,
                    "bedrag":        tot,
                })

    return items


# ---------------------------------------------------------------------------
# Main PDF processor
# ---------------------------------------------------------------------------

def process_pdf(pdf_bytes: bytes) -> tuple[list[dict], dict]:
    """
    Parse all pages of a PDF and return:
      - aggregated list (one entry per unique product, quantities summed)
      - invoice meta dict
    """
    raw_items: list[dict] = []
    full_text = ""

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            page_text = page.extract_text() or ""
            full_text += page_text + "\n"

            page_items: list[dict] = []

            # 1. Try table extraction first
            tables = page.extract_tables()
            for table in tables:
                page_items.extend(_parse_table_page(table))

            # 2. If tables yielded nothing, try text parsing
            if not page_items:
                page_items = _parse_text_page(page_text)

            # Tag each item with its page number
            for item in page_items:
                item["pagina"] = page_num

            raw_items.extend(page_items)

    # Filter items that are clearly not article lines
    raw_items = [
        it for it in raw_items
        if it["artikel"]
        and not SKIP_RE.search(it["artikel"])
        and len(it["artikel"]) >= 2
    ]

    # ── Aggregate by product name ──────────────────────────────────────────
    # Key = lowercased+stripped article name for case-insensitive grouping
    groups: dict[str, dict] = {}
    key_to_canonical: dict[str, str] = {}  # key → original display name

    for item in raw_items:
        key = re.sub(r"\s+", " ", item["artikel"].lower().strip())
        if key not in groups:
            groups[key] = {
                "artikel":        item["artikel"],   # first seen → canonical name
                "hoeveelheid":    0.0,
                "eenheid":        item["eenheid"] or "",
                "prijzen":        [],                # collect all unit prices
                "bedrag":         0.0,
                "paginas":        [],
                "regels":         0,
            }
        g = groups[key]
        # Accumulate
        if item["hoeveelheid"] is not None:
            g["hoeveelheid"] += item["hoeveelheid"]
        if item["bedrag"] is not None:
            g["bedrag"] += item["bedrag"]
        if item["eenheidsprijs"] is not None:
            g["prijzen"].append(item["eenheidsprijs"])
        if item["pagina"] not in g["paginas"]:
            g["paginas"].append(item["pagina"])
        g["regels"] += 1

    # Build final aggregated list
    aggregated = []
    for key, g in groups.items():
        # Determine eenheidsprijs: use single price or average if varies
        if g["prijzen"]:
            avg_price = sum(g["prijzen"]) / len(g["prijzen"])
            # If all prices are equal, use that; else signal variance
            if len(set(round(p, 4) for p in g["prijzen"])) == 1:
                unit_price = g["prijzen"][0]
                price_note = ""
            else:
                unit_price = avg_price
                price_note = "gem."
        else:
            # Derive from total / quantity if possible
            if g["hoeveelheid"] and g["hoeveelheid"] > 0 and g["bedrag"]:
                unit_price = g["bedrag"] / g["hoeveelheid"]
                price_note = "berekend"
            else:
                unit_price = None
                price_note = ""

        aggregated.append({
            "artikel":        g["artikel"],
            "hoeveelheid":    g["hoeveelheid"] if g["hoeveelheid"] else None,
            "eenheid":        g["eenheid"],
            "eenheidsprijs":  unit_price,
            "price_note":     price_note,
            "bedrag":         g["bedrag"] if g["bedrag"] else None,
            "paginas":        sorted(g["paginas"]),
            "regels":         g["regels"],
        })

    # Sort alphabetically by article name
    aggregated.sort(key=lambda x: x["artikel"].lower())

    # Extract invoice meta from full text
    meta = _extract_meta(full_text)

    logger.info(
        "Parsed %d raw lines → %d unique products", len(raw_items), len(aggregated)
    )
    return aggregated, meta, raw_items


def _extract_meta(text: str) -> dict:
    meta = {"factuurnummer": "", "datum": "", "leverancier": ""}

    m = re.search(
        r"(factuur\s*(?:nr|nummer|no)?\.?\s*:?\s*)([A-Z0-9\-/]{3,25})",
        text, re.IGNORECASE,
    )
    if m:
        meta["factuurnummer"] = m.group(2).strip()

    m = re.search(
        r"(?:factuurdatum|datum)[:\s]+(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})",
        text, re.IGNORECASE,
    )
    if not m:
        m = re.search(
            r"(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})", text
        )
    if m:
        meta["datum"] = m.group(1)

    # Supplier: line after "Leverancier:"
    m = re.search(r"leverancier\s*:\s*(.+)", text, re.IGNORECASE)
    if m:
        meta["leverancier"] = m.group(1).strip()
    else:
        for line in text.split("\n"):
            line = line.strip()
            if line and len(line) > 4 and not re.match(
                r"^(aankoopfactuur|invoice|factuur|bladzijde|pagina)",
                line, re.IGNORECASE
            ):
                meta["leverancier"] = line
                break

    return meta


# ---------------------------------------------------------------------------
# Excel builder
# ---------------------------------------------------------------------------

BROWN       = "8B7355"
BEIGE_ALT   = "F5F0E8"
BEIGE_WHITE = "FEFCF9"
BEIGE_DARK  = "D4C5A9"
GOLD        = "C9A96E"
GREEN       = "4A7A3E"
BLUE_GREY   = "6B7E8F"


def _border(color=BEIGE_DARK):
    s = Side(style="thin", color=color)
    return Border(left=s, right=s, top=s, bottom=s)


def _hdr_cell(ws, row, col, value, width=None):
    c = ws.cell(row=row, column=col, value=value)
    c.font  = Font(bold=True, color="FFFFFF", size=10)
    c.fill  = PatternFill("solid", fgColor=BROWN)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = _border("FFFFFF")
    if width and ws.column_dimensions[get_column_letter(col)].width < width:
        ws.column_dimensions[get_column_letter(col)].width = width
    return c


def build_excel(aggregated: list[dict], meta: dict, raw_items: list[dict]) -> bytes:
    wb = openpyxl.Workbook()

    # ── Sheet 1: Totalen per product ──────────────────────────────────────
    ws1 = wb.active
    ws1.title = "Totalen per product"
    _build_summary_sheet(ws1, aggregated, meta)

    # ── Sheet 2: Detail (alle ruwe regels) ────────────────────────────────
    ws2 = wb.create_sheet("Detail per pagina")
    _build_detail_sheet(ws2, raw_items, meta)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def _build_summary_sheet(ws, aggregated, meta):
    # Title
    ws.merge_cells("A1:H1")
    t = ws["A1"]
    t.value = "Intrastat – Totalen per product"
    t.font  = Font(bold=True, size=14, color="FFFFFF")
    t.fill  = PatternFill("solid", fgColor=BROWN)
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    # Meta
    def meta_cell(rng, val):
        ws.merge_cells(rng)
        start = rng.split(":")[0]
        c = ws[start]
        c.value = val
        c.font  = Font(italic=True, size=9, color="5C4A32")
        c.fill  = PatternFill("solid", fgColor="EDE5D0")
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    meta_cell("A2:C2", f"Leverancier: {meta.get('leverancier','')}")
    meta_cell("D2:F2", f"Factuurnummer: {meta.get('factuurnummer','')}")
    meta_cell("G2:H2", f"Datum: {meta.get('datum','')}  |  Export: {datetime.today().strftime('%d/%m/%Y')}")
    ws.row_dimensions[2].height = 16

    ws.append([])  # blank row

    # Headers (row 4)
    COLS = [
        ("Product / Artikel",        32),
        ("Totale hoeveelheid",       18),
        ("Eenheid",                  10),
        ("Eenheidsprijs (€)",        16),
        ("Totale waarde (€)",        18),
        ("Verschijnt op pagina's",   22),
        ("Aantal regels",            13),
        ("Notitie prijs",            13),
    ]
    HR = 4
    for ci, (label, w) in enumerate(COLS, 1):
        _hdr_cell(ws, HR, ci, label, w)
    ws.row_dimensions[HR].height = 32

    # Data
    DS = HR + 1
    fmt_qty  = '#,##0.##'
    fmt_eur  = '#,##0.00 "€"'

    for ri, item in enumerate(aggregated):
        r = DS + ri
        fill = PatternFill("solid", fgColor=BEIGE_ALT if ri % 2 == 0 else BEIGE_WHITE)
        brd  = _border()

        paginas_str = ", ".join(str(p) for p in item["paginas"])

        vals = [
            item["artikel"],
            item["hoeveelheid"],
            item["eenheid"],
            item["eenheidsprijs"],
            item["bedrag"],
            paginas_str,
            item["regels"],
            item["price_note"],
        ]

        for ci, v in enumerate(vals, 1):
            c = ws.cell(row=r, column=ci, value=v)
            c.fill   = fill
            c.border = brd
            c.alignment = Alignment(vertical="center", wrap_text=(ci == 1))
            if ci == 2 and isinstance(v, (int, float)):
                c.number_format = fmt_qty
            if ci in (4, 5) and isinstance(v, (int, float)):
                c.number_format = fmt_eur

        ws.row_dimensions[r].height = 18

    # Totals row
    TR = DS + len(aggregated)
    tot_fill = PatternFill("solid", fgColor=GOLD)
    tot_font = Font(bold=True, size=10)

    for ci in range(1, 9):
        c = ws.cell(row=TR, column=ci)
        c.fill   = tot_fill
        c.border = _border()
        c.font   = tot_font
        c.alignment = Alignment(horizontal="center", vertical="center")

    ws.cell(row=TR, column=1, value="TOTAAL")
    # Sum total waarde
    col_e = get_column_letter(5)
    ws.cell(row=TR, column=5).value = f"=SUM({col_e}{DS}:{col_e}{TR-1})"
    ws.cell(row=TR, column=5).number_format = fmt_eur
    ws.cell(row=TR, column=7).value = f"=SUM(G{DS}:G{TR-1})"
    ws.row_dimensions[TR].height = 22

    # Freeze & filter
    ws.freeze_panes = ws.cell(row=DS, column=1)
    ws.auto_filter.ref = f"A{HR}:{get_column_letter(len(COLS))}{TR-1}"

    # Footer
    nr = TR + 2
    ws.merge_cells(f"A{nr}:H{nr}")
    fc = ws[f"A{nr}"]
    fc.value = (
        f"Gegenereerd door Intrastat Factuur App  •  "
        f"{datetime.today().strftime('%d/%m/%Y %H:%M')}  •  © Lander Smits"
    )
    fc.font = Font(italic=True, size=8, color="9C8567")
    fc.alignment = Alignment(horizontal="center")


def _build_detail_sheet(ws, raw_items, meta):
    ws.merge_cells("A1:G1")
    t = ws["A1"]
    t.value = "Detail – alle regels per pagina"
    t.font  = Font(bold=True, size=12, color="FFFFFF")
    t.fill  = PatternFill("solid", fgColor=BLUE_GREY)
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    ws.append([])

    COLS = [
        ("Pagina",                8),
        ("Product / Artikel",    32),
        ("Hoeveelheid",          14),
        ("Eenheid",              10),
        ("Eenheidsprijs (€)",    16),
        ("Bedrag (€)",           14),
    ]
    HR = 3
    for ci, (label, w) in enumerate(COLS, 1):
        c = ws.cell(row=HR, column=ci, value=label)
        c.font  = Font(bold=True, color="FFFFFF", size=10)
        c.fill  = PatternFill("solid", fgColor=BLUE_GREY)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = _border("FFFFFF")
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.row_dimensions[HR].height = 28

    DS = HR + 1
    fmt_eur = '#,##0.00 "€"'
    fmt_qty = '#,##0.##'

    for ri, item in enumerate(raw_items):
        r = DS + ri
        fill = PatternFill("solid", fgColor=BEIGE_ALT if ri % 2 == 0 else BEIGE_WHITE)
        brd  = _border()

        vals = [
            item.get("pagina"),
            item.get("artikel"),
            item.get("hoeveelheid"),
            item.get("eenheid"),
            item.get("eenheidsprijs"),
            item.get("bedrag"),
        ]

        for ci, v in enumerate(vals, 1):
            c = ws.cell(row=r, column=ci, value=v)
            c.fill   = fill
            c.border = brd
            c.alignment = Alignment(vertical="center", horizontal="center" if ci in (1, 3, 4) else "left")
            if ci == 3 and isinstance(v, (int, float)):
                c.number_format = fmt_qty
            if ci in (5, 6) and isinstance(v, (int, float)):
                c.number_format = fmt_eur

        ws.row_dimensions[r].height = 16

    ws.freeze_panes = ws.cell(row=DS, column=1)
    ws.auto_filter.ref = f"A{HR}:{get_column_letter(len(COLS))}{DS + len(raw_items) - 1}"


# ---------------------------------------------------------------------------
# Flask routes
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    if "pdf" not in request.files:
        return jsonify({"error": "Geen bestand ontvangen."}), 400

    file = request.files["pdf"]
    if not file.filename or not file.filename.lower().endswith(".pdf"):
        return jsonify({"error": "Upload een geldig PDF-bestand."}), 400

    pdf_bytes = file.read()
    if not pdf_bytes:
        return jsonify({"error": "Het bestand is leeg."}), 400

    try:
        aggregated, meta, raw_items = process_pdf(pdf_bytes)
    except Exception as exc:
        logger.exception("PDF verwerking mislukt")
        return jsonify({"error": f"Fout bij verwerken: {exc}"}), 500

    if not aggregated:
        return jsonify({
            "error": (
                "Geen artikelregels gevonden. Controleer of de PDF leesbare "
                "tekst bevat (niet enkel gescande afbeeldingen)."
            )
        }), 422

    safe = secure_filename(file.filename).replace(".pdf", "")
    xlsx_name = f"intrastat_{safe}_{datetime.today().strftime('%Y%m%d')}.xlsx"

    try:
        excel_bytes = build_excel(aggregated, meta, raw_items)
    except Exception as exc:
        logger.exception("Excel aanmaken mislukt")
        return jsonify({"error": f"Excel fout: {exc}"}), 500

    token = str(uuid.uuid4())
    RESULT_STORE[token] = (excel_bytes, xlsx_name)

    return jsonify({
        "token":       token,
        "filename":    xlsx_name,
        "item_count":  len(aggregated),
        "raw_count":   len(raw_items),
        "meta":        meta,
    })


@app.route("/download/<token>")
def download(token: str):
    result = RESULT_STORE.pop(token, None)
    if result is None:
        abort(404)
    excel_bytes, filename = result
    return send_file(
        io.BytesIO(excel_bytes),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=filename,
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
