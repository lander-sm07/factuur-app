import io
import os
import re
import uuid
import logging
from datetime import datetime

import pdfplumber
import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side
)
from openpyxl.utils import get_column_letter
from flask import (
    Flask, request, jsonify, send_file,
    render_template, abort
)
from werkzeug.utils import secure_filename

# ---------------------------------------------------------------------------
# App setup
# ---------------------------------------------------------------------------
app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024  # 20 MB max upload

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

ALLOWED_EXTENSIONS = {"pdf"}

# In-memory store for generated Excel files (keyed by a UUID token)
RESULT_STORE: dict[str, bytes] = {}


def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


# ---------------------------------------------------------------------------
# PDF parsing helpers
# ---------------------------------------------------------------------------

# Regex patterns (case-insensitive) used to locate header rows in text tables
SKIP_PATTERNS = re.compile(
    r"(subtotaal|totaal|btw|vat|discount|korting|verzend|levering|"
    r"shipping|tax|total|sub-total|total\s+ht|tva|transport|porto)",
    re.IGNORECASE,
)

AMOUNT_RE = re.compile(r"[\d\s]{1,10}[.,]\d{2}")
NUMBER_RE = re.compile(r"^\d+([.,]\d+)?$")
CODE_RE   = re.compile(r"^[A-Z0-9\-]{3,20}$")


def clean_number(val: str) -> float | None:
    """Convert European number strings like '1.234,56' or '1234.56' to float."""
    if val is None:
        return None
    val = str(val).strip().replace(" ", "").replace("\xa0", "")
    # Remove currency symbols
    val = re.sub(r"[€$£]", "", val)
    if not val:
        return None
    # European format: 1.234,56
    if re.search(r"\d\.\d{3},", val):
        val = val.replace(".", "").replace(",", ".")
    # Only comma as decimal: 1234,56
    elif val.count(",") == 1 and val.count(".") == 0:
        val = val.replace(",", ".")
    # Comma as thousands: 1,234.56
    elif val.count(",") >= 1 and val.count(".") == 1:
        val = val.replace(",", "")
    try:
        return float(val)
    except ValueError:
        return None


def parse_table_rows(table: list[list]) -> list[dict]:
    """
    Try to interpret a pdfplumber table as invoice line items.
    Returns a list of dicts with keys: code, omschrijving, hoeveelheid,
    eenheid, eenheidsprijs, bedrag.
    """
    if not table or len(table) < 2:
        return []

    # Normalise cells: strip whitespace, replace None with ""
    rows = []
    for row in table:
        rows.append([str(c).strip() if c is not None else "" for c in row])

    # Detect header row (first row that looks like column titles)
    header_idx = 0
    header_keywords = re.compile(
        r"(omschr|descri|artikel|product|ref|code|qty|hoeveelh|aantal|"
        r"prijs|price|bedrag|amount|total|eenheid|unit|gewicht|weight)",
        re.IGNORECASE,
    )
    for i, row in enumerate(rows[:5]):
        if sum(1 for cell in row if header_keywords.search(cell)) >= 2:
            header_idx = i
            break

    headers = rows[header_idx]
    data_rows = rows[header_idx + 1:]

    # Map column indices to semantic roles
    col_map = {
        "code": None,
        "omschrijving": None,
        "hoeveelheid": None,
        "eenheid": None,
        "eenheidsprijs": None,
        "bedrag": None,
        "cn_code": None,
        "land": None,
        "massa": None,
    }

    for i, h in enumerate(headers):
        h_low = h.lower()
        if re.search(r"omschr|descri|product|goed|artikel\s*naam", h_low):
            col_map["omschrijving"] = i
        elif re.search(r"artikel|ref|code|sku|nr\.?$|nr\s", h_low):
            col_map["code"] = i
        elif re.search(r"hoeveelh|aantal|qty|quantit", h_low):
            col_map["hoeveelheid"] = i
        elif re.search(r"eenheid|unit", h_low):
            col_map["eenheid"] = i
        elif re.search(r"eenheidsprijs|stuksprijs|unit\s*price|prijs\s*p", h_low):
            col_map["eenheidsprijs"] = i
        elif re.search(r"bedrag|totaal|total|amount|waarde", h_low):
            col_map["bedrag"] = i
        elif re.search(r"cn|hs\s*code|goederencode|tarief", h_low):
            col_map["cn_code"] = i
        elif re.search(r"land|country|origine|origin", h_low):
            col_map["land"] = i
        elif re.search(r"massa|gewicht|weight", h_low):
            col_map["massa"] = i

    # Heuristic fallback: if no bedrag column found, use last numeric column
    if col_map["bedrag"] is None and col_map["omschrijving"] is not None:
        for i in range(len(headers) - 1, -1, -1):
            if i != col_map.get("hoeveelheid") and i != col_map.get("eenheidsprijs"):
                col_map["bedrag"] = i
                break

    items = []
    for row in data_rows:
        if not any(row):
            continue
        # Skip totals / summary rows
        combined = " ".join(row)
        if SKIP_PATTERNS.search(combined):
            continue
        # Skip rows with no numeric value at all
        if not re.search(r"\d", combined):
            continue

        def get(key):
            idx = col_map.get(key)
            if idx is not None and idx < len(row):
                return row[idx]
            return ""

        item = {
            "code": get("code"),
            "omschrijving": get("omschrijving"),
            "hoeveelheid": clean_number(get("hoeveelheid")),
            "eenheid": get("eenheid"),
            "eenheidsprijs": clean_number(get("eenheidsprijs")),
            "bedrag": clean_number(get("bedrag")),
            "cn_code": get("cn_code"),
            "land": get("land"),
            "massa": clean_number(get("massa")),
        }

        # Only keep row if at least omschrijving OR bedrag is present
        if item["omschrijving"] or item["bedrag"] is not None:
            items.append(item)

    return items


def parse_text_fallback(text: str) -> list[dict]:
    """
    Regex-based fallback parser for PDFs without recognisable tables.
    Looks for lines that contain both a description and at least one amount.
    """
    items = []
    lines = text.split("\n")

    for line in lines:
        line = line.strip()
        if not line or len(line) < 5:
            continue
        if SKIP_PATTERNS.search(line):
            continue

        amounts = AMOUNT_RE.findall(line)
        if not amounts:
            continue

        # Extract the last amount as the line total
        last_amount = clean_number(amounts[-1])
        if last_amount is None or last_amount <= 0:
            continue

        # Everything before the first amount is the description
        first_amt_pos = line.index(amounts[0])
        description = line[:first_amt_pos].strip()
        if len(description) < 3:
            continue

        # Try to extract quantity (first standalone number before description)
        qty = None
        qty_match = re.match(r"^(\d+(?:[.,]\d+)?)\s+", description)
        if qty_match:
            qty = clean_number(qty_match.group(1))
            description = description[qty_match.end():].strip()

        # Unit price: second-to-last amount if >= 2 amounts
        unit_price = None
        if len(amounts) >= 2:
            unit_price = clean_number(amounts[-2])

        items.append(
            {
                "code": "",
                "omschrijving": description,
                "hoeveelheid": qty,
                "eenheid": "",
                "eenheidsprijs": unit_price,
                "bedrag": last_amount,
                "cn_code": "",
                "land": "",
                "massa": None,
            }
        )

    return items


def extract_invoice_meta(text: str) -> dict:
    """Pull basic invoice header info from raw text."""
    meta = {"factuurnummer": "", "datum": "", "leverancier": ""}

    inv_match = re.search(
        r"(factuur\s*(?:nr|nummer|no)?\.?\s*:?\s*)([A-Z0-9\-/]{3,20})",
        text, re.IGNORECASE
    )
    if inv_match:
        meta["factuurnummer"] = inv_match.group(2).strip()

    date_match = re.search(
        r"(\d{1,2}[./-]\d{1,2}[./-]\d{2,4}|\d{4}[./-]\d{2}[./-]\d{2})",
        text
    )
    if date_match:
        meta["datum"] = date_match.group(1)

    # First non-empty line often contains the supplier name
    for line in text.split("\n"):
        line = line.strip()
        if line and len(line) > 3 and not re.match(r"^(factuur|invoice|rekening)", line, re.IGNORECASE):
            meta["leverancier"] = line
            break

    return meta


def process_pdf(pdf_bytes: bytes) -> tuple[list[dict], dict]:
    """
    Main entry point: parse a PDF and return (line_items, meta).
    Tries table extraction per page first; falls back to text parsing.
    """
    items: list[dict] = []
    full_text = ""

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            # 1. Table extraction
            tables = page.extract_tables()
            for table in tables:
                rows = parse_table_rows(table)
                items.extend(rows)

            # 2. Accumulate text for meta + fallback
            page_text = page.extract_text() or ""
            full_text += page_text + "\n"

    # If table extraction found nothing meaningful, use text fallback
    if not items:
        logger.info("No table items found, using text fallback parser")
        items = parse_text_fallback(full_text)

    # Deduplicate identical rows
    seen = set()
    unique_items = []
    for item in items:
        key = (item.get("omschrijving", ""), item.get("bedrag"))
        if key not in seen:
            seen.add(key)
            unique_items.append(item)

    meta = extract_invoice_meta(full_text)
    return unique_items, meta


# ---------------------------------------------------------------------------
# Excel generation
# ---------------------------------------------------------------------------

HEADER_BG   = "8B7355"   # warm brown
ALT_ROW_BG  = "F5F0E8"   # light beige
WHITE       = "FAFAF7"
BORDER_CLR  = "D4C5A9"
TOTAL_BG    = "C9A96E"   # golden beige


def _thin_border(color=BORDER_CLR):
    side = Side(style="thin", color=color)
    return Border(left=side, right=side, top=side, bottom=side)


def build_excel(items: list[dict], meta: dict, filename: str) -> bytes:
    """Generate a styled Intrastat Excel workbook and return as bytes."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Intrastat"

    # ── Title block ──────────────────────────────────────────────────────────
    ws.merge_cells("A1:I1")
    title_cell = ws["A1"]
    title_cell.value = "Intrastat – Aankoopfacturen"
    title_cell.font = Font(bold=True, size=14, color="FFFFFF")
    title_cell.fill = PatternFill("solid", fgColor=HEADER_BG)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    # Meta row
    ws.merge_cells("A2:C2")
    ws["A2"].value = f"Leverancier: {meta.get('leverancier', '')}"
    ws.merge_cells("D2:F2")
    ws["D2"].value = f"Factuurnummer: {meta.get('factuurnummer', '')}"
    ws.merge_cells("G2:I2")
    ws["G2"].value = f"Datum: {meta.get('datum', '')}   |   Export: {datetime.today().strftime('%d/%m/%Y')}"
    for col in ["A2", "D2", "G2"]:
        cell = ws[col]
        cell.font = Font(italic=True, size=10, color="5C4A32")
        cell.fill = PatternFill("solid", fgColor="EDE5D0")
        cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[2].height = 18

    ws.append([])  # blank row

    # ── Column headers ───────────────────────────────────────────────────────
    COLUMNS = [
        ("Artikelcode",        14),
        ("Omschrijving",       38),
        ("Goederencode (CN)",  18),
        ("Land v. Oorsprong",  18),
        ("Hoeveelheid",        13),
        ("Eenheid",            10),
        ("Netto massa (kg)",   16),
        ("Eenheidsprijs (€)",  16),
        ("Factuurwaarde (€)",  18),
    ]

    header_fill = PatternFill("solid", fgColor=HEADER_BG)
    header_font = Font(bold=True, color="FFFFFF", size=10)
    header_row = 4

    for col_idx, (label, width) in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=header_row, column=col_idx, value=label)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = _thin_border("FFFFFF")
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    ws.row_dimensions[header_row].height = 30

    # ── Data rows ────────────────────────────────────────────────────────────
    data_start = header_row + 1
    number_fmt = '#,##0.00 "€"'
    qty_fmt    = '#,##0.##'

    for row_idx, item in enumerate(items):
        r = data_start + row_idx
        fill_color = ALT_ROW_BG if row_idx % 2 == 0 else WHITE
        row_fill = PatternFill("solid", fgColor=fill_color)
        border = _thin_border()

        values = [
            item.get("code", ""),
            item.get("omschrijving", ""),
            item.get("cn_code", ""),
            item.get("land", ""),
            item.get("hoeveelheid"),
            item.get("eenheid", ""),
            item.get("massa"),
            item.get("eenheidsprijs"),
            item.get("bedrag"),
        ]

        for col_idx, val in enumerate(values, start=1):
            cell = ws.cell(row=r, column=col_idx, value=val)
            cell.fill = row_fill
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=(col_idx == 2))

            # Numeric formatting
            if col_idx in (5, 7):   # qty / massa
                if isinstance(val, (int, float)):
                    cell.number_format = qty_fmt
            if col_idx in (8, 9):   # prices
                if isinstance(val, (int, float)):
                    cell.number_format = number_fmt

        ws.row_dimensions[r].height = 18

    # ── Totals row ───────────────────────────────────────────────────────────
    total_row = data_start + len(items)
    total_fill = PatternFill("solid", fgColor=TOTAL_BG)
    total_font = Font(bold=True, size=10, color="3D3229")

    ws.cell(row=total_row, column=1, value="TOTAAL").font  = total_font
    ws.cell(row=total_row, column=1).fill = total_fill
    ws.cell(row=total_row, column=1).border = _thin_border()

    for col in range(2, 10):
        cell = ws.cell(row=total_row, column=col)
        cell.fill = total_fill
        cell.border = _thin_border()
        cell.font = total_font

    # SUM formulas for numeric columns
    if items:
        for col_idx, fmt in [(5, qty_fmt), (7, qty_fmt), (8, number_fmt), (9, number_fmt)]:
            col_letter = get_column_letter(col_idx)
            sum_cell = ws.cell(row=total_row, column=col_idx)
            sum_cell.value = f"=SUM({col_letter}{data_start}:{col_letter}{total_row - 1})"
            sum_cell.number_format = fmt
            sum_cell.font = total_font
            sum_cell.fill = total_fill
            sum_cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.row_dimensions[total_row].height = 22

    # ── Freeze panes & auto-filter ───────────────────────────────────────────
    ws.freeze_panes = ws.cell(row=data_start, column=1)
    ws.auto_filter.ref = (
        f"A{header_row}:{get_column_letter(len(COLUMNS))}{total_row - 1}"
    )

    # ── Footer note ──────────────────────────────────────────────────────────
    note_row = total_row + 2
    ws.merge_cells(f"A{note_row}:I{note_row}")
    note_cell = ws[f"A{note_row}"]
    note_cell.value = (
        f"Gegenereerd door Intrastat Factuur App  •  {datetime.today().strftime('%d/%m/%Y %H:%M')}  •  © Lander Smits"
    )
    note_cell.font = Font(italic=True, size=8, color="9C8567")
    note_cell.alignment = Alignment(horizontal="center")

    # ── Save to bytes ─────────────────────────────────────────────────────────
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    if "pdf" not in request.files:
        return jsonify({"error": "Geen bestand ontvangen."}), 400

    file = request.files["pdf"]
    if file.filename == "" or not allowed_file(file.filename):
        return jsonify({"error": "Ongeldig bestandstype. Upload een PDF-bestand."}), 400

    pdf_bytes = file.read()
    if len(pdf_bytes) == 0:
        return jsonify({"error": "Het geüploade bestand is leeg."}), 400

    try:
        items, meta = process_pdf(pdf_bytes)
    except Exception as exc:
        logger.exception("PDF parsing failed")
        return jsonify({"error": f"Fout bij het verwerken van de PDF: {exc}"}), 500

    if not items:
        return jsonify({
            "error": (
                "Geen artikelregels gevonden in de PDF. "
                "Controleer of de factuur leesbare tekst of tabellen bevat."
            )
        }), 422

    # Build Excel
    safe_name = secure_filename(file.filename).replace(".pdf", "")
    excel_filename = f"intrastat_{safe_name}_{datetime.today().strftime('%Y%m%d')}.xlsx"

    try:
        excel_bytes = build_excel(items, meta, excel_filename)
    except Exception as exc:
        logger.exception("Excel generation failed")
        return jsonify({"error": f"Fout bij het aanmaken van het Excel-bestand: {exc}"}), 500

    # Store result in memory with a unique token
    token = str(uuid.uuid4())
    RESULT_STORE[token] = (excel_bytes, excel_filename)

    return jsonify({
        "token": token,
        "filename": excel_filename,
        "item_count": len(items),
        "meta": meta,
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
