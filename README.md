# Intrastat PDF Extractor

Een Flask webapplicatie die aankoopfacturen (PDF) omzet naar NBB Intrastat-bestanden (Excel + CSV).  
Gemaakt door **Lander Smits**.

---

## Wat doet het?

- Upload een PDF-factuur (met QUANTITY / VALUE / CUSTOMS-kolommen)
- De app haalt alle productlijnen op, groepeert ze per CN/goederencode en telt bedragen, gewichten en hoeveelheden op
- Output: een kant-en-klaar `.xlsx` bestand in NBB Intrastat-formaat + een `.csv` voor directe import
- Automatische dubbelcheck: som van lijnen vs. factuurtotaal uit de PDF

---

## Gratis hosten op Render

### Eenmalige setup (5 minuten)

1. **Maak een gratis account aan op [render.com](https://render.com)**

2. **Klik op "New +" → "Web Service"**

3. **Verbind je GitHub-repo:**  
   Kies `lander-sm07/factuur-app` (of geef Render toegang tot je GitHub-account)

4. **Instellingen** (Render detecteert dit automatisch via `render.yaml`):
   | Veld | Waarde |
   |------|--------|
   | Name | `intrastat-pdf-extractor` |
   | Runtime | `Python 3` |
   | Build Command | `pip install -r requirements.txt` |
   | Start Command | `gunicorn app:app` |
   | Plan | **Free** |

5. **Klik "Create Web Service"** — Render bouwt en deployt automatisch

6. Na ~2 minuten krijg je een URL zoals:  
   `https://intrastat-pdf-extractor.onrender.com`

### Updates deployen

Elke keer dat je naar GitHub pusht wordt de app automatisch hergebouwd:

```bash
git add -A
git commit -m "update"
git push
```

### Let op bij de gratis tier

- De app **gaat slapen** na 15 minuten zonder gebruik — de eerste request duurt dan ~30 seconden
- Voor actief gebruik is de gratis tier prima; voor productie upgrade je naar de "Starter" tier ($7/maand)

---

## Lokaal draaien

```bash
# Installeer dependencies
pip install -r requirements.txt

# Start de app
python app.py
# of via gunicorn:
gunicorn app:app

# Open in browser
# http://localhost:5000
```

---

## Technische stack

| Component | Bibliotheek |
|-----------|------------|
| Webframework | Flask 3.0 |
| PDF-extractie | pdfplumber |
| Excel-output | openpyxl |
| Productieserver | gunicorn |

---

## Configuratie

Vaste NBB-waarden staan bovenaan `app.py` en zijn eenvoudig aan te passen:

```python
FIXED_GEWEST   = "1 Vlaams gewest"
FIXED_LAND     = "Italie"
FIXED_AARD     = "Rechtstreekse verkoop/aankoop..."
FIXED_INCOTERM = "EXW Af fabriek"
FIXED_VERVOER  = "3 Wegvervoer"
```

---

## Licentie

Privéproject — © Lander Smits
