# Intrastat PDF Extractor

**Wat doet dit?**
Je uploadt een aankoopfactuur als PDF. De app leest automatisch alle artikellijnen, groepeert ze per goederencode (CN-code) en genereert een kant-en-klaar Excel-bestand in het NBB Intrastat-formaat, plus een CSV voor rechtstreekse import. De app controleert ook of de berekende totalen overeenkomen met het factuurtotaal.

Gemaakt door **Lander Smits**.

---

## Wat heb je nodig?

- Een computer met **Windows, Mac of Linux**
- **Python 3.10 of nieuwer** — gratis te downloaden op [python.org](https://www.python.org/downloads/)
- Een internetverbinding (alleen voor de installatie)

---

## Stap 1 — Installeer Python

> Sla deze stap over als Python al geïnstalleerd is.

1. Ga naar [python.org/downloads](https://www.python.org/downloads/)
2. Klik op de grote gele knop **"Download Python 3.x.x"**
3. Open het gedownloade bestand en volg de installatie
   - ⚠️ **Belangrijk voor Windows:** vink onderaan het installatiescherm **"Add Python to PATH"** aan vóór je op Install klikt

Om te controleren of Python correct geïnstalleerd is, open je een terminal en typ:
```
python --version
```
Je zou iets moeten zien als `Python 3.12.0`.

---

## Stap 2 — Download de app

### Optie A — Via Git (aanbevolen)
Open een terminal en typ:
```
git clone https://github.com/lander-sm07/factuur-app.git
cd factuur-app
```

### Optie B — Zonder Git
1. Ga naar [github.com/lander-sm07/factuur-app](https://github.com/lander-sm07/factuur-app)
2. Klik op de groene knop **"Code"** → **"Download ZIP"**
3. Pak het ZIP-bestand uit
4. Open een terminal en navigeer naar de uitgepakte map:
   ```
   cd pad/naar/factuur-app
   ```

---

## Stap 3 — Installeer de benodigde pakketten

Typ dit in de terminal (éénmalig):

**Windows:**
```
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

**Mac / Linux:**
```
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Je ziet een hoop tekst voorbijkomen terwijl de pakketten installeren. Dat is normaal. Wacht tot het klaar is.

---

## Stap 4 — Start de app

**Windows:**
```
.venv\Scripts\activate
python app.py
```

**Mac / Linux:**
```
source .venv/bin/activate
python3 app.py
```

Je ziet iets als:
```
 * Running on http://127.0.0.1:5000
```

---

## Stap 5 — Gebruik de app

1. Open je browser (Chrome, Edge, Firefox, ...)
2. Ga naar **[http://localhost:5000](http://localhost:5000)**
3. Sleep je PDF-factuur naar het uploadvenster of klik om te bladeren
4. Klik op **"Generate Intrastat Excel"**
5. Download het Excel-bestand (.xlsx) of het CSV-bestand voor NBB-import

---

## De app stoppen

Klik in de terminal op **Ctrl + C**.

---

## De volgende keer opstarten

Je hoeft stap 1–3 maar één keer te doen. Daarna is het gewoon:

**Windows:**
```
cd factuur-app
.venv\Scripts\activate
python app.py
```

**Mac / Linux:**
```
cd factuur-app
source .venv/bin/activate
python3 app.py
```

---

## Problemen?

| Fout | Oplossing |
|------|-----------|
| `python` werkt niet | Probeer `python3` in plaats van `python` |
| `pip` werkt niet | Probeer `pip3` in plaats van `pip` |
| Poort 5000 bezet | Wijzig de poort: `python app.py` → de app kijkt naar de `PORT` omgevingsvariabele, of open `app.py` en verander `5000` naar `5001` |
| Pagina laadt niet | Controleer of de terminal nog actief is en geen foutmelding toont |

---

## Technische details (voor wie het wil weten)

| Onderdeel | Wat |
|-----------|-----|
| Webframework | Flask |
| PDF-uitlezen | pdfplumber |
| Excel aanmaken | openpyxl |
| Taal | Python 3 |

De vaste NBB-waarden (gewest, land, incoterm, ...) staan bovenaan `app.py` en zijn eenvoudig aan te passen.
