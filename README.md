# Intrastat Factuur Import App

Upload een aankoopfactuur (PDF) en ontvang automatisch een Intrastat-klaar Excel-bestand met alle artikelregels en bedragen.

## Lokaal draaien

```bash
pip install -r requirements.txt
python app.py
```

Open dan http://localhost:5000 in je browser.

## Online deployen (Render.com — gratis tier)

1. Push deze repo naar GitHub
2. Ga naar https://render.com → New → Web Service
3. Kies je GitHub repo
4. Render detecteert de `Procfile` automatisch
5. Start Command: `gunicorn app:app`
6. Klik **Deploy**

## Online deployen (Railway)

1. Push naar GitHub
2. Ga naar https://railway.app → New Project → Deploy from GitHub
3. Railway detecteert Python automatisch via `requirements.txt`
4. Deploy!

## Wat doet de app?

- Leest tabellen én vrijlopende tekst uit de PDF
- Herkent artikelregels op basis van kolomnamen (NL/EN)
- Maakt een gestijld Excel-bestand met Intrastat-kolommen:
  - Artikelcode, Omschrijving, Goederencode (CN), Land van Oorsprong,
    Hoeveelheid, Eenheid, Netto massa, Eenheidsprijs, Factuurwaarde
- Totaalrij met SUM-formules en autofilter

---

© Lander Smits