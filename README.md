## Object Description Generator

Dieses Projekt erstellt automatisch zweisprachige Katalogtexte (Deutsch/Englisch) für Objekte aus Excel-Listen anhand zugehöriger Bilder.  
Die Ergebnisse werden in einer Excel-Datei gespeichert.

---

## Projektstruktur

TMB-PIPELINE
── Objektbilder
── Liste_AEG Produktsammlung.xls
── Liste_AK Kommunikation.xls
── main.py
── .env
── object_descriptions.xlsx

---

## Voraussetzungen

- Python 3.10+
- Installation der Abhängigkeiten:

```bash
pip install pandas requests python-dotenv tqdm openpyxl
```

---

## Konfiguration

In einer .env Datei muss folgender Eintrag vorhanden sein:

OPENAI_API_KEY=der_api_key

---

## Verwendung

Skript starten:

python main.py

Während der Laufzeit wird der Fortschritt angezeigt.
Nach Abschluss liegt die Datei object_descriptions.xlsx mit den Ergebnissen vor.

---

## Konfigurierbare Parameter

Im Skript main.py können folgende Werte angepasst werden:
MAX_IMAGES = 3     
MAX_TOKENS = 1200  
API_WAIT = 10