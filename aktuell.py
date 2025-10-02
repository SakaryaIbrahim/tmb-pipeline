import os
import base64
import pandas as pd
import requests
import time
from dotenv import load_dotenv

# ----------------------------------------------------------
# Hilfsfunktionen
# ----------------------------------------------------------
def encode_image(image_path):
    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode("utf-8")

def safe_value(val):
    """
    Ersetzt NaN/leer durch 'Leere Zelle', sonst String.
    """
    if val is None:
        return "Leere Zelle"
    try:
        s = str(val).strip()
    except Exception:
        return "Leere Zelle"
    return "Leere Zelle" if s == "" or s.lower() == "nan" else s

def get_val(row, colname_lower):
    """
    Liest einen Wert aus der (bereits normalisierten) Zeile.
    colname_lower muss lowercase sein (weil wir df.columns normalisieren).
    """
    if colname_lower in row.index:
        return safe_value(row[colname_lower])
    return "Leere Zelle"

# ----------------------------------------------------------
# Setup
# ----------------------------------------------------------
load_dotenv()
API_KEY = os.getenv("OPENAI_API_KEY")
if not API_KEY:
    raise ValueError("❌ Kein OPENAI_API_KEY gefunden! Bitte .env prüfen.")

OPENAI_URL = "https://api.openai.com/v1/chat/completions"

# ----------------------------------------------------------
# Prompt-Template
# ----------------------------------------------------------
BASE_PROMPT_TEMPLATE = """
You are a museum cataloguer for the Technisches Museum Berlin. Describe only what the images justify; be conservative; no speculation.
The input includes metadata from an Excel row. If any field value equals the literal string "Leere Zelle", treat it as missing/absent.

Object ID: {object_id}
Title: {title}
Image List (local selection): {image_list}
Manufacturer: {manufacturer}
Material: {material}
Dimensions: {dimensions}
Year: {year}

Use title to choose the plain object type and function wording in Paragraph 1 if consistent with the images.
Facts from manufacturer, material, dimensions, year are authoritative list data and must be included in Paragraph 2 as a clearly labeled sentence beginning with „Angaben laut Liste:“ (German) and “According to the list:” (English).
Copy their values verbatim; do not rephrase.
Never invent missing values. If a field is absent or equals "Leere Zelle", omit it entirely from that list sentence.
Do not state date, maker, materials, or dimensions as if they were seen on the object unless they are visibly inscribed. For visible inscriptions, transcribe verbatim in quotes; use […] for illegible parts.
Ban hedging/guessing: avoid vermutlich, wahrscheinlich, wohl, offenbar, anscheinend, möglicherweise, könnte, dürfte, early/late, typical, domestic, etc.

OUTPUT (exactly two blocks, in this order)

1) CATALOGUE TEXT (DE → EN)
Write 2–3 short paragraphs in German, then provide an English translation with the same constraints.

Paragraph 1 (DE): Name the object type in plain words and its function/purpose (use title if it matches the images).

Paragraph 2 (DE): Strictly image-based description: form and construction (shape, handles, openings, moving parts, connectors); colours/finish; materials only if visually evident; transcribe any visible labels/marks verbatim (use […] for gaps).
At the end of Paragraph 2 (DE), add ONE sentence with Excel facts using EXACT formatting, but include only the fields that exist (omit any that are "Leere Zelle"):
Angaben laut Liste: {de_list_sentence}.

Paragraph 3 (DE): Condition notes only (e.g., Kratzer, Korrosion, Abplatzungen, fehlende Teile). No usage scenarios or history.

Then provide the EN translation mirroring the structure above.
In the English version of Paragraph 2, reproduce the list sentence as:
According to the list: {en_list_sentence}.

2) CAPTION (German)
One sentence (≤35 Wörter) suitable for a label: Objekttyp, Funktion/Zweck, sichtbares Material/Finish, datierungslos.
Do not include maker/date/provenance here unless they are visibly inscribed; Excel facts remain only in the “Angaben laut Liste” sentence within the description.
"""

# ----------------------------------------------------------
# Prompt bauen
# ----------------------------------------------------------
def build_list_sentences(manufacturer, material, dimensions, year):
    parts_de, parts_en = [], []
    if manufacturer != "Leere Zelle":
        parts_de.append(f"Hersteller: {manufacturer}")
        parts_en.append(f"Manufacturer: {manufacturer}")
    if material != "Leere Zelle":
        parts_de.append(f"Material: {material}")
        parts_en.append(f"Material: {material}")
    if dimensions != "Leere Zelle":
        parts_de.append(f"Maße: {dimensions}")
        parts_en.append(f"Dimensions: {dimensions}")
    if year != "Leere Zelle":
        parts_de.append(f"Jahr: {year}")
        parts_en.append(f"Year: {year}")
    return "; ".join(parts_de), "; ".join(parts_en)

def build_prompt(row_normalized, used_files, object_id_value):
    """
    row_normalized: Series mit lowercase/stripped Spaltennamen.
    object_id_value: kommt direkt aus dem Gruppier-Key (immer gesetzt).
    """
    # Objekt-ID direkt aus dem Group-Key, nicht durch safe_value verfälschen
    object_id = str(object_id_value).strip() if object_id_value is not None else "Leere Zelle"
    if object_id == "":
        object_id = "Leere Zelle"

    title        = get_val(row_normalized, "t10")
    manufacturer = get_val(row_normalized, "t2")
    material     = get_val(row_normalized, "t3")
    dimensions   = get_val(row_normalized, "t5")
    year         = get_val(row_normalized, "t14")

    image_list  = ", ".join([os.path.basename(p) for p in used_files]) if used_files else "Leere Zelle"

    de_list_sentence, en_list_sentence = build_list_sentences(manufacturer, material, dimensions, year)

    return BASE_PROMPT_TEMPLATE.format(
        object_id=object_id,
        title=title,
        image_list=image_list,
        manufacturer=manufacturer,
        material=material,
        dimensions=dimensions,
        year=year,
        de_list_sentence=de_list_sentence,
        en_list_sentence=en_list_sentence
    )

# ----------------------------------------------------------
# API Call
# ----------------------------------------------------------
def generate_catalog_text(image_paths, prompt_text, max_tokens=1200):
    image_contents = []
    for path in image_paths:
        base64_image = encode_image(path)
        image_contents.append({
            "type": "image_url",
            "image_url": {"url": f"data:image/jpeg;base64,{base64_image}"}
        })

    payload = {
        "model": "gpt-4o-mini",
        "messages": [{
            "role": "user",
            "content": [{"type": "text", "text": prompt_text}, *image_contents]
        }],
        "max_tokens": max_tokens
    }
    headers = {"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"}

    for attempt in range(3):
        try:
            response = requests.post(OPENAI_URL, headers=headers, json=payload, timeout=120)
            response.raise_for_status()
            return response.json()["choices"][0]["message"]["content"]
        except requests.exceptions.HTTPError as e:
            status = getattr(response, "status_code", None)
            if status == 429:
                wait = 5 * (attempt + 1)
                print(f"⏳ Rate limit erreicht, warte {wait}s...")
                time.sleep(wait)
                continue
            print(f"❌ API-Fehler: {e}")
            try:
                print("📩 Response:", response.text)
            except Exception:
                pass
            return f"❌ API-Fehler: {str(e)}"
        except requests.exceptions.RequestException as e:
            print(f"❌ Verbindungsfehler: {e}")
            return f"❌ Verbindungsfehler: {str(e)}"
    return "❌ Abgebrochen nach zu vielen Fehlversuchen"

# ----------------------------------------------------------
# Bildsuche
# ----------------------------------------------------------
def find_images_for_object(base_folder, object_id, max_images=3):
    year = next((p for p in object_id.split("/") if p.isdigit() and len(p) == 4), None)
    if not year:
        print("⚠️ Kein Jahr in Objekt-ID gefunden:", object_id)
        return []

    folder = os.path.join(base_folder, year)
    if not os.path.exists(folder):
        print(f"⚠️ Ordner nicht gefunden: {folder}")
        return []

    prefix = object_id.replace("/", "-").split()[0]
    candidates = [f for f in os.listdir(folder) if f.lower().endswith((".jpg", ".jpeg", ".png"))]
    matches = sorted([f for f in candidates if f.startswith(prefix)])

    use_count = min(len(matches), max_images)
    print(f"🔎 Suche in {folder} nach Prefix '{prefix}' → Gefunden: {len(matches)} → Verwende: {use_count}")
    if use_count > 0:
        print("   🖼️ Dateien:", ", ".join(matches[:use_count]))

    return [os.path.join(folder, f) for f in matches[:use_count]]

# ----------------------------------------------------------
# Excel-Verarbeitung
# ----------------------------------------------------------
def process_excel(excel_path, base_folder, row_start_1_based, row_end_1_based, max_objects=5):
    df = pd.read_excel(excel_path)

    # Spaltennamen normalisieren: trim + lowercase
    df.columns = [str(c).strip().lower() for c in df.columns]

    # Bereich 1-basiert → 0-basiert
    df = df.iloc[row_start_1_based - 1 : row_end_1_based]

    if "t1" not in df.columns:
        raise ValueError("❌ Erwartete Spalte 't1' fehlt (nach Normalisierung)!")

    grouped = df.groupby("t1")
    results = []

    print(f"\n📄 Verarbeite max. {max_objects} Objekte aus {excel_path} (Zeilen {row_start_1_based}-{row_end_1_based}) ...\n")

    for i, (object_id, group) in enumerate(grouped):
        if i >= max_objects:
            print("⏹️ Testlauf beendet nach", max_objects, "Objekten.")
            break

        try:
            object_id = str(object_id).strip() if object_id is not None else ""
            print(f"\n🔍 Bearbeite Objekt-ID: {object_id if object_id else '—'}")

            image_paths = find_images_for_object(base_folder, object_id, max_images=3)
            if not image_paths:
                print(f"⚠️ Keine gültigen Bilder für {object_id if object_id else '—'}")
                results.append({
                    "Quelle": os.path.basename(excel_path),
                    "Objekt-ID": object_id if object_id else "Leere Zelle",
                    "Bilder": "",
                    "Katalogtext": "❌ Kein gültiges Bild gefunden"
                })
                continue

            # Erste Zeile (mit normalisierten Spaltennamen)
            first_row = group.iloc[0]
            # Sicherstellen, dass auch first_row Index normalisiert ist:
            first_row.index = [str(c).strip().lower() for c in first_row.index]

            prompt = build_prompt(first_row, image_paths, object_id_value=object_id)

            print("\n📤 Prompt an API (gekürzt auf 1200 Zeichen):\n")
            snippet = prompt if len(prompt) <= 1200 else (prompt[:1200] + " ...[gekürzt]")
            print(snippet + "\n")

            catalog_text = generate_catalog_text(image_paths, prompt, max_tokens=1200)

            results.append({
                "Quelle": os.path.basename(excel_path),
                "Objekt-ID": object_id if object_id else "Leere Zelle",
                "Bilder": ", ".join([os.path.basename(p) for p in image_paths]),
                "Katalogtext": catalog_text
            })

            print(f"✅ {object_id if object_id else '—'}   → Text generiert")
            print("⏳ Warte 10 Sekunden, um Rate Limits zu vermeiden...\n")
            time.sleep(10)

        except Exception as e:
            print(f"❌ Fehler bei {object_id if object_id else '—'}: {e}")
            results.append({
                "Quelle": os.path.basename(excel_path),
                "Objekt-ID": object_id if object_id else "Leere Zelle",
                "Bilder": "",
                "Katalogtext": f"❌ Fehler: {str(e)}"
            })

    return results

# ----------------------------------------------------------
# Main
# ----------------------------------------------------------
if __name__ == "__main__":
    base_folder = "Objektbilder"

    all_results = []
    # A) Produktsammlung → Zeilen 3585–3650
    all_results.extend(
        process_excel("Liste_AEG Produktsammlung.xls", base_folder, 3585, 3650, max_objects=5)
    )
    # B) Kommunikation → Zeilen 308–314
    all_results.extend(
        process_excel("Liste_AK Kommunikation.xls", base_folder, 308, 314, max_objects=5)
    )

    out_path = "catalog_results_grouped.xlsx"
    pd.DataFrame(all_results).to_excel(out_path, index=False)
    print(f"\n📘 Ergebnisse gespeichert in: {out_path}")
