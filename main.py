import os
import base64
import pandas as pd
import requests
import time
from dotenv import load_dotenv
from tqdm import tqdm

# ----------------------------------------------------------
# Configuration constants
# ----------------------------------------------------------
MAX_IMAGES = 3
MAX_TOKENS = 1200
API_WAIT = 10  # seconds between API calls to avoid rate limits

# ----------------------------------------------------------
# Helper functions
# ----------------------------------------------------------
def encode_image(image_path):
    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode("utf-8")

def safe_value(val):
    if val is None:
        return "Empty Cell"
    try:
        s = str(val).strip()
    except Exception:
        return "Empty Cell"
    return "Empty Cell" if s == "" or s.lower() == "nan" else s

def get_val(row, colname_lower):
    if colname_lower in row.index:
        return safe_value(row[colname_lower])
    return "Empty Cell"

# ----------------------------------------------------------
# Setup
# ----------------------------------------------------------
load_dotenv()
API_KEY = os.getenv("OPENAI_API_KEY")
if not API_KEY:
    raise ValueError("Error: No OPENAI_API_KEY found! Please check your .env file.")

OPENAI_URL = "https://api.openai.com/v1/chat/completions"

# ----------------------------------------------------------
# Prompt template
# ----------------------------------------------------------
BASE_PROMPT_TEMPLATE = """
You are a museum cataloguer for the Technisches Museum Berlin.
Describe only what the images justify; be conservative; no speculation.
The input includes metadata from an Excel row.
If any field value equals the literal string "Empty Cell", treat it as missing/absent. 

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
Never invent missing values. If a field is absent or equals "Empty Cell", omit it entirely from that list sentence. 
Do not state date, maker, materials, or dimensions as if they are visibly inscribed. For visible inscriptions, transcribe verbatim in quotes; use […] for illegible parts. 
Ban hedging/guessing: avoid vermutlich, wahrscheinlich, wohl, offenbar, anscheinend, möglicherweise, könnte, dürfte, early/late, typical, domestic, etc. 

OUTPUT (exactly three parts, in this order) 
One sentence (≤35 Wörter) suitable for a label: Objekttyp, Funktion/Zweck, sichtbares Material/Finish, datierungslos. Do not include maker/date/provenance here unless they are visibly inscribed; Excel facts remain only in the “Angaben laut Liste” sentence within the description.

Write 3 short paragraphs in German, then provide an English translation with the same constraints. The German paragraphs should have one title above the first paragraph called "Deutsche Beschreibung". Then provide the EN translation of the title mirroring the placement used for the German title.
Content of the paragraph 1: Name the object type in plain words and its function/purpose (use title if it matches the images). 
Content of the paragraph 2: Strictly image-based description: form and construction (shape, handles, openings, moving parts, connectors); colours/finish; materials only if visually evident; transcribe any visible labels/marks verbatim (use […] for gaps). At the end of Paragraph 2, add ONE sentence with Excel facts using EXACT formatting, but include only the fields that exist (omit any that are "Empty Cell"): 
Angaben laut Liste: {de_list_sentence}. 
Content of the paragraph 3: Condition notes only (e.g., Kratzer, Korrosion, Abplatzungen, fehlende Teile). No usage scenarios or history. Then provide the EN translation mirroring the structure above. In the English version of Paragraph 2, reproduce the list sentence as: According to the list: {en_list_sentence}.
"""

# ----------------------------------------------------------
# Prompt building
# ----------------------------------------------------------
def build_list_sentences(manufacturer, material, dimensions, year):
    parts_de, parts_en = [], []
    if manufacturer != "Empty Cell":
        parts_de.append(f"Hersteller: {manufacturer}")
        parts_en.append(f"Manufacturer: {manufacturer}")
    if material != "Empty Cell":
        parts_de.append(f"Material: {material}")
        parts_en.append(f"Material: {material}")
    if dimensions != "Empty Cell":
        parts_de.append(f"Maße: {dimensions}")
        parts_en.append(f"Dimensions: {dimensions}")
    if year != "Empty Cell":
        parts_de.append(f"Jahr: {year}")
        parts_en.append(f"Year: {year}")
    return "; ".join(parts_de), "; ".join(parts_en)

def build_prompt(row, image_paths, object_id):
    object_id = str(object_id).strip() if object_id is not None else "Empty Cell"
    if object_id == "":
        object_id = "Empty Cell"

    title        = get_val(row, "t10")
    manufacturer = get_val(row, "t2")
    material     = get_val(row, "t3")
    dimensions   = get_val(row, "t5")
    year         = get_val(row, "t14")

    image_list  = ", ".join([os.path.basename(p) for p in image_paths]) if image_paths else "Empty Cell"

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
def generate_catalog_text(image_paths, prompt_text, max_tokens=MAX_TOKENS):
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
                time.sleep(wait)
                continue
            return f"Error: API error: {str(e)}"
        except requests.exceptions.RequestException as e:
            return f"Error: Connection error: {str(e)}"
    return "Error: Aborted after too many failed attempts"

# ----------------------------------------------------------
# Image search
# ----------------------------------------------------------
def find_images_for_object(base_folder, object_id, max_images=MAX_IMAGES):
    year = next((p for p in object_id.split("/") if p.isdigit() and len(p) == 4), None)
    if not year:
        return []

    folder = os.path.join(base_folder, year)
    if not os.path.exists(folder):
        return []

    prefix = object_id.replace("/", "-").split()[0]
    candidates = [f for f in os.listdir(folder) if f.lower().endswith((".jpg", ".jpeg", ".png"))]
    matches = sorted([f for f in candidates if f.startswith(prefix)])

    return [os.path.join(folder, f) for f in matches[:max_images]]

# ----------------------------------------------------------
# Excel processing
# ----------------------------------------------------------
def process_excel(excel_path, base_folder, row_start_1_based, row_end_1_based, max_objects=5):
    df = pd.read_excel(excel_path)
    df.columns = [str(c).strip().lower() for c in df.columns]
    df = df.iloc[row_start_1_based - 1 : row_end_1_based]

    if "t1" not in df.columns:
        raise ValueError("Error: Expected column 't1' not found (after normalization)!")

    grouped = df.groupby("t1")
    results = []

    with tqdm(total=min(max_objects, len(grouped)), desc=f"Processing {os.path.basename(excel_path)}", unit="Obj") as pbar:
        for i, (object_id, group) in enumerate(grouped):
            if i >= max_objects:
                break

            try:
                object_id = str(object_id).strip() if object_id is not None else "Empty Cell"
                image_paths = find_images_for_object(base_folder, object_id, max_images=MAX_IMAGES)

                if not image_paths:
                    results.append({
                        "Source": os.path.basename(excel_path),
                        "Object ID": object_id,
                        "Images": "",
                        "Description": "Error: No valid image found"
                    })
                    pbar.update(1)
                    continue

                first_row = group.iloc[0]
                first_row.index = [str(c).strip().lower() for c in first_row.index]

                prompt = build_prompt(first_row, image_paths, object_id)
                catalog_text = generate_catalog_text(image_paths, prompt, max_tokens=MAX_TOKENS)

                results.append({
                    "Source": os.path.basename(excel_path),
                    "Object ID": object_id,
                    "Images": ", ".join([os.path.basename(p) for p in image_paths]),
                    "Description": catalog_text
                })

                time.sleep(API_WAIT)

            except Exception as e:
                results.append({
                    "Source": os.path.basename(excel_path),
                    "Object ID": object_id,
                    "Images": "",
                    "Description": f"Error: Processing error: {str(e)}"
                })

            pbar.update(1)

    return results

# ----------------------------------------------------------
# Main
# ----------------------------------------------------------
if __name__ == "__main__":
    base_folder = "Objektbilder"

    all_results = []
    all_results.extend(process_excel("Liste_AEG Produktsammlung.xls", base_folder, 3585, 3650, max_objects=5))
    all_results.extend(process_excel("Liste_AK Kommunikation.xls", base_folder, 308, 314, max_objects=5))

    out_path = "object_descriptions.xlsx"
    pd.DataFrame(all_results).to_excel(out_path, index=False)
    print(f"\nResults saved in: {out_path}")
