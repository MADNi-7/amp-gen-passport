import os
import json
import time
from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import google.generativeai as genai
from dotenv import load_dotenv
import openpyxl

# ============================================================================
# 1. DYNAMIC PATH RESOLUTION (Bulletproof path finding)
# ============================================================================
SRC_DIR = Path(__file__).resolve().parent          # .../amp-gen-passport/src
PROJECT_DIR = SRC_DIR.parent                      # .../amp-gen-passport
WORKSPACE_DIR = PROJECT_DIR.parent                # .../amp-gen-task

# Load .env file from either project dir or workspace dir
load_dotenv(PROJECT_DIR / ".env")
load_dotenv(WORKSPACE_DIR / ".env")

OUTPUT_DIR = PROJECT_DIR / "output"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

def find_file(filename):
    """Searches for a file across probable directory locations."""
    search_paths = [
        WORKSPACE_DIR / "data" / filename,
        PROJECT_DIR / "data" / filename,
        PROJECT_DIR / filename,
        WORKSPACE_DIR / filename
    ]
    for path in search_paths:
        if path.exists():
            return path
    raise FileNotFoundError(f"Could not find '{filename}'. Searched in: {[str(p) for p in search_paths]}")

# ============================================================================
# 2. CONFIGURATION & KEYS
# ============================================================================
API_KEYS = [
    os.getenv("GEMINI_API_KEY"),
    os.getenv("GEMINI_API_KEY_2"),
    os.getenv("GEMINI_API_KEY_3")
]
# Keep valid non-empty keys
API_KEYS = [k for k in API_KEYS if k]

if not API_KEYS:
    raise ValueError("No Gemini API keys found! Please add GEMINI_API_KEY to your .env file.")

# Rotating Models
MODELS = [
    "gemini-3.6-flash",
    "gemini-3.7-flash",
    "gemini-3.5-flash",
    "gemini-3.1-pro-preview",
    "gemini-omni-flash-preview"
]

# Bonus B2: Mock EPD Database (Embodied Carbon - kgCO2e/kg)
EPD_DATABASE = {
    "Concrete": {"carbon_factor": 0.15, "source": "ICE Database V3.0"},
    "Steel": {"carbon_factor": 1.46, "source": "ICE Database V3.0"},
    "Brick": {"carbon_factor": 0.24, "source": "ICE Database V3.0"},
    "Timber": {"carbon_factor": 0.30, "source": "ICE Database V3.0"},
    "Glass": {"carbon_factor": 1.43, "source": "ICE Database V3.0"}
}

# ============================================================================
# 3. EXTRACTION FUNCTION (With Failover & Rotation)
# ============================================================================
def extract_data_from_pdf(pdf_path):
    """Extracts structured JSON from PDF using model & key rotation with backoff."""
    prompt = """
    You are an expert Quantity Surveyor. Analyze this scanned Bill of Quantities (BoQ) and extract:
    1. 'metadata': The building block metadata from Page 1 (Project name, location, date, etc.)
    2. 'line_items': An array of exactly 64 line items. For each item, extract:
       - 'id': Item number
       - 'description': Full description of the work/material
       - 'unit': Unit of measurement (e.g., Cum, Sqm, Kg)
       - 'quantity': The numerical quantity
       - 'discipline': Infer if this is 'Civil', 'MEP', 'Architecture', etc.
       - 'material_category': Infer the primary material (e.g., 'Concrete', 'Steel', 'Brick')
       
    Output STRICTLY in valid JSON format without markdown blocks.
    """
    
    key_idx = 0
    model_idx = 0
    
    while True:
        curr_key = API_KEYS[key_idx]
        curr_model = MODELS[model_idx]
        
        genai.configure(api_key=curr_key)
        print(f"\n[Attempt] Model: {curr_model} | Key: {key_idx + 1}/{len(API_KEYS)}")
        
        try:
            print(f"  -> Uploading PDF ({pdf_path.name}) to Gemini...")
            sample_file = genai.upload_file(path=str(pdf_path))
            
            model = genai.GenerativeModel(curr_model)
            
            print("  -> Extracting line items (timeout set to 600s)...")
            response = model.generate_content(
                [sample_file, prompt],
                request_options={"timeout": 600}
            )
            
            raw_json = response.text.strip().removeprefix('```json').removesuffix('```')
            data = json.loads(raw_json)
            
            print("✓ Extraction successful!")
            return data

        except Exception as e:
            print(f"✗ Extraction failed: {type(e).__name__} - {e}")
            
            key_idx = (key_idx + 1) % len(API_KEYS)
            model_idx = (model_idx + 1) % len(MODELS)
            
            print("  -> Holding for 15 seconds to avoid rate limits before resuming...\n")
            time.sleep(15)

# ============================================================================
# 4. PROCESS AND FILL TEMPLATE
# ============================================================================
def process_and_save(data, template_path):
    """Processes JSON data, fills the template with green/amber columns, and exports."""
    # Bonus B3: Save Metadata
    meta_path = OUTPUT_DIR / "building_meta.json"
    with open(meta_path, "w", encoding="utf-8") as f:
        json.dump(data.get("metadata", {}), f, indent=4)
        
    line_items = data.get("line_items", [])
    print(f"Filling official AMP Passport Template with {len(line_items)} items...")
    
    wb = openpyxl.load_workbook(str(template_path))
    ws = wb.active
    start_row = 6
    
    for idx, item in enumerate(line_items):
        current_row = start_row + idx
        
        ws.cell(row=current_row, column=2).value = item.get("id")
        ws.cell(row=current_row, column=5).value = item.get("description")
        ws.cell(row=current_row, column=7).value = item.get("discipline")
        ws.cell(row=current_row, column=10).value = item.get("material_category")
        
        try:
            qty = float(item.get("quantity", 0))
            ws.cell(row=current_row, column=14).value = qty
        except (ValueError, TypeError):
            ws.cell(row=current_row, column=14).value = item.get("quantity")
            qty = 0
            
        ws.cell(row=current_row, column=15).value = item.get("unit")
        
        # Carbon calculation (AMBER columns)
        category = item.get("material_category", "")
        if category in EPD_DATABASE:
            carbon_factor = EPD_DATABASE[category]["carbon_factor"]
            source = EPD_DATABASE[category]["source"]
            
            ws.cell(row=current_row, column=25).value = qty * carbon_factor
            ws.cell(row=current_row, column=26).value = carbon_factor
            ws.cell(row=current_row, column=50).value = f"EPD Source: {source}"
            
    excel_out = OUTPUT_DIR / "passport_filled.xlsx"
    wb.save(str(excel_out))
    
    json_out = OUTPUT_DIR / "passport.json"
    with open(json_out, "w", encoding="utf-8") as f:
        json.dump(line_items, f, indent=4)
        
    print(f"✓ Saved: {excel_out}")
    print(f"✓ Saved: {json_out}")
    return pd.DataFrame(line_items)

# ============================================================================
# 5. VISUALIZATION
# ============================================================================
def create_visualization(df):
    """Generates building-level material distribution chart."""
    if df.empty or "material_category" not in df.columns:
        print("⚠ Insufficient data for visualization.")
        return
        
    plt.figure(figsize=(10, 6))
    mat_counts = df["material_category"].value_counts()
    
    sns.barplot(x=mat_counts.values, y=mat_counts.index, palette="viridis")
    plt.title("Material Distribution across BoQ Line Items", fontsize=14)
    plt.xlabel("Number of Line Items", fontsize=12)
    plt.ylabel("Material Category", fontsize=12)
    plt.tight_layout()
    
    img_out = OUTPUT_DIR / "visualization.png"
    plt.savefig(str(img_out))
    plt.close()
    print(f"✓ Visualization saved: {img_out}")

# ============================================================================
# 6. MAIN RUNNER
# ============================================================================
if __name__ == "__main__":
    pdf_file = find_file("BoQ_CBRI_Principals_Residence.pdf")
    template_file = find_file("AMP_Passport_Template.xlsx")
    
    print(f"Using PDF: {pdf_file}")
    print(f"Using Template: {template_file}")
    
    extracted_data = extract_data_from_pdf(pdf_file)
    passport_df = process_and_save(extracted_data, template_file)
    create_visualization(passport_df)
    print("\n All tasks completed successfully!")