import streamlit as st
import pandas as pd
import json
from pathlib import Path
from PIL import Image

# Configure the page
st.set_page_config(page_title="AMP Material Passport", layout="wide")
st.title("🏗️ AMP Material Passport Extractor")
st.markdown("**Live Dashboard showing extracted BoQ data & embodied carbon calculations.**")

# Dynamically find the output directory
BASE_DIR = Path(__file__).resolve().parent.parent
OUTPUT_DIR = BASE_DIR / "output"

# 1. Display the Visualization
img_path = OUTPUT_DIR / "visualization.png"
if img_path.exists():
    st.subheader("📊 Material Distribution")
    image = Image.open(img_path)
    st.image(image, use_container_width=True)

# 2. Display the Extracted Excel Data
excel_path = OUTPUT_DIR / "passport_filled.xlsx"
if excel_path.exists():
    st.subheader("📋 Extracted Material Passport Data")
    # Load the data, skipping the top two formatting rows
    df = pd.read_excel(excel_path, header=2)
    
    # Drop completely empty columns for a cleaner view
    df_clean = df.dropna(axis=1, how='all')
    st.dataframe(df_clean, use_container_width=True)

# 3. Display the Building Metadata (Bonus B3)
meta_path = OUTPUT_DIR / "building_meta.json"
if meta_path.exists():
    st.subheader("🏢 Building Metadata")
    with open(meta_path, "r") as f:
        meta_data = json.load(f)
    st.json(meta_data)