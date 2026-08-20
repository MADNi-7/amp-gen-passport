import streamlit as st
import pandas as pd
import json
import os
import tempfile
from pathlib import Path
from PIL import Image

# Configure the page
st.set_page_config(page_title="AMP Material Passport", layout="wide")
st.title("🏗️ AMP Material Passport Extractor")
st.markdown("**Upload a scanned BoQ (PDF) to extract data and calculate embodied carbon.**")

# Attempt to load the extraction pipeline
try:
    from extract_and_process import extract_data_from_pdf, process_and_save, create_visualization, find_file
    pipeline_ready = True
except Exception as e:
    pipeline_ready = False
    st.error(f"⚠️ Pipeline offline. Please add GEMINI_API_KEY to Streamlit Secrets. (Error: {e})")

BASE_DIR = Path(__file__).resolve().parent.parent
OUTPUT_DIR = BASE_DIR / "output"

if 'processing_complete' not in st.session_state:
    st.session_state.processing_complete = False

# ==========================================
# 1. UPLOAD & RUN PIPELINE
# ==========================================
with st.expander("📤 Upload new BoQ (PDF)", expanded=True):
    uploaded_file = st.file_uploader("Choose a PDF file", type=["pdf"])
    
    if uploaded_file is not None and pipeline_ready:
        if st.button("🚀 Run AI Extraction Pipeline"):
            with st.spinner("🤖 Analyzing document with Gemini... (This takes 2-5 minutes. Please do not close the page)."):
                try:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
                        tmp_file.write(uploaded_file.getvalue())
                        tmp_pdf_path = Path(tmp_file.name)
                    
                    template_file = find_file("AMP_Passport_Template.xlsx")
                    extracted_data = extract_data_from_pdf(tmp_pdf_path)
                    passport_df = process_and_save(extracted_data, template_file)
                    create_visualization(passport_df)
                    
                    st.session_state.processing_complete = True
                    st.success("✅ Extraction Complete! Results updated below.")
                except Exception as e:
                    st.error(f"Pipeline failed: {e}")

st.divider()

# ==========================================
# 2. DISPLAY DYNAMIC RESULTS
# ==========================================
# Display Visualization
img_path = OUTPUT_DIR / "visualization.png"
if img_path.exists():
    st.subheader("📊 Material Distribution")
    image = Image.open(img_path)
    st.image(image, use_container_width=True)

# Display Excel Data
excel_path = OUTPUT_DIR / "passport_filled.xlsx"
if excel_path.exists():
    st.subheader("📋 Extracted Material Passport Data")
    df = pd.read_excel(excel_path, header=2)
    
    # 🛡️ THE SHIELD: Forcefully remove the template examples from the pandas dataframe
    if 'GMAP Id' in df.columns:
        df = df[df['GMAP Id'] != 'EXAMPLE 1']
    if 'BOQ Item No.' in df.columns:
        df = df[df['BOQ Item No.'].astype(str) != '1.1.1']
        df = df[df['BOQ Item No.'].astype(str) != '2.1.1.1']
    
    # Clean the dataframe
    df_clean = df.dropna(axis=1, how='all').dropna(axis=0, how='all')
    st.dataframe(df_clean, use_container_width=True)

# Display Metadata
meta_path = OUTPUT_DIR / "building_meta.json"
if meta_path.exists():
    st.subheader("🏢 Building Metadata")
    with open(meta_path, "r") as f:
        meta_data = json.load(f)
    st.json(meta_data)