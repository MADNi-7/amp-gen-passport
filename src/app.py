import streamlit as st
import pandas as pd
import json
import os
import tempfile
from pathlib import Path
from PIL import Image

# Configure the page to look modern and wide
st.set_page_config(page_title="AMP Material Passport", layout="wide", page_icon="🏗️")

# Attempt to load the extraction pipeline
try:
    from extract_and_process import extract_data_from_pdf, process_and_save, create_visualization, find_file
    pipeline_ready = True
except Exception as e:
    pipeline_ready = False

BASE_DIR = Path(__file__).resolve().parent.parent
OUTPUT_DIR = BASE_DIR / "output"

if 'processing_complete' not in st.session_state:
    st.session_state.processing_complete = False

# ==========================================
# SIDEBAR: UPLOAD & CONTROLS
# ==========================================
with st.sidebar:
    st.title("⚙️ Control Panel")
    st.markdown("Upload your scanned BoQ (PDF) to begin extraction.")
    st.divider()
    
    uploaded_file = st.file_uploader("📄 Choose a PDF file", type=["pdf"])
    
    run_pressed = False
    if uploaded_file is not None and pipeline_ready:
        run_pressed = st.button("🚀 Run AI Extraction", use_container_width=True, type="primary")
        
    if not pipeline_ready:
        st.error("⚠️ Pipeline offline. Missing API Key in Streamlit Secrets.")

# ==========================================
# MAIN HEADER
# ==========================================
st.title("🏗️ AMP Material Passport Dashboard")
st.markdown("Automated Bill of Quantities extraction and embodied carbon calculation using **Gemini AI**.")

# ==========================================
# PIPELINE EXECUTION
# ==========================================
if run_pressed:
    with st.spinner("🤖 Analyzing document with Gemini... (Takes 2-5 minutes. Please do not close)."):
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
                tmp_file.write(uploaded_file.getvalue())
                tmp_pdf_path = Path(tmp_file.name)
            
            template_file = find_file("AMP_Passport_Template.xlsx")
            extracted_data = extract_data_from_pdf(tmp_pdf_path)
            passport_df = process_and_save(extracted_data, template_file)
            create_visualization(passport_df)
            
            st.session_state.processing_complete = True
            st.sidebar.success("✅ Extraction Complete!")
        except Exception as e:
            st.error(f"Pipeline failed: {e}")

# ==========================================
# DYNAMIC RESULTS UI (TABS & METRICS)
# ==========================================
if st.session_state.processing_complete:
    st.divider()
    
    # Load the data defensively
    excel_path = OUTPUT_DIR / "passport_filled.xlsx"
    df_clean = pd.DataFrame()
    if excel_path.exists():
        df = pd.read_excel(excel_path, header=2)
        
        # UI Shield to guarantee no examples show up
        if 'GMAP Id' in df.columns:
            df = df[df['GMAP Id'] != 'EXAMPLE 1']
        if 'BOQ Item No.' in df.columns:
            df = df[df['BOQ Item No.'].astype(str) != '1.1.1']
            df = df[df['BOQ Item No.'].astype(str) != '2.1.1.1']
            
        df_clean = df.dropna(axis=1, how='all').dropna(axis=0, how='all')

    # Create 3 elegant tabs for better organization
    tab1, tab2, tab3 = st.tabs(["📊 Dashboard & Insights", "📋 Extracted Data", "🏢 Building Metadata"])
    
    with tab1:
        # High-level Metrics
        if not df_clean.empty:
            col1, col2, col3 = st.columns(3)
            col1.metric(label="Total Line Items Extracted", value=f"{len(df_clean)} items")
            col2.metric(label="Carbon Factors Applied", value="ICE v3.0")
            col3.metric(label="Data Schema", value="AMP Compliant")
            st.divider()
            
        img_path = OUTPUT_DIR / "visualization.png"
        if img_path.exists():
            st.subheader("Material Category Distribution")
            image = Image.open(img_path)
            # Center the image nicely using columns
            c1, c2, c3 = st.columns([1, 3, 1])
            with c2:
                st.image(image, use_container_width=True)

    with tab2:
        st.subheader("Raw Material Passport")
        st.markdown("This dataset maps directly to the official `AMP_Passport_Template.xlsx` schema.")
        if not df_clean.empty:
            st.dataframe(df_clean, use_container_width=True, height=600)
        else:
            st.info("No table data found.")
            
    with tab3:
        st.subheader("Project Metadata")
        meta_path = OUTPUT_DIR / "building_meta.json"
        if meta_path.exists():
            with open(meta_path, "r") as f:
                meta_data = json.load(f)
            st.json(meta_data)
else:
    # Welcome Screen
    st.info("👈 Please use the Control Panel on the left to upload a BoQ file and run the pipeline.")