# AMP-GEN Material Passport Extractor

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://amp-gen-passport-uhr9sy2rxzm8jfbmf5p4yt.streamlit.app/)

This repository contains an automated, AI-driven pipeline for extracting Material Passport data from scanned, hand-annotated Bills of Quantities (BoQ) documents. The data is mapped strictly to the AMP-GEN Material Passport schema and visualized via a modern web dashboard.

## 🔗 Quick Links
- **[🔴 LIVE DEMO: Streamlit Dashboard](https://amp-gen-passport-uhr9sy2rxzm8jfbmf5p4yt.streamlit.app/)**

---

## 🚀 Features & Bonuses Completed

- **Mandatory Task:** Extracted 64 line items from a dot-matrix scanned BoQ (PDF) and mapped them directly into the 37 mandatory green columns of the `AMP_Passport_Template.xlsx`.
- **[Bonus B1] Live Web Deployment:** Built a fully interactive UI using Streamlit. Features a sidebar control panel, tabbed data views, top-level metrics, and live PDF uploading so anyone can test the AI pipeline.
- **[Bonus B2] Embodied Carbon Calculations:** Programmatically mapped extracted materials to an EPD database (ICE v3.0) to estimate lifecycle embodied carbon ($A1-A3$) and populated the AMBER columns.
- **[Bonus B3] Building Metadata:** Isolated and extracted high-level project details (Project Name, Location, etc.) from Page 1 and saved them as `building_meta.json`.

---

## 🛠️ Repository Structure
```text
amp-gen-passport/
├── data/
│   ├── AMP_Passport_Template.xlsx        # Official target template
│   └── BoQ_CBRI_Principals_Residence.pdf # Scanned BoQ source document
├── output/
│   ├── passport_filled.xlsx              # Fully populated passport template
│   ├── passport.json                     # Structured JSON export
│   ├── building_meta.json                # Extracted project metadata (B3)
│   └── visualization.png                 # Material category distribution chart
├── src/
│   ├── app.py                            # Streamlit web application dashboard
│   └── extract_and_process.py            # Core Gemini extraction & excel logic
├── APPROACH.md                           # Detailed methodology & problem-solving
├── README.md                             # Project documentation
└── requirements.txt                      # Project dependencies
