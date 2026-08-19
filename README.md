# AMP-GEN Material Passport Extractor

This repository contains the automated pipeline for extracting Material Passport data from scanned, hand-annotated Bills of Quantities (BoQ) using multimodal LLMs, adhering strictly to the AMP-GEN Material Passport schema.

## Features
- **Multimodal Extraction:** Uses Gemini models with model and API key rotation with an automated failover loop to parse complex, dot-matrix scanned documents.
- **Structured Schema Mapping:** Directly maps extracted quantities, items, disciplines, and material categories into the official `AMP_Passport_Template.xlsx` while preserving layout and style.
- **Embodied Carbon Calculations (Bonus B2):** Programmatically maps extracted materials to an EPD database to estimate lifecycle embodied carbon ($A1-A3$).
- **Building Metadata Extraction (Bonus B3):** Automatically isolates and outputs project metadata into `building_meta.json`.
- **Visualization:** Programmatically generates material distribution charts (`visualization.png`).

---

## Repository Structure
```text
amp-gen-passport/
├── data/
│   ├── AMP_Passport_Template.xlsx        # Official target template
│   └── BoQ_CBRI_Principals_Residence.pdf # Scanned BoQ source document
├── output/
│   ├── passport_filled.xlsx              # Fully populated passport template
│   ├── passport.json                     # Structured JSON export
│   ├── building_meta.json                # Extracted project metadata (Bonus B3)
│   └── visualization.png                 # Material category distribution chart
├── src/
│   ├── app.py                            # Streamlit application (Bonus B1)
│   └── extract_and_process.py            # Core extraction & processing pipeline
├── .env                                  # Local environment configuration (API keys)
├── .gitignore                          
├── APPROACH.md                           # Detailed methodology & lessons learned
├── README.md                             # Project documentation & run guide
└── requirements.txt                      # Project dependencies