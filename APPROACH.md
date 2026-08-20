# Methodology and Approach

## 1. Tools Picked & Why
* **Google Gemini API (gemini-3.6-flash / 3.7-flash):** Standard OCR (like Tesseract) fails on dot-matrix, hand-annotated, and faded scans. Multimodal LLMs were chosen because they natively understand spatial document layouts, infer smudged characters contextually, and allow for "zero-shot" extraction directly into a structured JSON schema without brittle Regex.
* **Python (openpyxl & pandas):** `pandas` was used for data manipulation and carbon calculations. Crucially, `openpyxl` was selected to write the data back to Excel because it preserves the complex color-coding and formatting of the provided `AMP_Passport_Template.xlsx` (which `pandas.to_excel` would destroy).
* **Streamlit (Bonus B1):** Chosen for its ability to rapidly deploy data scripts into interactive web applications, allowing end-users to upload PDFs and view results dynamically without touching the terminal.
* **Matplotlib & Seaborn:** Chosen for generating clean, programmatic visualizations of the material distribution.

---

## 2. Technical Strategy & Problem Solving
* **API Resiliency & Failovers:** Extracting 64 detailed line items from a multi-page PDF requires a massive output token context, which occasionally leads to API timeouts or rate limits. I implemented a robust `while True` failover loop that automatically rotates through backup API keys and models if a timeout exception occurs.
* **Data Integrity (The List-to-String Bug):** During testing, the AI occasionally returned an array for the material description (e.g., `["soil"]`) instead of a raw string. When `openpyxl` attempted to write this list directly into an Excel cell, it threw a `ValueError`. I implemented a programmatic type checker that intercepts and flattens the AI output before writing.
* **Template Sanitization:** The provided template contained three sample rows. Simply overwriting the target cells left behind "ghost" data in the unmapped grey columns, which corrupted visualizations. I utilized `openpyxl`'s `ws.delete_rows(4, 3)` to physically scrub the template examples, and added a secondary "Pandas Shield" into the Streamlit app to ensure absolute data purity on the web dashboard.

---

## 3. What Worked (Successes & Bonuses)
* **Direct-to-Schema Extraction:** Prompting the LLM to act as a Quantity Surveyor and outputting a strict JSON array bypassed the need for complex text-parsing logic, mapping messy descriptions directly to the required template columns.
* **Live Web Dashboard (Bonus B1):** Successfully deployed a live UI with tabbed data views and a side-bar upload panel.
* **Embodied Carbon Calculations (Bonus B2):** Programmatically mapped inferred material categories to a mock EPD database (ICE v3.0) for $A1-A3$ calculations.
* **Building Metadata Extraction (Bonus B3):** Successfully extracted the building metadata (Project Name, Location, etc.) directly from Page 1 into `building_meta.json`.

---

## 4. What Did Not Work (Challenges)
* **Strict Mass Conversions:** To accurately calculate total embodied carbon, quantities need to be converted to mass (kg). Since the BoQ uses various units (Cum, Sqm, etc.), I had to hardcode assumed densities for major materials. This is not dynamically scalable without a comprehensive material database.
* **Absolute Confidence on Smudges:** Some hand-annotated fractions and dimensions are heavily degraded in the scan. While the LLM contextually guessed them well, absolute mathematical verification is impossible without human review of the original physical document.

---

## 5. What I Would Do With 2 More Weeks
1. **Human-in-the-Loop (HITL) Interface:** Expand the current Streamlit web application to display a side-by-side view of the PDF bounding boxes and the extracted JSON, allowing a QS to manually verify and correct smudged numbers before the final Excel export.
2. **Live EC3 API Integration:** Connect the script directly to the Embodied Carbon in Construction Calculator (EC3) API to fetch live, regional EPD data for carbon mapping instead of relying on a static mock dictionary.
3. **Advanced PyMuPDF Chunking:** To scale this pipeline for BoQs that are 100+ pages, I would implement PyMuPDF to crop the document row-by-row before passing it to the LLM, guaranteeing 100% precision on line-item counts without hitting context window token limits.
