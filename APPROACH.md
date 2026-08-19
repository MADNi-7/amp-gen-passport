# APPROACH.md

## Tools Picked & Why
* **Google Gemini API (gemini-3.6-flash / 1.5-flash):** Standard OCR (like Tesseract) fails on dot-matrix, hand-annotated, and faded scans. Multimodal LLMs were chosen because they natively understand spatial document layouts, infer smudged characters contextually, and allow for "zero-shot" extraction directly into a structured JSON schema without brittle Regex.
* **Python (openpyxl, pandas):** `pandas` was used for data manipulation and carbon calculations. Crucially, `openpyxl` was selected to write the data back to Excel because it preserves the complex color-coding and formatting of the provided `AMP_Passport_Template.xlsx` (which `pandas.to_excel` would destroy).
* **Matplotlib & Seaborn:** Chosen for generating clean, programmatic visualizations of the material distribution.
* **Custom Failover Loop:** Built using `time` and rotating API keys to handle the heavy token load of parsing a 64-item BoQ, ensuring the script survives API timeouts and rate limits.

---

## What Worked
* **Direct-to-Schema Extraction:** Prompting the LLM to act as a Quantity Surveyor and outputting a strict JSON array bypassed the need for complex text-parsing logic. It accurately mapped messy descriptions to the required template columns.
* **Resilience & Dynamic Paths:** The failover retry loop completely solved `504 Deadline Exceeded` errors. Additionally, using `pathlib` for dynamic path resolution ensured the script successfully found the inputs/outputs regardless of where the VS Code terminal was launched from.
* **Bonuses Achieved:** Successfully extracted the building metadata (Bonus B3) and programmatically mapped inferred material categories to a mock EPD database for Embodied Carbon calculations (Bonus B2).

---

## What Did Not Work (Challenges)
* **Strict Mass Conversions:** To accurately calculate total embodied carbon, quantities need to be converted to mass (kg). Since the BoQ uses various units (Cum, Sqm, etc.), I had to hardcode assumed densities for major materials. This is not dynamically scalable without a comprehensive material database.
* **Absolute Confidence on Smudges:** Some hand-annotated fractions and dimensions are heavily degraded in the scan. While the LLM contextually guessed them well, absolute mathematical verification is impossible without human review of the original physical document.

---

## What I Would Do With 2 More Weeks
1. **Human-in-the-Loop (HITL) Interface:** Build a Streamlit web application that displays a side-by-side view of the PDF bounding boxes and the extracted JSON, allowing a QS to manually verify and correct smudged numbers before the final Excel export.
2. **Live EC3 API Integration:** Connect the script directly to the Embodied Carbon in Construction Calculator (EC3) API to fetch live, regional EPD data for carbon mapping instead of relying on a static mock dictionary.
3. **Advanced PyMuPDF Chunking:** To scale this pipeline for BoQs that are 100+ pages, I would implement PyMuPDF to crop the document row-by-row before passing it to the LLM, guaranteeing 100% precision on line-item counts without hitting context window token limits.