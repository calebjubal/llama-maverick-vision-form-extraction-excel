# Vision Form Extraction (Groq Maverick)

[![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?logo=python&logoColor=white)](https://www.python.org/) [![Streamlit](https://img.shields.io/badge/Streamlit-1.52+-FF4B4B?logo=streamlit&logoColor=white)](https://streamlit.io/) [![OpenPyXL](https://img.shields.io/badge/Excel-Automation-217346?logo=microsoft-excel&logoColor=white)](https://openpyxl.readthedocs.io/)

Vision-first Streamlit workspace that OCRs up to 20 workshop forms at once with Groq's Maverick vision model and auto-appends the structured fields into Excel.

**Tags:** `ocr` · `groq` · `streamlit` · `excel` · `vision` · `data-entry` · `automation`

---

## Table of Contents
1. [Overview](#overview)
2. [Features](#features)
3. [How It Works](#how-it-works)
4. [Getting Started](#getting-started)
5. [Configuration](#configuration)
6. [Usage Notes](#usage-notes)
7. [Troubleshooting](#troubleshooting)
8. [Project Layout](#project-layout)
9. [Contributing](#contributing)

## Overview
This tool accelerates repetitive claim processing by transforming scanned repair forms into clean, structured rows that can be reviewed and committed to Excel. The pipeline is optimized for Groq's `meta-llama/llama-4-maverick-17b-128e-instruct` vision model and enforces domain constraints (date formats, numeric VIN fragments, uppercase dealer observations) before persisting data.

## Features
- 🔍 **Vision Extraction:** Streams Groq Maverick responses for faster OCR + understanding.
- 🧱 **Schema Guardrails:** Only allows the predefined target fields to reduce post-cleaning.
- 📝 **Inline Review:** Streamlit `st.data_editor` lets reviewers tweak values before saving.
- 📊 **Excel Append:** Uses `openpyxl` to append validated rows into any uploaded workbook.
- ♻️ **Multi-Image Batch:** Handles up to 20 images per run with previews to cross-check results.

## How It Works
1. Users upload scanned form images plus a destination `.xlsx` file.
2. Each image is base64-encoded and sent to the Groq API with a strict JSON-only prompt (see [app.py](app.py)).
3. The streamed response is concatenated and parsed into Python dicts.
4. Reviewers edit rows in-app and preview the original image alongside the data.
5. On submission, sanitized rows append to the uploaded workbook via `openpyxl`.

## Getting Started
### Prerequisites
- Python 3.10+
- Groq account + API key
- Local Excel workbook to append data into

### Setup
```bash
# 1. Clone or download this repository
# 2. Create and activate a virtual environment (recommended)
python -m venv .venv
. .venv/Scripts/activate        # Windows PowerShell

# 3. Install dependencies
pip install -r requirements.txt
```

### Run the App
```bash
streamlit run app.py
```
Open the provided localhost URL in your browser to interact with the UI.

## Configuration
Set environment variables in a `.env` file at the project root (loaded via `python-dotenv`).

| Variable | Description |
| --- | --- |
| `GROQ_API_KEY` | Required. Groq API key with access to the Maverick vision model. |

Example `.env`:
```ini
GROQ_API_KEY="sk_your_groq_key"
```

## Usage Notes
- Supported image types: PNG, JPG, JPEG (max 20 files per run).
- The uploaded Excel file is modified in-place; keep backups if historical state matters.
- `Dealer Observation` is uppercased automatically to satisfy downstream validations.
- Date outputs follow `dd-mm-yy`. Adjust prompt rules in [app.py](app.py) if your region differs.

## Troubleshooting
- **Streaming errors:** Confirm your Groq quota and that `GROQ_API_KEY` is valid.
- **Excel write failures:** Ensure the workbook is not open in another program (Windows locks files).
- **JSON parsing issues:** Inspect model output via Streamlit logs; tighten prompt instructions if needed.
- **Performance:** Large dependency stack (PaddleOCR, etc.) comes from `requirements.txt`. Remove unused packages if startup time is critical.

## Project Layout
```
.
├── app.py          # Streamlit application and Groq integration
├── requirements.txt
└── README.md
```

## Contributing
We welcome improvements that enhance extraction accuracy, UX polish, or deployment flexibility.

### Ground Rules
- Keep features scoped and documented.
- Add inline comments only where logic is non-obvious.
- Avoid committing secrets (use `.env` and `.streamlit/secrets.toml`).

### Workflow
1. **Fork** or create a feature branch (`git checkout -b feat/better-review-table`).
2. **Install** dependencies via `pip install -r requirements.txt`.
3. **Format & Lint** (if you add tooling, document it here).
4. **Test manually**: `streamlit run app.py`, upload sample images + Excel, verify append.
5. **Commit** with descriptive messages and open a PR describing motivation + screenshots.

### Issue Templates
When filing issues, please include:
- Repro steps (images + Excel sample if possible)
- Expected vs actual behavior
- Logs or tracebacks from the Streamlit console

Need ideas to tackle? See `README` TODOs (if any) or open a discussion before large refactors.
