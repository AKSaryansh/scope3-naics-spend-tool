# Scope 3 (Category 1 & 2) Spend-Based Emissions Tool (NAICS)

This tool calculates Scope 3 emissions using **spend-based emission factors** mapped to **NAICS** codes.

## What it does
- Upload a purchases Excel.
- Infer the spend column and description/category columns.
- Map each line to a NAICS code (using the provided reference workbook tables first; otherwise best-match suggestions).
- Convert INR to USD using an **INR per USD** value you enter.
- Uses **"Supply Chain Emission Factors with Margins"** (conservative) from the reference workbook.
- Outputs Cat 1 and Cat 2 tables and exports an Excel with **formulas in cells**.

## Run locally
From this folder:

```bash
python3.12 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

## Reference workbook
The app needs a reference workbook containing:
- `EF & Conversion` (NAICS + spend-based EF with margins)
- (optional) `Calculation Sheet` (Category/Product → NAICS mapping)

### Local dev default
If `../product wise sales and procurement details.xlsx` exists, the app will use it automatically.

### Deployed environments
For deployments, you typically **upload the reference workbook in the sidebar**, or mount it and set:
- `SCOPE3_REF_XLSX_PATH=/path/to/reference.xlsx`
