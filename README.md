# Beacon Regression DB Generator

Web app to generate regression test files for Beacon Lending System.

## Live App
https://beacon-regression.streamlit.app

## Features
- Company, Lender, Account generation
- Sanction Master with main + sub breakups
- 10 liability types (TL, WCDL, CC, STL, OD, LOC, ICD, ECB, Re-Finance, DA)
- TDS, Fees, Prepayment rules
- Download as individual files or ZIP

## Run Locally
```bash
pip install -r requirements.txt
streamlit run web_app.py