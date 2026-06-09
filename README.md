# Alcorn MSG Quote Extractor

Streamlit app to upload Outlook `.msg` files, extract attached quote PDFs, rename PDFs, extract Alcorn quote line items, and download one ZIP containing:

- `alcorn_quote_extraction.xlsx`
- `pdf/alcorn_YYYYMMDD_HHMMSS_mmm_###.pdf`

## Privacy / storage behavior

The app processes uploaded MSG files and PDF attachments in memory. It does not use `st.cache_data` or `st.cache_resource`, and it does not permanently store PDFs.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Notes

The extractor uses visual PDF coordinates for both header blocks and item columns, then falls back to raw text parsing. Version 3 fixes these issues found in testing: bad Customer Number values from the logo address, mixed Sold-To/Ship-To addresses, lead-time notes appended to descriptions, and generic MISC/PARTS & MISC item IDs swallowing customer item numbers.
