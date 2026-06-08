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

The extractor uses visual PDF column positions first, then falls back to text parsing. This fixes the issue where PDFs were found but 0 rows were extracted because the PDF text separated `Qty.` and `Ord.` into different text lines.
