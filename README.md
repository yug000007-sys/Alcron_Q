# Alcorn MSG Quote Extractor

Streamlit app to process Outlook `.msg` files containing Alcorn quote PDFs.

## What it does

- Upload one or more `.msg` files.
- Extract attached PDFs in memory only.
- Rename PDFs in the ZIP as `alcorn_YYYYMMDD_HHMMSS_mmm_###.pdf`.
- Extract quote line items into a fixed-header Excel file.
- Download one ZIP containing:
  - `alcorn_quote_extraction.xlsx`
  - renamed PDFs under `/pdf/`

## Important data rules

- Excel headers match the manual workbook layout.
- `item_id` uses the PDF **Item Number** column exactly, e.g. `MISC`, `PARTS & MISC`, `DYNA 56177`.
- Customer item numbers are used only for parsing descriptions, not exported as `item_id`.
- `List Price` is blank.
- `ContactEmail` uses only the first email in the **Ship To** block.
- Ship-To address is used for company/address/city/state/zip/country.
- ZIP+4 values are normalized to 5-digit ZIPs.
- Lead-time notes are ignored unless they are part of the manual description rule, such as `STK @ VENDOR`.
- `CustomerPONumber` is left blank to match the manual workbook.
- PDF column is written as `Alcorn_<QuoteNumber>.pdf`.

## Install

```bash
pip install -r requirements.txt
```

## Run

```bash
streamlit run app.py
```

## Privacy / storage

The app does not write uploaded MSG or extracted PDF files to disk. Processing is done in memory. The final ZIP is created in memory for download.
