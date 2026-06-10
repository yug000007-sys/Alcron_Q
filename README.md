# Alcorn MSG Quote Extractor

Streamlit app to upload Outlook `.msg` files, extract attached Alcorn quotation PDFs in memory, parse line-item data into the fixed Excel layout, and download one ZIP containing:

- `alcorn_quote_extraction.xlsx`
- renamed PDFs in `/pdf/`

## PDF naming rule

PDFs are renamed using:

`Alcorn_<CleanQuoteNumber>_<YYYYMMDD>.pdf`

Examples:

- `Alcorn_QT00042198_20260610_143522_381.pdf`
- `Alcorn_RQ8496_129_20260610_143522_382.pdf`
- `Alcorn_QT_ALCPT1698_20260610_143522_383.pdf`

Cleaning rule: all non-alphanumeric characters in the quote number are replaced with `_`, repeated underscores are collapsed, and leading/trailing underscores are removed. So `QT-ALCPT1698` becomes `QT_ALCPT1698`.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Privacy / storage

The app reads uploaded MSG/PDF content in memory and does not intentionally cache or persist PDFs. Output PDFs exist only inside the generated ZIP download.
