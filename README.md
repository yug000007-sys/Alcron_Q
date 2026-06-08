# Alcorn MSG Quote Extractor

Streamlit app to extract PDF quote attachments from Outlook `.msg` files, parse Alcorn quote data, and download one ZIP containing:

- `alcorn_quote_extraction.xlsx`
- renamed PDFs in `/pdf/`

## Key rules

- Uploaded `.msg` files are processed only during the Streamlit run.
- PDFs are not saved permanently and no Streamlit cache is used.
- PDFs inside the ZIP are renamed like:

```text
alcorn_YYYYMMDD_HHMMSS_millisecond_001.pdf
```

- Excel headers are fixed in `FIXED_HEADERS` in `app.py`.
- Generic item IDs like `MISC` and `PARTS & MISC` use the real customer item number when available.
- One Excel row is created per actual PDF line item.
- Notes such as lead times, `IN STOCK`, repair percentage, and vendor notes are not treated as separate rows.

## Local setup

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

## GitHub deployment

Upload these files to a GitHub repository:

```text
app.py
requirements.txt
README.md
.gitignore
```

Then deploy the repository on Streamlit Community Cloud.

## Output

Click **Extract quotes and build ZIP** after uploading `.msg` files. The download ZIP contains the Excel file and all renamed PDF attachments.
