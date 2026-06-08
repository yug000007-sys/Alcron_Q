import io
import re
import zipfile
from datetime import datetime
from decimal import Decimal, InvalidOperation
from typing import Dict, List, Tuple

import pandas as pd
import pdfplumber
import streamlit as st

try:
    import extract_msg
except ImportError:  # friendly app error if requirements were not installed
    extract_msg = None

APP_TITLE = "Alcorn MSG Quote Extractor"
PDF_PREFIX = "alcorn"

FIXED_HEADERS = [
    "ReferralManagerCode",
    "ReferralManagerName",
    "ReferralEmail",
    "Brand",
    "QuoteNumber",
    "QuoteVersion",
    "QuoteDate",
    "QuoteValidUntil",
    "Customer Number/ID",
    "Company",
    "Address",
    "Country",
    "City",
    "State",
    "ZipCode",
    "CountryName",
    "FirstName",
    "LastName",
    "ContactEmail",
    "ContactPhone",
    "Webaddress",
    "item_id",
    "item_desc",
    "UOM",
    "Quantity",
    "Unit Price",
    "List Price",
    "TotalSales",
    "Manufacturer",
    "manufacturer_part_number",
    "Writer Name",
    "CustomerPONumber",
    "PDF",
    "DemoQuote",
    "Duns",
    "SIC",
    "NAICS",
    "LineOfBusiness",
    "LinkedInProfile",
    "PhoneResearch",
    "PhoneSuppression",
    "ParentName",
]

EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")
MONEY_RE = re.compile(r"^-?\$?\d{1,3}(?:,\d{3})*(?:\.\d{2})$|^-?\$?\d+(?:\.\d{2})$")
ITEM_START_RE = re.compile(r"^\s*(\d+)\s+(.+?)\s+(-?\$?\d[\d,]*\.\d{2})\s+(-?\$?\d[\d,]*\.\d{2})\s*$")

STOP_PHRASES = (
    "tax summary", "subtotal", "total sales tax", "total order", "comments:",
    "order discount", "included tax", "less", "quotation continued", "please send your order"
)
NOTE_PATTERNS = [
    re.compile(r"^est\.?\s+\d", re.I),
    re.compile(r"^eta\s+", re.I),
    re.compile(r"^\d+\s*-\s*\d+\s*(day|days|week|weeks)", re.I),
    re.compile(r"^\d+\s*(day|days|week|weeks)\b", re.I),
    re.compile(r"^in stock", re.I),
    re.compile(r"^stk\s+@", re.I),
    re.compile(r"^quote is for", re.I),
    re.compile(r"^new tool\s+\$", re.I),
    re.compile(r"^repair is\s+", re.I),
    re.compile(r"^cert from", re.I),
]


def clean_text(value: str) -> str:
    value = value or ""
    value = value.replace("\u201c", '"').replace("\u201d", '"').replace("\u2019", "'")
    value = re.sub(r"\s+", " ", value).strip()
    return value


def money_to_float(value: str):
    if not value:
        return None
    try:
        return float(Decimal(value.replace("$", "").replace(",", "")))
    except (InvalidOperation, ValueError):
        return None


def should_ignore_note(line: str) -> bool:
    line = clean_text(line)
    return any(p.search(line) for p in NOTE_PATTERNS)


def split_city_state_zip(line: str) -> Tuple[str, str, str, str]:
    line = clean_text(line).replace(" ,", ",")
    country = ""
    if line.upper().endswith(" USA"):
        country = "USA"
        line = line[:-4].strip()
    m = re.match(r"^(.*?),\s*([A-Za-z]{2}|[A-Za-z]+),?\s+([A-Za-z0-9][A-Za-z0-9\- ]*)$", line)
    if m:
        return clean_text(m.group(1)), clean_text(m.group(2)), clean_text(m.group(3)), country
    m = re.match(r"^(.*?),\s*([A-Z]{2})\s+([0-9]{5}(?:-[0-9]{4})?)$", line)
    if m:
        return clean_text(m.group(1)), clean_text(m.group(2)), clean_text(m.group(3)), country
    return "", "", "", country


def parse_address_block(block_lines: List[str]) -> Dict[str, str]:
    cleaned = [clean_text(x) for x in block_lines if clean_text(x)]
    emails = []
    non_email = []
    for line in cleaned:
        emails.extend(EMAIL_RE.findall(line))
        stripped = EMAIL_RE.sub("", line)
        stripped = re.sub(r"(?i)email[:\-]?|<-.*|<--.*|\(email only\)|\(email\)|inv thru coupa|invoice[s]?.*", "", stripped)
        stripped = clean_text(stripped)
        if stripped:
            non_email.append(stripped)

    company = non_email[0] if non_email else ""
    first_name = ""
    address_parts = []
    city = state = zip_code = country = ""

    for line in non_email[1:]:
        if line.upper() in {"USA", "HOIST"}:
            country = "USA" if line.upper() == "USA" else country
            continue
        if line.lower().startswith("attn"):
            contact = re.sub(r"(?i)^attn[:\s]*", "", line).strip()
            first_name = contact
            continue
        c, s, z, co = split_city_state_zip(line)
        if c or s or z:
            city, state, zip_code = c, s, z
            if co:
                country = co
            continue
        if "customers only" in line.lower() or line.lower() == "counter sales":
            continue
        address_parts.append(line)

    return {
        "Company": company,
        "Address": ", ".join(address_parts),
        "City": city,
        "State": state,
        "ZipCode": zip_code,
        "Country": country or "USA",
        "CountryName": country or "USA",
        "FirstName": first_name,
        "LastName": "",
        "ContactEmail": "; ".join(dict.fromkeys(emails)),
    }


def extract_header_blocks(lines: List[str]) -> Tuple[List[str], List[str]]:
    sold_start = ship_start = None
    for i, line in enumerate(lines):
        if "Sold To:" in line and "Ship To:" in line:
            sold_start = i + 1
            ship_start = i + 1
            break
    if sold_start is None:
        return [], []

    end = len(lines)
    for j in range(sold_start, len(lines)):
        if "Reference" in lines[j] and "PO Number" in lines[j]:
            end = j
            break

    # pdfplumber often returns sold/ship blocks as separate lines in sequence.
    block = [clean_text(x) for x in lines[sold_start:end] if clean_text(x)]
    if len(block) <= 1:
        return block, block

    # Heuristic: split at the repeated company/name or at visible ship line marker if present.
    mid = len(block) // 2
    for idx in range(1, len(block)):
        if block[idx] == block[0] and idx >= 2:
            mid = idx
            break
    return block[:mid], block[mid:]


def parse_header(lines: List[str], fallback_pdf_name: str) -> Dict[str, str]:
    quote_date = ""
    quote_number = ""
    reference = ""
    po_number = ""
    customer_no = ""
    salesperson = ""
    ship_via = ""
    terms = ""

    for line in lines[:20]:
        m = re.search(r"\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{1,2},\s+\d{4}\b", line)
        if m and not quote_date:
            quote_date = m.group(0)
        m = re.search(r"\b(?:QT[\w\-]+|RQ[\w\-]+)\b", line)
        if m:
            quote_number = m.group(0)

    # Header row is usually right before quote number: reference/po/customer/sales/ship/terms.
    for i, line in enumerate(lines[:18]):
        if quote_number and quote_number in line and i > 0:
            prev = clean_text(lines[i - 1])
            tokens = prev.split()
            if len(tokens) >= 4:
                # Parse from right: terms, ship via, salesperson, customer no, remaining reference/po
                terms = tokens[-1]
                ship_via = tokens[-2]
                salesperson = tokens[-3]
                customer_no = tokens[-4]
                left = " ".join(tokens[:-4])
                reference = left
            break

    # Visual table sometimes has reference + PO in left side; infer PO when obvious in reference text is not personal name.
    if reference:
        # Known behavior from user's manual entry: PO is the middle text when the left side contains name + PO.
        pass

    sold_lines, ship_lines = extract_header_blocks(lines)
    ship = parse_address_block(ship_lines or sold_lines)
    sold = parse_address_block(sold_lines)
    if not ship.get("ContactEmail"):
        ship["ContactEmail"] = sold.get("ContactEmail", "")

    return {
        "ReferralManagerCode": salesperson,
        "ReferralManagerName": "",
        "ReferralEmail": "",
        "Brand": "Alcorn Industrial Inc",
        "QuoteNumber": quote_number or fallback_pdf_name.rsplit(".", 1)[0],
        "QuoteVersion": "",
        "QuoteDate": quote_date,
        "QuoteValidUntil": "",
        "Customer Number/ID": customer_no,
        "Company": ship.get("Company", ""),
        "Address": ship.get("Address", ""),
        "Country": ship.get("Country", ""),
        "City": ship.get("City", ""),
        "State": ship.get("State", ""),
        "ZipCode": ship.get("ZipCode", ""),
        "CountryName": ship.get("CountryName", ""),
        "FirstName": ship.get("FirstName", ""),
        "LastName": ship.get("LastName", ""),
        "ContactEmail": ship.get("ContactEmail", ""),
        "ContactPhone": "",
        "Webaddress": "",
        "Writer Name": "",
        "CustomerPONumber": po_number,
        "DemoQuote": "",
        "Duns": "",
        "SIC": "",
        "NAICS": "",
        "LineOfBusiness": "",
        "LinkedInProfile": "",
        "PhoneResearch": "",
        "PhoneSuppression": "",
        "ParentName": "",
        "ShipVia": ship_via,
        "Terms": terms,
        "Reference": reference,
    }


def split_item_fields(item_text: str) -> Tuple[str, str, str]:
    """Return item_number, customer_item_number, description."""
    tokens = item_text.split()
    if not tokens:
        return "", "", ""

    item_number = tokens[0]
    rest = tokens[1:]
    customer_item = ""

    if item_number.upper() in {"MISC", "PARTS", "CERT", "LABOR", "SUP", "MIT", "IR", "DYNA", "MAKITA", "MIL", "ARO", "T/C", "HUCK", "HONSA"}:
        if item_number.upper() == "PARTS" and rest and rest[0] == "&" and len(rest) >= 3 and rest[1].upper() == "MISC":
            item_number = "PARTS & MISC"
            rest = rest[2:]
        # For MISC/PARTS & MISC/CERT/SUP etc, the customer item is normally the next part number.
        if rest:
            # LABOR may include SN before description; keep model in item id and SN in customer item number.
            customer_item = rest[0]
            rest = rest[1:]
    else:
        # If second token is @, skip it; otherwise no separate customer item number.
        if rest and rest[0] == "@":
            rest = rest[1:]

    if rest and rest[0] == "@":
        rest = rest[1:]
    desc = " ".join(rest)
    return clean_text(item_number), clean_text(customer_item), clean_text(desc)


def normalize_item_id(item_number: str, customer_item: str) -> str:
    # User-validated rule: if item number is generic, use real customer item number.
    generic = {"MISC", "PARTS & MISC"}
    if item_number.upper() in generic and customer_item:
        return customer_item
    return item_number


def extract_items(lines: List[str]) -> List[Dict[str, object]]:
    rows = []
    in_items = False
    current = None

    for raw_line in lines:
        line = clean_text(raw_line)
        if not line:
            continue
        lower = line.lower()
        if "qty." in lower and "ord" in lower:
            in_items = True
            continue
        if not in_items:
            continue
        if lower.startswith(STOP_PHRASES):
            break
        if any(lower.startswith(x) for x in ["customer item number", "item number", "description unit price"]):
            continue

        m = ITEM_START_RE.match(line)
        if m:
            if current:
                rows.append(current)
            qty = int(m.group(1))
            item_text = clean_text(m.group(2))
            unit_price = money_to_float(m.group(3))
            total = money_to_float(m.group(4))
            item_number, customer_item, desc = split_item_fields(item_text)
            current = {
                "raw_item_number": item_number,
                "customer_item_number": customer_item,
                "item_id": normalize_item_id(item_number, customer_item),
                "item_desc": desc,
                "Quantity": qty,
                "Unit Price": unit_price,
                "List Price": unit_price,
                "TotalSales": total,
                "UOM": "",
                "Manufacturer": "",
                "manufacturer_part_number": "",
            }
            continue

        if current:
            if should_ignore_note(line):
                continue
            # If a continuation line has price pair at the end, it is likely a wrapped item; handled poorly by text extraction.
            # Otherwise append to description.
            if not MONEY_RE.match(line):
                current["item_desc"] = clean_text(f"{current['item_desc']} {line}")

    if current:
        rows.append(current)
    return rows



def group_words_by_line(words, y_tol: float = 3.0):
    """Group pdfplumber words into visual text lines by vertical position."""
    lines = []
    for w in sorted(words, key=lambda x: (x["top"], x["x0"])):
        placed = False
        for line in lines:
            if abs(line["top"] - w["top"]) <= y_tol:
                line["words"].append(w)
                line["top"] = min(line["top"], w["top"])
                placed = True
                break
        if not placed:
            lines.append({"top": w["top"], "words": [w]})
    for line in lines:
        line["words"].sort(key=lambda x: x["x0"])
        line["text"] = clean_text(" ".join(w["text"] for w in line["words"]))
    return sorted(lines, key=lambda x: x["top"])


def words_in_range(words, left: float, right: float) -> str:
    return clean_text(" ".join(w["text"] for w in words if left <= w["x0"] < right))


def extract_items_from_pdf_columns(pdf_bytes: bytes) -> List[Dict[str, object]]:
    """Extract line items from the visual columns, not only raw text.

    This is the primary parser. It handles the Alcorn table where pdf text may split
    Qty. and Ord. into separate lines or move customer-item values onto wrapped lines.
    """
    all_rows: List[Dict[str, object]] = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            words = page.extract_words(x_tolerance=1, y_tolerance=3) or []
            if not words:
                continue
            # Header and bottom of table.
            qty_tops = [w["top"] for w in words if w["text"].lower().startswith("qty") and w["x0"] < 70]
            tax_tops = [w["top"] for w in words if w["text"].lower().startswith("tax")]
            header_y = min(qty_tops) if qty_tops else 260
            bottom_y = min([y for y in tax_tops if y > header_y] or [page.height - 80])
            table_words = [w for w in words if header_y + 15 <= w["top"] < bottom_y - 2]
            lines = group_words_by_line(table_words)
            # Visual column boundaries for US Letter Alcorn quote PDFs.
            c_qty = (25, 55)
            c_item = (55, 155)
            c_customer = (155, 252)
            c_desc = (252, 455)
            c_unit = (455, 510)
            c_ext = (510, 590)

            row_indices = []
            for idx, line in enumerate(lines):
                qty_text = words_in_range(line["words"], *c_qty)
                if re.fullmatch(r"\d+", qty_text):
                    row_indices.append(idx)
            row_indices.append(len(lines))

            for pos in range(len(row_indices) - 1):
                start, end = row_indices[pos], row_indices[pos + 1]
                group = lines[start:end]
                if not group:
                    continue
                first_words = group[0]["words"]
                qty_text = words_in_range(first_words, *c_qty)
                if not qty_text.isdigit():
                    continue
                qty = int(qty_text)
                item_number = words_in_range(first_words, *c_item)
                customer_parts = []
                desc_parts = []
                unit_price = None
                total_sales = None

                for line in group:
                    ws = line["words"]
                    cust = words_in_range(ws, *c_customer)
                    desc = words_in_range(ws, *c_desc)
                    unit = words_in_range(ws, *c_unit)
                    ext = words_in_range(ws, *c_ext)
                    if cust:
                        customer_parts.append(cust)
                    if desc and not should_ignore_note(desc):
                        desc_parts.append(desc)
                    if unit and MONEY_RE.match(unit):
                        unit_price = money_to_float(unit)
                    if ext and MONEY_RE.match(ext):
                        total_sales = money_to_float(ext)

                customer_item = clean_text(" ".join(customer_parts))
                desc = clean_text(" ".join(desc_parts))
                if not item_number and not customer_item and not desc:
                    continue
                all_rows.append({
                    "raw_item_number": clean_text(item_number),
                    "customer_item_number": customer_item,
                    "item_id": normalize_item_id(clean_text(item_number), customer_item),
                    "item_desc": desc,
                    "Quantity": qty,
                    "Unit Price": unit_price,
                    "List Price": unit_price,
                    "TotalSales": total_sales,
                    "UOM": "",
                    "Manufacturer": "",
                    "manufacturer_part_number": "",
                })
    return all_rows


def extract_pdf_rows(pdf_bytes: bytes, pdf_name: str) -> List[Dict[str, object]]:
    text_lines = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text(x_tolerance=1, y_tolerance=3) or ""
            text_lines.extend(text.splitlines())
    header = parse_header(text_lines, pdf_name)

    # Primary extractor: visual columns. Fallback: raw text parser.
    items = extract_items_from_pdf_columns(pdf_bytes)
    if not items:
        items = extract_items(text_lines)

    rows = []
    for item in items:
        row = {h: "" for h in FIXED_HEADERS}
        for key, value in header.items():
            if key in row:
                row[key] = value
        for key in ["item_id", "item_desc", "UOM", "Quantity", "Unit Price", "List Price", "TotalSales", "Manufacturer", "manufacturer_part_number"]:
            row[key] = item.get(key, "")
        row["PDF"] = pdf_name
        rows.append(row)
    return rows


def extract_pdfs_from_msg(msg_bytes: bytes, base_dt: datetime) -> List[Tuple[str, bytes]]:
    if extract_msg is None:
        raise RuntimeError("extract-msg is not installed. Run: pip install -r requirements.txt")

    msg = extract_msg.Message(io.BytesIO(msg_bytes))
    pdfs: List[Tuple[str, bytes]] = []
    for i, attachment in enumerate(msg.attachments, start=1):
        long_name = getattr(attachment, "longFilename", None) or ""
        short_name = getattr(attachment, "shortFilename", None) or ""
        original_name = long_name or short_name or f"attachment_{i}.pdf"
        if not original_name.lower().endswith(".pdf"):
            continue
        data = attachment.data
        if not data:
            continue
        stamp = base_dt.strftime("%Y%m%d_%H%M%S")
        milli = f"{base_dt.microsecond // 1000:03d}"
        safe_name = f"{PDF_PREFIX}_{stamp}_{milli}_{i:03d}.pdf"
        pdfs.append((safe_name, data))
    return pdfs


def build_excel(rows: List[Dict[str, object]]) -> bytes:
    df = pd.DataFrame(rows, columns=FIXED_HEADERS)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Quotes")
        ws = writer.book["Quotes"]
        ws.freeze_panes = "A2"
        for cell in ws[1]:
            cell.style = "Headline 3"
        for col in ws.columns:
            letter = col[0].column_letter
            max_len = max(len(str(c.value or "")) for c in col[:200])
            ws.column_dimensions[letter].width = min(max(max_len + 2, 10), 40)
    return output.getvalue()


def build_zip(excel_bytes: bytes, pdfs: List[Tuple[str, bytes]]) -> bytes:
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("alcorn_quote_extraction.xlsx", excel_bytes)
        for pdf_name, pdf_bytes in pdfs:
            zf.writestr(f"pdf/{pdf_name}", pdf_bytes)
    return zip_buffer.getvalue()


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.title(APP_TITLE)
    st.caption("Upload Outlook .msg files. PDFs are processed in memory only and are not cached.")

    uploaded_msgs = st.file_uploader("Upload .msg files", type=["msg"], accept_multiple_files=True)
    if not uploaded_msgs:
        st.info("Upload one or more .msg files to extract PDFs and build the Excel + PDF ZIP.")
        return

    if st.button("Extract quotes and build ZIP", type="primary"):
        all_pdfs: List[Tuple[str, bytes]] = []
        all_rows: List[Dict[str, object]] = []
        errors: List[str] = []
        run_dt = datetime.now()

        for msg_file in uploaded_msgs:
            try:
                msg_bytes = msg_file.getvalue()
                pdfs = extract_pdfs_from_msg(msg_bytes, run_dt)
                all_pdfs.extend(pdfs)
                for pdf_name, pdf_bytes in pdfs:
                    try:
                        rows = extract_pdf_rows(pdf_bytes, pdf_name)
                        all_rows.extend(rows)
                    except Exception as exc:  # continue processing other PDFs
                        errors.append(f"{pdf_name}: {exc}")
            except Exception as exc:
                errors.append(f"{msg_file.name}: {exc}")

        if not all_pdfs:
            st.error("No PDF attachments were found in the uploaded MSG file(s).")
            return

        excel_bytes = build_excel(all_rows)
        zip_bytes = build_zip(excel_bytes, all_pdfs)

        st.success(f"Processed {len(all_pdfs)} PDF(s) and extracted {len(all_rows)} line item row(s).")
        if errors:
            st.warning("Some files had extraction issues:")
            st.write(errors)

        st.download_button(
            label="Download Excel + renamed PDFs ZIP",
            data=zip_bytes,
            file_name=f"alcorn_quote_output_{run_dt.strftime('%Y%m%d_%H%M%S')}.zip",
            mime="application/zip",
        )

        with st.expander("Preview extracted rows"):
            st.dataframe(pd.DataFrame(all_rows, columns=FIXED_HEADERS), use_container_width=True)


if __name__ == "__main__":
    main()
