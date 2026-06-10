import io
import re
import zipfile
from datetime import datetime, timedelta
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
    "ReferralManager",
    "ReferralEmail",
    "Brand",
    "QuoteNumber",
    "QuoteVersion",
    "QuoteDate",
    "QuoteValidDate",
    "Customer Number/ID",
    "Company",
    "Address",
    "County",
    "City",
    "State",
    "ZipCode",
    "Country",
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
    "Manufacturer_ID",
    "manufacturer_Name",
    "Writer Name",
    "CustomerPONumber",
    "PDF",
    "DemoQuote",
    "Duns",
    "SIC",
    "NAICS",
    "LineOfBusiness",
    "LinkedinProfile",
    "PhoneResearched",
    "PhoneSupplied",
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
    re.compile(r".*lead time.*", re.I),
    re.compile(r"^eta\s+", re.I),
    re.compile(r"^\d+\s*-\s*\d+\s*(day|days|week|weeks)", re.I),
    re.compile(r"^\d+\s*(day|days|week|weeks)\b", re.I),
    re.compile(r"^all .* are in", re.I),
    re.compile(r"^quoted closest", re.I),
    re.compile(r"^made to order", re.I),
    re.compile(r"^remaining \(.*\) is", re.I),
    re.compile(r"^\(\d+\)\s+are", re.I),
    re.compile(r"^stock$", re.I),
    re.compile(r"^length for", re.I),
    re.compile(r"^determined$", re.I),
    re.compile(r"^quote is for", re.I),
    re.compile(r"^uses ", re.I),
    re.compile(r"^includes:", re.I),
    re.compile(r"^new tool\s+\$", re.I),
    re.compile(r"^repair is\s+", re.I),
    re.compile(r"^cert from", re.I),
]


def clean_text(value: str) -> str:
    value = value or ""
    value = value.replace("\u201c", '"').replace("\u201d", '"').replace("\u2019", "'")
    value = re.sub(r"\s+", " ", value).strip()
    return value




def clean_quote_for_filename(quote_number: str) -> str:
    """Return a filename-safe quote number.

    Business rule: keep the quote readable, replace hyphens and other
    non-alphanumeric characters with underscores, collapse repeated
    underscores, and trim leading/trailing underscores.
    Example: QT-ALCPT1698 -> QT_ALCPT1698.
    """
    quote_number = clean_text(quote_number)
    quote_number = re.sub(r"[^A-Za-z0-9]+", "_", quote_number)
    quote_number = re.sub(r"_+", "_", quote_number).strip("_")
    return quote_number or "UNKNOWN"


def make_pdf_output_name(quote_number: str, stamp_dt: datetime) -> str:
    """Build final renamed PDF name.

    Format: Alcorn_<CleanQuoteNumber>_<YYYYMMDD>_<HHMMSS>_<milliseconds>.pdf
    """
    clean_quote = clean_quote_for_filename(quote_number)
    milliseconds = stamp_dt.microsecond // 1000
    return f"Alcorn_{clean_quote}_{stamp_dt:%Y%m%d}_{stamp_dt:%H%M%S}_{milliseconds:03d}.pdf"

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



def normalize_zip(zip_code: str) -> str:
    zip_code = clean_text(zip_code)
    # Manual XLSX uses 5-digit ZIPs for US ZIP+4 values.
    m = re.match(r"^(\d{5})-\d{4}$", zip_code)
    return m.group(1) if m else zip_code

def parse_address_block(block_lines: List[str]) -> Dict[str, str]:
    cleaned = [clean_text(x) for x in block_lines if clean_text(x)]
    emails = []
    non_email = []
    for line in cleaned:
        emails.extend(EMAIL_RE.findall(line))
        stripped = EMAIL_RE.sub("", line)
        # Remove operational notes from address parsing.
        if re.search(r"(?i)send invoice|no pack slip|packages|keep packages|copy of po|inv thru coupa|invoice[s]? only|email both|email cc", stripped):
            continue
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
            city, state, zip_code = c, s, normalize_zip(z)
            if co:
                country = co
            continue
        if "customers only" in line.lower() or line.lower() == "counter sales":
            continue
        address_parts.append(line)

    # Manual-entry rule: department/role ATTN lines are not stored as FirstName,
    # but a person name like "Eric" is kept.
    if first_name and re.search(r"(?i)account|payable|receiv|department|dept|store|room|maintenance", first_name):
        first_name = ""
    return {
        "Company": company,
        "Address": ", ".join(address_parts),
        "City": city,
        "State": state,
        "ZipCode": zip_code,
        "County": "",
        "Country": country or "USA",
        "FirstName": first_name,
        "LastName": "",
        "ContactEmail": (list(dict.fromkeys(emails))[0] if emails else ""),
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



def extract_visual_lines(page, bbox):
    words = page.extract_words(x_tolerance=1, y_tolerance=3) or []
    left, top, right, bottom = bbox
    selected = [w for w in words if left <= w["x0"] < right and top <= w["top"] < bottom]
    lines = []
    for w in sorted(selected, key=lambda x: (x["top"], x["x0"])):
        placed = False
        for line in lines:
            if abs(line["top"] - w["top"]) <= 3:
                line["words"].append(w)
                line["top"] = min(line["top"], w["top"])
                placed = True
                break
        if not placed:
            lines.append({"top": w["top"], "words": [w]})
    out = []
    for line in sorted(lines, key=lambda x: x["top"]):
        line["words"].sort(key=lambda x: x["x0"])
        txt = clean_text(" ".join(w["text"] for w in line["words"]))
        if txt:
            out.append(txt)
    return out


def clean_header_block(lines: List[str]) -> List[str]:
    out = []
    for line in lines:
        line = clean_text(line)
        if not line or line.lower() in {"sold to:", "ship to:", "to:"}:
            continue
        out.append(line)
    return out


def extract_header_values_from_page(page) -> Dict[str, str]:
    words = page.extract_words(x_tolerance=1, y_tolerance=3) or []

    def box_text(left, top, right, bottom):
        vals = [w for w in words if left <= w["x0"] < right and top <= w["top"] < bottom]
        return clean_text(" ".join(w["text"] for w in sorted(vals, key=lambda x: (x["top"], x["x0"]))) )

    # Right-top date/order box.
    quote_date = ""
    quote_number = ""
    top_box = extract_visual_lines(page, (455, 15, 590, 70))
    for line in top_box:
        m = re.search(r"\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{1,2},\s+\d{4}\b", line)
        if m:
            quote_date = m.group(0)
        m = re.search(r"\b(?:QT[\w\-]+|RQ[\w\-]+)\b", line)
        if m:
            quote_number = m.group(0)

    # Bottom values in the header cells. Header labels are at y~234, values at y~243-252.
    y1, y2 = 241, 254
    reference = box_text(25, y1, 135, y2)
    po_number = box_text(135, y1, 250, y2)
    customer_no = box_text(250, y1, 320, y2)
    salesperson = box_text(320, y1, 405, y2)
    order_date = box_text(405, y1, 470, y2)
    ship_via = box_text(470, y1, 520, y2)
    terms = box_text(520, y1, 590, y2)
    if order_date and not quote_date:
        quote_date = order_date

    sold_lines = clean_header_block(extract_visual_lines(page, (25, 135, 300, 225)))
    ship_lines = clean_header_block(extract_visual_lines(page, (295, 135, 590, 225)))
    return {
        "quote_date": quote_date,
        "quote_number": quote_number,
        "reference": reference,
        "po_number": po_number,
        "customer_no": customer_no,
        "salesperson": salesperson,
        "ship_via": ship_via,
        "terms": terms,
        "sold_lines": sold_lines,
        "ship_lines": ship_lines,
    }



def normalize_customer_no(value: str) -> str:
    value = clean_text(value)
    if re.fullmatch(r"0+\d+", value):
        return str(int(value))
    return value

def parse_header(lines: List[str], fallback_pdf_name: str, first_page=None) -> Dict[str, str]:
    # Prefer visual extraction. Raw text ordering often mixes Sold-To and Ship-To blocks and causes bad addresses.
    visual = extract_header_values_from_page(first_page) if first_page is not None else {}

    quote_date = visual.get("quote_date", "")
    quote_number = visual.get("quote_number", "")
    reference = visual.get("reference", "")
    po_number = visual.get("po_number", "")
    customer_no = visual.get("customer_no", "")
    salesperson = visual.get("salesperson", "")
    ship_via = visual.get("ship_via", "")
    terms = visual.get("terms", "")

    # Fallback only for missing date/order number.
    for line in lines[:25]:
        if not quote_date:
            m = re.search(r"\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{1,2},\s+\d{4}\b", line)
            if m:
                quote_date = m.group(0)
        if not quote_number:
            m = re.search(r"\b(?:QT[\w\-]+|RQ[\w\-]+)\b", line)
            if m:
                quote_number = m.group(0)

    sold_lines = visual.get("sold_lines", [])
    ship_lines = visual.get("ship_lines", [])
    if not sold_lines or not ship_lines:
        sold_lines, ship_lines = extract_header_blocks(lines)

    ship = parse_address_block(ship_lines or sold_lines)
    sold = parse_address_block(sold_lines)
    # If Ship-To has only delivery instructions/person and no usable address, keep Ship-To company
    # but fill missing address components from Sold-To.
    ship_had_city = bool(ship.get("City") or ship.get("State") or ship.get("ZipCode"))
    for field in ["Address", "City", "State", "ZipCode", "Country"]:
        if not ship.get(field) and sold.get(field):
            ship[field] = sold[field]
    if not ship_had_city and sold.get("Address") and re.search(r"(?i)will deliver|deliver", ship.get("Address", "")):
        ship["Address"] = sold.get("Address", "")
    # Do not copy sold-to email into ContactEmail unless it also appears in Ship-To.

    return {
        "ReferralManagerCode": salesperson,
        "ReferralManager": "",
        "ReferralEmail": "",
        "Brand": "Alcorn Industrial Inc",
        "QuoteNumber": quote_number or fallback_pdf_name.rsplit(".", 1)[0],
        "QuoteVersion": "",
        "QuoteDate": quote_date,
        "QuoteValidDate": "",
        "Customer Number/ID": normalize_customer_no(customer_no),
        "Company": ship.get("Company", ""),
        "Address": ship.get("Address", ""),
        "County": ship.get("County", ""),
        "City": ship.get("City", ""),
        "State": ship.get("State", ""),
        "ZipCode": ship.get("ZipCode", ""),
        "Country": ship.get("Country", ""),
        "FirstName": ship.get("FirstName", ""),
        "LastName": "",
        "ContactEmail": ship.get("ContactEmail", ""),
        "ContactPhone": "",
        "Webaddress": "",
        "Writer Name": "",
        "CustomerPONumber": "",
        "DemoQuote": "",
        "Duns": "",
        "SIC": "",
        "NAICS": "",
        "LineOfBusiness": "",
        "LinkedinProfile": "",
        "PhoneResearched": "",
        "PhoneSupplied": "",
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
    # Manual-entry rule confirmed from comparison: item_id is the PDF Item Number column.
    # Customer Item Number is not exported to the fixed CRM sheet.
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
                "List Price": "",
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
    parts = []
    for w in words:
        txt = w["text"]
        x0, x1 = w["x0"], w["x1"]
        if x1 <= left or x0 >= right:
            continue
        # Usually words are fully inside a column. If a very long token crosses a column
        # boundary, split proportionally so customer-item text does not swallow description.
        if x0 < left or x1 > right:
            if len(txt) >= 12 and x1 > x0:
                start_idx = max(0, int(round((max(left, x0) - x0) / (x1 - x0) * len(txt))))
                end_idx = min(len(txt), int(round((min(right, x1) - x0) / (x1 - x0) * len(txt))))
                piece = txt[start_idx:end_idx]
                if piece:
                    parts.append(piece)
            else:
                # For normal small tokens crossing by a point or two, keep it with its start column.
                if left <= x0 < right:
                    parts.append(txt)
        else:
            parts.append(txt)
    return clean_text(" ".join(parts))


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
                    "List Price": "",
                    "TotalSales": total_sales,
                    "UOM": "",
                    "Manufacturer_ID": "",
                    "manufacturer_Name": "",
                })
    return all_rows


def extract_pdf_rows(pdf_bytes: bytes, pdf_name: str) -> List[Dict[str, object]]:
    text_lines = []
    first_page = None
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        first_page = pdf.pages[0] if pdf.pages else None
        for page in pdf.pages:
            text = page.extract_text(x_tolerance=1, y_tolerance=3) or ""
            text_lines.extend(text.splitlines())
    header = parse_header(text_lines, pdf_name, first_page)

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
        for key in ["item_id", "item_desc", "UOM", "Quantity", "Unit Price", "List Price", "TotalSales", "Manufacturer_ID", "manufacturer_Name"]:
            row[key] = item.get(key, "")
        row["PDF"] = f"Alcorn_{header.get('QuoteNumber', '').strip()}.pdf" if header.get('QuoteNumber') else pdf_name
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
        # Keep the original attachment name temporarily. The final ZIP PDF name
        # is created after parsing the quote number from the PDF.
        safe_name = re.sub(r"[^A-Za-z0-9_.-]+", "_", original_name)
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

        pdf_counter = 0
        for msg_file in uploaded_msgs:
            try:
                msg_bytes = msg_file.getvalue()
                pdfs = extract_pdfs_from_msg(msg_bytes, run_dt)
                for original_pdf_name, pdf_bytes in pdfs:
                    pdf_counter += 1
                    try:
                        rows = extract_pdf_rows(pdf_bytes, original_pdf_name)
                        quote_number = ""
                        if rows:
                            quote_number = str(rows[0].get("QuoteNumber", ""))

                        # Add 1 millisecond per PDF so duplicate quote numbers in the
                        # same run still receive unique filenames.
                        pdf_stamp = run_dt + timedelta(milliseconds=pdf_counter - 1)
                        final_pdf_name = make_pdf_output_name(quote_number or original_pdf_name, pdf_stamp)

                        # Ensure the Excel PDF column exactly matches the renamed PDF
                        # inside the ZIP.
                        for row in rows:
                            row["PDF"] = final_pdf_name

                        all_pdfs.append((final_pdf_name, pdf_bytes))
                        all_rows.extend(rows)
                    except Exception as exc:  # continue processing other PDFs
                        errors.append(f"{original_pdf_name}: {exc}")
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
