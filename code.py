# app.py

import os
import re
import io
import zipfile
import tempfile
from copy import copy
from datetime import datetime, date, time
from collections import defaultdict

import streamlit as st
import fitz  # PyMuPDF
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from pypdf import PdfReader, PdfWriter


# ============================================================
# SETTINGS
# ============================================================

OUTPUT_FILE_NAME = "Populated_Template.xlsx"
MATCHED_PDF_FILE_NAME = "Matched_Manifest_Packet.pdf"

OPENDOCK_SHEET = "OPENDOCK"
MG_REPORT_SHEET = "MG REPORT"
DISPATCH_SHEET = "DISPATCH SHEET"

OPENDOCK_START_ROW = 2
MG_REPORT_START_ROW = 2

MAX_OPENDOCK_ROWS = 1000
MAX_MG_REPORT_ROWS = 1000


# ============================================================
# BASIC HELPERS
# ============================================================

def clean_spaces(value) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def normalize_header(value) -> str:
    if value is None:
        return ""
    return re.sub(r"[^A-Z0-9]", "", str(value).strip().upper())


def normalize_load_number(value) -> str:
    if value is None:
        return ""

    text = str(value).strip()
    text = re.sub(r"\.0$", "", text)
    text = re.sub(r"^LD", "", text, flags=re.IGNORECASE)
    text = re.sub(r"[^0-9]", "", text)

    return text if text.isdigit() else ""


def parse_number(value):
    if value is None:
        return 0

    text = str(value).replace(",", "").strip()

    if not text:
        return 0

    try:
        number = float(text)
        if number.is_integer():
            return int(number)
        return number
    except Exception:
        return 0


def normalize_date(value) -> str:
    if value is None:
        return ""

    if isinstance(value, datetime):
        return value.strftime("%m/%d/%Y")

    if isinstance(value, date):
        return value.strftime("%m/%d/%Y")

    text = clean_spaces(value)

    if not text:
        return ""

    match = re.search(r"(\d{1,2}/\d{1,2}/\d{2,4})", text)
    if match:
        text = match.group(1)

    for fmt in ["%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d"]:
        try:
            dt = datetime.strptime(text, fmt)
            return dt.strftime("%m/%d/%Y")
        except Exception:
            pass

    return ""


def normalize_time(value) -> str:
    if value is None:
        return ""

    if isinstance(value, datetime):
        return value.strftime("%H:%M")

    if isinstance(value, time):
        return value.strftime("%H:%M")

    if isinstance(value, (int, float)) and 0 <= value < 1:
        total_minutes = int(round(value * 24 * 60))
        hours = (total_minutes // 60) % 24
        minutes = total_minutes % 60
        return f"{hours:02d}:{minutes:02d}"

    text = clean_spaces(value)

    if not text:
        return ""

    match = re.search(r"\b(\d{1,2}):(\d{2})\b", text)
    if match and "AM" not in text.upper() and "PM" not in text.upper():
        return f"{int(match.group(1)):02d}:{match.group(2)}"

    compact_text = text.upper().replace(" ", "")

    try:
        dt = datetime.strptime(compact_text, "%I:%M%p")
        return dt.strftime("%H:%M")
    except Exception:
        pass

    return ""


def sort_datetime_key(record):
    date_text = record.get("appt_date", "")
    time_text = record.get("appt_time", "")

    if not date_text:
        return datetime.max

    if not time_text:
        time_text = "23:59"

    try:
        return datetime.strptime(f"{date_text} {time_text}", "%m/%d/%Y %H:%M")
    except Exception:
        return datetime.max


def get_header_map(ws, header_row=1):
    headers = {}

    for col in range(1, ws.max_column + 1):
        value = ws.cell(header_row, col).value

        if value not in [None, ""]:
            headers[normalize_header(value)] = col

    return headers


def find_col(header_map, possible_names, default_col=None):
    for name in possible_names:
        key = normalize_header(name)
        if key in header_map:
            return header_map[key]

    return default_col


def copy_cell_style(source_cell, target_cell):
    if source_cell.has_style:
        target_cell.font = copy(source_cell.font)
        target_cell.fill = copy(source_cell.fill)
        target_cell.border = copy(source_cell.border)
        target_cell.alignment = copy(source_cell.alignment)
        target_cell.number_format = source_cell.number_format
        target_cell.protection = copy(source_cell.protection)


def copy_row_format(ws, source_row, target_row, max_col):
    ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height

    for col in range(1, max_col + 1):
        source_cell = ws.cell(source_row, col)
        target_cell = ws.cell(target_row, col)

        copy_cell_style(source_cell, target_cell)


def clear_range_values(ws, start_row, end_row, columns):
    for row in range(start_row, end_row + 1):
        for col in columns:
            ws.cell(row, col).value = None


# ============================================================
# PART 1 - MG MANIFEST MATCHER
# ============================================================

def sample_text(pdf_path, pages=5):
    doc = fitz.open(pdf_path)
    text = ""

    for i in range(min(pages, len(doc))):
        text += doc[i].get_text("text") + "\n"

    doc.close()
    return text.upper()


def parse_load_from_mg_text(text):
    match = re.search(r"\bLOAD\s*:\s*([A-Z]{1,5}\d+)", text, re.IGNORECASE)
    return match.group(1).upper().strip() if match else None


def parse_pu_appt_from_mg_text(text):
    match = re.search(
        r"\bPU\s+APPT\s*:\s*(\d{1,2}/\d{1,2}/\d{4})\s+(\d{1,2}:\d{2})",
        text,
        re.IGNORECASE
    )

    if not match:
        return None

    try:
        return datetime.strptime(
            match.group(1) + " " + match.group(2),
            "%m/%d/%Y %H:%M"
        )
    except Exception:
        return None


def group_pages_by_load(pdf_path):
    doc = fitz.open(pdf_path)

    groups = {}
    current_load = None

    for i in range(len(doc)):
        text = doc[i].get_text("text")
        found_load = parse_load_from_mg_text(text)

        if found_load:
            current_load = found_load

            if current_load not in groups:
                groups[current_load] = {
                    "pages": [],
                    "text": ""
                }

        if current_load:
            groups[current_load]["pages"].append(i)
            groups[current_load]["text"] += "\n" + text

    doc.close()
    return groups


def extract_uploaded_manifest_files(uploaded_files, workdir):
    pdf_files = []

    for uploaded_file in uploaded_files:
        file_path = os.path.join(workdir, uploaded_file.name)

        with open(file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        if uploaded_file.name.lower().endswith(".zip"):
            folder = os.path.join(workdir, uploaded_file.name.replace(".zip", ""))
            os.makedirs(folder, exist_ok=True)

            with zipfile.ZipFile(file_path, "r") as z:
                z.extractall(folder)

            for root, dirs, files in os.walk(folder):
                for file in files:
                    if file.lower().endswith(".pdf"):
                        pdf_files.append(os.path.join(root, file))

        elif uploaded_file.name.lower().endswith(".pdf"):
            pdf_files.append(file_path)

    return pdf_files


def build_matched_packet(uploaded_files):
    with tempfile.TemporaryDirectory() as workdir:
        pdf_files = extract_uploaded_manifest_files(uploaded_files, workdir)

        if len(pdf_files) < 2:
            raise Exception("Need at least two PDFs or ZIPs containing PDFs.")

        loading_pdf = None
        shipping_pdf = None

        for pdf in pdf_files:
            text = sample_text(pdf)

            if "LOADING MANIFEST" in text and "SHIPPING MANIFEST" not in text:
                loading_pdf = pdf

            if "SHIPPING MANIFEST" in text or "PICKUP TOTAL" in text:
                shipping_pdf = pdf

        if not loading_pdf or not shipping_pdf:
            raise Exception("Could not identify loading and shipping PDFs.")

        loading_groups = group_pages_by_load(loading_pdf)
        shipping_groups = group_pages_by_load(shipping_pdf)

        all_loads = sorted(
            set(loading_groups.keys()) |
            set(shipping_groups.keys())
        )

        records = []

        for load in all_loads:
            loading_text = loading_groups.get(load, {}).get("text", "")
            shipping_text = shipping_groups.get(load, {}).get("text", "")

            pickup_datetime = parse_pu_appt_from_mg_text(loading_text or shipping_text)

            records.append({
                "load": load,
                "datetime": pickup_datetime,
                "loading_pages": loading_groups.get(load, {}).get("pages", []),
                "shipping_pages": shipping_groups.get(load, {}).get("pages", [])
            })

        records = sorted(
            records,
            key=lambda r: (
                r["datetime"] is None,
                r["datetime"] or datetime.max,
                r["load"]
            )
        )

        output_pdf = os.path.join(workdir, MATCHED_PDF_FILE_NAME)

        writer = PdfWriter()
        loading_reader = PdfReader(loading_pdf)
        shipping_reader = PdfReader(shipping_pdf)

        for record in records:
            for page_num in record["loading_pages"]:
                writer.add_page(loading_reader.pages[page_num])

            for page_num in record["shipping_pages"]:
                writer.add_page(shipping_reader.pages[page_num])

        with open(output_pdf, "wb") as f:
            writer.write(f)

        with open(output_pdf, "rb") as f:
            pdf_bytes = f.read()

        summary = {
            "Loading loads found": len(loading_groups),
            "Shipping loads found": len(shipping_groups),
            "Unique loads matched": len(records),
        }

        return pdf_bytes, summary


# ============================================================
# OPENDOCK PARSER
# ============================================================

def parse_opendock_excel(opendock_bytes: bytes) -> tuple[list[dict], dict]:
    wb = openpyxl.load_workbook(io.BytesIO(opendock_bytes), data_only=True)

    if "Appointments" in wb.sheetnames:
        ws = wb["Appointments"]
    else:
        ws = wb[wb.sheetnames[0]]

    headers = get_header_map(ws, 1)

    load_col = find_col(headers, ["Load Reference", "Load", "Load #"])
    appt_date_col = find_col(headers, ["Appt Date", "Appointment Date", "Date", "Pickup Date"])
    appt_time_col = find_col(headers, ["Appt Time", "Appointment Time", "Time", "Pickup Time"])
    arrival_date_col = find_col(headers, ["Arrival Date"])
    arrival_time_col = find_col(headers, ["Arrival Time"])
    departure_date_col = find_col(headers, ["Departure Date"])
    departure_time_col = find_col(headers, ["Departure Time"])
    status_col = find_col(headers, ["Status"])
    carrier_col = find_col(headers, ["Carrier Company", "Carrier"])
    load_type_col = find_col(headers, ["Load Type", "Type"])
    dock_col = find_col(headers, ["Dock", "Door"])
    direction_col = find_col(headers, ["Direction"])

    missing = []

    if not load_col:
        missing.append("Load Reference")

    if not appt_date_col:
        missing.append("Appt Date")

    if not appt_time_col:
        missing.append("Appt Time")

    if not carrier_col:
        missing.append("Carrier Company")

    if missing:
        raise ValueError("The Opendock report is missing these columns: " + ", ".join(missing))

    records = []
    skipped_inbound = []
    duplicate_tracker = defaultdict(int)

    for row in range(2, ws.max_row + 1):
        load_number = normalize_load_number(ws.cell(row, load_col).value)

        if not load_number:
            continue

        direction = clean_spaces(ws.cell(row, direction_col).value) if direction_col else ""

        if direction and direction.upper() != "OUTBOUND":
            skipped_inbound.append(load_number)
            continue

        duplicate_tracker[load_number] += 1

        record = {
            "load": load_number,
            "appt_date": normalize_date(ws.cell(row, appt_date_col).value),
            "appt_time": normalize_time(ws.cell(row, appt_time_col).value),
            "arrival_date": normalize_date(ws.cell(row, arrival_date_col).value) if arrival_date_col else "",
            "arrival_time": normalize_time(ws.cell(row, arrival_time_col).value) if arrival_time_col else "",
            "departure_date": normalize_date(ws.cell(row, departure_date_col).value) if departure_date_col else "",
            "departure_time": normalize_time(ws.cell(row, departure_time_col).value) if departure_time_col else "",
            "status": clean_spaces(ws.cell(row, status_col).value) if status_col else "",
            "carrier": clean_spaces(ws.cell(row, carrier_col).value) if carrier_col else "",
            "load_type": clean_spaces(ws.cell(row, load_type_col).value) if load_type_col else "",
            "dock": clean_spaces(ws.cell(row, dock_col).value) if dock_col else "",
            "direction": direction,
        }

        records.append(record)

    records.sort(key=lambda x: (sort_datetime_key(x), x.get("load", "")))

    duplicates = [
        load for load, count in duplicate_tracker.items()
        if count > 1
    ]

    summary = {
        "outbound_rows_loaded": len(records),
        "inbound_rows_skipped": len(skipped_inbound),
        "duplicate_outbound_loads": duplicates,
    }

    return records, summary


# ============================================================
# MG PDF PARSER
# ============================================================

def extract_pdf_pages(pdf_bytes: bytes) -> list[str]:
    pages = []

    with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
        for page in doc:
            pages.append(page.get_text("text"))

    return pages


def extract_mg_load_number(text: str) -> str:
    match = re.search(r"\bLOAD:\s*(?:LD)?\s*([A-Z0-9]+)", text, re.IGNORECASE)

    if not match:
        return ""

    return normalize_load_number(match.group(1))


def extract_mg_pu_appt(text: str) -> tuple[str, str]:
    match = re.search(
        r"\bPU\s+APPT:\s*(\d{1,2}/\d{1,2}/\d{4})\s+(\d{1,2}:\d{2})",
        text,
        re.IGNORECASE,
    )

    if not match:
        return "", ""

    return normalize_date(match.group(1)), normalize_time(match.group(2))


def extract_mg_carrier(text: str) -> str:
    match = re.search(r"\bCARR/SCT\s+TR:\s*(.+)", text, re.IGNORECASE)

    if not match:
        return ""

    carrier = clean_spaces(match.group(1).splitlines()[0])

    if carrier.upper().startswith("CPUC"):
        return "CPUC"

    carrier = re.sub(r"^[A-Z0-9]{3,6}\s+", "", carrier).strip()

    return clean_spaces(carrier)


def extract_mg_totals(text: str) -> tuple[float, float]:
    matches = re.findall(
        r"OUTBOUND\s+TOTALS:\s+([\d,]+(?:\.\d+)?)\s+([\d,]+(?:\.\d+)?)\s+([\d,]+(?:\.\d+)?)",
        text,
        re.IGNORECASE,
    )

    if not matches:
        return 0, 0

    weight, _volume, quantity = matches[-1]

    return parse_number(weight), parse_number(quantity)


def clean_customer_name(value: str) -> str:
    value = clean_spaces(value)

    if not value:
        return ""

    value = re.sub(r"^\d{5,12}\s*/?\s*", "", value)
    value = re.sub(r"^\d{5,12}\s+", "", value)
    value = re.sub(r"\s+TK$", "", value, flags=re.IGNORECASE)
    value = re.sub(r"\s+T$", "", value, flags=re.IGNORECASE)

    bad_fragments = [
        "RESER'S - TK - TOPEKA DISTRIBUTION CENTER",
        "TOPEKA DISTRIBUTION CENTER",
        "3121 SE 6TH AVE",
        "TOPEKA, KS",
        "PH:",
        "FX:",
        "PICKUP TOTAL",
        "DROP TOTAL",
        "OUTBOUND TOTALS",
        "TOTAL MILES",
        "PAGE ",
        "ORDER#",
        "WEIGHT",
        "CUSTID",
        "CUST",
        "NAME",
        "LOCATION",
        "DATE",
        "TIME",
    ]

    upper_value = value.upper()

    for bad in bad_fragments:
        if bad in upper_value:
            return ""

    value = re.sub(r"/\s*\d+[A-Z0-9\-\.]*$", "", value)
    value = re.sub(r"^/\s*", "", value)
    value = clean_spaces(value)

    if re.fullmatch(r"[\d\.\,\-]+", value):
        return ""

    if len(re.sub(r"[^A-Za-z]", "", value)) < 2:
        return ""

    return value


def extract_customers_from_loading_manifest(text: str) -> list[str]:
    customers = []

    for raw_line in text.splitlines():
        line = clean_spaces(raw_line)

        match = re.match(r"^[2-9]\s+(.+)$", line)
        if not match:
            continue

        candidate = clean_customer_name(match.group(1))

        if not candidate:
            continue

        if candidate not in customers:
            customers.append(candidate)

    return customers


def get_shipping_pickup_section(text: str) -> str:
    if "SHIPPING MANIFEST" not in text.upper():
        return ""

    if "Pickup TOTAL" in text:
        return text.split("Pickup TOTAL", 1)[0]

    if "PICKUP TOTAL" in text.upper():
        return re.split(r"PICKUP\s+TOTAL", text, flags=re.IGNORECASE)[0]

    return text


def flush_customer_parts(parts, customers):
    cleaned_parts = []

    for part in parts:
        part = clean_customer_name(part)

        if not part:
            continue

        if part.startswith("/"):
            continue

        cleaned_parts.append(part)

    if not cleaned_parts:
        return

    customer = clean_spaces(" ".join(cleaned_parts))

    customer = customer.replace("SPRINGFIEL D", "SPRINGFIELD")
    customer = customer.replace("GWILLIMBUR Y", "GWILLIMBURY")
    customer = customer.replace("DISTRIBUTO RS", "DISTRIBUTORS")
    customer = customer.replace("PENNSYLVAN IA", "PENNSYLVANIA")
    customer = customer.replace("CONNECTICU T", "CONNECTICUT")
    customer = customer.replace("SOWESTERN ONT", "SW ONTARIO")
    customer = customer.replace("SOUTHWESTERN ONT", "SW ONTARIO")
    customer = clean_spaces(customer)

    if customer and customer not in customers:
        customers.append(customer)


def extract_customers_from_shipping_manifest(text: str) -> list[str]:
    pickup_text = get_shipping_pickup_section(text)

    if not pickup_text:
        return []

    lines = [
        clean_spaces(line)
        for line in pickup_text.splitlines()
        if clean_spaces(line)
    ]

    customers = []
    current_parts = []
    in_order = False

    stop_words = [
        "RESERS FINE FOOD",
        "RESER'S - TK",
        "SHIPPING MANIFEST",
        "PRINTED BY",
        "LOAD:",
        "PU APPT",
        "CARR/SCT",
        "DRIVER",
        "TRACTOR",
        "PRO:",
        "LOCATION",
        "ORDER#",
        "WEIGHT",
        "VOL",
        "PCS",
        "PLT",
        "CUSTID",
        "CUST",
        "NAME",
        "TOPEKA",
        "3121 SE",
        "PH:",
        "APPT:",
        "PICKUP TOTAL",
    ]

    for line in lines:
        upper_line = line.upper()

        if re.match(r"^\d{8,12}-\d{3}", line):
            if current_parts:
                flush_customer_parts(current_parts, customers)

            current_parts = []
            in_order = True

            slash_match = re.search(r"\b[A-Z0-9]{1,10}/\s*(.+)$", line)
            if slash_match:
                possible = clean_customer_name(slash_match.group(1))
                if possible:
                    current_parts.append(possible)

            continue

        if not in_order:
            continue

        if any(word in upper_line for word in stop_words):
            if current_parts:
                flush_customer_parts(current_parts, customers)
            current_parts = []
            in_order = False
            continue

        if line.startswith("/"):
            if current_parts:
                flush_customer_parts(current_parts, customers)
            current_parts = []
            in_order = False
            continue

        if re.fullmatch(r"[\d\.\,\-]+", line):
            continue

        if re.fullmatch(r"\d{2,12}/?", line):
            continue

        if re.match(r"^\d{2,12}/", line):
            after_slash = line.split("/", 1)[1].strip()
            possible = clean_customer_name(after_slash)
            if possible:
                current_parts.append(possible)
            continue

        possible = clean_customer_name(line)

        if possible:
            current_parts.append(possible)

    if current_parts:
        flush_customer_parts(current_parts, customers)

    final_customers = []

    for customer in customers:
        customer = clean_spaces(customer)

        if not customer:
            continue

        if len(customer) > 70:
            words = customer.split()
            customer = " ".join(words[:8])

        if customer not in final_customers:
            final_customers.append(customer)

    return final_customers


def is_internal_transfer_customer(customer: str) -> bool:
    upper_customer = customer.upper()

    internal_terms = [
        "RESER'S -",
        "RESERS -",
        "DISTRIBUTION CENTER",
        "CENTURY DISTRIBUTION",
        "HALIFAX DISTRIBUTION",
        "REED TRUCKING",
        "BOZEL TRANSFER",
        "REET -",
        "BOZL -",
        "NC RESER",
        "DC RESER",
    ]

    return any(term in upper_customer for term in internal_terms)


def extract_mg_customers(loading_text: str, shipping_text: str) -> str:
    shipping_customers = extract_customers_from_shipping_manifest(shipping_text)
    loading_customers = extract_customers_from_loading_manifest(loading_text)

    if shipping_customers:
        selected = shipping_customers
    else:
        selected = loading_customers

    useful_loading = [
        c for c in loading_customers
        if not is_internal_transfer_customer(c)
    ]

    if useful_loading and not shipping_customers:
        selected = useful_loading

    final_customers = []

    for customer in selected:
        customer = clean_customer_name(customer)

        if not customer:
            continue

        if customer not in final_customers:
            final_customers.append(customer)

    return " / ".join(final_customers)


def parse_manifest_pdf(pdf_bytes: bytes) -> tuple[list[dict], list[dict]]:
    pages = extract_pdf_pages(pdf_bytes)

    pages_by_load = defaultdict(list)

    for page_text in pages:
        load_number = extract_mg_load_number(page_text)

        if load_number:
            pages_by_load[load_number].append(page_text)

    records = []
    duplicate_pdf_records = []
    seen_loads = set()

    for load_number, load_pages in pages_by_load.items():
        clean_load = normalize_load_number(load_number)

        if clean_load in seen_loads:
            duplicate_pdf_records.append({
                "Load": clean_load,
                "Reason": "Duplicate inside uploaded PDF"
            })
            continue

        seen_loads.add(clean_load)

        loading_pages = [
            p for p in load_pages
            if "LOADING MANIFEST" in p.upper()
        ]

        shipping_pages = [
            p for p in load_pages
            if "SHIPPING MANIFEST" in p.upper()
        ]

        if loading_pages:
            loading_text = "\n".join(loading_pages)
        else:
            loading_text = "\n".join(load_pages)

        shipping_text = "\n".join(shipping_pages)

        appt_date, appt_time = extract_mg_pu_appt(loading_text)
        carrier = extract_mg_carrier(loading_text)
        actual_weight, actual_quantity = extract_mg_totals(loading_text)
        customer = extract_mg_customers(loading_text, shipping_text)

        records.append({
            "load": clean_load,
            "appt_date": appt_date,
            "appt_time": appt_time,
            "carrier": carrier,
            "actual_weight": actual_weight,
            "actual_quantity": actual_quantity,
            "customer": customer,
        })

    records.sort(key=lambda x: (sort_datetime_key(x), x.get("load", "")))

    return records, duplicate_pdf_records


# ============================================================
# TEMPLATE POPULATION
# ============================================================

def validate_template(wb):
    missing = []

    for sheet_name in [OPENDOCK_SHEET, MG_REPORT_SHEET]:
        if sheet_name not in wb.sheetnames:
            missing.append(sheet_name)

    if missing:
        raise ValueError("Template is missing these sheets: " + ", ".join(missing))


def populate_opendock_sheet(wb, opendock_records):
    ws = wb[OPENDOCK_SHEET]
    headers = get_header_map(ws, 1)

    load_col = find_col(headers, ["Load Reference", "Load", "Load #"], 1)
    appt_date_col = find_col(headers, ["Appt Date", "Appointment Date", "Date"], 2)
    appt_time_col = find_col(headers, ["Appt Time", "Appointment Time", "Time"], 3)
    arrival_date_col = find_col(headers, ["Arrival Date"], 4)
    arrival_time_col = find_col(headers, ["Arrival Time"], 5)
    departure_date_col = find_col(headers, ["Departure Date"], 6)
    departure_time_col = find_col(headers, ["Departure Time"], 7)
    status_col = find_col(headers, ["Status"], 8)
    carrier_col = find_col(headers, ["Carrier Company", "Carrier"], 9)
    load_type_col = find_col(headers, ["Load Type", "Type"], 10)
    dock_col = find_col(headers, ["Dock", "Door"], 11)

    output_cols = [
        load_col,
        appt_date_col,
        appt_time_col,
        arrival_date_col,
        arrival_time_col,
        departure_date_col,
        departure_time_col,
        status_col,
        carrier_col,
        load_type_col,
        dock_col,
    ]

    output_cols = sorted(set([col for col in output_cols if col]))

    max_row = max(ws.max_row, OPENDOCK_START_ROW + MAX_OPENDOCK_ROWS)

    clear_range_values(ws, OPENDOCK_START_ROW, max_row, output_cols)

    for index, record in enumerate(opendock_records, start=OPENDOCK_START_ROW):
        copy_row_format(ws, OPENDOCK_START_ROW, index, max(output_cols))

        ws.cell(index, load_col).value = record.get("load", "")
        ws.cell(index, appt_date_col).value = record.get("appt_date", "")
        ws.cell(index, appt_time_col).value = record.get("appt_time", "")
        ws.cell(index, arrival_date_col).value = record.get("arrival_date", "")
        ws.cell(index, arrival_time_col).value = record.get("arrival_time", "")
        ws.cell(index, departure_date_col).value = record.get("departure_date", "")
        ws.cell(index, departure_time_col).value = record.get("departure_time", "")
        ws.cell(index, status_col).value = record.get("status", "")
        ws.cell(index, carrier_col).value = record.get("carrier", "")
        ws.cell(index, load_type_col).value = record.get("load_type", "")
        ws.cell(index, dock_col).value = record.get("dock", "")


def populate_mg_report_sheet(wb, mg_records):
    ws = wb[MG_REPORT_SHEET]
    headers = get_header_map(ws, 1)

    load_col = find_col(headers, ["Load #", "Load", "Load Number"], 1)
    weight_col = find_col(headers, ["Actual Weight", "Weight"], 2)
    quantity_col = find_col(headers, ["Actual Quantity", "Cases", "Quantity"], 3)
    customer_col = find_col(headers, ["Consignee Name", "Customer", "Customer(s)", "Destination"], 4)
    appt_date_col = find_col(headers, ["Appt Date", "Appointment Date", "Pickup Date"], None)
    appt_time_col = find_col(headers, ["Appt Time", "Appointment Time", "Pickup Time"], None)
    carrier_col = find_col(headers, ["Carrier", "Carrier Company"], None)

    output_cols = [
        load_col,
        weight_col,
        quantity_col,
        customer_col,
        appt_date_col,
        appt_time_col,
        carrier_col,
    ]

    output_cols = sorted(set([col for col in output_cols if col]))

    max_row = max(ws.max_row, MG_REPORT_START_ROW + MAX_MG_REPORT_ROWS)

    clear_range_values(ws, MG_REPORT_START_ROW, max_row, output_cols)

    for index, record in enumerate(mg_records, start=MG_REPORT_START_ROW):
        copy_row_format(ws, MG_REPORT_START_ROW, index, max(output_cols))

        ws.cell(index, load_col).value = record.get("load", "")
        ws.cell(index, weight_col).value = record.get("actual_weight", 0)
        ws.cell(index, quantity_col).value = record.get("actual_quantity", 0)
        ws.cell(index, customer_col).value = record.get("customer", "")

        if appt_date_col:
            ws.cell(index, appt_date_col).value = record.get("appt_date", "")

        if appt_time_col:
            ws.cell(index, appt_time_col).value = record.get("appt_time", "")

        if carrier_col:
            ws.cell(index, carrier_col).value = record.get("carrier", "")


def write_dispatch_load_count(wb, opendock_records):
    if DISPATCH_SHEET not in wb.sheetnames:
        return

    ws = wb[DISPATCH_SHEET]
    ws["G1"] = len(opendock_records)


def add_match_report_sheet(wb, opendock_records, mg_records):
    sheet_name = "MATCH REPORT"

    if sheet_name in wb.sheetnames:
        del wb[sheet_name]

    ws = wb.create_sheet(sheet_name)

    headers = [
        "Load",
        "Opendock Date",
        "Opendock Time",
        "MG Date",
        "MG Time",
        "Opendock Carrier",
        "MG Carrier",
        "Opendock Load Type",
        "MG Customer(s)",
        "MG Weight",
        "MG Cases",
        "Status",
    ]

    ws.append(headers)

    mg_by_load = {
        normalize_load_number(record.get("load")): record
        for record in mg_records
        if normalize_load_number(record.get("load"))
    }

    opendock_loads = set()

    for record in opendock_records:
        load_number = record.get("load", "")
        opendock_loads.add(load_number)

        mg = mg_by_load.get(load_number)

        if not mg:
            ws.append([
                load_number,
                record.get("appt_date", ""),
                record.get("appt_time", ""),
                "",
                "",
                record.get("carrier", ""),
                "",
                record.get("load_type", ""),
                "",
                0,
                0,
                "Missing in MG report",
            ])
            continue

        status_parts = []

        if mg.get("appt_date") and record.get("appt_date") and mg.get("appt_date") != record.get("appt_date"):
            status_parts.append("Date mismatch")

        if mg.get("appt_time") and record.get("appt_time") and mg.get("appt_time") != record.get("appt_time"):
            status_parts.append("Time mismatch")

        if not mg.get("customer"):
            status_parts.append("Customer missing from MG")

        status = "Matched" if not status_parts else " / ".join(status_parts)

        ws.append([
            load_number,
            record.get("appt_date", ""),
            record.get("appt_time", ""),
            mg.get("appt_date", ""),
            mg.get("appt_time", ""),
            record.get("carrier", ""),
            mg.get("carrier", ""),
            record.get("load_type", ""),
            mg.get("customer", ""),
            mg.get("actual_weight", 0),
            mg.get("actual_quantity", 0),
            status,
        ])

    for mg_load, mg in mg_by_load.items():
        if mg_load not in opendock_loads:
            ws.append([
                mg_load,
                "",
                "",
                mg.get("appt_date", ""),
                mg.get("appt_time", ""),
                "",
                mg.get("carrier", ""),
                "",
                mg.get("customer", ""),
                mg.get("actual_weight", 0),
                mg.get("actual_quantity", 0),
                "Missing in Opendock outbound report",
            ])

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    for column in range(1, ws.max_column + 1):
        ws.column_dimensions[get_column_letter(column)].width = 20

    ws.column_dimensions["I"].width = 55
    ws.column_dimensions["L"].width = 35

    ws.freeze_panes = "A2"


def format_output_workbook(wb):
    for sheet_name in [OPENDOCK_SHEET, MG_REPORT_SHEET]:
        if sheet_name not in wb.sheetnames:
            continue

        ws = wb[sheet_name]

        for col in range(1, ws.max_column + 1):
            letter = get_column_letter(col)

            if sheet_name == MG_REPORT_SHEET and letter in ["D", "E"]:
                ws.column_dimensions[letter].width = 55
            else:
                ws.column_dimensions[letter].width = 18

        try:
            ws.freeze_panes = "A2"
        except Exception:
            pass


def populate_template(template_bytes, opendock_records, mg_records):
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))

    validate_template(wb)

    populate_opendock_sheet(wb, opendock_records)
    populate_mg_report_sheet(wb, mg_records)
    write_dispatch_load_count(wb, opendock_records)
    add_match_report_sheet(wb, opendock_records, mg_records)
    format_output_workbook(wb)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return output.getvalue()


# ============================================================
# STREAMLIT APP
# ============================================================

st.set_page_config(
    page_title="MG + Opendock Worksheet Populator",
    layout="wide"
)

st.title("MG + Opendock Worksheet Populator")

st.write(
    "Upload the MG Loading/Shipping reports, the Excel template, and the Opendock report. "
    "The app will populate only the OPENDOCK and MG REPORT worksheets. "
    "The DISPATCH SHEET formulas will do the rest."
)

st.subheader("Step 1 - Upload all inputs")

mg_files = st.file_uploader(
    "Upload MG Loading Manifest and Shipping Manifest. PDFs or ZIPs are accepted.",
    type=["pdf", "zip"],
    accept_multiple_files=True
)

excel_template_file = st.file_uploader(
    "Upload Excel template with OPENDOCK, MG REPORT, and DISPATCH SHEET",
    type=["xlsx"]
)

opendock_file = st.file_uploader(
    "Upload Opendock report",
    type=["xlsx"]
)

if mg_files and excel_template_file and opendock_file:
    if st.button("Build populated workbook"):
        try:
            with st.spinner("Building matched manifest packet from MG reports..."):
                matched_pdf_bytes, matched_summary = build_matched_packet(mg_files)

            st.success("Matched manifest packet created successfully.")

            col1, col2, col3 = st.columns(3)

            col1.metric("Loading loads found", matched_summary["Loading loads found"])
            col2.metric("Shipping loads found", matched_summary["Shipping loads found"])
            col3.metric("Unique loads matched", matched_summary["Unique loads matched"])

            st.download_button(
                label="Download matched manifest packet",
                data=matched_pdf_bytes,
                file_name=MATCHED_PDF_FILE_NAME,
                mime="application/pdf"
            )

            with st.spinner("Reading Opendock report and keeping only Outbound rows..."):
                opendock_records, opendock_summary = parse_opendock_excel(
                    opendock_file.read()
                )

            with st.spinner("Parsing matched MG packet..."):
                mg_records, duplicate_pdf_records = parse_manifest_pdf(
                    matched_pdf_bytes
                )

            st.success("Source files parsed successfully.")

            c1, c2, c3, c4 = st.columns(4)

            c1.metric("Opendock Outbound Rows", opendock_summary["outbound_rows_loaded"])
            c2.metric("Inbound Rows Skipped", opendock_summary["inbound_rows_skipped"])
            c3.metric("MG Loads Found", len(mg_records))
            c4.metric("Duplicate MG Loads", len(duplicate_pdf_records))

            if opendock_summary["duplicate_outbound_loads"]:
                st.warning(
                    "Duplicate outbound loads in Opendock: "
                    + ", ".join(opendock_summary["duplicate_outbound_loads"])
                )

            if duplicate_pdf_records:
                st.warning("Duplicate loads found inside MG PDF.")
                st.dataframe(duplicate_pdf_records, use_container_width=True)

            with st.spinner("Populating OPENDOCK and MG REPORT worksheets..."):
                populated_workbook = populate_template(
                    excel_template_file.read(),
                    opendock_records,
                    mg_records
                )

            st.success("Workbook populated successfully.")

            st.subheader("Opendock worksheet preview")
            st.dataframe(opendock_records, use_container_width=True)

            st.subheader("MG Report worksheet preview")
            st.dataframe(mg_records, use_container_width=True)

            missing_customer_records = [
                record for record in mg_records
                if not record.get("customer", "")
            ]

            if missing_customer_records:
                st.warning(f"{len(missing_customer_records)} MG loads have no customer detected.")
                st.dataframe(missing_customer_records, use_container_width=True)
            else:
                st.success("Customer populated for all MG loads.")

            st.download_button(
                label="Download populated workbook",
                data=populated_workbook,
                file_name=OUTPUT_FILE_NAME,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error("Something went wrong while processing the files.")
            st.exception(e)

else:
    st.info("Upload the MG reports, Excel template, and Opendock report to start.")
