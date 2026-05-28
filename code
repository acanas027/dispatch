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
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from pypdf import PdfReader, PdfWriter


# ============================================================
# SETTINGS
# ============================================================

SHEET_NAME = "Outbound"
START_ROW = 4

# The code only uses A:D to detect the actual board table.
TABLE_MIN_COL = 1
TABLE_MAX_COL = 4

# Only copy/style the real board columns.
# A:N = 1:14. This prevents Python from touching S:T.
BOARD_MIN_COL = 1
BOARD_MAX_COL = 14

# Clear generated values from S:T so TT4 only shows in column I.
CLEAR_EXTRA_OUTPUT_COLS = [19, 20]  # S, T

# Template rows.
# Row 4 = day row format.
# Row 5 = load row format/formulas.
TEMPLATE_DAY_ROW = START_ROW
TEMPLATE_LOAD_ROW = START_ROW + 1

# Default board columns
COL_LOAD_OR_DAY = 1       # A
COL_CUSTOMER = 2          # B
COL_CARRIER_DEFAULT = 3   # C
COL_TIME = 4              # D
COL_CASES_DEFAULT = 6     # F
COL_TT4 = 9               # I

# Appointment Excel headers
APPT_LOAD_REF_HEADER = "Load Reference"
APPT_LOAD_TYPE_HEADER = "Load Type"
APPT_CARRIER_HEADER = "Carrier Company"

APPT_DATE_HEADERS = [
    "Appointment Date",
    "Appt Date",
    "Pickup Date",
    "PU Date",
    "Date",
    "Start Date",
]

APPT_TIME_HEADERS = [
    "Appointment Time",
    "Appt Time",
    "Pickup Time",
    "PU Time",
    "Time",
    "Start Time",
]

APPT_DATETIME_HEADERS = [
    "Appointment Date Time",
    "Appointment Datetime",
    "Appt Date Time",
    "Appt Datetime",
    "Pickup Date Time",
    "PU Date Time",
    "Scheduled Time",
    "Scheduled Date Time",
    "Start Date Time",
]

TT4_FOR_CANADA = True

TT4_CUSTOMER_KEYWORDS = [
    # "LOBLAWS",
    # "SOBEYS",
    # "COSTCO",
]

OUTPUT_FILE_NAME = "Updated_EXCEL_BOARD.xlsx"
MATCHED_PDF_FILE_NAME = "Matched_Manifest_Packet.pdf"


# ============================================================
# BASIC CLEANING
# ============================================================

def clean_spaces(value: str) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def parse_number(value):
    if value is None:
        return 0

    value = str(value).replace(",", "").strip()

    if not value:
        return 0

    try:
        number = float(value)
        if number.is_integer():
            return int(number)
        return number
    except ValueError:
        return 0


def normalize_load_number(value) -> str:
    if value is None:
        return ""

    text = str(value).strip()
    text = re.sub(r"\.0$", "", text)
    text = re.sub(r"^LD", "", text, flags=re.IGNORECASE)
    text = re.sub(r"[^0-9]", "", text)

    if text.isdigit():
        return text

    return ""


def normalize_header(value) -> str:
    if value is None:
        return ""

    return re.sub(r"[^A-Z0-9]", "", str(value).strip().upper())


def parse_sort_datetime(value: str):
    try:
        return datetime.strptime(value, "%m/%d/%Y %H:%M")
    except Exception:
        return datetime.max


def parse_date_any(value: str) -> str:
    if not value:
        return ""

    text = str(value).strip()

    for fmt in ["%m/%d/%Y", "%m/%d/%y"]:
        try:
            dt = datetime.strptime(text, fmt)
            return dt.strftime("%m/%d/%Y")
        except Exception:
            pass

    return ""


def format_date_short(date_text: str) -> str:
    try:
        dt = datetime.strptime(date_text, "%m/%d/%Y")
        return dt.strftime("%m/%d/%y")
    except Exception:
        return date_text


def get_day_name(date_text: str) -> str:
    try:
        dt = datetime.strptime(date_text, "%m/%d/%Y")
        return dt.strftime("%A")
    except Exception:
        return ""


def format_time_only(time_text: str) -> str:
    if not time_text:
        return ""

    text = str(time_text).strip()

    match = re.search(r"(\d{1,2}):(\d{2})", text)
    if match:
        return f"{int(match.group(1)):02d}:{match.group(2)}"

    return text


def build_datetime_from_date_time(date_text: str, time_text: str) -> str:
    date_text = parse_date_any(date_text) or date_text
    time_text = format_time_only(time_text)

    if not date_text:
        return ""

    if not time_text:
        time_text = "23:59"

    return f"{date_text} {time_text}"


def normalize_date_for_match(value) -> str:
    if value is None:
        return ""

    if isinstance(value, datetime):
        return value.strftime("%m/%d/%Y")

    if isinstance(value, date):
        return value.strftime("%m/%d/%Y")

    text = clean_spaces(value)

    if not text:
        return ""

    date_match = re.search(r"(\d{1,2}/\d{1,2}/\d{2,4})", text)
    if date_match:
        text = date_match.group(1)

    for fmt in ["%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d"]:
        try:
            dt = datetime.strptime(text, fmt)
            return dt.strftime("%m/%d/%Y")
        except Exception:
            pass

    return ""


def normalize_time_for_match(value) -> str:
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

    time_match = re.search(r"\b(\d{1,2}):(\d{2})\b", text)
    if time_match and "AM" not in text.upper() and "PM" not in text.upper():
        hour = int(time_match.group(1))
        minute = int(time_match.group(2))
        return f"{hour:02d}:{minute:02d}"

    compact_text = text.upper().replace(" ", "")

    try:
        dt = datetime.strptime(compact_text, "%I:%M%p")
        return dt.strftime("%H:%M")
    except Exception:
        pass

    return ""


def get_record_match_key(record: dict) -> tuple[str, str, str]:
    load_number = normalize_load_number(record.get("Load", ""))
    pickup_date = normalize_date_for_match(record.get("Pickup Date", ""))
    pickup_time = normalize_time_for_match(record.get("Pickup Time", ""))

    return load_number, pickup_date, pickup_time


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
    """
    Takes the two MG report uploads, identifies loading and shipping manifests,
    matches them by load number, and returns a matched PDF packet as bytes.
    """
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
# APPOINTMENT EXCEL PARSER
# ============================================================

def find_header_column(ws, header_name: str, header_row: int = 1) -> int | None:
    target = normalize_header(header_name)

    for col in range(1, ws.max_column + 1):
        value = ws.cell(header_row, col).value

        if normalize_header(value) == target:
            return col

    return None


def find_first_header_column(ws, possible_headers: list[str], header_row: int = 1) -> int | None:
    for header in possible_headers:
        col = find_header_column(ws, header, header_row)
        if col:
            return col
    return None


def parse_appointments_excel(appointments_bytes: bytes) -> dict:
    wb = openpyxl.load_workbook(io.BytesIO(appointments_bytes), data_only=True)

    if "Appointments" in wb.sheetnames:
        ws = wb["Appointments"]
    else:
        ws = wb[wb.sheetnames[0]]

    load_ref_col = find_header_column(ws, APPT_LOAD_REF_HEADER)
    load_type_col = find_header_column(ws, APPT_LOAD_TYPE_HEADER)
    carrier_col = find_header_column(ws, APPT_CARRIER_HEADER)

    datetime_col = find_first_header_column(ws, APPT_DATETIME_HEADERS)
    date_col = find_first_header_column(ws, APPT_DATE_HEADERS)
    time_col = find_first_header_column(ws, APPT_TIME_HEADERS)

    missing = []

    if not load_ref_col:
        missing.append(APPT_LOAD_REF_HEADER)

    if not load_type_col:
        missing.append(APPT_LOAD_TYPE_HEADER)

    if not carrier_col:
        missing.append(APPT_CARRIER_HEADER)

    if not datetime_col and not date_col:
        missing.append("Appointment Date / Pickup Date")

    if not datetime_col and not time_col:
        missing.append("Appointment Time / Pickup Time")

    if missing:
        raise ValueError(
            "The appointments Excel is missing these headers: "
            + ", ".join(missing)
        )

    lookup = defaultdict(list)

    for row in range(2, ws.max_row + 1):
        load_ref = ws.cell(row, load_ref_col).value
        load_number = normalize_load_number(load_ref)

        if not load_number:
            continue

        load_type = clean_spaces(ws.cell(row, load_type_col).value)
        carrier = clean_spaces(ws.cell(row, carrier_col).value)

        if datetime_col:
            datetime_value = ws.cell(row, datetime_col).value
            pickup_date = normalize_date_for_match(datetime_value)
            pickup_time = normalize_time_for_match(datetime_value)
        else:
            pickup_date = normalize_date_for_match(ws.cell(row, date_col).value)
            pickup_time = normalize_time_for_match(ws.cell(row, time_col).value)

        lookup[load_number].append({
            "Load Type": load_type,
            "Carrier": carrier,
            "Pickup Date": pickup_date,
            "Pickup Time": pickup_time,
            "Raw Load Reference": clean_spaces(load_ref),
        })

    return dict(lookup)


def build_records_from_appointments(pdf_records: list[dict], appointment_lookup: dict) -> tuple[list[dict], list[dict]]:
    report_match_issues = []
    output_records = []

    pdf_by_load = defaultdict(list)

    for pdf_record in pdf_records:
        pdf_load = normalize_load_number(pdf_record.get("Load", ""))

        if pdf_load:
            pdf_by_load[pdf_load].append(pdf_record)

    for appointment_load, appointment_rows in appointment_lookup.items():
        for appointment in appointment_rows:
            appt_date = normalize_date_for_match(appointment.get("Pickup Date", ""))
            appt_time = normalize_time_for_match(appointment.get("Pickup Time", ""))

            pdf_matches_for_load = pdf_by_load.get(appointment_load, [])

            exact_pdf_match = None

            for pdf_record in pdf_matches_for_load:
                _, pdf_date, pdf_time = get_record_match_key(pdf_record)

                if pdf_date == appt_date and pdf_time == appt_time:
                    exact_pdf_match = pdf_record
                    break

            if exact_pdf_match:
                record = exact_pdf_match.copy()

                record["Load"] = appointment_load
                record["Pickup Date"] = appt_date
                record["Pickup Time"] = appt_time
                record["Pickup DateTime"] = f"{appt_date} {appt_time}"
                record["Load Type"] = appointment.get("Load Type", "")
                record["Carrier"] = appointment.get("Carrier", "")
                record["Cases"] = parse_number(exact_pdf_match.get("Cases", 0))
                record["Appointment Match"] = "Yes"
                record["Report Match Status"] = "Matched"
                record["Source"] = "New"

                output_records.append(record)

            elif pdf_matches_for_load:
                first_pdf_record = pdf_matches_for_load[0].copy()

                pdf_dates = []
                pdf_times = []
                pdf_cases_total = 0

                for pdf_record in pdf_matches_for_load:
                    _, pdf_date, pdf_time = get_record_match_key(pdf_record)
                    pdf_dates.append(pdf_date)
                    pdf_times.append(pdf_time)
                    pdf_cases_total += parse_number(pdf_record.get("Cases", 0))

                record = first_pdf_record.copy()

                record["Load"] = appointment_load
                record["Pickup Date"] = appt_date
                record["Pickup Time"] = appt_time
                record["Pickup DateTime"] = f"{appt_date} {appt_time}"
                record["Load Type"] = appointment.get("Load Type", "")
                record["Carrier"] = appointment.get("Carrier", "")
                record["Cases"] = pdf_cases_total
                record["Appointment Match"] = "Date/Time Mismatch"
                record["Report Match Status"] = "Load found in PDF, but date/time did not match"
                record["Source"] = "New"

                output_records.append(record)

                report_match_issues.append({
                    "Load": appointment_load,
                    "Appointment Date": appt_date,
                    "Appointment Time": appt_time,
                    "PDF Date": " / ".join(pdf_dates),
                    "PDF Time": " / ".join(pdf_times),
                    "PDF Cases Counted": pdf_cases_total,
                    "Reason": "Load number found in PDF, but date/time does not match"
                })

            else:
                record = {
                    "Load": appointment_load,
                    "Customer": "",
                    "Carrier": appointment.get("Carrier", ""),
                    "Load Type": appointment.get("Load Type", ""),
                    "Pickup DateTime": f"{appt_date} {appt_time}",
                    "Pickup Date": appt_date,
                    "Pickup Time": appt_time,
                    "Cases": 0,
                    "Is Canada": False,
                    "Raw Text": "",
                    "TT4": "",
                    "Appointment Match": "No PDF Match",
                    "Report Match Status": "Load not found in PDF",
                    "Source": "New",
                }

                output_records.append(record)

                report_match_issues.append({
                    "Load": appointment_load,
                    "Appointment Date": appt_date,
                    "Appointment Time": appt_time,
                    "PDF Date": "",
                    "PDF Time": "",
                    "PDF Cases Counted": 0,
                    "Reason": "Load number was found in appointments Excel but not found in PDF"
                })

    output_records.sort(
        key=lambda x: (
            parse_sort_datetime(x.get("Pickup DateTime", "")),
            x.get("Load", "")
        )
    )

    return output_records, report_match_issues


# ============================================================
# BOARD COLUMN RESOLUTION
# ============================================================

def find_board_header_column(ws, accepted_headers: list[str], header_rows=(1, 2, 3)) -> int | None:
    accepted = {normalize_header(x) for x in accepted_headers}

    for row in header_rows:
        for col in range(1, ws.max_column + 1):
            value = ws.cell(row, col).value
            normalized = normalize_header(value)

            if normalized in accepted:
                return col

    return None


def resolve_board_columns(ws) -> dict:
    load_col = find_board_header_column(
        ws,
        ["LOAD #", "LOAD", "LOAD NUMBER"]
    ) or COL_LOAD_OR_DAY

    customer_col = find_board_header_column(
        ws,
        ["CUSTOMER", "DESTINATION"]
    ) or COL_CUSTOMER

    carrier_col = find_board_header_column(
        ws,
        ["CARRIER", "CARRIER COMPANY"]
    )

    type_col = find_board_header_column(
        ws,
        ["TYPE", "LOAD TYPE"]
    )

    carrier_type_col = find_board_header_column(
        ws,
        ["CARRIER-TYPE", "CARRIER TYPE", "CARRIERTYPE"]
    )

    time_col = find_board_header_column(
        ws,
        ["TIME", "APPT TIME", "APPOINTMENT TIME"]
    ) or COL_TIME

    tt4_col = find_board_header_column(
        ws,
        ["TT4"]
    ) or COL_TT4

    if not carrier_col:
        carrier_col = carrier_type_col or COL_CARRIER_DEFAULT

    return {
        "load_col": load_col,
        "customer_col": customer_col,
        "carrier_col": carrier_col,
        "type_col": type_col,
        "carrier_type_col": carrier_type_col,
        "time_col": time_col,
        "tt4_col": tt4_col,
    }


def get_carrier_output_value(record: dict, board_columns: dict) -> str:
    carrier = clean_spaces(record.get("Carrier", ""))
    load_type = clean_spaces(record.get("Load Type", ""))

    type_col = board_columns.get("type_col")
    carrier_type_col = board_columns.get("carrier_type_col")
    carrier_col = board_columns.get("carrier_col")

    if type_col and carrier_col != type_col:
        return carrier

    if carrier_type_col and not type_col:
        if load_type and carrier:
            return f"{load_type} - {carrier}"
        if carrier:
            return carrier
        return load_type

    if carrier:
        return carrier

    return load_type


# ============================================================
# PDF PARSER FOR THE MATCHED PACKET
# ============================================================

def extract_pdf_text_by_page(pdf_bytes: bytes) -> list[str]:
    pages = []

    with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
        for page in doc:
            pages.append(page.get_text("text"))

    return pages


def extract_load_number(text: str) -> str:
    match = re.search(r"\bLOAD:\s*(?:LD)?\s*([A-Z0-9]+)", text, re.IGNORECASE)

    if not match:
        return ""

    return normalize_load_number(match.group(1))


def extract_pickup_datetime(text: str) -> tuple[str, str, str]:
    match = re.search(
        r"\bPU\s+APPT:\s*(\d{1,2}/\d{1,2}/\d{4})\s+(\d{1,2}:\d{2})",
        text,
        re.IGNORECASE,
    )

    if not match:
        return "", "", ""

    date_text = match.group(1)
    time_text = match.group(2)
    full_datetime = f"{date_text} {time_text}"

    return full_datetime, date_text, time_text


def clean_carrier(raw_carrier: str) -> str:
    carrier = clean_spaces(raw_carrier)

    if not carrier:
        return ""

    upper_carrier = carrier.upper()

    if upper_carrier == "CPUC":
        return "CPU"

    if upper_carrier.startswith("CPUC"):
        return "CPU"

    match = re.match(r"^[A-Z0-9]{3,6}\s+(.+)$", carrier)

    if match:
        carrier = match.group(1).strip()

    carrier = clean_spaces(carrier)

    if carrier.upper().startswith("CPUC"):
        return "CPU"

    return carrier


def extract_carrier(text: str) -> str:
    match = re.search(r"\bCARR/SCT\s+TR:\s*(.+)", text, re.IGNORECASE)

    if not match:
        return ""

    raw_carrier = match.group(1).splitlines()[0]

    return clean_carrier(raw_carrier)


def extract_cases(text: str):
    matches = re.findall(
        r"OUTBOUND\s+TOTALS:\s+([\d,]+(?:\.\d+)?)\s+([\d,]+(?:\.\d+)?)\s+([\d,]+(?:\.\d+)?)",
        text,
        re.IGNORECASE,
    )

    if not matches:
        return 0

    last_match = matches[-1]
    return parse_number(last_match[2])


# ============================================================
# CUSTOMER EXTRACTION
# ============================================================

def clean_customer_destination(value: str) -> str:
    value = clean_spaces(value)
    value = re.sub(r"^\d{5,12}\s+", "", value)
    value = re.sub(r"\s+TK$", "", value, flags=re.IGNORECASE)
    value = re.sub(r"\s+T$", "", value, flags=re.IGNORECASE)
    value = value.replace("DISTRIBUTION CENTER", "")
    value = value.replace("WAREHOUSE", "")
    value = clean_spaces(value)
    return value


def extract_stop_customer_lines_from_loading_manifest(text: str) -> list[str]:
    customers = []

    for raw_line in text.splitlines():
        line = clean_spaces(raw_line)

        match = re.match(r"^[2-9]\s+(.+)$", line)
        if not match:
            continue

        candidate = match.group(1).strip()
        candidate_upper = candidate.upper()

        if candidate_upper.startswith(("TK ", "DC ", "SE ")):
            continue

        bad_words = [
            "DROP TOTAL",
            "OUTBOUND TOTALS",
            "SUGGESTED DELIVERY",
            "RESER'S - TK - TOPEKA",
            "PAGE ",
        ]

        if any(word in candidate_upper for word in bad_words):
            continue

        candidate = clean_customer_destination(candidate)

        if candidate:
            customers.append(candidate)

    return customers


def get_pickup_section_from_shipping_manifest(text: str) -> str:
    if "SHIPPING MANIFEST" not in text.upper():
        return ""

    parts = re.split(r"Pickup\s+TOTAL:", text, flags=re.IGNORECASE)

    if len(parts) < 2:
        return text

    return parts[0]


def extract_order_customer_lines_from_shipping_manifest(text: str) -> list[str]:
    pickup_text = get_pickup_section_from_shipping_manifest(text)

    if not pickup_text:
        return []

    lines = [clean_spaces(x) for x in pickup_text.splitlines() if clean_spaces(x)]
    customers = []

    stop_words = [
        "RESERS FINE FOOD",
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
        "CUSTID",
        "CUST",
        "NAME",
        "TOPEKA",
        "3121 SE",
        "PH:",
        "APPT:",
    ]

    i = 0

    while i < len(lines):
        line = lines[i]
        starts_customer = False
        initial_name_parts = []

        if re.match(r"^\d{1,12}/$", line):
            starts_customer = True

        elif re.search(r"\b\d{1,12}/\s*$", line):
            starts_customer = True

        elif re.search(r"\b[A-Z]{1,4}/\s*[A-Z]", line):
            starts_customer = True
            after_slash = line.split("/", 1)[1].strip()
            if after_slash:
                initial_name_parts.append(after_slash)

        if not starts_customer:
            i += 1
            continue

        name_parts = initial_name_parts[:]
        j = i + 1

        while j < len(lines):
            part = lines[j]
            part_upper = part.upper()

            if part.startswith("/"):
                break

            if re.match(r"^\d{8,12}-\d{3}", part):
                break

            if any(word in part_upper for word in stop_words):
                break

            if "TOTAL" in part_upper:
                break

            if re.fullmatch(r"[\d\.\-\,]+", part):
                break

            if re.match(r"^\d{1,12}/$", part):
                break

            name_parts.append(part)
            j += 1

        customer = clean_customer_destination(" ".join(name_parts))

        if customer:
            customers.append(customer)

        i = max(j, i + 1)

    cleaned = []

    for customer in customers:
        customer_upper = customer.upper()

        if not customer_upper:
            continue

        if customer_upper in ["CUST", "NAME", "PO"]:
            continue

        if len(customer_upper) <= 1:
            continue

        cleaned.append(customer)

    return cleaned


CUSTOMER_ABBREVIATIONS = {
    "US FOODS": "US FOODS",
    "U.S. FOODS": "US FOODS",
    "SYSCO": "SYSCO",
    "PFS": "PFS",
    "PERFORMANCE FOOD": "PFS",
    "LOBLAWS": "LOBLAWS",
    "LOWLAWS": "LOBLAWS",
    "SOBEYS": "SOBEYS",
    "COSTCO": "COSTCO",
    "AWG": "AWG",
    "ASSOC WHLS GROC": "AWG",
    "ASSOCIATED WHOLESALE": "AWG",
    "ASSOCIATED GROCERS": "AWG",
    "JEWEL OSCO": "JEWEL OSCO",
    "JEWEL": "JEWEL OSCO",
    "SAFEWAY": "SAFEWAY",
    "WAKEFERN": "WAKEFERN",
    "BURRIS": "BURRIS",
    "KROGER": "KROGER",
    "LIPARI": "LIPARI",
    "HEB": "HEB",
    "H.E.B": "HEB",
    "GFS": "GFS",
    "GORDON": "GFS",
    "METRO": "METRO",
    "ALDI": "ALDI",
    "FOOD LION": "FOOD LION",
    "PUBLIX": "PUBLIX",
    "WEGMANS": "WEGMANS",
    "MARKET BASKET": "MARKET BASKET",
    "C&S": "C&S",
    "CERTCO": "CERTCO",
    "RESTAURANT DEPOT": "RESTAURANT DEPOT",
    "REDNERS": "REDNERS",
    "INGLES": "INGLES",
    "DARDEN": "DARDEN",
    "CORE MARK": "CORE MARK",
    "PALMER": "PALMER",
    "BIG Y": "BIG Y",
    "CHENEY": "CHENEY",
    "HARRIS TEETER": "HARRIS TEETER",
    "PERISHABLE DIST": "PDI",
    "PDI": "PDI",
    "FRAN TOMALIS": "FRAN TOMALIS",
    "TOMALIS": "FRAN TOMALIS",
    "NORTHWEST": "NORTHWEST",
    "PRATTS": "PRATTS",
    "FLANAGAN": "FLANAGAN",
    "ZARKYS": "ZARKYS",
    "LORENZ": "LORENZ",
    "VOILA": "VOILA",
    "RESER'S": "RESER'S",
    "RESERS": "RESER'S",
}


STATE_NAME_TO_CODE = {
    "ALABAMA": "AL",
    "ALASKA": "AK",
    "ARIZONA": "AZ",
    "ARKANSAS": "AR",
    "CALIFORNIA": "CA",
    "COLORADO": "CO",
    "CONNECTICUT": "CT",
    "DELAWARE": "DE",
    "FLORIDA": "FL",
    "GEORGIA": "GA",
    "HAWAII": "HI",
    "IDAHO": "ID",
    "ILLINOIS": "IL",
    "INDIANA": "IN",
    "IOWA": "IA",
    "KANSAS": "KS",
    "KENTUCKY": "KY",
    "LOUISIANA": "LA",
    "MAINE": "ME",
    "MARYLAND": "MD",
    "MASSACHUSETTS": "MA",
    "MICHIGAN": "MI",
    "MINNESOTA": "MN",
    "MISSISSIPPI": "MS",
    "MISSOURI": "MO",
    "MONTANA": "MT",
    "NEBRASKA": "NE",
    "NEVADA": "NV",
    "NEW HAMPSHIRE": "NH",
    "NEW JERSEY": "NJ",
    "NEW MEXICO": "NM",
    "NEW YORK": "NY",
    "NORTH CAROLINA": "NC",
    "NORTH DAKOTA": "ND",
    "OHIO": "OH",
    "OKLAHOMA": "OK",
    "OREGON": "OR",
    "PENNSYLVANIA": "PA",
    "RHODE ISLAND": "RI",
    "SOUTH CAROLINA": "SC",
    "SOUTH DAKOTA": "SD",
    "TENNESSEE": "TN",
    "TEXAS": "TX",
    "UTAH": "UT",
    "VERMONT": "VT",
    "VIRGINIA": "VA",
    "WASHINGTON": "WA",
    "WEST VIRGINIA": "WV",
    "WISCONSIN": "WI",
    "WYOMING": "WY",
    "ONTARIO": "ON",
    "QUEBEC": "QC",
    "BRITISH COLUMBIA": "BC",
    "ALBERTA": "AB",
    "MANITOBA": "MB",
    "SASKATCHEWAN": "SK",
    "NOVA SCOTIA": "NS",
    "NEW BRUNSWICK": "NB",
}


VALID_STATE_CODES = {
    "AL", "AK", "AZ", "AR", "CA", "CO", "CT", "DE", "FL", "GA",
    "HI", "ID", "IL", "IN", "IA", "KS", "KY", "LA", "ME", "MD",
    "MA", "MI", "MN", "MS", "MO", "MT", "NE", "NV", "NH", "NJ",
    "NM", "NY", "NC", "ND", "OH", "OK", "OR", "PA", "RI", "SC",
    "SD", "TN", "TX", "UT", "VT", "VA", "WA", "WV", "WI", "WY",
    "ON", "QC", "BC", "AB", "MB", "SK", "NS", "NB", "NL", "PE",
    "YT", "NT", "NU",
}


def get_destination_text_only(combined_text: str) -> str:
    if "Pickup TOTAL" in combined_text:
        return combined_text.split("Pickup TOTAL", 1)[-1]

    return combined_text


def extract_state_codes_from_text(text: str) -> list[str]:
    states = []
    upper_text = text.upper()

    us_matches = re.findall(r",\s*([A-Z]{2})\s+\d{5}", upper_text)
    ca_matches = re.findall(r",\s*([A-Z]{2})\s+[A-Z]\d[A-Z]\s*\d[A-Z]\d", upper_text)

    name_matches = []

    for state_name, state_code in STATE_NAME_TO_CODE.items():
        if state_name in upper_text:
            name_matches.append(state_code)

    for state in us_matches + ca_matches + name_matches:
        state = state.strip().upper()

        if state in VALID_STATE_CODES and state not in states:
            states.append(state)

    return states


def extract_state_from_customer_line(customer_line: str, destination_text: str) -> str:
    text = clean_spaces(customer_line).upper()

    words = re.findall(r"\b[A-Z]{2}\b", text)

    for word in reversed(words):
        if word in VALID_STATE_CODES:
            return word

    for state_name, state_code in STATE_NAME_TO_CODE.items():
        if state_name in text:
            return state_code

    states = extract_state_codes_from_text(destination_text)

    if states:
        return states[0]

    return ""


def abbreviate_customer_name(raw_customer: str) -> str:
    if not raw_customer:
        return ""

    text = clean_spaces(raw_customer).upper()
    text = re.sub(r"^\d{5,12}\s+", "", text)

    noise_words = [
        "INC",
        "LLC",
        "CORP",
        "CORPORATION",
        "COMPANY",
        "CO",
        "DELI",
        "MEAT",
        "TK",
        "T",
        "TORTILLA",
        "SALADS",
        "BAKED",
        "FRESH DC",
        "REGION",
        "DISTRIBUTION CENTER",
        "WAREHOUSE",
        "CUSTOMER PICK UP CARRIER",
        "RETAIL",
        "FOODSERVICE",
        "FOOD SERVICE",
        "MARKET",
        "MARKETS",
        "GENERAL STORES",
    ]

    for word in noise_words:
        text = re.sub(rf"\b{re.escape(word)}\b", "", text)

    text = clean_spaces(text)

    for key, short_name in CUSTOMER_ABBREVIATIONS.items():
        if key in text:
            return short_name

    words = text.split()
    return " ".join(words[:3])


def build_short_customer_field(customer_lines: list[str], combined_text: str) -> str:
    destination_text = get_destination_text_only(combined_text)
    short_customers = []

    for customer_line in customer_lines:
        short_name = abbreviate_customer_name(customer_line)
        state = extract_state_from_customer_line(customer_line, destination_text)

        if state:
            short_customer = clean_spaces(f"{short_name} {state}")
        else:
            short_customer = clean_spaces(short_name)

        if not short_customer:
            continue

        if short_customer not in short_customers:
            short_customers.append(short_customer)

    return " / ".join(short_customers)


# ============================================================
# TT4 DETECTION
# ============================================================

def is_canadian_load(text: str) -> bool:
    canadian_patterns = [
        r"\bON\b",
        r"\bQC\b",
        r"\bBC\b",
        r"\bAB\b",
        r"\bMB\b",
        r"\bSK\b",
        r"\bNS\b",
        r"\bNB\b",
        r"\bNL\b",
        r"\bPE\b",
        r"\bYT\b",
        r"\bNT\b",
        r"\bNU\b",
        "ONTARIO",
        "QUEBEC",
        "CANADA",
        "BRAMPTON",
        "VAUGHAN",
        "MISSISSAUGA",
        "TORONTO",
        "MILTON",
        "AJAX",
        "CAMBRIDGE",
        "REGINA",
        "WINNIPEG",
        "VARENNES",
    ]

    upper_text = text.upper()

    for pattern in canadian_patterns:
        if pattern.startswith(r"\b"):
            if re.search(pattern, upper_text):
                return True
        else:
            if pattern in upper_text:
                return True

    return False


def detect_tt4_required(record: dict) -> str:
    customer = record.get("Customer", "").upper()
    raw_text = record.get("Raw Text", "").upper()

    if TT4_FOR_CANADA and record.get("Is Canada", False):
        return "TT4"

    for keyword in TT4_CUSTOMER_KEYWORDS:
        keyword_upper = keyword.upper()

        if keyword_upper in customer or keyword_upper in raw_text:
            return "TT4"

    return ""


def parse_manifest_pdf(pdf_bytes: bytes) -> tuple[list[dict], list[dict]]:
    pages = extract_pdf_text_by_page(pdf_bytes)

    pages_by_load = defaultdict(list)

    for page_text in pages:
        load_number = extract_load_number(page_text)

        if load_number:
            pages_by_load[load_number].append(page_text)

    records = []
    duplicate_pdf_records = []
    seen_pdf_loads = set()

    for load_number, load_pages in pages_by_load.items():
        clean_load = normalize_load_number(load_number)

        if clean_load in seen_pdf_loads:
            duplicate_pdf_records.append({
                "Load": clean_load,
                "Reason": "Duplicate inside uploaded PDF"
            })
            continue

        seen_pdf_loads.add(clean_load)

        loading_pages = [p for p in load_pages if "LOADING MANIFEST" in p.upper()]
        shipping_pages = [p for p in load_pages if "SHIPPING MANIFEST" in p.upper()]

        if not loading_pages:
            continue

        loading_text = "\n".join(loading_pages)
        shipping_text = "\n".join(shipping_pages)
        combined_text = loading_text + "\n" + shipping_text

        full_datetime, date_text, time_text = extract_pickup_datetime(loading_text)

        carrier = extract_carrier(loading_text)
        cases = extract_cases(loading_text)

        order_customer_lines = extract_order_customer_lines_from_shipping_manifest(shipping_text)
        stop_customer_lines = extract_stop_customer_lines_from_loading_manifest(loading_text)

        if order_customer_lines:
            customer_lines = order_customer_lines
        else:
            customer_lines = stop_customer_lines

        customer = build_short_customer_field(customer_lines, combined_text)

        record = {
            "Load": clean_load,
            "Customer": customer,
            "Carrier": carrier,
            "Load Type": "",
            "Pickup DateTime": full_datetime,
            "Pickup Date": date_text,
            "Pickup Time": time_text,
            "Cases": cases,
            "Is Canada": is_canadian_load(combined_text),
            "Raw Text": combined_text,
            "Source": "New",
        }

        record["TT4"] = detect_tt4_required(record)

        records.append(record)

    records.sort(
        key=lambda x: (
            parse_sort_datetime(x.get("Pickup DateTime", "")),
            x.get("Load", "")
        )
    )

    return records, duplicate_pdf_records


# ============================================================
# BOARD HELPERS
# ============================================================

def find_label_column(ws, label_text: str, search_rows=(1, 2, 3)):
    label_text = label_text.strip().upper()

    for row in search_rows:
        for col in range(1, ws.max_column + 1):
            value = ws.cell(row, col).value

            if value is None:
                continue

            if str(value).strip().upper() == label_text:
                return col

    return None


def find_last_used_board_row(ws, start_row: int) -> int:
    last_row = start_row - 1

    for row in range(start_row, ws.max_row + 1):
        row_has_value = False

        for col in range(TABLE_MIN_COL, TABLE_MAX_COL + 1):
            value = ws.cell(row, col).value

            if value not in [None, ""]:
                row_has_value = True
                break

        if row_has_value:
            last_row = row

    return last_row


def is_day_row_value(value) -> bool:
    if value is None:
        return False

    weekday_names = {
        "MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY",
        "FRIDAY", "SATURDAY", "SUNDAY"
    }

    return str(value).strip().upper() in weekday_names


def read_existing_board_records(ws, start_row: int, board_columns: dict) -> tuple[list[dict], dict]:
    records = []
    existing_day_totals = defaultdict(lambda: {"Cases": 0, "Qty": 0})

    last_row = find_last_used_board_row(ws, start_row)

    cases_label_col = find_label_column(ws, "Cases")
    qty_label_col = find_label_column(ws, "Qty")

    current_date = ""

    load_col = board_columns.get("load_col", COL_LOAD_OR_DAY)
    customer_col = board_columns.get("customer_col", COL_CUSTOMER)
    carrier_col = board_columns.get("carrier_col", COL_CARRIER_DEFAULT)
    type_col = board_columns.get("type_col")
    time_col = board_columns.get("time_col", COL_TIME)
    tt4_col = board_columns.get("tt4_col", COL_TT4)

    for row in range(start_row, last_row + 1):
        col_a = ws.cell(row, load_col).value
        col_b = ws.cell(row, customer_col).value
        col_d = ws.cell(row, time_col).value

        if is_day_row_value(col_a):
            current_date = parse_date_any(col_b)

            if current_date:
                if cases_label_col:
                    existing_day_totals[current_date]["Cases"] = parse_number(
                        ws.cell(row, cases_label_col + 1).value
                    )

                if qty_label_col:
                    existing_day_totals[current_date]["Qty"] = parse_number(
                        ws.cell(row, qty_label_col + 1).value
                    )

            continue

        load_number = normalize_load_number(col_a)

        if load_number:
            pickup_time = format_time_only(col_d)
            pickup_datetime = build_datetime_from_date_time(current_date, pickup_time)

            carrier_value = clean_spaces(ws.cell(row, carrier_col).value)
            type_value = ""

            if type_col:
                type_value = clean_spaces(ws.cell(row, type_col).value)

            records.append({
                "Load": load_number,
                "Customer": clean_spaces(ws.cell(row, customer_col).value),
                "Load Type": type_value,
                "Carrier": carrier_value,
                "Pickup DateTime": pickup_datetime,
                "Pickup Date": current_date,
                "Pickup Time": pickup_time,
                "Cases": 0,
                "TT4": clean_spaces(ws.cell(row, tt4_col).value),
                "Source": "Existing",
            })

    return records, dict(existing_day_totals)


def group_records_by_date(records: list[dict]) -> dict:
    grouped = defaultdict(list)

    for record in records:
        date_text = record.get("Pickup Date", "")
        grouped[date_text].append(record)

    for date_text in grouped:
        grouped[date_text].sort(
            key=lambda x: (
                parse_sort_datetime(x.get("Pickup DateTime", "")),
                x.get("Load", "")
            )
        )

    return dict(
        sorted(
            grouped.items(),
            key=lambda x: parse_sort_datetime(x[0] + " 00:00")
        )
    )


def total_cases_for_day(records_for_day: list[dict], existing_day_total: int = 0):
    new_cases = sum(
        parse_number(record.get("Cases"))
        for record in records_for_day
        if record.get("Source") == "New"
    )

    return parse_number(existing_day_total) + new_cases


def total_loads_for_day(records_for_day: list[dict]):
    return len(records_for_day)


# ============================================================
# EXCEL FORMATTING
# ============================================================

def copy_cell(source_cell, target_cell):
    target_cell.value = source_cell.value

    if source_cell.has_style:
        target_cell.font = copy(source_cell.font)
        target_cell.fill = copy(source_cell.fill)
        target_cell.border = copy(source_cell.border)
        target_cell.alignment = copy(source_cell.alignment)
        target_cell.number_format = source_cell.number_format
        target_cell.protection = copy(source_cell.protection)


def copy_row_template(ws, source_row: int, target_row: int):
    ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height

    for col in range(BOARD_MIN_COL, BOARD_MAX_COL + 1):
        source_cell = ws.cell(source_row, col)
        target_cell = ws.cell(target_row, col)
        copy_cell(source_cell, target_cell)


def make_thin_border():
    return Border(
        left=Side(style="thin", color="000000"),
        right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"),
        bottom=Side(style="thin", color="000000"),
    )


def apply_borders_to_range(ws, start_row: int, end_row: int):
    border = make_thin_border()

    for row in range(start_row, end_row + 1):
        for col in range(BOARD_MIN_COL, BOARD_MAX_COL + 1):
            ws.cell(row, col).border = border


def style_day_row(ws, row: int):
    dark_blue = "002060"
    border = make_thin_border()

    for col in range(BOARD_MIN_COL, BOARD_MAX_COL + 1):
        cell = ws.cell(row, col)
        cell.fill = PatternFill("solid", fgColor=dark_blue)
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border


def style_load_row(ws, row: int, board_columns: dict):
    input_columns = [
        board_columns.get("load_col", COL_LOAD_OR_DAY),
        board_columns.get("customer_col", COL_CUSTOMER),
        board_columns.get("carrier_col", COL_CARRIER_DEFAULT),
        board_columns.get("time_col", COL_TIME),
        board_columns.get("tt4_col", COL_TT4),
    ]

    if board_columns.get("type_col"):
        input_columns.append(board_columns.get("type_col"))

    for col in sorted(set(input_columns)):
        if col is None:
            continue

        cell = ws.cell(row, col)
        cell.font = Font(color="000000", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=False)
        cell.border = make_thin_border()


def clear_existing_table_values(ws, start_row: int, board_columns: dict):
    last_row = find_last_used_board_row(ws, start_row)

    if last_row < start_row:
        return

    cases_label_col = find_label_column(ws, "Cases")
    qty_label_col = find_label_column(ws, "Qty")

    columns_to_clear = [
        board_columns.get("load_col", COL_LOAD_OR_DAY),
        board_columns.get("customer_col", COL_CUSTOMER),
        board_columns.get("carrier_col", COL_CARRIER_DEFAULT),
        board_columns.get("time_col", COL_TIME),
        board_columns.get("tt4_col", COL_TT4),
    ]

    if board_columns.get("type_col"):
        columns_to_clear.append(board_columns.get("type_col"))

    if cases_label_col:
        columns_to_clear.append(cases_label_col + 1)

    if qty_label_col:
        columns_to_clear.append(qty_label_col + 1)

    columns_to_clear = sorted(set([col for col in columns_to_clear if col]))

    for row in range(start_row, last_row + 1):
        for col in columns_to_clear:
            ws.cell(row, col).value = None


def clear_extra_output_columns(ws, start_row: int):
    last_row = max(ws.max_row, start_row)

    for row in range(start_row, last_row + 1):
        for col in CLEAR_EXTRA_OUTPUT_COLS:
            ws.cell(row, col).value = None


# ============================================================
# EXCEL BOARD POPULATION
# ============================================================

def populate_board(
    excel_bytes: bytes,
    pdf_records: list[dict],
    appointment_lookup: dict
) -> tuple[bytes, list[dict], list[dict], list[dict], list[dict]]:
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))

    if SHEET_NAME not in wb.sheetnames:
        raise ValueError(
            f"Sheet '{SHEET_NAME}' was not found. Available sheets: {wb.sheetnames}"
        )

    ws = wb[SHEET_NAME]

    board_columns = resolve_board_columns(ws)

    appointment_records, report_match_issues = build_records_from_appointments(
        pdf_records,
        appointment_lookup
    )

    existing_records, existing_day_totals = read_existing_board_records(
        ws,
        START_ROW,
        board_columns
    )

    existing_loads = {normalize_load_number(record.get("Load")) for record in existing_records}

    added_records = []
    skipped_duplicates = []
    seen_new_upload = set()

    for record in appointment_records:
        load_number = normalize_load_number(record.get("Load", ""))

        if not load_number:
            skipped_duplicates.append({
                "Load": record.get("Load", ""),
                "Customer": record.get("Customer", ""),
                "Reason": "Missing or invalid load number"
            })
            continue

        if load_number in existing_loads:
            skipped_duplicates.append({
                "Load": load_number,
                "Customer": record.get("Customer", ""),
                "Reason": "Already exists in uploaded Excel board"
            })
            continue

        if load_number in seen_new_upload:
            skipped_duplicates.append({
                "Load": load_number,
                "Customer": record.get("Customer", ""),
                "Reason": "Duplicate inside appointments Excel upload"
            })
            continue

        seen_new_upload.add(load_number)
        record["Source"] = "New"
        added_records.append(record)

    final_records = existing_records + added_records

    final_records.sort(
        key=lambda x: (
            parse_sort_datetime(x.get("Pickup DateTime", "")),
            x.get("Load", "")
        )
    )

    clear_existing_table_values(ws, START_ROW, board_columns)

    grouped = group_records_by_date(final_records)

    current_row = START_ROW

    load_col = board_columns.get("load_col", COL_LOAD_OR_DAY)
    customer_col = board_columns.get("customer_col", COL_CUSTOMER)
    carrier_col = board_columns.get("carrier_col", COL_CARRIER_DEFAULT)
    type_col = board_columns.get("type_col")
    time_col = board_columns.get("time_col", COL_TIME)
    tt4_col = board_columns.get("tt4_col", COL_TT4)

    cases_label_col = find_label_column(ws, "Cases")
    qty_label_col = find_label_column(ws, "Qty")

    for date_text, records_for_day in grouped.items():
        copy_row_template(ws, TEMPLATE_DAY_ROW, current_row)
        style_day_row(ws, current_row)

        ws.cell(current_row, load_col).value = get_day_name(date_text)
        ws.cell(current_row, customer_col).value = format_date_short(date_text)

        if carrier_col:
            ws.cell(current_row, carrier_col).value = ""

        if type_col:
            ws.cell(current_row, type_col).value = ""

        ws.cell(current_row, time_col).value = ""

        existing_cases_for_day = existing_day_totals.get(date_text, {}).get("Cases", 0)

        if cases_label_col:
            cases_cell = ws.cell(current_row, cases_label_col + 1)
            cases_cell.value = total_cases_for_day(records_for_day, existing_cases_for_day)
            cases_cell.font = Font(color="FFFFFF", bold=True)
            cases_cell.alignment = Alignment(horizontal="center", vertical="center")
            cases_cell.border = make_thin_border()

        if qty_label_col:
            qty_cell = ws.cell(current_row, qty_label_col + 1)
            qty_cell.value = total_loads_for_day(records_for_day)
            qty_cell.font = Font(color="FFFFFF", bold=True)
            qty_cell.alignment = Alignment(horizontal="center", vertical="center")
            qty_cell.border = make_thin_border()

        current_row += 1

        for record in records_for_day:
            copy_row_template(ws, TEMPLATE_LOAD_ROW, current_row)

            ws.cell(current_row, load_col).value = record.get("Load", "")
            ws.cell(current_row, customer_col).value = record.get("Customer", "")
            ws.cell(current_row, carrier_col).value = get_carrier_output_value(record, board_columns)

            if type_col:
                ws.cell(current_row, type_col).value = record.get("Load Type", "")

            ws.cell(current_row, time_col).value = format_time_only(record.get("Pickup Time", ""))
            ws.cell(current_row, tt4_col).value = record.get("TT4", "")

            style_load_row(ws, current_row, board_columns)

            current_row += 1

    last_output_row = current_row - 1

    if last_output_row >= START_ROW:
        apply_borders_to_range(ws, START_ROW, last_output_row)

    old_last_row = find_last_used_board_row(ws, START_ROW)

    if old_last_row > last_output_row:
        columns_to_clear = [
            load_col,
            customer_col,
            carrier_col,
            time_col,
            tt4_col,
        ]

        if type_col:
            columns_to_clear.append(type_col)

        if cases_label_col:
            columns_to_clear.append(cases_label_col + 1)

        if qty_label_col:
            columns_to_clear.append(qty_label_col + 1)

        for row in range(last_output_row + 1, old_last_row + 1):
            for col in sorted(set(columns_to_clear)):
                ws.cell(row, col).value = None

    clear_extra_output_columns(ws, START_ROW)

    ws.column_dimensions[get_column_letter(load_col)].width = 14
    ws.column_dimensions[get_column_letter(customer_col)].width = 52
    ws.column_dimensions[get_column_letter(carrier_col)].width = 30
    ws.column_dimensions[get_column_letter(time_col)].width = 12
    ws.column_dimensions[get_column_letter(tt4_col)].width = 12

    if type_col:
        ws.column_dimensions[get_column_letter(type_col)].width = 18

    ws.freeze_panes = f"A{START_ROW}"

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return (
        output.getvalue(),
        added_records,
        skipped_duplicates,
        final_records,
        report_match_issues
    )


# ============================================================
# STREAMLIT APP
# ============================================================

st.set_page_config(
    page_title="MG Manifest Matcher + Excel Board Builder",
    layout="wide"
)

st.title("MG Manifest Matcher + Excel Board Builder")

st.write(
    "Upload the two MG reports, the Excel board, and the appointments Excel. "
    "The app will first build the matched manifest packet, then use that matched packet "
    "to populate the Excel board."
)

st.subheader("Step 1 - Upload all inputs")

mg_files = st.file_uploader(
    "Upload the two MG report files: Loading Manifest and Shipping Manifest. PDFs or ZIPs are accepted.",
    type=["pdf", "zip"],
    accept_multiple_files=True
)

excel_board_file = st.file_uploader(
    "Upload existing Excel Board",
    type=["xlsx"]
)

appointments_file = st.file_uploader(
    "Upload Appointments Excel",
    type=["xlsx"]
)

with st.expander("TT4 settings"):
    st.write("The manifest does not clearly show a TT4 field. This app flags TT4 by rules.")
    st.write(f"TT4 for Canada loads: {TT4_FOR_CANADA}")
    st.write("Extra customer keywords requiring TT4:")
    st.write(TT4_CUSTOMER_KEYWORDS if TT4_CUSTOMER_KEYWORDS else "No extra keywords added.")

if mg_files and excel_board_file and appointments_file:
    if st.button("Build matched packet and update Excel board"):
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

            with st.spinner("Reading appointments Excel..."):
                appointment_lookup = parse_appointments_excel(appointments_file.read())

            total_appointment_loads = sum(len(rows) for rows in appointment_lookup.values())
            st.success(f"Loaded {total_appointment_loads} appointment rows from the appointments Excel.")

            with st.spinner("Parsing matched manifest packet..."):
                pdf_records, duplicate_pdf_records = parse_manifest_pdf(matched_pdf_bytes)

            if not pdf_records:
                st.warning(
                    "No loading manifest records were found in the matched PDF. "
                    "The app will still add appointment loads, but they will be flagged as missing from PDF."
                )
            else:
                st.success(f"Found {len(pdf_records)} loads in the matched PDF.")

            with st.spinner("Updating Excel board..."):
                (
                    populated_excel,
                    added_records,
                    skipped_duplicates,
                    final_sorted_records,
                    report_match_issues
                ) = populate_board(
                    excel_board_file.read(),
                    pdf_records,
                    appointment_lookup
                )

            st.success("Excel board created successfully.")

            st.subheader("Loads added to Excel")

            if added_records:
                preview_added = []

                for record in added_records:
                    preview_added.append({
                        "Load": record.get("Load", ""),
                        "Customer": record.get("Customer", ""),
                        "Load Type": record.get("Load Type", ""),
                        "Carrier": record.get("Carrier", ""),
                        "Time": record.get("Pickup Time", ""),
                        "TT4": record.get("TT4", ""),
                        "Appointment Match": record.get("Appointment Match", ""),
                        "Report Match Status": record.get("Report Match Status", ""),
                    })

                st.dataframe(preview_added, use_container_width=True)
            else:
                st.info("No new loads were added.")

            all_skipped = skipped_duplicates + duplicate_pdf_records

            st.subheader("Duplicate / skipped loads")

            if all_skipped:
                st.warning(f"Skipped {len(all_skipped)} duplicate or invalid loads.")
                st.dataframe(all_skipped, use_container_width=True)
            else:
                st.success("No duplicates found.")

            st.subheader("Report match issues")

            if report_match_issues:
                st.warning(
                    f"{len(report_match_issues)} loads were flagged because the PDF was missing the load or the date/time did not match."
                )
                st.dataframe(report_match_issues, use_container_width=True)
            else:
                st.success("Every appointment load matched the PDF by load number, date, and time.")

            st.subheader("Final sorted board preview")

            final_preview = []

            for record in final_sorted_records:
                final_preview.append({
                    "Load": record.get("Load", ""),
                    "Customer": record.get("Customer", ""),
                    "Load Type": record.get("Load Type", ""),
                    "Carrier": record.get("Carrier", ""),
                    "Time": record.get("Pickup Time", ""),
                    "TT4": record.get("TT4", ""),
                    "Report Match Status": record.get("Report Match Status", ""),
                })

            st.dataframe(final_preview, use_container_width=True)

            st.download_button(
                label="Download updated and sorted Excel board",
                data=populated_excel,
                file_name=OUTPUT_FILE_NAME,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error("Something went wrong while processing the files.")
            st.exception(e)

else:
    st.info("Upload the two MG reports, the Excel board, and the appointments Excel to start.")
