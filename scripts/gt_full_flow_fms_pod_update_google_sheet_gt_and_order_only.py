import os
import re
import json
import base64
from datetime import datetime

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from google_auth_httplib2 import AuthorizedHttp
import httplib2


# ============================================================
# CONFIGURATION
# ============================================================

# Destination Google Spreadsheet containing the Full Flow FMS tab.
# This is the spreadsheet that should be updated by this script.
OUTPUT_SPREADSHEET_ID = "1jMEEFkAVfg7NHtxj3QCYuNY_V-AhU00ljQeRYTBQDio"

SHEET_NAME = "Full Flow FMS"

SERVICE_ACCOUNT_FILE = (
    r"C:\Users\diwak\.gcloud\automation-bot.json"
)

POD_FOLDER_ID = "1heBZTSPUiF-5wFddGpEw64qzCyQaMwf1"

DATA_START_ROW = 8

# Google Sheets API batch size.
BATCH_SIZE = 500


# ============================================================
# GOOGLE SHEET COLUMNS
# ============================================================

# 1-based Google Sheets column numbers.
COL_GT = 3       # C  = GT Number
COL_ORDER = 16   # P  = Order No

# Destination columns.
COL_AJ = 36      # AJ = Order Delivery Status
COL_AP = 42      # AP = Actual
COL_AU = 47      # AU = POD Link
COL_AV = 48      # AV = Delivered Date

# Read-only evidence columns for the AJ stage (2026-11-01).
COL_AE = 31      # AE = AWB
COL_AW = 49      # AW = Invoice No

# A:AW = 49 columns (AW added 2026-11-01 for the "Invoiced" stage).
LAST_READ_COLUMN = "AW"

# Value written into AJ when a POD is found.
DELIVERED_TEXT = "Delivered"

# ------------------------------------------------------------
# AJ "Order Delivery Status" stages (2026-11-01, user decision).
# Before, AJ was only ever "Delivered" (POD found) / "Cancelled" / "Virtual Invoice",
# so 3,040 open lines read blank whatever their progress. Now:
#   blank < Invoiced (AW Invoice No filled) < Dispatched (AE AWB filled) < Delivered (POD)
# AJ is written only when it is blank or holds a LOWER stage of this ladder: never
# downgraded, any other text (a human's note, "Returned", ...) is never replaced, and
# closed lines (Cancelled / Virtual Invoice / OOS in C or P:AW) are never touched.
# "Dispatched" = a real AWB in AE. AB / AG / AL "Actual" are NOT evidence: the planned-
# status step fills them on plan from X (resolve_stage_actual), and AS Dispatched Qty
# is the Vinculum Pick Qty of the invoice line (same source as AW).
INVOICED_TEXT = "Invoiced"
DISPATCHED_TEXT = "Dispatched"
AJ_STAGE_RANK = {"": 0, INVOICED_TEXT.lower(): 1, DISPATCHED_TEXT.lower(): 2,
                 DELIVERED_TEXT.lower(): 3}


# ============================================================
# GOOGLE DRIVE CONNECTION
# ============================================================

def connect_google_apis():
    print("\nConnecting to Google Drive + Google Sheets...")

    scopes = [
        "https://www.googleapis.com/auth/drive.readonly",
        "https://www.googleapis.com/auth/spreadsheets",
    ]

    raw = os.environ.get("GCP_SA_KEY")
    if raw:
        try:
            info = json.loads(raw)
        except json.JSONDecodeError:
            info = json.loads(base64.b64decode(raw))
        credentials = Credentials.from_service_account_info(info, scopes=scopes)
        print("Using service-account credentials from GCP_SA_KEY env var")
    else:
        if not os.path.exists(SERVICE_ACCOUNT_FILE):
            raise FileNotFoundError(
                "\nService account JSON not found:\n"
                f"{SERVICE_ACCOUNT_FILE}"
            )
        credentials = Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=scopes
        )
        print(f"Service account: {credentials.service_account_email}")

    http = httplib2.Http(timeout=60)

    authorized_http = AuthorizedHttp(
        credentials,
        http=http,
    )

    drive = build(
        "drive",
        "v3",
        http=authorized_http,
        cache_discovery=False,
    )

    sheets = build(
        "sheets",
        "v4",
        http=authorized_http,
        cache_discovery=False,
    )

    print("Google Drive connection successful.")
    print("Google Sheets connection successful.")

    return drive, sheets


# ============================================================
# NORMALIZE GT
# ============================================================

def normalize_gt(gt):

    gt = str(
        gt or ""
    ).strip().upper()

    # Example:
    #
    # GT310726A3R
    #       ↓
    # GT310726A3

    return re.sub(
        r"([0-9])[A-Z]$",
        r"\1",
        gt
    )


# ============================================================
# NORMALIZE ORDER
# ============================================================

def normalize_order(order):

    if order is None:
        return ""

    return str(
        order
    ).strip().upper()


# ============================================================
# PARSE DATE
# ============================================================

def parse_date(date_text):

    if not date_text:
        return None

    date_text = str(
        date_text
    ).strip()

    date_text = (
        date_text
        .replace("–", "-")
        .replace("—", "-")
        .replace("−", "-")
    )

    formats = [

        # ----------------------------------------------------
        # DD-MM-YYYY / D-M-YYYY
        # ----------------------------------------------------

        "%d-%m-%Y",
        "%d-%m-%Y %H:%M",
        "%d-%m-%Y %H:%M:%S",

        # ----------------------------------------------------
        # DD/MM/YYYY / D/M/YYYY
        # ----------------------------------------------------

        "%d/%m/%Y",
        "%d/%m/%Y %H:%M",
        "%d/%m/%Y %H:%M:%S",

        # ----------------------------------------------------
        # DD-MMM-YYYY
        # ----------------------------------------------------

        "%d-%b-%Y",
        "%d-%b-%Y %H:%M",
        "%d-%b-%Y %H:%M:%S",

        # ----------------------------------------------------
        # DD-MONTH-YYYY
        # ----------------------------------------------------

        "%d-%B-%Y",
        "%d-%B-%Y %H:%M",
        "%d-%B-%Y %H:%M:%S",
    ]

    for fmt in formats:

        try:

            return datetime.strptime(
                date_text,
                fmt
            )

        except ValueError:
            continue

    return None


# ============================================================
# EXTRACT DATE FROM FILENAME
# ============================================================

def extract_date_from_filename(filename):

    """
    Finds a date anywhere in the filename.

    Supported examples:

    31-08-2026
    31-8-2026
    31-08-2026 15:13
    31-08-2026 15:13:57

    14-04-2026 19:14:28

    07-5-2026

    26-AUG-2026
    """

    filename = str(
        filename or ""
    )

    filename = (
        filename
        .replace("–", "-")
        .replace("—", "-")
        .replace("−", "-")
    )

    # ========================================================
    # DD-MM-YYYY with optional time
    # Allows 1 or 2 digits for day/month
    # ========================================================

    match = re.search(
        r"\b"
        r"(\d{1,2}[-/]\d{1,2}[-/]\d{4}"
        r"(?:\s+\d{1,2}:\d{2}"
        r"(?::\d{2})?)?)"
        r"\b",
        filename
    )

    if match:

        parsed = parse_date(
            match.group(1)
        )

        if parsed:

            return parsed

    # ========================================================
    # DD-MMM-YYYY
    # ========================================================

    match = re.search(
        r"\b"
        r"(\d{1,2}-[A-Za-z]{3}-\d{4}"
        r"(?:\s+\d{1,2}:\d{2}"
        r"(?::\d{2})?)?)"
        r"\b",
        filename
    )

    if match:

        parsed = parse_date(
            match.group(1)
        )

        if parsed:

            return parsed

    return None


# ============================================================
# EXTRACT GT FROM FILENAME
# ============================================================

def extract_gt_from_filename(filename):

    """
    Finds GT number anywhere in filename.

    Examples:

    GT060426A9 | DELIVERY CONFIRMATION RECEIPT | ...

    Invoice_N012080_GT060426A9_2704N10469.pdf
    """

    filename = str(
        filename or ""
    ).upper()

    matches = re.findall(
        r"\bGT[0-9]+[A-Z0-9]*\b",
        filename
    )

    if matches:

        return matches[0].strip().upper()

    return ""


# ============================================================
# PARSE STRUCTURED POD FILE
# ============================================================

def parse_structured_pod_filename(filename):

    """
    Handles structured POD filenames.

    Examples:

    GT190826A1 | POD | 31-08-2026 15:13:57 | 313676200.jpeg

    GT260826A2 | POD | O94115 | 31-08-2026 | H90304949445

    GT190826A1 | POD| N051587 | 24-08-2026 | 26448310520800

    GT290426A1 | POD | N89484 | 07-5-2026 | DD01AJ9424.png
    """

    filename = str(
        filename or ""
    ).strip()

    if not filename:
        return None

    # --------------------------------------------------------
    # Split on | and ignore spaces around it
    # --------------------------------------------------------

    parts = re.split(
        r"\s*\|\s*",
        filename
    )

    parts = [
        p.strip()
        for p in parts
        if p.strip()
    ]

    if len(parts) < 3:
        return None

    # --------------------------------------------------------
    # First part should be GT
    # --------------------------------------------------------

    gt = parts[0].strip().upper()

    if not gt.startswith("GT"):

        return None

    # --------------------------------------------------------
    # Find POD marker
    # --------------------------------------------------------

    pod_index = None

    for i, part in enumerate(parts):

        if part.strip().upper() == "POD":

            pod_index = i

            break

    if pod_index is None:
        return None

    # --------------------------------------------------------
    # Everything after POD
    # --------------------------------------------------------

    after_pod = parts[
        pod_index + 1:
    ]

    if not after_pod:
        return None

    # --------------------------------------------------------
    # Find date
    # --------------------------------------------------------

    date_index = None
    pod_date = None

    for i, part in enumerate(after_pod):

        parsed_date = extract_date_from_filename(
            part
        )

        if parsed_date:

            date_index = i
            pod_date = parsed_date

            break

    if date_index is None:
        return None

    # --------------------------------------------------------
    # Order ID
    #
    # If something appears between POD and date,
    # treat first value as Order ID.
    # --------------------------------------------------------

    order_id = ""

    if date_index > 0:

        order_id = (
            after_pod[0]
            .strip()
            .upper()
        )

    return {
        "gt": gt,
        "order": order_id,
        "date": pod_date,
        "filename": filename,
        "structured": True,
    }


# ============================================================
# BUILD POD INDEX
# ============================================================

def build_pod_index(drive):

    print()
    print("=" * 80)
    print("READING POD FOLDER")
    print("=" * 80)

    # --------------------------------------------------------
    # Two groups:
    #
    # 1. Structured POD files
    # 2. Any file with GT + Date
    # --------------------------------------------------------

    structured_records = []

    all_gt_records = []

    page_token = None
    page_number = 0

    total_files = 0
    ignored_folders = 0

    parsed_structured = 0
    parsed_gt_only = 0

    skipped_no_gt = 0
    skipped_no_date = 0

    # ========================================================
    # READ DRIVE
    # ========================================================

    while True:

        page_number += 1

        print(
            f"\nReading POD page {page_number}..."
        )

        try:

            response = drive.files().list(

                q=(
                    f"'{POD_FOLDER_ID}' in parents "
                    "and trashed = false"
                ),

                spaces="drive",

                fields=(
                    "nextPageToken,"
                    "files("
                    "id,"
                    "name,"
                    "mimeType,"
                    "webViewLink"
                    ")"
                ),

                pageSize=100,

                pageToken=page_token,

                includeItemsFromAllDrives=True,

                supportsAllDrives=True,

            ).execute()

        except Exception as e:

            print()
            print("=" * 80)
            print("ERROR READING POD FOLDER")
            print("=" * 80)
            print(e)

            raise

        files = response.get(
            "files",
            []
        )

        print(
            f"Files received: {len(files)}"
        )

        # ====================================================
        # PROCESS FILES
        # ====================================================

        for file in files:

            total_files += 1

            name = file.get(
                "name",
                ""
            ).strip()

            if not name:
                continue

            # ------------------------------------------------
            # Ignore folders
            # ------------------------------------------------

            if (
                file.get("mimeType")
                == "application/vnd.google-apps.folder"
            ):

                ignored_folders += 1

                continue

            # ------------------------------------------------
            # Get URL
            # ------------------------------------------------

            url = file.get(
                "webViewLink"
            )

            if not url:

                file_id = file.get(
                    "id"
                )

                if file_id:

                    url = (
                        "https://drive.google.com/file/d/"
                        f"{file_id}/view"
                    )

            # =================================================
            # LOGIC 1:
            #
            # Structured POD + Order
            # =================================================

            structured = (
                parse_structured_pod_filename(
                    name
                )
            )

            if structured:

                structured["url"] = url

                structured["file_id"] = (
                    file.get("id")
                )

                structured_records.append(
                    structured
                )

                all_gt_records.append(
                    structured
                )

                parsed_structured += 1

                continue

            # =================================================
            # LOGIC 3:
            #
            # GT + ANY DATE
            #
            # Example:
            #
            # GT060426A9 |
            # DELIVERY CONFIRMATION RECEIPT |
            # 14-04-2026 19:14:28 |
            # ...
            #
            # This does NOT require "POD".
            # =================================================

            gt = extract_gt_from_filename(
                name
            )

            if not gt:

                skipped_no_gt += 1

                continue

            file_date = (
                extract_date_from_filename(
                    name
                )
            )

            if not file_date:

                skipped_no_date += 1

                continue

            record = {

                "gt": gt,

                "order": "",

                "date": file_date,

                "filename": name,

                "url": url,

                "file_id": file.get("id"),

                "structured": False,
            }

            all_gt_records.append(
                record
            )

            parsed_gt_only += 1

        # ====================================================
        # NEXT PAGE
        # ====================================================

        page_token = response.get(
            "nextPageToken"
        )

        if not page_token:

            print(
                "\nNo more POD pages."
            )

            break

    # ========================================================
    # BUILD EXACT GT + ORDER INDEX
    # ========================================================

    print()
    print("-" * 80)
    print("BUILDING POD INDEXES")
    print("-" * 80)

    exact_index = {}

    gt_index = {}

    # ========================================================
    # EXACT GT + ORDER
    # ========================================================

    for record in structured_records:

        gt = record["gt"]

        order = normalize_order(
            record["order"]
        )

        if not order:
            continue

        gt_norm = normalize_gt(
            gt
        )

        # ----------------------------------------------------
        # Exact GT + Order
        # ----------------------------------------------------

        key = (
            f"{gt}|{order}"
        )

        # Keep latest POD
        if (
            key not in exact_index
            or record["date"]
            > exact_index[key]["date"]
        ):

            exact_index[key] = record

        # ----------------------------------------------------
        # Normalized GT + Order
        # ----------------------------------------------------

        normalized_key = (
            f"{gt_norm}|{order}"
        )

        if (
            normalized_key not in exact_index
            or record["date"]
            > exact_index[
                normalized_key
            ]["date"]
        ):

            exact_index[
                normalized_key
            ] = record

    # ========================================================
    # GT-ONLY INDEX
    # ========================================================

    for record in all_gt_records:

        gt = record["gt"]

        gt_norm = normalize_gt(
            gt
        )

        # ----------------------------------------------------
        # Exact GT
        # ----------------------------------------------------

        if (
            gt not in gt_index
            or record["date"]
            > gt_index[gt]["date"]
        ):

            gt_index[gt] = record

        # ----------------------------------------------------
        # Normalized GT
        # ----------------------------------------------------

        if (
            gt_norm not in gt_index
            or record["date"]
            > gt_index[
                gt_norm
            ]["date"]
        ):

            gt_index[
                gt_norm
            ] = record

    # ========================================================
    # SUMMARY
    # ========================================================

    print(
        f"Total Drive items        : "
        f"{total_files}"
    )

    print(
        f"Folders ignored          : "
        f"{ignored_folders}"
    )

    print(
        f"Structured POD files     : "
        f"{parsed_structured}"
    )

    print(
        f"GT + Date fallback files : "
        f"{parsed_gt_only}"
    )

    print(
        f"Skipped - no GT          : "
        f"{skipped_no_gt}"
    )

    print(
        f"Skipped - no Date        : "
        f"{skipped_no_date}"
    )

    print(
        f"GT + Order index keys    : "
        f"{len(exact_index)}"
    )

    print(
        f"GT-only index keys       : "
        f"{len(gt_index)}"
    )

    print("=" * 80)

    return exact_index, gt_index


# POD indexes are populated once in main() and then used by the row matcher.
exact_index = {}
gt_index = {}
closed_skipped = 0   # rows with a POD that were left alone because they are closed


# ============================================================
# FIND POD
# ============================================================

def find_pod(
    exact_index,
    gt_index,
    gt_number,
    order_number
):
    """
    Find a POD ONLY when BOTH GT Number and Order No are present
    and BOTH match the POD data.

    Matching priority:
      1. Exact GT + Order
      2. Normalized GT + Order

    IMPORTANT:
    GT-only matching is intentionally disabled.
    A row with only GT Number, or only Order No, will NOT be matched.
    """

    gt = str(
        gt_number or ""
    ).strip().upper()

    order = normalize_order(
        order_number
    )

    # Both fields are mandatory. "Cancelled" (and the other closed markers) in Order No
    # is a sentinel written by the Full Flow job / GT_FMS, never an order number.
    if not gt or not order or order.lower() in CLOSED_MARKERS:
        return None, ""

    gt_norm = normalize_gt(
        gt
    )

    # ========================================================
    # LOGIC 1
    # Exact GT + Order
    # ========================================================

    key = f"{gt}|{order}"

    if key in exact_index:
        return (
            exact_index[key],
            "GT + Order"
        )

    # ========================================================
    # LOGIC 2
    # Normalized GT + Order
    # ========================================================

    key = f"{gt_norm}|{order}"

    if key in exact_index:
        return (
            exact_index[key],
            "Normalized GT + Order"
        )

    # ========================================================
    # NO GT-ONLY MATCHING
    # ========================================================
    #
    # Previously the script could fall back to:
    #   GT Only
    #   Normalized GT Only
    #
    # That behavior is now removed.
    #
    # Therefore:
    #   GT present + Order missing -> NO MATCH
    #   GT missing + Order present -> NO MATCH
    #   GT present + Order present but not matching -> NO MATCH
    #

    return None, ""


# ============================================================
# CHECK IF CELL IS BLANK
# ============================================================

def is_blank(cell):

    if cell.value is None:

        return True

    if isinstance(
        cell.value,
        str
    ):

        return cell.value.strip() == ""

    return False


# ============================================================
# MAIN
# ============================================================

def get_sheet_id(sheets):
    response = sheets.spreadsheets().get(
        spreadsheetId=OUTPUT_SPREADSHEET_ID,
        fields="sheets(properties(sheetId,title))",
    ).execute()

    for sheet in response.get("sheets", []):
        properties = sheet.get("properties", {})
        if properties.get("title") == SHEET_NAME:
            return properties.get("sheetId")

    raise ValueError(
        f"Google Sheet tab '{SHEET_NAME}' was not found in "
        f"spreadsheet {OUTPUT_SPREADSHEET_ID}."
    )


def read_full_flow_sheet(sheets):
    print()
    print("=" * 80)
    print("READING GOOGLE SHEET")
    print("=" * 80)
    print(f"Spreadsheet ID : {OUTPUT_SPREADSHEET_ID}")
    print(f"Tab            : {SHEET_NAME}")
    print(f"Range          : {SHEET_NAME}!A{DATA_START_ROW}:{LAST_READ_COLUMN}")

    result = sheets.spreadsheets().values().get(
        spreadsheetId=OUTPUT_SPREADSHEET_ID,
        range=f"'{SHEET_NAME}'!A{DATA_START_ROW}:{LAST_READ_COLUMN}",
        valueRenderOption="FORMATTED_VALUE",
        dateTimeRenderOption="FORMATTED_STRING",
    ).execute()

    values = result.get("values", [])

    print(f"Rows read      : {len(values):,}")

    return values


def get_cell(row, column_number):
    index = column_number - 1
    if index >= len(row):
        return ""

    value = row[index]
    return "" if value is None else str(value).strip()


def is_blank_value(value):
    return value is None or str(value).strip() == ""


# Lines a human (GT_FMS dashboard Cancel GT / Mark Virtual Invoice) or an earlier
# job closed. GT_FMS markVirtualGT_ writes "Virtual Invoice" into ONLY AJ on punched
# lines (P:AW on open lines); Cancel writes "Cancelled" into P:AW. Before 2026-10-09
# only C / P / AP were checked, so a punched Virtual Invoice line (AJ only) was
# overwritten with "Delivered" whenever a POD file matched.
CLOSED_MARKERS = {"cancelled", "virtual invoice", "out of stock", "oos"}
CLOSED_MARKER_COLUMNS = [COL_GT] + list(range(COL_ORDER, COL_AW + 1))   # C, P:AW


def row_is_cancelled(row):
    """
    True if this row is closed (Cancelled / Virtual Invoice / Out of stock / OOS
    in C or anywhere in P:AW — includes AJ). Closed rows are never written: no
    AJ stage, not AP / AU / AV.
    """
    for column_number in CLOSED_MARKER_COLUMNS:
        if get_cell(row, column_number).strip().lower() in CLOSED_MARKERS:
            return True
    return False


def _has_evidence(row, column_number):
    value = get_cell(row, column_number)
    return bool(value) and value.lower() not in CLOSED_MARKERS


def aj_stage_for_row(row, pod_found):
    """Highest AJ stage the row has evidence for, or "" (none).
    Delivered = POD matched; Dispatched = AE AWB; Invoiced = AW Invoice No."""
    if pod_found:
        return DELIVERED_TEXT
    if _has_evidence(row, COL_AE):
        return DISPATCHED_TEXT
    if _has_evidence(row, COL_AW):
        return INVOICED_TEXT
    return ""


def aj_should_write(current_aj, stage):
    """True when AJ may move to `stage`: AJ is blank or a lower stage of the ladder.
    Unknown text (not a stage) is never replaced; equal / higher is never downgraded."""
    if not stage:
        return False
    current_rank = AJ_STAGE_RANK.get(str(current_aj or "").strip().lower())
    if current_rank is None:
        return False
    return current_rank < AJ_STAGE_RANK[stage.lower()]


def build_sheet_update_requests(values):
    """
    Match every Full Flow FMS row against the POD index.

    IMPORTANT:
    - AP is updated only when currently blank.
    - AU is updated only when currently blank.
    - AV is updated only when currently blank.
    - AJ (Order Delivery Status) moves up the stage ladder
      blank < "Invoiced" (AW) < "Dispatched" (AE AWB) < "Delivered" (POD)
      (2026-11-01). "Delivered" ONLY when a real POD is found. Never written
      on a closed row (Cancelled / Virtual Invoice / OOS in C or P:AW), never
      downgraded, never over a non-stage text.
    - Existing manual values are NEVER overwritten (except AJ, which is
      promoted to a higher stage).
    """
    rows_checked = 0
    matched = 0
    missed = 0

    exact_matches = 0
    normalized_order_matches = 0
    gt_only_matches = 0
    normalized_gt_matches = 0

    updates = []
    global closed_skipped
    closed_skipped = 0

    for offset, row in enumerate(values):
        row_num = DATA_START_ROW + offset
        rows_checked += 1

        gt = get_cell(row, COL_GT)
        if not gt:
            continue

        order = get_cell(row, COL_ORDER)

        pod, match_type = find_pod(
            exact_index,
            gt_index,
            gt,
            order,
        )

        if pod is None:
            missed += 1
            # 2026-11-01: no POD -> AJ "Invoiced" / "Dispatched" from AW / AE.
            if not row_is_cancelled(row):
                stage = aj_stage_for_row(row, pod_found=False)
                if aj_should_write(get_cell(row, COL_AJ), stage):
                    updates.append({
                        "range": f"'{SHEET_NAME}'!AJ{row_num}",
                        "values": [[stage]],
                    })
            continue

        matched += 1

        # Closed line (Cancelled / Virtual Invoice / OOS): leave it exactly as is.
        if row_is_cancelled(row):
            closed_skipped += 1
            continue

        if match_type == "GT + Order":
            exact_matches += 1
        elif match_type == "Normalized GT + Order":
            normalized_order_matches += 1
        elif match_type == "GT Only":
            gt_only_matches += 1
        elif match_type == "Normalized GT Only":
            normalized_gt_matches += 1

        pod_date = pod["date"]
        display_date = pod_date.strftime("%d-%b-%Y")
        display_datetime = pod_date.strftime("%d-%b-%Y %H:%M:%S")

        # Google Sheets A1 notation for this row.
        aj_cell = f"AJ{row_num}"
        ap_cell = f"AP{row_num}"
        au_cell = f"AU{row_num}"
        av_cell = f"AV{row_num}"

        current_aj = get_cell(row, COL_AJ)
        current_ap = get_cell(row, COL_AP)
        current_au = get_cell(row, COL_AU)
        current_av = get_cell(row, COL_AV)

        # AP = Actual — only fill if blank.
        if is_blank_value(current_ap):
            updates.append({
                "range": f"'{SHEET_NAME}'!{ap_cell}",
                "values": [[display_datetime]],
            })

        # AV = Delivered Date — only fill if blank.
        if is_blank_value(current_av):
            updates.append({
                "range": f"'{SHEET_NAME}'!{av_cell}",
                "values": [[display_date]],
            })

        # AU = POD Link — only fill if blank.
        if is_blank_value(current_au) and pod.get("url"):
            formula = (
                f'=HYPERLINK('
                f'"{pod["url"]}",'
                f'"{display_date}")'
            )

            updates.append({
                "range": f"'{SHEET_NAME}'!{au_cell}",
                "values": [[formula]],
            })

        # AJ = Order Delivery Status -> "Delivered".
        #
        # A POD was matched for this row, so it is genuinely delivered.
        # Closed rows were skipped above. Written over blank / "Invoiced" /
        # "Dispatched" only (aj_should_write): not rewritten when AJ already
        # reads "Delivered", and a non-stage text in AJ is left alone.
        if aj_should_write(current_aj, DELIVERED_TEXT):
            updates.append({
                "range": f"'{SHEET_NAME}'!{aj_cell}",
                "values": [[DELIVERED_TEXT]],
            })

    return (
        rows_checked,
        matched,
        missed,
        exact_matches,
        normalized_order_matches,
        gt_only_matches,
        normalized_gt_matches,
        updates,
    )


def apply_sheet_updates(sheets, updates):
    if not updates:
        print("\nNo AJ/AP/AU/AV cells needed updating.")
        return 0

    print()
    print("=" * 80)
    print("UPDATING GOOGLE SHEET")
    print("=" * 80)
    print(f"Cells to update: {len(updates):,}")

    total = 0

    for start_index in range(0, len(updates), BATCH_SIZE):
        batch = updates[start_index:start_index + BATCH_SIZE]

        sheets.spreadsheets().values().batchUpdate(
            spreadsheetId=OUTPUT_SPREADSHEET_ID,
            body={
                "valueInputOption": "USER_ENTERED",
                "data": batch,
            },
        ).execute()

        total += len(batch)

        print(
            f"Updated {total:,} / {len(updates):,} cells"
        )

    return total


def main():

    print()
    print("=" * 80)
    print("POD UPDATE - GOOGLE SHEETS FULL FLOW FMS")
    print("=" * 80)

    print()
    print(f"Destination Spreadsheet ID:\n{OUTPUT_SPREADSHEET_ID}")
    print(f"Destination tab:\n{SHEET_NAME}")
    print(f"Data starts at row: {DATA_START_ROW}")
    print(
        "Update columns: AJ (Order Delivery Status), AP (Actual), "
        "AU (POD Link), AV (Delivered Date)"
    )
    print("AJ stages: Invoiced (AW) < Dispatched (AE AWB) < Delivered (confirmed POD); "
          "only over blank / a lower stage, never on Cancelled / Virtual Invoice / OOS rows.")
    print("Existing non-blank values will NOT be overwritten.")

    # ========================================================
    # CONNECT
    # ========================================================

    drive, sheets = connect_google_apis()

    # ========================================================
    # VERIFY DESTINATION TAB
    # ========================================================

    print()
    print("=" * 80)
    print("CHECKING DESTINATION GOOGLE SHEET")
    print("=" * 80)

    sheet_id = get_sheet_id(sheets)

    print(
        f"Found destination tab: '{SHEET_NAME}' "
        f"(sheet ID: {sheet_id})"
    )

    # ========================================================
    # BUILD POD INDEX
    # ========================================================

    global exact_index, gt_index

    exact_index, gt_index = build_pod_index(drive)

    # ========================================================
    # READ GOOGLE SHEET
    # ========================================================

    values = read_full_flow_sheet(sheets)

    # ========================================================
    # MATCH PODs
    # ========================================================

    print()
    print("=" * 80)
    print("MATCHING PODs TO GOOGLE SHEET")
    print("=" * 80)

    (
        rows_checked,
        matched,
        missed,
        exact_matches,
        normalized_order_matches,
        gt_only_matches,
        normalized_gt_matches,
        updates,
    ) = build_sheet_update_requests(values)

    # ========================================================
    # WRITE GOOGLE SHEET
    # ========================================================

    updated_cells = apply_sheet_updates(
        sheets,
        updates,
    )

    # ========================================================
    # FINAL SUMMARY
    # ========================================================

    aj_by_stage = {
        stage: sum(
            1 for item in updates
            if "'Full Flow FMS'!AJ" in item["range"] and item["values"][0][0] == stage
        )
        for stage in (DELIVERED_TEXT, DISPATCHED_TEXT, INVOICED_TEXT)
    }
    aj_updated = aj_by_stage[DELIVERED_TEXT]

    ap_updated = sum(
        1 for item in updates
        if "'Full Flow FMS'!AP" in item["range"]
    )

    au_updated = sum(
        1 for item in updates
        if "'Full Flow FMS'!AU" in item["range"]
    )

    av_updated = sum(
        1 for item in updates
        if "'Full Flow FMS'!AV" in item["range"]
    )

    print()
    print("=" * 80)
    print("POD UPDATE COMPLETED")
    print("=" * 80)

    print(f"Rows checked             : {rows_checked:,}")
    print(f"POD matched              : {matched:,}")
    print(f"POD not found            : {missed:,}")
    print(f"Closed rows left alone   : {closed_skipped:,} (Cancelled / Virtual Invoice / OOS with a POD)")

    print("-" * 80)

    print(f"Exact GT + Order         : {exact_matches:,}")
    print(f"Normalized GT + Order    : {normalized_order_matches:,}")
    print("GT-only matching         : DISABLED")

    print("-" * 80)

    print(f"AJ marked Delivered      : {aj_updated:,}")
    print(f"AJ marked Dispatched     : {aj_by_stage[DISPATCHED_TEXT]:,} (AE AWB, no POD)")
    print(f"AJ marked Invoiced       : {aj_by_stage[INVOICED_TEXT]:,} (AW Invoice No, no POD / AWB)")
    print(f"AP updated               : {ap_updated:,}")
    print(f"AU links updated         : {au_updated:,}")
    print(f"AV updated               : {av_updated:,}")
    print(f"Total cells updated      : {updated_cells:,}")

    print("=" * 80)

    print()
    print("Updated Google Spreadsheet:")
    print(OUTPUT_SPREADSHEET_ID)
    print(f"Updated tab: {SHEET_NAME}")
    print()


# ============================================================
# RUN
# ============================================================

if __name__ == "__main__":

    main()
