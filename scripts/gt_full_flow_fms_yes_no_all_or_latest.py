import io
import json
import base64
import os
import re
import warnings
from datetime import datetime

import pandas as pd

# Google API packages:
# pip install google-api-python-client google-auth google-auth-httplib2 google-auth-oauthlib openpyxl
from google.oauth2 import service_account
from google.oauth2.credentials import Credentials as UserCredentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

warnings.filterwarnings("ignore", message="Workbook contains no default style")


# ============================================================
# CONFIGURATION
# ============================================================

CREDENTIALS_FOLDER = r"C:\Users\diwak\.gcloud"

# Source spreadsheet containing the FMS tab
SOURCE_SPREADSHEET_ID = "1Bk80DiH6deaXZaFdkHX-mdJZZDs30_HcCu96Kzp_AmY"
FMS_SHEET_NAME = "FMS"

# Destination spreadsheet containing Full Flow FMS
OUTPUT_SPREADSHEET_ID = "1jMEEFkAVfg7NHtxj3QCYuNY_V-AhU00ljQeRYTBQDio"
FULL_FLOW_SHEET_NAME = "Full Flow FMS"

# Warehouse code -> warehouse name mapping is maintained in this tab
# of the destination spreadsheet. Column A = Warehouse Code,
# Column B = Warehouse Name.
WAREHOUSE_MAPPING_SHEET_NAME = "Warehouse Name and Code"

# Order Enquiry is now a SINGLE Google Spreadsheet containing multiple tabs.
# Leave ORDER_ENQUIRY_SHEET_NAMES empty [] to automatically process ALL tabs.
# If you want only selected tabs, put their exact tab names in this list.
# Example: ["Apr-May -26", "Jun-26", "Jul-26"]
# New tabs added to the spreadsheet will automatically be included when this
# list is empty.
ORDER_ENQUIRY_SPREADSHEET_ID = "1Jnt0yU-caJgGm95hX5FSjPljUR_6ub__TA_G7mdB8yI"
ORDER_ENQUIRY_SHEET_NAMES = []

PICK_PACK_FOLDER_ID = "1Xi66o4A0I3GdtlNrh61XcEWNkT5UKmey"

# ------------------------------------------------------------
# ONE-TIME HISTORICAL CANCELLATION LOGIC
# ------------------------------------------------------------
# Put "YES" to mark FMS rows that are NOT found in Order Enquiry
# as Cancelled from column T through AW, only up to the cutoff date.
# Put "NO" to completely disable this logic.
#
# 2026-10-09: set to "NO". "No Vinculum match" is NOT evidence of cancellation —
# the match was exact-GT only, so GT Finals with an R/T/F approval suffix never
# matched and were stamped "Cancelled" (then skipped by every later fill). Same
# failure as GT_FMS_v2_DEPRECATED invoice2.js (2026-09-23 record). Rows already
# stamped are NOT reverted automatically; see report_suspect_cancellations().
# See "Errors and their fixes/2026-10-09-full-flow-rows-empty-suffix-and-refresh.md".
ENABLE_HISTORICAL_CANCELLATION = "NO"

# Change this date whenever you need a different historical cutoff.
# The cutoff is inclusive: 15-Aug-2026 is included.
CANCELLATION_CUTOFF_DATE = "2026-08-15"

# ------------------------------------------------------------
# ORDER ENQUIRY PROCESSING
# ------------------------------------------------------------
# The Order Enquiry source is now one Google Spreadsheet with multiple tabs.
# ORDER_ENQUIRY_SHEET_NAMES = [] means READ ALL TABS automatically.
# To restrict processing to specific tabs, enter their exact names above.
# Incremental GT protection still applies to the destination: existing GTs in
# "Full Flow FMS" are preserved and not rebuilt.

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]


# ============================================================
# FMS COLUMNS
# ============================================================

# SOURCE FMS layout (physical A:AW in the source sheet). This is read
# POSITIONALLY, so this list must stay in the source sheet's real order.
# Party Name / Contact Person / Phone / Location physically live at P:S
# (source indices 15-18) and must NOT be reordered here.
FMS_SOURCE_COLUMNS = [
    "Form Received Time", "Day of Date", "GT Number", "Vinculum SKU",
    "SKU Code", "Size", "MRP", "CSP", "% Off From MRP", "Order Quantity",
    "Distributor Discount", "Per Unit Billing Price", "Value", "Name",
    "E-Mail ID", "Party Name", "Contact Person", "Contact Person Phone Number",
    "Location", "Status", "Time Delay", "Warehouse Team", "Planned.1",
    "Actual.1", "Status.1", "Time Delay.1", "Planned.2", "Actual.2",
    "Status.2", "Time Delay.2", "AWB", "Planned.3", "Actual.3",
    "Status.3", "Time Delay.3", "Order Delivery Status", "Planned.4",
    "Actual.4", "Status.4", "Time Delay.4", "Planned.5", "Actual.5",
    "Status.5", "Time Delay.5", "Dispatched Qty", "Open Qty", "POD Copy",
    "Delivered Date", "Invoice No",
]

# Working set of columns held in the in-memory FMS dataframe: the 49 source
# columns plus the six output-only fields. Order here does NOT matter because
# every read/write selects by column NAME, not by position.
FMS_COLUMNS = FMS_SOURCE_COLUMNS + [
    "Warehouse", "Invoice Amount",
    "Order No", "Team", "Planned", "Actual",
]

# ------------------------------------------------------------
# OUTPUT (Full Flow FMS) LAYOUT — decoupled from the source read order.
# ------------------------------------------------------------
# Target layout:
#   A:O   = source head (Form Received Time .. E-Mail ID)
#   P:S   = workflow fields  -> Order No, Team, Planned, Actual
#   T:AW  = source tail      (Status .. Invoice No)
#   AX    = Warehouse
#   AY    = Invoice Amount
#   AZ:BC = customer details -> Party Name, Contact Person,
#           Contact Person Phone Number, Location
#
# The customer-detail block that physically sits at source P:S is moved to the
# tail (AZ:BC), and the four workflow fields take its place at output P:S.
FMS_SOURCE_HEAD = FMS_SOURCE_COLUMNS[:15]            # A:O
CUSTOMER_DETAIL_COLUMNS = FMS_SOURCE_COLUMNS[15:19]  # Party Name .. Location
FMS_SOURCE_TAIL = FMS_SOURCE_COLUMNS[19:]            # Status .. Invoice No (T:AW)

WORKFLOW_COLUMNS = ["Order No", "Team", "Planned", "Actual"]

OUTPUT_COLUMNS = (
    FMS_SOURCE_HEAD
    + WORKFLOW_COLUMNS
    + FMS_SOURCE_TAIL
    + ["Warehouse", "Invoice Amount"]
    + CUSTOMER_DETAIL_COLUMNS
)


# Exact Order Enquiry columns used by this automation.
# Matching key: ExtOrderNo (column C) + SKU Code (column W)
# Output mappings:
#   Order No   <- column A            -> Full Flow P
#   Team       <- Created by          -> Full Flow Q
#   Actual     <- Order Date          -> Full Flow S
#   Invoice No <- Invoice No          -> Full Flow AW
#   Invoice Amount <- Total           -> Full Flow AY
#   Dispatched Qty <- Pick Qty        -> Full Flow AS
#   Warehouse  <- Originating warehouse (code -> name) -> Full Flow AX
# NOTE: Order Quantity (Full Flow J) is intentionally NOT overwritten anymore;
# the FMS source value at J is preserved.
SALES_REQUIRED_COLUMNS = [
    "Order No", "ExtOrderNo", "Order Date", "SKU Code",
    "Created by", "Invoice No", "Originating warehouse", "Pick Qty", "Total"
]

PICK_PACK_REQUIRED_COLUMNS = ["Order No", "Pick Date", "Pick User"]


# ============================================================
# GOOGLE AUTHENTICATION
# ============================================================

def find_credentials_file():
    if not os.path.isdir(CREDENTIALS_FOLDER):
        raise FileNotFoundError(
            f"Google credentials folder not found:\n{CREDENTIALS_FOLDER}"
        )

    json_files = [
        os.path.join(CREDENTIALS_FOLDER, f)
        for f in os.listdir(CREDENTIALS_FOLDER)
        if f.lower().endswith(".json") and not f.startswith("~$")
    ]

    if not json_files:
        raise FileNotFoundError(
            f"No JSON credential file found in:\n{CREDENTIALS_FOLDER}"
        )

    # Prefer a service-account JSON key.
    for path in json_files:
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if data.get("type") == "service_account":
                return path
        except Exception:
            pass

    # Otherwise use the first JSON file. This also supports an OAuth
    # authorized_user credential file if one was created in .gcloud.
    return json_files[0]


def get_google_credentials():
    # GitHub Actions path: key comes from the GCP_SA_KEY environment variable
    raw = os.environ.get("GCP_SA_KEY")
    if raw:
        try:
            data = json.loads(raw)
        except json.JSONDecodeError:
            data = json.loads(base64.b64decode(raw))
        creds = service_account.Credentials.from_service_account_info(
            data, scopes=SCOPES
        )
        print("Using service-account credentials from GCP_SA_KEY env var")
        return creds

    # Local path: read the key from your PC folder (unchanged)
    credentials_file = find_credentials_file()

    with open(credentials_file, "r", encoding="utf-8") as f:
        data = json.load(f)

    if data.get("type") == "service_account":
        creds = service_account.Credentials.from_service_account_file(
            credentials_file,
            scopes=SCOPES,
        )
        print(f"Using service-account credentials: {os.path.basename(credentials_file)}")
        print(f"Service account: {creds.service_account_email}")
        return creds

    if data.get("type") == "authorized_user":
        creds = UserCredentials.from_authorized_user_file(
            credentials_file,
            scopes=SCOPES,
        )
        if creds.expired and creds.refresh_token:
            creds.refresh(Request())
        print(f"Using OAuth credentials: {os.path.basename(credentials_file)}")
        return creds

    raise ValueError(
        f"Unsupported Google credential JSON type in: {credentials_file}\n"
        f"Expected 'service_account' or 'authorized_user'."
    )


creds = get_google_credentials()
drive_service = build("drive", "v3", credentials=creds, cache_discovery=False)
sheets_service = build("sheets", "v4", credentials=creds, cache_discovery=False)


# ============================================================
# GOOGLE DRIVE HELPERS
# ============================================================

def list_drive_files(folder_id, extension=None):
    query = (
        f"'{folder_id}' in parents "
        "and trashed = false"
    )

    if extension:
        extension = extension.lower()
        query += f" and name contains '{extension}'"

    files = []
    page_token = None

    while True:
        response = drive_service.files().list(
            q=query,
            fields="nextPageToken, files(id,name,mimeType,modifiedTime,size)",
            pageSize=1000,
            pageToken=page_token,
            orderBy="modifiedTime desc",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ).execute()

        files.extend(response.get("files", []))
        page_token = response.get("nextPageToken")
        if not page_token:
            break

    if extension:
        files = [f for f in files if f["name"].lower().endswith(extension)]

    return files


def download_drive_file(file_id):
    return _download_request(drive_service.files().get_media(fileId=file_id))


def export_drive_file(file_id, mime_type):
    """Native Google file (e.g. a Google Sheet) exported to mime_type."""
    return _download_request(
        drive_service.files().export_media(fileId=file_id, mimeType=mime_type))


def _download_request(request):
    buffer = io.BytesIO()
    downloader = MediaIoBaseDownload(buffer, request)
    done = False

    while not done:
        _, done = downloader.next_chunk()

    buffer.seek(0)
    return buffer


# ============================================================
# GOOGLE SHEETS HELPERS
# ============================================================

def get_sheet_metadata(spreadsheet_id):
    return sheets_service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        fields="sheets(properties(sheetId,title,index,gridProperties(rowCount,columnCount)))",
    ).execute()


def get_sheet_by_title(spreadsheet_id, title):
    metadata = get_sheet_metadata(spreadsheet_id)
    for sheet in metadata.get("sheets", []):
        if sheet["properties"]["title"] == title:
            return sheet["properties"]
    return None


def read_fms_sheet():
    print("\nReading FMS tab from SOURCE Google Sheets...")

    # Row 7 is the header; data starts at row 8.
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=SOURCE_SPREADSHEET_ID,
        range=f"'{FMS_SHEET_NAME}'!A7:AW",
        valueRenderOption="UNFORMATTED_VALUE",
        dateTimeRenderOption="FORMATTED_STRING",
    ).execute()

    values = result.get("values", [])
    if not values:
        raise ValueError(
            f"No data found in source tab '{FMS_SHEET_NAME}'."
        )

    # Google Sheets keeps duplicate header names exactly as they appear
    # (e.g. Planned, Planned, Planned), while pandas.read_excel()
    # automatically makes them unique (Planned, Planned.1, Planned.2, ...).
    # Google Sheets API omits trailing empty cells from the returned header row.
    # The FMS source is structurally A:AW (49 columns), so pad the header row
    # to the expected width before applying positional fallback names.
    raw_headers = list(values[0])
    if len(raw_headers) < len(FMS_SOURCE_COLUMNS):
        raw_headers += [""] * (len(FMS_SOURCE_COLUMNS) - len(raw_headers))
    else:
        raw_headers = raw_headers[:len(FMS_SOURCE_COLUMNS)]

    headers = []
    header_counts = {}

    for position, header in enumerate(raw_headers):
        header = "" if header is None else str(header).strip()

        # Use the expected FMS column name by position when Google Sheets
        # returns an empty header cell.
        if header == "" and position < len(FMS_SOURCE_COLUMNS):
            header = FMS_SOURCE_COLUMNS[position]

        count = header_counts.get(header, 0)

        if count == 0:
            unique_header = header
        else:
            unique_header = f"{header}.{count}"

        header_counts[header] = count + 1
        headers.append(unique_header)

    rows = values[1:]

    # Google Sheets may omit trailing blank cells in each data row as well.
    # Normalize every row to the same 49-column A:AW source width.
    normalized_rows = []
    for row in rows:
        row = list(row)
        if len(row) < len(headers):
            row += [None] * (len(headers) - len(row))
        else:
            row = row[:len(headers)]
        normalized_rows.append(row)

    fms_df = pd.DataFrame(normalized_rows, columns=headers)
    fms_df = fms_df.dropna(how="all").reset_index(drop=True)

    print("\nFMS headers detected after duplicate-header normalization:")
    print(" | ".join(str(c) for c in headers))

    # Validate the actual source FMS A:AW structure by HEADER NAME.
    # Order No / Team / Planned / Actual are output-only fields and are
    # intentionally NOT expected in the source.
    missing = [c for c in FMS_SOURCE_COLUMNS if c not in fms_df.columns]
    if missing:
        raise ValueError(
            "Missing FMS source columns:\n" + "\n".join(missing)
        )

    # Keep the source columns exactly in A:AW order. This preserves
    # Party Name / Contact Person / Phone / Location at source P:S.
    fms_df = fms_df[FMS_SOURCE_COLUMNS].copy()

    # These are output-only fields and will be populated later.
    fms_df["Warehouse"] = ""
    fms_df["Invoice Amount"] = ""
    fms_df["Order No"] = ""
    fms_df["Team"] = ""
    fms_df["Planned"] = ""
    fms_df["Actual"] = ""
    print(f"FMS rows: {len(fms_df):,}")
    return fms_df


def get_output_sheet():
    """
    Find the existing 'Full Flow FMS' tab in the DESTINATION spreadsheet.
    The tab is expected to already exist.
    """
    print(
        f"\nChecking destination spreadsheet for "
        f"'{FULL_FLOW_SHEET_NAME}' tab..."
    )

    output_props = get_sheet_by_title(
        OUTPUT_SPREADSHEET_ID,
        FULL_FLOW_SHEET_NAME,
    )

    if not output_props:
        raise ValueError(
            f"Destination spreadsheet does not contain the "
            f"'{FULL_FLOW_SHEET_NAME}' tab. "
            f"Please create the tab first."
        )

    print(
        f"Found destination tab: '{FULL_FLOW_SHEET_NAME}' "
        f"(sheet ID: {output_props['sheetId']})"
    )

    return output_props["sheetId"]


def clear_full_flow_data(sheet_id):
    print(f"Clearing old data from '{FULL_FLOW_SHEET_NAME}'...")

    sheets_service.spreadsheets().values().clear(
        spreadsheetId=OUTPUT_SPREADSHEET_ID,
        range=f"'{FULL_FLOW_SHEET_NAME}'!A8:BC",
        body={},
    ).execute()


def read_warehouse_mapping():
    """Read Warehouse Code -> Warehouse Name mapping from the destination tab."""
    print(f"\nReading warehouse code/name mapping from '{WAREHOUSE_MAPPING_SHEET_NAME}'...")

    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=OUTPUT_SPREADSHEET_ID,
        range=f"'{WAREHOUSE_MAPPING_SHEET_NAME}'!A:B",
        valueRenderOption="UNFORMATTED_VALUE",
    ).execute()

    values = result.get("values", [])
    if not values:
        raise ValueError(
            f"No warehouse mapping data found in '{WAREHOUSE_MAPPING_SHEET_NAME}'."
        )

    mapping = {}
    for row_number, row in enumerate(values[1:], start=2):
        code = row[0] if len(row) > 0 else ""
        name = row[1] if len(row) > 1 else ""
        code_key = normalize_value(code)
        if not code_key:
            continue
        mapping[code_key] = str(name).strip() if value_found(name) else ""

    if not mapping:
        raise ValueError(
            f"No valid warehouse code/name mappings found in '{WAREHOUSE_MAPPING_SHEET_NAME}'."
        )

    print(f"Warehouse mappings loaded: {len(mapping):,}")
    return mapping


def clean_for_google(value):
    if value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass

    if isinstance(value, pd.Timestamp):
        if pd.isna(value):
            return ""
        return value.strftime("%d-%b-%Y %H:%M:%S")

    if isinstance(value, datetime):
        return value.strftime("%d-%b-%Y %H:%M:%S")

    # Avoid numpy scalar objects in the Sheets JSON payload.
    if hasattr(value, "item"):
        try:
            value = value.item()
        except Exception:
            pass

    return value


def read_existing_full_flow():
    """Read existing destination rows so already-processed GTs are preserved."""
    print(f"\nReading existing '{FULL_FLOW_SHEET_NAME}' data...")

    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=OUTPUT_SPREADSHEET_ID,
        range=f"'{FULL_FLOW_SHEET_NAME}'!A8:BC",
        valueRenderOption="UNFORMATTED_VALUE",
        dateTimeRenderOption="FORMATTED_STRING",
    ).execute()

    values = result.get("values", [])
    if not values:
        print("  No existing Full Flow data found.")
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    normalized_rows = []
    for row in values:
        row = list(row)
        if len(row) < len(OUTPUT_COLUMNS):
            row += [""] * (len(OUTPUT_COLUMNS) - len(row))
        normalized_rows.append(row[:len(OUTPUT_COLUMNS)])

    existing_df = pd.DataFrame(normalized_rows, columns=OUTPUT_COLUMNS)
    existing_df = existing_df.dropna(how="all").reset_index(drop=True)

    print(f"  Existing Full Flow rows: {len(existing_df):,}")
    return existing_df


def write_full_flow(sheet_id, df, start_row=8):
    print(
        f"Writing Full Flow data to destination "
        f"'{FULL_FLOW_SHEET_NAME}'..."
    )
    print("  Output layout: A:O=FMS head | P:S=Order No/Team/Planned/Actual | "
          "T:AW=FMS tail | AX=Warehouse | AY=Invoice Amount | "
          "AZ:BC=Party Name/Contact Person/Phone/Location")

    def column_letter(n):
        result = ""
        while n:
            n, remainder = divmod(n - 1, 26)
            result = chr(65 + remainder) + result
        return result

    open_qty_index = OUTPUT_COLUMNS.index("Open Qty")
    order_qty_column = column_letter(OUTPUT_COLUMNS.index("Order Quantity") + 1)
    dispatched_qty_column = column_letter(OUTPUT_COLUMNS.index("Dispatched Qty") + 1)

    values = []
    for local_index, row in enumerate(df.itertuples(index=False, name=None)):
        output_row = [clean_for_google(value) for value in row]

        # Open Qty = Order Quantity (J) - Dispatched Qty (AS).
        # Use column indexes from OUTPUT_COLUMNS so this remains correct if
        # the output layout changes again.
        sheet_row = start_row + local_index

        output_row[open_qty_index] = (
            f'=IF(OR({order_qty_column}{sheet_row}="Cancelled",'
            f'{dispatched_qty_column}{sheet_row}="Cancelled"),'
            f'"Cancelled",IF({order_qty_column}{sheet_row}="","",'
            f'N({order_qty_column}{sheet_row})-N({dispatched_qty_column}{sheet_row})))'
        )

        values.append(output_row)

    if not values:
        print("No new Full Flow rows to write.")
        return

    # Expand the destination grid only when the new data requires it.
    required_rows = start_row - 1 + len(values)
    sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=OUTPUT_SPREADSHEET_ID,
        body={
            "requests": [
                {
                    "updateSheetProperties": {
                        "properties": {
                            "sheetId": sheet_id,
                            "gridProperties": {
                                "rowCount": required_rows,
                                "columnCount": len(OUTPUT_COLUMNS),
                            },
                        },
                        "fields": "gridProperties.rowCount,gridProperties.columnCount",
                    }
                }
            ]
        },
    ).execute()

    print(
        f"  Destination grid ready: {required_rows:,} rows x {len(OUTPUT_COLUMNS)} columns."
    )

    chunk_size = 500
    end_column = column_letter(len(OUTPUT_COLUMNS))

    for start in range(0, len(values), chunk_size):
        chunk = values[start:start + chunk_size]
        write_start_row = start_row + start
        end_row = write_start_row + len(chunk) - 1
        range_name = (
            f"'{FULL_FLOW_SHEET_NAME}'!A{write_start_row}:{end_column}{end_row}"
        )

        sheets_service.spreadsheets().values().update(
            spreadsheetId=OUTPUT_SPREADSHEET_ID,
            range=range_name,
            valueInputOption="USER_ENTERED",
            body={"values": chunk},
        ).execute()

        print(f"  Written rows {write_start_row:,} - {end_row:,}")

    # Apply date formatting to the newly written rows.
    # "Actual" here is the workflow field at S (Order Date). "Actual.1" is the
    # source milestone at X (Pick Date bridge).
    requests = []
    date_columns = ["Form Received Time", "Actual", "Actual.1", "Delivered Date"]
    date_col_indices = [OUTPUT_COLUMNS.index(c) for c in date_columns]
    for col_index in date_col_indices:
        requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": start_row - 1,
                    "endRowIndex": start_row - 1 + len(df),
                    "startColumnIndex": col_index,
                    "endColumnIndex": col_index + 1,
                },
                "cell": {
                    "userEnteredFormat": {
                        "numberFormat": {
                            "type": "DATE_TIME",
                            "pattern": "dd-mmm-yyyy hh:mm:ss",
                        }
                    }
                },
                "fields": "userEnteredFormat.numberFormat",
            }
        })

    sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=OUTPUT_SPREADSHEET_ID,
        body={"requests": requests},
    ).execute()


def normalize_value(value):
    if value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass

    value = str(value)

    # Remove HTML tags such as <br>, <br/>, <div>, etc.
    # Sales / Order Enquiry exports can contain HTML line-break tags.
    value = re.sub(r"<[^>]*>", "", value)

    # Remove non-breaking spaces and invisible BOM characters.
    value = value.replace("\xa0", " ")
    value = value.replace("\ufeff", "")

    # Remove quotes.
    value = value.replace('"', "").replace("'", "")

    # Remove all whitespace.
    value = value.strip()
    value = re.sub(r"\s+", "", value)

    if value.endswith(".0"):
        numeric_part = value[:-2]
        if numeric_part.isdigit():
            value = numeric_part

    return value.upper()


def create_match_key(order_no, sku):
    """Create a stable GT/Order + SKU matching key."""
    return normalize_value(order_no) + "||" + normalize_value(sku)


def value_found(value):
    if value is None:
        return False
    try:
        if pd.isna(value):
            return False
    except Exception:
        pass
    return str(value).replace("\xa0", " ").strip() != ""


# ============================================================
# PICK DATE PARSER
# ============================================================

def parse_pick_date(value):
    if value is None:
        return pd.NaT

    try:
        if pd.isna(value):
            return pd.NaT
    except Exception:
        pass

    value = str(value)
    value = value.replace("\xa0", " ")
    value = value.replace("\ufeff", "")
    value = value.replace('"', "").replace("'", "")
    value = value.strip()

    if not value:
        return pd.NaT

    # Excel serial date stored as text.
    if re.match(r"^\d+(\.\d+)?$", value):
        try:
            return pd.Timestamp("1899-12-30") + pd.Timedelta(days=float(value))
        except Exception:
            return pd.NaT

    match = re.match(
        r"^\s*(\d{1,2})[-/](\d{1,2})[-/](\d{4})"
        r"\s+(\d{1,2}):(\d{2})"
        r"(?:[\s:]+(\d{1,2}))?\s*$",
        value,
    )

    if match:
        try:
            return pd.Timestamp(
                year=int(match.group(3)),
                month=int(match.group(2)),
                day=int(match.group(1)),
                hour=int(match.group(4)),
                minute=int(match.group(5)),
                second=int(match.group(6) or 0),
            )
        except Exception:
            return pd.NaT

    # ISO yyyy-mm-dd[ hh:mm[:ss]] — the form real date cells take when a Google Sheet is
    # exported to xlsx and read as text, and the form of the May/June 2026 sheets
    # ("2026-05-02 9:34:53"). Must be handled BEFORE the day-first fallback below:
    # pd.to_datetime("2026-05-02 9:34:53", dayfirst=True) returns 2026-02-05.
    match = re.match(
        r"^(\d{4})-(\d{1,2})-(\d{1,2})"
        r"(?:[ T](\d{1,2}):(\d{2})(?::(\d{1,2}))?(?:\.\d+)?)?$",
        value,
    )
    if match:
        try:
            return pd.Timestamp(
                year=int(match.group(1)),
                month=int(match.group(2)),
                day=int(match.group(3)),
                hour=int(match.group(4) or 0),
                minute=int(match.group(5) or 0),
                second=int(match.group(6) or 0),
            )
        except Exception:
            return pd.NaT

    # Anything else: day-first best effort; a value that cannot be read is blank
    # (that row simply gets no X), it never aborts the run.
    try:
        return pd.to_datetime(value, dayfirst=True, errors="coerce")
    except Exception:
        return pd.NaT


# ============================================================
# SALES / ORDER ENQUIRY FROM GOOGLE DRIVE
# ============================================================

def normalize_sales_header(header):
    """Normalize an Order Enquiry header without changing its meaning."""
    if header is None:
        return ""

    header = str(header)
    header = header.replace("\ufeff", "").replace("\xa0", " ").strip()
    header = re.sub(r"\s+", " ", header)

    aliases = {
        "Ext OrderNo": "ExtOrderNo",
        "Ext Order No": "ExtOrderNo",
        "ExternOrderNo": "ExtOrderNo",
        "Extern OrderNo": "ExtOrderNo",
        "Extern Order No": "ExtOrderNo",
        "Order Date": "Order Date",
        "OrderDate": "Order Date",
        "InvoiceNo": "Invoice No",
        "Invoice Number": "Invoice No",
        "Order Quantity": "Orderqty",
        "Order Qty": "Orderqty",
        "OrderQty": "Orderqty",
        "Created By": "Created by",
    }

    return aliases.get(header, header)


def read_order_enquiry_tab(spreadsheet_id, tab_name):
    """Read one Order Enquiry tab from the configured Google Spreadsheet."""
    print(f"  Reading Order Enquiry tab: {tab_name}")

    response = sheets_service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f"'{tab_name}'",
        valueRenderOption="UNFORMATTED_VALUE",
        dateTimeRenderOption="FORMATTED_STRING",
    ).execute()

    values = response.get("values", [])
    if not values:
        print(f"    WARNING - tab '{tab_name}' is empty.")
        return pd.DataFrame()

    headers = [normalize_sales_header(h) for h in values[0]]
    rows = values[1:]

    # Remove completely blank header columns, but keep all actual data columns.
    width = len(headers)
    normalized_rows = [
        list(row) + [""] * max(0, width - len(row))
        for row in rows
    ]
    normalized_rows = [row[:width] for row in normalized_rows]

    df = pd.DataFrame(normalized_rows, columns=headers)
    df = df.dropna(how="all").reset_index(drop=True)

    # If duplicate headers exist, pandas would otherwise create ambiguity.
    # The required headers are expected to be unique in the Order Enquiry export.
    return df


def get_order_enquiry_tab_names():
    """Return the tabs to process; empty configuration means all tabs."""
    metadata = sheets_service.spreadsheets().get(
        spreadsheetId=ORDER_ENQUIRY_SPREADSHEET_ID,
        fields="sheets.properties(title,index)",
    ).execute()

    all_tabs = [
        s["properties"]["title"]
        for s in metadata.get("sheets", [])
        if s.get("properties", {}).get("title") is not None
    ]

    if not all_tabs:
        raise ValueError(
            f"Order Enquiry spreadsheet {ORDER_ENQUIRY_SPREADSHEET_ID} has no tabs."
        )

    configured = [str(x).strip() for x in ORDER_ENQUIRY_SHEET_NAMES if str(x).strip()]

    if not configured:
        print("\nORDER ENQUIRY TABS: ALL")
        print("Reading every tab in the Order Enquiry spreadsheet.")
        return all_tabs

    missing = [name for name in configured if name not in all_tabs]
    if missing:
        raise ValueError(
            "These configured Order Enquiry tabs were not found:\n"
            + "\n".join(f"  - {name}" for name in missing)
            + "\n\nAvailable tabs:\n"
            + "\n".join(f"  - {name}" for name in all_tabs)
        )

    print("\nORDER ENQUIRY TABS: SELECTED")
    print("Reading only these tabs:")
    for name in configured:
        print(f"  - {name}")
    return configured


def read_sales_file():
    """
    Read Order Enquiry data from one Google Spreadsheet containing multiple tabs.

    By default ORDER_ENQUIRY_SHEET_NAMES = [] reads ALL tabs, including tabs
    added in the future. If a list is provided, only those exact tab names are read.
    """
    tab_names = get_order_enquiry_tab_names()

    print("\n" + "=" * 80)
    print("ORDER ENQUIRY SOURCE")
    print("=" * 80)
    print(f"Spreadsheet ID: {ORDER_ENQUIRY_SPREADSHEET_ID}")
    print(f"Tabs selected:  {len(tab_names):,}")

    sales_dataframes = []
    included_tabs = []
    skipped_tabs = []

    for tab_name in tab_names:
        print("\nChecking Order Enquiry tab:")
        print(f"  {tab_name}")

        try:
            df = read_order_enquiry_tab(
                ORDER_ENQUIRY_SPREADSHEET_ID,
                tab_name,
            )

            if df.empty:
                skipped_tabs.append(tab_name)
                continue

            missing = [
                column
                for column in SALES_REQUIRED_COLUMNS
                if column not in df.columns
            ]

            if missing:
                print("  SKIPPED - missing required columns:")
                print("    " + " | ".join(missing))
                print(
                    "    Available columns: "
                    + " | ".join(str(c) for c in df.columns)
                )
                skipped_tabs.append(tab_name)
                continue

            df["_TAB"] = tab_name   # used to drop rows duplicated across month tabs
            sales_dataframes.append(df)
            included_tabs.append(tab_name)
            print(f"  INCLUDED - {len(df):,} rows")

        except Exception as exc:
            print(f"  ERROR reading tab '{tab_name}': {exc}")
            skipped_tabs.append(tab_name)

    if not sales_dataframes:
        raise ValueError(
            "None of the selected Order Enquiry tabs contains all required "
            "columns for Full Flow processing.\n\n"
            "Required columns:\n"
            + "\n".join(SALES_REQUIRED_COLUMNS)
        )

    sales_df = pd.concat(sales_dataframes, ignore_index=True)
    sales_df = sales_df.dropna(how="all").reset_index(drop=True)

    print("\n" + "=" * 80)
    print("ORDER ENQUIRY TAB SUMMARY")
    print("=" * 80)
    print(f"Tabs available:              {len(tab_names):,}")
    print(f"Tabs included:               {len(included_tabs):,}")
    print(f"Tabs skipped:                {len(skipped_tabs):,}")
    print(f"Combined Order Enquiry rows: {len(sales_df):,}")

    if included_tabs:
        print("\nIncluded tabs:")
        for name in included_tabs:
            print(f"  ✓ {name}")

    if skipped_tabs:
        print("\nSkipped tabs:")
        for name in skipped_tabs:
            print(f"  ✗ {name}")

    return sales_df


# Pick Pack file kinds. 2026-10-10: since May 2026 the daily reports arrive in the folder
# as native Google Sheets ("PickPackReport_08.10.2026", no extension); the old reader kept
# only names ending ".csv", so it saw April only and V/X stayed blank from May on.
GOOGLE_SHEET_MIME = "application/vnd.google-apps.spreadsheet"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _pick_pack_kind(file_info):
    """"csv" / "gsheet" / "xlsx", or None for a file that is not a Pick Pack report."""
    name = str(file_info.get("name", "")).lower()
    mime = file_info.get("mimeType", "")
    if mime == GOOGLE_SHEET_MIME:
        return "gsheet"
    if name.endswith(".csv") or mime == "text/csv":
        return "csv"
    if name.endswith((".xlsx", ".xlsm")) or mime == XLSX_MIME:
        return "xlsx"
    return None


def _clean_pick_pack_columns(df):
    # Header cells can carry a BOM, NBSPs and - in Sheets made by importing a CSV - the
    # literal quotes of the CSV header ('\ufeff"Order No"').
    df.columns = (
        df.columns.astype(str)
        .str.replace("\ufeff", "", regex=False)
        .str.replace("\xa0", " ", regex=False)
        .str.strip()
        .str.strip("\"'")
        .str.strip()
    )
    return df


def _read_pick_pack_csv(raw_bytes):
    for encoding in ["utf-8-sig", "cp1252", "latin1"]:
        try:
            return _clean_pick_pack_columns(pd.read_csv(
                io.BytesIO(raw_bytes),
                encoding=encoding,
                dtype=str,
                keep_default_na=False,
                low_memory=False,
            ))
        except Exception:
            pass
    raise ValueError("not readable as CSV (utf-8 / cp1252 / latin1)")


def _read_pick_pack_workbook(raw_bytes):
    """Every tab that has the Pick Pack columns (a Google Sheet export may have several
    tabs and the data tab is not always first). Real date cells come back as ISO text,
    which parse_pick_date reads year-first. With no such tab, the first tab is returned
    so the caller can report which columns it has."""
    tabs = pd.read_excel(io.BytesIO(raw_bytes), sheet_name=None, dtype=str,
                         keep_default_na=False, engine="openpyxl")
    tabs = {title: _clean_pick_pack_columns(df) for title, df in tabs.items()}
    good = [df for df in tabs.values()
            if all(c in df.columns for c in PICK_PACK_REQUIRED_COLUMNS)]
    if good:
        return pd.concat(good, ignore_index=True)
    return next(iter(tabs.values())) if tabs else pd.DataFrame()


def read_pick_pack_files():
    print("\nLooking for Pick Pack files (CSV / Google Sheet / Excel) in Google Drive...")

    files = list_drive_files(PICK_PACK_FOLDER_ID)
    candidates = [(f, _pick_pack_kind(f)) for f in files]
    ignored = [f for f, kind in candidates if kind is None]
    candidates = [(f, kind) for f, kind in candidates if kind]
    by_kind = {k: sum(1 for _, kind in candidates if kind == k) for k in ("csv", "gsheet", "xlsx")}

    print(f"Pick Pack files found: {len(candidates):,}  (csv {by_kind['csv']} | "
          f"google sheet {by_kind['gsheet']} | xlsx {by_kind['xlsx']}); "
          f"other files ignored: {len(ignored)}")
    for f in ignored:
        print(f"  ignored: {f.get('name')} ({f.get('mimeType')})")

    dataframes, skipped = [], []

    for file_info, kind in candidates:
        name = file_info.get("name", file_info.get("id"))
        try:
            if kind == "gsheet":
                df = _read_pick_pack_workbook(export_drive_file(file_info["id"], XLSX_MIME).getvalue())
            elif kind == "xlsx":
                df = _read_pick_pack_workbook(download_drive_file(file_info["id"]).getvalue())
            else:
                df = _read_pick_pack_csv(download_drive_file(file_info["id"]).getvalue())
        except Exception as exc:
            # One bad / unshared file must not cost the whole night.
            print(f"WARNING: Pick Pack file not readable, skipped: {name} [{kind}] "
                  f"-> {type(exc).__name__}: {exc}")
            skipped.append(name)
            continue

        missing = [c for c in PICK_PACK_REQUIRED_COLUMNS if c not in df.columns]
        if missing:
            print(f"WARNING: Pick Pack file skipped, missing columns {missing}: {name} [{kind}] "
                  f"(has: {list(df.columns)[:12]})")
            skipped.append(name)
            continue

        df = df.dropna(how="all")
        dataframes.append(df)
        print(f"Read: {name} [{kind}] -> {len(df):,} rows")

    if skipped:
        print(f"Pick Pack files skipped: {len(skipped)} of {len(candidates)}")

    if not dataframes:
        print("WARNING: no Pick Pack file could be read - Warehouse Team (V) / "
              "Actual.1 (X) will not be filled this run.")
        return pd.DataFrame(columns=PICK_PACK_REQUIRED_COLUMNS)

    pick_pack_df = pd.concat(dataframes, ignore_index=True)
    pick_pack_df = pick_pack_df.dropna(how="all").reset_index(drop=True)

    print(f"Pick Pack files read: {len(dataframes):,}")
    print(f"Total Pick Pack rows: {len(pick_pack_df):,}")
    return pick_pack_df


# ============================================================
# OPTIONAL HISTORICAL CANCELLATION
# ============================================================

def apply_historical_cancellation(fms_df, sales_df):
    """
    If enabled, mark FMS rows that have no Order Enquiry match as
    'Cancelled' from source T through AW (Status .. Invoice No), only when
    Form Received Time is on or before CANCELLATION_CUTOFF_DATE.

    The customer-detail block (source P:S = Party Name .. Location) is NEVER
    overwritten, because those values are relocated to output AZ:BC and must
    survive. The workflow fields (Order No / Team / Planned / Actual) are added
    separately and are not touched here.

    If disabled, return the FMS data unchanged.
    """
    if str(ENABLE_HISTORICAL_CANCELLATION).strip().upper() != "YES":
        print("\nHistorical cancellation logic: DISABLED (NO)")
        return fms_df.copy()

    print("\nHistorical cancellation logic: ENABLED (YES)")
    print(f"Cancellation cutoff date: {CANCELLATION_CUTOFF_DATE}")

    working = fms_df.copy()

    working["Form Received Time"] = pd.to_datetime(
        working["Form Received Time"],
        errors="coerce",
        dayfirst=True,
    )

    # Build the same GT + SKU key used by the normal Sales matching.
    working["_CANCEL_MATCH_KEY"] = working.apply(
        lambda row: create_match_key(row["GT Number"], row["Vinculum SKU"]),
        axis=1,
    )

    sales_keys = set(
        sales_df.apply(
            lambda row: create_match_key(
                row["ExtOrderNo"],
                row["SKU Code"],
            ),
            axis=1,
        )
    )

    cutoff = pd.Timestamp(CANCELLATION_CUTOFF_DATE).normalize()
    eligible_date = (
        working["Form Received Time"].notna()
        & (working["Form Received Time"].dt.normalize() <= cutoff)
    )

    missing_sales = ~working["_CANCEL_MATCH_KEY"].isin(sales_keys)
    cancel_mask = eligible_date & missing_sales

    cancel_count = int(cancel_mask.sum())

    # Source T:AW = Status .. Invoice No (source indices 19-48). This preserves
    # the customer-detail block at source indices 15-18.
    cancel_columns = FMS_SOURCE_COLUMNS[19:49]

    for column in cancel_columns:
        working.loc[cancel_mask, column] = "Cancelled"

    working.drop(columns=["_CANCEL_MATCH_KEY"], inplace=True)

    print(f"Historical rows marked Cancelled: {cancel_count:,}")
    print(f"Source columns affected: T:AW ({len(cancel_columns)} columns)")

    return working


# ============================================================
# BUILD FULL FLOW
# ============================================================

def build_sales_lookup(sales_df):
    """Prepare Order Enquiry rows (parse Order Date / Pick Qty / Total, add the GT+SKU
    match key) and group them by key. Mutates and returns sales_df, like before."""
    sales_df["_MATCH_KEY"] = sales_df.apply(
        lambda row: create_match_key(row["ExtOrderNo"], row["SKU Code"]),
        axis=1,
    )

    # Order Enquiry Order Date is yyyy-mm-dd hh:mm:ss. Parse it explicitly.
    sales_df["Order Date"] = pd.to_datetime(
        sales_df["Order Date"],
        format="%Y-%m-%d %H:%M:%S",
        errors="coerce",
    )

    # Fallback for any rows that contain another valid timestamp format.
    invalid_order_dates = sales_df["Order Date"].isna()
    if invalid_order_dates.any():
        sales_df.loc[invalid_order_dates, "Order Date"] = pd.to_datetime(
            sales_df.loc[invalid_order_dates, "Order Date"],
            errors="coerce",
            dayfirst=False,
        )

    # NOTE: Orderqty is intentionally NOT used. Order Quantity (Full Flow J)
    # keeps its FMS source value and is never overwritten from Order Enquiry.

    # Sales Pick Qty -> Full Flow Dispatched Qty (AS).
    sales_df["Pick Qty"] = pd.to_numeric(
        sales_df["Pick Qty"], errors="coerce"
    )

    # Sales Total -> Full Flow Invoice Amount (AY).
    sales_df["Total"] = pd.to_numeric(
        sales_df["Total"], errors="coerce"
    )

    sales_groups = {
        key: group
        for key, group in sales_df.groupby("_MATCH_KEY", sort=False)
    }
    return sales_groups, sales_df


def build_pick_pack_lookup(pick_pack_df):
    """Order No -> latest valid Pick Date + corresponding Pick User."""
    print("Parsing Pick Pack Pick Date...")
    pick_pack_df["_RAW_PICK_DATE"] = pick_pack_df["Pick Date"]
    pick_pack_df["Pick Date"] = pick_pack_df["Pick Date"].apply(parse_pick_date)

    valid_pick_dates = pick_pack_df["Pick Date"].notna().sum()
    print(f"Valid Pick Dates: {valid_pick_dates:,}")
    print(f"Invalid/blank Pick Dates: {len(pick_pack_df) - valid_pick_dates:,}")

    if valid_pick_dates == 0:
        raw = pick_pack_df["_RAW_PICK_DATE"].astype(str).str.strip()
        print("WARNING: 0 Pick Dates parsed. Raw samples:")
        for value in raw[raw != ""].head(15):
            print(f"  {value!r}")
        raise SystemExit("Pick Date parsing produced 0 valid dates.")

    # Pick Pack lookup - latest valid Pick Date + corresponding user.
    pick_pack_df["_NORMALIZED_ORDER_NO"] = pick_pack_df["Order No"].map(normalize_value)
    valid_pick_orders = pick_pack_df[
        pick_pack_df["_NORMALIZED_ORDER_NO"] != ""
    ].copy()

    pick_pack_lookup = {}
    for order_no, group in valid_pick_orders.groupby(
        "_NORMALIZED_ORDER_NO", sort=False
    ):
        valid_dates = group[group["Pick Date"].notna()].sort_values("Pick Date")

        if not valid_dates.empty:
            latest_row = valid_dates.iloc[-1]
            selected_pick_date = latest_row["Pick Date"]
            selected_pick_user = (
                latest_row["Pick User"]
                if value_found(latest_row["Pick User"])
                else None
            )
        else:
            selected_pick_date = None
            selected_pick_user = None

        if selected_pick_user is None:
            for _, row in group.iterrows():
                if value_found(row["Pick User"]):
                    selected_pick_user = row["Pick User"]
                    break

        pick_pack_lookup[order_no] = {
            "Pick Date": selected_pick_date,
            "Pick User": selected_pick_user,
        }
    return pick_pack_lookup


def build_full_flow(fms_df, assigned, pick_pack_lookup, warehouse_mapping):
    """Build output rows for FMS rows being appended.

    assigned[pos] = list of Order Enquiry lines (Series) that allocate_lines() gave
    to the pos-th fms_df row ([] = no Vinculum line yet; the row is appended with a
    blank Order No and the nightly fill picks it up later). One output row per line,
    so a GT+SKU punched as several Vinculum orders expands to several rows (as before).
    """
    print("\nBuilding Full Flow rows to append...")

    fms_df = fms_df.copy()
    fms_df["Form Received Time"] = pd.to_datetime(
        fms_df["Form Received Time"], errors="coerce", dayfirst=True
    )

    full_flow_rows = []
    matched_fms_count = 0
    unmatched_fms_count = 0
    expanded_fms_count = 0
    pick_pack_match_count = 0
    pick_pack_actual_count = 0
    pick_pack_user_count = 0
    warehouse_match_count = 0
    warehouse_missing_count = 0

    for pos, (_, fms_row) in enumerate(fms_df.iterrows()):
        base_row = {
            column: fms_row[column] if column in fms_row.index else ""
            for column in OUTPUT_COLUMNS
        }
        matched_sales_rows = assigned[pos] if pos < len(assigned) else []

        if matched_sales_rows:
            matched_fms_count += 1
            if len(matched_sales_rows) > 1:
                expanded_fms_count += 1
            outputs = []
            for sales_row in matched_sales_rows:
                output_row = base_row.copy()

                # Order Enquiry -> Full Flow mappings
                output_row["Order No"] = sales_row["Order No"]        # P
                output_row["Team"] = sales_row["Created by"]          # Q
                output_row["Actual"] = sales_row["Order Date"]        # S
                output_row["Invoice No"] = sales_row["Invoice No"]    # AW
                output_row["Invoice Amount"] = sales_row["Total"]     # AY
                # Order Quantity (J) is deliberately left as the FMS source value.

                warehouse_code = sales_row["Originating warehouse"]
                warehouse_key = normalize_value(warehouse_code)
                if warehouse_key and warehouse_key in warehouse_mapping:
                    output_row["Warehouse"] = warehouse_mapping[warehouse_key]
                    warehouse_match_count += 1
                elif warehouse_key:
                    output_row["Warehouse"] = str(warehouse_code).strip()
                    warehouse_missing_count += 1

                # Pick Qty 0 = not picked yet: leave AS blank so a later night can
                # fill the real quantity (the fill never overwrites a non-blank 0).
                dispatched_qty = _dispatched_qty(sales_row)
                if dispatched_qty is not None:
                    output_row["Dispatched Qty"] = dispatched_qty
                outputs.append(output_row)
        else:
            unmatched_fms_count += 1
            outputs = [base_row.copy()]

        for output_row in outputs:
            fms_order_no_key = normalize_value(output_row["Order No"])
            if fms_order_no_key and fms_order_no_key in pick_pack_lookup:
                pick_pack_match_count += 1
                pick_data = pick_pack_lookup[fms_order_no_key]

                pick_date = pick_data["Pick Date"]
                if pick_date is not None and not pd.isna(pick_date):
                    output_row["Actual.1"] = pd.Timestamp(pick_date).to_pydatetime()
                    pick_pack_actual_count += 1

                if (
                    not value_found(output_row["Warehouse Team"])
                    and value_found(pick_data["Pick User"])
                ):
                    output_row["Warehouse Team"] = pick_data["Pick User"]
                    pick_pack_user_count += 1
            full_flow_rows.append(output_row)

    full_flow_df = pd.DataFrame(full_flow_rows, columns=OUTPUT_COLUMNS)

    print("\n" + "=" * 80)
    print("PROCESSING SUMMARY (rows to append)")
    print("=" * 80)
    print(f"FMS rows to append:            {len(fms_df):,}")
    print(f"FMS rows with Sales match:     {matched_fms_count:,}")
    print(f"FMS rows without Sales match:  {unmatched_fms_count:,}")
    print(f"FMS rows expanded:             {expanded_fms_count:,}")
    print(f"Final Full Flow rows:          {len(full_flow_df):,}")
    print(f"Pick Pack matches:             {pick_pack_match_count:,}")
    print(f"Pick Dates -> Actual.1 (X):    {pick_pack_actual_count:,}")
    print(f"Pick Users -> Warehouse Team (V): {pick_pack_user_count:,}")
    print(f"Warehouse codes mapped -> AX:  {warehouse_match_count:,}")
    print(f"Warehouse codes unmapped:      {warehouse_missing_count:,}")

    return full_flow_df


# ============================================================
# NIGHTLY SYNC: suffix-tolerant matching, fill + refresh, new lines
# ============================================================
# History:
#  - 2026-10-05: backfill of rows with a BLANK Order No (exact GT + SKU match).
#  - 2026-10-09: "a lot of rows are empty". Three gaps closed here:
#      1. Match key was the exact GT Number. FMS "GT Number" is the order's GT Final,
#         which carries an R/T/F approval suffix on mixed orders (GT081026A1R / …T / …F,
#         GeneralTrade_FMS Order_Entry.js needsSuffix); Ops punch ExtOrderNo with or
#         without it. Every old v2 script stripped it on BOTH sides
#         (GT_FMS_v2_DEPRECATED/orderidandteam.js normalizeOrderId) and pod.py does too.
#         Now: exact GT first, then GT with ONE trailing R/T/F stripped (both sides),
#         then the alternate SKU fields (FMS SKU Code / Vinculum Bundle SKU Code) —
#         always within the SAME GT base, never by prefix (GT…A1 never meets GT…A10).
#      2. Only rows with a blank Order No were ever revisited, so pick qty / invoice no /
#         invoice amount that reach Vinculum after the punch never landed. Now every
#         row whose Order No matches a Vinculum line gets its BLANK cells of
#         P/Q/S/AS/AW/AX/AY/X/V filled every night (never overwritten).
#      3. Processing was append-only per GT, so a new SKU line for a GT already in
#         Full Flow never arrived. Now any FMS GT+SKU not yet in Full Flow is appended.
#    Cancelled / Virtual Invoice / OOS rows are never written.
#  - 2026-10-10: V (Warehouse Team) / X (Actual.1) blank from May 2026 on. Pick Pack
#    reports have arrived as native Google Sheets since May and read_pick_pack_files
#    kept only "*.csv" names. It now reads CSV + Google Sheet (xlsx export) + xlsx,
#    skips a bad file with a warning, parses ISO Pick Dates year-first, and also fills
#    V/X on rows whose P is not in Order Enquiry (keyed by P). Record:
#    "Errors and their fixes/2026-10-10-pick-pack-google-sheets-skipped.md".
#
# BACKFILL_MODE env: "write" (default in CI), "dry-run" (log only, this script writes
# NOTHING — no fills, no appends), "off" (no fills; appends as before).

BACKFILL_MODE = os.environ.get("BACKFILL_MODE", "dry-run").strip().lower()

# Safety net for gap 3: if this many FMS lines suddenly look "missing" from GTs that
# are already in Full Flow, something structural changed (e.g. SKU codes rewritten);
# refuse to append them and say so, rather than duplicate hundreds of rows.
NEW_LINES_FOR_EXISTING_GTS_MAX = 300

# Values that mark a line as closed by a human / earlier job. Such rows are never written.
CLOSED_MARKERS = {"cancelled", "virtual invoice", "out of stock", "oos"}

# Output columns in which a closed marker is looked for: J (Order Quantity) and P:AW.
_P_INDEX = OUTPUT_COLUMNS.index("Order No")
_AW_INDEX = OUTPUT_COLUMNS.index("Invoice No")
CLOSED_MARKER_COLUMNS = ["Order Quantity"] + OUTPUT_COLUMNS[_P_INDEX:_AW_INDEX + 1]


def _dispatched_qty(sale):
    """Vinculum Pick Qty as Dispatched Qty, or None when not picked yet (blank / 0)."""
    qty = sale["Pick Qty"]
    try:
        if pd.isna(qty) or float(qty) <= 0:
            return None
    except (TypeError, ValueError):
        return None
    return qty


# Output column <- how to get the value from a matched Order Enquiry row.
BACKFILL_FIELDS = [
    ("Order No", lambda s, wh: s["Order No"]),                    # P
    ("Team", lambda s, wh: s["Created by"]),                      # Q
    ("Actual", lambda s, wh: s["Order Date"]),                    # S
    ("Invoice No", lambda s, wh: s["Invoice No"]),                # AW
    ("Invoice Amount", lambda s, wh: s["Total"]),                 # AY
    ("Warehouse", lambda s, wh: wh.get(normalize_value(s["Originating warehouse"]),
                                       str(s["Originating warehouse"]).strip())),  # AX
    ("Dispatched Qty", lambda s, wh: _dispatched_qty(s)),         # AS
]


_GT_APPROVAL_SUFFIX = re.compile(r"(?<=\d)[RTF]$")


def gt_base(value):
    """GT081026A1R / GT081026A1T / GT081026A1F -> GT081026A1.

    Strips exactly ONE trailing R/T/F that follows a digit (GT bases are
    "GT" + ddmmyy + "A" + serial, always ending in a digit). Comparison is by
    equality only, so GT081026A1 and GT081026A10 can never be confused.
    """
    return _GT_APPROVAL_SUFFIX.sub("", normalize_value(value))


def row_is_closed(row):
    """True when a human or an earlier job closed this line (Cancelled / Virtual
    Invoice / OOS in J or anywhere in P:AW). Such rows are never written."""
    for column in CLOSED_MARKER_COLUMNS:
        try:
            value = row[column]
        except (KeyError, IndexError):
            continue
        if str(value if value is not None else "").strip().lower() in CLOSED_MARKERS:
            return True
    return False


# Kept for callers/tests of the 2026-10-05 version.
def _is_cancelled_row(row):
    return row_is_closed(row)


def _row_marker(row):
    for column in CLOSED_MARKER_COLUMNS:
        value = str(row.get(column, "") if hasattr(row, "get") else "").strip().lower()
        if value in CLOSED_MARKERS:
            return value
    return ""


def prepare_sales_lines(sales_df):
    """Parse Order Enquiry (build_sales_lookup), add GT/SKU keys, drop lines that are
    duplicated across month tabs. Returns a DataFrame indexed 0..n-1 (line id)."""
    _, sales_df = build_sales_lookup(sales_df)
    sales_df = sales_df.reset_index(drop=True)

    sales_df["_GT"] = sales_df["ExtOrderNo"].map(normalize_value)
    sales_df["_GT_BASE"] = sales_df["ExtOrderNo"].map(gt_base)
    sales_df["_SKU"] = sales_df["SKU Code"].map(normalize_value)
    if "Bundle SKU Code" in sales_df.columns:
        sales_df["_SKU_ALT"] = sales_df["Bundle SKU Code"].map(normalize_value)
    else:
        sales_df["_SKU_ALT"] = ""
    sales_df["_ORDER"] = sales_df["Order No"].map(normalize_value)

    before = len(sales_df)
    if "_TAB" in sales_df.columns and before:
        # The same Vinculum line exported into two month tabs (e.g. "Sep-26" and
        # "Sept-26") would otherwise count as a second shipment. Same-tab repeats are
        # kept (they can be genuine separate lines). Keep the most complete copy.
        sales_df["_COMPLETE"] = (
            sales_df["Invoice No"].map(value_found).astype(int) * 2
            + sales_df["Pick Qty"].fillna(0).gt(0).astype(int)
        )
        key_cols = ["_ORDER", "_GT", "_SKU"]
        ranked = sales_df.sort_values("_COMPLETE", ascending=False, kind="stable")
        tabs_per_key = ranked.groupby(key_cols)["_TAB"].transform("nunique")
        first_tab = ranked.groupby(key_cols)["_TAB"].transform("first")
        keep = (tabs_per_key <= 1) | (ranked["_TAB"] == first_tab)
        sales_df = ranked[keep].sort_index().drop(columns=["_COMPLETE"])
        sales_df = sales_df.reset_index(drop=True)
    print(f"Order Enquiry lines: {before:,} -> {len(sales_df):,} after dropping "
          f"cross-tab duplicates ({before - len(sales_df):,} dropped)")
    return sales_df


class _Slot:
    """One Full Flow row (existing) or FMS row (to be appended) taking part in matching."""
    __slots__ = ("kind", "ref", "gt", "base", "sku", "sku_alt", "order", "closed",
                 "anchor", "assigned", "tier")

    def __init__(self, kind, ref, row):
        self.kind = kind
        self.ref = ref
        self.gt = normalize_value(row["GT Number"])
        self.base = gt_base(row["GT Number"])
        self.sku = normalize_value(row["Vinculum SKU"])
        self.sku_alt = normalize_value(row["SKU Code"]) if "SKU Code" in row.index else ""
        order = normalize_value(row["Order No"]) if "Order No" in row.index else ""
        self.order = "" if order.lower() in CLOSED_MARKERS else order
        self.closed = row_is_closed(row)
        self.anchor = None      # line id already recorded in P (refresh source)
        self.assigned = []      # line ids given to a blank row / new row
        self.tier = None


def _match_tier(slot, line):
    """1 = exact GT + Vinculum SKU, 2 = same GT base (R/T/F stripped) + Vinculum SKU,
    3 = same GT base + alternate SKU field (FMS SKU Code / Vinculum Bundle SKU Code).
    None = no match. Caller only compares lines of the same GT base."""
    if slot.base == "" or slot.base != line["_GT_BASE"]:
        return None
    if slot.sku and slot.sku == line["_SKU"]:
        return 1 if slot.gt == line["_GT"] else 2
    slot_skus = {s for s in (slot.sku, slot.sku_alt) if s}
    line_skus = {s for s in (line["_SKU"], line["_SKU_ALT"]) if s}
    if slot_skus & line_skus:
        return 3
    return None


def allocate_lines(slots, lines):
    """Give each Vinculum line to at most one slot.

    1. Anchor: a row whose Order No (P) is already filled keeps the line with that
       Order No (best tier) — that line is the source for refreshing its blank cells.
    2. Fill: rows with a blank P (existing rows first, then rows being appended), tier
       by tier, take the first unused matching line.
    3. Expand: leftover tier-1/2 lines of a GT+SKU that is being appended go to the
       last appended row of that key (one output row per line, as before).
    Leftover matching lines are returned as unplaced (reported, not inserted).
    """
    lines_by_base = {}
    for line_id, base in lines["_GT_BASE"].items():
        if base:
            lines_by_base.setdefault(base, []).append(line_id)
    records = lines[["_GT", "_GT_BASE", "_SKU", "_SKU_ALT", "_ORDER"]].to_dict("index")

    slots_by_base = {}
    for slot in slots:
        if slot.base:
            slots_by_base.setdefault(slot.base, []).append(slot)

    unplaced = []
    for base, group in slots_by_base.items():
        line_ids = lines_by_base.get(base, [])
        if not line_ids:
            continue
        used = set()

        for slot in group:
            if not slot.order:
                continue
            best = None
            for lid in line_ids:
                if lid in used or records[lid]["_ORDER"] != slot.order:
                    continue
                tier = _match_tier(slot, records[lid])
                if tier is not None and (best is None or tier < best[0]):
                    best = (tier, lid)
            if best:
                slot.anchor, slot.tier = best[1], best[0]
                used.add(best[1])

        open_slots = [s for s in group if not s.order and not s.closed and s.kind == "existing"]
        open_slots += [s for s in group if s.kind == "new"]
        for tier_limit in (1, 2, 3):
            for slot in open_slots:
                if slot.assigned:
                    continue
                for lid in line_ids:
                    if lid in used:
                        continue
                    tier = _match_tier(slot, records[lid])
                    if tier is not None and tier <= tier_limit:
                        slot.assigned.append(lid)
                        slot.tier = tier
                        used.add(lid)
                        break

        new_slots = [s for s in group if s.kind == "new"]
        for lid in line_ids:
            if lid in used:
                continue
            takers = [s for s in new_slots
                      if (_match_tier(s, records[lid]) or 9) <= 2]
            if takers:
                takers[-1].assigned.append(lid)
                used.add(lid)
                continue
            if any(_match_tier(s, records[lid]) is not None for s in group if not s.closed):
                unplaced.append(lid)
    return unplaced


def _pick_pack_cells(row, order_key, pick_pack_lookup, sheet_row, reason):
    """X (Actual.1) <- latest Pick Date, V (Warehouse Team) <- its Pick User, for the
    Vinculum Order No order_key; blank cells only."""
    pick = pick_pack_lookup.get(order_key) if order_key else None
    cells = []
    if not pick:
        return cells
    if pick["Pick Date"] is not None and not pd.isna(pick["Pick Date"]) \
            and not value_found(row["Actual.1"]):
        cells.append((sheet_row, "Actual.1",
                      pd.Timestamp(pick["Pick Date"]).to_pydatetime(), reason))
    if value_found(pick["Pick User"]) and not value_found(row["Warehouse Team"]):
        cells.append((sheet_row, "Warehouse Team", pick["Pick User"], reason))
    return cells


def plan_existing_writes(existing_df, slots, lines, pick_pack_lookup, warehouse_mapping):
    """Cell writes for existing rows: blank cells only, never a closed row.
    Returns (writes, stats); writes = [(sheet_row, output_column, value, reason)]."""
    stats = {"blank_p_rows": 0, "filled_rows": 0, "fill_cells": 0,
             "still_unmatched": 0, "refreshed_rows": 0, "refresh_cells": 0,
             "tiers": {1: 0, 2: 0, 3: 0}, "anchored_rows": 0,
             "stale_zero_dispatched": 0, "order_no_without_line": 0}
    writes = []
    for slot in slots:
        if slot.kind != "existing" or slot.closed:
            continue
        i = slot.ref
        row = existing_df.loc[i]
        if not slot.order:
            stats["blank_p_rows"] += 1
            if not slot.assigned:
                stats["still_unmatched"] += 1
                continue
            line_id, reason = slot.assigned[0], "fill"
            stats["tiers"][slot.tier] += 1
        elif slot.anchor is not None:
            line_id, reason = slot.anchor, "refresh"
            stats["anchored_rows"] += 1
        else:
            # P is filled but its line is not in Order Enquiry. Pick Pack is keyed by
            # Order No, so V/X can still come from the row's own P (blank cells only).
            stats["order_no_without_line"] += 1
            row_cells = _pick_pack_cells(row, slot.order, pick_pack_lookup,
                                         8 + int(i), "refresh")
            if row_cells:
                stats["refreshed_rows"] += 1
                stats["refresh_cells"] += len(row_cells)
                writes.extend(row_cells)
            continue

        sale = lines.loc[line_id]
        sheet_row = 8 + int(i)
        row_cells = []
        for column, getter in BACKFILL_FIELDS:
            value = getter(sale, warehouse_mapping)
            if value is not None and value_found(value) and not value_found(row[column]):
                row_cells.append((sheet_row, column, value, reason))

        row_cells += _pick_pack_cells(row, normalize_value(sale["Order No"]),
                                      pick_pack_lookup, sheet_row, reason)

        current_as = row["Dispatched Qty"]
        try:
            if value_found(current_as) and float(current_as) == 0 \
                    and _dispatched_qty(sale) is not None:
                stats["stale_zero_dispatched"] += 1
        except (TypeError, ValueError):
            pass

        if row_cells:
            if reason == "fill":
                stats["filled_rows"] += 1
                stats["fill_cells"] += len(row_cells)
            else:
                stats["refreshed_rows"] += 1
                stats["refresh_cells"] += len(row_cells)
            writes.extend(row_cells)
    return writes, stats


def plan_backfill(existing_df, sales_lines, pick_pack_lookup, warehouse_mapping):
    """Existing rows only (no appends). Kept as the single entry point the tests use.
    Returns (writes, stats, slots, unplaced)."""
    slots = [_Slot("existing", i, row) for i, row in existing_df.iterrows()]
    unplaced = allocate_lines(slots, sales_lines)
    writes, stats = plan_existing_writes(existing_df, slots, sales_lines,
                                         pick_pack_lookup, warehouse_mapping)
    stats["unplaced_sales_lines"] = len(unplaced)
    return writes, stats, slots, unplaced


def find_new_fms_rows(fms_df, existing_df):
    """FMS rows whose GT Number + Vinculum SKU is not in Full Flow yet.

    Returns (new_gt_mask, new_line_mask): rows of GTs Full Flow has never seen, and
    rows of GTs it has but whose SKU line is missing (gap 3). Exact normalised keys:
    a line already in Full Flow in any state (incl. Cancelled) is never re-added.
    """
    ff_gts = set(existing_df["GT Number"].map(normalize_value)) - {""}
    ff_keys = set(
        create_match_key(gt, sku)
        for gt, sku in zip(existing_df["GT Number"], existing_df["Vinculum SKU"])
    )
    gts = fms_df["GT Number"].map(normalize_value)
    keys = pd.Series(
        [create_match_key(gt, sku) for gt, sku in zip(fms_df["GT Number"], fms_df["Vinculum SKU"])],
        index=fms_df.index,
    )
    has_gt = gts != ""
    new_gt_mask = has_gt & ~gts.isin(ff_gts)
    new_line_mask = has_gt & gts.isin(ff_gts) & ~keys.isin(ff_keys)
    return new_gt_mask, new_line_mask


def report_suspect_cancellations(existing_df, sales_lines, path=None):
    """READ-ONLY. Full Flow rows that read "Cancelled" but DO have a Vinculum line under
    the suffix-tolerant rule — likely fabricated by the historical-cancellation logic
    (no match -> "Cancelled"). Nothing is reverted; the CSV is for a human decision."""
    probe = []
    for i, row in existing_df.iterrows():
        if _row_marker(row) != "cancelled":
            continue
        slot = _Slot("existing", i, row)
        slot.order = ""
        probe.append((i, row, slot))

    by_base = {}
    for line_id, base in sales_lines["_GT_BASE"].items():
        if base:
            by_base.setdefault(base, []).append(line_id)
    records = sales_lines[["_GT", "_GT_BASE", "_SKU", "_SKU_ALT", "_ORDER"]].to_dict("index")

    rows = []
    for i, row, slot in probe:
        hits = [lid for lid in by_base.get(slot.base, [])
                if _match_tier(slot, records[lid]) is not None]
        if not hits:
            continue
        p_value = str(row["Order No"]).strip().lower()
        pattern = ("dashboard/manual cancel (P:AW)" if p_value == "cancelled"
                   else "historical auto-cancel (T:AW, P blank)")
        matched = sales_lines.loc[hits]
        invoices = [str(v) for v in matched["Invoice No"] if value_found(v)]
        rows.append({
            "sheet_row": 8 + int(i),
            "gt_number": row["GT Number"],
            "vinculum_sku": row["Vinculum SKU"],
            "form_received": row["Form Received Time"],
            "pattern": pattern,
            "vinculum_order_nos": ", ".join(dict.fromkeys(str(v) for v in matched["Order No"])),
            "vinculum_ext_order_nos": ", ".join(dict.fromkeys(str(v) for v in matched["ExtOrderNo"])),
            "vinculum_invoice_nos": ", ".join(dict.fromkeys(invoices)),
            "invoiced": "yes" if invoices else "no",
        })
    report = pd.DataFrame(rows, columns=[
        "sheet_row", "gt_number", "vinculum_sku", "form_received", "pattern",
        "vinculum_order_nos", "vinculum_ext_order_nos", "vinculum_invoice_nos", "invoiced"])

    print("\n" + "=" * 80)
    print("REPORT: 'Cancelled' rows that DO have a Vinculum line (read-only, nothing reverted)")
    print("=" * 80)
    print(f"Rows reading Cancelled:                 {len(probe):,}")
    print(f"... of which have a Vinculum match:     {len(report):,}")
    if len(report):
        for pattern, count in report["pattern"].value_counts().items():
            print(f"  {pattern:<42} {count:,}")
        print(f"  invoiced in Vinculum:                 {(report['invoiced'] == 'yes').sum():,}")
        for _, r in report.head(25).iterrows():
            print(f"  row {r['sheet_row']}: {r['gt_number']} / {r['vinculum_sku']} -> "
                  f"{r['vinculum_order_nos']} (ext {r['vinculum_ext_order_nos']}; "
                  f"inv {r['vinculum_invoice_nos'] or '-'}) [{r['pattern']}]")
    if path:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        report.to_csv(path, index=False)
        print(f"Full list written to {path} (workflow artifact 'fullflow-reports').")
    return report


def _print_unmatched_diagnostics(existing_df, slots, sales_lines, limit=15, path=None):
    """READ-ONLY. For blank-P rows that still have no line: is the GT base in Vinculum
    at all, which SKUs does Vinculum hold for it, and which ExtOrderNos merely CONTAIN
    the GT base (formatting variants). Writes a CSV for offline analysis."""
    bases = set(sales_lines["_GT_BASE"]) - {""}
    lines_by_base = {}
    for lid, base in sales_lines["_GT_BASE"].items():
        if base:
            lines_by_base.setdefault(base, []).append(lid)
    used = set()
    for slot in slots:
        if slot.anchor is not None:
            used.add(slot.anchor)
        used.update(slot.assigned)
    unique_ext = sorted(set(sales_lines["_GT"]) - {""})
    print("Order Enquiry columns: " + " | ".join(str(c) for c in sales_lines.columns
                                                  if not str(c).startswith("_")))
    gt_absent = gt_present = 0
    samples, rows = [], []
    for slot in slots:
        if slot.kind != "existing" or slot.closed or slot.order or slot.assigned:
            continue
        row = existing_df.loc[slot.ref]
        if slot.base in bases:
            gt_present += 1
            reason = "sku_differs"
            lids = lines_by_base[slot.base]
            vin = [f"{sales_lines.at[l, '_SKU']}={sales_lines.at[l, '_ORDER']}"
                   f"{'*' if l in used else ''}" for l in lids]
            near = []
            if len(samples) < limit:
                samples.append(f"  SKU differs: row {8 + int(slot.ref)} {slot.gt} / {slot.sku}"
                               f" (alt {slot.sku_alt or '-'}) ; Vinculum has "
                               + ", ".join(vin[:8]))
        else:
            gt_absent += 1
            reason = "gt_not_in_vinculum"
            vin = []
            near = [e for e in unique_ext
                    if slot.base and slot.base in e and e != slot.base
                    and not re.match(r"\d", e[e.index(slot.base) + len(slot.base):][:1] or "x")][:5]
        rows.append({
            "sheet_row": 8 + int(slot.ref), "gt_number": slot.gt, "gt_base": slot.base,
            "vinculum_sku": slot.sku, "sku_code": slot.sku_alt,
            "form_received": row["Form Received Time"], "party": row.get("Party Name", ""),
            "reason": reason, "vinculum_lines_for_gt (*=used)": " ".join(vin),
            "ext_order_nos_containing_gt": " ".join(near),
        })
    print(f"Still unmatched, GT base NOT in Vinculum (not punched yet): {gt_absent:,}")
    print(f"Still unmatched, GT base in Vinculum but no SKU match:      {gt_present:,}")
    near_count = sum(1 for r in rows if r["ext_order_nos_containing_gt"])
    print(f"  ... GT not in Vinculum but an ExtOrderNo CONTAINS it:     {near_count:,}")
    for s in samples:
        print(s)
    if path:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        pd.DataFrame(rows).to_csv(path, index=False)
        print(f"Unmatched rows written to {path} (workflow artifact 'fullflow-reports').")
        # Vinculum GT-looking orders since Aug whose GT base is on NO Full Flow row:
        # if ExtOrderNo formatting drifted, the "not in Vinculum" GTs show up here.
        ff_bases = {s.base for s in slots}
        recent = sales_lines["Order Date"].ge(pd.Timestamp("2026-08-01"))
        unclaimed = sales_lines[recent & ~sales_lines["_GT_BASE"].isin(ff_bases)]
        unclaimed = unclaimed[["ExtOrderNo", "Order No", "SKU Code", "Order Date", "Created by"]]
        unclaimed_path = os.path.join(os.path.dirname(path), "vinculum_lines_not_in_full_flow.csv")
        unclaimed.to_csv(unclaimed_path, index=False)
        print(f"Vinculum lines since Aug with a GT base on no Full Flow row: {len(unclaimed):,} "
              f"-> {unclaimed_path}")


FILL_RATE_COLUMNS = ["Order No", "Team", "Actual", "Warehouse Team", "Actual.1",
                     "Order Delivery Status", "Dispatched Qty", "POD Copy",
                     "Delivered Date", "Invoice No", "Warehouse", "Invoice Amount"]


def _print_fill_rates(existing_df, writes=None):
    """READ-ONLY. Filled cells per column and form month on open (not closed) rows.
    With writes, the table is the PROJECTION after this run's fills/refreshes."""
    closed = existing_df.apply(row_is_closed, axis=1)
    if writes is not None:
        existing_df = existing_df.copy()
        for sheet_row, column, value, _ in writes:
            existing_df.at[sheet_row - 8, column] = value
    open_rows = existing_df[~closed]
    months = open_rows["Form Received Time"].map(_month_of)
    title = ("PROJECTED fill rate on open rows after this run's fills/refreshes"
             if writes is not None else "Fill rate on open rows (filled / rows) by form month")
    print(f"\n{title}; closed rows: {int(closed.sum()):,}")
    header = f"{'month':<8}{'rows':>6} " + " ".join(f"{backfill_a1(8, c).split('!')[1][:-1]:>5}"
                                                  for c in FILL_RATE_COLUMNS)
    print(header)
    for month in sorted(months.unique()):
        part = open_rows[months == month]
        counts = [int(part[c].map(value_found).sum()) for c in FILL_RATE_COLUMNS]
        print(f"{month:<8}{len(part):>6} " + " ".join(f"{n:>5}" for n in counts))


def backfill_a1(sheet_row, column):
    n, letters = OUTPUT_COLUMNS.index(column) + 1, ""
    while n:
        n, r = divmod(n - 1, 26)
        letters = chr(65 + r) + letters
    return f"'{FULL_FLOW_SHEET_NAME}'!{letters}{sheet_row}"


def write_cells(writes):
    data = [{"range": backfill_a1(r, c), "values": [[clean_for_google(v)]]}
            for r, c, v, _ in writes]
    for start in range(0, len(data), 500):
        sheets_service.spreadsheets().values().batchUpdate(
            spreadsheetId=OUTPUT_SPREADSHEET_ID,
            body={"valueInputOption": "USER_ENTERED", "data": data[start:start + 500]},
        ).execute()
        print(f"  Written cells {start + 1:,} - {min(start + 500, len(data)):,}")
    print(f"Fill/refresh complete: {len(data):,} cells written.")


def _month_of(value):
    ts = pd.to_datetime(value, errors="coerce", dayfirst=True)
    return "unknown" if pd.isna(ts) else ts.strftime("%Y-%m")


# ============================================================
# MAIN
# ============================================================

def main():
    print("=" * 80)
    print("GT FULL FLOW FMS AUTOMATION - GOOGLE SHEETS")
    print("=" * 80)
    print(f"Source Spreadsheet ID:      {SOURCE_SPREADSHEET_ID}")
    print(f"Source FMS tab:             {FMS_SHEET_NAME}")
    print(f"Output Spreadsheet ID:      {OUTPUT_SPREADSHEET_ID}")
    print(f"Output tab:                 {FULL_FLOW_SHEET_NAME}")
    print(f"Order Enquiry Spreadsheet:  {ORDER_ENQUIRY_SPREADSHEET_ID}")
    print(f"Order Enquiry tabs config:  {ORDER_ENQUIRY_SHEET_NAMES if ORDER_ENQUIRY_SHEET_NAMES else 'ALL TABS'}")
    print(f"Pick Pack folder ID:        {PICK_PACK_FOLDER_ID}")
    print(f"Mode (BACKFILL_MODE):       {BACKFILL_MODE}")
    if BACKFILL_MODE not in ("write", "dry-run", "off"):
        raise SystemExit(f"Unknown BACKFILL_MODE {BACKFILL_MODE!r} (write / dry-run / off)")
    dry_run = BACKFILL_MODE == "dry-run"

    # --------------------------------------------------------
    # READ SOURCE DATA
    # --------------------------------------------------------
    fms_df = read_fms_sheet()
    sales_df = read_sales_file()
    pick_pack_df = read_pick_pack_files()
    warehouse_mapping = read_warehouse_mapping()

    # Disabled since 2026-10-09 (ENABLE_HISTORICAL_CANCELLATION = "NO"); kept so the
    # switch still works. It only ever affected rows being appended.
    fms_df = apply_historical_cancellation(fms_df, sales_df)

    output_sheet_id = get_output_sheet()
    existing_full_flow_df = read_existing_full_flow()

    sales_lines = prepare_sales_lines(sales_df.copy())
    try:
        pick_pack_lookup = build_pick_pack_lookup(pick_pack_df.copy())
    except SystemExit:
        pick_pack_lookup = {}

    # --------------------------------------------------------
    # WHICH FMS ROWS ARE MISSING FROM FULL FLOW
    # --------------------------------------------------------
    new_gt_mask, new_line_mask = find_new_fms_rows(fms_df, existing_full_flow_df)
    print("\n" + "=" * 80)
    print("FMS ROWS NOT YET IN FULL FLOW")
    print("=" * 80)
    print(f"Rows of GTs not in Full Flow (new GTs):              {int(new_gt_mask.sum()):,}")
    print(f"New SKU lines for GTs already in Full Flow:          {int(new_line_mask.sum()):,}")
    if new_line_mask.any():
        months = fms_df.loc[new_line_mask, "Form Received Time"].map(_month_of).value_counts().sort_index()
        print("  by form month: " + ", ".join(f"{m} {c}" for m, c in months.items()))
        for _, r in fms_df.loc[new_line_mask].head(15).iterrows():
            print(f"  new line: {r['GT Number']} / {r['Vinculum SKU']} qty {r['Order Quantity']}"
                  f" ({r['Form Received Time']})")
    if int(new_line_mask.sum()) > NEW_LINES_FOR_EXISTING_GTS_MAX:
        print(f"WARNING: more than {NEW_LINES_FOR_EXISTING_GTS_MAX} new lines for existing GTs — "
              f"NOT appending them (looks structural, check the samples above). New GTs still go in.")
        new_line_mask[:] = False
    fms_to_append = fms_df.loc[new_gt_mask | new_line_mask].copy()

    # --------------------------------------------------------
    # MATCH (existing rows + rows to append) AGAINST VINCULUM
    # --------------------------------------------------------
    slots = [_Slot("existing", i, row) for i, row in existing_full_flow_df.iterrows()]
    new_slots = [_Slot("new", pos, row)
                 for pos, (_, row) in enumerate(fms_to_append.iterrows())]
    unplaced = allocate_lines(slots + new_slots, sales_lines)

    writes, stats = plan_existing_writes(existing_full_flow_df, slots, sales_lines,
                                         pick_pack_lookup, warehouse_mapping)

    print("\n" + "=" * 80)
    print(f"FILL + REFRESH EXISTING ROWS  —  mode: {BACKFILL_MODE}")
    print("=" * 80)
    closed_rows = sum(1 for s in slots if s.closed)
    print(f"Existing rows:                                {len(slots):,}")
    print(f"Closed rows skipped (Cancelled/Virtual/OOS):  {closed_rows:,}")
    print(f"Rows with blank Order No considered:          {stats['blank_p_rows']:,}")
    print(f"  matched now (would be / are filled):        {stats['filled_rows']:,}  "
          f"(tier1 exact {stats['tiers'][1]:,} | tier2 R/T/F suffix {stats['tiers'][2]:,} | "
          f"tier3 alt SKU {stats['tiers'][3]:,})")
    print(f"  cells:                                      {stats['fill_cells']:,}")
    print(f"Rows still without a Vinculum match:          {stats['still_unmatched']:,}")
    print(f"Rows with Order No anchored to a Vinculum line: {stats['anchored_rows']:,}")
    print(f"  of which get blank cells refreshed:         {stats['refreshed_rows']:,}")
    print(f"  refresh cells:                              {stats['refresh_cells']:,}")
    print(f"Rows with an Order No not found in Vinculum (no refresh): {stats['order_no_without_line']:,}")
    print(f"Rows with Dispatched Qty 0 but Vinculum picked >0 (not overwritten): "
          f"{stats['stale_zero_dispatched']:,}")
    print(f"Extra Vinculum lines not placed (split shipments): {len(unplaced):,}")
    for lid in unplaced[:15]:
        line = sales_lines.loc[lid]
        print(f"  unplaced: {line['ExtOrderNo']} / {line['SKU Code']} -> {line['Order No']}")
    per_col = {}
    for _, column, _, reason in writes:
        per_col[(reason, column)] = per_col.get((reason, column), 0) + 1
    for (reason, column), count in sorted(per_col.items()):
        print(f"  {reason:<8} {column:<16} {count:,}")
    for sheet_row, column, value, reason in writes[:10]:
        print(f"  sample {reason}: row {sheet_row} {column} = {clean_for_google(value)!r}")
    tier2 = [s for s in slots if s.kind == "existing" and s.assigned and s.tier == 2][:10]
    for s in tier2:
        line = sales_lines.loc[s.assigned[0]]
        print(f"  suffix match: row {8 + int(s.ref)} {s.gt} / {s.sku} <- Vinculum "
              f"{line['ExtOrderNo']} / {line['SKU Code']} = {line['Order No']}")
    tier3 = [s for s in slots if s.kind == "existing" and s.assigned and s.tier == 3][:10]
    for s in tier3:
        line = sales_lines.loc[s.assigned[0]]
        print(f"  alt-SKU match: row {8 + int(s.ref)} {s.gt} / {s.sku} (SKU Code {s.sku_alt}) <- "
              f"Vinculum {line['ExtOrderNo']} / {line['SKU Code']} "
              f"(bundle {line.get('Bundle SKU Code', '')}) = {line['Order No']}")
    _print_unmatched_diagnostics(existing_full_flow_df, slots, sales_lines,
                                 path=os.path.join("reports", "unmatched_rows.csv"))

    _print_fill_rates(existing_full_flow_df)
    _print_fill_rates(existing_full_flow_df, writes=writes)

    report_suspect_cancellations(
        existing_full_flow_df, sales_lines,
        path=os.path.join("reports", "suspect_cancellations.csv"))

    # --------------------------------------------------------
    # BUILD ROWS TO APPEND
    # --------------------------------------------------------
    assigned = [[sales_lines.loc[lid] for lid in s.assigned] for s in new_slots]
    new_full_flow_df = build_full_flow(fms_to_append, assigned, pick_pack_lookup,
                                       warehouse_mapping) if len(fms_to_append) else \
        pd.DataFrame(columns=OUTPUT_COLUMNS)

    if dry_run:
        print("\nDRY RUN — nothing written (no fills, no refresh, no appends).")
        print(f"Would append {len(new_full_flow_df):,} rows after row "
              f"{7 + len(existing_full_flow_df):,}.")
        for _, r in new_full_flow_df.head(15).iterrows():
            print(f"  would append: {r['GT Number']} / {r['Vinculum SKU']} -> "
                  f"Order No {clean_for_google(r['Order No']) or '-'}")
        return

    if BACKFILL_MODE == "write" and writes:
        write_cells(writes)
    elif BACKFILL_MODE == "off":
        print("Fill/refresh disabled (off).")

    if new_full_flow_df.empty:
        print("\nNothing to append.")
        return

    existing_row_count = len(existing_full_flow_df)
    write_full_flow(output_sheet_id, new_full_flow_df, start_row=8 + existing_row_count)

    print("\n" + "=" * 80)
    print("COMPLETED SUCCESSFULLY")
    print("=" * 80)
    print(f"New Full Flow rows appended: {len(new_full_flow_df):,} "
          f"(after {existing_row_count:,} existing rows)")
    print("FMS source tab was not modified.")
    print(f"Cancellation logic: {str(ENABLE_HISTORICAL_CANCELLATION).strip().upper()}")


if __name__ == "__main__":
    main()
