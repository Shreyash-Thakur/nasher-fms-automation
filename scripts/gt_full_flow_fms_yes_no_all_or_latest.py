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
ENABLE_HISTORICAL_CANCELLATION = "YES"

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
    request = drive_service.files().get_media(fileId=file_id)
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


def get_processed_gt_numbers(existing_df):
    """Return GTs already present in the destination Full Flow FMS."""
    if existing_df.empty or "GT Number" not in existing_df.columns:
        return set()

    processed = set()
    for value in existing_df["GT Number"]:
        key = normalize_value(value)
        if key:
            processed.add(key)

    return processed


def filter_unprocessed_fms(fms_df, processed_gt_numbers):
    """Keep only FMS rows belonging to GTs not already processed."""
    if not processed_gt_numbers:
        print("  No previously processed GTs found. All FMS rows are eligible.")
        return fms_df.copy()

    gt_keys = fms_df["GT Number"].map(normalize_value)
    already_processed_mask = gt_keys.isin(processed_gt_numbers)

    skipped_rows = int(already_processed_mask.sum())
    new_df = fms_df.loc[~already_processed_mask].copy()

    print(f"  FMS rows already processed: {skipped_rows:,}")
    print(f"  FMS rows to process now:    {len(new_df):,}")
    print(f"  Already processed GTs:      {len(processed_gt_numbers):,}")

    return new_df


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

    return pd.to_datetime(value, dayfirst=True, errors="coerce")


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


def read_pick_pack_files():
    print("\nLooking for Pick Pack CSV files in Google Drive...")

    files = list_drive_files(PICK_PACK_FOLDER_ID)
    csv_files = [f for f in files if f["name"].lower().endswith(".csv")]

    if not csv_files:
        raise FileNotFoundError(
            "No Pick Pack CSV files found in the configured Google Drive folder."
        )

    print(f"Pick Pack CSV files found: {len(csv_files):,}")
    for f in csv_files:
        print(f"  {f['name']}")

    dataframes = []

    for file_info in csv_files:
        buffer = download_drive_file(file_info["id"])
        raw_bytes = buffer.getvalue()
        df = None

        for encoding in ["utf-8-sig", "cp1252", "latin1"]:
            try:
                df = pd.read_csv(
                    io.BytesIO(raw_bytes),
                    encoding=encoding,
                    dtype=str,
                    keep_default_na=False,
                    low_memory=False,
                )
                break
            except Exception:
                pass

        if df is None:
            print(f"ERROR reading {file_info['name']}; skipped.")
            continue

        df.columns = (
            df.columns.astype(str)
            .str.replace("\ufeff", "", regex=False)
            .str.replace("\xa0", " ", regex=False)
            .str.strip()
        )
        df = df.dropna(how="all")
        dataframes.append(df)
        print(f"Read: {file_info['name']} -> {len(df):,} rows")

    if not dataframes:
        raise ValueError("No Pick Pack CSV could be read.")

    pick_pack_df = pd.concat(dataframes, ignore_index=True)
    pick_pack_df = pick_pack_df.dropna(how="all").reset_index(drop=True)

    missing = [c for c in PICK_PACK_REQUIRED_COLUMNS if c not in pick_pack_df.columns]
    if missing:
        print("Actual Pick Pack columns found:")
        for col in pick_pack_df.columns:
            print(f"  [{col}]")
        raise ValueError("Missing Pick Pack columns:\n" + "\n".join(missing))

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

def build_full_flow(fms_df, sales_df, pick_pack_df, warehouse_mapping):
    print("\nCreating Sales match keys...")

    fms_df["_MATCH_KEY"] = fms_df.apply(
        lambda row: create_match_key(row["GT Number"], row["Vinculum SKU"]),
        axis=1,
    )

    sales_df["_MATCH_KEY"] = sales_df.apply(
        lambda row: create_match_key(row["ExtOrderNo"], row["SKU Code"]),
        axis=1,
    )

    fms_df["Form Received Time"] = pd.to_datetime(
        fms_df["Form Received Time"], errors="coerce", dayfirst=True
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

    # Sales lookup
    sales_groups = {
        key: group
        for key, group in sales_df.groupby("_MATCH_KEY", sort=False)
    }

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

    full_flow_rows = []
    matched_fms_count = 0
    unmatched_fms_count = 0
    expanded_fms_count = 0
    pick_pack_match_count = 0
    pick_pack_actual_count = 0
    pick_pack_user_count = 0
    warehouse_match_count = 0
    warehouse_missing_count = 0

    print("\nBuilding Full Flow FMS...")

    for _, fms_row in fms_df.iterrows():
        match_key = fms_row["_MATCH_KEY"]

        base_row = {
            column: fms_row[column] if column in fms_row.index else ""
            for column in OUTPUT_COLUMNS
        }

        # --------------------------------------------------------
        # SALES MATCH FOUND
        # --------------------------------------------------------
        if match_key in sales_groups:
            matched_sales_rows = sales_groups[match_key]
            matched_fms_count += 1

            if len(matched_sales_rows) > 1:
                expanded_fms_count += 1

            for _, sales_row in matched_sales_rows.iterrows():
                output_row = base_row.copy()

                # Order Enquiry -> Full Flow mappings
                output_row["Order No"] = sales_row["Order No"]        # P
                output_row["Team"] = sales_row["Created by"]          # Q
                output_row["Actual"] = sales_row["Order Date"]        # S
                output_row["Invoice No"] = sales_row["Invoice No"]    # AW
                output_row["Invoice Amount"] = sales_row["Total"]     # AY
                # Order Quantity (J) is deliberately left as the FMS source value.

                # Order Enquiry -> warehouse and dispatched quantity mappings.
                warehouse_code = sales_row["Originating warehouse"]
                warehouse_key = normalize_value(warehouse_code)
                if warehouse_key and warehouse_key in warehouse_mapping:
                    output_row["Warehouse"] = warehouse_mapping[warehouse_key]
                    warehouse_match_count += 1
                elif warehouse_key:
                    # Keep the source code as a diagnostic fallback if no mapping
                    # exists. The run summary reports these unmapped codes.
                    output_row["Warehouse"] = str(warehouse_code).strip()
                    warehouse_missing_count += 1

                dispatched_qty = sales_row["Pick Qty"]
                if pd.notna(dispatched_qty):
                    output_row["Dispatched Qty"] = dispatched_qty

                # Pick Pack bridge
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

        # --------------------------------------------------------
        # NO SALES MATCH
        # --------------------------------------------------------
        else:
            unmatched_fms_count += 1
            output_row = base_row.copy()

            # No cancellation logic here.
            # The unmatched FMS row is simply retained as-is,
            # then Pick Pack is bridged if an Order No is available.
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

    # Open Qty (AT) = Order Quantity (J) - Dispatched Qty (AS).
    # The per-row Google Sheets formula is written in write_full_flow().
    full_flow_df["Open Qty"] = full_flow_df["Open Qty"]

    invoice_amount_mapped = int(sales_df["Total"].notna().sum())

    print("\n" + "=" * 80)
    print("PROCESSING SUMMARY")
    print("=" * 80)
    print(f"Original FMS rows:             {len(fms_df):,}")
    print(f"FMS rows with Sales match:     {matched_fms_count:,}")
    print(f"FMS rows without Sales match:  {unmatched_fms_count:,}")
    print(f"FMS rows expanded:             {expanded_fms_count:,}")
    print(f"Final Full Flow rows:          {len(full_flow_df):,}")
    print(f"Pick Pack matches:             {pick_pack_match_count:,}")
    print(f"Pick Dates -> Actual.1 (X):    {pick_pack_actual_count:,}")
    print(f"Pick Users -> Warehouse Team (V): {pick_pack_user_count:,}")
    print(f"Warehouse codes mapped -> AX:  {warehouse_match_count:,}")
    print(f"Invoice Amounts mapped -> AY:  {invoice_amount_mapped:,}")
    print(f"Warehouse codes unmapped:      {warehouse_missing_count:,}")

    return full_flow_df


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

    # --------------------------------------------------------
    # READ SOURCE DATA
    # --------------------------------------------------------
    fms_df = read_fms_sheet()
    sales_df = read_sales_file()
    pick_pack_df = read_pick_pack_files()
    warehouse_mapping = read_warehouse_mapping()

    # --------------------------------------------------------
    # OPTIONAL HISTORICAL CANCELLATION
    # --------------------------------------------------------
    fms_df = apply_historical_cancellation(
        fms_df,
        sales_df,
    )

    # --------------------------------------------------------
    # INCREMENTAL PROCESSING
    # --------------------------------------------------------
    # Do NOT clear the destination. Existing processed GTs must keep
    # their current points/statuses/actuals. Only new GTs are processed.
    output_sheet_id = get_output_sheet()
    existing_full_flow_df = read_existing_full_flow()
    processed_gt_numbers = get_processed_gt_numbers(existing_full_flow_df)

    fms_to_process = filter_unprocessed_fms(
        fms_df,
        processed_gt_numbers,
    )

    if fms_to_process.empty:
        print("\nNo new GTs to process.")
        print("Existing Full Flow FMS data has been left completely unchanged.")
        return

    print("\nBuilding Full Flow only for new/unprocessed GTs...")
    new_full_flow_df = build_full_flow(
        fms_to_process,
        sales_df,
        pick_pack_df,
        warehouse_mapping,
    )

    existing_row_count = len(existing_full_flow_df)
    write_start_row = 8 + existing_row_count

    write_full_flow(
        output_sheet_id,
        new_full_flow_df,
        start_row=write_start_row,
    )

    print(
        f"\nIncremental update complete: "
        f"{len(new_full_flow_df):,} new rows appended."
    )
    print(
        f"Existing rows preserved: {existing_row_count:,}"
    )

    print("\n" + "=" * 80)
    print("COMPLETED SUCCESSFULLY")
    print("=" * 80)
    print(
        "Source Spreadsheet: "
        f"https://docs.google.com/spreadsheets/d/"
        f"{SOURCE_SPREADSHEET_ID}/edit"
    )
    print(
        "Destination Spreadsheet: "
        f"https://docs.google.com/spreadsheets/d/"
        f"{OUTPUT_SPREADSHEET_ID}/edit"
    )
    print(f"Output tab: {FULL_FLOW_SHEET_NAME}")
    # In incremental mode the newly generated dataframe is named
    # new_full_flow_df. Do not reference the old full_flow_df variable.
    print(f"New Full Flow rows appended: {len(new_full_flow_df):,}")
    print("FMS source tab was not modified.")
    print(
        "Cancellation logic: "
        f"{str(ENABLE_HISTORICAL_CANCELLATION).strip().upper()} "
        f"(cutoff={CANCELLATION_CUTOFF_DATE})"
    )


if __name__ == "__main__":
    main()
