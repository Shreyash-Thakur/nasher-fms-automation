import os
import json
import base64
from datetime import datetime, timedelta, date
import time
import random

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from google_auth_httplib2 import AuthorizedHttp
import httplib2


# ============================================================
# CONFIGURATION
# ============================================================

# MAIN OUTPUT GOOGLE SPREADSHEET
OUTPUT_SPREADSHEET_ID = (
    "1jMEEFkAVfg7NHtxj3QCYuNY_V-AhU00ljQeRYTBQDio"
)

MAIN_SHEET = "Full Flow FMS"
HOLIDAY_SHEET = "Holiday"

SERVICE_ACCOUNT_FILE = (
    r"C:\Users\diwak\.gcloud\automation-bot.json"
)

DATA_START_ROW = 8

# Google Sheets columns used by this process.
LAST_DATA_COLUMN = "AR"

# Use large batches so we stay well below Google Sheets' write-request
# quota. One API call can contain many cell updates.
BATCH_SIZE = 5000

# Retry Google Sheets 429 rate-limit responses automatically.
MAX_RETRIES = 8
INITIAL_RETRY_SECONDS = 2


# ============================================================
# COLUMN MAPPING
# ============================================================

# ------------------------------------------------------------
# SOURCE
# ------------------------------------------------------------

COL_FORM_RECEIVED = 1       # A


# ------------------------------------------------------------
# STAGE 1
# ------------------------------------------------------------

COL_R = 18                  # R = Planned
COL_S = 19                  # S = Actual
COL_T = 20                  # T = Status
COL_U = 21                  # U = Time Delay


# ------------------------------------------------------------
# STAGE 2
# ------------------------------------------------------------

COL_W = 23                  # W = Planned
COL_X = 24                  # X = Actual
COL_Y = 25                  # Y = Status
COL_Z = 26                  # Z = Time Delay


# ------------------------------------------------------------
# STAGE 3
# ------------------------------------------------------------

COL_AA = 27                 # AA = Planned
COL_AB = 28                 # AB = Actual
COL_AC = 29                 # AC = Status
COL_AD = 30                 # AD = Time Delay


# ------------------------------------------------------------
# STAGE 4
# ------------------------------------------------------------

COL_AF = 32                 # AF = Planned
COL_AG = 33                 # AG = Actual
COL_AH = 34                 # AH = Status
COL_AI = 35                 # AI = Time Delay


# ------------------------------------------------------------
# STAGE 5
# ------------------------------------------------------------

COL_AK = 37                 # AK = Planned
COL_AL = 38                 # AL = Actual
COL_AM = 39                 # AM = Status
COL_AN = 40                 # AN = Time Delay


# ------------------------------------------------------------
# STAGE 6
# ------------------------------------------------------------

COL_AO = 41                 # AO = Planned
COL_AP = 42                 # AP = Actual
COL_AQ = 43                 # AQ = Status
COL_AR = 44                 # AR = Time Delay


# ============================================================
# DATE FORMAT
# ============================================================

DATE_TIME_FORMAT = "dd-mmm-yyyy hh:mm:ss"

# ============================================================
# STATUS ALERT THRESHOLDS
# ============================================================
# A status cell is highlighted RED when:
#   Early -> more than 24 hours early
#   Delay -> 3 days or more late
#
# Change these values here if the business rule changes.
EARLY_ALERT_HOURS = 24
DELAY_ALERT_DAYS = 3


# ============================================================
# ============================================================
# === RECONSTRUCTED HELPERS - VERIFY AGAINST YOUR ORIGINALS ===
# ============================================================
# The file you shared did not include these four functions, but
# every stage in main() and the holiday loaders call them. The
# implementations below are BEST-GUESS reconstructions based on the
# surrounding code. Replace this whole block with your real versions
# before running against the production sheet - especially
# process_stage(), which decides what gets written and formatted.
# ============================================================

# Datetime parse formats. Python's %b is locale month abbrev (Aug),
# matching the DATE_TIME_FORMAT "dd-mmm-yyyy hh:mm:ss".
_DATETIME_FORMATS = [
    "%d-%b-%Y %H:%M:%S",
    "%d-%b-%Y %H:%M",
    "%d-%b-%Y",
    "%d-%m-%Y %H:%M:%S",
    "%d-%m-%Y %H:%M",
    "%d-%m-%Y",
    "%d/%m/%Y %H:%M:%S",
    "%d/%m/%Y %H:%M",
    "%d/%m/%Y",
]

_DATE_FORMATS = [
    "%d-%m-%Y",
    "%d/%m/%Y",
    "%d-%b-%Y",
    "%d-%B-%Y",
    "%Y-%m-%d",
]


def to_datetime(value):
    """Parse a formatted cell string into a datetime, or None."""
    if value is None:
        return None

    text = str(value).strip()
    if not text:
        return None

    text = (
        text.replace("\u2013", "-")
            .replace("\u2014", "-")
            .replace("\u2212", "-")
    )

    for fmt in _DATETIME_FORMATS:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue

    return None


def normalize_date(value):
    """Parse a holiday cell into a datetime.date, or None.

    Returns a date (not datetime) so holidays can be compared against
    the .date() of any timestamp in add_time_skip_holidays.
    """
    if value is None:
        return None

    text = str(value).strip()
    if not text:
        return None

    text = (
        text.replace("\u2013", "-")
            .replace("\u2014", "-")
            .replace("\u2212", "-")
    )

    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue

    # Fall back to the datetime parser, then drop the time.
    parsed = to_datetime(text)
    if parsed is not None:
        return parsed.date()

    return None


def add_time_skip_holidays(start, delta, holidays):
    """Add delta to start, pushing the result past any holiday days.

    While the resulting calendar day is a holiday, advance one whole
    day at a time until it lands on a non-holiday day.
    """
    if start is None:
        return None

    result = start + delta

    # Skip forward over holiday days.
    while result.date() in holidays:
        result = result + timedelta(days=1)

    return result


def format_time_delay(actual, planned):
    """Full-precision gap between actual and planned, WITHOUT a sign.

    Formatted the SAME way for every stage and in both directions:
    days + hours + minutes + seconds (the days part is shown only when
    it is non-zero). There is no longer any date-only special case, so
    a stage whose Actual is stored at midnight is measured with exactly
    the same precision as one that has a real time. This is what keeps
    stage 1 and stage 2 in the same format instead of one showing "88d"
    and the other "87d 11h 24m".

    Direction (early vs late) is carried by the Status column, so no
    leading '-' is added here.
    """
    total_seconds = int(abs((actual - planned).total_seconds()))

    days, rem = divmod(total_seconds, 86400)
    hours, rem = divmod(rem, 3600)
    minutes, seconds = divmod(rem, 60)

    if days:
        return f"{days}d {hours:02d}h {minutes:02d}m {seconds:02d}s"

    return f"{hours:02d}h {minutes:02d}m {seconds:02d}s"


def process_stage(
    row_values,
    row,
    col_planned,
    col_actual,
    col_status,
    col_time_delay,
    planned_dt,
    updates,
    format_requests,
    sheet_id,
    actual_override=None,
):
    """Write Planned/Status/Time Delay for one stage and queue alert
    formatting.

    Behaviour (reconstructed):
      - Planned cell is written only when currently blank.
      - When the Actual cell has a value, Status + Time Delay are
        computed by comparing Actual to Planned:
            Early   -> actual before planned
            Delay   -> actual after planned
            On Time -> equal
      - Red alert formatting is queued when Early by more than
        EARLY_ALERT_HOURS or Delay by DELAY_ALERT_DAYS or more;
        otherwise any existing red formatting is cleared.

    Returns the status string ("Early" / "Delay" / "On Time" / "").
    """
    if planned_dt is None:
        return ""

    planned_display = planned_dt.strftime("%d-%b-%Y %H:%M:%S")
    planned_letter = column_number_to_letter(col_planned)
    status_letter = column_number_to_letter(col_status)
    delay_letter = column_number_to_letter(col_time_delay)

    # --- Planned: only fill when blank ---
    current_planned = get_value(row_values, col_planned)
    if is_blank_value(current_planned):
        updates.append(
            make_update(
                f"'{MAIN_SHEET}'!{planned_letter}{row}",
                planned_display,
            )
        )

    # --- Status + Time Delay depend on the Actual value ---
    # An explicit override (used by the auto-filled internal stages)
    # takes priority over whatever is currently in the sheet cell,
    # because the sheet was read once at the start of the run and the
    # auto-filled value only exists in memory + the pending updates.
    if actual_override is not None:
        actual_dt = actual_override
    else:
        actual_dt = to_datetime(get_value(row_values, col_actual))

    if actual_dt is None:
        return ""

    diff_seconds = (actual_dt - planned_dt).total_seconds()

    # Same calendar day counts as On Time even when the clock times
    # differ. Planned carries a time-of-day (e.g. 20:18) while Actual
    # is recorded at midnight, so a same-day stage would otherwise be
    # mislabelled "Early" and left with an unreadable/blank cell.
    if actual_dt.date() == planned_dt.date():
        status = "On Time"
    elif diff_seconds < 0:
        status = "Early"
    elif diff_seconds > 0:
        status = "Delay"
    else:
        status = "On Time"

    updates.append(
        make_update(
            f"'{MAIN_SHEET}'!{status_letter}{row}",
            status,
        )
    )
    updates.append(
        make_update(
            f"'{MAIN_SHEET}'!{delay_letter}{row}",
            format_time_delay(actual_dt, planned_dt),
        )
    )

    # --- Alert formatting on the Status cell ---
    # --- Colour logic removed per requirement ---
    # This stage now writes Status text + Time Delay only; no cell
    # highlighting. sheet_id and format_requests remain in the
    # signature (and format_requests stays empty) so the rest of the
    # pipeline and apply_updates() do not need to change.
    return status


# Offsets of the internal stages between X (pick) and AO (POD planned), in order:
# AA = X + 30m, AF = AB + 6h, AK = AG + 30m, AO = AL + 24h.
STAGE_OFFSETS_AFTER_X = (
    timedelta(minutes=30), timedelta(hours=6), timedelta(minutes=30), timedelta(hours=24))


def stage6_planned(row_values, actual_AL, holidays):
    """Planned AO for the POD stage -> (datetime, basis) or (None, None).

    Normal: AO = AL + 24h (basis "AL").
    2026-10-31 fallback: X (Pick Actual) blank, so the chain X -> AB -> AG -> AL never
    ran, but AP (POD date) is a date. Without AO, AQ/AR stayed blank on 198 rows that
    have a POD. The next known EARLIER date of the row stands in for X: S (punch date),
    else A (form received). AO = that date + 30m + 6h + 30m + 24h, each step skipping
    holidays exactly like the chain (basis "S" / "A"). AA..AL are NOT written, and a
    row without a POD date gets no AO from a guess.
    """
    if actual_AL is not None:
        return add_time_skip_holidays(actual_AL, timedelta(hours=24), holidays), "AL"
    if to_datetime(get_value(row_values, COL_X)) is not None:
        return None, None
    if to_datetime(get_value(row_values, COL_AP)) is None:
        return None, None
    for column, basis in ((COL_S, "S"), (COL_FORM_RECEIVED, "A")):
        planned = to_datetime(get_value(row_values, column))
        if planned is None:
            continue
        for delta in STAGE_OFFSETS_AFTER_X:
            planned = add_time_skip_holidays(planned, delta, holidays)
        return planned, basis
    return None, None


def resolve_stage_actual(row_values, col_actual, planned_dt, row, updates):
    """Return the Actual datetime to use for an internal handoff stage.

    Internal stages (courier assignment, label generation, pickup) do
    not get their own logged timestamp. So:

      - If the Actual cell already has a value, parse and return it.
        A real recorded Actual is NEVER overwritten.
      - If the Actual cell is blank, treat the stage as completed on
        plan: return the Planned datetime AND queue writing it into the
        blank Actual cell, so the sheet and the in-memory cascade agree
        and the next stage's Planned can be computed in the same run.
    """
    existing = to_datetime(get_value(row_values, col_actual))
    if existing is not None:
        return existing

    letter = column_number_to_letter(col_actual)
    updates.append(
        make_update(
            f"'{MAIN_SHEET}'!{letter}{row}",
            planned_dt.strftime("%d-%b-%Y %H:%M:%S"),
        )
    )
    return planned_dt

# ============================================================
# === END RECONSTRUCTED HELPERS ==============================
# ============================================================


# ============================================================
# GOOGLE SHEETS CONNECTION
# ============================================================

def connect_google_sheets():
    print("\nConnecting to Google Sheets...")

    scopes = ["https://www.googleapis.com/auth/spreadsheets"]

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

    sheets = build(
        "sheets",
        "v4",
        http=authorized_http,
        cache_discovery=False,
    )

    print("Google Sheets connection successful.")
    return sheets


def get_sheet_id(sheets, sheet_name):
    response = sheets.spreadsheets().get(
        spreadsheetId=OUTPUT_SPREADSHEET_ID,
        fields="sheets(properties(sheetId,title,gridProperties(rowCount,columnCount)))",
    ).execute()

    for sheet in response.get("sheets", []):
        properties = sheet.get("properties", {})
        if properties.get("title") == sheet_name:
            return properties.get("sheetId"), properties.get("gridProperties", {})

    raise ValueError(
        f"Sheet '{sheet_name}' was not found in "
        f"spreadsheet {OUTPUT_SPREADSHEET_ID}."
    )


def column_number_to_letter(column_number):
    result = ""
    while column_number:
        column_number, remainder = divmod(column_number - 1, 26)
        result = chr(65 + remainder) + result
    return result


def read_main_sheet(sheets):
    print()
    print("=" * 80)
    print("READING MAIN GOOGLE SHEET")
    print("=" * 80)

    result = sheets.spreadsheets().values().get(
        spreadsheetId=OUTPUT_SPREADSHEET_ID,
        range=f"'{MAIN_SHEET}'!A{DATA_START_ROW}:{LAST_DATA_COLUMN}",
        valueRenderOption="FORMATTED_VALUE",
        dateTimeRenderOption="FORMATTED_STRING",
    ).execute()

    values = result.get("values", [])

    print(
        f"Spreadsheet : {OUTPUT_SPREADSHEET_ID}"
    )
    print(
        f"Sheet       : {MAIN_SHEET}"
    )
    print(
        f"Rows read   : {len(values):,}"
    )

    return values


def get_value(row_values, column_number):
    index = column_number - 1

    if index >= len(row_values):
        return ""

    value = row_values[index]

    if value is None:
        return ""

    return str(value).strip()


def is_blank_value(value):
    return value is None or str(value).strip() == ""


def read_holidays_from_google_sheet(sheets):
    """
    Read Holiday dates directly from the Holiday tab in the SAME
    output spreadsheet.

    Holiday dates are expected in column A. A header row is harmless
    because normalize_date() returns None for non-date text.
    """
    print()
    print("=" * 80)
    print("READING HOLIDAYS FROM GOOGLE SHEET")
    print("=" * 80)

    result = sheets.spreadsheets().values().get(
        spreadsheetId=OUTPUT_SPREADSHEET_ID,
        range=f"'{HOLIDAY_SHEET}'!A:A",
        valueRenderOption="FORMATTED_VALUE",
        dateTimeRenderOption="FORMATTED_STRING",
    ).execute()

    values = result.get("values", [])
    holidays = set()

    for row in values:
        if not row:
            continue

        holiday = normalize_date(row[0])

        if holiday:
            holidays.add(holiday)

    print(
        f"Holiday dates loaded: {len(holidays)}"
    )

    if holidays:
        print("Holiday dates:")

        for holiday in sorted(holidays):
            print(
                f"  {holiday.strftime('%d-%m-%Y')}"
            )

    return holidays


def make_update(range_name, value):
    return {
        "range": range_name,
        "values": [[value]],
    }


def make_red_format_request(sheet_id, cell_range):
    """
    Return a Google Sheets API request that makes the target cell
    red when an alert threshold is reached.
    """
    return {
        "repeatCell": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": cell_range["row"] - 1,
                "endRowIndex": cell_range["row"],
                "startColumnIndex": cell_range["column"] - 1,
                "endColumnIndex": cell_range["column"],
            },
            "cell": {
                "userEnteredFormat": {
                    "backgroundColor": {
                        "red": 1.0,
                        "green": 0.0,
                        "blue": 0.0,
                    },
                    "textFormat": {
                        "foregroundColor": {
                            "red": 1.0,
                            "green": 1.0,
                            "blue": 1.0,
                        },
                        "bold": True,
                    },
                }
            },
            "fields": "userEnteredFormat.backgroundColor,userEnteredFormat.textFormat",
        }
    }


def make_clear_format_request(sheet_id, cell_range):
    """
    Reset the target status cell to a plain, readable look (white
    background, black text) when it no longer meets an alert threshold.

    IMPORTANT:
    Do NOT use empty color objects like {} here. Google Sheets treats
    a color with missing red/green/blue channels as 0,0,0 = BLACK, so
    an empty backgroundColor paints the cell black and an empty
    foregroundColor paints the text black -> the cell looks blank.
    Setting explicit white/black keeps the cell readable.
    """
    return {
        "repeatCell": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": cell_range["row"] - 1,
                "endRowIndex": cell_range["row"],
                "startColumnIndex": cell_range["column"] - 1,
                "endColumnIndex": cell_range["column"],
            },
            "cell": {
                "userEnteredFormat": {
                    "backgroundColor": {
                        "red": 1.0,
                        "green": 1.0,
                        "blue": 1.0,
                    },
                    "textFormat": {
                        "foregroundColor": {
                            "red": 0.0,
                            "green": 0.0,
                            "blue": 0.0,
                        },
                        "bold": False,
                    },
                }
            },
            "fields": "userEnteredFormat.backgroundColor,userEnteredFormat.textFormat",
        }
    }


def make_green_format_request(sheet_id, cell_range):
    """
    Mark the status cell green when the stage landed on the planned
    day (On Time).
    """
    return {
        "repeatCell": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": cell_range["row"] - 1,
                "endRowIndex": cell_range["row"],
                "startColumnIndex": cell_range["column"] - 1,
                "endColumnIndex": cell_range["column"],
            },
            "cell": {
                "userEnteredFormat": {
                    "backgroundColor": {
                        "red": 0.20,
                        "green": 0.66,
                        "blue": 0.33,
                    },
                    "textFormat": {
                        "foregroundColor": {
                            "red": 1.0,
                            "green": 1.0,
                            "blue": 1.0,
                        },
                        "bold": True,
                    },
                }
            },
            "fields": "userEnteredFormat.backgroundColor,userEnteredFormat.textFormat",
        }
    }


def execute_with_retry(api_call_factory, operation_name):
    """
    Execute a Google Sheets API request with exponential backoff for
    temporary 429 / 5xx errors.

    This prevents the script from stopping when the per-minute write
    quota is temporarily exhausted.

    NOTE: api_call_factory MUST return an un-executed request object.
    execute_with_retry is the ONLY place .execute() is called, so it
    can rebuild and retry the request on quota errors.
    """
    delay = INITIAL_RETRY_SECONDS

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            return api_call_factory().execute()

        except Exception as e:
            status = getattr(
                getattr(e, "resp", None),
                "status",
                None,
            )

            if status not in (429, 500, 502, 503, 504):
                raise

            if attempt == MAX_RETRIES:
                print(
                    f"\n{operation_name} failed after "
                    f"{MAX_RETRIES} retries."
                )
                raise

            # Add jitter so repeated retries do not hit the quota
            # at exactly the same time.
            sleep_seconds = delay + random.uniform(0, 1)

            print(
                f"\nGoogle Sheets returned HTTP {status} for "
                f"{operation_name}."
            )
            print(
                f"Retry {attempt}/{MAX_RETRIES} in "
                f"{sleep_seconds:.1f} seconds..."
            )

            time.sleep(sleep_seconds)
            delay = min(delay * 2, 60)



def load_holidays(sheets):
    """
    Load holiday dates from the Google Sheets 'Holiday' tab.

    Holiday dates are expected in column A. Blank cells, headers, and
    unparseable values are ignored.
    """
    print()
    print("=" * 80)
    print("READING HOLIDAYS")
    print("=" * 80)

    # FIX: the lambda must return an un-executed request. Previously
    # it ended with .execute(), so it returned a dict and
    # execute_with_retry then failed on dict.execute().
    result = execute_with_retry(
        lambda: sheets.spreadsheets().values().get(
            spreadsheetId=OUTPUT_SPREADSHEET_ID,
            range=f"{HOLIDAY_SHEET}!A:A",
            majorDimension="ROWS",
        ),
        "Loading holidays",
    )

    values = result.get("values", [])
    holidays = set()

    for row in values:
        if not row:
            continue

        value = row[0]
        holiday = normalize_date(value)

        if holiday:
            holidays.add(holiday)

    print(f"Holiday dates loaded: {len(holidays)}")

    if holidays:
        print("First holidays:")
        for holiday in sorted(holidays)[:10]:
            print(f"  {holiday.strftime('%d-%m-%Y')}")

    return holidays


def apply_updates(
    sheets,
    sheet_id,
    updates,
    format_requests,
):
    """
    Write values and formatting to Google Sheets using large batches.

    IMPORTANT:
    The old version used BATCH_SIZE=500. With ~47,502 individual cell
    updates this created ~96 write requests, which exceeded the
    60 write-requests/minute quota.

    This version uses 5,000 updates per API call, reducing the value
    writes to roughly 10 requests. It also retries HTTP 429 responses.
    """

    if not updates:
        print("\nNo Google Sheet cells require updating.")
    else:
        print()
        print("=" * 80)
        print("WRITING UPDATES TO GOOGLE SHEET")
        print("=" * 80)
        print(
            f"Total cell updates: {len(updates):,}"
        )
        print(
            f"Batch size: {BATCH_SIZE:,} cells/API request"
        )

        updated = 0

        for start_index in range(0, len(updates), BATCH_SIZE):
            batch = updates[
                start_index:start_index + BATCH_SIZE
            ]

            execute_with_retry(
                lambda batch=batch: (
                    sheets.spreadsheets().values().batchUpdate(
                        spreadsheetId=OUTPUT_SPREADSHEET_ID,
                        body={
                            "valueInputOption": "USER_ENTERED",
                            "data": batch,
                        },
                    )
                ),
                f"value batch {start_index // BATCH_SIZE + 1}",
            )

            updated += len(batch)

            print(
                f"Updated {updated:,} / {len(updates):,} cells"
            )

        # Formatting requests are also write requests. Send them in
        # large groups rather than making hundreds of API calls.
        if format_requests:
            print(
                f"Applying red alert formatting to "
                f"{len(format_requests):,} cells..."
            )

            FORMAT_BATCH_SIZE = 5000

            formatting_done = 0

            for start_index in range(
                0,
                len(format_requests),
                FORMAT_BATCH_SIZE,
            ):
                batch = format_requests[
                    start_index:start_index + FORMAT_BATCH_SIZE
                ]

                execute_with_retry(
                    lambda batch=batch: (
                        sheets.spreadsheets().batchUpdate(
                            spreadsheetId=OUTPUT_SPREADSHEET_ID,
                            body={
                                "requests": batch,
                            },
                        )
                    ),
                    f"formatting batch {start_index // FORMAT_BATCH_SIZE + 1}",
                )

                formatting_done += len(batch)

                print(
                    f"Formatted {formatting_done:,} / "
                    f"{len(format_requests):,} alert cells"
                )

    return len(updates)


def main():

    print()
    print("=" * 80)
    print("PLANNED DATE / STATUS / TIME DELAY - GOOGLE SHEETS")
    print("=" * 80)

    print(
        f"Output Spreadsheet ID: {OUTPUT_SPREADSHEET_ID}"
    )
    print(
        f"Main Sheet: {MAIN_SHEET}"
    )
    print(
        f"Holiday Sheet: {HOLIDAY_SHEET}"
    )

    # ========================================================
    # CONNECT TO GOOGLE SHEETS
    # ========================================================

    sheets = connect_google_sheets()

    # ========================================================
    # VERIFY MAIN + HOLIDAY SHEETS
    # ========================================================

    print()
    print("=" * 80)
    print("CHECKING GOOGLE SHEETS TABS")
    print("=" * 80)

    main_sheet_id, main_grid = get_sheet_id(
        sheets,
        MAIN_SHEET,
    )

    holiday_sheet_id, holiday_grid = get_sheet_id(
        sheets,
        HOLIDAY_SHEET,
    )

    print(
        f"Main tab found    : '{MAIN_SHEET}' "
        f"(sheet ID: {main_sheet_id})"
    )
    print(
        f"Holiday tab found : '{HOLIDAY_SHEET}' "
        f"(sheet ID: {holiday_sheet_id})"
    )

    # ========================================================
    # LOAD HOLIDAYS FROM SAME OUTPUT SPREADSHEET
    # ========================================================

    holidays = load_holidays(sheets)

    # ========================================================
    # READ CURRENT FULL FLOW FMS
    # ========================================================

    values = read_main_sheet(sheets)

    # ========================================================
    # COUNTERS
    # ========================================================

    rows_processed = 0

    stage1_planned = 0
    stage2_planned = 0
    stage3_planned = 0
    stage4_planned = 0
    stage5_planned = 0
    stage6_planned_count = 0
    stage6_fallback_count = 0

    status_counts = {
        "Early": 0,
        "Delay": 0,
        "On Time": 0,
    }

    updates = []
    format_requests = []

    # ========================================================
    # PROCESS ROWS
    # ========================================================

    print()
    print("=" * 80)
    print("PROCESSING FULL FLOW FMS")
    print("=" * 80)

    for offset, row_values in enumerate(values):

        row = DATA_START_ROW + offset
        rows_processed += 1

        # ====================================================
        # STAGE 1
        # R = A + 24 hours
        # S = Actual
        # T = Status
        # U = Time Delay
        # ====================================================

        form_received = to_datetime(
            get_value(
                row_values,
                COL_FORM_RECEIVED,
            )
        )

        if form_received is not None:

            planned_R = add_time_skip_holidays(
                form_received,
                timedelta(hours=24),
                holidays,
            )

            status = process_stage(
                row_values,
                row,
                COL_R,
                COL_S,
                COL_T,
                COL_U,
                planned_R,
                updates,
                format_requests,
                main_sheet_id,
            )

            stage1_planned += 1

            if status in status_counts:
                status_counts[status] += 1

        # ====================================================
        # STAGE 2
        # W = S + 24 hours
        # ====================================================

        actual_S = to_datetime(
            get_value(
                row_values,
                COL_S,
            )
        )

        if actual_S is not None:

            planned_W = add_time_skip_holidays(
                actual_S,
                timedelta(hours=24),
                holidays,
            )

            status = process_stage(
                row_values,
                row,
                COL_W,
                COL_X,
                COL_Y,
                COL_Z,
                planned_W,
                updates,
                format_requests,
                main_sheet_id,
            )

            stage2_planned += 1

        # ====================================================
        # STAGE 3  (AA = X + 30 min ; AB = Actual)
        #
        # Internal handoff. If AB is blank we treat it as completed on
        # plan (AB = AA) and write it into the Actual cell. The value is
        # also carried forward in memory so Stage 4's Planned can be
        # computed in the same run. A real AB value is never overwritten.
        # ====================================================

        actual_X = to_datetime(
            get_value(
                row_values,
                COL_X,
            )
        )

        actual_AB = None

        if actual_X is not None:

            planned_AA = add_time_skip_holidays(
                actual_X,
                timedelta(minutes=30),
                holidays,
            )

            actual_AB = resolve_stage_actual(
                row_values,
                COL_AB,
                planned_AA,
                row,
                updates,
            )

            process_stage(
                row_values,
                row,
                COL_AA,
                COL_AB,
                COL_AC,
                COL_AD,
                planned_AA,
                updates,
                format_requests,
                main_sheet_id,
                actual_override=actual_AB,
            )

            stage3_planned += 1

        # ====================================================
        # STAGE 4  (AF = AB + 6 hours ; AG = Actual)
        # ====================================================

        actual_AG = None

        if actual_AB is not None:

            planned_AF = add_time_skip_holidays(
                actual_AB,
                timedelta(hours=6),
                holidays,
            )

            actual_AG = resolve_stage_actual(
                row_values,
                COL_AG,
                planned_AF,
                row,
                updates,
            )

            process_stage(
                row_values,
                row,
                COL_AF,
                COL_AG,
                COL_AH,
                COL_AI,
                planned_AF,
                updates,
                format_requests,
                main_sheet_id,
                actual_override=actual_AG,
            )

            stage4_planned += 1

        # ====================================================
        # STAGE 5  (AK = AG + 30 min ; AL = Actual)
        # ====================================================

        actual_AL = None

        if actual_AG is not None:

            planned_AK = add_time_skip_holidays(
                actual_AG,
                timedelta(minutes=30),
                holidays,
            )

            actual_AL = resolve_stage_actual(
                row_values,
                COL_AL,
                planned_AK,
                row,
                updates,
            )

            process_stage(
                row_values,
                row,
                COL_AK,
                COL_AL,
                COL_AM,
                COL_AN,
                planned_AK,
                updates,
                format_requests,
                main_sheet_id,
                actual_override=actual_AL,
            )

            stage5_planned += 1

        # ====================================================
        # STAGE 6  (AO = AL + 24 hours ; AP = Actual)
        #
        # Final stage. AP is a REAL external actual (delivery / POD) and
        # is NOT auto-filled. Status/Time Delay (AQ/AR) are computed only
        # when AP already has a value.
        # X blank + POD date present: AO from S (else A), see stage6_planned.
        # ====================================================

        planned_AO, ao_basis = stage6_planned(row_values, actual_AL, holidays)

        if planned_AO is not None:

            process_stage(
                row_values,
                row,
                COL_AO,
                COL_AP,
                COL_AQ,
                COL_AR,
                planned_AO,
                updates,
                format_requests,
                main_sheet_id,
            )

            stage6_planned_count += 1
            if ao_basis != "AL":
                stage6_fallback_count += 1

        if rows_processed % 500 == 0:
            print(
                f"Processed {rows_processed:,} rows..."
            )

    # ========================================================
    # WRITE TO GOOGLE SHEET
    # ========================================================

    updated_cells = apply_updates(
        sheets,
        main_sheet_id,
        updates,
        format_requests,
    )

    # ========================================================
    # SUMMARY
    # ========================================================

    print()
    print("=" * 80)
    print("PLANNED / STATUS UPDATE COMPLETED")
    print("=" * 80)

    print(
        f"Rows processed          : {rows_processed:,}"
    )
    print(
        f"Google Sheet cells updated: {updated_cells:,}"
    )
    print(
        f"Red alert cells           : {len(format_requests):,}"
    )

    print("-" * 80)

    print(
        f"R planned               : {stage1_planned:,}"
    )
    print(
        f"W planned               : {stage2_planned:,}"
    )
    print(
        f"AA planned              : {stage3_planned:,}"
    )
    print(
        f"AF planned              : {stage4_planned:,}"
    )
    print(
        f"AK planned              : {stage5_planned:,}"
    )
    print(
        f"AO planned              : {stage6_planned_count:,}"
        f"  (of which X blank, from S/A: {stage6_fallback_count:,})"
    )

    print("-" * 80)

    print(
        f"Early                   : {status_counts['Early']:,}"
    )
    print(
        f"Delay                   : {status_counts['Delay']:,}"
    )
    print(
        f"On Time                 : {status_counts['On Time']:,}"
    )

    print("-" * 80)

    print(
        f"Holidays loaded         : {len(holidays):,}"
    )

    print("=" * 80)

    print()
    print("Updated Google Spreadsheet:")
    print(OUTPUT_SPREADSHEET_ID)
    print(f"Updated tab: {MAIN_SHEET}")
    print(f"Holiday source: {HOLIDAY_SHEET}")
    print()




# ============================================================
# RUN
# ============================================================

if __name__ == "__main__":

    main()
