"""
Find (and, only with owner approval, repair) GT Full Flow FMS dates whose month and
day were swapped  -  2026-11-11.

The Pick Pack Google Sheets for 03-12 Jul 2026 were made by importing the dd/mm/yyyy
Vinculum CSV with a US locale: every Pick Date with day <= 12 became a real date with
month and day swapped (12-Jul -> 07-Dec). The nightly job copied them into X (Pick
Actual), e.g. GT100726A3 row 5431 X = 07-Dec-2026. The job now undoes the swap
(fix_swapped_pick_date); this script repairs the cells already written.

Reason rule - a cell is swapped back only when ALL hold:
  - its day is <= 12 (and day != month), so a swap is possible;
  - the swapped date falls in [order date, order date + 30 days];
  - the current value is impossible: after today ("future") or before the order date
    ("before_order"). With --include-late also "late_over_30d": more than 30 days after
    the order date (the 07-11 Jul picks, e.g. 10-Jul pick stored as 07-Oct). Those rows
    are always listed in the CSV; --apply changes them only with --include-late.
Order date = A (Form Received Time), or S (Order Date) when A is blank.

DRY RUN BY DEFAULT: reads the sheet, writes a CSV report, changes nothing.

    python scripts/repair_swapped_pick_dates.py                    # dry run -> CSV
    python scripts/repair_swapped_pick_dates.py --columns X,S      # also check S
    python scripts/repair_swapped_pick_dates.py --apply [--include-late]
        # REPAIR - needs owner approval. First copies every cell it will change
        # (A1, old value, new value) to a NEW tab "Swapped Dates Backup <stamp>",
        # then writes the swapped-back value as a date serial (RAW number, so no
        # locale re-parsing; the cell keeps its date format). Planned / Status /
        # Delay are recomputed by the planned-status job on its next run.

Credentials: GCP_SA_KEY env var (JSON or base64), else --sa-file, else the
planned-status script's SERVICE_ACCOUNT_FILE.
"""
import argparse
import csv
import os
import sys
from datetime import date, datetime, timedelta

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import gt_full_flow_fms_planned_status_google_sheet_alerts_quota_fixed_v3 as ps  # noqa: E402

BACKUP_TAB_PREFIX = "Swapped Dates Backup"
READ_LAST_COLUMN = "X"
WINDOW_DAYS = 30
EPOCH = datetime(1899, 12, 30)

COL_A, COL_C, COL_P, COL_S, COL_X = 1, 3, 16, 19, 24
CHECK_COLUMNS = {"S": COL_S, "X": COL_X}
DEFAULT_REASONS = ("future", "before_order")
LATE_REASON = "late_over_30d"

REPORT_HEADER = ["Row", "GT", "Order No", "Column", "Cell", "Order date",
                 "Current", "Proposed", "Reason", "Applied by --apply"]


def serial_to_datetime(value):
    """Sheets serial number (UNFORMATTED_VALUE + SERIAL_NUMBER) -> datetime, or None."""
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return None
    return EPOCH + timedelta(seconds=round(float(value) * 86400))


def datetime_to_serial(value):
    return round((value - EPOCH).total_seconds() / 86400, 8)


def _cell(row_values, col):
    return row_values[col - 1] if col - 1 < len(row_values) else ""


def order_date(row_values):
    """Midnight of A (Form Received Time), else of S (Order Date), else None."""
    for col in (COL_A, COL_S):
        dt = serial_to_datetime(_cell(row_values, col))
        if dt is not None:
            return datetime.combine(dt.date(), datetime.min.time())
    return None


def classify(current, base, today):
    """(proposed datetime, reason) when current is a provable month/day swap, else None."""
    if current.day > 12 or current.day == current.month:
        return None
    try:
        proposed = current.replace(month=current.day, day=current.month)
    except ValueError:
        return None
    if not base <= proposed <= base + timedelta(days=WINDOW_DAYS, hours=23, minutes=59, seconds=59):
        return None
    if current.date() > today:
        return proposed, "future"
    if current < base:
        return proposed, "before_order"
    if current > base + timedelta(days=WINDOW_DAYS, hours=23, minutes=59, seconds=59):
        return proposed, LATE_REASON
    return None


def find_swaps(values, columns=("X",), today=None, include_late=False,
               start_row=ps.DATA_START_ROW):
    """Report rows (REPORT_HEADER order) for swapped cells. values: UNFORMATTED rows
    (dates as serial numbers) from A. Pure: no I/O."""
    today = today or date.today()
    out = []
    for offset, row_values in enumerate(values):
        row = start_row + offset
        base = order_date(row_values)
        if base is None:
            continue
        for letter in columns:
            current = serial_to_datetime(_cell(row_values, CHECK_COLUMNS[letter]))
            if current is None:
                continue
            hit = classify(current, base, today)
            if not hit:
                continue
            proposed, reason = hit
            applied = reason in DEFAULT_REASONS or (include_late and reason == LATE_REASON)
            out.append([row, _cell(row_values, COL_C), _cell(row_values, COL_P), letter,
                        f"{letter}{row}", base.strftime("%d-%b-%Y"),
                        current.strftime("%d-%b-%Y %H:%M:%S"),
                        proposed.strftime("%d-%b-%Y %H:%M:%S"), reason,
                        "yes" if applied else "no"])
    return out


def summarize(report, rows_read):
    lines = [f"Rows read            : {rows_read:,}",
             f"Swapped cells found  : {len(report):,}"]
    by = {}
    for r in report:
        by.setdefault((r[3], r[8]), []).append(r)
    for (col, reason), rs in sorted(by.items()):
        gts = sorted({r[1] for r in rs})
        lines.append(f"  {col} {reason:<14}: {len(rs):,} cells, {len(gts)} GTs, rows "
                     f"{rs[0][0]}-{rs[-1][0]}, proposed {min(r[7][:11] for r in rs)}"
                     f" .. {max(r[7][:11] for r in rs)}")
    lines.append(f"Cells --apply would change: {sum(1 for r in report if r[9] == 'yes'):,}")
    return "\n".join(lines)


# ---------------------------------------------------------------- Google I/O

def connect(sa_file):
    if sa_file and not os.environ.get("GCP_SA_KEY"):
        ps.SERVICE_ACCOUNT_FILE = sa_file
    return ps.connect_google_sheets()


def read_sheet(sheets):
    result = ps.execute_with_retry(
        lambda: sheets.spreadsheets().values().get(
            spreadsheetId=ps.OUTPUT_SPREADSHEET_ID,
            range=f"'{ps.MAIN_SHEET}'!A{ps.DATA_START_ROW}:{READ_LAST_COLUMN}",
            valueRenderOption="UNFORMATTED_VALUE",
            dateTimeRenderOption="SERIAL_NUMBER",
        ),
        "read Full Flow FMS",
    )
    return result.get("values", [])


def add_tab(sheets, title, rows):
    """Create a NEW tab (fails if it already exists) and write rows to it."""
    ps.execute_with_retry(
        lambda: sheets.spreadsheets().batchUpdate(
            spreadsheetId=ps.OUTPUT_SPREADSHEET_ID,
            body={"requests": [{"addSheet": {"properties": {
                "title": title,
                "gridProperties": {"rowCount": max(len(rows) + 10, 100),
                                   "columnCount": max(len(rows[0]), 5)}}}}]},
        ),
        f"add tab {title}",
    )
    ps.execute_with_retry(
        lambda: sheets.spreadsheets().values().update(
            spreadsheetId=ps.OUTPUT_SPREADSHEET_ID,
            range=f"'{title}'!A1",
            valueInputOption="RAW",
            body={"values": rows},
        ),
        f"write tab {title}",
    )


def apply_repair(sheets, report, stamp):
    todo = [r for r in report if r[9] == "yes"]
    if not todo:
        print("Nothing to change.")
        return
    backup_tab = f"{BACKUP_TAB_PREFIX} {stamp}"
    backup = [["Sheet", "Cell", "GT", "Original value", "New value", "Reason", "Backed up at"]]
    backup += [[ps.MAIN_SHEET, r[4], r[1], r[6], r[7], r[8], stamp] for r in todo]
    add_tab(sheets, backup_tab, backup)
    print(f"Backup tab written: '{backup_tab}' ({len(todo):,} cells)")

    data = [{"range": f"'{ps.MAIN_SHEET}'!{r[4]}",
             "values": [[datetime_to_serial(datetime.strptime(r[7], "%d-%b-%Y %H:%M:%S"))]]}
            for r in todo]
    for start in range(0, len(data), 500):
        chunk = data[start:start + 500]
        ps.execute_with_retry(
            lambda chunk=chunk: sheets.spreadsheets().values().batchUpdate(
                spreadsheetId=ps.OUTPUT_SPREADSHEET_ID,
                body={"valueInputOption": "RAW", "data": chunk}),
            "write swapped-back dates",
        )
    print(f"Changed {len(data):,} cells (date serials).")


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--csv", default=os.path.join("reports", "swapped_dates_repair.csv"))
    ap.add_argument("--columns", default="X",
                    help="comma list of columns to check: X (Pick Actual), S (Order Date)")
    ap.add_argument("--include-late", action="store_true",
                    help=f"--apply also changes '{LATE_REASON}' cells")
    ap.add_argument("--sa-file", default=None)
    ap.add_argument("--apply", action="store_true",
                    help="REPAIR (owner approval needed): back up, then write the cells")
    args = ap.parse_args(argv)
    columns = [c.strip().upper() for c in args.columns.split(",") if c.strip()]
    bad = [c for c in columns if c not in CHECK_COLUMNS]
    if bad:
        ap.error(f"unsupported column(s) {bad}; use {sorted(CHECK_COLUMNS)}")

    sheets = connect(args.sa_file)
    values = read_sheet(sheets)
    report = find_swaps(values, columns, include_late=args.include_late)

    os.makedirs(os.path.dirname(os.path.abspath(args.csv)), exist_ok=True)
    with open(args.csv, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(REPORT_HEADER)
        writer.writerows(report)

    print(summarize(report, len(values)))
    print(f"CSV report: {os.path.abspath(args.csv)}")

    if args.apply:
        apply_repair(sheets, report, datetime.now().strftime("%Y-%m-%d %H%M%S"))
    else:
        print("DRY RUN: no sheet cells were changed (pass --apply to repair).")


if __name__ == "__main__":
    main()
