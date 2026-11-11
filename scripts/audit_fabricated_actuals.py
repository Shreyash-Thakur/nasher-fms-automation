"""
Audit (and, only with owner approval, repair) the FABRICATED stage Actuals in
GT Full Flow FMS  —  2026-11-11.

Until 2026-11-11 the planned-status step (resolve_stage_actual) wrote the Planned
datetime into every blank courier / label / dispatch Actual:
    AB = AA (courier), AG = AF (label), AL = AK (dispatch)
and then wrote AC / AH / AM "On Time" + AD / AI / AN delay from that fake value.
A real Actual equal to its Planned to the second is practically impossible, so an
exact match is the fingerprint of a fabricated value.

DRY RUN BY DEFAULT: reads the sheet, writes a CSV report, changes nothing.

    python scripts/audit_fabricated_actuals.py                 # dry run -> CSV
    python scripts/audit_fabricated_actuals.py --csv out.csv   # CSV elsewhere
    python scripts/audit_fabricated_actuals.py --apply-audit-tab
        # also writes the report to a NEW tab "Fabricated Actuals Audit"
    python scripts/audit_fabricated_actuals.py --apply
        # REPAIR - needs owner approval. First copies every cell it will clear
        # (with its A1 address) to a NEW tab "Fabricated Actuals Backup <stamp>",
        # then blanks the fingerprint Actual + its Status + Time Delay cells.
        # Planned cells are NOT touched.

Credentials: GCP_SA_KEY env var (JSON or base64), else --sa-file, else the
planned-status script's SERVICE_ACCOUNT_FILE.
"""
import argparse
import csv
import os
import sys
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import gt_full_flow_fms_planned_status_google_sheet_alerts_quota_fixed_v3 as ps  # noqa: E402

AUDIT_TAB = "Fabricated Actuals Audit"
BACKUP_TAB_PREFIX = "Fabricated Actuals Backup"
READ_LAST_COLUMN = "AS"

COL_GT = 3        # C
COL_ORDER = 16    # P
COL_AE = 31       # AE = AWB
COL_AS = 45       # AS = Dispatched Qty

# (stage name, planned, actual, status, delay)
STAGES = (
    ("courier", ps.COL_AA, ps.COL_AB, ps.COL_AC, ps.COL_AD),
    ("label", ps.COL_AF, ps.COL_AG, ps.COL_AH, ps.COL_AI),
    ("dispatch", ps.COL_AK, ps.COL_AL, ps.COL_AM, ps.COL_AN),
)

REPORT_HEADER = ["Row", "GT", "Order No", "Stage", "Planned col", "Planned",
                 "Actual col", "Actual", "Status col", "Status", "Delay col", "Delay",
                 "AWB (AE)", "Dispatched Qty (AS)"]


def L(col):
    return ps.column_number_to_letter(col)


def find_fingerprints(values, start_row=ps.DATA_START_ROW):
    """Report rows (lists, REPORT_HEADER order) where Actual == Planned exactly
    (non-blank) for a courier / label / dispatch stage. Pure: no I/O."""
    out = []
    for offset, row_values in enumerate(values):
        row = start_row + offset
        for name, c_plan, c_act, c_stat, c_delay in STAGES:
            planned = ps.get_value(row_values, c_plan)
            actual = ps.get_value(row_values, c_act)
            if not planned or actual != planned:
                continue
            out.append([
                row, ps.get_value(row_values, COL_GT), ps.get_value(row_values, COL_ORDER),
                name, L(c_plan), planned, L(c_act), actual,
                L(c_stat), ps.get_value(row_values, c_stat),
                L(c_delay), ps.get_value(row_values, c_delay),
                ps.get_value(row_values, COL_AE), ps.get_value(row_values, COL_AS),
            ])
    return out


def cells_to_clear(report):
    """[(a1, original value)] for the Actual + Status + Delay of each finding."""
    cells = []
    for r in report:
        row = r[0]
        cells.append((f"{r[6]}{row}", r[7]))
        cells.append((f"{r[8]}{row}", r[9]))
        cells.append((f"{r[10]}{row}", r[11]))
    return cells


def summarize(report, rows_read):
    by_stage = {name: 0 for name, *_ in STAGES}
    rows = {}
    for r in report:
        by_stage[r[3]] += 1
        rows.setdefault(r[0], set()).add(r[3])
    all_three = sum(1 for s in rows.values() if len(s) == 3)
    lines = [f"Rows read                 : {rows_read:,}",
             f"Rows with any fingerprint : {len(rows):,}",
             f"Rows with all three       : {all_three:,}"]
    lines += [f"  {name:<9} Actual==Planned : {n:,}" for name, n in by_stage.items()]
    lines.append(f"Cells a repair would clear: {len(cells_to_clear(report)):,}")
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
            valueRenderOption="FORMATTED_VALUE",
            dateTimeRenderOption="FORMATTED_STRING",
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
    for start in range(0, len(rows), 5000):
        chunk = rows[start:start + 5000]
        ps.execute_with_retry(
            lambda chunk=chunk, start=start: sheets.spreadsheets().values().update(
                spreadsheetId=ps.OUTPUT_SPREADSHEET_ID,
                range=f"'{title}'!A{start + 1}",
                valueInputOption="RAW",
                body={"values": chunk},
            ),
            f"write tab {title}",
        )


def apply_repair(sheets, report, stamp):
    cells = cells_to_clear(report)
    if not cells:
        print("Nothing to clear.")
        return
    backup_tab = f"{BACKUP_TAB_PREFIX} {stamp}"
    backup = [["Sheet", "Cell", "Original value", "Backed up at"]]
    backup += [[ps.MAIN_SHEET, a1, value, stamp] for a1, value in cells]
    add_tab(sheets, backup_tab, backup)
    print(f"Backup tab written: '{backup_tab}' ({len(cells):,} cells)")

    ranges = [f"'{ps.MAIN_SHEET}'!{a1}" for a1, _ in cells]
    for start in range(0, len(ranges), 5000):
        chunk = ranges[start:start + 5000]
        ps.execute_with_retry(
            lambda chunk=chunk: sheets.spreadsheets().values().batchClear(
                spreadsheetId=ps.OUTPUT_SPREADSHEET_ID, body={"ranges": chunk}),
            "clear fabricated cells",
        )
    print(f"Cleared {len(ranges):,} cells (Actual + Status + Delay).")


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--csv", default=os.path.join("reports", "fabricated_actuals_audit.csv"))
    ap.add_argument("--sa-file", default=None)
    ap.add_argument("--apply-audit-tab", action="store_true",
                    help=f"also write the report to a NEW tab '{AUDIT_TAB}'")
    ap.add_argument("--apply", action="store_true",
                    help="REPAIR (owner approval needed): back up, then blank the cells")
    args = ap.parse_args(argv)

    sheets = connect(args.sa_file)
    values = read_sheet(sheets)
    report = find_fingerprints(values)

    os.makedirs(os.path.dirname(os.path.abspath(args.csv)), exist_ok=True)
    with open(args.csv, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(REPORT_HEADER)
        writer.writerows(report)

    print(summarize(report, len(values)))
    print(f"CSV report: {os.path.abspath(args.csv)}")

    stamp = datetime.now().strftime("%Y-%m-%d %H%M%S")
    if args.apply_audit_tab:
        add_tab(sheets, AUDIT_TAB, [REPORT_HEADER] + [list(map(str, r)) for r in report])
        print(f"Audit tab written: '{AUDIT_TAB}'")
    if args.apply:
        apply_repair(sheets, report, stamp)
    else:
        print("DRY RUN: no sheet cells were changed (pass --apply to repair).")


if __name__ == "__main__":
    main()
