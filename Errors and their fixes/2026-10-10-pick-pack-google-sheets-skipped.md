# GT Full Flow FMS: Warehouse Team (V) and Pick date (X) blank for every row from May 2026

- **Date:** 2026-10-10
- **Reported as:** "in the GT Full Flow FMS, Warehouse Team (V) and Pick date (X / "Actual.1") are blank for every row from May 2026 onward."
- **Symptom:** nightly run 38004155464 (2026-10-09 23:23 UTC) log:
  `Pick Pack CSV files found: 10` (all `PickPackReport_2026-04-xx_to_2026-04-yy.csv`),
  `Total Pick Pack rows: 18,870`, `Rows with Order No anchored to a Vinculum line: 5,541 / refresh cells: 0`,
  and the fill-rate table on open rows:

  | month | rows | V | X |
  |---|---|---|---|
  | 2026-04 | 1385 | 1148 | 1148 |
  | 2026-05 | 999 | 0 | 0 |
  | 2026-06 | 1129 | 0 | 0 |
  | 2026-07 | 855 | 0 | 0 |
  | 2026-08 | 467 | 0 | 0 |
  | 2026-09 | 1315 | 0 | 0 |
  | 2026-10 | 178 | 0 | 0 |

- **Where:**
  - `scripts/gt_full_flow_fms_yes_no_all_or_latest.py:975` (before this fix), in `read_pick_pack_files`:
    `csv_files = [f for f in files if f["name"].lower().endswith(".csv")]`
  - same file `parse_pick_date` (before: `:768`) — last-resort `pd.to_datetime(value, dayfirst=True)`.
  - same file `plan_existing_writes` — rows whose P is filled but not in Order Enquiry got no V/X.
- **Root cause:** since May 2026 the daily Pick Pack reports land in Drive folder
  `1Xi66o4A0I3GdtlNrh61XcEWNkT5UKmey` as **native Google Sheets** (mimeType
  `application/vnd.google-apps.spreadsheet`, names like `PickPackReport_08.10.2026`, no extension), so the
  `.endswith(".csv")` filter dropped all 118 of them; the Pick Pack lookup only ever held April orders, so
  no row from May on could get V/X — on append or in the nightly refresh.
  Two latent faults sat behind it and would have bitten once the sheets were read:
  1. The May–June sheets (`PickPackReport_2026-05-01_to_2026-05-03` …) hold ISO dates
     (`2026-05-02 9:34:53`) and real date cells export as ISO text; the fallback
     `pd.to_datetime("2026-05-02 9:34:53", dayfirst=True)` returns **2026-02-05** (verified locally,
     pandas 2.2.3) — silently wrong X.
  2. At least one sheet's header cell is `﻿"Order No"` (BOM + literal CSV quotes from the import);
     the old header clean-up stripped the BOM but not the quotes, so "Order No" would be "missing" and the
     old code raised for the WHOLE concatenated frame.
- **Why it wasn't caught:** the reader never logged files it ignored, so "10 CSV files found" looked healthy;
  no alert on V/X fill rate; no test of the Drive listing / file types; the parser had no ISO test.
- **Fix:**
  - `read_pick_pack_files` reads every Pick Pack file by type: CSV (download, as before), Google Sheet
    (Drive `files().export_media(mimeType=xlsx)`), xlsx (download); all workbook tabs that carry
    Order No / Pick Date / Pick User are read (data tab need not be first). Other file types are listed as
    ignored. The listing was already paginated (`list_drive_files` follows `nextPageToken`) — now tested.
    xlsx export chosen over CSV export because CSV export returns only the first tab and the sheet's
    locale-formatted display text; xlsx keeps real dates as dates (read back as ISO text) and text as text.
    openpyxl was already in requirements.txt.
  - An unreadable file or one without the required columns is skipped with a `WARNING` line naming the
    file and the exact error; if nothing is readable an empty frame is returned (V/X just not filled that
    night) instead of aborting the run.
  - `parse_pick_date`: explicit ISO `yyyy-mm-dd[ hh:mm[:ss]]` branch before the day-first fallback; the
    fallback is wrapped so one bad value is blank, never fatal (dd/mm/yyyy `10:37 22` form unchanged).
  - Header clean-up strips surrounding quotes as well as BOM/NBSP.
  - `plan_existing_writes`: V/X logic moved into `_pick_pack_cells` (unchanged rules: blank cells only,
    closed rows never) and now also applied to rows whose P is filled but not found in Order Enquiry
    (keyed by the row's own P). Already-anchored rows (5,541) were already covered by the refresh — they
    were blank only because the lookup had no May+ orders.
  - Log: per-type file counts, per-file `Read: name [kind] -> n rows`, and a second
    "PROJECTED fill rate … after this run's fills/refreshes" table so a dry-run shows what V/X would become.
- **Evidence:**
  - Drive connector, 2026-10-10: folder holds 10 `text/csv` files (April) + 118 Google Sheets
    (95 `PickPackReport_dd.mm.yyyy`, 21 `PickPackReport_yyyy-mm-dd_to_yyyy-mm-dd` covering 2026-04-01 and
    2026-05-01 … 2026-06-28, 1 `29.06.2026`, 1 `PickPackReport_28.08.2026_29.08.2026`); no xlsx.
    Pick Date forms seen: `08/10/2026 10:37 22`, `28/07/2026 10:39 40`, `2026-05-02 9:34:53`, `2026-06-29 10:51:19`.
  - New tests `tests/test_pick_pack.py` against the OLD script: `8 failed, 43 passed`
    (reader, ISO dates ×3, V/X on P-without-line, projected table).
  - After the fix: `python -m pytest tests -q` → `51 passed in 0.42s`.
  - CI dry-run / write results: see "Live results" below.
- **Blast radius:** only `read_pick_pack_files` → `build_pick_pack_lookup` → V/X of Full Flow FMS (append
  path `build_full_flow` and nightly `plan_existing_writes`). Drive access is read-only
  (`drive.readonly` scope). POD and planned-status scripts do not read Pick Pack. Run time grows by one
  Drive export per sheet (~118 small files).
- **Still open:**
  - A sheet whose locale made Google parse a dd/mm text date as mm/dd at import time cannot be detected
    here; none seen (dd/mm forms with `hh:mm ss` stay text, ISO is unambiguous).
  - Pick Pack files are not de-duplicated by name (`PickPackReport_2026-04-01_to_2026-04-03` exists as both
    CSV and Sheet); harmless — the lookup keeps the latest Pick Date per Order No.
  - No alert when V/X fill rate drops; the projected table in the log is the check for now.
- **Related:** [2026-10-09-full-flow-rows-empty-suffix-and-refresh.md](2026-10-09-full-flow-rows-empty-suffix-and-refresh.md)
  (V/X 0 from May was visible in its fill-rate table; that fix added the refresh that will now carry V/X).

## Live results

- **Dry-run 38023719561 (commit af306e9) FAILED** after planning, in the new projected fill-rate table:
  `TypeError: Invalid value '2026-05-02 18:04:54' for dtype 'str'. Value should be a string or missing
  value, got 'datetime' instead.` (`_print_fill_rates` -> `existing_df.at[...] = value`). CI installs
  pandas 3.0.6, whose default string dtype refuses a datetime; local pandas 2.2.3 does not, so the test
  passed locally. Fix: project on `existing_df.astype(object)`; the test now uses a `"string"`-dtype frame
  and fails without the fix (`TypeError: Cannot set non-string value ... into a StringArray`).
  Everything before it was already right in that run:
  `Pick Pack files found: 128  (csv 10 | google sheet 118 | xlsx 0); other files ignored: 0`,
  `Pick Pack files read: 128`, `Total Pick Pack rows: 123,914`, `Valid Pick Dates: 123,914 / Invalid 0`,
  `Rows with Order No anchored to a Vinculum line: 5,541 — of which get blank cells refreshed: 4,373 —
  refresh cells: 8,746` (`refresh Actual.1 4,373`, `refresh Warehouse Team 4,373`), no other column
  written, sample `row 1695 Actual.1 = '02-May-2026 18:04:54'`, `Warehouse Team = 'VinayakB'`.
