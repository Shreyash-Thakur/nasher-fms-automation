# Full Flow details arrive 2-3 days late, and an unreadable file leaves the run green

- **Date:** 2026-10-13
- **Reported as:** "check if all the files are being fetched correctly and if it's working fine end to end — all the details should be updated at least within 24 hours."
- **Symptom:** In the 13-Oct nightly run (38332436931, started 07:00 IST) the newest Pick Pack,
  Order Enquiry and Sales Report files were all `10.10.2026` (3 days old) — every Vinculum export
  job has failed since 12-Oct ("Login failed — check VIN_LOGIN_ID / VIN_PASSWORD secrets.") and the
  Full Flow run was still green with no warning. On 11-Oct the run logged
  `WARNING: Pick Pack file not readable, skipped: PickPackReport_2026-06-09_to_2026-06-11 [gsheet] -> HttpError: <HttpError 403 ... returned "User rate limit exceeded." ... 'reason': 'userRateLimitExceeded'>`
  and also stayed green. Invoices (Q/AS/AW–AY) of day D reached Full Flow only on D+3 morning.
- **Where:**
  - `scripts/gt_full_flow_fms_yes_no_all_or_latest.py` `export_drive_file` / `download_drive_file`
    (no retry), `read_pick_pack_files` / `read_order_enquiry_daily_files` (skip = print only),
    `read_sales_file` (Sales Master only), `main()` / `__main__` (exit 0 regardless).
  - `.github/workflows/run_fms.yml` (one run a day, `cron: '30 19 * * *'`, starts 3–6 h late).
  - `requirements.txt` (unpinned: CI silently moved to pandas 3.0.6).
- **Root cause:** Three independent delays/blind spots stacked up:
  1. Invoice path has an extra hop: the Vinculum Sales Report for invoice day D is uploaded
     (folder `1iL94fl…`, file `DD.MM.YYYY`) on D+1 ~15:30–17:00 IST, Apps Script
     `VinculumSalesReportConsolidate` copies it into the Sales Master on its next daily trigger
     (D+2), and Full Flow — which read only the Sales Master — picks it up on the following nightly
     run (D+3 04:00–07:00 IST).
  2. Full Flow ran once a day and GitHub starts scheduled runs 3–6 h late, so anything landing in
     the afternoon waited ~12–15 h for the next run.
  3. Drive exports of 130 Pick Pack Sheets in a row hit the per-user rate limit; the reader caught
     every exception as "skip this file" with no retry and no effect on the exit code, and nothing
     checked whether the newest file of each source was recent.
- **Why it wasn't caught:** No freshness check and no non-zero exit for a skipped file; the
  only signal was one `WARNING:` line among ~1,000 log lines. The upstream Vinculum jobs fail in a
  different repo, so the Full Flow run looked healthy while reading 3-day-old data.
- **Fix:**
  - Read the daily Sales Report files of the last 7 days directly (`read_sales_daily_files`, via
    the Sheets API like the Sales Master tabs) and use the lines the Sales Master does not hold yet
    (same Order No + GT + SKU + Invoice No; kept out of the cross-tab de-dup so a second invoice of
    the same line is not dropped). Removes the consolidation day: D's invoices now land with the
    first Full Flow run after the file appears (same evening).
  - `with_drive_backoff`: retry 403 rate-limit / 429 / 5xx with 2-4-8-16-32 s backoff (a 403 for a
    missing permission is not retried); `list` calls use `execute(num_retries=5)`.
  - `RunHealth`: per-source found / read / skipped / rows / newest file printed, written to the
    GitHub job summary and annotated; the process exits 1 (after all writes) if a file dated within
    the last 3 days could not be read; a daily source whose newest file is > 2 days old is a
    `::warning::`.
  - `run_fms.yml`: second run at 12:47 UTC (18:17 IST) after the Vinculum files land, shared
    `concurrency: fms-full-flow` so runs never overlap, `timeout-minutes: 45`, pytest pinned.
  - `requirements.txt` pinned to the versions of the last green run (pandas 3.0.6 etc.).
- **Evidence:** `python -m pytest tests -q` → `128 passed` (100 before + 28 in
  `tests/test_freshness.py`), also under the pinned versions in a clean venv
  (pandas 3.0.6, `pip check`: No broken requirements found). Live dry-run 38362883163 (writes
  nothing; POD/planned skipped), green:
  `Sales Report daily files in window: 5 (folder has 96 files)` (06–10.10.2026 read) →
  `Sales Report daily lines: 3,558 read, 2,772 already in the Sales Master, 786 used (not consolidated yet)`
  → `matched now (would be / are filled): 10`, "Order No not in the Sales Master" 144 → 130; job
  summary table written; 3 `::warning::` lines "newest file is 10.10.2026 (3 days old)" for Pick
  Pack / Order Enquiry / Sales daily. So the Sales Master was missing ~786 invoice lines of files
  that had been in Drive for 2+ days — the consolidation lags by more than the expected day.
- **Blast radius:** `prepare_sales_lines` output now also holds not-yet-consolidated invoice lines
  → fills of P/Q/S/AS/AW/AX/AY, `OrderEnquiryIndex` (an order with a daily invoice line counts as
  invoiced), the unmatched / "not in Full Flow" reports. Exit code 1 now marks the Full Flow step
  red; POD and planned-status steps still run (`!cancelled()`). The evening run also runs POD and
  planned-status a second time a day (idempotent; ~62k planned-status cells, 13 API calls).
- **Still open:**
  - The root of today's staleness is upstream: Vinculum login has failed for every report job
    since 12-Oct ~16:20 IST (screenshot: login page, user `Diwakar`, password cleared). Needs a
    valid Vinculum user in the `VIN_LOGIN_ID` / `VIN_PASSWORD` secrets of
    nasher-miles-report-automation — see that repo's 2026-10-13 record.
  - Vinculum reports are exported for "yesterday" only, so an event on day D is in Full Flow on
    D+1 evening at the earliest (~20–34 h depending on the time of day); ≤24 h for every event
    needs intraday ("today so far") exports, not done (partial files would also be consolidated
    by the Apps Script).
  - The service account of this repo must be able to read folder `1iL94fl…`; if not, the job
    summary says "folder empty / not shared" and invoices fall back to the Sales Master path.
- **Related:** `2026-10-10-pick-pack-google-sheets-skipped.md`,
  `2026-10-10-order-no-blank-until-invoiced.md`,
  `2026-10-12-live-vinculum-gt-orders-missing-from-full-flow.md`.
