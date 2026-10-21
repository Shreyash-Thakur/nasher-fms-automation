# Nightly Full Flow append deletes column BD ("Vinculum Customer Code") and everything right of BC

- **Date:** 2026-10-21
- **Reported as:** N/A. Found in a code review while building the Vinculum Customer Code feature. The brief asked to "check whether it is used and whether any write can misalign or clear BD".
- **Symptom:** none was reported yet. On every run that appends rows, `write_full_flow` sends `updateSheetProperties` with `gridProperties.columnCount = 55` (A:BC). If the grid is wider, that request deletes columns BD onward together with their data. BD is the "Vinculum Customer Code" column, which the auto-punch project (NM_abhi_viniculum_testing, `punch_run.js:446-452`) writes and reads.
- **Where:** `scripts/gt_full_flow_fms_yes_no_all_or_latest.py:847` (before this fix: `"columnCount": len(OUTPUT_COLUMNS)` in `write_full_flow`). Related sites: `read_existing_full_flow` (range `A8:BC`), and `clear_full_flow_data` (cleared `A8:BC`, never called).
- **Root cause:** the grid resize was written to make sure A:BC exist. It *set* the column count to the job's own width instead of raising it to at least that width. The Sheets API treats a smaller `columnCount` as "delete the extra columns". Rows had already been fixed to grow only (2026-10-12, GT Pull realtime). Columns were missed because nothing lived beyond BC until the auto-punch project started using BD.
- **Why it wasn't caught:** `test_grid_is_never_shrunk` checked only `rowCount`. No test covered `columnCount` or anything right of BC.
- **Fix:**
  - `grid_column_count(current, required)` returns `max(current, required)`, so columns only grow. The job always keeps at least 56 columns, so BD exists.
  - The grid is resized *before* BD7 is read or written, because an A1 range outside the grid is refused.
  - BD7 handling:
    - If BD7 holds "Vinculum Customer Code", the job writes A:BD.
    - If BD7 is blank, the job sets that header and writes A:BD.
    - If BD7 holds anything else, the job writes A:BC only and prints a WARNING.
  - Existing rows never get a BD write (a filter in `plan_existing_writes`).
  - The reads are clamped to the real grid (`fms_read_last_column`, `full_flow_read_last_column`), so a narrow tab cannot make a range read fail.
  - The unused `clear_full_flow_data` was removed. It would have wiped A:BC but left BD behind, out of line with its rows.
- **Evidence:** `python -m pytest tests -q` gave `151 passed in 0.90s`. These tests in `tests/test_customer_code.py` cover the fix:
  - `test_column_count_never_shrinks[60-…-60 | 55-''-56 | 57-Remarks-57]`
  - `test_grid_expanded_before_bd7_is_read` (order: grid, then BD7 get; columnCount 56; write A8:BD8)
  - `test_read_ranges_never_pass_the_grid`
  - `test_existing_rows_never_get_a_bd_write`

  Running `python -m pytest tests/test_customer_code.py tests/test_realtime_rows.py -q` gave `30 passed`.
- **Blast radius:** column BD and anything right of BC in "Full Flow FMS", which only the auto-punch project uses. GT Pull `Realtime.js` writes BD by header in the same way (GT_Pull commit of 2026-10-21). The POD and planned-status scripts write single cells only, so they are not affected.
- **Still open:** the GT Pull realtime append and the nightly append are still not locked against each other. That race existed before this change, and BD has it like every other column. It is not addressed here.
- **Related:** `2026-10-12-live-vinculum-gt-orders-missing-from-full-flow.md` (the rows-never-shrink fix).
