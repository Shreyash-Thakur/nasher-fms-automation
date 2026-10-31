# 83 invoiced rows never got AW: P holds the wrong order of a split GT (73), and the 10.10 Sales Report is cut off (10)

- **Date:** 2026-10-31
- **Reported as:** "Im seeing so many gaps in the fms." Then: "whatever's mentioned in the screenshots try to fill as many as you can use the vin api any source you want".
- **Symptom:** in the 31-Oct export of "GT Full Flow FMS 2026-2027", 83 open rows have an Order No (P) but a blank Invoice No (AW) and Team (Q). Their orders are in the 10, 15 and 16-Oct daily Sales Report files, which the 23-Oct run read. Examples: row 8420 GT141026A1 O94134 and row 8213 GT081026A2T M892782. By GT:

  | GT | P on every row | rows | AS |
  |---|---|---|---|
  | GT141026A1 | O94134 | 70 | blank |
  | GT151026A5 | M892795 | 3 | blank |
  | GT081026A2T | M892782 | 8 | filled (4/6) |
  | GT081026A10 / GT091026A2 | N89678 / N89688 | 1 + 1 | filled |

- **Where:** `scripts/gt_full_flow_fms_yes_no_all_or_latest.py`:
  - `allocate_lines` (line 2464). The anchor step (line 2508) takes only a line whose Order No is in the row's P. Lines of any other order go to "unplaced (split shipments)".
  - The daily Sales Report file `10.10.2026` (Drive folder 1iL94fl…, created 11-Oct 10:18Z) is damaged. This is data, not this repo.
- **Root cause:** there are two separate causes.
  1. **Wrong order in P (73 rows: GT141026A1 70, GT151026A5 3).**
     - One Vinculum Order No was written into P of every row of the GT, but Vinculum split each GT into several orders.
     - The 15.10 and 16.10 files show GT141026A1 on O94134 (2 SKUs, 27/10/NM4/340), N012273 (6 SKUs, N10), M892791 (20 SKUs, NM1), N051641 (4 SKUs, NM7) and N89724 (10 SKUs, GJ1). GT151026A5's three Arc backpacks are on N012276 (N10, 27/10/N10/1792), not M892795.
     - Only the 2 rows whose SKU really is on O94134 (8475/8476) and the 3 Trail rows of M892795 were ever filled. The other lines were logged every night as "unplaced", for example `unplaced: GT141026A1 / NMH2409256VS -> N051641` in the 23-Oct run log.
     - This job did not write those P values. The runs of 15-Oct 17:22Z and 22:32Z and 16-Oct 16:58Z all log "filled from Order Enquiry (punched orders): 0". GT141026A1 is in neither the punch portal's ledger nor its Vinculum sync log. So the writer was most likely a manual entry.
  2. **Damaged Sales Report 10.10.2026 (10 of the 83 rows, plus 14 more rows with AS but no AW).**
     - 803 of its 869 rows end after column "Shipping IGST". Invoice No, Originating warehouse, Total and Created by are empty. The rows do carry an e-invoice IRN, so they are invoiced.
     - The job wrote AS from Pick Qty and could write nothing else.
     - 09.10, 15.10 and 16.10 are complete (0 rows without an invoice). Only 10.10 is affected. It is the last file made before the 11-Oct producer change.
     - M892782 / N89678 / N89688 are in no other Sales file, so no source the job reads has their invoice.
- **Why it wasn't caught:**
  - The run log printed "Extra Vinculum lines not placed (split shipments): 56" as information only.
  - Nothing checks that a daily Sales file has invoice numbers. A row with AS but no AW looks like "not invoiced yet".
- **Fix:**
  - `allocate_lines` step 4, `_other_order_slot` (line 2553). A leftover line goes to an existing row only when all of these hold:
    - the row is open, P is filled, and it has no anchor yet;
    - same GT and same Vinculum SKU (tier 1/2);
    - P's order(s) have at least one Sales line in this GT and none for this SKU;
    - the line's order is in no row's P of that GT (an order other rows claim keeps its lines).
  - The first such row wins, and each row takes one line. That row's blank Q/AS/AW/AX/AY (and V/X from Pick Pack by the line's order) are filled with reason `refresh_other_order`. P is never written.
  - New report `reports/order_no_mismatch.csv` (artifact "fullflow-reports") lists sheet_row, GT, SKU, P, the real Vinculum order and the invoice, so ops can correct P. The POD sync and the dashboards look up orders by P.
  - Anchor and fill take a copy of a line WITH an Invoice No before one without (line 2500). A re-exported 10.10 line will then win over the truncated copy already in the Sales Master.
  - `sales_rows_without_invoice` (line 1444). A daily Sales file with rows that have a blank Invoice No now prints a WARNING and a run-health warning ("re-export the file"). This does not fail the run.
- **Evidence:**
  - Sales Report 10.10.2026 (CSV export, 31-Oct): `_last: Shipping IGST 803, Created by 64`. GT lines: `26 blank inv 26 with IRN & blank inv 26`. These are M892782 (8 lines), N89678…N89694, N012269 and H01807, the same orders as the 24 "AS filled, AW blank" rows.
  - Files 09.10, 15.10 and 16.10: `blank inv 0`.
  - Offline dry run (`plan_backfill` on the 31-Oct export with the 09/10/15/16-Oct Sales files; no Sales Master, nothing written):
    - old code: `{}` (no writes), `unplaced 50`, `GT141026A1 rows 72 AW planned 0`.
    - new code: `('refresh_other_order', 'Invoice No') 46` (and the same for Team / Warehouse / Dispatched Qty / Invoice Amount), `GT141026A1 rows 72 AW planned 40`, `GT151026A5 rows 6 AW planned 3`, `unplaced 4`. The other 3 rows are GT151026A4 8525-8527 (P M892794 -> N012275, 27/10/N10/1793).
    - Sample: `(8420, 'Invoice No', '27/10/NM1/1273', 'refresh_other_order')`, `(8422, 'Invoice No', '27/10/GJ1/340', ...)`.
  - New tests `tests/test_other_order_lines.py` (11). Against the old script: `7 failed, 4 passed`. With the fix: the full suite gives `171 passed`.
- **Blast radius:**
  - Rows whose P names a different order of the same GT, where that order is invoiced without this SKU and the real order is on no row's P. They get their blank invoice-chain cells.
  - About 46 rows on the 15/16-Oct lines alone. Older GTs with the same pattern in the Sales Master will also be filled; all are listed in `order_no_mismatch.csv`.
  - No P, no filled cell, no closed row and no BD is written.
- **Still open:**
  - **Not merged to main** (branch `fill-gaps-31oct`).
  - The 15/16-Oct files are older than the default 7-day lookback. If their lines are not in the Sales Master (the `_ProcessedFiles` skip noted on 23-Oct), run "GT FMS Automation" once with `sales_lookback_days` = 17 after merging.
  - Fix P by hand from `order_no_mismatch.csv`.
  - Re-export Sales Report 10.10.2026 (nasher-miles-report-automation) and remove "10.10.2026" from the Sales Master `_ProcessedFiles` tab so it is re-imported. Then 10 + 14 rows (M892782, N89678…N89694, N012269, H01807) get AW from a run with `sales_lookback_days` = 22, or from the Sales Master.
  - The 83 count in the audit was found with Drive full-text search. "Order is in the file" did not mean "invoice is in the file".
- **Related:** `2026-10-23-portal-punched-rows-never-get-punch-date.md` (split P, `split_order_nos`), `2026-10-09-full-flow-rows-empty-suffix-and-refresh.md` (anchor/refresh), audit `gaps31/gt_audit.md`.
