# GT Full Flow FMS: columns after AA empty for orders from August onwards

- **Date:** 2026-10-05
- **Reported as:** "why are columns after AA empty post july os something broken refer to new gt fms and compare and give"
- **Symptom:** filled-cell counts per form month in `Full Flow FMS` (sheet `1jMEEFk…`, header row 7):
  Order No (P) Apr 1389/2181 · May 999/1351 · Jun 1129/1459 · Jul 855/1162 · **Aug 363/590 · Sep 40/1315 · Oct 0/25**;
  Actual (S), Dispatched Qty (AS), Warehouse (AX), Invoice Amount (AY) identical to P; POD (AU) / Delivered Date (AV) /
  Order Delivery Status (AJ): Sep 39, Oct 0. The New GT FMS (`Form_FMS`, `1TWMxOk6…`) has Order No for 794/1365 Sep rows,
  and the Order Enquiry source has full `Sep-26` / `Sept-26` / `Oct-26` tabs — the data exists, Full Flow never picks it up.
- **Where:**
  - `scripts/gt_full_flow_fms_yes_no_all_or_latest.py` `main()` — "INCREMENTAL PROCESSING … Only new GTs are processed"
    (`filter_unprocessed_fms`, `get_processed_gt_numbers`); `build_full_flow` matches GT Number + Vinculum SKU against
    Order Enquiry ExtOrderNo + SKU Code only for rows being appended.
  - Run log 2026-10-04 (run 37239885448): `FMS rows already processed: 7,974 · FMS rows to process now: 13 · FMS rows with Sales match: 0`.
  - Downstream: `gt_full_flow_fms_pod_update_google_sheet_gt_and_order_only.py` `find_pod` requires BOTH GT and Order No.
- **Root cause:** a GT is appended to Full Flow the first night after its form arrives — usually before Ops have punched
  the Vinculum order — so it is written with a blank Order No, and because processing is append-only it is never
  revisited; with no Order No the POD sync (GT + Order No) can't match it either, so every column after it stays empty.
  Before 2026-09-08 Diwakar rebuilt Full Flow manually, which re-matched rows; since the job moved to the nightly
  GitHub schedule (2026-09-08) only the append path runs.
- **Why it wasn't caught:** the job reports success every night (it does what it was written to do); no check counts
  rows with a blank Order No; the summary only reports matches for newly appended rows.
- **Fix:** new nightly step `run_backfill()` (before the "no new GTs" early exit): for existing rows with a blank Order No
  that aren't Cancelled, the same GT + Vinculum SKU match (`build_sales_lookup`, shared with `build_full_flow`) fills
  **only blank cells** of P Order No, Q Team, S Actual, AS Dispatched Qty, AW Invoice No, AX Warehouse, AY Invoice
  Amount, plus Pick-Pack X Actual.1 / V Warehouse Team. Order Enquiry lines already used by a filled row of the same
  GT+SKU are not reused; extra lines (split shipments) are reported, not inserted. Controlled by `BACKFILL_MODE`
  (`dry-run` default / `write` / `off`; workflow_dispatch input `backfill_mode`). The Order Enquiry + Pick-Pack
  lookup code was moved verbatim into `build_sales_lookup` / `build_pick_pack_lookup` so both paths share it.
  The POD step that runs right after then matches the newly filled rows on its own.
- **Evidence:**
  - `python tests/test_backfill.py` → `11 passed, 0 failed` (blank-only writes, never overwrite, cancelled skipped,
    used lines not reused, extras reported, Pick-Pack bridge, row/A1 mapping).
  - Refactor parity: old (HEAD) vs new `build_full_flow` on the same inputs (incl. one GT expanding to 2 rows and an
    unparseable date) → `PARITY OK: old and new build_full_flow identical on 4 output rows`.
  - Live dry-run (run 37291985617, 2026-10-05): `Rows with blank Order No considered: 1,404 · would be filled: 743 ·
    cells: 5,201 · still without a Vinculum match: 661 · extra Vinculum lines not placed: 1 (GT230926A3||NMHPP80069QM)`.
    Spot-check: row 6887 = GT050926A1 / NMH81133GS → would get `N051601`, Team `Process6`, Warehouse `GGN 1`, Dispatched 8
    of 10, Invoice Amount 14,136 (= 8 × 1,767); New GT FMS independently shows `N051601 / Process6` for all three lines.
  - Write enabled 2026-10-05 after review (workflow default `write`). First write run 37294051229: `Backfill complete:
    5,201 cells written.` (743 rows); POD step in the same run: `POD matched : 3,015 · Total cell updates: 21,971`.
  - Sheet re-counted after the run (gviz count per form month): Order No Sep **40 → 775** / 1,315, Oct 0 → 8 / 92;
    POD Copy / Delivered Date / Order Delivery Status Sep **39 → 256**. Aug unchanged at 363 / 590 — its 227 blank rows
    are 123 `Cancelled` (skipped by design) + 104 with no Vinculum line. Cross-check: New GT FMS has Order No for
    794 / 1,365 Sep rows, Full Flow now 775 / 1,315 — the remaining blanks are orders not (yet) in Vinculum.
- **Blast radius:** writes only blank cells in existing Full Flow rows (P, Q, S, AS, AW, AX, AY, X, V). Consumers:
  POD sync (will now match these rows), planned-status script, FMS_Dashboard (`dashboard.js`), GT Fulfilment web app
  (`Pod_*`). The append path for new GTs is unchanged (parity-tested).
- **Still open:**
  - GT+SKU keys with more Vinculum lines than Full Flow rows (split shipments) are reported, not expanded into new rows.
  - Stage columns that have no source at all (Full Flow X/AA/AB/AF/AG/AK/AL/AO only filled for April) are not addressed.
- **Related:** `GT_FMS/Errors and their fixes/2026-10-01-pending-pod-always-zero.md` (same sheet; POD rule).
