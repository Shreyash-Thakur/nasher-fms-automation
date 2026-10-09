# GT Full Flow FMS: "a lot of rows are empty" — suffix mismatch fabricated Cancelled, rest are data gaps

- **Date:** 2026-10-09
- **Reported as:** "it's not fetching data correctly — a lot of rows are empty"
- **Symptom:** nightly log (run 37861578756, 2026-10-08): `FMS rows with Sales match: 0` for new rows,
  `Rows with blank Order No considered: 758 · still without a Vinculum match: 744 · 98 cells written`,
  `Historical rows marked Cancelled: 1,908`. Live fill-rate table (dry-run 37930816537, open rows only,
  filled / rows):

  | form month | rows | P Order No | V/X pick | AJ/AU/AV POD | AS/AW/AX/AY |
  |---|---|---|---|---|---|
  | 2026-04 | 1385 | 1385 | 1148 | 648 | 1385 |
  | 2026-05 | 999 | 999 | **0** | 626 | 999 |
  | 2026-06 | 1129 | 1129 | **0** | 655 | 1129 |
  | 2026-07 | 855 | 855 | **0** | 675 | 855 |
  | 2026-08 | 467 | **363** | **0** | 176 | 363 |
  | 2026-09 | 1315 | **775** | **0** | 378 | 775 |
  | 2026-10 | 131 | 26 | **0** | 4 | 26 |

  plus **1,908 closed rows** (1,905 read "Cancelled" across T:AW).
- **Where:**
  - `scripts/gt_full_flow_fms_yes_no_all_or_latest.py:730` (before this fix) `create_match_key` — exact
    `normalize_value(GT) + "||" + SKU`; no R/T/F approval-suffix handling (every v2 script stripped it:
    `GT_FMS_v2_DEPRECATED/orderidandteam.js normalizeOrderId`; pod.py `normalize_gt` does too).
  - same file `:58` `ENABLE_HISTORICAL_CANCELLATION = "YES"` + `:1117` `cancel_mask = eligible_date & missing_sales`
    — "no exact Vinculum key" ⇒ stamp "Cancelled" into T:AW for forms up to 2026-08-15.
  - same file `:1412` `_is_cancelled_row` + `:1437` backfill skips those rows forever.
  - same file `:554` `filter_unprocessed_fms` — append-only per GT (a new SKU line of a known GT never lands).
  - same file `:1437` — only rows with a BLANK Order No were ever revisited (no refresh of later invoice/pick).
  - `scripts/gt_full_flow_fms_pod_update_google_sheet_gt_and_order_only.py:1057` `row_is_cancelled` looked at
    C / P / AP only, so a punched line marked "Virtual Invoice" (GT_FMS `markVirtualGT_` writes AJ only) was
    overwritten with "Delivered" when a POD matched.
  - `.github/workflows/run_fms.yml` — no `if:` on later steps: a Full Flow failure skipped POD + planned status.
- **Root cause:** FMS "GT Number" is the order's GT Final, which carries an R/T/F approval suffix on mixed orders
  (GeneralTrade_FMS `Order_Entry.js` `needsSuffix`), while Ops punch Vinculum ExtOrderNo with or without it; the
  match compared the two exactly, so lines like `GT140426A10R` (FMS) vs `GT140426A10` (Vinculum) never matched,
  the historical-cancellation step read "no match" as "cancelled" and stamped them, and every later fill skipped
  stamped rows. Live evidence: **206 rows / 13 GTs** read "Cancelled" yet have a Vinculum line under the
  suffix-tolerant rule — **all 206 invoiced**; 197 have the suffix on the FMS side, 9 on the Vinculum side
  (`GT180526A3` ↔ `GT180526A3T`), 0 match exactly. Biggest: GT140426A10R 75, GT140426A10T 52, GT150626A2T 33,
  GT150626A2R 22, GT090626A2R 13.
  The other empties are **not** a fetch bug (measured on live data, dry-run 37930420863):
  - **724 open rows with blank Order No**: 372 rows / 45 GTs whose GT is in no Vinculum ExtOrderNo at all (no
    ExtOrderNo even contains it; GT070926A5 105 rows, GT030926A1 88, GT220926A2T 63, GT260826A1 44 …) — not
    punched; 352 rows whose GT is in Vinculum but that SKU line is not — for 276 of them every Vinculum line of
    the GT is already on another row (partly punched orders), 76 have leftover lines under a different SKU
    (substitutions, or a set such as `NMH12284DSML` punched as `NMH12284DS/DM/DL`) — not safely auto-matchable.
    Suffix-tolerant matching rescues **0** of these today (it would have rescued the 206 above).
  - **V/X (Warehouse Team / pick date) empty from May**: the Pick Pack Drive folder (`1Xi66o4A…`) only holds
    ten `PickPackReport_2026-04-*.csv` files — no source after April.
  - **AJ/AU/AV (POD)**: POD files only exist for part of the orders; pod.py needs GT + Order No.
  - Rows with an Order No: **5,516 of 5,532** anchor to their Vinculum line and have **0** blank P/Q/S/AS/AW/AX/AY
    cells to refresh (the "refresh" gap was real in code but empty in data today).
- **Why it wasn't caught:** the job "succeeds" every night; no count of rows per match tier, of Cancelled rows
  that actually have a Vinculum order, or of per-column fill rate; tests only covered exact keys.
- **Fix:**
  - `gt_base()` strips exactly ONE trailing R/T/F after a digit (`GT081026A1R → GT081026A1`; `A10` never becomes
    `A1`; equality only, no prefix match). `allocate_lines()` matches every existing row AND every row being
    appended against Vinculum lines of the same GT base, tier by tier: 1 exact GT + Vinculum SKU, 2 R/T/F-stripped
    GT + Vinculum SKU (either side), 3 same GT base + alternate SKU field (FMS SKU Code / Vinculum Bundle SKU
    Code). A line is used once; rows that already have an Order No keep their own line first.
  - Refresh: rows whose Order No is a Vinculum line get BLANK cells of P/Q/S/AS/AW/AX/AY/X/V filled nightly;
    nothing non-blank is overwritten; closed rows (Cancelled / Virtual Invoice / OOS / Out of stock in J or
    P:AW) are never written. Pick Qty 0 is not written (would block the real quantity later).
  - New lines: `find_new_fms_rows()` appends any FMS GT + Vinculum SKU not yet in Full Flow (new GTs as before
    + new SKU lines of known GTs; a line present in any state, incl. Cancelled, is never re-added; >300 such
    lines in one night is refused as structural). Idempotent (test).
  - Order Enquiry lines duplicated across month tabs (`Sep-26` / `Sept-26`) are dropped (794 live) so they are
    not mistaken for split shipments.
  - `ENABLE_HISTORICAL_CANCELLATION = "NO"`. Existing stamps are NOT reverted; `report_suspect_cancellations()`
    lists them every run (log + CSV in workflow artifact `fullflow-reports/suspect_cancellations.csv`), plus
    `unmatched_rows.csv` and a fill-rate table.
  - pod.py: closed = marker in C or anywhere in P:AV (incl. AJ "Virtual Invoice"); closed rows get no write at all.
  - Workflow: `pytest` step (Full Flow step is skipped if tests fail); POD + planned status run
    `if: !cancelled()` so an earlier failure doesn't skip them; `backfill_mode` dry-run now means the Full Flow
    script writes nothing (no fills, no appends) and POD/planned-status are skipped.
- **Evidence:**
  - `python -m pytest tests -q` → `30 passed` (test_backfill 11, test_sync 15, test_pod 4). test_pod fails 3/4
    against the previous pod.py (`git show HEAD~:…pod…py`), passes after.
  - Branch dry-runs on live data (37929753099, 37930121174, 37930420863, 37930816537):
    `Rows with blank Order No considered: 749 · matched now: 25 (tier1 exact 25 | tier2 R/T/F suffix 0 | tier3 alt
    SKU 0) · cells: 175 · still without a Vinculum match: 724 · anchored 5,516 · refresh cells 0 · Extra Vinculum
    lines not placed: 1 (GT230926A3 / NMHPP80069QM -> O94127)`; `Rows reading Cancelled: 1,905 · ... have a
    Vinculum match: 206 · historical auto-cancel 206 · invoiced 206`; `New SKU lines for GTs already in Full
    Flow: 0`; `Rows of GTs not in Full Flow (new GTs): 47`.
  - Main-branch dry-run / write results: see "Results on main" below.
- **Blast radius:** writes only blank cells of existing open rows (P, Q, S, AS, AW, AX, AY, X, V) and appends rows
  at the bottom. Readers: pod.py (POD), planned-status script, GT_FMS web app / `dashboard.js`, FMS_Dashboard.
  pod.py now leaves closed rows completely alone (previously AP/AU/AV could still be filled on a virtual line).
- **Still open:**
  - **206 fabricated "Cancelled" rows** (13 GTs, all invoiced) — revert needs the user's OK: clear "Cancelled"
    from T:AW on the rows in `suspect_cancellations.csv`; the next nightly run then fills P..AY and POD.
    The remaining ~1,699 historical "Cancelled" rows have no Vinculum line under any rule.
  - 724 blank rows: Ops to punch / cancel / mark OOS (GT_FMS dashboard "Not punched yet" list); 76 have
    substitute or exploded-set SKUs in Vinculum — decide whether set codes (`…SML`) should map to S/M/L lines.
  - Pick Pack reports stopped after April 2026 → V/X empty from May. Someone must upload PickPackReport CSVs.
  - Vinculum ExtOrderNo variants not matched on purpose: `GT110826A1V` (41 lines), `GT080926A2_C` (1). Ask Ops what
    `V` / `_C` mean before matching them.
  - Vinculum GT orders whose GT never reached FMS (GT180926A20–A31, GT190926A1–A5, GT290926A3, GT290926A4R/T,
    GT180926A5) — GT_Pull / Order_Entry side, not this repo.
  - GT_FMS `Pod_Upload.js:110` (manual POD upload) should get the same Virtual Invoice protection; `dashboard.js`
    `FF_EXCL_STATUS` still lacks 'virtual invoice' (known gap in GT_FMS HANDOVER.md).
  - pod.py `find_pod` still needs an Order No (GT-only matching deliberately disabled); `gt_index` stays unused.
- **Related:** `2026-10-05-full-flow-blank-order-no-never-backfilled.md` (the blank-P backfill this extends);
  `GT_FMS_v2_DEPRECATED/Errors and their fixes/2026-09-23-invoice-sync-fabricates-cancellations.md` (same
  "no match ⇒ Cancelled" failure in invoice2.js).

## Results on main

Fix commit `0ba2a32`.

| run | mode | blank-P rows considered | matched (tier1/2/3) | fill cells | refresh cells | still unmatched | appended |
|---|---|---|---|---|---|---|---|
| 37861578756 (2026-10-08 nightly, old code) | write | 758 | 14 | 98 | — | 744 | 5 |
| 37931236509 | dry-run | 749 | 25 (25/0/0) | 175 | 0 | 724 | would 47 |
| 37931523767 | write | 749 | 25 (25/0/0) | **175 written** | 0 | 724 | **47 (rows 8,197–8,243)** |
| 37931839960 (idempotency) | dry-run | 771 | 0 | 0 | 0 | 771 | 0 |

- Write run 37931523767: every step succeeded (tests → Full Flow → artifact → POD → planned status).
  POD: `Rows checked 8,236 · matched 3,162 · Closed rows left alone 0 · Total cells updated 0`;
  planned status `21,386` cells (its normal nightly rewrite).
- Oct form-month Order No 26 → 51; no other month changed (their blanks are not in Vinculum).
- Suspect-cancellation report unchanged: 1,905 Cancelled rows, 206 with an invoiced Vinculum order
  (artifact `fullflow-reports` of any run, e.g. 37931523767).
