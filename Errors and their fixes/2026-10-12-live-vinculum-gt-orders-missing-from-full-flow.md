# Live Vinculum GT orders (not cancelled) missing from Full Flow FMS — nobody could see them

- **Date:** 2026-10-12
- **Reported as:** "There are 88 GT orders missing from the GT Full FMS, and their order status is not cancelled. Kindly check and update." (list not attached)
- **Symptom:** Reproduced from the daily Order Enquiry exports, 07-09 to 10-10 (no daily file exists for
  1–6 Sep). Of 325 live (not Cancelled) Vinculum GT orders, **84 Order Nos appear nowhere in Full Flow FMS**
  (`1jMEEFk…`):
  - `gt_missing` 42 orders / 37 GT refs. The GT never reached Full Flow: GT180926A5–A31, GT190926A1–A5,
    GT290926A2/A3/A4R/A4T, and GT080926A2_C.
  - `order_not_shown` 42 orders / 31 GTs. The GT is in Full Flow, but this order is not, mostly one
    `N4617xx` "Confirmed" order per GT (for example GT230926A3 → N461748, GT101026A2 → M892783, N012270,
    N461767, N89705).
- **Where:**
  - `scripts/gt_full_flow_fms_yes_no_all_or_latest.py` `find_new_fms_rows` (`:2104`). Full Flow only ever
    receives what is in GT Pull FMS (`read_fms_sheet`, `:343`). GTs missing upstream (`gt_missing`) cannot
    appear. The cause is in GT_Pull: lines were removed from GT Master `Orders` before GT Pull copied them.
    See the GT_Pull record.
  - `OrderEnquiryIndex.match` (`:1327`, `if len(candidates) > 1: return None, "ambiguous"`) and the
    line-level Sales Master fill. P is one Order No per SKU line. Order Enquiry has no SKU, so when a GT
    is split over several live Vinculum orders, the extra orders get no P until they are invoiced and
    matched by SKU (`order_not_shown`).
  - No report anywhere listed Vinculum GT orders that Full Flow lacks. `vinculum_lines_not_in_full_flow.csv`
    is all channels, invoiced lines only: 31,823 rows, of which only 69 are GT lines.
- **Root cause:** The pipeline is one-directional (Orders → GT Pull FMS → Full Flow), and nothing compares
  Vinculum's punched GT orders back against Full Flow. GTs whose Orders lines vanished before GT Pull ran,
  and split orders that P cannot hold, were therefore invisible until Ops checked by hand.
- **Why it wasn't caught:** The only Vinculum-side reports were `suspect_cancellations.csv` and
  `unmatched_rows.csv`. Both start from Full Flow rows, so a GT that has no Full Flow row is never looked at.
- **Fix:**
  - New READ-ONLY `report_vinculum_gt_orders_not_in_full_flow()`, called in `main()` after
    `report_suspect_cancellations`. Its universe is daily Order Enquiry orders with a GT ExtOrderNo that are
    not Cancelled, plus Sales Master GT orders, minus any order Order Enquiry calls Cancelled.
  - GT presence is checked on the exact GT, else the R/T/F base (equality only, so A1 never matches A10).
  - Each order gets one reason: `gt_missing` (in neither Full Flow nor GT Pull), `appending_tonight` (in
    GT Pull only) or `order_not_shown` (GT present, Order No in no P). The report prints counts and the
    missing GTs with their orders, and writes `reports/vinculum_gt_orders_not_in_full_flow.csv` to the
    `fullflow-reports` artifact.
  - It writes nothing to any sheet. The recovery for the 36 live `gt_missing` GTs is the GT_Pull
    backfill (`Backfill_DeletedGTs.js`).
- **Evidence:**
  - `python -m pytest tests -q` → `93 passed`, including 6 new tests in
    `tests/test_vinculum_missing_report.py`: reasons, A1≠A10, suffix-tolerant presence, Sales-Master
    orders unless cancelled, writes CSV and leaves Full Flow untouched, no Order Enquiry files.
  - Offline run of the new function on the real 2026-10-12 exports (30 daily Order Enquiry files, Full Flow,
    GT Pull FMS; Sales Master not available locally):
    `Live Vinculum GT orders: 325` / `gt_missing orders 42 GTs 37` / `appending_tonight orders 0` /
    `order_not_shown orders 42 GTs 31`, followed by the per-GT list (`GT180926A20: N89578 (Allocated)` …).
  - `find_new_fms_rows` on GT Pull FMS plus the 45 rows the GT_Pull backfill plans:
    `after backfill: new-GT rows 45, GTs 36, new lines for existing GTs 0`.
- **Blast radius:** read-only. It adds one CSV and a log section to the nightly run. No sheet writes and
  no change to matching, filling or appending.
- **Still open:**
  - `order_not_shown` (42 orders) is a model limitation: a SKU line has one P, and Order Enquiry is
    order-level. Showing every split order needs either an SKU-level Order Enquiry export or a separate
    "other Vinculum orders" column. That is an Ops/design decision and is not done here. Most are `N4617xx`
    orders that have sat in "Confirmed" since early Sep. Ops should check whether that location ever
    fulfils.
  - The 92 GTs whose every Full Flow row reads "Cancelled" (historical auto-cancel, P blank; excluded from
    the dashboard's "643 Active GTs") include about 24 with non-cancelled Vinculum evidence. Commit
    `2b8f439` ("Order No reads Cancelled on cancelled rows") will write "Cancelled" into P for those 823
    rows on the next write run. Revert the fabricated stamps first, or exclude them.
- **Related:** `GT_Pull/Errors and their fixes/2026-10-12-live-gt-orders-missing-from-full-flow.md`,
  `2026-10-10-order-no-blank-until-invoiced.md`, `2026-10-09-full-flow-rows-empty-suffix-and-refresh.md`.
