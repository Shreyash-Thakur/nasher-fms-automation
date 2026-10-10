# GT Full Flow FMS: Order No (P), Team (Q), Order Date (S) stay blank until the order is invoiced

- **Date:** 2026-10-10
- **Reported as:** "in the GT Full Flow FMS, Order No (P), Team (Q) and Order Date/stage-1 Actual (S) stay blank until an order is INVOICED, because the only source is the "Vinculum Sales Master" (Sales Report, by invoice date). Ops want Order No etc. as soon as the order is punched in Vinculum."
- **Symptom:** dry-run 38034854973 (2026-10-10, before this fix): `Rows with blank Order No considered: 771`,
  `matched now: 0`, `Still unmatched, GT base NOT in Vinculum (not punched yet): 419`,
  `GT base in Vinculum but no SKU match: 352`; open-row fill rate P / Q / S:
  2026-08 363 of 467 · 2026-09 775 of 1315 · **2026-10 51 of 178**.
  Yet those GTs ARE punched: e.g. `OrderEnquiry_06-10-2026` has `NM11832, GT061026A4, 06/10/2026 06:27 PM,
  Shipped complete` while Full Flow row 8185 (`GT061026A4`) is blank with reason `gt_not_in_vinculum`;
  rows 8075-8080 (`GT031026A1…A6`) are blank while `OrderEnquiry_05-10-2026` / `06-10-2026` hold
  N89650… / N89655… for them. Offline count over the 11 daily files 29-09…09-10: of the 127 blank-P
  October rows, 46 `gt_not_in_vinculum` rows have exactly one Vinculum order in Order Enquiry.
- **Where** (line numbers before this fix):
  - `scripts/gt_full_flow_fms_yes_no_all_or_latest.py:917` `read_sales_file()` — the ONLY Vinculum source:
    spreadsheet `1Jnt0yU-…` (`ORDER_ENQUIRY_SPREADSHEET_ID`, `:47`), which despite the name is the Sales
    Master fed by Apps Script `VinculumSalesReportConsolidate` from the Vinculum Sales Report (invoiced
    lines; its columns include Invoice No / Invoice Date / Invoiced).
  - same file `:1471-1474` `BACKFILL_FIELDS` and `:1348-1350` `build_full_flow` — P/Q/S are copied only from
    a matched line of that frame.
  - same file `:1697-1700` `plan_existing_writes` — a blank-P row with no Sales line is counted as
    `still_unmatched` and skipped; there was no other source to try.
- **Root cause:** Full Flow's Vinculum data came from a single invoice-driven feed (Sales Report), so a
  line exists only after invoicing; an order that is punched (Confirmed / Allocated / Packed / …) but not yet
  invoiced has no line, so allocation finds nothing and P/Q/S (and the V/X that hang off P) stay blank until
  the invoice lands — days later for GT orders, and never if the order is short-closed.
- **Why it wasn't caught:** the constant and the log both call the Sales Master "Order Enquiry", so the log
  line "GT base NOT in Vinculum (not punched yet)" read as "not punched" when it meant "not invoiced"; no
  check compared Full Flow against the punched-order list.
- **Fix:**
  - New reader `read_order_enquiry_daily_files()` for Drive folder `1rinwTyd3BseS7o_FHP4Y907egQz39zBO`: every
    file named `OrderEnquiry_*` (case-insensitive), Google Sheet (xlsx export) / CSV / xlsx, same code path as
    Pick Pack (`_read_workbook_tabs`, now shared), header aliases via `normalize_sales_header`
    (`Extern OrderNo` -> `ExtOrderNo`). The monthly dumps in the folder (`Jan 2026`, `feb-march 2026`,
    `Order Enquiry April 2026`, `may-june 2026`, `July -August 2026`) are NOT read (different, line-level
    format; one-off snapshot of 2026-09-07). A bad file is skipped with a `WARNING`; nothing readable = no
    OE fill that night, never a failed run.
  - `prepare_oe_orders()`: one row per Vinculum Order No; an order in several files keeps the most
    recently modified file's copy (Drive listing is modifiedTime desc); `Status` containing "cancel" marks it
    cancelled; Order Date `06/10/2026 11:58 PM` parsed day-first 12-hour by `parse_oe_date` (ISO / serial
    fall back to `parse_pick_date`).
  - **Matching rule** (`OrderEnquiryIndex`) — the file is ORDER-LEVEL (39 columns, no SKU, no Created By;
    verified on `OrderEnquiry_06-10-2026` and the xlsx export of `OrderEnquiry_01-10-2026`), and a GT is
    often split into several Vinculum orders (GT041026A2 -> 4, GT061026A6 -> 3), so a GT+SKU match is
    impossible. A blank-P open row that the Sales Master did not match gets an Order Enquiry order only when:
    1. its candidate set = all LIVE Vinculum Order Nos known for the GT — Order Enquiry (not Cancelled) ∪
       Sales Master ∪ Order Nos already in Full Flow P for that GT, minus every order Order Enquiry marks
       Cancelled — taken on the exact normalised GT first, else on the GT base (one trailing R/T/F stripped,
       equality only, so GT…A1 never meets GT…A10);
    2. has exactly ONE member, present in Order Enquiry and NOT invoiced (no Sales Master line). An invoiced
       order belongs to the line-level Sales Master: if that did not match this SKU, an order-level file
       cannot overrule it.
    Then P <- Order No and S <- Order Date (blank cells only), plus V/X from Pick Pack keyed by that Order No.
    Every row of that GT gets the same Order No (one order holds all its lines). Q is written only if a
    future export carries "Created By"; today Q arrives with the invoice: the next night the row's P anchors
    to the Sales Master line and the existing refresh fills Q/AW/AX/AY/AS (never overwriting S).
  - Sales Master stays first and stays the only source for AS/AW/AX/AY: Order Enquiry is consulted only for
    rows the Sales Master left without a line. Closed rows (Cancelled / Virtual Invoice / OOS) untouched.
  - Same rule on the append path (`build_full_flow(..., oe_records)`), so a new FMS row punched yesterday
    arrives with P/S.
  - Log: `filled from Sales Master` vs `filled from Order Enquiry` (rows, cells, by form month), sample
    `order-enquiry match:` lines, why Order Enquiry could not resolve the rest (`GT has >1 live order` /
    `only cancelled orders` / `unique order already invoiced` / …), and `unmatched_rows.csv` reasons are now
    `sku_differs` / `punched_not_invoiced_multi_order` / `vinculum_order_cancelled` / `gt_not_in_vinculum`
    with an `order_enquiry_orders_for_gt` column. Fill-rate tables unchanged (the projected one includes the
    new fills).
- **Evidence:**
  - Drive connector 2026-10-10: folder holds 30 daily `OrderEnquiry_DD-MM-YYYY` Google Sheets
    (07-09-2026 … 09-10-2026; 14-09, 19-09, 20-09 missing) + the 5 monthly dumps above. Each daily file =
    orders of ONE order date (e.g. `OrderEnquiry_06-10-2026`: 330 rows, 330 unique Order Nos, all
    `06/10/2026`); no Order No repeats across the 11 files checked. Statuses seen on GT orders: Packed,
    Allocated, Confirmed, Shipped complete, Part Allocated, Partially Shipped, Part Picked (Cancelled on
    non-GT orders).
  - In those 11 files 74 GT bases: 60 have one order, 14 several; all 14 multi-order GTs had every order
    punched on the same day (so a complete daily file shows the full split).
  - New `tests/test_order_enquiry.py` (24 tests) against the OLD script: `24 failed`
    (`AttributeError: … has no attribute 'read_order_enquiry_daily_files'` etc.).
  - After the fix: `python -m pytest tests -q` -> `75 passed in 0.57s` (51 existing + 24 new).
  - Reader on the real xlsx export of `OrderEnquiry_01-10-2026`: `364 rows -> 364 unique Order Nos …
    GT orders 18 (cancelled 0); unparsed Order Dates 0`, `N89645 GT011026A6 2026-10-01 18:19:00 Part Allocated`.
  - CI dry-run / write: see "Live results" below.
- **Blast radius:** only the Full Flow script (P, S, and V/X via Pick Pack on rows that get a P this way;
  appended rows likewise). Drive access read-only. POD script matches by GT + Order No, so rows with a new P
  become eligible for POD lookup — intended. Planned-status script reads S (stage-1 Actual) — rows now get
  their stage-1 actual on punch day instead of invoice day, which is the requested behaviour. Run time: +30
  small Drive exports.
- **Still open:**
  - Q (Team) cannot come from the daily export (no "Created By" column); it fills when invoiced. Adding the
    column to the export lives in `nasher-miles-report-automation` (not touched here).
  - GTs split over several Vinculum orders are not resolved (no SKU in the file) — listed as
    `punched_not_invoiced_multi_order`. A line-level export (the monthly dumps have SKU + Created by) would
    resolve them.
  - Status is a snapshot from the export day: an order cancelled after its export still looks live. Only
    blank cells are filled, so this can at worst put a cancelled order's number on a still-blank row.
  - A GT punched as one order and split again on a later day would have got the first order on all its
    rows; not seen in the data (all splits same-day), never overwritten automatically.
  - Daily files start 07-09-2026: August rows (and early-September orders in the 3 missing days) still rely
    on the Sales Master.
- **Related:** [2026-10-09-full-flow-rows-empty-suffix-and-refresh.md](2026-10-09-full-flow-rows-empty-suffix-and-refresh.md)
  (suffix rule reused), [2026-10-10-pick-pack-google-sheets-skipped.md](2026-10-10-pick-pack-google-sheets-skipped.md)
  (reader reused), [2026-10-05-full-flow-blank-order-no-never-backfilled.md](2026-10-05-full-flow-blank-order-no-never-backfilled.md).

## Live results

- **Dry-run 38035984741 (commit e931ec0): success.** `Order Enquiry daily files found: 30; other files in the
  folder ignored: 5` (the monthly dumps), `Order Enquiry files read: 30; rows: 11,631`, `11,631 unique Order
  Nos (0 repeated across files); GT orders 285 (cancelled 12); unparsed Order Dates 0`.
  `Rows with blank Order No considered: 771` -> `filled from Sales Master (invoiced lines): 18` (162 cells;
  invoices that landed since the morning run) + `filled from Order Enquiry (punched orders): 33` (66 cells:
  `oe_fill Order No 33`, `oe_fill Actual 33`, nothing else; all form month 2026-10), e.g.
  `row 8197 GT081026A1 / NMHPP0816PS <- Vinculum GT081026A1 = N012268 (Packed, 08-Oct-2026 00:00:00)`,
  `row 8193 GT071026A2 / NMH801912LS <- … = N89667 (Confirmed, 07-Oct-2026 18:47:00)`.
  Not resolved: `GT has >1 live order 337 | only cancelled orders 112 | unique order already invoiced (SKU
  not on it) 35 | GT in no Vinculum source 236`. Projected October P / Q / S: 51 / 51 / 51 -> 102 / 69 / 102
  of 178. No other month changes.
- **Write 38036427347: success** — `Fill/refresh complete: 228 cells written.` (162 + 66),
  `New Full Flow rows appended: 17` (GT091026A1…A12: not yet punched / GT091026A1 has 2 orders -> P blank),
  POD + planned-status steps success.
- **Verification dry-run 38036850174** (reads the sheet back): `filled from Sales Master: 0`,
  `filled from Order Enquiry: 0`, `refresh cells: 0`, `Would append 0 rows` (idempotent).
  Live open-row fill rate, P / Q / S:

  | form month | rows | P | Q | S |
  |---|---|---|---|---|
  | 2026-04 … 07 | 4,368 | 4,368 | 4,368 | 4,368 |
  | 2026-08 | 467 | 363 | 363 | 363 |
  | 2026-09 | 1,315 | 775 | 775 | 775 |
  | 2026-10 | 195 (178 + 17 appended) | **102** (was 51) | 69 | **102** |

  `Rows with an Order No not in the Sales Master (no invoice refresh yet): 49` (16 before + the 33 OE rows;
  they refresh Q/AW/AX/AY/AS once invoiced).
- Remaining blank P (737), `unmatched_rows.csv` reasons: `sku_differs 357` (GT invoiced, SKU not on an
  invoiced line; mostly big GTs split into several orders incl. an extra `N461…` "Confirmed" order),
  `punched_not_invoiced_multi_order 19`, `vinculum_order_cancelled 112` (105 of them GT070926A5: all six
  Vinculum orders Cancelled — needs an Ops decision), `gt_not_in_vinculum 249` (Aug 49, Sep 171, Oct 29 incl.
  the 12 new GT091026 rows: punched before 07-09, on the missing file days 14/19/20-09, or not punched).
