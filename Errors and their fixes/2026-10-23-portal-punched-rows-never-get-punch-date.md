# Full Flow "gap" 12-15 Oct: portal-punched GTs have an Order No (P) but no punch date (S), so the dashboard shows them as not punched

- **Date:** 2026-10-23
- **Reported as:** "why is there a gap in 12th to 15th October in Full Flow FMS please check and fix that"
- **Symptom:** no rows are missing. The gap is in the columns. Every GT approved on 12, 14, 15 and 16-Oct is in "Full Flow FMS". GT Master Orders, GT Pull FMS and Full Flow have the same lines. 13-Oct has no GT at all in GT Master Orders. 11, 17 and 18-Oct are Sun/Sat/Sun. What is blank (export of 2026-10-23 ~11:00 IST, form date in A):

  | form date | rows | P Order No | S punch date | V/X picked | AW invoice |
  |---|---|---|---|---|---|
  | 09-Oct | 16 | 13 | 13 | 9 | 0 |
  | 10-Oct | 40 | 25 | 25 | 1 | 0 |
  | 12-Oct | 120 | 96 | **5** | 1 | 3 |
  | 14-Oct | 93 | 72 | 72 | 72 | **0** |
  | 15-Oct | 37 | 36 | **12** | **0** | **0** |
  | 16-Oct | 117 | 117 | 97 | 19 | 97 |
  | 19-Oct | 145 | 94 | 39 | 37 | 38 |

  On 12 and 15-Oct, 115 rows have P but no S: 91 on 12-Oct (GT121026A1: N461768, M892787, N89722, N012271 and the split "N012271, N461768") and 24 on 15-Oct (M892792/93/96/97, N89725/26). The FMS_Dashboard (GT_FMS/dashboard.js `punch: !!ffToDate_(r[C_S1_ACTUAL])`) therefore lists them under "OMS Punch".
- **Where:** `scripts/gt_full_flow_fms_yes_no_all_or_latest.py`, `plan_existing_writes`, in the branch "P filled but no Sales Master line" (`order_no_without_line`). It only looked up Pick Pack (V/X). `allocate_lines` anchored a row only when P equalled the line's Order No, so a split P "N012271, N461768" never anchored.
- **Root cause:** since 12/13-Oct the auto-punch portal (NM_abhi_viniculum_testing `punch_run.js`, PN_GT_COL.VIN_ORDER = col P) writes the Vinculum Order No into P at punch time, but not S. The nightly job filled S only for a *blank* P (from daily Order Enquiry) or from an invoiced Sales Master line. So a portal-punched GT keeps S blank until it is invoiced.
- **Contributing (data, not this repo):**
  - Order Enquiry and Pick Pack export only `LOCATIONS_WITH_ORDERS = "A03,GJ1,NM1,NM4,NM7,NM9"` (nasher-miles-report-automation `common/vinculum_ui.py`, since 21-Oct). The GT warehouses N10 (N01 series), HO (N46), A06 (P19) and NM8 (O85) are left out. N012271 (N10) is in the company-wide Sales file 13.10.2026 (invoice 27/10/N10/1471), and N10 has 208 of that day's 828 lines, but the order is in no OrderEnquiry_* file. N461768 (HO) is not in any file. So the new S fill cannot reach these orders (on 12-Oct, 51 of the 91 rows) until the Order Enquiry backfill includes those locations. In the offline dry run below, only 40 of the 91 rows filled from OrderEnquiry_12-10 alone.
  - V/X stay blank because PickPackReport_11.10 … 18.10 are NM4-only (2-4 KB, made with the new login before the all-location change).
  - AW stays blank because the daily Sales files 12.10, 20.10 and 22.10 do not exist, and 14.10, 15.10 and 16.10 are empty (~1.5 KB). The Sales backfill run 39716075711 was cancelled. These are fixed in nasher-miles-report-automation the same day.
  - The job reads daily Sales files only 7 days back, so 12-15 Oct invoices would be missed by now even after a backfill, unless the Sales Master consolidation imports them.
- **Why it wasn't caught:**
  - The run log printed "Rows with an Order No not in the Sales Master (no invoice refresh yet): 388" with 0 refreshes as information only.
  - The October fill rate (P 614 vs S 389) was never alerted on.
  - The tests asserted the old behaviour: `test_never_overwrite_and_closed_untouched` expected no write at all for a row with P filled.
- **Fix:**
  - New `_order_no_cells`. For a row with P filled and no Sales line, it writes blank cells only and never writes P:
    - S (and Q when the export has "Created by") come from the daily Order Enquiry record of P's order. With several orders in P, the earliest Order Date is used. Cancelled orders are ignored.
    - V/X come from Pick Pack. With several orders in P, they are written only when every order has a Pick Date, and the latest date is used.
  - New `split_order_nos` / `_Slot.orders`: P "N012271, N461768" is split on `, ; / | &`. A Sales line of any of those orders anchors the row (invoice refresh).
  - The log prints a new line: "of which get S/Q (punch date) from Order Enquiry by their own P".
  - Env `SALES_DAILY_LOOKBACK_DAYS` (workflow input `sales_lookback_days`, default 7, 1-62) lets a manual catch-up run read older daily Sales files.
- **Evidence:**
  - GT Master / GT Pull / Full Flow comparison: 0 approved lines missing on any day from 09-Oct to 22-Oct. GT Master Orders has no timestamp on 13-Oct.
  - Drive full-text search finds N461768/M892787/N89722/N012271 in OrderEnquiry_12-10-2026 and N89725/M892792/M892796/N89726 in OrderEnquiry_15-10-2026, so the data to fill S is there.
  - Run 39865830465 (2026-10-22T22:33Z), from its log: "filled from Order Enquiry (punched orders): 0", "Rows with an Order No not in the Sales Master (no invoice refresh yet): 388", fill rate 2026-10 "rows 1097, P 614, S 389, V 258, AW 244".
  - Offline dry run: `plan_backfill` on the real Full Flow export (2026-10-23) with OrderEnquiry_12-10-2026 and no Sales lines. The new code plans `('2026-10-12', 'Actual', 'oe_refresh') 40` (M892787 ×18, N89722 ×21, O94133 ×1). The old code plans 0. Still blank: N461768 ×25, N012271 ×17, "N012271, N461768" ×8 (not in Order Enquiry), and the cancelled TEST order M892786.
  - `python -m pytest tests -q` gave 160 passed. The new file `tests/test_order_no_punch_date.py` covers the punch date, Team, split P (earliest date / all-picked rule), cancelled, untouched S/closed, the anchor of a split P, idempotence and the lookback override. `test_never_overwrite_and_closed_untouched` was updated (P still never overwritten; S now filled).
- **Blast radius:**
  - Existing open rows with a filled P and no Sales line: about 388 in October. They get S (+ Q if the export has "Created by"), and V/X where Pick Pack has them.
  - The planned-status script then computes T/U for them, and the dashboard stops offering "OMS Punch" for punched GTs.
  - No P, invoice or BD cell is written by the new path.
- **Still open:**
  - Re-run Pick Pack 11/10-18/10 (all locations) and the Sales backfill, then run "GT FMS Automation" with `sales_lookback_days` covering 12.10 onward (see the session report for the exact sequence).
  - The Sales Master consolidation skips files by name (`_ProcessedFiles`). Re-uploaded 14/15/16.10 files are therefore not re-imported unless their names are removed from that tab first.
  - P blank on GT121026A2/A3, GT141026A2 and GT151026A6 (46 rows): not punched yet, or the GT has more than one live order.
- **Related:** `2026-10-10-order-no-blank-until-invoiced.md`, nasher-miles-report-automation `2026-10-15-punched-gt-orders-still-shown-as-not-punched.md`, `2026-10-19-new-vinculum-login-exports-only-nm4.md`.
