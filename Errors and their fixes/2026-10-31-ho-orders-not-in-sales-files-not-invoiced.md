# Invoice / Team (AW/Q…) blank on 116-119 rows whose orders are in no daily Sales Report: Vinculum still shows them open (not invoiced)

- **Date:** 2026-10-31
- **Reported as:** "Im seeing so many gaps in the fms." Then: "whatever's mentioned in the screenshots try to fill as many as you can use the vin api any source you want".
- **Symptom:** P is filled but Q, AS, AW, AX and AY are blank on the rows of 9 orders that appear in no daily Sales Report file: N461768, N461774, N461775, N461777, N461789 (HO), N89737, N89738, N051649 and N051595. The audit counted 116 rows. The 31-Oct export has 119 open rows, including split P values such as "N012271, N461768".
- **Where:** no defect in `scripts/gt_full_flow_fms_yes_no_all_or_latest.py`. The invoice chain (`BACKFILL_FIELDS`) is filled only from a Sales line. The question was whether another source the job already reads carries these invoices.
- **Root cause:** these lines are not invoiced yet. Every source the job reads was checked:

  | Source the job reads | Invoice No? | Has these orders? |
  |---|---|---|
  | Sales Master (Order Enquiry spreadsheet 1Jnt0yU…, month tabs) | yes | no. It is consolidated from the same daily Sales files. |
  | Daily Sales Report files (1iL94fl…) | yes | no. Drive full-text search found none of the 9. |
  | Daily OrderEnquiry_* files (1rinwTy…) | no (Order No, ExtOrderNo, Order Date, Status only) | HO / N10 are not exported at all (`LOCATIONS_WITH_ORDERS`; another repo, out of scope). |
  | Pick Pack reports (1Xi66o…) | no | not relevant to AW. |

  Vinculum's own Open Qty report of 30-Oct 15:30 (`3010261530_vinopereport`, Drive folder 1mJZXLqEPCzEKTsZYtRllEU8RNnSSPNn9) still lists **110 of the 119 rows** (same order **and** SKU) with open quantity. These were not shipped and not invoiced on 30-Oct, so the blank AW is correct.
  - The other 9 rows are GT010926A1 N051595 NMB80583CONE, GT211026A1R N051649 (NM810216PSML, NM810258FSML: S3-set codes), and GT291026A3/A4 N89737/N89738 (NMACCTK801578ZS, NMB80303CM, NMB80313CM).
  - Their SKU is not open on the order. Most likely it was never on the order (set split into components, or short-supplied), not invoiced elsewhere.
- **Why it wasn't caught:**
  - The audit split blanks by "order found in a Sales file by Drive search". It had no positive signal for "not invoiced".
  - The job does not read the Open Qty report.
- **Fix:** none needed for these rows. They fill automatically, blank cells only, on the first nightly run after the invoice appears in a daily Sales Report within the lookback window.
  - Possible improvement, not built: read the newest `*_vinopereport` and mark open lines in the run log or the unmatched report as "open in Vinculum (not invoiced)", so they stop showing up as gaps.
- **Evidence:**
  - Script over the 31-Oct export and the 30-Oct Open Qty CSV (16,515 rows, 703 orders). Order level: `('N461768', 'open') 25`, `('N461789', 'open') 28`, `('N461774', 'open') 20`, … `('N051595', 'not-open') 1`, total 119. Order + SKU level: `Counter({'sku open': 110, 'sku not open': 9})`.
  - Drive full-text search for N461768 / N461774 / N461789 finds only the 27/28/30-Oct Open Qty reports and the punch portal's test sheet. No Sales Report file has them.
- **Blast radius:** none. No code change.
- **Still open:**
  - Ops: ship or close these HO lines.
  - The 9 rows whose SKU is not open on the order need an ops decision (OOS / not supplied).
  - **Not merged to main** does not apply (record only, branch `fill-gaps-31oct`).
- **Related:** audit `gaps31/gt_audit.md` (116 / 9 GTs bucket), `2026-10-23-portal-punched-rows-never-get-punch-date.md` (HO/N10 missing from Order Enquiry).
