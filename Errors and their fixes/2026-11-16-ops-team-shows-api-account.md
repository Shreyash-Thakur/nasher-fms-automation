# GT Full Flow col Q "Team" filled with the API account ("Diwakar") by the nightly job

- **Date:** 2026-11-16
- **Reported as:** "why is ops team coming as diwakar can we setup a way to fix the api pushed to avoid diwakar's name and have the person who punched it in the FMS"
- **Symptom:** Q = "Diwakar" (59 cells) and "finance.automation" (13) on GT Full Flow rows; already replaced by the real person earlier on 16-Nov (GT_FMS GtTeamFix). None left today.
- **Where:** `scripts/gt_full_flow_fms_yes_no_all_or_latest.py`: `OrderEnquiryIndex.__init__` team, new-row `output_row["Team"] = sales_row["Created by"]`, `BACKFILL_FIELDS` Team. Both refresh paths (`_oe_cells`, `_order_no_cells`) take the team from the index.
- **Root cause:** Vinculum "Created by" of every portal-punched order is the API key owner. The job copied it into blank Q.
- **Why it wasn't caught:** no test fed an API creator.
- **Fix:** `team_from_creator()`: blank or API account (diwakar, finance.automation) -> None, Q stays blank (the portal / GT_FMS GtTeamAuto fill the real person). `tests/test_team_api_account.py`, 188 tests pass.
- **Evidence:** `python -m pytest tests -q` -> 188 passed.
- **Blast radius:** Q only, blank cells only.
- **Still open:** the job needs to run for the change to apply (nightly Actions; committed locally, not pushed).
- **Related:** KASA_FMS `2026-11-16-ops-team-shows-api-account.md`.
