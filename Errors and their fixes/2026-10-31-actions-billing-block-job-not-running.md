# GT Full Flow not filled since 23-Oct: every GitHub Actions run refused for billing

- **Date:** 2026-10-31
- **Reported as:** Abhishek (30-Oct): "Im seeing so many gaps in the fms. Nothing close to what it should have been". User (31-Oct): "but i transfered the repo to my friend and i gave you the link still why isnt it running".
- **Symptom:** every scheduled "GT FMS Automation" run from 23-Oct evening to 30-Oct 23:04 UTC (16 runs) failed in ~4 s with no steps. The last success was 2026-10-23T07:33Z. Annotation on run 41142156435: "The job was not started because recent account payments have failed or your spending limit needs to be increased. Please check the 'Billing & plans' section in your settings". Full Flow columns S/T/U, Q/AW/AX/AY, V–AD and AQ/AR stopped filling for every GT after 23-Oct (GT audit of 31-Oct: ~342 S, ~251 invoice, ~51 Pick Pack, ~916 POD-status rows).
- **Where:** `.github/workflows/run_fms.yml` (job `run`, `runs-on: ubuntu-latest`); the repo was private on the personal account Shreyash-Thakur, whose Actions billing failed.
- **Root cause:** private-repo Actions minutes are billed to the owner's account; once that account's payment failed, GitHub refused to start any job. The repo had NOT been transferred: `gh repo view` still showed owner Shreyash-Thakur, private. sarthinasher/nasher-miles-report-automation is a different repo (it only uploads Vinculum report files to Drive) and never contained this job.
- **Why it wasn't caught:** a refused job sends only GitHub's default failure mail to the owner; nothing on the sheet or in the team's mail says the Full Flow stopped updating.
- **Fix:** the user chose to make the repo public ("make public whatever is necessary to make it fully functional and operastional"). Public repos run on standard GitHub-hosted runners for free, so the billing block no longer applies. No code change. A catch-up run was started by hand with `backfill_mode=write`, `sales_lookback_days=10` (covers the Sales files since 21-Oct).
- **Evidence:**
  - `gh repo edit Shreyash-Thakur/nasher-fms-automation --visibility public --accept-visibility-change-consequences` → `gh repo view … -q .visibility` = `PUBLIC`.
  - `gh workflow run run_fms.yml -f backfill_mode=write -f sales_lookback_days=10` → run 41191530872: `success`, every step success (tests, Full Flow, reports upload, POD update, planned status). Log: "Fill/refresh complete: 1,941 cells written.", "Nothing to append.", POD step "Total cells updated : 4", planned status "Google Sheet cells updated: 71,204".
- **Blast radius:** the code, the `Errors and their fixes/` notes, all run logs (GT numbers, Vinculum order and invoice numbers, SKUs) and the stored `fullflow-reports` artifacts (CSV lines incl. "Created by" staff names) are now public. The `GCP_SA_KEY` secret stays hidden. Scheduled runs (01:00 and 18:17 IST, earliest) resume.
- **Still open:**
  - The 46 older `fullflow-reports` artifacts are public until they expire (90 days) or are deleted by the owner (repo → Actions → run → Artifacts).
  - Account billing itself is still broken; making the repo private again would stop the job.
  - The remaining real gaps (N10/HO missing from the Vinculum exports, the 1,861-line POD backlog, portal not writing S) are not caused by the outage — see the 31-Oct GT audit and the records made with those fixes.
- **Related:** 2026-10-23-lookback-input-breaks-ci-tests.md, 2026-10-13-full-flow-data-older-than-24h.md
