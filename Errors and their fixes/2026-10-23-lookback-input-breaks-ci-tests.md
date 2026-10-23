# Full Flow catch-up run with sales_lookback_days=13 failed at "Run tests" and wrote nothing

- **Date:** 2026-10-23
- **Reported as:** N/A (found while running the 12–15 Oct gap catch-up: `gh workflow run run_fms.yml -f backfill_mode=write -f sales_lookback_days=13`, run 39918292343 → "completed failure")
- **Symptom:** step "Run tests" failed: `FAILED tests/test_freshness.py::test_read_sales_daily_files_reads_sales_report_tab - AssertionError: assert (('s12', 'Sales Report') == ('s12', 'Sales Report')`; the Full Flow step was then skipped, so nothing was filled.
- **Where:** `.github/workflows/run_fms.yml` job-level `env: SALES_DAILY_LOOKBACK_DAYS: ${{ github.event.inputs.sales_lookback_days || '7' }}` (added in f8f10ac); read at import by `scripts/gt_full_flow_fms_yes_no_all_or_latest.py:82`.
- **Root cause:** the lookback override was set for the whole job, so the offline test step imported the script with 13 days instead of the 7-day default the test is written for, and the test's "older file is not read" assertion failed.
- **Why it wasn't caught:** f8f10ac's tests ran locally without the env var; nothing exercised the workflow with the new input set.
- **Fix:** the env var is now set only on the "Run Full Flow FMS script" step, so tests always see the default and catch-up runs still get the override.
- **Evidence:** `SALES_DAILY_LOOKBACK_DAYS=13 python -m pytest tests -q` → `1 failed, 159 passed` (reproduces CI); without it → `160 passed`; workflow YAML parses (`yaml.safe_load` ok).
- **Blast radius:** only manual runs that pass `sales_lookback_days`; scheduled runs were unaffected (default 7).
- **Still open:** re-run the catch-up after this commit.
- **Related:** 2026-10-23-portal-punched-rows-never-get-punch-date.md
