# 2026-11-17 - Nightly POD job reports only (approval-only)

User: "everything goes via approval". scripts/gt_full_flow_fms_pod_update_google_sheet_gt_and_order_only.py (commit 32f840b) no longer writes AJ / AP / AU / AV
from Drive files: it only reports what it would link. PODs are written by a person (dashboard upload, Replace POD, review-panel Save / Approve).
Test: tests/test_pod.py. Committed locally only (no push).
