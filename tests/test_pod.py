"""pod.py: a found POD never overwrites a closed line (Cancelled / Virtual Invoice / OOS).

GT_FMS markVirtualGT_ writes "Virtual Invoice" into ONLY AJ on punched lines; before
2026-10-09 row_is_cancelled looked at C / P / AP only, so AJ became "Delivered".
"""
from datetime import datetime

import pytest


def sheet_row(pod, **cells):
    row = [""] * 48                           # A:AV
    letters = {"C": pod.COL_GT, "P": pod.COL_ORDER, "AJ": pod.COL_AJ, "AP": pod.COL_AP,
               "AU": pod.COL_AU, "AV": pod.COL_AV}
    for col, value in cells.items():
        row[letters[col] - 1] = value
    return row


@pytest.fixture
def with_pod(pod, monkeypatch):
    monkeypatch.setattr(pod, "pod_index", {
        "N1": [{"gt": "GT1", "date": datetime(2026, 9, 10), "url": "https://drive/x"}],
    })
    return pod


def test_virtual_invoice_in_aj_is_protected(with_pod):
    rows = [sheet_row(with_pod, C="GT1", P="N1", AJ="Virtual Invoice")]
    *_, updates = with_pod.build_sheet_update_requests(rows)
    assert updates == []


def test_virtual_invoice_stamped_across_p_aw_is_protected(with_pod):
    rows = [sheet_row(with_pod, C="GT1", P="N1", AP="Virtual Invoice")]
    *_, updates = with_pod.build_sheet_update_requests(rows)
    assert updates == []


def test_cancelled_is_protected(with_pod):
    rows = [sheet_row(with_pod, C="GT1", P="N1", AJ="Cancelled")]
    *_, updates = with_pod.build_sheet_update_requests(rows)
    assert updates == []


def test_open_line_with_pod_is_delivered(with_pod):
    rows = [sheet_row(with_pod, C="GT1", P="N1")]
    *_, updates = with_pod.build_sheet_update_requests(rows)
    ranges = {u["range"]: u["values"][0][0] for u in updates}
    assert ranges["'Full Flow FMS'!AJ8"] == "Delivered"
    assert "'Full Flow FMS'!AV8" in ranges and "'Full Flow FMS'!AU8" in ranges


# 17-Nov-2026 approval-only policy: the nightly job never writes a POD cell.
def test_nightly_pod_job_never_writes_to_the_sheet(with_pod):
    class Boom:
        def spreadsheets(self):
            raise AssertionError("the nightly POD job called the Sheets API to write")

    rows = [sheet_row(with_pod, C="GT1", P="N1")]
    *_, updates = with_pod.build_sheet_update_requests(rows)
    assert updates, "the plan itself still finds the link (reported only)"
    assert with_pod.POD_WRITES_ENABLED is False
    assert with_pod.apply_sheet_updates(Boom(), updates) == 0


# 17-Nov-2026: the AJ Invoiced / Dispatched stages (from AW / AE, not a POD) still go through; POD cells never do.
def test_approval_only_keeps_stage_writes_drops_pod_writes(with_pod):
    sent = []

    class Values:
        def batchUpdate(self, spreadsheetId=None, body=None):
            sent.extend(body["data"])

            class R:
                def execute(self_inner):
                    return {"totalUpdatedCells": len(body["data"])}
            return R()

    class Sheets:
        def spreadsheets(self):
            class S:
                def values(self_inner):
                    return Values()
            return S()

    updates = [
        {"range": "'X'!AJ5", "values": [["Invoiced"]], "pod": False},
        {"range": "'X'!AU6", "values": [["=HYPERLINK(\"u\",\"d\")"]]},
        {"range": "'X'!AJ6", "values": [["Delivered"]]},
    ]
    with_pod.apply_sheet_updates(Sheets(), updates)
    ranges = [d["range"] for d in sent]
    assert ranges == ["'X'!AJ5"]
    assert all("pod" not in d for d in sent)
