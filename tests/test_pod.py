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
    monkeypatch.setattr(pod, "exact_index", {
        "GT1|N1": {"date": datetime(2026, 9, 10), "url": "https://drive/x"},
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
