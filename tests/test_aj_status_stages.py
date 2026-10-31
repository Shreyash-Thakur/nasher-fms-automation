"""pod.py: AJ "Order Delivery Status" stages  —  2026-11-01.

AJ used to be written only as "Delivered" (POD found), so 3,040 open lines read blank.
Ladder: blank < Invoiced (AW Invoice No) < Dispatched (AE AWB) < Delivered (POD).
AJ moves only UP from blank / a lower stage; never on a closed line (Cancelled / Virtual
Invoice / OOS in C or P:AW); a non-stage text in AJ is never replaced.
"""
from datetime import datetime

import pytest

LETTERS = {"C": 3, "P": 16, "AE": 31, "AJ": 36, "AP": 42, "AS": 45, "AU": 47, "AV": 48,
           "AW": 49}


def sheet_row(**cells):
    row = [""] * 49                           # A:AW
    for col, value in cells.items():
        row[LETTERS[col] - 1] = value
    return row


@pytest.fixture
def p(pod, monkeypatch):
    monkeypatch.setattr(pod, "exact_index", {
        "GT1|N1": {"date": datetime(2026, 9, 10), "url": "https://drive/x"},
    })
    monkeypatch.setattr(pod, "gt_index", {})
    return pod


def aj_writes(pod, rows):
    *_, updates = pod.build_sheet_update_requests(rows)
    return {u["range"]: u["values"][0][0] for u in updates if "!AJ" in u["range"]}


def test_invoiced_when_aw_filled_and_no_pod(p):
    rows = [sheet_row(C="GT2", P="N2", AW="27/10/NM1/164", AS="2")]
    assert aj_writes(p, rows) == {"'Full Flow FMS'!AJ8": "Invoiced"}


def test_no_invoice_no_write(p):
    assert aj_writes(p, [sheet_row(C="GT2", P="N2")]) == {}
    assert aj_writes(p, [sheet_row(C="GT2")]) == {}


def test_dispatched_when_awb_present(p):
    rows = [sheet_row(C="GT2", P="N2", AE="AWB123456", AW="27/10/NM1/164")]
    assert aj_writes(p, rows) == {"'Full Flow FMS'!AJ8": "Dispatched"}


def test_invoiced_promoted_to_dispatched(p):
    rows = [sheet_row(C="GT2", P="N2", AE="AWB1", AW="27/10/NM1/164", AJ="Invoiced")]
    assert aj_writes(p, rows) == {"'Full Flow FMS'!AJ8": "Dispatched"}


@pytest.mark.parametrize("current", ["Invoiced", "Dispatched", "invoiced", ""])
def test_pod_promotes_to_delivered(p, current):
    rows = [sheet_row(C="GT1", P="N1", AW="27/10/NM1/164", AJ=current)]
    assert aj_writes(p, rows) == {"'Full Flow FMS'!AJ8": "Delivered"}


@pytest.mark.parametrize("current,cells", [
    ("Delivered", {"AW": "27/10/NM1/164"}),          # no POD now: never back to Invoiced
    ("Delivered", {"AE": "AWB1"}),
    ("Dispatched", {"AW": "27/10/NM1/164"}),         # AWB gone: never back to Invoiced
    ("Invoiced", {"AW": "27/10/NM1/164"}),           # same stage: no rewrite
])
def test_never_downgraded_or_rewritten(p, current, cells):
    rows = [sheet_row(C="GT2", P="N2", AJ=current, **cells)]
    assert aj_writes(p, rows) == {}


@pytest.mark.parametrize("current", ["Returned", "RTO", "Partially delivered"])
def test_non_stage_text_never_replaced(p, current):
    rows = [sheet_row(C="GT1", P="N1", AW="27/10/NM1/164", AJ=current)]   # POD found too
    assert aj_writes(p, rows) == {}


@pytest.mark.parametrize("cells", [
    {"AJ": "Cancelled", "AW": "27/10/NM1/164"},
    {"AJ": "Virtual Invoice", "AW": "27/10/NM1/164"},
    {"AW": "Cancelled"},                              # GT_FMS Cancel stamps P:AW
    {"AW": "Virtual Invoice"},
    {"AE": "Cancelled", "AW": "27/10/NM1/164"},
    {"P": "Cancelled", "AW": "27/10/NM1/164"},
    {"AS": "OOS", "AW": "27/10/NM1/164"},
])
def test_closed_rows_untouched(p, cells):
    rows = [sheet_row(C="GT2", **{"P": "N2", **cells})]
    *_, updates = p.build_sheet_update_requests(rows)
    assert updates == []


def test_closed_marker_in_aw_protects_pod_row(p):
    rows = [sheet_row(C="GT1", P="N1", AW="Cancelled")]
    *_, updates = p.build_sheet_update_requests(rows)
    assert updates == []


def test_awb_closed_marker_is_not_dispatch_evidence(p):
    assert p.aj_stage_for_row(sheet_row(AE="Cancelled", AW="27/10/NM1/164"), False) == "Invoiced"


def test_short_row_from_api_trims_trailing_blanks(p):
    """The Sheets API drops trailing blank cells: a row ending before AW is no evidence."""
    row = sheet_row(C="GT2", P="N2")[:36]
    assert aj_writes(p, [row]) == {}


def test_mixed_batch(p):
    rows = [
        sheet_row(C="GT1", P="N1", AW="27/10/NM1/1"),          # 8: POD -> Delivered
        sheet_row(C="GT2", P="N2", AW="27/10/NM1/2"),          # 9: Invoiced
        sheet_row(C="GT3", P="N3"),                            # 10: nothing
        sheet_row(C="GT4", P="N4", AW="Cancelled", AJ=""),     # 11: closed
        sheet_row(C="GT5", P="N5", AW="27/10/NM1/5", AJ="Delivered"),   # 12: kept
    ]
    assert aj_writes(p, rows) == {"'Full Flow FMS'!AJ8": "Delivered",
                                  "'Full Flow FMS'!AJ9": "Invoiced"}
