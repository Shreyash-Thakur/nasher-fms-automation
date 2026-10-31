"""
AQ/AR (POD Status / Time Delay) when X (Pick Actual) is blank  —  2026-10-31.

The status step planned AO = AL + 24h, and AL only exists through the chain
X -> AB (+30m) -> AG (+6h) -> AL (+30m). With X blank (N10 orders missing from Pick
Pack, the missing PickPackReport_20.08.2026, ...) AO was never planned, so 198 rows with
a POD (AP) had a blank AQ/AR.

Rule: X blank AND AP is a date -> stand in for X with the next known EARLIER date of the
row: S (punch date), else A (form received). AO = that date + 30m + 6h + 30m + 24h
(each step skipping holidays, exactly like the normal chain). AQ/AR are computed from it.
AA..AL are NOT written (no fabricated internal stage actuals), and a row without a POD
gets nothing (no planned AO from a guess).
"""
from datetime import datetime

import pytest

from conftest import _load


@pytest.fixture(scope="module")
def st():
    return _load("gt_full_flow_fms_planned_status_google_sheet_alerts_quota_fixed_v3.py",
                 "plannedstatus")


def _row(st, **cells):
    """Row of formatted strings as read from the sheet (A..AR), cells by column letter."""
    from string import ascii_uppercase as U
    letters = list(U) + ["A" + c for c in U[:18]]          # A..AR
    row = [""] * len(letters)
    for letter, value in cells.items():
        row[letters.index(letter)] = value
    return row


def test_x_blank_pod_present_uses_s(st):
    row = _row(st, A="18-Aug-2026 16:40:48", S="19-Aug-2026 00:00:00", AP="24-Aug-2026 00:00:00")
    planned, basis = st.stage6_planned(row, None, set())
    assert basis == "S"
    assert planned == datetime(2026, 8, 20, 7, 0)            # 19-Aug 00:00 + 31h


def test_falls_back_to_form_date_when_s_blank(st):
    row = _row(st, A="18-Aug-2026 16:40:48", AP="24-Aug-2026 00:00:00")
    planned, basis = st.stage6_planned(row, None, set())
    assert basis == "A" and planned == datetime(2026, 8, 19, 23, 40, 48)


def test_holidays_skipped_like_the_chain(st):
    row = _row(st, S="19-Aug-2026 00:00:00", AP="24-Aug-2026 00:00:00")
    hol = {datetime(2026, 8, 20).date()}
    planned, _ = st.stage6_planned(row, None, hol)
    assert planned == datetime(2026, 8, 21, 7, 0)


def test_normal_chain_unchanged(st):
    al = datetime(2026, 8, 21, 9, 0)
    row = _row(st, S="19-Aug-2026 00:00:00", X="21-Aug-2026 02:00:00", AP="24-Aug-2026 00:00:00")
    assert st.stage6_planned(row, al, set()) == (datetime(2026, 8, 22, 9, 0), "AL")


@pytest.mark.parametrize("ap", ["", "POD pending", None])
def test_no_pod_no_fallback(st, ap):
    row = _row(st, S="19-Aug-2026 00:00:00", AP=ap)
    assert st.stage6_planned(row, None, set()) == (None, None)


def test_no_date_at_all(st):
    assert st.stage6_planned(_row(st, AP="24-Aug-2026 00:00:00"), None, set()) == (None, None)


def test_status_and_delay_written_only_ao_aq_ar(st):
    row = _row(st, S="19-Aug-2026 00:00:00", AP="24-Aug-2026 00:00:00")
    planned, _ = st.stage6_planned(row, None, set())
    updates = []
    status = st.process_stage(row, 6465, st.COL_AO, st.COL_AP, st.COL_AQ, st.COL_AR,
                              planned, updates, [], 0)
    assert status == "Delay"
    ranges = [u["range"] for u in updates]
    assert ranges == ["'Full Flow FMS'!AO6465", "'Full Flow FMS'!AQ6465",
                      "'Full Flow FMS'!AR6465"]
    assert updates[0]["values"] == [["20-Aug-2026 07:00:00"]]
    assert updates[2]["values"] == [["3d 17h 00m 00s"]]
