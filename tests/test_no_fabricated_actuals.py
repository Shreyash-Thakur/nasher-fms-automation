"""
Planned-status step must never fabricate AB / AG / AL Actuals  —  2026-11-11.

resolve_stage_actual wrote the Planned datetime into a blank courier (AB), label (AG)
and dispatch (AL) Actual, then AC/AH/AM read "On Time" from that fake value: 6,640 of
6,640 datetime Actuals equalled Planned. Now:
  - a blank Actual stays blank, and its Status / Time Delay stay blank (no alert);
  - the Planned chain still runs, from the previous Planned when the Actual is blank;
  - a real Actual still gives Early / On Time / Delay + delay text;
  - AP (POD) is written date-only and AR is measured in days.
"""
from datetime import datetime

import pytest

from conftest import _load

SHEET = "'Full Flow FMS'!"


@pytest.fixture(scope="module")
def st():
    return _load("gt_full_flow_fms_planned_status_google_sheet_alerts_quota_fixed_v3.py",
                 "plannedstatus_nofab")


def _row(**cells):
    """Row of formatted strings as read from the sheet (A..AS), cells by column letter."""
    from string import ascii_uppercase as U
    letters = list(U) + ["A" + c for c in U[:19]]          # A..AS
    row = [""] * len(letters)
    for letter, value in cells.items():
        row[letters.index(letter)] = value
    return row


def _run(st, row_values, row=100, holidays=frozenset()):
    updates, fmt = [], []
    result = st.process_row(row_values, row, set(holidays), updates, fmt, 0)
    cells = {u["range"].replace(SHEET, ""): u["values"][0][0] for u in updates}
    return result, cells, fmt


PICKED = dict(A="18-Aug-2026 10:00:00", S="18-Aug-2026 00:00:00",
              X="19-Aug-2026 12:00:00")


def test_blank_actuals_are_never_written(st):
    _, cells, _ = _run(st, _row(**PICKED))
    for col in ("AB", "AG", "AL"):
        assert f"{col}100" not in cells
    assert st.internal_actual_writes(
        [{"range": f"{SHEET}{c}", "values": [[v]]} for c, v in cells.items()]) == []


def test_resolve_stage_actual_is_read_only(st):
    assert st.resolve_stage_actual(_row(**PICKED), st.COL_AB) is None
    assert st.resolve_stage_actual(_row(AB="19-Aug-2026 13:00:00"), st.COL_AB) == \
        datetime(2026, 8, 19, 13, 0)


def test_planned_still_chains_from_planned(st):
    result, cells, _ = _run(st, _row(**PICKED))
    assert cells["AA100"] == "19-Aug-2026 12:30:00"          # X + 30m
    assert cells["AF100"] == "19-Aug-2026 18:30:00"          # AA + 6h (AB blank)
    assert cells["AK100"] == "19-Aug-2026 19:00:00"          # AF + 30m (AG blank)
    assert cells["AO100"] == "20-Aug-2026 19:00:00"          # AK + 24h (AL blank)
    assert result["AO_basis"] == "AK"


def test_chain_uses_real_actual_when_present(st):
    _, cells, _ = _run(st, _row(**PICKED, AB="20-Aug-2026 09:00:00"))
    assert cells["AF100"] == "20-Aug-2026 15:00:00"          # real AB + 6h
    assert cells["AK100"] == "20-Aug-2026 15:30:00"


def test_chain_skips_holidays_from_planned(st):
    hol = {datetime(2026, 8, 20).date()}
    _, cells, _ = _run(st, _row(**PICKED), holidays=hol)
    assert cells["AO100"] == "21-Aug-2026 19:00:00"          # 20-Aug skipped


def test_no_status_or_alert_for_blank_ab_ag_al(st):
    _, cells, fmt = _run(st, _row(**PICKED))
    for col in ("AC", "AD", "AH", "AI", "AM", "AN"):
        assert f"{col}100" not in cells
    assert fmt == []                                          # no red alert formatting


def test_real_late_dispatch_gives_delay(st):
    _, cells, _ = _run(st, _row(**PICKED, AL="22-Aug-2026 11:00:00"))
    assert cells["AM100"] == "Delay"
    assert cells["AN100"] == "2d 16h 00m 00s"                 # vs AK 19-Aug 19:00
    assert cells["AO100"] == "23-Aug-2026 11:00:00"          # real AL + 24h
    assert "AC100" not in cells and "AH100" not in cells     # AB / AG still blank


def test_real_actual_same_day_is_on_time_and_earlier_is_early(st):
    _, cells, _ = _run(st, _row(**PICKED, AB="19-Aug-2026 14:00:00",
                                AG="19-Aug-2026 08:00:00"))
    assert cells["AC100"] == "On Time"                        # same day as AA
    assert cells["AD100"] == "01h 30m 00s"
    # AF = real AB 14:00 + 6h = 20:00; AG on 19-Aug 08:00 -> same day -> On Time
    assert cells["AH100"] == "On Time"
    _, cells, _ = _run(st, _row(**PICKED, AB="18-Aug-2026 20:00:00"))
    assert cells["AC100"] == "Early"


def test_no_pick_no_internal_stages(st):
    _, cells, _ = _run(st, _row(A="18-Aug-2026 10:00:00", S="18-Aug-2026 00:00:00"))
    assert not any(k.startswith(("AA", "AF", "AK", "AO")) for k in cells)


def test_guard_detects_internal_actual_writes(st):
    ups = [{"range": f"{SHEET}AB12", "values": [["x"]]},
           {"range": f"{SHEET}AA12", "values": [["x"]]},
           {"range": f"{SHEET}AL9", "values": [["x"]]}]
    assert st.internal_actual_writes(ups) == [f"{SHEET}AB12", f"{SHEET}AL9"]


def test_ap_delay_in_days(st):
    _, cells, _ = _run(st, _row(**PICKED, AP="23-Aug-2026 00:00:00"))
    assert cells["AQ100"] == "Delay"
    assert cells["AR100"] == "3d"                             # 20-Aug -> 23-Aug


# ---- POD step: AP date-only -------------------------------------------------

def _pod_row(pod, **cells):
    row = [""] * 49                                           # A:AW
    letters = {"C": pod.COL_GT, "P": pod.COL_ORDER}
    for col, value in cells.items():
        row[letters[col] - 1] = value
    return row


def test_ap_written_date_only(pod, monkeypatch):
    monkeypatch.setattr(pod, "exact_index", {
        "GT1|N1": {"date": datetime(2026, 9, 10), "url": "https://drive/x"}})
    *_, updates = pod.build_sheet_update_requests([_pod_row(pod, C="GT1", P="N1")])
    cells = {u["range"]: u["values"][0][0] for u in updates}
    assert cells["'Full Flow FMS'!AP8"] == "10-Sep-2026"
    assert cells["'Full Flow FMS'!AV8"] == "10-Sep-2026"
