"""audit_fabricated_actuals.py: finds AB==AA / AG==AF / AL==AK and is dry-run by default."""
import sys

import pytest

from conftest import SCRIPTS, _load


@pytest.fixture(scope="module")
def au():
    _load("gt_full_flow_fms_planned_status_google_sheet_alerts_quota_fixed_v3.py",
          "gt_full_flow_fms_planned_status_google_sheet_alerts_quota_fixed_v3")
    sys.path.insert(0, SCRIPTS)
    return _load("audit_fabricated_actuals.py", "audit_fab")


def _row(**cells):
    from string import ascii_uppercase as U
    letters = list(U) + ["A" + c for c in U[:19]]          # A..AS
    row = [""] * len(letters)
    for letter, value in cells.items():
        row[letters.index(letter)] = value
    return row


FAKE = dict(C="GT1", P="N1",
            AA="19-Aug-2026 12:30:00", AB="19-Aug-2026 12:30:00", AC="On Time", AD="00h 00m 00s",
            AF="19-Aug-2026 18:30:00", AG="19-Aug-2026 18:30:00", AH="On Time", AI="00h 00m 00s",
            AK="19-Aug-2026 19:00:00", AL="19-Aug-2026 19:00:00", AM="On Time", AN="00h 00m 00s")


def test_fingerprint_found_per_stage(au):
    real = dict(FAKE, AL="21-Aug-2026 10:00:00", AM="Delay", AN="1d 15h 00m 00s")
    blank = dict(C="GT3", AA="19-Aug-2026 12:30:00")
    report = au.find_fingerprints([_row(**FAKE), _row(**real), _row(**blank)])
    assert [(r[0], r[3]) for r in report] == [
        (8, "courier"), (8, "label"), (8, "dispatch"), (9, "courier"), (9, "label")]
    cells = au.cells_to_clear(report[:1])
    assert cells == [("AB8", "19-Aug-2026 12:30:00"), ("AC8", "On Time"), ("AD8", "00h 00m 00s")]
    assert "Rows with all three       : 1" in au.summarize(report, 3)


class _Recorder:
    """Fake Sheets client: serves one read, records every other call."""
    def __init__(self, values):
        self.values_data, self.calls = values, []

    def spreadsheets(self):
        return self

    def values(self):
        return self

    def get(self, **kw):
        return _Exec({"values": self.values_data})

    def __getattr__(self, name):                # update / batchUpdate / batchClear
        def call(**kw):
            self.calls.append(name)
            return _Exec({})
        return call


class _Exec:
    def __init__(self, result):
        self.result = result

    def execute(self):
        return self.result


def test_dry_run_writes_nothing(au, monkeypatch, tmp_path):
    fake = _Recorder([_row(**FAKE)])
    monkeypatch.setattr(au, "connect", lambda sa: fake)
    out = tmp_path / "r.csv"
    au.main(["--csv", str(out)])
    assert fake.calls == []
    assert out.read_text(encoding="utf-8").count("\n") == 4      # header + 3 stages


def test_apply_backs_up_before_clearing(au, monkeypatch, tmp_path):
    fake = _Recorder([_row(**FAKE)])
    monkeypatch.setattr(au, "connect", lambda sa: fake)
    au.main(["--csv", str(tmp_path / "r.csv"), "--apply"])
    assert fake.calls == ["batchUpdate", "update", "batchClear"]  # add tab, backup, clear
