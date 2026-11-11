"""
Pick Date month/day swap  -  2026-11-11.

PickPackReport_03.07 .. 12.07.2026 were imported into Google Sheets with a US locale:
day <= 12 Pick Dates became real dates with month and day swapped (12-Jul -> 07-Dec)
and X got 07-Dec-2026 (GT100726A3 row 5431). Contract:
  - a Pick Date outside [report date - 31d, report date + 1d] whose swap lands inside
    is swapped back; everything else is left as parsed
  - the report date comes from the file name, so read_pick_pack_files carries it
  - X is written as "dd-Mon-yyyy hh:mm:ss" (month name: no locale can swap it)
  - repair_swapped_pick_dates.py is dry-run by default, flags only provable swaps,
    and with --apply backs up to a new tab before writing serials
"""
import io
import sys
from datetime import date, datetime

import openpyxl
import pandas as pd
import pytest

from conftest import SCRIPTS, _load

GSHEET = "application/vnd.google-apps.spreadsheet"
HEADER = ["Order No", "Extern Order No", "Order Date", "SKU Code", "Pick Date",
          "Pick User", "Qty"]


def _xlsx(rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    for r in [HEADER] + rows:
        ws.append(list(r))
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


@pytest.mark.parametrize("value, report, expected, swapped", [
    # the real cases: 12.07 file held 07-Dec, 03.07 file held 07-Mar
    ("2026-12-07 14:48", "2026-07-12", "2026-07-12 14:48", True),
    ("2026-03-07 16:46", "2026-07-03", "2026-07-03 16:46", True),
    ("2026-08-07 15:14", "2026-07-08", "2026-07-08 15:14", True),    # 08.07 file: 07-Aug
    # inside the window: untouched
    ("2026-07-12 14:48", "2026-07-12", "2026-07-12 14:48", False),
    ("2026-07-05 10:00", "2026-07-12", "2026-07-05 10:00", False),
    # day > 12 cannot be a swap; day == month swap is a no-op
    ("2026-05-20 10:00", "2026-07-12", "2026-05-20 10:00", False),
    ("2026-07-07 10:00", "2026-09-07", "2026-07-07 10:00", False),
    # neither value near the report (old real date in a later file): untouched
    ("2026-05-02 09:34:53", "2026-10-09", "2026-05-02 09:34:53", False),
    # no report date: untouched
    ("2026-12-07 14:48", None, "2026-12-07 14:48", False),
])
def test_fix_swapped_pick_date(ff, value, report, expected, swapped):
    rd = date.fromisoformat(report) if report else None
    got, was = ff.fix_swapped_pick_date(pd.Timestamp(value), rd)
    assert (pd.Timestamp(got), was) == (pd.Timestamp(expected), swapped)


def test_nat_and_none_pass_through(ff):
    assert ff.fix_swapped_pick_date(pd.NaT, date(2026, 7, 12))[1] is False
    assert ff.fix_swapped_pick_date(None, date(2026, 7, 12)) == (None, False)


def test_july_sheet_end_to_end(ff, monkeypatch, capsys):
    """PickPackReport_12.07.2026 as exported: real date cell 07-Dec -> lookup 12-Jul."""
    files = [{"id": "g12", "name": "PickPackReport_12.07.2026", "mimeType": GSHEET},
             {"id": "g13", "name": "PickPackReport_13.07.2026", "mimeType": GSHEET}]
    contents = {
        "g12": _xlsx([["O9486", "GT100726A3", "x", "S", datetime(2026, 12, 7, 14, 48), "krishna", 1]]),
        "g13": _xlsx([["O9500", "GT1", "x", "S", "13/07/2026 10:52 05", "Bidyut", 1]]),
    }

    class _Files:
        def list(self, **kw):
            return type("R", (), {"execute": lambda s, **k: {"files": files}})()

    monkeypatch.setattr(ff, "drive_service", type("D", (), {"files": lambda s: _Files()})())
    monkeypatch.setattr(ff, "export_drive_file", lambda fid, mime: io.BytesIO(contents[fid]))
    df = ff.read_pick_pack_files()
    assert list(df["_FILE_DATE"]) == [date(2026, 7, 12), date(2026, 7, 13)]
    lookup = ff.build_pick_pack_lookup(df.copy())
    assert lookup["O9486"]["Pick Date"] == pd.Timestamp("2026-07-12 14:48")
    assert lookup["O9500"]["Pick Date"] == pd.Timestamp("2026-07-13 10:52:05")
    assert "month/day swapped by the Sheets import, corrected: 1" in capsys.readouterr().out


def test_lookup_without_file_date_unchanged(ff):
    df = pd.DataFrame({"Order No": ["O1"], "Pick Date": ["2026-12-07 14:48:00"],
                       "Pick User": ["K"]})
    assert ff.build_pick_pack_lookup(df)["O1"]["Pick Date"] == pd.Timestamp("2026-12-07 14:48")


def test_x_written_with_month_name(ff):
    # USER_ENTERED re-parses strings: a month name cannot be read month-first.
    assert ff.clean_for_google(pd.Timestamp("2026-07-12 14:48")) == "12-Jul-2026 14:48:00"
    assert ff.clean_for_google(datetime(2026, 7, 3, 9, 5, 1)) == "03-Jul-2026 09:05:01"


# ------------------------------------------------------------------ repair script

@pytest.fixture(scope="module")
def rp():
    _load("gt_full_flow_fms_planned_status_google_sheet_alerts_quota_fixed_v3.py",
          "gt_full_flow_fms_planned_status_google_sheet_alerts_quota_fixed_v3")
    sys.path.insert(0, SCRIPTS)
    return _load("repair_swapped_pick_dates.py", "repair_swap")


def _row(rp, A=None, S=None, X=None, gt="GT1"):
    row = [""] * 24
    for col, v in ((1, A), (19, S), (24, X)):
        if v is not None:
            row[col - 1] = rp.datetime_to_serial(datetime.fromisoformat(v))
    row[2], row[15] = gt, "O1"
    return row


TODAY = date(2026, 11, 11)


def test_reason_rule(rp):
    rows = [
        _row(rp, "2026-07-10 19:04:09", "2026-07-11", "2026-12-07 14:48", "GT100726A3"),  # future
        _row(rp, "2026-07-03 11:20:43", "2026-07-03", "2026-04-07 18:05", "GT030726A1"),  # before
        _row(rp, "2026-07-10 16:46:55", "2026-07-10", "2026-11-07 13:09", "GT100726A1"),  # late
        _row(rp, "2026-09-10 10:00:00", "2026-09-10", "2026-09-10 17:00:22", "GT100926A1"),  # plausible
        _row(rp, "2026-07-04 16:19:47", "2026-07-04", "2026-07-18 17:46:14", "GT040726A2"),  # day > 12
        _row(rp, "2026-04-01 13:10:57", "2026-01-04", "2026-04-02 11:24:43", "GT010426A1"),  # S swap
    ]
    rep = rp.find_swaps(rows, ["X"], today=TODAY)
    assert [(r[0], r[8], r[7], r[9]) for r in rep] == [
        (8, "future", "12-Jul-2026 14:48:00", "yes"),
        (9, "before_order", "04-Jul-2026 18:05:00", "yes"),
        (10, "late_over_30d", "11-Jul-2026 13:09:00", "no"),
    ]
    assert [r[9] for r in rp.find_swaps(rows, ["X"], today=TODAY, include_late=True)] == \
        ["yes", "yes", "yes"]
    s = rp.find_swaps(rows, ["S"], today=TODAY)
    assert [(r[4], r[7], r[8]) for r in s] == [("S13", "01-Apr-2026 00:00:00", "before_order")]


def test_serial_round_trip(rp):
    dt = datetime(2026, 7, 12, 14, 48)
    assert rp.serial_to_datetime(rp.datetime_to_serial(dt)) == dt
    assert rp.serial_to_datetime("Cancelled") is None and rp.serial_to_datetime("") is None


class _Exec:
    def __init__(self, result):
        self.result = result

    def execute(self):
        return self.result


class _Recorder:
    def __init__(self, values):
        self.values_data, self.calls, self.bodies = values, [], []

    def spreadsheets(self):
        return self

    def values(self):
        return self

    def get(self, **kw):
        assert kw["valueRenderOption"] == "UNFORMATTED_VALUE"
        return _Exec({"values": self.values_data})

    def __getattr__(self, name):
        def call(**kw):
            self.calls.append(name)
            self.bodies.append(kw.get("body"))
            return _Exec({})
        return call


def test_dry_run_writes_nothing(rp, monkeypatch, tmp_path):
    fake = _Recorder([_row(rp, "2026-07-10 19:04:09", None, "2026-12-07 14:48")])
    monkeypatch.setattr(rp, "connect", lambda sa: fake)
    out = tmp_path / "r.csv"
    rp.main(["--csv", str(out)])
    assert fake.calls == []
    assert "X8" in out.read_text(encoding="utf-8")


def test_apply_backs_up_then_writes_serials(rp, monkeypatch, tmp_path):
    fake = _Recorder([_row(rp, "2026-07-10 19:04:09", None, "2026-12-07 14:48")])
    monkeypatch.setattr(rp, "connect", lambda sa: fake)
    rp.main(["--csv", str(tmp_path / "r.csv"), "--apply"])
    assert fake.calls == ["batchUpdate", "update", "batchUpdate"]   # add tab, backup, write
    write = fake.bodies[2]
    assert write["valueInputOption"] == "RAW"
    assert write["data"][0]["range"].endswith("!X8")
    assert rp.serial_to_datetime(write["data"][0]["values"][0][0]) == datetime(2026, 7, 12, 14, 48)
