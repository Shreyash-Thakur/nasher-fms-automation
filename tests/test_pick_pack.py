"""
Pick Pack -> Full Flow V (Warehouse Team) / X (Actual.1)  —  2026-10-10.

Since May 2026 the daily Pick Pack reports land in the Drive folder as native
Google Sheets (no ".csv" in the name), so the CSV-only reader saw April only and
V/X stayed blank from May onward. Contract:
  - every Pick Pack file is read: CSV (download), Google Sheet (export as xlsx),
    Excel (download); other file types are ignored
  - the Drive listing is paginated
  - an unreadable file / a file without the required columns is skipped with a
    warning, never fails the night
  - Pick Date formats seen in the folder all parse to the right day (ISO dates are
    NOT read day-first)
  - V/X are filled on already-matched rows (blank cells only, closed rows untouched)

Run: python -m pytest tests -q
"""
import io
from datetime import datetime

import openpyxl
import pandas as pd
import pytest

HEADER = ["Order No", "Extern Order No", "Order Date", "SKU Code", "Pick Date",
          "Pick User", "Qty"]


def _csv_bytes(rows, header=HEADER, bom=True):
    text = ",".join(header) + "\n" + "\n".join(",".join(map(str, r)) for r in rows) + "\n"
    return (("﻿" if bom else "") + text).encode("utf-8")


def _xlsx_bytes(tabs):
    """tabs: {title: [header, row, row...]} -> xlsx bytes (cells keep their types)."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for title, rows in tabs.items():
        ws = wb.create_sheet(title)
        for r in rows:
            ws.append(list(r))
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


class _FakeList:
    def __init__(self, pages, calls):
        self.pages, self.calls = pages, calls

    def list(self, **kwargs):
        self.calls.append(kwargs.get("pageToken"))
        page = self.pages[kwargs.get("pageToken")]
        return type("Req", (), {"execute": lambda _self: page})()


class _FakeDrive:
    def __init__(self, pages):
        self.calls = []
        self._files = _FakeList(pages, self.calls)

    def files(self):
        return self._files


@pytest.fixture
def drive(ff, monkeypatch):
    """Fake Drive folder: listing in two pages + per-file bytes by id."""
    def install(files, contents, page_size=2):
        pages, token = {}, None
        for start in range(0, len(files), page_size):
            nxt = f"p{start + page_size}" if start + page_size < len(files) else None
            pages[token] = {"files": files[start:start + page_size], "nextPageToken": nxt}
            token = nxt
        fake = _FakeDrive(pages)
        monkeypatch.setattr(ff, "drive_service", fake)
        exported = []

        def download(file_id):
            data = contents[file_id]
            if isinstance(data, Exception):
                raise data
            return io.BytesIO(data)

        def export(file_id, mime_type):
            exported.append((file_id, mime_type))
            return download(file_id)

        monkeypatch.setattr(ff, "download_drive_file", download)
        monkeypatch.setattr(ff, "export_drive_file", export)
        return fake, exported
    return install


GSHEET = "application/vnd.google-apps.spreadsheet"
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def test_all_formats_read_with_pagination(ff, drive):
    files = [
        {"id": "c1", "name": "PickPackReport_2026-04-01_to_2026-04-03.csv", "mimeType": "text/csv"},
        {"id": "g1", "name": "PickPackReport_08.10.2026", "mimeType": GSHEET},
        {"id": "x1", "name": "PickPackReport_09.10.2026.xlsx", "mimeType": XLSX},
        {"id": "p1", "name": "notes.pdf", "mimeType": "application/pdf"},
    ]
    contents = {
        "c1": _csv_bytes([["O1", "E1", "01/04/2026 10:00 00", "S1", "02/04/2026 10:37 22", "Bappa", 1]]),
        # Google Sheet export: header cell carried a literal BOM + quotes from the CSV import,
        # Pick Date kept as text in the dd/mm/yyyy hh:mm ss form.
        "g1": _xlsx_bytes({"PickPackReport_08.10.2026": [
            ['﻿"Order No"'] + HEADER[1:],
            ["O2", "E2", "06/10/2026 04:03 50", "S2", "08/10/2026 10:37 22", "Bidyut", 1]]}),
        # Excel with a real datetime cell, on the SECOND tab (first tab is a cover sheet).
        "x1": _xlsx_bytes({"Cover": [["Report"], ["generated"]],
                           "Data": [HEADER, ["O3", "E3", "x", "S3",
                                             datetime(2026, 5, 2, 9, 34, 53), "Rahul", 1]]}),
    }
    fake, exported = drive(files, contents, page_size=2)

    df = ff.read_pick_pack_files()

    assert fake.calls == [None, "p2"]                       # both pages listed
    assert exported == [("g1", XLSX)]                       # Sheet exported, not downloaded
    assert sorted(df["Order No"]) == ["O1", "O2", "O3"]     # pdf ignored, all three read
    lookup = ff.build_pick_pack_lookup(df.copy())
    assert lookup["O1"]["Pick Date"] == pd.Timestamp("2026-04-02 10:37:22")
    assert lookup["O2"]["Pick Date"] == pd.Timestamp("2026-10-08 10:37:22")
    assert lookup["O2"]["Pick User"] == "Bidyut"
    assert lookup["O3"]["Pick Date"] == pd.Timestamp("2026-05-02 09:34:53")
    assert lookup["O3"]["Pick User"] == "Rahul"


def test_bad_file_skipped_not_fatal(ff, drive, capsys):
    files = [
        {"id": "g1", "name": "PickPackReport_07.10.2026", "mimeType": GSHEET},
        {"id": "g2", "name": "PickPackReport_06.10.2026", "mimeType": GSHEET},
        {"id": "g3", "name": "Random", "mimeType": GSHEET},
    ]
    contents = {
        "g1": PermissionError("403 insufficientFilePermissions"),
        "g2": _xlsx_bytes({"Sheet1": [HEADER, ["O9", "E", "x", "S", "06/10/2026 11:00 05", "Bidyut", 1]]}),
        "g3": _xlsx_bytes({"Sheet1": [["foo", "bar"], [1, 2]]}),       # no Pick Pack columns
    }
    drive(files, contents)

    df = ff.read_pick_pack_files()

    assert list(df["Order No"]) == ["O9"]
    out = capsys.readouterr().out
    assert "PickPackReport_07.10.2026" in out and "403 insufficientFilePermissions" in out
    assert "Random" in out and "missing" in out.lower()


def test_nothing_readable_returns_empty_frame(ff, drive):
    drive([{"id": "g1", "name": "PickPackReport_07.10.2026", "mimeType": GSHEET}],
          {"g1": RuntimeError("boom")})
    df = ff.read_pick_pack_files()
    assert df.empty and set(ff.PICK_PACK_REQUIRED_COLUMNS) <= set(df.columns)


@pytest.mark.parametrize("raw, expected", [
    ("08/10/2026 10:37 22", "2026-10-08 10:37:22"),     # Vinculum CSV / Sheet text
    ("08/10/2026 10:37:22", "2026-10-08 10:37:22"),
    ("2026-05-02 9:34:53", "2026-05-02 09:34:53"),      # ISO (May/June sheets): NOT day-first
    ("2026-10-08 10:37:22", "2026-10-08 10:37:22"),
    ("2026-10-08", "2026-10-08 00:00:00"),
    ("08/10/2026", "2026-10-08 00:00:00"),
    ("46303.4425", "2026-10-08 10:37:12"),              # serial (number cell exported as text)
])
def test_pick_date_formats(ff, raw, expected):
    got = ff.parse_pick_date(raw)
    assert pd.Timestamp(got).round("s") == pd.Timestamp(expected), (raw, got)


@pytest.mark.parametrize("raw", ["", "abc", "31/02/2026 10:00 00", "2026-13-45 10:00:00", None])
def test_bad_pick_date_is_blank_not_fatal(ff, raw):
    assert pd.isna(ff.parse_pick_date(raw))


def test_one_bad_date_does_not_abort_lookup(ff):
    df = pd.DataFrame({"Order No": ["O1", "O1", "O2"],
                       "Pick Date": ["garbage", "08/10/2026 10:37 22", "2026-13-45"],
                       "Pick User": ["X", "Bidyut", "Y"]})
    lookup = ff.build_pick_pack_lookup(df)
    assert lookup["O1"] == {"Pick Date": pd.Timestamp("2026-10-08 10:37:22"), "Pick User": "Bidyut"}
    assert lookup["O2"]["Pick Date"] is None and lookup["O2"]["Pick User"] == "Y"


PICK = {"N1": {"Pick Date": pd.Timestamp("2026-06-03 11:00"), "Pick User": "Bidyut"}}


def test_refresh_fills_v_x_on_matched_row(make):
    # Fully punched May row: everything filled except V/X.
    full = {"GT Number": "GT1", "Vinculum SKU": "S", "Order No": "N1", "Team": "T",
            "Actual": "x", "Invoice No": "I", "Invoice Amount": 1, "Warehouse": "W",
            "Dispatched Qty": 5}
    writes, stats = make.plan([full], [{"Order No": "N1", "ExtOrderNo": "GT1", "SKU Code": "S"}],
                              pick_pack=PICK)
    c = make.cells(writes)
    assert set(c) == {(8, "Warehouse Team"), (8, "Actual.1")}
    assert c[(8, "Warehouse Team")] == "Bidyut"
    assert c[(8, "Actual.1")] == datetime(2026, 6, 3, 11, 0)
    assert stats["refreshed_rows"] == 1


def test_v_x_never_overwritten(make):
    row = {"GT Number": "GT1", "Vinculum SKU": "S", "Order No": "N1",
           "Warehouse Team": "Manual", "Actual.1": "01-Jun-2026"}
    writes, _ = make.plan([row], [{"Order No": "N1", "ExtOrderNo": "GT1", "SKU Code": "S"}],
                          pick_pack=PICK)
    c = make.cells(writes)
    assert (8, "Warehouse Team") not in c and (8, "Actual.1") not in c


def test_v_x_filled_when_order_no_not_in_vinculum(make):
    # P filled by hand, line not in Order Enquiry: Pick Pack is keyed by Order No, so V/X
    # still come from the row's own P (and nothing else is written).
    writes, stats = make.plan([{"GT Number": "GT1", "Vinculum SKU": "S", "Order No": "N1"}],
                              [{"Order No": "N9", "ExtOrderNo": "GT9", "SKU Code": "S"}],
                              pick_pack=PICK)
    c = make.cells(writes)
    assert set(c) == {(8, "Warehouse Team"), (8, "Actual.1")}
    assert stats["order_no_without_line"] == 1


def test_closed_row_v_x_untouched(make):
    writes, _ = make.plan(
        [{"GT Number": "GT1", "Vinculum SKU": "S", "Order No": "N1", "Status": "Cancelled"},
         {"GT Number": "GT2", "Vinculum SKU": "S", "Order No": "N1", "Invoice No": "Virtual Invoice"}],
        [{"Order No": "N1", "ExtOrderNo": "GT1", "SKU Code": "S"}], pick_pack=PICK)
    assert writes == []


def test_projected_fill_rates(ff, make, capsys):
    existing = make.existing([
        {"GT Number": "GT1", "Vinculum SKU": "S", "Order No": "N1",
         "Form Received Time": "03/06/2026 10:00:00"}])
    ff._print_fill_rates(existing, writes=[(8, "Warehouse Team", "Bidyut", "refresh")])
    out = capsys.readouterr().out
    assert "projected" in out.lower()
