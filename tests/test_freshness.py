"""
Full Flow data freshness  —  2026-10-13.
Record: "Errors and their fixes/2026-10-13-full-flow-data-older-than-24h.md".

Contract:
  - Drive export/download is retried with backoff on a rate limit (403
    userRateLimitExceeded / rateLimitExceeded), 429 and 5xx; a 403 for a missing
    permission is not retried
  - a file that still cannot be read is recorded (RUN_HEALTH) and the run exits 1 when
    that file is dated within the last 3 days; older skips and a stale source (newest
    daily file > 2 days old) are warnings; a job summary is written
  - the daily Vinculum Sales Report files (DD.MM.YYYY, last 7 days) are read directly so
    an invoice does not wait for the Apps Script consolidation into the Sales Master;
    only lines the Sales Master lacks (Order No + GT + SKU + Invoice No) are used, and a
    second invoice of the same line is not mistaken for a duplicate
"""
import io
from datetime import date

import openpyxl
import pandas as pd
import pytest

GSHEET = "application/vnd.google-apps.spreadsheet"
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
PP_HEADER = ["Order No", "Ext Order No", "Order Date", "Shipment No", "Pick Date",
             "Pick User", "Qty"]


class FakeHttpError(Exception):
    def __init__(self, status, text):
        super().__init__(text)
        self.resp = type("Resp", (), {"status": status})()


RATE_LIMIT = ('<HttpError 403 when requesting https://www.googleapis.com/drive/v3/files/x/'
              'export returned "User rate limit exceeded.". Details: "[{\'message\': \'User '
              'rate limit exceeded.\', \'domain\': \'usageLimits\', \'reason\': '
              '\'userRateLimitExceeded\'}]">')


@pytest.fixture
def health(ff, monkeypatch):
    h = ff.RunHealth()
    monkeypatch.setattr(ff, "RUN_HEALTH", h)
    return h


def _xlsx(rows):
    wb = openpyxl.Workbook()
    for r in rows:
        wb.active.append(list(r))
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------- retry / backoff

@pytest.mark.parametrize("exc, retry", [
    (FakeHttpError(403, RATE_LIMIT), True),
    (FakeHttpError(403, "rateLimitExceeded"), True),
    (FakeHttpError(429, "Too Many Requests"), True),
    (FakeHttpError(503, "backend"), True),
    (FakeHttpError(403, "insufficientFilePermissions"), False),
    (FakeHttpError(404, "File not found"), False),
    (PermissionError("403 insufficientFilePermissions"), False),
    (ValueError("bad xlsx"), False),
])
def test_retryable_errors(ff, exc, retry):
    assert ff.is_retryable_drive_error(exc) is retry


def test_backoff_retries_rate_limit_then_succeeds(ff):
    calls, sleeps = [], []

    def call():
        calls.append(1)
        if len(calls) < 3:
            raise FakeHttpError(403, RATE_LIMIT)
        return "data"

    assert ff.with_drive_backoff(call, "export x", delays=(2, 4, 8), sleep=sleeps.append) == "data"
    assert len(calls) == 3 and len(sleeps) == 2
    assert 2 <= sleeps[0] < 3 and 4 <= sleeps[1] < 5          # exponential + jitter


def test_backoff_gives_up_and_raises(ff):
    sleeps = []
    with pytest.raises(FakeHttpError):
        ff.with_drive_backoff(lambda: (_ for _ in ()).throw(FakeHttpError(429, "x")),
                              "export x", delays=(1, 1), sleep=sleeps.append)
    assert len(sleeps) == 2


def test_backoff_does_not_retry_permission_error(ff):
    sleeps = []
    with pytest.raises(FakeHttpError):
        ff.with_drive_backoff(lambda: (_ for _ in ()).throw(
            FakeHttpError(403, "insufficientFilePermissions")), "x", sleep=sleeps.append)
    assert sleeps == []


def test_export_drive_file_survives_rate_limit(ff, monkeypatch):
    """The real export path (not monkeypatched away): 2 rate-limit errors, then data."""
    attempts = []

    class Files:
        def export_media(self, fileId, mimeType):
            return ("req", fileId)

    monkeypatch.setattr(ff, "drive_service", type("D", (), {"files": lambda self: Files()})())

    def download(request):
        attempts.append(request)
        if len(attempts) <= 2:
            raise FakeHttpError(403, RATE_LIMIT)
        return io.BytesIO(b"ok")

    monkeypatch.setattr(ff, "_download_request", download)
    monkeypatch.setattr(ff.time, "sleep", lambda s: None)
    assert ff.export_drive_file("f1", XLSX).getvalue() == b"ok"
    assert len(attempts) == 3


# ---------------------------------------------------------------- run health

@pytest.mark.parametrize("name, expected", [
    ("PickPackReport_10.10.2026", date(2026, 10, 10)),
    ("OrderEnquiry_10-10-2026", date(2026, 10, 10)),
    ("10.10.2026", date(2026, 10, 10)),
    ("PickPackReport_2026-04-01_to_2026-04-03.csv", date(2026, 4, 3)),
    ("PickPackReport_28.08.2026_29.08.2026", date(2026, 8, 29)),
    ("Oct-26", None),
    ("_ProcessedFiles", None),
])
def test_file_report_date(ff, name, expected):
    assert ff.file_report_date(name) == expected


def test_recent_skip_fails_run_old_skip_warns(ff, health, tmp_path, capsys):
    today = date(2026, 10, 13)
    health.source("pick_pack", "Pick Pack (V/X)")
    health.found("pick_pack", 3)
    health.read("pick_pack", "PickPackReport_12.10.2026", 900)
    health.skipped("pick_pack", "PickPackReport_11.10.2026", "HttpError 403 userRateLimitExceeded")
    health.skipped("pick_pack", "PickPackReport_2026-06-09_to_2026-06-11", "HttpError 403")
    summary = tmp_path / "summary.md"

    code = health.finish(today=today, summary_path=str(summary))

    assert code == 1
    out = capsys.readouterr().out
    assert "::error::Pick Pack (V/X): file not read, skipped: PickPackReport_11.10.2026" in out
    assert "::warning::Pick Pack (V/X): file not read, skipped: PickPackReport_2026-06-09" in out
    text = summary.read_text(encoding="utf-8")
    assert "| Pick Pack (V/X) | 3 | 1 | 2 | 900 | 12.10.2026 |" in text


def test_only_old_skip_keeps_run_green(ff, health):
    health.source("pick_pack", "Pick Pack (V/X)")
    health.read("pick_pack", "PickPackReport_12.10.2026", 10)
    health.skipped("pick_pack", "PickPackReport_2026-06-09_to_2026-06-11", "403")
    assert health.finish(today=date(2026, 10, 13), summary_path=None) == 0


def test_stale_daily_source_is_a_warning(ff, health, capsys):
    health.source("order_enquiry", "Order Enquiry daily (P/S)")
    health.read("order_enquiry", "OrderEnquiry_10-10-2026", 771)
    health.source("sales_master", "Sales Master", daily=False)
    health.read("sales_master", "Oct-26", 5000)
    errors, warns = health.problems(date(2026, 10, 13))
    assert errors == []
    assert len(warns) == 1 and "10.10.2026 (3 days old)" in warns[0]
    assert health.finish(today=date(2026, 10, 13), summary_path=None) == 0
    assert "::warning::Order Enquiry daily (P/S): newest file is 10.10.2026" in capsys.readouterr().out


def test_pick_pack_reader_records_skips_and_recovers_rate_limit(ff, health, monkeypatch):
    files = [{"id": "a", "name": "PickPackReport_12.10.2026", "mimeType": GSHEET},
             {"id": "b", "name": "PickPackReport_11.10.2026", "mimeType": GSHEET}]
    monkeypatch.setattr(ff, "list_drive_files", lambda folder: files)
    good = _xlsx([PP_HEADER, ["O1", "E", "x", "S", "12/10/2026 10:00 05", "Bidyut", 1]])
    tries = {"a": 0}

    def export(file_id, mime):
        if file_id == "a":           # rate limited once, then fine (retry inside export)
            tries["a"] += 1
            return io.BytesIO(good)
        raise FakeHttpError(403, "insufficientFilePermissions")

    monkeypatch.setattr(ff, "export_drive_file", export)
    df = ff.read_pick_pack_files()
    assert list(df["Order No"]) == ["O1"]
    src = health.sources["pick_pack"]
    assert src["found"] == 2 and src["read"] == 1
    assert [n for n, _ in src["skipped"]] == ["PickPackReport_11.10.2026"]
    assert health.finish(today=date(2026, 10, 13), summary_path=None) == 1


# ---------------------------------------------------------------- daily Sales Report

def test_sales_daily_candidates_window(ff):
    files = [{"name": n, "id": n} for n in
             ["12.10.2026", "06.10.2026", "05.10.2026", "July -August 2026", "13.10.2026 copy",
              "31.02.2026", "11.10.2026"]]
    got = ff.sales_daily_candidates(files, date(2026, 10, 13), lookback_days=7)
    assert [f["name"] for _, f in got] == ["06.10.2026", "11.10.2026", "12.10.2026"]


SALES_COLS = ["Order No", "ExtOrderNo", "Order Date", "SKU Code", "Created by", "Invoice No",
              "Originating warehouse", "Pick Qty", "Total"]


def test_read_sales_daily_files_reads_sales_report_tab(ff, health, monkeypatch):
    files = [{"id": "s12", "name": "12.10.2026", "mimeType": GSHEET},
             {"id": "s01", "name": "01.10.2026", "mimeType": GSHEET},       # outside window
             {"id": "s11", "name": "11.10.2026", "mimeType": GSHEET}]
    monkeypatch.setattr(ff, "list_drive_files", lambda folder: files)

    class Sheets:
        def spreadsheets(self):
            return self

        def get(self, spreadsheetId, fields):
            titles = ["Inventory Aging Report", "Sales Report"]
            return type("R", (), {"execute": lambda _s, **k: {
                "sheets": [{"properties": {"title": t}} for t in titles]}})()

    monkeypatch.setattr(ff, "sheets_service", Sheets())
    reads = []

    def read_tab(spreadsheet_id, tab):
        reads.append((spreadsheet_id, tab))
        if spreadsheet_id == "s11":
            raise FakeHttpError(403, "insufficientFilePermissions")
        row = ["N1", "GT121026A1", "2026-10-12 10:00:00", "SKU1", "Gauri", "27/10/NM7/9",
               "NM7", 2, 1000]
        return pd.DataFrame([row], columns=SALES_COLS)

    monkeypatch.setattr(ff, "read_order_enquiry_tab", read_tab)
    df = ff.read_sales_daily_files(today=date(2026, 10, 13))

    assert ("s12", "Sales Report") == reads[-1] and ("s01", "Sales Report") not in reads
    assert list(df["_TAB"]) == ["daily:12.10.2026"]
    src = health.sources["sales_daily"]
    assert src["read"] == 1 and src["newest"] == date(2026, 10, 12)
    assert [n for n, _ in src["skipped"]] == ["11.10.2026"]


def _sales(rows):
    base = {"Order No": "", "ExtOrderNo": "", "Order Date": "2026-10-12 10:00:00",
            "SKU Code": "", "Created by": "Gauri", "Invoice No": "",
            "Originating warehouse": "W1", "Pick Qty": 1, "Total": 100, "_TAB": "Oct-26"}
    return pd.DataFrame([{**base, **r} for r in rows])


def test_daily_lines_deduplicated_against_sales_master(ff, health):
    df = _sales([
        {"Order No": "N1", "ExtOrderNo": "GT1", "SKU Code": "S1", "Invoice No": "INV1"},
        # same line in the daily file -> dropped
        {"Order No": "N1", "ExtOrderNo": "GT1", "SKU Code": "S1", "Invoice No": "INV1",
         "_TAB": "daily:10.10.2026"},
        # second shipment of the same line, invoiced later -> kept
        {"Order No": "N1", "ExtOrderNo": "GT1", "SKU Code": "S1", "Invoice No": "INV2",
         "_TAB": "daily:12.10.2026"},
        # invoice not consolidated yet -> kept
        {"Order No": "N2", "ExtOrderNo": "GT2", "SKU Code": "S2", "Invoice No": "INV3",
         "_TAB": "daily:12.10.2026"},
        # the same new line in two daily files -> once
        {"Order No": "N2", "ExtOrderNo": "GT2", "SKU Code": "S2", "Invoice No": "INV3",
         "_TAB": "daily:13.10.2026"},
    ])
    lines = ff.prepare_sales_lines(df)
    got = sorted(zip(lines["_ORDER"], lines["Invoice No"]))
    assert got == [("N1", "INV1"), ("N1", "INV2"), ("N2", "INV3")]


def test_cross_tab_dedup_unchanged_for_master_tabs(ff):
    df = _sales([
        {"Order No": "N1", "ExtOrderNo": "GT1", "SKU Code": "S1", "Invoice No": "INV1",
         "_TAB": "Sep-26"},
        {"Order No": "N1", "ExtOrderNo": "GT1", "SKU Code": "S1", "Invoice No": "INV1",
         "_TAB": "Sept-26"},
    ])
    assert len(ff.prepare_sales_lines(df)) == 1


def test_daily_invoice_fills_blank_row_same_night(ff, make):
    """GT invoiced yesterday, file landed today, Sales Master not consolidated yet:
    the blank-P row gets Order No / invoice data tonight."""
    existing = make.existing([{"GT Number": "GT121026A1", "Vinculum SKU": "SKU1",
                               "Form Received Time": "12-Oct-2026 10:00:00"}])
    sales = ff.prepare_sales_lines(_sales([
        {"Order No": "M1", "ExtOrderNo": "GT0001", "SKU Code": "X", "Invoice No": "I0"},
        {"Order No": "N9", "ExtOrderNo": "GT121026A1", "SKU Code": "SKU1",
         "Invoice No": "27/10/NM7/9", "Total": 2500, "_TAB": "daily:12.10.2026"},
    ]))
    writes, _ = ff.plan_backfill(existing, sales, {}, {"W1": "Vadape"})[:2]
    cells = make.cells(writes)
    assert cells[(8, "Order No")] == "N9"
    assert cells[(8, "Invoice No")] == "27/10/NM7/9"
