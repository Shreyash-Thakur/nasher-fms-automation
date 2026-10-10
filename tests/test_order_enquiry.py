"""
Daily Vinculum Order Enquiry -> Full Flow P (Order No) / S (Order Date)  —  2026-10-10.

P/Q/S used to come only from the "Vinculum Sales Master" (Sales Report = INVOICED lines),
so a punched-but-not-invoiced order left P/Q/S blank. The daily OrderEnquiry_DD-MM-YYYY
exports hold every punched order, but ORDER-LEVEL only (no SKU, no Created By). Contract:
  - files named OrderEnquiry_* are read (Google Sheet / CSV / xlsx), others ignored,
    a bad file is skipped with a warning; an order in several files: latest file wins
  - a blank-P open row that the Sales Master did not match gets P + S from Order Enquiry
    ONLY when its GT (exact, else R/T/F-stripped base; never A1<->A10) has exactly ONE
    live Vinculum order (Order Enquiry + Sales Master + Full Flow P), and that order is
    not invoiced (an invoiced order is the Sales Master's, which is line-level)
  - cancelled orders are ignored; Sales Master wins when both have a value
  - blank cells only, closed rows untouched, idempotent
"""
import io

import openpyxl
import pandas as pd
import pytest

GSHEET = "application/vnd.google-apps.spreadsheet"
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
OE_HEADER = ["Order No", "Extern OrderNo", "Order Date", "Order Type", "Status", "Order Amount"]


def _xlsx(rows, header=OE_HEADER):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Export"
    ws.append(header)
    for r in rows:
        ws.append(list(r))
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


@pytest.fixture
def drive(ff, monkeypatch):
    def install(files, contents):
        class _Req:
            def __init__(self, page):
                self.page = page

            def execute(self):
                return self.page

        class _Files:
            def list(self, **kwargs):
                return _Req({"files": files, "nextPageToken": None})

        class _Drive:
            def files(self):
                return _Files()

        monkeypatch.setattr(ff, "drive_service", _Drive())

        def download(file_id):
            data = contents[file_id]
            if isinstance(data, Exception):
                raise data
            return io.BytesIO(data)

        monkeypatch.setattr(ff, "download_drive_file", download)
        monkeypatch.setattr(ff, "export_drive_file", lambda fid, mime: download(fid))
    return install


# ---------------------------------------------------------------- reader

def test_reader_reads_only_order_enquiry_files_latest_wins(ff, drive, capsys):
    files = [  # listing order = modifiedTime desc
        {"id": "a", "name": "OrderEnquiry_09-10-2026", "mimeType": GSHEET},
        {"id": "b", "name": "OrderEnquiry_08-10-2026", "mimeType": GSHEET},
        {"id": "c", "name": "OrderEnquiry_07-10-2026", "mimeType": GSHEET},
        {"id": "m", "name": "July -August 2026", "mimeType": GSHEET},          # not a daily file
    ]
    contents = {
        "a": _xlsx([["N2", "GT081026A2", "08/10/2026 12:00 AM", "Prepaid", "Cancelled", "1"]]),
        "b": _xlsx([["N1", "GT081026A1", "08/10/2026 06:27 PM", "Prepaid", "Confirmed", "1"],
                    ["N2", "GT081026A2", "08/10/2026 12:00 AM", "Prepaid", "Allocated", "1"]]),
        "c": PermissionError("403 insufficientFilePermissions"),
        "m": AssertionError("must not be read"),
    }
    drive(files, contents)

    raw = ff.read_order_enquiry_daily_files()
    out = capsys.readouterr().out
    assert "OrderEnquiry_07-10-2026" in out and "403" in out           # skipped, named
    assert sorted(raw["Order No"]) == ["N1", "N2", "N2"]
    assert "ExtOrderNo" in raw.columns                                   # "Extern OrderNo" aliased

    orders = ff.prepare_oe_orders(raw)
    assert len(orders) == 2
    n2 = orders[orders["_ORDER"] == "N2"].iloc[0]
    assert n2["Status"] == "Cancelled" and bool(n2["_CANCELLED"])        # latest file's status
    n1 = orders[orders["_ORDER"] == "N1"].iloc[0]
    assert n1["Order Date"] == pd.Timestamp("2026-10-08 18:27")


def test_reader_nothing_readable_is_not_fatal(ff, drive):
    drive([{"id": "a", "name": "OrderEnquiry_09-10-2026", "mimeType": GSHEET}],
          {"a": RuntimeError("boom")})
    raw = ff.read_order_enquiry_daily_files()
    assert raw.empty
    assert ff.prepare_oe_orders(raw).empty


@pytest.mark.parametrize("raw, expected", [
    ("06/10/2026 11:58 PM", "2026-10-06 23:58:00"),
    ("06/10/2026 12:00 AM", "2026-10-06 00:00:00"),
    ("06/10/2026 12:30 PM", "2026-10-06 12:30:00"),
    ("13/10/2026 09:05 am", "2026-10-13 09:05:00"),
    ("2026-10-06 23:58:00", "2026-10-06 23:58:00"),     # real date cell exported to xlsx
])
def test_oe_date_formats(ff, raw, expected):
    assert ff.parse_oe_date(raw) == pd.Timestamp(expected)


def test_oe_bad_date_blank(ff):
    assert pd.isna(ff.parse_oe_date("31/02/2026 10:00 AM"))
    assert pd.isna(ff.parse_oe_date(""))


# ---------------------------------------------------------------- fill

def test_oe_fills_p_s_for_uninvoiced_order(make):
    writes, stats = make.plan(
        [{"GT Number": "GT061026A4", "Vinculum SKU": "S1"},
         {"GT Number": "GT061026A4", "Vinculum SKU": "S2"}],
        [],                                                     # not invoiced: no Sales line
        oe_rows=[{"Order No": "NM11832", "ExtOrderNo": "GT061026A4",
                  "Order Date": "06/10/2026 06:27 PM", "Status": "Shipped complete"}])
    c = make.cells(writes)
    for r in (8, 9):                                            # one order, every line of the GT
        assert c[(r, "Order No")] == "NM11832"
        assert c[(r, "Actual")] == pd.Timestamp("2026-10-06 18:27")
        assert (r, "Team") not in c                             # export has no Created By
    assert {w[3] for w in writes} == {"oe_fill"}
    assert stats["oe_filled_rows"] == 2 and stats["filled_rows"] == 0
    for col in ("Invoice No", "Invoice Amount", "Warehouse", "Dispatched Qty"):
        assert (8, col) not in c                                # invoice fields: Sales Master only


def test_oe_team_when_export_has_created_by(make):
    writes, _ = make.plan([{"GT Number": "GT1", "Vinculum SKU": "S"}], [],
                          oe_rows=[{"Order No": "N1", "ExtOrderNo": "GT1", "Created by": "Gauri"}])
    assert make.cells(writes)[(8, "Team")] == "Gauri"


def test_oe_fill_brings_pick_pack_v_x(make):
    writes, _ = make.plan(
        [{"GT Number": "GT1", "Vinculum SKU": "S"}], [],
        pick_pack={"N1": {"Pick Date": pd.Timestamp("2026-10-07 10:00"), "Pick User": "Bidyut"}},
        oe_rows=[{"Order No": "N1", "ExtOrderNo": "GT1"}])
    c = make.cells(writes)
    assert c[(8, "Warehouse Team")] == "Bidyut" and (8, "Actual.1") in c


def test_sales_master_preferred(make):
    # Sales Master matches the row -> its line wins (P, Q, S, invoice); OE is not used.
    writes, stats = make.plan(
        [{"GT Number": "GT1", "Vinculum SKU": "S"}],
        [{"Order No": "N1", "ExtOrderNo": "GT1", "SKU Code": "S", "Created by": "Gauri",
          "Order Date": "2026-10-01 10:00:00"}],
        oe_rows=[{"Order No": "N1", "ExtOrderNo": "GT1", "Order Date": "02/10/2026 11:00 AM"}])
    c = make.cells(writes)
    assert c[(8, "Order No")] == "N1" and c[(8, "Team")] == "Gauri"
    assert c[(8, "Actual")] == pd.Timestamp("2026-10-01 10:00")
    assert {w[3] for w in writes} == {"fill"} and stats["oe_filled_rows"] == 0


def test_unique_order_already_invoiced_not_filled(make):
    # GT's only order is invoiced; this SKU is not on it -> Sales Master (line-level) says no.
    writes, stats = make.plan(
        [{"GT Number": "GT1", "Vinculum SKU": "S1"},
         {"GT Number": "GT1", "Vinculum SKU": "OTHER"}],
        [{"Order No": "N1", "ExtOrderNo": "GT1", "SKU Code": "S1"}],
        oe_rows=[{"Order No": "N1", "ExtOrderNo": "GT1"}])
    c = make.cells(writes)
    assert c[(8, "Order No")] == "N1" and (9, "Order No") not in c
    assert stats["oe_outcomes"]["unique_order_invoiced"] == 1


def test_multi_order_gt_ambiguous(make):
    writes, stats = make.plan(
        [{"GT Number": "GT041026A2", "Vinculum SKU": "S1"}], [],
        oe_rows=[{"Order No": "M1", "ExtOrderNo": "GT041026A2"},
                 {"Order No": "N2", "ExtOrderNo": "GT041026A2"}])
    assert writes == []
    assert stats["oe_outcomes"]["ambiguous"] == 1


def test_invoiced_sibling_order_makes_gt_ambiguous(make):
    # OE has one un-invoiced order, Sales Master holds another order of the same GT.
    writes, stats = make.plan(
        [{"GT Number": "GT1", "Vinculum SKU": "S1"}, {"GT Number": "GT1", "Vinculum SKU": "S2"}],
        [{"Order No": "N1", "ExtOrderNo": "GT1", "SKU Code": "S1"}],
        oe_rows=[{"Order No": "N2", "ExtOrderNo": "GT1"}])
    c = make.cells(writes)
    assert c[(8, "Order No")] == "N1" and (9, "Order No") not in c


def test_full_flow_p_counts_as_an_order(make):
    # Another row of the GT already carries a (manual) Order No -> two orders -> no guess.
    writes, _ = make.plan(
        [{"GT Number": "GT1", "Vinculum SKU": "S1", "Order No": "MANUAL1"},
         {"GT Number": "GT1", "Vinculum SKU": "S2"}], [],
        oe_rows=[{"Order No": "N2", "ExtOrderNo": "GT1"}])
    assert (9, "Order No") not in make.cells(writes)


def test_cancelled_orders_ignored(make):
    writes, stats = make.plan(
        [{"GT Number": "GT1", "Vinculum SKU": "S"}, {"GT Number": "GT2", "Vinculum SKU": "S"}], [],
        oe_rows=[{"Order No": "C1", "ExtOrderNo": "GT1", "Status": "Cancelled"},
                 {"Order No": "C2", "ExtOrderNo": "GT2", "Status": "Cancelled"},
                 {"Order No": "N2", "ExtOrderNo": "GT2", "Status": "Allocated"}])
    c = make.cells(writes)
    assert (8, "Order No") not in c                       # only a cancelled order
    assert c[(9, "Order No")] == "N2"                     # cancelled sibling does not count
    assert stats["oe_outcomes"]["only_cancelled"] == 1


def test_suffix_exact_first_then_base(make):
    writes, _ = make.plan(
        [{"GT Number": "GT081026A2R", "Vinculum SKU": "S"},     # exact R order is unique
         {"GT Number": "GT091026A5T", "Vinculum SKU": "S"}],    # no exact -> base, unique
        [],
        oe_rows=[{"Order No": "NR", "ExtOrderNo": "GT081026A2R"},
                 {"Order No": "NT1", "ExtOrderNo": "GT081026A2T"},
                 {"Order No": "NT2", "ExtOrderNo": "GT081026A2T"},
                 {"Order No": "N5", "ExtOrderNo": "GT091026A5"}])
    c = make.cells(writes)
    assert c[(8, "Order No")] == "NR" and c[(9, "Order No")] == "N5"


def test_a1_never_takes_a10(make):
    writes, _ = make.plan(
        [{"GT Number": "GT081026A1", "Vinculum SKU": "S"},
         {"GT Number": "GT081026A10R", "Vinculum SKU": "S"}], [],
        oe_rows=[{"Order No": "N10", "ExtOrderNo": "GT081026A10"}])
    c = make.cells(writes)
    assert (8, "Order No") not in c and c[(9, "Order No")] == "N10"


def test_never_overwrite_and_closed_untouched(make):
    writes, _ = make.plan(
        [{"GT Number": "GT1", "Vinculum SKU": "S", "Actual": "01-Oct-2026 10:00:00"},  # S kept
         {"GT Number": "GT2", "Vinculum SKU": "S", "Status": "Cancelled"},             # closed
         {"GT Number": "GT3", "Vinculum SKU": "S", "Order No": "N3"}],                 # P kept
        [],
        oe_rows=[{"Order No": "N1", "ExtOrderNo": "GT1"}, {"Order No": "N2", "ExtOrderNo": "GT2"},
                 {"Order No": "N3", "ExtOrderNo": "GT3"}])
    c = make.cells(writes)
    assert c == {(8, "Order No"): "N1"}


def test_idempotent(ff, make):
    existing = make.existing([{"GT Number": "GT1", "Vinculum SKU": "S1"},
                              {"GT Number": "GT1", "Vinculum SKU": "S2"}])
    oe = make.oe([{"Order No": "N1", "ExtOrderNo": "GT1"}])
    sales = make.sales([])
    writes, *_ = ff.plan_backfill(existing, sales, {}, {}, oe_orders=oe)
    assert len(writes) == 4
    after = existing.astype(object)
    for sheet_row, column, value, _ in writes:
        after.at[sheet_row - 8, column] = value
    again, stats, *_ = ff.plan_backfill(after, sales, {}, {}, oe_orders=oe)
    assert again == [] and stats["oe_filled_rows"] == 0


def test_without_order_enquiry_behaviour_unchanged(make):
    writes, stats = make.plan([{"GT Number": "GT1", "Vinculum SKU": "S"}], [])
    assert writes == [] and stats["still_unmatched"] == 1


# ---------------------------------------------------------------- append path + report

def test_new_row_filled_from_oe(ff, make):
    fms = pd.DataFrame([{**{c: "" for c in ff.FMS_COLUMNS},
                         "GT Number": "GT091026A3", "Vinculum SKU": "S"}])
    existing = make.existing([])
    oe_index = ff.OrderEnquiryIndex(make.oe([{"Order No": "N9", "ExtOrderNo": "GT091026A3",
                                              "Order Date": "09/10/2026 04:00 PM"}]),
                                    make.sales([]), existing)
    rec = oe_index.records[oe_index.match("GT091026A3", "GT091026A3")[0]]
    built = ff.build_full_flow(fms, [[]], {}, {}, oe_records=[rec])
    assert built.iloc[0]["Order No"] == "N9"
    assert built.iloc[0]["Actual"] == pd.Timestamp("2026-10-09 16:00")


def test_unmatched_classification(ff, make, tmp_path):
    existing = make.existing([
        {"GT Number": "GT1", "Vinculum SKU": "S"},     # GT has 2 live orders -> punched, ambiguous
        {"GT Number": "GT2", "Vinculum SKU": "S"},     # only cancelled
        {"GT Number": "GT3", "Vinculum SKU": "X"},     # invoiced, SKU differs
        {"GT Number": "GT4", "Vinculum SKU": "S"},     # nowhere
        {"GT Number": "GT5", "Vinculum SKU": "S"},     # filled from OE -> not listed
    ])
    sales = make.sales([{"Order No": "N3", "ExtOrderNo": "GT3", "SKU Code": "S"}])
    oe = make.oe([{"Order No": "A", "ExtOrderNo": "GT1"}, {"Order No": "B", "ExtOrderNo": "GT1"},
                  {"Order No": "C", "ExtOrderNo": "GT2", "Status": "Cancelled"},
                  {"Order No": "E", "ExtOrderNo": "GT5"}])
    writes, stats, slots, _ = ff.plan_backfill(existing, sales, {}, {}, oe_orders=oe)
    path = tmp_path / "u.csv"
    ff._print_unmatched_diagnostics(existing, slots, sales, path=str(path),
                                    oe_index=ff.OrderEnquiryIndex(oe, sales, existing))
    got = pd.read_csv(path)
    assert dict(zip(got["gt_number"], got["reason"])) == {
        "GT1": "punched_not_invoiced_multi_order", "GT2": "vinculum_order_cancelled",
        "GT3": "sku_differs", "GT4": "gt_not_in_vinculum"}
