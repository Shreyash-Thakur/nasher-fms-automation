"""
2026-10-21: Vinculum Customer Code.

GT order form -> Orders "Customer Code" -> GT Pull FMS "Vinculum Customer Code" (a column
appended after A:AW, found by header) -> Full Flow BD, which the auto-punch project reads
positionally and writes itself after a punch.

  1. read_fms_sheet finds the code by header right of AW (AX or further), blank if absent;
     A:AW stay positional
  2. build_full_flow puts the code in BD of every row it appends (incl. expanded rows)
  3. write_full_flow writes A:BD only when BD7 is (or becomes, if blank) the header;
     a foreign BD7 -> A:BC only; the column count never shrinks (it used to be set to 55,
     deleting BD)
  4. existing rows: BD is never written; reading A8:BD keeps every row aligned
"""
import pandas as pd
import pytest

from test_realtime_rows import FakeSheets

CODE = "Vinculum Customer Code"


def _fms_values(ff, header_extra, row_extra, gt="GT211026A1", sku="NMH801912LS"):
    head = list(ff.FMS_SOURCE_COLUMNS)
    row = [""] * len(head)
    row[head.index("Form Received Time")] = "21-Oct-2026 10:00:00"
    row[head.index("GT Number")] = gt
    row[head.index("Vinculum SKU")] = sku
    row[head.index("Order Quantity")] = 3
    row[head.index("Party Name")] = "CROSSWORD BOOKSTORES PRIVATE LIMITED"
    return head + header_extra, row + row_extra


def _read(ff, monkeypatch, values):
    monkeypatch.setattr(ff, "sheets_service", FakeSheets(get_values=values))
    return ff.read_fms_sheet()


def test_code_found_by_header_at_ax(ff, monkeypatch):
    head, row = _fms_values(ff, [CODE], [1004544])
    df = _read(ff, monkeypatch, [head, row])
    assert list(df[CODE]) == ["1004544"]
    assert df.loc[0, "GT Number"] == "GT211026A1"
    assert df.loc[0, "Party Name"] == "CROSSWORD BOOKSTORES PRIVATE LIMITED"


@pytest.mark.parametrize("raw", [1004544, 1004544.0, "1004544.0", " 1004544 "])
def test_code_normalised(ff, monkeypatch, raw):
    head, row = _fms_values(ff, ["Other", "", " vinculum customer code "], ["x", "", raw])
    df = _read(ff, monkeypatch, [head, row])
    assert list(df[CODE]) == ["1004544"]


def test_code_blank_when_header_missing(ff, monkeypatch):
    head, row = _fms_values(ff, ["Something Else"], ["1004544"])
    df = _read(ff, monkeypatch, [head, row])
    assert list(df[CODE]) == [""]


def test_a_aw_stay_positional_and_a_renamed_a_aw_header_still_fails_closed(ff, monkeypatch):
    head, row = _fms_values(ff, ["", CODE], ["", "1004544"])
    row[ff.FMS_SOURCE_COLUMNS.index("AWB")] = "AWB123"
    row[ff.FMS_SOURCE_COLUMNS.index("Invoice No")] = "INV9"     # AW, the last positional one
    df = _read(ff, monkeypatch, [head, row])
    assert (df.loc[0, "AWB"], df.loc[0, "Invoice No"], df.loc[0, CODE]) == ("AWB123", "INV9", "1004544")
    # The code header inside A:AW is never taken as the code (A:AW validation unchanged).
    head2, row2 = _fms_values(ff, [], [])
    head2[ff.FMS_SOURCE_COLUMNS.index("AWB")] = CODE
    with pytest.raises(ValueError, match="Missing FMS source columns"):
        _read(ff, monkeypatch, [head2, row2])


def test_codes_stay_with_their_rows_when_short_and_blank_rows_present(ff, monkeypatch):
    head, r1 = _fms_values(ff, [CODE], ["111"], gt="GT1")
    _, r2 = _fms_values(ff, [CODE], [], gt="GT2")      # Sheets omits the trailing blank cell
    _, r3 = _fms_values(ff, [CODE], ["333"], gt="GT3")
    df = _read(ff, monkeypatch, [head, r1, r2[:5], r3])
    assert list(zip(df["GT Number"], df[CODE])) == [("GT1", "111"), ("GT2", ""), ("GT3", "333")]


def test_build_full_flow_puts_code_in_bd_for_every_expanded_row(ff, make):
    fms = pd.DataFrame([{c: "" for c in ff.FMS_COLUMNS}] * 2, columns=ff.FMS_COLUMNS)
    fms.loc[0, ["GT Number", "Vinculum SKU", CODE]] = ["GT1", "S1", "1004544"]
    fms.loc[1, ["GT Number", "Vinculum SKU", CODE]] = ["GT2", "S2", ""]
    lines = make.sales([
        {"Order No": "N1", "ExtOrderNo": "GT1", "SKU Code": "S1"},
        {"Order No": "N2", "ExtOrderNo": "GT1", "SKU Code": "S1"},
    ])
    assigned = [[lines.iloc[0], lines.iloc[1]], []]
    out = ff.build_full_flow(fms, assigned, {}, {"W1": "Vadape"}, [None, None])
    assert list(out.columns) == ff.OUTPUT_COLUMNS
    assert ff.OUTPUT_COLUMNS.index(CODE) == 55          # BD
    assert list(zip(out["GT Number"], out["Order No"], out[CODE])) == [
        ("GT1", "N1", "1004544"), ("GT1", "N2", "1004544"), ("GT2", "", "")]


def test_drop_and_cancel_keep_bd_with_its_row(ff, make):
    new_df = make.existing([
        {"GT Number": "GT1", "Vinculum SKU": "S1", CODE: "111"},
        {"GT Number": "GT2", "Vinculum SKU": "S2", CODE: "222", "Status": "Cancelled"},
        {"GT Number": "GT3", "Vinculum SKU": "S3", CODE: "333"},
    ])
    kept, dropped = ff.drop_rows_now_in_full_flow(
        new_df, make.existing([{"GT Number": "GT1", "Vinculum SKU": "S1"}]))
    assert dropped == 1
    ff.apply_cancelled_order_no(kept)
    assert list(zip(kept["GT Number"], kept[CODE])) == [("GT2", "222"), ("GT3", "333")]


def _write(ff, monkeypatch, df, bd7, columns=56, rows=100, start_row=8):
    fake = FakeSheets(bd7=bd7)
    monkeypatch.setattr(ff, "sheets_service", fake)
    monkeypatch.setattr(ff, "get_sheet_by_title",
                        lambda *a: {"gridProperties": {"rowCount": rows, "columnCount": columns}})
    ff.write_full_flow(1, df, start_row=start_row)
    return fake


def _grid(fake):
    return fake.batch[0]["body"]["requests"][0]["updateSheetProperties"]["properties"]["gridProperties"]


def test_write_a_bd_when_bd7_is_the_header(ff, make, monkeypatch):
    df = make.existing([{"GT Number": "GT1", "Vinculum SKU": "S", CODE: "1004544"}])
    fake = _write(ff, monkeypatch, df, bd7=CODE)
    assert fake.header_writes == []
    assert fake.updates[0]["range"] == "'Full Flow FMS'!A8:BD8"
    assert fake.updates[0]["body"]["values"][0][55] == "1004544"


def test_blank_bd7_gets_the_header_then_a_bd(ff, make, monkeypatch):
    df = make.existing([{"GT Number": "GT1", "Vinculum SKU": "S", CODE: "1004544"}])
    fake = _write(ff, monkeypatch, df, bd7="")
    assert [w["body"]["values"] for w in fake.header_writes] == [[[CODE]]]
    assert fake.header_writes[0]["range"] == "'Full Flow FMS'!BD7"
    assert fake.updates[0]["range"] == "'Full Flow FMS'!A8:BD8"


def test_foreign_bd7_writes_a_bc_only(ff, make, monkeypatch, capsys):
    df = make.existing([{"GT Number": "GT1", "Vinculum SKU": "S", CODE: "1004544"}])
    fake = _write(ff, monkeypatch, df, bd7="Remarks")
    assert fake.header_writes == []
    assert fake.updates[0]["range"] == "'Full Flow FMS'!A8:BC8"
    assert len(fake.updates[0]["body"]["values"][0]) == 55
    assert "1004544" not in fake.updates[0]["body"]["values"][0]
    assert "WARNING" in capsys.readouterr().out


@pytest.mark.parametrize("current,bd7,expected", [(60, CODE, 60), (55, "", 56), (57, "Remarks", 57)])
def test_column_count_never_shrinks(ff, make, monkeypatch, current, bd7, expected):
    df = make.existing([{"GT Number": "GT1", "Vinculum SKU": "S"}])
    fake = _write(ff, monkeypatch, df, bd7=bd7, columns=current)
    assert _grid(fake)["columnCount"] == expected
    assert ff.grid_column_count(60, 56) == 60 and ff.grid_column_count(55, 56) == 56


def test_appended_start_row_unaffected(ff, make, monkeypatch):
    df = make.existing([{"GT Number": "GT1", "Vinculum SKU": "S"},
                        {"GT Number": "GT2", "Vinculum SKU": "S"}])
    fake = _write(ff, monkeypatch, df, bd7=CODE, start_row=8300, rows=9000)
    assert fake.updates[0]["range"] == "'Full Flow FMS'!A8300:BD8301"
    assert _grid(fake)["rowCount"] == 9000


def test_nothing_to_write_touches_nothing(ff, make, monkeypatch):
    fake = _write(ff, monkeypatch, make.existing([]), bd7="")
    assert fake.header_writes == [] and fake.updates == [] and fake.batch == []


def test_existing_rows_never_get_a_bd_write(ff, make, monkeypatch):
    rows = [{"GT Number": "GT1", "Vinculum SKU": "S1", CODE: ""},
            {"GT Number": "GT2", "Vinculum SKU": "S2", CODE: "999"}]
    sales = [{"Order No": "N1", "ExtOrderNo": "GT1", "SKU Code": "S1", "Invoice No": "I1"},
             {"Order No": "N2", "ExtOrderNo": "GT2", "SKU Code": "S2", "Invoice No": "I2"}]
    writes, _ = make.plan(rows, sales)
    assert writes and not any(c == CODE for _, c, _, _ in writes)
    # guard: even if a future fill rule targeted BD, it is filtered out
    monkeypatch.setattr(ff, "BACKFILL_FIELDS",
                        ff.BACKFILL_FIELDS + [(CODE, lambda s, wh: "1234567")])
    writes, _ = make.plan(rows, sales)
    assert not any(c == CODE for _, c, _, _ in writes)


def test_read_existing_full_flow_reads_bd_and_keeps_alignment(ff, monkeypatch):
    n = len(ff.OUTPUT_COLUMNS)
    r1 = ["x"] * n
    r1[2], r1[55] = "GT1", "111"
    r2 = ["x"] * 3
    r2[2] = "GT2"                                 # short row: BD omitted by the API
    r3 = [""] * n
    r3[2], r3[55] = "GT3", "333"
    seen = {}

    class Fake(FakeSheets):
        def get(self, **kw):
            seen["range"] = kw["range"]
            return super().get(**kw)

    monkeypatch.setattr(ff, "sheets_service", Fake(get_values=[r1, r2, r3]))
    df = ff.read_existing_full_flow()
    assert seen["range"] == "'Full Flow FMS'!A8:BD"
    assert list(zip(df["GT Number"], df[CODE])) == [("GT1", "111"), ("GT2", ""), ("GT3", "333")]


def test_clear_full_flow_data_is_gone(ff):
    assert not hasattr(ff, "clear_full_flow_data")


def test_grid_expanded_before_bd7_is_read(ff, make, monkeypatch):
    """A 55-column grid has no BD7: the grid grows to 56 FIRST, then BD7 is read/set
    (an A1 range outside the grid is refused by the Sheets API)."""
    order = []

    class OrderedFake(FakeSheets):
        def get(self, **kw):
            order.append(("get", kw.get("range", "")))
            return super().get(**kw)

        def batchUpdate(self, **kw):
            order.append(("grid", ""))
            return super().batchUpdate(**kw)

    fake = OrderedFake(bd7="")
    monkeypatch.setattr(ff, "sheets_service", fake)
    monkeypatch.setattr(ff, "get_sheet_by_title",
                        lambda *a: {"gridProperties": {"rowCount": 100, "columnCount": 55}})
    ff.write_full_flow(1, make.existing([{"GT Number": "GT1", "Vinculum SKU": "S", CODE: "1"}]),
                       start_row=8)
    kinds = [k for k, r in order if k == "grid" or r.endswith("!BD7")]
    assert kinds[:2] == ["grid", "get"]
    assert _grid(fake)["columnCount"] == 56
    assert fake.updates[0]["range"] == "'Full Flow FMS'!A8:BD8"


def test_read_ranges_never_pass_the_grid(ff):
    assert ff.fms_read_last_column(49) == "AW"
    assert ff.fms_read_last_column(52) == "AZ"
    assert ff.fms_read_last_column(500) == "CZ"
    assert ff.fms_read_last_column(0) == "CZ"          # metadata unknown
    assert ff.full_flow_read_last_column(55) == "BC"
    assert ff.full_flow_read_last_column(56) == "BD"
    assert ff.full_flow_read_last_column(0) == "BD"
