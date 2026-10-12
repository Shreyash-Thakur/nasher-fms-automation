"""
Read-only report: live Vinculum GT orders Full Flow does not show  —  2026-10-12.

Reported: "There are 88 GT orders missing from the GT Full FMS, and their order status is
not cancelled." Contract of report_vinculum_gt_orders_not_in_full_flow():
  - universe = daily Order Enquiry orders with a GT ExtOrderNo, not Cancelled, plus Sales
    Master (invoiced) GT orders; an order Order Enquiry calls Cancelled is never reported
  - gt_missing: GT (exact, else R/T/F base; never A1<->A10) in neither Full Flow nor GT Pull
  - appending_tonight: GT only in GT Pull FMS (this run appends it)
  - order_not_shown: GT in Full Flow but the Order No is in no row's P
  - an order whose number is in Full Flow P is not reported; non-GT channels are ignored
"""
import pandas as pd


def _fms(gts):
    return pd.DataFrame({"GT Number": list(gts)})


def _report(ff, make, existing_rows, oe_rows, sales_rows=(), fms_gts=(), path=None):
    return ff.report_vinculum_gt_orders_not_in_full_flow(
        make.oe(oe_rows), make.sales(list(sales_rows)), make.existing(existing_rows),
        _fms(fms_gts), path=path)


def _by_order(report):
    return dict(zip(report["order_no"], report["reason"]))


def test_reasons(ff, make):
    existing = [
        {"GT Number": "GT290926A1", "Vinculum SKU": "S1", "Order No": "N1"},
        {"GT Number": "GT240926A4R", "Vinculum SKU": "S2", "Order No": ""},
    ]
    oe = [
        {"Order No": "N1", "ExtOrderNo": "GT290926A1"},                  # shown in P
        {"Order No": "N461759", "ExtOrderNo": "GT290926A1"},             # GT present, order not
        {"Order No": "N461750", "ExtOrderNo": "GT240926A4R"},            # same (suffix GT)
        {"Order No": "N89637", "ExtOrderNo": "GT290926A2", "Status": "Allocated"},  # missing
        {"Order No": "N051624", "ExtOrderNo": "GT290926A4T"},            # missing (T split)
        {"Order No": "N9", "ExtOrderNo": "GT011026A3"},                  # in GT Pull only
        {"Order No": "M1", "ExtOrderNo": "406-1234-5678"},               # Amazon: ignored
        {"Order No": "N5", "ExtOrderNo": "GT290926A9", "Status": "Cancelled"},  # cancelled
    ]
    report = _report(ff, make, existing, oe, fms_gts=["GT290926A1", "GT011026A3"])
    assert _by_order(report) == {
        "N461759": "order_not_shown", "N461750": "order_not_shown",
        "N89637": "gt_missing", "N051624": "gt_missing", "N9": "appending_tonight"}


def test_base_match_is_equality_only(ff, make):
    """GT…A1 in Full Flow must not hide an order of GT…A10 (and vice versa)."""
    existing = [{"GT Number": "GT180926A1", "Vinculum SKU": "S", "Order No": "N1"}]
    report = _report(ff, make, existing, [{"Order No": "N2", "ExtOrderNo": "GT180926A10"}])
    assert _by_order(report) == {"N2": "gt_missing"}


def test_suffix_tolerant_presence(ff, make):
    """Vinculum GT290926A4 (no suffix) vs Full Flow GT290926A4R -> GT present."""
    existing = [{"GT Number": "GT290926A4R", "Vinculum SKU": "S", "Order No": "N7"}]
    report = _report(ff, make, existing, [{"Order No": "N7", "ExtOrderNo": "GT290926A4"},
                                          {"Order No": "N8", "ExtOrderNo": "GT290926A4"}])
    assert _by_order(report) == {"N8": "order_not_shown"}


def test_sales_master_orders_count_unless_cancelled(ff, make):
    sales = [
        {"Order No": "N89578", "ExtOrderNo": "GT180926A20", "SKU Code": "S1",
         "Invoice No": "INV1"},
        {"Order No": "N77", "ExtOrderNo": "GT180926A21", "SKU Code": "S1"},
    ]
    oe = [{"Order No": "N77", "ExtOrderNo": "GT180926A21", "Status": "Cancelled"}]
    report = _report(ff, make, [], oe, sales_rows=sales)
    assert _by_order(report) == {"N89578": "gt_missing"}
    assert report.loc[0, "source"] == "sales_master"


def test_writes_csv_and_changes_nothing(ff, make, tmp_path):
    existing_rows = [{"GT Number": "GT290926A1", "Vinculum SKU": "S1", "Order No": ""}]
    existing = make.existing(existing_rows)
    before = existing.copy()
    path = tmp_path / "reports" / "x.csv"
    ff.report_vinculum_gt_orders_not_in_full_flow(
        make.oe([{"Order No": "N1", "ExtOrderNo": "GT290926A2"}]), make.sales([]), existing,
        _fms([]), path=str(path))
    assert pd.read_csv(path)["order_no"].tolist() == ["N1"]
    pd.testing.assert_frame_equal(existing, before)


def test_no_order_enquiry_files(ff, make):
    """A run where the daily folder could not be read still reports Sales Master orders."""
    report = ff.report_vinculum_gt_orders_not_in_full_flow(
        None, make.sales([{"Order No": "N3", "ExtOrderNo": "GT010926A1", "SKU Code": "S"}]),
        make.existing([]), _fms([]))
    assert _by_order(report) == {"N3": "gt_missing"}
