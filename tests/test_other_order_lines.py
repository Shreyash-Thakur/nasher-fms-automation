"""
Invoiced rows never filled because P holds the WRONG order of a split GT  —  2026-10-31.

GT141026A1 (70 rows) and GT151026A5 (3 rows): someone wrote ONE Vinculum Order No into P
of every row of the GT (O94134 / M892795). The GT was really punched as several orders
(N012273, M892791, N051641, N89724 / N012276), each invoiced in the 15/16-Oct Sales
Report. A Sales line only anchors a row whose P lists the line's order, so these lines
stayed "unplaced (split shipments)" and AS/AW/AX/AY/Q stayed blank.

Contract (blank cells only; P is never written; the row is listed for ops to fix P):
  - a LEFTOVER line (no row anchored / filled / appended with it) goes to a row of the
    same GT and the same Vinculum SKU (tier 1/2) whose P orders
      * have at least one Sales line in this GT (invoiced, so their SKUs are known), and
      * have no line at all for this row's SKU;
  - and only when the line's order is in no row's P of that GT (an order claimed by
    other rows keeps its lines for them).
Also: when a line exists twice (truncated copy without Invoice No + full copy), the copy
WITH the invoice is the one anchored.
"""
import pandas as pd


GT = "GT141026A1"


def _rows(*specs):
    return [{"GT Number": GT, "Vinculum SKU": sku, "Order No": p} for sku, p in specs]


def _line(order, sku, inv, **extra):
    return {"Order No": order, "ExtOrderNo": GT, "SKU Code": sku, "Invoice No": inv, **extra}


def test_row_with_wrong_p_gets_invoice_of_other_order(make):
    writes, stats = make.plan(
        _rows(("A", "O94134"), ("B", "O94134")),
        [_line("O94134", "A", "27/10/NM4/340"), _line("M892791", "B", "27/10/NM1/1273")])
    c = make.cells(writes)
    assert c[(9, "Invoice No")] == "27/10/NM1/1273"
    assert c[(9, "Warehouse")] == "Vadape" and c[(9, "Team")] == "Gauri"
    assert c[(9, "Dispatched Qty")] == 5 and c[(9, "Invoice Amount")] == 1000
    assert (9, "Order No") not in c                       # P is never written
    assert {w[3] for w in writes if w[0] == 9} == {"refresh_other_order"}
    assert c[(8, "Invoice No")] == "27/10/NM4/340"        # normal anchor unchanged
    assert stats["other_order_rows"] == 1


def test_never_overwrites_filled_cells(make):
    writes, _ = make.plan(
        [{"GT Number": GT, "Vinculum SKU": "A", "Order No": "O1"},
         {"GT Number": GT, "Vinculum SKU": "B", "Order No": "O1", "Invoice No": "KEEP",
          "Team": "Prem"}],
        [_line("O1", "A", "I1"), _line("O2", "B", "I2")])
    c = make.cells(writes)
    assert (9, "Invoice No") not in c and (9, "Team") not in c
    assert c[(9, "Warehouse")] == "Vadape"


def test_order_claimed_by_another_row_is_not_reassigned(make):
    # O2 is in P of row 10, so its spare B line (second shipment) is not given to row 9.
    writes, stats = make.plan(
        _rows(("A", "O1"), ("B", "O1"), ("B", "O2")),
        [_line("O1", "A", "I1"), _line("O2", "B", "I2"), _line("O2", "B", "I3")])
    c = make.cells(writes)
    assert (9, "Invoice No") not in c
    assert c[(10, "Invoice No")] == "I2"
    assert stats["other_order_rows"] == 0


def test_p_order_not_invoiced_yet_is_left_alone(make):
    # O1 has no Sales line at all: its SKUs are unknown, so nothing is borrowed.
    writes, _ = make.plan(_rows(("B", "O1")), [_line("O2", "B", "I2")])
    assert not [w for w in writes if w[1] == "Invoice No"]


def test_p_order_has_this_sku_is_left_alone(make):
    # Two rows of the same SKU, P's order has a line for it (used by the first row).
    writes, _ = make.plan(
        _rows(("B", "O1"), ("B", "O1")),
        [_line("O1", "B", "I1"), _line("O2", "B", "I2")])
    c = make.cells(writes)
    assert c[(8, "Invoice No")] == "I1" and (9, "Invoice No") not in c


def test_closed_and_other_gt_rows_untouched(make):
    writes, _ = make.plan(
        [{"GT Number": GT, "Vinculum SKU": "A", "Order No": "O1"},
         {"GT Number": GT, "Vinculum SKU": "B", "Order No": "O1", "Status": "x",
          "Invoice No": "Cancelled"},
         {"GT Number": "GT141026A10", "Vinculum SKU": "B", "Order No": "O1"}],
        [_line("O1", "A", "I1"), _line("O2", "B", "I2")])
    assert {w[0] for w in writes} == {8}


def test_alt_sku_only_match_is_not_borrowed(make):
    writes, _ = make.plan(
        [{"GT Number": GT, "Vinculum SKU": "A", "Order No": "O1"},
         {"GT Number": GT, "Vinculum SKU": "B", "SKU Code": "BUNDLE", "Order No": "O1"}],
        [_line("O1", "A", "I1"), _line("O2", "OTHER", "I2", **{"Bundle SKU Code": "BUNDLE"})])
    assert {w[0] for w in writes} == {8}


def test_line_used_once_and_idempotent(make):
    rows = _rows(("A", "O1"), ("B", "O1"), ("B", "O1"))
    sales = [_line("O1", "A", "I1"), _line("O2", "B", "I2")]
    writes, _ = make.plan(rows, sales)
    got = [w[0] for w in writes if w[1] == "Invoice No"]
    assert got == [8, 9]
    rows[0]["Invoice No"], rows[1]["Invoice No"] = "I1", "I2"
    for col in ("Team", "Warehouse", "Dispatched Qty", "Invoice Amount"):
        rows[0][col] = rows[1][col] = "x"
    writes, _ = make.plan(rows, sales)
    assert not [w for w in writes if w[1] == "Invoice No"]


def test_mismatch_report_lists_rows_for_ops(ff, make):
    existing = make.existing(_rows(("A", "O1"), ("B", "O1")))
    sales = make.sales([_line("O1", "A", "I1"), _line("O2", "B", "I2")])
    _, _, slots, _ = ff.plan_backfill(existing, sales, {}, {})
    report = ff.report_order_no_mismatch(existing, slots, sales)
    assert report.to_dict("records") == [{
        "sheet_row": 9, "gt_number": GT, "vinculum_sku": "B", "order_no_in_p": "O1",
        "vinculum_order_for_sku": "O2", "vinculum_invoice_no": "I2"}]


def test_invoiced_copy_of_a_line_is_anchored(make):
    # 10.10.2026 Sales Report: 803 of 869 rows cut off after "Shipping IGST" (no Invoice
    # No / warehouse / Total / Created by). If the same line later arrives complete, the
    # complete copy must win, not the first (truncated) one.
    trunc = _line("M892782", "B", "", **{"Originating warehouse": "", "Total": None,
                                         "Created by": "", "_TAB": "Oct-26"})
    full = _line("M892782", "B", "27/10/NM1/9999", **{"_TAB": "daily:24.10.2026"})
    writes, _ = make.plan(_rows(("B", "M892782")), [trunc, full])
    c = make.cells(writes)
    assert c[(8, "Invoice No")] == "27/10/NM1/9999"


def test_daily_rows_without_invoice_are_counted(ff):
    df = pd.DataFrame({"Invoice No": ["27/10/NM7/1", "", None, "  "],
                       "status": ["Shipped complete"] * 4})
    assert ff.sales_rows_without_invoice(df) == 3
    assert ff.sales_rows_without_invoice(pd.DataFrame({"Invoice No": ["X"]})) == 0
