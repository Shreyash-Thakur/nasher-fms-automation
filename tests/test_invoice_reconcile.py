"""
Nightly Vinculum cross-check of Order No (P) and quantities  -  2026-11-16.

User: "can you not fix the quantiyt and order id overlapp all such issues at once it just
requires cross check with viniculum ... make sure this doesn't happen in future".
Contract:
  (a) invoice-proven P correction: an open row whose AW invoice is, in the Sales Master,
      a line of THIS GT + SKU on ONE order X, while P names other order(s) whose lines are
      all known (Sales / Pick Pack) and do NOT carry the SKU -> P := X. The one allowed
      overwrite of a filled cell; every change is logged and listed in
      reports/order_no_invoice_fixes.csv. Anything less certain is held (report only):
      P order also carries the SKU (split line), P order lines unknown, invoice without a
      line for the SKU, invoice of another GT, more fixes than the cap.
  (b) quantity cross-check per (P order, SKU): AS vs Vinculum Pick Qty, J vs Orderqty,
      and the same order + SKU on rows of two GTs (overlap) -> reports/qty_crosscheck.csv.
      AS is corrected only when the invoice proves one line and one row (unambiguous).
  (c) an Order No on rows of two unrelated GTs (not the R/T/F family of one GT)
      -> reports/duplicate_order_nos.csv + a count in the log.
Record: "Errors and their fixes/2026-11-16-gt-only-order-fill-guard.md".
"""
import pandas as pd

GT = "GT141026A1"


def _line(order, sku, inv, gt=GT, pick=1, orderqty=None):
    r = {"Order No": order, "ExtOrderNo": gt, "SKU Code": sku, "Invoice No": inv,
         "Pick Qty": pick}
    if orderqty is not None:
        r["Orderqty"] = orderqty
    return r


SALES = [_line("O94134", "NMH8113CL", "27/10/NM4/340"),
         _line("O94134", "NMH8113CM", "27/10/NM4/340"),
         _line("M892791", "NMHPP013CS", "27/10/NM1/1273"),
         _line("N89724", "NMHPP0127AS", "27/10/GJ1/340")]


def _setup(ff, make, rows, sales_rows, oe_rows=None, pick_pack_skus=None):
    existing = make.existing(rows)
    sales = make.sales(sales_rows)
    oe = make.oe(oe_rows) if oe_rows is not None else None
    slots = [ff._Slot("existing", i, row) for i, row in existing.iterrows()]
    ff.allocate_lines(slots, sales)
    index = ff.OrderEnquiryIndex(oe, sales, existing, pick_pack_skus)
    return existing, sales, slots, index


# ------------------------------------------------------------ (a) P from invoice

def test_invoice_proves_other_order_p_corrected_and_logged(ff, make, tmp_path, capsys):
    rows = [{"GT Number": GT, "Vinculum SKU": "NMH8113CL", "Order No": "O94134",
             "Invoice No": "27/10/NM4/340", "Dispatched Qty": 1},
            {"GT Number": GT, "Vinculum SKU": "NMHPP013CS", "Order No": "O94134",
             "Invoice No": "27/10/NM1/1273", "Dispatched Qty": 1},
            {"GT Number": GT, "Vinculum SKU": "NMHPP0127AS", "Order No": "O94134",
             "Invoice No": "27/10/GJ1/340", "Dispatched Qty": 1}]
    existing, sales, slots, index = _setup(ff, make, rows, SALES)
    path = tmp_path / "fix.csv"
    writes, report = ff.plan_invoice_order_fixes(existing, slots, sales, index, path=str(path))
    assert writes == [(9, "Order No", "M892791", "invoice_fix"),
                      (10, "Order No", "N89724", "invoice_fix")]
    out = capsys.readouterr().out
    assert ("invoice-proven P fix: row 9 GT GT141026A1 / NMHPP013CS P O94134 -> M892791 "
            "(invoice 27/10/NM1/1273)") in out
    got = pd.read_csv(path)
    assert got[got["action"] == "fixed"]["sheet_row"].tolist() == [9, 10]
    assert 8 not in got["sheet_row"].tolist()                 # P agrees with its invoice


def test_split_line_held_when_p_order_also_carries_sku(ff, make):
    # The SKU is really on both orders (remainder on a second order): no rewrite.
    rows = [{"GT Number": "GT1", "Vinculum SKU": "S", "Order No": "O1", "Invoice No": "I2"}]
    sales = [_line("O1", "S", "I1", gt="GT1"), _line("O2", "S", "I2", gt="GT1")]
    existing, sales, slots, index = _setup(ff, make, rows, sales)
    writes, report = ff.plan_invoice_order_fixes(existing, slots, sales, index)
    assert writes == []
    assert report.iloc[0]["action"] == "held:p_order_also_carries_sku"


def test_p_order_lines_unknown_held(ff, make):
    rows = [{"GT Number": "GT1", "Vinculum SKU": "S", "Order No": "N555", "Invoice No": "I2"}]
    existing, sales, slots, index = _setup(ff, make, rows, [_line("O2", "S", "I2", gt="GT1")])
    writes, report = ff.plan_invoice_order_fixes(existing, slots, sales, index)
    assert writes == [] and report.iloc[0]["action"] == "held:p_order_lines_unknown"


def test_pick_pack_lines_make_p_order_known(ff, make):
    rows = [{"GT Number": "GT1", "Vinculum SKU": "S", "Order No": "N555", "Invoice No": "I2"}]
    existing, sales, slots, index = _setup(ff, make, rows, [_line("O2", "S", "I2", gt="GT1")],
                                           pick_pack_skus={"N555": {"OTHER"}})
    writes, _ = ff.plan_invoice_order_fixes(existing, slots, sales, index)
    assert writes == [(8, "Order No", "O2", "invoice_fix")]


def test_invoice_without_line_for_sku_or_of_other_gt_held(ff, make):
    rows = [{"GT Number": "GT1", "Vinculum SKU": "X", "Order No": "O1", "Invoice No": "I2"},
            {"GT Number": "GT1", "Vinculum SKU": "S", "Order No": "O1", "Invoice No": "I9"}]
    sales = [_line("O1", "S", "I1", gt="GT1"), _line("O2", "S", "I2", gt="GT1"),
             _line("O9", "S", "I9", gt="GT2")]
    existing, sales, slots, index = _setup(ff, make, rows, sales)
    writes, report = ff.plan_invoice_order_fixes(existing, slots, sales, index)
    assert writes == []
    assert report["action"].tolist() == ["held:invoice_has_no_line_for_sku",
                                         "held:invoice_of_other_gt"]


def test_unknown_invoice_closed_and_blank_rows_untouched(ff, make):
    rows = [{"GT Number": "GT1", "Vinculum SKU": "S", "Order No": "O1", "Invoice No": "NEW/1"},
            {"GT Number": "GT1", "Vinculum SKU": "T", "Order No": "O1", "Invoice No": "I2",
             "Status": "Cancelled"},
            {"GT Number": "GT1", "Vinculum SKU": "T", "Invoice No": "I2"}]   # blank P: fill path
    sales = [_line("O1", "S", "I1", gt="GT1"), _line("O2", "T", "I2", gt="GT1")]
    existing, sales, slots, index = _setup(ff, make, rows, sales)
    writes, report = ff.plan_invoice_order_fixes(existing, slots, sales, index)
    assert writes == [] and report.empty


def test_cap_holds_every_fix(ff, make, capsys):
    rows = [{"GT Number": GT, "Vinculum SKU": "NMHPP013CS", "Order No": "O94134",
             "Invoice No": "27/10/NM1/1273"},
            {"GT Number": GT, "Vinculum SKU": "NMHPP0127AS", "Order No": "O94134",
             "Invoice No": "27/10/GJ1/340"}]
    existing, sales, slots, index = _setup(ff, make, rows, SALES)
    writes, report = ff.plan_invoice_order_fixes(existing, slots, sales, index, max_fixes=1)
    assert writes == [] and set(report["action"]) == {"held:over_cap"}
    assert "WARNING" in capsys.readouterr().out


def test_fixed_sheet_is_stable_next_run(ff, make):
    rows = [{"GT Number": GT, "Vinculum SKU": "NMHPP013CS", "Order No": "O94134",
             "Invoice No": "27/10/NM1/1273"}]
    existing, sales, slots, index = _setup(ff, make, rows, SALES)
    writes, _ = ff.plan_invoice_order_fixes(existing, slots, sales, index)
    existing.at[0, "Order No"] = writes[0][2]
    existing, sales, slots, index = _setup(ff, make, existing.to_dict("records"), SALES)
    again, report = ff.plan_invoice_order_fixes(existing, slots, sales, index)
    assert again == [] and report.empty


# ------------------------------------------------------------ (b) quantities

def test_dispatched_qty_fixed_when_invoice_line_unambiguous(ff, make, tmp_path):
    rows = [{"GT Number": "GT1", "Vinculum SKU": "S", "Order No": "O1", "Invoice No": "I1",
             "Dispatched Qty": "5"}]
    existing, sales, slots, _ = _setup(ff, make, rows, [_line("O1", "S", "I1", gt="GT1", pick=3)])
    path = tmp_path / "q.csv"
    writes, report = ff.plan_qty_crosscheck(existing, slots, sales, path=str(path))
    assert writes == [(8, "Dispatched Qty", 3, "qty_fix")]
    assert report["check"].tolist() == ["dispatched_fixed"]
    assert pd.read_csv(path)["check"].tolist() == ["dispatched_fixed"]


def test_dispatched_mismatch_reported_only_when_ambiguous(ff, make):
    rows = [{"GT Number": "GT1", "Vinculum SKU": "S", "Order No": "O1", "Invoice No": "I1",
             "Dispatched Qty": 2},
            {"GT Number": "GT1", "Vinculum SKU": "S", "Order No": "O1", "Invoice No": "I1",
             "Dispatched Qty": 2}]
    existing, sales, slots, _ = _setup(ff, make, rows, [_line("O1", "S", "I1", gt="GT1", pick=3)])
    writes, report = ff.plan_qty_crosscheck(existing, slots, sales)
    assert writes == []
    r = report.iloc[0]
    assert r["check"] == "dispatched_mismatch" and r["full_flow_qty"] == 4 \
        and r["vinculum_qty"] == 3 and r["sheet_rows"] == "8 9"


def test_matching_quantities_not_reported(ff, make):
    rows = [{"GT Number": "GT1", "Vinculum SKU": "S", "Order No": "O1", "Invoice No": "I1",
             "Dispatched Qty": 3, "Order Quantity": 3}]
    existing, sales, slots, _ = _setup(ff, make, rows,
                                       [_line("O1", "S", "I1", gt="GT1", pick=3, orderqty=3)])
    writes, report = ff.plan_qty_crosscheck(existing, slots, sales)
    assert writes == [] and report.empty


def test_order_qty_mismatch_reported_never_written(ff, make):
    rows = [{"GT Number": "GT1", "Vinculum SKU": "S", "Order No": "O1", "Invoice No": "I1",
             "Dispatched Qty": 3, "Order Quantity": 10}]
    existing, sales, slots, _ = _setup(ff, make, rows,
                                       [_line("O1", "S", "I1", gt="GT1", pick=3, orderqty=8)])
    writes, report = ff.plan_qty_crosscheck(existing, slots, sales)
    assert writes == []
    r = report.iloc[0]
    assert r["check"] == "order_qty_mismatch" and r["full_flow_qty"] == 10 and r["vinculum_qty"] == 8


def test_qty_overlap_same_order_sku_on_two_gts(ff, make):
    rows = [{"GT Number": "GT1", "Vinculum SKU": "S", "Order No": "O1", "Dispatched Qty": 3},
            {"GT Number": "GT2", "Vinculum SKU": "S", "Order No": "O1", "Dispatched Qty": 3}]
    existing, sales, slots, _ = _setup(ff, make, rows, [_line("O1", "S", "I1", gt="GT1", pick=3)])
    writes, report = ff.plan_qty_crosscheck(existing, slots, sales)
    assert writes == []
    over = report[report["check"] == "qty_overlap"]
    assert len(over) == 1 and over.iloc[0]["gt_number"] == "GT1 GT2"


def test_qty_check_uses_invoice_corrected_order(ff, make):
    rows = [{"GT Number": GT, "Vinculum SKU": "NMHPP013CS", "Order No": "O94134",
             "Invoice No": "27/10/NM1/1273", "Dispatched Qty": 2}]
    existing, sales, slots, index = _setup(ff, make, rows, SALES)
    # Before the P fix the row's (P order, SKU) has no Vinculum line: nothing to compare.
    assert ff.plan_qty_crosscheck(existing, slots, sales)[0] == []
    fixes, _ = ff.plan_invoice_order_fixes(existing, slots, sales, index)
    writes, report = ff.plan_qty_crosscheck(existing, slots, sales, order_fixes=fixes)
    assert writes == [(8, "Dispatched Qty", 1, "qty_fix")]       # M892791 picked 1


# ------------------------------------------------------------ (c) duplicate orders

def test_duplicate_order_on_unrelated_gts(ff, make, tmp_path, capsys):
    existing = make.existing([
        {"GT Number": "GT1", "Vinculum SKU": "S", "Order No": "O1"},
        {"GT Number": "GT2", "Vinculum SKU": "S", "Order No": "O1, O7"},
        {"GT Number": "GT3", "Vinculum SKU": "S", "Order No": "O3"},
        {"GT Number": "GT3R", "Vinculum SKU": "S", "Order No": "O3"},          # repunch family
        {"GT Number": "GT4", "Vinculum SKU": "S", "Order No": "Cancelled", "Status": "Cancelled"},
        {"GT Number": "GT5", "Vinculum SKU": "S", "Order No": "Cancelled", "Status": "Cancelled"},
    ])
    sales = make.sales([_line("O1", "S", "I1", gt="GT1")])
    path = tmp_path / "d.csv"
    report = ff.report_duplicate_order_nos(existing, sales, path=str(path))
    assert report["order_no"].tolist() == ["O1"]
    r = report.iloc[0]
    assert r["gt_numbers"] == "GT1 GT2" and r["sheet_rows"] == "8 9" and r["vinculum_gt"] == "GT1"
    assert "Order Nos on rows of two unrelated GTs: 1" in capsys.readouterr().out
    assert pd.read_csv(path)["order_no"].tolist() == ["O1"]
