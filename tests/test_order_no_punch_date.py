"""
Rows whose Order No (P) is already filled but not invoiced get S (punch date) from the
daily Order Enquiry  —  2026-10-23.

Since 12-Oct the auto-punch portal writes the Vinculum Order No into P at punch time
(one per warehouse part, comma-separated when a GT is split). The nightly job filled S
only for a BLANK P (Order Enquiry) or from a Sales Master line (invoice), so every
portal-punched, not-yet-invoiced GT kept S blank and the FMS_Dashboard showed it as
"OMS Punch" pending (the 12-15 Oct gap). Contract:
  - P filled, no Sales line, P's order in Order Enquiry -> S (+ Q if the export has
    Created by) from it; blank cells only; P never written
  - several orders in P -> earliest Order Date; cancelled orders ignored
  - V/X from Pick Pack for every order in P; with several orders only when all picked
  - a Sales Master line of any order in P still anchors the row (invoice refresh)
  - idempotent
"""
import pandas as pd


def test_split_order_nos(ff):
    assert ff.split_order_nos("N012271, N461768") == ("N012271", "N461768")
    assert ff.split_order_nos(" M892787 ") == ("M892787",)
    assert ff.split_order_nos("N1; N1 / N2") == ("N1", "N2")
    assert ff.split_order_nos("") == () and ff.split_order_nos("Cancelled") == ()


def test_portal_punched_row_gets_punch_date(make):
    writes, stats = make.plan(
        [{"GT Number": "GT121026A1", "Vinculum SKU": "S1", "Order No": "N461768"}], [],
        oe_rows=[{"Order No": "N461768", "ExtOrderNo": "GT121026A1",
                  "Order Date": "12/10/2026 03:15 PM"}])
    c = make.cells(writes)
    assert c == {(8, "Actual"): pd.Timestamp("2026-10-12 15:15")}
    assert stats["order_no_oe_rows"] == 1 and stats["oe_filled_rows"] == 0


def test_team_from_created_by_when_export_has_it(ff, make):
    existing = make.existing([{"GT Number": "GT1", "Vinculum SKU": "S", "Order No": "N1"}])
    oe = ff.prepare_oe_orders(pd.DataFrame([{
        "Order No": "N1", "ExtOrderNo": "GT1", "Order Date": "12/10/2026 10:00 AM",
        "Status": "Confirmed", "Created by": "Saloni"}]))
    writes, *_ = ff.plan_backfill(existing, make.sales([]), {}, {}, oe_orders=oe)
    c = make.cells(writes)
    assert c[(8, "Team")] == "Saloni" and (8, "Order No") not in c


def test_split_gt_earliest_date_and_pick_only_when_all_picked(make):
    pick = {"N012271": {"Pick Date": pd.Timestamp("2026-10-14 10:00"), "Pick User": "Bhagwan"},
            "N461768": {"Pick Date": pd.Timestamp("2026-10-15 09:00"), "Pick User": "Uttam"}}
    oe = [{"Order No": "N012271", "ExtOrderNo": "GT121026A1", "Order Date": "13/10/2026 11:00 AM"},
          {"Order No": "N461768", "ExtOrderNo": "GT121026A1", "Order Date": "12/10/2026 06:00 PM"}]
    rows = [{"GT Number": "GT121026A1", "Vinculum SKU": "S1", "Order No": "N012271, N461768"}]
    writes, _ = make.plan(rows, [], pick_pack=pick, oe_rows=oe)
    c = make.cells(writes)
    assert c[(8, "Actual")] == pd.Timestamp("2026-10-12 18:00")
    assert c[(8, "Actual.1")] == pd.Timestamp("2026-10-15 09:00")
    assert c[(8, "Warehouse Team")] == "Uttam"

    half = {"N012271": pick["N012271"]}
    writes, _ = make.plan(rows, [], pick_pack=half, oe_rows=oe)
    c = make.cells(writes)
    assert (8, "Actual") in c and (8, "Actual.1") not in c and (8, "Warehouse Team") not in c


def test_cancelled_order_in_p_gives_no_date(make):
    writes, _ = make.plan(
        [{"GT Number": "GT1", "Vinculum SKU": "S", "Order No": "N1"}], [],
        oe_rows=[{"Order No": "N1", "ExtOrderNo": "GT1", "Status": "Cancelled"}])
    assert writes == []


def test_existing_s_and_closed_rows_untouched(make):
    writes, _ = make.plan(
        [{"GT Number": "GT1", "Vinculum SKU": "S", "Order No": "N1",
          "Actual": "12-Oct-2026 10:00:00"},
         {"GT Number": "GT2", "Vinculum SKU": "S", "Order No": "N2",
          "Order Delivery Status": "Cancelled", "Status": "Cancelled"}], [],
        oe_rows=[{"Order No": "N1", "ExtOrderNo": "GT1"}, {"Order No": "N2", "ExtOrderNo": "GT2"}])
    assert writes == []


def test_invoiced_order_still_anchors_split_p(make):
    # A Sales Master line of one order in a split P anchors the row -> invoice refresh.
    writes, stats = make.plan(
        [{"GT Number": "GT1", "Vinculum SKU": "S1", "Order No": "N1, N2"}],
        [{"Order No": "N2", "ExtOrderNo": "GT1", "SKU Code": "S1", "Invoice No": "INV9",
          "Order Date": "2026-10-12 10:00:00"}])
    c = make.cells(writes)
    assert c[(8, "Invoice No")] == "INV9" and (8, "Order No") not in c
    assert stats["anchored_rows"] == 1


def test_idempotent(ff, make):
    existing = make.existing([{"GT Number": "GT1", "Vinculum SKU": "S", "Order No": "N1"}])
    oe = make.oe([{"Order No": "N1", "ExtOrderNo": "GT1"}])
    sales = make.sales([])
    writes, *_ = ff.plan_backfill(existing, sales, {}, {}, oe_orders=oe)
    assert len(writes) == 1
    after = existing.astype(object)
    for sheet_row, column, value, _ in writes:
        after.at[sheet_row - 8, column] = value
    again, stats, *_ = ff.plan_backfill(after, sales, {}, {}, oe_orders=oe)
    assert again == [] and stats["order_no_oe_rows"] == 0


def test_sales_lookback_days_override(ff):
    assert ff._lookback_days(None) == 7 and ff._lookback_days("") == 7
    assert ff._lookback_days("12") == 12 and ff._lookback_days(" 14 ") == 14
    assert ff._lookback_days("0") == 7 and ff._lookback_days("400") == 7
    assert ff._lookback_days("abc") == 7
