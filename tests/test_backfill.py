"""
Full Flow FMS nightly fill (yes_no.py plan_backfill) — existing rows only.

Contract:
  - rows with a blank Order No get the first unused matching Vinculum line
  - rows whose Order No is a Vinculum line get their BLANK cells refreshed (2026-10-09)
  - only BLANK cells are written; nothing is ever overwritten
  - Cancelled / Virtual Invoice / OOS rows are never written
  - a Vinculum line is used by at most one row; extras are reported, not inserted

Run: python -m pytest tests -q
"""
import pandas as pd


def test_blank_order_no_filled(make):
    writes, stats = make.plan(
        [{"GT Number": "GT010926A1", "Vinculum SKU": "SKU1", "Order Quantity": 5}],
        [{"Order No": "N100", "ExtOrderNo": "GT010926A1", "SKU Code": "SKU1", "Invoice No": "INV/1"}])
    c = make.cells(writes)
    assert c[(8, "Order No")] == "N100"
    assert c[(8, "Team")] == "Gauri"
    assert c[(8, "Invoice No")] == "INV/1"
    assert c[(8, "Invoice Amount")] == 1000
    assert c[(8, "Warehouse")] == "Vadape"
    assert c[(8, "Dispatched Qty")] == 5
    assert (8, "Actual") in c
    assert stats["filled_rows"] == 1 and stats["tiers"][1] == 1


def test_never_overwrite(make):
    writes, _ = make.plan(
        [{"GT Number": "GT1", "Vinculum SKU": "S", "Team": "Manual", "Invoice No": "KEEP"}],
        [{"Order No": "N1", "ExtOrderNo": "GT1", "SKU Code": "S", "Invoice No": "OTHER"}])
    c = make.cells(writes)
    assert (8, "Team") not in c and (8, "Invoice No") not in c, c
    assert c[(8, "Order No")] == "N1"


def test_fully_filled_row_untouched(make):
    full = {"GT Number": "GT1", "Vinculum SKU": "S", "Order No": "N1", "Team": "T",
            "Actual": "x", "Invoice No": "I", "Invoice Amount": 1, "Warehouse": "W",
            "Dispatched Qty": 5}
    writes, stats = make.plan([full], [{"Order No": "N1", "ExtOrderNo": "GT1", "SKU Code": "S"}])
    assert writes == []
    assert stats["blank_p_rows"] == 0 and stats["anchored_rows"] == 1


def test_cancelled_rows_skipped(make):
    writes, _ = make.plan(
        [{"GT Number": "GT1", "Vinculum SKU": "S", "Status": "Cancelled"},
         {"GT Number": "GT2", "Vinculum SKU": "S", "Order Quantity": "Cancelled"}],
        [{"Order No": "N1", "ExtOrderNo": "GT1", "SKU Code": "S"},
         {"Order No": "N2", "ExtOrderNo": "GT2", "SKU Code": "S"}])
    assert writes == []


def test_no_match(make):
    writes, stats = make.plan([{"GT Number": "GT1", "Vinculum SKU": "S"}],
                              [{"Order No": "N9", "ExtOrderNo": "GT9", "SKU Code": "S"}])
    assert writes == []
    assert stats["still_unmatched"] == 1


def test_case_and_space_insensitive(make):
    writes, _ = make.plan([{"GT Number": " gt1 ", "Vinculum SKU": "sku 1"}],
                          [{"Order No": "N1", "ExtOrderNo": "GT1", "SKU Code": "SKU1"}])
    assert make.cells(writes)[(8, "Order No")] == "N1"


def test_used_line_not_reused(make):
    # Row 8 already holds N1; row 9 (same GT+SKU) is blank -> gets N2, never N1 again.
    writes, _ = make.plan(
        [{"GT Number": "GT1", "Vinculum SKU": "S", "Order No": "N1"},
         {"GT Number": "GT1", "Vinculum SKU": "S"}],
        [{"Order No": "N1", "ExtOrderNo": "GT1", "SKU Code": "S"},
         {"Order No": "N2", "ExtOrderNo": "GT1", "SKU Code": "S"}])
    c = make.cells(writes)
    assert c[(9, "Order No")] == "N2", c
    assert (8, "Order No") not in c


def test_extra_lines_reported(make):
    writes, stats = make.plan(
        [{"GT Number": "GT1", "Vinculum SKU": "S"}],
        [{"Order No": "N1", "ExtOrderNo": "GT1", "SKU Code": "S"},
         {"Order No": "N2", "ExtOrderNo": "GT1", "SKU Code": "S"}])
    assert make.cells(writes)[(8, "Order No")] == "N1"
    assert stats["unplaced_sales_lines"] == 1


def test_pick_pack_bridge(make):
    writes, _ = make.plan(
        [{"GT Number": "GT1", "Vinculum SKU": "S"}],
        [{"Order No": "N1", "ExtOrderNo": "GT1", "SKU Code": "S"}],
        pick_pack={"N1": {"Pick Date": pd.Timestamp("2026-09-06 09:00"), "Pick User": "Bappa"}})
    c = make.cells(writes)
    assert c[(8, "Warehouse Team")] == "Bappa"
    assert (8, "Actual.1") in c


def test_rows_map_to_sheet_rows(make):
    writes, _ = make.plan(
        [{"GT Number": "GTX", "Vinculum SKU": "S", "Order No": "N0", "Team": "a", "Actual": "b",
          "Invoice No": "c", "Invoice Amount": 1, "Warehouse": "w", "Dispatched Qty": 1},
         {"GT Number": "GTX", "Vinculum SKU": "Q"},
         {"GT Number": "GTY", "Vinculum SKU": "S"}],
        [{"Order No": "N7", "ExtOrderNo": "GTY", "SKU Code": "S"}])
    assert {w[0] for w in writes} == {10}, writes


def test_a1(ff):
    assert ff.backfill_a1(8, "Order No") == "'Full Flow FMS'!P8"
    assert ff.backfill_a1(12, "Invoice Amount") == "'Full Flow FMS'!AY12"
    assert ff.backfill_a1(9, "Actual.1") == "'Full Flow FMS'!X9"
    assert ff.backfill_a1(9, "Dispatched Qty") == "'Full Flow FMS'!AS9"
    assert ff.backfill_a1(9, "Warehouse Team") == "'Full Flow FMS'!V9"
