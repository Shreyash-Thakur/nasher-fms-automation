"""
Order Enquiry GT-only match needs SKU proof before it fills P  -  2026-11-16.

The daily Order Enquiry export is ORDER-LEVEL (no SKU). OrderEnquiryIndex.match used
to fill P with a GT's only live, not-invoiced order on EVERY blank-P row of the GT.
Vinculum splits a GT into one order per warehouse and GTs are punched in parts, so
while only the first part existed its order was stamped on rows that belong to the
later parts (GT141026A1: 72 rows O94134, 40 of them on M892791 / N89724 / N012273 /
N051641). Contract now:
  - an Order Enquiry order fills P only on a row whose SKU (Vinculum SKU or SKU Code)
    is on that order in line-level data (Pick Pack Order No + SKU Code, Sales lines)
  - no proof -> P stays blank, outcome "no_sku_proof", and the run log lists
    "order-enquiry GT-only match skipped (no SKU proof): row N GT X order Y"
  - Sales Master (line-level) fills are unchanged
  - a filled P is never overwritten; P whose orders are known line by line and do not
    carry the row's SKU is reported (reports/order_no_suspect_no_sku.csv)
Record: "Errors and their fixes/2026-11-16-gt-only-order-fill-guard.md".
"""
import pandas as pd


GT = "GT141026A1"


def _rows(*skus, gt=GT):
    return [{"GT Number": gt, "Vinculum SKU": s} for s in skus]


def _apply(existing, writes):
    after = existing.astype(object)
    for sheet_row, column, value, _ in writes:
        after.at[sheet_row - 8, column] = value
    return after


# ------------------------------------------------------------ pick pack SKU index

def test_pick_pack_sku_index(ff):
    df = pd.DataFrame({"Order No": ["O1", "O1", "o2 ", "", "O3"],
                       "SKU Code": ["nmh8113cl", "NMH8113CM", "A", "B", ""],
                       "Pick Date": [""] * 5, "Pick User": [""] * 5})
    idx = ff.build_pick_pack_sku_index(df)
    assert idx == {"O1": {"NMH8113CL", "NMH8113CM"}, "O2": {"A"}}


def test_pick_pack_sku_index_without_sku_column_is_empty(ff):
    df = pd.DataFrame({"Order No": ["O1"], "Pick Date": [""], "Pick User": [""]})
    assert ff.build_pick_pack_sku_index(df) == {}
    assert ff.build_pick_pack_sku_index(None) == {}


# ------------------------------------------------------------ GT punched in two parts

def test_two_part_gt_first_run_fills_only_proven_skus(ff, make, capsys):
    # Run 1: only the first part (O94134) is punched. Its Pick Pack lines hold A and B.
    existing = make.existing(_rows("A", "B", "C", "D"))
    oe1 = make.oe([{"Order No": "O94134", "ExtOrderNo": GT}])
    skus1 = {"O94134": {"A", "B"}}
    writes, stats, slots, _ = ff.plan_backfill(existing, make.sales([]), {}, {},
                                               oe_orders=oe1, pick_pack_skus=skus1)
    c = make.cells(writes)
    assert c[(8, "Order No")] == "O94134" and c[(9, "Order No")] == "O94134"
    assert (10, "Order No") not in c and (11, "Order No") not in c   # C, D: not on O94134
    assert not any(r in (10, 11) for r, *_ in writes)                # no S either
    assert stats["oe_filled_rows"] == 2
    assert stats["oe_outcomes"]["no_sku_proof"] == 2
    assert [(s.ref, s.oe_unproven) for s in slots if s.oe_outcome == "no_sku_proof"] == \
        [(2, "O94134"), (3, "O94134")]

    skipped = ff.report_oe_no_sku_proof(slots)
    out = capsys.readouterr().out
    assert skipped == 2
    assert ("order-enquiry GT-only match skipped (no SKU proof): row 10 GT GT141026A1 "
            "order O94134") in out
    assert "row 11 GT GT141026A1 order O94134" in out

    # Run 2: the second part (M892791, C + D) is punched. Rows A/B keep O94134 (never
    # rewritten); C/D are not stamped with O94134 - the GT now has two orders.
    after = _apply(existing, writes)
    oe2 = make.oe([{"Order No": "O94134", "ExtOrderNo": GT},
                   {"Order No": "M892791", "ExtOrderNo": GT}])
    skus2 = {"O94134": {"A", "B"}, "M892791": {"C", "D"}}
    writes2, stats2, *_ = ff.plan_backfill(after, make.sales([]), {}, {},
                                           oe_orders=oe2, pick_pack_skus=skus2)
    c2 = make.cells(writes2)
    assert not any(col == "Order No" for _, col, *_ in writes2)
    assert c2.get((10, "Order No")) != "O94134" and c2.get((11, "Order No")) != "O94134"
    assert stats2["oe_outcomes"]["ambiguous"] == 2


def test_two_part_gt_not_picked_yet_nothing_filled(make):
    # Punched, not picked, not invoiced: no line-level data at all -> no P anywhere.
    writes, stats = make.plan(_rows("A", "B"), [],
                              oe_rows=[{"Order No": "O94134", "ExtOrderNo": GT}])
    assert writes == []
    assert stats["oe_filled_rows"] == 0 and stats["oe_outcomes"]["no_sku_proof"] == 2
    assert stats["still_unmatched"] == 2


# ------------------------------------------------------------ SKU-proven fills

def test_sku_proven_fill_still_works(make):
    writes, stats = make.plan(
        _rows("NMH8113CL", gt="GT1"), [],
        pick_pack={"N1": {"Pick Date": pd.Timestamp("2026-10-07 10:00"), "Pick User": "Bidyut"}},
        pick_pack_skus={"N1": {"NMH8113CL"}},
        oe_rows=[{"Order No": "N1", "ExtOrderNo": "GT1", "Order Date": "06/10/2026 06:27 PM"}])
    c = make.cells(writes)
    assert c[(8, "Order No")] == "N1"
    assert c[(8, "Actual")] == pd.Timestamp("2026-10-06 18:27")
    assert c[(8, "Warehouse Team")] == "Bidyut" and (8, "Actual.1") in c
    assert stats["oe_filled_rows"] == 1 and stats["oe_outcomes"]["no_sku_proof"] == 0


def test_proof_through_sku_code_alt_field(make):
    # FMS SKU Code (col E) is what Pick Pack lists for a bundle.
    writes, _ = make.plan(
        [{"GT Number": "GT1", "Vinculum SKU": "NMB_CHILD", "SKU Code": "NM_BUNDLE"}], [],
        pick_pack_skus={"N1": {"NM_BUNDLE"}},
        oe_rows=[{"Order No": "N1", "ExtOrderNo": "GT1"}])
    assert make.cells(writes)[(8, "Order No")] == "N1"


def test_single_order_gt_all_skus_proven(make):
    writes, stats = make.plan(
        _rows("S1", "S2", "S3", gt="GT061026A4"), [],
        pick_pack_skus={"NM11832": {"S1", "S2", "S3"}},
        oe_rows=[{"Order No": "NM11832", "ExtOrderNo": "GT061026A4"}])
    c = make.cells(writes)
    assert [c[(r, "Order No")] for r in (8, 9, 10)] == ["NM11832"] * 3
    assert stats["oe_filled_rows"] == 3 and stats["oe_outcomes"]["no_sku_proof"] == 0


def test_proof_of_another_order_does_not_count(make):
    # The SKU is on some order, but not on the GT's only Order Enquiry order.
    writes, stats = make.plan(
        _rows("A", gt="GT1"), [],
        pick_pack_skus={"OTHER": {"A"}},
        oe_rows=[{"Order No": "N1", "ExtOrderNo": "GT1"}])
    assert writes == [] and stats["oe_outcomes"]["no_sku_proof"] == 1


def test_sales_master_fill_unchanged(make):
    # Line-level Sales Master fills need no Pick Pack proof.
    writes, stats = make.plan(
        _rows("S", gt="GT1"), [{"Order No": "N1", "ExtOrderNo": "GT1", "SKU Code": "S"}],
        oe_rows=[{"Order No": "N1", "ExtOrderNo": "GT1"}])
    assert make.cells(writes)[(8, "Order No")] == "N1"
    assert stats["filled_rows"] == 1 and stats["oe_outcomes"]["no_sku_proof"] == 0


# ------------------------------------------------------------ append path

def test_match_without_skus_never_fills(ff, make):
    oe_index = ff.OrderEnquiryIndex(make.oe([{"Order No": "N9", "ExtOrderNo": "GT9"}]),
                                    make.sales([]), make.existing([]),
                                    order_skus={"N9": {"S"}})
    assert oe_index.match("GT9", "GT9") == (None, "no_sku_proof")          # GT only
    assert oe_index.match("GT9", "GT9", {"X"}) == (None, "no_sku_proof")
    assert oe_index.match("GT9", "GT9", {"S", ""}) == ("N9", "fill")
    assert oe_index.gt_only_order("GT9", "GT9") == "N9"


def test_new_rows_need_proof_too(ff, make, capsys):
    oe_index = ff.OrderEnquiryIndex(make.oe([{"Order No": "N9", "ExtOrderNo": "GT9"}]),
                                    make.sales([]), make.existing([]),
                                    order_skus={"N9": {"S"}})
    fms = pd.DataFrame([{**{c: "" for c in ff.FMS_COLUMNS}, "GT Number": "GT9",
                         "Vinculum SKU": sku} for sku in ("S", "T")])
    new_slots = [ff._Slot("new", pos, row) for pos, (_, row) in enumerate(fms.iterrows())]
    for s in new_slots:                                  # as main() does for appends
        s.oe_order = ff.match_slot_oe(s, oe_index)
    assert [s.oe_order for s in new_slots] == ["N9", None]
    assert new_slots[1].oe_outcome == "no_sku_proof" and new_slots[1].oe_unproven == "N9"
    assert ff.report_oe_no_sku_proof(new_slots) == 1
    assert ("order-enquiry GT-only match skipped (no SKU proof): new row 2 (append) GT GT9 "
            "order N9") in capsys.readouterr().out


# ------------------------------------------------------------ suspect P report

def test_suspect_order_no_reported_never_rewritten(ff, make, tmp_path):
    # A run before this fix stamped O94134 on C. O94134's lines are known (A, B) and the
    # GT's other order M892791 carries C. Report it; never write P.
    existing = make.existing([
        {"GT Number": GT, "Vinculum SKU": "A", "Order No": "O94134"},
        {"GT Number": GT, "Vinculum SKU": "C", "Order No": "O94134"},
        {"GT Number": GT, "Vinculum SKU": "E", "Order No": "N555"},   # N555 lines unknown
    ])
    oe = make.oe([{"Order No": "O94134", "ExtOrderNo": GT},
                  {"Order No": "M892791", "ExtOrderNo": GT},
                  {"Order No": "N555", "ExtOrderNo": GT}])
    skus = {"O94134": {"A", "B"}, "M892791": {"C"}}
    writes, _, slots, _ = ff.plan_backfill(existing, make.sales([]), {}, {},
                                           oe_orders=oe, pick_pack_skus=skus)
    assert not any(col == "Order No" for _, col, *_ in writes)
    index = ff.OrderEnquiryIndex(oe, make.sales([]), existing, order_skus=skus)
    path = tmp_path / "s.csv"
    report = ff.report_order_no_without_sku_proof(existing, slots, index, path=str(path))
    assert list(report["sheet_row"]) == [9]
    r = report.iloc[0]
    assert r["order_no_in_p"] == "O94134" and r["orders_with_sku"] == "M892791"
    assert pd.read_csv(path)["sheet_row"].tolist() == [9]
