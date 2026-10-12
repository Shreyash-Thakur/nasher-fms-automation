"""Order No (P) = "Cancelled" for rows whose status is Cancelled (2026-10-12).

Contract:
  - blank P + status Cancelled (T, a stage status or AJ; trimmed, any case) -> "Cancelled"
  - a real Vinculum Order No is never replaced
  - rows that are not cancelled are untouched
  - P "Cancelled" whose status is no longer Cancelled (no "Cancelled" left in J / P:AW)
    -> cleared, so the normal fill can run
  - idempotent; "Cancelled" in P is never used as a lookup key (Python + pod.py)
"""
from datetime import datetime

import pandas as pd


def _plan(ff, make, rows):
    return ff.plan_cancelled_order_no(make.existing(rows))


def test_blank_order_no_and_cancelled_status_marked(ff, make):
    marks, clears = _plan(ff, make, [
        {"GT Number": "GT1", "Vinculum SKU": "S", "Status": "Cancelled"},
        {"GT Number": "GT2", "Vinculum SKU": "S", "Status": "  cancelled "},          # trim / case
        {"GT Number": "GT3", "Vinculum SKU": "S", "Order Delivery Status": "CANCELLED"},  # AJ
        {"GT Number": "GT4", "Vinculum SKU": "S", "Status.3": "Cancelled"},          # stage status
    ])
    assert marks == [(r, "Order No", "Cancelled", "cancel_mark") for r in (8, 9, 10, 11)]
    assert clears == []


def test_real_order_no_kept(ff, make):
    marks, clears = _plan(ff, make, [
        {"GT Number": "GT1", "Vinculum SKU": "S", "Order No": "N123", "Status": "Cancelled"}])
    assert marks == [] and clears == []


def test_not_cancelled_untouched(ff, make):
    marks, clears = _plan(ff, make, [
        {"GT Number": "GT1", "Vinculum SKU": "S", "Status": "Delay"},
        {"GT Number": "GT2", "Vinculum SKU": "S", "Order Delivery Status": "Delivered"},
        {"GT Number": "GT3", "Vinculum SKU": "S", "Order No": "N1", "Status": "On Time"},
        {"GT Number": "GT4", "Vinculum SKU": "S", "Order Delivery Status": "Virtual Invoice"},
        {"GT Number": "GT5", "Vinculum SKU": "S", "Status": "Cancelled by customer"},   # not exact
    ])
    assert marks == [] and clears == []


def test_sentinel_cleared_when_status_no_longer_cancelled(ff, make):
    marks, clears = _plan(ff, make, [
        # status reverted (T:AW cleared) -> sentinel cleared
        {"GT Number": "GT1", "Vinculum SKU": "S", "Order No": "Cancelled", "Status": "Delay"},
        {"GT Number": "GT2", "Vinculum SKU": "S", "Order No": " cancelled "},
        # still cancelled (status, or "Cancelled" elsewhere in J / P:AW) -> kept
        {"GT Number": "GT3", "Vinculum SKU": "S", "Order No": "Cancelled", "Status": "Cancelled"},
        {"GT Number": "GT4", "Vinculum SKU": "S", "Order No": "Cancelled", "Invoice No": "Cancelled"},
        {"GT Number": "GT5", "Vinculum SKU": "S", "Order No": "Cancelled", "Order Quantity": "Cancelled"},
    ])
    assert marks == []
    assert clears == [(8, "Order No", "", "cancel_clear"), (9, "Order No", "", "cancel_clear")]


def test_idempotent(ff, make):
    rows = [{"GT Number": "GT1", "Vinculum SKU": "S", "Status": "Cancelled"}]
    marks, _ = _plan(ff, make, rows)
    rows[0]["Order No"] = marks[0][2]          # as written tonight
    assert _plan(ff, make, rows) == ([], [])


def test_historical_stamp_marked_and_closed_rows_never_filled(ff, make):
    # Historical auto-cancel: T:AW "Cancelled", P blank — even with a Vinculum line.
    stamp = {c: "Cancelled" for c in ff.FMS_SOURCE_TAIL}
    row = {"GT Number": "GT1", "Vinculum SKU": "S", **stamp}
    marks, _ = _plan(ff, make, [row])
    assert marks == [(8, "Order No", "Cancelled", "cancel_mark")]
    writes, _ = make.plan([row], [{"Order No": "N1", "ExtOrderNo": "GT1", "SKU Code": "S"}])
    assert writes == []


def test_new_rows_marked_blank_only(ff, make):
    df = make.existing([
        {"GT Number": "GT1", "Vinculum SKU": "S", "Status": "Cancelled"},
        {"GT Number": "GT2", "Vinculum SKU": "S", "Status": "Cancelled", "Order No": "N2"},
        {"GT Number": "GT3", "Vinculum SKU": "S"},
    ])
    assert ff.apply_cancelled_order_no(df) == 1
    assert list(df["Order No"]) == ["Cancelled", "N2", ""]
    assert ff.apply_cancelled_order_no(df) == 0                    # idempotent
    assert ff.apply_cancelled_order_no(pd.DataFrame(columns=ff.OUTPUT_COLUMNS)) == 0


# ------------------------------------------------ sentinel is never a lookup key

def test_sentinel_not_an_order_key_in_matching(ff, make):
    row = make.existing([{"GT Number": "GT1", "Vinculum SKU": "S", "Order No": "Cancelled",
                          "Status": "Cancelled"}]).loc[0]
    slot = ff._Slot("existing", 0, row)
    assert slot.order == "" and slot.closed
    # Vinculum / Pick Pack data that happens to carry "CANCELLED" as an order key.
    sales = [{"Order No": "Cancelled", "ExtOrderNo": "GT1", "SKU Code": "S", "Invoice No": "I"}]
    pick = {"CANCELLED": {"Pick Date": pd.Timestamp("2026-09-01"), "Pick User": "X"}}
    writes, stats = make.plan([{"GT Number": "GT1", "Vinculum SKU": "S", "Order No": "Cancelled",
                                "Status": "Cancelled"}], sales, pick_pack=pick)
    assert writes == []
    assert stats["anchored_rows"] == 0 and stats["order_no_without_line"] == 0
    assert stats["blank_p_rows"] == 0                         # not counted as "blank Order No"


def test_sentinel_not_a_candidate_order_for_siblings(ff, make):
    # A cancelled sibling's "Cancelled" in P must not make GT1 look ambiguous or give
    # the open row "Cancelled" as its order.
    writes, stats = make.plan(
        [{"GT Number": "GT1", "Vinculum SKU": "A", "Order No": "Cancelled", "Status": "Cancelled"},
         {"GT Number": "GT1", "Vinculum SKU": "B"}],
        [], oe_rows=[{"Order No": "N7", "ExtOrderNo": "GT1"}])
    c = make.cells(writes)
    assert c[(9, "Order No")] == "N7"
    assert stats["oe_outcomes"]["ambiguous"] == 0


def test_pod_never_matches_sentinel(pod, monkeypatch):
    monkeypatch.setattr(pod, "exact_index", {
        "GT1|CANCELLED": {"date": datetime(2026, 9, 10), "url": "https://drive/x"}})
    assert pod.find_pod(pod.exact_index, {}, "GT1", "Cancelled") == (None, "")
    row = [""] * 48
    row[pod.COL_GT - 1], row[pod.COL_ORDER - 1] = "GT1", "Cancelled"
    rows_checked, matched, *_, updates = pod.build_sheet_update_requests([row])
    assert matched == 0 and updates == []


def test_reverted_row_cleared_then_filled_same_night(ff, make):
    # Same order of steps as main(): plan, clear in memory, fill, merge.
    existing = make.existing([
        {"GT Number": "GT1", "Vinculum SKU": "S", "Order No": "Cancelled"},     # reverted
        {"GT Number": "GT2", "Vinculum SKU": "S", "Order No": "Cancelled"},     # reverted, no line
        {"GT Number": "GT3", "Vinculum SKU": "S", "Status": "Cancelled"},       # to mark
    ])
    marks, clears = ff.plan_cancelled_order_no(existing)
    for r, *_ in clears:
        existing.at[r - 8, "Order No"] = ""
    writes, _, _, _ = ff.plan_backfill(
        existing, make.sales([{"Order No": "N1", "ExtOrderNo": "GT1", "SKU Code": "S"},
                              {"Order No": "N3", "ExtOrderNo": "GT3", "SKU Code": "S"}]),
        {}, {"W1": "Vadape"})
    writes, refilled = ff.merge_order_no_writes(writes, marks, clears)
    p = [(w[0], w[2], w[3]) for w in writes if w[1] == "Order No"]
    assert sorted(p) == [(8, "N1", "fill"), (9, "", "cancel_clear"),
                         (10, "Cancelled", "cancel_mark")]
    assert refilled == 1


def test_genuine_fill_wins_over_mark(ff):
    writes, _ = ff.merge_order_no_writes([(8, "Order No", "N1", "fill")],
                                         [(8, "Order No", "Cancelled", "cancel_mark")], [])
    assert writes == [(8, "Order No", "N1", "fill")]
