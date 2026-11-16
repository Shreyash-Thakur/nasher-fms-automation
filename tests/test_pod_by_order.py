"""
pod.py maps PODs by ORDER NO, not by GT + Order  —  2026-11-16.

User: "also stop maping POD because of GT number fetch it based on order id".
Before, find_pod looked up exact_index["GT|ORDER"] then ["normGT|ORDER"], so a POD file
named with a wrong GT (typo / sister order's GT) or no GT never matched. Now (mirrors
GT_FMS Pod_Upload.js gtpOrderRows_):
  - POD index keyed by normalize_order alone; col P split into its Order Nos.
  - All rows carrying the order are one normalize_gt family (GT + its R / T repunch)
    -> any file of that order is accepted.
  - Order on UNRELATED GTs -> the file's GT is the tie-breaker; a file naming none of
    them -> AMBIGUOUS, nothing written, reported.
Also: AJ "RTV" (Return to Vendor, GT_FMS dashboard RTV button) closes the line like
"Cancelled" — never overwritten with "Delivered", AP / AU / AV not filled.
"""
from datetime import datetime

import pytest

LETTERS = {"C": 3, "P": 16, "AE": 31, "AJ": 36, "AP": 42, "AU": 47, "AV": 48, "AW": 49}


def sheet_row(**cells):
    row = [""] * 49                           # A:AW
    for col, value in cells.items():
        row[LETTERS[col] - 1] = value
    return row


def rec(gt, order, day=10, name=None):
    return {"gt": gt, "order": order, "date": datetime(2026, 9, day),
            "url": f"https://drive/{gt or 'nogt'}-{order}-{day}",
            "filename": name or f"{gt} | POD | {order} | {day:02d}-09-2026 | x.pdf"}


def index(pod, *records):
    out = {}
    for r in records:
        for order in pod.split_orders(r["order"]):
            out.setdefault(order, []).append(r)
    return out


def cells(pod, rows, monkeypatch, *records):
    monkeypatch.setattr(pod, "pod_index", index(pod, *records))
    *_, updates = pod.build_sheet_update_requests(rows)
    return {u["range"].split("!")[1]: u["values"][0][0] for u in updates}


def delivered_rows(c):
    return sorted(int(k[2:]) for k, v in c.items() if k.startswith("AJ") and v == "Delivered")


# ---------------------------------------------------------------- file names

def test_structured_name_without_gt_is_parsed(pod):
    p = pod.parse_structured_pod_filename("POD | N89705 | 05-11-2026 | x.pdf")
    assert p["gt"] == "" and p["order"] == "N89705"
    # GT_FMS gtpPodFileName_ with a blank GT
    p = pod.parse_structured_pod_filename(" | POD | N89705 | 05-11-2026 | x.pdf")
    assert p["gt"] == "" and p["order"] == "N89705"
    # still not a POD name without the marker
    assert pod.parse_structured_pod_filename("Invoice | N89705 | 05-11-2026") is None


def test_split_orders(pod):
    assert pod.split_orders(" n012271, N461768 ") == ["N012271", "N461768"]
    assert pod.split_orders("N1;N2 N1") == ["N1", "N2"]
    assert pod.split_orders("Cancelled") == [] and pod.split_orders("RTV") == []
    assert pod.split_orders("") == []


# ---------------------------------------------------------------- order-first mapping

def test_wrong_gt_in_file_name_still_matches_by_order(pod, monkeypatch):
    rows = [sheet_row(C="GT101026A2", P="N89705"), sheet_row(C="GT101026A2", P="N2")]
    c = cells(pod, rows, monkeypatch, rec("GT999999A9", "N89705"))
    assert delivered_rows(c) == [8]
    assert c["AV8"] == "10-Sep-2026" and c["AP8"] == "10-Sep-2026"
    assert "https://drive/GT999999A9-N89705-10" in c["AU8"]


def test_missing_gt_in_file_name_matches_by_order(pod, monkeypatch):
    rows = [sheet_row(C="GT101026A2", P="N89705"), sheet_row(C="GT111026A1", P="N3")]
    c = cells(pod, rows, monkeypatch, rec("", "N89705"))
    assert delivered_rows(c) == [8]


def test_order_listed_with_others_in_col_p(pod, monkeypatch):
    rows = [sheet_row(C="GT101026A2", P="N1, n89705 ")]
    c = cells(pod, rows, monkeypatch, rec("GT101026A2", "N89705"))
    assert delivered_rows(c) == [8]


def test_gt_and_repunch_r_share_an_order_both_written(pod, monkeypatch):
    rows = [sheet_row(C="GT030426A3", P="N100"), sheet_row(C="GT030426A3R", P="N100"),
            sheet_row(C="GT030426A3", P="N200")]
    c = cells(pod, rows, monkeypatch, rec("GT030426A3R", "N100"))
    assert delivered_rows(c) == [8, 9]
    assert pod.ambiguous_orders == {}


def test_unrelated_gts_without_tie_breaker_are_ambiguous(pod, monkeypatch):
    rows = [sheet_row(C="GT030426A3", P="N100"), sheet_row(C="GT050526A1", P="N100")]
    for file_gt in ("", "GT777777A1"):
        c = cells(pod, rows, monkeypatch, rec(file_gt, "N100"))
        assert c == {}, file_gt
        assert set(pod.ambiguous_orders) == {"N100"}
        assert pod.ambiguous_orders["N100"][0] == ["GT030426A3", "GT050526A1"]


def test_unrelated_gts_file_gt_breaks_the_tie(pod, monkeypatch):
    rows = [sheet_row(C="GT030426A3", P="N100"), sheet_row(C="GT050526A1", P="N100")]
    c = cells(pod, rows, monkeypatch, rec("GT050526A1R", "N100"))
    assert delivered_rows(c) == [9]
    assert pod.ambiguous_orders == {}
    rows_checked, matched, missed, by_order, tie_break, *_ = \
        pod.build_sheet_update_requests(rows)
    assert (matched, by_order, tie_break) == (1, 0, 1)


def test_latest_pod_of_the_order_wins(pod, monkeypatch):
    rows = [sheet_row(C="GT1", P="N1")]
    c = cells(pod, rows, monkeypatch, rec("GT1", "N1", day=5), rec("GT9", "N1", day=12))
    assert c["AV8"] == "12-Sep-2026"


def test_order_not_in_pod_index_no_match(pod, monkeypatch):
    rows = [sheet_row(C="GT1", P="N1")]
    assert cells(pod, rows, monkeypatch, rec("GT1", "N999")) == {}


def test_blanks_only_writes_kept(pod, monkeypatch):
    rows = [sheet_row(C="GT1", P="N1", AP="01-Sep-2026", AU="=HYPERLINK(\"x\",\"y\")",
                      AV="01-Sep-2026")]
    c = cells(pod, rows, monkeypatch, rec("GT9", "N1"))
    assert c == {"AJ8": "Delivered"}


# ---------------------------------------------------------------- RTV

def test_rtv_is_a_closed_marker(pod):
    assert "rtv" in pod.CLOSED_MARKERS
    assert pod.row_is_cancelled(sheet_row(C="GT1", P="N1", AJ="RTV"))
    assert pod.row_is_cancelled(sheet_row(C="GT1", P="N1", AJ=" rtv "))


def test_aj_rtv_untouched_when_pod_matches(pod, monkeypatch):
    rows = [sheet_row(C="GT1", P="N1", AJ="RTV", AW="27/10/GJ1/302")]
    assert cells(pod, rows, monkeypatch, rec("GT1", "N1")) == {}
    assert pod.closed_skipped == 1


def test_aj_rtv_untouched_without_pod(pod, monkeypatch):
    rows = [sheet_row(C="GT1", P="N1", AJ="RTV", AE="AWB12345678", AW="27/10/GJ1/302")]
    assert cells(pod, rows, monkeypatch) == {}


def test_full_flow_treats_rtv_as_closed(ff):
    import pandas as pd
    row = pd.Series({c: "" for c in ff.OUTPUT_COLUMNS})
    row["Order Delivery Status"] = "RTV"
    assert ff.row_is_closed(row)
    assert ff.split_order_nos("RTV") == ()
