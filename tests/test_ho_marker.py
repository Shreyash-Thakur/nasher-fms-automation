"""Head Office (HO) lines (2026-11-16): AJ "Head Office (HO)" is a closed marker for the nightly jobs.

GeneralTrade_FMS GtHoMark.js marks lines of Vinculum orders located at H01 with "Head Office (HO)" in AJ and the dashboards hide
them. The nightly job must leave such a line exactly as marked (like Cancelled / Virtual Invoice / OOS), and a POD match must
never turn it into "Delivered" (the pod script treats it as closed too).
"""
from datetime import datetime

import pytest


def test_ho_is_a_closed_marker(ff, pod):
    assert "head office (ho)" in ff.CLOSED_MARKERS
    assert "head office (ho)" in pod.CLOSED_MARKERS


def test_row_with_ho_in_aj_is_closed(ff):
    assert ff.row_is_closed({"Order Delivery Status": "Head Office (HO)"}) is True
    assert ff.row_is_closed({"Order Delivery Status": " head office (ho) "}) is True
    assert ff.row_is_closed({"Order Delivery Status": "Delivered"}) is False


def test_ho_row_never_written_by_backfill(make):
    rows = [{"GT Number": "GT1", "Vinculum SKU": "S", "Order No": "N461001", "Order Delivery Status": "Head Office (HO)"}]
    sales = [{"Order No": "N461001", "ExtOrderNo": "GT1", "SKU Code": "S", "Invoice No": "I/1"}]
    writes, _ = make.plan(rows, sales)
    assert writes == []


def test_pod_never_turns_ho_into_delivered(pod, monkeypatch):
    monkeypatch.setattr(pod, "exact_index", {"GT1|N461001": {"date": datetime(2026, 9, 10), "url": "https://drive/x"}})
    row = [""] * 48
    row[pod.COL_GT - 1] = "GT1"
    row[pod.COL_ORDER - 1] = "N461001"
    row[pod.COL_AJ - 1] = "Head Office (HO)"
    *_, updates = pod.build_sheet_update_requests([row])
    assert updates == []
