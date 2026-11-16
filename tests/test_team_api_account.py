"""Team (Q) is never the Vinculum API account (16-Nov-2026): "Diwakar" / "finance.automation" are not people."""
import pandas as pd
import pytest


@pytest.mark.parametrize("name", ["Diwakar", " diwakar ", "DIWAKAR", "finance.automation",
                                  "finance.automation@nashermiles.com"])
def test_api_account_is_not_a_team(ff, name):
    assert ff.team_from_creator(name) is None


def test_real_person_kept(ff):
    assert ff.team_from_creator("Gauri") == "Gauri"
    assert ff.team_from_creator("Diwakar Rao") == "Diwakar Rao"
    assert ff.team_from_creator(None) is None and ff.team_from_creator("") is None


def test_invoiced_line_with_api_creator_leaves_team_blank(make):
    writes, _ = make.plan(
        [{"GT Number": "GT1", "Vinculum SKU": "S", "Order No": "N1"}],
        [{"Order No": "N1", "ExtOrderNo": "GT1", "SKU Code": "S", "Created by": "Diwakar"}])
    c = make.cells(writes)
    assert (8, "Team") not in c


def test_order_enquiry_with_api_creator_leaves_team_blank(ff, make):
    existing = make.existing([{"GT Number": "GT1", "Vinculum SKU": "S", "Order No": "N1"}])
    oe = ff.prepare_oe_orders(pd.DataFrame([{
        "Order No": "N1", "ExtOrderNo": "GT1", "Order Date": "12/10/2026 10:00 AM",
        "Status": "Confirmed", "Created by": "Diwakar"}]))
    writes, *_ = ff.plan_backfill(existing, make.sales([]), {}, {}, oe_orders=oe)
    c = make.cells(writes)
    assert (8, "Team") not in c and (8, "Actual") in c
