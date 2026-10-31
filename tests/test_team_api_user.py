"""
Q (Team) never shows the Vinculum API user  -  2026-11-01.

Orders punched through the auto-punch portal reach Vinculum via the API key of user
"Diwakar", so their Order Enquiry / Sales "Created by" is "Diwakar" and Q showed
Diwakar on ~345 Oct rows. The portal now writes the real ops name into Q at punch time
(and backfills old rows). Contract for every Q write path of the nightly job:
  - an API-user name (API_USER_NAMES, case/space-insensitive) is "no team info": never
    written into Q
  - Q is written only when blank or holding an API-user name; a real name (the portal's)
    is never overwritten
  - an existing API-user Q is never blanked by this job (the portal backfill replaces it)
  - the fill-rate table counts an API-user Q as blank
"""
import pandas as pd
import pytest


@pytest.mark.parametrize("value", ["Diwakar", "diwakar", " DIWAKAR ", "Diwakar\xa0", "di wakar"])
def test_api_user_detected(ff, value):
    assert ff.is_api_user(value)


@pytest.mark.parametrize("value", ["", None, "Gauri", "Diwakar Singh", float("nan")])
def test_not_api_user(ff, value):
    assert not ff.is_api_user(value)


# ------------------------------------------------ existing rows: Sales Master (fill/refresh)

def test_sales_fill_skips_api_user(make):
    writes, _ = make.plan([{"GT Number": "GT1", "Vinculum SKU": "S"}],
                          [{"Order No": "N1", "ExtOrderNo": "GT1", "SKU Code": "S",
                            "Created by": "Diwakar"}])
    c = make.cells(writes)
    assert c[(8, "Order No")] == "N1"                     # rest of the row still fills
    assert (8, "Team") not in c


def test_sales_refresh_keeps_portal_name(make):
    writes, _ = make.plan([{"GT Number": "GT1", "Vinculum SKU": "S", "Order No": "N1",
                            "Team": "Saloni"}],
                          [{"Order No": "N1", "ExtOrderNo": "GT1", "SKU Code": "S",
                            "Created by": "Gauri", "Invoice No": "INV1"}])
    c = make.cells(writes)
    assert c[(8, "Invoice No")] == "INV1" and (8, "Team") not in c


def test_sales_refresh_replaces_api_user_q_with_real_name(make):
    writes, _ = make.plan([{"GT Number": "GT1", "Vinculum SKU": "S", "Order No": "N1",
                            "Team": " diwakar "}],
                          [{"Order No": "N1", "ExtOrderNo": "GT1", "SKU Code": "S",
                            "Created by": "Gauri"}])
    assert make.cells(writes)[(8, "Team")] == "Gauri"


def test_api_user_q_not_blanked_or_rewritten(make):
    writes, _ = make.plan([{"GT Number": "GT1", "Vinculum SKU": "S", "Order No": "N1",
                            "Team": "Diwakar"}],
                          [{"Order No": "N1", "ExtOrderNo": "GT1", "SKU Code": "S",
                            "Created by": "Diwakar"}])
    assert (8, "Team") not in make.cells(writes)


# ------------------------------------------------ existing rows: daily Order Enquiry

def test_oe_fill_skips_api_user(make):
    writes, _ = make.plan([{"GT Number": "GT1", "Vinculum SKU": "S"}], [],
                          oe_rows=[{"Order No": "N1", "ExtOrderNo": "GT1", "Created by": "DIWAKAR"}])
    c = make.cells(writes)
    assert c[(8, "Order No")] == "N1" and (8, "Team") not in c


def test_order_no_row_skips_api_user_and_keeps_portal_name(make):
    oe = [{"Order No": "N1", "ExtOrderNo": "GT1", "Created by": "Diwakar"},
          {"Order No": "N2", "ExtOrderNo": "GT2", "Created by": "Gauri"}]
    writes, _ = make.plan([{"GT Number": "GT1", "Vinculum SKU": "S", "Order No": "N1"},
                           {"GT Number": "GT2", "Vinculum SKU": "S", "Order No": "N2",
                            "Team": "Saloni"}], [], oe_rows=oe)
    c = make.cells(writes)
    assert (8, "Actual") in c and (8, "Team") not in c
    assert (9, "Team") not in c


def test_order_no_row_replaces_api_user_q(make):
    writes, _ = make.plan([{"GT Number": "GT1", "Vinculum SKU": "S", "Order No": "N1",
                            "Team": "Diwakar"}], [],
                          oe_rows=[{"Order No": "N1", "ExtOrderNo": "GT1", "Created by": "Gauri"}])
    assert make.cells(writes)[(8, "Team")] == "Gauri"


# ------------------------------------------------ append path (build_full_flow)

def _fms(ff, n=1):
    fms = pd.DataFrame([{c: "" for c in ff.FMS_COLUMNS}] * n, columns=ff.FMS_COLUMNS)
    for i in range(n):
        fms.loc[i, ["GT Number", "Vinculum SKU"]] = [f"GT{i + 1}", "S"]
    return fms


def test_append_sales_line_skips_api_user(ff, make):
    lines = make.sales([{"Order No": "N1", "ExtOrderNo": "GT1", "SKU Code": "S",
                         "Created by": "Diwakar"},
                        {"Order No": "N2", "ExtOrderNo": "GT2", "SKU Code": "S",
                         "Created by": "Gauri"}])
    out = ff.build_full_flow(_fms(ff, 2), [[lines.iloc[0]], [lines.iloc[1]]], {},
                             {"W1": "Vadape"}, [None, None])
    assert list(out["Order No"]) == ["N1", "N2"]
    assert list(out["Team"]) == ["", "Gauri"]


def test_append_oe_record_skips_api_user(ff):
    rec = {"order_no": "N9", "order_date": pd.Timestamp("2026-10-09 16:00"),
           "team": "Diwakar", "status": "Confirmed", "ext_order_no": "GT1"}
    out = ff.build_full_flow(_fms(ff), [[]], {}, {}, oe_records=[rec])
    assert out.iloc[0]["Order No"] == "N9" and out.iloc[0]["Team"] == ""


# ------------------------------------------------ fill-rate report

def test_fill_rate_counts_api_user_as_blank(ff, make, capsys):
    df = make.existing([{"GT Number": "GT1", "Form Received Time": "2026-10-05",
                         "Team": "Diwakar"},
                        {"GT Number": "GT2", "Form Received Time": "2026-10-05",
                         "Team": "Saloni"}])
    ff._print_fill_rates(df)
    out = capsys.readouterr().out.splitlines()
    header = out[-2].split()
    row = out[-1].split()
    team_col = ff.backfill_a1(8, "Team").split("!")[1][:-1]       # "Q"
    assert row[1] == "2" and row[header.index(team_col)] == "1"
