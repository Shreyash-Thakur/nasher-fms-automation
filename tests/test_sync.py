"""
2026-10-09 "it's not fetching data correctly — a lot of rows are empty".

  1. GT Final R/T/F approval suffix tolerated on either side, never across GTs (A1 vs A10)
  2. rows that already have an Order No get BLANK cells refreshed (never overwritten)
  3. new SKU lines for a GT already in Full Flow are appended, once (idempotent)
  4. historical "no match -> Cancelled" is off; suspect cancellations are only reported
"""
import pandas as pd


def fms(ff, rows):
    out = []
    for r in rows:
        row = {c: "" for c in ff.FMS_COLUMNS}
        row.update(r)
        out.append(row)
    return pd.DataFrame(out, columns=ff.FMS_COLUMNS)


# ---------------------------------------------------------------- 1. suffix matching

def test_gt_base(ff):
    assert ff.gt_base("GT081026A1R") == "GT081026A1"
    assert ff.gt_base("gt081026a1t ") == "GT081026A1"
    assert ff.gt_base("GT081026A10F") == "GT081026A10"
    assert ff.gt_base("GT081026A1") == "GT081026A1"
    assert ff.gt_base("GT081026A10") == "GT081026A10"   # never trims a digit
    assert ff.gt_base("GT081026AR") == "GT081026AR"     # suffix only after a digit


def test_suffix_on_fms_side(make):
    writes, stats = make.plan(
        [{"GT Number": "GT081026A1R", "Vinculum SKU": "S"}],
        [{"Order No": "N1", "ExtOrderNo": "GT081026A1", "SKU Code": "S"}])
    assert make.cells(writes)[(8, "Order No")] == "N1"
    assert stats["tiers"][2] == 1


def test_suffix_on_vinculum_side(make):
    writes, stats = make.plan(
        [{"GT Number": "GT081026A1", "Vinculum SKU": "S"}],
        [{"Order No": "N1", "ExtOrderNo": "GT081026A1T", "SKU Code": "S"}])
    assert make.cells(writes)[(8, "Order No")] == "N1"
    assert stats["tiers"][2] == 1


def test_exact_beats_suffix(make):
    # Both R and T rows for the same SKU; Vinculum has both, exact pairs win.
    writes, stats = make.plan(
        [{"GT Number": "GT081026A1R", "Vinculum SKU": "S"},
         {"GT Number": "GT081026A1T", "Vinculum SKU": "S"}],
        [{"Order No": "NT", "ExtOrderNo": "GT081026A1T", "SKU Code": "S"},
         {"Order No": "NR", "ExtOrderNo": "GT081026A1R", "SKU Code": "S"}])
    c = make.cells(writes)
    assert c[(8, "Order No")] == "NR" and c[(9, "Order No")] == "NT"
    assert stats["tiers"][1] == 2


def test_a1_never_matches_a10(make):
    writes, stats = make.plan(
        [{"GT Number": "GT081026A1", "Vinculum SKU": "S"},
         {"GT Number": "GT081026A10R", "Vinculum SKU": "S"}],
        [{"Order No": "N10", "ExtOrderNo": "GT081026A10", "SKU Code": "S"}])
    c = make.cells(writes)
    assert (8, "Order No") not in c          # A1 must not take A10's line
    assert c[(9, "Order No")] == "N10"
    writes, _ = make.plan(
        [{"GT Number": "GT081026A10", "Vinculum SKU": "S"}],
        [{"Order No": "N1", "ExtOrderNo": "GT081026A1R", "SKU Code": "S"}])
    assert writes == []


def test_alternate_sku_field_same_gt_only(make):
    # FMS SKU Code (E) equals the Vinculum SKU Code when D does not; same GT base only.
    writes, stats = make.plan(
        [{"GT Number": "GT1R", "Vinculum SKU": "OLDCODE", "SKU Code": "NEWCODE"},
         {"GT Number": "GT2", "Vinculum SKU": "OLDCODE", "SKU Code": "NEWCODE"}],
        [{"Order No": "N1", "ExtOrderNo": "GT1", "SKU Code": "NEWCODE"}])
    c = make.cells(writes)
    assert c[(8, "Order No")] == "N1" and (9, "Order No") not in c
    assert stats["tiers"][3] == 1


# ---------------------------------------------------------------- 2. refresh

def test_refresh_fills_blank_cells_of_row_with_order_no(make):
    # Punched earlier (P, Q, S filled); invoice + pick arrived later in Vinculum.
    writes, stats = make.plan(
        [{"GT Number": "GT1R", "Vinculum SKU": "S", "Order No": "N1", "Team": "Gauri",
          "Actual": "05-Sep-2026 10:00:00", "Invoice Amount": 999}],
        [{"Order No": "N1", "ExtOrderNo": "GT1", "SKU Code": "S", "Invoice No": "INV/9",
          "Pick Qty": 4, "Total": 1000}],
        pick_pack={"N1": {"Pick Date": pd.Timestamp("2026-09-06 09:00"), "Pick User": "Bappa"}})
    c = make.cells(writes)
    assert c[(8, "Invoice No")] == "INV/9"
    assert c[(8, "Dispatched Qty")] == 4
    assert c[(8, "Warehouse")] == "Vadape"
    assert c[(8, "Warehouse Team")] == "Bappa" and (8, "Actual.1") in c
    for kept in ("Order No", "Team", "Actual", "Invoice Amount"):
        assert (8, kept) not in c, kept           # non-blank: never overwritten
    assert {w[3] for w in writes} == {"refresh"}
    assert stats["refreshed_rows"] == 1


def test_refresh_uses_the_row_own_order_no(make):
    writes, _ = make.plan(
        [{"GT Number": "GT1", "Vinculum SKU": "S", "Order No": "N2"}],
        [{"Order No": "N1", "ExtOrderNo": "GT1", "SKU Code": "S", "Invoice No": "INV/1"},
         {"Order No": "N2", "ExtOrderNo": "GT1", "SKU Code": "S", "Invoice No": "INV/2"}])
    assert make.cells(writes)[(8, "Invoice No")] == "INV/2"


def test_zero_pick_qty_not_written(make):
    writes, _ = make.plan([{"GT Number": "GT1", "Vinculum SKU": "S"}],
                          [{"Order No": "N1", "ExtOrderNo": "GT1", "SKU Code": "S", "Pick Qty": 0}])
    assert (8, "Dispatched Qty") not in make.cells(writes)


def test_closed_rows_never_written(make):
    rows = [
        {"GT Number": "GT1", "Vinculum SKU": "S", "Order No": "N1",
         "Order Delivery Status": "Virtual Invoice"},                 # punched virtual: AJ only
        {"GT Number": "GT2", "Vinculum SKU": "S", "Order No": "Virtual Invoice",
         "Team": "Virtual Invoice"},                                   # open virtual: P:AW
        {"GT Number": "GT3", "Vinculum SKU": "S", "Order No": "Cancelled"},
        {"GT Number": "GT4", "Vinculum SKU": "S", "Status": "OOS"},
    ]
    sales = [{"Order No": f"N{i}", "ExtOrderNo": f"GT{i}", "SKU Code": "S", "Invoice No": "I"}
             for i in range(1, 5)]
    writes, _ = make.plan(rows, sales)
    assert writes == []


# ---------------------------------------------------------------- 3. new lines

def test_new_line_for_existing_gt_appended_once(ff, make):
    existing = make.existing([{"GT Number": "GT1", "Vinculum SKU": "S1", "Order No": "N1"},
                              {"GT Number": "GT9", "Vinculum SKU": "S", "Status": "Cancelled"}])
    source = fms(ff, [{"GT Number": "GT1", "Vinculum SKU": "S1"},
                      {"GT Number": "GT1", "Vinculum SKU": "S2", "Form Received Time": "08-Oct-2026 10:00:00"},
                      {"GT Number": "GT9", "Vinculum SKU": "S"},          # cancelled line: never re-added
                      {"GT Number": "GT5", "Vinculum SKU": "S"}])
    new_gt, new_line = ff.find_new_fms_rows(source, existing)
    assert list(source.loc[new_line, "Vinculum SKU"]) == ["S2"]
    assert list(source.loc[new_gt, "GT Number"]) == ["GT5"]

    to_append = source.loc[new_gt | new_line]
    lines = make.sales([{"Order No": "N1", "ExtOrderNo": "GT1", "SKU Code": "S1"},
                        {"Order No": "N2", "ExtOrderNo": "GT1R", "SKU Code": "S2"}])
    slots = [ff._Slot("existing", i, r) for i, r in existing.iterrows()]
    new_slots = [ff._Slot("new", p, r) for p, (_, r) in enumerate(to_append.iterrows())]
    ff.allocate_lines(slots + new_slots, lines)
    assigned = [[lines.loc[l] for l in s.assigned] for s in new_slots]
    built = ff.build_full_flow(to_append, assigned, {}, {"W1": "Vadape"})
    assert len(built) == 2
    s2 = built[built["Vinculum SKU"] == "S2"].iloc[0]
    assert s2["Order No"] == "N2"                          # suffix-tolerant on the append path too

    after = pd.concat([existing, built], ignore_index=True)
    new_gt, new_line = ff.find_new_fms_rows(source, after)
    assert not new_gt.any() and not new_line.any()          # idempotent: nothing to add next night


def test_new_row_does_not_take_line_used_by_existing_row(ff, make):
    existing = make.existing([{"GT Number": "GT1R", "Vinculum SKU": "S", "Order No": "N1"}])
    to_append = fms(ff, [{"GT Number": "GT1T", "Vinculum SKU": "S"}])
    lines = make.sales([{"Order No": "N1", "ExtOrderNo": "GT1", "SKU Code": "S"}])
    slots = [ff._Slot("existing", i, r) for i, r in existing.iterrows()]
    new_slots = [ff._Slot("new", 0, to_append.iloc[0])]
    ff.allocate_lines(slots + new_slots, lines)
    assert new_slots[0].assigned == []


def test_cross_tab_duplicates_dropped(make):
    lines = make.sales([
        {"Order No": "N1", "ExtOrderNo": "GT1", "SKU Code": "S", "_TAB": "Sep-26"},
        {"Order No": "N1", "ExtOrderNo": "GT1", "SKU Code": "S", "_TAB": "Sept-26", "Invoice No": "INV"},
        {"Order No": "N2", "ExtOrderNo": "GT2", "SKU Code": "S", "_TAB": "Sep-26"},
        {"Order No": "N2", "ExtOrderNo": "GT2", "SKU Code": "S", "_TAB": "Sep-26"},   # same tab: kept
    ])
    assert len(lines) == 3
    assert lines.loc[lines["_ORDER"] == "N1", "Invoice No"].tolist() == ["INV"]


# ---------------------------------------------------------------- 4. cancellations

def test_historical_cancellation_off(ff):
    assert ff.ENABLE_HISTORICAL_CANCELLATION == "NO"
    source = fms(ff, [{"GT Number": "GT1", "Vinculum SKU": "S",
                       "Form Received Time": "01-Jul-2026 10:00:00"}])
    sales = pd.DataFrame([{"ExtOrderNo": "GT9", "SKU Code": "S"}])
    out = ff.apply_historical_cancellation(source, sales)
    assert "Cancelled" not in out.astype(str).values


def test_suspect_cancellation_report(ff, make, tmp_path):
    existing = make.existing([
        {"GT Number": "GT1R", "Vinculum SKU": "S", "Status": "Cancelled", "Invoice No": "Cancelled"},
        {"GT Number": "GT2", "Vinculum SKU": "S", "Order No": "Cancelled"},
        {"GT Number": "GT3", "Vinculum SKU": "S", "Status": "Cancelled"},     # no Vinculum line
    ])
    lines = make.sales([{"Order No": "N1", "ExtOrderNo": "GT1", "SKU Code": "S", "Invoice No": "INV"},
                        {"Order No": "N2", "ExtOrderNo": "GT2", "SKU Code": "S"}])
    path = tmp_path / "r" / "suspect.csv"
    report = ff.report_suspect_cancellations(existing, lines, path=str(path))
    assert list(report["sheet_row"]) == [8, 9]
    assert report.iloc[0]["pattern"].startswith("historical") and report.iloc[0]["invoiced"] == "yes"
    assert report.iloc[1]["pattern"].startswith("dashboard")
    assert path.exists()
    assert (existing == make.existing([
        {"GT Number": "GT1R", "Vinculum SKU": "S", "Status": "Cancelled", "Invoice No": "Cancelled"},
        {"GT Number": "GT2", "Vinculum SKU": "S", "Order No": "Cancelled"},
        {"GT Number": "GT3", "Vinculum SKU": "S", "Status": "Cancelled"}])).all().all()  # read-only
