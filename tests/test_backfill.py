"""
Offline tests for the Full Flow FMS backfill (blank Order No -> Order Enquiry match).

Why: the nightly job only processed NEW GTs. A GT appended before its Vinculum order
existed kept a blank Order No (P) forever, so P/Q/S/AS/AW/AX/AY stayed empty and the POD
sync (GT + Order No) could never match it — Sep-2026: 40 of 1,315 rows had an Order No.

Contract of plan_backfill():
  - only rows whose Order No is blank and that are not Cancelled
  - same GT Number + Vinculum SKU key as the normal build
  - Order Enquiry lines already used by a filled row of the same key are not reused
  - only BLANK cells are written; nothing is ever overwritten
  - more Vinculum lines than blank rows -> extras reported, not inserted

Run: python tests/test_backfill.py
"""
import os
import sys
import types
import importlib.util

import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPT = os.path.join(ROOT, "scripts", "gt_full_flow_fms_yes_no_all_or_latest.py")


def load_script():
    """Import the script with every Google client stubbed (it connects at import time)."""
    def mod(name, **attrs):
        m = types.ModuleType(name)
        for k, v in attrs.items():
            setattr(m, k, v)
        sys.modules[name] = m
        return m

    class FakeCreds:
        @staticmethod
        def from_service_account_info(info, scopes=None):
            return object()

    mod("google")
    mod("google.oauth2", service_account=types.SimpleNamespace(Credentials=FakeCreds))
    mod("google.oauth2.service_account", Credentials=FakeCreds)
    mod("google.oauth2.credentials", Credentials=object)
    mod("google.auth")
    mod("google.auth.transport")
    mod("google.auth.transport.requests", Request=object)
    mod("googleapiclient")
    mod("googleapiclient.discovery", build=lambda *a, **k: object())
    mod("googleapiclient.http", MediaIoBaseDownload=object)
    os.environ["GCP_SA_KEY"] = '{"type": "service_account"}'

    spec = importlib.util.spec_from_file_location("fullflow", SCRIPT)
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m


ff = load_script()
COLS = ff.OUTPUT_COLUMNS


def existing(rows):
    """rows: list of dicts with output-column names; everything else blank."""
    out = []
    for r in rows:
        row = {c: "" for c in COLS}
        row.update(r)
        out.append(row)
    return pd.DataFrame(out, columns=COLS)


def sales(rows):
    base = {"Order No": "", "ExtOrderNo": "", "Order Date": "2026-09-05 10:00:00", "SKU Code": "",
            "Created by": "Gauri", "Invoice No": "", "Originating warehouse": "W1", "Pick Qty": 5,
            "Total": 1000}
    return pd.DataFrame([{**base, **r} for r in rows])


def plan(existing_rows, sales_rows, warehouse_mapping=None, pick_pack=None):
    sales_groups, _ = ff.build_sales_lookup(sales(sales_rows))
    return ff.plan_backfill(existing(existing_rows), sales_groups, pick_pack or {},
                            warehouse_mapping or {"W1": "Vadape"})


def cells(writes):
    return {(w[0], w[1]): w[2] for w in writes}


PASSED = 0
FAILED = 0


def test(name, fn):
    global PASSED, FAILED
    try:
        fn()
        PASSED += 1
        print("  ok   " + name)
    except AssertionError as e:
        FAILED += 1
        print("  FAIL " + name + "\n       " + str(e))


print("Full Flow backfill")


def t_basic():
    writes, stats = plan(
        [{"GT Number": "GT010926A1", "Vinculum SKU": "SKU1", "Order Quantity": 5}],
        [{"Order No": "N100", "ExtOrderNo": "GT010926A1", "SKU Code": "SKU1", "Invoice No": "INV/1"}])
    c = cells(writes)
    assert c[(8, "Order No")] == "N100", c
    assert c[(8, "Team")] == "Gauri"
    assert c[(8, "Invoice No")] == "INV/1"
    assert c[(8, "Invoice Amount")] == 1000
    assert c[(8, "Warehouse")] == "Vadape"
    assert c[(8, "Dispatched Qty")] == 5
    assert (8, "Actual") in c
    assert stats["rows_filled"] == 1


test("blank Order No + matching Vinculum line -> P,Q,S,AS,AW,AX,AY filled on sheet row 8", t_basic)


def t_never_overwrite():
    writes, _ = plan(
        [{"GT Number": "GT1", "Vinculum SKU": "S", "Team": "Manual", "Invoice No": "KEEP"}],
        [{"Order No": "N1", "ExtOrderNo": "GT1", "SKU Code": "S", "Invoice No": "OTHER"}])
    c = cells(writes)
    assert (8, "Team") not in c and (8, "Invoice No") not in c, c
    assert c[(8, "Order No")] == "N1"


test("cells that already have a value are never overwritten", t_never_overwrite)


def t_filled_rows_untouched():
    writes, stats = plan(
        [{"GT Number": "GT1", "Vinculum SKU": "S", "Order No": "N1"}],
        [{"Order No": "N1", "ExtOrderNo": "GT1", "SKU Code": "S"}])
    assert writes == [], writes
    assert stats["rows_considered"] == 0


test("rows that already have an Order No are not touched", t_filled_rows_untouched)


def t_cancelled():
    writes, _ = plan(
        [{"GT Number": "GT1", "Vinculum SKU": "S", "Status": "Cancelled"},
         {"GT Number": "GT2", "Vinculum SKU": "S", "Order Quantity": "Cancelled"}],
        [{"Order No": "N1", "ExtOrderNo": "GT1", "SKU Code": "S"},
         {"Order No": "N2", "ExtOrderNo": "GT2", "SKU Code": "S"}])
    assert writes == [], writes


test("Cancelled rows are skipped", t_cancelled)


def t_no_match():
    writes, stats = plan([{"GT Number": "GT1", "Vinculum SKU": "S"}],
                         [{"Order No": "N9", "ExtOrderNo": "GT9", "SKU Code": "S"}])
    assert writes == []
    assert stats["rows_still_unmatched"] == 1


test("no Order Enquiry line yet -> nothing written, counted as still unmatched", t_no_match)


def t_case_space():
    writes, _ = plan([{"GT Number": " gt1 ", "Vinculum SKU": "sku 1"}],
                     [{"Order No": "N1", "ExtOrderNo": "GT1", "SKU Code": "SKU1"}])
    assert cells(writes)[(8, "Order No")] == "N1"


test("match key ignores case and spaces (same normalize_value as the build)", t_case_space)


def t_used_line_not_reused():
    # Row 8 already holds N1; row 9 (same GT+SKU) is blank -> gets N2, never N1 again.
    writes, _ = plan(
        [{"GT Number": "GT1", "Vinculum SKU": "S", "Order No": "N1"},
         {"GT Number": "GT1", "Vinculum SKU": "S"}],
        [{"Order No": "N1", "ExtOrderNo": "GT1", "SKU Code": "S"},
         {"Order No": "N2", "ExtOrderNo": "GT1", "SKU Code": "S"}])
    c = cells(writes)
    assert c[(9, "Order No")] == "N2", c
    assert all(w[0] == 9 for w in writes)


test("a Vinculum line already used by a filled row is not reused", t_used_line_not_reused)


def t_extra_lines_reported():
    writes, stats = plan(
        [{"GT Number": "GT1", "Vinculum SKU": "S"}],
        [{"Order No": "N1", "ExtOrderNo": "GT1", "SKU Code": "S"},
         {"Order No": "N2", "ExtOrderNo": "GT1", "SKU Code": "S"}])
    assert cells(writes)[(8, "Order No")] == "N1"
    assert stats["unplaced_sales_lines"] == 1, stats
    assert "GT1||S" in stats["unplaced_keys"]


test("more Vinculum lines than blank rows -> first fills, extras reported (not inserted)", t_extra_lines_reported)


def t_pick_pack():
    writes, _ = plan(
        [{"GT Number": "GT1", "Vinculum SKU": "S"}],
        [{"Order No": "N1", "ExtOrderNo": "GT1", "SKU Code": "S"}],
        pick_pack={"N1": {"Pick Date": pd.Timestamp("2026-09-06 09:00"), "Pick User": "Bappa"}})
    c = cells(writes)
    assert c[(8, "Warehouse Team")] == "Bappa"
    assert (8, "Actual.1") in c


test("Pick-Pack bridge fills Warehouse Team (V) and Actual.1 (X) when blank", t_pick_pack)


def t_rows_map_to_sheet_rows():
    writes, _ = plan(
        [{"GT Number": "GTX", "Vinculum SKU": "S", "Order No": "N0"},
         {"GT Number": "GTX", "Vinculum SKU": "Q"},
         {"GT Number": "GTY", "Vinculum SKU": "S"}],
        [{"Order No": "N7", "ExtOrderNo": "GTY", "SKU Code": "S"}])
    assert {w[0] for w in writes} == {10}, writes


test("dataframe index i maps to sheet row 8 + i", t_rows_map_to_sheet_rows)


def t_a1():
    assert ff.backfill_a1(8, "Order No") == "'Full Flow FMS'!P8"
    assert ff.backfill_a1(12, "Invoice Amount") == "'Full Flow FMS'!AY12"
    assert ff.backfill_a1(9, "Actual.1") == "'Full Flow FMS'!X9"


test("A1 addresses follow OUTPUT_COLUMNS (P, AY, X)", t_a1)

print(f"{PASSED} passed, {FAILED} failed")
if __name__ == "__main__":
    sys.exit(1 if FAILED else 0)
