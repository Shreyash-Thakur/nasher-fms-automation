"""Shared test setup: import the nightly scripts offline (Google clients stubbed —
yes_no.py connects at import time) and build small Full Flow / Vinculum frames."""
import importlib.util
import os
import sys
import types

import pandas as pd
import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPTS = os.path.join(ROOT, "scripts")


def _stub_google():
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

        from_service_account_file = from_service_account_info

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
    mod("google_auth_httplib2", AuthorizedHttp=object)
    mod("httplib2", Http=object)
    os.environ["GCP_SA_KEY"] = '{"type": "service_account"}'


def _load(filename, name):
    _stub_google()
    spec = importlib.util.spec_from_file_location(name, os.path.join(SCRIPTS, filename))
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m


@pytest.fixture(scope="session")
def ff():
    return _load("gt_full_flow_fms_yes_no_all_or_latest.py", "fullflow")


@pytest.fixture(scope="session")
def pod():
    return _load("gt_full_flow_fms_pod_update_google_sheet_gt_and_order_only.py", "podsync")


@pytest.fixture
def make(ff):
    """Builders shared by the tests."""
    cols = ff.OUTPUT_COLUMNS

    def existing(rows):
        out = []
        for r in rows:
            row = {c: "" for c in cols}
            row.update(r)
            out.append(row)
        return pd.DataFrame(out, columns=cols)

    def sales(rows):
        base = {"Order No": "", "ExtOrderNo": "", "Order Date": "2026-09-05 10:00:00",
                "SKU Code": "", "Created by": "Gauri", "Invoice No": "",
                "Originating warehouse": "W1", "Pick Qty": 5, "Total": 1000}
        return ff.prepare_sales_lines(pd.DataFrame([{**base, **r} for r in rows]))

    def plan(existing_rows, sales_rows, warehouse_mapping=None, pick_pack=None):
        writes, stats, slots, unplaced = ff.plan_backfill(
            existing(existing_rows), sales(sales_rows), pick_pack or {},
            warehouse_mapping or {"W1": "Vadape"})
        return writes, stats

    def cells(writes):
        return {(w[0], w[1]): w[2] for w in writes}

    return types.SimpleNamespace(existing=existing, sales=sales, plan=plan, cells=cells)
