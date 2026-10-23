"""Data-integrity tests: extraction, workbook, validation and KPI formulas."""
import json
import shutil
from decimal import Decimal

import pytest
from openpyxl import load_workbook

from pipeline import dataset
from pipeline.build_site_data import read_observations, validate
from pipeline.build_workbook import FIELD_COL, HEADER_ROW, SHEET_CH, SHEET_FIN, WORKBOOK
from pipeline.metrics import REPORTED, STATUS_CONFLICT, STATUS_MATCHED

# Representative values re-read by hand from the filings (early and late years, both scopes).
GOLDEN = {
    (2020, "revenues"): 37403, (2020, "nike_direct_revenue"): 12382, (2020, "nike_brand_wholesale_revenue"): 23156,
    (2020, "inventories"): 7367, (2021, "ebit_reported"): 6923, (2022, "nike_direct_revenue"): 18726,
    (2023, "nike_brand_wholesale_revenue"): 27397, (2024, "inventories"): 7519, (2025, "demand_creation_expense"): 4689,
    (2026, "revenues"): 46398, (2026, "nike_direct_revenue"): 17720, (2026, "nike_brand_wholesale_revenue"): 27453,
    (2026, "nike_brand_revenue"): 45222, (2026, "net_income"): 3108, (2026, "interest_expense_net"): -50,
    (2023, "nike_brand_digital_sales"): Decimal("12.6"),
}


@pytest.fixture(scope="module")
def facts():
    return dataset.load_facts()


@pytest.fixture(scope="module")
def observations(facts):
    obs = dataset.build_observations(facts, dataset.load_sec_check())
    results = dataset.compute(obs)
    dataset.apply_cross_check_statuses(obs, results)
    return obs, results


def test_every_metric_and_year_present_once(observations):
    obs, _ = observations
    keys = [o["key"] for o in obs]
    assert len(keys) == len(set(keys)) == len(REPORTED) * len(dataset.YEARS)
    missing = [o["key"] for o in obs if o["value"] is None]
    assert missing == ["FY2026|nike_direct_share_stated"]


@pytest.mark.parametrize("key,value", sorted(GOLDEN.items()))
def test_golden_values(observations, key, value):
    obs, _ = observations
    o = next(o for o in obs if (o["fy"], o["metric"]["id"]) == key)
    assert o["value"] == value


def test_sec_api_cross_check_all_match():
    results = dataset.load_sec_check()
    assert results and set(results.values()) == {"match"}


def test_only_disclosed_restatement_is_digital_fy2023(observations):
    obs, _ = observations
    restated = [o["key"] for o in obs if o["restatement"].startswith("ĐIỀU CHỈNH")]
    assert restated == ["FY2023|nike_brand_digital_sales"]


def test_reconciliations_and_comparisons_within_tolerance(observations):
    from pipeline.metrics import CHECKS, COMPARISONS
    _, results = observations
    for c in CHECKS + COMPARISONS:
        for y in dataset.YEARS:
            v = results[y][c["id"]]
            if v is not None:
                assert abs(v) <= Decimal(str(c["tol"])), (c["id"], y, v)


def test_kpi_formulas_against_hand_calculation(observations):
    _, r = observations
    assert r[2020]["gross_margin"] == Decimal(16241) / Decimal(37403)
    assert r[2026]["direct_share"] == Decimal(17720) / Decimal(45222)
    assert r[2021]["direct_growth"] == Decimal(16370) / Decimal(12382) - 1
    assert r[2026]["ebit"] == 3108 + 792 - 50
    assert r[2020]["average_inventory"] == (Decimal(7367) + Decimal(5622)) / 2
    assert r[2024]["days_inventory"] == 365 / (Decimal(28475) / ((Decimal(7519) + Decimal(8454)) / 2))


def test_missing_values_never_become_zero(observations):
    _, r = observations
    assert r[2026]["cmp_direct_share"] is None
    assert dataset.evaluate("a / b", {"a": Decimal(1), "b": None}) is None


def test_mdna_only_values_are_not_marked_matched_without_a_check(observations):
    obs, _ = observations
    for o in obs:
        if o["metric"]["id"].endswith("_pct_change_cn") or o["metric"]["id"] == "nike_brand_digital_sales":
            assert o["status"] != STATUS_MATCHED, o["key"]


# ---------------------------------------------------------------- workbook validation
def test_workbook_passes_validation(facts):
    errors, _ = validate(read_observations(WORKBOOK), facts)
    assert errors == []


def _mutated_copy(tmp_path, sheet, key, field, new_value):
    path = tmp_path / "wb.xlsx"
    shutil.copyfile(WORKBOOK, path)
    wb = load_workbook(path)
    ws = wb[sheet]
    for row in ws.iter_rows(min_row=HEADER_ROW + 2):
        if row[FIELD_COL["key"]].value == key:
            row[FIELD_COL[field]].value = new_value
            break
    else:
        raise AssertionError(key)
    wb.save(path)
    return path


@pytest.mark.parametrize("sheet,key,field,value,expected", [
    (SHEET_FIN, "FY2024|revenues", "value", 51000, "differs from the 10-K extraction"),
    (SHEET_CH, "FY2022|nike_direct_revenue", "value", None, "must be marked missing"),
    (SHEET_CH, "FY2021|nike_direct_pct_change_cn", "value", 30, "whole percent"),
    (SHEET_FIN, "FY2020|cost_of_sales", "unit", "USD tỷ", "unit"),
    (SHEET_CH, "FY2025|nike_brand_revenue", "scope", "NIKE, Inc.", "reporting scope"),
    (SHEET_FIN, "FY2023|gross_profit", "status", STATUS_CONFLICT, "conflicting evidence"),
    (SHEET_FIN, "FY2022|net_income", "page", None, "incomplete provenance"),
    (SHEET_FIN, "FY2021|inventories", "prior", 6000, "comparative"),
    (SHEET_CH, "FY2026|nike_direct_revenue", "key", "FY2026|revenues", "Duplicate"),
])
def test_validator_rejects_bad_data(tmp_path, facts, sheet, key, field, value, expected):
    errors, _ = validate(read_observations(_mutated_copy(tmp_path, sheet, key, field, value)), facts)
    assert any(expected in e for e in errors), errors


def test_workbook_formulas_recalculate_independently():
    from pipeline.verify_workbook import verify
    n, problems = verify(WORKBOOK)
    assert n > 400 and problems == []


def test_site_json_matches_workbook(observations, site_data):
    obs, results = observations
    for o in obs:
        got = site_data["observations"][o["metric"]["id"]][str(o["fy"])]["value"]
        assert got == (None if o["value"] is None else float(o["value"])), o["key"]
    for kid, k in site_data["kpis"].items():
        for y in dataset.YEARS:
            want = results[y][kid]
            assert k["values"][str(y)] == (None if want is None else float(want)), (kid, y)
