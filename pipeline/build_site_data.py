"""Validate the handover workbook and export the dashboard data layer.

Usage:  python -m pipeline.build_site_data

Reads the reported observations from sheets 02 and 03 of the workbook (the source of
truth students may annotate), validates them, recomputes every KPI with the shared
catalogue formulas, and writes site/data/nike-financials.json. Any failed check stops
the build with a non-zero exit code; nothing is exported from data that failed.
"""
import datetime
import json
import math
import shutil
import sys
from decimal import Decimal
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

from . import dataset
from .build_workbook import FIELD_COL, HEADER_ROW, OBS_COLUMNS, SHEET_CH, SHEET_FIN, SHEET_KPI, WORKBOOK
from .metrics import (CHECKS, COMPARISONS, KPIS, REPORTED, REPORTED_BY_ID, STATUS_CONFLICT, STATUS_MISSING, STATUSES)
from .sources import FILINGS, document_title, filing, sec_document_url

ROOT = Path(__file__).resolve().parent.parent
SITE_DATA = ROOT / "site" / "data" / "nike-financials.json"
SITE_DOWNLOAD = ROOT / "site" / "downloads" / WORKBOOK.name
HEADERS = {field: h for h, field, _ in OBS_COLUMNS}
MAY_BE_NEGATIVE = {"interest_expense_net", "other_expense_net", "corporate_revenue"}


class ValidationError(Exception):
    pass


def read_observations(path=WORKBOOK):
    frames = []
    for sheet in (SHEET_FIN, SHEET_CH):
        df = pd.read_excel(path, sheet_name=sheet, header=HEADER_ROW, dtype={HEADERS["page"]: "Int64"})
        missing_cols = [h for h in HEADERS.values() if h not in df.columns]
        if missing_cols:
            raise ValidationError(f"{sheet}: missing columns {missing_cols}")
        df["sheet_name"] = sheet
        frames.append(df)
    return pd.concat(frames, ignore_index=True).rename(columns={v: k for k, v in HEADERS.items()})


def _num(v):
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return None
    return Decimal(str(v))


def validate(df, facts):
    errors, notes = [], []
    keys = df["key"].astype(str)
    dup = keys[keys.duplicated()].tolist()
    if dup:
        errors.append(f"Duplicate observation keys: {dup}")
    expected = {f"FY{y}|{m['id']}" for y in dataset.YEARS for m in REPORTED}
    missing_rows = sorted(expected - set(keys))
    extra_rows = sorted(set(keys) - expected)
    if missing_rows:
        errors.append(f"Observations missing from workbook: {missing_rows}")
    if extra_rows:
        errors.append(f"Unexpected observations (unknown metric or fiscal year outside FY{dataset.YEARS[0]}–FY{dataset.YEARS[-1]}): {extra_rows}")
    for r in df.itertuples():
        mid = r.metric_id
        m = REPORTED_BY_ID.get(mid)
        if m is None:
            continue
        fy = int(str(r.fy_label).replace("FY", ""))
        where = f"{r.key} ({r.sheet_name})"
        if r.key != f"FY{fy}|{mid}":
            errors.append(f"{where}: key does not match fiscal year and metric")
        expected_sheet = SHEET_FIN if m["sheet"] == "financial" else SHEET_CH
        if r.sheet_name != expected_sheet:
            errors.append(f"{where}: belongs on sheet {expected_sheet}")
        if r.unit != m["unit"]:
            errors.append(f"{where}: unit '{r.unit}' should be '{m['unit']}'")
        if r.scope != m["scope"]:
            errors.append(f"{where}: reporting scope '{r.scope}' should be '{m['scope']}'")
        if r.status not in STATUSES:
            errors.append(f"{where}: unknown data status '{r.status}'")
        value = _num(r.value)
        if value is None and r.status != STATUS_MISSING:
            errors.append(f"{where}: value is empty but status is '{r.status}' (empty values must be marked missing)")
        if value is not None and r.status == STATUS_MISSING:
            errors.append(f"{where}: has a value but is marked missing")
        if r.status == STATUS_CONFLICT:
            errors.append(f"{where}: conflicting evidence must be resolved before publishing")
        if value is not None:
            if value < 0 and mid not in MAY_BE_NEGATIVE and m["unit"] != "%":
                errors.append(f"{where}: unexpected negative value {value}")
            if m["unit"] == "%" and abs(value) > 5:
                errors.append(f"{where}: percentage {value} looks like it was entered in whole percent instead of a fraction")
            if r.page is pd.NA or not r.section or str(r.section) == "—" or not r.evidence or not r.accession:
                errors.append(f"{where}: incomplete provenance (page, section, evidence and accession are required)")
            if r.accession != filing(fy)["accession"]:
                errors.append(f"{where}: accession {r.accession} is not the FY{fy} 10-K")
            method = dataset.method_of(m)
            src = facts.get((fy, mid, fy, method))
            if src is None or src["value"] != value:
                errors.append(f"{where}: value {value} differs from the 10-K extraction "
                              f"({src['value'] if src else 'not found'}); correct the workbook or re-run extraction")
            prior = _num(r.prior)
            src_prior = facts.get((fy, mid, fy - 1, method))
            if m["unit"] != "%" and src_prior is not None and prior != src_prior["value"]:
                errors.append(f"{where}: prior-year value {prior} differs from the 10-K comparative {src_prior['value']}")
    return errors, notes


def observations_from(df):
    obs = []
    for r in df.itertuples():
        m = REPORTED_BY_ID[r.metric_id]
        fy = int(str(r.fy_label).replace("FY", ""))
        obs.append(dict(key=r.key, fy=fy, metric=m, value=_num(r.value), prior=_num(r.prior), status=r.status,
                        page=None if r.page is pd.NA else int(r.page), section=r.section, evidence=r.evidence,
                        cross_check=r.cross_check, restatement=r.restatement,
                        verified_by="" if pd.isna(r.verified_by) else str(r.verified_by),
                        verified_on="" if pd.isna(r.verified_on) else str(r.verified_on)))
    return obs


def check_results(results):
    errors = []
    for c in CHECKS + COMPARISONS:
        for y in dataset.YEARS:
            v = results[y][c["id"]]
            if v is None:
                if c in CHECKS:
                    errors.append(f"Check {c['id']} FY{y}: cannot be computed (missing input)")
                continue
            if abs(v) > Decimal(str(c["tol"])):
                errors.append(f"Check {c['id']} FY{y}: difference {v} exceeds tolerance {c['tol']}")
    return errors


def compare_with_workbook_formulas(results, path=WORKBOOK):
    """Cached results of the sheet 04 formulas must equal the Python results."""
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb[SHEET_KPI]
    errors, compared = [], 0
    rows = list(ws.iter_rows(min_row=HEADER_ROW + 1, values_only=True))
    header = rows[0]
    cols = {int(str(h)[2:]): i for i, h in enumerate(header) if isinstance(h, str) and h.startswith("FY")}
    ids = {k["id"] for k in KPIS + CHECKS + COMPARISONS}
    for row in rows[1:]:
        if row[0] not in ids:
            continue
        for y, i in cols.items():
            want, got = results[y][row[0]], row[i]
            compared += 1
            if want is None:
                if got not in ("–", None):
                    errors.append(f"Sheet 04 {row[0]} FY{y}: expected '–', workbook shows {got!r}")
            elif not isinstance(got, (int, float)) or not math.isclose(float(want), got, rel_tol=1e-9, abs_tol=1e-9):
                errors.append(f"Sheet 04 {row[0]} FY{y}: workbook {got!r} vs recomputed {float(want)}")
    if compared == 0:
        errors.append("Sheet 04: no KPI rows found to compare")
    return errors, compared


def _f(v):
    return None if v is None else float(v)


def export(obs, results, extracted_on):
    years = dataset.YEARS
    data = {
        "meta": {
            "title": "NIKE, Inc. – phân phối và kết quả tài chính FY2020–FY2026",
            "generated_on": datetime.date.today().isoformat(),
            "extracted_on": extracted_on,
            "workbook": f"downloads/{WORKBOOK.name}",
            "currency": "USD",
            "fiscal_year_end": "31/5",
            "status_counts": {s: sum(1 for o in obs if o["status"] == s) for s in STATUSES},
            "status_labels": STATUSES,
        },
        "years": years,
        "filings": {str(fy): {"accession": a, "filed": filed, "period_end": pe, "url": sec_document_url(filing(fy)),
                              "title": document_title(filing(fy))} for fy, a, _, filed, pe in FILINGS},
        "metrics": {m["id"]: {k: m[k] for k in ("vi", "en", "unit", "scope", "channel", "kind", "definition", "location")}
                    for m in REPORTED},
        "observations": {m["id"]: {} for m in REPORTED},
        "kpis": {k["id"]: {"vi": k["vi"], "unit": k["unit"], "group": k["group"], "formula": dataset.IDENT.sub(
            lambda t: t.group(0), k["expr"]), "note": k["note"], "values": {str(y): _f(results[y][k["id"]]) for y in years}}
                 for k in KPIS},
        "checks": {c["id"]: {"vi": c["vi"], "tol": c["tol"], "values": {str(y): _f(results[y][c["id"]]) for y in years}}
                   for c in CHECKS + COMPARISONS},
    }
    for o in obs:
        data["observations"][o["metric"]["id"]][str(o["fy"])] = {
            "value": _f(o["value"]), "prior": _f(o["prior"]), "status": o["status"], "page": o["page"],
            "section": o["section"], "evidence": o["evidence"], "cross_check": o["cross_check"],
            "restatement": o["restatement"], "verified_by": o["verified_by"], "verified_on": o["verified_on"]}
    SITE_DATA.parent.mkdir(parents=True, exist_ok=True)
    SITE_DATA.write_text(json.dumps(data, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")
    SITE_DOWNLOAD.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(WORKBOOK, SITE_DOWNLOAD)
    return data


def main():
    facts = dataset.load_facts()
    extracted_on = json.loads((ROOT / "data" / "extracted" / "meta.json").read_text())["extracted_on"]
    df = read_observations()
    errors, _ = validate(df, facts)
    if errors:
        print("VALIDATION FAILED – dashboard data not exported:")
        for e in errors:
            print("  ✗", e)
        return 1
    obs = observations_from(df)
    results = dataset.compute(obs)
    errors = check_results(results)
    wb_errors, compared = compare_with_workbook_formulas(results)
    errors += wb_errors
    if errors:
        print("VALIDATION FAILED – dashboard data not exported:")
        for e in errors:
            print("  ✗", e)
        return 1
    data = export(obs, results, extracted_on)
    n_values = sum(1 for o in obs if o["value"] is not None)
    print(f"✓ {len(obs)} observations validated ({n_values} with values, "
          f"{data['meta']['status_counts'][STATUS_MISSING]} marked missing)")
    print(f"✓ {len(CHECKS)} arithmetic reconciliations and {len(COMPARISONS)} comparisons with NIKE-reported "
          f"figures pass for FY{dataset.YEARS[0]}–FY{dataset.YEARS[-1]}")
    print(f"✓ {compared} sheet 04 formula results match the recomputed KPIs")
    print(f"✓ Wrote {SITE_DATA.relative_to(ROOT)} and {SITE_DOWNLOAD.relative_to(ROOT)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
