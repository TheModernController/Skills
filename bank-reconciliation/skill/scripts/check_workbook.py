"""Verify a review workbook from its own contents.

Structure: every tab is present, Setup and Checks are hidden the ordinary way, and the stored records are readable.
Formulas: where LibreOffice is available, a copy is recalculated and every formula result the script stored under
"Formula expectations" is compared; any error cell fails the check. Without LibreOffice, Excel recalculates when the
workbook is opened, and the Checks line on Reconciliation shows any failure.
"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

import review_state as S
from foundation.sections import read_kv, read_section
from foundation.xlsx import recalculated

RECALC_CHECK = "Formulas recalculated with LibreOffice agree with the script's values"
ERRORS = {"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"}


def same(expected, got):
    if expected in (None, "") and got in (None, ""):
        return True
    if isinstance(expected, (int, float)) and not isinstance(expected, bool):
        try:
            return abs(float(got) - float(expected)) < 0.005
        except (TypeError, ValueError):
            return False
    return str(expected).strip() == ("" if got is None else str(got).strip())


def compare(calc, expectations):
    """Problems where a recalculated workbook's values differ from the expectations, plus any error cells."""
    problems, errors = [], []
    for e in expectations:
        sheet, cell = str(e["Cell"]).split("!")
        got = calc[sheet][cell].value if sheet in calc.sheetnames else None
        if not same(e["Expected"], got):
            problems.append(f"{e['Cell']}: {got!r} (expected {e['Expected']!r})")
    for ws in calc.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if c.data_type == "e" or (isinstance(c.value, str) and (c.value in ERRORS or c.value.startswith("Err:"))):
                    errors.append(f"{ws.title}!{c.coordinate}")
    if errors:
        problems.append(f"{len(errors)} cell(s) show an error after recalculation, first {errors[0]}")
    return problems, errors[:50]


def check(path, diagnostics=None):
    wb = load_workbook(path, data_only=False)
    out = []

    def add(label, ok, detail=""):
        out.append({"check": label, "result": "PASS" if ok else "FAIL", "detail": detail})

    missing = [t for t in S.TABS if t not in wb.sheetnames]
    add("All tabs present", not missing, ", ".join(missing))
    if missing:
        return {"ok": False, "checks": out, "formulas": "Not checked: tabs are missing."}
    add("Setup and Checks are hidden the ordinary way (Unhide shows them)",
        wb[S.SETUP].sheet_state == "hidden" and wb[S.CHECKS].sheet_state == "hidden")
    add("Every other tab is visible", all(wb[t].sheet_state == "visible" for t in S.TABS if t not in (S.SETUP, S.CHECKS)))
    add("Workbook format recognized", S.is_review_workbook(wb))
    try:
        S.setup_of(wb)
        ok_setup = True
    except Exception as e:  # report, don't crash
        ok_setup, why = False, str(e)
    add("Setup reads back", ok_setup, "" if ok_setup else why)
    run = read_kv(wb[S.CHECKS], S.S_RUN)
    exp = read_section(wb[S.CHECKS], S.S_EXPECT)
    add("Formula expectations stored", bool(exp))
    completion = str(run.get("Completion") or "")
    add("Validation result is recorded", bool(completion), completion)
    shown = wb[S.REC].cell(S.REVIEW_DATE_ROW, 3).value
    add("Review completed on is blank or a date, and matches the record", (shown in (None, "") and not run.get(
        "Review completed on (entered by the reviewer)")) or (isinstance(shown, datetime) and shown.date().isoformat() ==
        str(run.get("Review completed on (entered by the reviewer)"))), f"{shown} vs {run.get('Review completed on (entered by the reviewer)')}")
    if completion.startswith("Complete"):
        add("A complete workbook has its carry forward", bool(read_kv(wb[S.CHECKS], S.S_CARRY_FACTS)))
        add("A complete workbook's engine status is reconciled", str(run.get("Engine status", "")).startswith("Reconciled"),
            run.get("Engine status"))

    recalc = {"state": "not_available", "problems": [], "error_cells": []}
    calc, info = recalculated(path)
    recalc["tool"] = info
    if not info.get("tool"):
        recalc["summary"] = ("LibreOffice is not available here, so the formulas were not recalculated. Excel recalculates them on "
                             "opening, and the Checks line on Reconciliation shows any failure.")
    elif calc is None:
        recalc["state"] = "failed"
        recalc["problems"].append("LibreOffice did not produce a readable recalculated copy"
                                  + (f" (exit code {info['exit_code']})" if "exit_code" in info else "")
                                  + (f": {info['error']}" if info.get("error") else ""))
    else:
        problems, errors = compare(calc, exp)
        recalc["problems"], recalc["error_cells"] = problems, errors
        recalc["state"] = "failed" if problems else "passed"
        recalc["compared"] = len(exp)
    if recalc["state"] != "not_available":
        add(RECALC_CHECK, recalc["state"] == "passed", "; ".join(recalc["problems"][:10]))
        recalc["summary"] = (f"Formulas were recalculated here with LibreOffice: all {len(exp)} stored results agree and no cell shows "
                             "an error." if recalc["state"] == "passed" else
                             "Formula recalculation with LibreOffice did not pass, so this workbook is not final: " +
                             "; ".join(recalc["problems"][:10]))
    res = {"ok": all(c["result"] == "PASS" for c in out), "workbook": Path(path).name, "completion": completion, "checks": out,
           "formulas": recalc["summary"], "recalculation": recalc}
    if recalc["state"] == "failed" and diagnostics:
        save_diagnostics(res, recalc, path, diagnostics)
    return res


def save_diagnostics(res, recalc, path, folder):
    import json, shutil
    folder = Path(folder)
    folder.mkdir(parents=True, exist_ok=True)
    stem = Path(path).stem
    saved = []
    if recalc.get("tool", {}).get("output"):
        copy = folder / f"{stem}-libreoffice-copy.xlsx"
        shutil.copy(recalc["tool"]["output"], copy)
        saved.append(str(copy))
    report = folder / f"{stem}-recalculation-diagnostics.json"
    saved.append(str(report))
    recalc["diagnostics_saved"] = saved
    report.write_text(json.dumps(res, indent=1, default=str, ensure_ascii=False), encoding="utf-8")
