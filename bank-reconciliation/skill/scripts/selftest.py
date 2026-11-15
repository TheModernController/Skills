"""Self-test with fictional data: run it in Claude's environment to verify the workbook formulas with LibreOffice there.

  python bankrec.py selftest [--out DIR]

It builds a fictional March workbook, reviews it with scripted edits (approvals, a split line, a new account, a missing
account, a JE number entered before approval, a changed source amount), validates it, then starts April from the checked
March workbook and does the same. After each stage it recalculates the workbook with LibreOffice and compares every
formula result the script stored. It reads no user files and makes no network connections.
"""
from __future__ import annotations

import json, shutil, subprocess, sys, tempfile
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

HERE = Path(__file__).resolve().parent

BANK_MAR = """line_id,date,description,reference,amount
BK-001,2026-03-03,DEPOSIT,INV-1041,3200.00
BK-002,2026-03-05,ACH DEBIT NORTHLINE SUPPLY,ACH-5521,-1860.50
BK-003,2026-03-09,CHECK PAID,2051,-812.40
BK-004,2026-03-11,MOBILE DEPOSIT,,2150.00
BK-005,2026-03-12,ACH DEBIT ASHGROVE REALTY,ACH-5522,-4500.00
BK-006,2026-03-19,CHECK PAID,2052,-450.00
BK-007,2026-03-19,DEPOSIT,INV-1043,480.00
BK-008,2026-03-23,ACH DEBIT CEDAR FREIGHT,ACH-5523,-925.00
BK-009,2026-03-24,RETURNED ITEM,,-480.00
BK-010,2026-03-26,INCOMING WIRE,,1500.00
BK-011,2026-03-31,MONTHLY SERVICE FEE,,-35.00
"""
LEDGER_MAR = """line_id,date,type,reference,party,amount
CB-001,2026-03-02,Receipt,INV-1041,Birchwood Cafe,3200.00
CB-002,2026-03-04,Payment,ACH-5521,Northline Supply,-1860.50
CB-003,2026-03-05,Check,2051,Tallow & Co,-812.40
CB-004,2026-03-10,Receipt,INV-1042,Harbor Yoga,2150.00
CB-005,2026-03-12,Payment,ACH-5522,Ashgrove Realty,-4500.00
CB-006,2026-03-16,Check,2052,Mossy Print,-540.00
CB-007,2026-03-18,Receipt,INV-1043,Fern Street Gym,480.00
CB-008,2026-03-20,Payment,ACH-5523,Cedar Freight,-925.00
CB-009,2026-03-20,Payment,ACH-5523,Cedar Freight,-925.00
CB-010,2026-03-30,Check,2053,Ridgeline Tools,-1275.00
CB-011,2026-03-31,Receipt,INV-1044,Oakline Studio,2640.00
"""
BANK_APR = """line_id,date,description,reference,amount
BK-011,2026-03-31,MONTHLY SERVICE FEE,,-35.00
BK-101,2026-04-02,DEPOSIT,INV-1044,2640.00
BK-102,2026-04-06,DEPOSIT,INV-1045,1980.00
BK-103,2026-04-08,ACH DEBIT NORTHLINE SUPPLY,ACH-5524,-2310.00
BK-104,2026-04-13,CHECK PAID,2053,-1275.00
BK-105,2026-04-20,CHECK PAID,2054,-615.25
BK-106,2026-04-30,MONTHLY SERVICE FEE,,-35.00
"""
LEDGER_APR = """line_id,date,type,reference,party,amount
CB-101,2026-04-01,Journal,JE-0457,Bank fees,-35.00
CB-102,2026-04-01,Journal,JE-0458,Mossy Print,90.00
CB-103,2026-04-01,Journal,JE-0459,Cedar Freight,925.00
CB-104,2026-04-02,Journal,,Customer deposits,1500.00
CB-105,2026-04-06,Receipt,INV-1045,Willow Bakery,1980.00
CB-106,2026-04-08,Payment,ACH-5524,Northline Supply,-2310.00
CB-107,2026-04-14,Check,2054,Mossy Print,-615.25
CB-108,2026-04-28,Check,2055,Cedar Freight,-740.00
CB-109,2026-04-29,Payment,ACH-5525,Ashgrove Realty,-1150.00
CB-110,2026-04-30,Receipt,INV-1046,Harbor Yoga,860.00
"""
CHART = "Number,Name\n1010,Operating Cash\n1200,Accounts receivable\n2000,Accounts payable\n2300,Customer deposits\n6150,Bank fees\n"


def profile():
    src = lambda desc, extra: {"file_type": "csv", "sheet": None, "header_row": 1,
                               "columns": {"id": "line_id", "date": "date", "description": desc, "reference": "reference", "amount": "amount", **extra},
                               "date_format": "%Y-%m-%d", "sign": "as_is", "summary_rows": "none"}
    bank = src("description", {})
    bank["header_fingerprint"] = ["line_id", "date", "description", "reference", "amount"]
    ledger = src("party", {"type": "type"})
    ledger["type_map"] = {"Receipt": "receipt", "Payment": "payment", "Check": "check"}
    ledger["header_fingerprint"] = ["line_id", "date", "type", "reference", "party", "amount"]
    return {"profile_format": 1, "profile_version": 1, "confirmed": True,
            "account": {"company": "Alder Creek Outfitters (fictional)", "nickname": "Self-test account", "last4": "4821", "currency": "USD",
                        "accounting_system": "Fictional ledger"},
            "process": {"reconciled_by": "Fictional reviewer"}, "sources": {"bank": bank, "ledger": ledger},
            "rules": {"date_window_days": 3, "deposit_in_transit_days": 3, "fee_descriptions": ["MONTHLY SERVICE FEE"], "returned_item_words": ["RETURNED"]},
            "accounts": {"cash": "1010 Operating Cash", "bank_fees": "6150 Bank fees"}}


def cli(*args):
    cp = subprocess.run([sys.executable, str(HERE / "bankrec.py")] + [str(a) for a in args], capture_output=True, text=True)
    try:
        return json.loads(cp.stdout)
    except ValueError:
        return {"stopped": (cp.stdout + cp.stderr)[-1500:]}


def find_rows(path):
    """Where each Ref sits: Reconciliation row, and each entry's header, line and total rows."""
    wb = load_workbook(path)
    rec, je = wb["Reconciliation"], wb["Journal entries"]
    rec_rows = {}
    r = 21
    while rec.cell(r, 2).value:
        rec_rows.setdefault(rec.cell(r, 2).value, r)
        r += 1
    entries = {}
    for r in range(5, je.max_row + 1):
        v = je.cell(r, 2).value
        if v and str(v).startswith("R-"):
            t = r + 1
            while je.cell(t, 5).value != "Total":
                t += 1
            entries[v] = {"head": r, "lines": list(range(r + 1, t)), "total": t,
                          "accounts": {je.cell(x, 4).value: x for x in range(r + 1, t)}}
    return rec_rows, entries


def review(src, dst, edits):
    """Apply scripted edits with openpyxl, as a reviewer would in Excel (a split inserts a row under the yellow line)."""
    wb = load_workbook(src)
    for e in edits:
        ws = wb[e[0]]
        if e[1] == "insert":
            ws.insert_rows(e[2])
        else:
            ws[e[1]] = e[2]
    wb.save(dst)


def run(out=None):
    tmp = Path(tempfile.mkdtemp(prefix="bankrec_selftest_"))
    out = Path(out) if out else tmp / "out"
    out.mkdir(parents=True, exist_ok=True)
    stages, ok = [], True

    def stage(name, workbook, extra=None):
        nonlocal ok
        chk = cli("check", workbook)
        rec = chk.get("recalculation", {})
        passed = chk.get("ok", False) if "stopped" not in chk else False
        ok = ok and passed
        stages.append({"stage": name, "workbook": Path(workbook).name, "passed": passed, "libreoffice": rec.get("state", "unknown"),
                       "formulas_compared": rec.get("compared"), "problems": (rec.get("problems") or [])[:5] or
                       [c["check"] + ": " + c["detail"] for c in chk.get("checks", []) if c["result"] == "FAIL"][:5] or
                       ([chk["stopped"]] if "stopped" in chk else []), **(extra or {})})

    def expect(name, cond, detail=""):
        nonlocal ok
        ok = ok and bool(cond)
        if not cond:
            stages.append({"stage": name, "passed": False, "problems": [detail]})

    (tmp / "bank-mar.csv").write_text(BANK_MAR, encoding="utf-8")
    (tmp / "ledger-mar.csv").write_text(LEDGER_MAR, encoding="utf-8")
    (tmp / "bank-apr.csv").write_text(BANK_APR, encoding="utf-8")
    (tmp / "ledger-apr.csv").write_text(LEDGER_APR, encoding="utf-8")
    (tmp / "chart.csv").write_text(CHART, encoding="utf-8")
    (tmp / "profile.json").write_text(json.dumps(profile()), encoding="utf-8")
    (tmp / "march.json").write_text(json.dumps({"profile": "profile.json", "bank_file": "bank-mar.csv", "ledger_file": "ledger-mar.csv",
                                                "period_start": "2026-03-01", "period_end": "2026-03-31", "statement_opening": "48250.00",
                                                "statement_closing": "46517.10", "ledger_opening": "48250.00", "chart_file": "chart.csv",
                                                "decisions": "decisions.json", "out_dir": str(out / "draft")}), encoding="utf-8")

    d = cli("run", tmp / "march.json", "--on", "2026-04-01")
    if "stopped" in d:
        return {"ok": False, "stages": [{"stage": "March draft", "passed": False, "problems": [d["stopped"]]}]}
    stage("March draft", d["workbook"], {"remaining_difference": d.get("remaining_difference")})
    rr, ee = find_rows(d["workbook"])
    JE, REC = "Journal entries", "Reconciliation"
    acct = lambda ref, n=1: f"D{ee[ref]['lines'][n]}" if len(ee[ref]["lines"]) > n else f"D{ee[ref]['lines'][0]}"
    # identify entries by their proposal text
    wb = load_workbook(d["workbook"])
    props = {ref: str(wb[JE].cell(ee[ref]["head"], 9).value or "") for ref in ee}
    by = lambda word: next(ref for ref, p in props.items() if word in p)
    fee, corr, dup, ret, wire = by("bank fee"), by("Correct the ledger"), by("duplicate"), by("Returned"), by("Who sent")
    timing = [ref for ref, row in rr.items() if wb[REC].cell(row, 10).value == "Timing item—no JE"]
    other_line = lambda ref: next(x for a, x in ee[ref]["accounts"].items() if a != "1010 Operating Cash")

    # review 1: problems (missing account, JE number before approval, split onto a new account, bank error choice, source edit)
    w = ee[wire]
    review(d["workbook"], tmp / "review1.xlsx", [
        (JE, f"J{ee[corr]['head']}", "Confirm bank error—no JE"),
        (JE, f"J{ee[dup]['head']}", "Approved: reverse duplicate"),
        (JE, f"K{ee[ret]['head']}", "JE-0999"),
        (JE, f"D{other_line(wire)}", "2300 Customer deposits"), (JE, f"G{other_line(wire)}", 1200),
        (JE, "insert", w["total"]), (JE, f"D{w['total']}", "2350 Unearned revenue"), (JE, f"G{w['total']}", 300),
        (JE, f"J{w['head']}", "Approved"),
        ("Bank raw", "E8", -4600),
    ])
    v1 = cli("validate", tmp / "review1.xlsx", "--record", "--work", tmp / "v1", "--out", out / "review1", "--on", "2026-04-02")
    if "stopped" in v1:
        return {"ok": False, "stages": stages + [{"stage": "March review with problems", "passed": False, "problems": [v1["stopped"]]}]}
    expect("Problems are reported", len(v1.get("blocking", [])) >= 3 and v1.get("questions"), json.dumps(v1.get("blocking")))
    expect("Not complete while problems remain", str(v1.get("completion", "")).startswith("Not complete"), v1.get("completion"))
    expect("The uploaded workbook is unchanged", v1.get("uploaded_file_unchanged"), "")
    stage("March checked, with problems", v1["workbook"])

    # review 2: everything approved per the fictional facts, JE numbers for three entries
    rr2, ee2 = find_rows(d["workbook"])
    edits = [(JE, f"J{ee2[corr]['head']}", "Approve ledger correction"), (JE, f"D{other_line(corr)}", "2000 Accounts payable"),
             (JE, f"K{ee2[corr]['head']}", "JE-0458"),
             (JE, f"J{ee2[dup]['head']}", "Approved: reverse duplicate"), (JE, f"D{other_line(dup)}", "2000 Accounts payable"),
             (JE, f"K{ee2[dup]['head']}", "JE-0459"),
             (JE, f"J{ee2[ret]['head']}", "Approved"),
             (JE, f"J{ee2[wire]['head']}", "Approved"), (JE, f"D{other_line(wire)}", "2300 Customer deposits"),
             (JE, f"J{ee2[fee]['head']}", "Approved"), (JE, f"K{ee2[fee]['head']}", "JE-0457"), (JE, f"M{ee2[fee]['head']}", "Posted 3 Apr")]
    edits += [(REC, f"I{rr2[t]}", "Approved") for t in timing] + [(REC, "C6", datetime(2026, 4, 3))]
    review(d["workbook"], tmp / "review2.xlsx", edits)
    v2 = cli("validate", tmp / "review2.xlsx", "--record", "--work", tmp / "v2", "--out", out / "march", "--on", "2026-04-03")
    if "stopped" in v2:
        return {"ok": False, "stages": stages + [{"stage": "March review", "passed": False, "problems": [v2["stopped"]]}]}
    expect("March completes", str(v2.get("completion", "")).startswith("Complete"), v2.get("completion"))
    expect("The review date is recorded", v2.get("review_completed_on") == "2026-04-03", str(v2.get("review_completed_on")))
    expect("March remaining difference is 0.00", str(v2.get("figures", {}).get("remaining_difference")) in ("0.00", "0"), "")
    stage("March checked, complete", v2["workbook"], {"completion": v2.get("completion")})

    # April from the checked March workbook only
    c = cli("carry", v2["workbook"], "--work", tmp / "april")
    if "stopped" in c:
        return {"ok": False, "stages": stages + [{"stage": "April setup", "passed": False, "problems": [c["stopped"]]}]}
    (tmp / "april.json").write_text(json.dumps({"profile": c["setup_written"], "bank_file": str(tmp / "bank-apr.csv"),
                                                "ledger_file": str(tmp / "ledger-apr.csv"), "period_start": "2026-04-01",
                                                "period_end": "2026-04-30", "statement_opening": "46517.10", "statement_closing": "46901.85",
                                                "ledger_opening": "45882.10", "prior_workbook": v2["workbook"], "accounts_from": v2["workbook"],
                                                "decisions": "decisions.json", "out_dir": str(out / "april-draft")}), encoding="utf-8")
    a = cli("run", tmp / "april.json", "--on", "2026-05-02")
    if "stopped" in a:
        return {"ok": False, "stages": stages + [{"stage": "April draft", "passed": False, "problems": [a["stopped"]]}]}
    stage("April draft", a["workbook"], {"remaining_difference": a.get("remaining_difference")})
    rr3, ee3 = find_rows(a["workbook"])
    wb3 = load_workbook(a["workbook"])
    edits = []
    for ref, row in rr3.items():
        jr = str(wb3[REC].cell(row, 10).value or "")
        if jr == "Timing item—no JE":
            label = "Approved: outstanding payment" if "Ashgrove" in str(wb3[REC].cell(row, 6).value) else "Approved"
            edits.append((REC, f"I{row}", label))
        elif jr.startswith("Posting of a carried"):
            edits.append((REC, f"I{row}", "Approved: posting of 2026-03 " + wire))
    for ref, e in ee3.items():
        if str(wb3[JE].cell(e["head"], 10).value) == "Pending":
            edits.append((JE, f"J{e['head']}", "Approved"))
    review(a["workbook"], tmp / "review3.xlsx", edits)
    v3 = cli("validate", tmp / "review3.xlsx", "--record", "--work", tmp / "v3", "--out", out / "april", "--on", "2026-05-03")
    if "stopped" in v3:
        return {"ok": False, "stages": stages + [{"stage": "April review", "passed": False, "problems": [v3["stopped"]]}]}
    expect("April completes", str(v3.get("completion", "")).startswith("Complete"), v3.get("completion"))
    stage("April checked, complete", v3["workbook"], {"completion": v3.get("completion")})

    lo = {s_.get("libreoffice") for s_ in stages if "libreoffice" in s_}
    if lo == {"passed"}:
        summary = "LibreOffice recalculated every stage, and every formula result agrees with the script."
    elif "failed" in lo:
        summary = "LibreOffice recalculation did not agree at one or more stages. See the problems listed."
    else:
        summary = "LibreOffice isn't available here, so the formulas were not recalculated. The workflow itself was tested."
    res = {"ok": ok and "failed" not in lo, "libreoffice_verified": lo == {"passed"}, "summary": summary, "stages": stages,
           "workbooks_kept_in": str(out)}
    if not out or str(out).startswith(str(tmp)):
        res["workbooks_kept_in"] = None
        shutil.rmtree(tmp, ignore_errors=True)
    return res
