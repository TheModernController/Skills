"""Read a reviewed workbook back, validate the accountant's edits, and build the checked workbook.

Rules (the user's, 2026-11-08):
- Only a "validate my reviewed workbook" request records approvals (--record), and only items explicitly marked Approved,
  under the name in Reconciled by. Without --record nothing is recorded and no workbook is written.
- The uploaded file is never changed. The checked workbook is a new file.
- Valid edits and comments are kept. Problems are flagged, not silently fixed.
- Edits to fixed data are not applied: the returned copy restores the source values and logs each edit. A changed source
  amount blocks completion until the user confirms it was accidental (--accept-restore) or provides corrected source data.
- Completion is set only when every check passes and nothing is pending, never merely because the difference is zero.
- 0.5.1: one choice per item records the decision. "Review completed on" is the reviewer's own date: it is recorded, and it
  never completes a reconciliation that has unresolved items or failed checks.
- A 0.5.0 upload (format review-1) is validated only with from_050=True: its two renamed amount-mismatch choices are read
  as the new ones, its script-written Completion text is set aside, and every conversion is logged.
"""
from __future__ import annotations

import csv, hashlib, json
from datetime import date, datetime, timezone
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook, load_workbook

import bank_rules
import review_state as S
import review_workbook as W
from foundation import profiles
from foundation.sections import read_kv, read_section
from foundation.tables import InputError, text

ZERO = Decimal("0.00")


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def dec_or_none(v):
    if v in (None, ""):
        return None
    try:
        return Decimal(str(v)).quantize(Decimal("0.01"))
    except Exception:
        return None


def same(a, b):
    return W.canon_cell(a) == W.canon_cell(b)


def to_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    try:
        return datetime.strptime(str(v)[:10], "%Y-%m-%d").date()
    except (TypeError, ValueError):
        return None


# ------------------------------------------------------------------ reading what Claude stored

def stored(wb):
    ck = wb[S.CHECKS]
    run = read_kv(ck, S.S_RUN)
    snap = {}
    for side in ("bank", "ledger"):
        rows = read_section(ck, S.S_SNAP[side])
        grid = []
        for r_ in rows:
            vals = [r_[k] for k in r_ if k != "Row"]
            while vals and vals[-1] in (None, ""):
                vals.pop()
            grid.append(vals)
        snap[side] = grid
    proposed = {}
    for p in read_section(ck, S.S_PROPOSED):
        e = proposed.setdefault(p["Ref"], {"proposal": p.get("Proposal") or "", "title": p.get("Title") or "",
                                           "date": to_date(p.get("Date")), "lines": []})
        e["lines"].append(dict(account=p.get("Account") or "", description=p.get("Description") or "",
                               debit=dec_or_none(p.get("Debit")), credit=dec_or_none(p.get("Credit")), fixed=p.get("Fixed") == "Yes"))
    refs = {r_["Key"]: r_["Ref"] for r_ in read_section(ck, S.S_REFS)}
    where = {r_["Ref"]: r_.get("Where") for r_ in read_section(ck, S.S_REFS)}
    recorded = read_section(ck, S.S_DECISIONS)
    recorded = [d for d in recorded if d.get("Ref") and d["Ref"] != "None yet"]
    rec_rows = read_section(ck, S.S_REC_ROWS)
    return dict(run=run, snap=snap, proposed=proposed, refs=refs, where=where, recorded=recorded, rec_rows=rec_rows)


def read_rec(wb):
    ws = wb[S.REC]
    reconciled_by = text(ws.cell(S.HEADER_ROWS["Reconciled by"], 3).value).strip()
    rows = []
    r = S.FIRST_ROW
    while ws.cell(r, 2).value not in (None, ""):
        rows.append({"row": r, "Ref": text(ws.cell(r, 2).value).strip(),
                     "cells": [ws.cell(r, c).value for c in range(3, 9)], "approval": ws.cell(r, 9).value})
        r += 1
    return reconciled_by, rows


def read_je(wb, refs):
    """Entries keyed by Ref: approval, JE number, comment and every line between the entry's header row and its Total row.
    Rows with content outside an entry are returned separately, so they are flagged rather than lost."""
    ws = wb[S.JE]
    out, orphans = {}, []
    r = S.JE_FIRST
    last = ws.max_row
    cur = None
    while r <= last:
        b = text(ws.cell(r, 2).value).strip()
        e_ = text(ws.cell(r, 5).value).strip()
        vals = [ws.cell(r, c).value for c in (4, 5, 6, 7)]
        if b in refs:
            cur = out.setdefault(b, {"header_row": r, "date": ws.cell(r, 3).value, "title": ws.cell(r, 5).value,
                                     "proposal": ws.cell(r, 9).value, "approval": text(ws.cell(r, 10).value).strip(),
                                     "je_number": text(ws.cell(r, 11).value).strip(), "comment": text(ws.cell(r, 13).value).strip(),
                                     "lines": []})
        elif cur is not None and e_ == "Total" and not text(ws.cell(r, 4).value):
            cur = None
        elif any(v not in (None, "") for v in vals):
            line = dict(account=text(vals[0]).strip(), description=text(vals[1]).strip(), debit=dec_or_none(vals[2]),
                        credit=dec_or_none(vals[3]))
            bad = [v for v in vals[2:] if v not in (None, "") and dec_or_none(v) is None]
            if bad:
                line["unreadable"] = True
            if cur is not None:
                cur["lines"].append(line)
            else:
                orphans.append({"row": r, **line})
        r += 1
    return out, orphans


def read_raw(wb, tab):
    ws = wb[tab]
    grid = []
    for r in range(3, ws.max_row + 1):
        grid.append([ws.cell(r, c).value for c in range(1, ws.max_column + 1)])
    while grid and not any(v not in (None, "") for v in grid[-1]):
        grid.pop()
    return grid


# ------------------------------------------------------------------ rebuilding the source files from the stored copy

def rebuild_source(grid, kind, sheet, path):
    path = Path(path)
    if kind == "csv":
        with open(path, "w", newline="", encoding="utf-8") as fh:
            w = csv.writer(fh)
            for row in grid:
                w.writerow(["" if v is None else v for v in row])
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = sheet or "Sheet1"
        for row in grid:
            ws.append(list(row))
        wb.save(path)
    return path


# ------------------------------------------------------------------ validation

def review_date(v):
    """The reviewer's 'Review completed on' cell: (date or None, problem or None)."""
    if v in (None, ""):
        return None, None
    if isinstance(v, datetime):
        return v.date(), None
    if isinstance(v, date):
        return v, None
    s = text(v).strip()
    for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%m/%d/%y"):
        try:
            return datetime.strptime(s, fmt).date(), None
        except ValueError:
            pass
    return None, f"Review completed on reads '{s}', which isn't a date. What date did you finish your review?"


def validate(workbook, work, out_dir, record=False, confirm_accounts=(), accept_restore=False, bank_file=None, ledger_file=None,
             on=None, version="", from_050=False):
    upload = Path(workbook)
    before = sha(upload)
    work = Path(work).resolve()
    out_dir = Path(out_dir).resolve()
    work.mkdir(parents=True, exist_ok=True)
    wb = load_workbook(upload, data_only=False)
    if not S.is_review_workbook(wb):
        raise InputError(f"{upload.name} isn't a review workbook from this version of the skill.")
    fmt = S.workbook_format(wb)
    if fmt == "review-1" and not from_050:
        raise InputError("MIGRATION NEEDED: this workbook was made by skill 0.5.0. 0.5.1 renamed the two amount-mismatch choices "
                         "('Approved: correct ledger' becomes 'Approve ledger correction'; 'Approved: bank error' becomes "
                         "'Confirm bank error—no JE') and replaced the Completion text with a 'Review completed on' date. "
                         "Ask the user whether to convert it (validate again with --from-0.5.0; every conversion is listed and "
                         "every other edit is checked as usual) or to keep validating it with 0.5.0.")
    migrating = fmt == "review-1"
    st = stored(wb)
    run = st["run"]
    prof = S.setup_of(wb)
    accounts = S.accounts_of(wb)
    today = on or date.today().isoformat()
    edits, blocking, questions = [], [], []
    if migrating:
        edits.append(("Workbook made by skill 0.5.0", "Converted to the 0.5.1 format at your request; conversions are listed below."))
    roles = {side: W.roles_of(prof["sources"][side]) for side in ("bank", "ledger")}

    # 1. the stored source copy must be intact, unless corrected exports are given
    corrected = {"bank": bank_file, "ledger": ledger_file}
    if W.snapshot_digest(st["snap"]) != run.get("Snapshot digest") and not (bank_file and ledger_file):
        raise InputError("The hidden copy of the source files in this workbook was changed, so it can't be used to check your edits. "
                         "Please upload the original bank and ledger exports with the workbook.")

    # 2. edits to the visible source tabs: never applied; amounts block completion until confirmed or corrected
    files = {}
    for side, tab in (("bank", S.BANK), ("ledger", S.LEDGER)):
        grid = st["snap"][side]
        cfg = prof["sources"][side]
        name, ftype, sheet = run.get(f"{side.title()} file"), run.get(f"{side.title()} file type"), run.get(f"{side.title()} sheet")
        if corrected[side]:
            from foundation.tables import read_grid, sha256
            sheet_name, grid, _ = read_grid(corrected[side], cfg.get("sheet"))
            name, ftype, sheet = Path(corrected[side]).name, Path(corrected[side]).suffix.lower().lstrip("."), sheet_name or ""
            edits.append((f"{tab}: corrected export {name} uploaded", "Used instead of the original. Items were matched again."))
            src_path = corrected[side]
            fsha = sha256(corrected[side])
        else:
            src_path = rebuild_source(grid, "csv" if ftype in ("csv", "txt", "tsv") else "xlsx", sheet, work / f"{side}-source.{'csv' if ftype in ('csv', 'txt', 'tsv') else 'xlsx'}")
            fsha = run.get(f"{side.title()} file SHA-256")
        header_row = next((i + 1 for i, row in enumerate(grid) if
                           {text(c).casefold() for c in row if text(c)} >= {v.casefold() for v in cfg["columns"].values() if v}), 1)
        files[side] = {"name": name, "grid": grid, "sha": fsha, "type": ftype, "sheet": sheet, "path": str(src_path),
                       "header_row": header_row}
        if corrected[side]:
            continue
        expect = W.display_grid(grid, cfg, header_row)
        shown = read_raw(wb, tab)
        head = [text(c) for c in grid[header_row - 1]] if grid else []
        amount_cols = {j for j, h in enumerate(head) if roles[side].get(h.casefold()) in ("amount", "debit", "credit")}
        changed_amount = False
        for i in range(max(len(expect), len(shown))):
            a = expect[i] if i < len(expect) else []
            b = shown[i] if i < len(shown) else []
            for j in range(max(len(a), len(b))):
                va = a[j] if j < len(a) else None
                vb = b[j] if j < len(b) else None
                if not same(va, vb):
                    is_amount = j in amount_cols and i >= header_row
                    changed_amount = changed_amount or is_amount
                    cell = f"{tab}!{chr(65 + j) if j < 26 else 'col ' + str(j + 1)}{i + 3}"
                    edits.append((f"{cell}: {W.canon_cell(va) or '(blank)'} changed to {W.canon_cell(vb) or '(blank)'}",
                                  "Not applied: source data comes from your export. The original value is restored in the returned copy."
                                  + (" Completion is blocked until you confirm this was accidental or upload a corrected export."
                                     if is_amount and not accept_restore else "")))
        if changed_amount:
            if accept_restore:
                edits.append((f"{tab}: you confirmed the amount edit(s) were accidental", "Originals restored; completion no longer blocked."))
            else:
                blocking.append(f"A source amount on {tab} was changed. Was that accidental (I'll restore it), or will you upload a corrected export?")

    # 3. re-run the rules on the same source data, to rebuild the items and their Refs
    profile_path = work / "profile.json"
    profiles.save(prof, profile_path)
    period = {"profile": str(profile_path), "bank_file": files["bank"]["path"], "ledger_file": files["ledger"]["path"],
              "period_start": run["Period start"], "period_end": run["Period end"],
              "statement_opening": str(run["Stated bank opening balance"]) if run.get("Stated bank opening balance") is not None else None,
              "statement_closing": str(run["Stated bank closing balance"]) if run.get("Stated bank closing balance") is not None else None,
              "ledger_opening": str(run["Ledger opening balance"]) if run.get("Ledger opening balance") is not None else None,
              "ledger_closing": str(run["Ledger closing balance (stated)"]) if run.get("Ledger closing balance (stated)") not in (None, "") else None,
              "posted_before_opening": [x for x in str(run.get("Posted before opening") or "").split(", ") if x],
              "decisions": str(work / "decisions.json"), "out_dir": str(out_dir)}
    if read_kv(wb[S.CHECKS], S.S_CARRIED_FACTS):
        period["prior_embedded"] = str(upload.resolve())
    period_path = work / "period.json"
    period_path.write_text(json.dumps({k: v for k, v in period.items() if v is not None}, indent=1), encoding="utf-8")
    (work / "decisions.json").write_text("[]", encoding="utf-8")
    r0 = bank_rules.run(period_path)
    refs, mref, iref = W.assign_refs(r0, st["refs"])
    items_by_ref = {iref[i.id]: i for i in r0["items"]}
    for k, ref in st["refs"].items():
        if ref not in items_by_ref and not ref.startswith("M-"):
            edits.append((f"{ref}: no longer in this run", "Its source lines changed, so any decision on it was not applied."))

    # 4. what the accountant entered
    reconciled_by, rec_rows = read_rec(wb)
    entry_refs = {ref for ref, i in items_by_ref.items() if any(o["je"] for o in (([W.fixed_option(i, r0)] if W.fixed_option(i, r0) else W.options_for(i, r0))))}
    je, orphans = read_je(wb, entry_refs)
    for o in orphans:
        blocking.append(f"Journal entries row {o['row']} ({o['account'] or 'no account'}) is outside any entry. Which entry does it belong to?")
        edits.append((f"Journal entries row {o['row']}: line outside an entry ({o['account']}, {o['description']}, "
                      f"{o['debit'] or ''} {o['credit'] or ''})", "Not applied. Move it inside its entry and upload again, or tell Claude which entry."))
    rec_by_ref = {}
    for row in rec_rows:
        rec_by_ref.setdefault(row["Ref"], []).append(row)
    # fixed Reconciliation cells are display only: restore and log any edit
    expected_rec = {}
    for x in st["rec_rows"]:
        expected_rec.setdefault(x["Ref"], []).append(x)
    for ref, exp_rows in expected_rec.items():
        got = rec_by_ref.get(ref, [])
        if not got:
            edits.append((f"Reconciliation: row(s) for {ref} missing", "Restored in the returned copy."))
            continue
        for k, x in enumerate(exp_rows):
            if k >= len(got):
                break
            want = [x.get(c) for c in ("Bank date", "Ledger date", "Bank description", "Ledger description", "Bank amount", "Ledger amount")]
            for name_, a, b in zip(("Bank date", "Ledger date", "Bank description", "Ledger description", "Bank amount", "Ledger amount"),
                                   want, got[k]["cells"]):
                if not same(a, b):
                    edits.append((f"Reconciliation {ref} {name_}: {W.canon_cell(a) or '(blank)'} changed to {W.canon_cell(b) or '(blank)'}",
                                  "Not applied: this comes from your files. Restored in the returned copy."))

    # 4b. the reviewer's own date (review-1 wrote the completion text in this cell instead)
    cell6 = wb[S.REC].cell(S.REVIEW_DATE_ROW, 3).value
    if migrating:
        review_on = None
        if text(cell6).strip() != text(run.get("Completion")).strip():
            edits.append((f"Reconciliation C{S.REVIEW_DATE_ROW} (0.5.0 Completion cell) was changed to '{text(cell6)}'",
                          "Not used: in 0.5.1 this cell is 'Review completed on'. Enter your review date there and upload again."))
        else:
            edits.append((f"Reconciliation C{S.REVIEW_DATE_ROW}: 0.5.0 Completion text set aside",
                          "Now 'Review completed on', left blank for your date. Completion is in the validation result."))
    else:
        review_on, problem = review_date(cell6)
        if problem:
            questions.append(problem)
            edits.append((f"Reconciliation C{S.REVIEW_DATE_ROW}: '{text(cell6)}'", "Not a date, so not recorded."))

    # 5. approvals and entries
    approvals, entries = {}, {}
    cash = prof["accounts"]["cash"]
    acct_set = {a.casefold() for a in accounts}
    for a in confirm_accounts:
        if a.casefold() not in acct_set:
            accounts.append(a)
            acct_set.add(a.casefold())
            edits.append((f"New account confirmed: {a}", "Added to the Accounts tab."))
    for ref, i in items_by_ref.items():
        fx = W.fixed_option(i, r0)
        opts = [fx] if fx else W.options_for(i, r0)
        labels = [o["label"] for o in opts]
        where = "JE" if any(o["je"] for o in opts) else ("Rec" if opts and not fx else "fixed")
        if where == "JE" and not fx:
            v = (je.get(ref) or {}).get("approval") or W.PENDING
        elif where == "Rec":
            v = text((rec_by_ref.get(ref) or [{}])[0].get("approval")).strip() or W.PENDING
        else:
            continue
        if migrating and v in W.OLD_LABELS and W.OLD_LABELS[v] in labels:
            edits.append((f"{ref}: 0.5.0 choice '{v}'", f"Read as '{W.OLD_LABELS[v]}', the same treatment in 0.5.1."))
            v = W.OLD_LABELS[v]
        if v != W.PENDING and v not in labels:
            questions.append(f"{ref}: the approval '{v}' isn't one of the choices ({', '.join(labels)}). Which did you mean?")
            edits.append((f"{ref}: approval '{v}' not understood", "Treated as Pending until you say which choice you meant."))
            v = W.PENDING
        approvals[ref] = v
    for ref in sorted(entry_refs):
        i = items_by_ref[ref]
        prop = st["proposed"].get(ref)
        got = je.get(ref)
        fx = W.fixed_option(i, r0)
        if prop is None:  # an entry this upload didn't have before (only after corrected exports)
            continue
        if got is None:
            edits.append((f"Journal entries: entry {ref} missing", "Restored from the proposal."))
            entries[ref] = {"lines": [dict(x) for x in prop["lines"]], "je_number": "", "comment": ""}
            continue
        lines = [dict(x) for x in got["lines"]]
        for x in lines:
            if x.pop("unreadable", False):
                questions.append(f"{ref}: an amount on the line '{x['account']}' can't be read as a number. What should it be?")
        fixed_lines = [x for x in prop["lines"] if x["fixed"]]
        if fx:  # an entry approved in an earlier period: its lines are fixed
            if [(x["account"], x["debit"], x["credit"]) for x in lines] != [(x["account"], x["debit"], x["credit"]) for x in prop["lines"]]:
                edits.append((f"Journal entries {ref}: lines of an entry approved earlier were changed", "Not applied: restored."))
            lines = [dict(x) for x in prop["lines"]]
        else:
            for fl in fixed_lines:
                hit = next((x for x in lines if x["account"].casefold() == fl["account"].casefold() and x["debit"] == fl["debit"]
                            and x["credit"] == fl["credit"]), None)
                if hit:
                    hit["fixed"] = True
                    hit["description"] = hit["description"] or fl["description"]
                else:
                    changed = next((x for x in lines if x["account"].casefold() == fl["account"].casefold()), None)
                    edits.append((f"Journal entries {ref}: the cash line was changed"
                                  + (f" ({changed['debit'] or changed['credit']})" if changed else " or removed"),
                                  "Not applied: the cash line comes from the bank or ledger line. Restored."))
                    if changed:
                        lines.remove(changed)
                    pos = prop["lines"].index(fl)
                    lines.insert(min(pos, len(lines)), dict(fl))
            for x in lines:
                x.setdefault("fixed", False)
            orig_other = [x for x in prop["lines"] if not x["fixed"]]
            other = [x for x in lines if not x["fixed"]]
            if len(other) > len(orig_other):
                edits.append((f"Journal entries {ref}: split into {len(other)} lines", "Kept."))
            for x in other:
                o_ = next((y for y in orig_other), None)
                if o_ and (x["account"] != o_["account"] or x["debit"] != o_["debit"] or x["credit"] != o_["credit"]) and len(other) == 1:
                    edits.append((f"Journal entries {ref}: changed to {x['account']} {x['debit'] or x['credit']}", "Kept."))
        for fieldname, want, have in (("title", prop["title"], got.get("title")), ("proposal", prop["proposal"], got.get("proposal"))):
            if text(want).strip() != text(have).strip():
                edits.append((f"Journal entries {ref}: {fieldname} changed", "Not applied: restored."))
        entries[ref] = {"lines": lines, "je_number": got.get("je_number") or "", "comment": got.get("comment") or ""}
        if got.get("comment"):
            edits.append((f"Journal entries {ref}: comment \"{got['comment']}\"", "Kept."))

    # 6. what blocks completion, per entry and item
    m_probe = W.build_model(r0, {"refs": st["refs"], "approvals": approvals, "entries": entries, "accounts": accounts,
                                 "proposed": st["proposed"], "files": files})
    checks = {e["ref"]: W.entry_check(m_probe, e) for e in m_probe["entries"]}
    new_accounts = sorted({x["account"] for e in m_probe["entries"] for x in e["lines"]
                           if x.get("account") and not str(x["account"]).startswith("[") and x["account"].casefold() not in acct_set
                           and (W.is_approved(e["approval"]) or e["fixed"])})
    for a in new_accounts:
        questions.append(f"'{a}' isn't in your chart of accounts. Add it as a new account, or choose another in the workbook?")
    for e in m_probe["entries"]:
        c = checks[e["ref"]]
        if c == "Not approved":
            blocking.append(f"{e['ref']} has a JE number but isn't approved. Approve it, or remove the JE number.")
        elif c == "Remove JE number":
            blocking.append(f"{e['ref']}: '{e['approval']}' needs no journal entry, so it shouldn't have a JE number. Remove it, "
                            "or choose the treatment with an entry.")
        elif W.is_approved(e["approval"]) and c in ("Needs an account", "Not balanced"):
            blocking.append(f"{e['ref']}: {c.lower()}.")
    pending = [ref for ref, v in approvals.items() if v == W.PENDING]
    no_treatment = [ref for ref, i in items_by_ref.items() if not W.fixed_option(i, r0) and not W.options_for(i, r0)]
    if (record and any(W.is_approved(v) for v in approvals.values())) and not reconciled_by:
        raise InputError("Reconciled by is blank. Ask who is reconciling, put that name in Reconciled by, and validate again.")

    # 7. decisions from explicit approvals only
    prev = {(d["Ref"], W.OLD_LABELS.get(d.get("Decision"), d.get("Decision"))): d for d in st["recorded"]}
    decisions, recorded = [], []
    for ref, v in approvals.items():
        if not W.is_approved(v):
            continue
        i = items_by_ref[ref]
        o = next(o for o in W.options_for(i, r0) if o["label"] == v)
        lines = (entries.get(ref) or {}).get("lines") or []
        others = [x for x in lines if not x.get("fixed") and (x.get("account") or "")]
        account = others[0]["account"] if len(others) == 1 else ("Split: " + "; ".join(x["account"] for x in others) if others else W.PLACEHOLDER)
        if o["treatment"] in ("accept_outstanding", "accept_in_transit", "accept_outstanding_payment", "bank_error_to_correct",
                              "second_payment_outstanding", "posting_of_carried"):
            account = ""
        p = prev.get((ref, v))
        when = p["Approved on"] if p and p.get("Approved by") == reconciled_by else today
        decisions.append({"item_id": i.id, "treatment": o["treatment"], "reviewed_amount": str(i.amount), "line_ref": o["line_ref"],
                          "account": account, "note": "", "approved_by": reconciled_by, "approved_on": when, "fingerprint": i.fingerprint})
        bank_desc = next((r0["by_id"][x].description for x in i.lines if r0["by_id"][x].source == "Bank"), "")
        recorded.append({"Ref": ref, "Decision": v, "Approved by": reconciled_by, "Approved on": when, "Type": i.type,
                         "Treatment": o["treatment"], "Account": account, "Bank description": bank_desc, "Fingerprint": i.fingerprint})
    for d in st["recorded"]:
        if (d["Ref"], W.OLD_LABELS.get(d.get("Decision"), d.get("Decision"))) not in {(x["Ref"], x["Decision"]) for x in recorded}:
            edits.append((f"{d['Ref']}: earlier approval '{d.get('Decision')}' withdrawn", "Recorded as withdrawn; the item is open again."))

    out = {"recorded": False, "edits": [{"where": a, "done": b} for a, b in edits], "blocking": blocking, "questions": questions,
           "approved": sorted(ref for ref, v in approvals.items() if W.is_approved(v)), "pending": sorted(pending),
           "no_treatment": sorted(no_treatment), "reconciled_by": reconciled_by,
           "review_completed_on": review_on.isoformat() if review_on else None, "migrated_from_050": migrating}
    if not record:
        after = sha(upload)
        out.update(uploaded_file_unchanged=before == after, note="Nothing was recorded and no workbook was written (no --record).")
        return out

    (work / "decisions.json").write_text(json.dumps(decisions, indent=1), encoding="utf-8")
    r = bank_rules.run(period_path, refs=iref)
    if r["issues"]:
        raise InputError("Some approvals could not be applied: " + "; ".join(f"{x['item_id']}: {x['why']}" for x in r["issues"]))

    reasons = list(blocking) + questions
    if pending:
        reasons.append(f"{len(pending)} item(s) still pending: {', '.join(sorted(pending))}")
    if no_treatment:
        reasons.append(f"{len(no_treatment)} item(s) have no treatment here: {', '.join(sorted(no_treatment))}")
    if not r["status"].startswith("Reconciled"):
        reasons.append(r["status"])
    if reasons:
        completion = "Not complete: " + reasons[0]
    else:
        stamp = datetime.strptime(today, "%Y-%m-%d").strftime("%m/%d/%Y")
        m_final = W.build_model(r, {"refs": st["refs"], "approvals": approvals, "entries": entries, "accounts": accounts,
                                    "proposed": st["proposed"], "files": files})
        final_checks = {e["ref"]: W.entry_check(m_final, e) for e in m_final["entries"]}
        need = [e for e in m_final["entries"] if W.is_approved(e["approval"]) and final_checks[e["ref"]] != "Not needed"]
        to_post = [e for e in need if not e["je_number"]]
        reported = [e for e in need if e["je_number"]]
        if to_post:
            completion = (f"Complete subject to posting ({stamp}): {len(to_post)} approved entr{'y' if len(to_post) == 1 else 'ies'} "
                          f"without a JE number yet" + (f"; {len(reported)} reported posted, to be confirmed in next period's ledger"
                                                        if reported else ""))
        elif reported:
            completion = f"Complete ({stamp}): every entry reported posted; next period's ledger confirms them"
        else:
            completion = f"Complete ({stamp})"
    ctx = {"refs": st["refs"], "approvals": approvals, "entries": entries, "reconciled_by": reconciled_by, "accounts": accounts,
           "accounts_source": run.get("Accounts source") or "", "edits": edits or [("No edits needed attention", "")],
           "recorded": recorded, "validated_on": today, "completion": completion, "proposed": st["proposed"], "files": files,
           "created_on": run.get("Draft created"), "version": version, "check_engine": True,
           "validated_at": datetime.now(timezone.utc).isoformat(timespec="seconds"), "review_completed_on": review_on,
           "migrated_from": "skill 0.5.0 (review-1)" if migrating else ""}
    path = Path(out_dir) / W.out_name(r)
    if path.resolve() == upload.resolve():
        path = path.with_name(path.stem + "-checked.xlsx")
    W.build(r, ctx, path)
    fig = r["figures"]
    m = W.build_model(r, ctx)
    fx = W.figures(m)
    after = sha(upload)
    out.update(recorded=True, workbook=str(path), completion=completion, uploaded_file_unchanged=before == after,
               figures={"bank_closing": fig.get("bank_closing"), "ledger_closing": fig.get("ledger_closing"),
                        "remaining_difference": fx.get("remaining")},
               passed=not blocking and not questions)
    if review_on and not completion.startswith("Complete"):
        out["review_date_note"] = (f"Review completed on {review_on:%m/%d/%Y} is recorded, but it doesn't complete the "
                                   "reconciliation while items are unresolved.")
    return out
