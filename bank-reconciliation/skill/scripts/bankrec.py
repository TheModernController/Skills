"""Bank reconciliation commands (0.5.1). Claude runs these; the user installs nothing.

  inspect FILE... [--sheet NAME]     describe a file's layout before asking the user anything
  chart FILE [--sheet NAME]          read a chart of accounts export
  profile PROFILE [--previous OLD --change TEXT --write OUT]
                                     validate the setup Claude wrote from the user's answers, and version a change
  preview PERIOD                     apply the mapping and show what was read, before the run
  run PERIOD                         reconcile and write the review workbook (first run or next period)
  validate WORKBOOK [--record] [--confirm-account NAME]... [--accept-restore] [--bank FILE --ledger FILE] [--from-0.5.0]
                                     read a reviewed workbook back and validate the edits. Only --record records approvals
                                     and writes the checked workbook; use it only when the user asked to validate
  carry WORKBOOK --work DIR          start a next period from last period's checked workbook (no separate setup file)
  check WORKBOOK [--diagnostics DIR] verify a workbook independently; recalculates with LibreOffice where available
  selftest [--out DIR]               build a fictional workbook, apply scripted edits, recalculate with LibreOffice, and
                                     compare every formula with the script's values

All figures Claude reports must come from this script's output. Exit code 1 means stopped: read the message.
"""
from __future__ import annotations

import argparse, json, sys, warnings
from datetime import date
from decimal import Decimal
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from foundation import profiles, tables  # noqa: E402
from foundation.tables import InputError, LayoutError  # noqa: E402


def emit(obj):
    print(json.dumps(obj, indent=1, default=lambda v: str(v) if isinstance(v, (Decimal, date)) else repr(v), ensure_ascii=False))


def version():
    from check_env import SKILL_VERSION
    return SKILL_VERSION


def cmd_inspect(a):
    emit([tables.inspect_file(p, a.sheet) for p in a.files])


def cmd_chart(a):
    from foundation.chart import read_chart
    acc = read_chart(a.file, a.sheet)
    emit({"accounts": len(acc), "first": acc[:8], "last": acc[-3:]})


def cmd_profile(a):
    p = profiles.load(a.profile)
    problems = profiles.validate(p)
    out = {"valid": not problems, "problems": problems, "read_back": profiles.describe(p)}
    if a.previous:
        old = profiles.load(a.previous)
        out["changes"] = profiles.diff(old, p)
        if a.write:
            if not a.change:
                sys.exit("Stopped: --change is required with --write, to log what changed and why")
            new = profiles.bumped(old, p, a.change, a.on or date.today().isoformat())
            profiles.save(new, a.write)
            out["written"] = a.write
            out["profile_version"] = new["profile_version"]
    emit(out)
    if problems:
        sys.exit(1)


def cmd_preview(a):
    from bank_rules import load_period
    import bank_rules
    period = load_period(a.period)
    prof = profiles.load(period["profile"])
    problems = profiles.validate(prof)
    if problems:
        sys.exit("Stopped: setup is not ready: " + "; ".join(problems))
    out = {}
    files = {side: tables.read_source(period[key], prof["sources"][side], side.title(), prefix, period["period_start"], period["period_end"],
                                      prof["account"]["currency"])
             for side, key, prefix in (("bank", "bank_file", "BK"), ("ledger", "ledger_file", "CB"))}
    if period.get("prior_workbook"):
        bank_rules.mark_repeats(bank_rules.read_prior(period["prior_workbook"], period, prof), files["bank"], files["ledger"])
    for side, f in files.items():
        inc = [l for l in f.rows if l.included]
        amts = [l.amount for l in inc if l.amount is not None]
        out[side] = {"file": f.name, "rows": len(f.rows), "included": len(inc),
                     "excluded": [{"row": l.row, "reason": l.reason} for l in f.rows if not l.included],
                     "inflows": sum((x for x in amts if x > 0), Decimal("0.00")), "outflows": sum((x for x in amts if x < 0), Decimal("0.00")),
                     "dates": [min((l.date for l in inc if l.date), default=None), max((l.date for l in inc if l.date), default=None)],
                     "problems": [{"row": l.row, "problems": l.problems} for l in inc if l.problems],
                     "totals_rows": [{"check": c[0], "total_row": c[1], "rows_read": c[2]} for c in f.checks],
                     "layout_notes": f.layout_notes, "unmapped_type_values": f.unmapped_types}
        if side == "bank":
            out[side]["running_balance"] = tables.balance_continuity(f.rows)
    import coding
    hist, sources = coding.load_history(period, prof)
    if sources:
        out["coded_history"] = {"sources": sources, "lines": len(hist), "accounts": sorted({h["account"] for h in hist}),
                                "note": "Used only for suggestions."}
    emit(out)


def stopped(e):
    if isinstance(e, LayoutError):
        emit({"stopped": str(e), "layout_change": {"source": e.source, "missing_mapped_columns": e.missing, "new_columns": e.new,
                                                   "columns_found": e.header}})
    else:
        emit({"stopped": str(e)})
    sys.exit(1)


def accounts_for(period, prof, sug):
    """The account list for the Accounts tab, and where it came from."""
    from foundation.chart import read_chart
    out, source = [], ""
    if period.get("chart_file"):
        out = read_chart(period["chart_file"])
        source = f"uploaded {Path(period['chart_file']).name}"
    elif period.get("accounts_from"):
        from openpyxl import load_workbook
        import review_state
        out = review_state.accounts_of(load_workbook(period["accounts_from"]))
        rec = review_state.run_record(load_workbook(period["accounts_from"]))
        source = f"kept from the {rec.get('Period end', '')[:7]} workbook"
    else:
        source = "no chart uploaded"
        for a in (prof["accounts"].get("cash"), prof["accounts"].get("bank_fees")):
            if a and a not in out:
                out.append(a)
        for line in ((sug or {}).get("lines") or []):
            for h in line.get("history_lines") or []:
                if h.get("account") and h["account"] not in out:
                    out.append(h["account"])
    added = []
    for a in period.get("accounts_add") or []:
        if a not in out:
            out.append(a)
            added.append(a)
    if added:
        source += f"; added: {', '.join(added)}"
    return out, source


def setup_changes(period, prof):
    """Confirmed settings that differ from last period's checked workbook, so a changed rule is always flagged.

    Empty on a first run, or when nothing changed. Reported in the run output only; the workbook keeps the change log
    in the Setup tab's technical settings, as before.
    """
    if not period.get("prior_workbook"):
        return {}
    from openpyxl import load_workbook
    import review_state
    old = review_state.setup_of(load_workbook(period["prior_workbook"]))
    diffs = [d for d in profiles.diff(old, prof) if not d.startswith("profile_version")]
    out = {"rules": [d for d in diffs if d.startswith(("rules.", "coding_rules"))],
           "mapping": [d for d in diffs if d.startswith("sources.")]}
    out["other"] = [d for d in diffs if d not in out["rules"] and d not in out["mapping"]]
    out["logged_as"] = [c.get("change") for c in prof.get("change_log", []) if c not in old.get("change_log", [])]
    return {k: v for k, v in out.items() if v}


def source_files(r):
    out = {}
    for side, f in (("bank", r["bank_file"]), ("ledger", r["ledger_file"])):
        _, grid, _ = tables.read_grid(f.path, f.sheet or None)
        out[side] = {"name": f.name, "grid": grid, "sha": f.sha, "type": Path(f.path).suffix.lower().lstrip("."), "sheet": f.sheet,
                     "header_row": f.header_row}
    return out


def cmd_run(a):
    import bank_rules, coding, review_workbook as W
    r = bank_rules.run(a.period)
    sug = coding.suggest(r, a.period)  # optional; reads the finished run and never changes it
    period = r["period"]
    accounts, source = accounts_for(period, r["profile"], sug)
    open_n = sum(1 for i in r["items"] if i.status in ("Open", "Rejected"))
    if r["status"].startswith("Incomplete"):
        completion = "Not complete: " + r["status"][len("Incomplete: "):]
    elif open_n:
        completion = f"Not complete: {open_n} item(s) need your review"
    else:
        completion = "Not complete: upload the reviewed workbook for validation"
    ctx = {"accounts": accounts, "accounts_source": source, "coding": sug, "files": source_files(r),
           "reconciled_by": r["profile"].get("process", {}).get("reconciled_by") or r["profile"].get("process", {}).get("approver") or "",
           "created_on": (a.on or date.today().isoformat()), "completion": completion, "version": version()}
    path, m = W.build(r, ctx, Path(period["out_dir"]) / W.out_name(r))
    fx = W.figures(m)
    fig = r["figures"]
    failed = [c[0] for c in r["checks"] if not c[3]]
    out = {"workbook": str(path), "completion": completion, "status": r["status"],
           "bank_closing": fig.get("bank_closing"), "ledger_closing": fig.get("ledger_closing"), "remaining_difference": fx.get("remaining"),
           "statement_check": "agrees" if not fig.get("statement_difference") else f"differs by {fig.get('statement_difference')}",
           "items_to_review": open_n, "matched_pairs": len(r["matches"]), "failed_checks": failed, "notes": r["notes"],
           "rows_read": {"bank": sum(1 for l in r["bank_file"].rows if l.included), "ledger": sum(1 for l in r["ledger_file"].rows if l.included)},
           "excluded_rows": sum(1 for f in (r["bank_file"], r["ledger_file"]) for l in f.rows if not l.included),
           "carried_in": [{"item_id": x, "outcome": y, "detail": z} for x, y, z in r["carried_out"]] +
                         [{"item_id": p.id, "outcome": p.status, "detail": p.note} for p in r["pending"]] if r.get("prior") else [],
           "accounts": {"count": len(accounts), "source": source}}
    changes = setup_changes(period, r["profile"])
    if changes:
        out["setup_changes"] = changes
    Path(a.period).resolve().parent.joinpath("run-summary.json").write_text(json.dumps(out, indent=1, default=str), encoding="utf-8")
    emit(out)


def cmd_validate(a):
    import validate
    res = validate.validate(a.workbook, a.work, a.out, record=a.record, confirm_accounts=a.confirm_account or [],
                            accept_restore=a.accept_restore, bank_file=a.bank, ledger_file=a.ledger, on=a.on, version=version(),
                            from_050=a.from_050)
    emit(res)


def cmd_carry(a):
    """Read last period's checked workbook: the setup, accounts, and what the next period starts from."""
    from datetime import datetime, timedelta
    import calendar
    from openpyxl import load_workbook
    import review_state as S
    from foundation.sections import read_kv, read_section
    wb = load_workbook(a.workbook)
    if not S.is_review_workbook(wb):
        sys.exit(emit({"stopped": "This isn't a checked workbook from this version of the skill. Start with a first reconciliation."}) or 1)
    prof = S.setup_of(wb)
    rec = S.run_record(wb)
    facts = read_kv(wb[S.CHECKS], S.S_CARRY_FACTS)
    if not str(rec.get("Completion") or "").startswith("Complete") or not facts:
        emit({"stopped": f"Last period's workbook says '{rec.get('Completion')}'. Upload the checked workbook Claude returned after "
                         "validation, once it says Complete."})
        sys.exit(1)
    work = Path(a.work).resolve()
    work.mkdir(parents=True, exist_ok=True)
    profiles.save(prof, work / "profile.json")
    pe = datetime.strptime(str(rec["Period end"])[:10], "%Y-%m-%d").date()
    ns = pe + timedelta(days=1)
    ne = ns.replace(day=calendar.monthrange(ns.year, ns.month)[1]) if pe.day == calendar.monthrange(pe.year, pe.month)[1] else None
    carry = read_section(wb[S.CHECKS], S.S_CARRY)
    emit({"setup_written": str(work / "profile.json"), "read_back": profiles.describe(prof),
          "previous_period": f"{rec.get('Period start')} to {rec.get('Period end')}", "previous_completion": rec.get("Completion"),
          "next_period": {"period_start": str(ns), "period_end": str(ne) if ne else None},
          "statement_opening_from_previous_closing": facts.get("Stated bank closing balance"),
          "ledger_opening_from_previous_actual_closing": facts.get("Actual ledger closing balance"),
          "carried_items": len(carry), "reconciled_by": prof.get("process", {}).get("reconciled_by") or "",
          "accounts": {"count": len(S.accounts_of(wb)), "source": rec.get("Accounts source")},
          "period_file_keys": {"profile": str(work / "profile.json"), "prior_workbook": str(Path(a.workbook)),
                               "accounts_from": str(Path(a.workbook))}})


def cmd_check(a):
    import check_workbook
    res = check_workbook.check(a.workbook, a.diagnostics)
    emit(res)
    if not res["ok"]:
        sys.exit(1)


def cmd_selftest(a):
    import selftest
    res = selftest.run(a.out)
    emit(res)
    if not res["ok"]:
        sys.exit(1)


def main(argv=None):
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = p.add_subparsers(dest="cmd", required=True)
    s = sub.add_parser("inspect"); s.add_argument("files", nargs="+"); s.add_argument("--sheet"); s.set_defaults(fn=cmd_inspect)
    s = sub.add_parser("chart"); s.add_argument("file"); s.add_argument("--sheet"); s.set_defaults(fn=cmd_chart)
    s = sub.add_parser("profile"); s.add_argument("profile"); s.add_argument("--previous"); s.add_argument("--change"); s.add_argument("--write")
    s.add_argument("--on"); s.set_defaults(fn=cmd_profile)
    s = sub.add_parser("preview"); s.add_argument("period"); s.set_defaults(fn=cmd_preview)
    s = sub.add_parser("run"); s.add_argument("period"); s.add_argument("--on"); s.set_defaults(fn=cmd_run)
    s = sub.add_parser("validate"); s.add_argument("workbook"); s.add_argument("--record", action="store_true")
    s.add_argument("--confirm-account", action="append"); s.add_argument("--accept-restore", action="store_true")
    s.add_argument("--bank"); s.add_argument("--ledger"); s.add_argument("--work", default="/tmp/bankrec/validate")
    s.add_argument("--out", default="/mnt/user-data/outputs"); s.add_argument("--on")
    s.add_argument("--from-0.5.0", dest="from_050", action="store_true",
                   help="convert a 0.5.0 workbook explicitly, listing each conversion"); s.set_defaults(fn=cmd_validate)
    s = sub.add_parser("carry"); s.add_argument("workbook"); s.add_argument("--work", default="/tmp/bankrec"); s.set_defaults(fn=cmd_carry)
    s = sub.add_parser("check"); s.add_argument("workbook"); s.add_argument("--diagnostics"); s.set_defaults(fn=cmd_check)
    s = sub.add_parser("selftest"); s.add_argument("--out"); s.set_defaults(fn=cmd_selftest)
    a = p.parse_args(argv)
    # Excel keeps some formatting in extensions openpyxl doesn't read. They are rebuilt in every workbook Claude returns.
    warnings.filterwarnings("ignore", message=".*extension is not supported.*")
    try:
        a.fn(a)
    except InputError as e:
        stopped(e)
    except ModuleNotFoundError as e:
        emit({"stopped": f"A required Python package is missing here: {e.name}. Run check_env.py for an explanation."})
        sys.exit(1)
    except (ValueError, KeyError, OSError) as e:
        emit({"stopped": f"{type(e).__name__}: {e}"})
        sys.exit(1)


if __name__ == "__main__":
    main()
