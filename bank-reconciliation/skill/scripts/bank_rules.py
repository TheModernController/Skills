"""Bank reconciliation rules: matching, review items, decisions, carry-forward, balances, and status.

Ported from the episode 002 pilot (automation/reconcile.py), which stays unchanged. Rules:
  R1  same reference, same amount, one candidate each side (equal numbers of identical lines pair by date)
  R2  same direction and amount, dates within the window, no conflicting references, unique candidates
  H1  same reference, different amount                     -> held for review
  H2  more than one possible partner                       -> held for review
  P1  ledger line carrying a carried adjustment's posting reference and amount -> that adjustment is posted
  C1  unmatched ledger check                               -> proposed outstanding check
  C2  unmatched ledger receipt in the last N days          -> proposed deposit in transit
  C3  unmatched bank line on the fee description list      -> proposed bank fee entry
  C4  anything else unmatched                              -> held for review
  D1  source row with a missing or unreadable date or amount -> held; fix the export and rerun
Nothing moves the reconciliation until the user approves it. There is no balancing line.
"""
from __future__ import annotations

import json
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from decimal import Decimal
from pathlib import Path

from foundation import decisions as dec
from foundation import profiles, tables
from foundation.tables import InputError, Row

ZERO = Decimal("0.00")
BANK_CATEGORIES = ["Deposit in transit", "Outstanding check", "Outstanding payment", "Bank error"]
BOOK_CATEGORIES = ["Bank fee", "Amount correction", "Duplicate reversal", "Returned item", "Other bank item"]
CARRIED_TYPES = {"deposit_in_transit": ("Carried deposit in transit", "Bank", "Deposit in transit"),
                 "outstanding_check": ("Carried outstanding check", "Bank", "Outstanding check"),
                 "outstanding_payment": ("Carried outstanding payment", "Bank", "Outstanding payment"),
                 "bank_error": ("Carried bank error", "Bank", "Bank error"),
                 "pending_adjustment": ("Carried adjustment not yet posted", "Ledger", "")}
KIND_OF_CATEGORY = {"Deposit in transit": "deposit_in_transit", "Outstanding check": "outstanding_check",
                    "Outstanding payment": "outstanding_payment", "Bank error": "bank_error"}
ALLOWED = {
    "Proposed outstanding check": {"accept_outstanding", "reject"},
    "Proposed deposit in transit": {"accept_in_transit", "reject"},
    "Proposed bank fee entry": {"accept_fee_entry", "reject"},
    "Amount mismatch": {"correct_cashbook_to_bank", "bank_error_to_correct"},
    "Multiple candidates": {"reverse_duplicate", "second_payment_outstanding"},
    "Returned item": {"reverse_application_reopen_receivable"},
    "Unidentified bank item": {"book_bank_item"},
    "Unmatched ledger payment": {"accept_outstanding_payment", "posting_of_carried"},
    "Unmatched ledger receipt": {"posting_of_carried"},
}
MEANING = {
    "accept_outstanding": "accept as an outstanding check",
    "accept_in_transit": "accept as a deposit in transit",
    "accept_fee_entry": "accept the bank fee; propose a ledger entry",
    "reject": "reject the proposal; the item stays unresolved",
    "correct_cashbook_to_bank": "the bank is right; propose a ledger correction",
    "bank_error_to_correct": "the bank is wrong; show as a bank error until the bank corrects it",
    "reverse_duplicate": "a duplicate ledger entry; propose reversing the named line",
    "second_payment_outstanding": "a genuine second payment; treat the named line as outstanding",
    "reverse_application_reopen_receivable": "returned payment; propose reversing the receipt and reopening the receivable",
    "book_bank_item": "propose a ledger entry for this bank line",
    "accept_outstanding_payment": "accept as an outstanding payment (not yet cleared by the bank)",
    "posting_of_carried": "this ledger line is the posting of the named carried adjustment",
}


@dataclass
class Item:
    id: str
    type: str
    rule: str
    lines: list
    amount: Decimal
    note: str = ""
    side: str = ""
    category: str = ""
    recipe: list = field(default_factory=list)  # (sign, line id) pairs that define the effect
    fixed_effect: Decimal | None = None          # carried items, whose lines belong to an earlier period
    status: str = "Open"
    treatment: str = ""
    possible_original: str = ""
    suggestion: str = ""
    origin: str = ""
    fingerprint: str = ""
    posting_ref: str = ""
    account: str = ""
    approved_by: str = ""
    approved_on: str = ""
    description: str = ""
    reference: str = ""
    line_date: str = ""


def d(v):
    return None if v in (None, "") else Decimal(str(v)).quantize(Decimal("0.01"))


def money(v):
    return f"{v:,.2f}"


# ---------------------------------------------------------------- period file

def load_period(path):
    p = Path(path)
    data = json.loads(p.read_text(encoding="utf-8"))
    base = p.parent

    def resolve(key):
        v = data.get(key)
        return str((base / v).resolve()) if v else None

    for k in ("period_start", "period_end"):
        if not data.get(k):
            raise InputError(f"Period file: {k} is required")
    out = dict(data)
    out["period_start"] = datetime.strptime(data["period_start"], "%Y-%m-%d").date()
    out["period_end"] = datetime.strptime(data["period_end"], "%Y-%m-%d").date()
    if out["period_end"] < out["period_start"]:
        raise InputError("Period file: period_end is before period_start")
    for k in ("profile", "bank_file", "ledger_file", "prior_workbook", "opening_items", "decisions", "out_dir",
              "history_file", "coding_decisions",  # optional, for coding suggestions only
              "prior_embedded", "chart_file", "accounts_from"):  # 0.5.0: carried-in data kept in this period's workbook; chart of accounts
        out[k] = resolve(k)
    out["accounts_add"] = list(data.get("accounts_add", []))
    for k in ("profile", "bank_file", "ledger_file"):
        if not out[k]:
            raise InputError(f"Period file: {k} is required")
    for k in ("statement_opening", "statement_closing", "ledger_opening", "ledger_closing"):
        out[k] = d(data.get(k))
    out["posted_before_opening"] = list(data.get("posted_before_opening", []))
    out["out_dir"] = out["out_dir"] or str(base)
    return out


# ---------------------------------------------------------------- prior period

def canon(v):
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, (int, float, Decimal)) and not isinstance(v, bool):
        return str(Decimal(str(v)).quantize(Decimal("0.01")))
    return str(v)


CARRY_COLS = ["Kind", "Item ID", "Origin period", "Date", "Reference", "Description", "Amount", "Category", "Account", "Posting ref", "Approved by", "Note"]


def carry_digest(facts, rows):
    import hashlib
    blob = json.dumps([[k, canon(v)] for k, v in sorted(facts.items())] + [[canon(r.get(c)) for c in CARRY_COLS] for r in rows])
    return hashlib.sha256(blob.encode()).hexdigest()[:16]


def read_prior(path, period, profile, embedded=False):
    from foundation.xlsx import read_table
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=False)
    import review_state  # the 0.5.0 review workbook keeps its carry forward on the hidden Checks tab
    if review_state.is_review_workbook(wb):
        return review_state.read_prior(wb, path, period, profile, embedded)
    if embedded:
        raise InputError(f"{Path(path).name} is not a workbook from this version of the skill.")
    if "Carry forward" not in wb.sheetnames:
        raise InputError(f"{Path(path).name} has no Carry forward sheet. Use the workbook this skill produced for the previous period.")
    facts = {r["Field"]: r.get("Value") for r in read_table(wb, "Carry forward", "Field")}
    rows = read_table(wb, "Carry forward", "Kind")
    record = {r["Field"]: r.get("Value") for r in read_table(wb, "Run record", "Field")}
    edited = carry_digest(facts, rows) != record.get("Carry forward digest")
    pe = datetime.strptime(str(facts["Period end"])[:10], "%Y-%m-%d").date()
    ps = datetime.strptime(str(facts["Period start"])[:10], "%Y-%m-%d").date()
    if ps == period["period_start"] and pe == period["period_end"]:
        raise InputError("The previous-period workbook given is for this same period. Use the workbook for the period immediately before.")
    if pe + timedelta(days=1) != period["period_start"]:
        raise InputError(f"The previous-period workbook ends {pe}, but this period starts {period['period_start']}. "
                         "Use the workbook for the period immediately before.")
    if not str(facts.get("Status", "")).startswith("Reconciled"):
        raise InputError(f"The previous period's status is '{facts.get('Status')}'. Finish that reconciliation before carrying it forward.")
    acct = profile["account"]
    if facts.get("Account") != acct["nickname"] or facts.get("Currency") != acct["currency"]:
        raise InputError("The previous-period workbook is for a different account or currency than this profile.")
    if edited and not period.get("accept_edited_prior"):
        raise InputError("The previous period's Carry forward sheet was changed after that run. Ask the user whether the change was "
                         "intended; if so, set accept_edited_prior to true in the period file.")
    lines = {}
    for sheet in ("Bank rows", "Ledger rows"):
        for r in read_table(wb, sheet, "Line ID"):
            if r.get("Included") == "Yes":
                lines[(sheet, canon(r.get("Date")), str(r.get("Reference") or ""), str(r.get("Description") or ""), canon(r.get("Amount")))] = r["Line ID"]
    history = {}
    by_line = {r["Line ID"]: r for r in read_table(wb, "Bank rows", "Line ID")}
    for r in read_table(wb, "Review items", "Item ID"):
        if str(r.get("Status", "")).startswith("Approved") and r.get("Treatment"):
            for lid in str(r.get("Source lines") or "").split(", "):
                if lid in by_line:
                    history[str(by_line[lid].get("Description") or "").upper()] = (r.get("Type"), r.get("Treatment"), r.get("Account") or "")
    return dict(facts=facts, rows=rows, edited=edited, lines=lines, history=history, period_label=pe.strftime("%Y-%m"),
                files={record.get("Bank file SHA-256"), record.get("Ledger file SHA-256")} - {None},
                statement_closing=d(facts.get("Stated bank closing balance")),
                ledger_closing=d(facts.get("Actual ledger closing balance")))


def mark_repeats(prior, bank_file, ledger_file):
    """Rows dated before this period that the previous workbook already reconciled are confirmed repeated imports."""
    for f, sheet in ((bank_file, "Bank rows"), (ledger_file, "Ledger rows")):
        for r in f.rows:
            if not r.included and r.reason.startswith("Dated outside"):
                key = (sheet, str(r.date), r.reference, r.description, str(r.amount))
                if key in prior["lines"]:
                    r.reason = f"Repeated import: already in the {prior['period_label']} reconciliation as {prior['lines'][key]}"


def read_opening_items(path):
    rows = tables.read_grid(path)[1]
    head = [tables.text(c) for c in rows[0]]
    out = []
    for r in rows[1:]:
        if not any(tables.text(c) for c in r):
            continue
        rec = {h: (r[j] if j < len(r) else None) for j, h in enumerate(head)}
        kind = tables.text(rec.get("kind"))
        if kind not in CARRIED_TYPES:
            raise InputError(f"Opening items: kind must be one of {', '.join(CARRIED_TYPES)}, not {kind!r}")
        out.append({"Kind": kind, "Item ID": f"OPEN-{len(out) + 1:03d}", "Origin period": "before first run",
                    "Date": tables.text(rec.get("date")), "Reference": tables.text(rec.get("reference")).upper(),
                    "Description": tables.text(rec.get("description")), "Amount": tables.text(rec.get("amount")),
                    "Category": tables.text(rec.get("category")), "Account": tables.text(rec.get("account")),
                    "Posting ref": tables.text(rec.get("posting_ref")), "Approved by": "Opening list", "Note": tables.text(rec.get("note"))})
    return out


# ---------------------------------------------------------------- matching

def match_lines(bank, book, window, period_start):
    """Return (matches, held groups, used ids). Input order never affects the result."""
    bank = sorted(bank, key=lambda l: l.line_id)
    book = sorted(book, key=lambda l: l.line_id)
    matches, held, used = [], [], set()

    def pair_identical(b, c, rule, within=None):
        if len(b) == len(c) and len({l.amount for l in b + c}) == 1:
            pairs = list(zip(sorted(b, key=lambda l: (l.date, l.line_id)), sorted(c, key=lambda l: (l.date, l.line_id))))
            if within is None or all(within(x, y) for x, y in pairs):
                for x, y in pairs:
                    matches.append((x, y, rule, "Identical lines paired in date order" if len(b) > 1 else ""))
                return True
        return False

    groups = defaultdict(lambda: ([], []))
    for l in bank:
        if l.reference:
            groups[l.reference][0].append(l)
    for l in book:
        if l.reference:
            groups[l.reference][1].append(l)
    for ref in sorted(groups):
        b, c = groups[ref]
        if not b or not c:
            continue
        used.update(l.line_id for l in b + c)
        if pair_identical(b, c, "R1"):
            continue
        held.append(("H1" if len(b) == 1 and len(c) == 1 else "H2", b + c))

    def anchor(l):
        return max(l.date, period_start) if l.source == "Carried" else l.date

    def within(b, c):
        return abs((b.date - anchor(c)).days) <= window

    free_bank = [l for l in bank if l.line_id not in used]
    free_book = [l for l in book if l.line_id not in used]

    def candidate(b, c):
        conflict = bool(b.reference and c.reference and b.reference != c.reference)
        return (b.amount > 0) == (c.amount > 0) and not conflict and b.amount == c.amount and within(b, c)

    links = {l.line_id: [] for l in free_bank + free_book}
    for b in free_bank:
        for c in free_book:
            if candidate(b, c):
                links[b.line_id].append(c)
                links[c.line_id].append(b)
    seen = set()
    for b in free_bank:
        if b.line_id in seen or not links[b.line_id]:
            continue
        group, stack = {}, [b]
        while stack:
            l = stack.pop()
            if l.line_id not in group:
                group[l.line_id] = l
                stack.extend(links[l.line_id])
        seen.update(group)
        members = sorted(group.values(), key=lambda l: l.line_id)
        gb = [l for l in members if l.source == "Bank"]
        gc = [l for l in members if l.source != "Bank"]
        if not pair_identical(gb, gc, "R2", within):
            held.append(("H2", members))
        used.update(group)
    return matches, held, used


# ---------------------------------------------------------------- run

def run(period_path, extra_decisions=None, refs=None):
    """refs (0.5.0): item id -> the stable Ref shown in the workbook (R-01). Entries then take their posting reference
    from it, so a reference never changes when decisions are added or changed."""
    period = load_period(period_path)
    profile = profiles.load(period["profile"])
    problems = profiles.validate(profile)
    if problems:
        raise InputError("Profile is not ready: " + "; ".join(problems))
    rules = profile["rules"]
    start, end = period["period_start"], period["period_end"]
    label = end.strftime("%Y-%m")
    ccy = profile["account"]["currency"]
    if Path(period["bank_file"]).is_file() and Path(period["ledger_file"]).is_file() and \
            tables.sha256(period["bank_file"]) == tables.sha256(period["ledger_file"]):
        raise InputError("The bank file and the ledger file are the same file.")
    bank_file = tables.read_source(period["bank_file"], profile["sources"]["bank"], "Bank", "BK", start, end, ccy)
    ledger_file = tables.read_source(period["ledger_file"], profile["sources"]["ledger"], "Ledger", "CB", start, end, ccy)
    if bank_file.sha == ledger_file.sha:
        raise InputError("The bank file and the ledger file are the same file.")

    prior = None
    carried = []
    if period.get("prior_embedded"):  # validating a next-period workbook: the previous period's carry forward is kept inside it
        prior = read_prior(period["prior_embedded"], period, profile, embedded=True)
        carried = prior["rows"]
        mark_repeats(prior, bank_file, ledger_file)
    elif period.get("prior_workbook"):
        prior = read_prior(period["prior_workbook"], period, profile)
        for f in (bank_file, ledger_file):
            if f.sha in prior["files"]:
                raise InputError(f"{f.name} is the same file used for {prior['period_label']}. This is a repeated import; ask for this period's export.")
        carried = prior["rows"]
        mark_repeats(prior, bank_file, ledger_file)
    elif period.get("opening_items"):
        carried = read_opening_items(period["opening_items"])

    bank_all = [r for r in bank_file.rows if r.included]
    book_all = [r for r in ledger_file.rows if r.included]
    ids = [r.line_id for r in bank_all + book_all]
    if len(set(ids)) != len(ids):
        raise InputError("A line id appears in both files. Line ids must be unique across the bank and ledger files.")
    valid = lambda r: r.date is not None and r.amount is not None and not r.problems
    bank = [r for r in bank_all if valid(r)]
    book = [r for r in book_all if valid(r)]

    items = []
    posted_before = set(period["posted_before_opening"])
    carried_lines, pending = [], []
    for c in carried:
        kind = c["Kind"]
        tname, side, category = CARRIED_TYPES[kind]
        amt = d(c["Amount"])
        cid = c["Item ID"] if str(c["Item ID"]).startswith(("CF-", "OPEN-")) else f"CF-{c['Item ID']}"
        when = datetime.strptime(str(c["Date"])[:10], "%Y-%m-%d").date() if c.get("Date") else start
        if kind == "pending_adjustment":
            it = Item(cid, tname, "", [], abs(amt), c.get("Note") or "", "Ledger", c.get("Category") or "", fixed_effect=amt,
                      status="Approved (carried)", origin=str(c.get("Origin period") or ""), posting_ref=str(c.get("Posting ref") or ""),
                      account=str(c.get("Account") or ""), approved_by=str(c.get("Approved by") or ""), description=str(c.get("Description") or ""),
                      reference=str(c.get("Reference") or ""), line_date=str(when))
            if cid in posted_before or c["Item ID"] in posted_before:
                it.status = "Posted before this period's opening"
            pending.append(it)
        else:
            line = Row(cid, "Carried", "previous workbook", "Carry forward", 0, when, str(c.get("Reference") or "").upper(),
                       str(c.get("Description") or ""), amt, kind="check" if kind == "outstanding_check" else "receipt" if amt > 0 else "payment")
            line.carried = c  # type: ignore[attr-defined]
            carried_lines.append(line)
    unknown_posted = posted_before - {p.id for p in pending} - {p.id[3:] for p in pending}
    if unknown_posted:
        raise InputError(f"posted_before_opening names items that are not carried adjustments: {', '.join(sorted(unknown_posted))}")

    # P1: postings of carried adjustments, recognized by their posting reference
    consumed = set()
    for p in pending:
        if p.status != "Approved (carried)" or not p.posting_ref:
            continue
        ref = p.posting_ref.upper()
        hits = [l for l in book if l.line_id not in consumed and (l.reference == ref or ref in l.description.upper())]
        if len(hits) == 1 and hits[0].amount == p.fixed_effect:
            l = hits[0]
            consumed.add(l.line_id)
            p.status = "Posted in this period"
            p.note = f"Posted by ledger line {l.line_id}"
            items.append(Item(f"PST-{l.line_id}", "Posting of carried adjustment", "P1", [l.line_id], abs(l.amount),
                              f"Posts {p.id} ({p.posting_ref})", status="Approved", treatment="posting reference found",
                              category="Posting of carried adjustment", possible_original=p.id))
        elif hits:
            for l in hits:
                consumed.add(l.line_id)
            items.append(Item(f"X-{hits[0].line_id}", "Carried adjustment posting does not agree", "P1", [l.line_id for l in hits],
                              abs(sum(l.amount for l in hits)), f"Ledger line(s) carry {p.posting_ref} but the amount or count differs from "
                              f"{p.id} ({money(p.fixed_effect)}). Correct the ledger or the export and rerun.", possible_original=p.id))

    window = rules["date_window_days"]
    match_book = [l for l in book if l.line_id not in consumed] + carried_lines
    matches, held, used = match_lines(bank, match_book, window, start)
    used |= consumed

    for rule, group in held:
        bt = sum(l.amount for l in group if l.source == "Bank")
        ct = sum(l.amount for l in group if l.source != "Bank")
        items.append(Item(f"X-{group[0].line_id}", "Amount mismatch" if rule == "H1" else "Multiple candidates", rule,
                          [l.line_id for l in group], abs(bt - ct),
                          "Same reference, different amount" if rule == "H1" else
                          "More than one possible partner, so nothing was matched" + (" (possible duplicate)" if len({l.amount for l in group}) == 1 else "")))

    transit_from = end - timedelta(days=rules["deposit_in_transit_days"] - 1)
    fees = {x.casefold() for x in rules.get("fee_descriptions", [])}
    returned_words = [w.upper() for w in rules.get("returned_item_words", ["RETURNED"])]
    hist = prior["history"] if prior else {}
    pending_open = [p for p in pending if p.status == "Approved (carried)"]
    for l in sorted(book, key=lambda l: l.line_id):
        if l.line_id in used:
            continue
        if l.kind == "check" and l.amount < 0:
            items.append(Item(f"P-{l.line_id}", "Proposed outstanding check", "C1", [l.line_id], abs(l.amount), "Proposal only", "Bank",
                              "Outstanding check", [(1, l.line_id)]))
        elif l.amount > 0 and l.kind == "receipt" and l.date >= transit_from:
            items.append(Item(f"P-{l.line_id}", "Proposed deposit in transit", "C2", [l.line_id], abs(l.amount), "Proposal only", "Bank",
                              "Deposit in transit", [(1, l.line_id)]))
        else:
            t = "Unmatched ledger payment" if l.amount < 0 else "Unmatched ledger receipt"
            it = Item(f"X-{l.line_id}", t, "C4", [l.line_id], abs(l.amount), "No matching bank line")
            hints = [f"Same amount as carried adjustment {p.id} ({p.posting_ref}), which has no posting reference here: it may be that posting "
                     f"(treatment posting_of_carried, line_ref {p.id})" for p in pending_open if p.fixed_effect == l.amount]
            if l.amount < 0 and l.date >= end - timedelta(days=4):
                hints.append("Dated near period end, so it may not have reached the bank yet (treatment accept_outstanding_payment)")
            it.suggestion = "; ".join(hints)
            items.append(it)
    for l in sorted(bank, key=lambda l: l.line_id):
        if l.line_id in used:
            continue
        h = hist.get(l.description.upper())
        basis = f"In {prior['period_label']}, '{l.description}' was approved as {h[0].lower()} ({h[1]}{', ' + h[2] if h[2] else ''}). Suggestion only." if h else ""
        if l.amount < 0 and l.description.casefold() in fees:
            items.append(Item(f"P-{l.line_id}", "Proposed bank fee entry", "C3", [l.line_id], abs(l.amount), "Proposal only", "Ledger",
                              "Bank fee", [(1, l.line_id)], suggestion=basis))
        elif any(w in l.description.upper() for w in returned_words):
            items.append(Item(f"X-{l.line_id}", "Returned item", "C4", [l.line_id], abs(l.amount), "Bank returned a payment", suggestion=basis))
        else:
            items.append(Item(f"X-{l.line_id}", "Unidentified bank item", "C4", [l.line_id], abs(l.amount), "No matching ledger line", suggestion=basis))
    for l in sorted([r for r in bank_all + book_all if not valid(r)], key=lambda r: r.line_id):
        items.append(Item(f"X-{l.line_id}", "Incomplete source row", "D1", [l.line_id], abs(l.amount or ZERO),
                          "; ".join(l.problems) + ". Correct the export and rerun; no treatment applies."))

    # carried bank-side items: cleared by a match, or still outstanding
    matched_ids = {x.line_id for m in matches for x in m[:2]}
    carried_out = []
    for cl in carried_lines:
        c = cl.carried
        tname, side, category = CARRIED_TYPES[c["Kind"]]
        if cl.line_id in matched_ids:
            partner = next(m[0].line_id for m in matches if m[1].line_id == cl.line_id)
            carried_out.append((cl.line_id, "Cleared", f"Cleared by bank line {partner}"))
            continue
        if cl.line_id in used:
            carried_out.append((cl.line_id, "Held for review", "Has more than one possible clearing line; see Review items"))
            continue
        items.append(Item(cl.line_id, tname, "CF", [], abs(cl.amount), str(c.get("Note") or ""), side, category, fixed_effect=cl.amount,
                          status="Approved (carried)", origin=str(c.get("Origin period") or ""), approved_by=str(c.get("Approved by") or ""),
                          description=cl.description, reference=cl.reference, line_date=str(cl.date)))
        carried_out.append((cl.line_id, "Still outstanding", f"Carried from {c.get('Origin period')}; {(end - cl.date).days} days old at period end"))

    by_id = {l.line_id: l for l in bank_all + book_all + carried_lines}
    for i in items:
        i.fingerprint = dec.fingerprint(i.type, [by_id[x].content() for x in i.lines]) if i.lines else dec.fingerprint(i.type, [(i.id, str(i.fixed_effect))])
        if i.type == "Returned item":
            target = by_id[i.lines[0]]
            i.possible_original = ", ".join(sorted(m[1].line_id for m in matches if m[1].amount == -target.amount and m[1].date <= target.date))
    items.sort(key=lambda i: i.id)

    decisions = dec.load(period.get("decisions"))
    if extra_decisions:
        decisions, _ = dec.merge(decisions, [dec.clean(x) for x in extra_decisions])
    issues = apply_decisions(items, decisions, by_id, matches, profile, pending)

    # dispositions
    for b, c, rule, note in matches:
        for l, other in ((b, c), (c, b)):
            l.disposition, l.rule, l.link = "Matched", rule, other.line_id
    for i in items:
        for x in i.lines:
            l = by_id[x]
            l.disposition, l.rule, l.link = f"Review item ({i.status})", i.rule, i.id
    for r in bank_file.rows + ledger_file.rows:
        if not r.included:
            r.disposition = "Excluded: " + r.reason

    # postings via decisions
    for i in items:
        if i.treatment == "posting_of_carried" and i.status == "Approved":
            target = next(p for p in pending if p.id == i.possible_original)
            target.status, target.note = "Posted in this period", f"Posted by ledger line {i.lines[0]} (confirmed by the user)"
    for p in pending:
        if p.status == "Approved (carried)":
            items.append(p)
    items.sort(key=lambda i: i.id)

    # proposed entries and posting references
    cash = profile["accounts"]["cash"]
    entries = []
    n = 0
    for i in items:
        if i.side == "Ledger" and i.status == "Approved":
            n += 1
            i.posting_ref = f"{label}-{refs[i.id]}" if refs and i.id in refs else f"REC-{label}-{n:02d}"
            eff = effect_of(i, by_id)
            dr, cr = (cash, i.account) if eff > 0 else (i.account, cash)
            entries.append(dict(ref=i.posting_ref, item=i.id, origin=label, debit=dr, credit=cr, amount=abs(eff), memo=f"{i.posting_ref} {i.note}".strip(),
                                status="New this period, not posted"))
    for p in pending:
        if p.status == "Approved (carried)":
            eff = p.fixed_effect
            dr, cr = (cash, p.account) if eff > 0 else (p.account, cash)
            entries.append(dict(ref=p.posting_ref, item=p.id, origin=p.origin, debit=dr, credit=cr, amount=abs(eff), memo=f"{p.posting_ref} {p.description}".strip(),
                                status=f"Carried from {p.origin}, still not posted"))

    # balances
    s_open = period["statement_opening"]
    l_open = period["ledger_opening"]
    notes = []
    if s_open is None and prior and prior["statement_closing"] is not None:
        s_open = prior["statement_closing"]
        notes.append("Bank opening balance taken from the previous period's stated closing balance; not independently checked.")
    if l_open is None and prior and prior["ledger_closing"] is not None:
        l_open = prior["ledger_closing"] + sum((p.fixed_effect for p in pending if p.status.startswith("Posted before")), ZERO)
        notes.append("Ledger opening balance derived from the previous period's actual closing balance; not independently checked.")
    bank_act = sum((l.amount for l in bank_all if l.amount is not None), ZERO)
    book_act = sum((l.amount for l in book_all if l.amount is not None), ZERO)
    approved = [i for i in items if i.status in ("Approved", "Approved (carried)")]
    bank_side = sum((effect_of(i, by_id) for i in approved if i.side == "Bank"), ZERO)
    book_side = sum((effect_of(i, by_id) for i in approved if i.side == "Ledger"), ZERO)
    fig = dict(statement_opening=s_open, statement_closing=period["statement_closing"], bank_activity=bank_act,
               ledger_opening=l_open, ledger_activity=book_act, ledger_closing_stated=period["ledger_closing"],
               bank_side=bank_side, book_side=book_side)
    if s_open is not None:
        fig["bank_closing"] = s_open + bank_act
        fig["statement_difference"] = fig["bank_closing"] - fig["statement_closing"] if fig["statement_closing"] is not None else None
    if l_open is not None:
        fig["ledger_closing"] = l_open + book_act
        fig["ledger_difference"] = fig["ledger_closing"] - fig["ledger_closing_stated"] if fig["ledger_closing_stated"] is not None else None
    if s_open is not None and l_open is not None:
        fig["raw_difference"] = fig["bank_closing"] - fig["ledger_closing"]
        fig["adjusted_bank"] = fig["bank_closing"] + bank_side
        fig["after_bank_side"] = fig["adjusted_bank"] - fig["ledger_closing"]
        fig["adjusted_ledger"] = fig["ledger_closing"] + book_side
        fig["final_difference"] = fig["adjusted_bank"] - fig["adjusted_ledger"]
        opening_bank = sum((cl.amount for cl in carried_lines), ZERO)
        opening_book = sum((p.fixed_effect for p in pending if not p.status.startswith("Posted before")), ZERO)
        fig["opening_difference"] = (s_open + opening_bank) - (l_open + opening_book)
    if prior:
        if period["statement_opening"] is not None and prior["statement_closing"] is not None:
            fig["bank_opening_vs_prior"] = period["statement_opening"] - prior["statement_closing"]
        if period["ledger_opening"] is not None and prior["ledger_closing"] is not None:
            fig["ledger_opening_vs_prior"] = period["ledger_opening"] - (prior["ledger_closing"] + sum(
                (p.fixed_effect for p in pending if p.status.startswith("Posted before")), ZERO))

    result = dict(period=period, profile=profile, label=label, bank_file=bank_file, ledger_file=ledger_file, prior=prior,
                  matches=sorted(matches, key=lambda m: m[0].line_id), items=items, issues=issues, entries=entries, by_id=by_id,
                  carried_lines=carried_lines, carried_out=carried_out, pending=pending, figures=fig, notes=notes,
                  decisions=decisions)
    result["checks"] = integrity_checks(result)
    result["status"], result["missing"] = status_of(result)
    return result


# ---------------------------------------------------------------- decisions

def apply_decisions(items, decisions, by_id, matches, profile, pending):
    """Apply valid decisions. Anything doubtful stays open and is listed as an issue."""
    issues = []
    known = {i.id: i for i in items}
    counts = defaultdict(int)
    for x in decisions:
        counts[x["item_id"]] += 1
    matched_book = {m[1].line_id for m in matches}
    linked = set()
    carried_pending = {p.id: p for p in pending}
    by_fp = {i.fingerprint: i.id for i in items}

    def issue(x, why):
        issues.append(dict(item_id=x["item_id"], treatment=x.get("treatment", ""), why=why))
        x["result"] = "Not applied: " + why

    for x in decisions:
        item = known.get(x["item_id"])
        if counts[x["item_id"]] > 1:
            issue(x, "Item has more than one decision, so none was applied")
            continue
        stale = dec.freshness(x, item)
        if stale:
            if item is None and x.get("fingerprint") in by_fp:
                stale += f" Item {by_fp[x['fingerprint']]} has the same content; confirm the decision for it instead."
            issue(x, stale)
            continue
        if item.status != "Open":
            issue(x, f"Item is already {item.status.lower()}; no decision is needed")
            continue
        t, ref, account = x["treatment"], x.get("line_ref", ""), x.get("account", "")
        if t not in ALLOWED.get(item.type, set()):
            issue(x, f"Treatment is not allowed for {item.type}")
            continue
        ids = item.lines
        if t == "reject":
            item.status, item.treatment, item.note = "Rejected", t, x.get("note") or "Rejected, still unresolved"
            x["result"] = "Applied"
            item.approved_by, item.approved_on = x.get("approved_by", ""), x.get("approved_on", "")
            continue
        if t == "accept_fee_entry":
            account = account or profile["accounts"].get("bank_fees", "")
            if not account:
                issue(x, "No bank fee account in the profile; give the account")
                continue
        elif t == "correct_cashbook_to_bank":
            if not account:
                issue(x, "Offset account is required")
                continue
            b, c = (by_id[i] for i in ids)
            item.side, item.category, item.recipe = "Ledger", "Amount correction", [(1, b.line_id), (-1, c.line_id)]
        elif t == "bank_error_to_correct":
            b, c = (by_id[i] for i in ids)
            item.side, item.category, item.recipe = "Bank", "Bank error", [(1, c.line_id), (-1, b.line_id)]
        elif t in ("reverse_duplicate", "second_payment_outstanding"):
            group = [by_id[i] for i in ids]
            bl = [l for l in group if l.source == "Bank"]
            cl = [l for l in group if l.source == "Ledger"]
            if len(bl) != 1 or len(cl) != 2 or len({l.amount for l in group}) != 1 or ref not in {l.line_id for l in cl}:
                issue(x, "Needs one bank line, two equal ledger lines, and a line_ref naming one of them")
                continue
            if t == "reverse_duplicate":
                if not account:
                    issue(x, "Offset account is required")
                    continue
                item.side, item.category, item.recipe = "Ledger", "Duplicate reversal", [(-1, ref)]
            else:
                kind = by_id[ref].kind
                item.side, item.category, item.recipe = "Bank", "Outstanding check" if kind == "check" else "Outstanding payment", [(1, ref)]
        elif t == "reverse_application_reopen_receivable":
            original, returned = by_id.get(ref), by_id[ids[0]]
            ok = (original is not None and original.source == "Ledger" and original.kind == "receipt" and original.line_id in matched_book
                  and original.amount == -returned.amount and original.date <= returned.date and ref not in linked)
            if not ok:
                issue(x, "line_ref must be an earlier matched receipt of the same amount, not already reversed")
                continue
            linked.add(ref)
            item.side, item.category, item.recipe = "Ledger", "Returned item", [(1, returned.line_id)]
            account = account or f"Accounts receivable - {original.description} (reopen {original.reference})"
            item.note = f"Reopens {original.reference} by reversing {original.line_id}; no new invoice is created"
        elif t == "book_bank_item":
            if not account:
                issue(x, "Offset account is required")
                continue
            item.side, item.category, item.recipe = "Ledger", "Other bank item", [(1, ids[0])]
        elif t == "accept_outstanding_payment":
            item.side, item.category, item.recipe = "Bank", "Outstanding payment", [(1, ids[0])]
        elif t == "posting_of_carried":
            target = carried_pending.get(ref)
            line = by_id[ids[0]]
            if target is None or target.status != "Approved (carried)" or target.fixed_effect != line.amount or ref in linked:
                issue(x, "line_ref must name a carried adjustment that is still unposted and has exactly this amount")
                continue
            linked.add(ref)
            item.category, item.possible_original = "Posting of carried adjustment", ref
            item.note = f"Posts {ref} ({target.posting_ref}); confirmed by the user"
        item.status, item.treatment, item.account = "Approved", t, account
        item.approved_by, item.approved_on = x.get("approved_by", ""), x.get("approved_on", "")
        if x.get("note") and t not in ("reverse_application_reopen_receivable", "posting_of_carried"):
            item.note = x["note"]
        x["result"] = "Applied"
    return issues


def effect_of(item, by_id):
    if item.fixed_effect is not None:
        return item.fixed_effect
    return sum((s * by_id[i].amount for s, i in item.recipe), ZERO)


# ---------------------------------------------------------------- checks and status

def integrity_checks(r):
    """Script checks. Each is (label, expected, actual, passed)."""
    out = []
    lines = [l for f in (r["bank_file"], r["ledger_file"]) for l in f.rows if l.included]
    uses = defaultdict(int)
    for b, c, *_ in r["matches"]:
        uses[b.line_id] += 1
        uses[c.line_id] += 1
    for i in r["items"]:
        for x in i.lines:
            uses[x] += 1
    bad = [l.line_id for l in lines if uses[l.line_id] != 1]
    out.append(("Included source rows accounted for exactly once", "0 exceptions", f"{len(bad)} exceptions" + (f": {', '.join(bad[:10])}" if bad else ""), not bad))
    excluded_used = [l.line_id for f in (r["bank_file"], r["ledger_file"]) for l in f.rows if not l.included and uses[l.line_id]]
    out.append(("Excluded rows not used", "0", str(len(excluded_used)), not excluded_used))
    keys = [(l.file, l.sheet, l.row) for f in (r["bank_file"], r["ledger_file"]) for l in f.rows]
    out.append(("Each source row (file, sheet, row) appears once", str(len(keys)), str(len(set(keys))), len(keys) == len(set(keys))))
    for f in (r["bank_file"], r["ledger_file"]):
        inc = [l for l in f.rows if l.included]
        total = sum((l.amount for l in inc if l.amount is not None), ZERO)
        by_disp = sum((l.amount for l in inc if l.amount is not None and (l.disposition == "Matched" or l.disposition.startswith("Review item"))), ZERO)
        out.append((f"{f.source} rows: included amounts equal amounts by disposition", money(total), money(by_disp), total == by_disp))
        for label, exp, act in f.checks:
            out.append((label, exp, act, Decimal(exp) == Decimal(act)))
    pairs_bad = [b.line_id for b, c, *_ in r["matches"] if b.amount != c.amount]
    out.append(("Matched pairs have equal amounts", "0", str(len(pairs_bad)), not pairs_bad))
    cont = tables.balance_continuity(r["bank_file"].rows)
    if cont:
        out.append(("Bank running balance is continuous", "no breaks", ", ".join(map(str, cont["breaks_at_rows"])) or "no breaks", not cont["breaks_at_rows"]))
        f = r["figures"]
        if f.get("statement_opening") is not None and not cont["breaks_at_rows"]:
            out.append(("Running balance implies the stated opening balance", money(f["statement_opening"]), money(cont["implied_opening"]),
                        cont["implied_opening"] == f["statement_opening"]))
        if f.get("statement_closing") is not None and not cont["breaks_at_rows"]:
            out.append(("Running balance implies the stated closing balance", money(f["statement_closing"]), money(cont["implied_closing"]),
                        cont["implied_closing"] == f["statement_closing"]))
    return out


def status_of(r):
    f, p = r["figures"], r["period"]
    missing = []
    if f.get("statement_closing") is None:
        missing.append("the bank statement closing balance")
    if f.get("statement_opening") is None:
        missing.append("the bank statement opening balance")
    if f.get("ledger_opening") is None:
        missing.append("the ledger opening balance")
    if not r["prior"] and not p.get("opening_items") and f.get("opening_difference") not in (None, ZERO):
        missing.append(f"the outstanding items from the previous reconciliation (opening balances differ by {money(f['opening_difference'])})")
    if missing:
        return "Incomplete: missing " + "; ".join(missing), missing
    fails = []
    if f.get("statement_difference"):
        fails.append(f"bank activity does not agree to the stated closing balance (difference {money(f['statement_difference'])})")
    if f.get("ledger_difference"):
        fails.append(f"ledger activity does not agree to the stated ledger closing balance (difference {money(f['ledger_difference'])})")
    for k, text in (("opening_difference", "opening balances and carried items do not agree"),
                    ("bank_opening_vs_prior", "bank opening balance differs from the previous stated closing balance"),
                    ("ledger_opening_vs_prior", "ledger opening balance differs from the previous actual ledger closing balance")):
        if f.get(k):
            fails.append(f"{text} (difference {money(f[k])})")
    failed = [c[0] for c in r["checks"] if not c[3]]
    if failed:
        fails.append("checks failed: " + "; ".join(failed))
    if fails:
        return "Not reconciled: " + "; ".join(fails), []
    if f["final_difference"] != 0:
        return f"Not reconciled: difference of {money(f['final_difference'])} remains", []
    open_n = sum(1 for i in r["items"] if i.status in ("Open", "Rejected"))
    if open_n or r["issues"]:
        extra = f" and {len(r['issues'])} decision(s) could not be applied" if r["issues"] else ""
        return f"Not reconciled: zero difference, but {open_n} item(s) need decisions{extra}", []
    pend = [e for e in r["entries"]]
    if pend:
        net = sum((e["amount"] if e["debit"] == r["profile"]["accounts"]["cash"] else -e["amount"] for e in pend), ZERO)
        return f"Reconciled subject to posting {len(pend)} approved adjustment(s) (net {money(net)})", []
    return "Reconciled, awaiting sign-off", []
