"""Build the review workbook (0.5.1, format review-2): the accountant reviews and edits it in Excel, then uploads it back.

Visible: Reconciliation, Journal entries, Bank raw, Ledger raw, Guide, Accounts. Hidden (ordinary hiding): Setup and
Checks, which keep the setup, the controls, and every record needed to validate an upload and start the next period.

Every figure the workbook shows is a live formula; the script works out what each formula should show, stores those
expectations on the Checks tab, and `check` compares them after recalculation. Nothing here decides anything: an
item is approved only when the accountant marks it Approved and Claude validates the upload.
"""
from __future__ import annotations

import hashlib, json, re
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

from bank_rules import CARRY_COLS, KIND_OF_CATEGORY, MEANING, carry_digest, effect_of, money
from foundation import decisions as dec
from foundation import profiles
from foundation.sections import write_section
from foundation.tables import parse_amount, parse_date, text
import review_state as S

ZERO = Decimal("0.00")
GREEN, PINK, READY, EDIT, FLOW = "E2EFDA", "FCE4E4", "FFE699", "FFF9DB", "F2F2F2"
MUTED = "808080"
fill = lambda c: PatternFill("solid", start_color=c, end_color=c)
hair = Side(style="thin", color="D9D9D9")
thin = Side(style="thin", color="808080")
dark = Side(style="thin", color="000000")
gold = Side(style="thin", color="BF9000")
MONEY = '#,##0.00;(#,##0.00);0.00'
DATE = "mm/dd/yyyy"
TIMING = "Timing item—no JE"
BANK_ERROR_ROW = "Bank error—no JE"
PENDING = "Pending"
NO_TREATMENT = "No treatment"
LEDGER_FIX = "Approve ledger correction"
BANK_ERROR = "Confirm bank error—no JE"
# 0.5.0 (review-1) labels for the same two treatments, read only when migrating a 0.5.0 upload
OLD_LABELS = {"Approved: correct ledger": LEDGER_FIX, "Approved: bank error": BANK_ERROR}
PLACEHOLDER = "[choose account]"
JEQ = f"'{S.JE}'"
CKQ = S.CHECKS
# the rules' own item notes, which say why an item was raised, not why it was approved
GENERIC_NOTES = ("Proposal only", "No matching", "Same reference, different amount", "More than one possible partner",
                 "Bank returned a payment")


def f(v):
    return None if v is None else float(v)


def is_approved(label):
    """A choice records the decision itself: any listed choice except Pending (one decision per item)."""
    return bool(str(label or "").strip()) and str(label).strip() not in (PENDING, NO_TREATMENT)


def xl_approved(cell):
    """The same test as an Excel formula fragment."""
    return f'AND({cell}<>"",{cell}<>"{PENDING}",{cell}<>"{NO_TREATMENT}")'


def slug(s):
    return re.sub(r"[^a-z0-9]+", "-", s.lower()).strip("-") or "account"


def out_name(r):
    return f"bankrec-{slug(r['profile']['account']['nickname'])}-{r['label']}.xlsx"


def month_label(label):
    return datetime.strptime(label, "%Y-%m").strftime("%b %Y")


# ------------------------------------------------------------------ refs

def match_key(b, c):
    return dec.fingerprint("Matched", [b.content(), c.content()])


def assign_refs(r, existing=None):
    """Stable Refs: M-01 for matched pairs, R-01 for every other item. Existing Refs (from an uploaded workbook)
    are kept by content; anything new gets the next number. Sorting rows or changing decisions never renumbers."""
    existing = dict(existing or {})
    used = set(existing.values())
    by_id = r["by_id"]
    keys = []
    for b, c, *_ in sorted(r["matches"], key=lambda m: (m[0].date, m[0].line_id)):
        keys.append(("M", match_key(b, c), (b, c)))
    items = [i for i in r["items"]] + [p for p in r["pending"] if posted_by_user(p)]

    def when(i):
        ds = [by_id[x].date for x in i.lines if by_id[x].date] if i.lines else []
        return (min(ds) if ds else r["period"]["period_end"], i.id)

    for i in sorted(items, key=when):
        keys.append(("R", i.fingerprint or i.id, i))
    refs, out_m, out_i = {}, {}, {}
    nxt = {p: max([int(v[2:]) for v in used if v.startswith(p + "-")] + [0]) for p in ("M", "R")}
    for p, k, obj in keys:
        ref = existing.get(k)
        if ref is None or ref in refs.values():
            nxt[p] += 1
            ref = f"{p}-{nxt[p]:02d}"
        refs[k] = ref
        if p == "M":
            out_m[(obj[0].line_id, obj[1].line_id)] = ref
        else:
            out_i[obj.id] = ref
    return refs, out_m, out_i


# ------------------------------------------------------------------ choices for each item

def opt(label, treatment, bank=ZERO, ledger=ZERO, je=False, line_ref=""):
    return dict(label=label, treatment=treatment, bank=bank, ledger=ledger, je=je, line_ref=line_ref)


def carried_label(p):
    """'CF-2026-03-R-05' -> '2026-03 R-05'."""
    s = p.id[3:] if p.id.startswith("CF-") else p.id
    m = re.match(r"(\d{4}-\d{2})-(.*)$", s)
    return f"{m.group(1)} {m.group(2)}" if m else s


def options_for(i, r):
    """The approval choices the workbook offers for an item, each with its effect on the bank and ledger sides.

    Every choice for an item moves the remaining difference by the same amount; they differ in which side moves
    and whether a journal entry is needed."""
    by_id = r["by_id"]
    lines = [by_id[x] for x in i.lines]
    bank = [l for l in lines if l.source == "Bank"]
    book = [l for l in lines if l.source != "Bank"]
    t = i.type
    if t == "Proposed outstanding check":
        return [opt("Approved", "accept_outstanding", bank=book[0].amount)]
    if t == "Proposed deposit in transit":
        return [opt("Approved", "accept_in_transit", bank=book[0].amount)]
    if t in ("Unmatched ledger payment", "Unmatched ledger receipt"):
        out = [opt("Approved: outstanding payment", "accept_outstanding_payment", bank=book[0].amount)] if t == "Unmatched ledger payment" else []
        for p in r["pending"]:
            still_open = p.status == "Approved (carried)"
            posted_here = p.status == "Posted in this period" and "confirmed by the user" in (p.note or "") and i.possible_original == p.id
            if (still_open or posted_here) and p.fixed_effect == book[0].amount:
                out.append(opt(f"Approved: posting of {carried_label(p)}", "posting_of_carried", ledger=-p.fixed_effect, line_ref=p.id))
        return out
    if t == "Proposed bank fee entry":
        return [opt("Approved", "accept_fee_entry", ledger=bank[0].amount, je=True)]
    if t == "Unidentified bank item":
        return [opt("Approved", "book_bank_item", ledger=bank[0].amount, je=True)]
    if t == "Returned item":
        originals = [x for x in (i.possible_original or "").split(", ") if x]
        if len(originals) == 1:
            return [opt("Approved", "reverse_application_reopen_receivable", ledger=bank[0].amount, je=True, line_ref=originals[0])]
        return []
    if t == "Amount mismatch" and len(bank) == 1 and len(book) == 1:
        b, c = bank[0], book[0]
        return [opt(LEDGER_FIX, "correct_cashbook_to_bank", ledger=b.amount - c.amount, je=True),
                opt(BANK_ERROR, "bank_error_to_correct", bank=c.amount - b.amount)]
    if t == "Multiple candidates" and len(bank) == 1 and len(book) == 2 and len({l.amount for l in lines}) == 1:
        later = max(book, key=lambda l: (l.date, l.row, l.line_id))
        return [opt("Approved: reverse duplicate", "reverse_duplicate", ledger=-later.amount, je=True, line_ref=later.line_id),
                opt("Approved: second payment", "second_payment_outstanding", bank=later.amount, line_ref=later.line_id)]
    return []


def posted_by_user(p):
    return p.status == "Posted in this period" and "confirmed by the user" in (p.note or "")


def fixed_option(i, r):
    """Items already settled (carried from last period, or a posting found by its JE number): one fixed choice."""
    if i.type == "Posting of carried adjustment" and i.status == "Approved" and i.rule == "P1":
        return opt("Matched", "posting reference found")
    if posted_by_user(i):  # its posting was confirmed on the Reconciliation tab; that line's choice reverses this effect
        when = month_label(i.origin) if re.match(r"^\d{4}-\d{2}$", i.origin or "") else (i.origin or "earlier")
        return opt(f"Approved ({when})", "carried", ledger=i.fixed_effect)
    if i.status == "Approved (carried)":
        eff = i.fixed_effect
        when = month_label(i.origin) if re.match(r"^\d{4}-\d{2}$", i.origin or "") else (i.origin or "earlier")
        if i.side == "Bank":
            return opt(f"Approved ({when})", "carried", bank=eff)
        return opt(f"Approved ({when})", "carried", ledger=eff, je=True)
    return None


def short(l):
    return f"{l.date:%m/%d} {l.description}" if l.date else l.description


def proposal_for(i, r, chosen_account, basis):
    by_id = r["by_id"]
    lines = [by_id[x] for x in i.lines]
    t = i.type
    if t == "Proposed bank fee entry":
        return "Record bank fee" + (f" ({basis})" if basis else "")
    if t == "Unidentified bank item":
        if chosen_account and chosen_account != PLACEHOLDER:
            return f"Record: {chosen_account} suggested" + (f" ({basis})" if basis else "")
        return "Who sent it, or what was it for? Choose the account" if lines[0].amount > 0 else "What was it for? Choose the account"
    if t == "Amount mismatch":
        b = next(l for l in lines if l.source == "Bank")
        c = next(l for l in lines if l.source != "Bank")
        return (f"Ledger {money(abs(c.amount))}, bank {money(abs(b.amount))}. Correct the ledger by {money(b.amount - c.amount)}, "
                "or confirm a bank error (no JE)")
    if t == "Multiple candidates":
        book = sorted((l for l in lines if l.source != "Bank"), key=lambda l: (l.date, l.row, l.line_id))
        later = book[-1]
        twin = len(book) == 2 and (book[0].date, book[0].description) == (later.date, later.description)
        return (f"Possible duplicate: reverse the second {short(later)} entry, or choose second payment" if twin else
                f"Possible duplicate: reverse {short(later)}, or choose second payment")
    if t == "Returned item":
        orig = by_id.get((i.possible_original or "").split(", ")[0])
        return (f"Returned payment: reverse {short(orig)}{' (' + orig.reference + ')' if orig and orig.reference else ''}, reopen the receivable"
                if orig else "Returned payment")
    if i.status == "Approved (carried)":
        return f"Approved in {i.origin}; not yet seen in the ledger" + (f". {i.note}" if i.note else "")
    return ""


def default_account(i, r, accounts, coding_by_item):
    """A suggested account for the non-cash line, with its basis. Suggestions only."""
    prof = r["profile"]
    t = i.type
    hist = (r["prior"] or {}).get("history", {}) if r.get("prior") else {}
    line = r["by_id"][i.lines[0]] if i.lines else None
    h = hist.get(line.description.upper()) if line is not None else None
    if t == "Proposed bank fee entry":
        a = prof["accounts"].get("bank_fees") or next((x for x in accounts if "fee" in x.lower()), "")
        return (a or PLACEHOLDER), ("as in " + r["prior"]["period_label"] if h else "")
    if t == "Unidentified bank item":
        s = coding_by_item.get(i.id)
        if s and s.get("suggested_account"):
            return s["suggested_account"], "coded this way before"
        if h and h[2]:
            return h[2], "as in " + r["prior"]["period_label"]
        return PLACEHOLDER, ""
    if t == "Returned item":
        ar = [x for x in accounts if "receivable" in x.lower()]
        if len(ar) == 1:
            return ar[0], ""
        orig = r["by_id"].get((i.possible_original or "").split(", ")[0])
        return (f"Accounts receivable - {orig.description}" if orig else PLACEHOLDER), ""
    return PLACEHOLDER, ""


def proposed_lines(i, o, r, account, description):
    """The proposed entry: the cash line (fixed, from the bank or ledger lines) and one line for the user's account."""
    cash = r["profile"]["accounts"]["cash"]
    e = o["ledger"]
    title = entry_title(i, r)
    cash_line = dict(account=cash, description=title, debit=e if e > 0 else None, credit=-e if e < 0 else None, fixed=True)
    other = dict(account=account, description=description or title, debit=-e if e < 0 else None, credit=e if e > 0 else None, fixed=False)
    return [cash_line, other] if e > 0 else [other, cash_line]


def entry_title(i, r):
    by_id = r["by_id"]
    lines = [by_id[x] for x in i.lines]
    bank = [l for l in lines if l.source == "Bank"]
    if i.type == "Multiple candidates":
        later = max((l for l in lines if l.source != "Bank"), key=lambda l: (l.date, l.row, l.line_id))
        return f"Duplicate {later.description} {later.reference}".strip()
    if i.type == "Amount mismatch":
        c = next(l for l in lines if l.source != "Bank")
        return f"Correction: {c.description} {c.reference}".strip()
    if i.status == "Approved (carried)":
        return i.description or i.category
    return (bank[0] if bank else lines[0]).description if lines else (i.description or i.type)


def entry_date(i, r):
    by_id = r["by_id"]
    ds = [by_id[x].date for x in i.lines if by_id[x].date]
    if ds:
        return min(ds)
    try:
        return datetime.strptime(str(i.line_date)[:10], "%Y-%m-%d").date()
    except ValueError:
        return r["period"]["period_end"]


# ------------------------------------------------------------------ model

def build_model(r, ctx):
    """Everything the workbook shows, from the run and the review state. ctx keys (all optional):
    refs (key -> Ref), approvals (Ref -> label), entries (Ref -> {lines, je_number, comment}), reconciled_by, accounts,
    accounts_source, edits (list), recorded (list), validated_on, completion, coding, proposed (Ref -> stored proposal),
    files, created_on, version."""
    by_id = r["by_id"]
    refs, mref, iref = assign_refs(r, ctx.get("refs"))
    accounts = list(ctx.get("accounts") or [])
    coding_by_item = {s["item_id"]: s for s in ((ctx.get("coding") or {}).get("lines") or [])}
    approvals = dict(ctx.get("approvals") or {})
    user_entries = dict(ctx.get("entries") or {})
    stored = dict(ctx.get("proposed") or {})
    rows, entries, effects, options_rec, linemap, refs_rows, proposed_rows = [], [], [], [], [], [], []

    for (b_id, c_id), ref in sorted(mref.items(), key=lambda kv: kv[1]):
        b, c = by_id[b_id], by_id[c_id]
        rows.append(dict(ref=ref, bank=b, ledger=c, approval=("fixed", "Matched"), jeref=("text", "")))
        linemap += [(ref, "Bank", b.line_id), (ref, "Ledger" if c.source != "Carried" else "Carried", c.line_id)]
        refs_rows.append((ref, match_key(b, c), f"{b.line_id}+{c.line_id}", "Matched pair", "none"))
    pending_by_id = {p.id: p for p in r["pending"]}
    shown = list(r["items"]) + [p for p in r["pending"] if posted_by_user(p)]
    for i in sorted(shown, key=lambda i: iref[i.id]):
        ref = iref[i.id]
        lines = [by_id[x] for x in i.lines]
        bank = [l for l in lines if l.source == "Bank"]
        book = [l for l in lines if l.source != "Bank"]
        fx = fixed_option(i, r)
        opts = [fx] if fx else options_for(i, r)
        where = "JE" if any(o["je"] for o in opts) else ("Rec" if opts and not fx else ("fixed" if fx else "none"))
        refs_rows.append((ref, i.fingerprint or i.id, i.id, i.type, where))
        for o in opts:
            options_rec.append((ref, o["label"], o["treatment"], f(o["bank"]), f(o["ledger"]), "Yes" if o["je"] else "No", o["line_ref"]))
        effects.append(dict(ref=ref, opts=opts))
        for l in lines:
            linemap.append((ref, "Bank" if l.source == "Bank" else "Ledger", l.line_id))
        # reconciliation rows: one per bank/ledger line pair; carried items show what they carry
        if where == "JE":
            approval, jeref = ("flow", None), ("flow", None)
        elif where == "Rec":
            approval = ("edit", approvals.get(ref, PENDING), [o["label"] for o in opts] + [PENDING])
            jeref = ("text", TIMING if any(o["treatment"] != "posting_of_carried" for o in opts) else "Posting of a carried entry—no JE")
        elif where == "fixed":
            if fx["treatment"] == "posting reference found":
                p = pending_by_id.get(i.possible_original)
                approval = ("fixed", "Matched")
                jeref = ("text", f"Posts {carried_label(p) if p else i.possible_original}" + (f" ({p.posting_ref})" if p and p.posting_ref else ""))
            elif posted_by_user(i):
                approval = ("fixed", fx["label"])
                jeref = ("text", "Posted: " + (i.note.split("Posted by ")[-1].split(" (")[0] if "Posted by " in (i.note or "") else "confirmed"))
            else:
                approval = ("fixed", fx["label"]) if i.side == "Bank" else ("flow", None)
                jeref = ("text", BANK_ERROR_ROW if i.category == "Bank error" else TIMING) if i.side == "Bank" else ("flow", None)
        else:
            approval = ("fixed", NO_TREATMENT)
            reason = ("Correct the export and rerun" if i.type == "Incomplete source row" else
                      "Name the original receipt: tell Claude" if i.type == "Returned item" else
                      "No treatment here: check the ledger or export")
            jeref = ("text", reason)
        if lines:
            for k in range(max(len(bank), len(book))):
                rows.append(dict(ref=ref, bank=bank[k] if k < len(bank) else None, ledger=book[k] if k < len(book) else None,
                                 approval=approval, jeref=jeref))
        else:
            origin = i.origin or ""
            what = "entry not yet seen in the ledger" if i.side == "Ledger" else "still outstanding"
            what = "posted this period" if posted_by_user(i) else what
            rows.append(dict(ref=ref, bank=None, ledger=None, carried=dict(
                date=entry_date(i, r), description=f"Carried from {origin}: {i.category or i.description} ({what})",
                amount=i.fixed_effect), approval=approval, jeref=jeref))
        # journal entry
        if where == "JE" or (where == "fixed" and fx["je"]):
            o = next(o for o in opts if o["je"])
            st = stored.get(ref)
            if st:
                proposal, plines, title, edate = st["proposal"], st["lines"], st["title"], st["date"]
            else:
                if i.status == "Approved (carried)":
                    cash = r["profile"]["accounts"]["cash"]
                    e = i.fixed_effect
                    plines = [dict(account=i.account or PLACEHOLDER, description=i.description or i.category,
                                   debit=-e if e < 0 else None, credit=e if e > 0 else None, fixed=True),
                              dict(account=cash, description=i.description or i.category, debit=e if e > 0 else None,
                                   credit=-e if e < 0 else None, fixed=True)]
                    if e > 0:
                        plines.reverse()
                    acct, basis = i.account, ""
                else:
                    acct, basis = default_account(i, r, accounts, coding_by_item)
                    s = coding_by_item.get(i.id)
                    desc = s.get("expected_from_history") if s and s.get("expected_from_history") else None
                    plines = proposed_lines(i, o, r, acct, desc)
                proposal = proposal_for(i, r, acct, basis)
                title, edate = entry_title(i, r), entry_date(i, r)
            ue = user_entries.get(ref) or {}
            cur_lines = ue.get("lines") or [dict(x) for x in plines]
            je_number = ue.get("je_number") if ue.get("je_number") is not None else (i.posting_ref if i.status == "Approved (carried)" else "")
            entries.append(dict(ref=ref, date=edate, title=title, proposal=proposal, lines=cur_lines, comment=ue.get("comment") or "",
                                je_number=je_number or "", fixed=bool(fx), approval=(fx["label"] if fx else approvals.get(ref, PENDING)),
                                labels=[x["label"] for x in opts] + ([] if fx else [PENDING])))
            for n, pl in enumerate(plines, 1):
                proposed_rows.append((ref, n, pl["account"], pl["description"], f(pl["debit"]), f(pl["credit"]),
                                      "Yes" if pl["fixed"] else "No", proposal, title, edate))
    return dict(r=r, ctx=ctx, refs=refs, mref=mref, iref=iref, rows=rows, entries=entries, effects=effects,
                options=options_rec, linemap=linemap, refs_rows=refs_rows, proposed=proposed_rows, accounts=accounts)


# ------------------------------------------------------------------ what each formula should show

def current_choice(m, ref):
    for e in m["entries"]:
        if e["ref"] == ref:
            return e["approval"] if e["approval"] else PENDING
    for row in m["rows"]:
        if row["ref"] == ref:
            a = row["approval"]
            if a[0] == "edit":
                return a[1]
            if a[0] == "fixed":
                return a[1]
    return PENDING


def entry_check(m, e):
    accts = {a.casefold() for a in m["accounts"]}
    lines = e["lines"]
    approved = is_approved(e["approval"])
    eff = next(x for x in m["effects"] if x["ref"] == e["ref"])
    chosen = next((o for o in eff["opts"] if o["label"] == e["approval"]), None)
    needs = chosen["je"] if chosen else True
    if e["je_number"] and not approved:
        return "Not approved"
    if approved and not needs:
        return "Remove JE number" if e["je_number"] else "Not needed"
    missing = sum(1 for l in lines if not (l.get("account") or "").strip() and (l.get("debit") is not None or l.get("credit") is not None))
    placeholder = sum(1 for l in lines if str(l.get("account") or "").startswith("["))
    if missing + placeholder:
        return "Needs an account"
    dr = sum((Decimal(str(l["debit"])) for l in lines if l.get("debit") not in (None, "")), ZERO)
    cr = sum((Decimal(str(l["credit"])) for l in lines if l.get("credit") not in (None, "")), ZERO)
    if dr.quantize(Decimal("0.01")) != cr.quantize(Decimal("0.01")) or dr == 0:
        return "Not balanced"
    unknown = sum(1 for l in lines if (l.get("account") or "").strip() and not str(l["account"]).startswith("[")
                  and str(l["account"]).strip().casefold() not in accts)
    return "New account: confirm" if unknown else "Balanced"


def figures(m):
    """Adjustments and remaining difference as the workbook's formulas compute them."""
    r = m["r"]
    fig = r["figures"]
    bank_adj = ledger_adj = pending_eff = ZERO
    chosen = {}
    for e in m["effects"]:
        c = current_choice(m, e["ref"])
        o = next((o for o in e["opts"] if o["label"] == c), None)
        chosen[e["ref"]] = o
        if o:
            bank_adj += o["bank"]
            ledger_adj += o["ledger"]
        elif e["opts"]:
            pending_eff += e["opts"][0]["bank"] - e["opts"][0]["ledger"]
    out = dict(bank_adj=bank_adj, ledger_adj=ledger_adj, chosen=chosen)
    if fig.get("bank_closing") is not None and fig.get("ledger_closing") is not None:
        out["adjusted_bank"] = fig["bank_closing"] + bank_adj
        out["adjusted_ledger"] = fig["ledger_closing"] + ledger_adj
        out["remaining"] = out["adjusted_bank"] - out["adjusted_ledger"]
        out["unexplained"] = out["remaining"] + pending_eff
    return out


LIVE_LABELS = ["Bank closing differs from the statement", "Bank lines not used exactly once", "Ledger lines not used exactly once",
               "Matched items whose amounts differ", "Difference not explained by pending items",
               "A journal entry needs fixing (see its Check)"]


def live_values(m):
    r = m["r"]
    fig = r["figures"]
    fx = figures(m)
    checks = {e["ref"]: entry_check(m, e) for e in m["entries"]}
    fixing = sum(1 for e in m["entries"] if is_approved(e["approval"]) and checks[e["ref"]] in ("Needs an account", "Not balanced")) + \
        sum(1 for v in checks.values() if v in ("Not approved", "Remove JE number"))
    stmt = fig.get("statement_difference") or ZERO
    return [stmt, 0, 0, 0, fx.get("unexplained", ZERO), fixing], checks


def checks_line(m):
    vals, _ = live_values(m)
    for label, v in zip(LIVE_LABELS, vals):
        if Decimal(str(v)).quantize(Decimal("0.01")) != 0:
            return "Check failed: " + (label if not label.startswith("Bank closing") else
                                      f"Bank closing differs from the statement ({money(m['r']['figures']['statement_closing'])})")
    return "All pass. Your edits are validated on upload."


# ------------------------------------------------------------------ rendering helpers

def heading(ws, row, col, labels, right=()):
    for i, t in enumerate(labels):
        c = ws.cell(row, col + i, t)
        c.font = Font(bold=True)
        c.border = Border(bottom=thin)
        c.alignment = Alignment(horizontal="right" if t in right else "left", vertical="bottom", wrap_text=True)


def editable(c):
    c.fill = fill(EDIT)
    c.border = Border(left=gold, right=gold, top=gold, bottom=gold)


def flow(c):
    c.fill = fill(FLOW)
    c.border = Border(bottom=hair)
    c.font = Font(color="404040")


def muted(c):
    c.font = Font(color=MUTED, italic=True, size=10)


def outline(ws, r1, c1, r2, c2, side=dark):
    for rr in range(r1, r2 + 1):
        for cc in range(c1, c2 + 1):
            cell = ws.cell(rr, cc)
            b = cell.border
            cell.border = Border(left=side if cc == c1 else b.left, right=side if cc == c2 else b.right,
                                 top=side if rr == r1 else b.top, bottom=side if rr == r2 else b.bottom)


def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w


def num(v):
    return None if v in (None, "") else float(Decimal(str(v)))


# ------------------------------------------------------------------ source tabs

def roles_of(cfg):
    return {v.casefold(): k for k, v in (cfg.get("columns") or {}).items() if v}


def display_grid(grid, cfg, header_row):
    """The export as received, with dates and amounts shown as dates and numbers where they read cleanly."""
    roles = roles_of(cfg)
    head = [text(c) for c in grid[header_row - 1]] if grid and header_row and header_row <= len(grid) else []
    role_at = {j: roles.get(h.casefold()) for j, h in enumerate(head) if h}
    out = []
    for i, row in enumerate(grid):
        new = []
        for j, v in enumerate(row):
            role = role_at.get(j) if i >= header_row else None
            if isinstance(v, str) and role in ("amount", "debit", "credit", "balance"):
                a, prob = parse_amount(v)
                new.append(float(a) if a is not None and not prob else v)
            elif isinstance(v, str) and role == "date":
                d = parse_date(v, cfg.get("date_format"))
                new.append(datetime(d.year, d.month, d.day) if d else v)
            else:
                new.append(v)
        out.append(new)
    return out


def canon_cell(v):
    """Comparable form of a source cell, so a CSV value and the number Excel shows compare equal."""
    if v is None or (isinstance(v, str) and not v.strip()):
        return ""
    if isinstance(v, datetime):
        return v.date().isoformat() if not (v.hour or v.minute or v.second) else v.isoformat()
    if isinstance(v, date):
        return v.isoformat()
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, (int, float, Decimal)):
        return str(Decimal(repr(v) if isinstance(v, float) else v).quantize(Decimal("0.01")))
    return str(v).strip()


def snapshot_digest(grids):
    def trim(row):
        vals = [canon_cell(c) for c in row]
        while vals and vals[-1] == "":
            vals.pop()
        return vals
    blob = json.dumps({k: [trim(row) for row in g] for k, g in sorted(grids.items())})
    return hashlib.sha256(blob.encode()).hexdigest()[:16]


def write_source_tab(ws, name, grid, cfg, header_row):
    ws["A1"] = f"As uploaded: {name}. Kept as received."
    muted(ws["A1"])
    shown = display_grid(grid, cfg, header_row)
    start = 3
    width = {}
    for i, row in enumerate(shown):
        for j, v in enumerate(row):
            if v is None:
                continue
            c = ws.cell(start + i, j + 1, v)
            if i == header_row - 1:
                c.font = Font(bold=True)
                c.border = Border(bottom=thin)
            if isinstance(v, datetime):
                c.number_format = DATE
            elif isinstance(v, float):
                c.number_format = MONEY
            width[j] = max(width.get(j, 8), min(40, len(text(v)) + 2))
    for j, w in width.items():
        ws.column_dimensions[ws.cell(1, j + 1).column_letter].width = w
    ws.freeze_panes = ws.cell(start + header_row, 1).coordinate if header_row else "A4"


# ------------------------------------------------------------------ the workbook

def render(m, path):
    r, ctx = m["r"], m["ctx"]
    prof, fig, period = r["profile"], r["figures"], r["period"]
    acct = prof["account"]
    wb = Workbook()
    wb.properties.creator = ""
    wb.properties.lastModifiedBy = ""
    wb.properties.title = "Bank reconciliation"
    wb.properties.created = wb.properties.modified = datetime(2026, 1, 1)
    rec = wb.active
    rec.title = S.REC
    je = wb.create_sheet(S.JE)
    bank_ws = wb.create_sheet(S.BANK)
    led_ws = wb.create_sheet(S.LEDGER)
    guide = wb.create_sheet(S.GUIDE)
    accts = wb.create_sheet(S.ACCOUNTS)
    setup = wb.create_sheet(S.SETUP)
    ck = wb.create_sheet(S.CHECKS)
    for ws in (guide, accts, setup, ck):
        ws.sheet_properties.tabColor = "BFBFBF"
    files = ctx["files"]

    # ---- source tabs
    for ws, side in ((bank_ws, "bank"), (led_ws, "ledger")):
        fi = files[side]
        write_source_tab(ws, fi["name"], fi["grid"], prof["sources"][side], fi["header_row"])

    # ---- Checks first, so other tabs can refer to its rows
    ck.sheet_view.showGridLines = False
    widths(ck, {"A": 2, "B": 46, "C": 16, "D": 14, "E": 16, "F": 14, "G": 14, "H": 14, "I": 16, "J": 14, "K": 14, "L": 14,
                "M": 16, "N": 14, "O": 14, "P": 14, "Q": 14, "R": 14, "S": 14})
    ck["B1"] = "Control checks and records (hidden supporting tab). Failures show on Reconciliation under Checks."
    muted(ck["B1"])

    # sections whose position the formulas need: effects (per Ref), included lines, line map
    eff_title_row = 40
    n_eff = len(m["effects"])
    eff_first = eff_title_row + 2
    eff_last = eff_first + max(n_eff, 1) - 1
    row = eff_last + 3
    lines_rows = []
    for l in sorted([x for x in r["bank_file"].rows + r["ledger_file"].rows if x.included], key=lambda x: (x.source, x.line_id)):
        lines_rows.append((l.source, l.line_id, datetime(l.date.year, l.date.month, l.date.day) if l.date else None, l.reference,
                           l.description, f(l.amount)))
    lines_title = row
    lines_first = row + 2
    lines_last = lines_first + max(len(lines_rows), 1) - 1
    row = write_section(ck, row, S.S_LINES, ["Side", "Line ID", "Date", "Reference", "Description", "Amount"], lines_rows,
                        number_formats={"Amount": MONEY, "Date": DATE})
    map_title = row
    map_first = row + 2
    map_last = map_first + max(len(m["linemap"]), 1) - 1
    row = write_section(ck, row, S.S_LINEMAP, ["Ref", "Side", "Line ID"], m["linemap"])

    E = lambda col: f"{CKQ}!${col}${eff_first}:${col}${eff_last}"
    L = lambda col: f"{CKQ}!${col}${lines_first}:${col}${lines_last}"
    MP = lambda col: f"{CKQ}!${col}${map_first}:${col}${map_last}"
    RR = lambda col: f"{S.REC}!${col}${S.FIRST_ROW}:${col}${S.LAST_REC_ROW}"

    # effects table: Ref | Current choice | 3 x (Choice, Bank, Ledger, JE) | Chosen bank | Chosen ledger | Needs JE | Pending effect
    eh = ["Ref", "Current choice"]
    for k in (1, 2, 3):
        eh += [f"Choice {k}", f"Bank {k}", f"Ledger {k}", f"JE {k}"]
    eh += ["Chosen bank", "Chosen ledger", "Needs JE", "Pending effect"]
    ck.cell(eff_title_row, 2, S.S_EFFECTS).font = Font(bold=True)
    for j, h in enumerate(eh):
        c = ck.cell(eff_title_row + 1, 2 + j, h)
        c.font = Font(bold=True)
        c.border = Border(bottom=thin)
    for n, e in enumerate(m["effects"]):
        rr = eff_first + n
        ck.cell(rr, 2, e["ref"])
        ck.cell(rr, 3, f'=IFERROR(INDEX({RR("I")},MATCH($B{rr},{RR("B")},0))&"","")')
        for k in range(3):
            o = e["opts"][k] if k < len(e["opts"]) else None
            base = 4 + k * 4
            if o:
                ck.cell(rr, base, o["label"])
                ck.cell(rr, base + 1, f(o["bank"])).number_format = MONEY
                ck.cell(rr, base + 2, f(o["ledger"])).number_format = MONEY
                ck.cell(rr, base + 3, 1 if o["je"] else 0)
        # columns: D..O hold the choices; P chosen bank, Q chosen ledger, R needs JE, S pending effect
        pick = lambda off: (f'IF($C{rr}=$D{rr},{chr(ord("D") + off)}{rr},IF($C{rr}=$H{rr},{chr(ord("H") + off)}{rr},'
                            f'IF($C{rr}=$L{rr},{chr(ord("L") + off)}{rr},0)))')
        ck.cell(rr, 16, f'=IF($C{rr}="",0,{pick(1)})').number_format = MONEY
        ck.cell(rr, 17, f'=IF($C{rr}="",0,{pick(2)})').number_format = MONEY
        ck.cell(rr, 18, f'=IF($C{rr}="",1,IF($C{rr}=$D{rr},N(G{rr}),IF($C{rr}=$H{rr},N(K{rr}),IF($C{rr}=$L{rr},N(O{rr}),1))))')
        ck.cell(rr, 19, f'=IF(OR($D{rr}="",AND($C{rr}<>"",OR($C{rr}=$D{rr},$C{rr}=$H{rr},$C{rr}=$L{rr}))),0,N(E{rr})-N(F{rr}))').number_format = MONEY

    # live checks (rows 4-9) and adjustments by kind
    heading(ck, 3, 2, ["Problem count (updates as you edit; 0 passes)", "Value", "Result"])
    stmt_close = fig.get("statement_closing")
    je_rng = lambda col: f"{JEQ}!${col}${S.JE_FIRST}:${col}$3000"
    live_formulas = [
        f"=ROUND({S.REC}!C{S.CLOSE_ROW}-{float(stmt_close) if stmt_close is not None else 0},2)",
        f'=SUMPRODUCT(({L("B")}="Bank")*(COUNTIF({MP("D")},{L("C")})<>1))',
        f'=SUMPRODUCT(({L("B")}="Ledger")*(COUNTIF({MP("D")},{L("C")})<>1))',
        f'=SUMPRODUCT(({RR("I")}="Matched")*({RR("G")}<>"")*({RR("H")}<>"")*({RR("G")}<>{RR("H")}))',
        f'=ROUND({S.REC}!C{S.REMAINING_ROW}+SUM({E("S")}),2)',
        f'=SUMPRODUCT(({je_rng("J")}<>"")*({je_rng("J")}<>"{PENDING}")*(({je_rng("L")}="Needs an account")+({je_rng("L")}="Not balanced")))'
        f'+COUNTIF({je_rng("L")},"Not approved")+COUNTIF({je_rng("L")},"Remove JE number")',
    ]
    labels = list(LIVE_LABELS)
    if stmt_close is not None:
        labels[0] = f"Bank closing differs from the statement ({money(stmt_close)})"
    for k, (label, fml) in enumerate(zip(labels, live_formulas)):
        rr = S.LIVE_FIRST + k
        ck.cell(rr, 2, label)
        ck.cell(rr, 3, fml).number_format = MONEY if k in (0, 4) else "0"
        ck.cell(rr, 4, f'=IF(ROUND(C{rr},2)=0,"PASS","FAIL")')
    ck.conditional_formatting.add(f"D{S.LIVE_FIRST}:D{S.LIVE_LAST}", FormulaRule(formula=[f'D{S.LIVE_FIRST}="FAIL"'], fill=fill(PINK)))
    heading(ck, 11, 2, ["Reconciling adjustments by kind", "Bank", "Ledger"])
    ck.cell(12, 2, "Timing items and bank errors (approved)")
    ck.cell(12, 3, f"=SUM({E('P')})").number_format = MONEY
    ck.cell(13, 2, "Entries approved (including any reported posted)")
    ck.cell(13, 4, f"=SUM({E('Q')})").number_format = MONEY

    val = ctx.get("validated_on")
    ck["B16"] = (f"Last validated by Claude: {val}" if val else
                 f"Last validated by Claude: {ctx.get('created_on')}, when this draft was made")
    ck["B16"].font = Font(bold=True)
    ck["B17"] = (f"Validation result: {ctx.get('completion')}" if val else "Your edits are validated when you upload the workbook back.")
    ck["B18"] = f"Source files: {files['bank']['name']}, {files['ledger']['name']} (fingerprints stored)"
    ck["B19"] = f"Setup version {prof.get('profile_version', 1)}. Accounts: {ctx.get('accounts_source') or 'not given'}."

    # records, after the line map
    row = write_section(ck, row, S.S_EDITS, ["Where and what", "What Claude did"],
                        [(e[0], e[1]) for e in (ctx.get("edits") or [])] or [("No upload validated yet", "")])
    row = write_section(ck, row, S.S_DECISIONS, ["Ref", "Decision", "Approved by", "Approved on", "Type", "Treatment", "Account",
                                                  "Bank description", "Fingerprint"],
                        [(d["Ref"], d["Decision"], d["Approved by"], d["Approved on"], d["Type"], d["Treatment"], d["Account"],
                          d["Bank description"], d["Fingerprint"]) for d in (ctx.get("recorded") or [])] or
                        [("None yet", "Approvals are recorded when you upload the reviewed workbook for validation.", "", "", "", "", "", "", "")])
    carry, cfacts = carry_forward(m) if str(ctx.get("completion") or "").startswith("Complete") else ([], {})
    if cfacts:
        row = write_section(ck, row, S.S_CARRY_FACTS, ["Field", "Value"], list(cfacts.items()))
        row = write_section(ck, row, S.S_CARRY, CARRY_COLS, [[c.get(k) for k in CARRY_COLS] for c in carry] or [])
    else:
        row = write_section(ck, row, S.S_CARRY, ["Kind"], [("Written when the reviewed workbook is validated as complete.",)])
    row = write_section(ck, row, S.S_REFS, ["Ref", "Key", "Item ID", "Type", "Where"], m["refs_rows"])
    row = write_section(ck, row, S.S_OPTIONS, ["Ref", "Label", "Treatment", "Bank effect", "Ledger effect", "Needs JE", "Line ref"],
                        m["options"])
    row = write_section(ck, row, S.S_PROPOSED, ["Ref", "Line", "Account", "Description", "Debit", "Credit", "Fixed", "Proposal",
                                                 "Title", "Date"],
                        [(p[0], p[1], p[2], p[3], p[4], p[5], p[6], p[7], p[8],
                          datetime(p[9].year, p[9].month, p[9].day) if isinstance(p[9], date) else p[9]) for p in m["proposed"]])
    prior = r.get("prior")
    if prior:
        row = write_section(ck, row, S.S_CARRIED_FACTS, ["Field", "Value"],
                            [(k, v) for k, v in prior["facts"].items()] +
                            [("Bank file SHA-256", next(iter(sorted(prior["files"])), None))])
        row = write_section(ck, row, S.S_CARRIED_ROWS, CARRY_COLS, [[x.get(k) for k in CARRY_COLS] for x in prior["rows"]])
        row = write_section(ck, row, S.S_CARRIED_LINES, ["Side", "Line ID", "Date", "Reference", "Description", "Amount"],
                            [(x.get("Side"), x.get("Line ID"), x.get("Date"), x.get("Reference"), x.get("Description"), x.get("Amount"))
                             for x in prior.get("raw_lines", [])] or
                            [("Bank" if k[0] == "Bank rows" else "Ledger", v, k[1], k[2], k[3], k[4]) for k, v in prior["lines"].items()])
        row = write_section(ck, row, S.S_CARRIED_HISTORY, ["Bank description", "Type", "Treatment", "Account"],
                            [(k, *v) for k, v in prior["history"].items()])
    for side in ("bank", "ledger"):
        grid = files[side]["grid"]
        ncol = max((len(x) for x in grid), default=1)
        row = write_section(ck, row, S.S_SNAP[side], ["Row"] + [f"C{j + 1}" for j in range(ncol)],
                            [[i + 1] + list(x) + [None] * (ncol - len(x)) for i, x in enumerate(grid)])
    expect_title = row
    # (the expectations are written after the other tabs, below)

    # ---- Reconciliation
    rec.sheet_view.showGridLines = False
    widths(rec, {"A": 2, "B": 24, "C": 13, "D": 13, "E": 30, "F": 30, "G": 14, "H": 14, "I": 22, "J": 30})
    period_text = f"{period['period_start']:%m/%d/%Y} to {period['period_end']:%m/%d/%Y}"
    acct_text = f"{acct.get('nickname')}" + (f" ending {acct['last4']}" if acct.get("last4") else "") + f" ({acct.get('currency')})"
    review_on = ctx.get("review_completed_on")
    head = {"Company": acct.get("company") or "", "Account": acct_text, "Period": period_text,
            "Reconciled by": ctx.get("reconciled_by") or "",
            "Review completed on": datetime(review_on.year, review_on.month, review_on.day) if isinstance(review_on, date) else None}
    for k, rr in S.HEADER_ROWS.items():
        rec.cell(rr, 2, k)
        if k in head:
            rec.cell(rr, 3, head[k])
    rec.merge_cells(start_row=S.HEADER_ROWS["Reconciled by"], start_column=3, end_row=S.HEADER_ROWS["Reconciled by"], end_column=4)
    editable(rec.cell(S.HEADER_ROWS["Reconciled by"], 3))
    editable(rec.cell(S.HEADER_ROWS["Reconciled by"], 4))
    rd = rec.cell(S.REVIEW_DATE_ROW, 3)
    rd.number_format = DATE
    editable(rd)
    ddv = DataValidation(type="date", operator="greaterThan", formula1="36526", allow_blank=True, showErrorMessage=True,
                         errorStyle="warning", errorTitle="Review completed on", error="Enter a date, such as 04/03/2026.",
                         showInputMessage=True, promptTitle="Review completed on",
                         prompt="The date you finished your review. It records your sign-off; Claude's validation decides completion.")
    rec.add_data_validation(ddv)
    ddv.add(rd)
    fails = f'COUNTIF({CKQ}!$D${S.LIVE_FIRST}:$D${S.LIVE_LAST},"FAIL")'
    crow = S.HEADER_ROWS["Checks"]
    rec.cell(crow, 3, f'=IF({fails}=0,"All pass. Your edits are validated on upload.",'
                      f'"Check failed: "&INDEX({CKQ}!$B${S.LIVE_FIRST}:$B${S.LIVE_LAST},MATCH("FAIL",{CKQ}!$D${S.LIVE_FIRST}:$D${S.LIVE_LAST},0)))')
    rec.conditional_formatting.add(f"C{crow}", FormulaRule(formula=[f"{fails}>0"], font=Font(color="C00000", bold=True)))
    heading(rec, S.SUMMARY_HEAD, 3, ["Bank", "Ledger"], right=("Bank", "Ledger"))
    s_open, l_open = fig.get("statement_opening"), fig.get("ledger_opening")
    summary = [
        (S.OPEN_ROW, "Opening balance", f(s_open), f(l_open)),
        (S.IN_ROW, "Money in", f'=SUMIFS({L("G")},{L("B")},"Bank",{L("G")},">0")', f'=SUMIFS({L("G")},{L("B")},"Ledger",{L("G")},">0")'),
        (S.OUT_ROW, "Money out", f'=-SUMIFS({L("G")},{L("B")},"Bank",{L("G")},"<0")', f'=-SUMIFS({L("G")},{L("B")},"Ledger",{L("G")},"<0")'),
        (S.CLOSE_ROW, "Closing balance", f"=C{S.OPEN_ROW}+C{S.IN_ROW}-C{S.OUT_ROW}", f"=D{S.OPEN_ROW}+D{S.IN_ROW}-D{S.OUT_ROW}"),
        (S.ADJ_ROW, "Reconciling adjustments", f"=SUM({E('P')})", f"=SUM({E('Q')})"),
        (S.ADJUSTED_ROW, "Adjusted balance", f"=C{S.CLOSE_ROW}+C{S.ADJ_ROW}", f"=D{S.CLOSE_ROW}+D{S.ADJ_ROW}"),
    ]
    for rr, label, b, l in summary:
        rec.cell(rr, 2, label)
        for col, v in ((3, b), (4, l)):
            rec.cell(rr, col, v).number_format = MONEY
    for col in (3, 4):
        rec.cell(S.CLOSE_ROW, col).border = Border(top=hair)
        rec.cell(S.ADJUSTED_ROW, col).border = Border(top=hair)
    rr = S.REMAINING_ROW
    rec.cell(rr, 2, "Remaining difference").font = Font(bold=True, size=12)
    c = rec.cell(rr, 3, f"=ROUND(C{S.ADJUSTED_ROW}-D{S.ADJUSTED_ROW},2)")
    c.number_format = MONEY
    c.font = Font(bold=True, size=12)
    c.border = Border(top=thin, bottom=Side(style="double", color="808080"))
    rec.conditional_formatting.add(f"C{rr}", FormulaRule(formula=[f"ROUND($C${rr},2)=0"], fill=fill(GREEN)))
    rec.conditional_formatting.add(f"C{rr}", FormulaRule(formula=[f"ROUND($C${rr},2)<>0"], fill=fill(PINK)))
    outline(rec, 2, 2, S.REMAINING_ROW, 5)
    for cc in range(2, 11):
        rec.cell(S.LINE_ROW, cc).border = Border(bottom=dark)
    lg = S.LEGEND_ROW
    for cell, txt, colour in ((f"B{lg}", "Matched or done", GREEN), (f"C{lg}", "Ready to post", READY), (f"E{lg}", "Needs your review", PINK)):
        rec[cell] = txt
        rec[cell].fill = fill(colour)
        rec[cell].font = Font(size=9)
        rec[cell].alignment = Alignment(horizontal="center")
    rec.merge_cells(f"C{lg}:D{lg}")
    rec[f"F{lg}"] = "Grey cells fill in from the Journal entries tab. Timing items are approved here."
    muted(rec[f"F{lg}"])
    heading(rec, S.TABLE_HEAD, 2, S.REC_COLS, right=("Bank amount", "Ledger amount"))
    dv_cache = {}
    first = S.FIRST_ROW
    rec_written = []
    for n, row_ in enumerate(m["rows"]):
        rr = first + n
        b, l = row_.get("bank"), row_.get("ledger")
        carried = row_.get("carried")
        ld = (l.date if l else carried["date"] if carried else None)
        ldesc = ((("Carried: " if l.source == "Carried" else "") +
                  (f"{l.description}, check {l.reference}" if l.kind == "check" and l.reference else l.description)) if l else
                 carried["description"] if carried else None)
        lamt = l.amount if l else carried["amount"] if carried else None
        vals = [row_["ref"], b.date if b else None, ld, b.description if b else None, ldesc, f(b.amount) if b else None, f(lamt)]
        for j, v in enumerate(vals, 2):
            if isinstance(v, date) and not isinstance(v, datetime):
                v = datetime(v.year, v.month, v.day)
            cell = rec.cell(rr, j, v)
            cell.border = Border(bottom=hair)
            if j in (3, 4):
                cell.number_format = DATE
            if j in (7, 8):
                cell.number_format = MONEY
        rec_written.append((row_["ref"], vals[1], vals[2], vals[3], vals[4], vals[5], vals[6]))
        a, jc = rec.cell(rr, 9), rec.cell(rr, 10)
        kind = row_["approval"][0]
        if kind == "fixed":
            a.value = row_["approval"][1]
            flow(a)
        elif kind == "flow":
            a.value = (f'=IF(COUNTIF({JEQ}!$B:$B,$B{rr})=0,"{PENDING}",IF(INDEX({JEQ}!$J:$J,MATCH($B{rr},{JEQ}!$B:$B,0))="","{PENDING}",'
                       f'INDEX({JEQ}!$J:$J,MATCH($B{rr},{JEQ}!$B:$B,0))&""))')
            flow(a)
        else:
            a.value = row_["approval"][1]
            editable(a)
            labels_ = row_["approval"][2]
            key = tuple(labels_)
            if key not in dv_cache:
                timing = any(x == "Approved" for x in labels_)
                dv = DataValidation(type="list", formula1='"' + ",".join(labels_) + '"', allow_blank=True, showErrorMessage=True,
                                    errorTitle="Approval", error="Choose one of the listed options.", showInputMessage=True,
                                    promptTitle="Timing item" if timing else "Approval",
                                    prompt=("Approving confirms the treatment as a timing item. It doesn't mean the item has cleared the bank."
                                            if timing else "Choose what this line is. Approving applies that treatment."))
                rec.add_data_validation(dv)
                dv_cache[key] = dv
            dv_cache[key].add(a)
        jk = row_["jeref"][0]
        if jk == "flow":
            jc.value = f'=IF(COUNTIF({JEQ}!$B:$B,$B{rr})=0,"",INDEX({JEQ}!$K:$K,MATCH($B{rr},{JEQ}!$B:$B,0))&"")'
        else:
            jc.value = row_["jeref"][1] or None
        flow(jc)
    last = first + max(len(m["rows"]), 1) - 1
    needs = f'IFERROR(INDEX({E("R")},MATCH($B{first},{E("B")},0)),0)'
    chk = f'IFERROR(INDEX({JEQ}!$L:$L,MATCH($B{first},{JEQ}!$B:$B,0)),"")'
    ok = f'OR({chk}="Balanced",{chk}="New account: confirm")'
    approved = xl_approved(f"$I{first}")
    done = f'OR($I{first}="Matched",AND({approved},OR({needs}=0,AND($J{first}<>"",{ok}))))'
    ready = f'AND({approved},{needs}=1,$J{first}="",{ok})'
    rng = f"B{first}:H{last}"
    rec.conditional_formatting.add(rng, FormulaRule(formula=[done], fill=fill(GREEN), stopIfTrue=True))
    rec.conditional_formatting.add(rng, FormulaRule(formula=[ready], fill=fill(READY), stopIfTrue=True))
    rec.conditional_formatting.add(rng, FormulaRule(formula=[f'$B{first}<>""'], fill=fill(PINK)))
    outline(rec, S.TABLE_HEAD, 2, last, 10, side=thin)
    rec.freeze_panes = f"C{first}"

    # ---- Journal entries
    je.sheet_view.showGridLines = False
    widths(je, {"A": 2, "B": 8, "C": 12, "D": 30, "E": 34, "F": 12, "G": 12, "H": 3, "I": 39.4, "J": 26, "K": 26, "L": 20, "M": 30})
    je["B1"] = "Proposed entries, one per item. Standard review format: not an import file for your system."
    muted(je["B1"])
    je["B2"] = "Yellow cells are yours to edit. Select an account from the list. To split a line, insert a row under a yellow line."
    muted(je["B2"])
    heading(je, S.JE_HEAD, 2, ["Ref", "Date", "Account", "Description", "Debit", "Credit"], right=("Debit", "Credit"))
    heading(je, S.JE_HEAD, 9, ["Proposal", "Approval status", "JE number\n(once approved and posted)", "Check", "Comments"])
    je.row_dimensions[S.JE_HEAD].height = 30
    adv = DataValidation(type="list", formula1=f"={S.ACCOUNTS}!$A$2:$A$2000", allow_blank=True, showErrorMessage=True,
                         errorStyle="warning", errorTitle="Not in your chart of accounts",
                         error="This account isn't in your chart of accounts. You can keep it; Claude will ask you to confirm it when you upload.")
    je.add_data_validation(adv)
    jdv_cache = {}
    rr = S.JE_FIRST
    je_written = []
    for e in m["entries"]:
        h = rr
        je.cell(h, 2, e["ref"]).font = Font(bold=True)
        dcell = je.cell(h, 3, datetime(e["date"].year, e["date"].month, e["date"].day) if isinstance(e["date"], date) else e["date"])
        dcell.number_format = DATE
        je.cell(h, 5, e["title"]).font = Font(bold=True)
        je.cell(h, 9, e["proposal"]).alignment = Alignment(wrap_text=True, vertical="top")
        ap = je.cell(h, 10, e["approval"])
        if e["fixed"]:
            flow(ap)
        else:
            editable(ap)
            key = tuple(e["labels"])
            if key not in jdv_cache:
                dv = DataValidation(type="list", formula1='"' + ",".join(e["labels"]) + '"', allow_blank=True, showErrorMessage=True,
                                    errorTitle="Approval status", error="Choose one of the listed options.")
                je.add_data_validation(dv)
                jdv_cache[key] = dv
            jdv_cache[key].add(ap)
        je.cell(h, 11, e["je_number"] or None)
        editable(je.cell(h, 11))
        je.cell(h, 13, e["comment"] or None).alignment = Alignment(wrap_text=True, vertical="top")
        editable(je.cell(h, 13))
        for col in (2, 3, 5, 10, 11, 12):
            je.cell(h, col).alignment = Alignment(vertical="top")
        for k, ln in enumerate(e["lines"], 1):
            lr = h + k
            for col, v in ((4, ln.get("account")), (5, ln.get("description")), (6, num(ln.get("debit"))), (7, num(ln.get("credit")))):
                c = je.cell(lr, col, v if v not in ("",) else None)
                if col in (6, 7):
                    c.number_format = MONEY
                if not ln.get("fixed") and not e["fixed"]:
                    editable(c)
            if not ln.get("fixed") and not e["fixed"]:
                adv.add(je.cell(lr, 4))
        t = h + len(e["lines"]) + 1
        je.cell(t, 5, "Total").alignment = Alignment(horizontal="right")
        for col in "FG":
            c = je[f"{col}{t}"]
            c.value = f"=SUM({col}{h}:INDEX({col}:{col},ROW()-1))"
            c.number_format = MONEY
            c.border = Border(top=thin)
            c.font = Font(bold=True)
        outline(je, h, 2, t, 7)
        rg = lambda col: f"{col}{h}:INDEX(${col}:${col},ROW($F${t})-1)"
        missing = f'SUMPRODUCT(({rg("D")}="")*((({rg("F")}<>"")+({rg("G")}<>""))>0))'
        placeholder = f'SUMPRODUCT((LEFT({rg("D")},1)="[")*1)'
        unknown = f'SUMPRODUCT(({rg("D")}<>"")*(LEFT({rg("D")},1)<>"[")*(COUNTIF({S.ACCOUNTS}!$A$2:$A$2000,{rg("D")})=0))'
        needs_e = f'IFERROR(INDEX({E("R")},MATCH(B{h},{E("B")},0)),1)'
        appr = xl_approved(f"J{h}")
        je.cell(h, 12, f'=IF(AND(K{h}<>"",NOT({appr})),"Not approved",IF(AND({appr},{needs_e}=0),IF(K{h}<>"","Remove JE number","Not needed"),'
                       f'IF({missing}+{placeholder}>0,"Needs an account",IF(OR(ROUND(F{t}-G{t},2)<>0,F{t}=0),"Not balanced",'
                       f'IF({unknown}>0,"New account: confirm","Balanced")))))')
        je_written.append((e["ref"], h, t))
        rr = t + 2
    last_je = max(rr - 2, S.JE_FIRST)
    je.conditional_formatting.add(f"L{S.JE_FIRST}:L{last_je}", FormulaRule(formula=[f'L{S.JE_FIRST}="Balanced"'], fill=fill(GREEN)))
    je.conditional_formatting.add(f"L{S.JE_FIRST}:L{last_je}", FormulaRule(formula=[f'L{S.JE_FIRST}="New account: confirm"'], fill=fill(READY)))
    je.conditional_formatting.add(f"L{S.JE_FIRST}:L{last_je}", FormulaRule(
        formula=[f'OR(L{S.JE_FIRST}="Needs an account",L{S.JE_FIRST}="Not balanced",L{S.JE_FIRST}="Not approved",'
                 f'L{S.JE_FIRST}="Remove JE number")'], fill=fill(PINK)))
    je.freeze_panes = f"A{S.JE_HEAD + 1}"
    if not m["entries"]:
        je.cell(S.JE_FIRST, 2, "No entries are needed for this period.")
        muted(je.cell(S.JE_FIRST, 2))

    # ---- Guide
    write_guide(guide)

    # ---- Accounts
    heading(accts, 1, 1, ["Account"])
    accts.column_dimensions["A"].width = 34
    for i, a in enumerate(m["accounts"], 2):
        accts.cell(i, 1, a)
    accts["C1"] = ("Your chart of accounts. Add new accounts at the bottom, or upload a new chart to Claude to replace the list."
                   if ctx.get("accounts_source") and "no chart" not in ctx.get("accounts_source", "") else
                   "No chart of accounts was uploaded. Add accounts at the bottom, or upload your chart to Claude.")
    muted(accts["C1"])

    # ---- Setup (settings, rules, technical settings)
    write_setup(setup, prof, ctx, files)

    # ---- run record and expectations
    rec_section_row = expect_title
    rec_section_row = write_section(ck, rec_section_row, S.S_REC_ROWS, ["Ref", "Bank date", "Ledger date", "Bank description",
                                                                       "Ledger description", "Bank amount", "Ledger amount"], rec_written)
    snapshot = snapshot_digest({k: files[k]["grid"] for k in ("bank", "ledger")})
    run = run_record(m, snapshot, carry if str(ctx.get("completion") or "").startswith("Complete") else None,
                     cfacts if str(ctx.get("completion") or "").startswith("Complete") else None)
    rec_section_row = write_section(ck, rec_section_row, S.S_RUN, ["Field", "Value"], list(run.items()))
    exp = expectations(m, je_written, first)
    write_section(ck, rec_section_row, S.S_EXPECT, ["Cell", "Expected"], exp)

    setup.sheet_state = "hidden"
    ck.sheet_state = "hidden"
    wb.active = 0
    for ws in wb.worksheets:
        ws.sheet_view.tabSelected = ws.title == S.REC
    wb.calculation.fullCalcOnLoad = True
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def write_guide(guide):
    guide.sheet_view.showGridLines = False
    guide.page_setup.orientation = "landscape"
    guide.page_setup.fitToWidth = 1
    guide.page_setup.fitToHeight = 1
    guide.sheet_properties.pageSetUpPr.fitToPage = True
    widths(guide, {"A": 2, "B": 4, "C": 112})
    guide["B2"] = "How to review"
    guide["B2"].font = Font(bold=True, size=12)
    steps = ["Journal entries: check each proposed entry and edit the yellow cells, then choose its Approval status. One choice "
             "records your decision; some choices need no entry, such as Confirm bank error—no JE. Add a comment if you need to.",
             "Reconciliation: approve timing items (outstanding checks and payments, deposits in transit) in the yellow cells. "
             "Everything else fills in from the Journal entries tab.",
             "Post the approved entries in your accounting system, then type each JE number on the Journal entries tab.",
             "Upload the workbook back to Claude and ask it to validate your reviewed workbook. Claude validates your edits and "
             "returns the checked workbook. Next month, upload that checked workbook with the new bank and ledger exports."]
    for i, s in enumerate(steps, 1):
        guide.cell(3 + i, 2, f"{i}.").alignment = Alignment(vertical="top")
        guide.cell(3 + i, 3, s).alignment = Alignment(wrap_text=True, vertical="top")
        guide.row_dimensions[3 + i].height = 30
    guide["B9"] = "Good to know"
    guide["B9"].font = Font(bold=True)
    notes = ["Approving applies the treatment you chose. It doesn't guarantee a zero difference. Approving a timing item confirms "
             "it's a timing item; it doesn't mean the item has cleared the bank.",
             "A zero difference on its own doesn't mean the reconciliation is complete. Claude's validation decides completion after "
             "your upload, only when every check passes and nothing is pending. Review completed on is your own sign-off date; it "
             "doesn't override the checks.",
             "Check (Journal entries): Balanced means debits equal credits and every account is in your chart. Otherwise it says what "
             "to fix: Needs an account, Not balanced, New account: confirm, Not approved (a JE number was entered before approval), "
             "or Remove JE number (the choice needs no entry).",
             "Supporting tabs: Setup (settings and rules) and Checks (control detail and records) are hidden. Right-click any tab and "
             "choose Unhide to see them. Their status shows on Reconciliation under Checks."]
    for i, s in enumerate(notes):
        guide.cell(10 + i, 2, "-").alignment = Alignment(vertical="top")
        guide.cell(10 + i, 3, s).alignment = Alignment(wrap_text=True, vertical="top")
        guide.row_dimensions[10 + i].height = 32


def write_setup(setup, prof, ctx, files):
    widths(setup, {"A": 2, "B": 40, "C": 60, "D": 12})
    setup["B1"] = "Setup and rules. Claude reads these next month instead of a separate setup file. Tell Claude to change a setting."
    muted(setup["B1"])
    a, rules, accs = prof["account"], prof["rules"], prof["accounts"]
    src = lambda side: (f"{files[side]['name']}: " + "; ".join(f"{k} = {v}" for k, v in prof["sources"][side]["columns"].items() if v))
    settings = [("Company", a.get("company") or ""), ("Account", a.get("nickname", "") + (f" ending {a['last4']}" if a.get("last4") else "")),
                ("Currency", a.get("currency")), ("Accounting system", a.get("accounting_system") or ""),
                ("Cash GL account", accs.get("cash")), ("Bank fee account", accs.get("bank_fees") or "not set"),
                ("Reconciled by", ctx.get("reconciled_by") or ""), ("Bank file columns", src("bank")), ("Ledger file columns", src("ledger")),
                ("Matching", f"Exact amounts; dates within {rules.get('date_window_days')} days; deposits in transit from the last "
                             f"{rules.get('deposit_in_transit_days')} days"),
                ("JE format", "Standard review format"), ("Chart of accounts", ctx.get("accounts_source") or "not given"),
                ("Setup version", str(prof.get("profile_version", 1)))]
    row = write_section(setup, 3, S.S_SETTINGS, ["Setting", "Value"], settings)
    plain = [(f'Bank says "{d}"', accs.get("bank_fees") or "Bank fee entry", "Confirmed at setup") for d in rules.get("fee_descriptions", [])]
    plain += [(f'Bank description contains "{w}"', "Hold as a returned item", "Standard rule") for w in rules.get("returned_item_words", [])]
    for cr in prof.get("coding_rules", []):
        if cr.get("status", "active") == "active":
            plain.append((f"{cr.get('counterparty')}" + (f" with {', '.join(cr['words'])}" if cr.get("words") else "")
                          + (f", amount {cr['amount']}" if cr.get("amount") else ""), cr.get("account") or cr.get("description") or "",
                          cr.get("basis") or ""))
    row = write_section(setup, row, S.S_RULES, ["Rule", "Suggests", "Basis"], plain or [("No rules yet", "", "")])
    setup.cell(row - 1, 2, "Rules only suggest. Every entry still needs your approval.")
    muted(setup.cell(row - 1, 2))
    row += 1
    write_section(setup, row, S.S_TECH, ["Setting", "Value", "Type"], profiles.flatten(prof))


def carry_forward(m):
    """What the next period must carry: approved bank-side items and approved entries, with any reported JE number."""
    r = m["r"]
    by_id = r["by_id"]
    label = r["label"]
    je_by_ref = {e["ref"]: e for e in m["entries"]}
    rows = []
    for i in r["items"]:
        if i.status not in ("Approved", "Approved (carried)"):
            continue
        ref = m["iref"][i.id]
        item_id = i.id if i.status == "Approved (carried)" else f"{label}-{ref}"
        if i.side == "Bank" and (i.recipe or i.fixed_effect is not None):
            src = by_id[i.recipe[0][1]] if i.recipe else None
            rows.append({"Kind": KIND_OF_CATEGORY[i.category], "Item ID": item_id, "Origin period": i.origin or label,
                         "Date": str(src.date) if src else i.line_date, "Reference": src.reference if src else i.reference,
                         "Description": src.description if src else i.description, "Amount": f(effect_of(i, by_id)), "Category": i.category,
                         "Account": "", "Posting ref": "", "Approved by": i.approved_by, "Note": i.note})
        elif i.side == "Ledger":
            src = by_id[i.recipe[0][1]] if i.recipe else None
            e = je_by_ref.get(ref)
            accts = [l for l in (e["lines"] if e else []) if (l.get("account") or "") != r["profile"]["accounts"]["cash"]]
            account = "; ".join(f"{l.get('account')} {money(Decimal(str(l.get('debit') or l.get('credit') or 0)))}" for l in accts) \
                if len(accts) > 1 else (accts[0].get("account") if accts else i.account)
            rows.append({"Kind": "pending_adjustment", "Item ID": item_id, "Origin period": i.origin or label,
                         "Date": str(src.date) if src else i.line_date, "Reference": src.reference if src else i.reference,
                         "Description": (i.description or f"{i.category}: {src.description if src else ''}").strip(),
                         "Amount": f(effect_of(i, by_id)), "Category": i.category, "Account": account,
                         "Posting ref": (e["je_number"] if e else "") or "", "Approved by": i.approved_by, "Note": i.note})
    # the reason and the reviewer's comment travel with the item, so next period can follow it without this workbook open
    items_by_carry = {(i.id if i.status == "Approved (carried)" else f"{label}-{m['iref'][i.id]}"): i for i in r["items"]
                      if i.status in ("Approved", "Approved (carried)")}
    for row in rows:
        i = items_by_carry.get(row["Item ID"])
        if i is not None and i.status == "Approved":
            e = je_by_ref.get(m["iref"][i.id])
            reason = i.note if i.note and not i.note.startswith(GENERIC_NOTES) else MEANING.get(i.treatment or "", "")
            reason = reason[:1].upper() + reason[1:] if reason else ""
            comment = (e or {}).get("comment") or ""
            row["Note"] = "; ".join(x for x in (reason, f"Comment: {comment}" if comment else "") if x)
        for k in CARRY_COLS:
            row.setdefault(k, "")
            if row[k] is None:
                row[k] = ""
    fig = r["figures"]
    acct = r["profile"]["account"]
    facts = {"Period start": str(r["period"]["period_start"]), "Period end": str(r["period"]["period_end"]), "Status": r["status"],
             "Account": acct["nickname"], "Currency": acct["currency"], "Stated bank closing balance": f(fig.get("statement_closing")),
             "Actual ledger closing balance": f(fig.get("ledger_closing")), "Completion": m["ctx"].get("completion") or ""}
    return rows, facts


def run_record(m, snapshot, carry, cfacts):
    r, ctx = m["r"], m["ctx"]
    fig, period = r["figures"], r["period"]
    acct = r["profile"]["account"]
    files = ctx["files"]
    rd = ctx.get("review_completed_on")
    out = {"Workbook format": S.FORMAT, "Skill version": ctx.get("version"), "Account": acct["nickname"], "Currency": acct["currency"],
           "Period start": str(period["period_start"]), "Period end": str(period["period_end"]),
           "Stated bank opening balance": f(fig.get("statement_opening")), "Stated bank closing balance": f(fig.get("statement_closing")),
           "Ledger opening balance": f(fig.get("ledger_opening")), "Ledger closing balance (stated)": f(period.get("ledger_closing")),
           "Bank file": files["bank"]["name"], "Bank file SHA-256": files["bank"]["sha"], "Bank file type": files["bank"]["type"],
           "Bank sheet": files["bank"].get("sheet") or "", "Ledger file": files["ledger"]["name"], "Ledger file SHA-256": files["ledger"]["sha"],
           "Ledger file type": files["ledger"]["type"], "Ledger sheet": files["ledger"].get("sheet") or "", "Snapshot digest": snapshot,
           "Carry forward digest": carry_digest(cfacts, carry) if cfacts is not None else "", "Draft created": ctx.get("created_on"),
           "Last validated": ctx.get("validated_on") or "", "Validated at": ctx.get("validated_at") or "",
           "Completion": ctx.get("completion") or "Not complete",
           "Review completed on (entered by the reviewer)": rd.isoformat() if isinstance(rd, date) else "",
           "Migrated from": ctx.get("migrated_from") or "",
           "Engine status": r["status"], "Previous period": (r["prior"] or {}).get("period_label", "") if r.get("prior") else "",
           "Posted before opening": ", ".join(period.get("posted_before_opening") or []),
           "Accounts source": ctx.get("accounts_source") or ""}
    return out


def expectations(m, je_written, first):
    """(cell, expected value) for formulas whose result the script knows. `check` compares them after recalculation."""
    r = m["r"]
    fig = r["figures"]
    fx = figures(m)
    out = []
    put = lambda cell, v: out.append((cell, f(v) if isinstance(v, Decimal) else v))
    if fig.get("bank_closing") is not None:
        put(f"{S.REC}!C{S.CLOSE_ROW}", fig["bank_closing"])
    if fig.get("ledger_closing") is not None:
        put(f"{S.REC}!D{S.CLOSE_ROW}", fig["ledger_closing"])
    put(f"{S.REC}!C{S.ADJ_ROW}", fx["bank_adj"])
    put(f"{S.REC}!D{S.ADJ_ROW}", fx["ledger_adj"])
    if "remaining" in fx:
        put(f"{S.REC}!C{S.ADJUSTED_ROW}", fx["adjusted_bank"])
        put(f"{S.REC}!D{S.ADJUSTED_ROW}", fx["adjusted_ledger"])
        put(f"{S.REC}!C{S.REMAINING_ROW}", fx["remaining"])
    put(f"{S.REC}!C{S.HEADER_ROWS['Checks']}", checks_line(m))
    vals, checks = live_values(m)
    for k, v in enumerate(vals):
        put(f"{S.CHECKS}!C{S.LIVE_FIRST + k}", Decimal(str(v)))
        put(f"{S.CHECKS}!D{S.LIVE_FIRST + k}", "PASS" if Decimal(str(v)).quantize(Decimal("0.01")) == 0 else "FAIL")
    je_by_ref = {e["ref"]: e for e in m["entries"]}
    for n, row_ in enumerate(m["rows"]):
        rr = first + n
        kind = row_["approval"][0]
        if kind == "flow":
            e = je_by_ref.get(row_["ref"])
            put(f"{S.REC}!I{rr}", (e["approval"] if e and e["approval"] else PENDING))
        if row_["jeref"][0] == "flow":
            e = je_by_ref.get(row_["ref"])
            put(f"{S.REC}!J{rr}", (e["je_number"] if e else "") or "")
    for ref, h, t in je_written:
        e = je_by_ref[ref]
        put(f"{S.JE}!L{h}", checks[ref])
        dr = sum((Decimal(str(l["debit"])) for l in e["lines"] if l.get("debit") not in (None, "")), ZERO)
        cr = sum((Decimal(str(l["credit"])) for l in e["lines"] if l.get("credit") not in (None, "")), ZERO)
        put(f"{S.JE}!F{t}", dr)
        put(f"{S.JE}!G{t}", cr)
    return out


def build(r, ctx, path):
    """Build the model, cross-check it against the engine, and write the workbook. Returns (path, model)."""
    m = build_model(r, ctx)
    if ctx.get("check_engine"):
        fx = figures(m)
        if fx["bank_adj"] != r["figures"]["bank_side"] or fx["ledger_adj"] != r["figures"]["book_side"]:
            raise ValueError(f"Workbook adjustments ({fx['bank_adj']}, {fx['ledger_adj']}) differ from the script's "
                             f"({r['figures']['bank_side']}, {r['figures']['book_side']}). Nothing was written.")
    render(m, path)
    return path, m
