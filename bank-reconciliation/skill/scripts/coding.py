"""Coding and naming suggestions: optional, and kept apart from the reconciliation.

For each bank line that needs a coding decision (an unidentified bank item), this compares the line with a
coded-transactions history and any saved coding rules, and reports:

- the description the historical naming pattern would give, with any part it cannot establish left as [..?]
- a suggested account, with the history or saved rule it rests on, or the conflict if history disagrees
- the variable parts on the line (months, years, invoice and installment numbers) and what established their meaning
- the questions that remain

It reads the finished reconciliation and never feeds back into it: matching, review items, decisions, figures, and
status are the same with or without it. Every suggestion is unapproved until the user accepts it, and accepting
wording or coding here never approves an entry; that is a reconciliation decision, recorded when the reviewed workbook is validated.
Source descriptions are never changed.
"""
from __future__ import annotations

import hashlib, json, re
from collections import Counter
from datetime import date
from decimal import Decimal
from pathlib import Path

from foundation import tables
from foundation.tables import InputError

SUGGEST_FOR = {"Unidentified bank item"}
STATUSES = ("Suggested", "Accepted", "Changed by user", "Rejected")
SHEET = "Coding suggestions"
COLS = ["Line ID", "Date", "Amount", "Source description (unchanged)", "Expected from history",
        "Recommended wording (Claude, not approved)", "Suggested account", "Basis", "Variable parts on the line",
        "Questions and uncertainty", "Saved rule", "Description status", "Final description", "Account status",
        "Final account", "Confirmed values", "Decided by", "Decided on", "Note", "Fingerprint", "Evidence"]

GENERIC = {"ACH", "CARD", "DEBIT", "CREDIT", "POS", "WIRE", "IN", "OUT", "PAYMENT", "PMT", "INV", "INVOICE", "REF", "NO",
           "TRANSFER", "TFR", "DD", "SO", "CHECK", "CHQ", "PURCHASE", "ONLINE", "BILL", "THE", "AND", "OF", "FOR", "TO",
           "FROM", "DEPOSIT", "WITHDRAWAL", "PAID"}
MONTH_NUM = {m: i for i, names in enumerate([("JAN", "JANUARY"), ("FEB", "FEBRUARY"), ("MAR", "MARCH"), ("APR", "APRIL"),
                                             ("MAY",), ("JUN", "JUNE"), ("JUL", "JULY"), ("AUG", "AUGUST"),
                                             ("SEP", "SEPT", "SEPTEMBER"), ("OCT", "OCTOBER"), ("NOV", "NOVEMBER"),
                                             ("DEC", "DECEMBER")], 1) for m in names}
MONTH_NAME = ["", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October",
              "November", "December"]
PATTERNS = [  # order matters: earlier kinds claim their text first
    ("invoice number", re.compile(r"\b(INV(?:OICE)?|BILL)\b\.?\s*(?:NO\.?\s*|#\s*)?(\d{3,})\b", re.I)),
    ("n of m", re.compile(r"\b(\d{1,3})\s+OF\s+(\d{1,3})\b", re.I)),
    ("n/m", re.compile(r"(?<![\d/])(\d{1,3})/(\d{1,3})(?![\d/])")),
    ("month", re.compile(r"\b(" + "|".join(sorted(MONTH_NUM, key=len, reverse=True)) + r")\b\.?(?:\s+(20\d{2}))?", re.I)),
    ("year", re.compile(r"\b(20\d{2})\b")),
    ("number", re.compile(r"\b(\d+)\b")),
]


# ---------------------------------------------------------------- reading descriptions

def variable_parts(text):
    """The variable parts of a description, with what (if anything) establishes their meaning."""
    taken, out = [], []
    for kind, rx in PATTERNS:
        for m in rx.finditer(text or ""):
            if any(m.start() < e and s < m.end() for s, e in taken):
                continue
            taken.append((m.start(), m.end()))
            part = {"kind": kind, "text": m.group(0), "span": [m.start(), m.end()]}
            if kind == "invoice number":
                part.update(value=m.group(2), meaning="invoice number", established_by=f"labelled '{m.group(1)}' on the line")
            elif kind == "n of m":
                part.update(value=f"{m.group(1)} of {m.group(2)}", meaning="n of m", established_by="written as 'n of m' on the line")
            elif kind == "n/m":
                part.update(value=f"{m.group(1)}/{m.group(2)}", n=m.group(1), m=m.group(2), meaning=None, established_by=None)
            elif kind == "month":
                mon = MONTH_NUM[m.group(1).upper()]
                part.update(value=MONTH_NAME[mon] + (f" {m.group(2)}" if m.group(2) else ""), month=mon,
                            year=int(m.group(2)) if m.group(2) else None, meaning="month", established_by="named on the line")
            elif kind == "year":
                part.update(value=m.group(1), meaning="year", established_by="on the line")
            else:
                part.update(value=m.group(1), meaning=None, established_by=None)
            out.append(part)
    return sorted(out, key=lambda p: p["span"][0])


def mask(text):
    """The description with its variable parts replaced by placeholders: the naming pattern."""
    out, pos = [], 0
    for p in variable_parts(text):
        s, e = p["span"]
        out.append(text[pos:s])
        out.append({"invoice number": p["text"][: p["text"].upper().rfind(p["value"])] + "{invoice number}",
                    "n of m": "{n} of {m}", "n/m": "{n}/{m}", "month": "{month}", "year": "{year}",
                    "number": "{number}"}[p["kind"]])
        pos = e
    out.append(text[pos:])
    return re.sub(r"\s+", " ", "".join(out)).strip()


def significant(text):
    """Words that identify the counterparty and purpose: no generic banking words, months, or numbers."""
    plain = text or ""
    for p in reversed(variable_parts(plain)):
        s, e = p["span"]
        plain = plain[:s] + " " + plain[e:]
    words = re.findall(r"[A-Za-z][A-Za-z&']+", plain.upper())
    return [w for w in words if w not in GENERIC and w not in MONTH_NUM]


def key(text):
    """The counterparty key: the first two significant words."""
    return tuple(significant(text)[:2])


def months_between(d, year, month):
    return (d.year * 12 + d.month) - (year * 12 + month)


def month_label(d, back):
    n = d.year * 12 + d.month - 1 - back
    return f"{MONTH_NAME[n % 12 + 1]} {n // 12}"


def evidence_tokens(text):
    """Numbers and month names in a text, for checking that wording adds nothing beyond the evidence."""
    up = (text or "").upper()
    return set(re.findall(r"\d+", up)) | {MONTH_NAME[MONTH_NUM[w]] for w in re.findall(r"[A-Z]+", up) if w in MONTH_NUM}


# ---------------------------------------------------------------- history

def read_history(path, cfg, label):
    """Coded lines from a history file (or a ledger export with an account column), as dicts."""
    cols = cfg["columns"]
    sheet, rows, _ = tables.read_grid(path, cfg.get("sheet"))
    name = Path(path).name + (f", sheet {sheet}" if sheet else "")
    needed = [v for v in cols.values() if v]
    h = tables.locate_header(rows, needed, cfg.get("header_fingerprint", needed), label)
    header = [tables.text(c) for c in rows[h]]
    pos = {c.casefold(): j for j, c in enumerate(header) if c}
    get = lambda r, role: r[pos[cols[role].casefold()]] if cols.get(role) and pos.get(cols[role].casefold()) is not None \
        and pos[cols[role].casefold()] < len(r) else None
    out = []
    for i, r in enumerate(rows[h + 1:], start=h + 2):
        cells = [tables.text(c) for c in r]
        if not any(cells) or any(tables.is_summary_text(c) for c in cells if c):
            continue
        d = tables.parse_date(get(r, "date"), cfg.get("date_format"))
        if cols.get("amount"):
            amt = tables.parse_amount(get(r, "amount"))[0]
            if amt is not None and cfg.get("sign") == "reverse":
                amt = -amt
        else:
            dr, cr = tables.parse_amount(get(r, "debit"))[0], tables.parse_amount(get(r, "credit"))[0]
            inflow_credit = cfg.get("inflow") == "credit"
            amt = (abs(cr) if inflow_credit else -abs(cr)) if cr else (-abs(dr) if inflow_credit else abs(dr)) if dr else None
        desc, account = tables.text(get(r, "description")), tables.text(get(r, "account"))
        if d is None or amt is None or not desc or not account:
            continue
        ref = tables.text(get(r, "id")) or tables.text(get(r, "reference")) or f"{label[:1]}-R{i:04d}"
        out.append({"id": ref, "date": d, "description": desc, "amount": amt, "account": account, "source": f"{name} row {i}"})
    return out, {"file": name, "sha256": tables.sha256(path), "lines": len(out)}


def load_history(period, profile):
    hist, sources = [], []
    hcfg = profile.get("sources", {}).get("history")
    if period.get("history_file"):
        if not hcfg:
            raise InputError("A history file was given, but the setup has no confirmed history mapping (sources.history).")
        lines, info = read_history(period["history_file"], hcfg, "History")
        hist += lines
        sources.append(info)
    lcfg = profile["sources"]["ledger"]
    if lcfg.get("columns", {}).get("account"):
        lines, info = read_history(period["ledger_file"], lcfg, "Ledger")
        hist += lines
        sources.append({**info, "note": "ledger lines with their account"})
    return hist, sources


# ---------------------------------------------------------------- suggestions

def rule_applies(rule, line):
    if rule.get("status", "active") != "active" or key(rule["counterparty"]) != key(line.description):
        return False
    up = line.description.upper()
    if any(w.upper() not in up for w in rule.get("words") or []):
        return False
    if rule.get("amount") and Decimal(str(rule["amount"])) != abs(line.amount):
        return False
    if rule.get("direction") and rule["direction"] != ("in" if line.amount > 0 else "out"):
        return False
    return True


def rule_text(rule):
    when = [f"the counterparty is {rule['counterparty']}"]
    if rule.get("words"):
        when.append("the line mentions " + ", ".join(f"'{w}'" for w in rule["words"]))
    if rule.get("amount"):
        when.append(f"the amount is {rule['amount']}")
    then = [x for x in (f"account {rule['account']}" if rule.get("account") else "",
                        f"description '{rule['description']}'" if rule.get("description") else "",
                        "n/m on the line means installment n of m" if rule.get("slash_means") else "") if x]
    return (f"Saved rule {rule['id']} (saved {rule['saved_on']} by {rule['saved_by']}): when " + " and ".join(when)
            + ", suggest " + "; ".join(then) + f". Basis: {rule['basis']}")


def fill(template, parts, slash_installment):
    """Fill a naming pattern only with values the line establishes; anything else stays a visible question."""
    get = lambda kind: next((p for p in parts if p["kind"] == kind), None)
    inv, nofm, slash, month, year = (get(k) for k in ("invoice number", "n of m", "n/m", "month", "year"))
    out = template
    out = out.replace("{invoice number}", inv["value"] if inv else "[invoice number?]")
    if nofm:
        n_of_m = nofm["value"]
    elif slash and slash_installment:
        n_of_m = f"{slash['n']} of {slash['m']}"
    elif slash:
        n_of_m = f"[{slash['n']} of {slash['m']}?]"
    else:
        n_of_m = "[n of m?]"
    out = out.replace("{n} of {m}", n_of_m).replace("{n}/{m}", slash["value"] if slash else "[n/m?]")
    out = out.replace("{month}", month["value"] if month else "[service month?]")
    out = out.replace("{year}", year["value"] if year else "[year?]")
    return out.replace("{number}", "[number?]").replace("{purpose}", "[purpose?]")


def counterparty_name(matched, line):
    for h in sorted(matched, key=lambda h: h["date"], reverse=True):
        if " - " in h["description"]:
            return h["description"].split(" - ")[0].strip()
    return " ".join(w.title() for w in key(line.description))


def suggest_line(line, hist, rules):
    parts = variable_parts(line.description)
    k = key(line.description)
    inflow = line.amount > 0
    matched = [h for h in hist if k and key(h["description"])[:len(k)] == k and (h["amount"] > 0) == inflow]
    matched.sort(key=lambda h: (h["date"], h["id"]))
    rule = next((x for x in rules if rule_applies(x, line)), None)
    questions, basis = [], []

    slash = next((p for p in parts if p["kind"] == "n/m"), None)
    slash_installment = bool(rule and rule.get("slash_means") == "installment")
    if slash and slash_installment:
        slash.update(meaning="installment n of m", established_by=f"saved rule {rule['id']}")

    accounts = Counter(h["account"] for h in matched)
    refs = {a: [h["id"] for h in matched if h["account"] == a] for a in accounts}
    conflict = None
    suggested = None
    history_accounts = set(accounts)
    if rule:
        basis.append(rule_text(rule) + " A saved rule only triggers this suggestion: it doesn't establish what this "
                     "transaction was for, and it doesn't outweigh conflicting history.")
    if rule and rule.get("account") and history_accounts <= {rule["account"]}:
        suggested = rule["account"]
        if matched:
            basis.append(f"History agrees: {', '.join(refs[rule['account']])}")
    elif len(accounts) == 1 and not (rule and rule.get("account")):
        suggested = next(iter(accounts))
        basis.append(f"{len(matched)} of {len(matched)} similar earlier lines were coded to {suggested}: {', '.join(refs[suggested])}")
    elif accounts:  # history disagrees with itself, or with the saved rule: a question, never a rule
        conflict = [{"account": a, "lines": refs[a]} for a in sorted(accounts, key=lambda a: (-accounts[a], a))]
        if rule and rule.get("account") and rule["account"] not in accounts:
            conflict.insert(0, {"account": rule["account"], "lines": [f"saved rule {rule['id']}"]})
        elif rule and rule.get("account"):
            next(c for c in conflict if c["account"] == rule["account"])["lines"].append(f"saved rule {rule['id']}")
        basis.append("The evidence disagrees: " + "; ".join(f"{c['account']}: {', '.join(c['lines'])}" for c in conflict))
        questions.append("The evidence points to more than one account (" +
                         "; ".join(f"{c['account']}: {', '.join(c['lines'])}" for c in conflict) + "). Which applies to this one?")

    templates = {mask(h["description"]) for h in matched}
    if rule and rule.get("description") and templates <= {rule["description"]}:
        template = rule["description"]
    elif rule and rule.get("description") or len(templates) > 1:  # a rule can't settle a purpose history disagrees on
        template = counterparty_name(matched, line) + " - {purpose}" if matched else rule["description"]
    elif templates:
        template = templates.pop()
    else:
        template = None
    expected = fill(template, parts, slash_installment) if template else None

    period = None
    if template and "{month}" in template and not any(p["kind"] == "month" for p in parts):
        lags = []
        for h in matched:
            for p in variable_parts(h["description"]):
                if p["kind"] == "month" and p.get("year"):
                    lags.append((h["id"], months_between(h["date"], p["year"], p["month"])))
                    break
        if len(lags) >= 2 and len({lag for _, lag in lags}) == 1:
            lag = lags[0][1]
            period = {"value": month_label(line.date, lag), "status": "inferred, needs confirmation",
                      "basis": f"In {', '.join(i for i, _ in lags)} the month named was {lag} month(s) before the payment month. "
                               "This is a pattern, not something the line states."}
            questions.append(f"Is this for {period['value']}? That is inferred from the pattern in {', '.join(i for i, _ in lags)}, "
                             "not stated on the line.")
        else:
            questions.append("Which month or period does this cover? The line doesn't say, and history gives no consistent pattern.")
    if slash and not slash_installment:
        questions.append(f"What does '{slash['value']}' on the line mean? It might be payment {slash['n']} of {slash['m']}, "
                         "but the line and the setup don't establish that.")
    if template and "{purpose}" in template:
        questions.append("What was this for? Earlier descriptions for this counterparty differ, and the line doesn't say.")
    if not matched and not rule:
        questions.append("No similar coded history, so there is no suggestion. The reconciliation asks about this line as usual.")

    purpose_words = significant(line.description)[len(k):]
    same_amount = len(matched) >= 2 and len({abs(h["amount"]) for h in matched} | {abs(line.amount)}) == 1
    if rule:
        rule_offer = {"possible": False, "reason": f"Saved rule {rule['id']} already applies."}
    elif len(matched) < 2:
        rule_offer = {"possible": False, "reason": "No recurring pattern: fewer than two similar earlier lines."}
    elif not (purpose_words or same_amount):
        rule_offer = {"possible": False, "reason": "Nothing on the line beyond the counterparty tells these transactions apart, "
                                                     "so a rule would cover the supplier alone. Record this transaction only."}
    else:
        rule_offer = {"possible": True, "counterparty": counterparty_name(matched, line),
                      "words": purpose_words or None, "amount": str(abs(line.amount)) if same_amount else None}

    return {"line_id": line.line_id, "date": str(line.date), "amount": str(line.amount),
            "source_description": line.description, "reference": line.reference,
            "expected_from_history": expected, "naming_pattern": template,
            "suggested_account": suggested, "account_conflict": conflict, "basis": basis,
            "history_lines": [{k2: (str(v) if isinstance(v, (date, Decimal)) else v) for k2, v in h.items()} for h in matched],
            "variable_parts": [{k2: v for k2, v in p.items() if k2 != "span"} for p in parts],
            "period_suggestion": period, "questions": questions, "rule": rule["id"] if rule else None,
            "save_rule_offer": rule_offer, "fingerprint": fingerprint(line),
            "evidence": evidence_digest(expected, suggested, period, rule)}


def evidence_digest(expected, account, period, rule):
    """What a suggestion rests on. A rejection holds until this changes (new history or a changed rule)."""
    blob = json.dumps([expected, account, (period or {}).get("value"), rule["id"] if rule else None])
    return hashlib.sha256(blob.encode()).hexdigest()[:12]


def fingerprint(line):
    return hashlib.sha256(json.dumps(list(line.content())).encode()).hexdigest()[:16]


def suggest(r, period_path):
    """Suggestions for this run, or None when there is neither history nor a saved rule."""
    from bank_rules import load_period
    period = load_period(period_path)
    profile = r["profile"]
    hist, sources = load_history(period, profile)
    rules = [x for x in profile.get("coding_rules", []) if x.get("status", "active") == "active"]
    if not hist and not rules:
        return None
    decided = {x["line_id"]: x for x in load_decisions(period.get("coding_decisions"))}
    lines = []
    for it in r["items"]:
        if it.type not in SUGGEST_FOR:
            continue
        s = suggest_line(r["by_id"][it.lines[0]], hist, rules)
        s["item_id"] = it.id
        d = decided.get(s["line_id"])
        if d and d.get("fingerprint") and d["fingerprint"] != s["fingerprint"]:
            s["coding_decision"] = None
            s["stale_decision"] = "The line changed since its coding was recorded, so that coding is not applied."
        else:
            s["coding_decision"] = d
            rejected = d and "Rejected" in (d.get("description_status"), d.get("account_status"))
            if rejected and d.get("evidence") and d["evidence"] != s["evidence"]:
                s["rejection"] = "Rejected earlier, but the evidence has changed since, so it may be offered again."
            elif rejected:
                s["rejection"] = "Rejected for this line. Don't propose it again unless the evidence changes."
        lines.append(s)
    return {"history_sources": sources, "history_lines": len(hist), "active_rules": [x["id"] for x in rules], "lines": lines,
            "note": "Suggestions only. They don't change matching, decisions, figures, or status, and accepting one approves no entry."}


# ---------------------------------------------------------------- coding decisions (separate from reconciliation decisions)

FIELDS = ["line_id", "recommended", "description_status", "final_description", "account_status", "final_account", "evidence",
          "confirmed_values", "decided_by", "decided_on", "note", "fingerprint"]


def load_decisions(path):
    if not path or not Path(path).exists():
        return []
    return [{k: x.get(k) for k in FIELDS} for x in json.loads(Path(path).read_text(encoding="utf-8"))]


def check_decision(x, s):
    """Problems with one coding decision against its suggestion. Wording must stay within the evidence."""
    probs = []
    if s is None:
        return [f"{x.get('line_id')}: not a line with coding suggestions in this run"]
    confirmed = x.get("confirmed_values") or {}
    allowed = evidence_tokens(s["source_description"] + " " + (s.get("reference") or "")) | evidence_tokens(" ".join(map(str, confirmed.values())))
    rec = x.get("recommended") or ""
    if rec:
        extra = evidence_tokens(rec) - allowed
        inferred = evidence_tokens(s["period_suggestion"]["value"]) if s.get("period_suggestion") else set()
        if "to confirm" in rec.lower():
            extra -= inferred
        if extra:
            probs.append(f"{x['line_id']}: recommended wording uses {sorted(extra)}, which the line doesn't state and the user hasn't "
                         "confirmed. Use a placeholder or mark an inferred value '(to confirm)'.")
    for k in ("description_status", "account_status"):
        if (x.get(k) or "Suggested") not in STATUSES:
            probs.append(f"{x['line_id']}: {k} must be one of {', '.join(STATUSES)}")
    ds, acs = x.get("description_status") or "Suggested", x.get("account_status") or "Suggested"
    if ds in ("Accepted", "Changed by user") and not x.get("final_description"):
        probs.append(f"{x['line_id']}: final_description is needed when the description is {ds.lower()}")
    if ds == "Accepted" and x.get("final_description"):
        fd = x["final_description"]
        if "to confirm" in fd.lower() or "?]" in fd:
            probs.append(f"{x['line_id']}: the accepted description still has an unconfirmed part. Ask, then record the answer in confirmed_values.")
        extra = evidence_tokens(fd) - allowed
        if extra:
            probs.append(f"{x['line_id']}: the accepted description uses {sorted(extra)}, which isn't on the line or confirmed.")
    choices = [c["account"] for c in s.get("account_conflict") or []]
    if acs == "Accepted" and not s.get("suggested_account") and x.get("final_account") not in choices:
        probs.append(f"{x['line_id']}: there is no suggested account to accept. If the user picked one of the conflicting accounts, "
                     "give it as final_account; otherwise use 'Changed by user'.")
    if acs == "Changed by user" and not x.get("final_account"):
        probs.append(f"{x['line_id']}: final_account is needed when the user names a different account")
    return probs


def complete(x, s, decided_by, decided_on):
    x = {k: x.get(k) for k in FIELDS}  # statuses left unset keep any earlier value when merged
    if x["account_status"] == "Accepted" and not x["final_account"] and s:
        x["final_account"] = s.get("suggested_account")
    x["confirmed_values"] = x["confirmed_values"] or {}
    user_decision = any(x[k] not in (None, "", "Suggested") for k in ("description_status", "account_status")) or x["confirmed_values"]
    if user_decision:  # Claude's own wording alone is not anyone's decision, so it carries no decider
        x["decided_by"] = x["decided_by"] or decided_by or ""
        x["decided_on"] = x["decided_on"] or decided_on
    x["fingerprint"] = s["fingerprint"] if s else x.get("fingerprint")
    x["evidence"] = s["evidence"] if s else x.get("evidence")
    return x


def merge(existing, new):
    by = {x["line_id"]: x for x in existing}
    for x in new:
        old = by.get(x["line_id"], {})
        by[x["line_id"]] = {**old, **{k: v for k, v in x.items() if v not in (None, "", {})}}
    for x in by.values():
        x["description_status"] = x.get("description_status") or "Suggested"
        x["account_status"] = x.get("account_status") or "Suggested"
    return list(by.values())


# ---------------------------------------------------------------- workbook sheet

def fill_sheet(ws, coding):
    from openpyxl.styles import Alignment, Font
    from foundation.xlsx import MONEY, table
    ws["A1"] = ("Suggestions only. Nothing on this sheet approves an entry or changes the reconciliation; "
                "entries are approved in the Decision log. Source descriptions are as exported.")
    ws["A1"].font = Font(italic=True)
    ws["A2"] = "History used: " + ("; ".join(f"{s['file']} ({s['lines']} coded lines)" for s in coding["history_sources"]) or "none") \
        + ("; saved rules: " + ", ".join(coding["active_rules"]) if coding["active_rules"] else "")
    table(ws, COLS, [10, 11, 11, 32, 40, 40, 22, 50, 36, 60, 10, 13, 40, 13, 22, 24, 16, 12, 30, 18, 14], start_row=4)
    for s in coding["lines"]:
        d = s.get("coding_decision") or {}
        parts = "; ".join(f"{p['text']}: {p['meaning'] or 'meaning not established'}"
                          + (f" ({p['established_by']})" if p.get("established_by") else "") for p in s["variable_parts"])
        qs = list(s["questions"])
        if s.get("period_suggestion"):
            qs.insert(0, f"Period suggestion: {s['period_suggestion']['value']} ({s['period_suggestion']['status']}). "
                         + s["period_suggestion"]["basis"])
        if s.get("stale_decision"):
            qs.insert(0, s["stale_decision"])
        if s.get("rejection"):
            qs.insert(0, s["rejection"])
        ws.append([s["line_id"], s["date"], float(s["amount"]), s["source_description"], s["expected_from_history"] or "(no history)",
                   d.get("recommended") or "", s["suggested_account"] or "(none)", " | ".join(s["basis"]) or "No similar history",
                   parts, " | ".join(qs), s["rule"] or "", d.get("description_status") or "Suggested", d.get("final_description") or "",
                   d.get("account_status") or "Suggested", d.get("final_account") or "",
                   "; ".join(f"{k}: {v}" for k, v in (d.get("confirmed_values") or {}).items()),
                   d.get("decided_by") or "", d.get("decided_on") or "", d.get("note") or "", s["fingerprint"], d.get("evidence") or ""])
        ws.cell(ws.max_row, 3).number_format = MONEY
        for c in ws[ws.max_row]:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "B5"
