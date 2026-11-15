"""The reusable profile: permanent, confirmed settings only.

It holds no transactions, balances, outstanding items, or approvals. Those belong to each period's workbook.
"""
from __future__ import annotations

import copy, hashlib, json
from pathlib import Path

FORMAT = 1
SOURCE_KEYS = {"file_type", "sheet", "header_row", "columns", "date_format", "sign", "inflow", "type_map",
               "checks_by_numeric_reference", "summary_rows", "header_fingerprint"}


def load(path):
    try:
        return json.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, ValueError) as e:
        raise ValueError(f"Profile cannot be read: {e}")


def save(profile, path):
    Path(path).write_text(json.dumps(profile, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")


def digest(profile):
    return hashlib.sha256(json.dumps(profile, sort_keys=True).encode()).hexdigest()


def validate(p):
    """Return a list of problems. An empty list means the profile can be used."""
    errs = []
    if p.get("profile_format") != FORMAT:
        errs.append(f"profile_format must be {FORMAT}")
    if not p.get("confirmed"):
        errs.append("confirmed must be true: the user has not confirmed this setup")
    acct = p.get("account", {})
    for k in ("nickname", "currency"):
        if not acct.get(k):
            errs.append(f"account.{k} is required")
    last4 = str(acct.get("last4", ""))
    if last4 and (not last4.isdigit() or len(last4) > 4):
        errs.append("account.last4 may hold at most the last four digits, never a full account number")
    if not str(p.get("profile_version", "")).isdigit():
        errs.append("profile_version must be a whole number")
    for side in ("bank", "ledger"):
        s = p.get("sources", {}).get(side)
        if not s:
            errs.append(f"sources.{side} is required")
            continue
        check_source(side, s, errs)
    h = p.get("sources", {}).get("history")
    if h:  # optional: a coded-transactions history, used only for naming and coding suggestions
        check_source("history", h, errs)
        if not h.get("columns", {}).get("account"):
            errs.append("sources.history.columns.account is required (the account or category each line was coded to)")
    for i, rule in enumerate(p.get("coding_rules", [])):
        errs += [f"coding_rules[{i}]: {e}" for e in rule_problems(rule)]
    r = p.get("rules", {})
    for k in ("date_window_days", "deposit_in_transit_days"):
        if not isinstance(r.get(k), int) or r[k] < 0:
            errs.append(f"rules.{k} must be a whole number of days")
    if not isinstance(r.get("fee_descriptions", []), list):
        errs.append("rules.fee_descriptions must be a list")
    if not p.get("accounts", {}).get("cash"):
        errs.append("accounts.cash (the ledger cash account name) is required")
    return errs


def check_source(side, s, errs):
    unknown = set(s) - SOURCE_KEYS
    if unknown:
        errs.append(f"sources.{side}: unknown keys {sorted(unknown)}")
    c = s.get("columns", {})
    for k in ("date", "description"):
        if not c.get(k):
            errs.append(f"sources.{side}.columns.{k} is required")
    if c.get("amount"):
        if s.get("sign") not in ("as_is", "reverse"):
            errs.append(f"sources.{side}.sign must be as_is or reverse when one amount column is used")
    elif c.get("debit") and c.get("credit"):
        if s.get("inflow") not in ("debit", "credit"):
            errs.append(f"sources.{side}.inflow must say which column (debit or credit) increases cash")
    else:
        errs.append(f"sources.{side}.columns needs amount, or both debit and credit")
    if not s.get("date_format"):
        errs.append(f"sources.{side}.date_format is required (excel, or a format such as %d/%m/%Y)")
    if not s.get("header_fingerprint"):
        errs.append(f"sources.{side}.header_fingerprint is required (the confirmed column names, in order)")
    if s.get("summary_rows") not in (None, "exclude", "none"):
        errs.append(f"sources.{side}.summary_rows must be exclude or none")


RULE_KEYS = {"id", "counterparty", "words", "amount", "direction", "account", "description", "slash_means", "basis",
             "saved_on", "saved_by", "status", "retired_on", "note"}


def rule_problems(rule):
    """A saved coding rule must say when it applies (more than the supplier alone) and what it suggests."""
    out = []
    unknown = set(rule) - RULE_KEYS
    if unknown:
        out.append(f"unknown keys {sorted(unknown)}")
    for k in ("id", "counterparty", "basis", "saved_on", "saved_by"):
        if not rule.get(k):
            out.append(f"{k} is required")
    if not (rule.get("words") or rule.get("amount")):
        out.append("say when it applies beyond the supplier: purpose words on the line, or an exact amount")
    if not (rule.get("account") or rule.get("description") or rule.get("slash_means")):
        out.append("it must suggest an account, a description, or what n/m means on these lines")
    if rule.get("slash_means") not in (None, "installment"):
        out.append("slash_means can only be installment")
    if rule.get("direction") not in (None, "in", "out"):
        out.append("direction must be in or out")
    if rule.get("status", "active") not in ("active", "retired"):
        out.append("status must be active or retired")
    return out


def describe(p):
    """Plain-language lines for reading the setup back to the user."""
    a, r = p.get("account", {}), p.get("rules", {})
    out = [f"Account: {a.get('nickname')}" + (f" (ending {a['last4']})" if a.get("last4") else "") + f", {a.get('currency')}",
           f"Accounting system: {a.get('accounting_system') or 'not stated'}"
           + (f"; reconciled {a['frequency']}" if a.get("frequency") else "")
           + (f"; approved by {p['process']['approver']}" if p.get("process", {}).get("approver") else "")]
    for side in ("bank", "ledger"):
        s = p.get("sources", {}).get(side, {})
        c = s.get("columns", {})
        amount = (f"amount from '{c['amount']}'" + (" with signs reversed" if s.get("sign") == "reverse" else ""))\
            if c.get("amount") else (f"'{c.get(s.get('inflow'))}' increases cash and "
                                     f"'{c.get('debit' if s.get('inflow') == 'credit' else 'credit')}' decreases it")
        out.append(f"{side.title()} file: {s.get('file_type', '?')}, sheet {s.get('sheet') or 'first'}, date from '{c.get('date')}' "
                   f"({s.get('date_format')}), description from '{c.get('description')}', reference from "
                   f"'{c.get('reference') or 'none'}', {amount}" + (f", type from '{c['type']}'" if c.get("type") else "")
                   + (", summary rows excluded" if s.get("summary_rows") == "exclude" else ""))
    out.append(f"Matching: exact amounts; dates within {r.get('date_window_days')} days; deposits in transit from the last "
               f"{r.get('deposit_in_transit_days')} days; fee descriptions {r.get('fee_descriptions', [])}")
    acc = p.get("accounts", {})
    out.append(f"Accounts: cash '{acc.get('cash')}', bank fees '{acc.get('bank_fees', 'not set')}'")
    h = p.get("sources", {}).get("history")
    if h:
        c = h.get("columns", {})
        out.append(f"Coded history (suggestions only): {h.get('file_type', '?')}, description from '{c.get('description')}', "
                   f"account from '{c.get('account')}'")
    for rule in p.get("coding_rules", []):
        when = [rule.get("counterparty", "?")] + [f"'{w}'" for w in rule.get("words") or []] + ([f"amount {rule['amount']}"] if rule.get("amount") else [])
        out.append(f"Saved coding rule {rule.get('id')} ({rule.get('status', 'active')}, saved {rule.get('saved_on')} by {rule.get('saved_by')}): "
                   f"when {' + '.join(when)}; basis: {rule.get('basis')}")
    return out


def diff(old, new, prefix=""):
    """Paths whose values differ between two profiles."""
    out = []
    keys = sorted(set(old) | set(new)) if isinstance(old, dict) and isinstance(new, dict) else []
    if not keys:
        return [] if old == new else [f"{prefix or '(root)'}: {json.dumps(old)} -> {json.dumps(new)}"]
    for k in keys:
        if k == "change_log":
            continue
        out += diff(old.get(k), new.get(k), f"{prefix}.{k}" if prefix else k) if isinstance(old.get(k), dict) and isinstance(new.get(k), dict) \
            else ([] if old.get(k) == new.get(k) else [f"{prefix + '.' if prefix else ''}{k}: {json.dumps(old.get(k))} -> {json.dumps(new.get(k))}"])
    return out


def _esc(k):
    return str(k).replace("\\", "\\\\").replace(".", "\\.")


def _split(path):
    out, cur, i = [], "", 0
    while i < len(path):
        ch = path[i]
        if ch == "\\" and i + 1 < len(path):
            cur += path[i + 1]
            i += 2
            continue
        if ch == ".":
            out.append(cur)
            cur = ""
        else:
            cur += ch
        i += 1
    out.append(cur)
    return out


def flatten(obj, prefix=""):
    """Rows (setting path, value, type) so the setup can live in the workbook as plain cells.

    Paths use dots between keys and [n] for list positions; a dot inside a key is escaped. Empty lists and
    dicts, and nulls, keep their type so the setup reads back exactly.
    """
    rows = []
    if isinstance(obj, dict):
        if not obj and prefix:
            rows.append((prefix, None, "dict"))
        for k, v in obj.items():
            rows += flatten(v, f"{prefix}.{_esc(k)}" if prefix else _esc(k))
    elif isinstance(obj, list):
        if not obj:
            rows.append((prefix, None, "list"))
        for i, v in enumerate(obj):
            rows += flatten(v, f"{prefix}.[{i}]")
    elif obj is None:
        rows.append((prefix, None, "null"))
    elif isinstance(obj, bool):
        rows.append((prefix, "true" if obj else "false", "bool"))
    elif isinstance(obj, int):
        rows.append((prefix, obj, "int"))
    elif isinstance(obj, float):
        rows.append((prefix, obj, "float"))
    else:
        rows.append((prefix, str(obj), "text"))
    return rows


def unflatten(rows):
    """The inverse of flatten: rows of (path, value, type) back to the setup."""
    root = {}

    def container(seg_next):
        return [] if seg_next.startswith("[") and seg_next.endswith("]") else {}

    for path, value, kind in rows:
        segs = _split(str(path))
        node = root
        for j, seg in enumerate(segs):
            last = j == len(segs) - 1
            key = int(seg[1:-1]) if seg.startswith("[") and seg.endswith("]") and seg[1:-1].isdigit() else seg
            if last:
                v = ({"null": None, "dict": {}, "list": []}.get(kind) if kind in ("null", "dict", "list") else
                     (str(value).lower() == "true") if kind == "bool" else int(value) if kind == "int" else
                     float(value) if kind == "float" else ("" if value is None else str(value)))
            else:
                v = container(segs[j + 1])
            if isinstance(node, list):
                while len(node) <= key:
                    node.append(None)
                if last or node[key] is None:
                    node[key] = v
                node = node[key]
            else:
                if last or key not in node:
                    node[key] = v
                node = node[key]
    return root


def bumped(old, new, change, on):
    """Return a copy of new with the version raised and the change logged, if anything differs."""
    if not diff(old, new):
        return copy.deepcopy(old)
    out = copy.deepcopy(new)
    out["profile_version"] = int(old.get("profile_version", 1)) + 1
    out["change_log"] = list(old.get("change_log", [])) + [{"version": out["profile_version"], "date": on, "change": change}]
    return out
