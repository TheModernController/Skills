"""Read Excel and CSV exports, inspect their layout, and apply a confirmed column mapping.

Every row read keeps its file, sheet, and row number. Nothing is dropped silently: a row that is not
a transaction is returned with the reason it was excluded, and blank rows are counted.
All Excel access is in read_grid, so changing the Excel library changes only that function.
"""
from __future__ import annotations

import csv, hashlib, io, re
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

CENT = Decimal("0.01")
SUMMARY_START = ("total", "totals", "opening balance", "closing balance", "balance brought forward",
                 "balance carried forward", "brought forward", "carried forward", "balance b/f", "balance c/f")
DATE_FORMATS = ["%Y-%m-%d", "%Y/%m/%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y", "%m-%d-%Y", "%d.%m.%Y",
                "%d/%m/%y", "%m/%d/%y", "%d %b %Y", "%d-%b-%Y", "%b %d, %Y", "%d %B %Y", "%B %d, %Y"]
ROLE_WORDS = {
    "id": ["line id", "transaction id", "id", "line_id"],
    "date": ["date", "posted", "posting date", "transaction date", "value date"],
    "reference": ["reference", "ref", "num", "number", "check", "cheque", "doc", "document"],
    "description": ["description", "details", "memo", "narrative", "payee", "party", "name"],
    "amount": ["amount", "value", "net"],
    "debit": ["debit", "withdrawal", "withdrawals", "paid out", "money out", "dr"],
    "credit": ["credit", "deposit", "deposits", "paid in", "money in", "cr"],
    "balance": ["balance", "running balance"],
    "type": ["type", "transaction type", "kind"],
    "currency": ["currency", "ccy"],
    "account": ["account", "gl account", "category", "split", "nominal", "coding", "account name"],
}


class InputError(Exception):
    """A file cannot be read as confirmed, so the run stops instead of guessing."""


class LayoutError(InputError):
    """The file's columns no longer contain every mapped column."""

    def __init__(self, source, missing, new, header):
        self.source, self.missing, self.new, self.header = source, missing, new, header
        super().__init__(f"{source} file layout changed: mapped column(s) not found: {', '.join(missing)}."
                         + (f" New column(s): {', '.join(new)}." if new else ""))


def sha256(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def text(v):
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.date().isoformat() if not (v.hour or v.minute or v.second) else v.isoformat(sep=" ")
    if isinstance(v, date):
        return v.isoformat()
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def read_grid(path, sheet=None):
    """Return (sheet name or None, rows as lists of cell values, all sheet names)."""
    p = Path(path)
    if not p.is_file():
        raise InputError(f"File not found: {p.name}")
    ext = p.suffix.lower()
    if ext in (".xlsx", ".xlsm"):
        from openpyxl import load_workbook  # the only Excel dependency for reading
        wb = load_workbook(p, read_only=True, data_only=True)
        names = list(wb.sheetnames)
        name = sheet or names[0]
        if name not in names:
            raise InputError(f"{p.name}: sheet {name!r} not found. Sheets: {', '.join(names)}")
        rows = [list(r) for r in wb[name].iter_rows(values_only=True)]
        wb.close()
        return name, rows, names
    if ext in (".csv", ".tsv", ".txt"):
        raw = p.read_bytes()
        for enc in ("utf-8-sig", "cp1252"):
            try:
                content = raw.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        if ext == ".tsv":
            delim = "\t"
        else:
            try:
                delim = csv.Sniffer().sniff(content[:4096], delimiters=",;\t|").delimiter
            except csv.Error:
                delim = ","
        return None, [list(r) for r in csv.reader(io.StringIO(content), delimiter=delim)], []
    if ext == ".xls":
        raise InputError(f"{p.name}: older .xls files are not supported. Save it as .xlsx or CSV.")
    raise InputError(f"{p.name}: only Excel (.xlsx) and CSV files are supported.")


# ---------------------------------------------------------------- values

def parse_amount(v):
    """Return (Decimal or None, problem text or '')."""
    if v is None or (isinstance(v, str) and v.strip() == ""):
        return None, ""
    if isinstance(v, bool):
        return None, f"not an amount: {v!r}"
    if isinstance(v, (int, float, Decimal)):
        d = Decimal(repr(v)) if isinstance(v, float) else Decimal(v)
        if abs(d - d.quantize(CENT)) < Decimal("0.000001"):
            d = d.quantize(CENT)
    else:
        s, neg = str(v).strip(), False
        if s.startswith("(") and s.endswith(")"):
            neg, s = True, s[1:-1]
        s = re.sub(r"[\s,$£€¥]", "", s)
        if s.endswith("-"):
            neg, s = not neg, s[:-1]
        try:
            d = Decimal(s)
        except InvalidOperation:
            return None, f"not an amount: {v!r}"
        if neg:
            d = -d
    if d != d.quantize(CENT):
        return None, f"more than two decimals: {v!r}"
    return d.quantize(CENT), ""


def parse_date(v, fmt):
    """Return a date or None. Excel date cells are read directly; text uses the confirmed format."""
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = text(v)
    if not s or not fmt or fmt == "excel":
        return None
    for candidate in (s, s.split(" ")[0], s.split("T")[0]):
        try:
            return datetime.strptime(candidate, fmt).date()
        except ValueError:
            continue
    return None


def is_summary_text(s):
    s = s.casefold().strip()
    return any(s == w or s.startswith(w + " ") or s.startswith(w + ":") for w in SUMMARY_START)


# ---------------------------------------------------------------- inspection

def guess_role(name):
    n = name.casefold().strip().replace("_", " ")
    for role in ("id", "balance", "debit", "credit", "date", "reference", "amount", "type", "currency", "account", "description"):
        if any(n == w or n.startswith(w + " ") or n.endswith(" " + w) for w in ROLE_WORDS[role]):
            return role
    return ""


def find_header_row(rows, limit=30):
    """The first row with at least three text cells followed by rows holding dates or numbers."""
    for i, r in enumerate(rows[:limit]):
        cells = [c for c in r if text(c)]
        texts = [c for c in cells if isinstance(c, str) and parse_amount(c)[0] is None]
        if len(texts) < 3:
            continue
        below = rows[i + 1:i + 6]
        if any(any(isinstance(c, (int, float, date)) or parse_amount(c)[0] is not None for c in b if text(c)) for b in below):
            return i
    return 0


def date_evidence(values):
    cells = [v for v in values if text(v)]
    if not cells:
        return {"kind": "empty"}
    real = sum(isinstance(v, (date, datetime)) for v in cells)
    if real == len(cells):
        return {"kind": "excel date cells", "format": "excel"}
    strings = [text(v) for v in cells if not isinstance(v, (date, datetime))]
    fits = [f for f in DATE_FORMATS if all(parse_date(s, f) for s in strings)]
    result = {"kind": "text dates" if not real else "mixed date cells and text", "formats_that_fit": fits}
    if any("%d/%m" in f for f in fits) and any("%m/%d" in f for f in fits):
        result["ambiguous"] = "Every day and month is 12 or under, so day-first and month-first both fit. Ask."
    return result


FEE_LIKE = re.compile(r"\b(FEES?|CHARGES?)\b", re.I)
RETURN_LIKE = re.compile(r"\bRETURN(ED|S)?\b", re.I)


def inspect_file(path, sheet=None):
    """Describe a file so mappings can be proposed. Reads only; decides nothing."""
    name, rows, sheets = read_grid(path, sheet)
    h = find_header_row(rows)
    header = [text(c) for c in rows[h]] if rows else []
    everything = rows[h + 1:]
    summary_rows, blank, body = [], 0, []
    for i, r in enumerate(everything):
        cells = [text(c) for c in r]
        if not any(cells):
            blank += 1
        elif any(is_summary_text(c) for c in cells if c):
            summary_rows.append({"row": h + i + 2, "cells": [c for c in cells if c]})
        else:
            body.append(r)
    width = max([len(header)] + [len(r) for r in everything]) if rows else 0
    cols = []
    for j in range(width):
        values = [r[j] if j < len(r) else None for r in body]
        filled = [v for v in values if text(v)]
        amounts = [parse_amount(v)[0] for v in filled]
        nums = [a for a in amounts if a is not None]
        col = {"column": header[j] if j < len(header) else f"(column {j + 1})", "role_guess": guess_role(header[j]) if j < len(header) else "",
               "filled": len(filled), "samples": [text(v) for v in filled[:3]]}
        if nums and len(nums) >= len(filled) * 0.8:
            col["numbers"] = {"count": len(nums), "sum": str(sum(nums)), "negatives": sum(n < 0 for n in nums),
                              "positives": sum(n > 0 for n in nums)}
        dates = date_evidence(filled)
        if dates.get("kind") == "excel date cells" or dates.get("formats_that_fit"):
            col["dates"] = dates
            ds = [parse_date(v, dates.get("format") or (dates["formats_that_fit"] or [None])[0]) for v in filled]
            ds = [d for d in ds if d]
            if ds:
                col["dates"]["range"] = [str(min(ds)), str(max(ds))]
                col["dates"]["order"] = "oldest first" if ds == sorted(ds) else "newest first" if ds == sorted(ds, reverse=True) else "mixed"
        if 1 < len(set(text(v) for v in filled)) <= 8 and not nums:
            col["distinct_values"] = sorted(set(text(v) for v in filled))
        if not nums and col["role_guess"] in ("description", ""):
            for key, pattern in (("descriptions_that_look_like_fees", FEE_LIKE), ("descriptions_that_look_like_returned_items", RETURN_LIKE)):
                found = sorted({text(v) for v in filled if pattern.search(text(v))})
                if found:
                    col[key] = found
        cols.append(col)
    return {
        "file": Path(path).name, "sha256": sha256(path), "sheet": name, "sheets": sheets,
        "header_row": h + 1, "rows_above_header": [[text(c) for c in r if text(c)] for r in rows[:h] if any(text(c) for c in r)],
        "data_rows": len(body), "blank_rows": blank, "columns": cols,
        "sample_rows": [{(header[j] if j < len(header) and header[j] else f"(column {j + 1})"): text(v) for j, v in enumerate(r) if text(v)}
                        for r in body[:3]],
        "column_statistics": ("Calculated without the possible summary rows listed below" if summary_rows
                              else "Calculated from every non-blank row below the header"),
        "possible_summary_rows": summary_rows,
        "note": "Summary rows (totals, opening or closing balances) are excluded only after the user confirms. Their values can be checked against the rows read.",
    }


# ---------------------------------------------------------------- mapping

@dataclass
class Row:
    line_id: str
    source: str
    file: str
    sheet: str
    row: int
    date: date | None
    reference: str
    description: str
    amount: Decimal | None
    kind: str = ""
    type_text: str = ""
    balance: Decimal | None = None
    included: bool = True
    reason: str = ""
    problems: list = field(default_factory=list)
    notes: list = field(default_factory=list)
    disposition: str = ""
    rule: str = ""
    link: str = ""

    def content(self):
        return (self.source, str(self.date), self.reference, self.description, str(self.amount), self.kind)


@dataclass
class SourceFile:
    source: str
    path: str
    name: str
    sha: str
    sheet: str
    header_row: int
    header: list
    rows: list
    blank_rows: int
    layout_notes: list
    checks: list  # (label, expected, actual)
    unmapped_types: list


def locate_header(rows, needed, fingerprint, source):
    want = {n.casefold() for n in needed}
    for i, r in enumerate(rows[:30]):
        have = {text(c).casefold() for c in r if text(c)}
        if want <= have:
            return i
    best = max(range(min(30, len(rows))), key=lambda i: len({text(c).casefold() for c in rows[i]} & {f.casefold() for f in fingerprint}), default=0)
    header = [text(c) for c in rows[best]] if rows else []
    have = {c.casefold() for c in header if c}
    missing = [n for n in needed if n.casefold() not in have]
    new = [c for c in header if c and c.casefold() not in {f.casefold() for f in fingerprint}]
    raise LayoutError(source, missing, new, [c for c in header if c])


def read_source(path, cfg, source, prefix, period_start=None, period_end=None, currency=None):
    """Apply a confirmed mapping. Returns a SourceFile whose rows include excluded rows with reasons."""
    cols = cfg["columns"]
    name, rows, _ = read_grid(path, cfg.get("sheet"))
    needed = [v for v in cols.values() if v]
    fingerprint = cfg.get("header_fingerprint", [])
    h = locate_header(rows, needed, fingerprint, source)
    header = [text(c) for c in rows[h]]
    pos = {c.casefold(): j for j, c in enumerate(header) if c}
    idx = {role: pos[col.casefold()] for role, col in cols.items() if col}
    notes = []
    current = [c for c in header if c]
    if fingerprint and current != fingerprint:
        added = [c for c in current if c not in fingerprint]
        gone = [c for c in fingerprint if c not in current]
        if added or gone:
            notes.append("Unmapped column changes: " + "; ".join(filter(None, [
                f"new {', '.join(added)}" if added else "", f"no longer present {', '.join(gone)}" if gone else ""])))
        else:
            notes.append("Columns are in a different order; mapping is by name, so results are unaffected.")
    if cfg.get("header_row") and cfg["header_row"] != h + 1:
        notes.append(f"Header is now on row {h + 1} (was {cfg['header_row']}).")

    def cell(r, role):
        j = idx.get(role)
        return r[j] if j is not None and j < len(r) else None

    fmt = cfg.get("date_format")
    type_map = {k.casefold(): v for k, v in (cfg.get("type_map") or {}).items()}
    out, ids, blank, unmapped = [], set(), 0, set()
    col_sums = {role: Decimal("0.00") for role in ("amount", "debit", "credit") if role in idx}
    totals = []
    for i, r in enumerate(rows[h + 1:]):
        rownum = h + i + 2
        cells = [text(c) for c in r]
        if not any(cells):
            blank += 1
            continue
        raw_date = cell(r, "date")
        when = parse_date(raw_date, fmt)
        desc = text(cell(r, "description"))
        line_id = text(cell(r, "id")) if "id" in idx else f"{prefix}-R{rownum:04d}"
        rec = Row(line_id, source, Path(path).name, name or "", rownum, when, text(cell(r, "reference")).upper(), desc, None)
        summary = (is_summary_text(desc) and desc.casefold().strip() in {w for w in SUMMARY_START}) or \
                  (when is None and any(is_summary_text(c) for c in cells if c))
        if summary:
            if cfg.get("summary_rows") != "exclude":
                raise InputError(f"{Path(path).name} row {rownum} looks like a summary row ({', '.join(c for c in cells if c)}). "
                                 "Ask the user, then set summary_rows to \"exclude\" in the profile if confirmed.")
            rec.included, rec.reason = False, "Summary row, not a transaction (confirmed in setup)"
            vals = {role: parse_amount(cell(r, role))[0] for role in ("amount", "debit", "credit", "balance") if role in idx}
            rec.notes.append("; ".join(f"{k} {v}" for k, v in vals.items() if v is not None))
            if any(c.casefold().startswith("total") for c in cells if c):
                totals.append((rownum, vals))
            out.append(rec)
            continue
        if not line_id:
            line_id = rec.line_id = f"{prefix}-R{rownum:04d}"
        if line_id in ids:
            raise InputError(f"{Path(path).name} row {rownum}: line id {line_id} appears twice")
        ids.add(line_id)
        # amount
        if "amount" in idx:
            amt, prob = parse_amount(cell(r, "amount"))
            if amt is not None:
                col_sums["amount"] += amt
                if cfg.get("sign") == "reverse":
                    amt = -amt
        else:
            dr, p1 = parse_amount(cell(r, "debit"))
            cr, p2 = parse_amount(cell(r, "credit"))
            prob = p1 or p2
            if dr is not None:
                col_sums["debit"] += dr
            if cr is not None:
                col_sums["credit"] += cr
            amt = None
            if not prob:
                if dr and cr:
                    prob = "both debit and credit are filled"
                elif dr or cr:
                    inflow_is_credit = cfg.get("inflow") == "credit"
                    if dr:
                        amt = -abs(dr) if inflow_is_credit else abs(dr)
                    else:
                        amt = abs(cr) if inflow_is_credit else -abs(cr)
        if prob:
            rec.problems.append(f"Amount unreadable ({prob})")
        elif amt is None or amt == 0:
            rec.problems.append("Amount missing or zero")
            amt = None
        rec.amount = amt
        if when is None:
            rec.problems.append(f"Date missing or unreadable ({text(raw_date)!r})" if text(raw_date) else "Date missing")
        if "balance" in idx:
            rec.balance = parse_amount(cell(r, "balance"))[0]
        if "currency" in idx and currency:
            ccy = text(cell(r, "currency")).upper()
            if ccy and ccy != currency.upper():
                raise InputError(f"{Path(path).name} row {rownum}: currency {ccy} is not {currency}. Version one reconciles one currency per run.")
        t = text(cell(r, "type"))
        rec.type_text = t
        mapped = type_map.get(t.casefold()) if t else None
        if t and mapped is None:
            unmapped.add(t)
        sign = (rec.amount or 0)
        if mapped in ("check", "payment", "receipt"):
            rec.kind = mapped
        elif cfg.get("checks_by_numeric_reference") and sign < 0 and rec.reference.isdigit():
            rec.kind = "check"
        else:
            rec.kind = "receipt" if sign > 0 else "payment" if sign < 0 else ""
        if when and period_start and period_end and not period_start <= when <= period_end:
            rec.included, rec.reason = False, f"Dated outside the period ({when})"
        out.append(rec)

    checks = []
    for rownum, vals in totals:
        for role, total in vals.items():
            if role in col_sums and total is not None:
                checks.append((f"{Path(path).name}: totals row {rownum} {role} agrees to the rows read", str(total), str(col_sums[role])))
    seen = {}
    for rec in out:
        if rec.included and rec.amount is not None:
            key = rec.content()
            if key in seen:
                rec.notes.append(f"Identical to row {seen[key].row}; kept, since identical transactions can be genuine")
                seen[key].notes.append(f"Identical to row {rec.row}; kept, since identical transactions can be genuine")
            else:
                seen[key] = rec
    return SourceFile(source, str(path), Path(path).name, sha256(path), name or "", h + 1, current, out, blank, notes, checks, sorted(unmapped))


def balance_continuity(rows):
    """Check a running balance column. Returns dict with order, breaks, implied opening and closing, or None."""
    rs = [r for r in rows if r.included and r.balance is not None and r.amount is not None]
    if len(rs) < 2:
        return None
    fwd = [i for i in range(1, len(rs)) if rs[i].balance != rs[i - 1].balance + rs[i].amount]
    rev = [i for i in range(len(rs) - 1) if rs[i].balance != rs[i + 1].balance + rs[i].amount]
    if len(fwd) <= len(rev):
        chron, breaks = rs, [rs[i].row for i in fwd]
    else:
        chron, breaks = list(reversed(rs)), [rs[i].row for i in rev]
    return {"order": "oldest first" if chron is rs else "newest first", "breaks_at_rows": breaks,
            "implied_opening": chron[0].balance - chron[0].amount, "implied_closing": chron[-1].balance}
