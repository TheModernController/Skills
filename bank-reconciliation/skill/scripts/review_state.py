"""The review workbook's layout constants, and reading what Claude stored in it.

The workbook is the working surface: the accountant reviews and edits it in Excel and uploads it back. The hidden
Setup and Checks tabs keep everything needed to validate an upload and to start the next period, so no separate
setup file is used.
"""
from __future__ import annotations

from datetime import datetime, timedelta
from decimal import Decimal
from pathlib import Path

from foundation import profiles
from foundation.sections import read_kv, read_section
from foundation.tables import InputError

# tabs
REC, JE, BANK, LEDGER, GUIDE, ACCOUNTS, SETUP, CHECKS = ("Reconciliation", "Journal entries", "Bank raw", "Ledger raw", "Guide",
                                                         "Accounts", "Setup", "Checks")
TABS = [REC, JE, BANK, LEDGER, GUIDE, ACCOUNTS, SETUP, CHECKS]

# Reconciliation layout (the user's saved copy of revision 3)
HEADER_ROWS = {"Company": 2, "Account": 3, "Period": 4, "Reconciled by": 5, "Review completed on": 6, "Checks": 7}
REVIEW_DATE_ROW = HEADER_ROWS["Review completed on"]  # the reviewer's own date; review-1 (0.5.0) wrote the completion text here
SUMMARY_HEAD, OPEN_ROW, IN_ROW, OUT_ROW, CLOSE_ROW, ADJ_ROW, ADJUSTED_ROW, REMAINING_ROW = 9, 10, 11, 12, 13, 14, 15, 17
LINE_ROW, LEGEND_ROW, TABLE_HEAD, FIRST_ROW = 18, 19, 20, 21
REC_COLS = ["Ref", "Bank date", "Ledger date", "Bank description", "Ledger description", "Bank amount", "Ledger amount",
            "Approval", "JE ref"]  # columns B..J
LAST_REC_ROW = 2000

# Journal entries layout
JE_HEAD, JE_FIRST = 3, 5
JE_COLS = {"Ref": "B", "Date": "C", "Account": "D", "Description": "E", "Debit": "F", "Credit": "G",
           "Proposal": "I", "Approval status": "J", "JE number": "K", "Check": "L", "Comments": "M"}

# Checks tab: fixed rows for the live checks (Reconciliation shows the first failure)
LIVE_FIRST, LIVE_LAST = 4, 9

# section titles
S_EFFECTS = "Effects of each choice"
S_OPTIONS = "Choices offered"
S_EDITS = "Edits found at the last upload"
S_DECISIONS = "Decisions recorded"
S_CARRY = "Carry forward to next period"
S_CARRY_FACTS = "Carry forward facts"
S_LINEMAP = "Where each line went"
S_REFS = "Refs"
S_LINES = "Included lines"
S_PROPOSED = "Entries as proposed"
S_REC_ROWS = "Reconciliation rows as written"
S_CARRIED_FACTS = "Carried in: facts"
S_CARRIED_ROWS = "Carried in: items"
S_CARRIED_LINES = "Carried in: previous period's lines"
S_CARRIED_HISTORY = "Carried in: approved descriptions"
S_SNAP = {"bank": "Source rows as received: bank", "ledger": "Source rows as received: ledger"}
S_RUN = "Run record"
S_EXPECT = "Formula expectations"
S_SETTINGS = "Settings"
S_RULES = "Rules"
S_TECH = "Technical settings (read by Claude)"

CARRY_COLS = ["Kind", "Item ID", "Origin period", "Date", "Reference", "Description", "Amount", "Category", "Account", "Posting ref",
              "Approved by", "Note"]


# review-1: skill 0.5.0. review-2: 0.5.1 (one choice per item, and the reviewer's "Review completed on" date).
# Both keep the same carry forward, so either can start the next period; validating a review-1 upload needs --from-0.5.0.
FORMAT = "review-2"
FORMATS = ("review-1", "review-2")


def workbook_format(wb):
    if CHECKS in wb.sheetnames and SETUP in wb.sheetnames and REC in wb.sheetnames:
        return read_kv(wb[CHECKS], S_RUN).get("Workbook format")
    return None


def is_review_workbook(wb):
    return workbook_format(wb) in FORMATS


def setup_of(wb):
    """The confirmed setup (profile) kept on the hidden Setup tab."""
    rows = read_section(wb[SETUP], S_TECH)
    if not rows:
        raise InputError("This workbook's Setup tab has no technical settings. Use a workbook this skill produced.")
    return profiles.unflatten([(r["Setting"], r.get("Value"), r.get("Type")) for r in rows])


def run_record(wb):
    return read_kv(wb[CHECKS], S_RUN)


def accounts_of(wb):
    ws = wb[ACCOUNTS]
    out = []
    for r in range(2, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if v not in (None, "") and str(v).strip():
            out.append(str(v).strip())
    return out


def D(v):
    if v in (None, ""):
        return None
    return Decimal(str(v)).quantize(Decimal("0.01"))


def _date(v):
    if v in (None, ""):
        return None
    if isinstance(v, datetime):
        return v.date()
    return datetime.strptime(str(v)[:10], "%Y-%m-%d").date()


def read_prior(wb, path, period, profile, embedded=False):
    """The previous period as bank_rules expects it, from a checked workbook of this version.

    embedded=False: wb is last period's checked workbook (starting a next period).
    embedded=True:  wb is this period's own workbook, which keeps what it carried in (validating an upload).
    """
    from bank_rules import carry_digest
    ck = wb[CHECKS]
    if embedded:
        facts = read_kv(ck, S_CARRIED_FACTS)
        if not facts:
            return None
        rows = [{c: r.get(c) for c in CARRY_COLS} for r in read_section(ck, S_CARRIED_ROWS)]
        lines_src = read_section(ck, S_CARRIED_LINES)
        hist_src = read_section(ck, S_CARRIED_HISTORY)
        files = {facts.get("Bank file SHA-256"), facts.get("Ledger file SHA-256")} - {None}
        edited = False
    else:
        record = read_kv(ck, S_RUN)
        facts = read_kv(ck, S_CARRY_FACTS)
        rows = [{c: r.get(c) for c in CARRY_COLS} for r in read_section(ck, S_CARRY)]
        if not facts:
            raise InputError(f"{Path(path).name} has no carry forward yet. Upload the checked workbook Claude returned after "
                             "validating last period, not the first draft.")
        completion = str(read_kv(ck, S_RUN).get("Completion") or "")
        if not completion.startswith("Complete"):
            raise InputError(f"Last period's workbook says '{completion}'. Finish and validate that reconciliation first, then "
                             "upload the checked workbook Claude returned.")
        edited = carry_digest(facts, rows) != record.get("Carry forward digest")
        if edited and not period.get("accept_edited_prior"):
            raise InputError("The carry forward kept in last period's workbook was changed after Claude checked it. Ask the user "
                             "whether that was intended; if so, set accept_edited_prior to true in the period file.")
        lines_src = read_section(ck, S_LINES)
        hist_src = [r for r in read_section(ck, S_DECISIONS) if r.get("Bank description")]
        files = {record.get("Bank file SHA-256"), record.get("Ledger file SHA-256")} - {None}
    ps, pe = _date(facts.get("Period start")), _date(facts.get("Period end"))
    if ps == period["period_start"] and pe == period["period_end"]:
        raise InputError("The previous-period workbook given is for this same period. Use the workbook for the period immediately before.")
    if pe + timedelta(days=1) != period["period_start"]:
        raise InputError(f"The previous-period workbook ends {pe}, but this period starts {period['period_start']}. "
                         "Use the workbook for the period immediately before.")
    acct = profile["account"]
    if facts.get("Account") != acct["nickname"] or facts.get("Currency") != acct["currency"]:
        raise InputError("The previous-period workbook is for a different account or currency than this setup.")
    lines = {}
    for r in lines_src:
        sheet = "Bank rows" if r.get("Side") == "Bank" else "Ledger rows"
        d = _date(r.get("Date"))
        lines[(sheet, str(d), str(r.get("Reference") or ""), str(r.get("Description") or ""), str(D(r.get("Amount"))))] = r.get("Line ID")
    history = {}
    for r in hist_src:
        if r.get("Bank description"):
            history[str(r["Bank description"]).upper()] = (r.get("Type") or "", r.get("Treatment") or "", r.get("Account") or "")
    return dict(facts=facts, rows=rows, edited=edited, lines=lines, history=history, period_label=pe.strftime("%Y-%m"),
                files=files, statement_closing=D(facts.get("Stated bank closing balance")),
                ledger_closing=D(facts.get("Actual ledger closing balance")),
                raw_lines=lines_src, raw_history=hist_src)
