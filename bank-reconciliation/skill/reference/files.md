# File formats

## Setup (permanent settings, kept in the workbook)

Claude writes the setup as JSON in its working folder from the user's answers, validates it with `profile`, and the workbook keeps it on the hidden Setup tab ("Technical settings"). The user keeps no separate file: next period, `carry` reads it back from the checked workbook. It holds confirmed settings only: no transactions, balances, outstanding items, or approvals. Columns are named exactly as in the file's header row.

**Required fields:**
- **Top level:**
  - `profile_format` (1).
  - `profile_version`.
  - `confirmed` (true once Claude has stated the interpretation and the user has answered every genuine ambiguity; no blanket confirmation is asked).
  - `rules.date_window_days` and `rules.deposit_in_transit_days`.
  - `accounts.cash`.
- **Account:** `account.nickname` and `account.currency`.
- **For each of `sources.bank` and `sources.ledger`:**
  - `columns.date` and `columns.description`.
  - Either `columns.amount` with `sign`, or both `columns.debit` and `columns.credit` with `inflow`.
  - `date_format`.
  - `header_fingerprint`.
- **Everything else is optional.** `bankrec.py profile` lists anything missing.

```json
{
  "profile_format": 1,
  "profile_version": 1,
  "confirmed": true,
  "account": {"company": "Example Ltd", "nickname": "Operating account", "last4": "4821", "currency": "USD",
              "accounting_system": "optional: only what the user volunteered (e.g. \"In-house ledger\"); leave it out otherwise, never inferred"},
  "process": {"reconciled_by": "the name the user gave"},
  "sources": {
    "bank": {
      "file_type": "xlsx", "sheet": "Account Activity", "header_row": 4,
      "columns": {"date": "Posting Date", "description": "Details", "reference": "Ref",
                  "debit": "Withdrawals", "credit": "Deposits", "balance": "Balance"},
      "date_format": "excel", "inflow": "credit", "summary_rows": "exclude",
      "header_fingerprint": ["Posting Date", "Details", "Ref", "Withdrawals", "Deposits", "Balance"]
    },
    "ledger": {
      "file_type": "csv", "sheet": null, "header_row": 5,
      "columns": {"date": "Date", "description": "Name", "reference": "Num", "debit": "Debit", "credit": "Credit", "type": "Type"},
      "date_format": "%d/%m/%Y", "inflow": "debit",
      "type_map": {"Deposit": "receipt", "Bill Payment": "payment", "Check": "check"},
      "summary_rows": "exclude",
      "header_fingerprint": ["Date", "Type", "Num", "Name", "Memo", "Debit", "Credit", "Balance"]
    }
  },
  "rules": {"date_window_days": 3, "deposit_in_transit_days": 3, "fee_descriptions": ["SERVICE CHARGE - MONTHLY"], "returned_item_words": ["RETURNED"]},
  "accounts": {"cash": "1010 Operating Cash", "bank_fees": "Bank fees"},
  "change_log": [{"version": 1, "date": "YYYY-MM-DD", "change": "Initial setup confirmed"}]
}
```

### Columns

- **`columns`:**
  - Required: `date`, `description`, and either `amount` or both `debit` and `credit`.
  - Optional: `reference`, `type`, `id` (a unique line id column), `balance` (running balance, bank only), and `currency`.
  - With no `id` column, line ids are made from the row number (`BK-R0007`, `CB-R0012`).

### Amount sign

- **`sign`** (with `amount`):
  - `as_is` when money in is positive.
  - `reverse` when money in is negative.
- **`inflow`** (with debit and credit columns): the role whose column increases cash.
  - **Bank exports:** usually `credit`, whatever the column is called.
  - **Ledger cash accounts:** usually `debit`.

### Dates

- **`date_format`:**
  - `excel` for real Excel date cells.
  - Otherwise a format such as `%Y-%m-%d`, `%d/%m/%Y`, or `%m/%d/%Y`.
  - Never guess day-first or month-first. If inspection says ambiguous, ask.

### Row handling

- **`summary_rows`:**
  - `exclude` only after the user confirms that totals, opening balance, and closing balance rows are not transactions.
  - `none` if there are none.
- **`type_map`:** maps the file's type values to `check`, `payment`, or `receipt`.
  - Values not in the map are classified by sign and reported.
  - `checks_by_numeric_reference: true` treats outflows with an all-digit reference as checks when there is no type column.

### Layout and version

- **`header_fingerprint`:** the confirmed column names in order. Next period, a missing mapped column stops the run and is asked about. Other column changes are reported only.
- **`last4`:** at most four digits. Never a full account number.
- **Changes:** for a later change, run `profile NEW --previous OLD --change "..." --write OUT`. It raises `profile_version` and appends to `change_log`.

### Coding suggestions (optional)

These settings are used only for naming and coding suggestions. They never affect matching or the result.

- **`sources.history`:** a coded-transactions history, mapped like the other sources. It needs `columns.account` (the account or category each line was coded to), plus `date`, `description`, and an amount. `columns.id` (for example `Ref`) is optional.
  - **Amounts:** read as the cash account sees them, with money out negative, as in the bank file. Suggestions compare direction with the bank line.
  - **Positive payments:** if the history shows payments as positive (an expense-side export), use `sign: reverse`, or `inflow` for debit and credit columns.
- **`sources.ledger.columns.account`:** set this instead, when the ledger export itself shows each line's account.
- **`coding_rules`:** reusable preferences the user chose to save. Each rule says when it applies and what it suggests:

```json
{"id": "CR-1", "counterparty": "Cedarline Leasing", "amount": "389.00", "slash_means": "installment",
 "account": "Vehicle lease expense", "basis": "BK-209, March 2026: the user confirmed 7/24 is payment 7 of 24",
 "saved_on": "2026-04-02", "saved_by": "Fictional reviewer", "status": "active"}
```

**When it applies:**
- `counterparty` is required.
- It also needs `words` (purpose words that must appear on the line) or an exact `amount`, because the supplier alone is not enough.
- `direction` (`in` or `out`) is optional.

**What it suggests:** at least one of:
- `account`.
- `description`, a template using `{invoice number}`, `{month}`, `{n} of {m}`, or `{purpose}`.
- `slash_means`: `installment`, meaning n/m on these lines is installment n of m.

`saved_on` is the date the user chose to save the rule.

**Status:**
- **Retiring a rule:** set `status` to `retired` and add `retired_on`.
- **Changing or retiring:** every change goes through `profile --previous --change --write`.
- **What a rule gives:** a suggestion only, never a decision. It doesn't establish a transaction's purpose. When history disagrees with it, the line shows a conflict and a question.

## Period file (one per run, written by Claude)

```json
{
  "profile": "/tmp/bankrec/profile.json",
  "bank_file": "/mnt/user-data/uploads/bank-april.csv",
  "ledger_file": "/mnt/user-data/uploads/ledger-april.csv",
  "period_start": "2026-04-01", "period_end": "2026-04-30",
  "statement_opening": "46517.10", "statement_closing": "46901.85",
  "ledger_opening": "45882.10", "ledger_closing": null,
  "chart_file": "/mnt/user-data/uploads/chart-of-accounts.csv",
  "prior_workbook": "/mnt/user-data/uploads/bankrec-operating-account-2026-03.xlsx",
  "accounts_from": null, "accounts_add": [],
  "opening_items": null,
  "posted_before_opening": [],
  "accept_edited_prior": false,
  "decisions": "decisions.json",
  "out_dir": "/mnt/user-data/outputs"
}
```

Relative paths are resolved from the period file's folder.

### Balances

- **The statement balances:** amounts as text, as the user states them or as the bank export shows them.
- **`ledger_closing`:** optional. If given, the ledger roll-forward is checked against it.

### Accounts (the Accounts tab and the account lists)

- **`chart_file`:** an uploaded chart of accounts (Excel or CSV). `chart` shows what it reads: a number or code column with a name column, or a single account column.
- **`accounts_from`:** last period's checked workbook, to reuse its Accounts tab (including accounts the user added at the bottom).
- **`accounts_add`:** new or changed account names the user gave in chat.
- **None of these:** the Accounts tab starts from the cash and bank-fee accounts. An account typed in the workbook that isn't on the list is asked about at validation.

### Previous period

- **`prior_workbook`:** last period's checked workbook from 0.5.0 or later, after validation set Completion to Complete. Its hidden Checks tab carries the outstanding items, the entries not yet posted (with any JE numbers the user reported), the closing balances, and the lines already reconciled.
- **`opening_items`:** for a first run with earlier outstanding items, a CSV with the columns `kind,date,reference,description,amount,category,account,posting_ref,note`.
  - `kind`: `outstanding_check`, `outstanding_payment`, `deposit_in_transit`, `bank_error`, or `pending_adjustment`.
  - `amount`: signed as the ledger shows it. Outstanding checks are negative, deposits in transit positive, and a pending adjustment is its effect on the ledger.
- **`posted_before_opening`:** carried adjustment ids (`CF-...`) that the user posted dated in the previous period, so they are already in the ledger opening balance.
- **`history_file`:** optional coded-transactions history, read with `sources.history`, used only to suggest accounts.

## The review workbook

**Visible tabs:** Reconciliation, Journal entries, Bank raw, Ledger raw, Guide, Accounts. **Hidden (ordinary hiding):** Setup and Checks.

**What the user edits (yellow cells):**
- Reconciliation: Reconciled by, Review completed on (the reviewer's own sign-off date), and the Approval of timing items (outstanding checks and payments, deposits in transit, postings of carried entries). Items with an entry show their Journal entries choice here as a grey, display-only cell.
- Journal entries: each entry's non-cash lines (account, description, amounts, added split rows), Approval status, JE number, and Comments.
- Accounts: new accounts at the bottom.

**Fixed:** the raw tabs, the Refs, each entry's cash line, the proposals, and the Reconciliation columns from the files. Edits to them are not applied: `validate` restores them in the returned copy and logs each one. A changed source amount blocks completion until the user confirms it was accidental (`--accept-restore`) or uploads a corrected export (`--bank`, `--ledger`).

**Approval choices:** one dropdown per item, and the choice is the decision. "Approved" or "Pending". Items with two valid treatments offer both: "Approve ledger correction" or "Confirm bank error—no JE"; "Approved: reverse duplicate" or "Approved: second payment"; "Approved: outstanding payment" or "Approved: posting of <period> <Ref>". A choice without a journal entry needs no account and no JE number (its Check says "Not needed", or "Remove JE number" if one was typed). A confirmed bank error is carried forward until a bank line with its reference and amount clears it.

**Review completed on:** blank in the draft. `validate` records the date the reviewer enters (Run record, "Review completed on (entered by the reviewer)") next to its own result ("Completion") and timestamp ("Validated at"). The date never completes a reconciliation that has unresolved items or failed checks.

**Workbook format:** "review-2" (0.5.1). A "review-1" workbook (0.5.0) is validated only with `--from-0.5.0`: "Approved: correct ledger" and "Approved: bank error" are read as the two choices above, the Completion text 0.5.0 wrote in C6 is set aside, and each conversion is listed in "Edits found at the last upload". A checked review-1 workbook starts the next period without conversion.

**Hidden Checks tab (read by `validate`, `carry`, and `check`):** the live control checks, the effect of each choice, the decisions recorded (who, when), the carry forward, the Refs and line map, the entries as proposed, a copy of the source rows as received, the run record, and the formula expectations `check` compares after recalculation.

## validate

`validate WORKBOOK` reads the uploaded workbook, rebuilds the source files from its stored copy, reruns the rules, and lists what it found: approvals, pending items, edits kept or not applied, blocking issues, and questions. It records nothing and writes nothing.

With `--record` (only when the user asked to validate their reviewed workbook) it records the explicit approvals under Reconciled by, reruns the rules with them, and writes the checked workbook. Completion is set only when every check passes, nothing is pending, no question is open, and the rules say reconciled:
- **"Complete subject to posting (date): N approved entries without a JE number yet"**, when some approved entries have no JE number.
- **"Complete (date): every entry reported posted; next period's ledger confirms them"**, or **"Complete (date)"** when nothing needs posting.

The uploaded file is never changed.
