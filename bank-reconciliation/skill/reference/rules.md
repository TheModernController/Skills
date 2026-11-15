# Rules, items, treatments, and statuses

## Matching rules

Each line is used once. Amounts must match exactly. Input order never changes the result.

| Rule | Condition | Result |
| --- | --- | --- |
| R1 | Same reference or check number, same amount, one candidate each side. Equal numbers of identical lines on both sides are paired in date order, since identical transactions can be genuine. | Matched |
| R2 | Same direction and amount, dates within the window, no conflicting references, unique candidates (identical groups paired as in R1) | Matched |
| H1 | Same reference, different amount | Held: Amount mismatch |
| H2 | More than one possible partner | Held: Multiple candidates (noted as a possible duplicate when the amounts are equal) |
| P1 | A ledger line carries a carried entry's JE number and the same amount | That entry is posted |
| C1 | Unmatched ledger check | Proposed outstanding check |
| C2 | Unmatched ledger receipt in the last N days of the period | Proposed deposit in transit |
| C3 | Unmatched bank debit whose description is on the fee list | Proposed bank fee entry |
| C4 | Anything else unmatched | Held for review |
| D1 | Row with a missing or unreadable date or amount | Held; correct the export and rerun |
| CF | Item carried from an earlier period, not yet cleared | Stays approved and is carried again |

Carried bank-side items match new bank lines by reference (any date) or by amount and direction within the window from the period start.

## Treatments

From 0.5.0 the user chooses a treatment in the workbook, with one choice per item: the Approval status on the Journal entries tab, or the Approval cell on the Reconciliation tab for items with no entry. `validate --record` turns each explicit approval into the treatment below. `reviewed_amount` equals the item's amount.

| Item type | Workbook choice | Treatment | Meaning |
| --- | --- | --- | --- |
| Proposed outstanding check | Approved (Reconciliation) | `accept_outstanding` | Not yet cleared by the bank |
| Proposed deposit in transit | Approved (Reconciliation) | `accept_in_transit` | Deposited, not yet on the statement |
| Proposed bank fee entry | Approved (Journal entries) | `accept_fee_entry` | A ledger entry to the account on the entry |
| Amount mismatch | Approve ledger correction | `correct_cashbook_to_bank` | The bank is right; a ledger correction entry |
| Amount mismatch | Confirm bank error—no JE | `bank_error_to_correct` | The bank is wrong; no entry, no account, no JE number; carried as a bank error until a bank line with its reference and amount clears it |
| Multiple candidates (one bank line, two equal ledger lines) | Approved: reverse duplicate | `reverse_duplicate` | The later ledger line is a duplicate; a reversing entry |
| Multiple candidates (same shape) | Approved: second payment | `second_payment_outstanding` | A genuine second payment not yet cleared, no entry |
| Returned item (one possible original receipt) | Approved (Journal entries) | `reverse_application_reopen_receivable` | Reverse the receipt and reopen the receivable; no new invoice |
| Unidentified bank item | Approved (Journal entries) | `book_bank_item` | A ledger entry to the account on the entry |
| Unmatched ledger payment | Approved: outstanding payment | `accept_outstanding_payment` | A payment not yet cleared by the bank |
| Unmatched ledger payment or receipt | Approved: posting of <period> <Ref> | `posting_of_carried` | This line posts a carried entry with exactly this amount |

Pending leaves the item open, so the period cannot be completed. Items with no choice (an incomplete source row, a returned item with no single original, other multiple-candidate shapes, unmatched ledger receipts outside the deposit-in-transit window) need corrected files or the user's explanation.

There is no treatment that reverses an unmatched ledger line: an unmatched ledger line is not enough evidence that it was wrong.

## Proposed entries

Every item that needs a ledger entry gets one entry on the Journal entries tab, identified by its Ref (R-01). The Ref never changes when rows are sorted, entries are split, or decisions change. The cash line comes from the bank or ledger line and is fixed; the user chooses the other account(s) and may split them. Nothing is posted.

When the user posts an entry, they type its JE number. Next period, a ledger line carrying that JE number (as its reference or in its description) and the same amount closes the carried entry (P1). Without the number, the run offers "Approved: posting of <period> <Ref>" on the matching ledger line for the user to confirm. An entry still not posted is carried again and never proposed twice.

## Statuses (set by the script; Completion is set by `validate --record`)

| Status | Meaning |
| --- | --- |
| `Incomplete: missing ...` | Evidence needed for a reconciliation is missing. The workbook is a matching review only |
| `Not reconciled: ...` | A check failed (statement roll-forward, ledger roll-forward, opening balances against the previous period, totals rows, running balance, row accounting), a difference remains, or items need decisions |
| `Reconciled subject to posting N approved adjustment(s) (net X)` | Adjusted bank equals adjusted ledger, every check passes, nothing is open, but the ledger does not show the reconciled balance until those entries are posted |
| `Reconciled, awaiting sign-off` | As above, with nothing left to post |

## Balances

- **Bank closing** = stated opening + included bank rows. It must equal the stated closing balance.
- **Actual ledger closing** = ledger opening + included ledger rows. This is what the ledger shows, and next period's ledger opening is checked against it.
- **Adjusted bank** = bank closing + approved deposits in transit − approved outstanding checks and payments ± bank errors, including carried items.
- **Adjusted ledger** = actual ledger closing + approved adjustments not yet posted. This is the reconciled balance.
- **Opening check (next period):**
  - The bank opening must equal the previous stated closing.
  - The ledger opening must equal the previous actual ledger closing plus any adjustments posted dated before the opening.
  - Bank opening + carried bank-side items must equal ledger opening + carried adjustments not yet posted.
