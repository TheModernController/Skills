# Fact sheet: Alder Creek Outfitters (fictional)

Everything here is fictional. It is what is true for the examples, not a script. Answer Claude in your own words, and only what it asks, as you would at work. If Claude asks for something not covered here, say you don't know.

## The company and account (both examples)

| Fact | Value |
| --- | --- |
| Company | Alder Creek Outfitters (fictional) |
| Bank account | The operating account, ending 4821. No other digits exist. |
| Currency | USD |
| Accounting system | A fictional ledger, not any named product. Claude shouldn't need to ask. If it does, say "a fictional ledger" |
| How often it is reconciled | Monthly |
| Who reconciles | "Fictional reviewer" |
| Ledger cash account | "1010 Cash - Operating account" |
| Chart of accounts | `01-first-month/chart-of-accounts.csv` (attach it when Claude asks) |
| Journal entry template | None. Say you don't have one |
| Matching rules | The standard rules are fine. There are no special day limits |

## 01: March 2026, the first reconciliation

- **What you're doing:** the first reconciliation of this account with this skill, for March 2026 (1 to 31 March).
- **Files you have:** the bank transactions export, the ledger activity for the cash account, and the chart of accounts. There is no earlier reconciliation and no outstanding-items list. Nothing was outstanding at the end of February.
- **Known issues:** none.
- **Bank statement:** opening balance 48,250.00, closing balance 46,517.10.
- **Ledger:** the cash account's balance on 1 March was 48,250.00.

**What actually happened in March.** These are the facts behind your review. Apply them in the workbook (Journal entries tab, and the yellow Approval cells on the Reconciliation tab):

| Transaction | Fact | In the workbook |
| --- | --- | --- |
| Check 2053, 1,275.00 | Issued 30 March. Not yet presented at the bank | Approved (Reconciliation) |
| INV-1044 receipt, 2,640.00 | Deposited after the bank's cut-off on 31 March | Approved (Reconciliation) |
| Monthly service fee, 35.00 | A genuine bank charge | Approved, 6150 Bank fees |
| Check 2052 | The bank paid 450.00, but the ledger shows 540.00. The bank is right: the ledger amount was keyed wrong. It's Mossy Print's payable | Approve ledger correction, 2000 Accounts payable |
| ACH-5523 to Cedar Freight | Entered twice in the ledger and paid once. The second ledger entry is the duplicate. It's Cedar Freight's payable | Approved: reverse duplicate, 2000 Accounts payable |
| Returned item, 480.00 | Fern Street Gym's payment of INV-1043 bounced. Reverse that receipt and reopen the receivable | Approved, 1200 Accounts receivable |
| Incoming wire, 1,500.00 | A customer prepayment | Approved, 2300 Customer deposits |

**Posting, before you upload the reviewed workbook.** On 3 April you post three of the entries in the fictional ledger and type their JE numbers in the workbook:

| Entry | JE number |
| --- | --- |
| Bank fee | JE-0457 |
| Check 2052 correction | JE-0458 |
| ACH-5523 duplicate reversal | JE-0459 |

The returned-item reversal isn't posted yet. The prepayment entry was posted on 2 April, but you don't have its number to hand, so leave its JE number blank. Then upload the workbook and ask Claude to validate your reviewed workbook.

## 03: April 2026, the next period, in a new conversation

- **What you're doing:** next period, for April 2026 (1 to 30 April).
- **Files you have:** the checked March workbook Claude returned in example 01, plus the April bank and ledger exports. There is no separate setup file.
- **Chart of accounts:** use the saved one.
- **Known issues:** none beyond the March entries below.
- **Bank statement:** opening balance 46,517.10, closing balance 46,901.85.
- **Ledger:** the cash account's balance on 1 April was 45,882.10.
- **Bank export:** the bank renamed its description column to "details". It is the same field.

**March's entries.** None of them was posted dated in March.

| Entry | Fact |
| --- | --- |
| Bank fee, check 2052 correction, ACH-5523 duplicate reversal | Posted in April as JE-0457, JE-0458 and JE-0459. The April ledger shows those numbers |
| Returned item reversal | Not posted yet |
| Customer prepayment, 1,500.00 | Posted on 2 April without a reference. It is the ledger receipt of 1,500.00 on 2 April: confirm it as the posting of the March prepayment entry |

**What actually happened in April:**

| Transaction | Fact |
| --- | --- |
| Check 2055 to Cedar Freight, 740.00 | Issued 28 April. Not yet presented |
| INV-1046 receipt from Harbor Yoga, 860.00 | Deposited after the bank's cut-off on 30 April |
| ACH-5525 to Ashgrove Realty, 1,150.00 | Sent 29 April. It settles at the bank on 1 May (an outstanding payment) |
| Monthly service fee, 35.00 | A genuine bank charge. Record it to 6150 Bank fees |

You don't post any April entries before validating.
