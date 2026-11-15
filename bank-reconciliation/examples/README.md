# Example files

Fictional data for a fictional company, Alder Creek Outfitters, and its operating account ending 4821, in USD. These are the same two months as in the video.

The [fact sheet](fact-sheet.md) has the company's details and what actually happened each month. Use it to answer Claude's questions in your own words and to make your review decisions.

## 01: first month (March 2026)

Start a new conversation. When Claude asks for files, upload:
- `01-first-month/bank_activity.csv`
- `01-first-month/cashbook.csv`
- `01-first-month/chart-of-accounts.csv`

Review the workbook in Excel as the fact sheet describes, and type the JE numbers it gives. Then upload the workbook and ask Claude to validate your reviewed workbook. Keep the checked workbook Claude returns, because April starts from it.

**What to expect:**
- The draft's remaining difference is **635.00**, with 7 items to review.
- After you approve every item as the fact sheet describes, the difference is **0.00**.
- Validation reports the reconciliation complete subject to posting: 2 approved entries without a JE number yet, and 3 reported posted.

## 03: next month (April 2026)

Start a new conversation, say "next period", and upload together:
- the checked March workbook from 01;
- `03-next-month/bank_activity_2026-04.csv`;
- `03-next-month/cashbook_2026-04.csv`.

The bank export now calls its description column "details", and it repeats one line from 31 March. Claude should ask you about the renamed column, and leave out the repeated line.

**What to expect:**
- The remaining difference before your April decisions is **−505.00**.
- After you approve every item as the fact sheet describes, it's **0.00**.

Claude's wording will vary from run to run, but the figures shouldn't.
