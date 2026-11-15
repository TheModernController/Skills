# Bank reconciliation skill for Claude

Version 0.5.1: the version shown in Day 1 of the Month-End Automation series.

### [Download the skill: bank-reconciliation-0.5.1.zip](https://github.com/TheModernController/Skills/raw/main/bank-reconciliation/bank-reconciliation-0.5.1.zip)

This is the file you upload to Claude. Don't unzip it. GitHub's green Code > Download ZIP button downloads the whole repository instead, which Claude can't install.

## What it does

It reconciles one bank account for one period from your Excel or CSV exports.

1. Claude reads your bank and ledger files and returns a review workbook.
2. You review it in Excel: approve items, adjust the proposed journal entries, and type JE numbers once you've posted them.
3. You upload the workbook back. Claude validates your edits and returns a checked workbook.
4. Next month starts from that checked workbook. There's no separate setup file.

The scripts do every calculation and match. Claude runs them and explains the results, and you make the decisions.

## Requirements

- A Claude plan with **code execution and file creation** turned on, and custom skills allowed. On Team and Enterprise plans, an owner may need to allow skills first.
- **Excel**, to review the workbook.
- For each period:
  - your bank's transaction export, as Excel or CSV;
  - your ledger's cash-account activity, as Excel or CSV;
  - the statement's opening and closing balances;
  - optionally, a chart of accounts export, which gives you proper account lists in the workbook.

There's nothing to install on your computer. The skill's scripts run in Claude's hosted code execution environment, using the Python and openpyxl it provides.

**Your files are processed in that hosted environment, on Anthropic's systems.** Use real data only if your organization's policy allows it.

## Install

These steps were tested with version 0.5.1 on claude.ai:

1. Download [bank-reconciliation-0.5.1.zip](https://github.com/TheModernController/Skills/raw/main/bank-reconciliation/bank-reconciliation-0.5.1.zip). Don't unzip it.
2. Turn on code execution and file creation in Claude's settings.
3. Upload the ZIP as a custom skill.
4. Optional, once: in a new conversation, ask Claude to "run the bank reconciliation self-test". It builds a fictional workbook and checks its formulas.
5. Start a new conversation and say something like "I'd like to reconcile my operating bank account." Claude asks for your files when it needs them.

### Where to find these settings

Claude's menus change from time to time, so the labels you see may differ:

- **Uploading the skill:** when 0.5.1 was tested in November 2026, the ZIP was uploaded at Customize > Skills on claude.ai. In the video, the installed skill appears on a page titled Plugins.
- **Code execution and file creation:** these are in Claude's settings, but their current menu label hasn't been confirmed. If you can't find them, search the settings for "code execution".

## Try it with the example files

The [examples](examples/README.md) folder has a fictional company's March and April files: the same two months as in the video.

## Limits

- One account, one currency and one period per run.
- Excel and CSV only. PDF statements, batched deposits, card settlements net of fees, foreign currency and scheduled runs aren't handled.
- **Nothing is posted.** Journal entries are proposals for an accountant to review, approve and post in their own system.
- Approvals are recorded under the name in Reconciled by when you ask Claude to validate. They aren't an authenticated sign-off.
- Entries use a standard review format, not your own JE template or an import file.
- Check the workbook, not only the chat. The workbook holds the figures and the controls.

## What's in this folder

| Item | What it is |
| --- | --- |
| `bank-reconciliation-0.5.1.zip` | The skill to upload to Claude |
| `skill/` | The same skill unzipped, so you can read every instruction and script before uploading |
| `examples/` | Fictional files to practise on, with a fact sheet and the expected results |

## Version

0.5.1 was verified in Claude on fictional data with a hosted self-test, which recalculated every formula, and a manual two-month trial.

## License

Licensed under the [MIT License](../LICENSE). No third-party code is bundled: the scripts use Python's standard library, plus openpyxl, which Claude's environment provides.
