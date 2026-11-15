"""Read a chart of accounts export into a list of account names ("1010 Operating Cash").

Accepts Excel or CSV with a header row. Uses a number or code column and a name column when both exist, or a single
account column. Anything it can't read is reported rather than guessed.
"""
from __future__ import annotations

import re
from pathlib import Path

from .tables import InputError, read_grid, text

NUMBER = re.compile(r"^(account\s*)?(no\.?|number|num|code|#|acct\s*(no|number|#)?)$", re.I)
NAME = re.compile(r"^(account\s*)?(name|title|description)$", re.I)
SINGLE = re.compile(r"^(account|gl\s*account|account\s*name|ledger\s*account)$", re.I)


def read_chart(path, sheet=None):
    _, rows, _ = read_grid(path, sheet)
    for h, row in enumerate(rows[:15]):
        head = [text(c).strip() for c in row]
        num = next((j for j, c in enumerate(head) if NUMBER.match(c)), None)
        name = next((j for j, c in enumerate(head) if NAME.match(c)), None)
        single = next((j for j, c in enumerate(head) if SINGLE.match(c)), None)
        if (num is not None and name is not None) or single is not None or name is not None:
            out = []
            for r in rows[h + 1:]:
                get = lambda j: text(r[j]).strip() if j is not None and j < len(r) else ""
                if num is not None and name is not None:
                    a = f"{get(num)} {get(name)}".strip()
                elif single is not None:
                    a = get(single)
                else:
                    a = get(name)
                if a and a not in out:
                    out.append(a)
            if not out:
                raise InputError(f"{Path(path).name}: no accounts found under the header row.")
            return out
    raise InputError(f"{Path(path).name}: couldn't find an account name column. Which column holds the account names?")
