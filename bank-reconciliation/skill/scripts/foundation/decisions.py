"""Decisions the user has stated, tied to the exact item they reviewed.

Each decision stores the item's fingerprint: a hash of the type and the content of its source lines.
If the item later changes or disappears, the decision is not applied and is listed for re-confirmation.
Whether a treatment is allowed is decided by the accounting rules, not here.
"""
from __future__ import annotations

import hashlib, json
from decimal import Decimal, InvalidOperation
from pathlib import Path

FIELDS = ["item_id", "treatment", "reviewed_amount", "line_ref", "account", "note", "approved_by", "approved_on", "fingerprint"]


def fingerprint(item_type, line_contents):
    blob = json.dumps([item_type] + sorted(list(map(str, c)) for c in line_contents))
    return hashlib.sha256(blob.encode()).hexdigest()[:16]


def clean(d):
    out = {k: str(d.get(k, "") or "").strip() for k in FIELDS}
    return out


def load(path):
    """Read decisions from a JSON list, or from the Decision log sheet of an earlier workbook for the same period."""
    if not path:
        return []
    p = Path(path)
    if not p.exists():
        return []
    if p.suffix.lower() in (".xlsx", ".xlsm"):
        from .xlsx import read_table
        rows = read_table(p, "Decision log", "Item ID")
        key = {"Item ID": "item_id", "Treatment": "treatment", "Reviewed amount": "reviewed_amount", "Line ref": "line_ref",
               "Account": "account", "Note": "note", "Approved by": "approved_by", "Approved on": "approved_on", "Fingerprint": "fingerprint"}
        return [clean({key[k]: v for k, v in r.items() if k in key}) for r in rows]
    data = json.loads(p.read_text(encoding="utf-8"))
    return [clean(d) for d in data]


def save(decisions, path):
    Path(path).write_text(json.dumps([clean(d) for d in decisions], indent=2) + "\n", encoding="utf-8")


def merge(existing, new):
    """New decisions replace earlier ones for the same item. Returns (merged list, replaced item ids)."""
    by_id = {d["item_id"]: d for d in existing}
    replaced = [d["item_id"] for d in new if d["item_id"] in by_id]
    for d in new:
        by_id[d["item_id"]] = d
    return list(by_id.values()), replaced


def freshness(decision, item):
    """Return '' if the decision still describes this item, otherwise the reason it cannot be applied."""
    if item is None:
        return "Item is no longer in this run (its source lines changed or were removed). Confirm again if still wanted."
    if decision.get("fingerprint") and decision["fingerprint"] != item.fingerprint:
        return "Item's source lines changed after this decision. Review it again."
    try:
        reviewed = Decimal(decision.get("reviewed_amount") or "x")
    except InvalidOperation:
        return "Reviewed amount is missing or invalid"
    if reviewed != item.amount:
        return f"Reviewed amount {reviewed} does not equal the item's amount {item.amount}"
    return ""
