"""Titled tables on a worksheet, written and read back by title.

A section is a bold title cell, a header row under it, and data rows until the first row whose first cell is
empty. The supporting tabs (Setup and Checks) keep everything Claude needs next time in sections like this, so
no separate setup file is needed.
"""
from __future__ import annotations

from openpyxl.styles import Border, Font, Side

LINE = Side(style="thin", color="808080")


def write_section(ws, row, title, headers, rows, col=2, number_formats=None):
    """Write a section starting at `row`. Returns the first row after it (leaving one blank row)."""
    ws.cell(row, col, title).font = Font(bold=True)
    for j, h in enumerate(headers):
        c = ws.cell(row + 1, col + j, h)
        c.font = Font(bold=True)
        c.border = Border(bottom=LINE)
    for i, values in enumerate(rows):
        for j, v in enumerate(values):
            c = ws.cell(row + 2 + i, col + j, v)
            if number_formats and headers[j] in number_formats and v is not None and not isinstance(v, str):
                c.number_format = number_formats[headers[j]]
    return row + 2 + len(rows) + 1


def find(ws, title, col=2):
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, col).value == title:
            return r
    return None


def read_section(ws, title, col=2):
    """Rows of the section as dicts keyed by header. An absent section gives []."""
    r = find(ws, title, col)
    if r is None:
        return []
    headers = []
    j = col
    while ws.cell(r + 1, j).value not in (None, ""):
        headers.append(ws.cell(r + 1, j).value)
        j += 1
    out = []
    i = r + 2
    while ws.cell(i, col).value not in (None, ""):
        out.append({h: ws.cell(i, col + k).value for k, h in enumerate(headers)})
        i += 1
    return out


def read_kv(ws, title, col=2):
    """A two-column section (Field, Value) as a dict."""
    return {r.get("Field"): r.get("Value") for r in read_section(ws, title, col)}


def section_rows(ws, title, col=2):
    """(first data row, last data row) of a section, for building formulas that refer to it."""
    r = find(ws, title, col)
    if r is None:
        return None
    first = r + 2
    last = first - 1
    while ws.cell(last + 1, col).value not in (None, ""):
        last += 1
    return first, last
