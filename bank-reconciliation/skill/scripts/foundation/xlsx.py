"""Workbook helpers: styled tables, reading tables back, and formula recalculation where a tool exists.

openpyxl writes formulas but cannot calculate them. Excel calculates them on opening. If LibreOffice is
available, recalculated() returns the calculated values so they can be compared with the script's own.
"""
from __future__ import annotations

import shutil, subprocess, tempfile, time
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

MONEY = '#,##0.00;(#,##0.00);"-"'
MONEY_ZERO = "#,##0.00;(#,##0.00);0.00"
DATE = "yyyy-mm-dd"
HEAD = PatternFill("solid", fgColor="1F3864")
BAND = PatternFill("solid", fgColor="D9E2F3")
KEY = PatternFill("solid", fgColor="FFF2CC")
ALERT = PatternFill("solid", fgColor="FCE4D6")
GOOD = PatternFill("solid", fgColor="E2EFDA")
BOLD = Font(bold=True)


def new_workbook(title):
    wb = Workbook()
    p = wb.properties
    p.creator = p.lastModifiedBy = "bank-reconciliation skill"
    p.title = title
    p.created = p.modified = datetime(2026, 1, 1)  # fixed, so repeated runs give identical metadata
    wb.calculation.fullCalcOnLoad = True
    return wb


def table(ws, cols, widths, start_row=1):
    for j, (c, w) in enumerate(zip(cols, widths), 1):
        cell = ws.cell(start_row, j, c)
        cell.font, cell.fill = Font(bold=True, color="FFFFFF"), HEAD
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.column_dimensions[cell.column_letter].width = w
    if start_row == 1:
        ws.freeze_panes = "A2"


def name(wb, label, sheet, ref):
    wb.defined_names[label] = DefinedName(label, attr_text=f"'{sheet}'!{ref}")


def read_table(path_or_wb, sheet, first_header):
    """Rows of the table on `sheet` whose header row starts with `first_header`, as dicts."""
    wb = path_or_wb if hasattr(path_or_wb, "sheetnames") else load_workbook(path_or_wb, data_only=False)
    if sheet not in wb.sheetnames:
        return []
    rows = list(wb[sheet].iter_rows(values_only=True))
    for i, r in enumerate(rows):
        if r and r[0] == first_header:
            head = [h for h in r]
            out = []
            for body in rows[i + 1:]:
                if not body or body[0] in (None, ""):
                    break
                out.append({h: v for h, v in zip(head, body) if h})
            return out
    return []


def soffice():
    return shutil.which("soffice") or shutil.which("libreoffice")


def recalculated(path):
    """Recalculate a copy of the workbook with LibreOffice.

    Returns (workbook, info). The workbook is a data-only copy with LibreOffice's calculated values, or None.
    `info` records what happened, so a failure can be diagnosed: the tool found (None if there is none), the
    command, exit code, the tail of its output, whether a copy was produced, and where it is.
    """
    exe = soffice()
    info = {"tool": exe, "attempted": False}
    if not exe:
        return None, info
    out = Path(tempfile.mkdtemp(prefix="bankrec_recalc_"))
    cmd = [exe, "--headless", "--convert-to", "xlsx", "--outdir", str(out), str(path)]  # as proven by the environment probe
    info.update(attempted=True, command=" ".join(cmd))
    start = time.monotonic()
    try:
        cp = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
        info.update(exit_code=cp.returncode, stdout=cp.stdout[-2000:], stderr=cp.stderr[-2000:])
    except (OSError, subprocess.SubprocessError) as e:
        info["error"] = f"{type(e).__name__}: {e}"[:500]
    info["seconds"] = round(time.monotonic() - start, 1)
    q = out / Path(path).name
    info["output_found"] = q.exists()
    if not q.exists():
        return None, info
    info["output"] = str(q)
    try:
        return load_workbook(q, data_only=True), info
    except Exception as e:  # keep the reason; the caller reports it as a failed recalculation
        info["error"] = f"Could not read LibreOffice's copy: {type(e).__name__}: {e}"[:500]
        return None, info
