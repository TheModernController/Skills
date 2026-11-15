"""Capability check: run first in every conversation, before any user file is read.

  python check_env.py

Uses only the Python standard library, so it works even when a required package is missing. Prints JSON:
`ok` is false when the skill cannot run here, and `explanation` says why in plain language. It makes no
network connections and reads no user files. It does not run LibreOffice, only looks for it.
"""
from __future__ import annotations

import importlib, json, os, shutil, sys, tempfile

SKILL_VERSION = "0.5.1"  # shown in every capability check and in each workbook's Run record
MIN_PYTHON = (3, 10)
UPLOADS, OUTPUTS = "/mnt/user-data/uploads", "/mnt/user-data/outputs"


def check():
    required, optional, explanation = [], [], []

    ok_py = sys.version_info[:2] >= MIN_PYTHON
    required.append({"check": f"Python {MIN_PYTHON[0]}.{MIN_PYTHON[1]} or later", "result": "PASS" if ok_py else "FAIL",
                     "detail": sys.version.split()[0]})
    if not ok_py:
        explanation.append(f"This environment has Python {sys.version.split()[0]}; the skill needs {MIN_PYTHON[0]}.{MIN_PYTHON[1]} or later.")

    xl = {"check": "openpyxl can write and read an Excel workbook", "result": "FAIL", "detail": ""}
    try:
        openpyxl = importlib.import_module("openpyxl")
        d = tempfile.mkdtemp(prefix="bankrec_env_")
        try:
            p = os.path.join(d, "check.xlsx")
            wb = openpyxl.Workbook()
            wb.active["A1"], wb.active["A2"] = 1.5, "=A1*2"
            wb.save(p)
            back = openpyxl.load_workbook(p).active
            if back["A1"].value == 1.5 and back["A2"].value == "=A1*2":
                xl["result"], xl["detail"] = "PASS", getattr(openpyxl, "__version__", "present")
            else:
                xl["detail"] = "a test workbook did not read back as written"
        finally:
            shutil.rmtree(d, ignore_errors=True)
    except ModuleNotFoundError:
        xl["detail"] = "not installed"
    except Exception as e:  # report, don't crash: this check exists to explain problems
        xl["detail"] = f"{type(e).__name__}: {e}"[:300]
    required.append(xl)
    if xl["result"] == "FAIL":
        explanation.append("The Python package openpyxl, which reads and writes Excel files, is not working here "
                           f"({xl['detail']}). On claude.ai it is normally supplied by the code execution environment. "
                           "Without it the skill cannot read the files or produce the workbook, so it stops rather than "
                           "attempting the reconciliation another way.")

    exe = shutil.which("soffice") or shutil.which("libreoffice")
    optional.append({"check": "LibreOffice, to recalculate the workbook's formulas before delivery",
                     "result": "available" if exe else "not found", "detail": exe or ""})
    if not exe:
        explanation.append("LibreOffice was not found, so the workbook's formulas will not be recalculated here. "
                           "Excel recalculates them when the workbook is opened, and the Checks line on the Reconciliation "
                           "tab shows any failure.")

    folders = {"uploads": {"path": UPLOADS, "exists": os.path.isdir(UPLOADS)},
               "outputs": {"path": OUTPUTS, "exists": os.path.isdir(OUTPUTS),
                           "writable": os.path.isdir(OUTPUTS) and os.access(OUTPUTS, os.W_OK)}}
    if not folders["outputs"]["writable"]:
        explanation.append(f"The download folder {OUTPUTS} is missing or not writable, so the workbook may "
                           "not be offered for download. On claude.ai, check that code execution and file creation is on "
                           "(Settings > Capabilities).")

    ok = all(c["result"] == "PASS" for c in required)
    return {"ok": ok, "skill_version": SKILL_VERSION, "required": required, "optional": optional, "folders": folders, "explanation": explanation}


if __name__ == "__main__":
    res = check()
    print(json.dumps(res, indent=1))
    sys.exit(0 if res["ok"] else 1)
