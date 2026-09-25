"""Shared QA helpers: recalculate a workbook in LibreOffice and inspect computed values."""
import os
import subprocess
import tempfile

import openpyxl

ERR_MARKERS = ("#NAME", "#VALUE", "#REF", "#DIV", "#N/A", "#NUM", "Err:", "#NULL")


def recalc_open(path):
    """Returns an openpyxl workbook (data_only) with values computed by LibreOffice.

    Requires LibreOffice Calc with OOXMLRecalcMode=0 in the profile at $LO_HOME (see ops/setup.sh).
    """
    env = dict(os.environ, HOME=os.environ.get("LO_HOME", "/tmp/lohome"))
    with tempfile.TemporaryDirectory() as td:
        subprocess.run(["soffice", "--headless", "--norestore", "--convert-to", "xlsx", "--outdir", td, path],
                       check=True, capture_output=True, env=env, timeout=240)
        return openpyxl.load_workbook(os.path.join(td, os.path.basename(path)), data_only=True)


def errors(wb):
    out = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith(ERR_MARKERS):
                    out.append(f"{ws.title}!{c.coordinate} = {c.value}")
    return out


def named(wb, name):
    dn = wb.defined_names[name]
    (sheet, ref), = list(dn.destinations)
    return wb[sheet][ref.replace("$", "")].value


def near(a, b, tol=0.005):
    return abs((a or 0) - (b or 0)) <= tol
