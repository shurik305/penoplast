"""Edge-case regression harness: copy a built workbook, change inputs, recalculate in LibreOffice, read results.

Based on the adversarial harness from roast #1 (lens B). Inputs are addressed by defined name or 'Sheet'!A1.
"""
import os
import subprocess
import tempfile

import openpyxl

from qakit import errors  # noqa: F401  (re-exported for edge_cases.py)


def resolve(wb, key):
    if "!" in key:
        sh, ref = key.rsplit("!", 1)
        return sh.strip("'"), ref.replace("$", "")
    (sh, ref), = list(wb.defined_names[key].destinations)
    return sh, ref.replace("$", "")


def run_cases(src_by_key, cases, workdir=None):
    """cases: {name: (src_key, {input_ref: value})}. Returns {name: recalculated openpyxl workbook (data_only)}."""
    workdir = workdir or tempfile.mkdtemp(prefix="cases-")
    inp, out = os.path.join(workdir, "in"), os.path.join(workdir, "out")
    os.makedirs(inp, exist_ok=True)
    os.makedirs(out, exist_ok=True)
    files = []
    for name, (src_key, sets) in cases.items():
        wb = openpyxl.load_workbook(src_by_key[src_key])  # keeps formulas
        for key, val in sets.items():
            sh, ref = resolve(wb, key)
            wb[sh][ref].value = val
        dst = os.path.join(inp, f"{name}.xlsx")
        wb.save(dst)
        files.append(dst)
    env = dict(os.environ, HOME=os.environ.get("LO_HOME", "/tmp/lohome"))
    subprocess.run(["soffice", "--headless", "--norestore", "--convert-to", "xlsx", "--outdir", out] + files,
                   check=True, capture_output=True, env=env, timeout=1800)
    return {name: openpyxl.load_workbook(os.path.join(out, f"{name}.xlsx"), data_only=True) for name in cases}


def value(wb, key):
    sh, ref = resolve(wb, key)
    return wb[sh][ref].value


class Checker:
    def __init__(self):
        self.failures = []
        self.passed = 0

    def check(self, name, ok, detail=""):
        if ok:
            self.passed += 1
        else:
            self.failures.append(f"{name}: {detail}")

    def report(self, title):
        print(f"{title}: {self.passed} checks passed, {len(self.failures)} failed")
        for f in self.failures:
            print("  FAIL", f)
        return 1 if self.failures else 0
