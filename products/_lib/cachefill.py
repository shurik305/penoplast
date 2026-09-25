"""Writes LibreOffice-computed results into the cached <v> values of formula cells.

xlsxwriter stores 0 as the cached result of every formula. Excel recalculates on open, but previews (phones,
Gmail, macOS Quick Look, Excel Protected View) show the cached zeros. This post-processor recalculates a copy in
LibreOffice and patches the original file's XML in place, keeping all formatting, validation and protection.
"""
import datetime as dt
import os
import re
import subprocess
import tempfile
import zipfile
from xml.sax.saxutils import escape

import openpyxl

NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CELL_RE = re.compile(r'<c r="([A-Z]+[0-9]+)"((?: s="[0-9]+")?)><f>(.*?)</f><v>0</v></c>', re.S)
EPOCH = dt.datetime(1899, 12, 30)


def _computed(path):
    env = dict(os.environ, HOME=os.environ.get("LO_HOME", "/tmp/lohome"))
    with tempfile.TemporaryDirectory() as td:
        subprocess.run(["soffice", "--headless", "--norestore", "--convert-to", "xlsx", "--outdir", td, path],
                       check=True, capture_output=True, env=env, timeout=300)
        wb = openpyxl.load_workbook(os.path.join(td, os.path.basename(path)), data_only=True)
        return {ws.title: {c.coordinate: c.value for row in ws.iter_rows() for c in row if c.value is not None}
                for ws in wb.worksheets}


def _sheet_files(z):
    wbxml = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    target = dict(re.findall(r'<Relationship Id="(rId[0-9]+)" Type="[^"]+/worksheet" Target="([^"]+)"', rels))
    out = {}
    for name, rid in re.findall(r'<sheet name="([^"]+)" sheetId="[0-9]+" r:id="(rId[0-9]+)"', wbxml):
        name = name.replace("&amp;", "&").replace("&apos;", "'").replace("&quot;", '"')
        out["xl/" + target[rid].lstrip("/").removeprefix("xl/")] = name
    return out


def _cell(ref, style, formula, v):
    if v is None:
        return f'<c r="{ref}"{style} t="str"><f>{formula}</f><v></v></c>'
    if isinstance(v, bool):
        return f'<c r="{ref}"{style} t="b"><f>{formula}</f><v>{int(v)}</v></c>'
    if isinstance(v, (int, float)):
        return f'<c r="{ref}"{style}><f>{formula}</f><v>{repr(float(v)) if isinstance(v, float) else v}</v></c>'
    if isinstance(v, dt.datetime):
        serial = (v - EPOCH).total_seconds() / 86400
        return f'<c r="{ref}"{style}><f>{formula}</f><v>{serial}</v></c>'
    if isinstance(v, dt.date):
        return f'<c r="{ref}"{style}><f>{formula}</f><v>{(dt.datetime.combine(v, dt.time()) - EPOCH).days}</v></c>'
    if isinstance(v, dt.time):
        frac = (v.hour * 3600 + v.minute * 60 + v.second) / 86400
        return f'<c r="{ref}"{style}><f>{formula}</f><v>{frac}</v></c>'
    if isinstance(v, dt.timedelta):
        return f'<c r="{ref}"{style}><f>{formula}</f><v>{v.total_seconds() / 86400}</v></c>'
    s = str(v)
    if s.startswith("#"):
        return f'<c r="{ref}"{style} t="e"><f>{formula}</f><v>{escape(s)}</v></c>'
    return f'<c r="{ref}"{style} t="str"><f>{formula}</f><v>{escape(s)}</v></c>'


def fill(path):
    """Patches cached formula results in place. Returns the number of cells updated."""
    values = _computed(path)
    tmp = path + ".tmp"
    n = 0
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        sheets = _sheet_files(zin)
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename in sheets:
                vals = values.get(sheets[item.filename], {})

                def repl(m):
                    nonlocal n
                    ref, style, formula = m.group(1), m.group(2), m.group(3)
                    n += 1
                    return _cell(ref, style, formula, vals.get(ref))
                data = CELL_RE.sub(repl, data.decode("utf-8")).encode("utf-8")
            zout.writestr(item, data)
    os.replace(tmp, path)
    return n


if __name__ == "__main__":
    import sys
    for p in sys.argv[1:]:
        print(p, fill(p), "formula cells filled")
