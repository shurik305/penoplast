#!/usr/bin/env python3
"""Edge-case regression tests (from roast #1). Usage: python3 edge_cases.py dist/   (exit 1 on failure)."""
import datetime as dt
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "_lib"))
from casekit import Checker, errors, run_cases, value  # noqa: E402

S, L, F = "'Settings'!", "'Turnover Log'!", "'Cleaning Fee Calculator'!"


def logrow(r, prop, start, finish):
    return {f"{L}B{r}": dt.datetime(2026, 9, 20), f"{L}C{r}": prop, f"{L}E{r}": start, f"{L}F{r}": finish,
            f"{L}K{r}": "No"}


def tiles(wb):
    lg = wb["Turnover Log"]
    return {lg.cell(row=4, column=c).value: lg.cell(row=5, column=c).value for c in range(2, 13)}


def main(dist):
    src = {"K": os.path.join(dist, "Short-Let-Turnover-Kit.xlsx")}
    t = dt.time
    cases = {
        "renamed_property": ("K", {f"{S}B11": "Harbour View Apt"}),
        "start_midnight": ("K", logrow(13, "Garden Cottage", t(0, 0), t(2, 0))),
        "custom_item_blank_scale": ("K", {f"{S}B40": "Scented candle", f"{S}C40": 4.0, f"{S}D40": 1, f"{S}E40": None}),
        "scale_typo": ("K", {f"{S}E30": "guests"}),
        "fee_150": ("K", {"FFee": 1.5}),
        "swapped_times": ("K", logrow(13, "Garden Cottage", t(13, 25), t(11, 0))),
        "unknown_log_property": ("K", logrow(13, "Harbor View", t(10, 0), t(12, 0))),
        "default": ("K", {}),
    }
    r = run_cases(src, cases)
    c = Checker()
    st = {n: value(wb, f"{F}D19") or "" for n, wb in r.items()}
    c.check("renamed property: no silent fee", value(r["renamed_property"], "FFeeOut") == 0,
            value(r["renamed_property"], "FFeeOut"))
    c.check("renamed property warns", st["renamed_property"].startswith("⚠ Property"), st["renamed_property"])
    lg = r["start_midnight"]["Turnover Log"]
    c.check("00:00 start gives minutes", lg["G13"].value == 120, lg["G13"].value)
    c.check("00:00 start pays by minutes", abs((lg["J13"].value or 0) - 40) < 0.01, lg["J13"].value)
    c.check("blank scale counted once per stay", abs((value(r["custom_item_blank_scale"], "FCons") or 0) - 21.04) < 0.01,
            value(r["custom_item_blank_scale"], "FCons"))
    c.check("scale typo warns", st["scale_typo"].startswith("⚠"), st["scale_typo"])
    fee = value(r["fee_150"], "FFeeOut")
    c.check("fee >= 100% stays sane", isinstance(fee, (int, float)) and fee >= 0, fee)
    c.check("swapped times flagged", (tiles(r["swapped_times"]).get("Rows to check") or 0) >= 1, tiles(r["swapped_times"]))
    c.check("unknown log property flagged", (tiles(r["unknown_log_property"]).get("Rows to check") or 0) >= 1,
            tiles(r["unknown_log_property"]))
    c.check("default status ok", st["default"].startswith("✔"), st["default"])
    for n, wb in r.items():
        c.check(f"{n}: no formula errors", not errors(wb), errors(wb)[:3])
    return c.report("shortlet-turnover-kit edge cases")


if __name__ == "__main__":
    sys.exit(main(sys.argv[1] if len(sys.argv) > 1 else "dist"))
