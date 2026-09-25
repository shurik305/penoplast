#!/usr/bin/env python3
"""QA for the Commercial Cleaning Bid Calculator. Usage: python3 qa.py dist/  (exit 1 on any problem)."""
import glob
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "_lib"))
from qakit import errors, named, near, recalc_open  # noqa: E402

NAMES = ["TotalArea", "RoutineHrsWeek", "MaxVisits", "BusiestVisit", "PeriodicHrsMonth", "RoutineHrsMonth",
         "SupHrs", "TotalHrsMonth", "CostMonth", "PriceMargin", "ConsCost", "ConsPrice", "PeriodicPrice", "FinalNet",
         "NetMonthly", "VATAmt", "Gross", "Annual", "Cleaners", "EffRate", "BERate", "ProfitMonth", "RatePerHr",
         "CurrencyLabel", "PeriodicList", "FixedPerHr", "Overhead"]


def check(path):
    wb = recalc_open(path)
    v = {n: named(wb, n) for n in NAMES}
    probs = errors(wb)
    cp = wb["Client Proposal"]
    fee = {}
    for row in cp.iter_rows(min_row=1, max_row=cp.max_row):
        label = row[1].value
        if isinstance(label, str) and label.startswith(("Regular cleaning", "Periodic services (", "Consumables (",
                                                         "Subtotal per month", "TOTAL PER MONTH")):
            fee[label.split(" ")[0]] = row[4].value
    lines = sum(fee.get(k) or 0 for k in ("Regular", "Periodic", "Consumables"))
    if not v["FinalNet"] or v["FinalNet"] <= 0:
        probs.append("FinalNet not positive")
    if not near(lines, fee.get("Subtotal")):
        probs.append(f"proposal lines {lines:.2f} != subtotal {fee.get('Subtotal')}")
    if not near(v["NetMonthly"] + v["VATAmt"], v["Gross"]):
        probs.append("net + VAT != gross")
    if not near(fee.get("TOTAL"), v["Gross"]):
        probs.append("proposal total != gross")
    if not (v["Cleaners"] and v["Cleaners"] >= 1):
        probs.append(f"cleaners needed = {v['Cleaners']}")
    if (fee.get("Regular") or 0) <= 0:
        probs.append("regular cleaning line not positive")
    if abs(v["Overhead"] - v["FixedPerHr"]) > 0.5:
        probs.append(f"default overhead {v['Overhead']} vs fixed/hour {v['FixedPerHr']:.2f} inconsistent")
    scope = next((c.value for row in cp.iter_rows() for c in row
                  if isinstance(c.value, str) and c.value.startswith("General office")), None)
    if not scope:
        probs.append("scope of work text missing")
    return v, fee, probs


if __name__ == "__main__":
    bad = 0
    for fp in sorted(glob.glob(os.path.join(sys.argv[1] if len(sys.argv) > 1 else "dist", "*.xlsx"))):
        v, fee, probs = check(fp)
        print(f"== {os.path.basename(fp)}")
        print(f"   area {v['TotalArea']} · {v['RoutineHrsWeek']:.2f} h/week · busiest visit {v['BusiestVisit']:.2f} h"
              f" · cleaners {v['Cleaners']} · periodic {v['PeriodicHrsMonth']:.2f} h/mo · total {v['TotalHrsMonth']:.1f}"
              f" h/mo")
        print(f"   net/month {v['NetMonthly']} {v['CurrencyLabel']} · VAT {v['VATAmt']} · gross {v['Gross']} · annual "
              f"{v['Annual']} · eff/h {v['EffRate']:.2f} · BE/h {v['BERate']:.2f} · profit/mo {v['ProfitMonth']:.2f}")
        print(f"   proposal fee lines {fee}")
        print(f"   periodic: {v['PeriodicList']}")
        for pr in probs:
            print("   PROBLEM:", pr)
        bad += bool(probs)
    sys.exit(1 if bad else 0)
