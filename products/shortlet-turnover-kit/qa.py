#!/usr/bin/env python3
"""QA for the Short-Let Turnover Kit. Usage: python3 qa.py dist/  (exit 1 on any problem).

Expected values are worked out by hand from the defaults in build.py (see comments).
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "_lib"))
from qakit import errors, named, near, recalc_open  # noqa: E402


def main(dist):
    wb = recalc_open(os.path.join(dist, "Short-Let-Turnover-Kit.xlsx"))
    probs = errors(wb)
    v = {n: named(wb, n) for n in ("FCleaner", "FLaundry", "FCons", "FSupplies", "FTotal", "FFeeOut", "FKeep",
                                   "RTotal", "RGuests")}
    # Harbour View: cleaner 55 fixed; laundry 2 beds × 8; consumables 17.04 for 4 guests/1 bath/2 bedrooms; supplies 3
    expect = {"FCleaner": 55, "FLaundry": 16, "FCons": 17.04, "FSupplies": 3, "FTotal": 91.04, "FFeeOut": 110,
              "FKeep": 110 * 0.97 - 91.04}
    for k, e in expect.items():
        if not near(v[k], e, 0.01):
            probs.append(f"{k} = {v[k]} (expected {e})")
    lg = wb["Turnover Log"]
    minutes = [lg.cell(row=r, column=7).value for r in (10, 11, 12)]
    pay = [lg.cell(row=r, column=10).value for r in (10, 11, 12)]
    if minutes != [145, 220, 85]:
        probs.append(f"log minutes {minutes} (expected [145, 220, 85])")
    if not all(near(a, b, 0.01) for a, b in zip(pay, [55, 220 / 60 * 20, 40])):
        probs.append(f"log pay {pay}")
    tiles = {lg.cell(row=4, column=c).value: lg.cell(row=5, column=c).value for c in (2, 4, 5, 7, 9, 11)}
    if tiles.get("Turnovers logged") != 3 or tiles.get("Issues reported") != 1:
        probs.append(f"log tiles {tiles}")
    if not near(tiles.get("Unpaid"), 220 / 60 * 20 + 40, 0.01):
        probs.append(f"unpaid tile {tiles.get('Unpaid')}")
    fc = wb["Cleaning Fee Calculator"]
    table = [[fc.cell(row=r, column=c).value for c in range(2, 10)] for r in range(24, 34)]
    filled = [row for row in table if row[0]]
    if len(filled) != 3:
        probs.append(f"all-properties table rows {len(filled)} (expected 3): {table[:4]}")
    rs = wb["Restock List"]
    first_item = [rs.cell(row=9, column=c).value for c in range(2, 8)]
    if first_item[:2] != ["Toilet paper roll", 8]:
        probs.append(f"restock first row {first_item} (expected toilet paper, 8 needed)")
    print(f"fee calc: {v}")
    print(f"log minutes {minutes} pay {pay} tiles {tiles}")
    print(f"all properties: {filled}")
    print(f"restock first row {first_item}, total {v['RTotal']}")
    for p in probs:
        print("PROBLEM:", p)
    return 1 if probs else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1] if len(sys.argv) > 1 else "dist"))
