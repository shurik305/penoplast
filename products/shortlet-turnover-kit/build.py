#!/usr/bin/env python3
"""Builds the "Short-Let Turnover Kit" workbook (Airbnb / holiday-let cleaning).

Usage: python3 build.py [outdir]
Sheets: Start Here, Settings, Turnover Checklist (print), Turnover Log, Cleaning Fee Calculator, Restock List,
Issue Report (print). Task list and consumables defaults are shared with printables.py.
"""
import datetime as dt
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "_lib"))
from xlsxkit import Book, cell, col, print_setup, start_here  # noqa: E402

PRODUCT = "Short-Let Turnover Kit"
BRAND = "Hutsol"
VERSION = "1.0"

CHECKLIST = [
    ("On arrival", ["Check for guest belongings left behind — log lost property",
                    "Photograph any damage or mess before you start",
                    "Open windows to air the property",
                    "Strip beds and start the first laundry load (towels first)"]),
    ("Kitchen", ["Wash, dry and put away all dishes; empty the dishwasher",
                 "Clean inside the microwave; wipe oven door and hob",
                 "Remove guest food from fridge (per host policy); wipe shelves",
                 "Wipe worktops, splashback, cupboard fronts and handles",
                 "Clean and polish sink and taps",
                 "Empty coffee machine and kettle; descale if needed",
                 "Empty all bins, wash if dirty, fit new liners",
                 "Check crockery, cutlery and glasses count"]),
    ("Bathrooms", ["Clean and disinfect toilet (bowl, seat, base, behind)",
                   "Clean shower/bath, screen and tiles; clear hair from drains",
                   "Clean basin, taps and mirror",
                   "Replace towels: bath, hand, face and bath mat",
                   "Restock toilet paper (2 spare rolls), soap and toiletries",
                   "Empty bins, fit new liners, mop floor"]),
    ("Bedrooms", ["Check mattress protectors and pillows for stains",
                  "Make beds with fresh linen (hotel-style corners)",
                  "Dust bedside tables, lamps and surfaces",
                  "Check under beds, drawers and wardrobes",
                  "Hangers, spare blankets and pillows in place"]),
    ("Living areas", ["Dust surfaces, TV and shelves; disinfect remotes",
                      "Vacuum sofa and check under cushions",
                      "Plump cushions, fold throws",
                      "Clean glass tables and mirrors",
                      "Arrange furniture to match the listing photos"]),
    ("Floors and touch points", ["Vacuum all floors, rugs and stairs",
                                 "Mop hard floors",
                                 "Wipe light switches, door handles and handrails",
                                 "Spot-clean walls and doors",
                                 "Check every light bulb works"]),
    ("Final checks", ["Restock tea, coffee, sugar, dish soap, sponges and bin bags",
                      "Heating / air-con set to the host's standard",
                      "Windows closed and locked",
                      "Wi-Fi card, house manual and welcome items in place",
                      "Check smoke / CO alarm indicator lights (report faults)",
                      "Count linen sets out and in; laundry started or collected",
                      "Take 'ready' photos of every room",
                      "Lock up, reset key box, message the host"]),
]

# item, unit cost, quantity per stay, scales with (stay/guest/bathroom/bedroom)
CONSUMABLES = [("Toilet paper roll", 0.60, 2, "bathroom"), ("Hand soap refill", 0.80, 1, "bathroom"),
               ("Shampoo (mini)", 0.50, 1, "guest"), ("Conditioner (mini)", 0.50, 1, "guest"),
               ("Body wash (mini)", 0.50, 1, "guest"), ("Coffee pods", 0.40, 2, "guest"),
               ("Tea bags", 0.05, 2, "guest"), ("Sugar sachets", 0.03, 2, "guest"),
               ("Dish soap (portion)", 0.30, 1, "stay"), ("Dishwasher tablets", 0.25, 2, "stay"),
               ("Sponge", 0.30, 1, "stay"), ("Bin liners", 0.10, 4, "stay"), ("Paper towel roll", 1.00, 0.5, "stay"),
               ("Welcome snack", 2.00, 1, "stay"), ("Laundry detergent (per load)", 0.30, 2, "bedroom")]

# name, bedrooms, bathrooms, beds to make, typical guests, pay per turnover, hourly pay, std minutes, laundry/set, sets
PROPERTIES = [("Harbour View Apartment", 2, 1, 2, 4, 55, 0, 150, 8, 6),
              ("Garden Cottage", 3, 2, 4, 6, 0, 20, 210, 8, 10),
              ("City Studio", 0, 1, 1, 2, 40, 0, 90, 8, 3)]


def build(outdir):
    path = os.path.join(outdir, "Short-Let-Turnover-Kit.xlsx")
    b = Book(path)
    wb, f = b.wb, b.f
    wb.set_properties({"title": PRODUCT, "author": BRAND, "company": BRAND,
                       "comments": f"Version {VERSION}. For use in the purchaser's own business. Not for resale."})
    names = ["Start Here", "Settings", "Turnover Checklist", "Turnover Log", "Cleaning Fee Calculator",
             "Restock List", "Issue Report"]
    ws = {n: wb.add_worksheet(n) for n in names}
    f_time = wb.add_format({"font_name": "Arial", "font_size": 10, "num_format": "hh:mm", "locked": False})
    f_box = wb.add_format({"border": 1, "border_color": "#1F5FA6"})
    f_task = wb.add_format({"font_name": "Arial", "font_size": 10, "valign": "vcenter", "locked": False})
    f_line = wb.add_format({"bottom": 1, "bottom_color": "#8895A3", "locked": False, "font_name": "Arial",
                            "font_size": 10})
    f_min = wb.add_format({"num_format": "0", "font_name": "Arial", "font_size": 10})

    # ------------------------------------------------------------------ Settings
    S = "Settings"
    s = ws[S]
    s.hide_gridlines(2)
    s.set_column("A:A", 2)
    s.set_column("B:B", 30)
    s.set_column("C:M", 12)
    s.set_row(0, 30)
    s.write("B1", "Settings", f["title"])
    s.write("B2", "Your properties, pay rates and guest consumables. Yellow cells are yours to change.",
            f["subtitle"])
    b.section(s, 3, "General", 1, 12)
    for r_, (nm, lab, val, fm, note) in enumerate([
            ("CurrencyLabel", "Currency (symbol or code)", "$", "in_text", "Shown in labels."),
            ("SuppliesTurn", "Cleaning supplies per turnover", 3.00, "in_num", "Chemicals, cloths, mop heads."),
            ("RoundTo", "Round cleaning fees up to the nearest", 5, "in_num0", "Use 0 for no rounding.")], start=4):
        s.write(r_, 1, lab, f["label"])
        s.write(r_, 2, val, f[fm])
        s.merge_range(r_, 3, r_, 6, note, f["note"])
        b.name(nm, S, r_, 2)
    b.section(s, 8, "Properties (up to 10)", 1, 12)
    heads = ["Property", "Bedrooms", "Bathrooms", "Beds to make", "Typical guests", "Pay per turnover (fixed)",
             "…or hourly pay", "Standard minutes", "Laundry cost per linen set", "Linen sets owned",
             "Recommended sets (3 per bed)", "Sets short"]
    s.write_row(9, 1, heads, f["th"])
    s.set_row(9, 44)
    p0, p1 = 10, 19
    for i, rr in enumerate(range(p0, p1 + 1)):
        vals = PROPERTIES[i] if i < len(PROPERTIES) else ("",) + ("",) * 9
        s.write(rr, 1, vals[0], f["in_text"])
        for j, v in enumerate(vals[1:], start=2):
            s.write(rr, j, v, f["in_num"] if j in (6, 7, 9) else f["in_num0"])
        s.write_formula(rr, 11, f'=IF({cell(rr, 1)}="","",N({cell(rr, 4)})*3)', f["c_num0"])
        s.write_formula(rr, 12, f'=IF({cell(rr, 1)}="","",MAX(0,{cell(rr, 11)}-N({cell(rr, 10)})))', f["c_num0"])
    for nm, c_ in (("PropNames", 1), ("PropBedrooms", 2), ("PropBaths", 3), ("PropBeds", 4), ("PropGuests", 5),
                   ("PropFixed", 6), ("PropHourly", 7), ("PropMinutes", 8), ("PropLaundry", 9)):
        b.name(nm, S, p0, c_, p1, c_)
    s.merge_range(p1 + 1, 1, p1 + 1, 12, "Pay per turnover: a fixed amount per clean. Leave it 0 to pay by the hour "
                                          "(hourly pay × standard minutes). Linen: keep about 3 sets per bed — one on "
                                          "the bed, one in the laundry, one spare.", f["note"])
    s.set_row(p1 + 1, 26)
    b.section(s, 22, "Guest consumables per stay", 1, 12)
    s.write_row(23, 1, ["Item", "Unit cost", "Qty per stay", "Scales with"], f["th"])
    s.merge_range(23, 5, 23, 8, "stay = once per stay · guest = per guest · bathroom / bedroom = per room", f["note"])
    c0 = 24
    for i, (it, uc, q, sc) in enumerate(CONSUMABLES):
        rr = c0 + i
        s.write(rr, 1, it, f["in_text"]); s.write(rr, 2, uc, f["in_num"]); s.write(rr, 3, q, f["in_dec"])
        s.write(rr, 4, sc, f["in_text"])
    c1 = c0 + len(CONSUMABLES) + 4  # spare rows for the buyer's own items
    for rr in range(c0 + len(CONSUMABLES), c1 + 1):
        s.write(rr, 1, "", f["in_text"]); s.write(rr, 2, "", f["in_num"]); s.write(rr, 3, "", f["in_dec"])
        s.write(rr, 4, "", f["in_text"])
    for nm, c_ in (("ConsItems", 1), ("ConsCost", 2), ("ConsQty", 3), ("ConsScale", 4)):
        b.name(nm, S, c0, c_, c1, c_)
    s.data_validation(c0, 4, c1, 4, {"validate": "list", "source": ["stay", "guest", "bathroom", "bedroom"]})
    s.freeze_panes(2, 0)

    def cons_per_stay(guests, baths, bedrooms):
        return (f'SUMPRODUCT(ConsCost,ConsQty,(ConsScale="stay")+(ConsScale="")+(ConsScale="guest")*{guests}+'
                f'(ConsScale="bathroom")*{baths}+(ConsScale="bedroom")*MAX(1,{bedrooms}))')

    def lookup(rng, key):
        return f"INDEX({rng},MATCH({key},PropNames,0))"

    # ------------------------------------------------------------------ Cleaning Fee Calculator
    F = "Cleaning Fee Calculator"
    fc = ws[F]
    fc.hide_gridlines(2)
    for c_, w in zip(range(10), [2, 34, 16, 14, 14, 14, 14, 14, 16, 3]):
        fc.set_column(c_, c_, w)
    fc.set_row(0, 30)
    fc.write("B1", "Cleaning Fee Calculator", f["title"])
    fc.write("B2", "What a turnover really costs you — and the cleaning fee that covers it.", f["subtitle"])
    b.section(fc, 3, "1. Choose the stay", 1, 3)
    fc.write(4, 1, "Property", f["label"]); fc.merge_range(4, 2, 4, 3, PROPERTIES[0][0], f["in_text"])
    b.name("FProp", F, 4, 2)
    fc.data_validation(4, 2, 4, 2, {"validate": "list", "source": b.lst("PropNames")})
    fc.write(5, 1, "Guests on this stay", f["label"]); fc.write(5, 2, PROPERTIES[0][4], f["in_int"])
    b.name("FGuests", F, 5, 2)
    fc.write(6, 1, "Mark-up on your turnover cost", f["label"]); fc.write(6, 2, 0.15, f["in_pct"])
    b.name("FMarkup", F, 6, 2)
    fc.write(7, 1, "Platform / payment fee taken from the cleaning fee", f["label"]); fc.write(7, 2, 0.03, f["in_pct"])
    b.name("FFee", F, 7, 2)
    fc.data_validation(7, 2, 7, 2, {"validate": "decimal", "criteria": "between", "minimum": 0, "maximum": 0.5,
                                    "error_message": "Use 0% to 50%."})
    fc.data_validation(6, 2, 6, 2, {"validate": "decimal", "criteria": "between", "minimum": 0, "maximum": 3})
    fc.data_validation(5, 2, 5, 2, {"validate": "integer", "criteria": "between", "minimum": 0, "maximum": 30})
    # hidden checks: property still exists in Settings; consumables with an unknown "scales with" value
    fc.set_column(10, 10, 8, None, {"hidden": True})
    fc.write_formula(4, 10, "=ISNUMBER(MATCH(FProp,PropNames,0))")
    b.name("FPropOK", F, 4, 10)
    fc.write_formula(5, 10, '=SUMPRODUCT((ConsItems<>"")*(ConsScale<>"")*(ConsScale<>"stay")*(ConsScale<>"guest")'
                            '*(ConsScale<>"bathroom")*(ConsScale<>"bedroom"))')
    b.name("ConsBad", F, 5, 10)
    fc.merge_range(7, 3, 8, 7, "Enter the percentage your booking platform or payment provider deducts from the "
                               "cleaning fee (0% if none). Check your own platform's current host fees.", f["note"])
    b.section(fc, 10, "2. Turnover cost", 1, 3)
    key = "FProp"
    rows = [("Cleaner", f"=IFERROR(IF(N({lookup('PropFixed', key)})>0,{lookup('PropFixed', key)},"
                        f"N({lookup('PropHourly', key)})*N({lookup('PropMinutes', key)})/60),0)", "FCleaner"),
            ("Laundry (beds × cost per linen set)", f"=IFERROR(N({lookup('PropBeds', key)})*"
                                                    f"N({lookup('PropLaundry', key)}),0)", "FLaundry"),
            ("Guest consumables", "=IFERROR(" + cons_per_stay("N(FGuests)", f"N({lookup('PropBaths', key)})",
                                                              f"N({lookup('PropBedrooms', key)})") + ",0)", "FCons"),
            ("Cleaning supplies", "=SuppliesTurn", "FSupplies")]
    rr = 11
    for lab, fml, nm in rows:
        fc.write(rr, 1, lab, f["label"])
        fc.write_formula(rr, 2, fml, f["c_num"])
        b.name(nm, F, rr, 2)
        rr += 1
    fc.write_formula(rr, 1, '="Total turnover cost ("&CurrencyLabel&")"', f["label_b"])
    fc.write_formula(rr, 2, "=IF(FPropOK,FCleaner+FLaundry+FCons+FSupplies,0)", f["c_num_b"])
    b.name("FTotal", F, rr, 2)
    rr += 2
    b.section(fc, rr, "3. Recommended cleaning fee", 1, 3); rr += 1
    fc.set_row(rr, 34)
    fc.write_formula(rr, 1, '="Cleaning fee to charge ("&CurrencyLabel&")"', f["big_label"])
    fee_row = rr
    fc.write_formula(rr, 2, "=IF(FTotal<=0,0,IF(RoundTo>0,CEILING(FTotal*(1+FMarkup)/(1-MIN(MAX(FFee,0),0.9)),RoundTo),"
                            "FTotal*(1+FMarkup)/(1-MIN(MAX(FFee,0),0.9))))", f["big"])
    b.name("FFeeOut", F, rr, 2)
    rr += 1
    fc.write(rr, 1, "What you keep after the fee and your costs", f["label"])
    fc.write_formula(rr, 2, "=FFeeOut*(1-MIN(MAX(FFee,0),0.9))-FTotal", f["c_num"])
    b.name("FKeep", F, rr, 2)
    fc.merge_range(fee_row, 3, rr, 7, "", f["status"])
    fc.write_formula(fee_row, 3, '=IF(NOT(FPropOK),"⚠ Property not found in Settings — choose it again from the list.",'
                                 'IF(ConsBad>0,"⚠ "&ConsBad&" consumable(s) have an unknown \'scales with\' value — use '
                                 'stay, guest, bathroom or bedroom.",IF(FKeep<0,"⚠ This fee does not cover the '
                                 'turnover cost.","✔ The fee covers the turnover cost and your mark-up.")))', f["status"])
    fc.conditional_format(fee_row, 3, rr, 7, {"type": "formula", "criteria": f'=LEFT($D${fee_row + 1},1)="✔"',
                                              "format": f["green"]})
    fc.conditional_format(fee_row, 3, rr, 7, {"type": "formula", "criteria": f'=LEFT($D${fee_row + 1},1)="⚠"',
                                              "format": f["red"]})
    rr += 2
    b.section(fc, rr, "All properties at a glance (typical number of guests)", 1, 8); rr += 1
    fc.write_row(rr, 1, ["Property", "Guests", "Cleaner", "Laundry", "Consumables", "Supplies", "Total cost",
                         "Recommended fee"], f["th"])
    rr += 1
    g0 = rr
    for i in range(p1 - p0 + 1):
        src = p0 + i
        P = f"'{S}'!{cell(src, 1)}"

        def sref(c_):
            return f"N('{S}'!{cell(src, c_)})"
        cleaner = f"IF({sref(6)}>0,{sref(6)},{sref(7)}*{sref(8)}/60)"
        laundry = f"{sref(4)}*{sref(9)}"
        cons = cons_per_stay(sref(5), sref(3), sref(2))
        fc.write_formula(rr, 1, f'=IF({P}="","",{P})', f["label"])
        fc.write_formula(rr, 2, f'=IF({P}="","",{sref(5)})', f["c_num0"])
        fc.write_formula(rr, 3, f'=IF({P}="","",{cleaner})', f["c_num"])
        fc.write_formula(rr, 4, f'=IF({P}="","",{laundry})', f["c_num"])
        fc.write_formula(rr, 5, f'=IF({P}="","",{cons})', f["c_num"])
        fc.write_formula(rr, 6, f'=IF({P}="","",SuppliesTurn)', f["c_num"])
        tot = f"SUM({cell(rr, 3)}:{cell(rr, 6)})"
        fc.write_formula(rr, 7, f'=IF({P}="","",{tot})', f["c_num_b"])
        fc.write_formula(rr, 8, f'=IF({P}="","",IF(RoundTo>0,CEILING({tot}*(1+FMarkup)/(1-MIN(MAX(FFee,0),0.9)),RoundTo),'
                                f'{tot}*(1+FMarkup)/(1-MIN(MAX(FFee,0),0.9))))', f["c_num_b"])
        rr += 1
    fc.protect("", {"select_locked_cells": True, "select_unlocked_cells": True})

    # ------------------------------------------------------------------ Restock List
    R = "Restock List"
    rs = ws[R]
    rs.hide_gridlines(2)
    for c_, w in zip(range(8), [2, 32, 16, 14, 12, 12, 12, 3]):
        rs.set_column(c_, c_, w)
    rs.set_row(0, 30)
    rs.write("B1", "Restock List", f["title"])
    rs.write("B2", "Count what's in the cupboard — see exactly what to buy before the next stays.", f["subtitle"])
    rs.write(3, 1, "Property", f["label"]); rs.merge_range(3, 2, 3, 3, PROPERTIES[0][0], f["in_text"])
    b.name("RProp", R, 3, 2)
    rs.merge_range(3, 4, 3, 6, "", f["note"])
    rs.write_formula(3, 4, '=IF(ISNUMBER(MATCH(RProp,PropNames,0)),"","⚠ Property not found in Settings — choose it again.")',
                     f["overdue"])
    rs.data_validation(3, 2, 3, 2, {"validate": "list", "source": b.lst("PropNames")})
    rs.write(4, 1, "Stock up for how many stays?", f["label"]); rs.write(4, 2, 4, f["in_int"])
    b.name("RStays", R, 4, 2)
    rs.write(5, 1, "Guests per stay", f["label"])
    rs.write_formula(5, 2, "=IFERROR(INDEX(PropGuests,MATCH(RProp,PropNames,0)),2)", f["in_int"])
    b.name("RGuests", R, 5, 2)
    rs.write_row(7, 1, ["Item", "Needed", "In stock (count)", "To buy", "Unit cost", "Cost"], f["th"])
    rs.set_row(7, 30)
    n = c1 - c0 + 1
    for i in range(n):
        rr = 8 + i
        src = c0 + i
        it = f"'{S}'!{cell(src, 1)}"
        scale = (f"IF('{S}'!{cell(src, 4)}=\"guest\",RGuests,IF('{S}'!{cell(src, 4)}=\"bathroom\","
                 f"IFERROR(INDEX(PropBaths,MATCH(RProp,PropNames,0)),1),IF('{S}'!{cell(src, 4)}=\"bedroom\","
                 f"MAX(1,IFERROR(INDEX(PropBedrooms,MATCH(RProp,PropNames,0)),1)),1)))")
        rs.write_formula(rr, 1, f'=IF({it}="","",{it})', f["label"])
        rs.write_formula(rr, 2, f"=IF({it}=\"\",\"\",ROUNDUP(N('{S}'!{cell(src, 3)})*{scale}*RStays,0))", f["c_num0"])
        rs.write(rr, 3, "", f["in_num0"])
        rs.write_formula(rr, 4, f'=IF({it}="","",MAX(0,{cell(rr, 2)}-N({cell(rr, 3)})))', f["c_num0"])
        rs.write_formula(rr, 5, f"=IF({it}=\"\",\"\",N('{S}'!{cell(src, 2)}))", f["c_num"])
        rs.write_formula(rr, 6, f'=IF({it}="","",{cell(rr, 4)}*{cell(rr, 5)})', f["c_num"])
    last = 8 + n - 1
    rs.write_formula(last + 1, 1, '="Shopping total ("&CurrencyLabel&")"', f["label_b"])
    rs.write_formula(last + 1, 6, f"=SUM({cell(8, 6)}:{cell(last, 6)})", f["c_num_b"])
    b.name("RTotal", R, last + 1, 6)
    rs.conditional_format(8, 4, last, 4, {"type": "cell", "criteria": ">", "value": 0, "format": f["amber"]})
    rs.protect("", {"select_locked_cells": True, "select_unlocked_cells": True})

    # ------------------------------------------------------------------ Turnover Log
    L = "Turnover Log"
    lg = ws[L]
    lg.hide_gridlines(2)
    for c_, w in zip(range(13), [2, 12, 24, 16, 9, 9, 10, 9, 10, 11, 9, 30, 3]):
        lg.set_column(c_, c_, w)
    lg.set_row(0, 30)
    lg.write("B1", "Turnover Log", f["title"])
    lg.write("B2", "One row per clean: minutes and cleaner pay are calculated. Rows 10 onwards.", f["subtitle"])
    first, lastr = 9, 308
    rg = lambda c_: f"{col(c_)}{first + 1}:{col(c_)}{lastr + 1}"  # noqa: E731
    tiles = [("Turnovers logged", f"=COUNT({rg(1)})", "tile_val", (1, 2)),
             ("This month", f'=COUNTIFS({rg(1)},">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),{rg(1)},"<="&TODAY())',
              "tile_val", (3, 3)),
             ("Average minutes", f'=IFERROR(AVERAGE({rg(6)}),0)', "tile_val", (4, 5)),
             ("Cleaner pay (all)", f"=SUM({rg(9)})", "tile_money", (6, 7)),
             ("Unpaid", f'=SUMIF({rg(10)},"No",{rg(9)})', "tile_money", (8, 9)),
             ("Issues reported", f'=COUNTIF({rg(8)},"Yes")', "tile_val", (10, 10)),
             ("Rows to check", f'=SUMPRODUCT(({rg(2)}<>"")*ISNA(MATCH({rg(2)},PropNames,0)))+COUNTIF({rg(6)},">720")',
              "tile_val", (11, 11))]
    for lab, fml, fk, (a, z) in tiles:
        if a == z:
            lg.write(3, a, lab, f["tile_label"]); lg.write_formula(4, a, fml, f[fk])
        else:
            lg.merge_range(3, a, 3, z, lab, f["tile_label"]); lg.merge_range(4, a, 4, z, "", f[fk])
            lg.write_formula(4, a, fml, f[fk])
    lg.set_row(3, 28); lg.set_row(4, 28)
    lg.write_row(8, 1, ["Date", "Property", "Cleaner", "Start", "Finish", "Minutes", "Next guests", "Issues?",
                        "Pay", "Paid?", "Notes"], f["th"])
    lg.set_row(8, 30)
    today = dt.date.today()
    samples = [(today - dt.timedelta(days=3), "Harbour View Apartment", "Anna", dt.time(11, 0), dt.time(13, 25), 4,
                "No", "Yes", "Example row — delete me"),
               (today - dt.timedelta(days=2), "Garden Cottage", "Marek", dt.time(10, 30), dt.time(14, 10), 6, "Yes",
                "No", "Example: red wine stain on rug — see Issue Report"),
               (today - dt.timedelta(days=1), "City Studio", "Anna", dt.time(11, 15), dt.time(12, 40), 2, "No", "No",
                "Example row — delete me")]
    for i, rr in enumerate(range(first, lastr + 1)):
        smp = samples[i] if i < len(samples) else None
        if smp:
            lg.write_datetime(rr, 1, dt.datetime.combine(smp[0], dt.time()), f["log_date"])
            lg.write(rr, 2, smp[1], f["log_text"]); lg.write(rr, 3, smp[2], f["log_text"])
            lg.write_datetime(rr, 4, dt.datetime.combine(dt.date(1899, 12, 31), smp[3]), f_time)
            lg.write_datetime(rr, 5, dt.datetime.combine(dt.date(1899, 12, 31), smp[4]), f_time)
            lg.write(rr, 7, smp[5], f["log_text"]); lg.write(rr, 8, smp[6], f["log_text"])
            lg.write(rr, 10, smp[7], f["log_text"]); lg.write(rr, 11, smp[8], f["log_text"])
        else:
            lg.write_blank(rr, 1, None, f["log_date"])
            for c_ in (2, 3, 7, 8, 10, 11):
                lg.write_blank(rr, c_, None, f["log_text"])
            lg.write_blank(rr, 4, None, f_time)
            lg.write_blank(rr, 5, None, f_time)
        E, Fc, C = cell(rr, 4), cell(rr, 5), cell(rr, 2)
        lg.write_formula(rr, 6, f'=IF(AND(ISNUMBER({E}),ISNUMBER({Fc})),ROUND(MOD({Fc}-{E},1)*1440,0),"")', f_min)
        lg.write_formula(rr, 9, f'=IF({C}="","",IFERROR(IF(N(INDEX(PropFixed,MATCH({C},PropNames,0)))>0,'
                                f'INDEX(PropFixed,MATCH({C},PropNames,0)),N(INDEX(PropHourly,MATCH({C},PropNames,0)))*'
                                f'IF(N({cell(rr, 6)})>0,{cell(rr, 6)},N(INDEX(PropMinutes,MATCH({C},PropNames,0))))/60)'
                                f',""))', f["log_num"])
    lg.data_validation(first, 2, lastr, 2, {"validate": "list", "source": b.lst("PropNames")})
    lg.data_validation(first, 8, lastr, 8, {"validate": "list", "source": ["Yes", "No"]})
    lg.data_validation(first, 10, lastr, 10, {"validate": "list", "source": ["Yes", "No"]})
    lg.conditional_format(first, 8, lastr, 8, {"type": "cell", "criteria": "==", "value": '"Yes"', "format": f["red"]})
    lg.conditional_format(first, 2, lastr, 2, {"type": "formula", "criteria": f'=AND($C{first + 1}<>"",'
                                                                           f'ISNA(MATCH($C{first + 1},PropNames,0)))',
                                               "format": f["red"]})
    lg.conditional_format(first, 6, lastr, 6, {"type": "cell", "criteria": ">", "value": 720, "format": f["red"]})
    lg.conditional_format(first, 10, lastr, 10, {"type": "cell", "criteria": "==", "value": '"No"',
                                                 "format": f["amber"]})
    lg.freeze_panes(first, 0)

    # ------------------------------------------------------------------ Turnover Checklist (printable)
    C = "Turnover Checklist"
    ck = ws[C]
    ck.hide_gridlines(2)
    for c_, w in zip(range(5), [2, 4, 64, 18, 2]):
        ck.set_column(c_, c_, w)
    ck.set_row(0, 28)
    ck.merge_range(0, 1, 0, 3, "Turnover Checklist", f["title"])
    hdr = ["Property", "Date", "Cleaner", "Guest checked out", "Next check-in / guests"]
    rr = 2
    for lab in hdr:
        ck.merge_range(rr, 1, rr, 2, "", f["label"])
        ck.write(rr, 1, lab + ":", f["label_b"])
        ck.write(rr, 3, PROPERTIES[0][0] if lab == "Property" else "", f_line)
        rr += 1
    ck.data_validation(2, 3, 2, 3, {"validate": "list", "source": b.lst("PropNames")})
    rr += 1
    for sec, tasks in CHECKLIST:
        ck.merge_range(rr, 1, rr, 3, sec, f["section"]); rr += 1
        for t in tasks:
            ck.write_blank(rr, 1, None, f_box)
            ck.merge_range(rr, 2, rr, 3, t, f_task)
            ck.set_row(rr, 17)
            rr += 1
        rr += 1
    ck.merge_range(rr, 1, rr, 3, "Restock needed / issues found", f["section"]); rr += 1
    for _ in range(4):
        ck.merge_range(rr, 1, rr, 3, "", f_line); ck.set_row(rr, 20); rr += 1
    rr += 1
    ck.merge_range(rr, 1, rr, 3, "Finished at: ________   Cleaner signature: _______________________", f["label"])
    ck_last = rr
    ck.set_paper(9)
    ck.set_portrait()
    ck.set_margins(left=0.5, right=0.5, top=0.5, bottom=0.5)
    ck.print_area(0, 0, ck_last, 4)
    ck.fit_to_pages(1, 0)

    # ------------------------------------------------------------------ Issue Report (printable)
    I_ = "Issue Report"
    ir = ws[I_]
    ir.hide_gridlines(2)
    for c_, w in zip(range(5), [2, 30, 24, 30, 2]):
        ir.set_column(c_, c_, w)
    ir.set_row(0, 28)
    ir.merge_range(0, 1, 0, 3, "Damage & Issue Report", f["title"])
    ir.merge_range(1, 1, 1, 3, "Record it the same day, with photos — hosts and platforms usually need evidence "
                               "quickly to support a claim.", f["subtitle"])
    fields = ["Property", "Date found", "Found by", "Booking / reservation reference", "Guest check-out date",
              "Issue type (damage / missing item / maintenance / cleanliness / other)", "Room / location",
              "Photos taken (how many, file names)", "Estimated repair / replacement cost", "Reported to host on",
              "Action taken"]
    rr = 3
    for fl in fields:
        ir.write(rr, 1, fl, f["label_b"])
        ir.merge_range(rr, 2, rr, 3, "", f_line)
        ir.set_row(rr, 22)
        rr += 1
    ir.write(3, 2, PROPERTIES[0][0], f_line)
    ir.data_validation(3, 2, 3, 2, {"validate": "list", "source": b.lst("PropNames")})
    rr += 1
    ir.merge_range(rr, 1, rr, 3, "Description", f["section"]); rr += 1
    for _ in range(8):
        ir.merge_range(rr, 1, rr, 3, "", f_line); ir.set_row(rr, 22); rr += 1
    rr += 1
    ir.merge_range(rr, 1, rr, 3, "Signature: ______________________________     Date: ______________", f["label"])
    ir.set_paper(9)
    ir.set_portrait()
    ir.set_margins(left=0.6, right=0.6, top=0.6, bottom=0.6)
    ir.print_area(0, 0, rr, 4)
    ir.fit_to_pages(1, 1)

    # ------------------------------------------------------------------ Start Here
    steps = [("Settings", "Add your properties (bedrooms, bathrooms, beds, typical guests), how you pay cleaners "
                          "(per turnover or per hour), laundry cost per linen set and your guest consumables."),
             ("Turnover Checklist", "Choose the property at the top and print the checklist for each clean — or "
                                    "use the blank printable PDF included in your download."),
             ("Turnover Log", "Log every clean: start and finish times give minutes and cleaner pay automatically; "
                              "mark issues and whether the cleaner has been paid."),
             ("Cleaning Fee Calculator", "See the real cost of a turnover (cleaner, laundry, consumables, supplies) "
                                         "and the cleaning fee that covers it, for every property."),
             ("Restock List", "Count what's in the cupboard and get a shopping list for the next stays."),
             ("Issue Report", "Record damage or missing items the same day, with photos, for the host or a claim.")]
    tips = ["Keep three linen sets per bed: one on the bed, one in the laundry, one spare.",
            "Photograph every room before and after the clean — it protects you and the host.",
            "Strip the beds and start the laundry first; it is always the bottleneck.",
            "Time a few turnovers and update the standard minutes — pay and fees depend on it.",
            "Recheck your cleaning fee whenever laundry, wages or consumables go up."]
    start_here(b, ws["Start Here"], PRODUCT, "Works with any currency", VERSION, BRAND, steps, tips, "VAT/sales tax",
               extra_note="Included in your download: this workbook, a blank printable Turnover Checklist (PDF) and a "
                          "printable Damage & Issue Report (PDF).")
    print_setup(ws["Settings"], True)
    print_setup(ws["Cleaning Fee Calculator"], False)
    print_setup(ws["Restock List"], False)
    print_setup(ws["Turnover Log"], True, (0, 0, 60, 12))
    print_setup(ws["Start Here"], False)
    wb.close()
    return path


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(__file__), "dist")
    os.makedirs(out, exist_ok=True)
    print(build(out))
