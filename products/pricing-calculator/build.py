#!/usr/bin/env python3
"""Builds the "Cleaning Price & Quote Calculator" workbook in several regional editions.

Usage: python3 build.py [outdir]
All formulas are plain Excel functions that also work in Google Sheets and LibreOffice.
"""
import os
import sys
import datetime as dt

import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell, xl_col_to_name

PRODUCT = "Cleaning Price & Quote Calculator"
BRAND = "Hutsol"
VERSION = "1.0"

# Production rates are m² per cleaner-hour; they are starting points the buyer is told to calibrate.
SERVICES = [
    ("Regular clean", 35, "Dust reachable surfaces; vacuum and mop all floors; kitchen worktops, hob, sink and "
     "appliance fronts; clean and disinfect bathrooms (toilet, basin, shower/bath, mirrors); empty bins; tidy."),
    ("Deep clean", 14, "Everything in a regular clean plus skirting boards, doors and frames, light switches and "
     "sockets, inside microwave, descaling of taps and shower screens, cobwebs, radiators, window sills and frames "
     "(inside), under and behind movable furniture."),
    ("End of tenancy / move-out", 11, "Deep clean of an empty property plus inside all cupboards and drawers, oven, "
     "hob and extractor, fridge/freezer (defrosted), interior windows and frames, spot-cleaning of walls, limescale "
     "removal, balconies swept."),
    ("After builders", 8, "Removal of construction dust from all surfaces incl. walls and ledges, paint/plaster "
     "spots removed from floors and glass, HEPA vacuuming, interior windows, sockets and switches, multiple "
     "wipe-downs until dust-free."),
    ("Office regular clean", 170, "Empty bins and recycling; dust and wipe clear desks and surfaces; kitchen/break "
     "area; clean and restock toilets; vacuum and mop floors; entrance glass; disinfect touch points (handles, "
     "switches, rails)."),
    ("Office deep clean", 50, "Office regular clean plus high-level dusting, chairs and upholstery vacuumed, skirting, "
     "glass partitions, inside kitchen appliances, machine scrub of hard floors where agreed."),
    ("Short-let / holiday-let turnover", 28, "Strip and remake beds with fresh linen, fresh towels, full kitchen and bathroom "
     "clean, dishes, restock consumables from the host list, empty bins, floors, check for damage and lost property "
     "with photos, short report to the host."),
]

CONDITIONS = [("Light", 0.85), ("Normal", 1.0), ("Heavy", 1.3), ("Very heavy", 1.6)]

# name, discount, visits per month (0 = not recurring)
FREQUENCIES = [
    ("One-off", 0.0, 0), ("Weekly", 0.15, 4.33), ("Fortnightly", 0.10, 2.17), ("Every 4 weeks", 0.05, 1.08),
    ("Monthly", 0.05, 1.0), ("Twice a week", 0.18, 8.67), ("Weekdays (Mon-Fri)", 0.22, 21.67),
]

# name, minutes per unit, fixed price per unit (0 = time-based only)
EXTRAS = [
    ("Inside oven", 60, 0), ("Inside fridge / freezer", 30, 0), ("Inside kitchen cupboards", 45, 0),
    ("Interior window (each)", 6, 0), ("Change bed linen (per bed)", 10, 0), ("Ironing (per hour)", 60, 0),
    ("Laundry load (wash, dry, fold)", 15, 0), ("Balcony / terrace", 20, 0), ("Blinds (each)", 10, 0),
    ("Wall spot-cleaning (per room)", 15, 0), ("Pet hair surcharge", 30, 0), ("Extra bathroom (beyond first)", 30, 0),
    ("Carpet / upholstery (per room, fixed)", 0, 35), ("Waste removal (per bag)", 5, 3),
]

ROOMS_M2 = [("Studio", 30), ("1 bed", 50), ("2 bed", 70), ("3 bed", 90), ("4 bed", 115), ("5 bed", 140),
            ("6+ bed", 170)]
ROOMS_M2_US = [("Studio", 45), ("1 bed", 65), ("2 bed", 95), ("3 bed", 140), ("4 bed", 190), ("5 bed", 240),
               ("6+ bed", 300)]

PRESETS = {
    "UK": dict(
        edition="UK edition (GBP, VAT)", currency="GBP", tax_name="VAT", tax=0.20, registered="Yes", incl="Yes",
        area="m²", dist="miles", pay=13.00, oncost=0.20, supplies=1.20, overhead=3.50, margin=0.25,
        min_charge=55, min_hours=2, free_radius=5, travel_rate=0.45, round_to=5,
        pay_terms="Payment is due within 7 days of each clean by bank transfer.",
        biz="Sparkle & Co Cleaning", contact="07700 900123 · hello@example.com", web="www.example.com",
        area_served="Bristol and surrounding areas", client="Sarah Jones", addr="12 Harbour Road, Bristol BS1",
        email="sarah.jones@example.com", rooms=ROOMS_M2, bands=[30, 50, 70, 90, 110, 130, 150, 200, 250, 300],
        fixed=[("Public liability insurance", 30), ("Vehicle (lease, fuel, servicing)", 260),
               ("Phone & internet", 30), ("Booking / invoicing software", 25), ("Marketing & advertising", 60),
               ("Accountant & bank fees", 45), ("Equipment replacement", 30), ("Other", 20)],
        hours_week=30, weeks=46,
    ),
    "EU": dict(
        edition="EU & Ireland edition (EUR, VAT)", currency="EUR", tax_name="VAT", tax=0.21, registered="Yes",
        incl="Yes", area="m²", dist="km", pay=14.00, oncost=0.25, supplies=1.30, overhead=3.80, margin=0.25,
        min_charge=60, min_hours=2, free_radius=10, travel_rate=0.40, round_to=5,
        pay_terms="Payment is due within 8 days of the invoice date by bank transfer.",
        biz="Clean Point Services", contact="+31 20 000 0000 · info@example.com", web="www.example.com",
        area_served="Utrecht and surrounding area", client="Anna de Vries", addr="Oudegracht 120, Utrecht",
        email="anna@example.com", rooms=ROOMS_M2, bands=[30, 50, 70, 90, 110, 130, 150, 200, 250, 300],
        fixed=[("Business insurance", 35), ("Vehicle (lease, fuel, servicing)", 300), ("Phone & internet", 35),
               ("Booking / invoicing software", 25), ("Marketing & advertising", 60),
               ("Accountant & bank fees", 50), ("Equipment replacement", 30), ("Other", 20)],
        hours_week=30, weeks=46,
    ),
    "AU": dict(
        edition="Australia & NZ edition (AUD, GST)", currency="AUD", tax_name="GST", tax=0.10, registered="Yes",
        incl="Yes", area="m²", dist="km", pay=32.00, oncost=0.15, supplies=2.00, overhead=6.00, margin=0.25,
        min_charge=120, min_hours=2, free_radius=15, travel_rate=0.90, round_to=5,
        pay_terms="Payment is due within 7 days of each clean by bank transfer.",
        biz="Harbour Clean Co", contact="0491 570 156 · hello@example.com", web="www.example.com",
        area_served="Newcastle and Lake Macquarie", client="Emma Wilson", addr="8 Ocean Street, Merewether NSW",
        email="emma@example.com", rooms=ROOMS_M2, bands=[30, 50, 70, 90, 110, 130, 150, 200, 250, 300],
        fixed=[("Public liability insurance", 60), ("Vehicle (lease, fuel, servicing)", 450),
               ("Phone & internet", 60), ("Booking / invoicing software", 40), ("Marketing & advertising", 100),
               ("Accountant & bank fees", 80), ("Equipment replacement", 50), ("Other", 30)],
        hours_week=30, weeks=46,
    ),
    "US": dict(
        edition="US & Canada edition (USD, sq ft)", currency="USD", tax_name="Sales tax", tax=0.0, registered="No",
        incl="No", area="sq ft", dist="miles", pay=18.00, oncost=0.15, supplies=1.50, overhead=4.00, margin=0.25,
        min_charge=90, min_hours=2, free_radius=10, travel_rate=0.70, round_to=5,
        pay_terms="Payment is due on the day of service by card or bank transfer.",
        biz="Bright Home Cleaning", contact="(555) 010-2030 · hello@example.com", web="www.example.com",
        area_served="Austin, TX and nearby", client="Jessica Miller", addr="1402 Oak Hollow Dr, Austin TX",
        email="jessica@example.com", rooms=ROOMS_M2_US,
        bands=[500, 750, 1000, 1250, 1500, 1750, 2000, 2500, 3000, 3500],
        fixed=[("Liability insurance & bonding", 80), ("Vehicle (lease, gas, maintenance)", 450),
               ("Phone & internet", 60), ("Booking / invoicing software", 40), ("Marketing & advertising", 120),
               ("Accountant & bank fees", 60), ("Equipment replacement", 40), ("Other", 30)],
        hours_week=30, weeks=48,
    ),
}

for _p in PRESETS.values():
    _p["hours_week"] = round(sum(v for _, v in _p["fixed"]) / _p["overhead"] * 12 / _p["weeks"])
SQFT = 10.7639

NAVY, BLUE, LIGHT, INPUT, INPUT_BORDER = "#17324D", "#1F5FA6", "#EAF1FA", "#FFF4CC", "#D9A400"
GREY, LINE = "#5B6B7B", "#C9D6E3"


class Book:
    def __init__(self, path):
        self.wb = xlsxwriter.Workbook(path)
        self.f = {}
        self.refs = {}
        base = {"font_name": "Arial", "font_size": 10, "valign": "vcenter"}

        def fmt(name, **kw):
            d = dict(base)
            d.update(kw)
            self.f[name] = self.wb.add_format(d)

        fmt("title", font_size=20, bold=True, font_color=NAVY)
        fmt("subtitle", font_size=10, font_color=GREY, italic=True)
        fmt("section", bold=True, font_color="white", bg_color=BLUE, font_size=11)
        fmt("th", bold=True, font_color=NAVY, bg_color=LIGHT, bottom=1, bottom_color=LINE, text_wrap=True)
        fmt("label", text_wrap=True)
        fmt("label_b", bold=True)
        fmt("note", font_size=9, font_color=GREY, italic=True, text_wrap=True, valign="top")
        inp = dict(bg_color=INPUT, border=1, border_color=INPUT_BORDER, locked=False)
        fmt("in_text", **inp)
        fmt("in_wrap", text_wrap=True, valign="top", **inp)
        fmt("in_num", num_format="#,##0.00", **inp)
        fmt("in_num0", num_format="#,##0", **inp)
        fmt("in_int", num_format="0", **inp)
        fmt("in_dec", num_format="0.00", **inp)
        fmt("in_pct", num_format="0%", **inp)
        fmt("in_date", num_format="dd mmm yyyy", align="left", **inp)
        calc = dict(bg_color=LIGHT, border=1, border_color=LINE)
        fmt("c_num", num_format="#,##0.00", **calc)
        fmt("c_num0", num_format="#,##0", **calc)
        fmt("c_dec1", num_format="0.0", **calc)
        fmt("c_dec2", num_format="0.00", **calc)
        fmt("c_pct", num_format="0%", **calc)
        fmt("c_text", text_wrap=True, **calc)
        fmt("c_num_b", num_format="#,##0.00", bold=True, **calc)
        fmt("big", num_format="#,##0.00", bold=True, font_size=22, font_color=NAVY, bg_color=LIGHT, border=1,
            border_color=LINE, align="right")
        fmt("big_label", bold=True, font_size=11, font_color=NAVY, text_wrap=True)
        fmt("status", text_wrap=True, bold=True, border=1, border_color=LINE, valign="vcenter")
        fmt("step_no", bold=True, font_color="white", bg_color=BLUE, align="center", font_size=12, border=2, border_color="white")
        fmt("step", text_wrap=True, valign="vcenter")
        fmt("step_b", text_wrap=True, valign="vcenter", bold=True, font_color=NAVY)
        fmt("bullet", align="center", valign="vcenter", font_color=BLUE, bold=True)
        fmt("legend_in", bg_color=INPUT, border=1, border_color=INPUT_BORDER)
        fmt("legend_c", bg_color=LIGHT, border=1, border_color=LINE)
        # client quote formats
        fmt("q_biz", font_size=18, bold=True, font_color=NAVY)
        fmt("q_title", font_size=20, bold=True, font_color=BLUE, align="right")
        fmt("q_small", font_size=9, font_color=GREY)
        fmt("q_small_r", font_size=9, font_color=GREY, align="right")
        fmt("q_val_r", align="right")
        fmt("q_date_r", align="right", num_format="dd mmm yyyy")
        fmt("q_cap", font_size=8, bold=True, font_color=GREY)
        fmt("q_client", bold=True, font_size=11)
        fmt("q_desc", bold=True, font_size=11, font_color=NAVY, text_wrap=True)
        fmt("q_th", bold=True, font_color="white", bg_color=NAVY)
        fmt("q_th_r", bold=True, font_color="white", bg_color=NAVY, align="right")
        fmt("q_row", bottom=1, bottom_color=LINE, text_wrap=True)
        fmt("q_amt", bottom=1, bottom_color=LINE, num_format='#,##0.00;-#,##0.00;"–"')
        fmt("q_sub", bold=True, top=1, top_color=NAVY)
        fmt("q_sub_amt", bold=True, top=1, top_color=NAVY, num_format="#,##0.00")
        fmt("q_total", bold=True, font_size=13, font_color="white", bg_color=BLUE)
        fmt("q_total_amt", bold=True, font_size=13, font_color="white", bg_color=BLUE, num_format="#,##0.00")
        fmt("q_h", bold=True, font_color=NAVY, font_size=10, top=1, top_color=LINE)
        fmt("q_body", text_wrap=True, valign="top", font_size=9)
        fmt("q_thanks", italic=True, font_color=BLUE, font_size=11)
        fmt("log_date", num_format="dd mmm yyyy", locked=False)
        fmt("log_num", num_format="#,##0.00", locked=False)
        fmt("log_text", locked=False)
        fmt("tile_label", font_size=9, font_color=GREY, bg_color=LIGHT, text_wrap=True)
        fmt("tile_val", bold=True, font_size=14, font_color=NAVY, bg_color=LIGHT, num_format="#,##0")
        fmt("tile_money", bold=True, font_size=14, font_color=NAVY, bg_color=LIGHT, num_format="#,##0.00")
        fmt("tile_pct", bold=True, font_size=14, font_color=NAVY, bg_color=LIGHT, num_format="0%")
        fmt("green", bg_color="#E3F4E8", font_color="#1E6B37")
        fmt("red", bg_color="#FDE7E7", font_color="#9B1C1C")
        fmt("amber", bg_color="#FFF1D6", font_color="#8A5A00")
        fmt("greyed", font_color="#8895A3")
        fmt("overdue", font_color="#C62828", bold=True)

    def name(self, nm, sheet, r, c, r2=None, c2=None):
        ref = f"'{sheet}'!{xl_rowcol_to_cell(r, c, True, True)}"
        if r2 is not None:
            ref += f":{xl_rowcol_to_cell(r2, c2, True, True)}"
        self.wb.define_name(nm, "=" + ref)
        self.refs[nm] = ref
        return ref

    def lst(self, nm):
        """Direct range reference for data-validation lists (Google Sheets imports these more reliably)."""
        return "=" + self.refs[nm]


def cell(r, c):
    return xl_rowcol_to_cell(r, c)


def build(preset_key, outdir):
    p = PRESETS[preset_key]
    fname = f"Cleaning-Price-Quote-Calculator_{preset_key}.xlsx"
    path = os.path.join(outdir, fname)
    b = Book(path)
    wb, f = b.wb, b.f
    wb.set_properties({"title": f"{PRODUCT} – {p['edition']}", "author": BRAND, "company": BRAND,
                       "comments": f"Version {VERSION}. For use in the purchaser's own business. Not for resale."})

    names = ["Start Here", "Settings", "Quote Calculator", "Client Quote", "Price List", "Break-even", "Quote Log"]
    ws = {n: wb.add_worksheet(n) for n in names}
    tax = p["tax_name"]

    # ------------------------------------------------------------------ Settings
    s = ws["Settings"]
    s.hide_gridlines(2)
    s.set_column("A:A", 2)
    s.set_column("B:B", 46)
    s.set_column("C:C", 18)
    s.set_column("D:D", 70)
    s.set_row(0, 30)
    s.write("B1", "Settings", f["title"])
    s.write("B2", "Fill in the yellow cells once. Every other sheet uses these numbers.", f["subtitle"])
    S = "Settings"
    r = 3

    def section(ws_, row, text, last_col=3):
        ws_.merge_range(row, 1, row, last_col, text, f["section"])
        ws_.set_row(row, 20)

    def setting(nm, label, value, fm, note="", formula=False):
        nonlocal r
        s.write(r, 1, label, f["label"])
        if formula:
            s.write_formula(r, 2, value, fm)
        else:
            s.write(r, 2, value, fm)
        if note:
            s.write(r, 3, note, f["note"])
        b.name(nm, S, r, 2)
        r += 1

    section(s, r, "Your business (shown on the client quote)"); r += 1
    setting("BizName", "Business name", p["biz"], f["in_text"])
    setting("BizContact", "Phone · email", p["contact"], f["in_text"])
    setting("BizWeb", "Website", p["web"], f["in_text"])
    setting("BizArea", "Area you serve", p["area_served"], f["in_text"])
    setting("ValidDays", "Quotes are valid for (days)", 30, f["in_int"])
    setting("PayTerms", "Payment terms (printed on the quote)", p["pay_terms"], f["in_text"])
    r += 1
    section(s, r, f"Currency, {tax} and units"); r += 1
    setting("CurrencyLabel", "Currency (code or symbol)", p["currency"], f["in_text"], "Shown in labels and on the quote.")
    setting("VATRate", f"{tax} rate", p["tax"], f["in_pct"], "Use your country's standard rate for cleaning services.")
    setting("VATReg", f"Are you {tax} registered?", p["registered"], f["in_text"],
            f"Choose No if you do not charge {tax} — prices are then calculated without it.")
    setting("InclVAT", f"Show client prices including {tax}?", p["incl"], f["in_text"],
            f"Yes = rounding is applied to the {tax}-inclusive price (typical for households).")
    setting("AreaUnit", "Area unit you measure in", p["area"], f["in_text"], "m² or sq ft. Maths runs in m² internally.")
    setting("DistUnit", "Distance unit", p["dist"], f["in_text"], "km or miles (label only).")
    setting("VATUsed", f"{tax} rate applied (calculated)", f'=IF(VATReg="Yes",VATRate,0)', f["c_pct"], formula=True)
    # validations for the yes/no + unit cells (rows are known from order above)
    s.data_validation(r - 5, 2, r - 5, 2, {"validate": "list", "source": ["Yes", "No"]})
    s.data_validation(r - 4, 2, r - 4, 2, {"validate": "list", "source": ["Yes", "No"]})
    s.data_validation(r - 3, 2, r - 3, 2, {"validate": "list", "source": ["m²", "sq ft"]})
    s.data_validation(r - 2, 2, r - 2, 2, {"validate": "list", "source": ["km", "miles"]})
    r += 1
    section(s, r, "Costs and margin (excl. " + tax + ")"); r += 1
    setting("PayRate", "Cleaner pay per hour (or the hourly pay you want yourself)", p["pay"], f["in_num"],
            "Owner-operator? Enter the hourly wage you want to take home before tax.")
    setting("OnCost", "Employer on-costs (holiday pay, pension, insurance, payroll taxes)", p["oncost"], f["in_pct"],
            "Solo self-employed with no staff: use 10–15% to cover your own holidays and sick days.")
    setting("LabourCost", "True labour cost per hour (calculated)", "=PayRate*(1+OnCost)", f["c_num"], formula=True)
    setting("Supplies", "Cleaning supplies per labour hour", p["supplies"], f["in_num"],
            "Chemicals, cloths, bags, mop heads. Total your monthly spend ÷ labour hours.")
    setting("Overhead", "Overheads per labour hour", p["overhead"], f["in_num"],
            "Insurance, vehicle, phone, software, marketing. Use the Break-even sheet to calculate it.")
    margin_row = r
    setting("Margin", "Target profit margin (% of the net price)", p["margin"], f["in_pct"],
            "20–30% is common for small cleaning companies. Profit pays for growth, tax and bad months.")
    s.data_validation(margin_row, 2, margin_row, 2, {"validate": "decimal", "criteria": "between", "minimum": 0,
                                                      "maximum": 0.95, "error_message": "Use 0% to 95%."})
    setting("CostPerHr", "Total cost per labour hour (calculated)", "=LabourCost+Supplies+Overhead", f["c_num"],
            formula=True)
    setting("RatePerHr", "Your price per labour hour at target margin (calculated)",
            "=CostPerHr/(1-MIN(MAX(Margin,0),0.95))", f["c_num_b"], formula=True)
    r += 1
    section(s, r, "Minimums, travel and rounding"); r += 1
    setting("MinCharge", f"Minimum charge per visit (excl. {tax})", p["min_charge"], f["in_num"],
            "Protects you on very small jobs.")
    setting("MinHours", "Minimum labour hours per visit", p["min_hours"], f["in_dec"])
    setting("FreeRadius", "Free travel radius (one way)", p["free_radius"], f["in_num0"],
            "No travel charge inside this distance.")
    setting("TravelRate", "Travel charge per distance unit beyond the radius", p["travel_rate"], f["in_num"],
            "Charged for the return trip (distance beyond radius × 2).")
    setting("RoundTo", "Round client prices up to the nearest", p["round_to"], f["in_num0"],
            "e.g. 5 → 83.40 becomes 85. Use 0 for no rounding.")
    r += 1

    # Service table
    section(s, r, "Service types and production rates  —  time 3 of your own jobs and adjust!"); r += 1
    s.write(r, 1, "Service type", f["th"])
    s.write_formula(r, 2, '="Area per cleaner-hour ("&AreaUnit&")"', f["th"])
    s.write(r, 3, "What's included (printed on the quote)", f["th"])
    s.set_row(r, 30)
    r += 1
    svc_r0 = r
    for nm_, rate, scope in SERVICES:
        s.write(r, 1, nm_, f["in_text"])
        s.write(r, 2, rate if p["area"] == "m²" else round(rate * SQFT / 5) * 5, f["in_num0"])
        s.write(r, 3, scope, f["in_wrap"])
        s.set_row(r, 54)
        r += 1
    svc_r1 = r - 1
    b.name("ServiceNames", S, svc_r0, 1, svc_r1, 1)
    b.name("ServiceRates", S, svc_r0, 2, svc_r1, 2)
    b.name("ServiceScopes", S, svc_r0, 3, svc_r1, 3)
    s.write(r, 1, "Rates are the floor area (in your area unit) one cleaner completes per hour in normal condition, e.g. "
                  "90 m² ÷ 14 m² per hour = 6.4 labour hours. Renamed an item? Choose it again on Quote Calculator — "
                  "the calculator warns you when a name doesn't match.", f["note"])
    s.set_row(r, 26)
    r += 2

    section(s, r, "Condition of the property"); r += 1
    s.write_row(r, 1, ["Condition", "Time multiplier"], f["th"]); r += 1
    cond_r0 = r
    for nm_, mult in CONDITIONS:
        s.write(r, 1, nm_, f["in_text"]); s.write(r, 2, mult, f["in_dec"]); r += 1
    b.name("CondNames", S, cond_r0, 1, r - 1, 1)
    b.name("CondMults", S, cond_r0, 2, r - 1, 2)
    r += 1

    section(s, r, "Frequency discounts"); r += 1
    s.write_row(r, 1, ["Frequency", "Discount", "Visits per month (0 = one-off)"], f["th"]); r += 1
    fr0 = r
    for nm_, disc, visits in FREQUENCIES:
        s.write(r, 1, nm_, f["in_text"]); s.write(r, 2, disc, f["in_pct"]); s.write(r, 3, visits, f["in_dec"])
        r += 1
    b.name("FreqNames", S, fr0, 1, r - 1, 1)
    b.name("FreqDiscs", S, fr0, 2, r - 1, 2)
    b.name("FreqVisits", S, fr0, 3, r - 1, 3)
    s.write(r, 1, "Give recurring discounts only for committed schedules — they fill your diary and cut travel "
                  "and set-up time.", f["note"])
    r += 2

    section(s, r, "Extras and add-ons"); r += 1
    s.write_row(r, 1, ["Extra", "Minutes per unit",
                       f"Fixed price per unit (optional, excl. {tax}) — used instead of time"], f["th"]); r += 1
    ex0 = r
    for nm_, mins, fixed in EXTRAS:
        s.write(r, 1, nm_, f["in_text"]); s.write(r, 2, mins, f["in_num0"]); s.write(r, 3, fixed, f["in_num"])
        r += 1
    b.name("ExtraNames", S, ex0, 1, r - 1, 1)
    b.name("ExtraMins", S, ex0, 2, r - 1, 2)
    b.name("ExtraFixed", S, ex0, 3, r - 1, 3)
    r += 1

    section(s, r, "Bedrooms → approximate floor area (used when the area is unknown)"); r += 1
    s.write_row(r, 1, ["Property", "Typical area (m²)", "Same in sq ft (for reference)"], f["th"]); r += 1
    rm0 = r
    for nm_, a in p["rooms"]:
        s.write(r, 1, nm_, f["in_text"]); s.write(r, 2, a, f["in_num0"])
        s.write_formula(r, 3, f"={cell(r, 2)}*10.7639", f["c_num0"])
        r += 1
    b.name("RoomNames", S, rm0, 1, r - 1, 1)
    b.name("RoomAreas", S, rm0, 2, r - 1, 2)
    s.freeze_panes(2, 0)

    # ------------------------------------------------------------------ Quote Calculator
    Q = "Quote Calculator"
    q = ws[Q]
    q.hide_gridlines(2)
    q.set_column("A:A", 2); q.set_column("B:B", 36); q.set_column("C:C", 24); q.set_column("D:D", 3)
    q.set_column("E:E", 40); q.set_column("F:F", 18); q.set_column("G:G", 3)
    q.set_column("H:J", 12, None, {"hidden": True})
    q.set_row(0, 30)
    q.merge_range("B1:F1", "Quote Calculator", f["title"])
    q.merge_range("B2:F2", "Fill in the yellow cells — the price and the printable client quote update instantly.",
                  f["subtitle"])

    left = {}

    def qin(row, key, label, value, fm, formula=False, label_formula=False):
        if label_formula:
            q.write_formula(row, 1, label, f["label"])
        else:
            q.write(row, 1, label, f["label"])
        if formula:
            q.write_formula(row, 2, value, fm)
        else:
            q.write(row, 2, value, fm)
        b.name(key, Q, row, 2)
        left[key] = row

    section(q, 3, "1. Client and job", 2)
    qin(4, "QNum", "Quote number", "Q-1001", f["in_text"])
    qin(5, "QDate", "Quote date", "=TODAY()", f["in_date"], formula=True)
    qin(6, "QClient", "Client name", p["client"], f["in_text"])
    qin(7, "QAddr", "Address / postcode", p["addr"], f["in_text"])
    qin(8, "QContact", "Client email / phone", p["email"], f["in_text"])
    section(q, 10, "2. The clean", 2)
    qin(11, "QService", "Service type", "Deep clean", f["in_text"])
    qin(12, "QArea", '="Floor area ("&AreaUnit&") — leave blank to use bedrooms"',
        p["bands"][3] if p["area"] == "m²" else 1500, f["in_num0"], label_formula=True)
    qin(13, "QBeds", "…or number of bedrooms (if area unknown)", "3 bed", f["in_text"])
    qin(14, "QCond", "Condition", "Normal", f["in_text"])
    qin(15, "QFreq", "Frequency", "One-off", f["in_text"])
    qin(16, "QCrew", "Cleaners on site", 2, f["in_int"])
    qin(17, "QDist", '="Travel distance, one way ("&DistUnit&")"', 9, f["in_num0"], label_formula=True)
    qin(18, "QAdj", "Manual adjustment % (+ stairs/parking, − easy job)", 0, f["in_pct"])
    qin(19, "QPromo", "Promotional discount %", 0, f["in_pct"])
    section(q, 21, "3. Extras (optional)", 2)
    q.write_row(22, 1, ["Extra", "Quantity"], f["th"])
    q.write(22, 7, "mins", f["note"]); q.write(22, 8, "fixed", f["note"])
    ex_defaults = [("Inside oven", 1), ("Inside fridge / freezer", 1), ("Interior window (each)", 8), ("", ""),
                   ("", ""), ("", "")]
    ex_rows = list(range(23, 29))
    for rr, (nm_, qty) in zip(ex_rows, ex_defaults):
        q.write(rr, 1, nm_, f["in_text"])
        q.write(rr, 2, qty, f["in_int"])
        B_, C_ = cell(rr, 1), cell(rr, 2)
        q.write_formula(rr, 7, f'=IFERROR(IF(AND({B_}<>"",N({C_})>0),INDEX(ExtraMins,MATCH({B_},ExtraNames,0))*{C_},0),0)')
        q.write_formula(rr, 8, f'=IFERROR(IF(AND({B_}<>"",N({C_})>0),INDEX(ExtraFixed,MATCH({B_},ExtraNames,0))*{C_},0),0)')
        q.write_formula(rr, 9, f'=IF(AND({B_}<>"",N({C_})>0),IF(ISNUMBER(MATCH({B_},ExtraNames,0)),0,1),0)')
    # extras list text for the client quote
    parts = "&".join(f'IF(AND({cell(rr, 1)}<>"",N({cell(rr, 2)})>0,ISNUMBER(MATCH({cell(rr, 1)},ExtraNames,0))),'
                     f'{cell(rr, 1)}&" × "&{cell(rr, 2)}&"; ","")' for rr in ex_rows)
    q.write_formula(29, 7, "=" + parts)
    q.write_formula(29, 8, f'=IF({cell(29, 7)}="","None",LEFT({cell(29, 7)},LEN({cell(29, 7)})-2))')
    b.name("ExtrasList", Q, 29, 8)
    # lookup checks (hidden): anything typed that no longer matches Settings is flagged in the status
    for rr_, nm, fml in ((31, "SvcOK", "=ISNUMBER(MATCH(QService,ServiceNames,0))"),
                         (32, "CondOK", "=ISNUMBER(MATCH(QCond,CondNames,0))"),
                         (33, "FreqOK", "=ISNUMBER(MATCH(QFreq,FreqNames,0))"),
                         (34, "ExtrasBad", f"=SUM({cell(23, 9)}:{cell(28, 9)})")):
        q.write_formula(rr_, 7, fml)
        b.name(nm, Q, rr_, 7)
    q.merge_range("B31:C33", "Tip: time your next three jobs and update the production rates in Settings. "
                             "Accurate quotes come from your own numbers, not from guesses.", f["note"])

    # validations
    q.data_validation(left["QService"], 2, left["QService"], 2, {"validate": "list", "source": b.lst("ServiceNames")})
    q.data_validation(left["QBeds"], 2, left["QBeds"], 2, {"validate": "list", "source": b.lst("RoomNames")})
    q.data_validation(left["QCond"], 2, left["QCond"], 2, {"validate": "list", "source": b.lst("CondNames")})
    q.data_validation(left["QFreq"], 2, left["QFreq"], 2, {"validate": "list", "source": b.lst("FreqNames")})
    q.data_validation(ex_rows[0], 1, ex_rows[-1], 1, {"validate": "list", "source": b.lst("ExtraNames")})
    q.data_validation(left["QCrew"], 2, left["QCrew"], 2,
                      {"validate": "integer", "criteria": "between", "minimum": 1, "maximum": 20})
    q.data_validation(left["QAdj"], 2, left["QAdj"], 2,
                      {"validate": "decimal", "criteria": "between", "minimum": -0.9, "maximum": 2})
    q.data_validation(left["QPromo"], 2, left["QPromo"], 2,
                      {"validate": "decimal", "criteria": "between", "minimum": 0, "maximum": 0.9})
    q.data_validation(left["QDist"], 2, left["QDist"], 2,
                      {"validate": "decimal", "criteria": ">=", "value": 0, "ignore_blank": True})
    q.data_validation(left["QArea"], 2, left["QArea"], 2,
                      {"validate": "decimal", "criteria": ">=", "value": 0, "ignore_blank": True})

    # right panel — workings first (names are referenced by the result cells)
    W = {}

    def wrow(row, key, label, formula, fm, label_formula=False):
        if label_formula:
            q.write_formula(row, 4, label, f["label"])
        else:
            q.write(row, 4, label, f["label"])
        q.write_formula(row, 5, formula, fm)
        b.name(key, Q, row, 5)
        W[key] = row

    q.merge_range(27, 4, 27, 5, "How the price is built", f["section"])
    wrow(28, "AreaM2", "Floor area used for the maths (m²)",
         '=IF(AND(ISNUMBER(QArea),N(QArea)>0),IF(AreaUnit="sq ft",QArea/10.7639,QArea),'
         'IFERROR(INDEX(RoomAreas,MATCH(QBeds,RoomNames,0)),0))', f["c_num0"])
    wrow(29, "ProdRate", "Production rate used (m² per cleaner-hour)",
         f"=IFERROR(INDEX(ServiceRates,MATCH(QService,ServiceNames,0))/IF(AreaUnit=\"sq ft\",{SQFT},1),0)", f["c_num0"])
    wrow(30, "CondMult", "Condition multiplier", "=IFERROR(INDEX(CondMults,MATCH(QCond,CondNames,0)),1)", f["c_dec2"])
    wrow(31, "BaseHrs", "Base cleaning hours", "=IF(AND(ProdRate>0,AreaM2>0),AreaM2/ProdRate*CondMult,0)", f["c_dec2"])
    wrow(32, "ExtraHrs", "Extras hours", f"=SUM({cell(23, 7)}:{cell(28, 7)})/60", f["c_dec2"])
    wrow(33, "LabHrs", "Labour hours (minimum applied)",
         "=IF(BaseHrs+ExtraHrs>0,MAX(MinHours,BaseHrs+ExtraHrs),0)", f["c_dec2"])
    wrow(34, "CostTotal", "Cost: labour + supplies + overheads", "=LabHrs*CostPerHr", f["c_num"])
    wrow(35, "PriceMargin", "Price at your target margin", "=CostTotal/(1-MIN(MAX(Margin,0),0.95))", f["c_num"])
    wrow(36, "FixedExtras", "Fixed-price extras", f"=SUM({cell(23, 8)}:{cell(28, 8)})", f["c_num"])
    wrow(37, "FreqDiscAmt", "Frequency discount",
         "=-(PriceMargin+FixedExtras)*IFERROR(INDEX(FreqDiscs,MATCH(QFreq,FreqNames,0)),0)", f["c_num"])
    wrow(38, "AdjAmt", "Manual adjustment", "=(PriceMargin+FixedExtras+FreqDiscAmt)*N(QAdj)", f["c_num"])
    wrow(39, "PromoAmt", "Promotional discount", "=-(PriceMargin+FixedExtras+FreqDiscAmt+AdjAmt)*N(QPromo)",
         f["c_num"])
    wrow(40, "TravelChg", "Travel charge (return trip beyond free radius)",
         "=MAX(0,N(QDist)-FreeRadius)*2*TravelRate", f["c_num"])
    wrow(41, "Subtotal0", "Subtotal before minimum and rounding",
         "=PriceMargin+FixedExtras+FreqDiscAmt+AdjAmt+PromoAmt+TravelChg", f["c_num"])
    wrow(42, "AfterMin", "After minimum charge", "=IF(Subtotal0>0,MAX(MinCharge,Subtotal0),0)", f["c_num"])
    wrow(43, "FinalNet", f"Final net price after rounding (excl. {tax})",
         '=IF(AfterMin<=0,0,IF(InclVAT="Yes",IF(RoundTo>0,CEILING(AfterMin*(1+VATUsed),RoundTo),'
         'AfterMin*(1+VATUsed))/(1+VATUsed),IF(RoundTo>0,CEILING(AfterMin,RoundTo),AfterMin)))', f["c_num_b"])
    wrow(44, "Visits", "Visits per month", "=IFERROR(INDEX(FreqVisits,MATCH(QFreq,FreqNames,0)),0)", f["c_dec2"])

    # results panel
    q.merge_range(3, 4, 3, 5, "Price", f["section"])
    q.set_row(4, 34)
    q.write_formula(4, 4, f'="Price to client ("&IF(VATUsed=0,"no {tax}",IF(InclVAT="Yes","incl. {tax}","excl. {tax}"))'
                          f'&", "&CurrencyLabel&")"', f["big_label"])
    q.write_formula(5, 4, '=IF(QFreq="One-off","One-off clean",QFreq&" — price per visit")', f["note"])
    q.write_formula(6, 4, f'="Net price (excl. {tax})"', f["label"])
    q.write_formula(6, 5, "=ROUND(FinalNet,2)", f["c_num"])
    b.name("NetPrice", Q, 6, 5)
    q.write_formula(7, 4, f'="{tax} ("&TEXT(VATUsed,"0%")&")"', f["label"])
    q.write_formula(8, 4, f'="Total incl. {tax}"', f["label_b"])
    q.write_formula(8, 5, "=ROUND(FinalNet*(1+VATUsed),2)", f["c_num_b"])
    b.name("Gross", Q, 8, 5)
    q.write_formula(7, 5, "=Gross-NetPrice", f["c_num"])
    b.name("VATAmt", Q, 7, 5)
    q.write_formula(4, 5, '=IF(InclVAT="Yes",Gross,NetPrice)', f["big"])
    q.write(9, 4, f"Monthly value if recurring (incl. {tax})", f["label"])
    q.write_formula(9, 5, "=IF(Visits>0,Gross*Visits,0)", f["c_num"])
    b.name("Monthly", Q, 9, 5)

    q.merge_range(11, 4, 11, 5, "Job details", f["section"])
    q.write_formula(12, 4, '="Area ("&AreaUnit&")"', f["label"])
    q.write_formula(12, 5, '=IF(AreaUnit="sq ft",AreaM2*10.7639,AreaM2)', f["c_num0"])
    q.write(13, 4, "Total labour hours (all cleaners)", f["label"])
    q.write_formula(13, 5, "=LabHrs", f["c_dec1"])
    q.write(14, 4, "Time on site with your team (hours)", f["label"])
    q.write_formula(14, 5, "=IF(LabHrs>0,LabHrs/MAX(1,N(QCrew)),0)", f["c_dec1"])
    b.name("OnSite", Q, 14, 5)
    q.write_formula(15, 4, f'="Price per "&AreaUnit&" (excl. {tax})"', f["label"])
    q.write_formula(15, 5, '=IF(AreaM2>0,FinalNet/IF(AreaUnit="sq ft",AreaM2*10.7639,AreaM2),0)', f["c_num"])
    q.write(16, 4, f"Effective rate per labour hour (excl. {tax})", f["label"])
    q.write_formula(16, 5, "=IF(LabHrs>0,FinalNet/LabHrs,0)", f["c_num"])
    b.name("EffRate", Q, 16, 5)

    q.merge_range(18, 4, 18, 5, "Profit check", f["section"])
    q.write(19, 4, "Labour cost", f["label"])
    q.write_formula(19, 5, "=LabHrs*LabourCost", f["c_num"])
    q.write(20, 4, "Supplies + overheads", f["label"])
    q.write_formula(20, 5, "=LabHrs*(Supplies+Overhead)", f["c_num"])
    q.write(21, 4, "Profit on this job", f["label_b"])
    q.write_formula(21, 5, f"=FinalNet-{cell(19, 5)}-{cell(20, 5)}", f["c_num_b"])
    b.name("ProfitJob", Q, 21, 5)
    q.write(22, 4, "Profit margin", f["label"])
    q.write_formula(22, 5, "=IF(FinalNet>0,ProfitJob/FinalNet,0)", f["c_pct"])
    b.name("JobMargin", Q, 22, 5)
    q.merge_range(23, 4, 25, 5, "", f["status"])
    q.write_formula(23, 4, '=IF(AreaM2=0,"Enter the floor area or choose the number of bedrooms.",'
                           'IF(NOT(SvcOK),"⚠ Service type not found in Settings — choose it again from the list.",'
                           'IF(NOT(CondOK),"⚠ Condition not found in Settings — choose it again from the list.",'
                           'IF(NOT(FreqOK),"⚠ Frequency not found in Settings — choose it again from the list.",'
                           'IF(ExtrasBad>0,"⚠ "&ExtrasBad&" extra(s) not found in Settings (renamed?) — choose them '
                           'again; they are not priced.",'
                           'IF(FinalNet<=0,"⚠ Discounts reduce the price to zero — check adjustment and promotion.",'
                           'IF(EffRate<BERate,"⚠ Below your break-even rate of "&FIXED(BERate,2)&" per labour hour — '
                           'check the Break-even sheet.",'
                           'IF(JobMargin<MIN(MAX(Margin,0),0.95)-0.005,"△ Covers your costs, but the margin is "'
                           '&TEXT(JobMargin,"0%")&" — below your "&TEXT(Margin,"0%")&" target (discounts?).",'
                           '"✔ Meets your target margin ("&TEXT(JobMargin,"0%")&")."))))))))', f["status"])
    q.conditional_format(23, 4, 25, 5, {"type": "formula", "criteria": '=LEFT($E$24,1)="✔"', "format": f["green"]})
    q.conditional_format(23, 4, 25, 5, {"type": "formula", "criteria": '=LEFT($E$24,1)="⚠"', "format": f["red"]})
    q.conditional_format(23, 4, 25, 5, {"type": "formula", "criteria": '=LEFT($E$24,1)="△"', "format": f["amber"]})
    q.freeze_panes(3, 0)
    q.protect("", {"select_locked_cells": True, "select_unlocked_cells": True})

    # ------------------------------------------------------------------ Client Quote
    cq = ws["Client Quote"]
    cq.hide_gridlines(2)
    cq.set_column("A:A", 2); cq.set_column("B:B", 52); cq.set_column("C:C", 14); cq.set_column("D:D", 18)
    cq.set_column("E:E", 2)
    cq.set_row(0, 28)
    cq.write_formula(0, 1, "=BizName", f["q_biz"])
    cq.merge_range(0, 2, 0, 3, "QUOTE", f["q_title"])
    cq.write_formula(1, 1, "=BizContact", f["q_small"])
    cq.write_formula(2, 1, "=BizWeb", f["q_small"])
    cq.write_formula(3, 1, "=BizArea", f["q_small"])
    cq.write(5, 1, "PREPARED FOR", f["q_cap"])
    cq.write(5, 2, "Quote no.", f["q_small_r"]); cq.write_formula(5, 3, "=QNum", f["q_val_r"])
    cq.write_formula(6, 1, "=QClient", f["q_client"])
    cq.write(6, 2, "Date", f["q_small_r"]); cq.write_formula(6, 3, "=QDate", f["q_date_r"])
    cq.write_formula(7, 1, "=QAddr", f["label"])
    cq.write(7, 2, "Valid until", f["q_small_r"]); cq.write_formula(7, 3, "=QDate+ValidDays", f["q_date_r"])
    cq.write_formula(8, 1, "=QContact", f["label"])
    cq.set_row(10, 30)
    cq.merge_range(10, 1, 10, 3, "", f["q_desc"])
    cq.write_formula(10, 1, '=QService&" — approx. "&FIXED(IF(AreaUnit="sq ft",AreaM2*10.7639,AreaM2),0)&" "'
                            '&AreaUnit&", "&LOWER(QCond)&" condition"', f["q_desc"])
    cq.merge_range(11, 1, 11, 3, "", f["q_small"])
    cq.write_formula(11, 1, '=IF(QFreq="One-off","One-off clean",QFreq&" clean, price per visit")&" · "&MAX(1,N(QCrew))'
                            '&" cleaner(s), about "&FIXED(OnSite,1)&" hours on site"', f["q_small"])
    cq.write(13, 1, "Description", f["q_th"])
    cq.write(13, 2, "", f["q_th"])
    cq.write_formula(13, 3, '="Amount ("&CurrencyLabel&")"', f["q_th_r"])
    lines = [
        ('="Cleaning service — "&QService', "=ROUND(PriceMargin,2)"),
        ('="Extras at fixed prices"', "=ROUND(FixedExtras,2)"),
        ('="Frequency discount ("&QFreq&")"', "=ROUND(FreqDiscAmt,2)"),
        ('="Adjustment"', "=ROUND(AdjAmt,2)"),
        ('="Promotional discount"', "=ROUND(PromoAmt,2)"),
        ('="Travel"', "=ROUND(TravelChg,2)"),
    ]
    rr = 14
    for lab, amt in lines:
        cq.merge_range(rr, 1, rr, 2, "", f["q_row"])
        cq.write_formula(rr, 1, lab, f["q_row"])
        cq.write_formula(rr, 3, amt, f["q_amt"])
        rr += 1
    # balancing line so the printed lines always add up exactly to the subtotal
    cq.merge_range(rr, 1, rr, 2, "", f["q_row"])
    cq.write_formula(rr, 1, '="Minimum charge / rounding"', f["q_row"])
    cq.write_formula(rr, 3, f"=NetPrice-SUM({cell(14, 3)}:{cell(rr - 1, 3)})", f["q_amt"])
    rr += 1
    cq.merge_range(rr, 1, rr, 2, "", f["q_sub"])
    cq.write_formula(rr, 1, f'=IF(VATUsed>0,"Subtotal (excl. {tax})","Subtotal")', f["q_sub"])
    cq.write_formula(rr, 3, "=NetPrice", f["q_sub_amt"]); rr += 1
    cq.merge_range(rr, 1, rr, 2, "", f["label"])
    cq.write_formula(rr, 1, f'=IF(VATUsed>0,"{tax} "&TEXT(VATUsed,"0%"),"")', f["label"])
    cq.write_formula(rr, 3, '=IF(VATUsed>0,VATAmt,"")', f["q_amt"]); rr += 1
    cq.set_row(rr, 22)
    cq.merge_range(rr, 1, rr, 2, "TOTAL", f["q_total"])
    cq.write_formula(rr, 3, "=Gross", f["q_total_amt"]); rr += 1
    cq.merge_range(rr, 1, rr, 3, "", f["q_small"])
    cq.write_formula(rr, 1, f'=IF(Visits>0,"Recurring: about "&FIXED(Visits,1)&" visits per month ≈ "'
                            f'&FIXED(Monthly,2)&" "&CurrencyLabel&" per month"&IF(VATUsed>0," incl. {tax}","")&".","")', f["q_small"])
    rr += 2
    cq.merge_range(rr, 1, rr, 3, "What's included", f["q_h"]); rr += 1
    cq.set_row(rr, 58)
    cq.merge_range(rr, 1, rr, 3, "", f["q_body"])
    cq.write_formula(rr, 1, '=IFERROR(INDEX(ServiceScopes,MATCH(QService,ServiceNames,0)),"")', f["q_body"]); rr += 1
    cq.merge_range(rr, 1, rr, 3, "Extras included", f["q_h"]); rr += 1
    cq.set_row(rr, 30)
    cq.merge_range(rr, 1, rr, 3, "", f["q_body"])
    cq.write_formula(rr, 1, "=ExtrasList", f["q_body"]); rr += 1
    cq.merge_range(rr, 1, rr, 3, "Terms", f["q_h"]); rr += 1
    cq.set_row(rr, 58)
    cq.merge_range(rr, 1, rr, 3, "", f["q_body"])
    cq.write_formula(rr, 1, f'=PayTerms&" Prices are in "&CurrencyLabel&IF(VATUsed>0,", {tax} shown separately","")&". '
                            'This quote is based on the information you gave us. If the size or condition of the '
                            'property is noticeably different on the day, we will agree any change with you before '
                            'we start."', f["q_body"]); rr += 2
    cq.write(rr, 1, "Accepted by (name and signature): ____________________________", f["label"])
    cq.merge_range(rr, 2, rr, 3, "Date: _______________", f["label"]); rr += 2
    cq.write(rr, 1, "Thank you for your enquiry — we look forward to working with you!", f["q_thanks"])
    cq.set_paper(9)
    cq.set_portrait()
    cq.set_margins(left=0.5, right=0.5, top=0.6, bottom=0.6)
    cq.print_area(0, 0, rr, 4)
    cq.fit_to_pages(1, 1)
    cq.protect("", {"select_locked_cells": True, "select_unlocked_cells": True})

    # ------------------------------------------------------------------ Price List
    pl = ws["Price List"]
    pl.hide_gridlines(2)
    pl.set_column("A:A", 2); pl.set_column("B:B", 16); pl.set_column("C:I", 15)
    pl.set_row(0, 30)
    pl.merge_range("B1:I1", "Price List Builder", f["title"])
    pl.merge_range("B2:I2", "", f["subtitle"])
    pl.write_formula(1, 1, f'="One-off clean, normal condition, 1 visit — prices "&IF(VATUsed=0,"with no {tax}",IF(InclVAT="Yes",'
                           f'"incl. {tax}","excl. {tax}"))&", in "&CurrencyLabel&". Minimum charge and rounding applied. '
                           'Change the sizes in the yellow cells."', f["subtitle"])
    pl.write_formula(3, 1, '="Size ("&AreaUnit&")"', f["th"])
    rate_row = 15
    for j in range(len(SERVICES)):
        col = 2 + j
        pl.write_formula(3, col, f"=INDEX(ServiceNames,{j + 1})", f["th"])
        pl.write_formula(rate_row, col, f"=INDEX(ServiceRates,{j + 1})", f["c_num0"])
    pl.set_row(3, 42)
    pl.write_formula(rate_row, 1, '=AreaUnit&" per cleaner-hour"', f["note"])

    def price_expr(size_ref, rate_ref, disc="0"):
        # size and rate are both in the buyer's area unit, so size / rate = labour hours
        net0 = f"MAX(MinCharge,MAX(MinHours,{size_ref}/{rate_ref})*CostPerHr/(1-MIN(MAX(Margin,0),0.95))*(1-{disc}))"
        return (f'IF(AND(N({size_ref})>0,N({rate_ref})>0),IF(InclVAT="Yes",IF(RoundTo>0,CEILING({net0}*(1+VATUsed),'
                f'RoundTo),{net0}*(1+VATUsed)),IF(RoundTo>0,CEILING({net0},RoundTo),{net0})),"")')

    for i, band in enumerate(p["bands"]):
        row = 4 + i
        pl.write(row, 1, band, f["in_num0"])
        for j in range(len(SERVICES)):
            col = 2 + j
            size_ref = xl_rowcol_to_cell(row, 1, col_abs=True)
            rate_ref = xl_rowcol_to_cell(rate_row, col, row_abs=True)
            pl.write_formula(row, col, "=" + price_expr(size_ref, rate_ref), f["c_num"])
    # recurring table
    r2 = 18
    pl.merge_range(r2, 1, r2, 8, "Recurring prices per visit", f["section"])
    pl.write(r2 + 1, 1, "Service:", f["label_b"])
    pl.merge_range(r2 + 1, 2, r2 + 1, 3, "Regular clean", f["in_text"])
    pl.data_validation(r2 + 1, 2, r2 + 1, 2, {"validate": "list", "source": b.lst("ServiceNames")})
    svc_cell = xl_rowcol_to_cell(r2 + 1, 2, True, True)
    rate_cell = xl_rowcol_to_cell(r2 + 1, 5, True, True)
    pl.write_formula(r2 + 1, 4, '=AreaUnit&" per cleaner-hour"', f["note"])
    pl.write_formula(r2 + 1, 5, f"=IFERROR(INDEX(ServiceRates,MATCH({svc_cell},ServiceNames,0)),0)", f["c_num0"])
    pl.write_formula(r2 + 2, 1, '="Size ("&AreaUnit&")"', f["th"])
    for j in range(5):
        pl.write_formula(r2 + 2, 2 + j, f"=INDEX(FreqNames,{j + 1})", f["th"])
    pl.set_row(r2 + 2, 30)
    for i in range(len(p["bands"])):
        row = r2 + 3 + i
        pl.write_formula(row, 1, f"={xl_rowcol_to_cell(4 + i, 1)}", f["c_num0"])
        for j in range(5):
            size_ref = xl_rowcol_to_cell(row, 1, col_abs=True)
            disc = f"INDEX(FreqDiscs,{j + 1})"
            pl.write_formula(row, 2 + j, "=" + price_expr(size_ref, rate_cell, disc), f["c_num"])
    pl.protect("", {"select_locked_cells": True, "select_unlocked_cells": True})

    # ------------------------------------------------------------------ Break-even
    be = ws["Break-even"]
    be.hide_gridlines(2)
    be.set_column("A:A", 2); be.set_column("B:B", 48); be.set_column("C:C", 18); be.set_column("D:D", 62)
    be.set_row(0, 30)
    be.write("B1", "Break-even & Hourly Rate Check", f["title"])
    be.write("B2", "The minimum you must earn per labour hour to cover every cost — and whether your prices do.",
             f["subtitle"])
    section(be, 3, f"Monthly fixed costs (excl. {tax})")
    rr = 4
    fx0 = rr
    for lab, amt in p["fixed"]:
        be.write(rr, 1, lab, f["in_text"]); be.write(rr, 2, amt, f["in_num"]); rr += 1
    be.write(rr, 1, "Total fixed costs per month", f["label_b"])
    be.write_formula(rr, 2, f"=SUM({cell(fx0, 2)}:{cell(rr - 1, 2)})", f["c_num_b"])
    b.name("FixedMonth", "Break-even", rr, 2); rr += 2
    section(be, rr, "Your capacity"); rr += 1
    be.write(rr, 1, "Paid cleaning hours per week (whole team)", f["label"])
    be.write(rr, 2, p["hours_week"], f["in_num0"]); hw = cell(rr, 2); rr += 1
    be.write(rr, 1, "Working weeks per year", f["label"])
    be.write(rr, 2, p["weeks"], f["in_num0"]); wk = cell(rr, 2); rr += 1
    be.write(rr, 1, "Paid cleaning hours per month", f["label"])
    be.write_formula(rr, 2, f"={hw}*{wk}/12", f["c_num0"]); b.name("HoursMonth", "Break-even", rr, 2); rr += 2
    section(be, rr, "Results"); rr += 1
    be.write(rr, 1, "Fixed costs per labour hour", f["label"])
    be.write_formula(rr, 2, "=IF(HoursMonth>0,FixedMonth/HoursMonth,0)", f["c_num"])
    b.name("FixedPerHr", "Break-even", rr, 2); rr += 1
    be.write(rr, 1, "Labour cost per hour (from Settings)", f["label"])
    be.write_formula(rr, 2, "=LabourCost", f["c_num"]); rr += 1
    be.write(rr, 1, "Supplies per hour (from Settings)", f["label"])
    be.write_formula(rr, 2, "=Supplies", f["c_num"]); rr += 1
    be.write(rr, 1, f"Break-even rate per labour hour (excl. {tax})", f["label_b"])
    be.write_formula(rr, 2, "=LabourCost+Supplies+FixedPerHr", f["c_num_b"])
    b.name("BERate", "Break-even", rr, 2); rr += 1
    be.write(rr, 1, "Target rate with your profit margin", f["label"])
    be.write_formula(rr, 2, "=BERate/(1-MIN(MAX(Margin,0),0.95))", f["c_num"]); rr += 1
    be.write(rr, 1, "Revenue needed per month just to break even", f["label"])
    be.write_formula(rr, 2, "=BERate*HoursMonth", f["c_num"]); rr += 1
    be.write(rr, 1, "Overheads per labour hour currently in Settings", f["label"])
    be.write_formula(rr, 2, "=Overhead", f["c_num"])
    be.write_formula(rr, 3, '=IF(ABS(Overhead-FixedPerHr)>0.5,"Tip: set Settings → Overheads per labour hour to "'
                            '&FIXED(FixedPerHr,2)&" to match your real fixed costs.","✔ Settings match your fixed '
                            'costs.")', f["note"]); rr += 1
    be.write(rr, 1, "Your price per labour hour from Settings", f["label"])
    be.write_formula(rr, 2, "=RatePerHr", f["c_num"])
    be.write_formula(rr, 3, '=IF(RatePerHr<BERate,"⚠ Your settings price below break-even — raise your margin or '
                            'cut costs.","✔ Your settings price above break-even.")', f["note"]); rr += 2
    be.merge_range(rr, 1, rr + 1, 3, "Owner-operator? Put the hourly wage you want to take home as 'Cleaner pay per "
                                      "hour' in Settings, so break-even already includes paying yourself.", f["note"])
    be.protect("", {"select_locked_cells": True, "select_unlocked_cells": True})

    # ------------------------------------------------------------------ Quote Log
    lg = ws["Quote Log"]
    lg.hide_gridlines(2)
    widths = [2, 11, 13, 22, 24, 10, 16, 12, 12, 12, 12, 13, 34]
    for i, w in enumerate(widths):
        lg.set_column(i, i, w)
    lg.set_row(0, 30)
    lg.write("B1", "Quote Log & Win Rate", f["title"])
    lg.write("B2", "Log every quote you send. Status: Sent → Follow-up → Won or Lost. Rows 10 onwards.", f["subtitle"])
    first, last = 9, 208
    rng = lambda col: f"{xl_col_to_name(col)}{first + 1}:{xl_col_to_name(col)}{last + 1}"
    tiles = [
        ("Quotes logged", f"=COUNTA({rng(1)})", "tile_val"),
        ("Won", f'=COUNTIF({rng(10)},"Won")', "tile_val"),
        ("Lost", f'=COUNTIF({rng(10)},"Lost")', "tile_val"),
        ("Win rate", f'=IF(COUNTIF({rng(10)},"Won")+COUNTIF({rng(10)},"Lost")>0,COUNTIF({rng(10)},"Won")/'
                     f'(COUNTIF({rng(10)},"Won")+COUNTIF({rng(10)},"Lost")),0)', "tile_pct"),
        (f"Won monthly value (incl. {tax})", f'=SUMIF({rng(10)},"Won",{rng(9)})', "tile_money"),
        (f"Open pipeline (incl. {tax})", f'=SUMIF({rng(10)},"Sent",{rng(8)})+SUMIF({rng(10)},"Follow-up",{rng(8)})',
         "tile_money"),
        ("Follow-ups overdue", f'=COUNTIFS({rng(10)},"Sent",{rng(11)},"<"&TODAY())+COUNTIFS({rng(10)},"Follow-up",'
                               f'{rng(11)},"<"&TODAY())', "tile_val"),
    ]
    # tiles: (first_col, last_col) blocks laid over the log columns
    tile_spans = [(1, 2), (3, 3), (4, 4), (5, 6), (7, 8), (9, 10), (11, 12)]
    for (lab, fml, fk), (c0, c1) in zip(tiles, tile_spans):
        if c0 == c1:
            lg.write(3, c0, lab, f["tile_label"])
            lg.write_formula(4, c0, fml, f[fk])
        else:
            lg.merge_range(3, c0, 3, c1, lab, f["tile_label"])
            lg.merge_range(4, c0, 4, c1, "", f[fk])
            lg.write_formula(4, c0, fml, f[fk])
    lg.set_row(3, 28); lg.set_row(4, 28)
    heads = ["Quote no.", "Date", "Client", "Service", "Area", "Frequency", f"Net price", f"Incl. {tax}",
             "Monthly value", "Status", "Follow-up", "Notes"]
    lg.write_row(8, 1, heads, f["th"])
    lg.set_row(8, 30)
    today = dt.date.today()
    samples = [
        ("Q-1001", today - dt.timedelta(days=12), "Example: Sarah Jones", "Deep clean", 90, "One-off", 250, 300, 0,
         "Won", None, "Example row — delete me"),
        ("Q-1002", today - dt.timedelta(days=5), "Example: Mark Evans", "Regular clean", 70, "Fortnightly", 70.83, 85,
         184.45, "Sent", today + dt.timedelta(days=2), "Example row — delete me"),
        ("Q-1003", today - dt.timedelta(days=20), "Example: Lee Office Ltd", "Office regular clean", 400, "Weekly",
         112.5, 135, 584.55, "Lost", None, "Example row — delete me"),
    ]
    for i in range(first, last + 1):
        for col in range(1, 13):
            fm = f["log_text"]
            if col in (2, 11):
                fm = f["log_date"]
            elif col in (7, 8, 9):
                fm = f["log_num"]
            lg.write_blank(i, col, None, fm)
    for i, row in enumerate(samples):
        for col, v in enumerate(row, start=1):
            if v is None:
                continue
            fm = f["log_text"]
            if col in (2, 11):
                fm = f["log_date"]
                lg.write_datetime(first + i, col, dt.datetime.combine(v, dt.time()), fm)
                continue
            if col in (7, 8, 9):
                fm = f["log_num"]
            lg.write(first + i, col, v, fm)
    lg.data_validation(first, 4, last, 4, {"validate": "list", "source": b.lst("ServiceNames")})
    lg.data_validation(first, 6, last, 6, {"validate": "list", "source": b.lst("FreqNames")})
    lg.data_validation(first, 10, last, 10, {"validate": "list", "source": ["Sent", "Follow-up", "Won", "Lost"]})
    stat = f"$K{first + 1}"
    lg.conditional_format(first, 10, last, 10, {"type": "cell", "criteria": "==", "value": '"Won"', "format": f["green"]})
    lg.conditional_format(first, 10, last, 10, {"type": "cell", "criteria": "==", "value": '"Lost"', "format": f["red"]})
    lg.conditional_format(first, 10, last, 10,
                          {"type": "cell", "criteria": "==", "value": '"Follow-up"', "format": f["amber"]})
    lg.conditional_format(first, 11, last, 11,
                          {"type": "formula", "criteria": f'=AND($L{first + 1}<>"",$L{first + 1}<TODAY(),'
                                                          f'OR({stat}="Sent",{stat}="Follow-up"))',
                           "format": f["overdue"]})
    lg.freeze_panes(9, 0)

    # ------------------------------------------------------------------ Start Here
    st = ws["Start Here"]
    st.hide_gridlines(2)
    st.set_column("A:A", 2); st.set_column("B:B", 5); st.set_column("C:C", 100)
    st.set_row(0, 34)
    st.write("B1", PRODUCT, f["title"])
    st.write("B2", f"{p['edition']} · version {VERSION} · by {BRAND}", f["subtitle"])
    steps = [
        ("Settings", f"Enter your business details, currency, {tax}, your labour costs and target margin. "
                     "Yellow cells are yours to change; blue cells calculate automatically."),
        ("Production rates", "Still on Settings: check the area per cleaner-hour for each service. Time your next "
                             "3 jobs and adjust — this single step makes every quote accurate."),
        ("Quote Calculator", "For each enquiry enter the area (or bedrooms), service, condition, frequency, number of "
                             "cleaners, travel distance and any extras. The price, hours and profit update instantly."),
        ("Client Quote", "A print-ready A4 quote is filled in for you. Print it or save as PDF "
                         "(File → Print → Save as PDF) and send it to your client."),
        ("Quote Log", "Record every quote and its status to see your win rate, pipeline and overdue follow-ups."),
        ("Price List", "Generate a standard price table by property size for your website, flyers or social media."),
        ("Break-even", "Enter your monthly fixed costs and hours to find the minimum rate you must charge — the "
                       "calculator warns you if a quote falls below it."),
    ]
    st.write("B4", "How to use it (5 minutes)", f["label_b"])
    r = 4
    for i, (hd, tx) in enumerate(steps, start=1):
        st.set_row(r, 36)
        st.write(r, 1, i, f["step_no"])
        st.write_rich_string(r, 2, f["step_b"], hd + ": ", f["step"], tx, f["step"])
        r += 1
    r += 1
    st.write(r, 1, "", f["legend_in"]); st.write(r, 2, "Yellow = your input", f["label"]); r += 1
    st.write(r, 1, "", f["legend_c"]); st.write(r, 2, "Blue = calculated (protected so formulas are not overwritten; "
                                                   "Review → Unprotect sheet if you want to customise)", f["label"])
    r += 2
    st.write(r, 2, "Google Sheets", f["label_b"]); r += 1
    st.set_row(r, 30)
    st.write(r, 2, "Upload the file to Google Drive and open it with Google Sheets (or File → Import in Sheets). "
                   "Only standard spreadsheet functions are used. Sheet protection is not carried over, so be careful with blue cells.", f["step"])
    r += 2
    st.write(r, 2, "Pricing tips from practice", f["label_b"]); r += 1
    tips = [
        "Set a minimum charge. Small jobs cost you almost as much travel and set-up time as big ones.",
        "Quote by area and condition, not by guesswork — and always ask for photos of heavy jobs before quoting.",
        "Charge travel beyond a sensible radius, and group recurring clients by area to cut driving time.",
        "Give frequency discounts only for committed schedules — they fill your diary and save set-up time.",
        "Review your prices every year when wages and supplies go up. Your Break-even sheet tells you when.",
    ]
    for tp in tips:
        st.set_row(r, 18)
        st.write(r, 1, "•", f["bullet"]); st.write(r, 2, tp, f["step"]); r += 1
    r += 1
    st.write(r, 2, "Terms of use", f["label_b"]); r += 1
    st.set_row(r, 44)
    st.write(r, 2, "Licensed for use in the purchaser's own business. Please don't resell, share or redistribute "
                   "the file. The default rates are examples to start from; results depend on the numbers you enter. "
                   f"This is not tax, legal or financial advice — check {tax} rules with your tax authority or "
                   "accountant.", f["step"])
    r += 2
    st.write(r, 2, "Questions or ideas? Message us through Etsy.",
             f["subtitle"])
    st.activate()

    # print setup for every sheet (A4, fit to width)
    for nm_, land, area in (("Start Here", False, None), ("Settings", False, None),
                            ("Quote Calculator", True, (0, 0, 44, 6)), ("Price List", True, (0, 0, 31, 9)),
                            ("Break-even", False, None), ("Quote Log", True, (0, 0, 60, 13))):
        w_ = ws[nm_]
        w_.set_paper(9)
        w_.set_landscape() if land else w_.set_portrait()
        w_.set_margins(left=0.4, right=0.4, top=0.5, bottom=0.5)
        w_.fit_to_pages(1, 1 if land else 0)
        if area:
            w_.print_area(*area)
    wb.close()
    return path


if __name__ == "__main__":
    outdir = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(__file__), "dist")
    os.makedirs(outdir, exist_ok=True)
    for key in PRESETS:
        print(build(key, outdir))
