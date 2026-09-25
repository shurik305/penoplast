#!/usr/bin/env python3
"""Builds the "Commercial Cleaning Bid Calculator" workbook in four regional editions.

Usage: python3 build.py [outdir]
Method: cleanable area per area type × frequency ÷ production rate (m² per cleaner-hour) → weekly hours →
monthly hours (× 52 ÷ 12) + periodic services + supervision → cost → price at target margin → monthly fee.
"""
import datetime as dt
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "_lib"))
from xlsxkit import Book, break_even_sheet, cell, log_range, log_sheet, print_setup, start_here  # noqa: E402

PRODUCT = "Commercial Cleaning Bid Calculator"
BRAND = "Hutsol"
VERSION = "1.0"

AREA_TYPES = [
    ("General office / open plan", 250, "Empty bins and replace liners; dust and wipe clear desks, sills and ledges; "
     "vacuum carpets and mop hard floors; spot-clean glass and partitions; disinfect touch points."),
    ("Private offices", 180, "Empty bins; dust and wipe clear desks and surfaces; vacuum or mop floors; spot-clean "
     "doors and glass; disinfect handles and switches."),
    ("Meeting rooms", 220, "Wipe tables and chairs; clean cleared whiteboards; empty bins; vacuum or mop; straighten "
     "chairs."),
    ("Reception / lobby", 180, "Clean entrance glass and doors; wipe reception desk; vacuum/mop floors and entrance "
     "mats; tidy seating; empty bins."),
    ("Corridors", 400, "Vacuum or mop floors; spot-clean walls and doors; dust ledges and handrails; empty bins."),
    ("Stairwells", 120, "Sweep and mop stairs and landings; wipe handrails; spot-clean walls; dust ledges."),
    ("Toilets / washrooms", 40, "Clean and disinfect toilets, urinals, basins and taps; descale; clean mirrors and "
     "dispensers; refill consumables; empty bins; mop floors with disinfectant."),
    ("Kitchen / break room", 70, "Clean worktops, sinks and taps; wipe appliance fronts, tables and chairs; "
     "dishwasher on request; empty bins; mop floor."),
    ("Retail sales floor", 300, "Vacuum/mop sales floor; dust fixtures and shelf edges; clean counters and glass; "
     "empty bins; entrance mats."),
    ("Clinic / treatment rooms", 90, "Disinfect touch points, couches and work surfaces; clean basins; empty general "
     "waste (not clinical waste); mop with disinfectant."),
    ("Warehouse / workshop", 600, "Sweep floors (machine where agreed); empty bins; dust accessible ledges."),
    ("Gym / changing rooms", 60, "Disinfect benches, lockers and agreed equipment touch points; clean showers and "
     "toilets; mop with disinfectant; empty bins."),
]
FREQUENCIES = [("Daily (7× per week)", 7), ("Weekdays (5× per week)", 5), ("3× per week", 3), ("2× per week", 2),
               ("Weekly", 1), ("Fortnightly", 0.5), ("Monthly", 0.23)]
TRAFFIC = [("Low", 0.85), ("Normal", 1.0), ("High", 1.25)]
# service, unit ("area" = rate is m² per cleaner-hour; "each" = rate is minutes per item), rate, default times/year
PERIODIC = [("Interior window cleaning", "area", 40, 4), ("Carpet deep extraction", "area", 35, 2),
            ("Hard floor machine scrub", "area", 150, 12), ("Hard floor strip & re-seal", "area", 25, 1),
            ("High-level dusting", "area", 250, 4), ("Upholstered chairs", "each", 8, 2),
            ("Kitchen appliances deep clean", "each", 30, 12), ("Blinds", "each", 12, 2)]

M2_AREAS = [("General office / open plan", 220, "Weekdays (5× per week)", "Normal"),
            ("Private offices", 60, "Weekdays (5× per week)", "Normal"),
            ("Meeting rooms", 40, "3× per week", "Normal"),
            ("Reception / lobby", 25, "Weekdays (5× per week)", "High"),
            ("Toilets / washrooms", 20, "Weekdays (5× per week)", "High"),
            ("Kitchen / break room", 18, "Weekdays (5× per week)", "Normal"),
            ("Corridors", 35, "3× per week", "Normal")]
M2_PERIODIC = [("Interior window cleaning", 60, 4), ("Carpet deep extraction", 280, 2),
               ("Hard floor machine scrub", 80, 12)]
SQFT = 10.7639

COMMON_FIXED = [("Insurance", 1), ("Vehicle (lease, fuel, servicing)", 8), ("Phone & internet", 1),
                ("Software (scheduling, invoicing)", 1), ("Marketing & tenders", 2), ("Accountant & bank fees", 1.5),
                ("Equipment (machines, replacement)", 2), ("Other", 0.5)]


def preset(currency, tax_name, tax, registered, area, pay, oncost, supplies, overhead, min_monthly, fixed_unit,
           edition, client, contact, site, email, time_txt, pay_terms, biz, biz_contact, web, area_served,
           consumables, weeks=47):
    k = 1 if area == "m²" else SQFT
    fixed = [(n, round(w * fixed_unit)) for n, w in COMMON_FIXED]
    return dict(currency=currency, tax_name=tax_name, tax=tax, registered=registered, incl="No", area=area,
                pay=pay, oncost=oncost, supervision=0.05, supplies=supplies, overhead=overhead, margin=0.15,
                min_monthly=min_monthly, round_to=10, cons_markup=0.15, edition=edition, client=client,
                contact=contact, site=site, email=email, window=3, time_txt=time_txt, pay_terms=pay_terms,
                biz=biz, biz_contact=biz_contact, web=web, area_served=area_served,
                areas=[(t, round(a * k, -1) if k != 1 else a, fr, tr) for t, a, fr, tr in M2_AREAS],
                periodic=[(s, round(q * k, -1) if k != 1 else q, n) for s, q, n in M2_PERIODIC],
                consumables=consumables, fixed=fixed,
                # capacity chosen so the default overhead per hour matches the default fixed costs
                hours_week=round(sum(v for _, v in fixed) / overhead * 12 / weeks), weeks=weeks)


PRESETS = {
    "UK": preset("GBP", "VAT", 0.20, "Yes", "m²", 12.75, 0.18, 0.80, 2.50, 250, 30, "UK edition (GBP, VAT)",
                 "Harbourside Accountants Ltd", "James Carter", "Unit 4, Temple Quay, Bristol BS1",
                 "james.carter@example.com", "Weekdays after 17:30",
                 "Invoiced monthly in advance; payment within 14 days by bank transfer.", "Sparkle & Co Cleaning",
                 "0117 496 0123 · hello@example.com", "www.example.com", "Bristol and surrounding areas",
                 [("Toilet paper & hand towels", 38), ("Hand soap & sanitiser", 12), ("Bin liners", 9)]),
    "EU": preset("EUR", "VAT", 0.21, "Yes", "m²", 14.00, 0.25, 0.90, 2.80, 280, 35, "EU & Ireland edition (EUR, VAT)",
                 "Van Dijk Architecten B.V.", "Sanne van Dijk", "Maliebaan 45, Utrecht", "sanne@example.com",
                 "Weekdays after 18:00", "Invoiced monthly in advance; payment within 14 days.",
                 "Clean Point Services", "+31 20 000 0000 · info@example.com", "www.example.com",
                 "Utrecht and surrounding area",
                 [("Toilet paper & hand towels", 42), ("Hand soap & sanitiser", 14), ("Bin liners", 10)]),
    "AU": preset("AUD", "GST", 0.10, "Yes", "m²", 31.00, 0.15, 1.50, 4.50, 500, 60,
                 "Australia & NZ edition (AUD, GST)", "Coastline Physio Pty Ltd", "Olivia Brown",
                 "22 Hunter Street, Newcastle NSW", "olivia@example.com", "Weekdays after 18:00",
                 "Invoiced monthly in advance; payment within 14 days.", "Harbour Clean Co",
                 "02 5550 1234 · hello@example.com", "www.example.com",
                 "Newcastle and Lake Macquarie",
                 [("Toilet paper & hand towels", 70), ("Hand soap & sanitiser", 22), ("Bin liners", 16)]),
    "US": preset("USD", "Sales tax", 0.0, "No", "sq ft", 17.00, 0.15, 1.00, 3.00, 400, 45,
                 "US & Canada edition (USD, sq ft)", "Oakridge Dental Group", "Dr. Maria Lopez",
                 "8200 Oakridge Blvd, Austin TX", "maria@example.com", "Weekdays after 6 pm",
                 "Invoiced monthly in advance; payment due within 15 days (Net 15).", "Bright Commercial Cleaning",
                 "(555) 010-2030 · bids@example.com", "www.example.com", "Austin, TX and nearby",
                 [("Toilet paper & paper towels", 55), ("Hand soap & sanitizer", 18), ("Trash liners", 12)]),
}


def build(key, outdir):
    p = PRESETS[key]
    tax = p["tax_name"]
    path = os.path.join(outdir, f"Commercial-Cleaning-Bid-Calculator_{key}.xlsx")
    b = Book(path)
    wb, f = b.wb, b.f
    wb.set_properties({"title": f"{PRODUCT} – {p['edition']}", "author": BRAND, "company": BRAND,
                       "comments": f"Version {VERSION}. For use in the purchaser's own business. Not for resale."})
    names = ["Start Here", "Settings", "Bid Calculator", "Client Proposal", "Break-even", "Bid Log"]
    ws = {n: wb.add_worksheet(n) for n in names}

    # ------------------------------------------------------------------ Settings
    S = "Settings"
    s = ws[S]
    s.hide_gridlines(2)
    s.set_column("A:A", 2)
    s.set_column("B:B", 46)
    s.set_column("C:C", 18)
    s.set_column("D:D", 18)
    s.set_column("E:E", 70)
    s.set_row(0, 30)
    s.write("B1", "Settings", f["title"])
    s.write("B2", "Fill in the yellow cells once. Every other sheet uses these numbers.", f["subtitle"])
    r = 3

    def setting(nm, label, value, fm, note="", formula=False):
        nonlocal r
        s.write(r, 1, label, f["label"])
        (s.write_formula if formula else s.write)(r, 2, value, fm)
        if note:
            s.merge_range(r, 3, r, 4, note, f["note"])
        b.name(nm, S, r, 2)
        r += 1

    b.section(s, r, "Your business (shown on the proposal)", 1, 4); r += 1
    setting("BizName", "Business name", p["biz"], f["in_text"])
    setting("BizContact", "Phone · email", p["biz_contact"], f["in_text"])
    setting("BizWeb", "Website", p["web"], f["in_text"])
    setting("BizArea", "Area you serve", p["area_served"], f["in_text"])
    setting("ValidDays", "Proposals are valid for (days)", 30, f["in_int"])
    setting("PayTerms", "Payment terms (printed on the proposal)", p["pay_terms"], f["in_text"])
    r += 1
    b.section(s, r, f"Currency, {tax} and units", 1, 4); r += 1
    setting("CurrencyLabel", "Currency (code or symbol)", p["currency"], f["in_text"])
    setting("VATRate", f"{tax} rate", p["tax"], f["in_pct"], "Your country's standard rate for cleaning services.")
    reg_row = r
    setting("VATReg", f"Are you {tax} registered?", p["registered"], f["in_text"])
    setting("InclVAT", f"Show the monthly fee including {tax}?", p["incl"], f["in_text"],
            f"Business clients usually compare prices excluding {tax}.")
    setting("AreaUnit", "Area unit you measure in", p["area"], f["in_text"], "m² or sq ft. Maths runs in m².")
    setting("VATUsed", f"{tax} rate applied (calculated)", '=IF(VATReg="Yes",VATRate,0)', f["c_pct"], formula=True)
    s.data_validation(reg_row, 2, reg_row + 1, 2, {"validate": "list", "source": ["Yes", "No"]})
    s.data_validation(reg_row + 2, 2, reg_row + 2, 2, {"validate": "list", "source": ["m²", "sq ft"]})
    r += 1
    b.section(s, r, f"Costs and margin (excl. {tax})", 1, 4); r += 1
    setting("PayRate", "Cleaner pay per hour", p["pay"], f["in_num"])
    setting("OnCost", "Employer on-costs (holiday pay, pension, insurance, payroll taxes)", p["oncost"], f["in_pct"])
    setting("LabourCost", "True labour cost per hour (calculated)", "=PayRate*(1+OnCost)", f["c_num"], formula=True)
    setting("Supervision", "Supervision & quality checks (extra % of hours)", p["supervision"], f["in_pct"],
            "Time for spot checks, key holding and client contact.")
    setting("Supplies", "Cleaning supplies per labour hour", p["supplies"], f["in_num"])
    setting("Overhead", "Overheads per labour hour", p["overhead"], f["in_num"],
            "Use the Break-even sheet to calculate it from your real monthly costs.")
    s.data_validation(r, 2, r, 2, {"validate": "decimal", "criteria": "between", "minimum": 0, "maximum": 0.95,
                                   "error_message": "Use 0% to 95%."})
    setting("Margin", "Target profit margin (% of the net fee)", p["margin"], f["in_pct"],
            "Commercial contracts often run at 10–20%.")
    setting("CostPerHr", "Total cost per labour hour (calculated)", "=LabourCost+Supplies+Overhead", f["c_num"],
            formula=True)
    setting("RatePerHr", "Your price per labour hour at target margin (calculated)",
            "=CostPerHr/(1-MIN(MAX(Margin,0),0.95))", f["c_num_b"], formula=True)
    r += 1
    b.section(s, r, "Contract terms (printed on the proposal)", 1, 4); r += 1
    setting("MinMonthly", f"Minimum monthly fee (excl. {tax})", p["min_monthly"], f["in_num"])
    setting("RoundTo", "Round the monthly fee up to the nearest", p["round_to"], f["in_num0"], "Use 0 for no rounding.")
    setting("ConsMarkup", "Handling mark-up on consumables you supply", p["cons_markup"], f["in_pct"])
    setting("ContractTerm", "Contract term", "12 months", f["in_text"])
    setting("Notice", "Notice period", "3 months' written notice", f["in_text"])
    setting("PriceReview", "Price review clause", "Prices are reviewed once a year in line with wage and cost "
                                                  "increases.", f["in_text"])
    setting("Insurance", "Insurance statement", "[Describe your insurance cover here — edit this text in Settings.]",
            f["in_text"], "Printed on the proposal. Only state cover you actually hold.")
    setting("Access", "Keys and access", "Keys and alarm codes are held securely and their use is logged.",
            f["in_text"])
    r += 1
    b.section(s, r, "Area types and production rates  —  calibrate with your own timings!", 1, 4); r += 1
    s.write(r, 1, "Area type", f["th"])
    s.write_formula(r, 2, '="Area per cleaner-hour ("&AreaUnit&")"', f["th"])
    s.merge_range(r, 3, r, 4, "Tasks each visit (printed in the scope of work)", f["th"]); r += 1
    a0 = r
    for nm_, rate, scope in AREA_TYPES:
        s.write(r, 1, nm_, f["in_text"])
        s.write(r, 2, rate if p["area"] == "m²" else round(rate * SQFT / 5) * 5, f["in_num0"])
        s.merge_range(r, 3, r, 4, scope, f["in_wrap"])
        s.set_row(r, 42)
        r += 1
    b.name("AreaNames", S, a0, 1, r - 1, 1)
    b.name("AreaRates", S, a0, 2, r - 1, 2)
    b.name("AreaScopes", S, a0, 3, r - 1, 3)
    s.merge_range(r, 1, r, 4, "Rates are the floor area (in your area unit) a cleaner finishes per hour for routine "
                              "cleaning at normal traffic, e.g. 220 m² of open-plan office ÷ 250 = 0.9 hours per visit. "
                              "Renamed a type? Choose it again on Bid Calculator — it warns you about names that don't "
                              "match.", f["note"])
    s.set_row(r, 26)
    r += 2
    b.section(s, r, "Frequencies", 1, 4); r += 1
    s.write_row(r, 1, ["Frequency", "Visits per week"], f["th"]); r += 1
    f0 = r
    for nm_, v in FREQUENCIES:
        s.write(r, 1, nm_, f["in_text"]); s.write(r, 2, v, f["in_dec"]); r += 1
    b.name("FreqNames", S, f0, 1, r - 1, 1)
    b.name("FreqVisits", S, f0, 2, r - 1, 2)
    r += 1
    b.section(s, r, "Traffic / soiling level", 1, 4); r += 1
    s.write_row(r, 1, ["Traffic", "Time multiplier"], f["th"]); r += 1
    t0 = r
    for nm_, m in TRAFFIC:
        s.write(r, 1, nm_, f["in_text"]); s.write(r, 2, m, f["in_dec"]); r += 1
    b.name("TrafficNames", S, t0, 1, r - 1, 1)
    b.name("TrafficMults", S, t0, 2, r - 1, 2)
    r += 1
    b.section(s, r, "Periodic services", 1, 4); r += 1
    s.write_row(r, 1, ["Service", "Unit: area or each", "Rate", "Rate means…"], f["th"]); r += 1
    p0 = r
    for nm_, unit, rate, _ in PERIODIC:
        s.write(r, 1, nm_, f["in_text"]); s.write(r, 2, unit, f["in_text"])
        s.write(r, 3, rate if unit == "each" or p["area"] == "m²" else round(rate * SQFT / 5) * 5, f["in_num0"])
        s.write_formula(r, 4, f'=IF({cell(r, 2)}="each","minutes per item",AreaUnit&" per cleaner-hour")', f["note"])
        r += 1
    b.name("PerNames", S, p0, 1, r - 1, 1)
    b.name("PerUnits", S, p0, 2, r - 1, 2)
    b.name("PerRates", S, p0, 3, r - 1, 3)
    s.data_validation(p0, 2, r - 1, 2, {"validate": "list", "source": ["area", "each"]})
    s.freeze_panes(2, 0)

    # ------------------------------------------------------------------ Bid Calculator
    Q = "Bid Calculator"
    q = ws[Q]
    q.hide_gridlines(2)
    for c_, w in zip(range(11), [2, 30, 12, 22, 11, 11, 11, 11, 3, 44, 16]):
        q.set_column(c_, c_, w)
    q.set_row(0, 30)
    q.merge_range("B1:K1", "Bid Calculator", f["title"])
    q.merge_range("B2:K2", "Fill in the yellow cells — the monthly fee, staffing and the printable proposal update "
                           "instantly.", f["subtitle"])
    b.section(q, 3, "1. Client and site", 1, 7)

    def qin(row, key, label, value, fm, formula=False, span=True):
        q.write(row, 1, label, f["label"])
        if span:
            q.merge_range(row, 2, row, 3, "", fm)
        (q.write_formula if formula else q.write)(row, 2, value, fm)
        b.name(key, Q, row, 2)

    qin(4, "QNum", "Bid number", "B-2001", f["in_text"])
    qin(5, "QDate", "Bid date", "=TODAY()", f["in_date"], formula=True)
    qin(6, "QCompany", "Client company", p["client"], f["in_text"])
    qin(7, "QContact", "Contact person", p["contact"], f["in_text"])
    qin(8, "QSite", "Site address", p["site"], f["in_text"])
    qin(9, "QEmail", "Contact email / phone", p["email"], f["in_text"])
    qin(10, "QWindow", "Cleaning window per visit (hours available)", p["window"], f["in_dec"], span=False)
    qin(11, "QTime", "Cleaning time (printed on the proposal)", p["time_txt"], f["in_text"])

    b.section(q, 13, "2. Areas to clean", 1, 7)
    q.write(14, 1, "Area type", f["th"])
    q.write_formula(14, 2, '="Area ("&AreaUnit&")"', f["th"])
    q.write_row(14, 3, ["Frequency", "Traffic", "Visits / week", "Hours / visit", "Hours / week"], f["th"])
    q.set_row(14, 30)
    ar0, ar1 = 15, 26
    for i, rr in enumerate(range(ar0, ar1 + 1)):
        t, a, fr, tr = p["areas"][i] if i < len(p["areas"]) else ("", "", "", "")
        q.write(rr, 1, t, f["in_text"]); q.write(rr, 2, a, f["in_num0"])
        q.write(rr, 3, fr, f["in_text"]); q.write(rr, 4, tr, f["in_text"])
        B_, C_, D_, E_, F_, G_ = (cell(rr, c_) for c_ in range(1, 7))
        q.write_formula(rr, 5, f'=IF({B_}="","",IFERROR(INDEX(FreqVisits,MATCH({D_},FreqNames,0)),0))', f["c_dec2"])
        q.write_formula(rr, 6, f'=IF(OR({B_}="",N({C_})=0),"",IFERROR(N({C_})/'
                               f'INDEX(AreaRates,MATCH({B_},AreaNames,0))*IFERROR(INDEX(TrafficMults,'
                               f'MATCH({E_},TrafficNames,0)),1),""))', f["c_dec2"])
        q.write_formula(rr, 7, f'=IF({G_}="","",{G_}*N({F_}))', f["c_dec2"])
        q.write_formula(rr, 13, f'=IF({B_}="",IF(N({C_})<>0,1,0),IF(AND(ISNUMBER(MATCH({B_},AreaNames,0)),'
                                f'ISNUMBER(MATCH({D_},FreqNames,0)),N({C_})>0,OR({E_}="",ISNUMBER(MATCH({E_},'
                                f'TrafficNames,0)))),0,1))')
    tot = ar1 + 1
    q.write(tot, 1, "Total", f["label_b"])
    q.write_formula(tot, 2, f'=SUMPRODUCT(--({cell(ar0, 6)}:{cell(ar1, 6)}<>""),{cell(ar0, 2)}:{cell(ar1, 2)})',
                    f["c_num0"])
    b.name("TotalArea", Q, tot, 2)
    q.write_formula(tot, 7, f"=SUM({cell(ar0, 7)}:{cell(ar1, 7)})", f["c_num_b"])
    b.name("RoutineHrsWeek", Q, tot, 7)
    q.write_formula(tot, 5, f"=MAX({cell(ar0, 5)}:{cell(ar1, 5)})", f["c_dec2"])
    b.name("MaxVisits", Q, tot, 5)
    q.write_formula(tot, 6, f'=SUMIF({cell(ar0, 5)}:{cell(ar1, 5)},">=1",{cell(ar0, 6)}:{cell(ar1, 6)})', f["c_dec2"])
    b.name("BusiestVisit", Q, tot, 6)
    q.data_validation(ar0, 1, ar1, 1, {"validate": "list", "source": b.lst("AreaNames")})
    q.data_validation(ar0, 2, ar1, 2, {"validate": "decimal", "criteria": ">=", "value": 0})
    q.data_validation(ar0, 3, ar1, 3, {"validate": "list", "source": b.lst("FreqNames")})
    q.data_validation(ar0, 4, ar1, 4, {"validate": "list", "source": b.lst("TrafficNames")})

    b.section(q, 29, "3. Periodic services", 1, 7)
    q.write_row(30, 1, ["Service", "Quantity", "Unit", "Times / year", "Hours / year", "Hours / month"], f["th"])
    pr0, pr1 = 31, 38
    for i, rr in enumerate(range(pr0, pr1 + 1)):
        sv, qty, n = p["periodic"][i] if i < len(p["periodic"]) else ("", "", "")
        q.write(rr, 1, sv, f["in_text"]); q.write(rr, 2, qty, f["in_num0"]); q.write(rr, 4, n, f["in_num0"])
        B_, C_, D_, E_, F_ = (cell(rr, c_) for c_ in range(1, 6))
        q.write_formula(rr, 3, f'=IF({B_}="","",IF(IFERROR(INDEX(PerUnits,MATCH({B_},PerNames,0)),"")="each",'
                               f'"each",AreaUnit))', f["c_text"])
        q.write_formula(rr, 5, f'=IF(OR({B_}="",N({C_})=0,N({E_})=0),"",IFERROR(IF({D_}="each",'
                               f'N({C_})*INDEX(PerRates,MATCH({B_},PerNames,0))/60,N({C_})/'
                               f'INDEX(PerRates,MATCH({B_},PerNames,0)))*N({E_}),""))', f["c_dec1"])
        q.write_formula(rr, 6, f'=IF({F_}="","",{F_}/12)', f["c_dec2"])
        q.write_formula(rr, 13, f'=IF({B_}="",IF(OR(N({C_})<>0,N({E_})<>0),1,0),IF(AND(ISNUMBER(MATCH({B_},PerNames,0)),'
                                f'N({C_})>0,N({E_})>0),0,1))')
    q.write(pr1 + 1, 1, "Total", f["label_b"])
    q.write_formula(pr1 + 1, 6, f"=SUM({cell(pr0, 6)}:{cell(pr1, 6)})", f["c_num_b"])
    b.name("PeriodicHrsMonth", Q, pr1 + 1, 6)
    q.data_validation(pr0, 1, pr1, 1, {"validate": "list", "source": b.lst("PerNames")})
    q.data_validation(pr0, 2, pr1, 2, {"validate": "decimal", "criteria": ">=", "value": 0})
    q.data_validation(pr0, 4, pr1, 4, {"validate": "decimal", "criteria": ">=", "value": 0})
    per_list = "&".join(f'IF(AND({cell(rr, 1)}<>"",N({cell(rr, 2)})>0,N({cell(rr, 4)})>0,ISNUMBER(MATCH('
                        f'{cell(rr, 1)},PerNames,0))),{cell(rr, 1)}&" — "&FIXED({cell(rr, 2)},0)&" "&{cell(rr, 3)}'
                        f'&", "&{cell(rr, 4)}&"× per year; ","")'
                        for rr in range(pr0, pr1 + 1))
    # helper text for the proposal lives in hidden column M
    q.set_column(12, 13, 30, None, {"hidden": True})
    q.write_formula(pr0 + 3, 12, f"=SUM({cell(ar0, 13)}:{cell(ar1, 13)})+SUM({cell(pr0, 13)}:{cell(pr1, 13)})")
    b.name("LinesBad", Q, pr0 + 3, 12)
    q.write_formula(pr0, 12, "=" + per_list)
    q.write_formula(pr0 + 1, 12, f'=IF({cell(pr0, 12)}="","None",LEFT({cell(pr0, 12)},LEN({cell(pr0, 12)})-2))')
    b.name("PeriodicList", Q, pr0 + 1, 12)

    b.section(q, 41, "4. Consumables you supply (optional)", 1, 7)
    q.write(42, 1, "Item", f["th"])
    q.write_formula(42, 2, f'="Cost / month (excl. {tax})"', f["th"])
    q.set_row(42, 30)
    c0, c1 = 43, 47
    for i, rr in enumerate(range(c0, c1 + 1)):
        it, amt = p["consumables"][i] if i < len(p["consumables"]) else ("", "")
        q.write(rr, 1, it, f["in_text"]); q.write(rr, 2, amt, f["in_num"])
    q.write(c1 + 1, 1, "Total consumables (your cost)", f["label_b"])
    q.write_formula(c1 + 1, 2, f"=SUM({cell(c0, 2)}:{cell(c1, 2)})", f["c_num_b"])
    b.name("ConsCost", Q, c1 + 1, 2)

    # results panel (J:K) — workings first so names exist
    W = {}

    def wrow(row, key, label, formula, fm, label_formula=False):
        (q.write_formula if label_formula else q.write)(row, 9, label, f["label"])
        q.write_formula(row, 10, formula, fm)
        if key:
            b.name(key, Q, row, 10)
        W[key] = row

    q.merge_range(29, 9, 29, 10, "How the fee is built", f["section"])
    wrow(30, "RoutineHrsMonth", "Routine hours per month (weekly × 52 ÷ 12)", "=RoutineHrsWeek*52/12", f["c_dec1"])
    wrow(31, "PerHrs", "Periodic services hours per month", "=PeriodicHrsMonth", f["c_dec1"])
    wrow(32, "SupHrs", "Supervision & quality checks", "=(RoutineHrsMonth+PerHrs)*Supervision", f["c_dec1"])
    wrow(33, "TotalHrsMonth", "Total labour hours per month", "=RoutineHrsMonth+PerHrs+SupHrs", f["c_dec1"])
    wrow(34, "CostMonth", "Cost: labour + supplies + overheads", "=TotalHrsMonth*CostPerHr", f["c_num"])
    wrow(35, "PriceMargin", "Price at your target margin", "=CostMonth/(1-MIN(MAX(Margin,0),0.95))", f["c_num"])
    wrow(36, "ConsPrice0", '="Consumables incl. "&TEXT(ConsMarkup,"0%")&" handling"', "=ConsCost*(1+ConsMarkup)",
         f["c_num"], label_formula=True)
    wrow(37, "Subtotal0", "Subtotal", "=PriceMargin+ConsPrice0", f["c_num"])
    wrow(38, "AfterMin", "After minimum monthly fee", "=IF(Subtotal0>0,MAX(MinMonthly,Subtotal0),0)", f["c_num"])
    wrow(39, "FinalNet", f"Final net monthly fee after rounding (excl. {tax})",
         '=IF(AfterMin<=0,0,IF(InclVAT="Yes",IF(RoundTo>0,CEILING(AfterMin*(1+VATUsed),RoundTo),'
         'AfterMin*(1+VATUsed))/(1+VATUsed),IF(RoundTo>0,CEILING(AfterMin,RoundTo),AfterMin)))', f["c_num_b"])
    wrow(40, "PeriodicPrice", "Periodic services share (for the proposal)",
         "=ROUND(PerHrs*(1+Supervision)*CostPerHr/(1-MIN(MAX(Margin,0),0.95)),2)", f["c_num"])
    wrow(41, "ConsPrice", "Consumables share (for the proposal)", "=ROUND(ConsPrice0,2)", f["c_num"])

    q.merge_range(3, 9, 3, 10, "Monthly fee", f["section"])
    q.set_row(4, 34)
    q.write_formula(4, 9, f'="Monthly fee ("&IF(VATUsed=0,"no {tax}",IF(InclVAT="Yes","incl. {tax}","excl. {tax}"))&", "'
                          f'&CurrencyLabel&")"', f["big_label"])
    q.write_formula(5, 9, '="Based on "&FIXED(RoutineHrsWeek,1)&" cleaning hours per week"', f["note"])
    q.write(6, 9, f"Net monthly fee (excl. {tax})", f["label"])
    q.write_formula(6, 10, "=ROUND(FinalNet,2)", f["c_num"])
    b.name("NetMonthly", Q, 6, 10)
    q.write_formula(7, 9, f'="{tax} ("&TEXT(VATUsed,"0%")&")"', f["label"])
    q.write(8, 9, f"Total per month incl. {tax}", f["label_b"])
    q.write_formula(8, 10, "=ROUND(FinalNet*(1+VATUsed),2)", f["c_num_b"])
    b.name("Gross", Q, 8, 10)
    q.write_formula(7, 10, "=Gross-NetMonthly", f["c_num"])
    b.name("VATAmt", Q, 7, 10)
    q.write_formula(4, 10, '=IF(InclVAT="Yes",Gross,NetMonthly)', f["big"])
    q.write(9, 9, f"Annual contract value (excl. {tax})", f["label"])
    q.write_formula(9, 10, "=NetMonthly*12", f["c_num"])
    b.name("Annual", Q, 9, 10)

    q.merge_range(11, 9, 11, 10, "Workload and staffing", f["section"])
    q.write(12, 9, "Routine cleaning hours per week", f["label"])
    q.write_formula(12, 10, "=RoutineHrsWeek", f["c_dec1"])
    q.write(13, 9, "Total labour hours per month (all work)", f["label"])
    q.write_formula(13, 10, "=TotalHrsMonth", f["c_dec1"])
    q.write(14, 9, "Busiest visit: all daily/weekly areas (hours)", f["label"])
    q.write_formula(14, 10, "=BusiestVisit", f["c_dec1"])
    q.write_formula(15, 9, '="Cleaners needed for a "&FIXED(QWindow,1)&"-hour window"', f["label"])
    q.write_formula(15, 10, '=IF(N(QWindow)>0,MAX(1,ROUNDUP(BusiestVisit/QWindow,0)),"")', f["c_num0"])
    b.name("Cleaners", Q, 15, 10)
    q.write_formula(16, 9, f'="Price per "&AreaUnit&" per month (excl. {tax})"', f["label"])
    q.write_formula(16, 10, "=IF(TotalArea>0,FinalNet/TotalArea,0)", f["c_num"])
    q.write(17, 9, f"Effective rate per labour hour (excl. {tax})", f["label"])
    q.write_formula(17, 10, "=IF(TotalHrsMonth>0,(FinalNet-ConsPrice)/TotalHrsMonth,0)", f["c_num"])
    b.name("EffRate", Q, 17, 10)

    q.merge_range(19, 9, 19, 10, "Profit check (per month)", f["section"])
    q.write(20, 9, "Labour cost", f["label"])
    q.write_formula(20, 10, "=TotalHrsMonth*LabourCost", f["c_num"])
    q.write(21, 9, "Supplies + overheads", f["label"])
    q.write_formula(21, 10, "=TotalHrsMonth*(Supplies+Overhead)", f["c_num"])
    q.write(22, 9, "Consumables (your cost)", f["label"])
    q.write_formula(22, 10, "=ConsCost", f["c_num"])
    q.write(23, 9, "Profit per month", f["label_b"])
    q.write_formula(23, 10, "=FinalNet-K21-K22-K23", f["c_num_b"])
    b.name("ProfitMonth", Q, 23, 10)
    q.write(24, 9, "Profit margin", f["label"])
    q.write_formula(24, 10, "=IF(FinalNet>0,ProfitMonth/FinalNet,0)", f["c_pct"])
    b.name("BidMargin", Q, 24, 10)
    q.merge_range(25, 9, 27, 10, "", f["status"])
    q.write_formula(25, 9, '=IF(LinesBad>0,"⚠ "&LinesBad&" line(s) not priced: a name no longer matches Settings, or '
                           'the type, size, frequency or times per year is missing. Choose them again.",'
                           'IF(TotalArea=0,"Add the areas to clean (type, size and frequency).",'
                           'IF(EffRate<BERate,"⚠ Below your break-even rate of "&FIXED(BERate,2)&" per labour hour — '
                           'check the Break-even sheet.",'
                           'IF(BidMargin<MIN(MAX(Margin,0),0.95)-0.005,"△ Covers your costs, but the margin is "'
                           '&TEXT(BidMargin,"0%")&" — below your "&TEXT(Margin,"0%")&" target (minimum fee or '
                           'rounding?).","✔ Meets your target margin ("&TEXT(BidMargin,"0%")&")."))))', f["status"])
    q.conditional_format(25, 9, 27, 10, {"type": "formula", "criteria": '=LEFT($J$26,1)="✔"', "format": f["green"]})
    q.conditional_format(25, 9, 27, 10, {"type": "formula", "criteria": '=LEFT($J$26,1)="⚠"', "format": f["red"]})
    q.conditional_format(25, 9, 27, 10, {"type": "formula", "criteria": '=LEFT($J$26,1)="△"', "format": f["amber"]})
    q.freeze_panes(3, 0)
    q.protect("", {"select_locked_cells": True, "select_unlocked_cells": True})

    # ------------------------------------------------------------------ Client Proposal
    cp = ws["Client Proposal"]
    cp.hide_gridlines(2)
    for c_, w in zip(range(6), [1, 37, 11, 23, 14, 1]):
        cp.set_column(c_, c_, w)
    cp.set_row(0, 28)
    cp.write_formula(0, 1, "=BizName", f["q_biz"])
    cp.merge_range(0, 2, 0, 4, "PROPOSAL", f["q_title"])
    cp.write_formula(1, 1, "=BizContact", f["q_small"])
    cp.write_formula(2, 1, "=BizWeb", f["q_small"])
    cp.write_formula(3, 1, "=BizArea", f["q_small"])
    cp.write(5, 1, "PREPARED FOR", f["q_cap"])
    for rr, lab, fml, fk in ((5, "Proposal no.", "=QNum", "q_val_r"), (6, "Date", "=QDate", "q_date_r"),
                             (7, "Valid until", "=QDate+ValidDays", "q_date_r")):
        cp.write(rr, 3, lab, f["q_small_r"])
        cp.write_formula(rr, 4, fml, f[fk])
    cp.write_formula(6, 1, "=QCompany", f["q_client"])
    cp.write_formula(7, 1, '="Attn: "&QContact', f["label"])
    cp.write_formula(8, 1, "=QSite", f["label"])
    cp.write_formula(9, 1, "=QEmail", f["label"])
    cp.set_row(11, 30)
    cp.merge_range(11, 1, 11, 4, "", f["q_desc"])
    cp.write_formula(11, 1, '="Commercial cleaning — "&FIXED(TotalArea,0)&" "&AreaUnit&", about "'
                            '&FIXED(RoutineHrsWeek,1)&" cleaning hours per week"', f["q_desc"])
    cp.merge_range(12, 1, 12, 4, "", f["q_small"])
    cp.write_formula(12, 1, '="Cleaning time: "&QTime&IF(N(Cleaners)>0," · Team of "&Cleaners,"")&" · Contract: "'
                            '&ContractTerm', f["q_small"])
    cp.write(14, 1, "Area", f["q_th"])
    cp.write_formula(14, 2, '="Size ("&AreaUnit&")"', f["q_th_c"])
    cp.write(14, 3, "Frequency", f["q_th"])
    cp.write(14, 4, "Hours / visit", f["q_th_r"])
    for i in range(ar1 - ar0 + 1):
        src = ar0 + i
        rr = 15 + i
        B_ = f"'{Q}'!{cell(src, 1)}"
        cp.write_formula(rr, 1, f'=IF({B_}="","",{B_})', f["q_row"])
        cp.write_formula(rr, 2, f"=IF({B_}=\"\",\"\",'{Q}'!{cell(src, 2)})", f["q_int_c"])
        cp.write_formula(rr, 3, f"=IF({B_}=\"\",\"\",'{Q}'!{cell(src, 3)})", f["q_row"])
        cp.write_formula(rr, 4, f"=IF({B_}=\"\",\"\",IF('{Q}'!{cell(src, 13)}=1,\"⚠ check\",'{Q}'!{cell(src, 6)}))",
                         f["q_amt"])
    rr = 15 + (ar1 - ar0 + 1) + 1
    cp.merge_range(rr, 1, rr, 4, "Periodic services", f["q_h"]); rr += 1
    cp.set_row(rr, 30)
    cp.merge_range(rr, 1, rr, 4, "", f["q_body"])
    cp.write_formula(rr, 1, "=PeriodicList", f["q_body"]); rr += 2
    cp.write(rr, 1, "Monthly fee", f["q_th"])
    cp.write(rr, 2, "", f["q_th"]); cp.write(rr, 3, "", f["q_th"])
    cp.write_formula(rr, 4, '="Amount ("&CurrencyLabel&")"', f["q_th_r"]); rr += 1
    fee0 = rr
    for lab, fml in (("Regular cleaning as per the scope of work", None),
                     ("Periodic services (averaged per month)", "=PeriodicPrice"),
                     ("Consumables (paper, soap, liners)", "=ConsPrice")):
        cp.merge_range(rr, 1, rr, 3, lab, f["q_row"])
        if fml:
            cp.write_formula(rr, 4, fml, f["q_amt"])
        rr += 1
    cp.write_formula(fee0, 4, f"=NetMonthly-SUM({cell(fee0 + 1, 4)}:{cell(fee0 + 2, 4)})", f["q_amt"])
    cp.merge_range(rr, 1, rr, 3, "", f["q_sub"])
    cp.write_formula(rr, 1, f'=IF(VATUsed>0,"Subtotal per month (excl. {tax})","Subtotal per month")', f["q_sub"])
    cp.write_formula(rr, 4, "=NetMonthly", f["q_sub_amt"]); rr += 1
    cp.merge_range(rr, 1, rr, 3, "", f["label"])
    cp.write_formula(rr, 1, f'=IF(VATUsed>0,"{tax} "&TEXT(VATUsed,"0%"),"")', f["label"])
    cp.write_formula(rr, 4, '=IF(VATUsed>0,VATAmt,"")', f["q_amt"]); rr += 1
    cp.set_row(rr, 22)
    cp.merge_range(rr, 1, rr, 3, "TOTAL PER MONTH", f["q_total"])
    cp.write_formula(rr, 4, "=Gross", f["q_total_amt"]); rr += 1
    cp.merge_range(rr, 1, rr, 4, "", f["q_small"])
    cp.write_formula(rr, 1, f'="Annual contract value: "&FIXED(Annual,2)&" "&CurrencyLabel&IF(VATUsed>0," excl. {tax}","")&"."',
                     f["q_small"]); rr += 1
    page2 = rr + 1
    cp.set_h_pagebreaks([page2])
    rr = page2
    cp.merge_range(rr, 1, rr, 4, "Scope of work — tasks each visit", f["q_h"]); rr += 1
    # one row per area line; repeated area types are listed once
    for src in range(ar0, ar1 + 1):
        a_ = f"'{Q}'!{cell(src, 1)}"
        first = f"COUNTIF('{Q}'!{cell(ar0, 1, True)}:{cell(src, 1)},{a_})=1"
        cp.merge_range(rr, 1, rr, 4, "", f["q_body"])
        cp.write_formula(rr, 1, f'=IF(AND({a_}<>"",{first}),{a_}&": "&IFERROR(INDEX(AreaScopes,MATCH({a_},'
                                f'AreaNames,0)),""),"")', f["q_body"])
        cp.set_row(rr, 25)
        rr += 1
    rr += 1
    cp.merge_range(rr, 1, rr, 4, "Terms", f["q_h"]); rr += 1
    cp.merge_range(rr, 1, rr + 2, 4, "", f["q_body"])
    cp.write_formula(rr, 1, '="Contract term: "&ContractTerm&". Notice period: "&Notice&". "&PayTerms&" "&PriceReview'
                            '&" "&Insurance&" "&Access&" This proposal is based on the site visit and the information '
                            'provided; any change in scope will be agreed in writing."', f["q_body"])
    for k in range(3):
        cp.set_row(rr + k, 22)
    rr += 4
    line = "_" * 44 + "      Date: " + "_" * 14
    cp.merge_range(rr, 1, rr, 4, "For the client — name and signature:", f["label"]); rr += 2
    cp.merge_range(rr, 1, rr, 4, line, f["label"]); rr += 2
    cp.merge_range(rr, 1, rr, 4, "", f["label"])
    cp.write_formula(rr, 1, '="For "&BizName&" — name and signature:"', f["label"]); rr += 2
    cp.merge_range(rr, 1, rr, 4, line, f["label"]); rr += 2
    cp.merge_range(rr, 1, rr, 4, "Thank you for the opportunity to quote — we look forward to working with you!",
                   f["q_thanks"])
    cp.set_paper(9)
    cp.set_portrait()
    cp.set_margins(left=0.5, right=0.5, top=0.6, bottom=0.6)
    cp.print_area(0, 0, rr, 5)
    # no fit-to-page: it would make Excel ignore the manual page break; columns fit A4 at 100%
    cp.protect("", {"select_locked_cells": True, "select_unlocked_cells": True})

    # ------------------------------------------------------------------ Break-even, Bid Log, Start Here
    break_even_sheet(b, ws["Break-even"], tax, p["fixed"], p["hours_week"], p["weeks"])

    today = dt.date.today()
    heads = ["Bid no.", "Date", "Client", "Site", "Area", "Visits / week", "Monthly fee", "Annual value", "Status",
             "Follow-up", "Notes"]
    t = log_range
    tiles = [("Bids logged", f"=COUNTA({t(1)})", "tile_val", 2),
             ("Won", f'=COUNTIF({t(9)},"Won")', "tile_val", 1),
             ("Lost", f'=COUNTIF({t(9)},"Lost")', "tile_val", 1),
             ("Win rate", f'=IF(COUNTIF({t(9)},"Won")+COUNTIF({t(9)},"Lost")>0,COUNTIF({t(9)},"Won")/'
                          f'(COUNTIF({t(9)},"Won")+COUNTIF({t(9)},"Lost")),0)', "tile_pct", 2),
             ("Won monthly revenue", f'=SUMIF({t(9)},"Won",{t(7)})', "tile_money", 2),
             ("Open pipeline (monthly)", f'=SUMIF({t(9)},"Sent",{t(7)})+SUMIF({t(9)},"Follow-up",{t(7)})',
              "tile_money", 2),
             ("Follow-ups overdue", f'=COUNTIFS({t(9)},"Sent",{t(10)},"<"&TODAY())+COUNTIFS({t(9)},"Follow-up",'
                                    f'{t(10)},"<"&TODAY())', "tile_val", 1)]
    k = 1 if p["area"] == "m²" else SQFT
    samples = [("B-2001", today - dt.timedelta(days=15), "Example: Harbourside Ltd", "Bristol BS1", round(418 * k),
                5, 1540, 18480, "Won", None, "Example row — delete me"),
               ("B-2002", today - dt.timedelta(days=6), "Example: Northside Dental", "Clifton", round(180 * k), 5,
                890, 10680, "Follow-up", today + dt.timedelta(days=1), "Example row — delete me"),
               ("B-2003", today - dt.timedelta(days=25), "Example: City Gym", "Redland", round(600 * k), 7, 2350,
                28200, "Lost", None, "Example row — delete me")]
    log_sheet(b, ws["Bid Log"], "Bid Log & Win Rate", "Log every proposal you send. Status: Sent → Follow-up → "
              "Won or Lost. Rows 10 onwards.", heads, [2, 11, 13, 26, 18, 10, 11, 13, 14, 12, 13, 34], tiles,
              samples, [(9, ["Sent", "Follow-up", "Won", "Lost"])], status_col=9, followup_col=10,
              date_cols=(2, 10), num_cols=(5, 6, 7, 8))

    steps = [("Settings", f"Enter your business details, currency and {tax}, your wages, on-costs, supplies, overheads "
                          "and target margin. Yellow cells are yours; blue cells calculate."),
             ("Production rates", "Check the area per cleaner-hour for each area type. Time a few real visits and "
                                  "adjust — this is what makes every bid accurate."),
             ("Bid Calculator", "Enter the client, then each area to clean with its size, frequency and traffic. "
                                "Add periodic services (windows, carpets, floors) and any consumables you supply."),
             ("Read the result", "Monthly fee, annual value, hours per week, team size for the cleaning window and "
                                 "your monthly profit — with a warning if you are below break-even."),
             ("Client Proposal", "A two-page A4 proposal with scope of work, fee and terms is filled in for you. "
                                 "File → Print → Save as PDF and send it."),
             ("Bid Log", "Track every bid: win rate, won monthly revenue, open pipeline and overdue follow-ups."),
             ("Break-even", "Enter your monthly fixed costs to find the minimum rate per labour hour.")]
    tips = ["Walk the site before bidding and measure (or ask for floor plans) — guessed areas lose money.",
            "Price toilets and kitchens separately: they take 5–10× longer per m² than open offices.",
            "Check the cleaning window: a short window means a bigger team, not fewer hours.",
            "Put periodic work (windows, carpets, floor care) in the bid so it isn't done for free.",
            "Include an annual price review — wages rise every year, your contract price should too."]
    start_here(b, ws["Start Here"], PRODUCT, p["edition"], VERSION, BRAND, steps, tips, tax)

    print_setup(ws["Settings"], False)
    print_setup(ws["Bid Calculator"], True, (0, 0, 48, 10))
    print_setup(ws["Bid Log"], True, (0, 0, 60, 12))
    print_setup(ws["Start Here"], False)
    print_setup(ws["Break-even"], False)
    wb.close()
    return path


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(__file__), "dist")
    os.makedirs(out, exist_ok=True)
    for key in PRESETS:
        print(build(key, out))
