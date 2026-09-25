"""Shared workbook styling and helpers for all spreadsheet products (xlsxwriter).

Formulas written with these helpers stick to functions that work in Excel 2010+, Google Sheets and LibreOffice.
"""
import datetime as dt

import xlsxwriter
from xlsxwriter.utility import xl_col_to_name, xl_rowcol_to_cell

NAVY, BLUE, LIGHT, INPUT, INPUT_BORDER = "#17324D", "#1F5FA6", "#EAF1FA", "#FFF4CC", "#D9A400"
GREY, LINE = "#5B6B7B", "#C9D6E3"


def cell(r, c, abs_=False):
    return xl_rowcol_to_cell(r, c, abs_, abs_)


def col(c):
    return xl_col_to_name(c)


class Book:
    def __init__(self, path):
        # Arial default font: column widths then match in Excel, Google Sheets and LibreOffice (Liberation Sans)
        self.wb = xlsxwriter.Workbook(path, {"default_format_properties": {"font_name": "Arial", "font_size": 10}})
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
        fmt("th_r", bold=True, font_color=NAVY, bg_color=LIGHT, bottom=1, bottom_color=LINE, text_wrap=True,
            align="right")
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
        fmt("step_no", bold=True, font_color="white", bg_color=BLUE, align="center", font_size=12, border=2,
            border_color="white")
        fmt("step", text_wrap=True, valign="vcenter")
        fmt("step_b", text_wrap=True, valign="vcenter", bold=True, font_color=NAVY)
        fmt("bullet", align="center", valign="vcenter", font_color=BLUE, bold=True)
        fmt("legend_in", bg_color=INPUT, border=1, border_color=INPUT_BORDER)
        fmt("legend_c", bg_color=LIGHT, border=1, border_color=LINE)
        # printable client documents
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
        fmt("q_row_r", bottom=1, bottom_color=LINE, align="right")
        fmt("q_amt", bottom=1, bottom_color=LINE, num_format='#,##0.00;-#,##0.00;"–"')
        fmt("q_int", bottom=1, bottom_color=LINE, num_format='#,##0;-#,##0;""', align="right")
        fmt("q_int_c", bottom=1, bottom_color=LINE, num_format='#,##0;-#,##0;""', align="center")
        fmt("q_th_c", bold=True, font_color="white", bg_color=NAVY, align="center")
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

    def section(self, ws, row, text, first_col=1, last_col=3):
        ws.merge_range(row, first_col, row, last_col, text, self.f["section"])
        ws.set_row(row, 20)


def print_setup(ws, landscape, area=None, tall=0):
    ws.set_paper(9)
    ws.set_landscape() if landscape else ws.set_portrait()
    ws.set_margins(left=0.4, right=0.4, top=0.5, bottom=0.5)
    ws.fit_to_pages(1, tall if tall else (1 if landscape else 0))
    if area:
        ws.print_area(*area)


def start_here(b, ws, product, edition, version, brand, steps, tips, tax, extra_note=None):
    f = b.f
    ws.hide_gridlines(2)
    ws.set_column("A:A", 2)
    ws.set_column("B:B", 5)
    ws.set_column("C:C", 100)
    ws.set_row(0, 34)
    ws.write("B1", product, f["title"])
    ws.write("B2", f"{edition} · version {version} · by {brand}", f["subtitle"])
    ws.write("B4", "How to use it", f["label_b"])
    r = 4
    for i, (hd, tx) in enumerate(steps, start=1):
        ws.set_row(r, 36)
        ws.write(r, 1, i, f["step_no"])
        ws.write_rich_string(r, 2, f["step_b"], hd + ": ", f["step"], tx, f["step"])
        r += 1
    r += 1
    ws.write(r, 1, "", f["legend_in"])
    ws.write(r, 2, "Yellow = your input", f["label"])
    r += 1
    ws.write(r, 1, "", f["legend_c"])
    ws.write(r, 2, "Blue = calculated (protected so formulas are not overwritten; Review → Unprotect sheet if you "
                   "want to customise)", f["label"])
    r += 2
    ws.write(r, 2, "Google Sheets", f["label_b"])
    r += 1
    ws.set_row(r, 30)
    ws.write(r, 2, "Upload the file to Google Drive and open it with Google Sheets (or File → Import in Sheets). "
                   "All formulas work. Sheet protection is not carried over, so be careful with blue cells.",
             f["step"])
    r += 2
    ws.write(r, 2, "Tips from practice", f["label_b"])
    r += 1
    for tp in tips:
        ws.set_row(r, 18)
        ws.write(r, 1, "•", f["bullet"])
        ws.write(r, 2, tp, f["step"])
        r += 1
    r += 1
    if extra_note:
        ws.set_row(r, 30)
        ws.write(r, 2, extra_note, f["step"])
        r += 2
    ws.write(r, 2, "Terms of use", f["label_b"])
    r += 1
    ws.set_row(r, 44)
    ws.write(r, 2, "Licensed for use in the purchaser's own business. Please don't resell, share or redistribute "
                   "the file. The default rates are examples to start from; results depend on the numbers you "
                   f"enter. This is not tax, legal or financial advice — check {tax} rules with your tax authority "
                   "or accountant.", f["step"])
    r += 2
    ws.write(r, 2, "Questions or ideas? Message us through Etsy — we usually reply within 1–2 working days.",
             f["subtitle"])
    ws.activate()


def log_sheet(b, ws, title, subtitle, heads, widths, tiles, samples, validations, status_col, followup_col,
              first=9, last=208, date_cols=(), num_cols=()):
    """Generic tracking log with KPI tiles. Columns are 1-based from B. `tiles` = [(label, formula, fmt, span)]."""
    f = b.f
    ws.hide_gridlines(2)
    for i, w in enumerate(widths):
        ws.set_column(i, i, w)
    ws.set_row(0, 30)
    ws.write("B1", title, f["title"])
    ws.write("B2", subtitle, f["subtitle"])
    c = 1
    for lab, fml, fk, span in tiles:
        if span == 1:
            ws.write(3, c, lab, f["tile_label"])
            ws.write_formula(4, c, fml, f[fk])
        else:
            ws.merge_range(3, c, 3, c + span - 1, lab, f["tile_label"])
            ws.merge_range(4, c, 4, c + span - 1, "", f[fk])
            ws.write_formula(4, c, fml, f[fk])
        c += span
    ws.set_row(3, 28)
    ws.set_row(4, 28)
    ws.write_row(8, 1, heads, f["th"])
    ws.set_row(8, 30)
    ncols = len(heads)
    for i in range(first, last + 1):
        for cc in range(1, ncols + 1):
            fm = f["log_date"] if cc in date_cols else f["log_num"] if cc in num_cols else f["log_text"]
            ws.write_blank(i, cc, None, fm)
    for i, row in enumerate(samples):
        for cc, v in enumerate(row, start=1):
            if v is None:
                continue
            if cc in date_cols:
                ws.write_datetime(first + i, cc, dt.datetime.combine(v, dt.time()), f["log_date"])
            else:
                ws.write(first + i, cc, v, f["log_num"] if cc in num_cols else f["log_text"])
    for cc, source in validations:
        ws.data_validation(first, cc, last, cc, {"validate": "list", "source": source})
    sc, fc = col(status_col), col(followup_col)
    rng = (first, status_col, last, status_col)
    ws.conditional_format(*rng, {"type": "cell", "criteria": "==", "value": '"Won"', "format": f["green"]})
    ws.conditional_format(*rng, {"type": "cell", "criteria": "==", "value": '"Lost"', "format": f["red"]})
    ws.conditional_format(*rng, {"type": "cell", "criteria": "==", "value": '"Follow-up"', "format": f["amber"]})
    ws.conditional_format(first, followup_col, last, followup_col,
                          {"type": "formula",
                           "criteria": f'=AND(${fc}{first + 1}<>"",${fc}{first + 1}<TODAY(),'
                                       f'OR(${sc}{first + 1}="Sent",${sc}{first + 1}="Follow-up"))',
                           "format": f["overdue"]})
    ws.freeze_panes(first, 0)


def log_range(col_idx, first=9, last=208):
    return f"{col(col_idx)}{first + 1}:{col(col_idx)}{last + 1}"


def break_even_sheet(b, ws, tax, fixed, hours_week, weeks):
    """Break-even sheet. Needs names LabourCost, Supplies, Overhead, Margin, RatePerHr. Defines BERate etc."""
    f = b.f
    ws.hide_gridlines(2)
    ws.set_column("A:A", 2)
    ws.set_column("B:B", 48)
    ws.set_column("C:C", 18)
    ws.set_column("D:D", 62)
    ws.set_row(0, 30)
    ws.write("B1", "Break-even & Hourly Rate Check", f["title"])
    ws.write("B2", "The minimum you must earn per labour hour to cover every cost — and whether your prices do.",
             f["subtitle"])
    b.section(ws, 3, f"Monthly fixed costs (excl. {tax})")
    rr = 4
    fx0 = rr
    for lab, amt in fixed:
        ws.write(rr, 1, lab, f["in_text"])
        ws.write(rr, 2, amt, f["in_num"])
        rr += 1
    ws.write(rr, 1, "Total fixed costs per month", f["label_b"])
    ws.write_formula(rr, 2, f"=SUM({cell(fx0, 2)}:{cell(rr - 1, 2)})", f["c_num_b"])
    b.name("FixedMonth", ws.name, rr, 2)
    rr += 2
    b.section(ws, rr, "Your capacity")
    rr += 1
    ws.write(rr, 1, "Paid cleaning hours per week (whole team)", f["label"])
    ws.write(rr, 2, hours_week, f["in_num0"])
    hw = cell(rr, 2)
    rr += 1
    ws.write(rr, 1, "Working weeks per year", f["label"])
    ws.write(rr, 2, weeks, f["in_num0"])
    wk = cell(rr, 2)
    rr += 1
    ws.write(rr, 1, "Paid cleaning hours per month", f["label"])
    ws.write_formula(rr, 2, f"={hw}*{wk}/12", f["c_num0"])
    b.name("HoursMonth", ws.name, rr, 2)
    rr += 2
    b.section(ws, rr, "Results")
    rr += 1
    rows = [
        ("Fixed costs per labour hour", "=IF(HoursMonth>0,FixedMonth/HoursMonth,0)", "c_num", "FixedPerHr"),
        ("Labour cost per hour (from Settings)", "=LabourCost", "c_num", None),
        ("Supplies per hour (from Settings)", "=Supplies", "c_num", None),
        (f"Break-even rate per labour hour (excl. {tax})", "=LabourCost+Supplies+FixedPerHr", "c_num_b", "BERate"),
        ("Target rate with your profit margin", "=IF(Margin<1,BERate/(1-Margin),BERate)", "c_num", None),
        ("Revenue needed per month just to break even", "=BERate*HoursMonth", "c_num", None),
    ]
    for lab, fml, fk, nm in rows:
        ws.write(rr, 1, lab, f["label_b"] if nm == "BERate" else f["label"])
        ws.write_formula(rr, 2, fml, f[fk])
        if nm:
            b.name(nm, ws.name, rr, 2)
        rr += 1
    ws.write(rr, 1, "Overheads per labour hour currently in Settings", f["label"])
    ws.write_formula(rr, 2, "=Overhead", f["c_num"])
    ws.write_formula(rr, 3, '=IF(ABS(Overhead-FixedPerHr)>0.5,"Tip: set Settings → Overheads per labour hour to "'
                            '&TEXT(FixedPerHr,"0.00")&" to match your real fixed costs.","✔ Settings match your '
                            'fixed costs.")', f["note"])
    rr += 1
    ws.write(rr, 1, "Your price per labour hour from Settings", f["label"])
    ws.write_formula(rr, 2, "=RatePerHr", f["c_num"])
    ws.write_formula(rr, 3, '=IF(RatePerHr<BERate,"⚠ Your settings price below break-even — raise your margin or '
                            'cut costs.","✔ Your settings price above break-even.")', f["note"])
    rr += 2
    ws.merge_range(rr, 1, rr + 1, 3, "Owner-operator? Put the hourly wage you want to take home as 'Cleaner pay per "
                                      "hour' in Settings, so break-even already includes paying yourself.", f["note"])
    ws.protect("", {"select_locked_cells": True, "select_unlocked_cells": True})
