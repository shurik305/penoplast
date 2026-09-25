#!/usr/bin/env python3
"""Builds the Quick-Start Guide PDF delivered with the calculator.

Usage: python3 guide.py listing/images/03-calculator.jpg dist/Quick-Start-Guide.pdf
"""
import os
import sys

from reportlab.lib.colors import HexColor
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

HERE = os.path.dirname(os.path.abspath(__file__))
FONTS = os.path.join(HERE, "..", "_assets", "fonts")
for w in (400, 600, 800):
    pdfmetrics.registerFont(TTFont(f"Inter{w}", os.path.join(FONTS, f"Inter-{w}.ttf")))
NAVY, BLUE, INK, GREY, SOFT = (HexColor(h) for h in ("#17324D", "#1F5FA6", "#2B4A6B", "#5B6B7B", "#EAF1FA"))
PW, PH = A4
M = 20 * mm


def para(c, text, x, y, width, font="Inter400", size=10.5, color=INK, leading=1.45):
    c.setFont(font, size)
    c.setFillColor(color)
    words, line = text.split(), ""
    for wd in words:
        t = (line + " " + wd).strip()
        if pdfmetrics.stringWidth(t, font, size) <= width:
            line = t
        else:
            c.drawString(x, y, line)
            y -= size * leading
            line = wd
    if line:
        c.drawString(x, y, line)
        y -= size * leading
    return y


def heading(c, text, y, size=15):
    c.setFont("Inter800", size)
    c.setFillColor(NAVY)
    c.drawString(M, y, text)
    return y - size * 1.2


def footer(c, page):
    c.setFont("Inter400", 8.5)
    c.setFillColor(GREY)
    c.drawString(M, 12 * mm, "Cleaning Price & Quote Calculator · Quick-Start Guide · by Hutsol")
    c.drawRightString(PW - M, 12 * mm, f"Page {page}")


def build(shot, out):
    c = canvas.Canvas(out, pagesize=A4)
    c.setTitle("Cleaning Price & Quote Calculator — Quick-Start Guide")
    c.setAuthor("Hutsol")
    # page 1
    c.setFillColor(SOFT)
    c.rect(0, PH - 95 * mm, PW, 95 * mm, stroke=0, fill=1)
    c.setFillColor(NAVY)
    c.setFont("Inter800", 28)
    c.drawString(M, PH - 32 * mm, "Cleaning Price &")
    c.drawString(M, PH - 44 * mm, "Quote Calculator")
    c.setFont("Inter600", 13)
    c.setFillColor(BLUE)
    c.drawString(M, PH - 55 * mm, "Quick-Start Guide — set up in 5 minutes")
    y = para(c, "Thank you for your purchase! Your download contains four ready-made editions of the same tool. "
                "Open the one that matches your country — you can change the currency and tax rate in any of them.",
             M, PH - 66 * mm, PW - 2 * M)
    img = ImageReader(shot)
    iw, ih = img.getSize()
    w = PW - 2 * M
    h = w * ih / iw
    c.drawImage(img, M, PH - 100 * mm - h, w, h)
    y = PH - 108 * mm - h
    y = heading(c, "Which file should I open?", y)
    files = [("Cleaning-Price-Quote-Calculator_UK.xlsx", "GBP, VAT 20%, miles"),
             ("Cleaning-Price-Quote-Calculator_EU.xlsx", "EUR, VAT (set your country's rate), km"),
             ("Cleaning-Price-Quote-Calculator_AU.xlsx", "AUD, GST 10%, km (NZ: set NZD and 15%)"),
             ("Cleaning-Price-Quote-Calculator_US.xlsx", "USD, sales tax off by default, sq ft, miles")]
    for fn, desc in files:
        c.setFont("Inter600", 10)
        c.setFillColor(NAVY)
        c.drawString(M + 4 * mm, y, "• " + fn)
        c.setFont("Inter400", 10)
        c.setFillColor(INK)
        c.drawString(M + 100 * mm, y, desc)
        y -= 15
    footer(c, 1)
    c.showPage()
    # page 2
    y = PH - 25 * mm
    y = heading(c, "Set up once (5 minutes)", y, 18) - 4
    steps = [
        ("Open the Settings sheet.", "Yellow cells are yours to change; blue cells calculate. Enter your business "
                                     "name and contact details — they appear on every client quote."),
        ("Currency and tax.", "Type your currency, your VAT/GST rate and whether you are registered. If you are not "
                              "registered, choose No and prices are calculated without tax."),
        ("Your costs and margin.", "Enter the hourly pay (or the wage you want to pay yourself), on-costs, supplies "
                                   "and overheads per hour, and your target margin. Use the Break-even sheet to "
                                   "work out overheads from your real monthly costs."),
        ("Production rates.", "Each service has an m² per cleaner-hour rate. They are starting points — time your "
                              "next three jobs and adjust. This is what makes your quotes accurate."),
        ("Minimums, travel and rounding.", "Set a minimum charge and minimum hours per visit, a free travel radius "
                                           "and a charge per km/mile beyond it, and how to round prices."),
    ]
    for i, (t, s) in enumerate(steps, 1):
        c.setFillColor(BLUE)
        c.circle(M + 4 * mm, y + 1.2 * mm, 3.6 * mm, stroke=0, fill=1)
        c.setFillColor(HexColor("#FFFFFF"))
        c.setFont("Inter800", 10)
        c.drawCentredString(M + 4 * mm, y - 0.2 * mm, str(i))
        c.setFont("Inter600", 11)
        c.setFillColor(NAVY)
        c.drawString(M + 11 * mm, y, t)
        y = para(c, s, M + 11 * mm, y - 15, PW - 2 * M - 11 * mm) - 8
    y = heading(c, "Quote a job", y - 6, 18) - 4
    y = para(c, "Go to Quote Calculator. Enter the client, the service type, the floor area (or choose the number "
                "of bedrooms), condition, frequency, number of cleaners, travel distance and any extras. The price, "
                "hours on site, price per m² and your profit update instantly. The status line is green when the "
                "price meets your target margin, amber when it only covers your costs (usually because of "
                "discounts), and red when it is below break-even or an item needs choosing again.", M, y, PW - 2 * M) - 6
    y = para(c, "Then open Client Quote — it is already filled in. Use File > Print > Save as PDF and send it to "
                "your client. Record the quote in Quote Log to track your win rate and follow-ups.", M, y,
             PW - 2 * M) - 10
    y = heading(c, "Using Google Sheets", y, 18) - 4
    y = para(c, "Upload the .xlsx file to Google Drive, then open it with Google Sheets (or in Sheets: File > "
                "Import > Upload). Only standard spreadsheet functions are used. Sheet protection is not carried over to Google Sheets, so take "
                "care not to type over the blue calculated cells.", M, y, PW - 2 * M) - 10
    y = heading(c, "Licence and support", y, 18) - 4
    y = para(c, "Licensed for use in your own business. Please do not resell, share or redistribute the files. "
                "The default rates are examples — results depend on the numbers you enter. This is not tax, legal "
                "or financial advice; check VAT/GST rules with your tax authority or accountant.", M, y, PW - 2 * M)
    y = para(c, "Questions or suggestions? Message us through Etsy.",
             M, y - 6, PW - 2 * M, font="Inter600", color=BLUE)
    footer(c, 2)
    c.save()
    return out


if __name__ == "__main__":
    print(build(sys.argv[1], sys.argv[2]))
