#!/usr/bin/env python3
"""Etsy listing images for the Short-Let Turnover Kit. Usage: python3 mockups.py dist/ listing/images/"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "_lib"))
import pymupdf  # noqa: E402
from PIL import Image, ImageDraw  # noqa: E402

import mockkit as mk  # noqa: E402

TABS = ["Start Here", "Settings", "Turnover Checklist", "Turnover Log", "Cleaning Fee Calculator", "Restock List"]
FNAME = "Short-Let-Turnover-Kit.xlsx"
TITLES = {"Short-Let Turnover Kit": "Start Here", "Settings": "Settings", "Turnover Checklist": "Checklist",
          "Turnover Log": "Log", "Cleaning Fee Calculator": "Fee", "Restock List": "Restock",
          "Damage & Issue Report": "Issue"}


def pdf_page(path, i=0, dpi=200):
    pix = pymupdf.open(path)[i].get_pixmap(dpi=dpi)
    return Image.frombytes("RGB", (pix.width, pix.height), pix.samples)


def main(dist, outdir):
    os.makedirs(outdir, exist_ok=True)
    pg = mk.pdf_pages(os.path.join(dist, "Short-Let-Turnover-Kit.xlsx"), TITLES)
    raw = os.path.join(os.path.dirname(outdir.rstrip("/")), "raw")
    os.makedirs(raw, exist_ok=True)
    for k, ims in pg.items():
        mk.trim(ims[0]).save(os.path.join(raw, {"Fee": "cleaning-fee-calculator"}.get(k, k.lower()) + "-1.png"))
    t = mk.trim
    printable = pdf_page(os.path.join(dist, "Turnover-Checklist-Printable.pdf"))
    report = pdf_page(os.path.join(dist, "Damage-Issue-Report-Printable.pdf"))

    c = mk.gradient()
    d = ImageDraw.Draw(c)
    d.ellipse((1900, -500, 3500, 1100), fill=(226, 236, 248))
    mk.pill(d, 150, 150, "PRINTABLE PDF + EXCEL / GOOGLE SHEETS", 40)
    y = mk.headline(d, 150, 300, ["Short-Let", "Turnover Kit"], 124)
    for ln in ["For Airbnb & holiday-let hosts", "and the cleaners who turn them over"]:
        d.text((150, y + 20), ln, font=mk.font(600, 62), fill=mk.BLUE)
        y += 80
    y += 90
    for item in ["40+ point printable checklist", "Turnover log with cleaner pay", "Cleaning fee calculator",
                 "Restock shopping list", "Damage & issue report"]:
        mk.check_item(d, 150, y, item, 58)
        y += 108
    mk.laptop(c, t(pg["Fee"][0]), 1600, 470, 1250, TABS, "Cleaning Fee Calculator", FNAME)
    mk.paper(c, printable, 1380, 1150, 560, angle=-5)
    mk.footer(d)
    c.save(os.path.join(outdir, "01-hero.jpg"), quality=92)

    mk.grid([(printable, "Turnover Checklist", "Printable, 7 sections, 40+ tasks"),
             (t(pg["Log"][0]), "Turnover Log", "Minutes, pay, issues, unpaid"),
             (t(pg["Fee"][0]), "Cleaning Fee Calculator", "True cost per turnover"),
             (t(pg["Restock"][0]), "Restock List", "From count to shopping list"),
             (report, "Issue Report", "Damage & missing items"),
             (t(pg["Settings"][0]), "Settings", "Up to 10 properties")],
            "What's inside", "One workbook + two printables — everything a turnover needs.",
            os.path.join(outdir, "02-whats-inside.jpg"))
    mk.feature(printable, "A checklist cleaners actually follow",
               "Seven sections, from arrival photos to lock-up. Print or laminate.",
               ["Nothing forgotten between guests", "Same standard from every cleaner",
                "Restock and issues noted on the spot", "Blank PDF + a version per property"],
               os.path.join(outdir, "03-checklist.jpg"), landscape=False)
    mk.feature(t(pg["Fee"][0]), "Know what a turnover really costs",
               "Cleaner, laundry, consumables and supplies — and the fee to charge.",
               ["Fixed or hourly cleaner pay", "Consumables per guest, bathroom, bedroom",
                "Platform / payment fee allowance", "All properties compared"],
               os.path.join(outdir, "04-fee-calculator.jpg"))
    mk.feature(t(pg["Log"][0]), "Every clean, logged",
               "Start and finish times give minutes and pay automatically.",
               ["Turnovers this month", "Average minutes per clean", "Unpaid cleaner pay",
                "Issues flagged in red"], os.path.join(outdir, "05-turnover-log.jpg"))
    mk.feature(t(pg["Restock"][0]), "Never run out mid-stay",
               "Count the cupboard, get the shopping list for the next stays.",
               ["Scales with guests, bathrooms, bedrooms", "Shopping total before you go",
                "15 common items + your own"], os.path.join(outdir, "06-restock.jpg"), landscape=False)
    mk.how_it_works([("Download", "Instant download: the workbook, a printable checklist, an issue report and a "
                                  "quick-start guide (PDF)."),
                     ("Add your properties", "Beds, bathrooms, guests, cleaner pay, laundry and consumables — once."),
                     ("Turn over with confidence", "Print the checklist, log each clean, price your cleaning fee and "
                                                   "restock before you run out.")],
                    "Microsoft Excel 2010 or newer · Excel for Mac · Microsoft 365 · Google Sheets · LibreOffice",
                    "Not affiliated with Airbnb or any booking platform. No macros, no subscriptions.",
                    os.path.join(outdir, "07-how-it-works.jpg"))
    return sorted(os.listdir(outdir))


if __name__ == "__main__":
    print(main(sys.argv[1] if len(sys.argv) > 1 else "dist", sys.argv[2] if len(sys.argv) > 2 else "listing/images"))
