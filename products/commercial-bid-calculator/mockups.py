#!/usr/bin/env python3
"""Etsy listing images for the Commercial Cleaning Bid Calculator. Usage: python3 mockups.py dist/ listing/images/"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "_lib"))
sys.path.insert(0, HERE)
from PIL import ImageDraw  # noqa: E402

import mockkit as mk  # noqa: E402
from build import PRESETS  # noqa: E402

TABS = ["Start Here", "Settings", "Bid Calculator", "Client Proposal", "Break-even", "Bid Log"]
FNAME = "Commercial-Cleaning-Bid-Calculator.xlsx"


def pages(dist, key):
    tm = {"Commercial Cleaning Bid Calculator": "Start Here", "Settings": "Settings", "Bid Calculator": "Bid Calculator",
          PRESETS[key]["biz"]: "Proposal", "Scope of work — tasks each visit": "Proposal",
          "Break-even & Hourly Rate Check": "Break-even", "Bid Log & Win Rate": "Bid Log"}
    return mk.pdf_pages(os.path.join(dist, f"Commercial-Cleaning-Bid-Calculator_{key}.xlsx"), tm)


def hero(pg, out):
    c = mk.gradient()
    d = ImageDraw.Draw(c)
    d.ellipse((1900, -500, 3500, 1100), fill=(226, 236, 248))
    mk.pill(d, 150, 150, "EXCEL + GOOGLE SHEETS · INSTANT DOWNLOAD", 40)
    y = mk.headline(d, 150, 300, ["Commercial Cleaning", "Bid Calculator"], 124)
    for ln in ["Price office contracts by m²:", "monthly fee, team size & profit"]:
        d.text((150, y + 20), ln, font=mk.font(600, 66), fill=mk.BLUE)
        y += 84
    y += 90
    for t in ["12 area types × frequency", "Monthly fee, VAT & annual value", "Team size for your time window",
              "Periodic services & consumables", "2-page client proposal"]:
        mk.check_item(d, 150, y, t, 58)
        y += 108
    mk.laptop(c, mk.trim(pg["Bid Calculator"][0]), 1600, 470, 1250, TABS, "Bid Calculator", FNAME)
    mk.paper(c, mk.trim(pg["Proposal"][0], 60), 1380, 1150, 560, angle=-5)
    mk.footer(d)
    c.save(out, quality=92)


def main(dist, outdir):
    os.makedirs(outdir, exist_ok=True)
    ed = {k: pages(dist, k) for k in ("UK", "EU", "AU", "US")}
    uk = ed["UK"]
    raw = os.path.join(os.path.dirname(outdir.rstrip("/")), "raw")
    os.makedirs(raw, exist_ok=True)
    for k, ims in uk.items():
        for i, im in enumerate(ims):
            mk.trim(im).save(os.path.join(raw, f"{k.lower().replace(' ', '-')}-{i + 1}.png"))
    hero(uk, os.path.join(outdir, "01-hero.jpg"))
    t = mk.trim
    mk.grid([(t(uk["Bid Calculator"][0]), "Bid Calculator", "Hours, monthly fee, team & profit"),
             (t(uk["Proposal"][0]), "Client Proposal", "Fee, scope and areas — page 1"),
             (t(uk["Proposal"][1]), "Scope & Terms", "Tasks per area, terms, signatures"),
             (t(uk["Settings"][0]), "Settings", "Rates, wages, VAT, contract terms"),
             (t(uk["Break-even"][0]), "Break-even", "Your minimum hourly rate"),
             (t(uk["Bid Log"][0]), "Bid Log", "Win rate & monthly pipeline")],
            "What's inside", "6 connected sheets — set up once, then bid any site in minutes.",
            os.path.join(outdir, "02-whats-inside.jpg"))
    mk.feature(t(uk["Bid Calculator"][0]), "Bid any site in minutes",
               "Area × frequency ÷ m² per hour gives hours and the monthly fee.",
               ["Offices, toilets, kitchens, stairs, clinics & more", "Traffic level per area",
                "Windows, carpets, floor care per year", "Warns you below break-even"],
               os.path.join(outdir, "03-calculator.jpg"))
    mk.feature(t(uk["Proposal"][0]), "A proposal clients take seriously",
               "Two A4 pages, filled in automatically. Save as PDF and send.",
               ["Areas, frequencies and hours", "Monthly fee, VAT and annual value",
                "Scope of work for every area type", "Contract term, notice & price review"],
               os.path.join(outdir, "04-proposal.jpg"), landscape=False)
    mk.feature(t(uk["Settings"][0]), "Your rates, your terms",
               "12 area types and 8 periodic services — all editable.",
               ["Production rates in m² per hour", "Wages, on-costs, supplies, overheads",
                "Minimum fee and rounding", "Contract clauses printed for you"],
               os.path.join(outdir, "05-settings.jpg"))
    labels = {"UK": "UK · GBP · VAT 20%", "EU": "EU & Ireland · EUR · VAT", "AU": "Australia & NZ · AUD · GST",
              "US": "US & Canada · USD · sq ft"}
    cards = []
    for k in ("UK", "EU", "AU", "US"):
        im = t(ed[k]["Bid Calculator"][0])
        cards.append((im.crop((0, 0, im.width, int(im.height * 0.6))), labels[k], ""))
    mk.grid(cards, "4 editions included", "Ready-made settings for your country — or set any currency and tax rate.",
            os.path.join(outdir, "06-editions.jpg"), cols=2, card=(1300, 780), gap=(100, 90), caption=False)
    mk.how_it_works([("Download", "Instant download after purchase: 4 editions + a quick-start guide (PDF)."),
                     ("Set up once", "Enter wages, supplies, overheads, margin and tax. Adjust m² per hour rates "
                                     "to your own timings."),
                     ("Bid in minutes", "Add the areas, frequencies and periodic work. Send the 2-page proposal "
                                        "as PDF.")],
                    "Microsoft Excel 2010 or newer · Excel for Mac · Microsoft 365 · Google Sheets · LibreOffice",
                    "No macros, no subscriptions, no sign-ups. An editable spreadsheet you keep.",
                    os.path.join(outdir, "07-how-it-works.jpg"))
    return sorted(os.listdir(outdir))


if __name__ == "__main__":
    print(main(sys.argv[1] if len(sys.argv) > 1 else "dist", sys.argv[2] if len(sys.argv) > 2 else "listing/images"))
