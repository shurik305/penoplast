#!/usr/bin/env python3
"""Quick-Start Guide PDF for the Commercial Cleaning Bid Calculator.

Usage: python3 guide.py listing/raw/bid-calculator-1.png dist/Quick-Start-Guide.pdf
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "_lib"))
import guidekit  # noqa: E402

LICENCE = ("Licensed for use in your own business. Please do not resell, share or redistribute the files. The default "
           "rates are examples — results depend on the numbers you enter. This is not tax, legal or financial advice; "
           "check VAT/GST rules with your tax authority or accountant, and have your own contract terms reviewed.")

if __name__ == "__main__":
    shot, out = sys.argv[1], sys.argv[2]
    print(guidekit.build(
        out, "Commercial Cleaning Bid Calculator", ["Commercial Cleaning", "Bid Calculator"],
        "Thank you for your purchase! Your download contains four ready-made editions of the same tool. Open the one "
        "that matches your country — you can change the currency, tax rate and area unit in any of them.",
        shot,
        [("Commercial-Cleaning-Bid-Calculator_UK.xlsx", "GBP, VAT 20%, m²"),
         ("Commercial-Cleaning-Bid-Calculator_EU.xlsx", "EUR, VAT (set your country's rate), m²"),
         ("Commercial-Cleaning-Bid-Calculator_AU.xlsx", "AUD, GST 10%, m² (NZ: set NZD and 15%)"),
         ("Commercial-Cleaning-Bid-Calculator_US.xlsx", "USD, sales tax off by default, sq ft")],
        [("Open the Settings sheet.", "Yellow cells are yours to change; blue cells calculate. Enter your business "
                                      "details — they appear on every proposal."),
         ("Currency and tax.", "Type your currency and VAT/GST rate, say whether you are registered, and choose "
                               "whether the monthly fee is shown including tax (business clients usually compare "
                               "prices excluding tax)."),
         ("Costs and margin.", "Enter cleaner pay, on-costs, supervision time, supplies and overheads per hour, and "
                               "your target margin. The Break-even sheet works out overheads from your real costs."),
         ("Production rates.", "Each area type has an m² per cleaner-hour rate for routine cleaning. They are "
                               "starting points — time a few real visits and adjust. Toilets and kitchens are slow; "
                               "corridors and warehouses are fast."),
         ("Contract terms.", "Set the minimum monthly fee, rounding, consumables mark-up, contract term, notice "
                             "period and the clauses printed on the proposal.")],
        [("Price a bid", ["Go to Bid Calculator. Enter the client and site, the cleaning window, then each area with "
                          "its size, frequency and traffic level. Add periodic services (windows, carpets, floor care) "
                          "with how many times a year, and any consumables you will supply.",
                          "Read the monthly fee, annual value, weekly hours, the team you need for the cleaning "
                          "window and your monthly profit. A red message warns you when the fee falls below your "
                          "break-even rate."]),
         ("Send the proposal", ["Open Client Proposal — both pages are filled in. Use File > Print > Save as PDF and "
                                "send it. Record the bid in Bid Log to track your win rate and follow-ups."]),
         ("Using Google Sheets", ["Upload the .xlsx file to Google Drive and open it with Google Sheets (or File > "
                                  "Import > Upload in Sheets). All formulas work; sheet protection is not carried over, "
                                  "so take care not to type over the blue cells."]),
         ("Licence", [LICENCE])]))
