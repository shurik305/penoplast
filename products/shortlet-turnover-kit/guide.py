#!/usr/bin/env python3
"""Quick-Start Guide PDF for the Short-Let Turnover Kit.

Usage: python3 guide.py listing/raw/cleaning-fee-calculator-1.png dist/Quick-Start-Guide.pdf
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "_lib"))
import guidekit  # noqa: E402

if __name__ == "__main__":
    shot, out = sys.argv[1], sys.argv[2]
    print(guidekit.build(
        out, "Short-Let Turnover Kit", ["Short-Let", "Turnover Kit"],
        "Thank you for your purchase! The kit has one workbook for planning and tracking, plus two blank printables "
        "your cleaners can use straight away. Works with any currency — set yours on the Settings sheet.",
        shot,
        [("Short-Let-Turnover-Kit.xlsx", "workbook: settings, log, fees, restock"),
         ("Turnover-Checklist-Printable.pdf", "blank 2-page checklist to print"),
         ("Damage-Issue-Report-Printable.pdf", "blank 1-page report to print")],
        [("Open the Settings sheet.", "Yellow cells are yours to change; blue cells calculate. Set your currency, "
                                      "the supplies cost per turnover and how cleaning fees are rounded."),
         ("Add your properties.", "Up to 10: bedrooms, bathrooms, beds to make, typical guests, how you pay the "
                                  "cleaner (fixed per turnover, or hourly with standard minutes), laundry cost per "
                                  "linen set and the linen sets you own."),
         ("Guest consumables.", "Check the list of toiletries, coffee, tea and cleaning items with their unit cost, "
                                "quantity per stay and whether they scale per stay, guest, bathroom or bedroom."),
         ("Print the checklist.", "Use the Turnover Checklist sheet (choose the property at the top) or the blank "
                                  "PDF. Laminate it or print one per clean."),
         ("Log every clean.", "In Turnover Log enter the date, property, cleaner, start and finish time. Minutes "
                              "and pay are calculated; mark issues and whether the cleaner has been paid.")],
        [("Price your cleaning fee", ["Cleaning Fee Calculator shows what one turnover really costs (cleaner, "
                                      "laundry, consumables, supplies) and the fee that covers it with your mark-up "
                                      "and any platform or payment fee. The table compares all properties."]),
         ("Restock and report", ["Restock List turns a quick cupboard count into a shopping list for the next "
                                 "stays. Use the Issue Report for damage or missing items the same day, with photos."]),
         ("Using Google Sheets", ["Upload the .xlsx file to Google Drive and open it with Google Sheets. Only standard "
                                  "spreadsheet functions are used; sheet protection is not carried over, so take care "
                                  "with blue cells."]),
         ("Licence", ["Licensed for use in your own business. Please do not resell, share or redistribute the files. "
                      "This kit is independent and not affiliated with or endorsed by any booking platform."])]))
