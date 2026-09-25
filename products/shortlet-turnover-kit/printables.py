#!/usr/bin/env python3
"""Blank printable PDFs: Turnover Checklist (A4) and Damage & Issue Report (A4).

Usage: python3 printables.py dist/
Task list comes from build.py so the workbook and the PDF never drift apart.
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(HERE, "..", "_lib"))
import guidekit  # noqa: E402,F401  (registers the Inter fonts)
from build import CHECKLIST  # noqa: E402
from reportlab.lib.colors import HexColor  # noqa: E402
from reportlab.lib.pagesizes import A4  # noqa: E402
from reportlab.lib.units import mm  # noqa: E402
from reportlab.pdfgen import canvas  # noqa: E402

NAVY, BLUE, INK, GREY, LINE = (HexColor(h) for h in ("#17324D", "#1F5FA6", "#2B4A6B", "#5B6B7B", "#9AA8B6"))
PW, PH = A4
M = 16 * mm


def field_line(c, x, y, label, width):
    c.setFont("Inter600", 9.5)
    c.setFillColor(NAVY)
    c.drawString(x, y, label)
    lw = c.stringWidth(label, "Inter600", 9.5) + 3 * mm
    c.setStrokeColor(LINE)
    c.setLineWidth(0.6)
    c.line(x + lw, y - 1.2 * mm, x + width, y - 1.2 * mm)


def footer(c, text):
    c.setFont("Inter400", 8)
    c.setFillColor(GREY)
    c.drawString(M, 9 * mm, text)


def checklist(out):
    c = canvas.Canvas(out, pagesize=A4)
    c.setTitle("Short-Let Turnover Checklist")
    c.setAuthor("Hutsol")

    def header(page):
        c.setFont("Inter800", 20)
        c.setFillColor(NAVY)
        c.drawString(M, PH - 20 * mm, "Turnover Checklist")
        c.setFont("Inter400", 9)
        c.setFillColor(GREY)
        c.drawRightString(PW - M, PH - 20 * mm, f"page {page} of 2")
        y = PH - 30 * mm
        if page == 1:
            half = (PW - 2 * M - 8 * mm) / 2
            field_line(c, M, y, "Property:", half)
            field_line(c, M + half + 8 * mm, y, "Date:", half)
            y -= 8 * mm
            field_line(c, M, y, "Cleaner:", half)
            field_line(c, M + half + 8 * mm, y, "Guest checked out:", half)
            y -= 8 * mm
            field_line(c, M, y, "Next check-in:", half)
            field_line(c, M + half + 8 * mm, y, "Next guests:", half)
            y -= 10 * mm
        return y

    page = 1
    y = header(page)
    for sec, tasks in CHECKLIST:
        need = 9 * mm + len(tasks) * 6.4 * mm
        if y - need < 24 * mm:
            footer(c, "Short-Let Turnover Kit · by Hutsol")
            c.showPage()
            page += 1
            y = header(page)
        c.setFillColor(BLUE)
        c.roundRect(M, y - 2.2 * mm, PW - 2 * M, 7 * mm, 1.5 * mm, stroke=0, fill=1)
        c.setFillColor(HexColor("#FFFFFF"))
        c.setFont("Inter600", 10.5)
        c.drawString(M + 3 * mm, y, sec)
        y -= 9 * mm
        for t in tasks:
            c.setStrokeColor(BLUE)
            c.setLineWidth(0.9)
            c.rect(M + 1.5 * mm, y - 1 * mm, 4.2 * mm, 4.2 * mm, stroke=1, fill=0)
            c.setFont("Inter400", 10)
            c.setFillColor(INK)
            c.drawString(M + 9 * mm, y, t)
            y -= 6.4 * mm
        y -= 3 * mm
    c.setFont("Inter600", 10.5)
    c.setFillColor(NAVY)
    c.drawString(M, y, "Restock needed / issues found")
    y -= 3 * mm
    for _ in range(4):
        y -= 7 * mm
        c.setStrokeColor(LINE)
        c.line(M, y, PW - M, y)
    y -= 10 * mm
    half = (PW - 2 * M - 8 * mm) / 2
    field_line(c, M, y, "Finished at:", half)
    field_line(c, M + half + 8 * mm, y, "Cleaner signature:", half)
    footer(c, "Short-Let Turnover Kit · by Hutsol")
    c.save()
    return page


def issue_report(out):
    c = canvas.Canvas(out, pagesize=A4)
    c.setTitle("Damage & Issue Report")
    c.setAuthor("Hutsol")
    c.setFont("Inter800", 20)
    c.setFillColor(NAVY)
    c.drawString(M, PH - 20 * mm, "Damage & Issue Report")
    c.setFont("Inter400", 9.5)
    c.setFillColor(GREY)
    c.drawString(M, PH - 27 * mm, "Record it the same day, with photos — hosts and platforms usually need evidence "
                                  "quickly to support a claim.")
    y = PH - 40 * mm
    for fl in ["Property", "Date found", "Found by", "Booking / reservation reference", "Guest check-out date",
               "Room / location", "Photos taken (how many, file names)", "Estimated repair / replacement cost",
               "Reported to host on", "Action taken"]:
        field_line(c, M, y, fl + ":", PW - 2 * M)
        y -= 10 * mm
    c.setFont("Inter600", 9.5)
    c.setFillColor(NAVY)
    c.drawString(M, y, "Issue type:")
    x = M + 25 * mm
    for opt in ["Damage", "Missing item", "Maintenance", "Cleanliness", "Other"]:
        c.setStrokeColor(BLUE)
        c.rect(x, y - 1 * mm, 4 * mm, 4 * mm, stroke=1, fill=0)
        c.setFont("Inter400", 9.5)
        c.setFillColor(INK)
        c.drawString(x + 6 * mm, y, opt)
        x += 6 * mm + c.stringWidth(opt, "Inter400", 9.5) + 8 * mm
    y -= 12 * mm
    c.setFont("Inter600", 10.5)
    c.setFillColor(NAVY)
    c.drawString(M, y, "Description")
    c.setStrokeColor(LINE)
    c.rect(M, y - 78 * mm, PW - 2 * M, 74 * mm, stroke=1, fill=0)
    y -= 92 * mm
    half = (PW - 2 * M - 8 * mm) / 2
    field_line(c, M, y, "Signature:", half)
    field_line(c, M + half + 8 * mm, y, "Date:", half)
    footer(c, "Short-Let Turnover Kit · by Hutsol")
    c.save()


if __name__ == "__main__":
    d = sys.argv[1] if len(sys.argv) > 1 else "dist"
    pages = checklist(os.path.join(d, "Turnover-Checklist-Printable.pdf"))
    issue_report(os.path.join(d, "Damage-Issue-Report-Printable.pdf"))
    print("checklist pages:", pages)
