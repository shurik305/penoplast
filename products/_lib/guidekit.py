"""Shared two-page Quick-Start Guide PDF generator (reportlab + Inter)."""
import os

from reportlab.lib.colors import HexColor
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

FONTS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "_assets", "fonts")
for _w in (400, 600, 800):
    pdfmetrics.registerFont(TTFont(f"Inter{_w}", os.path.join(FONTS, f"Inter-{_w}.ttf")))
NAVY, BLUE, INK, GREY, SOFT = (HexColor(h) for h in ("#17324D", "#1F5FA6", "#2B4A6B", "#5B6B7B", "#EAF1FA"))
PW, PH = A4
M = 20 * mm


def _para(c, text, x, y, width, font="Inter400", size=10, color=INK, leading=1.4):
    c.setFont(font, size)
    c.setFillColor(color)
    line = ""
    for wd in text.split():
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


def _heading(c, text, y, size=15):
    c.setFont("Inter800", size)
    c.setFillColor(NAVY)
    c.drawString(M, y, text)
    return y - size * 1.2


def _footer(c, product, page):
    c.setFont("Inter400", 8.5)
    c.setFillColor(GREY)
    c.drawString(M, 12 * mm, f"{product} · Quick-Start Guide · by Hutsol")
    c.drawRightString(PW - M, 12 * mm, f"Page {page}")


def build(out, product, title_lines, intro, shot, files, steps, sections):
    """files: [(filename, description)]; steps: [(title, text)]; sections: [(heading, [paragraphs])].

    Text must avoid glyphs missing from Inter's latin subset (e.g. use '>' instead of an arrow).
    """
    c = canvas.Canvas(out, pagesize=A4)
    c.setTitle(f"{product} — Quick-Start Guide")
    c.setAuthor("Hutsol")
    c.setFillColor(SOFT)
    c.rect(0, PH - 95 * mm, PW, 95 * mm, stroke=0, fill=1)
    c.setFillColor(NAVY)
    c.setFont("Inter800", 28)
    for i, ln in enumerate(title_lines):
        c.drawString(M, PH - (32 + 12 * i) * mm, ln)
    c.setFont("Inter600", 13)
    c.setFillColor(BLUE)
    c.drawString(M, PH - 55 * mm, "Quick-Start Guide — set up in 5 minutes")
    _para(c, intro, M, PH - 66 * mm, PW - 2 * M)
    img = ImageReader(shot)
    iw, ih = img.getSize()
    w = PW - 2 * M
    h = min(w * ih / iw, 120 * mm)
    w = h * iw / ih
    c.drawImage(img, M, PH - 100 * mm - h, w, h)
    y = _heading(c, "Which file should I open?", PH - 108 * mm - h)
    for fn, desc in files:
        c.setFont("Inter600", 9.5)
        c.setFillColor(NAVY)
        c.drawString(M + 4 * mm, y, "• " + fn)
        c.setFont("Inter400", 9.5)
        c.setFillColor(INK)
        c.drawString(M + 98 * mm, y, desc)
        y -= 14
    _footer(c, product, 1)
    c.showPage()
    y = _heading(c, "Set up once (5 minutes)", PH - 25 * mm, 18) - 4
    for i, (t, s) in enumerate(steps, 1):
        c.setFillColor(BLUE)
        c.circle(M + 4 * mm, y + 1.2 * mm, 3.6 * mm, stroke=0, fill=1)
        c.setFillColor(HexColor("#FFFFFF"))
        c.setFont("Inter800", 10)
        c.drawCentredString(M + 4 * mm, y - 0.2 * mm, str(i))
        c.setFont("Inter600", 11)
        c.setFillColor(NAVY)
        c.drawString(M + 11 * mm, y, t)
        y = _para(c, s, M + 11 * mm, y - 15, PW - 2 * M - 11 * mm) - 8
    for hd, paras in sections:
        y = _heading(c, hd, y - 6, 18) - 4
        for ptxt in paras:
            y = _para(c, ptxt, M, y, PW - 2 * M) - 6
    _para(c, "Questions or suggestions? Message us through Etsy.",
          M, y - 4, PW - 2 * M, font="Inter600", color=BLUE)
    _footer(c, product, 2)
    c.save()
    return out
