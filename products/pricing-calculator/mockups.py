#!/usr/bin/env python3
"""Renders Etsy listing images (3000x2250) from the real workbook output.

Usage: python3 mockups.py dist/ listing/images/
Needs LibreOffice Calc (with recalculation on load enabled, see ops/setup.sh) and PyMuPDF.
"""
import os
import subprocess
import sys
import tempfile

import pymupdf
from PIL import Image, ImageDraw, ImageFilter, ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))
FONTS = os.path.join(HERE, "..", "_assets", "fonts")
W, H = 3000, 2250
NAVY, BLUE, INK, SOFT = (23, 50, 77), (31, 95, 166), (43, 74, 107), (234, 241, 250)
GOLD = (242, 183, 5)
TABS = ["Start Here", "Settings", "Quote Calculator", "Client Quote", "Price List", "Break-even", "Quote Log"]


def font(weight, size):
    return ImageFont.truetype(os.path.join(FONTS, f"Inter-{weight}.ttf"), size)


def pdf_pages(xlsx, dpi=220):
    """Returns {sheet title: PIL image} for each sheet page of the workbook."""
    env = dict(os.environ, HOME=os.environ.get("LO_HOME", "/tmp/lohome"))
    with tempfile.TemporaryDirectory() as td:
        subprocess.run(["soffice", "--headless", "--norestore", "--convert-to", "pdf", "--outdir", td, xlsx],
                       check=True, capture_output=True, env=env, timeout=240)
        pdf = os.path.join(td, os.path.splitext(os.path.basename(xlsx))[0] + ".pdf")
        doc = pymupdf.open(pdf)
        out = {}
        for pg in doc:
            text = pg.get_text()
            first = text.strip().splitlines()[0] if text.strip() else ""
            key = {"Cleaning Price & Quote Calculator": "Start Here", "Settings": "Settings",
                   "Quote Calculator": "Quote Calculator", "Price List Builder": "Price List",
                   "Break-even & Hourly Rate Check": "Break-even", "Quote Log & Win Rate": "Quote Log"}.get(first)
            if key is None and "QUOTE" in text and "PREPARED FOR" in text:
                key = "Client Quote"
            if key and key not in out:
                pix = pg.get_pixmap(dpi=dpi)
                out[key] = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        return out


def trim(img, pad=30):
    """Crops white margins."""
    gray = img.convert("L").point(lambda v: 255 if v < 250 else 0)
    box = gray.getbbox()
    if not box:
        return img
    l, t, r, b = box
    return img.crop((max(0, l - pad), max(0, t - pad), min(img.width, r + pad), min(img.height, b + pad)))


def gradient(top=(246, 249, 253), bottom=(220, 232, 246)):
    base = Image.new("RGB", (W, H), top)
    d = ImageDraw.Draw(base)
    for y in range(H):
        t = y / H
        d.line([(0, y), (W, y)], fill=tuple(int(top[i] + (bottom[i] - top[i]) * t) for i in range(3)))
    return base


def shadow(canvas, box, radius=40, blur=40, offset=(0, 30), opacity=90):
    x0, y0, x1, y1 = box
    sh = Image.new("L", (W, H), 0)
    ImageDraw.Draw(sh).rounded_rectangle((x0 + offset[0], y0 + offset[1], x1 + offset[0], y1 + offset[1]),
                                         radius=radius, fill=opacity)
    sh = sh.filter(ImageFilter.GaussianBlur(blur))
    canvas.paste((20, 40, 70), (0, 0), sh)


def fit(img, w, h, anchor="top"):
    """Scales img to width w and crops/pads to height h."""
    scale = w / img.width
    img = img.resize((w, int(img.height * scale)), Image.LANCZOS)
    if img.height >= h:
        return img.crop((0, 0, w, h)) if anchor == "top" else img.crop((0, img.height - h, w, img.height))
    out = Image.new("RGB", (w, h), "white")
    out.paste(img, (0, 0))
    return out


def app_window(shot, w, h, active="Quote Calculator", fname="Cleaning-Price-Quote-Calculator.xlsx"):
    """Generic spreadsheet window: title bar, content, sheet tabs (no third-party branding)."""
    win = Image.new("RGB", (w, h), "white")
    d = ImageDraw.Draw(win)
    bar = int(h * 0.055)
    d.rectangle((0, 0, w, bar), fill=(237, 240, 244))
    for i, c in enumerate(((236, 95, 90), (245, 190, 80), (98, 197, 84))):
        cx = int(bar * 0.6) + i * int(bar * 0.55)
        d.ellipse((cx - bar * 0.16, bar / 2 - bar * 0.16, cx + bar * 0.16, bar / 2 + bar * 0.16), fill=c)
    ft = font(500, int(bar * 0.38))
    tw = d.textlength(fname, font=ft)
    d.text(((w - tw) / 2, bar * 0.3), fname, font=ft, fill=(80, 90, 100))
    tabs_h = int(h * 0.06)
    content = fit(shot, w, h - bar - tabs_h)
    win.paste(content, (0, bar))
    ty = h - tabs_h
    d.rectangle((0, ty, w, h), fill=(241, 243, 246))
    d.line((0, ty, w, ty), fill=(210, 215, 222), width=2)
    x = 20
    ftab = font(500, int(tabs_h * 0.36))
    for t in TABS:
        tw = d.textlength(t, font=ftab) + 44
        if x + tw > w:
            break
        if t == active:
            d.rectangle((x, ty, x + tw, h), fill="white")
            d.line((x, ty, x + tw, ty), fill=BLUE, width=6)
        d.text((x + 22, ty + tabs_h * 0.3), t, font=ftab, fill=NAVY if t == active else (95, 105, 115))
        x += tw + 6
    return win


def laptop(canvas, shot, x, y, w, active="Quote Calculator"):
    """Draws a generic laptop with the screenshot. Returns bounding box."""
    sw = w
    sh_ = int(w * 0.625)
    bezel = int(w * 0.03)
    body = (x - bezel, y - bezel, x + sw + bezel, y + sh_ + bezel)
    shadow(canvas, (body[0], body[1], body[2], body[3] + 60), radius=50, blur=50, offset=(0, 40), opacity=110)
    d = ImageDraw.Draw(canvas)
    d.rounded_rectangle(body, radius=int(bezel * 1.4), fill=(34, 38, 44))
    canvas.paste(app_window(shot, sw, sh_, active=active), (x, y))
    base_y = body[3]
    d.rounded_rectangle((body[0] - int(w * 0.07), base_y, body[2] + int(w * 0.07), base_y + int(w * 0.035)),
                        radius=20, fill=(196, 202, 210))
    d.rounded_rectangle((x + sw // 2 - int(w * 0.09), base_y, x + sw // 2 + int(w * 0.09), base_y + int(w * 0.012)),
                        radius=10, fill=(170, 176, 186))
    return body


def paper(canvas, page, x, y, w, angle=-4):
    """A4 sheet with shadow, slightly rotated."""
    h = int(w * 1.414)
    sheet = Image.new("RGB", (w, h), "white")
    pg = page.resize((w, int(page.height * w / page.width)), Image.LANCZOS)
    sheet.paste(pg.crop((0, 0, w, min(h, pg.height))), (0, 0))
    rgba = sheet.convert("RGBA").rotate(angle, expand=True, resample=Image.BICUBIC, fillcolor=(0, 0, 0, 0))
    mask = rgba.split()[3]
    sh = Image.new("L", (W, H), 0)
    sh.paste(mask.filter(ImageFilter.GaussianBlur(1)), (x + 18, y + 34))
    sh = sh.filter(ImageFilter.GaussianBlur(36)).point(lambda v: int(v * 0.45))
    canvas.paste((20, 40, 70), (0, 0), sh)
    canvas.paste(rgba, (x, y), rgba)


def pill(d, x, y, text, size=44, fill=BLUE, fg="white"):
    ft = font(700, size)
    tw = d.textlength(text, font=ft)
    d.rounded_rectangle((x, y, x + tw + size * 1.4, y + size * 1.9), radius=size, fill=fill)
    d.text((x + size * 0.7, y + size * 0.42), text, font=ft, fill=fg)
    return x + tw + size * 1.4


def check_item(d, x, y, text, size=56, color=NAVY):
    r = size * 0.55
    cx, cy = x + r, y + size * 0.62
    d.ellipse((cx - r, cy - r, cx + r, cy + r), fill=BLUE)
    d.line((cx - r * 0.45, cy + r * 0.02, cx - r * 0.1, cy + r * 0.38, cx + r * 0.5, cy - r * 0.35), fill="white",
           width=max(4, int(size * 0.11)), joint="curve")
    d.text((x + r * 2 + 28, y), text, font=font(600, size), fill=color)


def wrap(d, text, ft, width):
    words, lines, cur = text.split(), [], ""
    for wd in words:
        t = (cur + " " + wd).strip()
        if d.textlength(t, font=ft) <= width:
            cur = t
        else:
            lines.append(cur)
            cur = wd
    lines.append(cur)
    return lines


def headline(d, x, y, lines, size, color=NAVY, weight=800, gap=1.12):
    ft = font(weight, size)
    for ln in lines:
        d.text((x, y), ln, font=ft, fill=color)
        y += int(size * gap)
    return y


def footer(d, text="by Hutsol · instant digital download"):
    d.text((150, H - 120), text, font=font(500, 40), fill=(95, 115, 138))


def img_hero(pages, out):
    c = gradient()
    d = ImageDraw.Draw(c)
    d.ellipse((1900, -500, 3500, 1100), fill=(226, 236, 248))
    pill(d, 150, 150, "EXCEL + GOOGLE SHEETS · INSTANT DOWNLOAD", 40)
    y = headline(d, 150, 300, ["Cleaning Price", "& Quote Calculator"], 124)
    ft = font(600, 66)
    for ln in ["Price every clean by m²,", "with VAT, in 60 seconds"]:
        d.text((150, y + 20), ln, font=ft, fill=BLUE)
        y += 84
    y += 90
    for t in ["Metric m² (or sq ft)", "VAT / GST built in", "Printable client quote", "Profit & break-even check",
              "4 regional editions"]:
        check_item(d, 150, y, t, 58)
        y += 108
    laptop(c, trim(pages["Quote Calculator"]), 1600, 470, 1250)
    paper(c, trim(pages["Client Quote"], 60), 1380, 1150, 560, angle=-5)
    footer(d)
    c.save(out, quality=92)


def img_inside(pages, out):
    c = gradient()
    d = ImageDraw.Draw(c)
    headline(d, 150, 130, ["What's inside"], 120)
    d.text((150, 300), "7 connected sheets — fill in your numbers once, quote every job in a minute.",
           font=font(500, 56), fill=INK)
    items = [("Quote Calculator", "Price, hours & profit for each job"), ("Client Quote", "Print-ready A4 quote"),
             ("Price List", "Standard prices by size & frequency"), ("Break-even", "Your minimum hourly rate"),
             ("Settings", "Rates, VAT, margin, extras"), ("Quote Log", "Win rate, pipeline & follow-ups")]
    cw, ch = 860, 700
    for i, (key, cap) in enumerate(items):
        col, row = i % 3, i // 3
        x, y = 150 + col * (cw + 70), 470 + row * (ch + 110)
        shadow(c, (x, y, x + cw, y + ch), radius=30, blur=30, offset=(0, 18), opacity=70)
        d.rounded_rectangle((x, y, x + cw, y + ch), radius=30, fill="white")
        shot = fit(trim(pages[key]), cw - 60, ch - 190)
        c.paste(shot, (x + 30, y + 30))
        d.text((x + 34, y + ch - 145), key, font=font(700, 50), fill=NAVY)
        d.text((x + 34, y + ch - 80), cap, font=font(500, 38), fill=INK)
    c.save(out, quality=92)


def img_feature(pages, key, title, sub, bullets, out, landscape=True):
    c = gradient()
    d = ImageDraw.Draw(c)
    y = headline(d, 150, 130, [title], 110)
    d.text((150, y + 10), sub, font=font(500, 54), fill=INK)
    shot = trim(pages[key])
    if landscape:
        box_w = 1850
        s = fit(shot, box_w, int(box_w * shot.height / shot.width))
        x, yy = 150, 470
        shadow(c, (x, yy, x + s.width, yy + s.height), radius=24, blur=34, offset=(0, 22), opacity=80)
        c.paste(s, (x, yy))
        bx = x + s.width + 110
    else:
        paper(c, shot, 230, 430, 1150, angle=0)
        bx = 1600
    by = 520
    for b in bullets:
        ft = font(600, 54)
        lines = wrap(d, b, ft, W - bx - 150)
        check_item(d, bx, by, lines[0], 54)
        for ln in lines[1:]:
            by += 70
            d.text((bx + 88, by), ln, font=ft, fill=NAVY)
        by += 130
    footer(d)
    c.save(out, quality=92)


def img_editions(edition_pages, out):
    c = gradient()
    d = ImageDraw.Draw(c)
    headline(d, 150, 130, ["4 editions included"], 120)
    d.text((150, 300), "Same tool, ready-made settings for your country — or set any currency and tax rate.",
           font=font(500, 54), fill=INK)
    labels = {"UK": "UK · GBP · VAT 20%", "EU": "EU & Ireland · EUR · VAT", "AU": "Australia & NZ · AUD · GST",
              "US": "US & Canada · USD · sq ft"}
    cw, ch = 1300, 780
    for i, key in enumerate(["UK", "EU", "AU", "US"]):
        col, row = i % 2, i // 2
        x, y = 150 + col * (cw + 100), 470 + row * (ch + 90)
        shadow(c, (x, y, x + cw, y + ch), radius=30, blur=30, offset=(0, 18), opacity=70)
        d.rounded_rectangle((x, y, x + cw, y + ch), radius=30, fill="white")
        shot = trim(edition_pages[key]["Quote Calculator"])
        s = fit(shot.crop((0, 0, shot.width, int(shot.height * 0.62))), cw - 60, ch - 150)
        c.paste(s, (x + 30, y + 30))
        d.text((x + 34, y + ch - 100), labels[key], font=font(700, 52), fill=NAVY)
    c.save(out, quality=92)


def img_how(out):
    c = gradient()
    d = ImageDraw.Draw(c)
    headline(d, 150, 150, ["How it works"], 130)
    steps = [("1", "Download", "Instant download after purchase: 4 editions + a quick-start guide (PDF)."),
             ("2", "Set up once", "Enter your wages, supplies, overheads, margin and VAT. Adjust the m² rates to "
                                  "your own timings."),
             ("3", "Quote in a minute", "Enter area, service, condition and extras. Send the ready A4 quote as "
                                        "PDF.")]
    y = 480
    for n, t, s in steps:
        d.ellipse((150, y, 330, y + 180), fill=BLUE)
        tw = d.textlength(n, font=font(800, 100))
        d.text((240 - tw / 2, y + 25), n, font=font(800, 100), fill="white")
        d.text((400, y + 5), t, font=font(700, 72), fill=NAVY)
        yy = y + 100
        for ln in wrap(d, s, font(500, 52), 2400):
            d.text((400, yy), ln, font=font(500, 52), fill=INK)
            yy += 66
        y = max(y + 330, yy + 80)
    d.rounded_rectangle((150, 1650, 2850, 1990), radius=40, fill="white")
    d.text((230, 1710), "Works in", font=font(700, 56), fill=NAVY)
    d.text((230, 1800), "Microsoft Excel 2010 or newer · Excel for Mac · Microsoft 365 · Google Sheets · "
                        "LibreOffice", font=font(500, 50), fill=INK)
    d.text((230, 1885), "No macros, no subscriptions, no sign-ups. Not an app — an editable spreadsheet you keep.",
           font=font(500, 44), fill=(95, 115, 138))
    c.save(out, quality=92)


def main(dist, outdir):
    os.makedirs(outdir, exist_ok=True)
    ed = {}
    for key in ("UK", "EU", "AU", "US"):
        ed[key] = pdf_pages(os.path.join(dist, f"Cleaning-Price-Quote-Calculator_{key}.xlsx"))
    uk = ed["UK"]
    raw = os.path.join(os.path.dirname(outdir.rstrip("/")), "raw")
    os.makedirs(raw, exist_ok=True)
    for k, im in uk.items():
        trim(im).save(os.path.join(raw, k.lower().replace(" ", "-") + ".png"))
    img_hero(uk, os.path.join(outdir, "01-hero.jpg"))
    img_inside(uk, os.path.join(outdir, "02-whats-inside.jpg"))
    img_feature(uk, "Quote Calculator", "Quote every job in a minute",
                "From area or bedrooms to hours, price, VAT and profit.",
                ["Service, condition, frequency & extras", "Travel beyond your free radius",
                 "Minimum charge and rounding", "Warns you below break-even"], os.path.join(outdir, "03-calculator.jpg"))
    img_feature(uk, "Client Quote", "A professional quote, ready to send",
                "Filled in automatically. Print it or save as PDF.",
                ["Clean layout with your business details", "Line items, VAT and total",
                 "What's included for each service", "Validity date and acceptance line"],
                os.path.join(outdir, "04-client-quote.jpg"), landscape=False)
    img_feature(uk, "Price List", "Your price list in seconds", "By size for 7 services, plus recurring prices.",
                ["Post it on your website or flyers", "Weekly, fortnightly & monthly rates",
                 "Updates when your costs change"], os.path.join(outdir, "05-price-list.jpg"))
    img_feature(uk, "Break-even", "Know your minimum hourly rate",
                "Fixed costs and hours give the rate you must never go below.",
                ["Covers wages, supplies & overheads", "Revenue you need each month",
                 "Checks your settings against reality"], os.path.join(outdir, "06-break-even.jpg"), landscape=False)
    img_editions(ed, os.path.join(outdir, "07-editions.jpg"))
    img_how(os.path.join(outdir, "08-how-it-works.jpg"))
    return sorted(os.listdir(outdir))


if __name__ == "__main__":
    print(main(sys.argv[1] if len(sys.argv) > 1 else "dist", sys.argv[2] if len(sys.argv) > 2 else "listing/images"))
