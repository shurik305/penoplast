"""Shared drawing helpers for Etsy listing images (3000x2250) built from real workbook renders."""
import os
import subprocess
import tempfile

import pymupdf
from PIL import Image, ImageDraw, ImageFilter, ImageFont

FONTS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "_assets", "fonts")
W, H = 3000, 2250
NAVY, BLUE, INK = (23, 50, 77), (31, 95, 166), (43, 74, 107)
MUTED = (95, 115, 138)


def font(weight, size):
    return ImageFont.truetype(os.path.join(FONTS, f"Inter-{weight}.ttf"), size)


def missing_glyphs(text):
    """Characters the bundled Inter subset cannot draw (they would render as empty boxes)."""
    from fontTools.ttLib import TTFont
    cmap = TTFont(os.path.join(FONTS, "Inter-400.ttf")).getBestCmap()
    return sorted({ch for ch in text if not ch.isspace() and ord(ch) not in cmap})


def pdf_pages(xlsx, title_map, dpi=220):
    """Renders the workbook via LibreOffice. title_map: {first text line on the page: key}.

    Pages whose first line is not in the map are ignored; the first page seen for each key wins.
    Returns {key: [PIL images]} so multi-page documents keep all their pages.
    """
    env = dict(os.environ, HOME=os.environ.get("LO_HOME", "/tmp/lohome"))
    with tempfile.TemporaryDirectory() as td:
        subprocess.run(["soffice", "--headless", "--norestore", "--convert-to", "pdf", "--outdir", td, xlsx],
                       check=True, capture_output=True, env=env, timeout=240)
        doc = pymupdf.open(os.path.join(td, os.path.splitext(os.path.basename(xlsx))[0] + ".pdf"))
        out, last = {}, None
        for pg in doc:
            lines = pg.get_text().strip().splitlines()
            first = lines[0] if lines else ""
            key = title_map.get(first)
            if key is None and last and callable(title_map.get("__continuation__")):
                key = title_map["__continuation__"](last, first)
            if key:
                pix = pg.get_pixmap(dpi=dpi)
                out.setdefault(key, []).append(Image.frombytes("RGB", (pix.width, pix.height), pix.samples))
                last = key
        return out


def trim(img, pad=30):
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
    canvas.paste((20, 40, 70), (0, 0), sh.filter(ImageFilter.GaussianBlur(blur)))


def fit(img, w, h):
    img = img.resize((w, int(img.height * w / img.width)), Image.LANCZOS)
    if img.height >= h:
        return img.crop((0, 0, w, h))
    out = Image.new("RGB", (w, h), "white")
    out.paste(img, (0, 0))
    return out


def app_window(shot, w, h, tabs, active, fname):
    """Generic spreadsheet window (title bar, content, sheet tabs) without third-party branding."""
    win = Image.new("RGB", (w, h), "white")
    d = ImageDraw.Draw(win)
    bar = int(h * 0.055)
    d.rectangle((0, 0, w, bar), fill=(237, 240, 244))
    for i, c in enumerate(((236, 95, 90), (245, 190, 80), (98, 197, 84))):
        cx = int(bar * 0.6) + i * int(bar * 0.55)
        d.ellipse((cx - bar * 0.16, bar / 2 - bar * 0.16, cx + bar * 0.16, bar / 2 + bar * 0.16), fill=c)
    ft = font(500, int(bar * 0.38))
    d.text(((w - d.textlength(fname, font=ft)) / 2, bar * 0.3), fname, font=ft, fill=(80, 90, 100))
    tabs_h = int(h * 0.06)
    win.paste(fit(shot, w, h - bar - tabs_h), (0, bar))
    ty = h - tabs_h
    d.rectangle((0, ty, w, h), fill=(241, 243, 246))
    d.line((0, ty, w, ty), fill=(210, 215, 222), width=2)
    x = 20
    ftab = font(500, int(tabs_h * 0.36))
    for t in tabs:
        tw = d.textlength(t, font=ftab) + 44
        if x + tw > w:
            break
        if t == active:
            d.rectangle((x, ty, x + tw, h), fill="white")
            d.line((x, ty, x + tw, ty), fill=BLUE, width=6)
        d.text((x + 22, ty + tabs_h * 0.3), t, font=ftab, fill=NAVY if t == active else (95, 105, 115))
        x += tw + 6
    return win


def laptop(canvas, shot, x, y, w, tabs, active, fname):
    sh_ = int(w * 0.625)
    bezel = int(w * 0.03)
    body = (x - bezel, y - bezel, x + w + bezel, y + sh_ + bezel)
    shadow(canvas, (body[0], body[1], body[2], body[3] + 60), radius=50, blur=50, offset=(0, 40), opacity=110)
    d = ImageDraw.Draw(canvas)
    d.rounded_rectangle(body, radius=int(bezel * 1.4), fill=(34, 38, 44))
    canvas.paste(app_window(shot, w, sh_, tabs, active, fname), (x, y))
    base_y = body[3]
    d.rounded_rectangle((body[0] - int(w * 0.07), base_y, body[2] + int(w * 0.07), base_y + int(w * 0.035)),
                        radius=20, fill=(196, 202, 210))
    d.rounded_rectangle((x + w // 2 - int(w * 0.09), base_y, x + w // 2 + int(w * 0.09), base_y + int(w * 0.012)),
                        radius=10, fill=(170, 176, 186))
    return body


def paper(canvas, page, x, y, w, angle=-4):
    h = int(w * 1.414)
    sheet = Image.new("RGB", (w, h), "white")
    pg = page.resize((w, int(page.height * w / page.width)), Image.LANCZOS)
    sheet.paste(pg.crop((0, 0, w, min(h, pg.height))), (0, 0))
    rgba = sheet.convert("RGBA").rotate(angle, expand=True, resample=Image.BICUBIC, fillcolor=(0, 0, 0, 0))
    sh = Image.new("L", (W, H), 0)
    sh.paste(rgba.split()[3].filter(ImageFilter.GaussianBlur(1)), (x + 18, y + 34))
    canvas.paste((20, 40, 70), (0, 0), sh.filter(ImageFilter.GaussianBlur(36)).point(lambda v: int(v * 0.45)))
    canvas.paste(rgba, (x, y), rgba)


def pill(d, x, y, text, size=44, fill=BLUE, fg="white"):
    ft = font(700, size)
    tw = d.textlength(text, font=ft)
    d.rounded_rectangle((x, y, x + tw + size * 1.4, y + size * 1.9), radius=size, fill=fill)
    d.text((x + size * 0.7, y + size * 0.42), text, font=ft, fill=fg)


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
    d.text((150, H - 120), text, font=font(500, 40), fill=MUTED)


def bullets(d, x, y, items, size=54, width=None):
    for b in items:
        ft = font(600, size)
        lines = wrap(d, b, ft, width or (W - x - 150))
        check_item(d, x, y, lines[0], size)
        for ln in lines[1:]:
            y += int(size * 1.3)
            d.text((x + int(size * 1.63), y), ln, font=ft, fill=NAVY)
        y += int(size * 2.4)
    return y


def feature(shot, title, sub, items, out, landscape=True, footer_text=None):
    c = gradient()
    d = ImageDraw.Draw(c)
    y = headline(d, 150, 130, [title], 110)
    d.text((150, y + 10), sub, font=font(500, 54), fill=INK)
    if landscape:
        box_w = 1850
        s = fit(shot, box_w, int(box_w * shot.height / shot.width))
        s = s.crop((0, 0, s.width, min(s.height, H - 620)))
        shadow(c, (150, 470, 150 + s.width, 470 + s.height), radius=24, blur=34, offset=(0, 22), opacity=80)
        c.paste(s, (150, 470))
        bx = 150 + s.width + 110
    else:
        paper(c, shot, 230, 430, 1150, angle=0)
        bx = 1600
    bullets(d, bx, 520, items)
    footer(d, footer_text) if footer_text else footer(d)
    c.save(out, quality=92)


def grid(cards, title, sub, out, cols=3, card=(860, 700), gap=(70, 110), top=470, caption=True):
    """cards: [(image, heading, caption)]."""
    c = gradient()
    d = ImageDraw.Draw(c)
    headline(d, 150, 130, [title], 120)
    d.text((150, 300), sub, font=font(500, 56), fill=INK)
    cw, ch = card
    for i, (img, hd, cap) in enumerate(cards):
        colx, row = i % cols, i // cols
        x, y = 150 + colx * (cw + gap[0]), top + row * (ch + gap[1])
        shadow(c, (x, y, x + cw, y + ch), radius=30, blur=30, offset=(0, 18), opacity=70)
        d.rounded_rectangle((x, y, x + cw, y + ch), radius=30, fill="white")
        c.paste(fit(img, cw - 60, ch - (190 if caption else 150)), (x + 30, y + 30))
        d.text((x + 34, y + ch - (145 if caption else 100)), hd, font=font(700, 50), fill=NAVY)
        if caption:
            d.text((x + 34, y + ch - 80), cap, font=font(500, 38), fill=INK)
    c.save(out, quality=92)


def how_it_works(steps, works_in, note, out):
    c = gradient()
    d = ImageDraw.Draw(c)
    headline(d, 150, 150, ["How it works"], 130)
    y = 480
    for n, (t, s) in enumerate(steps, 1):
        d.ellipse((150, y, 330, y + 180), fill=BLUE)
        tw = d.textlength(str(n), font=font(800, 100))
        d.text((240 - tw / 2, y + 25), str(n), font=font(800, 100), fill="white")
        d.text((400, y + 5), t, font=font(700, 72), fill=NAVY)
        yy = y + 100
        for ln in wrap(d, s, font(500, 52), 2400):
            d.text((400, yy), ln, font=font(500, 52), fill=INK)
            yy += 66
        y = max(y + 330, yy + 80)
    d.rounded_rectangle((150, 1650, 2850, 1990), radius=40, fill="white")
    d.text((230, 1710), "Works in", font=font(700, 56), fill=NAVY)
    d.text((230, 1800), works_in, font=font(500, 50), fill=INK)
    d.text((230, 1885), note, font=font(500, 44), fill=MUTED)
    c.save(out, quality=92)
