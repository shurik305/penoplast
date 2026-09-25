#!/usr/bin/env python3
"""Builds a manual Etsy upload kit for each product: numbered photos, digital files and a copy-paste sheet.

Usage: python3 tools/make_upload_kit.py [--made-by a|b] [product-dir ...]
Output: _upload_kit/<product>/ and _upload_kit/etsy-upload-kit.zip (git-ignored: contains the paid files).
Run each product's build/mockups/guide first so dist/ and listing/images/ exist.
"""
import datetime as dt
import glob
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "_upload_kit")
MADE_BY = {
    "a": "Designed by Hutsol — a small cleaning and property-service company in West Jutland, Denmark — with the "
         "help of AI tools.",
    "b": "Designed by Hutsol with the help of AI tools.",
}
# Brand names that must never appear in titles or tags (trademark bots); stale or unverified claims
TRADEMARKS = ["airbnb", "vrbo", "booking.com", "canva", "excel®"]
STALE_CLAIMS = ["All formulas were tested", "rename freely", "1–2 working days", "2010 or newer", "All formulas work"]
DEFAULT_PRODUCTS = ["products/pricing-calculator", "products/commercial-bid-calculator",
                    "products/shortlet-turnover-kit"]


def sheet(spec, desc):
    tags = "\n".join(f"  {i}. {t}" for i, t in enumerate(spec["tags"], 1))
    return f"""ETSY LISTING — COPY AND PASTE
============================

1. PHOTOS: upload the files in 01-photos/ in this order (01 is the main photo).

2. TITLE:
{spec["title"]}

3. ABOUT THIS LISTING:
  Who made it?            I did   (if the form asks "designed by": your shop)
  What is it?             A finished product
  When was it made?       2020 - 2026
  Category:               search "templates" and choose the closest template / business form category
  Type:                   Digital files
  AI question/checkbox:   answer truthfully — yes, designed with the help of AI tools.
  Please screenshot the "About this listing" part of the form and send it to the agent (Etsy changes it often).

4. DESCRIPTION: paste everything between the lines.
----------------------------------------------------------------
{desc.strip()}
----------------------------------------------------------------

5. PRICE: {spec["price_dkk"]} DKK (if your shop currency is USD: {spec["price_usd"]} USD)
   QUANTITY: {spec["quantity"]}

6. DIGITAL FILES: upload all files in 02-digital-files/ (max 5 files per listing).

7. TAGS (13, copy one by one):
{tags}

8. MATERIALS: {", ".join(spec["materials"])}

9. RENEWAL: automatic (each renewal costs USD 0.20 — within the approved budget only).

10. Publish. Then write the listing link and date in the chat so the agent can log it.
"""


def sha(path):
    return hashlib.sha256(open(path, "rb").read()).hexdigest()[:16]


def checks(pd, spec, desc):
    """Refuses stale or risky kits: trademarks in title/tags, stale claims, outputs older than their sources."""
    problems = []
    title_tags = " ".join([spec["title"]] + spec["tags"]).lower()
    problems += [f"trademark '{t}' in title/tags" for t in TRADEMARKS if re.search(rf"\b{re.escape(t)}\b", title_tags)]
    problems += [f"stale claim '{c}' in description" for c in STALE_CLAIMS if c in desc]
    sources = [os.path.join(pd, f) for f in ("build.py", "mockups.py", "guide.py", "printables.py") if
               os.path.exists(os.path.join(pd, f))] + glob.glob(os.path.join(ROOT, "products", "_lib", "*.py"))
    newest_src = max(os.path.getmtime(x) for x in sources)
    outputs = [os.path.join(pd, "listing", i) for i in spec["images"]] + \
              [os.path.normpath(os.path.join(pd, "listing", f)) for f in spec["files"]]
    stale = [os.path.relpath(o, ROOT) for o in outputs if os.path.getmtime(o) < newest_src]
    if stale:
        problems.append(f"outputs older than their sources (run tools/release.py): {stale[:4]}")
    mock = open(os.path.join(pd, "mockups.py"), encoding="utf-8").read().lower()
    problems += [f"trademark '{t}' in image text" for t in TRADEMARKS if re.search(rf"\b{re.escape(t)}\b", mock)]
    return problems


def build(product_dirs, made_by):
    os.makedirs(OUT, exist_ok=True)
    made = []
    manifest = [f"Etsy upload kit — built {dt.datetime.now(dt.timezone.utc):%Y-%m-%d %H:%M} UTC",
                "git commit: " + subprocess.run(["git", "-C", ROOT, "rev-parse", "--short", "HEAD"],
                                                capture_output=True, text=True).stdout.strip(), ""]
    for pd in product_dirs:
        base = os.path.join(ROOT, pd)
        ldir = os.path.join(base, "listing")
        spec = json.load(open(os.path.join(ldir, "listing.json"), encoding="utf-8"))
        desc = open(os.path.join(ldir, spec["description_file"]), encoding="utf-8").read()
        desc = desc.replace("{{MADE_BY}}", MADE_BY[made_by])
        problems = checks(base, spec, desc)
        if problems:
            raise SystemExit(f"{pd}: refusing to build the kit:\n  - " + "\n  - ".join(problems))
        name = os.path.basename(pd)
        kit = os.path.join(OUT, name)
        shutil.rmtree(kit, ignore_errors=True)
        os.makedirs(os.path.join(kit, "01-photos"))
        os.makedirs(os.path.join(kit, "02-digital-files"))
        for img in spec["images"]:
            shutil.copy(os.path.join(ldir, img), os.path.join(kit, "01-photos", os.path.basename(img)))
        missing = []
        for fp in spec["files"]:
            src = os.path.normpath(os.path.join(ldir, fp))
            if os.path.exists(src):
                shutil.copy(src, os.path.join(kit, "02-digital-files", os.path.basename(src)))
            else:
                missing.append(fp)
        if missing:
            raise SystemExit(f"{name}: missing files {missing} — run the product build first")
        with open(os.path.join(kit, "LISTING.txt"), "w", encoding="utf-8") as fh:
            fh.write(sheet(spec, desc))
        made.append(kit)
        manifest.append(f"{name}: {spec['price_dkk']} DKK / {spec['price_usd']} USD — {spec['title']}")
        for sub in ("01-photos", "02-digital-files"):
            for fn in sorted(os.listdir(os.path.join(kit, sub))):
                fp = os.path.join(kit, sub, fn)
                manifest.append(f"  {sub}/{fn}  {os.path.getsize(fp) // 1024} KB  sha256:{sha(fp)}")
        manifest.append("")
    with open(os.path.join(OUT, "MANIFEST.txt"), "w", encoding="utf-8") as fh:
        fh.write("\n".join(manifest))
    zp = os.path.join(OUT, "etsy-upload-kit.zip")
    with zipfile.ZipFile(zp, "w", zipfile.ZIP_DEFLATED) as z:
        z.write(os.path.join(OUT, "MANIFEST.txt"), "MANIFEST.txt")
        for kit in made:
            for dp, _, fns in os.walk(kit):
                for fn in fns:
                    full = os.path.join(dp, fn)
                    z.write(full, os.path.relpath(full, OUT))
    return made, zp


if __name__ == "__main__":
    args = sys.argv[1:]
    mb = "b"
    if "--made-by" in args:
        i = args.index("--made-by")
        mb = args[i + 1]
        del args[i:i + 2]
    kits, zp = build(args or DEFAULT_PRODUCTS, mb)
    for k in kits:
        print(k)
    print(zp, os.path.getsize(zp) // 1024, "KB")
