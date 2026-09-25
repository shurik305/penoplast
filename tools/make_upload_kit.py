#!/usr/bin/env python3
"""Builds a manual Etsy upload kit for each product: numbered photos, digital files and a copy-paste sheet.

Usage: python3 tools/make_upload_kit.py [--made-by a|b] [product-dir ...]
Output: _upload_kit/<product>/ and _upload_kit/etsy-upload-kit.zip (git-ignored: contains the paid files).
Run each product's build/mockups/guide first so dist/ and listing/images/ exist.
"""
import json
import os
import shutil
import sys
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "_upload_kit")
MADE_BY = {
    "a": "Designed by Hutsol — a small cleaning and property-service company in West Jutland, Denmark — with the "
         "help of AI tools.",
    "b": "Designed by Hutsol with the help of AI tools.",
}
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
  Who made it?            I did
  What is it?             A finished product
  When was it made?       2020 - 2026
  Category:               search "templates" and choose the closest template / business form category
  Type:                   Digital files
  If Etsy asks whether AI was used: answer truthfully (yes — designed with the help of AI tools).

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


def build(product_dirs, made_by):
    os.makedirs(OUT, exist_ok=True)
    made = []
    for pd in product_dirs:
        base = os.path.join(ROOT, pd)
        ldir = os.path.join(base, "listing")
        spec = json.load(open(os.path.join(ldir, "listing.json"), encoding="utf-8"))
        desc = open(os.path.join(ldir, spec["description_file"]), encoding="utf-8").read()
        desc = desc.replace("{{MADE_BY}}", MADE_BY[made_by])
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
    zp = os.path.join(OUT, "etsy-upload-kit.zip")
    with zipfile.ZipFile(zp, "w", zipfile.ZIP_DEFLATED) as z:
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
