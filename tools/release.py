#!/usr/bin/env python3
"""One-command release: rebuild every product, fill cached values, run QA + edge cases, render images and guides,
then build the manual Etsy upload kit (with freshness, banned-string checks and a manifest).

Usage: source ops/setup.sh && python3 tools/release.py
Stops at the first failure. Nothing here publishes anything.
"""
import glob
import os
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "products", "_lib"))
import cachefill  # noqa: E402

PRODUCTS = [
    ("products/pricing-calculator", ["python3 build.py dist"], "listing/raw/quote-calculator.png"),
    ("products/commercial-bid-calculator", ["python3 build.py dist"], "listing/raw/bid-calculator-1.png"),
    ("products/shortlet-turnover-kit", ["python3 build.py dist", "python3 printables.py dist"],
     "listing/raw/cleaning-fee-calculator-1.png"),
]


def run(cmd, cwd):
    print(f"  $ {cmd}")
    res = subprocess.run(cmd, shell=True, cwd=cwd, capture_output=True, text=True)
    tail = (res.stdout + res.stderr).strip().splitlines()[-3:]
    for ln in tail:
        print(f"    {ln[:160]}")
    if res.returncode != 0:
        raise SystemExit(f"FAILED: {cmd} (in {cwd})")


def main():
    for rel, builds, shot in PRODUCTS:
        cwd = os.path.join(ROOT, rel)
        print(f"== {rel}")
        for b in builds:
            run(b, cwd)
        for x in sorted(glob.glob(os.path.join(cwd, "dist", "*.xlsx"))):
            print(f"  cached values: {os.path.basename(x)} ({cachefill.fill(x)} formula cells)")
        run("python3 qa.py dist", cwd)
        run("python3 edge_cases.py dist", cwd)
        run("python3 mockups.py dist listing/images", cwd)
        run(f"python3 guide.py {shot} dist/Quick-Start-Guide.pdf", cwd)
    print("== upload kit")
    run("python3 tools/make_upload_kit.py", ROOT)
    print("RELEASE OK")


if __name__ == "__main__":
    main()
