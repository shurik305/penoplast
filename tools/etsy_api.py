#!/usr/bin/env python3
"""Minimal Etsy Open API v3 client for the owner's own shop (Seller App access).

Secrets never live in the repository. Configure via environment variables (cloud environment secrets):
  ETSY_KEYSTRING, ETSY_SHARED_SECRET, ETSY_REFRESH_TOKEN
Tokens obtained during a session are cached in ~/.config/penoplast/etsy_tokens.json (outside the repo).

The refresh token is created by the owner locally with tools/etsy_oauth_local.py (never pasted into chat).

Commands:
  me                            show user id and shop id
  taxonomy <search>             find seller taxonomy node ids
  listings                      active listings with views and favourites (metrics)
  receipts                      paid orders (independent proof of sales)
  publish <listing.json> [--draft] [--made-by TEXT]   create listing, upload images and files, activate
Network: requires api.etsy.com, openapi.etsy.com and www.etsy.com to be allowed.
"""
import json
import os
import sys
import time

import requests

API = "https://openapi.etsy.com/v3/application"
TOKEN_URL = "https://api.etsy.com/v3/public/oauth/token"
REDIRECT = os.environ.get("ETSY_REDIRECT_URI", "http://localhost:3003/oauth/redirect")
SCOPES = "listings_r listings_w shops_r transactions_r"
CACHE = os.path.expanduser("~/.config/penoplast/etsy_tokens.json")


def _key():
    ks = os.environ["ETSY_KEYSTRING"]
    ss = os.environ.get("ETSY_SHARED_SECRET", "")
    # Etsy accepts "keystring:shared_secret" in x-api-key; fall back to the bare keystring if no secret is set.
    return f"{ks}:{ss}" if ss else ks


def _save(path, data):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w") as fh:
        json.dump(data, fh)
    os.chmod(path, 0o600)


def _load(path):
    try:
        with open(path) as fh:
            return json.load(fh)
    except FileNotFoundError:
        return {}


def _store_tokens(tok):
    tok["expires_at"] = time.time() + int(tok.get("expires_in", 3600)) - 60
    _save(CACHE, tok)
    return tok


def _token():
    tok = _load(CACHE)
    if tok.get("access_token") and tok.get("expires_at", 0) > time.time():
        return tok["access_token"]
    refresh = tok.get("refresh_token") or os.environ.get("ETSY_REFRESH_TOKEN")
    if not refresh:
        raise SystemExit("no token — set ETSY_REFRESH_TOKEN (see tools/etsy_oauth_local.py)")
    r = requests.post(TOKEN_URL, data={"grant_type": "refresh_token", "client_id": os.environ["ETSY_KEYSTRING"],
                                       "refresh_token": refresh}, timeout=30)
    r.raise_for_status()
    return _store_tokens(r.json())["access_token"]


def call(method, path, **kw):
    headers = {"x-api-key": _key(), "Authorization": f"Bearer {_token()}"}
    r = requests.request(method, API + path, headers=headers, timeout=120, **kw)
    if r.status_code >= 400:
        raise SystemExit(f"{method} {path} -> {r.status_code}: {r.text[:500]}")
    return r.json() if r.text else {}


def me():
    u = call("GET", "/users/me")
    return u["user_id"], u["shop_id"]


def shop_currency(shop_id):
    return call("GET", f"/shops/{shop_id}").get("currency_code", "USD")


def taxonomy(search):
    nodes, out = call("GET", "/seller-taxonomy/nodes")["results"], []

    def walk(ns, trail):
        for n in ns:
            t = trail + [n["name"]]
            if search.lower() in n["name"].lower():
                out.append((n["id"], " > ".join(t)))
            walk(n.get("children", []), t)
    walk(nodes, [])
    return out


def publish(listing_path, draft=False, made_by=None):
    base = os.path.dirname(os.path.abspath(listing_path))
    spec = json.load(open(listing_path))
    _, shop_id = me()
    cur = shop_currency(shop_id)
    price = spec["price_dkk"] if cur == "DKK" else spec["price_usd"] if cur == "USD" else None
    if price is None:
        raise SystemExit(f"shop currency {cur}: add price_{cur.lower()} to the listing spec")
    desc = open(os.path.join(base, spec["description_file"])).read()
    desc = desc.replace("{{MADE_BY}}", made_by or "Designed by Hutsol with the help of AI tools.")
    tax = spec.get("taxonomy_id") or taxonomy(spec["taxonomy_search"])[0][0]
    data = {"quantity": spec["quantity"], "title": spec["title"], "description": desc, "price": price,
            "who_made": spec["who_made"], "when_made": spec["when_made"], "taxonomy_id": tax,
            "type": spec["type"], "is_supply": str(spec["is_supply"]).lower(),
            "should_auto_renew": str(spec["should_auto_renew"]).lower(),
            "tags": ",".join(spec["tags"]), "materials": ",".join(spec["materials"])}
    lst = call("POST", f"/shops/{shop_id}/listings", data=data)
    lid = lst["listing_id"]
    for i, img in enumerate(spec["images"], 1):
        with open(os.path.join(base, img), "rb") as fh:
            call("POST", f"/shops/{shop_id}/listings/{lid}/images", files={"image": fh}, data={"rank": i})
    for i, fp in enumerate(spec["files"], 1):
        with open(os.path.join(base, fp), "rb") as fh:
            call("POST", f"/shops/{shop_id}/listings/{lid}/files", files={"file": fh},
                 data={"name": os.path.basename(fp), "rank": i})
    if not draft:
        call("PATCH", f"/shops/{shop_id}/listings/{lid}", data={"state": "active"})
    return lid


def listings():
    _, shop_id = me()
    res = call("GET", f"/shops/{shop_id}/listings", params={"state": "active", "limit": 100})["results"]
    return [(x["listing_id"], x.get("views"), x.get("num_favorers"), x["title"][:60]) for x in res]


def receipts():
    _, shop_id = me()
    res = call("GET", f"/shops/{shop_id}/receipts", params={"was_paid": "true", "limit": 100})["results"]
    return [(x["receipt_id"], x.get("create_timestamp"), x["grandtotal"]["amount"] / x["grandtotal"]["divisor"],
             x["grandtotal"]["currency_code"]) for x in res]


if __name__ == "__main__":
    cmd, args = (sys.argv[1] if len(sys.argv) > 1 else "help"), sys.argv[2:]
    if cmd == "me":
        print(me())
    elif cmd == "taxonomy":
        for row in taxonomy(args[0]):
            print(*row)
    elif cmd == "listings":
        for row in listings():
            print(*row)
    elif cmd == "receipts":
        for row in receipts():
            print(*row)
    elif cmd == "publish":
        mb = args[args.index("--made-by") + 1] if "--made-by" in args else None
        print("listing", publish(args[0], draft="--draft" in args, made_by=mb))
    else:
        print(__doc__)
