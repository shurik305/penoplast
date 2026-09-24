#!/usr/bin/env python3
"""Run this ON YOUR OWN COMPUTER (not in the cloud session) to connect the agent to your Etsy shop.

It opens Etsy's consent page, catches the redirect on http://localhost:3003, exchanges the code and prints a
refresh token. Save that token as the environment variable ETSY_REFRESH_TOKEN in the Claude cloud environment
settings. Never paste it into the chat.

Usage:  python3 etsy_oauth_local.py YOUR_ETSY_KEYSTRING
Needs only the Python standard library (Python 3.8+).
"""
import base64
import hashlib
import http.server
import json
import secrets
import sys
import urllib.parse
import urllib.request
import webbrowser

REDIRECT = "http://localhost:3003/oauth/redirect"
SCOPES = "listings_r listings_w shops_r transactions_r"


def main(keystring):
    verifier = secrets.token_urlsafe(64)[:96]
    challenge = base64.urlsafe_b64encode(hashlib.sha256(verifier.encode()).digest()).rstrip(b"=").decode()
    state = secrets.token_urlsafe(16)
    url = "https://www.etsy.com/oauth/connect?" + urllib.parse.urlencode({
        "response_type": "code", "redirect_uri": REDIRECT, "scope": SCOPES, "client_id": keystring,
        "state": state, "code_challenge": challenge, "code_challenge_method": "S256"})
    result = {}

    class Handler(http.server.BaseHTTPRequestHandler):
        def do_GET(self):
            qs = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
            result.update({k: v[0] for k, v in qs.items()})
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write("<h2>Done — you can close this tab and return to the terminal.</h2>".encode())

        def log_message(self, *a):
            pass

    print("Opening Etsy in your browser. If it doesn't open, visit:\n" + url)
    webbrowser.open(url)
    with http.server.HTTPServer(("localhost", 3003), Handler) as srv:
        while "code" not in result and "error" not in result:
            srv.handle_request()
    if result.get("state") != state or "code" not in result:
        sys.exit(f"Authorisation failed: {result}")
    body = urllib.parse.urlencode({"grant_type": "authorization_code", "client_id": keystring,
                                   "redirect_uri": REDIRECT, "code": result["code"],
                                   "code_verifier": verifier}).encode()
    with urllib.request.urlopen(urllib.request.Request("https://api.etsy.com/v3/public/oauth/token", data=body),
                                timeout=30) as resp:
        tok = json.load(resp)
    print("\nSuccess. Add this as the environment variable ETSY_REFRESH_TOKEN in the Claude environment settings")
    print("(do not share it in chat):\n")
    print(tok["refresh_token"])


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    main(sys.argv[1])
