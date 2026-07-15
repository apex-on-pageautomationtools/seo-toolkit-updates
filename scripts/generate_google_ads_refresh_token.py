"""
One-time helper: generates a Google Ads API refresh token via the standard
OAuth2 "installed app" loopback flow (opens a browser for you to log in and
consent, then captures the authorization code on http://localhost - the
modern replacement for the deprecated out-of-band/copy-paste-a-code flow).

Run this ONCE, by whoever is setting up Google Ads API access for the whole
team - the resulting refresh_token gets stored centrally (Keys sheet) and
reused server-side for every user from then on. No one else ever needs to
log in.

Usage:
    python scripts/generate_google_ads_refresh_token.py --client-id ... --client-secret ...

Prints the refresh_token to paste into the Keys sheet as
google_ads_refresh_token, alongside google_ads_client_id / _client_secret /
_developer_token / _customer_id.
"""
import argparse
import http.server
import json
import threading
import urllib.parse
import urllib.request
import webbrowser

AUTH_URL = "https://accounts.google.com/o/oauth2/v2/auth"
TOKEN_URL = "https://oauth2.googleapis.com/token"
SCOPE = "https://www.googleapis.com/auth/adwords"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--client-id", required=True)
    ap.add_argument("--client-secret", required=True)
    ap.add_argument("--port", type=int, default=8901,
                     help="Local port for the loopback redirect (must be free on this machine)")
    args = ap.parse_args()

    redirect_uri = f"http://localhost:{args.port}/oauth_callback"
    result = {}

    class Handler(http.server.BaseHTTPRequestHandler):
        def do_GET(self):
            parsed = urllib.parse.urlparse(self.path)
            qs = urllib.parse.parse_qs(parsed.query)
            if "code" in qs:
                result["code"] = qs["code"][0]
                self.send_response(200)
                self.send_header("Content-Type", "text/html")
                self.end_headers()
                self.wfile.write(b"<html><body><h2>Authorized. You can close this tab and return to the terminal.</h2></body></html>")
            else:
                self.send_response(400)
                self.end_headers()
                self.wfile.write(b"No authorization code received.")

        def log_message(self, *a):
            pass  # quiet - don't spam the terminal with HTTP access logs

    server = http.server.HTTPServer(("localhost", args.port), Handler)
    t = threading.Thread(target=server.handle_request, daemon=True)
    t.start()

    auth_url = AUTH_URL + "?" + urllib.parse.urlencode({
        "client_id": args.client_id,
        "redirect_uri": redirect_uri,
        "response_type": "code",
        "scope": SCOPE,
        "access_type": "offline",   # required to get a refresh_token back
        "prompt": "consent",        # forces a refresh_token even if previously authorized
    })
    print("Opening your browser to log in and authorize...")
    print("If it doesn't open automatically, visit this URL:\n")
    print(auth_url)
    print()
    webbrowser.open(auth_url)

    t.join(timeout=300)
    if "code" not in result:
        print("[ERROR] No authorization code received within 5 minutes. Try again.")
        return

    print("Exchanging authorization code for tokens...")
    data = urllib.parse.urlencode({
        "code": result["code"],
        "client_id": args.client_id,
        "client_secret": args.client_secret,
        "redirect_uri": redirect_uri,
        "grant_type": "authorization_code",
    }).encode()
    req = urllib.request.Request(TOKEN_URL, data=data,
                                  headers={"Content-Type": "application/x-www-form-urlencoded"})
    with urllib.request.urlopen(req, timeout=15) as r:
        tokens = json.loads(r.read().decode())

    if "error" in tokens:
        print(f"[ERROR] Token exchange failed: {tokens.get('error_description', tokens['error'])}")
        return
    if "refresh_token" not in tokens:
        print("[ERROR] No refresh_token in the response - this usually means the account already "
              "granted consent before without 'prompt=consent'. This script already forces that, "
              "so if you still see this, revoke prior access at https://myaccount.google.com/permissions "
              "for this app and try again.")
        return

    print("\n=== SUCCESS ===")
    print("Add this to the Keys sheet as: google_ads_refresh_token")
    print(tokens["refresh_token"])


if __name__ == "__main__":
    main()
