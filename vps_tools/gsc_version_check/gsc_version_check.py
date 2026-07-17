"""
GSC Version Check - VPS-hosted replacement/primary path for the Apps
Script's _runVersionCheckSweep(), which was hitting Apps Script's daily
UrlFetchApp quota (shared account-wide with every other script on that
account) running the sweep against a 3000+ row GSC_Config sheet.

This is NOT a new sheet and does NOT touch any existing OAuth token - it
reuses a COPY of an already-authorized GSC account's refresh token (see
config.example.json) to call the Search Console API directly from here,
and a separate Google service account (additive - just needs Editor access
shared on the SAME existing GSC_Config sheet) to read/write that sheet via
the Sheets API. The Apps Script version stays in place as a fallback/manual
option; this script is meant to become the primary way version checks run.

Replicates the exact same logic as the Apps Script's two-phase check
(_prepareDomainForVersionCheck / _resultFromInspection in the Apps Script
source), so results are equivalent regardless of which one actually ran:

  1. Resolve which GSC property (if any) the connected account has for this
     domain (_getAccountSiteEntries + resolveSiteProperty equivalent).
       - A domain property (sc-domain:) covers every URL variant -> Correct,
         no live inspection needed.
       - No property found at all -> Wrong Version.
       - A URL-prefix property that doesn't match the domain's current live
         host -> needs a live inspection call.
  2. For URL-prefix properties, call the URL Inspection API and compare its
     GSC-verified canonical URL's host against the registered property's
     host. Match -> Correct. Mismatch -> Wrong Version. Inconclusive (not
     yet crawled, API error) -> left blank, retried next run.

Run as a cron job (see README.md for the crontab line) - each run processes
up to BATCH_LIMIT domains that don't yet have a Version_Check_Status, so
repeated runs sweep through the whole sheet over time without ever holding
a single execution open long enough to hit any per-run time limit.
"""
import json
import os
import sys
import time
import datetime
import urllib.request
import urllib.error
import urllib.parse

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(SCRIPT_DIR, "config.json")

SHEETS_API = "https://sheets.googleapis.com/v4/spreadsheets"
OAUTH_TOKEN_URL = "https://oauth2.googleapis.com/token"
GSC_SITES_URL = "https://www.googleapis.com/webmasters/v3/sites"
URL_INSPECTION_URL = "https://searchconsole.googleapis.com/v1/urlInspection/index:inspect"

BATCH_LIMIT = 200          # domains processed per run - keep well under any API rate limit
REQUEST_DELAY_S = 0.15     # pause between GSC API calls, same courtesy the Apps Script used


def log(msg):
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{ts}] {msg}", flush=True)


def load_config():
    if not os.path.exists(CONFIG_PATH):
        log(f"[ERROR] {CONFIG_PATH} not found - copy config.example.json to config.json "
           "and fill it in first (see README.md).")
        sys.exit(1)
    with open(CONFIG_PATH, encoding="utf-8") as f:
        return json.load(f)


# --------------------------------------------------------------------------- #
# Google auth - two SEPARATE credentials, matching two separate concerns.
# --------------------------------------------------------------------------- #
def get_gsc_access_token(cfg):
    """Refresh the copied GSC OAuth account's access token. This is a COPY of
    an existing refresh_token from an already-connected account - using it
    here does not revoke or affect its original use anywhere else; Google
    allows a refresh token to be used concurrently from multiple places."""
    body = urllib.parse.urlencode({
        "client_id": cfg["gsc_client_id"],
        "client_secret": cfg["gsc_client_secret"],
        "refresh_token": cfg["gsc_refresh_token"],
        "grant_type": "refresh_token",
    }).encode()
    req = urllib.request.Request(OAUTH_TOKEN_URL, data=body, method="POST")
    with urllib.request.urlopen(req, timeout=20) as r:
        return json.loads(r.read().decode())["access_token"]


def get_service_account_token(cfg):
    """JWT-signed service account token for Sheets API access ONLY - this
    account never touches Search Console, just reads/writes the sheet it's
    been shared on (Editor access, granted manually - see README.md)."""
    import base64
    import hashlib
    try:
        from cryptography.hazmat.primitives import hashes, serialization
        from cryptography.hazmat.primitives.asymmetric import padding
    except ImportError:
        log("[ERROR] Missing dependency: pip install cryptography")
        sys.exit(1)

    with open(cfg["service_account_key_file"], encoding="utf-8") as f:
        sa = json.load(f)

    now = int(time.time())
    header = {"alg": "RS256", "typ": "JWT"}
    claims = {
        "iss": sa["client_email"],
        "scope": "https://www.googleapis.com/auth/spreadsheets",
        "aud": OAUTH_TOKEN_URL,
        "exp": now + 3600,
        "iat": now,
    }

    def b64(d):
        return base64.urlsafe_b64encode(json.dumps(d, separators=(",", ":")).encode()).rstrip(b"=")

    signing_input = b64(header) + b"." + b64(claims)
    private_key = serialization.load_pem_private_key(sa["private_key"].encode(), password=None)
    signature = private_key.sign(signing_input, padding.PKCS1v15(), hashes.SHA256())
    jwt = signing_input + b"." + base64.urlsafe_b64encode(signature).rstrip(b"=")

    body = urllib.parse.urlencode({
        "grant_type": "urn:ietf:params:oauth:grant-type:jwt-bearer",
        "assertion": jwt.decode(),
    }).encode()
    req = urllib.request.Request(OAUTH_TOKEN_URL, data=body, method="POST")
    with urllib.request.urlopen(req, timeout=20) as r:
        return json.loads(r.read().decode())["access_token"]


# --------------------------------------------------------------------------- #
# Sheets API - read/write GSC_Config, same shape the Apps Script uses.
# --------------------------------------------------------------------------- #
def sheets_get_values(token, sheet_id, range_a1):
    url = f"{SHEETS_API}/{sheet_id}/values/{urllib.parse.quote(range_a1)}"
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {token}"})
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.loads(r.read().decode()).get("values", [])


def sheets_batch_update_values(token, sheet_id, data):
    """data: list of {"range": "Sheet!A1", "values": [[...]]}. ONE API call
    for the whole batch - the exact lesson learned fixing the Apps Script's
    own timeout bug (N individual writes vs one bulk write)."""
    url = f"{SHEETS_API}/{sheet_id}/values:batchUpdate"
    body = json.dumps({"valueInputOption": "RAW", "data": data}).encode()
    req = urllib.request.Request(url, data=body, method="POST",
                                 headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.loads(r.read().decode())


# --------------------------------------------------------------------------- #
# GSC property resolution + inspection - mirrors resolveSiteProperty /
# inspectUrl / _resultFromInspection from the Apps Script exactly.
# --------------------------------------------------------------------------- #
def bare_host(url):
    import re
    url = (url or "").lower()
    url = re.sub(r"^sc-domain:", "", url)
    url = re.sub(r"^https?://", "", url)
    url = re.sub(r"^www\.", "", url)
    return url.split("/")[0]


_site_entries_cache = {}


def get_account_site_entries(token):
    if token in _site_entries_cache:
        return _site_entries_cache[token]
    req = urllib.request.Request(GSC_SITES_URL, headers={"Authorization": f"Bearer {token}"})
    try:
        with urllib.request.urlopen(req, timeout=20) as r:
            data = json.loads(r.read().decode())
    except urllib.error.HTTPError:
        return None
    entries = [e for e in (data.get("siteEntry") or []) if e.get("permissionLevel") != "siteUnverifiedUser"]
    _site_entries_cache[token] = entries
    return entries


def resolve_site_property(domain, token):
    entries = get_account_site_entries(token)
    if entries is None:
        return {"error": "Could not list GSC properties"}
    target = bare_host(domain)
    for e in entries:
        if e["siteUrl"].startswith("sc-domain:") and bare_host(e["siteUrl"]) == target:
            return {"siteUrl": e["siteUrl"], "type": "domain"}
    host_prefixes = [e["siteUrl"] for e in entries
                     if not e["siteUrl"].startswith("sc-domain:") and bare_host(e["siteUrl"]) == target]
    if host_prefixes:
        return {"mismatch": True, "available": host_prefixes}
    return {"notFound": True}


def inspect_url(inspection_url, site_url, token):
    body = json.dumps({"inspectionUrl": inspection_url, "siteUrl": site_url}).encode()
    req = urllib.request.Request(URL_INSPECTION_URL, data=body, method="POST",
                                 headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=20) as r:
            return r.status, json.loads(r.read().decode())
    except urllib.error.HTTPError as e:
        try:
            return e.code, json.loads(e.read().decode())
        except Exception:
            return e.code, {}


def result_from_inspection(property_url, code, body):
    if code != 200:
        return None
    idx = (body.get("inspectionResult") or {}).get("indexStatusResult") or {}
    canonical = idx.get("googleCanonical") or idx.get("userCanonical")
    if not canonical:
        return None
    if bare_host(canonical) == bare_host(property_url):
        return {"status": "Correct", "detail": f"GSC-verified canonical matches registered property: {canonical}"}
    return {"status": "Wrong Version",
            "detail": f"GSC-verified canonical is {canonical} but GSC access is only for: {property_url}"}


def check_one_domain(domain, token):
    prop = resolve_site_property(domain, token)
    if prop.get("error"):
        return {"status": "Error", "detail": prop["error"]}
    if prop.get("type") == "domain":
        return {"status": "Correct", "detail": f"Domain property (sc-domain:) - covers all URL variants of {domain}."}
    if prop.get("notFound"):
        return {"status": "Wrong Version", "detail": f"No GSC property found for {domain} on this account at all."}
    if not prop.get("mismatch") or not prop.get("available"):
        return None  # inconclusive - retried next run
    property_url = prop["available"][0]
    time.sleep(REQUEST_DELAY_S)
    code, body = inspect_url(property_url, property_url, token)
    return result_from_inspection(property_url, code, body)


# --------------------------------------------------------------------------- #
def main():
    cfg = load_config()
    log("Getting GSC access token...")
    gsc_token = get_gsc_access_token(cfg)
    log("Getting service account token for Sheets access...")
    sa_token = get_service_account_token(cfg)

    sheet_id = cfg["sheet_id"]
    tab = cfg.get("gsc_config_tab", "GSC_Config")

    log(f"Reading {tab}...")
    rows = sheets_get_values(sa_token, sheet_id, tab)
    if not rows:
        log("[ERROR] Sheet is empty or tab not found.")
        sys.exit(1)
    headers = rows[0]

    def col(name):
        return headers.index(name) if name in headers else -1

    domain_col = col("Domain")
    status_col = col("Version_Check_Status")
    if domain_col == -1:
        log("[ERROR] 'Domain' column not found in the sheet.")
        sys.exit(1)
    if status_col == -1:
        # Same on-demand column creation the Apps Script does - appends to
        # the header row, doesn't touch any existing column.
        status_col = len(headers)
        headers = headers + ["Version_Check_Status", "Version_Check_At", "Version_Check_Detail"]
        at_col, detail_col = status_col + 1, status_col + 2
        sheets_batch_update_values(sa_token, sheet_id, [
            {"range": f"{tab}!A1", "values": [headers]}])
        log("Added Version_Check_Status/At/Detail columns.")
    else:
        at_col = col("Version_Check_At")
        detail_col = col("Version_Check_Detail")

    pending = []
    for i, row in enumerate(rows[1:], start=2):
        domain = row[domain_col].strip() if domain_col < len(row) and row[domain_col] else ""
        status = row[status_col].strip() if status_col < len(row) and row[status_col] else ""
        if domain and not status:
            pending.append((i, domain))
    log(f"{len(pending)} domain(s) need a version check.")
    if not pending:
        return

    batch = pending[:BATCH_LIMIT]
    now = datetime.datetime.now().strftime("%d %b %Y %H:%M")
    updates = []
    checked, correct, wrong, inconclusive = 0, 0, 0, 0
    for row_num, domain in batch:
        result = check_one_domain(domain, gsc_token)
        if result is None:
            inconclusive += 1
            continue
        checked += 1
        if result["status"] == "Correct":
            correct += 1
        elif result["status"] == "Wrong Version":
            wrong += 1
        col_letter_status = _col_letter(status_col + 1)
        col_letter_detail = _col_letter(detail_col + 1)
        updates.append({"range": f"{tab}!{col_letter_status}{row_num}", "values": [[result["status"]]]})
        updates.append({"range": f"{tab}!{_col_letter(at_col + 1)}{row_num}", "values": [[now]]})
        updates.append({"range": f"{tab}!{col_letter_detail}{row_num}", "values": [[result["detail"]]]})

    if updates:
        log(f"Writing {len(updates)} cell update(s)...")
        sheets_batch_update_values(sa_token, sheet_id, updates)
    log(f"Done. {checked} checked ({correct} correct, {wrong} wrong version), "
       f"{inconclusive} inconclusive (retried next run), {len(pending) - len(batch)} left for later runs.")


def _col_letter(n):
    """1 -> A, 26 -> Z, 27 -> AA, ..."""
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


if __name__ == "__main__":
    main()
