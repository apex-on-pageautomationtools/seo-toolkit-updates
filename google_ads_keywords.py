"""
Google Ads API keyword search-volume lookups (real Keyword Planner data) -
avg monthly searches, competition, CPC range, and 12-month history, filtered
by location/language. Uses the Google Ads REST interface directly via
urllib (no google-ads SDK - that pulls in grpcio/protobuf, too heavy for the
embedded Python bundle), same style as every other Google API integration in
this app (Gemini, GSC OAuth).

Needs 5 credentials, synced from the Keys sheet into CONFIG:
  google_ads_client_id, google_ads_client_secret, google_ads_refresh_token,
  google_ads_developer_token, google_ads_customer_id
Plus, since a Developer Token can only be issued from a Manager (MCC)
account, almost always also:
  google_ads_manager_customer_id
"""
import contextlib
import json
import socket
import urllib.error
import urllib.parse
import urllib.request

API_VERSION = "v23"  # v17 was long since sunset (confirmed live: 404 "was not
                     # found on this server") - v23 is Google's current latest
                     # major version as of Feb 2026, supported into 2027; bump
                     # again once Google announces its own sunset.
OAUTH_TOKEN_URL = "https://oauth2.googleapis.com/token"
BASE_URL = f"https://googleads.googleapis.com/{API_VERSION}"


@contextlib.contextmanager
def _force_ipv4():
    """Temporarily forces IPv4-only DNS resolution for the request(s) made
    inside this block. Confirmed real root cause via PowerShell
    Test-NetConnection on a live machine: googleads.googleapis.com's IPv6
    addresses are unreachable on that network ("TCP connect ... failed" for
    every IPv6 address it resolves to), while oauth2.googleapis.com's IPv6
    address works fine - a browser silently races IPv4/IPv6 and uses
    whichever works ("Happy Eyeballs"), but Python's urllib tries each
    resolved address in order with no such racing, so it was burning the
    ENTIRE request timeout on broken IPv6 addresses before ever reaching a
    working IPv4 one - the exact cause of every Keyword Search Volume
    request timing out at 65s with no result. Scoped as a context manager
    (not a permanent global patch) so it only affects these specific calls,
    not the rest of the app's networking."""
    orig_getaddrinfo = socket.getaddrinfo

    def _ipv4_only(host, port, family=0, type=0, proto=0, flags=0):
        return orig_getaddrinfo(host, port, socket.AF_INET, type, proto, flags)

    socket.getaddrinfo = _ipv4_only
    try:
        yield
    finally:
        socket.getaddrinfo = orig_getaddrinfo

# GenerateKeywordHistoricalMetrics' documented per-request keyword limit is
# lower than the 50-keyword tool-wide limit this app exposes - chunk
# internally so the caller never has to think about it.
MAX_KEYWORDS_PER_CALL = 10

# Every language Google Ads' own UI offers as a targetable Keyword Planner
# language, not just the original 10 (confirmed against Google's own
# documented criterion IDs - never guess/hardcode an ID otherwise, since a
# wrong one silently targets the wrong language with no error).
LANGUAGE_CONSTANTS = {
    "English": "languageConstants/1000",
    "German": "languageConstants/1001",
    "French": "languageConstants/1002",
    "Spanish": "languageConstants/1003",
    "Italian": "languageConstants/1004",
    "Japanese": "languageConstants/1005",
    "Danish": "languageConstants/1009",
    "Dutch": "languageConstants/1010",
    "Finnish": "languageConstants/1011",
    "Korean": "languageConstants/1012",
    "Norwegian": "languageConstants/1013",
    "Portuguese": "languageConstants/1014",
    "Swedish": "languageConstants/1015",
    "Chinese (simplified)": "languageConstants/1017",
    "Chinese (traditional)": "languageConstants/1018",
    "Arabic": "languageConstants/1019",
    "Bulgarian": "languageConstants/1020",
    "Czech": "languageConstants/1021",
    "Greek": "languageConstants/1022",
    "Hindi": "languageConstants/1023",
    "Hungarian": "languageConstants/1024",
    "Indonesian": "languageConstants/1025",
    "Icelandic": "languageConstants/1026",
    "Hebrew": "languageConstants/1027",
    "Latvian": "languageConstants/1028",
    "Lithuanian": "languageConstants/1029",
    "Polish": "languageConstants/1030",
    "Russian": "languageConstants/1031",
    "Romanian": "languageConstants/1032",
    "Slovak": "languageConstants/1033",
    "Slovenian": "languageConstants/1034",
    "Serbian": "languageConstants/1035",
    "Ukrainian": "languageConstants/1036",
    "Turkish": "languageConstants/1037",
    "Catalan": "languageConstants/1038",
    "Croatian": "languageConstants/1039",
    "Vietnamese": "languageConstants/1040",
    "Urdu": "languageConstants/1041",
    "Filipino": "languageConstants/1042",
    "Estonian": "languageConstants/1043",
    "Thai": "languageConstants/1044",
    "Bengali": "languageConstants/1056",
    "Persian": "languageConstants/1064",
    "Gujarati": "languageConstants/1072",
    "Kannada": "languageConstants/1086",
    "Malayalam": "languageConstants/1098",
    "Marathi": "languageConstants/1101",
    "Malay": "languageConstants/1102",
    "Punjabi": "languageConstants/1110",
    "Tamil": "languageConstants/1130",
    "Telugu": "languageConstants/1131",
}


class GoogleAdsConfigError(Exception):
    """Raised when a required google_ads_* key is missing from CONFIG."""


def build_config(cfg_get):
    """cfg_get: a function like CONFIG.get. Returns the config dict every
    function below needs, or raises GoogleAdsConfigError naming exactly
    which key is missing."""
    required = ["google_ads_client_id", "google_ads_client_secret", "google_ads_refresh_token",
                "google_ads_developer_token", "google_ads_customer_id"]
    missing = [k for k in required if not (cfg_get(k, "") or "").strip()]
    if missing:
        raise GoogleAdsConfigError(
            "Missing Google Ads credential(s) in the Keys sheet: " + ", ".join(missing))
    return {
        "client_id": cfg_get("google_ads_client_id", "").strip(),
        "client_secret": cfg_get("google_ads_client_secret", "").strip(),
        "refresh_token": cfg_get("google_ads_refresh_token", "").strip(),
        "developer_token": cfg_get("google_ads_developer_token", "").strip(),
        "customer_id": cfg_get("google_ads_customer_id", "").strip().replace("-", ""),
        "manager_customer_id": (cfg_get("google_ads_manager_customer_id", "") or "").strip().replace("-", "") or None,
    }


def _get_access_token(client_id, client_secret, refresh_token):
    data = urllib.parse.urlencode({
        "client_id": client_id,
        "client_secret": client_secret,
        "refresh_token": refresh_token,
        "grant_type": "refresh_token",
    }).encode()
    req = urllib.request.Request(OAUTH_TOKEN_URL, data=data,
                                  headers={"Content-Type": "application/x-www-form-urlencoded"})
    with _force_ipv4(), urllib.request.urlopen(req, timeout=15) as r:
        tokens = json.loads(r.read().decode())
    if "error" in tokens:
        raise Exception(f"Google Ads OAuth refresh failed: {tokens.get('error_description', tokens['error'])}")
    return tokens["access_token"]


def _ads_request(path, body, access_token, config):
    url = f"{BASE_URL}/{path}"
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {access_token}",
        "developer-token": config["developer_token"],
    }
    if config.get("manager_customer_id"):
        headers["login-customer-id"] = config["manager_customer_id"]
    req = urllib.request.Request(url, data=json.dumps(body).encode("utf-8"), headers=headers, method="POST")
    try:
        with _force_ipv4(), urllib.request.urlopen(req, timeout=40) as r:
            return json.loads(r.read().decode())
    except urllib.error.HTTPError as e:
        err_body = e.read().decode("utf-8", "ignore")
        try:
            msg = json.loads(err_body).get("error", {}).get("message", err_body)
        except Exception:
            msg = err_body
        raise Exception(f"Google Ads API error ({e.code}): {msg}")


def suggest_geo_target(query_text, config):
    """Look up the real geoTargetConstant ID(s) for a free-text location name
    (e.g. "United States", "Sydney, Australia") via Google's own suggest API -
    never guess/hardcode an ID, since a wrong one would silently return data
    for the wrong place with no obvious error."""
    access_token = _get_access_token(config["client_id"], config["client_secret"], config["refresh_token"])
    body = {"locationNames": {"names": [query_text]}}
    result = _ads_request("geoTargetConstants:suggest", body, access_token, config)
    suggestions = result.get("geoTargetConstantSuggestions", [])
    return [{
        "id": s["geoTargetConstant"]["id"],
        "name": s["geoTargetConstant"].get("name", ""),
        "country_code": s["geoTargetConstant"].get("countryCode", ""),
        "target_type": s["geoTargetConstant"].get("targetType", ""),
        "resource_name": s["geoTargetConstant"]["resourceName"],
    } for s in suggestions]


def get_keyword_historical_metrics(keywords, geo_target_resource_names, language_resource_name, config):
    """keywords: list of raw keyword strings (tool-wide limit of 50 is
    enforced by the caller; chunked into MAX_KEYWORDS_PER_CALL-sized API
    calls internally). Returns a list of dicts: keyword, avg_monthly_searches,
    competition, competition_index, low_cpc, high_cpc, monthly_searches."""
    access_token = _get_access_token(config["client_id"], config["client_secret"], config["refresh_token"])
    all_results = []
    for i in range(0, len(keywords), MAX_KEYWORDS_PER_CALL):
        chunk = keywords[i:i + MAX_KEYWORDS_PER_CALL]
        body = {
            "keywords": chunk,
            "keywordPlanNetwork": "GOOGLE_SEARCH",
            "geoTargetConstants": geo_target_resource_names,
            "language": language_resource_name,
        }
        result = _ads_request(f"customers/{config['customer_id']}:generateKeywordHistoricalMetrics",
                               body, access_token, config)

        def _micros_to_currency(v):
            return (int(v) / 1_000_000) if v else None

        for item in result.get("results", []):
            m = item.get("keywordMetrics", {}) or {}
            all_results.append({
                "keyword": item.get("text", ""),
                "avg_monthly_searches": int(m.get("avgMonthlySearches", 0) or 0),
                "competition": m.get("competition", "UNSPECIFIED"),
                "competition_index": m.get("competitionIndex"),
                "low_cpc": _micros_to_currency(m.get("lowTopOfPageBidMicros")),
                "high_cpc": _micros_to_currency(m.get("highTopOfPageBidMicros")),
                "monthly_searches": [
                    {"year": mv.get("year"), "month": mv.get("month"),
                     "searches": int(mv.get("monthlySearches", 0) or 0)}
                    for mv in (m.get("monthlySearchVolumes") or [])
                ],
            })
    return all_results
