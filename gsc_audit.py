"""
SEO Toolkit Pro — GSC Audit Module
Handles Google OAuth, GSC API data, screenshots, and 5 PPTX report formats.
"""

import os
import sys
import json
import time
import hashlib
import tempfile
import urllib.request
import urllib.parse
from datetime import datetime, timedelta
from io import BytesIO

BUNDLE_DIR = os.path.dirname(os.path.abspath(__file__))
GSC_AUTH_FILE = os.path.join(BUNDLE_DIR, ".gsc_accounts")

OAUTH_AUTH_URL = "https://accounts.google.com/o/oauth2/v2/auth"
OAUTH_TOKEN_URL = "https://oauth2.googleapis.com/token"
USERINFO_URL = "https://www.googleapis.com/oauth2/v3/userinfo"
SCOPES = "https://www.googleapis.com/auth/webmasters https://www.googleapis.com/auth/userinfo.email"

GSC_API_BASE = "https://www.googleapis.com/webmasters/v3"
SEARCH_ANALYTICS_URL = "https://searchconsole.googleapis.com/webmasters/v3"
URL_INSPECTION_URL = "https://searchconsole.googleapis.com/v1/urlInspection/index:inspect"

# ---------------------------------------------------------------------------
# Account storage
# ---------------------------------------------------------------------------

def _load_accounts():
    try:
        if os.path.exists(GSC_AUTH_FILE):
            with open(GSC_AUTH_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return {}


def _save_accounts(accounts):
    try:
        with open(GSC_AUTH_FILE, "w", encoding="utf-8") as f:
            json.dump(accounts, f, indent=2)
    except Exception:
        pass


def list_accounts():
    accs = _load_accounts()
    return [{"email": k, "has_refresh": bool(v.get("refresh_token"))}
            for k, v in accs.items()]


def remove_account(email):
    accs = _load_accounts()
    accs.pop(email.lower(), None)
    _save_accounts(accs)


# ---------------------------------------------------------------------------
# OAuth via Selenium popup
# ---------------------------------------------------------------------------

def get_oauth_url(client_id, redirect_uri="urn:ietf:wg:oauth:2.0:oob"):
    params = urllib.parse.urlencode({
        "client_id": client_id,
        "redirect_uri": redirect_uri,
        "response_type": "code",
        "scope": SCOPES,
        "access_type": "offline",
        "prompt": "consent",
    })
    return f"{OAUTH_AUTH_URL}?{params}"


def oauth_login_selenium(driver, client_id, client_secret, log_fn=None):
    """Open Google OAuth in the Selenium browser and wait for the user to authorize.
    Returns the email of the connected account."""
    if log_fn is None:
        log_fn = print

    redirect_uri = "http://localhost:19876/oauth_callback"
    params = urllib.parse.urlencode({
        "client_id": client_id,
        "redirect_uri": redirect_uri,
        "response_type": "code",
        "scope": SCOPES,
        "access_type": "offline",
        "prompt": "consent",
    })
    auth_url = f"{OAUTH_AUTH_URL}?{params}"

    original_window = driver.current_window_handle
    driver.execute_script("window.open(arguments[0], '_blank', 'width=500,height=700');", auth_url)

    # Wait for the OAuth redirect
    import time as _t
    new_handles = [h for h in driver.window_handles if h != original_window]
    if new_handles:
        driver.switch_to.window(new_handles[0])

    log_fn("  Waiting for Google authorization (log in and click Allow)...")

    code = None
    for _ in range(300):  # 5 minutes max
        _t.sleep(1)
        try:
            current_url = driver.current_url
            if "oauth_callback" in current_url and "code=" in current_url:
                parsed = urllib.parse.urlparse(current_url)
                qs = urllib.parse.parse_qs(parsed.query)
                code = qs.get("code", [None])[0]
                break
            if "approval_code" in current_url or "/oauthchooseaccount" not in current_url:
                # Check if the page shows an authorization code
                try:
                    page_text = driver.find_element("tag name", "body").text
                    if "4/" in page_text:
                        for line in page_text.split("\n"):
                            line = line.strip()
                            if line.startswith("4/"):
                                code = line
                                break
                except Exception:
                    pass
        except Exception:
            # Window might be closed
            break

    # Close the OAuth popup
    try:
        if driver.current_window_handle != original_window:
            driver.close()
    except Exception:
        pass
    try:
        driver.switch_to.window(original_window)
    except Exception:
        pass

    if not code:
        raise Exception("OAuth authorization failed — no code received. Please try again.")

    # Exchange code for tokens
    tokens = _exchange_code(code, client_id, client_secret, redirect_uri)
    email = _get_user_email(tokens["access_token"])

    accs = _load_accounts()
    accs[email] = {
        "email": email,
        "refresh_token": tokens.get("refresh_token", ""),
        "access_token": tokens["access_token"],
        "expires_at": time.time() + tokens.get("expires_in", 3600) - 60,
    }
    _save_accounts(accs)
    log_fn(f"  Connected: {email}")
    return email


def _exchange_code(code, client_id, client_secret, redirect_uri):
    data = urllib.parse.urlencode({
        "code": code,
        "client_id": client_id,
        "client_secret": client_secret,
        "redirect_uri": redirect_uri,
        "grant_type": "authorization_code",
    }).encode()
    req = urllib.request.Request(OAUTH_TOKEN_URL, data=data,
                                headers={"Content-Type": "application/x-www-form-urlencoded"})
    with urllib.request.urlopen(req, timeout=15) as r:
        tokens = json.loads(r.read().decode())
    if "error" in tokens:
        raise Exception(f"Token exchange failed: {tokens.get('error_description', tokens['error'])}")
    return tokens


def _get_user_email(access_token):
    req = urllib.request.Request(USERINFO_URL,
                                headers={"Authorization": f"Bearer {access_token}"})
    with urllib.request.urlopen(req, timeout=10) as r:
        info = json.loads(r.read().decode())
    return info.get("email", "").lower()


def _refresh_token(email):
    """Refresh an expired access token. Returns the new access_token."""
    accs = _load_accounts()
    acc = accs.get(email)
    if not acc or not acc.get("refresh_token"):
        raise Exception(f"No refresh token for {email}. Please reconnect the account.")

    config = _load_gsc_config()
    data = urllib.parse.urlencode({
        "client_id": config.get("client_id", ""),
        "client_secret": config.get("client_secret", ""),
        "refresh_token": acc["refresh_token"],
        "grant_type": "refresh_token",
    }).encode()
    req = urllib.request.Request(OAUTH_TOKEN_URL, data=data,
                                headers={"Content-Type": "application/x-www-form-urlencoded"})
    with urllib.request.urlopen(req, timeout=15) as r:
        tokens = json.loads(r.read().decode())
    if "error" in tokens:
        remove_account(email)
        raise Exception(f"Session expired for {email}. Please reconnect.")

    acc["access_token"] = tokens["access_token"]
    acc["expires_at"] = time.time() + tokens.get("expires_in", 3600) - 60
    accs[email] = acc
    _save_accounts(accs)
    return tokens["access_token"]


def get_access_token(email):
    """Get a valid access token, refreshing if needed."""
    accs = _load_accounts()
    acc = accs.get(email.lower())
    if not acc:
        raise Exception(f"Account {email} not connected.")
    if acc.get("access_token") and acc.get("expires_at", 0) > time.time():
        return acc["access_token"]
    return _refresh_token(email.lower())


def _load_gsc_config():
    cfg_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
    try:
        if os.path.exists(cfg_path):
            with open(cfg_path, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return {}


# ---------------------------------------------------------------------------
# GSC API helpers
# ---------------------------------------------------------------------------

def _api_get(url, token, timeout=15):
    req = urllib.request.Request(url, headers={
        "Authorization": f"Bearer {token}",
        "User-Agent": "SEOToolkitPro-GSC/1.0",
    })
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.loads(r.read().decode())


def _api_post(url, token, body, timeout=30):
    data = json.dumps(body).encode()
    req = urllib.request.Request(url, data=data, headers={
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
        "User-Agent": "SEOToolkitPro-GSC/1.0",
    })
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.loads(r.read().decode())


def list_properties(token):
    """List all GSC properties accessible by this token."""
    resp = _api_get(f"{GSC_API_BASE}/sites", token)
    return resp.get("siteEntry", [])


def resolve_property(token, domain):
    """Find the best GSC property URL for a domain."""
    props = list_properties(token)
    domain_clean = domain.lower().replace("www.", "").strip("/")
    # Prefer domain property, then URL prefix
    for p in props:
        url = p.get("siteUrl", "")
        if url == f"sc-domain:{domain_clean}":
            return url
    for p in props:
        url = p.get("siteUrl", "")
        if domain_clean in url.lower():
            return url
    if props:
        return props[0].get("siteUrl", "")
    raise Exception(f"No GSC property found for '{domain}'. Make sure the domain is added to Google Search Console.")


def fetch_sitemaps(token, property_url):
    encoded = urllib.parse.quote(property_url, safe="")
    resp = _api_get(f"{GSC_API_BASE}/sites/{encoded}/sitemaps", token, timeout=20)
    return resp.get("sitemap", [])


def fetch_search_analytics(token, property_url, start_date, end_date, dimensions=None, row_limit=25):
    encoded = urllib.parse.quote(property_url, safe="")
    body = {
        "startDate": start_date,
        "endDate": end_date,
        "dimensions": dimensions or ["query"],
        "rowLimit": row_limit,
    }
    resp = _api_post(
        f"{SEARCH_ANALYTICS_URL}/sites/{encoded}/searchAnalytics/query",
        token, body, timeout=30
    )
    return resp.get("rows", [])


def fetch_performance_daily(token, property_url, start_date, end_date):
    return fetch_search_analytics(token, property_url, start_date, end_date,
                                  dimensions=["date"], row_limit=100)


def fetch_top_queries(token, property_url, start_date, end_date, limit=25):
    return fetch_search_analytics(token, property_url, start_date, end_date,
                                  dimensions=["query"], row_limit=limit)


def fetch_top_pages(token, property_url, start_date, end_date, limit=25):
    return fetch_search_analytics(token, property_url, start_date, end_date,
                                  dimensions=["page"], row_limit=limit)


def fetch_image_perf(token, property_url, start_date, end_date, limit=10):
    encoded = urllib.parse.quote(property_url, safe="")
    body = {
        "startDate": start_date,
        "endDate": end_date,
        "dimensions": ["query"],
        "rowLimit": limit,
        "searchType": "image",
    }
    resp = _api_post(
        f"{SEARCH_ANALYTICS_URL}/sites/{encoded}/searchAnalytics/query",
        token, body, timeout=30
    )
    return resp.get("rows", [])


def inspect_url(token, property_url, url_to_inspect):
    body = {
        "inspectionUrl": url_to_inspect,
        "siteUrl": property_url,
    }
    resp = _api_post(URL_INSPECTION_URL, token, body, timeout=30)
    result = resp.get("inspectionResult", {})
    return {
        "indexStatusResult": result.get("indexStatusResult", {}),
        "richResultsResult": result.get("richResultsResult", {}),
    }


def fetch_all_api_data(token, property_url, start_date, end_date, log_fn=None):
    if log_fn is None:
        log_fn = print
    data = {}
    log_fn("  Fetching sitemaps...")
    data["sitemaps"] = fetch_sitemaps(token, property_url)
    log_fn("  Fetching performance data...")
    data["perfDaily"] = fetch_performance_daily(token, property_url, start_date, end_date)
    log_fn("  Fetching top queries...")
    data["topQueries"] = fetch_top_queries(token, property_url, start_date, end_date)
    log_fn("  Fetching top pages...")
    data["topPages"] = fetch_top_pages(token, property_url, start_date, end_date)
    log_fn("  Fetching image performance...")
    data["imagePerf"] = fetch_image_perf(token, property_url, start_date, end_date)
    return data


def run_inspections(token, property_url, top_pages, log_fn=None):
    if log_fn is None:
        log_fn = print
    urls = []
    for p in top_pages[:10]:
        keys = p.get("keys", [])
        if keys:
            urls.append(keys[0])
    if not urls and property_url.startswith("http"):
        urls = [property_url.rstrip("/") + "/"]
    elif not urls and property_url.startswith("sc-domain:"):
        urls = [f"https://{property_url.replace('sc-domain:', '')}/"]

    inspections = []
    for i, url in enumerate(urls):
        log_fn(f"  Inspecting URL {i+1}/{len(urls)}: {url[:60]}...")
        try:
            result = inspect_url(token, property_url, url)
            inspections.append({"url": url, **result})
        except Exception as e:
            inspections.append({"url": url, "error": str(e)})
    return inspections


# ---------------------------------------------------------------------------
# GSC page screenshots via Selenium
# ---------------------------------------------------------------------------

def build_gsc_url(page, property_url, email=None):
    """Build a GSC console URL for a specific page."""
    resource_id = property_url
    base = "https://search.google.com/search-console"
    qs = f"resource_id={urllib.parse.quote(resource_id, safe='')}"
    return f"{base}/{page}?{qs}"


def capture_gsc_screenshots(driver, property_url, email, out_dir, pages=None, log_fn=None):
    """Capture GSC page screenshots using the existing Selenium browser.
    The browser must already be logged into a Google account with GSC access."""
    if log_fn is None:
        log_fn = print
    if pages is None:
        pages = [
            {"key": "sitemap",  "page": "sitemaps",                     "wait": 8},
            {"key": "manual",   "page": "manual-actions",               "wait": 12},
            {"key": "perf",     "page": "performance/search-analytics", "wait": 15},
            {"key": "security", "page": "security-issues",              "wait": 8},
            {"key": "removals", "page": "removals",                     "wait": 8},
        ]

    os.makedirs(out_dir, exist_ok=True)
    screenshots = {}

    for p in pages:
        url = build_gsc_url(p["page"], property_url, email)
        log_fn(f"  Capturing {p['key']}...")
        try:
            driver.get(url)
            time.sleep(p["wait"])
            ss_path = os.path.join(out_dir, f"gsc_{p['key']}.png")
            driver.save_screenshot(ss_path)
            screenshots[p["key"]] = ss_path
            log_fn(f"  {p['key']} captured.")
        except Exception as e:
            log_fn(f"  {p['key']} capture failed: {e}")

    return screenshots


# ---------------------------------------------------------------------------
# PPTX Report Builders
# ---------------------------------------------------------------------------

def _init_pptx():
    """Import python-pptx and return the module."""
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    return Presentation, Inches, Pt, Emu, RGBColor, PP_ALIGN, MSO_ANCHOR


def _add_image_slide(prs, img_path, Inches, Pt, RGBColor, PP_ALIGN,
                     slide_w=13.33, slide_h=7.5,
                     img_x=0.5, img_y=1.8, img_w=12.33, img_h=5.1):
    """Add an image to a slide, fitting within the given bounds (contain sizing)."""
    from pptx.util import Emu
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank

    if img_path and os.path.exists(img_path):
        from PIL import Image
        with Image.open(img_path) as im:
            iw, ih = im.size
        aspect = iw / ih
        target_w = Inches(img_w)
        target_h = Inches(img_h)
        if aspect > (img_w / img_h):
            final_w = target_w
            final_h = int(target_w / aspect)
        else:
            final_h = target_h
            final_w = int(target_h * aspect)
        cx = Inches(img_x) + (target_w - final_w) // 2
        cy = Inches(img_y) + (target_h - final_h) // 2
        slide.shapes.add_picture(img_path, cx, cy, final_w, final_h)

    return slide


def _text(slide, text_str, x, y, w, h, font_size, font_name, color, bold=False,
          align=None, char_spacing=None, italic=False):
    """Add a text box to a slide."""
    from pptx.util import Inches, Pt, Emu
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN

    txBox = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = str(text_str)
    run = p.runs[0] if p.runs else p.add_run()
    run.text = str(text_str)
    run.font.size = Pt(font_size)
    run.font.name = font_name
    if isinstance(color, str):
        color = RGBColor(int(color[1:3], 16), int(color[3:5], 16), int(color[5:7], 16))
    run.font.color.rgb = color
    run.font.bold = bold
    if italic:
        run.font.italic = True
    if align:
        p.alignment = align
    if char_spacing is not None:
        run.font.spacing = Pt(char_spacing)
    return txBox


def _rect(slide, x, y, w, h, color):
    """Add a rectangle shape."""
    from pptx.util import Inches
    from pptx.dml.color import RGBColor

    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE.RECTANGLE
        Inches(x), Inches(y), Inches(w), Inches(h)
    )
    if isinstance(color, str):
        color = RGBColor(int(color[1:3], 16), int(color[3:5], 16), int(color[5:7], 16))
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    return shape


def _fmt_num(n):
    if n is None:
        return "0"
    if isinstance(n, float) and n < 1:
        return f"{n:.1%}"
    if isinstance(n, float):
        return f"{n:.1f}"
    return f"{n:,}" if isinstance(n, int) else str(n)


def _totals(rows):
    clicks = sum(r.get("clicks", 0) for r in rows)
    impressions = sum(r.get("impressions", 0) for r in rows)
    ctr = clicks / impressions if impressions else 0
    positions = [r.get("position", 0) for r in rows if r.get("position")]
    avg_pos = sum(positions) / len(positions) if positions else 0
    return clicks, impressions, ctr, avg_pos


# --- James Full (19 slides) ---

def build_james_full(data, out_path, log_fn=None):
    if log_fn is None:
        log_fn = print
    Presentation, Inches, Pt, Emu, RGBColor, PP_ALIGN, MSO_ANCHOR = _init_pptx()

    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)

    C = lambda h: RGBColor(int(h[1:3],16), int(h[3:5],16), int(h[5:7],16))
    TEAL_BG = C("#165060"); TEAL_HDR = C("#0D3D4A"); GOLD = C("#C9A84C")
    WHITE = RGBColor(0xFF,0xFF,0xFF); OFF_WHITE = C("#E8F4F6")
    CARD_BG = C("#EEF6FA"); CARD_BG2 = C("#F5F8FA")
    TEXT_DARK = C("#1A1A2E"); TEXT_SOFT = C("#5A6B7A")
    KPI_ORANGE = C("#E8A020"); KPI_BLUE = C("#2980B9")
    KPI_GREEN = C("#27AE60"); KPI_PURPLE = C("#8E44AD")
    STATUS_OK = C("#27AE60"); STATUS_WARN = C("#E8A020"); STATUS_ERR = C("#C0392B")

    domain = data.get("domain", "")
    property_url = data.get("propertyUrl", data.get("property_url", ""))
    start_date = data.get("startDate", data.get("start_date", ""))
    end_date = data.get("endDate", data.get("end_date", ""))
    sitemaps = data.get("sitemaps", [])
    top_queries = data.get("topQueries", data.get("top_queries", []))
    top_pages = data.get("topPages", data.get("top_pages", []))
    perf_daily = data.get("perfDaily", data.get("perf_daily", []))
    image_perf = data.get("imagePerf", data.get("image_perf", []))
    inspections = data.get("inspections", [])
    screenshots = data.get("screenshots", {})

    def header_slide(title, desc="", status_text=""):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        _rect(slide, 0, 0, 13.33, 1.0, "#0D3D4A")
        _rect(slide, 0, 1.0, 13.33, 0.08, "#C9A84C")
        _text(slide, title, 0.5, 0.15, 12.3, 0.55, 26, "Trebuchet MS", WHITE, bold=True)
        if desc:
            _text(slide, desc, 0.5, 0.62, 12.3, 0.3, 14, "Calibri", OFF_WHITE)
        if status_text:
            _text(slide, status_text, 0.5, 1.18, 12.3, 0.3, 14, "Calibri", STATUS_OK)
        return slide

    def add_table(slide, headers, rows, col_widths, start_y=1.8):
        from pptx.util import Inches as I, Pt as P
        total_w = sum(col_widths)
        tbl_shape = slide.shapes.add_table(
            len(rows) + 1, len(headers), I(0.5), I(start_y), I(total_w), I(0.4 * (len(rows) + 1))
        )
        table = tbl_shape.table
        for i, w in enumerate(col_widths):
            table.columns[i].width = I(w)
        for i, h in enumerate(headers):
            cell = table.cell(0, i)
            cell.text = h
            for p in cell.text_frame.paragraphs:
                for r in p.runs:
                    r.font.size = P(11)
                    r.font.bold = True
                    r.font.color.rgb = WHITE
                    r.font.name = "Calibri"
            cell.fill.solid()
            cell.fill.fore_color.rgb = TEAL_HDR
        for ri, row in enumerate(rows):
            for ci, val in enumerate(row):
                cell = table.cell(ri + 1, ci)
                cell.text = str(val)
                for p in cell.text_frame.paragraphs:
                    for r in p.runs:
                        r.font.size = P(10)
                        r.font.name = "Calibri"
                        r.font.color.rgb = TEXT_DARK

    def kpi_cards(slide, values, y=2.4):
        """values = list of (label, value, color_hex)"""
        from pptx.util import Inches as I, Pt as P
        start_x = 0.5
        card_w = 2.95
        card_h = 2.2
        gap = 0.25
        for i, (label, val, clr) in enumerate(values[:4]):
            cx = start_x + i * (card_w + gap)
            _rect(slide, cx, y, card_w, card_h, "#EEF6FA")
            _rect(slide, cx, y, card_w, 0.12, clr)
            _text(slide, _fmt_num(val), cx, y + 0.5, card_w, 0.6, 30, "Calibri",
                  C(clr), bold=True, align=PP_ALIGN.CENTER)
            _text(slide, label, cx, y + 1.3, card_w, 0.4, 12, "Calibri",
                  TEXT_SOFT, align=PP_ALIGN.CENTER)

    # --- Slide 1: Title ---
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _rect(s, 0, 0, 13.33, 7.5, "#165060")
    _rect(s, 0, 0, 0.4, 7.5, "#C9A84C")
    _text(s, "GSC AUDIT REPORT", 1.2, 2.3, 10, 0.8, 36, "Trebuchet MS", WHITE, bold=True)
    _rect(s, 1.2, 3.4, 5, 0.06, "#C9A84C")
    _text(s, domain, 1.2, 3.6, 10, 0.5, 22, "Trebuchet MS", GOLD)

    # --- Slide 2: Introduction ---
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _rect(s, 0, 0, 13.33, 7.5, "#F5F8FA")
    _rect(s, 0, 0, 13.33, 1.0, "#0D3D4A")
    _rect(s, 0, 1.0, 13.33, 0.08, "#C9A84C")
    _text(s, "Introduction", 0.5, 0.15, 12.3, 0.55, 26, "Trebuchet MS", WHITE, bold=True)
    boxes = [
        ("Property", property_url),
        ("Pages Inspected", str(len(inspections))),
        ("Sections", "Sitemaps, Indexing, Performance, Manual Actions, Security"),
    ]
    for i, (label, val) in enumerate(boxes):
        bx = 0.47 + i * 4.3
        _rect(s, bx, 2.5, 4.0, 2.3, "#FFFFFF")
        _text(s, label, bx + 0.3, 2.7, 3.4, 0.4, 14, "Calibri", TEXT_SOFT, bold=True)
        _text(s, val, bx + 0.3, 3.2, 3.4, 1.0, 12, "Calibri", TEXT_DARK)

    # --- Slide 3: Sitemap Status ---
    s = header_slide("Sitemap Status", f"{len(sitemaps)} sitemap(s) found")
    sitemap_rows = []
    for sm in sitemaps:
        path = sm.get("path", "")
        sm_type = sm.get("type", "")
        is_index = "Yes" if sm.get("isSitemapsIndex") else "No"
        errors = str(sm.get("errors", 0))
        warnings = str(sm.get("warnings", 0))
        pending = "Yes" if sm.get("isPending") else "No"
        sitemap_rows.append([path, sm_type, is_index, errors, warnings, pending])
    add_table(s, ["Sitemap URL", "Type", "Index?", "Errors", "Warnings", "Pending"],
              sitemap_rows or [["No sitemaps found", "", "", "", "", ""]],
              [5.3, 1.7, 1.5, 1.4, 1.2, 1.2])

    # --- Slide 4: Sitemap Last Read ---
    s = header_slide("Sitemap Last Read")
    lr_rows = []
    for sm in sitemaps:
        contents = sm.get("contents", [])
        submitted = contents[0].get("submitted", "N/A") if contents else "N/A"
        lr_rows.append([sm.get("path", ""), sm.get("lastSubmitted", "N/A"),
                        sm.get("lastDownloaded", "N/A"), submitted])
    add_table(s, ["Sitemap URL", "Last Submitted", "Last Downloaded", "URLs Submitted"],
              lr_rows or [["N/A", "N/A", "N/A", "N/A"]],
              [5.5, 2.3, 2.3, 2.2])

    # --- Slides 5-10: Inspection-based slides ---
    insp_valid = [i for i in inspections if not i.get("error")]

    # Slide 5: URL Indexing
    s = header_slide("URL Indexing Status", f"{len(insp_valid)} URLs inspected")
    idx_rows = []
    for ins in insp_valid:
        isr = ins.get("indexStatusResult", {})
        idx_rows.append([ins["url"][:50], isr.get("verdict", "N/A"),
                         isr.get("coverageState", "N/A"), isr.get("indexingState", "N/A")])
    add_table(s, ["URL", "Verdict", "Coverage", "Indexing State"],
              idx_rows or [["No inspection data", "", "", ""]],
              [5.5, 1.5, 2.8, 2.5])

    # Slide 6: Canonical Check
    s = header_slide("Canonical URL Check")
    canon_rows = []
    for ins in insp_valid:
        isr = ins.get("indexStatusResult", {})
        user_c = isr.get("userCanonical", "N/A")
        google_c = isr.get("googleCanonical", "N/A")
        match = "Yes" if user_c == google_c else "No"
        canon_rows.append([ins["url"][:35], user_c[:35] if user_c else "N/A",
                           google_c[:35] if google_c else "N/A", match])
    add_table(s, ["URL", "User Canonical", "Google Canonical", "Match"],
              canon_rows or [["No data", "", "", ""]],
              [3.8, 4.0, 4.0, 0.5])

    # Slide 7: Robots.txt
    s = header_slide("Robots.txt Status")
    robot_rows = []
    for ins in insp_valid:
        isr = ins.get("indexStatusResult", {})
        robot_rows.append([ins["url"][:70], isr.get("robotsTxtState", "N/A")])
    add_table(s, ["URL", "Robots.txt State"],
              robot_rows or [["No data", ""]],
              [9.3, 3.0])

    # Slide 8: Fetchability
    s = header_slide("Page Fetchability")
    fetch_rows = []
    for ins in insp_valid:
        isr = ins.get("indexStatusResult", {})
        fetch_rows.append([ins["url"][:60], isr.get("pageFetchState", "N/A"),
                           isr.get("lastCrawlTime", "N/A")[:19]])
    add_table(s, ["URL", "Fetch State", "Last Crawl"],
              fetch_rows or [["No data", "", ""]],
              [7.5, 2.5, 2.3])

    # Slide 9: Crawl Type
    s = header_slide("Crawl Type")
    crawl_rows = []
    for ins in insp_valid:
        isr = ins.get("indexStatusResult", {})
        crawl_rows.append([ins["url"][:70], isr.get("crawledAs", "N/A")])
    add_table(s, ["URL", "Crawled As"],
              crawl_rows or [["No data", ""]],
              [9.3, 3.0])

    # Slide 10: Rich Results
    s = header_slide("Rich Results")
    rich_rows = []
    for ins in insp_valid:
        rr = ins.get("richResultsResult", {})
        verdict = rr.get("verdict", "N/A")
        items = rr.get("detectedItems", [])
        item_str = ", ".join(i.get("richResultType", "") for i in items) if items else "None"
        rich_rows.append([ins["url"][:50], item_str, verdict])
    add_table(s, ["URL", "Detected Items", "Verdict"],
              rich_rows or [["No data", "", ""]],
              [6.0, 4.3, 2.0])

    # --- Slide 11: Image Search Performance ---
    s = header_slide("Image Search Performance")
    img_clicks, img_impr, img_ctr, img_pos = _totals(image_perf)
    kpi_cards(s, [
        ("Total Clicks", img_clicks, "#E8A020"),
        ("Total Impressions", img_impr, "#2980B9"),
        ("Avg CTR", img_ctr, "#27AE60"),
        ("Avg Position", round(img_pos, 1) if img_pos else 0, "#8E44AD"),
    ])
    img_rows = []
    for q in image_perf[:5]:
        keys = q.get("keys", [""])
        img_rows.append([keys[0], _fmt_num(q.get("clicks", 0)), _fmt_num(q.get("impressions", 0)),
                         f"{q.get('ctr', 0):.1%}", f"{q.get('position', 0):.1f}"])
    if img_rows:
        add_table(s, ["#", "Query", "Clicks", "Impressions", "CTR", "Pos"],
                  [[str(i+1)] + r for i, r in enumerate(img_rows)],
                  [0.5, 5.3, 1.7, 2.0, 1.7, 1.6], start_y=5.0)

    # --- Slide 12: Performance Overview ---
    s = header_slide("Performance Overview", f"{start_date} to {end_date}")
    total_clicks, total_impr, total_ctr, avg_pos = _totals(perf_daily)
    kpi_cards(s, [
        ("Total Clicks", total_clicks, "#E8A020"),
        ("Total Impressions", total_impr, "#2980B9"),
        ("Avg CTR", total_ctr, "#27AE60"),
        ("Avg Position", round(avg_pos, 1) if avg_pos else 0, "#8E44AD"),
    ])

    # --- Slide 13: Top Queries ---
    s = header_slide("Top Search Queries", f"Top {len(top_queries)} queries by clicks")
    q_rows = []
    for i, q in enumerate(top_queries):
        keys = q.get("keys", [""])
        q_rows.append([str(i+1), keys[0], _fmt_num(q.get("clicks", 0)),
                       _fmt_num(q.get("impressions", 0)),
                       f"{q.get('ctr', 0):.1%}", f"{q.get('position', 0):.1f}"])
    add_table(s, ["#", "Query", "Clicks", "Impressions", "CTR", "Position"],
              q_rows or [["", "No data", "", "", "", ""]],
              [0.5, 5.8, 1.4, 1.8, 1.4, 1.4])

    # --- Slide 14: Top Pages ---
    s = header_slide("Top Pages", f"Top {len(top_pages)} pages by clicks")
    p_rows = []
    for i, pg in enumerate(top_pages):
        keys = pg.get("keys", [""])
        p_rows.append([str(i+1), keys[0][:50], _fmt_num(pg.get("clicks", 0)),
                       _fmt_num(pg.get("impressions", 0)),
                       f"{pg.get('ctr', 0):.1%}", f"{pg.get('position', 0):.1f}"])
    add_table(s, ["#", "Page", "Clicks", "Impressions", "CTR", "Position"],
              p_rows or [["", "No data", "", "", "", ""]],
              [0.5, 5.8, 1.4, 1.8, 1.4, 1.4])

    # --- Slide 15: Keyword Opportunities ---
    s = header_slide("Keyword Opportunities", "High-impression, low-position queries")
    opps = sorted([q for q in top_queries if q.get("position", 0) > 10],
                  key=lambda x: x.get("impressions", 0), reverse=True)[:10]
    opp_rows = []
    for q in opps:
        keys = q.get("keys", [""])
        opp_rows.append([keys[0], _fmt_num(q.get("impressions", 0)),
                         _fmt_num(q.get("clicks", 0)),
                         f"{q.get('ctr', 0):.1%}", f"{q.get('position', 0):.1f}"])
    add_table(s, ["Query", "Impressions", "Clicks", "CTR", "Current Position"],
              opp_rows or [["No keyword opportunities found", "", "", "", ""]],
              [5.5, 1.7, 1.4, 1.4, 2.3])

    # --- Slides 16-18: Screenshots ---
    for shot_key, title in [("manual", "Manual Action"), ("security", "Security Issues"),
                            ("removals", "Removals")]:
        s = header_slide(title)
        img_path = screenshots.get(shot_key, "")
        if img_path and os.path.exists(img_path):
            _add_image_slide.__wrapped__ if hasattr(_add_image_slide, '__wrapped__') else None
            # Add image directly to this slide
            from PIL import Image
            with Image.open(img_path) as im:
                iw, ih = im.size
            aspect = iw / ih
            target_w = Inches(12.33)
            target_h = Inches(5.1)
            if aspect > (12.33 / 5.1):
                fw = target_w; fh = int(target_w / aspect)
            else:
                fh = target_h; fw = int(target_h * aspect)
            cx = Inches(0.5) + (target_w - fw) // 2
            cy = Inches(1.8) + (target_h - fh) // 2
            s.shapes.add_picture(img_path, cx, cy, fw, fh)
        else:
            _text(s, "Screenshot not available. Please check Google Search Console.",
                  0.5, 3.5, 12.3, 0.5, 16, "Calibri", TEXT_SOFT, align=PP_ALIGN.CENTER)

    # --- Slide 19: Thank You ---
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _rect(s, 0, 0, 13.33, 7.5, "#165060")
    _rect(s, 0, 0, 0.4, 7.5, "#C9A84C")
    _text(s, "THANK YOU", 1.2, 2.8, 10, 1.0, 48, "Trebuchet MS", WHITE, bold=True)
    _rect(s, 1.2, 4.0, 5, 0.06, "#C9A84C")
    _text(s, domain, 1.2, 4.3, 10, 0.5, 18, "Trebuchet MS", GOLD)

    prs.save(out_path)
    log_fn(f"  James Full report saved: {os.path.basename(out_path)}")
    return out_path


# --- James Short (8 slides) ---

def build_james_short(data, out_path, log_fn=None):
    if log_fn is None:
        log_fn = print
    Presentation, Inches, Pt, Emu, RGBColor, PP_ALIGN, MSO_ANCHOR = _init_pptx()

    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)

    C = lambda h: RGBColor(int(h[1:3],16), int(h[3:5],16), int(h[5:7],16))
    BG = C("#00292E"); PANEL = C("#054854"); CARD = C("#107082")
    TEAL_MD = C("#64B2C1"); BEIGE = C("#F0CDA1")
    WHITE = RGBColor(0xFF,0xFF,0xFF); YELLOW = C("#FFFF00")

    domain = data.get("domain", "")
    screenshots = data.get("screenshots", {})

    def screenshot_slide(title, desc, shot_key, status_text=""):
        s = prs.slides.add_slide(prs.slide_layouts[6])
        _rect(s, 0, 0, 13.33, 7.5, "#00292E")
        _rect(s, 11.5, 0, 1.83, 0.28, "#F0CDA1")
        _text(s, title, 0.689, 0.25, 12.0, 0.70, 32, "Calibri", WHITE, bold=True)
        _rect(s, 0.689, 0.87, 5.5, 0.05, "#F0CDA1")
        if desc:
            _text(s, desc, 0.622, 1.06, 12.0, 0.4, 20, "Calibri", TEAL_MD)
        if status_text:
            txb = _text(s, "Status: ", 0.622, 1.96, 2.0, 0.3, 16, "Calibri", WHITE)
            _text(s, status_text, 2.5, 1.96, 5.0, 0.3, 16, "Calibri", YELLOW, bold=True)

        # Screenshot frame
        _rect(s, 0.38, 2.55, 12.57, 4.65, "#F0CDA1")
        _rect(s, 0.42, 2.59, 12.49, 4.57, "#00292E")

        img_path = screenshots.get(shot_key, "")
        if img_path and os.path.exists(img_path):
            from PIL import Image
            with Image.open(img_path) as im:
                iw, ih = im.size
            aspect = iw / ih
            target_w = Inches(12.33); target_h = Inches(4.4)
            if aspect > (12.33/4.4):
                fw = target_w; fh = int(target_w / aspect)
            else:
                fh = target_h; fw = int(target_h * aspect)
            cx = Inches(0.5) + (target_w - fw) // 2
            cy = Inches(2.7) + (target_h - fh) // 2
            s.shapes.add_picture(img_path, cx, cy, fw, fh)
        else:
            _text(s, "Screenshot not available", 3, 4.5, 7, 0.5, 18, "Calibri",
                  TEAL_MD, align=PP_ALIGN.CENTER)
        return s

    # Slide 1: Title
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _rect(s, 0, 0, 13.33, 7.5, "#00292E")
    _rect(s, 0, 0, 3.2, 7.5, "#054854")
    _rect(s, 3.15, 0, 0.15, 7.5, "#F0CDA1")
    _text(s, "GOOGLE SEARCH CONSOLE", 4.0, 2.0, 8.0, 0.8, 38, "Calibri", WHITE, bold=True)
    _text(s, "AUDIT REPORT", 4.0, 2.8, 8.0, 0.8, 38, "Calibri", BEIGE, bold=True)
    _rect(s, 4.0, 4.72, 7.0, 0.65, "#107082")
    _text(s, domain, 4.2, 4.77, 6.6, 0.55, 18, "Calibri", BEIGE, italic=True)

    # Slide 2: Introduction
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _rect(s, 0, 0, 13.33, 7.5, "#00292E")
    _rect(s, 8.35, 0, 4.98, 7.5, "#054854")
    _rect(s, 1.255, 2.988, 7.549, 3.838, "#107082")
    _text(s, "INTRODUCTION", 1.5, 3.2, 7.0, 0.8, 44, "Calibri", WHITE, bold=True)
    _text(s, f"This report provides a GSC audit overview for {domain}, "
             "covering sitemaps, performance, manual actions, security issues, and removals.",
          1.5, 4.2, 7.0, 2.0, 13, "Calibri", TEAL_MD)

    # Slides 3-7: Screenshots
    screenshot_slide("Sitemap Status", "Overview of submitted sitemaps", "sitemap")
    screenshot_slide("Performance Overview", "Search performance analytics", "perf")
    screenshot_slide("Manual Action", "Manual actions status", "manual")
    screenshot_slide("Security Issues", "Security issues status", "security")
    screenshot_slide("Removals", "URL removals status", "removals")

    # Slide 8: Thank You
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _rect(s, 0, 0, 13.33, 7.5, "#00292E")
    _rect(s, 9.8, 0, 3.53, 7.5, "#054854")
    _text(s, "THANK YOU", 1.0, 3.0, 8.0, 1.0, 60, "Calibri", WHITE, bold=True)
    _rect(s, 1.0, 4.3, 5, 0.06, "#F0CDA1")

    prs.save(out_path)
    log_fn(f"  James Short report saved: {os.path.basename(out_path)}")
    return out_path


# --- Sigma (10 slides) ---

def build_sigma(data, out_path, log_fn=None):
    if log_fn is None:
        log_fn = print
    Presentation, Inches, Pt, Emu, RGBColor, PP_ALIGN, MSO_ANCHOR = _init_pptx()

    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)

    C = lambda h: RGBColor(int(h[1:3],16), int(h[3:5],16), int(h[5:7],16))
    CREAM = C("#EFEDE3"); INK = C("#191B0E"); INK_SOFT = C("#6B6A5C")
    GOLD = C("#E6C069"); GOLD_DK = C("#B8923B"); GREEN = C("#00B050"); RED = C("#C00000")
    WHITE = RGBColor(0xFF,0xFF,0xFF); ROW_ALT = C("#F6F2E8")

    domain = data.get("domain", "")
    property_url = data.get("propertyUrl", data.get("property_url", ""))
    screenshots = data.get("screenshots", {})
    top_queries = data.get("topQueries", data.get("top_queries", []))
    top_pages = data.get("topPages", data.get("top_pages", []))
    generated = data.get("generatedDate", datetime.now().strftime("%B %d, %Y"))

    def header_slide(title, desc=""):
        s = prs.slides.add_slide(prs.slide_layouts[6])
        _rect(s, 0, 0, 13.33, 7.5, "#EFEDE3")
        _rect(s, 11.5, 0, 1.83, 0.28, "#E6C069")
        _text(s, title, 0.689, 0.25, 12.0, 0.6, 32, "Calibri", INK, bold=True)
        _rect(s, 0.689, 0.95, 5.5, 0.05, "#E6C069")
        if desc:
            _text(s, desc, 0.689, 1.1, 12.0, 0.4, 16, "Calibri", INK_SOFT)
        return s

    def screenshot_slide(title, shot_key, desc=""):
        s = header_slide(title, desc)
        _rect(s, 0.38, 2.55, 12.57, 4.65, "#E6C069")
        _rect(s, 0.42, 2.59, 12.49, 4.57, "#EFEDE3")
        img_path = screenshots.get(shot_key, "")
        if img_path and os.path.exists(img_path):
            from PIL import Image
            with Image.open(img_path) as im:
                iw, ih = im.size
            aspect = iw / ih
            tw = Inches(12.33); th = Inches(4.4)
            if aspect > (12.33/4.4): fw = tw; fh = int(tw / aspect)
            else: fh = th; fw = int(th * aspect)
            cx = Inches(0.5) + (tw - fw) // 2
            cy = Inches(2.7) + (th - fh) // 2
            s.shapes.add_picture(img_path, cx, cy, fw, fh)
        else:
            _text(s, "Screenshot not available", 3, 4.5, 7, 0.5, 18, "Calibri",
                  INK_SOFT, align=PP_ALIGN.CENTER)
        return s

    def data_table(slide, headers, rows, col_widths, y=2.1):
        from pptx.util import Inches as I, Pt as P
        total_w = sum(col_widths)
        tbl_shape = slide.shapes.add_table(
            len(rows) + 1, len(headers), I(0.5), I(y), I(total_w), I(0.35 * (len(rows) + 1))
        )
        table = tbl_shape.table
        for i, w in enumerate(col_widths):
            table.columns[i].width = I(w)
        for i, h in enumerate(headers):
            cell = table.cell(0, i)
            cell.text = h
            for p in cell.text_frame.paragraphs:
                for r in p.runs:
                    r.font.size = P(11); r.font.bold = True; r.font.name = "Calibri"
                    r.font.color.rgb = INK
            cell.fill.solid()
            cell.fill.fore_color.rgb = GOLD
        for ri, row in enumerate(rows):
            for ci, val in enumerate(row):
                cell = table.cell(ri+1, ci)
                cell.text = str(val)
                for p in cell.text_frame.paragraphs:
                    for r in p.runs:
                        r.font.size = P(10); r.font.name = "Calibri"; r.font.color.rgb = INK
                if ri % 2 == 1:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = ROW_ALT

    # Slide 1: Title
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _rect(s, 0, 0, 13.33, 7.5, "#EFEDE3")
    _rect(s, 0, 0, 13.33, 0.35, "#E6C069")
    _rect(s, 0, 7.15, 13.33, 0.35, "#E6C069")
    _text(s, "WEBMASTER AUDIT REPORT", 0, 2.8, 13.33, 0.9, 44, "Calibri", INK,
          bold=True, align=PP_ALIGN.CENTER)
    _text(s, domain, 0, 4.0, 13.33, 0.5, 22, "Calibri", GREEN, align=PP_ALIGN.CENTER)
    _text(s, f"Generated: {generated}", 0, 4.7, 13.33, 0.4, 14, "Calibri",
          INK_SOFT, align=PP_ALIGN.CENTER)

    # Slide 2: Introduction
    s = header_slide("Introduction")
    _rect(s, 0.689, 1.5, 0.12, 4.5, "#E6C069")
    _text(s, f"This audit report covers the Google Search Console data for {domain}. "
             "It includes performance metrics, top queries, top pages, sitemap status, "
             "manual actions, security issues, and removals.",
          1.0, 1.6, 11.5, 3.0, 18, "Calibri", INK)

    # Slide 3: Performance screenshot
    screenshot_slide("Performance Overview", "perf", "Search performance analytics")

    # Slide 4: Top Queries table
    s = header_slide("Top Search Queries", f"Top {len(top_queries)} queries by clicks")
    q_rows = []
    for q in top_queries:
        keys = q.get("keys", [""])
        q_rows.append([keys[0], _fmt_num(q.get("clicks", 0)),
                       _fmt_num(q.get("impressions", 0)),
                       f"{q.get('ctr', 0):.1%}", f"{q.get('position', 0):.1f}"])
    data_table(s, ["Query", "Clicks", "Impressions", "CTR", "Position"],
               q_rows or [["No data", "", "", "", ""]],
               [6.05, 1.4, 1.8, 1.3, 1.4])

    # Slide 5: Top Pages table
    s = header_slide("Top Pages", f"Top {len(top_pages)} pages by clicks")
    pg_rows = []
    for pg in top_pages:
        keys = pg.get("keys", [""])
        pg_rows.append([keys[0][:50], _fmt_num(pg.get("clicks", 0)),
                        _fmt_num(pg.get("impressions", 0)),
                        f"{pg.get('ctr', 0):.1%}", f"{pg.get('position', 0):.1f}"])
    data_table(s, ["Page", "Clicks", "Impressions", "CTR", "Position"],
               pg_rows or [["No data", "", "", "", ""]],
               [6.05, 1.4, 1.8, 1.3, 1.4])

    # Slides 6-9: Screenshots
    screenshot_slide("Sitemap Status", "sitemap", "Submitted sitemaps overview")
    screenshot_slide("Manual Action", "manual", "Manual actions from Google")
    screenshot_slide("Security Issues", "security", "Security issues detected")
    screenshot_slide("Removals", "removals", "URL removals status")

    # Slide 10: Thank You
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _rect(s, 0, 0, 13.33, 7.5, "#EFEDE3")
    _rect(s, 0, 0, 13.33, 0.35, "#E6C069")
    _rect(s, 0, 7.15, 13.33, 0.35, "#E6C069")
    _text(s, "THANK YOU", 0, 3.0, 13.33, 1.0, 60, "Calibri", INK,
          bold=True, align=PP_ALIGN.CENTER)
    _rect(s, 4.66, 4.3, 4.0, 0.06, "#E6C069")

    prs.save(out_path)
    log_fn(f"  Sigma report saved: {os.path.basename(out_path)}")
    return out_path


# --- Omega (8 slides) ---

def build_omega(data, out_path, log_fn=None):
    if log_fn is None:
        log_fn = print
    Presentation, Inches, Pt, Emu, RGBColor, PP_ALIGN, MSO_ANCHOR = _init_pptx()

    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)

    C = lambda h: RGBColor(int(h[1:3],16), int(h[3:5],16), int(h[5:7],16))
    NAVY = C("#1B2A41"); NAVY_DARK = C("#0F1A2C"); BRONZE = C("#B08D57")
    BRONZE_SOFT = C("#D4B891"); OFF_WHITE = C("#F8F7F4")
    TEXT_DARK = C("#2C2C2C"); TEXT_SOFT = C("#6B7280"); BORDER = C("#E5E0D5")
    STATUS_OK = C("#4A7C59")
    WHITE = RGBColor(0xFF,0xFF,0xFF)

    domain = data.get("domain", "")
    screenshots = data.get("screenshots", {})

    def header_slide(title, desc=""):
        s = prs.slides.add_slide(prs.slide_layouts[6])
        _rect(s, 0, 0, 13.33, 7.5, "#F8F7F4")
        _rect(s, 0, 0, 13.33, 1.2, "#1B2A41")
        _rect(s, 0, 1.2, 13.33, 0.06, "#B08D57")
        _rect(s, 0, 7.1, 13.33, 0.4, "#1B2A41")
        _text(s, title.upper(), 0.5, 0.3, 12.3, 0.6, 28, "Georgia", OFF_WHITE,
              bold=True, char_spacing=2)
        if desc:
            _text(s, desc, 0.5, 1.5, 12.3, 0.4, 14, "Calibri", TEXT_SOFT, italic=True)
        return s

    def screenshot_slide(title, shot_key, desc=""):
        s = header_slide(title, desc)
        # White card frame
        _rect(s, 0.7, 2.15, 11.93, 4.85, "#FFFFFF")
        img_path = screenshots.get(shot_key, "")
        if img_path and os.path.exists(img_path):
            from PIL import Image
            with Image.open(img_path) as im:
                iw, ih = im.size
            aspect = iw / ih
            tw = Inches(11.73); th = Inches(4.65)
            if aspect > (11.73/4.65): fw = tw; fh = int(tw / aspect)
            else: fh = th; fw = int(th * aspect)
            cx = Inches(0.8) + (tw - fw) // 2
            cy = Inches(2.25) + (th - fh) // 2
            s.shapes.add_picture(img_path, cx, cy, fw, fh)
        else:
            _text(s, "Screenshot not available", 3, 4.5, 7, 0.5, 18, "Calibri",
                  TEXT_SOFT, align=PP_ALIGN.CENTER)
        return s

    # Slide 1: Title
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _rect(s, 0, 0, 13.33, 7.5, "#F8F7F4")
    _rect(s, 0, 0, 13.33, 0.5, "#B08D57")
    _rect(s, 0, 7.0, 13.33, 0.5, "#1B2A41")
    _text(s, "WEBMASTER AUDIT REPORT", 0, 2.8, 13.33, 0.8, 40, "Georgia", TEXT_DARK,
          bold=True, align=PP_ALIGN.CENTER)
    _rect(s, 5.66, 3.8, 2.0, 0.06, "#B08D57")
    _text(s, domain, 0, 4.2, 13.33, 0.5, 24, "Georgia", TEXT_SOFT,
          italic=True, align=PP_ALIGN.CENTER)

    # Slide 2: Introduction
    s = header_slide("Introduction")
    _text(s, f"Domain: {domain}", 0.5, 2.0, 12.3, 0.4, 16, "Calibri", TEXT_DARK, bold=True)
    _text(s, "This report covers your Google Search Console audit including sitemaps, "
             "manual actions, performance, security issues, and other notable items.",
          0.5, 2.6, 12.3, 2.0, 14, "Calibri", TEXT_SOFT)

    # Slides 3-6: Screenshots
    screenshot_slide("Sitemap Status", "sitemap", "Submitted sitemaps and their status")
    screenshot_slide("Manual Action", "manual", "Manual actions from Google Search team")
    screenshot_slide("Performance", "perf", "Search performance analytics overview")
    screenshot_slide("Security Issues", "security", "Security issues detected by Google")

    # Slide 7: Other Issues (Removals or Not Found)
    s = header_slide("Other Issues")
    img_path = screenshots.get("removals", "")
    if img_path and os.path.exists(img_path):
        _rect(s, 0.7, 2.15, 11.93, 4.85, "#FFFFFF")
        from PIL import Image
        with Image.open(img_path) as im:
            iw, ih = im.size
        aspect = iw / ih
        tw = Inches(11.73); th = Inches(4.65)
        if aspect > (11.73/4.65): fw = tw; fh = int(tw / aspect)
        else: fh = th; fw = int(th * aspect)
        cx = Inches(0.8) + (tw - fw) // 2
        cy = Inches(2.25) + (th - fh) // 2
        s.shapes.add_picture(img_path, cx, cy, fw, fh)
    else:
        _rect(s, 3.5, 3.5, 6.3, 1.2, "#FFFFFF")
        _text(s, "NOT FOUND", 3.5, 3.7, 6.3, 0.8, 32, "Georgia", STATUS_OK,
              align=PP_ALIGN.CENTER)

    # Slide 8: Thank You
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _rect(s, 0, 0, 13.33, 7.5, "#1B2A41")
    _text(s, "THANK YOU", 0, 2.8, 13.33, 1.0, 56, "Georgia", BRONZE_SOFT,
          bold=True, align=PP_ALIGN.CENTER, char_spacing=8)
    _rect(s, 4.66, 4.2, 4.0, 0.06, "#B08D57")
    _text(s, domain, 0, 4.6, 13.33, 0.5, 16, "Georgia", WHITE, align=PP_ALIGN.CENTER)

    prs.save(out_path)
    log_fn(f"  Omega report saved: {os.path.basename(out_path)}")
    return out_path


# --- Neon (8 slides) ---

def build_neon(data, out_path, log_fn=None):
    if log_fn is None:
        log_fn = print
    Presentation, Inches, Pt, Emu, RGBColor, PP_ALIGN, MSO_ANCHOR = _init_pptx()

    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)

    C = lambda h: RGBColor(int(h[1:3],16), int(h[3:5],16), int(h[5:7],16))
    DARK = C("#0F1F2E"); DARK_MID = C("#162D40"); OFF_WHITE = C("#F4F7FA")
    TEAL = C("#B0BEC5"); TEAL_SOFT = C("#CFD8DC"); TEAL_DIM = C("#78909C")
    GOLD_N = C("#FFC100"); GOLD_SOFT = C("#FFD44D")
    TEXT_DARK = C("#1A2E3D"); TEXT_MID = C("#2D4A62")
    TEXT_SOFT = C("#5A7A8F"); TEXT_LIGHT = C("#C8D8E4"); BORDER_N = C("#D8E6F0")
    WHITE = RGBColor(0xFF,0xFF,0xFF)

    domain = data.get("domain", "")
    screenshots = data.get("screenshots", {})

    accent_colors = {"sitemap": TEAL, "manual": GOLD_N, "perf": TEAL_SOFT, "security": GOLD_SOFT}

    def header_slide(title, desc=""):
        s = prs.slides.add_slide(prs.slide_layouts[6])
        _rect(s, 0, 0, 13.33, 7.5, "#F4F7FA")
        _rect(s, 0, 0, 13.33, 1.3, "#0F1F2E")
        _rect(s, 0, 0, 0.18, 1.3, "#B0BEC5")
        _rect(s, 0, 1.25, 13.33, 0.05, "#FFC100")
        _rect(s, 0, 7.2, 13.33, 0.3, "#0F1F2E")
        _rect(s, 0, 7.2, 0.18, 0.3, "#FFC100")
        _text(s, title.upper(), 0.5, 0.35, 12.3, 0.6, 28, "Trebuchet MS", WHITE,
              bold=True, char_spacing=3)
        if desc:
            _text(s, desc, 0.5, 1.45, 12.3, 0.4, 13, "Calibri", TEXT_SOFT, italic=True)
        return s

    def screenshot_slide(title, shot_key, desc="", accent_hex="#B0BEC5"):
        s = header_slide(title, desc)
        # Frame
        _rect(s, 0.5, 2.1, 12.33, 4.85, "#FFFFFF")
        _rect(s, 0.5, 2.1, 12.33, 0.1, accent_hex)
        img_path = screenshots.get(shot_key, "")
        if img_path and os.path.exists(img_path):
            from PIL import Image
            with Image.open(img_path) as im:
                iw, ih = im.size
            aspect = iw / ih
            tw = Inches(12.21); th = Inches(4.65)
            if aspect > (12.21/4.65): fw = tw; fh = int(tw / aspect)
            else: fh = th; fw = int(th * aspect)
            cx = Inches(0.56) + (tw - fw) // 2
            cy = Inches(2.25) + (th - fh) // 2
            s.shapes.add_picture(img_path, cx, cy, fw, fh)
        else:
            _text(s, "Screenshot not available", 3, 4.5, 7, 0.5, 18, "Calibri",
                  TEXT_SOFT, align=PP_ALIGN.CENTER)
        return s

    # Slide 1: Title
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _rect(s, 0, 0, 13.33, 7.5, "#0F1F2E")
    _rect(s, 0, 0, 0.18, 7.5, "#B0BEC5")
    # Gold top-right L
    _rect(s, 12.0, 0, 1.33, 0.15, "#FFC100")
    _rect(s, 13.18, 0, 0.15, 1.5, "#FFC100")
    # Teal bottom-left
    _rect(s, 0, 7.0, 1.5, 0.15, "#B0BEC5")
    _text(s, "WEBMASTER", 1.0, 2.3, 11.0, 1.0, 54, "Trebuchet MS", TEXT_LIGHT, bold=True)
    _text(s, "AUDIT REPORT", 1.0, 3.3, 11.0, 1.0, 54, "Trebuchet MS", TEAL, bold=True)
    _rect(s, 1.0, 4.6, 4.5, 0.06, "#FFC100")
    _text(s, domain, 1.0, 5.0, 11.0, 0.5, 26, "Trebuchet MS", GOLD_SOFT)

    # Slide 2: Introduction
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _rect(s, 0, 0, 13.33, 7.5, "#F4F7FA")
    _rect(s, 0, 0, 13.33, 1.3, "#0F1F2E")
    _rect(s, 0, 0, 0.18, 1.3, "#B0BEC5")
    _rect(s, 0, 1.25, 13.33, 0.05, "#FFC100")
    _text(s, "INTRODUCTION", 0.5, 0.35, 12.3, 0.6, 28, "Trebuchet MS", WHITE,
          bold=True, char_spacing=3)
    _text(s, f"GSC audit report for {domain}. This report covers sitemap status, "
             "manual actions, performance analytics, and security issues.",
          0.5, 1.6, 12.3, 1.5, 16, "Calibri", TEXT_MID)
    # Domain card
    _rect(s, 0.5, 3.5, 12.33, 0.8, "#0F1F2E")
    _rect(s, 0.5, 3.5, 0.12, 0.8, "#B0BEC5")
    _text(s, domain, 0.8, 3.55, 11.8, 0.7, 16, "Calibri", WHITE)
    # Section preview boxes
    sections = [("Sitemap", "#B0BEC5"), ("Manual Action", "#FFC100"),
                ("Performance", "#CFD8DC"), ("Security", "#FFD44D")]
    for i, (label, clr) in enumerate(sections):
        bx = 0.5 + i * 3.1
        _rect(s, bx, 4.8, 2.9, 1.5, "#FFFFFF")
        _rect(s, bx, 4.8, 2.9, 0.08, clr)
        _text(s, label, bx + 0.2, 5.1, 2.5, 0.4, 13, "Calibri", TEXT_MID, bold=True)

    # Slides 3-6: Screenshots
    screenshot_slide("Sitemap Status", "sitemap", "Submitted sitemaps overview", "#B0BEC5")
    screenshot_slide("Manual Action", "manual", "Manual actions from Google", "#FFC100")
    screenshot_slide("Performance", "perf", "Search performance analytics", "#CFD8DC")
    screenshot_slide("Security Issues", "security", "Security issues detected", "#FFD44D")

    # Slide 7: Other Issues
    s = header_slide("Other Issues")
    _rect(s, 4.16, 3.0, 5.0, 1.7, "#0F1F2E")
    _rect(s, 4.16, 3.0, 5.0, 0.08, "#B0BEC5")
    _rect(s, 4.16, 4.62, 5.0, 0.08, "#FFC100")
    _text(s, "NOT FOUND", 4.16, 3.3, 5.0, 1.0, 40, "Trebuchet MS", TEAL,
          align=PP_ALIGN.CENTER, char_spacing=8)

    # Slide 8: Thank You
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _rect(s, 0, 0, 13.33, 7.5, "#0F1F2E")
    _rect(s, 13.15, 0, 0.18, 7.5, "#B0BEC5")
    # Gold bottom-left L
    _rect(s, 0, 7.35, 1.5, 0.15, "#FFC100")
    _rect(s, 0, 6.0, 0.15, 1.5, "#FFC100")
    _text(s, "THANK", 1.0, 2.5, 11.0, 1.2, 72, "Trebuchet MS", TEXT_LIGHT, bold=True)
    _text(s, "YOU", 1.0, 3.7, 11.0, 1.2, 72, "Trebuchet MS", TEXT_LIGHT, bold=True)
    _rect(s, 1.0, 5.1, 4.5, 0.06, "#FFC100")
    _text(s, domain, 1.0, 5.4, 11.0, 0.5, 20, "Trebuchet MS", GOLD_SOFT)

    prs.save(out_path)
    log_fn(f"  Neon report saved: {os.path.basename(out_path)}")
    return out_path


# ---------------------------------------------------------------------------
# Format registry & main entry point
# ---------------------------------------------------------------------------

GSC_FORMATS = {
    "james":       {"label": "James New (Full, 19 slides)", "builder": build_james_full},
    "james_short": {"label": "James Short (8 slides)",      "builder": build_james_short},
    "sigma":       {"label": "Sigma (10 slides)",           "builder": build_sigma},
    "omega":       {"label": "Omega (8 slides)",            "builder": build_omega},
    "neon":        {"label": "Neon (8 slides)",             "builder": build_neon},
}


def run_gsc_audit(domain, email, fmt="james", out_dir=None, driver=None,
                  period_days=28, end_offset=3, log_fn=None):
    """Run a complete GSC audit: fetch data, capture screenshots, build PPTX."""
    if log_fn is None:
        log_fn = print

    if out_dir is None:
        out_dir = tempfile.mkdtemp(prefix="gsc_audit_")
    os.makedirs(out_dir, exist_ok=True)

    token = get_access_token(email)
    log_fn(f"  Account: {email}")

    # Resolve property
    log_fn("  Resolving GSC property...")
    property_url = resolve_property(token, domain)
    log_fn(f"  Property: {property_url}")

    # Date range
    end_date = (datetime.now() - timedelta(days=end_offset)).strftime("%Y-%m-%d")
    start_date = (datetime.now() - timedelta(days=end_offset + period_days - 1)).strftime("%Y-%m-%d")
    log_fn(f"  Period: {start_date} to {end_date}")

    report_data = {
        "domain": domain,
        "propertyUrl": property_url,
        "startDate": start_date,
        "endDate": end_date,
        "periodDays": period_days,
        "generatedDate": datetime.now().strftime("%B %d, %Y"),
    }

    # Fetch API data for James Full and Sigma
    if fmt in ("james", "sigma"):
        log_fn("  Fetching GSC API data...")
        api_data = fetch_all_api_data(token, property_url, start_date, end_date, log_fn)
        report_data.update(api_data)

        if fmt == "james":
            log_fn("  Running URL inspections...")
            inspections = run_inspections(token, property_url,
                                          api_data.get("topPages", []), log_fn)
            report_data["inspections"] = inspections

    # Capture screenshots via Selenium
    screenshots = {}
    if driver:
        log_fn("  Capturing GSC screenshots...")
        ss_dir = os.path.join(out_dir, "gsc_screenshots")
        screenshots = capture_gsc_screenshots(driver, property_url, email, ss_dir, log_fn=log_fn)
        log_fn(f"  {len(screenshots)} screenshot(s) captured.")
    report_data["screenshots"] = screenshots

    # Build PPTX
    format_info = GSC_FORMATS.get(fmt, GSC_FORMATS["james"])
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    out_file = os.path.join(out_dir, f"GSC_Audit_{domain}_{fmt}_{timestamp}.pptx")

    log_fn(f"  Building {format_info['label']} report...")
    format_info["builder"](report_data, out_file, log_fn)

    return out_file
