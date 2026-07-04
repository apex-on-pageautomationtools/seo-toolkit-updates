"""
generate_seo_onpage_phase2.py — On-Page SEO Phase 2 report engine for Report Studio.

Given a domain + a "target pages & keywords" list, it crawls each target page, audits
the Phase-2 on-page layer, and produces the same deliverables the manual skill makes:

    1. On Page-Analysis-Report - <domain>.docx     (branded cover page + per-site narrative + screenshots)
    2. Meta Suggestions - <domain>.xlsx            (existing + suggested title/desc/H1; suggested via free Gemini or heuristic)
    3. Image Alt Tag Suggestions - <domain>.xlsx   (self-hosted + Shopify-CDN images only; alt suggestions)
    4. Canonical Tag Suggestions - <domain>.xlsx   (existing vs recommended; self-canonical -> "No Changes Needed…")
    5. Target Pages and Keywords - <domain>.xlsx   (the complete target-pages + keywords list)
    6. sitemap (existing - review).xml             (the site's REAL sitemap if found) OR
       sitemap.xml                                 (a freshly-generated working sitemap when none exists)

…all bundled into a single ZIP: "SEO On-Page Phase 2 - <domain>.zip".

The .xlsx are produced by CLONING the templates in backend/ (template_*.xlsx) so the
formatting matches the house style exactly. The .docx is built from scratch (dynamic
narrative + live screenshots). Page 1 is the branded cover (backend/onpage_cover.png,
A4 full-bleed; per-domain override seo_onpage_screens/<domain>/cover.png; text cover if
absent). Suggested meta copy uses Google Gemini's FREE tier when GEMINI_API_KEY is set,
else a heuristic — no paid APIs.

Run:
    python generate_seo_onpage_phase2.py --domain example.com --targets targets.json
    python generate_seo_onpage_phase2.py --domain example.com --targets targets.json --dry-run

`targets.json` is either:
    [{"page": "https://example.com/", "keywords": ["kw one", "kw two"]}, ...]
  or the flat rows form:
    [{"keyword": "kw one", "page": "https://example.com/"}, ...]
"""
import os
import re
import sys
import json
import shutil
import zipfile
import tempfile
import argparse
import datetime
import urllib.parse
from pathlib import Path

ROOT = Path(__file__).parent.resolve()
# The bundled python runs ISOLATED (a python312._pth file is present), so it does
# NOT put this script's own folder on sys.path. Without this, sibling imports such
# as `import generate_health_report` (used by the docx builder) raise
# ModuleNotFoundError and no on-page report is ever produced on user machines.
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
OUTPUT_DIR = ROOT / "output"
TPL_META = ROOT / "template_meta_suggestions.xlsx"
TPL_ALT = ROOT / "template_alt_suggestions.xlsx"
TPL_CANON = ROOT / "template_canonical_suggestions.xlsx"
# Branded cover image for page 1 (drop-in). Per-domain override:
# seo_onpage_screens/<domain>/cover.png ; else this shared file.
COVER_IMG = ROOT / "onpage_cover.png"

MISSING = "Missing!"


def log(msg):
    print(msg, flush=True)


# --------------------------------------------------------------------------- io
def safe_domain(domain):
    d = (domain or "").strip().lower()
    d = re.sub(r"^https?://", "", d).rstrip("/")
    d = re.sub(r"^www\.", "", d)
    return d.split("/")[0]


def site_root(domain):
    return f"https://{safe_domain(domain)}"


def brand_from(domain, homepage_title=None, homepage_h1=None, og_site=None):
    """Best-effort brand name. Prefer og:site_name (most reliable), then the part of
    the title/H1 before a separator; else title-case the domain's second-level label."""
    if og_site and 2 <= len(og_site.strip()) <= 40:
        return og_site.strip()
    for cand in (homepage_title, homepage_h1):
        if cand and cand != MISSING:
            # take the part after the last separator (brand usually trails the title)
            parts = [p.strip() for p in re.split(r"[|\-–—:·]", cand) if p.strip()]
            if parts:
                tail = parts[-1]
                # a short tail is usually the brand; otherwise use the head
                pick = tail if 2 <= len(tail) <= 30 else parts[0]
                if 2 <= len(pick) <= 40:
                    return pick
    root = safe_domain(domain).split(".")[0]
    return root.replace("-", " ").replace("_", " ").title()


def load_targets(targets_path, domain):
    """Return ordered list of {page, keywords:[...]} grouped by page (primary first)."""
    rows = []
    p = Path(targets_path)
    if p.suffix.lower() in (".xlsx", ".xlsm"):
        import openpyxl
        wb = openpyxl.load_workbook(p, data_only=True)
        ws = wb.active
        # find the header row that has "Keyword" + "Target Page"
        header_row = None
        for r in range(1, min(ws.max_row, 30) + 1):
            vals = [str(ws.cell(r, c).value or "").lower() for c in range(1, ws.max_column + 1)]
            if any("keyword" in v for v in vals) and any("page" in v for v in vals):
                header_row = r
                kw_col = next(c for c, v in enumerate(vals, 1) if "keyword" in v)
                pg_col = next(c for c, v in enumerate(vals, 1) if "page" in v)
                break
        if header_row:
            for r in range(header_row + 1, ws.max_row + 1):
                kw = ws.cell(r, kw_col).value
                pg = ws.cell(r, pg_col).value
                if kw and pg:
                    rows.append({"keyword": str(kw).strip(), "page": str(pg).strip()})
    else:
        data = json.loads(p.read_text(encoding="utf-8"))
        if data and isinstance(data[0], dict) and "keywords" in data[0]:
            # already grouped
            for item in data:
                rows.extend({"keyword": k, "page": item["page"]} for k in item["keywords"])
        else:
            rows = [{"keyword": str(d.get("keyword", "")).strip(), "page": str(d.get("page", "")).strip()} for d in data]

    # group, preserving first-seen page order and keyword order
    grouped = {}
    for row in rows:
        page = normalize_url(row["page"], domain)
        if not page:
            continue
        grouped.setdefault(page, [])
        if row["keyword"]:
            grouped[page].append(row["keyword"])
    return [{"page": pg, "keywords": kws} for pg, kws in grouped.items()]


def discover_targets(domain, dry_run=False, limit=12):
    """When no keyword sheet is supplied (e.g. Bulk runs), auto-pick target pages
    by crawling the homepage and collecting its same-site internal links."""
    root = site_root(domain)
    if dry_run:
        return [{"page": root + "/", "keywords": []}]
    home = crawl_page(root + "/", [], dry_run=False)
    site_reg = _registrable(safe_domain(domain))
    pages = [root + "/"]
    for u in home.get("internal_links", []):
        clean = u.split("#")[0].split("?")[0].rstrip("/")
        if not clean or clean in pages:
            continue
        if _registrable(urllib.parse.urlparse(clean).netloc) != site_reg:
            continue
        if re.search(r"\.(png|jpe?g|gif|svg|webp|pdf|zip|css|js|ico|xml|json)$", clean, re.I):
            continue
        pages.append(clean)
        if len(pages) >= limit:
            break
    log(f"   -> auto-discovered {len(pages)} target page(s)")
    return [{"page": p, "keywords": []} for p in dict.fromkeys(pages)]


def normalize_url(url, domain):
    url = (url or "").strip()
    if not url:
        return ""
    if not url.startswith("http"):
        url = site_root(domain) + "/" + url.lstrip("/")
    return url.rstrip()


def _registrable(host):
    """Last two labels of a host, e.g. www.x.com / api.x.com -> x.com."""
    host = (host or "").lower().split(":")[0]
    parts = [p for p in host.split(".") if p]
    return ".".join(parts[-2:]) if len(parts) >= 2 else host


_SHOPIFY_CDN = re.compile(r"(?:^|\.)(?:shopify\.com|myshopify\.com|shopifycdn\.com)$")


def _allowed_image_host(src, site_host):
    """Keep only self-hosted images (same registrable domain as the site) and
    Shopify-CDN images. Drops trackers / fonts / analytics pixels / third-party
    CDNs (facebook.com/tr, fonts.gstatic.com, *.blob.core.windows.net, …)."""
    host = urllib.parse.urlparse(src).netloc.lower()
    if not host:
        return True  # same-origin relative URL (already resolved against the page)
    if _registrable(host) == _registrable(site_host):
        return True
    return bool(_SHOPIFY_CDN.search(host))


# -------------------------------------------------------------------- crawling
def _mock_page(url, keywords):
    kw = keywords[0] if keywords else "page"
    return {
        "url": url,
        "title": MISSING,
        "description": MISSING,
        "h1": MISSING,
        "canonical": f'<link rel="canonical" href="{url}" />',
        "lang": "en",
        "viewport": True,
        "images": [
            {"src": f"{url.rstrip('/')}/assets/hero.png", "alt": ""},
            {"src": f"{url.rstrip('/')}/assets/logo.png", "alt": ""},
        ],
        "headings": [("h1", kw.title())],
        "internal_links": [],
        "external_links": [],
        "status": 200,
    }


# --- selenium-rendered crawl (fallback when patchright isn't bundled) ----------
# The embedded python ships selenium + Chrome but NOT patchright/playwright, so on
# user machines the patchright path below never runs. Without this fallback every
# page would be fetched as raw HTML — JS-rendered / SPA sites then report a missing
# H1, zero images, zero headings and zero internal links. Rendering with Chrome
# fixes that. One driver is reused for the whole run.
_op_driver = None


def _get_op_driver():
    global _op_driver
    if _op_driver is not None:
        try:
            _ = _op_driver.title
            return _op_driver
        except Exception:
            _op_driver = None
    try:
        import tempfile
        from selenium.webdriver import Chrome, ChromeOptions
        opts = ChromeOptions()
        opts.add_argument(f"--user-data-dir={tempfile.mkdtemp(prefix='seo_op_')}")
        opts.add_argument("--headless=new")
        opts.add_argument("--no-sandbox")
        opts.add_argument("--disable-dev-shm-usage")
        opts.add_argument("--disable-gpu")
        opts.add_argument("--disable-notifications")
        opts.add_argument("--mute-audio")
        opts.add_argument("--window-size=1366,900")
        opts.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                          "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")
        _op_driver = Chrome(options=opts)
        _op_driver.set_page_load_timeout(45)
        return _op_driver
    except Exception as e:
        log(f"   [warn] could not launch Chrome for rendering: {type(e).__name__}: {e}")
        return None


def _close_op_driver():
    global _op_driver
    if _op_driver:
        try:
            _op_driver.quit()
        except Exception:
            pass
        _op_driver = None


def _op_http_status(url):
    """Best-effort HTTP status (selenium doesn't expose it)."""
    import urllib.request, urllib.error
    for method in ("HEAD", "GET"):
        try:
            req = urllib.request.Request(url, method=method,
                                         headers={"User-Agent": "Mozilla/5.0 SEOPhase2Bot"})
            with urllib.request.urlopen(req, timeout=15) as r:
                return getattr(r, "status", 200) or 200
        except urllib.error.HTTPError as e:
            return e.code
        except Exception:
            continue
    return None


def _crawl_selenium(url, keywords):
    """Render the page in Chrome (waits for the SPA to hydrate) and parse it.
    Returns None if the browser is unavailable so the caller can fall back."""
    driver = _get_op_driver()
    if not driver:
        return None
    import time
    try:
        driver.get(url)
        last, stable = -1, 0
        for _ in range(22):                       # wait up to ~9s for hydration
            try:
                ready = driver.execute_script("return document.readyState")
                blen = driver.execute_script(
                    "return (document.body ? document.body.innerText.length : 0)")
            except Exception:
                break
            if ready == "complete" and blen > 0 and blen == last:
                stable += 1
                if stable >= 2:                   # body text stable ~0.8s -> settled
                    break
            else:
                stable = 0
            last = blen
            time.sleep(0.4)
        html = driver.page_source or ""
        final_url = driver.current_url or url
    except Exception as e:
        log(f"   [warn] selenium crawl failed for {url}: {type(e).__name__}: {e}")
        return None
    if not html or len(html) < 200:
        return None
    return _parse_html(html, final_url, _op_http_status(url))


def _crawl_rendered(url, keywords):
    """Prefer a real browser render (selenium+Chrome), fall back to raw HTTP."""
    sel = _crawl_selenium(url, keywords)
    if sel is not None:
        return sel
    return _crawl_requests(url, keywords)


def crawl_page(url, keywords, dry_run=False):
    if dry_run:
        return _mock_page(url, keywords)
    try:
        from patchright.sync_api import sync_playwright
    except Exception:
        return _crawl_rendered(url, keywords)      # no patchright -> selenium/raw
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            ctx = browser.new_context(user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"))
            page = ctx.new_page()
            resp = page.goto(url, wait_until="networkidle", timeout=45000)
            status = resp.status if resp else None
            html = page.content()
            final_url = page.url
            browser.close()
        return _parse_html(html, final_url or url, status)
    except Exception as e:
        log(f"   [warn] browser crawl failed for {url}: {type(e).__name__}: {e}")
        return _crawl_rendered(url, keywords)      # patchright failed -> selenium/raw


def _crawl_requests(url, keywords):
    try:
        import urllib.request
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 SEOPhase2Bot"})
        with urllib.request.urlopen(req, timeout=30) as r:
            html = r.read().decode("utf-8", "ignore")
            return _parse_html(html, url, r.status)
    except Exception as e:
        log(f"   [warn] http crawl failed for {url}: {e}")
        return _mock_page(url, keywords)


def _parse_html(html, url, status):
    try:
        from bs4 import BeautifulSoup
        soup = BeautifulSoup(html, "html.parser")
    except Exception:
        return _regex_parse(html, url, status)

    def meta(name=None, prop=None):
        if name:
            t = soup.find("meta", attrs={"name": name})
        else:
            t = soup.find("meta", attrs={"property": prop})
        return (t.get("content") or "").strip() if t and t.get("content") else ""

    import html as html_mod
    raw_title = (soup.title.get_text(strip=True) if soup.title else "") or MISSING
    title = html_mod.unescape(raw_title)
    desc = html_mod.unescape(meta(name="description") or MISSING)
    og_site = html_mod.unescape(meta(prop="og:site_name"))
    h1s = [html_mod.unescape(h.get_text(strip=True)) for h in soup.find_all("h1")]
    h1 = h1s[0] if h1s else MISSING
    canon_tag = soup.find("link", attrs={"rel": "canonical"})
    canonical = (f'<link rel="canonical" href="{canon_tag.get("href")}" />'
                 if canon_tag and canon_tag.get("href") else "")
    html_tag = soup.find("html")
    lang = (html_tag.get("lang") if html_tag else "") or ""
    viewport = bool(soup.find("meta", attrs={"name": "viewport"}))
    site_host = urllib.parse.urlparse(url).netloc
    images = []
    base = url
    for im in soup.find_all("img"):
        src = im.get("src") or im.get("data-src") or ""
        if not src or src.startswith("data:"):
            continue
        src = urllib.parse.urljoin(base, src)
        # Only the site's own images + Shopify CDN — drop trackers, fonts,
        # analytics pixels and third-party CDNs.
        if _allowed_image_host(src, site_host):
            images.append({"src": src, "alt": (im.get("alt") or "").strip()})
    headings = [(h.name, h.get_text(strip=True)) for h in soup.find_all(re.compile(r"^h[1-6]$"))]
    site_reg = _registrable(urllib.parse.urlparse(url).netloc)
    internal, external = [], []
    for a in soup.find_all("a", href=True):
        href = (a["href"] or "").strip()
        # skip mailto:, tel:, javascript:, fragments and data URIs — not links
        if not href or href.startswith(("mailto:", "tel:", "javascript:", "#", "data:")):
            continue
        full = urllib.parse.urljoin(base, href)
        if not full.startswith("http"):
            continue
        full = full.split("#")[0]  # ignore in-page anchors
        if _registrable(urllib.parse.urlparse(full).netloc) == site_reg:
            internal.append(full)        # www / non-www both count as internal
        else:
            external.append(full)
    return {
        "url": url, "title": title, "description": desc, "h1": h1, "h1s": h1s,
        "og_site_name": og_site, "canonical": canonical, "lang": lang, "viewport": viewport,
        "images": images, "headings": headings, "internal_links": list(dict.fromkeys(internal)),
        "external_links": list(dict.fromkeys(external)), "status": status,
    }


def _regex_parse(html, url, status):
    def find(pat):
        m = re.search(pat, html, re.I | re.S)
        return m.group(1).strip() if m else ""
    title = find(r"<title[^>]*>(.*?)</title>") or MISSING
    desc = find(r'<meta[^>]+name=["\']description["\'][^>]+content=["\'](.*?)["\']') or MISSING
    h1 = find(r"<h1[^>]*>(.*?)</h1>") or MISSING
    canon = find(r'<link[^>]+rel=["\']canonical["\'][^>]+href=["\'](.*?)["\']')
    canonical = f'<link rel="canonical" href="{canon}" />' if canon else ""
    lang = find(r'<html[^>]+lang=["\'](.*?)["\']')
    return {"url": url, "title": re.sub("<[^>]+>", "", title), "description": desc,
            "h1": re.sub("<[^>]+>", "", h1), "h1s": [], "canonical": canonical, "lang": lang,
            "viewport": "viewport" in html.lower(), "images": [], "headings": [],
            "internal_links": [], "external_links": [], "status": status}


# ----------------------------------------------------------------- deliverables
def _ai_suggest(prompt):
    """Suggestion copy via Google Gemini's FREE tier when GEMINI_API_KEY is set
    (free key: https://aistudio.google.com/apikey). Plain REST, no SDK, no paid API.
    Returns parsed JSON or None (→ heuristic fallback)."""
    key = os.environ.get("GEMINI_API_KEY", "").strip()
    if not key:
        return None
    try:
        import urllib.request
        model = os.environ.get("GEMINI_MODEL", "gemini-1.5-flash")
        url = (f"https://generativelanguage.googleapis.com/v1beta/models/"
               f"{model}:generateContent?key={key}")
        body = json.dumps({
            "contents": [{"parts": [{"text": prompt}]}],
            "generationConfig": {"temperature": 0.7, "responseMimeType": "application/json"},
        }).encode("utf-8")
        req = urllib.request.Request(url, data=body, headers={"Content-Type": "application/json"})
        with urllib.request.urlopen(req, timeout=40) as r:
            data = json.loads(r.read())
        text = data["candidates"][0]["content"]["parts"][0]["text"]
        return json.loads(re.search(r"\{.*\}|\[.*\]", text, re.S).group(0))
    except Exception as e:
        log(f"   [warn] Gemini meta suggestion failed, using heuristic: {e}")
        return None


def suggest_meta(page_data, keywords, brand):
    """Existing title/desc/H1 from the crawl. Suggested columns are left empty —
    the user checks ranking/traffic manually and fills them in as needed."""
    existing_title = page_data["title"]
    existing_desc = page_data["description"]
    existing_h1 = page_data["h1"]

    check_msg = "Check manually as per ranking and traffic"
    return {
        "page": page_data["url"],
        "existing_title": existing_title,
        "suggested_title": check_msg,
        "existing_description": existing_desc,
        "suggested_description": check_msg,
        "existing_h1": existing_h1,
        "suggested_h1": check_msg,
    }


# Tracking pixels, font icons, analytics — skip entirely (not real images)
_SKIP_PATTERNS = re.compile(
    r"facebook\.com/tr\?|"
    r"fonts\.gstatic\.com|"
    r"fonts\.googleapis\.com|"
    r"google-analytics\.com|"
    r"googletagmanager\.com|"
    r"doubleclick\.net|"
    r"analytics\.|"
    r"pixel\.|"
    r"\.gif\?|"   # tracking pixels are usually tiny GIFs with query strings
    r"data:image/",  # inline base64 images
    re.I
)

# Self-hosted / easily updatable from admin (WordPress, Shopify, common CMS)
_SELF_HOSTED_PATTERNS = re.compile(
    r"/wp-content/|"
    r"/wp-includes/|"
    r"cdn\.shopify\.com|"
    r"/assets/|"
    r"/uploads/|"
    r"/media/|"
    r"/images/|"
    r"/img/|"
    r"/static/",
    re.I
)


def _is_self_hosted(src, domain):
    """Check if image is on the same domain or a known CMS CDN."""
    parsed = urllib.parse.urlparse(src)
    host = parsed.netloc.lower()
    d = safe_domain(domain)
    if not host or d in host:
        return True
    if re.search(r"cdn\.shopify\.com", host, re.I):
        return True
    if _SELF_HOSTED_PATTERNS.search(parsed.path):
        return True
    return False


def suggest_alt(page_data, keywords, brand, domain=""):
    """List all images with their existing alt text. Suggested alt is left empty —
    the user checks the actual image and fills it in manually.
    Only includes self-hosted / easily updatable images.
    External CDN images are listed separately with a dev-check note."""
    self_hosted = []
    external_cdn = []

    for img in page_data.get("images", []):
        src = img.get("src", "")
        if not src or _SKIP_PATTERNS.search(src):
            continue
        existing_alt = img.get("alt", "") or ""
        entry = {"page": page_data["url"], "image": src, "existing_alt": existing_alt, "suggested_alt": ""}

        if _is_self_hosted(src, domain):
            self_hosted.append(entry)
        else:
            entry["suggested_alt"] = "Ask developer if alt text can be updated"
            external_cdn.append(entry)

    return self_hosted, external_cdn


def recommend_canonical(page_data):
    url = page_data["url"]
    existing = page_data.get("canonical") or MISSING
    m = re.search(r'href="([^"]+)"', existing)
    existing_href = m.group(1).strip() if m else None
    # Self-referencing canonical already points at this page → nothing to change.
    if existing_href and existing_href.rstrip("/") == url.rstrip("/"):
        recommended = "No Changes Needed as Self Canonical Found"
    else:
        recommended = f'<link rel="canonical" href="{url}" />'
    return {"page": url, "existing": existing, "recommended": recommended}


# ----------------------------------------------------------------- xlsx output
def _clone_and_clear(template, header_rows=1):
    import openpyxl
    wb = openpyxl.load_workbook(template)
    ws = wb.active
    # remember the style of the first data row to reuse
    style_row = header_rows + 1
    proto = [ws.cell(style_row, c)._style for c in range(1, ws.max_column + 1)]
    # clear everything below the header
    if ws.max_row > header_rows:
        ws.delete_rows(header_rows + 1, ws.max_row - header_rows)
    return wb, ws, proto


def _write_rows(ws, proto, rows, start=2):
    for i, row in enumerate(rows):
        r = start + i
        for c, val in enumerate(row, 1):
            cell = ws.cell(r, c, val)
            if c - 1 < len(proto):
                cell._style = proto[c - 1]


def write_meta_xlsx(metas, out_path):
    from openpyxl.styles import Alignment
    wb, ws, proto = _clone_and_clear(TPL_META)
    rows = [[m["page"], m["existing_title"], m["suggested_title"], m["existing_description"],
             m["suggested_description"], m["existing_h1"], m["suggested_h1"]] for m in metas]
    _write_rows(ws, proto, rows)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            cell.alignment = left
    wb.save(out_path)


def write_alt_xlsx(self_hosted, external_cdn, out_path):
    import openpyxl
    from openpyxl.styles import Font, Alignment
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Image Alt Tag Suggestions"
    headers = ["Page URL", "Image URL", "Existing Alt Text", "Suggested Alt Tag"]
    hfont = Font(bold=True, size=11)
    for c, h in enumerate(headers, 1):
        cell = ws.cell(1, c, h)
        cell.font = hfont
    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 60
    ws.column_dimensions["C"].width = 30
    ws.column_dimensions["D"].width = 30

    row = 2
    for a in self_hosted:
        ws.cell(row, 1, a["page"])
        ws.cell(row, 2, a["image"])
        ws.cell(row, 3, a["existing_alt"])
        ws.cell(row, 4, a["suggested_alt"])
        row += 1

    if external_cdn:
        row += 1
        note_cell = ws.cell(row, 1, "Below images are on external CDN — ask developer if alt text can be updated:")
        note_cell.font = Font(bold=True, color="C00000", size=11)
        row += 1
        for a in external_cdn:
            ws.cell(row, 1, a["page"])
            ws.cell(row, 2, a["image"])
            ws.cell(row, 3, a["existing_alt"])
            ws.cell(row, 4, a["suggested_alt"])
            row += 1

    wb.save(out_path)


def write_canonical_xlsx(canons, out_path):
    wb, ws, proto = _clone_and_clear(TPL_CANON)
    rows = [[c["page"], c["existing"], c["recommended"]] for c in canons]
    _write_rows(ws, proto, rows)
    wb.save(out_path)


def write_targets_xlsx(targets, out_path):
    """The complete list of target pages + their keywords, as its own sheet."""
    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Target Pages"
    header_fill = PatternFill("solid", fgColor="2F5496")
    header_font = Font(bold=True, size=14, color="FFFFFF")
    ws.append(["Keywords", "Target Pages"])
    for c in ws[1]:
        c.font = header_font
        c.fill = header_fill
        c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    for t in targets:
        kws = t.get("keywords") or [""]
        for kw in kws:
            ws.append([kw, t["page"]])
    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 58
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.alignment = left
    wb.save(out_path)


# --------------------------------------------------------------- sitemap output
def find_existing_sitemap(domain, dry_run=False):
    """Look for the site's REAL sitemap (robots.txt 'Sitemap:' line, then common
    locations). Returns (url, body) if found, else (None, None). We never fabricate
    a sitemap from just the target pages — that would miss most of the site."""
    if dry_run:
        return None, None
    root = site_root(domain)
    candidates = []
    _, rtxt, _ = _http(root + "/robots.txt")
    for m in re.finditer(r"(?im)^\s*sitemap:\s*(\S+)", rtxt or ""):
        candidates.append(m.group(1).strip())
    candidates += [root + "/sitemap.xml", root + "/sitemap_index.xml",
                   root + "/sitemap-index.xml", root + "/sitemap1.xml"]
    for u in dict.fromkeys(candidates):
        st, body, _ = _http(u)
        if st == 200 and ("<urlset" in body or "<sitemapindex" in body):
            return u, body
    return None, None


def _xml_escape(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def build_sitemap_xml(urls):
    """A valid, working <urlset> sitemap built from the given URLs (used when the
    site has no sitemap of its own)."""
    seen = list(dict.fromkeys((u.split("#")[0].rstrip("/") or u) for u in urls if u and u.startswith("http")))
    today = datetime.date.today().isoformat()
    items = "\n".join(
        f"  <url>\n    <loc>{_xml_escape(u)}</loc>\n    <lastmod>{today}</lastmod>\n"
        f"    <changefreq>weekly</changefreq>\n    <priority>0.8</priority>\n  </url>"
        for u in seen
    )
    return ('<?xml version="1.0" encoding="UTF-8"?>\n'
            '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n'
            f"{items}\n</urlset>\n")


# ------------------------------------------------------------------ docx output
def _kebab(slug):
    s = re.sub(r"([a-z0-9])([A-Z])", r"\1-\2", slug)
    return re.sub(r"[_\s]+", "-", s).lower()


def _http(url, method="GET", timeout=20):
    import urllib.request
    import urllib.error
    try:
        req = urllib.request.Request(url, method=method,
                                     headers={"User-Agent": "Mozilla/5.0 SEOPhase2Bot"})
        with urllib.request.urlopen(req, timeout=timeout) as r:
            body = r.read().decode("utf-8", "ignore") if method == "GET" else ""
            return r.status, body, r.geturl()
    except urllib.error.HTTPError as e:
        return e.code, "", url
    except Exception:
        return 0, "", url


def check_gsc_indexing(urls, token, property_url):
    """Call GSC URL Inspection API for each target page. Returns list of dicts."""
    import urllib.request
    results = []
    for url in urls:
        try:
            body = json.dumps({
                "inspectionUrl": url,
                "siteUrl": property_url,
                "languageCode": "en-US"
            }).encode()
            req = urllib.request.Request(
                "https://searchconsole.googleapis.com/v1/urlInspection/index:inspect",
                data=body,
                headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
                method="POST"
            )
            resp = urllib.request.urlopen(req, timeout=15)
            data = json.loads(resp.read())
            ir = data.get("inspectionResult", {})
            idx = ir.get("indexStatusResult", {})
            results.append({
                "url": url,
                "verdict": idx.get("verdict", "UNKNOWN"),
                "coverageState": idx.get("coverageState", ""),
                "robotsTxtState": idx.get("robotsTxtState", ""),
                "indexingState": idx.get("indexingState", ""),
                "lastCrawlTime": idx.get("lastCrawlTime", ""),
            })
        except Exception as e:
            results.append({"url": url, "verdict": "ERROR", "coverageState": str(e)})
    return results


def audit_site(domain, pages_data, dry_run=False):
    """Detect the real per-site findings that drive the narrative report."""
    d = safe_domain(domain)
    root = site_root(domain)
    # use the actual homepage (root path) for homepage-scoped checks, not just page[0]
    home = next((pd for pd in pages_data if urllib.parse.urlparse(pd["url"]).path in ("", "/")),
                pages_data[0] if pages_data else {})
    f = {"target_pages": [pd["url"] for pd in pages_data]}

    home_html = ""
    if not dry_run:
        _, home_html, _ = _http(root + "/")

    # robots.txt
    if dry_run:
        f["robots_found"], f["robots_has_sitemap"] = False, False
    else:
        rst, rtxt, _ = _http(root + "/robots.txt")
        f["robots_found"] = rst == 200 and ("user-agent" in rtxt.lower() or "disallow" in rtxt.lower())
        f["robots_has_sitemap"] = "sitemap" in rtxt.lower()
        f["robots_body"] = rtxt if rst == 200 else ""

    # sitemap_found / sitemap_url are set in main() via find_existing_sitemap()

    # www vs non-www (issue if www serves 200 without redirecting to non-www)
    if dry_run:
        f["www_redirect_issue"] = True
    else:
        wst, _, wfinal = _http(f"https://www.{d}/")
        f["www_redirect_issue"] = wst == 200 and "www." in urllib.parse.urlparse(wfinal).netloc

    # custom 404 (a clearly-missing URL should return 404, not 200/redirect)
    f["has_custom_404"] = (not dry_run) and _http(root + "/sos-404-probe-zzz")[0] == 404

    text_all = (home_html or "").lower()
    links = " ".join(home.get("internal_links", [])).lower()
    # a "blog" can also be a News / Articles / Insights / Media section
    f["has_blog"] = bool(re.search(r"/(blog|news|article|insight|press|media-?news)", links)) \
        or bool(re.search(r"\b(blog|latest news|our news)\b", text_all[:8000]))
    f["has_faq"] = "/faq" in links or "faq" in text_all[:10000]

    socials = [u for u in home.get("external_links", [])
               if re.search(r"facebook|instagram|twitter|linkedin|x\.com|youtube|t\.me|tiktok", u, re.I)]
    f["social_found"] = bool(socials)

    cur_year = datetime.date.today().year
    # take the MAX year near a copyright mark (sites often show "© 2026 … est. 2022")
    yrs = [int(y) for y in re.findall(r"(?:©|&copy;|copyright)[^0-9]{0,25}(20\d\d)", home_html, re.I)]
    f["copyright_year"] = max(yrs) if yrs else None
    f["copyright_stale"] = bool(f["copyright_year"] and f["copyright_year"] < cur_year)

    imgs = [im for pd in pages_data for im in pd.get("images", [])
            if not _SKIP_PATTERNS.search(im.get("src", ""))]
    f["alt_total"] = len(imgs)
    f["alt_missing"] = sum(1 for im in imgs if not im.get("alt"))

    # footer logo: check if <footer> contains an <img> tag
    f["has_footer_logo"] = bool(re.search(r"<footer[^>]*>[\s\S]*?<img\b", home_html[:50000], re.I))

    f["ext_count"] = len(home.get("external_links", []))
    f["lang"] = home.get("lang", "") or ""
    f["viewport"] = any(pd.get("viewport") for pd in pages_data)

    m = re.search(r'href="([^"]+)"', home.get("canonical", "") or "")
    f["home_canonical"] = m.group(1) if m else None
    f["canonical_issue"] = bool(f["home_canonical"] and f["home_canonical"].rstrip("/") != root)

    # noindex check — look for meta robots noindex in each target page's HTML
    noindex_pages = []
    for pd in pages_data:
        page_url = pd["url"]
        if dry_run:
            continue
        try:
            _, phtml, _ = _http(page_url)
            if re.search(r'<meta[^>]+name=["\']robots["\'][^>]+content=["\'][^"\']*noindex', phtml, re.I):
                noindex_pages.append(page_url)
            elif re.search(r'<meta[^>]+content=["\'][^"\']*noindex[^"\']*["\'][^>]+name=["\']robots["\']', phtml, re.I):
                noindex_pages.append(page_url)
        except Exception:
            pass
    f["noindex_pages"] = noindex_pages

    # mixed content — check if HTTPS pages load HTTP resources
    mixed_content_pages = []
    for pd in pages_data:
        page_url = pd["url"]
        if dry_run or not page_url.startswith("https"):
            continue
        try:
            _, phtml, _ = _http(page_url)
            http_resources = re.findall(r'(?:src|href|action)=["\']http://[^"\']+["\']', phtml, re.I)
            if http_resources:
                mixed_content_pages.append(page_url)
        except Exception:
            pass
    f["mixed_content_pages"] = mixed_content_pages

    # sucuri site check — query the free Sucuri SiteCheck API
    f["sucuri_clean"] = None  # None = couldn't check, True = clean, False = issues
    if not dry_run:
        try:
            suc_url = f"https://sitecheck.sucuri.net/api/v3/?scan={root}"
            _, suc_body, _ = _http(suc_url)
            suc_data = json.loads(suc_body) if suc_body else {}
            warnings = suc_data.get("warnings", {})
            blacklisted = suc_data.get("blacklists", {})
            has_malware = bool(suc_data.get("malware", []))
            has_blacklist = any(v for v in blacklisted.values() if v)
            f["sucuri_clean"] = not has_malware and not has_blacklist
        except Exception:
            f["sucuri_clean"] = None

    # broken links — check each external link for 404/5xx
    broken_links = []
    if not dry_run:
        ext_links = list(dict.fromkeys(home.get("external_links", [])[:30]))  # cap at 30
        for link in ext_links:
            try:
                lst, _, _ = _http(link)
                if lst and lst >= 400:
                    broken_links.append(link)
            except Exception:
                pass
    f["broken_links"] = broken_links

    f["url_changes"] = []
    for pd in pages_data:
        path = urllib.parse.urlparse(pd["url"]).path.strip("/")
        if path and re.search(r"[A-Z]", path):
            f["url_changes"].append((f"{root}/{path}", f"{root}/{_kebab(path)}"))
    return f


def capture_onpage_screenshots(domain, sitemap_url=None):
    """Live, HEADLESS screenshots for the report — all public pages/tools, so no
    Google login is needed. Captured with selenium + Chrome via CDP: the embedded
    python ships selenium but NOT patchright/playwright, so the old patchright
    capture path silently produced NO screenshots on user machines. The sitemap
    shot is taken of the ACTUAL found sitemap (only if one exists)."""
    import base64, tempfile
    import time as _t

    driver = _get_op_driver()
    if not driver:
        log("   [warn] no browser available — screenshots skipped")
        return {}

    out = {}
    out_dir = tempfile.mkdtemp(prefix=f"onpage_shots_{safe_domain(domain)}_")
    root = f"https://{domain}"

    def path(k):
        return os.path.join(out_dir, f"{domain}_seo_{k}.png")

    def _shot(key, url, height=900, view_source=False, sucuri=False):
        p = path(key)
        try:
            driver.get(("view-source:" + url) if view_source else url)
            _t.sleep(15 if sucuri else 4)          # let the page (or Sucuri scan) settle
            # Google serves headless browsers a reCAPTCHA / "unusual traffic" wall
            # instead of results — a screenshot of that is useless, so skip it (the
            # indexing section falls back to a text note / GSC data).
            if key == "serp":
                src = (driver.page_source or "").lower()
                if any(m in src for m in ("unusual traffic", "not a robot", "recaptcha",
                                          "/sorry/", "detected unusual", "before you continue")):
                    log("   [warn] Google SERP blocked (captcha) — skipping serp screenshot")
                    return
            try:
                driver.execute_script(
                    "document.querySelectorAll('.cookie-banner,.consent-banner,"
                    "[class*=cookie],[class*=consent],#cookie-law-info-bar')"
                    ".forEach(e=>e.remove()); window.scrollTo(0,0);")
            except Exception:
                pass
            _t.sleep(0.6)
            if view_source:
                cdp = {"format": "png", "captureBeyondViewport": False}
            else:
                try:
                    _w = driver.execute_script(
                        "return Math.max(document.documentElement.clientWidth||0,"
                        " window.innerWidth||0, 1366);")
                except Exception:
                    _w = 1366
                # Sucuri's verdict sits at the very top — clip a fixed top region of
                # the document so we never grab a scrolled 'check another URL' view.
                h = 1300.0 if sucuri else float(height)
                cdp = {"format": "png", "captureBeyondViewport": True,
                       "clip": {"x": 0, "y": 0, "width": float(_w or 1366), "height": h, "scale": 1}}
            result = driver.execute_cdp_cmd("Page.captureScreenshot", cdp)
            with open(p, "wb") as f:
                f.write(base64.b64decode(result["data"]))
            out[key] = p
            log(f"   -> captured [{key}]")
        except Exception as e:
            log(f"   [warn] capture {key} failed: {type(e).__name__}: {e}")
            try:                                   # last resort: plain viewport shot
                driver.save_screenshot(p)
                out[key] = p
                log(f"   -> captured [{key}] (fallback)")
            except Exception:
                pass

    _shot("homepage",  root + "/", height=900)
    _shot("robots",    root + "/robots.txt", height=700)
    _shot("canonical", root + "/", view_source=True)
    if sitemap_url:                                # only shoot a sitemap that exists
        _shot("sitemap", sitemap_url, height=760)
    _shot("serp",      f"https://www.google.com/search?q=site:{domain}", height=900)
    _shot("wayback",   f"https://web.archive.org/web/2/https://{domain}/", height=900)
    _shot("viewport",  root + "/", height=900)
    _shot("sucuri",    f"https://sitecheck.sucuri.net/results/https/{domain}", sucuri=True)

    # Broken-links image — reuse health_audit's pure (non-patchright) helpers.
    try:
        import sys as _sys
        _repo = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        if _repo not in _sys.path:
            _sys.path.insert(0, _repo)
        import health_audit as _ha
        checked, broken = _ha.check_broken_links(domain)
        _ha._render_broken_links_image(domain, checked, broken, path("brokenlinks"))
        out["brokenlinks"] = path("brokenlinks")
        out["_broken"] = len(broken) if hasattr(broken, "__len__") else 0
        log(f"   -> broken-link check: {checked} links, {out['_broken']} broken")
    except Exception as e:
        log(f"   [warn] broken-link check failed: {type(e).__name__}: {e}")

    _close_op_driver()
    return out


def _cover_source(domain):
    """Per-domain cover override, else the shared cover, else None (text fallback)."""
    dropin = Path("seo_onpage_screens") / safe_domain(domain) / "cover.png"
    if dropin.exists():
        return dropin
    return COVER_IMG if COVER_IMG.exists() else None


def _img_dims(path):
    from PIL import Image
    with Image.open(path) as im:
        return im.size


INTRO_TEXT = (
    "On-page optimization refers to all measures that can be taken directly within the website "
    "in order to improve its position in the search rankings. Optimize the content or improve "
    "the Meta tags."
)


def _compose_cover(cover_path, domain):
    """Bake the site URL + intro text onto the cover's white bottom band, returning
    a temp PNG path. Matches the supplied 'ON PAGE SEO REPORT' cover design."""
    from PIL import Image, ImageDraw, ImageFont
    img = Image.open(cover_path).convert("RGB")
    W, H = img.size
    draw = ImageDraw.Draw(img)

    def font(px, bold=False):
        for name in (("arialbd.ttf",) if bold else ("arial.ttf",)):
            try:
                return ImageFont.truetype("C:/Windows/Fonts/" + name, px)
            except Exception:
                pass
        return ImageFont.load_default()

    url = site_root(domain) + "/"
    title_f = font(int(W * 0.030), bold=True)
    body_f = font(int(W * 0.0175))
    maxw = int(W * 0.80)

    def wrap(text, f):
        lines, cur = [], ""
        for w in text.split():
            t = (cur + " " + w).strip()
            if draw.textlength(t, font=f) <= maxw:
                cur = t
            else:
                lines.append(cur)
                cur = w
        if cur:
            lines.append(cur)
        return lines

    y = int(H * 0.785)  # white band near the bottom of the cover
    uw = draw.textlength(url, font=title_f)
    draw.text(((W - uw) / 2, y), url, fill=(17, 17, 17), font=title_f)
    y += int(title_f.size * 1.7)
    for ln in wrap(INTRO_TEXT, body_f):
        lw = draw.textlength(ln, font=body_f)
        draw.text(((W - lw) / 2, y), ln, fill=(45, 45, 45), font=body_f)
        y += int(body_f.size * 1.5)

    # normalise to A4 portrait ratio (1:1.414) so it fills the page edge-to-edge
    img = img.resize((1240, 1754))
    fd, tmp = tempfile.mkstemp(suffix=".png", prefix="seocover_")
    os.close(fd)
    img.save(tmp)
    return tmp


def _text_cover(doc, domain):
    """Plain styled cover used when no cover image is available."""
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Pt
    for _ in range(7):
        doc.add_paragraph()
    t = doc.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = t.add_run("ON-PAGE SEO REPORT")
    r.bold = True
    r.font.size = Pt(34)
    u = doc.add_paragraph()
    u.alignment = WD_ALIGN_PARAGRAPH.CENTER
    ru = u.add_run(site_root(domain) + "/")
    ru.bold = True
    ru.font.size = Pt(16)
    i = doc.add_paragraph()
    i.alignment = WD_ALIGN_PARAGRAPH.CENTER
    ri = i.add_run(INTRO_TEXT)
    ri.font.size = Pt(11)


def _setup_docx(domain):
    """Create a Document with cover page and return (doc, helpers_dict)."""
    import generate_health_report as hr
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.section import WD_SECTION
    from docx.shared import Inches, Pt, Cm, RGBColor

    A4_W, A4_H = Cm(21), Cm(29.7)
    FONT, SIZE = "Calibri", 11

    d = safe_domain(domain)
    root = site_root(domain)

    doc = Document()
    for sname in ("Normal", "List Paragraph"):
        try:
            st = doc.styles[sname]
            st.font.name = FONT
            st.font.size = Pt(SIZE)
        except KeyError:
            pass

    sec = doc.sections[0]
    sec.page_width, sec.page_height = A4_W, A4_H
    cover = _cover_source(domain)
    if cover:
        sec.left_margin = sec.right_margin = sec.top_margin = sec.bottom_margin = Cm(0)
        try:
            composed = _compose_cover(cover, domain)
            cp = doc.add_paragraph()
            cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cp.paragraph_format.space_before = Pt(0)
            cp.paragraph_format.space_after = Pt(0)
            cp.paragraph_format.line_spacing = 1.0
            cp.add_run().add_picture(composed, width=A4_W)
            try:
                os.unlink(composed)
            except OSError:
                pass
        except Exception as e:
            log(f"   [warn] cover compose failed: {type(e).__name__}: {e}")
        body = doc.add_section(WD_SECTION.NEW_PAGE)
        body.page_width, body.page_height = A4_W, A4_H
        body.left_margin = body.right_margin = Inches(0.49)
        body.top_margin = body.bottom_margin = Inches(1)
    else:
        log("   [info] no cover image (backend/onpage_cover.png) — using a text cover.")
        sec.left_margin = sec.right_margin = Inches(0.49)
        sec.top_margin = sec.bottom_margin = Inches(1)
        _text_cover(doc, domain)
        doc.add_page_break()

    def _style_run(run, bold=False):
        run.font.name = FONT
        run.font.size = Pt(SIZE)
        run.font.bold = bold

    def label_body(label, body_text=""):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        if label:
            _style_run(p.add_run(label), bold=True)
        if body_text:
            _style_run(p.add_run(body_text), bold=False)
        return p

    def para(text="", bold=False):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        if text:
            _style_run(p.add_run(text), bold=bold)
        return p

    def para_red(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(text)
        r.font.name = FONT
        r.font.size = Pt(SIZE)
        r.font.bold = True
        r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
        return p

    def ribbon(text):
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(6)
        r = p.add_run(text)
        r.font.name = FONT
        r.font.size = Pt(14)
        r.font.bold = True
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        shd = OxmlElement('w:shd')
        shd.set(qn('w:fill'), '000066')
        shd.set(qn('w:val'), 'clear')
        p._p.get_or_add_pPr().append(shd)
        return p

    def result(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(text)
        r.font.name = FONT
        r.font.size = Pt(SIZE)
        r.font.bold = True
        r.font.color.rgb = RGBColor(0x1D, 0x54, 0x89)
        return p

    def green(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(text)
        r.font.name = FONT
        r.font.size = Pt(SIZE)
        r.font.bold = True
        r.font.color.rgb = RGBColor(0x00, 0xB0, 0x50)
        return p

    def shot(key, captured):
        src = (captured or {}).get(key)
        if src and Path(src).exists():
            green("Screenshot:")
            hr._add_bordered_image(doc, src)

    def summary_table(findings):
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement
        tbl = doc.add_table(rows=2, cols=2)
        tbl.style = 'Table Grid'
        hdr = tbl.rows[0]
        for i, txt in enumerate(["SEO Factors", "On Page Analysis Information"]):
            cell = hdr.cells[i]
            cell.text = ""
            p = cell.paragraphs[0]
            r = p.add_run(txt)
            r.font.name = FONT
            r.font.size = Pt(SIZE)
            r.font.bold = True
            r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            shd = OxmlElement('w:shd')
            shd.set(qn('w:fill'), '1D5489')
            shd.set(qn('w:val'), 'clear')
            cell._tc.get_or_add_tcPr().append(shd)
        row1 = tbl.rows[1]
        row1.cells[0].text = ""
        p0 = row1.cells[0].paragraphs[0]
        r0 = p0.add_run("Image ALT Tags")
        r0.font.name = FONT
        r0.font.size = Pt(SIZE)
        r0.font.bold = True
        row1.cells[1].text = ""
        p1 = row1.cells[1].paragraphs[0]
        alt_missing = findings.get("alt_missing", 0)
        if alt_missing > 0:
            status = f"Status: {alt_missing} image(s) without ALT Tags found."
        else:
            status = "Status: ALT Tags are found on images."
        r1 = p1.add_run(status)
        r1.font.name = FONT
        r1.font.size = Pt(SIZE)
        r1.font.color.rgb = RGBColor(0x1D, 0x54, 0x89)
        r1.font.bold = True
        p1.add_run("\n")
        rec = p1.add_run("Recommendation: All the Used images should be optimized using proper "
                         "ALT Attributes. Images should have relevant and useful text as ALT Attributes.")
        rec.font.name = FONT
        rec.font.size = Pt(SIZE)
        rec.font.color.rgb = RGBColor(0x00, 0xB0, 0x50)
        rec.font.bold = True
        para("")

    return doc, {"label_body": label_body, "para": para, "para_red": para_red, "shot": shot,
                 "ribbon": ribbon, "result": result, "green": green, "summary_table": summary_table,
                 "d": d, "root": root, "FONT": FONT, "SIZE": SIZE, "_style_run": _style_run}


# ---- Reusable section writers ----
def _sec_intro(h, root):
    p_url = h["para"].__self__ if hasattr(h["para"], "__self__") else None
    # Just use the helpers directly
    h["para"](root + "/", bold=True)
    h["para"]("On-page optimization refers to all measures that can be taken directly within the "
              "website in order to improve its position in the search ranking. This includes measures "
              "to optimize the content and the source code of a page.")

def _sec_on_page_analysis(h):
    h["result"]("On Page Analysis:")
    h["para"]("These are the most important factors to be checked before optimizing a website. "
              "Proper implementation will help keyword ranking in search engine for your website.")

def _sec_additional_notes(h, findings, captured, year):
    h["label_body"]("Additional Note:")
    h["label_body"]("Note for Content Optimization:", " Kindly find the attached doc for content optimization.")
    h["label_body"]("Note for Meta Suggestion:", " Kindly find the attached Sheet for Meta Suggestion.")
    if not findings.get("has_faq"):
        h["para"]("")
        h["label_body"]("Note for FAQ Page:", " As we analyzed your website, we did not find a FAQ page. "
                        "If possible, we kindly recommend creating a new FAQ page, as it will help with SEO.")
    if not findings.get("social_found"):
        h["para"]("")
        h["label_body"]("Note for Social Media Icons:", " Your website is not connected to social media "
                        "using the API's provided by Facebook, Instagram, Twitter etc. So, we suggest you "
                        "to connect your website to all social media.")
        h["shot"]("homepage", captured)
    if not findings.get("has_footer_logo"):
        h["para"]("")
        h["label_body"]("Note for Footer Logo:", " As we analyzed your website, we noticed that there is no "
                        "logo in the footer section, which can affect brand consistency so we would suggest "
                        "you to add the logo in the footer.")
    if findings.get("copyright_stale"):
        h["para"]("")
        h["label_body"]("Note for Copyright:", f" When we analyzed your website, we noticed that your copyright "
                        f"year is not updated. So, we suggest you to update the copyright year to {year} "
                        f"instead of {findings['copyright_year']}.")
        h["shot"]("homepage", captured)

def _sec_alt_tags(h, findings):
    if findings.get("alt_missing", 0) > 0:
        h["result"](f"Result: {findings['alt_missing']} image(s) without Alt tags found on the target pages "
                    f"checked. It is not good from an SEO point of view. Kindly find the attached Image Alt "
                    f"Tag Suggestions sheet.")
    else:
        h["result"]("Result: Suitable Image Alt tags are found on the target pages checked. It is good from "
                    "an SEO point of view.")

def _sec_robots(h, findings, captured):
    h["para"]("Robots.txt is a regular text file that through its name has special meaning to the majority "
              "of \"honorable\" robots on the web. By defining a few rules in this text file, you can "
              "instruct robots to not crawl & index certain files, directories within a site or the entire "
              "site at all.")
    if findings.get("robots_found"):
        h["result"]("Result: Existing robots.txt file is optimized. It is good from an SEO point of view.")
    else:
        h["result"]("Result: An optimized robots.txt file was created; please find it attached and upload it "
                    "to the root folder of the website.")
    if not findings.get("robots_has_sitemap"):
        h["result"]("Note: Wrong sitemap URL found in the robots.txt. So, we have added a working sitemap URL "
                    "in the robots.txt file.")
    h["shot"]("robots", captured)

def _sec_indexing(h, findings, captured):
    h["ribbon"]("Indexing Status of All Target Pages Check")
    h["para"]("Making your page appear on search engine search results is called indexing. Getting your "
              "webpages indexed by search engine is extremely important. Pages that are not indexed by "
              "Google cannot rank and attract search traffic.")

    gsc_idx = findings.get("gsc_indexing")
    if gsc_idx:
        indexed = [r for r in gsc_idx if r["verdict"] == "PASS"]
        not_indexed = [r for r in gsc_idx if r["verdict"] not in ("PASS", "ERROR")]
        errors = [r for r in gsc_idx if r["verdict"] == "ERROR"]

        h["result"](f"Result: {len(indexed)} of {len(gsc_idx)} target page(s) are indexed by Google.")
        if indexed:
            h["green"]("Indexed Pages:")
            for r in indexed:
                h["para"](f"  {r['url']}  —  Indexed")
        if not_indexed:
            h["para_red"]("Not Indexed / Issues:")
            for r in not_indexed:
                state = r.get("coverageState") or r["verdict"]
                h["para"](f"  {r['url']}  —  {state}")
        if errors:
            h["para"]("Could not check:")
            for r in errors:
                h["para"](f"  {r['url']}  —  {r.get('coverageState', 'API error')}")
    else:
        h["para"](f"{len(findings['target_pages'])} target page(s) found:", bold=True)
        for u in findings["target_pages"]:
            h["para"](u)
        h["para_red"]("Please verify the indexing status of these pages manually using Google Search Console "
                      "or by searching site:domain.com on Google.")
    h["shot"]("serp", captured)

def _sec_sitemap(h, findings, captured):
    h["para"]("Sitemap: A sitemap is a file where you can list the web pages of your site to tell Google "
              "and other search engines about the organization of your site content. Search engine web "
              "crawlers read this file to more intelligently crawl your site.")
    if findings.get("sitemap_found"):
        h["result"]("Result: Existing sitemap.xml file found. Please review the current sitemap once and "
                    "verify all important pages are included.")
        h["green"](f"Reference URL: {findings.get('sitemap_url')}")
        h["shot"]("sitemap", captured)
    else:
        h["result"]("Result: Sitemap.xml file not found on the website.")
        h["para_red"]("Please create a sitemap for the website. Different platforms have different methods:\n"
                      "- WordPress: Generate from admin using Yoast SEO or Rank Math plugin\n"
                      "- Shopify: Sitemap is auto-generated at /sitemap.xml\n"
                      "- Custom/Static sites: Create a static sitemap.xml and upload to root folder\n"
                      "- Other CMS: Check your platform's documentation for sitemap generation")

def _sec_redirection(h, findings):
    if findings.get("www_redirect_issue"):
        h["result"]("Redirection: The website runs with both www and non-www versions. We suggest redirecting "
                    "the www version to the non-www version using a 301 permanent redirect.")
    else:
        h["result"]("No redirection issue found on the website. It is good from a search engine point of view.")

def _sec_404(h, findings, root):
    if not findings.get("has_custom_404"):
        h["label_body"]("404 Custom Page Suggestion:")
        h["para"](f"We did not find a custom 404 page. Opening a mistyped URL such as {root}/asdfg redirects "
                  "to the default page instead of a 404 page. We recommend creating a custom 404 page.")

def _sec_canonical(h, findings, captured):
    d = h["d"]
    h["para"]("A canonical problem occurs when a site is running in multiple versions like www version "
              f"(http://www.{d}), the non-www version (http://{d}), trailing slash, "
              "or /index.html. Search engines may treat these as duplicates; a canonical tag tells them "
              "the preferred version.")
    if findings.get("canonical_issue"):
        h["result"]("Result: An incorrect / conflicting canonical tag was found. Kindly find the attached "
                    "Canonical Tag Suggestions sheet.")
    else:
        h["result"]("Result: Canonical tags are correctly set. Good from an SEO point of view.")
    h["shot"]("canonical", captured)

def _sec_external_links(h, findings, captured):
    h["result"](f"{findings.get('ext_count', 0)} external links found on the target pages checked.")
    if findings.get('ext_count', 0) > 0:
        h["para_red"]("Please verify whether any external links are harmful or beneficial as per need.")
    h["shot"]("externallinks", captured)

def _sec_broken_links(h, findings, captured):
    h["label_body"]("Broken Link Optimization")
    h["para"]("Though the broken links do not hurt directly but it affects user experience and if users "
              "do not find desired info they will not come frequently and search engine will rank the "
              "website down.", bold=True)
    broken_links = findings.get("broken_links", [])
    tp_count = len(findings["target_pages"])
    if broken_links:
        h["result"](f"Result: {len(broken_links)} broken link(s) found on the {tp_count} target page(s) "
                    "checked. It's not good from SEO point of view.")
        for bl in broken_links:
            h["para"](bl)
    else:
        h["result"](f"Result: No broken links found on the {tp_count} target page(s) checked. "
                    "It is good from an SEO point of view.")
    h["para_red"]("Note: Only external links on the target pages were checked. Please verify broken "
                  "links across the full website using Screaming Frog, Ahrefs, or Google Search Console.")
    h["shot"]("brokenlinks", captured)

def _sec_internal_linking(h, home):
    h["result"]("Internal Linking: Internal linking is called perfect when every live webpage is "
                "accessed/visited from every other live and available webpage in a website.")
    int_count = len(home.get("internal_links", []))
    h["result"](f"Result: {int_count} internal links found on the target pages checked.")
    h["para_red"]("Please verify the internal linking structure across the full website as per need.")

def _sec_page_speed(h, root):
    h["label_body"]("Page Optimization Score:")
    h["para_red"](f"Need to check it once. Check manually at: "
                  f"https://pagespeed.web.dev/analysis?url={root}/")

def _sec_url_structure(h, findings):
    root = h["root"]
    h["green"]("Recommendations: Search Engines like static URLs instead of dynamic one. And presently "
               "URL structure of your website's inner pages is user and search engine friendly.")
    if findings.get("url_changes"):
        h["result"]("Result: The URL structure of the following pages should be changed:")
        for ex, rec in findings["url_changes"]:
            h["para"](f"Existing  ->  {ex}\nRecommended  ->  {rec}", bold=True)
        h["para"]("Note: After changing the URL structure, 301-redirect each old URL to the new one.", bold=True)
    else:
        h["result"]("Result: Yes, Good. Your URL structure is search engine-friendly. It is good from an SEO "
                    "point of view.")
        for u in findings["target_pages"]:
            h["para"](u)

def _sec_hyperlinking(h):
    h["green"]("Recommendations: Hyperlinks are connections established between a word/phrase/image and a "
               "website/file. Effective hyper-linking between different pages of a website helps a website "
               "to rank better.")
    h["para_red"]("Result: Need to check it once.")

def _sec_mixed_content(h, findings, captured):
    h["ribbon"]("Mixed Content Optimization")
    h["para"]("Mixed content occurs when a webpage containing a combination of both secure (HTTPS) and "
              "non-secure (HTTP) content is delivered over SSL to the browser. It occurs when a website "
              "contains both HTTP and HTTPS content.")
    mixed_pages = findings.get("mixed_content_pages", [])
    tp_count = len(findings["target_pages"])
    if mixed_pages:
        h["result"](f"Result: Mixed content issues found on {len(mixed_pages)} of the {tp_count} target "
                    "page(s) checked. It's not good from an SEO point of view.")
        for mp in mixed_pages:
            h["para"](mp)
    else:
        h["result"](f"Result: No mixed content issues found on the {tp_count} target page(s) checked. "
                    "It's good from an SEO point of view.")
    h["shot"]("mixedcontent", captured)

def _sec_sucuri(h, findings, captured):
    d = h["d"]
    h["ribbon"]("Sucuri Site Scan")
    h["para"]("The Sucuri Site Check scanner helps to prevent security threats. It will check malware, "
              "viruses, blacklisting status, website errors, out-of-date software and malicious code.")
    sucuri_clean = findings.get("sucuri_clean")
    if sucuri_clean is True:
        h["result"]("Result: We have not found any malware/security issue on the website. It is good from "
                    "search engine point of view.")
    elif sucuri_clean is False:
        h["result"]("Result: Security issues were detected on the website. Please review and fix them.")
        h["para_red"]("Please check the details at the reference URL below.")
    else:
        h["para_red"](f"Please check the website security manually at: "
                      f"https://sitecheck.sucuri.net/results/https/{d}/")
    h["shot"]("sucuri", captured)
    h["green"](f"Reference URL: https://sitecheck.sucuri.net/results/https/{d}/")

def _sec_noindex(h, findings, captured):
    h["ribbon"]("No-index on Target Pages Check")
    h["para"]("Some pages of the website serve a purpose, and helps to improve the ranking and traffic "
              "to the site. These pages need to be there, as glue for other pages. But sometimes a few "
              "pages have a noindex tag that prevents them from being indexed.")
    noindex_pages = findings.get("noindex_pages", [])
    tp_count = len(findings["target_pages"])
    if noindex_pages:
        h["result"](f"Result: Noindex tag found on {len(noindex_pages)} of the {tp_count} target page(s) "
                    "checked. These pages will NOT get indexed on search engine. Please remove the noindex "
                    "tag.")
        for nip in noindex_pages:
            h["para"](nip)
    else:
        h["result"](f"Result: Noindex not found on the {tp_count} target page(s) checked. These pages "
                    "will get indexed on search engine.")
    h["para_red"]("Note: Only the target pages were checked. Please verify noindex tags across the full "
                  "website as per need.")
    h["shot"]("noindex", captured)

def _sec_web_archive(h, captured):
    root = h["root"]
    h["ribbon"]("Old Web Archive Status Check")
    h["para"]("It's the way to explore, find and retrieve historical and \"lost\" information from "
              "websites, to serve as evidence that something existed online, and was modified over time.")
    h["para_red"](f"Please check the web archive status at: "
                  f"https://web.archive.org/web/*/{root}/")

def _sec_viewport(h, findings, captured):
    h["ribbon"]("Meta viewport")
    h["para"]("The viewport meta tag allows you to tell the mobile browser what size this virtual "
              "viewport should be. This is often useful if you're not actually changing any visible "
              "content size but just the zoom level.")
    if findings.get("viewport"):
        h["result"]("Result: Yes, The Website pages have a viewport meta tag. It will look good on mobile "
                    "devices and will get a high position in mobile search results.")
    else:
        h["result"]("Result: Viewport meta tag missing — add it so pages render well on mobile.")
    h["shot"]("viewport", captured)

def _sec_lang(h, findings, captured):
    h["ribbon"]("lang Attribute")
    h["para"]("The \"lang\" attribute is an HTML attribute used to specify the language of the content "
              "within an HTML element. It helps search engines and assistive technologies understand "
              "the language of the page.")
    lang = findings.get("lang", "")
    if lang and "-" not in lang:
        h["result"](f'Result: We found lang="{lang}". If your target market is a specific region, we recommend '
                    f'a region-specific value (e.g. lang="en-US").')
    elif lang:
        h["result"](f'Result: We found the lang="{lang}" attribute on the website, it is good from search '
                    f'engine point of view.')
    else:
        h["result"]('Result: No lang attribute found — add an html lang attribute (e.g. lang="en-US").')
    h["shot"]("lang", captured)


# ---- Format: James (Driftzine) ----
def _build_docx_james(domain, pages_data, findings, captured, brand, out_path):
    doc, h = _setup_docx(domain)
    home = next((pd for pd in pages_data if urllib.parse.urlparse(pd["url"]).path in ("", "/")),
                pages_data[0] if pages_data else {})
    year = datetime.date.today().year
    root = h["root"]

    _sec_intro(h, root)
    _sec_on_page_analysis(h)
    h["summary_table"](findings)
    _sec_additional_notes(h, findings, captured, year)
    _sec_alt_tags(h, findings)
    _sec_robots(h, findings, captured)
    _sec_indexing(h, findings, captured)
    _sec_sitemap(h, findings, captured)
    _sec_redirection(h, findings)
    _sec_404(h, findings, root)
    _sec_canonical(h, findings, captured)
    _sec_external_links(h, findings, captured)
    _sec_broken_links(h, findings, captured)
    _sec_internal_linking(h, home)
    _sec_page_speed(h, root)
    _sec_url_structure(h, findings)
    _sec_hyperlinking(h)
    _sec_mixed_content(h, findings, captured)
    _sec_sucuri(h, findings, captured)
    _sec_noindex(h, findings, captured)
    _sec_web_archive(h, captured)
    _sec_viewport(h, findings, captured)
    _sec_lang(h, findings, captured)

    doc.save(out_path)


# ---- Format: Omega (alltechco) ----
def _build_docx_omega(domain, pages_data, findings, captured, brand, out_path):
    doc, h = _setup_docx(domain)
    home = next((pd for pd in pages_data if urllib.parse.urlparse(pd["url"]).path in ("", "/")),
                pages_data[0] if pages_data else {})
    year = datetime.date.today().year
    root = h["root"]

    _sec_intro(h, root)
    _sec_on_page_analysis(h)
    h["summary_table"](findings)
    _sec_additional_notes(h, findings, captured, year)
    _sec_alt_tags(h, findings)
    _sec_robots(h, findings, captured)
    _sec_indexing(h, findings, captured)
    _sec_sitemap(h, findings, captured)
    _sec_redirection(h, findings)
    _sec_404(h, findings, root)
    _sec_canonical(h, findings, captured)
    _sec_external_links(h, findings, captured)
    _sec_broken_links(h, findings, captured)
    _sec_internal_linking(h, home)
    _sec_page_speed(h, root)
    _sec_url_structure(h, findings)
    _sec_hyperlinking(h)
    _sec_mixed_content(h, findings, captured)
    _sec_sucuri(h, findings, captured)
    _sec_noindex(h, findings, captured)
    _sec_web_archive(h, captured)
    _sec_viewport(h, findings, captured)
    _sec_lang(h, findings, captured)

    doc.save(out_path)


# ---- Format: Neon (sumitechengineers) ----
def _build_docx_neon(domain, pages_data, findings, captured, brand, out_path):
    """Neon section order: Additional Suggestions → URLs → Hyperlinking → Robots →
    Image Alt → Internal Linking → External Links → Broken Links → Page Speed →
    Canonical → Mixed Content → No Index → URL Redirection → Sitemap → Site Security"""
    doc, h = _setup_docx(domain)
    home = next((pd for pd in pages_data if urllib.parse.urlparse(pd["url"]).path in ("", "/")),
                pages_data[0] if pages_data else {})
    year = datetime.date.today().year
    root = h["root"]

    _sec_intro(h, root)
    _sec_on_page_analysis(h)
    h["summary_table"](findings)
    _sec_additional_notes(h, findings, captured, year)
    _sec_url_structure(h, findings)
    _sec_hyperlinking(h)
    _sec_robots(h, findings, captured)
    _sec_alt_tags(h, findings)
    _sec_internal_linking(h, home)
    _sec_external_links(h, findings, captured)
    _sec_broken_links(h, findings, captured)
    _sec_page_speed(h, root)
    _sec_canonical(h, findings, captured)
    _sec_mixed_content(h, findings, captured)
    _sec_noindex(h, findings, captured)
    _sec_redirection(h, findings)
    _sec_sitemap(h, findings, captured)
    _sec_sucuri(h, findings, captured)
    _sec_indexing(h, findings, captured)
    _sec_web_archive(h, captured)
    _sec_viewport(h, findings, captured)
    _sec_lang(h, findings, captured)
    _sec_404(h, findings, root)

    doc.save(out_path)


# ---- Format: Xenon ----
def _build_docx_xenon(domain, pages_data, findings, captured, brand, out_path):
    """Xenon section order matches the Xenon reference: navy headers, status badges,
    alternating rows, light blue info bars."""
    doc, h = _setup_docx(domain)
    home = next((pd for pd in pages_data if urllib.parse.urlparse(pd["url"]).path in ("", "/")),
                pages_data[0] if pages_data else {})
    year = datetime.date.today().year
    root = h["root"]

    _sec_intro(h, root)
    _sec_on_page_analysis(h)
    h["summary_table"](findings)
    _sec_additional_notes(h, findings, captured, year)
    _sec_alt_tags(h, findings)
    _sec_robots(h, findings, captured)
    _sec_indexing(h, findings, captured)
    _sec_sitemap(h, findings, captured)
    _sec_redirection(h, findings)
    _sec_canonical(h, findings, captured)
    _sec_404(h, findings, root)
    _sec_url_structure(h, findings)
    _sec_hyperlinking(h)
    _sec_internal_linking(h, home)
    _sec_external_links(h, findings, captured)
    _sec_broken_links(h, findings, captured)
    _sec_page_speed(h, root)
    _sec_mixed_content(h, findings, captured)
    _sec_sucuri(h, findings, captured)
    _sec_noindex(h, findings, captured)
    _sec_web_archive(h, captured)
    _sec_viewport(h, findings, captured)
    _sec_lang(h, findings, captured)

    doc.save(out_path)


# ---- Format: Gamma (Hawkeev) ----
def _build_docx_gamma(domain, pages_data, findings, captured, brand, out_path):
    """Gamma: Book Antiqua headings, light blue #51C3F9 section headers,
    dark gold #984806 conclusions, green #00B050 screenshot labels,
    amber #FFC000 table headers."""
    doc, h = _setup_docx(domain)
    home = next((pd for pd in pages_data if urllib.parse.urlparse(pd["url"]).path in ("", "/")),
                pages_data[0] if pages_data else {})
    year = datetime.date.today().year
    root = h["root"]

    _sec_intro(h, root)
    _sec_on_page_analysis(h)
    h["summary_table"](findings)
    _sec_additional_notes(h, findings, captured, year)
    _sec_alt_tags(h, findings)
    _sec_robots(h, findings, captured)
    _sec_sitemap(h, findings, captured)
    _sec_canonical(h, findings, captured)
    _sec_redirection(h, findings)
    _sec_external_links(h, findings, captured)
    _sec_internal_linking(h, home)
    _sec_broken_links(h, findings, captured)
    _sec_hyperlinking(h)
    _sec_url_structure(h, findings)
    _sec_mixed_content(h, findings, captured)
    _sec_sucuri(h, findings, captured)
    _sec_noindex(h, findings, captured)
    _sec_indexing(h, findings, captured)
    _sec_web_archive(h, captured)
    _sec_viewport(h, findings, captured)
    _sec_lang(h, findings, captured)
    _sec_page_speed(h, root)
    _sec_404(h, findings, root)

    doc.save(out_path)


def build_content_suggestion_docx(domain, pages_data, targets, out_path):
    """Generate the Content Suggestion DOCX — per-page content optimization notes."""
    from docx import Document
    from docx.shared import Pt, RGBColor
    FONT, SIZE = "Arial", 11

    doc = Document()
    for sname in ("Normal",):
        try:
            st = doc.styles[sname]
            st.font.name = FONT
            st.font.size = Pt(SIZE)
        except KeyError:
            pass

    def _run(p, text, bold=False, color=None, size=None):
        r = p.add_run(text)
        r.font.name = FONT
        r.font.size = Pt(size or SIZE)
        r.font.bold = bold
        if color:
            r.font.color.rgb = color
        return r

    p = doc.add_paragraph()
    _run(p, "Content Optimization", bold=True, size=14)

    for i, t in enumerate(targets):
        page_url = t.get("page", "")
        keywords = t.get("keywords", [])
        kw_str = ", ".join(keywords) if keywords else ""

        if i == 0:
            p = doc.add_paragraph()
            _run(p, "Fresh Content Required: ", bold=True)
            _run(p, "We didn't find enough amount of keywords-oriented content integrated web-page. "
                     "We highly recommend you add some unique and relevant content (about 150-200 words) "
                     "containing targeted keywords within it on the following pages-", bold=True)
            p = doc.add_paragraph()
            _run(p, f"Page URL: {page_url}", bold=True)
        else:
            p = doc.add_paragraph()
            sep = "-" * 70
            _run(p, f"{sep}\nPage URL: {page_url}", bold=True)

        if kw_str:
            p = doc.add_paragraph()
            _run(p, f"Keywords: {kw_str}", bold=True)

        p = doc.add_paragraph()
        _run(p, "Section", bold=True, color=RGBColor(0xFF, 0x00, 0x00))

        p = doc.add_paragraph()
        _run(p, "Screenshots:", bold=True)

    p = doc.add_paragraph()
    _run(p, "=" * 82, bold=True)

    doc.save(out_path)


def build_onpage_docx(domain, pages_data, findings, captured, brand, out_path, fmt="james"):
    """Dispatch to the format-specific DOCX builder — builds EXACTLY the selected
    format, or raises rather than silently defaulting to another one."""
    builders = {
        "james": _build_docx_james, "omega": _build_docx_omega,
        "neon": _build_docx_neon, "xenon": _build_docx_xenon,
        "gamma": _build_docx_gamma,
    }
    fn = builders.get(str(fmt or "").strip().lower())
    if not fn:
        raise ValueError(f"Unknown on-page format '{fmt}'. "
                         f"Available: {', '.join(sorted(builders))}")
    fn(domain, pages_data, findings, captured, brand, out_path)


# ------------------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("domain", help="Target domain (positional, e.g. example.com)")
    ap.add_argument("--targets", help="targets.json or .xlsx (keywords + target pages). "
                                       "If omitted (e.g. Bulk), target pages are auto-discovered.")
    ap.add_argument("--out", default=str(OUTPUT_DIR))
    ap.add_argument("--dry-run", action="store_true", help="use mock crawl data (no network)")
    ap.add_argument("--no-capture", action="store_true", help="skip live screenshots (text-only docx)")
    ap.add_argument("--format", default="james", choices=["james", "omega", "neon", "xenon", "gamma"],
                    help="Report sub-format: james (Driftzine), omega (alltechco), neon (sumitechengineers), xenon, gamma (Hawkeev)")
    ap.add_argument("--gsc-token", default=None, help="GSC API access token for URL inspection")
    ap.add_argument("--property-url", default=None, help="GSC property URL (e.g. sc-domain:example.com)")
    # tolerate flags the shared task-runner may append (e.g. --account)
    args, _ = ap.parse_known_args()

    raw_domain = args.domain.strip()           # used verbatim for the output filename (glob match)
    domain = safe_domain(args.domain)          # normalized, for URL building
    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)
    work = out_dir / f"_seo_p2_{domain}"
    if work.exists():
        shutil.rmtree(work, ignore_errors=True)
    work.mkdir(parents=True)

    if args.targets:
        targets = load_targets(args.targets, domain)
    else:
        log("[*] No targets file — auto-discovering target pages from the site…")
        targets = discover_targets(domain, dry_run=args.dry_run)
    if not targets:
        log("[ERROR] No target pages found. Provide --targets, or make sure the site is reachable.")
        sys.exit(2)
    total = len(targets)
    log(f"[1/5] Loaded {total} target page(s)")

    # crawl + collect
    pages_data, homepage = [], None
    for i, t in enumerate(targets, 1):
        log(f"   -> [crawl {i}/{total}] {t['page']}")
        pd = crawl_page(t["page"], t["keywords"], dry_run=args.dry_run)
        pd["keywords"] = t["keywords"]
        pages_data.append(pd)
        if homepage is None or urllib.parse.urlparse(pd["url"]).path in ("", "/"):
            homepage = pd
    _close_op_driver()          # done crawling — free the render browser
    log(f"[2/5] Crawled {total} page(s)")

    brand = brand_from(domain, homepage.get("title") if homepage else None,
                       homepage.get("h1") if homepage else None,
                       homepage.get("og_site_name") if homepage else None)

    # deliverable data — Meta = existing + suggested (free Gemini if GEMINI_API_KEY,
    # else heuristic); plus alt suggestions + canonical recommendations.
    metas = [suggest_meta(pd, pd["keywords"], brand) for pd in pages_data]
    all_self_hosted, all_external_cdn = [], []
    for pd in pages_data:
        sh, ec = suggest_alt(pd, pd["keywords"], brand, domain=domain)
        all_self_hosted.extend(sh)
        all_external_cdn.extend(ec)
    canons = [recommend_canonical(pd) for pd in pages_data]
    log(f"[3/5] Built sheets (meta x{len(metas)}, alt x{len(all_self_hosted)}+{len(all_external_cdn)} ext, canonical x{len(canons)})")

    # audit findings + the site's REAL sitemap (not fabricated from target pages)
    findings = audit_site(domain, pages_data, dry_run=args.dry_run)
    sitemap_url, sitemap_body = find_existing_sitemap(domain, dry_run=args.dry_run)
    findings["sitemap_found"] = bool(sitemap_body)
    findings["sitemap_url"] = sitemap_url

    # GSC URL Inspection — check indexing status of target pages via API
    findings["gsc_indexing"] = None
    if args.gsc_token and args.property_url:
        log("[3.5/5] Checking indexing status via GSC API...")
        findings["gsc_indexing"] = check_gsc_indexing(
            [pd["url"] for pd in pages_data], args.gsc_token, args.property_url)

    captured = {}
    if not (args.dry_run or args.no_capture):
        log("[4/5] Capturing screenshots (Sucuri, robots, indexing, wayback…)")
        try:
            captured = capture_onpage_screenshots(domain, sitemap_url=sitemap_url)
        except Exception as e:
            log(f"   [warn] screenshot capture skipped: {type(e).__name__}: {e}")

    # deliverable files — names vary by format
    fmt = args.format
    log(f"   [format] {fmt}")

    # DOCX filename per format
    if fmt == "neon":
        doc_name = f"On-Page Suggestion Report - {domain}.docx"
    else:
        doc_name = f"On Page-Analysis-Report - {domain}.docx"

    meta_x = work / f"Meta Suggestions - {domain}.xlsx"
    can_x = work / f"Canonical Tag Suggestions - {domain}.xlsx"
    doc_f = work / doc_name

    write_meta_xlsx(metas, meta_x)
    write_canonical_xlsx(canons, can_x)

    # Omega: no separate Alt Tag XLSX; has Target-pages-and-keywords-Report
    # Neon + James: include Alt Tag XLSX + Target Pages
    if fmt == "omega":
        targets_x = work / f"Target-pages-and-keywords-Report - {domain}.xlsx"
        write_targets_xlsx(targets, targets_x)
    else:
        alt_x = work / f"Image Alt Tag Suggestions - {domain}.xlsx"
        write_alt_xlsx(all_self_hosted, all_external_cdn, alt_x)
        targets_x = work / f"Target Pages & Keywords Report - {domain}.xlsx"
        write_targets_xlsx(targets, targets_x)

    # robots.txt: attach the existing one
    if findings.get("robots_body"):
        (work / "Robots.txt").write_text(findings["robots_body"], encoding="utf-8")

    # sitemap: attach the existing one for review only — never generate from target pages
    if sitemap_body:
        (work / "sitemap (existing - review).xml").write_text(sitemap_body, encoding="utf-8")
    build_onpage_docx(domain, pages_data, findings, captured, brand, doc_f, fmt=fmt)

    content_doc = work / f"Content Suggestion - {domain}.docx"
    build_content_suggestion_docx(domain, pages_data, targets, content_doc)

    log("[4/5] Wrote docx + content suggestion + xlsx + sitemap")

    # bundle (filename uses the raw domain so the task-runner's glob matches)
    zip_path = out_dir / f"SEO On-Page Phase 2 - {raw_domain}.zip"
    if zip_path.exists():
        zip_path.unlink()
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        for f in sorted(work.iterdir()):
            if f.is_file():
                z.write(f, f.name)
    shutil.rmtree(work, ignore_errors=True)
    log(f"[5/5] Bundled -> {zip_path.name}")
    log(f"[DONE] {zip_path}")


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except Exception as e:
        # Surface any unexpected failure as a clean [ERROR] line so the task
        # store marks it failed and the UI shows a proper error toast.
        log(f"[ERROR] {type(e).__name__}: {e}")
        sys.exit(1)
