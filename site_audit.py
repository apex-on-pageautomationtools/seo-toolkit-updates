"""
Site Audit Module for SEO Toolkit Pro
Standalone technical/on-page crawl audit - given ONLY a domain (no manual page
list, no SE Ranking account), discovers pages (sitemap first, homepage-link
crawl fallback) and reports the kinds of issues a tool like SE Ranking's Site
Audit reports: broken links, missing/duplicate titles, missing/duplicate meta
descriptions, missing/multiple H1s, missing image alt text, thin content,
redirect chains, noindex pages, and sitemap/robots.txt problems.

Discovery and the broken-link / status-code / robots / sitemap checks reuse the
existing health_audit.py and scripts/generate_seo_onpage_phase2.py /
generate_geo_report.py logic rather than reimplementing sitemap parsing or
link-checking. Per-page title/meta/H1/alt-text extraction is done with a
lightweight requests-free (urllib + regex) fetch here, NOT brief_analysis.py's
_get_page/_render_html - those render every page through a real Selenium
browser, which is the right tradeoff for a handful of pages but far too slow
for the ~100-300 pages a site audit covers, so a plain HTTP fetch + a small
thread pool is used instead (same ThreadPoolExecutor pattern already used by
health_audit.check_broken_links).
"""
import os
import re
import sys
import time
import html as _html
import concurrent.futures
import urllib.request as _ur
import urllib.error as _ue

import health_audit

SCRIPTS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "scripts")
if SCRIPTS_DIR not in sys.path:
    sys.path.insert(0, SCRIPTS_DIR)
import generate_seo_onpage_phase2 as _op          # reuse homepage-link crawl fallback
import generate_geo_report as _geo                 # reuse sitemap discovery (get_sitemap_urls)

_UA = "Mozilla/5.0 (compatible; SEOToolkitPro-SiteAuditBot/1.0)"
THIN_CONTENT_WORDS = 200


# --------------------------------------------------------------------------- #
# Page discovery - reuses existing sitemap-parsing / homepage-crawl logic
# --------------------------------------------------------------------------- #
def discover_pages(domain, cap=300, log_fn=print):
    """Up to `cap` page URLs for domain: sitemap first (generate_geo_report's
    get_sitemap_urls, which already follows one level of sitemap-index nesting),
    falling back to a shallow homepage-link crawl (generate_seo_onpage_phase2's
    discover_targets) when no sitemap exists. Returns (urls, source)."""
    urls = []
    try:
        urls = _geo.get_sitemap_urls(domain, cap=cap)
    except Exception as e:
        log_fn(f"   [warn] sitemap discovery failed: {type(e).__name__}: {e}")
        urls = []
    if urls:
        log_fn(f"   -> {len(urls)} page(s) found via sitemap")
        return urls, "sitemap"

    log_fn("   -> no sitemap found, falling back to a homepage-link crawl")
    try:
        targets = _op.discover_targets(domain, dry_run=False, limit=cap)
        urls = [t["page"] for t in targets if t.get("page")]
    except Exception as e:
        log_fn(f"   [warn] homepage-link crawl failed: {type(e).__name__}: {e}")
        urls = []
    if not urls:
        urls = [f"https://{domain}/"]
    log_fn(f"   -> {len(urls)} page(s) found via homepage-link crawl")
    return urls, "crawl"


# --------------------------------------------------------------------------- #
# Per-page fetch + content extraction (title/meta/H1/alt/word-count/redirects)
# --------------------------------------------------------------------------- #
class _RedirectCounter(_ur.HTTPRedirectHandler):
    """Counts hops so a real redirect CHAIN (2+ hops before the final page) can
    be flagged, not just a single ordinary redirect (e.g. http -> https)."""
    def __init__(self):
        self.hops = []

    def redirect_request(self, req, fp, code, msg, headers, newurl):
        self.hops.append(newurl)
        return super().redirect_request(req, fp, code, msg, headers, newurl)


def _fetch_page(url, timeout=15):
    counter = _RedirectCounter()
    opener = _ur.build_opener(counter)
    try:
        req = _ur.Request(url, headers={"User-Agent": _UA})
        with opener.open(req, timeout=timeout) as r:
            html_body = r.read().decode("utf-8", "ignore")
            return {"url": url, "final_url": r.geturl(), "status": r.status,
                    "html": html_body, "redirect_hops": len(counter.hops), "error": None}
    except _ue.HTTPError as e:
        return {"url": url, "final_url": url, "status": e.code, "html": "",
                "redirect_hops": len(counter.hops), "error": None}
    except Exception as e:
        return {"url": url, "final_url": url, "status": 0, "html": "",
                "redirect_hops": 0, "error": str(e)}


def _extract_page_data(html_body):
    titles, descs, robots_vals = [], [], []
    for m in re.finditer(r'<title[^>]*>(.*?)</title>', html_body, re.I | re.S):
        t = re.sub(r'\s+', ' ', m.group(1)).strip()
        if t:
            titles.append(t)
    for m in re.finditer(r'<meta[^>]+name=["\']description["\'][^>]*content=["\']([^"\']*)["\']', html_body, re.I):
        descs.append(m.group(1).strip())
    for m in re.finditer(r'<meta[^>]+content=["\']([^"\']*)["\'][^>]*name=["\']description["\']', html_body, re.I):
        descs.append(m.group(1).strip())
    for m in re.finditer(r'<meta[^>]+name=["\']robots["\'][^>]*content=["\']([^"\']*)["\']', html_body, re.I):
        robots_vals.append(m.group(1).strip())
    for m in re.finditer(r'<meta[^>]+content=["\']([^"\']*)["\'][^>]*name=["\']robots["\']', html_body, re.I):
        robots_vals.append(m.group(1).strip())

    h1s = re.findall(r'<h1[^>]*>(.*?)</h1>', html_body, re.I | re.S)
    h1s = [re.sub(r'<[^>]+>', ' ', h).strip() for h in h1s]

    imgs = re.findall(r'<img\s[^>]*>', html_body, re.I)
    missing_alt = 0
    for tag in imgs:
        alt_m = re.search(r'alt\s*=\s*["\']([^"\']*)["\']', tag, re.I)
        if not alt_m or not alt_m.group(1).strip():
            missing_alt += 1

    text = re.sub(r'<(script|style)[^>]*>.*?</\1>', ' ', html_body, flags=re.I | re.S)
    text = re.sub(r'<[^>]+>', ' ', text)
    text = _html.unescape(text)
    word_count = len(re.findall(r"[A-Za-z0-9][\w'-]*", text))

    return {
        "title": titles[0] if titles else "",
        "title_count": len(titles),
        "description": descs[0] if descs else "",
        "description_count": len(descs),
        "robots": robots_vals[0] if robots_vals else "",
        "h1_count": len(h1s),
        "h1s": h1s[:3],
        "image_count": len(imgs),
        "missing_alt_count": missing_alt,
        "word_count": word_count,
    }


def _analyze_pages(urls, log_fn=print, max_workers=10, stop_event=None):
    """Fetch + extract on-page data for every discovered URL concurrently -
    same ThreadPoolExecutor pattern health_audit.check_broken_links already
    uses for checking many links at once, applied here to many pages."""
    results = []
    total = len(urls)
    done = 0
    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as pool:
        futures = {pool.submit(_fetch_page, u): u for u in urls}
        for fut in concurrent.futures.as_completed(futures):
            page = fut.result()
            if page["html"]:
                page.update(_extract_page_data(page["html"]))
            results.append(page)
            done += 1
            if done % 20 == 0 or done == total:
                log_fn(f"   -> analyzed {done}/{total} page(s)")
            if stop_event is not None and stop_event.is_set():
                break
    return results


# --------------------------------------------------------------------------- #
# Aggregation
# --------------------------------------------------------------------------- #
_SUMMARY_KEYS = (
    "broken_links", "missing_titles", "duplicate_titles", "missing_meta",
    "duplicate_meta", "missing_h1", "multiple_h1", "missing_alt_images",
    "thin_content_pages", "redirect_chains", "noindex_pages",
    "sitemap_issues", "robots_issues",
)


def _aggregate(pages, status_results, broken_links, robots_check, sitemap_check):
    issues = {k: [] for k in _SUMMARY_KEYS}

    # Pages that are themselves unreachable/erroring (checked via the existing
    # check_status200), plus links found on pages that lead nowhere - both are
    # "broken" from a site-audit point of view, so they share one bucket.
    for r in status_results:
        if not r["ok"]:
            issues["broken_links"].append({"url": r["url"], "code": r["status"],
                                            "found_on": ["(page itself)"], "suggested_redirect": None})
    issues["broken_links"].extend(broken_links)

    title_map, desc_map = {}, {}
    for p in pages:
        if p.get("error") or not p.get("status") or not (200 <= p["status"] < 300):
            continue
        url = p["url"]
        title = p.get("title", "")
        desc = p.get("description", "")
        if title:
            title_map.setdefault(title, []).append(url)
        else:
            issues["missing_titles"].append({"url": url})
        if desc:
            desc_map.setdefault(desc, []).append(url)
        else:
            issues["missing_meta"].append({"url": url})

        h1c = p.get("h1_count", 0)
        if h1c == 0:
            issues["missing_h1"].append({"url": url})
        elif h1c > 1:
            issues["multiple_h1"].append({"url": url, "count": h1c})

        if p.get("missing_alt_count", 0) > 0:
            issues["missing_alt_images"].append({"url": url, "count": p["missing_alt_count"]})

        if p.get("word_count", 0) and p["word_count"] < THIN_CONTENT_WORDS:
            issues["thin_content_pages"].append({"url": url, "word_count": p["word_count"]})

        if p.get("redirect_hops", 0) >= 2:
            issues["redirect_chains"].append({"url": url, "hops": p["redirect_hops"],
                                               "final_url": p.get("final_url")})

        if "noindex" in (p.get("robots") or "").lower():
            issues["noindex_pages"].append({"url": url, "robots": p.get("robots")})

    for title, urls_ in title_map.items():
        if len(urls_) > 1:
            issues["duplicate_titles"].append({"title": title, "urls": urls_})
    for desc, urls_ in desc_map.items():
        if len(urls_) > 1:
            issues["duplicate_meta"].append({"description": desc, "urls": urls_})

    if not robots_check.get("ok", True):
        issues["robots_issues"].append({"summary": robots_check.get("summary", "")})
    if not sitemap_check.get("ok", True):
        issues["sitemap_issues"].append({"summary": sitemap_check.get("summary", "")})

    summary = {k: len(v) for k, v in issues.items()}
    return summary, issues


# --------------------------------------------------------------------------- #
# Entry point
# --------------------------------------------------------------------------- #
def run_site_audit(domain, cap=300, log_fn=print, stop_event=None):
    domain = (domain or "").strip().lower()
    for prefix in ("https://", "http://"):
        if domain.startswith(prefix):
            domain = domain[len(prefix):]
    domain = domain.split("/")[0]
    if not domain:
        raise ValueError("Domain is required.")

    log_fn(f"[1/4] Discovering pages for {domain} (cap {cap})...")
    urls, source = discover_pages(domain, cap=cap, log_fn=log_fn)

    log_fn("[2/4] Checking robots.txt and sitemap...")
    robots_check = health_audit.check_robots_txt(domain)
    sitemap_check = health_audit.check_sitemap(domain, robots_check)

    log_fn(f"[3/4] Checking status codes and broken links across {len(urls)} page(s)...")
    status_results = health_audit.check_status200(urls, domain)
    total_links, broken_links = health_audit.check_broken_links(domain, target_pages=urls)
    log_fn(f"   -> {len(broken_links)} broken link(s) found out of {total_links} link(s) checked")

    log_fn(f"[4/4] Analyzing {len(urls)} page(s) for title/meta/H1/alt/content issues...")
    pages = _analyze_pages(urls, log_fn=log_fn, stop_event=stop_event)

    summary, issues = _aggregate(pages, status_results, broken_links, robots_check, sitemap_check)

    log_fn("Done.")
    return {
        "domain": domain,
        "pages_discovered": len(urls),
        "pages_analyzed": len(pages),
        "discovery_source": source,
        "summary": summary,
        "issues": issues,
        "robots": robots_check,
        "sitemap": sitemap_check,
        "checked_at": time.strftime("%Y-%m-%d %H:%M:%S"),
    }
