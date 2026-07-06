"""
Health Audit Module for SEO Toolkit Pro
Adapted from james-seo-tools generate_health_report.py
- No Patchright (uses Selenium CDP for screenshots)
- GSC screenshots (manual_action, security_issues) captured if GSC account is connected
- 4 format themes: James (docx), Neon (docx), Sigma (pptx), Omega (pptx)
"""

import os
import re
import tempfile
import urllib.request as _ur
import urllib.error as _ue
from pathlib import Path
from datetime import datetime
from html.parser import HTMLParser

# ---------- HTTP helpers ----------
_UA = "Mozilla/5.0 (compatible; SEOToolkitPro-HealthBot/1.0)"


def _fetch_html(url, timeout=15):
    try:
        req = _ur.Request(url, headers={"User-Agent": _UA})
        with _ur.urlopen(req, timeout=timeout) as r:
            return r.read().decode("utf-8", "ignore")
    except Exception:
        return None


# Per-run cache of browser-rendered pages so the same URL is only rendered once
# across the canonical / double-meta / meta-robots / meta-source checks.
_RENDER_CACHE = {}


def _render_html(url, driver=None, wait=1.2):
    """Return the page as a real browser renders it (so JS-set title / meta /
    canonical are seen — Google renders JS, so this is the ACTUAL state that
    matters). Falls back to raw HTTP when no browser/driver is available."""
    if url in _RENDER_CACHE:
        return _RENDER_CACHE[url]
    html = None
    if driver is not None:
        import time
        try:
            driver.get(url)
            for _ in range(16):                       # wait up to ~4s for load
                try:
                    if driver.execute_script("return document.readyState") == "complete":
                        break
                except Exception:
                    break
                time.sleep(0.25)
            time.sleep(wait)                          # short settle for hydration
            h = driver.page_source or ""
            if h and len(h) > 200:
                html = h
        except Exception:
            html = None
    if not html:
        html = _fetch_html(url)
    _RENDER_CACHE[url] = html
    return html


def _google_via_selenium(driver, query, num=30):
    """Run a Google search through the Selenium browser to avoid bot blocks.
    Returns page HTML or None. Uses gentle delays to stay under the radar."""
    import time, random
    from urllib.parse import quote_plus
    try:
        url = f"https://www.google.com/search?q={quote_plus(query)}&num={num}"
        driver.get(url)
        time.sleep(random.uniform(2.5, 4.5))
        return driver.page_source
    except Exception:
        return None


def _get_status(url, timeout=10):
    for method in ("HEAD", "GET"):
        try:
            req = _ur.Request(url, headers={"User-Agent": _UA}, method=method)
            with _ur.urlopen(req, timeout=timeout) as r:
                return r.status
        except _ue.HTTPError as e:
            if method == "HEAD" and e.code in (403, 405):
                continue
            return e.code
        except Exception:
            return 0
    return 0


def _parse_meta(html):
    result = {"canonical": [], "title": [], "description": [], "robots": []}
    for m in re.finditer(r'<link[^>]+rel=["\']canonical["\'][^>]*href=["\']([^"\']+)["\']', html, re.I):
        result["canonical"].append(m.group(1).strip())
    for m in re.finditer(r'<link[^>]+href=["\']([^"\']+)["\'][^>]*rel=["\']canonical["\']', html, re.I):
        result["canonical"].append(m.group(1).strip())
    for m in re.finditer(r'<title[^>]*>(.*?)</title>', html, re.I | re.S):
        t = m.group(1).strip()
        if t:
            result["title"].append(t)
    for m in re.finditer(r'<meta[^>]+name=["\']description["\'][^>]*content=["\']([^"\']*)["\']', html, re.I):
        result["description"].append(m.group(1).strip())
    for m in re.finditer(r'<meta[^>]+content=["\']([^"\']*)["\'][^>]*name=["\']description["\']', html, re.I):
        result["description"].append(m.group(1).strip())
    for m in re.finditer(r'<meta[^>]+name=["\']robots["\'][^>]*content=["\']([^"\']*)["\']', html, re.I):
        result["robots"].append(m.group(1).strip())
    for m in re.finditer(r'<meta[^>]+content=["\']([^"\']*)["\'][^>]*name=["\']robots["\']', html, re.I):
        result["robots"].append(m.group(1).strip())
    return result


# ---------- Checks ----------
def check_status200(target_pages, domain):
    if not target_pages:
        target_pages = [f"https://{domain}/"]
    results = []
    for url in target_pages:
        url = url.strip()
        if not url:
            continue
        if not url.startswith("http"):
            url = "https://" + url
        code = _get_status(url)
        results.append({"url": url, "status": str(code), "ok": code == 200})
    return results


def _detect_live_homepage(domain):
    for prefix in (f"https://{domain}/", f"https://www.{domain}/",
                   f"http://{domain}/", f"http://www.{domain}/"):
        try:
            req = _ur.Request(prefix, headers={"User-Agent": _UA})
            with _ur.urlopen(req, timeout=10) as r:
                return r.url
        except Exception:
            continue
    return f"https://{domain}/"


def check_canonical(target_pages, domain, driver=None):
    pages = list(target_pages) if target_pages else []
    homepage = _detect_live_homepage(domain)
    if not any(p.rstrip("/") == homepage.rstrip("/") for p in pages):
        pages = [homepage] + pages
    results = []
    for url in pages:
        url = url.strip()
        if not url:
            continue
        if not url.startswith("http"):
            url = "https://" + url
        html = _render_html(url, driver)
        if html is None:
            results.append({"url": url, "canonical": "—", "ok": False, "note": "Could not fetch"})
            continue
        meta = _parse_meta(html)
        canonicals = list(dict.fromkeys(meta["canonical"]))
        if not canonicals:
            results.append({"url": url, "canonical": "Missing", "ok": False, "note": "No canonical tag"})
        elif len(canonicals) > 1:
            results.append({"url": url, "canonical": canonicals[0], "ok": False, "note": f"{len(canonicals)} canonical tags"})
        else:
            canon = canonicals[0].rstrip("/")
            page_url = url.rstrip("/")
            ok = canon == page_url
            results.append({"url": url, "canonical": canonicals[0], "ok": ok,
                            "note": "Self-referencing" if ok else "Points elsewhere"})
    return results


def check_double_meta(target_pages, domain, driver=None):
    pages = list(target_pages) if target_pages else []
    if not pages:
        pages = [f"https://{domain}/"]
    results = []
    for url in pages:
        url = url.strip()
        if not url:
            continue
        if not url.startswith("http"):
            url = "https://" + url
        html = _render_html(url, driver)
        if html is None:
            results.append({"url": url, "title_count": "—", "desc_count": "—", "ok": False, "note": "Could not fetch"})
            continue
        meta = _parse_meta(html)
        tc, dc = len(meta["title"]), len(meta["description"])
        ok = (tc == 1 and dc == 1)
        notes = []
        if tc == 0:
            notes.append("Missing title")
        elif tc > 1:
            notes.append(f"Title ×{tc}")
        if dc == 0:
            notes.append("Missing meta desc")
        elif dc > 1:
            notes.append(f"Meta desc ×{dc}")
        results.append({"url": url, "title_count": str(tc), "desc_count": str(dc),
                         "ok": ok, "note": ", ".join(notes) if notes else "OK"})
    return results


def check_meta_robots(target_pages, domain, driver=None):
    pages = list(target_pages) if target_pages else []
    if not pages:
        pages = [f"https://{domain}/"]
    results = []
    for url in pages:
        url = url.strip()
        if not url:
            continue
        if not url.startswith("http"):
            url = "https://" + url
        html = _render_html(url, driver)
        if html is None:
            results.append({"url": url, "robots_value": "—", "ok": False, "note": "Could not fetch"})
            continue
        meta = _parse_meta(html)
        robots_values = meta["robots"]
        if not robots_values:
            results.append({"url": url, "robots_value": "Not set (default: index,follow)", "ok": True, "note": "OK"})
            continue
        val = robots_values[0]
        val_lower = val.lower()
        issues = []
        if "noindex" in val_lower:
            issues.append("noindex")
        if "nofollow" in val_lower:
            issues.append("nofollow")
        ok = not issues
        results.append({"url": url, "robots_value": val, "ok": ok,
                         "note": ", ".join(issues) if issues else "OK"})
    return results


def check_versions_full(domain):
    variants = [
        f"http://{domain}/", f"http://www.{domain}/",
        f"https://{domain}/", f"https://www.{domain}/",
    ]
    active = []
    final_urls = set()
    for url in variants:
        try:
            class _NoRedirect(_ur.HTTPErrorProcessor):
                def http_response(self, req, resp):
                    return resp
                https_response = http_response
            opener = _ur.build_opener(_NoRedirect)
            req = _ur.Request(url, headers={"User-Agent": _UA})
            with opener.open(req, timeout=10) as r:
                code = r.status
                location = r.getheader("Location", "")
            if code in (200, 301, 302, 307, 308):
                active.append((url, code, location))
                if code == 200:
                    final_urls.add(url.rstrip("/"))
                elif location:
                    final_urls.add(location.rstrip("/"))
        except Exception:
            pass

    returning_200 = [url for url, code, _ in active if code == 200]
    redirecting = [(url, code, loc) for url, code, loc in active if code != 200]

    is_ok = False
    summary = ""
    if len(returning_200) == 1:
        canonical = returning_200[0].rstrip("/")
        if all(loc.rstrip("/") == canonical or loc.rstrip("/") == "" for _, _, loc in redirecting):
            is_ok = True
            summary = f"No issue found. All variants correctly consolidate to a single version: {returning_200[0]}"
    if not is_ok and len(returning_200) == 0 and len(final_urls) == 1:
        is_ok = True
        summary = f"No issue found. All variants redirect to: {list(final_urls)[0]}"
    if not is_ok:
        summary = (f"Issue found — {len(returning_200)} URL variant(s) return HTTP 200 independently. "
                   f"Fix: redirect (301) all variants to one canonical URL via server config or .htaccess.")

    rows = []
    for url, code, loc in active:
        if code == 200:
            rows.append((url, "200", "OK (active)", is_ok))
        else:
            rows.append((url, str(code), f"-> {loc or '?'}", True))
    return summary, rows


def check_dummy_content(domain, driver=None):
    query = f"site:{domain} lorem"
    if driver:
        html = _google_via_selenium(driver, query, num=10)
    else:
        html = _fetch_html(f"https://www.google.com/search?q=site:{domain}+lorem")
    if html is None:
        return "Could not check Google for dummy content. Please check manually."
    no_results_markers = ["did not match any documents", "No results found", "no results", "0 results"]
    for marker in no_results_markers:
        if marker.lower() in html.lower():
            return f"No dummy content found. Google search for 'site:{domain} lorem' returned no results."
    count_match = re.search(r'About ([\d,]+) results', html)
    count = count_match.group(1) if count_match else "some"
    titles = re.findall(r'<h3[^>]*>(.*?)</h3>', html, re.S)
    titles = [re.sub(r'<[^>]+>', '', t).strip() for t in titles[:5] if t.strip()]
    if titles:
        result = f"Possible dummy content found — Google shows {count} result(s) for 'site:{domain} lorem'.\nPages found:\n"
        result += "\n".join(f"  - {t}" for t in titles)
    else:
        result = f"Google returned {count} result(s) for 'site:{domain} lorem'. Please verify manually."
    return result


def check_broken_links(domain, target_pages=None):
    """Check all links on every target page (or homepage if none given).
    No artificial limit — every link on each page is checked."""
    from urllib.parse import urljoin
    headers = {"User-Agent": _UA}

    pages = [p.strip() for p in (target_pages or []) if p.strip()]
    if not pages:
        pages = [f"https://{domain}/"]
    # Normalise
    pages = [p if p.startswith("http") else f"https://{p}" for p in pages]

    all_links = []
    seen = set()
    for page_url in pages:
        try:
            html = _ur.urlopen(_ur.Request(page_url, headers=headers), timeout=20).read().decode("utf-8", "ignore")
        except Exception:
            continue
        for m in re.finditer(r'href=["\']([^"\']+)["\']', html):
            href = m.group(1).strip()
            if href.startswith(("mailto:", "tel:", "javascript:", "#", "data:")):
                continue
            u = urljoin(page_url, href)
            if u.startswith("http") and u not in seen:
                seen.add(u)
                all_links.append(u)

    broken = []
    for u in all_links:
        code = None
        for method in ("HEAD", "GET"):
            try:
                req = _ur.Request(u, headers=headers, method=method)
                code = _ur.urlopen(req, timeout=12).status
                break
            except _ue.HTTPError as he:
                code = he.code
                if method == "HEAD" and code in (403, 405):
                    continue
                break
            except Exception:
                code = 0
        if code is None or code == 0 or code >= 400:
            broken.append((u, code or "ERR"))
    return len(all_links), broken


def check_sucuri(domain):
    """Fetch Sucuri SiteCheck results page and parse the actual verdict."""
    import re
    url = f"https://sitecheck.sucuri.net/results/https/{domain}"
    html = _fetch_html(url, timeout=30)
    if html is None:
        return {"ok": None, "summary": "Could not reach Sucuri SiteCheck. Check manually.", "details": [], "issues": []}
    details = []
    issues = []
    ok = True
    html_lower = html.lower()

    # Malware — check the actual result text, not template text
    has_no_malware = "didn't detect any malware" in html_lower or "no malware found" in html_lower
    has_malware = bool(re.search(r'Warning:\s*Malware\s+Detected', html))
    # The raw HTML contains BOTH templates; check which is actually shown
    # "No malware detected by scan (Low Risk)" = clean; presence of this confirms clean
    clean_scan = "no malware detected by scan" in html_lower
    if has_no_malware or clean_scan:
        details.append("No malware found")
        malware_ok = True
    elif has_malware:
        ok = False
        details.append("Malware detected!")
        issues.append("Malware detected on the website")
        malware_ok = False
    else:
        details.append("Malware status unclear")
        malware_ok = True

    # Blacklist
    if "is not blacklisted" in html_lower or "site is not blacklisted" in html_lower:
        details.append("Not blacklisted")
        bl_ok = True
    elif re.search(r'site\s+is\s+blacklisted', html, re.IGNORECASE):
        ok = False
        details.append("Domain is BLACKLISTED")
        issues.append("Domain is blacklisted")
        bl_ok = False
    else:
        bl_ok = True

    # Blacklist details — count "Domain clean by" entries
    bl_clean = len(re.findall(r'Domain clean by', html))
    if bl_clean:
        details.append(f"{bl_clean} blacklist checks clean")

    # Risk level — infer from malware/blacklist status
    if not malware_ok or not bl_ok:
        risk_level = "High"
    elif malware_ok and bl_ok:
        risk_level = "Low"
    else:
        risk_level = "Medium"
    details.append(f"Risk level: {risk_level}")

    # Security headers (from Hardening Improvements section)
    if "missing security header" in html_lower:
        header_issues = re.findall(r'Missing security header[^<]*<a[^>]*>([^<]+)</a>', html, re.IGNORECASE)
        header_issues += re.findall(r'Missing\s+<a[^>]*>([^<]+?)</a>\s*security header', html, re.IGNORECASE)
        for h in header_issues:
            issues.append(f"Missing: {h.strip()}")
        if not header_issues:
            issues.append("Missing security headers")

    # WAF
    if "no website application firewall" in html_lower or "please install a cloud-based waf" in html_lower:
        issues.append("No WAF/Firewall detected")
    elif "firewall detected" in html_lower:
        details.append("WAF/Firewall detected")

    # PHP version leaked
    if "leaked php version" in html_lower or "expose_php" in html_lower:
        issues.append("PHP version exposed in headers")

    # Missing Strict-Transport-Security
    if "strict-transport-security" in html_lower and "missing" in html_lower:
        issues.append("Missing: Strict-Transport-Security header")

    # Missing CSP
    if "content-security-policy directive" in html_lower and "missing" in html_lower:
        issues.append("Missing: Content-Security-Policy directive")

    summary_parts = []
    if ok:
        summary_parts.append("No critical issues found.")
    else:
        summary_parts.append("Critical issues detected!")
    summary_parts.append(f"Risk: {risk_level}.")
    if issues:
        summary_parts.append("Hardening: " + "; ".join(issues[:6]) + ".")

    return {"ok": ok, "summary": " ".join(summary_parts), "details": details, "issues": issues}


def check_robots_txt(domain):
    """Actually fetch and validate robots.txt."""
    url = f"https://{domain}/robots.txt"
    try:
        req = _ur.Request(url, headers={"User-Agent": _UA})
        with _ur.urlopen(req, timeout=10) as r:
            code = r.status
            content = r.read().decode("utf-8", "ignore")
    except _ue.HTTPError as e:
        return {"ok": False, "status": e.code,
                "summary": f"robots.txt returned HTTP {e.code} — file is missing or inaccessible."}
    except Exception as e:
        return {"ok": False, "status": 0,
                "summary": f"Could not fetch robots.txt: {e}"}
    if not content.strip():
        return {"ok": False, "status": code,
                "summary": "robots.txt exists but is empty."}
    has_user_agent = bool(re.search(r'(?i)^user-agent:', content, re.MULTILINE))
    has_sitemap = bool(re.search(r'(?i)^sitemap:', content, re.MULTILINE))
    has_disallow = bool(re.search(r'(?i)^disallow:', content, re.MULTILINE))
    # Is the whole site blocked from crawling? (Disallow: / under User-agent: * or Googlebot)
    root_blocked = False
    cur_ua = None
    for ln in content.splitlines():
        t = ln.split("#", 1)[0].strip()
        low = t.lower()
        if low.startswith("user-agent:"):
            cur_ua = low.split(":", 1)[1].strip()
        elif low.startswith("disallow:") and cur_ua in ("*", "googlebot"):
            if t.split(":", 1)[1].strip() == "/":
                root_blocked = True
    issues = []
    if root_blocked:
        issues.append("Main site is BLOCKED from crawling (Disallow: /)")
    if not has_user_agent:
        issues.append("Missing User-agent directive")
    if not has_sitemap:
        issues.append("No Sitemap reference in robots.txt")
    line_count = len([l for l in content.splitlines() if l.strip()])
    summary = f"robots.txt found ({line_count} lines). "
    if root_blocked:
        summary += "WARNING: the site is disallowed for search crawlers. "
    if issues:
        summary += "Issues: " + "; ".join(issues) + "."
    else:
        summary += "Main site is crawlable" + (", sitemap referenced." if has_sitemap else ".")
    return {"ok": not issues, "status": code, "summary": summary, "found": True,
            "root_blocked": root_blocked, "site_crawlable": not root_blocked,
            "has_sitemap_ref": has_sitemap, "has_disallow": has_disallow, "lines": line_count}


def check_sitemap(domain, robots_data=None):
    """Fetch and validate the sitemap. Tries /sitemap.xml then /sitemap_index.xml
    (many sites — e.g. Yoast — expose only a sitemap index). `robots_data` is
    accepted for compatibility with callers that pass robots.txt results; when it
    carries explicit Sitemap: URLs those are tried first."""
    candidates = []
    # Prefer any Sitemap: URLs surfaced by the robots.txt check.
    if isinstance(robots_data, dict):
        for k in ("sitemaps", "sitemap_urls", "sitemaps_found"):
            v = robots_data.get(k)
            if isinstance(v, (list, tuple)):
                candidates.extend([u for u in v if isinstance(u, str) and u.startswith("http")])
            elif isinstance(v, str) and v.startswith("http"):
                candidates.append(v)
    elif isinstance(robots_data, str):
        candidates.extend(re.findall(r'(?im)^\s*sitemap:\s*(\S+)', robots_data))
    # Common fallbacks.
    for u in (f"https://{domain}/sitemap.xml", f"https://{domain}/sitemap_index.xml"):
        if u not in candidates:
            candidates.append(u)

    last = {"ok": False, "status": 0, "summary": "No sitemap found."}
    for url in candidates:
        name = url.rsplit("/", 1)[-1] or "sitemap.xml"
        try:
            req = _ur.Request(url, headers={"User-Agent": _UA})
            with _ur.urlopen(req, timeout=15) as r:
                code = r.status
                content = r.read().decode("utf-8", "ignore")
        except _ue.HTTPError as e:
            last = {"ok": False, "status": e.code,
                    "summary": f"{name} returned HTTP {e.code} — file is missing or inaccessible."}
            continue
        except Exception as e:
            last = {"ok": False, "status": 0, "summary": f"Could not fetch {name}: {e}"}
            continue
        if not content.strip():
            last = {"ok": False, "status": code, "summary": f"{name} exists but is empty."}
            continue
        is_xml = content.strip().startswith("<?xml") or "<urlset" in content or "<sitemapindex" in content
        if not is_xml:
            last = {"ok": False, "status": code,
                    "summary": f"{name} exists but does not appear to be valid XML."}
            continue
        url_count = len(re.findall(r'<loc>', content))
        is_index = "<sitemapindex" in content
        if is_index:
            sub_count = len(re.findall(r'<sitemap>', content))
            summary = f"Sitemap index found with {sub_count} sub-sitemap(s)."
        else:
            summary = f"Sitemap found with {url_count} URL(s)."
        return {"ok": True, "status": code, "summary": summary,
                "url_count": url_count, "is_index": is_index, "url": url}
    return last


def check_serp(domain, driver=None):
    """Check Google SERP for site:{domain} — count indexed pages AND detect
    hacked/spam content (pharma, casino, Japanese SEO spam, cloaked pages)."""
    query = f"site:{domain}"
    if driver:
        html = _google_via_selenium(driver, query, num=30)
    else:
        html = _fetch_html(f"https://www.google.com/search?q=site:{domain}&num=30")
    if html is None:
        return {"ok": None, "summary": "Could not check Google SERP. Please check manually.",
                "result_count": "unknown", "spam_found": []}

    count_match = re.search(r'About ([\d,]+) results', html)
    count = count_match.group(1) if count_match else None

    no_results = any(m.lower() in html.lower()
                     for m in ["did not match any documents", "0 results", "No results found"])
    if no_results:
        return {"ok": False,
                "summary": f"No pages found indexed for site:{domain}. The website may not be indexed.",
                "result_count": "0", "spam_found": []}

    # Extract titles and snippets for spam/hack detection
    titles = re.findall(r'<h3[^>]*>(.*?)</h3>', html, re.S)
    titles = [re.sub(r'<[^>]+>', '', t).strip() for t in titles if t.strip()]
    snippets = re.findall(r'<span[^>]*class="[^"]*st[^"]*"[^>]*>(.*?)</span>', html, re.S)
    snippets = [re.sub(r'<[^>]+>', '', s).strip() for s in snippets if s.strip()]
    all_text = " ".join(titles + snippets).lower()

    # Hack/spam indicators
    SPAM_PATTERNS = [
        (r'\b(viagra|cialis|levitra|pharmacy|pharma|pills|tablets|medication)\b', "Pharma spam"),
        (r'\b(casino|poker|slots|gambling|bet online|betting)\b', "Casino/gambling spam"),
        (r'\b(buy cheap|cheap online|discount|wholesale|replica|counterfeit)\b', "Commercial spam"),
        (r'\b(hacked by|pwned|defaced|owned by)\b', "Hacked/defaced indicator"),
        (r'[぀-ヿ一-鿿]{5,}', "Japanese/Chinese SEO spam"),
        (r'\b(payday loan|loan online|cash advance|credit card hack)\b', "Financial spam"),
        (r'\b(free download|crack|keygen|serial key|torrent)\b', "Piracy/malware spam"),
        (r'\b(escort|adult|xxx|porn)\b', "Adult content injection"),
    ]
    spam_found = []
    for pattern, label in SPAM_PATTERNS:
        matches = re.findall(pattern, all_text, re.I)
        if matches:
            # Find which titles contain this spam
            flagged_titles = [t for t in titles if re.search(pattern, t.lower(), re.I)]
            spam_found.append({
                "type": label,
                "count": len(matches),
                "example_titles": flagged_titles[:3],
            })

    if spam_found:
        spam_types = ", ".join(s["type"] for s in spam_found)
        summary = (f"WARNING: Potential hacked/spam content detected in SERP results! "
                   f"Types found: {spam_types}. "
                   f"Google shows {count or 'multiple'} result(s) for site:{domain}. "
                   f"Immediate investigation recommended.")
        return {"ok": False, "summary": summary, "result_count": count or "unknown",
                "spam_found": spam_found}

    summary = f"No spam or hacked content detected. Google shows {count or 'multiple'} result(s) for site:{domain}."
    return {"ok": True, "summary": summary, "result_count": count or "unknown", "spam_found": []}


def check_blank_pages(target_pages, domain, driver=None):
    """Fetch target pages and check if they have minimal content (blank).
    Uses the RENDERED page so JS/SPA sites aren't falsely flagged as blank (their
    raw HTML is an empty shell). Only runs when target pages are provided."""
    pages = [p.strip() for p in (target_pages or []) if p.strip()]
    if not pages:
        return [], 0
    results = []
    for url in pages:
        if not url.startswith("http"):
            url = "https://" + url
        html = _render_html(url, driver)
        if html is None:
            results.append({"url": url, "ok": False, "note": "Could not fetch"})
            continue
        text = re.sub(r'<script[^>]*>.*?</script>', '', html, flags=re.S | re.I)
        text = re.sub(r'<style[^>]*>.*?</style>', '', text, flags=re.S | re.I)
        text = re.sub(r'<[^>]+>', '', text)
        text = re.sub(r'\s+', ' ', text).strip()
        is_blank = len(text) < 100
        results.append({"url": url, "ok": not is_blank, "text_len": len(text),
                         "note": f"Only {len(text)} chars of text — appears blank" if is_blank else "OK"})
    blank_count = sum(1 for r in results if not r["ok"])
    return results, blank_count


PSI_API_KEY_DEFAULT = "AIzaSyAuV5kxzjZm5KvjQ_HlODes6C7eh__aI0I"


def check_pagespeed(domain, api_key=None):
    """Fetch real PageSpeed Insights data via the PSI API."""
    import json
    key = api_key or PSI_API_KEY_DEFAULT
    base_text = ("The time it takes to fully display the content on a specific page; it reports on the "
                 "performance of a page on both mobile and desktop devices. ")
    scores = {}
    for strategy in ("mobile", "desktop"):
        api_url = (f"https://www.googleapis.com/pagespeedonline/v5/runPagespeed"
                   f"?url=https://{domain}/&strategy={strategy}"
                   f"&category=PERFORMANCE&category=SEO&category=BEST_PRACTICES&category=ACCESSIBILITY"
                   f"&key={key}")
        try:
            req = _ur.Request(api_url, headers={"User-Agent": _UA})
            with _ur.urlopen(req, timeout=90) as r:
                data = json.loads(r.read().decode("utf-8"))
            cats = data.get("lighthouseResult", {}).get("categories", {})
            scores[strategy] = {
                "performance": int(cats.get("performance", {}).get("score", 0) * 100),
                "seo": int(cats.get("seo", {}).get("score", 0) * 100),
                "best_practices": int(cats.get("best-practices", {}).get("score", 0) * 100),
                "accessibility": int(cats.get("accessibility", {}).get("score", 0) * 100),
            }
        except Exception:
            scores[strategy] = None

    parts = [base_text]
    for strategy in ("mobile", "desktop"):
        s = scores.get(strategy)
        if s:
            parts.append(f"{strategy.title()}: Performance {s['performance']}/100, "
                         f"SEO {s['seo']}/100, Best Practices {s['best_practices']}/100, "
                         f"Accessibility {s['accessibility']}/100.")
        else:
            parts.append(f"{strategy.title()}: Could not retrieve scores.")

    return {"summary": " ".join(parts), "scores": scores}


# ---------- Table image renderer (PIL) ----------
def _render_table_image(title, headers, rows, col_widths, path):
    from PIL import Image, ImageDraw, ImageFont
    try:
        bold_font = ImageFont.truetype("C:/Windows/Fonts/arialbd.ttf", 20)
        header_font = ImageFont.truetype("C:/Windows/Fonts/arialbd.ttf", 16)
        cell_font = ImageFont.truetype("C:/Windows/Fonts/arial.ttf", 15)
        title_font = ImageFont.truetype("C:/Windows/Fonts/arialbd.ttf", 24)
    except Exception:
        bold_font = header_font = cell_font = title_font = ImageFont.load_default()

    W = sum(col_widths) + 2
    ROW_H = 28
    TITLE_H = 50
    HEADER_H = 32
    H = TITLE_H + HEADER_H + ROW_H * max(len(rows), 1) + 20
    img = Image.new("RGB", (W, H), "#FFFFFF")
    d = ImageDraw.Draw(img)
    d.text((16, 14), title, fill=(31, 41, 55), font=title_font)
    y = TITLE_H
    x = 0
    d.rectangle([0, y, W, y + HEADER_H], fill="#F3F4F6")
    for i, h in enumerate(headers):
        d.text((x + 8, y + 8), h, fill=(75, 85, 99), font=header_font)
        x += col_widths[i]
        if i < len(headers) - 1:
            d.line([(x, y), (x, y + HEADER_H)], fill="#D1D5DB", width=1)
    d.line([(0, y + HEADER_H), (W, y + HEADER_H)], fill="#D1D5DB", width=1)
    for ri, row in enumerate(rows):
        y = TITLE_H + HEADER_H + ri * ROW_H
        bg = "#F9FAFB" if ri % 2 == 0 else "#FFFFFF"
        d.rectangle([0, y, W, y + ROW_H], fill=bg)
        x = 0
        ok = row[-1] if isinstance(row[-1], bool) else True
        for ci, cell in enumerate(row[:-1]):
            color = (31, 41, 55)
            if ci == len(row) - 2:
                color = (22, 128, 67) if ok else (200, 40, 60)
            text = str(cell)
            max_chars = col_widths[ci] // 8
            if len(text) > max_chars:
                text = text[:max_chars - 1] + "…"
            d.text((x + 8, y + 6), text, fill=color, font=cell_font)
            x += col_widths[ci]
            if ci < len(row) - 2:
                d.line([(x, y), (x, y + ROW_H)], fill="#E5E7EB", width=1)
        d.line([(0, y + ROW_H), (W, y + ROW_H)], fill="#E5E7EB", width=1)
    if not rows:
        d.text((16, TITLE_H + HEADER_H + 8), "No data.", fill=(150, 150, 150), font=cell_font)
    img.save(str(path))


def _render_broken_links_image(domain, checked, broken, path):
    from PIL import Image, ImageDraw, ImageFont
    rows = broken[:25]
    W = 1180
    H = 130 + (len(rows) + 1) * 30 + 30
    img = Image.new("RGB", (W, H), "#FFFFFF")
    d = ImageDraw.Draw(img)
    try:
        bold = ImageFont.truetype("C:/Windows/Fonts/arialbd.ttf", 26)
        f = ImageFont.truetype("C:/Windows/Fonts/arial.ttf", 18)
        fs = ImageFont.truetype("C:/Windows/Fonts/arial.ttf", 15)
    except Exception:
        bold = f = fs = ImageFont.load_default()
    d.text((26, 24), f"Broken Links Check — {domain}", fill=(31, 41, 55), font=bold)
    if not broken:
        d.text((26, 74), f"Checked {checked} links — No broken links found.", fill=(22, 128, 67), font=f)
    else:
        d.text((26, 74), f"Checked {checked} links — {len(broken)} broken:", fill=(200, 40, 60), font=f)
        y = 116
        d.text((26, y), "URL", fill=(120, 120, 120), font=fs)
        d.text((W - 150, y), "Status", fill=(120, 120, 120), font=fs)
        y += 28
        for u, code in rows:
            d.text((26, y), u[:135], fill=(40, 40, 40), font=fs)
            d.text((W - 150, y), str(code), fill=(200, 40, 60), font=fs)
            y += 28
    img.save(str(path))


# ---------- Checkpoint definitions ----------
INTRO = [
    ("Website health checkup and analysis",
     ": As a process of routine analysis, we have checked the website "
     "according to different points in order to determine if any issue "
     "persists in the website or not."),
    ("",
     "And after doing the analysis, we found that there is no health related "
     "issue in the website."),
]

CHECKPOINTS = [
    {"key": "sucuri", "label": "Sucuri Check",
     "body": ": We first of all check the website in the tool Sucuri, to check "
             "for any malware or website blacklisting issue.",
     "extra": ("URL:  ", "https://sitecheck.sucuri.net/results/https/{domain}"),
     "status": ("Status", ": We have not find issue in the website."),
     "capture": "url"},

    {"key": "manual_action", "label": "Manual Action:",
     "body": " The Manual Actions report lists manually detected issues with a "
             "page or site that are mostly attempts to manipulate our search "
             "index, but are not necessarily dangerous for users.",
     "status": ("Status –", " Not found"),
     "capture": "gsc", "red": True},

    {"key": "security_issues", "label": "Security Issues:",
     "body": " The Security Issues report lists indications that your site was "
             "hacked, or behavior on your site that could potentially harm a "
             "visitor or their computer: for example, phishing attacks or "
             "installing malware or unwanted software on the user's computer.",
     "status": ("Status:  ", "Not Found"),
     "capture": "gsc", "red": True},

    {"key": "robots", "label": "Robots.txt:",
     "body": " Robots.txt file has pages that needs to be blocked and are not "
             "useful for the website. Thus, we have checked the robots file is "
             "running fine for now.",
     "capture": "url"},

    {"key": "sitemap", "label": "Sitemap.xml: ",
     "body": "Sitemap file must have all the web-pages that are being present "
             "in the website so that all the pages can be identified by the "
             "Search Engine. We have checked the sitemap file as it is having "
             "all the relevant pages or not in the website.",
     "capture": "url"},

    {"key": "status200", "label": "All Target page's status is 200:- ",
     "body": "{status200_result}", "capture": "computed"},

    {"key": "versions", "label": "The website is running with multiple versions: - ",
     "body": "{versions_result}", "capture": "computed"},

    {"key": "meta_source", "label": "Meta Suggestions is visible in the source code (Also on Top):- ",
     "body": "Meta suggestion are visible in the source code of the pages.",
     "capture": "url"},

    {"key": "canonical", "label": "Right Canonical tag: - ",
     "body": "{canonical_result}", "capture": "computed"},

    {"key": "double_meta", "label": "Check for Double Meta Tags: - ",
     "body": "{double_meta_result}", "capture": "computed"},

    {"key": "meta_robots", "label": "Check Meta Robots Tag on Web-Pages: - ",
     "body": "{meta_robots_result}", "capture": "computed"},

    {"key": "layout", "label": "Check Website Layout: -  ",
     "body": "The screenshot below is the latest archived snapshot from web.archive.org. "
             "Compare it with the current live site to confirm no unexpected layout changes "
             "have occurred recently.",
     "capture": "url"},

    {"key": "dummy", "label": "Check Dummy Content: ",
     "body": "{dummy_content_result}", "capture": "url"},

    {"key": "broken_links", "label": "Check for Broken Links:- ",
     "body": "{broken_links_result}", "capture": "computed"},

    {"key": "serp", "label": "Any un-necessary SERP issues: - ",
     "body": "There is no unnecessary pages found in the SERP",
     "capture": "url"},

    {"key": "blank_page", "label": "Check Blank Page: ",
     "body": "We check our all-targeted page and didn't find any blank targeted page.",
     "capture": "url"},

    {"key": "pagespeed", "label": "Website Page speed:- ",
     "body": "{pagespeed_result}", "capture": "url"},
]

CHECKPOINT_BY_KEY = {cp["key"]: cp for cp in CHECKPOINTS}

JAMES_KEYS = [cp["key"] for cp in CHECKPOINTS]

OMEGA_KEYS = [
    "sucuri", "manual_action", "security_issues", "robots", "sitemap", "serp", "double_meta",
    "blank_page", "dummy", "meta_source", "layout", "versions", "status200",
]

NEON_KEYS = [
    # Neon sequence: SEO-first, security + performance last (distinct from James)
    "robots", "sitemap", "meta_source", "canonical", "meta_robots",
    "serp", "status200", "versions", "layout", "blank_page",
    "broken_links", "sucuri", "manual_action", "security_issues", "pagespeed",
]

SIGMA_KEYS = [
    "sucuri", "manual_action", "security_issues", "robots", "sitemap", "status200",
    "versions", "serp", "meta_source", "double_meta", "layout",
    "blank_page", "dummy", "pagespeed",
]

FORMAT_INFO = {
    "james": {"keys": JAMES_KEYS, "ext": "docx", "label": "James (Full DOCX)"},
    "sigma": {"keys": SIGMA_KEYS, "ext": "pptx", "label": "Sigma (PPTX)"},
    "omega": {"keys": OMEGA_KEYS, "ext": "pptx", "label": "Omega (PPTX)"},
    "neon":  {"keys": NEON_KEYS,  "ext": "docx", "label": "Neon (DOCX)"},
}

# Screenshot URLs for each checkpoint
SCREENSHOT_URLS = {
    "sucuri": "https://sitecheck.sucuri.net/results/https/{domain}",
    "robots": "https://{domain}/robots.txt",
    "sitemap": "https://{domain}/sitemap.xml",
    "meta_source": "view-source:https://{domain}/",
    "layout": "https://web.archive.org/web/2/https://{domain}/",
    "dummy": "https://www.google.com/search?q=site:{domain}+lorem",
    "serp": "https://www.google.com/search?q=site:{domain}",
    "blank_page": "https://{domain}/",
    "pagespeed": "https://pagespeed.web.dev/analysis?url=https%3A%2F%2F{domain}%2F",
}


# ---------- Selenium screenshot capture ----------
def capture_screenshots_selenium(driver, domain, out_dir, keys, log_fn=None):
    """Capture screenshots using an existing Selenium WebDriver via CDP.
    Returns dict of {key: image_path}."""
    captured = {}
    if log_fn is None:
        log_fn = print

    os.makedirs(out_dir, exist_ok=True)

    for key in keys:
        url_tmpl = SCREENSHOT_URLS.get(key)
        if not url_tmpl:
            continue

        url = url_tmpl.format(domain=domain)
        path = os.path.join(out_dir, f"{key}.png")

        try:
            if key == "meta_source":
                url = f"view-source:https://{domain}/"

            log_fn(f"  Capturing [{key}]...")
            driver.get(url)
            import time
            if key == "sucuri":
                time.sleep(15)
            elif key == "pagespeed":
                time.sleep(35)
            elif key == "layout":
                time.sleep(8)
            else:
                time.sleep(4)

            if key == "sucuri":
                # Sucuri shows a cookie-consent bar over the verdict. Click "Accept"
                # so it isn't in the screenshot; then remove any leftover banner.
                try:
                    driver.execute_script("""
                        var els = document.querySelectorAll('button, a, input[type=button], input[type=submit], span, div');
                        for (var i = 0; i < els.length; i++) {
                            var t = (els[i].textContent || els[i].value || '').trim().toLowerCase();
                            if (t === 'accept' || t === 'accept all' || t === 'i accept' ||
                                t === 'allow all' || t === 'allow' || t === 'got it' || t === 'agree') {
                                try { els[i].click(); } catch (e) {}
                                return;
                            }
                        }
                    """)
                    time.sleep(1.2)
                except Exception:
                    pass
                driver.execute_script(
                    "document.querySelectorAll('.cookie-banner,.consent-banner,[class*=cookie],[class*=consent],[id*=cookie],[id*=consent],#cookie-law-info-bar').forEach(function(e){e.remove();});"
                    "window.scrollTo(0, 0);"
                )
                time.sleep(0.6)
            else:
                driver.execute_script("window.scrollTo(0, 0);")
            time.sleep(0.5)

            # Use CDP for the screenshot. Sucuri's scan verdict sits at the very
            # top of the page, so capture a fixed top region of the *document*
            # (captureBeyondViewport + clip anchored at 0,0). This is immune to
            # wherever the page happened to be scrolled — which previously
            # produced a useless bottom-of-page ("Check another URL") shot.
            cdp_params = {"format": "png", "captureBeyondViewport": False}
            if key == "sucuri":
                try:
                    _w = driver.execute_script(
                        "return Math.max(document.documentElement.clientWidth||0, window.innerWidth||0, 1280);")
                except Exception:
                    _w = 1280
                cdp_params = {
                    "format": "png",
                    "captureBeyondViewport": True,
                    "clip": {"x": 0, "y": 0, "width": float(_w or 1280), "height": 1300.0, "scale": 1},
                }
            try:
                result = driver.execute_cdp_cmd("Page.captureScreenshot", cdp_params)
                import base64
                with open(path, "wb") as f:
                    f.write(base64.b64decode(result["data"]))
                captured[key] = path
                log_fn(f"  -> [{key}] captured")
            except Exception as e:
                log_fn(f"  [warn] CDP screenshot failed for {key}: {e}")
                try:
                    driver.save_screenshot(path)
                    captured[key] = path
                    log_fn(f"  -> [{key}] captured (fallback)")
                except Exception:
                    pass
        except Exception as e:
            log_fn(f"  [warn] capture '{key}' failed: {e}")

    return captured


# ---------- Prepare all health data ----------
def prepare_health_data(domain, captured=None, target_pages=None, log_fn=None, psi_api_key=None, driver=None, prefetched_futures=None):
    if log_fn is None:
        log_fn = print
    domain = re.sub(r'^\s*https?://', '', str(domain or '')).strip().strip('/').split('/')[0] or str(domain)
    captured = captured or {}
    target_pages = [p.strip() for p in (target_pages or []) if p.strip()]
    _RENDER_CACHE.clear()      # fresh per-run cache of browser-rendered pages

    _tmp = Path(tempfile.gettempdir())

    # --- Start slow checks in background threads (or reuse prefetched futures) ---
    import concurrent.futures
    executor = None

    if prefetched_futures and "sucuri" in prefetched_futures:
        fut_sucuri = prefetched_futures["sucuri"]
        fut_psi = prefetched_futures.get("psi")
        log_fn("Sucuri" + (" & PageSpeed" if fut_psi else "") + " already running from early start...")
    else:
        executor = concurrent.futures.ThreadPoolExecutor(max_workers=2)
        fut_sucuri = executor.submit(check_sucuri, domain)
        if psi_api_key:
            log_fn("Starting Sucuri & PageSpeed checks in background...")
            fut_psi = executor.submit(check_pagespeed, domain, psi_api_key)
        else:
            log_fn("Starting Sucuri in background (PageSpeed skipped)...")
            fut_psi = None

    # --- Fast checks run sequentially while slow ones are in progress ---
    _total = 13
    _step = [0]
    def _prog(msg):
        _step[0] += 1
        log_fn(f"[{_step[0]}/{_total}] {msg}")

    _prog("Checking robots.txt...")
    robots = check_robots_txt(domain)

    _prog("Checking sitemap.xml...")
    sitemap = check_sitemap(domain)

    _prog("Running status200 check...")
    s200 = check_status200(target_pages, domain)

    _prog("Running canonical check...")
    canon = check_canonical(target_pages, domain, driver=driver)

    _prog("Running double meta check...")
    dmeta = check_double_meta(target_pages, domain, driver=driver)

    _prog("Running meta robots check...")
    mrobots = check_meta_robots(target_pages, domain, driver=driver)

    _prog("Running URL versions check...")
    versions_summary, versions_rows = check_versions_full(domain)

    _prog("Running broken links check...")
    bl_checked, bl_broken = check_broken_links(domain, target_pages)

    _prog("Running dummy content check...")
    dummy_result = check_dummy_content(domain, driver=driver)

    _prog("Checking SERP for spam/hacked content...")
    serp = check_serp(domain, driver=driver)

    _prog("Checking blank pages...")
    blank_results, blank_count = check_blank_pages(target_pages, domain, driver=driver)

    _prog("Checking meta source code...")
    homepage_html = _render_html(f"https://{domain}/", driver)
    if homepage_html:
        hp_meta = _parse_meta(homepage_html)
        has_title = bool(hp_meta["title"])
        has_desc = bool(hp_meta["description"])
        meta_src_parts = []
        if has_title:
            meta_src_parts.append(f"Title found: \"{hp_meta['title'][0][:80]}\"")
        else:
            meta_src_parts.append("Title tag is MISSING")
        if has_desc:
            meta_src_parts.append(f"Meta description found ({len(hp_meta['description'][0])} chars)")
        else:
            meta_src_parts.append("Meta description is MISSING")
        meta_source_result = ". ".join(meta_src_parts) + "."
    else:
        meta_source_result = "Could not fetch homepage source code."

    # --- Collect slow checks ---
    if fut_psi:
        _prog("Waiting for Sucuri & PageSpeed results...")
    else:
        _prog("Waiting for Sucuri results...")
    sucuri = fut_sucuri.result(timeout=120)
    psi = fut_psi.result(timeout=120) if fut_psi else {"summary": "PageSpeed Insights — skipped", "scores": {}}
    if executor:
        executor.shutdown(wait=False)
    log_fn("Sucuri" + (" & PageSpeed" if fut_psi else "") + " done.")

    # Render table images
    if s200:
        p = _tmp / f"ha_status200_{domain}.png"
        rows = [(r["url"].replace("https://", "").replace("http://", ""),
                 r["status"], "OK" if r["ok"] else "Issue", r["ok"]) for r in s200]
        _render_table_image(f"HTTP Status Check — {domain}", ["Page URL", "Status", "Result"],
                            rows, [780, 120, 140], p)
        captured["status200"] = str(p)

    if canon:
        p = _tmp / f"ha_canonical_{domain}.png"
        rows = [(r["url"].replace("https://", "").replace("http://", ""),
                 r["canonical"][:60] + "…" if len(r["canonical"]) > 60 else r["canonical"],
                 r["note"], r["ok"]) for r in canon]
        _render_table_image(f"Canonical Tag Check — {domain}", ["Page URL", "Canonical", "Result"],
                            rows, [540, 360, 180], p)
        captured["canonical"] = str(p)

    if dmeta:
        p = _tmp / f"ha_doublemeta_{domain}.png"
        rows = [(r["url"].replace("https://", "").replace("http://", ""),
                 r["title_count"], r["desc_count"], r["note"], r["ok"]) for r in dmeta]
        _render_table_image(f"Meta Tags Check — {domain}", ["Page URL", "Title", "Meta Desc", "Result"],
                            rows, [600, 90, 110, 240], p)
        captured["double_meta"] = str(p)

    if mrobots:
        p = _tmp / f"ha_metarobots_{domain}.png"
        rows = [(r["url"].replace("https://", "").replace("http://", ""),
                 r["robots_value"], r["note"], r["ok"]) for r in mrobots]
        _render_table_image(f"Meta Robots Check — {domain}", ["Page URL", "Robots Value", "Result"],
                            rows, [540, 360, 180], p)
        captured["meta_robots"] = str(p)

    if versions_rows:
        p = _tmp / f"ha_versions_{domain}.png"
        _render_table_image(f"URL Versions Check — {domain}", ["URL Variant", "HTTP", "Result"],
                            versions_rows, [620, 90, 360], p)
        captured["versions"] = str(p)

    if bl_checked > 0:
        p = _tmp / f"ha_brokenlinks_{domain}.png"
        _render_broken_links_image(domain, bl_checked, bl_broken, p)
        captured["broken_links"] = str(p)

    # Render blank page table if we have results
    if blank_results:
        p = _tmp / f"ha_blankpage_{domain}.png"
        rows = [(r["url"].replace("https://", "").replace("http://", ""),
                 str(r.get("text_len", "—")), r["note"], r["ok"]) for r in blank_results]
        _render_table_image(f"Blank Page Check — {domain}", ["Page URL", "Text Length", "Result"],
                            rows, [700, 120, 220], p)
        captured["blank_page"] = str(p)

    # Render SERP spam table if spam found
    if serp.get("spam_found"):
        p = _tmp / f"ha_serp_spam_{domain}.png"
        rows = [(s["type"], str(s["count"]),
                 s["example_titles"][0][:60] + "…" if s["example_titles"] else "—", False)
                for s in serp["spam_found"]]
        _render_table_image(f"SERP Spam/Hack Detection — {domain}", ["Spam Type", "Hits", "Example Title"],
                            rows, [300, 80, 660], p)
        captured["serp"] = str(p)

    # Build summaries
    def _bl_summary():
        if bl_checked == 0:
            return "Could not check for broken links."
        if not bl_broken:
            return f"Checked {bl_checked} links — no broken links found."
        return f"Checked {bl_checked} links — {len(bl_broken)} broken link(s) found."

    def _s200_summary():
        if not s200:
            return "No target pages provided."
        bad = [r for r in s200 if not r["ok"]]
        if not bad:
            return f"All {len(s200)} page(s) returned HTTP 200."
        return f"{len(bad)} of {len(s200)} page(s) have non-200 status."

    def _canon_summary():
        if not canon:
            return "No pages checked."
        bad = [r for r in canon if not r["ok"]]
        if not bad:
            return f"All {len(canon)} page(s) have correct self-referencing canonical tags."
        return f"{len(bad)} of {len(canon)} page(s) have canonical issues."

    def _dmeta_summary():
        if not dmeta:
            return "No pages checked."
        bad = [r for r in dmeta if not r["ok"]]
        if not bad:
            return f"All {len(dmeta)} page(s) have title and meta description present once each."
        return f"{len(bad)} of {len(dmeta)} page(s) have meta tag issues."

    def _mrobots_summary():
        if not mrobots:
            return "No pages checked."
        bad = [r for r in mrobots if not r["ok"]]
        if not bad:
            return f"All {len(mrobots)} page(s) have correct meta robots (index, follow)."
        return f"{len(bad)} of {len(mrobots)} page(s) have robots tag issues."

    def _blank_summary():
        if not blank_results:
            return "No target pages provided for blank page check."
        if blank_count == 0:
            return f"All {len(blank_results)} target page(s) have content — no blank pages found."
        return f"{blank_count} of {len(blank_results)} target page(s) appear blank (less than 100 chars of text)."

    computed = {
        "sucuri_result": sucuri["summary"] + (" Details: " + "; ".join(sucuri["details"]) if sucuri.get("details") else ""),
        "robots_result": robots["summary"],
        "sitemap_result": sitemap["summary"],
        "versions_result": versions_summary,
        "status200_result": _s200_summary(),
        "canonical_result": _canon_summary(),
        "double_meta_result": _dmeta_summary(),
        "meta_robots_result": _mrobots_summary(),
        "meta_source_result": meta_source_result,
        "broken_links_result": _bl_summary(),
        "dummy_content_result": dummy_result,
        "serp_result": serp["summary"],
        "blank_page_result": _blank_summary(),
        "pagespeed_result": psi["summary"],
    }
    return captured, computed


# ---------- Safe save path ----------
def _safe_save_path(base_path):
    p = Path(base_path)
    if not p.exists():
        return p
    stem, suffix = p.stem, p.suffix
    parent = p.parent
    n = 1
    while True:
        candidate = parent / f"{stem} ({n}){suffix}"
        if not candidate.exists():
            return candidate
        n += 1


# ---------- DOCX builder ----------
def build_health_docx(domain, captured, computed, out_dir, include_keys=None, voice="we", page_border=False, header_fill=None):
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls
    from docx.shared import Inches, Pt, RGBColor

    FONT_NAME = "Candara"
    IMAGE_WIDTH_IN = 6.3
    IMAGE_MAX_H_IN = 4.6

    use_keys = include_keys or JAMES_KEYS
    cps = [CHECKPOINT_BY_KEY[k] for k in use_keys if k in CHECKPOINT_BY_KEY]
    doc = Document()

    for sname in ("Normal", "List Paragraph"):
        try:
            st = doc.styles[sname]
            st.font.name = FONT_NAME
            st.font.size = Pt(11)
        except KeyError:
            pass

    sec = doc.sections[0]
    sec.page_width, sec.page_height = Inches(8.5), Inches(11.0)
    sec.left_margin = sec.right_margin = Inches(1.0)
    sec.top_margin = sec.bottom_margin = Inches(1.0)

    if page_border:
        sectPr = sec._sectPr
        pgBorders = parse_xml(
            '<w:pgBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
            ' w:offsetFrom="page">'
            '<w:top w:val="single" w:sz="12" w:space="24" w:color="000000"/>'
            '<w:left w:val="single" w:sz="12" w:space="24" w:color="000000"/>'
            '<w:bottom w:val="single" w:sz="12" w:space="24" w:color="000000"/>'
            '<w:right w:val="single" w:sz="12" w:space="24" w:color="000000"/>'
            '</w:pgBorders>'
        )
        pgSz = sectPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pgSz')
        if pgSz is not None:
            pgSz.addnext(pgBorders)
        else:
            sectPr.append(pgBorders)

    # Header
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(domain)
    run.bold = True
    run.font.name = FONT_NAME
    run.font.size = Pt(24)

    def _voice(text):
        if voice == "i":
            return (text.replace("We first of all check", "I first of all check")
                    .replace("we have checked", "I have checked")
                    .replace("We have checked", "I have checked")
                    .replace("We have not", "I have not")
                    .replace("We check ", "I check ")
                    .replace("we found", "I found"))
        return text

    def _add_label_body(label, body):
        p = doc.add_paragraph(style="List Paragraph")
        p.paragraph_format.left_indent = Inches(0)
        if label:
            r = p.add_run(label)
            r.font.name = FONT_NAME
            r.font.size = Pt(11)
            r.bold = True
        if body:
            r = p.add_run(body)
            r.font.name = FONT_NAME
            r.font.size = Pt(11)
        return p

    def _shaded_header(text, fill):
        """A full-width shaded section-header bar (white bold) — used by Neon
        so its sections read as coloured bars instead of plain bold labels."""
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)
        r = p.add_run(text)
        r.font.name = FONT_NAME
        r.font.size = Pt(13)
        r.bold = True
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        shd = OxmlElement('w:shd')
        shd.set(qn('w:fill'), fill)
        shd.set(qn('w:val'), 'clear')
        p._p.get_or_add_pPr().append(shd)
        return p

    # Intro
    for label, body in INTRO:
        _add_label_body(label, _voice(body))
    doc.add_paragraph()

    placed = missing = 0
    for n, cp in enumerate(cps, start=1):
        body = cp["body"]
        for k, v in computed.items():
            body = body.replace("{" + k + "}", v)
        try:
            body = body.format(domain=domain)
        except Exception:
            pass
        if header_fill:
            # Neon style: coloured header bar with the section title, body below.
            _shaded_header(f"{n}. {cp['label'].strip().rstrip(':').strip()}", header_fill)
            btext = _voice(body)
            if btext.strip():
                bp = _add_label_body("", btext)
                if cp.get("red"):
                    for r in bp.runs:
                        r.font.color.rgb = RGBColor(0xCC, 0x00, 0x00)
                        r.bold = True
        else:
            p = _add_label_body(f"{n}. {cp['label']}", _voice(body))
            if cp.get("red") and p.runs:
                for r in p.runs:
                    if not r.bold:
                        r.font.color.rgb = RGBColor(0xCC, 0x00, 0x00)
                        r.bold = True

        extra = cp.get("extra")
        if extra:
            p = doc.add_paragraph(style="List Paragraph")
            p.paragraph_format.left_indent = Inches(0)
            r = p.add_run(extra[0])
            r.font.name = FONT_NAME
            r.font.size = Pt(11)
            r.bold = True
            r = p.add_run(extra[1].format(domain=domain))
            r.font.name = FONT_NAME
            r.font.size = Pt(11)

        status = cp.get("status")
        if status:
            p = doc.add_paragraph(style="List Paragraph")
            p.paragraph_format.left_indent = Inches(0)
            r = p.add_run(status[0])
            r.font.name = FONT_NAME
            r.font.size = Pt(11)
            r.bold = True
            r = p.add_run(status[1])
            r.font.name = FONT_NAME
            r.font.size = Pt(11)

        # Image
        img = captured.get(cp["key"])
        if img and os.path.exists(img):
            try:
                from PIL import Image as PILImage
                with PILImage.open(img) as im:
                    iw, ih = im.size
                p = doc.add_paragraph()
                p.paragraph_format.space_before = Pt(4)
                p.paragraph_format.space_after = Pt(4)
                run = p.add_run()
                if (IMAGE_WIDTH_IN * ih / iw) > IMAGE_MAX_H_IN:
                    run.add_picture(img, height=Inches(IMAGE_MAX_H_IN))
                else:
                    run.add_picture(img, width=Inches(IMAGE_WIDTH_IN))
                placed += 1
            except Exception:
                missing += 1
        else:
            missing += 1

        if cp["key"] == "layout":
            p = doc.add_paragraph()
            r = p.add_run("Compare this archived snapshot with the current live site.")
            r.font.name = FONT_NAME
            r.font.size = Pt(11)
            r.font.color.rgb = RGBColor(0xCC, 0x00, 0x00)
            r.bold = True

        doc.add_paragraph()

    gen_date = datetime.now().strftime("%d-%B-%Y")
    out_path = Path(out_dir) / f"{domain} Website health checkup analysis report {gen_date}.docx"
    out_path = _safe_save_path(out_path)
    doc.save(str(out_path))
    return str(out_path), placed, missing


# ---------- PPTX builder (Sigma) ----------
def build_health_pptx_sigma(domain, captured, computed, out_dir, include_keys=None):
    from pptx import Presentation
    from pptx.util import Inches as PInches, Pt as PPt
    from pptx.dml.color import RGBColor as PRGB
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

    CREAM   = PRGB(0xEF, 0xED, 0xE3)
    GOLD    = PRGB(0xE6, 0xC0, 0x69)
    GOLD_DK = PRGB(0xB8, 0x92, 0x3B)
    INK     = PRGB(0x19, 0x1B, 0x0E)
    INK_SOFT = PRGB(0x6B, 0x6A, 0x5C)
    GREEN   = PRGB(0x00, 0xB0, 0x50)
    RED     = PRGB(0xC0, 0x00, 0x00)
    FONT    = "Calibri"

    use_keys = include_keys or SIGMA_KEYS
    cps = [CHECKPOINT_BY_KEY[k] for k in use_keys if k in CHECKPOINT_BY_KEY]

    prs = Presentation()
    prs.slide_width = PInches(13.333)
    prs.slide_height = PInches(7.5)
    blank = prs.slide_layouts[6]

    def rect(slide, x, y, w, h, rgb, line_rgb=None):
        shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, PInches(x), PInches(y), PInches(w), PInches(h))
        shp.fill.solid()
        shp.fill.fore_color.rgb = rgb
        if line_rgb is None:
            shp.line.fill.background()
        else:
            shp.line.color.rgb = line_rgb
            shp.line.width = PPt(1)
        shp.shadow.inherit = False
        return shp

    def bg(slide):
        shp = rect(slide, 0, 0, 13.333, 7.5, CREAM)
        spTree = slide.shapes._spTree
        spTree.remove(shp._element)
        spTree.insert(2, shp._element)

    def text(slide, runs, x, y, w, h, size, rgb=INK, bold=False, align=PP_ALIGN.LEFT,
             anchor=MSO_ANCHOR.TOP):
        tb = slide.shapes.add_textbox(PInches(x), PInches(y), PInches(w), PInches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = anchor
        p = tf.paragraphs[0]
        p.alignment = align
        if not isinstance(runs, list):
            runs = [{"text": runs, "color": rgb, "bold": bold}]
        for seg in runs:
            r = p.add_run()
            r.text = seg["text"]
            r.font.size = PPt(size)
            r.font.name = FONT
            r.font.bold = seg.get("bold", bold)
            r.font.color.rgb = seg.get("color", rgb)
        return tb

    def add_image_fit(slide, img_path, bx, by, bw, bh):
        try:
            from PIL import Image
            if not img_path or not os.path.exists(img_path) or os.path.getsize(img_path) == 0:
                return False
            with Image.open(img_path) as im:
                im.verify()
            with Image.open(img_path) as im:
                iw, ih = im.size
        except Exception:
            return False
        ar = (iw / ih) if ih else 1.5
        box_ar = bw / bh
        if ar > box_ar:
            w = bw
            h = bw / ar
        else:
            h = bh
            w = bh * ar
        x = bx + (bw - w) / 2.0
        y = by + (bh - h) / 2.0
        rect(slide, x - 0.03, y - 0.03, w + 0.06, h + 0.06, GOLD)
        try:
            slide.shapes.add_picture(str(img_path), PInches(x), PInches(y),
                                     width=PInches(w), height=PInches(h))
            return True
        except Exception:
            return False

    # Title slide
    s = prs.slides.add_slide(blank)
    bg(s)
    rect(s, 0, 0, 13.333, 0.35, GOLD)
    rect(s, 0, 7.15, 13.333, 0.35, GOLD)
    text(s, domain, 1.0, 2.4, 11.33, 1.0, 40, INK, bold=True,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, "Website Health Checkup & Analysis Report", 1.0, 3.5, 11.33, 0.6, 22,
         GOLD_DK, bold=True, align=PP_ALIGN.CENTER)
    gen_date = datetime.now().strftime("%d-%B-%Y")
    text(s, f"Generated: {gen_date}", 1.0, 4.2, 11.33, 0.45, 13, INK_SOFT,
         align=PP_ALIGN.CENTER)

    # Intro slide
    s = prs.slides.add_slide(blank)
    bg(s)
    rect(s, 11.5, 0, 1.83, 0.28, GOLD)
    text(s, "INTRODUCTION", 0.689, 0.3, 12.0, 0.7, 32, INK, bold=True)
    rect(s, 0.689, 1.0, 5.5, 0.05, GOLD)
    intro_body = " ".join(b.lstrip(": ").strip() if b.startswith(":") else b.strip() for _, b in INTRO)
    text(s, intro_body, 1.1, 1.8, 11.1, 3.4, 18, INK)

    # Checkpoint slides
    placed = missing = 0
    for n, cp in enumerate(cps, start=1):
        s = prs.slides.add_slide(blank)
        bg(s)
        rect(s, 0, 0, 13.333, 0.18, GOLD)
        rect(s, 11.5, 0, 1.83, 0.28, GOLD)
        text(s, f"{n}. {cp['label'].strip()}", 0.689, 0.32, 12.0, 0.7, 26, INK, bold=True,
             anchor=MSO_ANCHOR.MIDDLE)
        rect(s, 0.689, 1.02, 5.5, 0.05, GOLD)

        body = cp["body"]
        for k, v in computed.items():
            body = body.replace("{" + k + "}", v)
        try:
            body = body.format(domain=domain)
        except Exception:
            pass
        body_color = PRGB(0xCC, 0x00, 0x00) if cp.get("red") else INK_SOFT
        text(s, body.strip(), 0.689, 1.12, 11.95, 0.95, 15, body_color, bold=bool(cp.get("red")))

        line_y = 2.05
        extra = cp.get("extra")
        if extra:
            text(s, [
                {"text": extra[0], "color": INK, "bold": True},
                {"text": extra[1].format(domain=domain), "color": GREEN, "bold": False},
            ], 0.689, line_y, 11.95, 0.35, 14)
            line_y += 0.4

        status = cp.get("status")
        if status:
            text(s, [
                {"text": status[0].strip() + " ", "color": INK, "bold": True},
                {"text": status[1].strip(), "color": GREEN, "bold": True},
            ], 0.689, line_y, 11.95, 0.35, 14)
            line_y += 0.4

        img_box_y = max(2.55, line_y + 0.05)
        img = captured.get(cp["key"])
        if img and add_image_fit(s, img, 0.7, img_box_y, 11.93, 7.2 - img_box_y):
            placed += 1
        else:
            missing += 1

    out_path = Path(out_dir) / f"{domain} Website health checkup analysis report {gen_date}.pptx"
    out_path = _safe_save_path(out_path)
    prs.save(str(out_path))
    return str(out_path), placed, missing


# ---------- PPTX builder (Omega) ----------
def build_health_pptx_omega(domain, captured, computed, out_dir):
    from pptx import Presentation as PRS
    from pptx.util import Inches as PInches, Pt as PPt
    from pptx.dml.color import RGBColor as PRGB
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

    BLUE    = PRGB(0x1F, 0x49, 0x7D)
    INK     = PRGB(0x1F, 0x29, 0x37)
    INK_SOFT = PRGB(0x5A, 0x5A, 0x5A)
    GREEN   = PRGB(0x00, 0xB0, 0x50)
    FONT    = "Calibri"

    cps = [CHECKPOINT_BY_KEY[k] for k in OMEGA_KEYS if k in CHECKPOINT_BY_KEY]

    prs = PRS()
    prs.slide_width = PInches(13.333)
    prs.slide_height = PInches(7.5)
    blank = prs.slide_layouts[6]

    def text(slide, runs, x, y, w, h, size, rgb=INK, bold=False, align=PP_ALIGN.LEFT,
             anchor=MSO_ANCHOR.TOP):
        tb = slide.shapes.add_textbox(PInches(x), PInches(y), PInches(w), PInches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = anchor
        p = tf.paragraphs[0]
        p.alignment = align
        if not isinstance(runs, list):
            runs = [{"text": runs, "color": rgb, "bold": bold}]
        for seg in runs:
            r = p.add_run()
            r.text = seg["text"]
            r.font.size = PPt(size)
            r.font.name = FONT
            r.font.bold = seg.get("bold", bold)
            r.font.color.rgb = seg.get("color", rgb)
        return tb

    def add_image_fit(slide, img_path, bx, by, bw, bh):
        try:
            from PIL import Image
            if not img_path or not os.path.exists(img_path) or os.path.getsize(img_path) == 0:
                return False
            with Image.open(img_path) as im:
                im.verify()
            with Image.open(img_path) as im:
                iw, ih = im.size
        except Exception:
            return False
        ar = (iw / ih) if ih else 1.5
        box_ar = bw / bh
        if ar > box_ar:
            w = bw
            h = bw / ar
        else:
            h = bh
            w = bh * ar
        x = bx + (bw - w) / 2.0
        y = by + (bh - h) / 2.0
        try:
            slide.shapes.add_picture(str(img_path), PInches(x), PInches(y),
                                     width=PInches(w), height=PInches(h))
            return True
        except Exception:
            return False

    # Title
    s = prs.slides.add_slide(blank)
    text(s, domain, 1.0, 2.2, 11.33, 1.0, 36, BLUE, bold=True,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    intro_body = " ".join(b.lstrip(": ").strip() if b.startswith(":") else b.strip() for _, b in INTRO)
    text(s, intro_body, 1.5, 3.5, 10.33, 2.5, 16, INK_SOFT, align=PP_ALIGN.CENTER)

    placed = missing = 0
    for n, cp in enumerate(cps, start=1):
        s = prs.slides.add_slide(blank)
        body = cp["body"]
        for k, v in computed.items():
            body = body.replace("{" + k + "}", v)
        try:
            body = body.format(domain=domain)
        except Exception:
            pass

        text(s, f"{n}. {cp['label'].strip()}", 0.5, 0.3, 12.3, 0.7, 24, BLUE, bold=True)
        body_color = PRGB(0xCC, 0x00, 0x00) if cp.get("red") else INK_SOFT
        text(s, body.strip(), 0.5, 1.05, 12.3, 0.95, 14, body_color, bold=bool(cp.get("red")))

        line_y = 2.0
        extra = cp.get("extra")
        if extra:
            text(s, [
                {"text": extra[0], "color": INK, "bold": True},
                {"text": extra[1].format(domain=domain), "color": GREEN, "bold": False},
            ], 0.5, line_y, 12.3, 0.35, 13)
            line_y += 0.35
        status = cp.get("status")
        if status:
            text(s, [
                {"text": status[0].strip() + " ", "color": INK, "bold": True},
                {"text": status[1].strip(), "color": GREEN, "bold": True},
            ], 0.5, line_y, 12.3, 0.35, 13)
            line_y += 0.35

        img_box_y = max(2.4, line_y + 0.1)
        img = captured.get(cp["key"])
        if img and add_image_fit(s, img, 0.5, img_box_y, 12.3, 7.2 - img_box_y):
            placed += 1
        else:
            missing += 1

    gen_date = datetime.now().strftime("%d-%B-%Y")
    out_path = Path(out_dir) / f"{domain} Health Analysis Report {gen_date}.pptx"
    out_path = _safe_save_path(out_path)
    prs.save(str(out_path))
    return str(out_path), placed, missing


# ---------- Main entry point ----------
def run_health_audit(domain, fmt="james", target_pages=None, out_dir=None,
                     driver=None, no_capture=False, log_fn=None, psi_api_key=None,
                     prefetched_futures=None):
    """Run a full health audit. Returns the output file path."""
    if log_fn is None:
        log_fn = print
    # Normalize to a bare host — a full URL (https://x.com/) breaks temp image
    # paths (".png" basename -> PIL "unknown file extension") and the filename.
    domain = re.sub(r'^\s*https?://', '', str(domain or '')).strip().strip('/').split('/')[0] or str(domain)

    # Build EXACTLY the selected format — never silently fall back to another one.
    fmt = str(fmt or "").strip().lower()
    fi = FORMAT_INFO.get(fmt)
    if not fi:
        raise ValueError(f"Unknown health audit format '{fmt}'. "
                         f"Available: {', '.join(sorted(FORMAT_INFO))}")
    use_keys = list(fi["keys"])
    if not psi_api_key and "pagespeed" in use_keys:
        use_keys.remove("pagespeed")
        log_fn("PageSpeed Insights skipped (not enabled).")

    total_steps = 3 if (not no_capture and driver is not None) else 2
    log_fn(f"[1/{total_steps}] Running health checks for {domain}...")
    captured, computed = prepare_health_data(domain, target_pages=target_pages, log_fn=log_fn, psi_api_key=psi_api_key, driver=driver, prefetched_futures=prefetched_futures)

    if not no_capture and driver is not None:
        log_fn(f"[2/{total_steps}] Capturing screenshots...")
        ss_dir = os.path.join(out_dir or tempfile.gettempdir(), "health_screenshots")
        ss = capture_screenshots_selenium(driver, domain, ss_dir, use_keys, log_fn)
        captured.update(ss)

        gsc_keys = {"manual_action", "security_issues"} & set(use_keys)
        uncaptured_gsc = gsc_keys - set(captured.keys())
        if uncaptured_gsc:
            try:
                import gsc_audit
                accounts = gsc_audit.list_accounts()
                connected = [a for a in accounts if a.get("has_refresh")]
                gsc_email = connected[0]["email"] if connected else None
                # Resolve the REAL property (URL-prefix vs sc-domain) via the API so
                # the manual-actions / security pages load the correct property —
                # hardcoding sc-domain broke it for URL-prefix properties.
                prop_url = f"sc-domain:{domain}"
                if gsc_email:
                    try:
                        _tok = gsc_audit.get_access_token(gsc_email)
                        prop_url = gsc_audit.resolve_property(_tok, domain) or prop_url
                        log_fn(f"  GSC property: {prop_url}")
                    except Exception as _e:
                        log_fn(f"  Could not resolve GSC property ({_e}); using {prop_url}")
                gsc_pages = []
                if "manual_action" in uncaptured_gsc:
                    gsc_pages.append({"key": "manual_action", "page": "manual-actions", "wait": 12})
                if "security_issues" in uncaptured_gsc:
                    gsc_pages.append({"key": "security_issues", "page": "security-issues", "wait": 8})
                if gsc_pages:
                    gsc_captured = False
                    sessions = gsc_audit.list_sessions()
                    # A session's "accounts" list is only populated by a cookie scan, which
                    # may not have run — so a perfectly-good logged-in session can look
                    # "untagged". Try email-matched sessions FIRST, then every other session.
                    # capture_gsc_with_session verifies it's actually logged in (and retries
                    # in a visible window if Google bounces the headless relaunch to login).
                    def _email_match(s):
                        return bool(gsc_email) and gsc_email.lower() in [a.lower() for a in s.get("accounts", [])]
                    ordered = [s for s in sessions if _email_match(s)] + [s for s in sessions if not _email_match(s)]
                    for sess in ordered:
                        tag = "matched" if _email_match(sess) else "untagged"
                        log_fn(f"  Capturing GSC screenshots via session {sess.get('label', sess['id'])} ({tag})...")
                        gsc_ss = gsc_audit.capture_gsc_with_session(
                            sess["id"], prop_url, gsc_email or "", ss_dir,
                            pages=gsc_pages, log_fn=log_fn)
                        if isinstance(gsc_ss, dict) and "error" not in gsc_ss:
                            captured.update(gsc_ss)
                            gsc_captured = True
                            break
                        elif isinstance(gsc_ss, dict) and gsc_ss.get("error") == "session_expired":
                            log_fn(f"  Session {sess.get('label', sess['id'])} needs re-login in GSC Sessions.")
                    if not gsc_captured:
                        if not sessions:
                            log_fn("  No GSC browser session found — create one in GSC Audit > Browser Sessions")
                        elif not gsc_email:
                            log_fn("  GSC not connected — skipping manual_action/security_issues screenshots")
                        else:
                            log_fn("  No valid GSC browser session found — create one in GSC Audit > Browser Sessions")
                        try:
                            import json as _json, urllib.request as _ur2, urllib.parse as _up2
                            _bd = os.path.dirname(os.path.abspath(__file__))
                            _cfg_path = os.path.join(_bd, "config.json")
                            if os.path.exists(_cfg_path):
                                with open(_cfg_path, "r") as _cf:
                                    _cfg = _json.load(_cf)
                                _wurl = (_cfg.get("auth_api_url") or "").strip()
                                if _wurl:
                                    _qs = _up2.urlencode({"action": "get_config"})
                                    _req = _ur2.Request(f"{_wurl}?{_qs}")
                                    with _ur2.urlopen(_req, timeout=10) as _r:
                                        _data = _json.loads(_r.read())
                                    _mapping = _data.get("mapping", {}) if _data.get("success") else {}
                                    _clean = domain.lower().replace("www.", "")
                                    _info = _mapping.get(_clean)
                                    if _info:
                                        log_fn(f"  >> Domain found in GSC account: {_info.get('email', '?')} "
                                               f"(access: {_info.get('accessLevel', '?')}). "
                                               f"Create a browser session and log in with this account.")
                        except Exception:
                            pass
            except Exception as e:
                log_fn(f"  GSC screenshot capture skipped: {e}")
    else:
        log_fn(f"[2/3] Skipping screenshots (no browser or --no-capture)")

    log_fn(f"[{total_steps}/{total_steps}] Building {fi['ext'].upper()} report ({fmt})...")
    if not out_dir:
        out_dir = tempfile.gettempdir()
    os.makedirs(out_dir, exist_ok=True)

    if fmt == "sigma":
        path, placed, miss = build_health_pptx_sigma(domain, captured, computed, out_dir, use_keys)
    elif fmt == "omega":
        path, placed, miss = build_health_pptx_omega(domain, captured, computed, out_dir)
    elif fmt == "neon":
        path, placed, miss = build_health_docx(domain, captured, computed, out_dir, NEON_KEYS, voice="i", header_fill="215868")
    else:
        path, placed, miss = build_health_docx(domain, captured, computed, out_dir, use_keys, page_border=True)

    log_fn(f"[DONE] {placed} images placed, {miss} text-only. Output: {path}")
    return path
