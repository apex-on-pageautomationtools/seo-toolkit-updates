"""
generate_index_coverage_report.py - Turns Search Console's Page Indexing report
into the client-ready "Index Coverage Report" workbook the team builds by hand
(reference: Index-coverage-report-<domain>.xlsx).

The official Search Console API does not expose the Page Indexing report's
per-reason URL lists at all - this is a real, documented gap, not something
this script works around by choice. The URL lists have to come from somewhere
that already scraped them: gsc_audit.py's capture_index_coverage_urls()
(Selenium, driving the same already-authenticated GSC browser session every
other GSC tool in this app already uses) is the intended source, but this
module only cares about the shape: {reason_name: [url, ...], ...}. Anything
that can produce that dict can feed this.

For every URL, this:
  - Checks its REAL current HTTP status + follows redirects (Search Console's
    own reason was recorded whenever it last crawled - can be stale; a page
    listed under "Crawled - currently not indexed" may actually 301 or 404
    TODAY, which changes what the real fix is).
  - Classifies Indexability / Indexability Status from that live check (not
    invented - "Redirected" only when the live check found a redirect,
    "Client Error" only from a real 4xx, etc.).
  - For "Not found (404)" specifically, suggests a real redirect target - the
    site's own sitemap's best-matching live URL by slug similarity, never a
    fabricated URL. No match found -> left blank, not guessed.

Output: <domain>'s Index Coverage Report.xlsx - "Index" summary tab (reasons
that need real action, with a one-line fix each) + one tab per reason
actually present (title + definition row + data), matching the reference
workbook's layout exactly.

Run:
    python generate_index_coverage_report.py --domain example.com \
        --reason-urls reason_urls.json --out "Index Coverage Report.xlsx"
"""
import os
import re
import sys
import json
import argparse
import urllib.parse
import urllib.request
import urllib.error
from pathlib import Path

ROOT = Path(__file__).parent.resolve()
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

import generate_seo_onpage_phase2 as onpage2   # live page fetch (title/canonical/meta robots)
import generate_geo_report as georpt           # get_sitemap_urls - real sitemap discovery


def log(msg):
    try:
        print(msg, flush=True)
    except UnicodeEncodeError:
        enc = sys.stdout.encoding or "ascii"
        print(str(msg).encode(enc, errors="replace").decode(enc), flush=True)


# --------------------------------------------------------------------------- #
# Reason definitions - Google's own documented Page Indexing reasons (Search
# Console Help: "Why pages aren't indexed"), not just the subset that happened
# to show up in any one reference report. A reason not in this table (Google
# adds new ones occasionally) still gets its own tab with a generic fallback
# definition, same as the reference workbook does for its one non-standard
# reason - never silently dropped.
# --------------------------------------------------------------------------- #
# needs_action: whether this reason belongs in the "Index" summary tab at all
# (healthy/expected-by-design reasons like "Page with redirect" or "Alternate
# page with proper canonical tag" still get their own detail tab, just not a
# row in the action-needed summary - matches the reference workbook, whose
# Index tab lists only 4 of its 9 reasons).
REASON_INFO = {
    "Discovered - currently not indexed": {
        "definition": "Google has found this URL (via a link or sitemap) but has not crawled it yet, "
                      "so it cannot appear in search. Often a sign of crawl-budget limits, weak internal "
                      "linking, or thin value.",
        "needs_action": True,
        "fix_summary": "Strengthen internal linking to these pages and confirm they carry real value - "
                       "weak/thin pages rarely get crawled.",
        "default_action": "Improve internal linking and page value so Google prioritizes crawling this URL.",
    },
    "Crawled - currently not indexed": {
        "definition": "Google crawled the page but decided not to index it, so it will not appear in "
                      "search. Usually means the content is thin, near-duplicate, or low value in "
                      "Google's judgement.",
        "needs_action": True,
        "fix_summary": "Google crawled the page but chose not to index it. Enhance the content to make "
                       "it more useful, original, and distinct from other pages.",
        "default_action": "Need to index this page",
    },
    "Soft 404": {
        "definition": "The server returns 200 OK, telling Google the page works, but the content is "
                      "empty, thin, or reads like 'not found', so Google sees no real value. Common "
                      "causes: blank or thin pages (empty tag pages, out-of-stock products with no "
                      "alternatives), deleted URLs redirected to the homepage instead of returning 404, "
                      "or content hidden from bots by JavaScript or CSS. Fix it by enriching the content "
                      "if the page should live, or removing it properly with a 404 or 410, or a 301 to "
                      "a relevant page.",
        "needs_action": True,
        "fix_summary": "Enrich the content if the page should exist, or remove it properly (404/410, "
                       "or a 301 to a relevant page).",
        "default_action": "Enrich the content, or remove it properly (404/410, or 301 to a relevant page)",
    },
    "Not found (404)": {
        "definition": "The URL returns 404 Not Found, so it cannot be indexed or ranked. Any inbound "
                      "links and past ranking value are lost unless the URL is redirected to a relevant "
                      "live page.",
        "needs_action": True,
        "fix_summary": "The URL returns 404. Redirect it to the closest relevant live page, or restore "
                       "the page if it should exist.",
        "default_action": None,   # per-row: filled from the real redirect suggestion, see build_report()
    },
    "Not found (410)": {
        "definition": "The URL deliberately returns 410 Gone, telling Google the page was intentionally "
                      "and permanently removed. Stronger than a 404 - only use this when the page should "
                      "never come back.",
        "needs_action": False,
        "fix_summary": "Working as intended if this removal was deliberate; redirect it instead if the "
                       "content moved rather than disappeared.",
        "default_action": "No action needed if deliberate - otherwise redirect to a relevant live page",
    },
    "Duplicate without user-selected canonical": {
        "definition": "Google found multiple near-identical URLs and no canonical tag told it which one "
                      "you prefer, so it picked one itself (which may not be the one you want ranked).",
        "needs_action": True,
        "fix_summary": "Add a self-referencing canonical tag on the version you want indexed.",
        "default_action": "Add a canonical tag pointing to the preferred version of this page",
    },
    "Duplicate, Google chose different canonical than user": {
        "definition": "You set a canonical but Google overrode it and chose a different page. Usually "
                      "Google trusts another version more; if your choice should win, strengthen its "
                      "internal links, sitemap entry, and canonical.",
        "needs_action": False,
        "fix_summary": "No action needed unless the page Google chose isn't the one you actually want "
                       "ranked - if so, strengthen the preferred version's internal links and sitemap entry.",
        "default_action": "No Action Needed",
    },
    "Duplicate, submitted URL not selected as canonical": {
        "definition": "You submitted this exact URL (e.g. via sitemap), but Google indexed a different "
                      "duplicate/variant of it as the canonical instead.",
        "needs_action": False,
        "fix_summary": "No action needed unless the canonical Google picked isn't the version you want "
                       "ranked - if so, strengthen this URL's own signals instead.",
        "default_action": "No Action Needed",
    },
    "Alternate page with proper canonical tag": {
        "definition": "This is a duplicate or variant (such as a paginated or parameter URL) that "
                      "correctly points to its canonical version, so Google indexes the canonical "
                      "instead. Working as intended.",
        "needs_action": False,
        "fix_summary": "Working as intended - no action needed.",
        "default_action": "No action needed",
    },
    "Blocked by robots.txt": {
        "definition": "robots.txt tells Googlebot not to crawl this URL, so Google cannot read it (it "
                      "may still show as a bare link). Fine when intentional; a problem only if it "
                      "blocks pages you want ranked.",
        "needs_action": False,
        "fix_summary": "No action needed for pages you don't want crawled - otherwise remove the "
                       "robots.txt rule blocking this URL.",
        "default_action": "Not Important Page - No action needed",
    },
    "Blocked due to unauthorized request (401)": {
        "definition": "The page returned 401 Unauthorized to Googlebot, so it couldn't be crawled or "
                      "indexed. Usually login-gated content, or a WAF/bot-check wrongly challenging "
                      "Googlebot's own requests.",
        "needs_action": True,
        "fix_summary": "Remove the login/auth requirement for pages you want indexed, or confirm this "
                       "page should genuinely stay gated.",
        "default_action": "Check whether Googlebot should have access to this page",
    },
    "Blocked due to access forbidden (403)": {
        "definition": "The page returned 403 Forbidden to Googlebot, so it couldn't be crawled or "
                      "indexed. Often a firewall/bot-protection rule blocking crawlers, not just users.",
        "needs_action": True,
        "fix_summary": "Whitelist Googlebot in the firewall/bot-protection rule blocking this page, if "
                       "the page should be indexed.",
        "default_action": "Check server/firewall rules blocking Googlebot from this page",
    },
    "Blocked due to other 4xx issue": {
        "definition": "The page returned some other 4xx client-error status (not 401/403/404) to "
                      "Googlebot, so it couldn't be indexed.",
        "needs_action": True,
        "fix_summary": "Inspect the URL directly to see the exact status code and fix the server-side "
                       "issue causing it.",
        "default_action": "Inspect the URL in Search Console for the exact error and fix it",
    },
    "Blocked by page removal tool": {
        "definition": "Someone used Search Console's Removals tool to temporarily hide this URL from "
                      "search results. Temporary by design - it doesn't remove the page from the index "
                      "permanently.",
        "needs_action": False,
        "fix_summary": "No action needed if this removal was intentional - it expires automatically "
                       "after ~6 months unless renewed.",
        "default_action": "No action needed if intentional",
    },
    "Page with redirect": {
        "definition": "This URL redirects to another URL, so the redirecting URL is not indexed; the "
                      "destination is what can rank. Normal for moved or canonical URLs. Only redirect "
                      "chains (several hops) or loops hurt performance and need fixing.",
        "needs_action": False,
        "fix_summary": "Working as intended for a normal single-hop redirect - only fix if this is part "
                       "of a redirect chain or loop.",
        "default_action": "No Action Needed",
    },
    "Redirect error": {
        "definition": "Googlebot could not follow this URL's redirect - a redirect loop, chain too long, "
                      "a redirect to an empty URL, or the redirect URL exceeded the max length.",
        "needs_action": True,
        "fix_summary": "Fix the redirect - remove loops/long chains and redirect straight to the final "
                       "live destination in one hop.",
        "default_action": "Fix the redirect to point directly at the final destination in one hop",
    },
    "Server error (5xx)": {
        "definition": "The server returned a 5xx error when Googlebot requested this URL, so it "
                      "couldn't be crawled or indexed. Usually a server-side problem, not the page "
                      "itself.",
        "needs_action": True,
        "fix_summary": "Investigate the server error (check server logs / hosting) - this is a hosting "
                       "issue, not a content issue.",
        "default_action": "Investigate the server error with your hosting provider",
    },
    "Excluded by 'noindex' tag": {
        "definition": "A meta robots or X-Robots-Tag 'noindex' directive on this page explicitly tells "
                      "Google not to index it. Working as intended if deliberate; a real problem if this "
                      "page should actually rank.",
        "needs_action": True,
        "fix_summary": "Remove the noindex tag if this page should be indexed and ranked; leave it if "
                       "the exclusion is intentional.",
        "default_action": "Need to remove noindex tag from important pages",
    },
    "Not indexed due to legal removal": {
        "definition": "This URL was removed from Google's index due to a legal request (e.g. a valid "
                      "DMCA takedown), not a technical or content issue.",
        "needs_action": False,
        "fix_summary": "No action available from the site side - this is a legal removal on Google's end.",
        "default_action": "No action available - legal removal",
    },
}

# Fallback for any reason Google shows that isn't in the table above (new
# reason types, or a name variant) - same generic text the reference workbook
# itself uses, never silently skipped.
_GENERIC_REASON_INFO = {
    "definition": "This reason is not in the standard list. Inspect the URL in Search Console to see "
                  "Google's own explanation, then act accordingly.",
    "needs_action": True,
    "fix_summary": "Inspect a sample URL in Search Console's URL Inspection tool for the specific reason.",
    "default_action": "Inspect in Search Console and act accordingly",
}


def _reason_info(reason):
    return REASON_INFO.get(reason, _GENERIC_REASON_INFO)


# --------------------------------------------------------------------------- #
# Live per-URL status/redirect check - never trusts GSC's own (possibly
# stale) recorded reason on its own; every URL gets checked fresh.
# --------------------------------------------------------------------------- #
_UA = "Mozilla/5.0 IndexCoverageReportBot"


def _safe_url(url):
    try:
        parts = urllib.parse.urlsplit(url)
        path = urllib.parse.quote(parts.path, safe="/%")
        query = urllib.parse.quote(parts.query, safe="=&%")
        return urllib.parse.urlunsplit((parts.scheme, parts.netloc, path, query, parts.fragment))
    except Exception:
        return url


def check_live_status(url):
    """Real current HTTP status + final URL after following redirects. Never
    raises - a network failure reports as status=None rather than crashing
    the whole run over one bad URL."""
    try:
        req = urllib.request.Request(_safe_url(url), headers={"User-Agent": _UA}, method="GET")
        with urllib.request.urlopen(req, timeout=15) as r:
            return {"status": r.status, "final_url": r.geturl(), "html": r.read().decode("utf-8", "ignore")}
    except urllib.error.HTTPError as e:
        return {"status": e.code, "final_url": e.geturl() if hasattr(e, "geturl") else url, "html": ""}
    except Exception as e:
        log(f"   [warn] could not check {url}: {type(e).__name__}: {e}")
        return {"status": None, "final_url": url, "html": ""}


def classify_indexability(url, live):
    """(Indexability, Indexability Status) columns - derived from the REAL
    live check just performed, never from GSC's own possibly-stale reason."""
    status = live.get("status")
    final_url = live.get("final_url") or url
    if status is None:
        return "Unknown", "Could not check - please verify manually"
    if 300 <= status < 400 or (final_url and final_url.rstrip("/") != url.rstrip("/")):
        return "Non-Indexable", "Redirected"
    if status in (401, 403):
        return "Non-Indexable", "Blocked"
    if status == 404 or status == 410:
        return "Non-Indexable", "Client Error"
    if status >= 500:
        return "Non-Indexable", "Server Error"
    if status == 200:
        html = live.get("html") or ""
        pd = onpage2._parse_html(html, final_url, status) if html else None
        if pd:
            robots_meta = re.search(r'<meta[^>]+name=["\']robots["\'][^>]+content=["\']([^"\']*noindex[^"\']*)',
                                    html, re.I)
            if robots_meta:
                return "Non-Indexable", "noindex"
            canon = pd.get("canonical") or ""
            m = re.search(r'href="([^"]+)"', canon)
            if m and m.group(1).rstrip("/") != url.rstrip("/"):
                return "Non-Indexable", "Canonicalised"
        return "Indexable", "-"
    return "Unknown", "Could not check - please verify manually"


_slug_word_re = re.compile(r"[a-z0-9]+")


def _slug_words(url):
    path = urllib.parse.urlparse(url).path
    return set(_slug_word_re.findall(path.lower()))


def suggest_redirect_target(dead_url, sitemap_urls, homepage_url):
    """Best-matching REAL live page for a 404'd URL, by slug word overlap
    against the site's own sitemap - never a fabricated URL. Falls back to
    the homepage only when no sitemap match scores above the minimum
    threshold, exactly like a human triaging redirects would."""
    dead_words = _slug_words(dead_url)
    if not dead_words or not sitemap_urls:
        return homepage_url
    best_url, best_score = None, 0
    for cand in sitemap_urls:
        cand_words = _slug_words(cand)
        if not cand_words:
            continue
        overlap = len(dead_words & cand_words)
        score = overlap / max(len(dead_words), 1)
        if score > best_score:
            best_score, best_url = score, cand
    return best_url if best_score >= 0.34 else homepage_url


# --------------------------------------------------------------------------- #
# Output workbook - matches the reference layout: title row, definition row,
# header row, then data - per reason tab, plus an "Index" summary tab.
# --------------------------------------------------------------------------- #
TITLE_FONT = Font(bold=True, size=13)
HEADER_FILL = PatternFill("solid", fgColor="2F5496")
HEADER_FONT = Font(bold=True, color="FFFFFF")
WRAP = Alignment(horizontal="left", vertical="top", wrap_text=True)
NO_WRAP = Alignment(horizontal="left", vertical="center", wrap_text=False)


def _safe_sheet_name(name, used):
    safe = re.sub(r"[\[\]:\\/?*]", "_", name)[:31] or "Sheet"
    base, n = safe, 1
    while safe in used:
        n += 1
        safe = f"{base[:28]}_{n}"
    used.add(safe)
    return safe


def build_report(domain, reason_rows, out_path, brand=None):
    """reason_rows: {reason: [{"url":, "status":, "indexability":,
    "indexability_status":, "action":}, ...], ...} - already live-checked
    (see process_reason() below, which is how main() builds this)."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    used_names = set()

    # --- Index summary tab (action-needed reasons only, matching reference) ---
    ws = wb.create_sheet(_safe_sheet_name("Index", used_names))
    ws.append(["Index Coverage Report"])
    ws["A1"].font = TITLE_FONT
    ws.append([domain])
    ws.append(["Every page-indexing reason from Search Console except healthy indexed pages. "
               "Each tab opens with a definition of the reason."])
    ws.append(["Reason", "Action", "Recommended fix"])
    for c in ws[4]:
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
    for reason in reason_rows:
        info = _reason_info(reason)
        if info["needs_action"]:
            ws.append([reason, "Action needed", info["fix_summary"]])
    for col, width in zip("ABC", (45, 16, 90)):
        ws.column_dimensions[col].width = width
    for row in ws.iter_rows(min_row=5):
        for cell in row:
            cell.alignment = WRAP

    # --- One tab per reason actually found ---
    for reason, rows in reason_rows.items():
        info = _reason_info(reason)
        sheet_name = _safe_sheet_name(reason, used_names)
        ws = wb.create_sheet(sheet_name)
        ws.append([reason])
        ws["A1"].font = TITLE_FONT
        ws.append([info["definition"]])
        ws["A2"].alignment = WRAP
        ws.row_dimensions[2].height = 60

        last_col_label = "Recommended Redirect URL" if reason == "Not found (404)" else "Action"
        headers = ["Address", "Status Code", "Indexability", "Indexability Status", last_col_label]
        ws.append(headers)
        for c in ws[3]:
            c.fill = HEADER_FILL
            c.font = HEADER_FONT

        for r in rows:
            ws.append([r["url"], r.get("status") or "-", r.get("indexability", "-"),
                      r.get("indexability_status", "-"), r.get("action", "")])

        for i, h in enumerate(headers, 1):
            col_letter = ws.cell(3, i).column_letter
            ws.column_dimensions[col_letter].width = 60 if i == 1 else (22 if i == 5 else 16)
        for row in ws.iter_rows(min_row=4):
            for cell in row:
                cell.alignment = NO_WRAP
        ws.freeze_panes = "A4"

    wb.save(out_path)
    log(f"[DONE] {out_path}")


def process_reason(reason, urls, domain, sitemap_urls, homepage_url):
    """Live-check every URL for one reason and build its row list - the
    shared per-reason pipeline main() and any other caller (e.g. a future
    web_app_batch.py route) should both use, so the live-check/redirect-
    suggestion logic never has to be duplicated."""
    info = _reason_info(reason)
    out = []
    for i, url in enumerate(urls, 1):
        log(f"   [{i}/{len(urls)}] Checking {url}")
        live = check_live_status(url)
        indexability, indexability_status = classify_indexability(url, live)
        if reason == "Not found (404)":
            action = (suggest_redirect_target(url, sitemap_urls, homepage_url)
                      if live.get("status") == 404 else "No action needed - page is live")
        else:
            action = info["default_action"] or "Check manually"
        out.append({"url": url, "status": live.get("status"), "indexability": indexability,
                   "indexability_status": indexability_status, "action": action})
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--domain", required=True)
    ap.add_argument("--reason-urls", required=True,
                    help="JSON file: {reason_name: [url, ...], ...} - see gsc_audit.capture_index_coverage_urls()")
    ap.add_argument("--out", required=True)
    ap.add_argument("--brand", default=None)
    args = ap.parse_args()

    with open(args.reason_urls, "r", encoding="utf-8") as f:
        reason_urls = json.load(f)

    log(f"Discovering {args.domain}'s real sitemap for redirect suggestions...")
    sitemap_urls = georpt.get_sitemap_urls(args.domain, cap=2000)
    log(f"   {len(sitemap_urls)} sitemap URL(s) found.")
    homepage_url = f"https://{onpage2.safe_domain(args.domain)}/"

    onpage2.set_run_scale(sum(len(v) for v in reason_urls.values()))

    reason_rows = {}
    for i, (reason, urls) in enumerate(reason_urls.items(), 1):
        log(f"[{i}/{len(reason_urls)}] '{reason}' ({len(urls)} URL(s))...")
        reason_rows[reason] = process_reason(reason, urls, args.domain, sitemap_urls, homepage_url)

    build_report(args.domain, reason_rows, args.out, brand=args.brand)


if __name__ == "__main__":
    main()
