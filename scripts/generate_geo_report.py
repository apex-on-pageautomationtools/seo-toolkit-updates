"""
GEO / AI Optimization report generator - the "apart from performance" core of the
monthly GEO deliverable set (see the real reference: onlinesafetytraining.ca month 7).

Given a domain + target pages, this:
  1. Crawls each target page's real content (reusing generate_seo_onpage_phase2's crawler)
  2. Generates AI-searchable keywords + long-tail/NLP FAQ queries + answers, grounded in
     that page's actual content (via Gemini, same free-tier pattern as suggest_meta())
  3. For each generated query, checks - WITHOUT login, WITHOUT any paid API - whether the
     site is already being cited in real AI-answer results for that query:
       - Perplexity's public search (no login required) - reliable, real answer text +
         cited source domains extracted from the live page.
       - Google's AI Overview - best-effort only. Headless automated requests get a
         captcha/"unusual traffic" wall from Google almost every time (confirmed live,
         same limitation the existing SERP screenshot in generate_seo_onpage_phase2.py
         already has to skip past) - so this is attempted but not relied on.
       - ChatGPT (chatgpt.com) - best-effort only. Its real chat interface requires a
         logged-in account to get an actual answer (confirmed live) - anonymous access
         is not reliably available, so this is attempted and skipped gracefully rather
         than faked.
  4. Writes a "Keywords and Content (FAQs) Work for GEO" xlsx matching the real
     reference's columns, plus extra columns showing which terms the site is ALREADY
     being cited for right now (an "opportunity" vs "already appearing" signal).

This intentionally does NOT build the Performance Report (rank tracking / Ahrefs /
GA4 screenshots) - that is separate, bigger scope (Ahrefs + GSC + rank-tracker
integration) to be scoped once this core piece is tested.
"""
import argparse
import datetime
import json
import os
import re
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))
import generate_seo_onpage_phase2 as op  # reuse crawler, AI helper, browser driver

OUTPUT_DIR = ROOT.parent / "output"


def log(msg):
    print(msg, flush=True)


# --------------------------------------------------------------------------- #
# Step 1: content-grounded AI-searchable keyword + FAQ generation
# --------------------------------------------------------------------------- #
def generate_ai_faqs(page_data, brand, n=5, seed_keywords=None):
    """AI-searchable keyword + long-tail/NLP FAQ query + answer, per page - grounded
    in that page's own crawled content (title/H1/body text), matching the real
    reference xlsx's columns exactly: AI-Searchable Keyword | Long-Tail/NLP Query
    (FAQ) | Answers. Falls back to a simple heuristic (from title/H1) if no Gemini
    key is configured or the call fails - same fallback philosophy as suggest_meta().

    seed_keywords: optional list of keywords the team already targets for this
    site (given up front instead of only auto-discovered from content) - when
    given, generation is steered to build FAQs around THOSE terms first, so the
    team doesn't have to separately go find AI-searchable phrasing themselves."""
    title = page_data.get("title") or ""
    h1 = page_data.get("h1") or ""
    body = (page_data.get("body_text") or "")[:2500]
    url = page_data["url"]

    result = None
    if body or title or h1:
        seed_line = ""
        if seed_keywords:
            seed_line = (
                "\nThe team already targets these existing SEO keywords for this site - "
                "ground the entries around these terms where relevant to this page, rather "
                f"than only inventing new topics: {', '.join(seed_keywords)}\n"
            )
        prompt = (
            "You are an AI-search (GEO/answer-engine) optimization specialist writing "
            f"content for {brand}'s page at {url}.\n"
            f"Page title: {title}\nPage H1: {h1}\n"
            f"Page content excerpt: {body or '(not available)'}\n"
            f"{seed_line}\n"
            f"Generate exactly {n} distinct entries. Each entry needs:\n"
            '- "keyword": a short AI-searchable keyword phrase someone would type into '
            "ChatGPT/Perplexity/Google AI Overview to find this page's topic\n"
            '- "query": a natural-language, long-tail question phrased the way a real '
            "person would ask an AI assistant (not a bare keyword)\n"
            '- "answer": a 2-4 sentence answer to that question, written from THIS '
            "page's actual content (not generic filler), naturally mentioning the "
            f"brand ({brand}) once\n\n"
            'Return ONLY a JSON array: [{"keyword": "...", "query": "...", "answer": "..."}, ...]'
        )
        result = op._ai_suggest(prompt)

    if isinstance(result, list) and result:
        out = []
        for item in result[:n]:
            if isinstance(item, dict) and item.get("query") and item.get("answer"):
                out.append({
                    "keyword": str(item.get("keyword", "")).strip(),
                    "query": str(item["query"]).strip(),
                    "answer": str(item["answer"]).strip(),
                })
        if out:
            return out

    # Heuristic fallback - no Gemini key, or the call failed.
    topic = h1 if h1 and h1 != op.MISSING else (title if title and title != op.MISSING else "this page")
    return [{
        "keyword": topic.lower(),
        "query": f"What should I know about {topic}?",
        "answer": f"{brand} covers {topic} on this page - see {url} for full details.",
    }]


# --------------------------------------------------------------------------- #
# Step 2: no-login AI-visibility checks
# --------------------------------------------------------------------------- #
def _domain_cited(text, domain):
    if not text:
        return False
    root = domain[4:] if domain.startswith("www.") else domain
    brand_token = root.split(".")[0]
    return root.lower() in text.lower() or (len(brand_token) > 3 and brand_token.lower() in text.lower())


def _check_perplexity(driver, query, domain):
    """Perplexity's public search - no login required, confirmed live and reliable.
    Returns {checked, cited, snippet}."""
    try:
        driver.get(f"https://www.perplexity.ai/search?q={query.replace(' ', '+')}")
        # Poll for the answer to actually render instead of a fixed sleep - same
        # hydration-wait idea as generate_seo_onpage_phase2's _wait_settled().
        text = ""
        for _ in range(20):
            time.sleep(1)
            try:
                text = driver.execute_script("return document.body.innerText;") or ""
            except Exception:
                continue
            if len(text) > 800 and "Searching the web" not in text[-200:]:
                break
        cited = _domain_cited(text, domain)
        return {"engine": "perplexity", "checked": True, "cited": cited,
                "snippet": text[:400]}
    except Exception as e:
        return {"engine": "perplexity", "checked": False, "cited": False,
                "snippet": f"error: {type(e).__name__}"}


def _check_google_ai_overview(driver, query, domain):
    """Google AI Overview - BEST EFFORT ONLY. Headless requests get a captcha/
    'unusual traffic' wall from Google almost every time (confirmed live) - this is
    the exact same limitation the existing SERP screenshot already has to skip past
    in generate_seo_onpage_phase2.capture_onpage_screenshots(). Reports checked=False
    (not a false negative) whenever blocked, rather than pretending it ran."""
    try:
        url = f"https://www.google.com/search?q={query.replace(' ', '+')}&hl=en&gl=us"
        driver.get(url)
        time.sleep(3)
        src = (driver.page_source or "").lower()
        if any(m in src for m in ("unusual traffic", "not a robot", "recaptcha",
                                   "/sorry/", "detected unusual", "before you continue")):
            return {"engine": "google_ai_overview", "checked": False, "cited": False,
                    "snippet": "blocked (captcha)"}
        has_overview = any(m in src for m in ("ai overview", "ai-generated", "generative ai is experimental"))
        if not has_overview:
            return {"engine": "google_ai_overview", "checked": True, "cited": False,
                    "snippet": "no AI Overview shown for this query"}
        cited = _domain_cited(src, domain)
        return {"engine": "google_ai_overview", "checked": True, "cited": cited, "snippet": ""}
    except Exception as e:
        return {"engine": "google_ai_overview", "checked": False, "cited": False,
                "snippet": f"error: {type(e).__name__}"}


def _check_chatgpt_anon(driver, query, domain):
    """ChatGPT (chatgpt.com) - BEST EFFORT ONLY. Its real chat interface requires a
    logged-in account to get an actual answer (confirmed live) - this detects the
    login wall and reports checked=False rather than faking a result."""
    try:
        driver.get("https://chatgpt.com/")
        time.sleep(3)
        src = (driver.page_source or "").lower()
        if any(m in src for m in ("log in", "sign up", "stay logged out")):
            return {"engine": "chatgpt", "checked": False, "cited": False,
                    "snippet": "login required - not available anonymously"}
        # If no login wall is detected, this deployment allows anonymous chat -
        # not currently reachable to verify further without a live account, so
        # still report unchecked rather than guessing at a selector.
        return {"engine": "chatgpt", "checked": False, "cited": False,
                "snippet": "anonymous chat UI present but not yet wired up"}
    except Exception as e:
        return {"engine": "chatgpt", "checked": False, "cited": False,
                "snippet": f"error: {type(e).__name__}"}


def check_ai_visibility(driver, query, domain):
    """Runs all three engine checks for one query. Perplexity is the reliable one;
    Google/ChatGPT are best-effort and commonly report checked=False."""
    return [
        _check_perplexity(driver, query, domain),
        _check_google_ai_overview(driver, query, domain),
        _check_chatgpt_anon(driver, query, domain),
    ]


# --------------------------------------------------------------------------- #
# Step 3: combine into the reference-matching xlsx
# --------------------------------------------------------------------------- #
def build_geo_keywords_xlsx(domain, pages_data, brand, out_path, check_visibility=True, seed_keywords=None):
    """seed_keywords: optional list of SEO keywords the team already has for this
    domain (given up front, e.g. from their existing keyword research) - these
    steer the per-page FAQ generation toward those terms AND are checked directly
    for AI-citation as their own rows, so the team can see right away whether a
    keyword they already have is already showing up in AI answers, instead of
    only getting brand-new AI-generated terms."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment

    wb = Workbook()
    ws = wb.active
    ws.title = "Keywords"
    headers = ["URL", "AI-Searchable Keyword", "Long-Tail/NLP Query (FAQ)", "Answers", "Source"]
    if check_visibility:
        headers += ["Cited on Perplexity", "Cited on Google AI Overview", "Cited on ChatGPT"]
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F3864")
        cell.alignment = Alignment(wrap_text=True, vertical="center")

    driver = op._get_op_driver() if check_visibility else None

    def _visibility_cells(query):
        if not (check_visibility and driver):
            return ["Not checked", "Not checked", "Not checked"] if check_visibility else []
        log(f"   -> checking AI visibility: {query[:70]}")
        results = check_ai_visibility(driver, query, domain)
        by_engine = {r["engine"]: r for r in results}
        return [
            "Yes" if by_engine["perplexity"]["cited"] else ("Not checked" if not by_engine["perplexity"]["checked"] else "No"),
            "Yes" if by_engine["google_ai_overview"]["cited"] else ("Not checked" if not by_engine["google_ai_overview"]["checked"] else "No"),
            "Yes" if by_engine["chatgpt"]["cited"] else ("Not checked" if not by_engine["chatgpt"]["checked"] else "No"),
        ]

    total_rows = 0
    for pd in pages_data:
        faqs = generate_ai_faqs(pd, brand, seed_keywords=seed_keywords)
        for item in faqs:
            row = [pd["url"], item["keyword"], item["query"], item["answer"], "Auto-generated from page content"]
            row += _visibility_cells(item["query"])
            ws.append(row)
            total_rows += 1

    # Direct check of the team's own given keywords, as-typed - not just the
    # AI-rephrased question derived from them, so they can see their EXACT
    # existing keyword's current AI-citation status.
    if seed_keywords:
        home_url = pages_data[0]["url"] if pages_data else f"https://{domain}/"
        for kw in seed_keywords:
            row = [home_url, kw, kw, "", "User-provided keyword (checked as-is)"]
            row += _visibility_cells(kw)
            ws.append(row)
            total_rows += 1

    if driver:
        op._close_op_driver()

    ws.column_dimensions["A"].width = 45
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 45
    ws.column_dimensions["D"].width = 60
    ws.column_dimensions["E"].width = 30
    for col in "FGH":
        ws.column_dimensions[col].width = 22
    for row in ws.iter_rows(min_row=2, max_col=4):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    wb.save(out_path)
    log(f"   -> {total_rows} keyword/FAQ row(s) written to {out_path}")
    return total_rows


# --------------------------------------------------------------------------- #
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("domain")
    ap.add_argument("--targets", help="targets.json or .xlsx (target pages). "
                                       "If omitted, target pages are auto-discovered.")
    ap.add_argument("--out", default=str(OUTPUT_DIR))
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--no-visibility-check", action="store_true",
                     help="skip the live no-login AI-visibility checks (faster, keyword/FAQ generation only)")
    ap.add_argument("--keywords", default=None,
                     help="Comma or newline separated SEO keywords the team already has for this domain - "
                          "used to steer FAQ generation and checked directly for AI-citation as their own rows.")
    args = ap.parse_args()

    seed_keywords = None
    if args.keywords:
        seed_keywords = [k.strip() for k in re.split(r"[,\n]", args.keywords) if k.strip()]

    domain = op.safe_domain(args.domain)
    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)

    if args.targets:
        targets = op.load_targets(args.targets, domain)
    else:
        log("[*] No targets file - auto-discovering target pages...")
        targets = op.discover_targets(domain, dry_run=args.dry_run)
    if not targets:
        log("[ERROR] No target pages found.")
        sys.exit(2)
    log(f"[1/3] Loaded {len(targets)} target page(s)")

    pages_data = []
    for i, t in enumerate(targets, 1):
        log(f"   -> [crawl {i}/{len(targets)}] {t['page']}")
        pages_data.append(op.crawl_page(t["page"], t.get("keywords", []), dry_run=args.dry_run))
    op._close_op_driver()
    log(f"[2/3] Crawled {len(pages_data)} page(s)")

    homepage = pages_data[0] if pages_data else None
    brand = op.brand_from(domain, homepage.get("title") if homepage else None,
                           homepage.get("h1") if homepage else None,
                           homepage.get("og_site_name") if homepage else None)

    out_path = out_dir / f"{domain} - Keywords and Content (FAQs) Work for GEO.xlsx"
    log("[3/3] Generating AI-searchable keywords/FAQs" +
        ("" if args.no_visibility_check else " + checking live AI visibility") + "...")
    build_geo_keywords_xlsx(domain, pages_data, brand, out_path,
                             check_visibility=not args.no_visibility_check,
                             seed_keywords=seed_keywords)
    log(f"[DONE] {out_path}")


if __name__ == "__main__":
    main()
