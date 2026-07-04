#!/usr/bin/env python
"""
Capture GSC screenshots for the PPTX (James / James-Short / Omega / Neon) reports
using the saved patchright auth_states session — the SAME engine the health
report uses. This replaces the old headless-Puppeteer capture that Google blocks.

Prints a single JSON line on stdout the Node worker reads:
  {"ok": true, "prop": "...", "account": "...",
   "screens": {
       "sitemap":     {"path": "...", "text": ""},
       "manual":      {"path": "...", "text": "<innerText>"},
       "security":    {"path": "...", "text": "<innerText>"},
       "removal":     {"path": "...", "text": "<innerText>"},
       "performance": {"path": "...", "text": ""}
   }}

All progress noise from generate_report is sent to stderr so stdout stays pure
JSON.

Usage:
    python capture_gsc.py <domain> [account_key] [access_level]
"""
import sys
import json
from urllib.parse import quote

import generate_report as gr
try:                                    # patchright isn't bundled — keep importable
    from patchright.sync_api import sync_playwright
except Exception:
    sync_playwright = None

CAPTURE_KEYS = ["sitemap", "manual", "security", "removal", "performance"]
TEXT_KEYS = {"manual", "security", "removal"}
PERF_CLIP = {"x": 0, "y": 0, "width": gr.VIEWPORT_W, "height": 900}


def _grab_text(page):
    try:
        return page.evaluate("() => document.body.innerText") or ""
    except Exception:
        return ""


def _run(domain, account_hint, format_hint):
    """Do the capture and return a result dict (no stdout printing here)."""
    csv_map = gr.load_accounts_csv()
    entry = csv_map.get(domain)
    if entry:
        account_hint = account_hint or entry["account"]
        format_hint = format_hint or entry["access_level"]

    auth_files = gr.get_auth_files()
    if not auth_files:
        return {"ok": False, "error": "no auth_states session — run: python 01_auth_setup.py <accountKey>"}
    if account_hint:
        import re
        hl = re.sub(r'@gmail\.com$', '', account_hint.lower())
        matched = [f for f in auth_files if f.stem.lower() == hl or f.stem.lower().startswith(hl)]
        if matched:
            auth_files = matched

    gr.SCREENSHOTS_DIR.mkdir(exist_ok=True)
    result = {"ok": False, "screens": {}}

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = None
        prop = None
        for af in auth_files:
            ctx = browser.new_context(
                storage_state=str(af),
                viewport={"width": gr.VIEWPORT_W, "height": gr.VIEWPORT_H},
            )
            pg = ctx.new_page()
            try:
                wp = gr.find_working_property(pg, domain, format_hint=format_hint)
                if wp:
                    page = pg
                    prop = wp
                    result["account"] = af.stem
                    break
                ctx.close()
            except gr.SessionExpired:
                ctx.close()
                continue
            except Exception:
                ctx.close()
                continue

        if not prop:
            browser.close()
            return {"ok": False, "error": f"'{domain}' did not match any logged-in account (session expired or no access)"}

        result["prop"] = prop
        urls = gr.build_gsc_urls(prop)
        urls["performance"] = (
            f"https://search.google.com/search-console/performance/search-analytics?resource_id={quote(prop, safe='')}"
        )

        for key in CAPTURE_KEYS:
            url = urls.get(key)
            # ABSOLUTE path — the Node worker reads these files from a different
            # working directory (project root), so a relative path would not be
            # found and the screenshot would silently drop from the PPTX.
            path = (gr.SCREENSHOTS_DIR / f"{domain}_pptx_{key}.png").resolve()
            try:
                if key in gr.CLIP_REGIONS:
                    gr.capture_gsc_page(page, key, url, path)  # tuned clip + sitemap auto-submit
                else:
                    page.goto(url, wait_until="networkidle", timeout=60000)
                    page.wait_for_timeout(gr.EXTRA_WAIT_MS)
                    page.screenshot(path=str(path), clip=PERF_CLIP, full_page=False)
                text = _grab_text(page) if key in TEXT_KEYS else ""
                result["screens"][key] = {"path": str(path), "text": text}
            except Exception as e:
                result["screens"][key] = {"path": None, "text": "", "error": type(e).__name__}

        browser.close()

    result["ok"] = True
    return result


def main():
    if len(sys.argv) < 2:
        print(json.dumps({"ok": False, "error": "usage: capture_gsc.py <domain> [account] [access_level]"}))
        return

    domain = gr.clean_domain(sys.argv[1])
    account_hint = sys.argv[2] if len(sys.argv) > 2 else None
    format_hint = sys.argv[3] if len(sys.argv) > 3 else None

    # Send all of generate_report's progress prints to stderr; keep stdout JSON-only.
    real_stdout = sys.stdout
    sys.stdout = sys.stderr
    try:
        result = _run(domain, account_hint, format_hint)
    except Exception as e:
        result = {"ok": False, "error": f"{type(e).__name__}: {e}"}
    finally:
        sys.stdout = real_stdout

    print(json.dumps(result))


if __name__ == "__main__":
    main()
