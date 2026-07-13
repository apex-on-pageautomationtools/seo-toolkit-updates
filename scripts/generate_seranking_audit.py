"""
generate_seranking_audit.py - Turns a SEranking Site Audit Excel export into the
polished, suggestion-filled final workbook the team currently builds by hand.

Input: a SEranking Site Audit .xlsx export (one sheet per issue type - JS/CSS
issues, broken links, missing/duplicate H1, title/description issues, etc).
Output: the SAME sheets (whatever is actually present - not a fixed 29), with:
  - Merged "Suggestion" cells forward-filled so every row carries its value.
  - Judgment columns (Suggested Title / Suggested H1 / Suggested Description /
    Broken link Suggestion) - REGENERATED from the page's own live content via
    the same title/description suggestion logic already used by the SEO On-Page
    report, so it's the same "don't touch what's already fine" behavior, not a
    blind rewrite. A page that's already optimized (or ranks well) gets left
    alone with a "No changes needed" note, exactly like the On-Page report.
  - Everything else (raw crawl-fact sheets: JS/CSS/image/nofollow-link issues)
    passed through UNCHANGED - that data and its generic advice text is already
    correct; this tool never invents new claims about it.

This never edits a live site - it only produces a report for human review.

Run:
    python generate_seranking_audit.py --in "Website Audit Report.xlsx" --out "Final Audit.xlsx" --brand "Acme"
"""
import os
import re
import sys
import argparse
import urllib.request
import urllib.error
from pathlib import Path

ROOT = Path(__file__).parent.resolve()
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

# Reuse the exact same "is this actually optimized" + AI/heuristic suggestion
# logic the On-Page report uses, so behavior is consistent across both tools.
import generate_seo_onpage_phase2 as onpage2


def log(msg):
    print(msg, flush=True)


# --------------------------------------------------------------------------- #
# Column-name based classification - works for "however many sheets" the input
# actually has, not a fixed list, since exports vary run to run.
# --------------------------------------------------------------------------- #
TITLE_COLS = {"suggested title"}
DESC_COLS = {"suggested description"}
H1_COLS = {"suggested h1"}
PAGE_COLS = ("page with issues", "page url", "target pages", "url")
EXISTING_TITLE_COLS = {"existing title", "existing title and h1 tag"}
EXISTING_DESC_COLS = {"existing description"}
EXISTING_H1_COLS = {"existing h1"}


def _norm(s):
    return re.sub(r"\s+", " ", str(s or "").strip().lower())


def _find_col(headers, wanted):
    for i, h in enumerate(headers):
        if _norm(h) in wanted:
            return i
    return None


def _find_page_col(headers):
    for i, h in enumerate(headers):
        n = _norm(h)
        if any(n == p or n.startswith(p) for p in PAGE_COLS):
            return i
    return None


# --------------------------------------------------------------------------- #
# Merged-cell handling - a merged "Suggestion" cell only carries its value on
# the top-left anchor; every other cell in the range reads as None otherwise.
# --------------------------------------------------------------------------- #
def _forward_fill_merged(ws):
    for merged_range in list(ws.merged_cells.ranges):
        min_col, min_row = merged_range.min_col, merged_range.min_row
        max_col, max_row = merged_range.max_col, merged_range.max_row
        top_left = ws.cell(min_row, min_col).value
        ws.unmerge_cells(str(merged_range))
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                ws.cell(r, c).value = top_left


def _read_sheet(ws):
    _forward_fill_merged(ws)
    headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if all(v in (None, "") for v in row):
            continue
        rows.append(list(row))
    return headers, rows


# --------------------------------------------------------------------------- #
# Live page fetch - best-effort, never raises. Same shape suggest_meta() wants.
# --------------------------------------------------------------------------- #
_UA = "Mozilla/5.0 SERankingAuditBot"


def _fetch_page_data(url):
    try:
        req = urllib.request.Request(url, headers={"User-Agent": _UA})
        with urllib.request.urlopen(req, timeout=15) as r:
            html = r.read().decode("utf-8", "ignore")
    except Exception as e:
        log(f"   [warn] could not fetch {url}: {type(e).__name__}: {e}")
        html = ""
    return onpage2._parse_html(html, url, 200 if html else 0) if html else {
        "url": url, "title": onpage2.MISSING, "description": onpage2.MISSING,
        "h1": onpage2.MISSING, "body_text": "",
    }


def _guess_keywords(page_url):
    """No keyword sheet is guaranteed to exist for this input - fall back to the
    page's own slug words as a rough target keyword, same idea as auto-discovery
    in the On-Page report when no keyword list is supplied."""
    path = re.sub(r"^https?://[^/]+/?", "", page_url).strip("/")
    words = re.split(r"[-_/]+", path)
    words = [w for w in words if w and not w.isdigit()]
    return [" ".join(words[-4:])] if words else []


_brand_cache = {}


def _suggest_for_page(page_url, brand):
    if page_url in _brand_cache:
        return _brand_cache[page_url]
    pd = _fetch_page_data(page_url)
    keywords = _guess_keywords(page_url)
    result = onpage2.suggest_meta(pd, keywords, brand)
    _brand_cache[page_url] = result
    return result


# --------------------------------------------------------------------------- #
# Broken-link judgment - a dead link either has a discoverable redirect target
# (suggest that) or genuinely doesn't (suggest removal). Never invents a URL.
# --------------------------------------------------------------------------- #
def _suggest_broken_link_fix(url):
    try:
        req = urllib.request.Request(url, headers={"User-Agent": _UA}, method="HEAD")
        with urllib.request.urlopen(req, timeout=12) as r:
            final = r.geturl()
            if final and final.rstrip("/") != url.rstrip("/"):
                return f"Redirects to {final} - update the link to point there directly."
            return "No changes needed - link is live."
    except urllib.error.HTTPError as e:
        if e.code in (404, 410):
            return "Remove This URL"
        return f"HTTP {e.code} - please verify manually."
    except Exception:
        return "Could not verify automatically - please check manually."


# --------------------------------------------------------------------------- #
# Per-sheet suggestion filling
# --------------------------------------------------------------------------- #
# Sheet-NAME based fallback for inputs that don't already carry pre-built
# "Suggested X" columns (a plain issue-list export, not SEranking's own template
# with those columns already present but blank) - if the sheet is clearly about
# one of these issue types and has a page/URL column, the right suggestion
# column is added rather than requiring it to already exist.
_NAME_JUDGMENT_RULES = [
    (re.compile(r"h1", re.I), "h1"),
    (re.compile(r"title", re.I), "title"),
    (re.compile(r"description", re.I), "desc"),
    (re.compile(r"4xx|broken|dead.?link", re.I), "broken"),
]


def _apply_suggestions(sheet_name, headers, rows, brand):
    page_col = _find_page_col(headers)
    title_col = _find_col(headers, TITLE_COLS)
    desc_col = _find_col(headers, DESC_COLS)
    h1_col = _find_col(headers, H1_COLS)
    broken_col = _find_col(headers, {"broken link suggestion"})

    has_any_suggestion_col = any(c is not None for c in (title_col, desc_col, h1_col, broken_col))
    if page_col is not None and not has_any_suggestion_col:
        # No pre-built suggestion column - infer what's needed from the sheet name
        # and add it, so a plain (non-SEranking-template) issue list still gets
        # real suggestions instead of silently passing through untouched.
        kinds = {kind for pattern, kind in _NAME_JUDGMENT_RULES if pattern.search(sheet_name)}
        if kinds:
            if "title" in kinds:
                headers = headers + ["Suggested Title"]; title_col = len(headers) - 1
            if "desc" in kinds:
                headers = headers + ["Suggested Description"]; desc_col = len(headers) - 1
            if "h1" in kinds:
                headers = headers + ["Suggested H1"]; h1_col = len(headers) - 1
            if "broken" in kinds:
                headers = headers + ["Broken link Suggestion"]; broken_col = len(headers) - 1
            rows = [list(r) + [None] * (len(headers) - len(r)) for r in rows]

    if page_col is None or not (title_col is not None or desc_col is not None
                                 or h1_col is not None or broken_col is not None):
        return headers, rows  # not a judgment sheet - pass through unchanged

    log(f"   Generating suggestions for '{sheet_name}' ({len(rows)} row(s))...")
    out = []
    for row in rows:
        row = list(row)
        page_url = str(row[page_col] or "").strip() if page_col < len(row) else ""
        if page_url and page_url.startswith("http"):
            if broken_col is not None:
                row[broken_col] = _suggest_broken_link_fix(page_url)
            if title_col is not None or desc_col is not None or h1_col is not None:
                s = _suggest_for_page(page_url, brand)
                if title_col is not None:
                    row[title_col] = s["suggested_title"]
                if desc_col is not None:
                    row[desc_col] = s["suggested_description"]
                if h1_col is not None:
                    row[h1_col] = s["suggested_h1"]
        out.append(row)
    return headers, out


# --------------------------------------------------------------------------- #
# Output workbook - house style: navy header, wrapped body, sane row height.
# --------------------------------------------------------------------------- #
HEADER_FILL = PatternFill("solid", fgColor="2F5496")
HEADER_FONT = Font(bold=True, color="FFFFFF")
WRAP = Alignment(horizontal="left", vertical="center", wrap_text=True)


def _write_sheet(wb_out, name, headers, rows):
    safe_name = re.sub(r"[\[\]:\\/?*]", "_", name)[:31] or "Sheet"
    base = safe_name
    n = 1
    while safe_name in wb_out.sheetnames:
        n += 1
        safe_name = f"{base[:28]}_{n}"
    ws = wb_out.create_sheet(safe_name)
    ws.append(headers)
    for c in ws[1]:
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.alignment = WRAP
    for row in rows:
        ws.append(row)
    max_len = 1
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row_max = 1
        for cell in row:
            cell.alignment = WRAP
            row_max = max(row_max, len(str(cell.value or "")))
        max_len = max(max_len, row_max)
        lines = max(1, -(-row_max // 45))
        ws.row_dimensions[row[0].row].height = max(18, lines * 15)
    for i, h in enumerate(headers, 1):
        col_letter = ws.cell(1, i).column_letter
        ws.column_dimensions[col_letter].width = min(60, max(14, len(str(h)) + 4))
    ws.freeze_panes = "A2"


def process_workbook(in_path, out_path, brand):
    log(f"Reading: {in_path}")
    wb_in = openpyxl.load_workbook(in_path, data_only=True)
    wb_out = openpyxl.Workbook()
    wb_out.remove(wb_out.active)

    for i, sheet_name in enumerate(wb_in.sheetnames):
        log(f"[{i + 1}/{len(wb_in.sheetnames)}] Processing '{sheet_name}'...")
        ws_in = wb_in[sheet_name]
        headers, rows = _read_sheet(ws_in)
        if not headers:
            continue
        headers, rows = _apply_suggestions(sheet_name, headers, rows, brand)
        _write_sheet(wb_out, sheet_name, headers, rows)

    wb_out.save(out_path)
    log(f"[DONE] {out_path}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--in", dest="in_path", required=True, help="SEranking Site Audit .xlsx export")
    ap.add_argument("--out", dest="out_path", required=True, help="Output .xlsx path")
    ap.add_argument("--brand", default="", help="Brand name for suggested titles/descriptions")
    args = ap.parse_args()
    process_workbook(args.in_path, args.out_path, args.brand)


if __name__ == "__main__":
    main()
