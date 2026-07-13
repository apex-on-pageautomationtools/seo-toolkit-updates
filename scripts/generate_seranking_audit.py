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
import zlib
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
# PDF input (e.g. a SEMrush Site Audit export) - stdlib-only text extraction
# (zlib, no new dependency to distribute to every machine). PDFs like this are
# typically a summary/overview rather than a per-URL table, so this pulls out
# readable text into its own sheet rather than attempting risky, likely-wrong
# structured-table extraction from what's usually prose + stat blocks.
# --------------------------------------------------------------------------- #
def _pdf_decompress_streams(raw):
    """Every FlateDecode-compressed stream in the PDF, decompressed. Best-effort -
    a PDF with other filters (rare for text-based exports) just yields fewer/no
    chunks rather than raising."""
    chunks = []
    for m in re.finditer(rb"<<(.*?)>>\s*stream\r?\n(.*?)endstream", raw, re.S):
        dict_txt, body = m.group(1), m.group(2)
        if b"FlateDecode" not in dict_txt:
            continue
        try:
            chunks.append(zlib.decompress(body))
        except Exception:
            continue
    return chunks


def _pdf_extract_text_from_content(content):
    """Pull the text-showing operators (Tj / TJ) out of one decompressed PDF
    content stream, resolving PDF's backslash string escapes."""
    out = []
    for m in re.finditer(rb"\((?:\\.|[^\\()])*\)", content):
        s = m.group(0)[1:-1]
        s = re.sub(rb"\\([()\\])", rb"\1", s)
        s = s.replace(b"\\n", b" ").replace(b"\\r", b" ")
        try:
            out.append(s.decode("latin-1"))
        except Exception:
            continue
    return out


_PDF_CONTROL_CHARS = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def _looks_like_real_text(s):
    """FlateDecode also compresses embedded font programs and other binary
    resources, not just page content - a '(...)' pulled out of one of those
    decodes to noise, not text. Require mostly-printable content before keeping
    a line, so that garbage never reaches the output sheet (openpyxl also
    outright rejects certain control characters)."""
    if not s or len(s) < 2:
        return False
    printable = sum(1 for c in s if c == " " or (32 <= ord(c) < 127) or ord(c) > 160)
    return (printable / len(s)) > 0.85


def _ensure_pdfminer():
    """pdfminer.six properly decodes embedded custom font glyph encodings that
    the lightweight stdlib extractor below can't - confirmed against a real
    SEranking PDF export where the stdlib path recovered ~0 real words and
    pdfminer recovered the full report cleanly. Installed on first use into
    this same embedded interpreter (pure top-level import check, no restart
    needed); if pip/PyPI isn't reachable (offline machine, blocked network),
    this just returns None and the caller falls back to the stdlib extractor."""
    try:
        from pdfminer.high_level import extract_text
        return extract_text
    except ImportError:
        pass
    try:
        import subprocess
        subprocess.run([sys.executable, "-m", "pip", "install", "pdfminer.six", "--quiet"],
                        timeout=120, capture_output=True, check=True)
        from pdfminer.high_level import extract_text
        return extract_text
    except Exception:
        return None


def extract_pdf_text(path):
    """Return the PDF's visible text as a list of lines. Tries pdfminer.six
    first (handles real-world custom-font-encoded PDFs correctly); falls back
    to a lightweight stdlib-only extraction (works for PDFs with literal text
    strings, e.g. simpler/older exports) if pdfminer is unavailable or fails."""
    extract_text = _ensure_pdfminer()
    if extract_text is not None:
        try:
            import logging
            logging.getLogger("pdfminer").setLevel(logging.ERROR)
            text = extract_text(path)
            lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
            if lines:
                return lines
        except Exception:
            pass

    with open(path, "rb") as f:
        raw = f.read()
    lines = []
    for stream in _pdf_decompress_streams(raw):
        pieces = _pdf_extract_text_from_content(stream)
        if pieces:
            line = _PDF_CONTROL_CHARS.sub("", "".join(pieces)).strip()
            if line and _looks_like_real_text(line):
                lines.append(line)
    return lines


def build_pdf_summary_rows(pdf_path):
    """[[line_number, text], ...] ready to drop straight into a sheet via
    _write_sheet - the PDF's content preserved for human review, since a
    summary-style PDF rarely has enough per-URL structure to safely auto-suggest
    from without risking wrong/invented claims.

    extract_pdf_text() tries pdfminer.six first, which correctly decodes the
    embedded custom font glyph encodings some PDFs (confirmed with real
    SEranking/SEMrush exports) use instead of literal text strings - the
    lightweight stdlib fallback can only recover font metadata noise from
    those, never the actual words. The quality check below is a last-resort
    safety net for when even pdfminer can't get real text (e.g. a genuinely
    scanned/image-only PDF, or pdfminer/PyPI unavailable on this machine) -
    showing that noise as if it were real content would be actively
    misleading, so it refuses to return it and reports the limitation
    honestly instead."""
    lines = extract_pdf_text(pdf_path)
    words = " ".join(lines).split()
    has_real_words = sum(1 for w in words if w.isalpha() and len(w) > 2)
    if not lines or has_real_words < 5:
        return [[1, "This PDF's text could not be reliably extracted - it likely renders "
                    "text through an embedded custom font rather than literal text (common "
                    "in SEMrush/design-tool exports), and pdfminer.six could not be installed "
                    "or failed to decode it either. Please review the PDF manually, or "
                    "provide the underlying data as an .xlsx export instead."]]
    return [[i + 1, line] for i, line in enumerate(lines)]


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


# --------------------------------------------------------------------------- #
# Zip-of-.xls input (SEranking's "Pages" export downloaded per-issue, zipped
# together by the team instead of the single combined .xlsx). Each file is a
# legacy binary .xls (Composite Document / BIFF format, not .xlsx) named just
# "pages_<domain>_<timestamp>.xls" - no issue name in the filename or inside
# the file, so which issue each one covers has to be inferred from its column
# set instead of a sheet name (the way the combined-.xlsx path infers it).
# --------------------------------------------------------------------------- #
def _ensure_xlrd():
    """xlrd is the only maintained reader for legacy binary .xls (openpyxl only
    reads .xlsx). Installed on first use into this same embedded interpreter,
    same pattern as _ensure_pdfminer()."""
    try:
        import xlrd
        return xlrd
    except ImportError:
        pass
    try:
        import subprocess
        subprocess.run([sys.executable, "-m", "pip", "install", "xlrd==2.0.1", "--quiet"],
                        timeout=120, capture_output=True, check=True)
        import xlrd
        return xlrd
    except Exception:
        return None


def _read_seranking_zip(zip_path):
    """Returns [(filename, headers, rows), ...] for every .xls/.xlsx member.
    ignore_workbook_corruption=True is required - confirmed against a real
    SEranking zip export where 3 of 5 files raised a CompDocError without it
    despite being perfectly readable files (a known false-positive in xlrd's
    strict OLE2 stream validation, not actual corruption)."""
    import zipfile
    xlrd = _ensure_xlrd()
    out = []
    with zipfile.ZipFile(zip_path) as zf:
        names = [n for n in zf.namelist() if n.lower().endswith((".xls", ".xlsx")) and not n.startswith("__MACOSX")]
        for name in names:
            data = zf.read(name)
            try:
                if name.lower().endswith(".xlsx"):
                    import io
                    wb_in = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
                    ws_in = wb_in[wb_in.sheetnames[0]]
                    headers, rows = _read_sheet(ws_in)
                elif xlrd is not None:
                    wb_in = xlrd.open_workbook(file_contents=data, ignore_workbook_corruption=True)
                    sh = wb_in.sheet_by_index(0)
                    headers = [str(v).strip() if v is not None else "" for v in sh.row_values(0)]
                    rows = [list(sh.row_values(r)) for r in range(1, sh.nrows)
                            if any(v not in (None, "") for v in sh.row_values(r))]
                else:
                    log(f"   [warn] xlrd unavailable - skipping '{name}'")
                    continue
            except Exception as e:
                log(f"   [warn] could not read '{name}': {type(e).__name__}: {e}")
                continue
            if headers:
                out.append((name, headers, rows))
    return out


# Column-signature -> display name. Matched against a normalized set of header
# names; order matters (more specific signatures checked first). The chosen
# name is fed into the SAME _apply_suggestions()/_NAME_JUDGMENT_RULES path the
# combined-.xlsx sheets use, so e.g. "Title Issues" still gets a real
# "Suggested Title" column generated - no separate suggestion logic needed.
def _classify_xls_report(headers):
    cols = {_norm(h) for h in headers}
    has = lambda *names: all(n in cols for n in names)
    if has("title", "duplicate title"):
        return "Title Issues"
    if has("description", "duplicate description"):
        return "Description Issues"
    if has("h1", "duplicate h1"):
        return "H1 Issues"
    if "target url" in cols and "target url http status code" in cols:
        return "Redirect Chains"
    if any(n in cols for n in ("blocked by robots.txt", "robots meta tag", "x-robots-tag")):
        return "Robots Blocking Issues"
    if "page http status code" in cols:
        return "4XX HTTP Status Codes"
    # Unrecognized column set - name it from whatever non-generic column it has,
    # so it's still identifiable in the output rather than a bare "Sheet".
    distinctive = [h for h in headers if _norm(h) not in
                   ("page url", "referring pages", "is in sitemap", "url length")]
    return (distinctive[0] if distinctive else "Pages") + " Issues"


def process_workbook(in_path, out_path, brand, pdf_path=None, zip_path=None):
    wb_out = openpyxl.Workbook()
    wb_out.remove(wb_out.active)

    if in_path:
        log(f"Reading: {in_path}")
        wb_in = openpyxl.load_workbook(in_path, data_only=True)
        for i, sheet_name in enumerate(wb_in.sheetnames):
            log(f"[{i + 1}/{len(wb_in.sheetnames)}] Processing '{sheet_name}'...")
            ws_in = wb_in[sheet_name]
            headers, rows = _read_sheet(ws_in)
            if not headers:
                continue
            headers, rows = _apply_suggestions(sheet_name, headers, rows, brand)
            _write_sheet(wb_out, sheet_name, headers, rows)

    if zip_path:
        log(f"Reading zip: {zip_path}")
        members = _read_seranking_zip(zip_path)
        for i, (fname, headers, rows) in enumerate(members):
            display_name = _classify_xls_report(headers)
            log(f"[{i + 1}/{len(members)}] '{fname}' -> '{display_name}' ({len(rows)} row(s))...")
            if not headers:
                continue
            headers, rows = _apply_suggestions(display_name, headers, rows, brand)
            _write_sheet(wb_out, display_name, headers, rows)

    if pdf_path:
        log(f"Reading PDF: {pdf_path}")
        rows = build_pdf_summary_rows(pdf_path)
        log(f"   Extracted {len(rows)} line(s) of text.")
        _write_sheet(wb_out, "PDF Summary", ["#", "Text"], rows)

    wb_out.save(out_path)
    log(f"[DONE] {out_path}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--in", dest="in_path", default=None, help="SEranking Site Audit .xlsx export")
    ap.add_argument("--pdf", dest="pdf_path", default=None, help="A PDF audit export (e.g. SEMrush overview)")
    ap.add_argument("--zip", dest="zip_path", default=None,
                     help="A zip of per-issue SEranking 'Pages' .xls exports")
    ap.add_argument("--out", dest="out_path", required=True, help="Output .xlsx path")
    ap.add_argument("--brand", default="", help="Brand name for suggested titles/descriptions")
    args = ap.parse_args()
    if not args.in_path and not args.pdf_path and not args.zip_path:
        ap.error("Provide --in (.xlsx) and/or --pdf and/or --zip")
    process_workbook(args.in_path, args.out_path, args.brand, args.pdf_path, args.zip_path)


if __name__ == "__main__":
    main()
