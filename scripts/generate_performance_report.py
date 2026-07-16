"""
generate_performance_report.py - Standalone monthly Performance Report (pptx),
separate from every other report tool in this app.

Covers what's buildable with data this app already has real API access to:
  - Google Search Console (via gsc_audit.py's existing OAuth/Search Analytics
    plumbing - clicks/impressions/CTR/position trend, top queries, top pages,
    top countries)
  - Google Analytics 4 (via gsc_audit.py's new GA4 Admin/Data API helpers -
    users/sessions trend, traffic by channel, device breakdown)

Does NOT attempt to replicate the rank-tracker (SE Ranking) or GSC Security/
Manual-Actions/Links sections some client-facing reports also include - those
need a different data source (SE Ranking API access, or GSC report types
Google no longer exposes via API) not wired up yet. This is "james" format -
the first Performance Report format; more will be added the same way the
On-Page formats were, once real client references for them exist.

Run:
    python generate_performance_report.py example.com --gsc-account you@x.com \
        --ga4-property properties/123456789 --out "SEO Performance Report.pptx"
"""
import argparse
import datetime
import os
import sys
from pathlib import Path

ROOT = Path(__file__).parent.resolve()
REPO_ROOT = ROOT.parent
for p in (str(ROOT), str(REPO_ROOT)):
    if p not in sys.path:
        sys.path.insert(0, p)

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION

import gsc_audit

NAVY = RGBColor(0x1F, 0x38, 0x64)
BLUE = RGBColor(0x2F, 0x54, 0x96)
LIGHT_BLUE = RGBColor(0xDE, 0xEA, 0xF6)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GREY = RGBColor(0x59, 0x59, 0x59)

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)


def log(msg):
    print(msg, flush=True)


# --------------------------------------------------------------------------- #
# Data fetch - GSC (reuses gsc_audit.py's existing, already-live functions)
# --------------------------------------------------------------------------- #
def fetch_gsc_data(token, property_url, start_date, end_date):
    return {
        "daily": gsc_audit.fetch_performance_daily(token, property_url, start_date, end_date),
        "queries": gsc_audit.fetch_top_queries(token, property_url, start_date, end_date, limit=10),
        "pages": gsc_audit.fetch_top_pages(token, property_url, start_date, end_date, limit=10),
        "countries": gsc_audit.fetch_top_countries(token, property_url, start_date, end_date, limit=10),
    }


# --------------------------------------------------------------------------- #
# Data fetch - GA4
# --------------------------------------------------------------------------- #
def fetch_ga4_data(token, property_name, start_date, end_date):
    return {
        "daily": gsc_audit.run_ga4_report(
            token, property_name, start_date, end_date,
            dimensions=["date"], metrics=["activeUsers", "sessions", "newUsers"], limit=200),
        "channels": gsc_audit.run_ga4_report(
            token, property_name, start_date, end_date,
            dimensions=["sessionDefaultChannelGroup"], metrics=["sessions"], limit=10),
        "devices": gsc_audit.run_ga4_report(
            token, property_name, start_date, end_date,
            dimensions=["deviceCategory"], metrics=["activeUsers"], limit=10),
    }


# --------------------------------------------------------------------------- #
# Slide-building helpers
# --------------------------------------------------------------------------- #
def _bg(slide, color):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = color


def _title(slide, text, color=WHITE, size=32, top=Inches(0.4)):
    box = slide.shapes.add_textbox(Inches(0.6), top, SLIDE_W - Inches(1.2), Inches(1))
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(size)
    p.font.bold = True
    p.font.color.rgb = color
    return box


def _subtitle(slide, text, top, color=GREY, size=14):
    box = slide.shapes.add_textbox(Inches(0.6), top, SLIDE_W - Inches(1.2), Inches(0.8))
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(size)
    p.font.color.rgb = color
    return box


def _stat_box(slide, left, top, width, value, label):
    box = slide.shapes.add_textbox(left, top, width, Inches(1.1))
    tf = box.text_frame
    tf.word_wrap = True
    p1 = tf.paragraphs[0]
    p1.text = value
    p1.font.size = Pt(28)
    p1.font.bold = True
    p1.font.color.rgb = NAVY
    p1.alignment = PP_ALIGN.CENTER
    p2 = tf.add_paragraph()
    p2.text = label
    p2.font.size = Pt(12)
    p2.font.color.rgb = GREY
    p2.alignment = PP_ALIGN.CENTER


def build_title_slide(prs, domain, report_date):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(slide, NAVY)
    box = slide.shapes.add_textbox(Inches(0.8), Inches(2.8), SLIDE_W - Inches(1.6), Inches(1.5))
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Performance Report"
    p.font.size = Pt(44)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p2 = tf.add_paragraph()
    p2.text = f"{domain}  |  {report_date}"
    p2.font.size = Pt(18)
    p2.font.color.rgb = LIGHT_BLUE
    return slide


def build_section_divider(prs, number, title):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(slide, NAVY)
    num_box = slide.shapes.add_textbox(Inches(0.8), Inches(2.6), Inches(2), Inches(1.5))
    p = num_box.text_frame.paragraphs[0]
    p.text = number
    p.font.size = Pt(60)
    p.font.bold = True
    p.font.color.rgb = LIGHT_BLUE
    _title(slide, title, color=WHITE, size=36, top=Inches(3.3))
    return slide


def build_overview_slide(prs, domain, start_date, end_date, has_gsc, has_ga4):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(slide, WHITE)
    _title(slide, "Report Overview", color=NAVY)
    parts = [f"This report covers {domain}'s search and site performance from {start_date} to {end_date}."]
    if has_gsc:
        parts.append("Google Search Console data shows real organic search visibility - clicks, impressions, "
                     "and which queries/pages/countries are driving traffic.")
    if has_ga4:
        parts.append("Google Analytics 4 data shows how visitors actually behave once they land on the site - "
                     "user volume, traffic sources, and device usage.")
    body = slide.shapes.add_textbox(Inches(0.6), Inches(1.6), SLIDE_W - Inches(1.2), Inches(4))
    tf = body.text_frame
    tf.word_wrap = True
    for i, text in enumerate(parts):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = text
        p.font.size = Pt(16)
        p.font.color.rgb = GREY
        p.space_after = Pt(14)
    return slide


def _line_chart_slide(prs, title, subtitle, categories, series):
    """series: list of (name, values) tuples."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(slide, WHITE)
    _title(slide, title, color=NAVY)
    _subtitle(slide, subtitle, top=Inches(1.1))
    chart_data = CategoryChartData()
    chart_data.categories = categories
    for name, values in series:
        chart_data.add_series(name, values)
    x, y, cx, cy = Inches(0.6), Inches(1.8), SLIDE_W - Inches(1.2), Inches(5.2)
    gframe = slide.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, x, y, cx, cy, chart_data)
    chart = gframe.chart
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False
    return slide


def _bar_chart_slide(prs, title, subtitle, categories, series_name, values, headline=None):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(slide, WHITE)
    _title(slide, title, color=NAVY)
    _subtitle(slide, subtitle, top=Inches(1.1))
    if headline:
        _stat_box(slide, SLIDE_W - Inches(3.2), Inches(0.4), Inches(2.6), headline[0], headline[1])
    chart_data = CategoryChartData()
    chart_data.categories = categories
    chart_data.add_series(series_name, values)
    x, y, cx, cy = Inches(0.6), Inches(1.8), SLIDE_W - Inches(1.2), Inches(5.2)
    gframe = slide.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, x, y, cx, cy, chart_data)
    gframe.chart.has_legend = False
    return slide


def _pie_chart_slide(prs, title, subtitle, categories, values):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(slide, WHITE)
    _title(slide, title, color=NAVY)
    _subtitle(slide, subtitle, top=Inches(1.1))
    chart_data = CategoryChartData()
    chart_data.categories = categories
    chart_data.add_series("Share", values)
    x, y, cx, cy = Inches(2.5), Inches(1.8), Inches(8.3), Inches(5.2)
    gframe = slide.shapes.add_chart(XL_CHART_TYPE.PIE, x, y, cx, cy, chart_data)
    chart = gframe.chart
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.RIGHT
    chart.legend.include_in_layout = False
    return slide


def build_summary_slide(prs, domain):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(slide, NAVY)
    _title(slide, "Final Summary", color=WHITE, top=Inches(2.6))
    _subtitle(slide, f"Thank you for reviewing {domain}'s performance report.",
              top=Inches(3.4), color=LIGHT_BLUE, size=16)
    return slide


# --------------------------------------------------------------------------- #
# Report assembly
# --------------------------------------------------------------------------- #
def build_report(domain, out_path, gsc_data=None, ga4_data=None, start_date=None, end_date=None):
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    today = datetime.date.today().isoformat()
    build_title_slide(prs, domain, today)
    build_overview_slide(prs, domain, start_date, end_date, bool(gsc_data), bool(ga4_data))

    if gsc_data:
        build_section_divider(prs, "01", "Google Search Console")

        daily = gsc_data["daily"]
        cats = [r["keys"][0] for r in daily]
        clicks = [r.get("clicks", 0) for r in daily]
        impressions = [r.get("impressions", 0) for r in daily]
        total_clicks = sum(clicks)
        total_impr = sum(impressions)
        if daily:
            slide = _line_chart_slide(
                prs, "Traffic Status", f"Clicks & impressions, {start_date} to {end_date}",
                cats, [("Clicks", clicks), ("Impressions", impressions)])
            _stat_box(slide, SLIDE_W - Inches(3.2), Inches(0.4), Inches(1.5), f"{total_clicks:,}", "Total Clicks")
            _stat_box(slide, SLIDE_W - Inches(1.6), Inches(0.4), Inches(1.5), f"{total_impr:,}", "Total Impressions")

        queries = gsc_data["queries"]
        if queries:
            _bar_chart_slide(
                prs, "Top Searches by Keywords", "Top 10 queries by clicks in this period",
                [r["keys"][0] for r in queries], "Clicks", [r.get("clicks", 0) for r in queries],
                headline=(str(len(queries)), "Keywords Shown"))

        pages = gsc_data["pages"]
        if pages:
            _bar_chart_slide(
                prs, "Top Searches by Pages", "Top 10 pages by clicks in this period",
                [r["keys"][0] for r in pages], "Clicks", [r.get("clicks", 0) for r in pages])

        countries = gsc_data["countries"]
        if countries:
            _bar_chart_slide(
                prs, "Top Searches by Country", "Top 10 countries by clicks in this period",
                [r["keys"][0].upper() for r in countries], "Clicks", [r.get("clicks", 0) for r in countries])

    if ga4_data:
        build_section_divider(prs, "02", "Google Analytics")

        daily = ga4_data["daily"]
        if daily:
            daily_sorted = sorted(daily, key=lambda r: r.get("date", ""))
            cats = [r.get("date", "") for r in daily_sorted]
            users = [int(r.get("activeUsers", 0) or 0) for r in daily_sorted]
            sessions = [int(r.get("sessions", 0) or 0) for r in daily_sorted]
            slide = _line_chart_slide(
                prs, "Audience Trend", f"Active users & sessions, {start_date} to {end_date}",
                cats, [("Active Users", users), ("Sessions", sessions)])
            _stat_box(slide, SLIDE_W - Inches(3.2), Inches(0.4), Inches(1.5), f"{sum(users):,}", "Total Active Users")
            _stat_box(slide, SLIDE_W - Inches(1.6), Inches(0.4), Inches(1.5), f"{sum(sessions):,}", "Total Sessions")

        channels = ga4_data["channels"]
        if channels:
            _pie_chart_slide(
                prs, "Traffic Acquisition", "Sessions by channel in this period",
                [r.get("sessionDefaultChannelGroup", "Unknown") for r in channels],
                [int(r.get("sessions", 0) or 0) for r in channels])

        devices = ga4_data["devices"]
        if devices:
            _pie_chart_slide(
                prs, "Demographic Details", "Active users by device category",
                [r.get("deviceCategory", "Unknown").title() for r in devices],
                [int(r.get("activeUsers", 0) or 0) for r in devices])

    build_summary_slide(prs, domain)
    prs.save(out_path)
    log(f"[DONE] {out_path}")


# --------------------------------------------------------------------------- #
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("domain", help="Target domain, e.g. example.com")
    ap.add_argument("--out", required=True, help="Output .pptx path")
    ap.add_argument("--gsc-account", default=None, help="Connected GSC account email")
    ap.add_argument("--ga4-property", default=None, help='GA4 property resource name, e.g. "properties/123456789"')
    ap.add_argument("--days", type=int, default=28, help="How many days back to report on")
    args = ap.parse_args()

    if not args.gsc_account and not args.ga4_property:
        log("[ERROR] Provide --gsc-account and/or --ga4-property - nothing to report on otherwise.")
        sys.exit(2)

    end_date = datetime.date.today()
    start_date = end_date - datetime.timedelta(days=args.days)
    start_s, end_s = start_date.isoformat(), end_date.isoformat()

    gsc_data = None
    ga4_data = None

    if args.gsc_account:
        log(f"[1/3] Resolving GSC access for {args.gsc_account}...")
        try:
            token = gsc_audit.get_access_token(args.gsc_account)
            property_url = gsc_audit.resolve_property(token, args.domain)
            log(f"   -> Property: {property_url}")
            log("[2/3] Fetching Search Console data...")
            gsc_data = fetch_gsc_data(token, property_url, start_s, end_s)
        except Exception as e:
            log(f"   [warn] GSC data skipped: {type(e).__name__}: {e}")

    if args.ga4_property:
        log(f"[2/3] Fetching Google Analytics data for {args.ga4_property}...")
        try:
            token = gsc_audit.get_access_token(args.gsc_account) if args.gsc_account else None
            if not token:
                raise Exception("GA4 requires --gsc-account too (same OAuth token is used for both).")
            ga4_data = fetch_ga4_data(token, args.ga4_property, start_s, end_s)
        except Exception as e:
            log(f"   [warn] GA4 data skipped: {type(e).__name__}: {e}")

    if not gsc_data and not ga4_data:
        log("[ERROR] Could not fetch any GSC or GA4 data - nothing to build a report from.")
        sys.exit(2)

    log("[3/3] Building report...")
    build_report(args.domain, args.out, gsc_data=gsc_data, ga4_data=ga4_data, start_date=start_s, end_date=end_s)


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except Exception as e:
        log(f"[ERROR] {type(e).__name__}: {e}")
        sys.exit(1)
