"""
Blog Optimization Suggestions report generator.

Given a single blog post URL (+ optionally a team-supplied list of target
pages/keywords), scrapes the REAL live page, generates AI suggestions
grounded in that real content, discovers REAL internal-link targets from the
site's own sitemap, and writes a .docx matching the "james" reference format
(D:\\Report Formats\\Blog Optimization Suggestions\\...) exactly.

Nothing here is fabricated: every "Existing" field is scraped, every
"Suggested" field is an AI suggestion grounded in the real scraped content
(never invented from nothing), internal-link anchor text is always a REAL
sentence pulled verbatim from the page, and the image suggestion is either a
real image already on the site or an explicit red-flagged "needs a new
image" note - never a guessed/placeholder URL.

Run as a subprocess (same convention as generate_seo_onpage_phase2.py) so
its stdout can be streamed live to the UI.
"""
import argparse
import html as html_lib
import io
import json
import os
import re
import sys
import time
from datetime import datetime
from urllib.parse import urljoin, urlparse

import warnings

import requests
from bs4 import BeautifulSoup

try:  # sitemaps and feeds are read with the HTML parser on purpose
    from bs4 import XMLParsedAsHTMLWarning
    warnings.filterwarnings("ignore", category=XMLParsedAsHTMLWarning)
except Exception:
    pass

import docx
from docx.shared import Inches, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.enum.text import WD_COLOR_INDEX

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
# Reuse the existing OpenAI/Gemini/Groq/OpenRouter fallback chain instead of
# duplicating the HTTP calls / rate-limit bookkeeping - see that module for
# the actual provider implementations.
import generate_seo_onpage_phase2 as _onpage

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")
MISSING = "Missing!"
NOT_PUBLISHED = "Not published yet"


def log(msg):
    print(msg, flush=True)


# --------------------------------------------------------------------------- #
# Fetching
# --------------------------------------------------------------------------- #
# Sites that refused a plain request and had to be read through the browser.
_BROWSER_HOSTS = set()


def fetch(url, timeout=20, retries=2):
    """Retries on connection/timeout errors before giving up - confirmed real
    case: a single transient timeout to the target blog's own server used to
    kill the entire run (exit code 1, no report at all) even though the site
    was reachable a moment later."""
    if urlparse(url).netloc in _BROWSER_HOSTS:
        return _browser_fetch(url)
    last_err = None
    for attempt in range(retries):
        try:
            return requests.get(url, headers={"User-Agent": UA}, timeout=timeout)
        except Exception as e:
            last_err = e
            if attempt < retries - 1:
                time.sleep(3)
    log(f"   [warn] fetch failed for {url}: {type(last_err).__name__}: {last_err}")
    return None


def site_root(url):
    p = urlparse(url)
    return f"{p.scheme}://{p.netloc}"


# --------------------------------------------------------------------------- #
# Page facts read straight from the live HTML (never guessed)
# --------------------------------------------------------------------------- #
def detect_platform(url, html):
    low = html.lower()
    if "cdn.shopify.com" in low or "shopify.theme" in low or "/cdn/shop/" in low:
        return "shopify"
    if "/wp-content/" in low or "/wp-includes/" in low or "wp-json" in low:
        return "wordpress"
    if "index.php?route=" in low or "catalog/view/" in low or "journal3" in low:
        return "opencart"
    return "unknown"


# What each platform's blog editor can actually take (per the team: Shopify
# blogs have no categories; OpenCart has no category or featured-image alt
# field). A report never asks for something its platform can't do.
PLATFORMS = {
    "wordpress": {"label": "WordPress", "category": True, "cta": True, "contact_form": True,
                  "new_image": True, "featured_alt": True, "content_image_alts": True, "links": 2},
    "shopify": {"label": "Shopify", "category": False, "cta": False, "contact_form": False,
                "new_image": False, "featured_alt": True, "content_image_alts": True, "links": 5},
    "opencart": {"label": "OpenCart", "category": False, "cta": False, "contact_form": False,
                 "new_image": False, "featured_alt": False, "content_image_alts": True, "links": 5},
}


def has_read_more(soup):
    """True when the post already shows previous/next or related-post links."""
    if soup.select('a[rel~="prev"], a[rel~="next"], .post-navigation, .nav-links, '
                   '.nav-previous, .nav-next, [class*="related-post"], [class*="related_post"], '
                   '[class*="relatedpost"], [class*="yarpp"], [class*="jp-relatedposts"]'):
        return True
    for h in soup.find_all(["h2", "h3", "h4"]):
        if re.search(r"\b(related|read more|you may also like|recent posts|more articles)\b",
                     h.get_text(" ", strip=True), re.I):
            return True
    return False


def has_contact_form(soup):
    """True when the post page itself carries a contact/enquiry form. Comment
    and search forms don't count, and neither does an email-only newsletter box."""
    for form in soup.find_all("form"):
        ident = " ".join([form.get("id") or "", " ".join(form.get("class") or []),
                          form.get("action") or "", form.get("role") or ""]).lower()
        if re.search(r"comment|search|login|cart|newsletter|subscribe", ident):
            continue
        if form.find_parent(id=re.compile(r"comments|respond", re.I)):
            continue
        fields = form.find_all(["input", "textarea", "select"])
        visible = [f for f in fields if (f.get("type") or "text").lower()
                   not in ("hidden", "submit", "button", "search")]
        has_email = any((f.get("type") or "").lower() == "email"
                        or "email" in (f.get("name") or "").lower() for f in visible)
        if form.find("textarea") or (has_email and len(visible) >= 3):
            return True
    return False


def extract_faqs(soup, cap=10):
    """The post's own visible FAQ questions and answers, copied verbatim. An
    existing FAQPage block is trusted first, then a visible FAQ section.
    Returns [] when the page has no FAQ, and the report says so."""
    for tag in soup.find_all("script", attrs={"type": "application/ld+json"}):
        try:
            data = json.loads(tag.string or "")
        except Exception:
            continue
        nodes = data if isinstance(data, list) else data.get("@graph", [data]) if isinstance(data, dict) else []
        for node in nodes:
            if isinstance(node, dict) and node.get("@type") == "FAQPage":
                out = []
                for q in node.get("mainEntity") or []:
                    ans = (q.get("acceptedAnswer") or {}).get("text") or ""
                    ans = BeautifulSoup(ans, "html.parser").get_text(" ", strip=True)
                    if q.get("name") and ans:
                        out.append({"q": q["name"].strip(), "a": ans})
                if out:
                    return out[:cap]

    faq_head = None
    for h in soup.find_all(["h2", "h3", "h4"]):
        if re.search(r"\bfaqs?\b|frequently asked|veelgestelde vragen|common questions",
                     h.get_text(" ", strip=True), re.I):
            faq_head = h
    if not faq_head:
        return []
    out = []
    for el in faq_head.find_all_next(["h2", "h3", "h4", "h5", "strong", "b", "summary", "dt", "p"]):
        q = el.get_text(" ", strip=True)
        if el.name == faq_head.name and el.name == "h2" and not q.endswith("?"):
            break
        if not q.endswith("?") or len(q) > 250 or any(x["q"] == q for x in out):
            continue
        if el.name == "p" and el.find(["strong", "b"]):
            continue  # the bold question inside it is handled on its own
        start = el.find_parent("p") if el.name in ("strong", "b") and el.find_parent("p") else el
        a = ""
        if start is not el:
            a = start.get_text(" ", strip=True).replace(q, "", 1).strip()
        if not a:
            nxt = start.find_next(["p", "dd", "div"])
            while nxt is not None and not nxt.get_text(strip=True):
                nxt = nxt.find_next(["p", "dd", "div"])
            a = nxt.get_text(" ", strip=True) if nxt is not None else ""
        if a and not a.endswith("?") and a != q:
            out.append({"q": q, "a": a})
        if len(out) >= cap:
            break
    return out


def _page_date(soup, meta_prop, ld_key):
    """The post's real publish/modified date (YYYY-MM-DD), or '' if the page
    doesn't state one."""
    m = soup.find("meta", attrs={"property": meta_prop})
    val = m.get("content", "") if m else ""
    if not val:
        for tag in soup.find_all("script", attrs={"type": "application/ld+json"}):
            hit = re.search(r'"%s"\s*:\s*"([^"]+)"' % ld_key, tag.string or "")
            if hit:
                val = hit.group(1)
                break
    if not val and ld_key == "datePublished":
        t = soup.find("time", attrs={"datetime": True})
        val = t["datetime"] if t else ""
    hit = re.match(r"\d{4}-\d{2}-\d{2}", val.strip())
    return hit.group(0) if hit else ""


def _post_body(node):
    """The deepest element that still holds most of the post's real
    paragraphs - so header logos, menus and footer banners don't get listed
    as post images."""
    ps = [p for p in node.find_all("p") if len(p.get_text(" ", strip=True).split()) >= 8]
    if not ps:
        return node
    counts, depth = {}, {}
    for p in ps:
        for d, anc in enumerate(reversed(list(p.parents))):
            counts[id(anc)] = counts.get(id(anc), 0) + 1
            depth[id(anc)] = (d, anc)
    need = max(1, int(len(ps) * 0.7))
    best = max((depth[k] for k, c in counts.items() if c >= need), key=lambda x: x[0], default=None)
    return best[1] if best else node


# --------------------------------------------------------------------------- #
# Scrape the target blog page for real
# --------------------------------------------------------------------------- #
def docx_to_html(path):
    """A blog content document (.docx) turned into simple HTML, so a post
    that is not published yet goes through the same reading code as a live
    page. Heading styles become headings; a short all-bold line counts as a
    heading too. "Meta Title:" / "Meta Description:" lines are picked out as
    the post's planned meta, not treated as body text."""
    d = docx.Document(path)
    parts, meta, has_h1, images = [], {}, False, {}
    for para in d.paragraphs:
        # Pictures placed in the document are the post's images - keep each
        # one, with its alt text, at the spot where it sits in the content.
        for drawing in para._p.iter(qn("w:drawing")):
            blip = next(drawing.iter(qn("a:blip")), None)
            rid = blip.get(qn("r:embed")) if blip is not None else None
            if not rid or rid not in d.part.related_parts:
                continue
            doc_pr = next(drawing.iter(qn("wp:docPr")), None)
            key = f"docimg:{len(images) + 1}"
            images[key] = d.part.related_parts[rid].blob
            alt = (doc_pr.get("descr") or "") if doc_pr is not None else ""
            parts.append(("img", (key, alt)))
        text = para.text.strip()
        if not text:
            continue
        m = re.match(r"(?i)^(meta\s*title|seo\s*title|title\s*tag|meta\s*description|meta\s*desc)"
                     r"\s*[:\-\u2013]\s*(.+)$", text)
        if m:
            meta.setdefault("meta" if "desc" in m.group(1).lower() else "title", m.group(2).strip())
            continue
        style = (para.style.name or "").lower()
        hm = re.match(r"heading (\d)", style)
        runs = [r for r in para.runs if r.text.strip()]
        if style == "title":
            tag = "h1"
        elif hm:
            tag = f"h{min(int(hm.group(1)), 6)}"
        elif len(text) <= 120 and runs and all(r.bold for r in runs):
            tag = "h2"
        else:
            tag = "p"
        if tag == "h1":
            tag = "h2" if has_h1 else "h1"
            has_h1 = True
        parts.append((tag, text))
    if not has_h1 and parts and parts[0][0] == "h2":
        parts[0] = ("h1", parts[0][1])
    body = "".join(f'<img src="{x[0]}" alt="{html_lib.escape(x[1], quote=True)}">' if t == "img"
                   else f"<{t}>{html_lib.escape(x)}</{t}>" for t, x in parts)
    return f"<html><body><article>{body}</article></body></html>", meta, images


def scrape_blog(url, content_doc=None):
    site_html, doc_meta, doc_images = "", {}, {}
    if content_doc:
        page_html, doc_meta, doc_images = docx_to_html(content_doc)
        home = fetch(site_root(url) + "/")
        site_html = home.text if home and home.status_code < 400 else ""
        if not site_html:
            site_html = fetch_in_browser(site_root(url) + "/")
        if not site_html:
            log("   [warn] could not open the site's homepage - internal links, contact page and "
                "platform detection may be missing.")
        r = None
    else:
        r = fetch(url)
    if content_doc:
        pass
    elif r and r.status_code < 400:
        page_html = r.text
    else:
        # Some sites answer a plain request with an error page (bot protection)
        # but load normally in a real browser - read the same live page there.
        log(f"   Direct request failed{f' (HTTP {r.status_code})' if r else ''} - retrying in a browser...")
        page_html = fetch_in_browser(url)
        if not page_html:
            raise Exception(f"Could not fetch {url}"
                            + (f" (HTTP {r.status_code})" if r else ""))
    soup = BeautifulSoup(page_html, "html.parser")

    title_tag = soup.find("title")
    existing_title = title_tag.get_text(strip=True) if title_tag else MISSING

    meta_desc_tag = soup.find("meta", attrs={"name": "description"})
    existing_meta = (meta_desc_tag.get("content", "").strip()
                     if meta_desc_tag and meta_desc_tag.get("content") else MISSING)

    h1_tag = soup.find("h1")
    existing_h1 = h1_tag.get_text(strip=True) if h1_tag else MISSING
    if content_doc:
        existing_title = doc_meta.get("title") or NOT_PUBLISHED
        existing_meta = doc_meta.get("meta") or NOT_PUBLISHED

    og_image = soup.find("meta", attrs={"property": "og:image"})
    featured_image = og_image.get("content", "").strip() if og_image and og_image.get("content") else ""
    if not featured_image and not content_doc:
        first_img = soup.select_one("article img, .entry-content img, main img")
        if first_img and first_img.get("src"):
            featured_image = urljoin(url, first_img["src"])

    # Existing category - WordPress convention: rel="category tag" links.
    cat_links = soup.select('a[rel*="category"]')
    existing_category = cat_links[0].get_text(strip=True) if cat_links else "Missing"

    # Real body paragraphs (for internal-linking anchor text and AI context) -
    # prefer the main content container over the whole page (nav/footer noise).
    # Pick whichever candidate container has the MOST real paragraphs, rather
    # than trusting selector priority blindly - confirmed real case: a site's
    # first <article> match was a "suggested content" teaser card with only
    # one short paragraph, while the real post body (224 paragraphs) was
    # elsewhere in <main> - selector-priority order alone picked the wrong,
    # nearly-empty container and silently produced zero usable paragraphs.
    def _real_paragraphs(node):
        ps = [p.get_text(" ", strip=True) for p in node.find_all("p")]
        return [p for p in ps if len(p.split()) >= 8]  # skip short/nav fragments

    candidates = [c for c in (soup.select_one("article"), soup.select_one(".entry-content"),
                              soup.select_one("main"), soup.body) if c]
    paragraphs, best_node = [], soup
    for cand in candidates:
        found = _real_paragraphs(cand)
        if len(found) > len(paragraphs):
            paragraphs, best_node = found, cand
    if not paragraphs:
        paragraphs = _real_paragraphs(soup)
    # Keep only the post's own paragraphs - confirmed real case: a store's
    # "Showroom, gym, testen en advies" block sat in the same container and
    # got offered as a sentence to link from.
    in_post = _real_paragraphs(_post_body(best_node))
    if len(in_post) >= 3:
        paragraphs = in_post

    # Real images inside the post body, with the alt text they have today.
    content_images, seen_src = [], set()
    featured_name = os.path.basename(urlparse(featured_image).path) if featured_image else ""
    for img in _post_body(best_node).find_all("img"):
        src = img.get("src") or img.get("data-src") or ""
        if not src or src.startswith("data:"):
            continue
        src = src if src.startswith("docimg:") else urljoin(url, src)
        name = os.path.basename(urlparse(src).path)
        if src in seen_src or (featured_name and name == featured_name):
            continue
        seen_src.add(src)
        near = img.find_previous(["h1", "h2", "h3"])
        content_images.append({"url": src, "alt": (img.get("alt") or "").strip(),
                               "near": near.get_text(" ", strip=True) if near else ""})
        if len(content_images) >= 10:
            break
    featured_alt = ""
    if featured_image:
        for img in soup.find_all("img"):
            if featured_name and featured_name in (img.get("src") or img.get("data-src") or ""):
                featured_alt = (img.get("alt") or "").strip()
                break
    body_text = "\n".join(paragraphs)

    # Best-guess Contact page - a real link on the site, not invented.
    contact_url = ""
    for a in (BeautifulSoup(site_html, "html.parser") if content_doc else soup).find_all("a", href=True):
        if re.search(r"contact", a.get_text(" ", strip=True), re.I) or re.search(r"/contact", a["href"], re.I):
            contact_url = urljoin(url, a["href"])
            break

    return {
        "url": url,
        "site_root": site_root(url),
        "platform": detect_platform(url, site_html if content_doc else page_html),
        "from_doc": bool(content_doc),
        "site_html": site_html,
        "doc_images": doc_images,
        "has_read_more": has_read_more(soup),
        "has_contact_form": has_contact_form(soup),
        "faqs": extract_faqs(soup),
        "date_published": _page_date(soup, "article:published_time", "datePublished"),
        "date_modified": _page_date(soup, "article:modified_time", "dateModified"),
        "existing_title": existing_title,
        "existing_meta": existing_meta,
        "existing_h1": existing_h1,
        "featured_image": featured_image,
        "featured_alt": featured_alt,
        "content_images": content_images,
        "existing_category": existing_category,
        "paragraphs": paragraphs,
        "body_text": body_text[:8000],  # keep the AI prompt bounded
        "contact_url": contact_url,
        "html": page_html,
    }


# --------------------------------------------------------------------------- #
# Sitemap crawl - real internal-link/CTA targets, not invented URLs
# --------------------------------------------------------------------------- #
def crawl_sitemap(root, cap=300):
    """Fetch sitemap.xml (following one level of sub-sitemaps) and return a
    flat list of real page URLs on this site. Best-effort - returns [] if the
    site has no sitemap, which the caller treats as "no internal-link target
    found" rather than guessing one."""
    seen = []
    queue = [urljoin(root, "/sitemap.xml"), urljoin(root, "/sitemap_index.xml")]
    visited_sitemaps = set()
    while queue and len(seen) < cap:
        sm_url = queue.pop(0)
        if sm_url in visited_sitemaps:
            continue
        visited_sitemaps.add(sm_url)
        r = fetch(sm_url, timeout=15)
        if not r or r.status_code >= 400:
            continue
        try:
            soup = BeautifulSoup(r.text, "xml")
        except Exception:
            soup = BeautifulSoup(r.text, "html.parser")
        locs = [loc.get_text(strip=True) for loc in soup.find_all("loc")]
        for loc in locs:
            if loc.endswith(".xml") and len(visited_sitemaps) < 6:
                queue.append(loc)
            elif re.search(r"\.(jpe?g|png|gif|webp|svg|pdf|mp4)(\?|$)", loc, re.I):
                continue  # image/file entries inside a sitemap are not pages
            elif loc not in seen:
                seen.append(loc)
            if len(seen) >= cap:
                break
    return seen


def _topic_words(keywords, sitemap_urls):
    """The distinctive words of the post's topic. Words that appear in a large
    share of the site's own URLs (a brand or product word such as "kratom" on
    a kratom store) say nothing about which page is relevant, so they are
    dropped - confirmed real case: whole-phrase matching found no target at
    all for phrases like "GMP kratom vendor requirements"."""
    words = {w for k in (keywords or []) for w in re.findall(r"[a-z0-9]{3,}", k.lower())}
    words -= {"the", "and", "for", "with", "your", "from", "that", "this", "what", "how", "are"}
    if not sitemap_urls:
        return words
    paths = [urlparse(u).path.lower() for u in sitemap_urls]
    distinctive = {w for w in words if sum(1 for pth in paths if w in pth) <= len(paths) * 0.3}
    return distinctive or words


def score_candidates(sitemap_urls, current_url, target_keywords):
    """Cheap URL scoring by how many distinctive topic words a page's path
    contains, so only the most plausible candidates get their <title> fetched
    instead of every page on the site."""
    words = _topic_words(target_keywords, sitemap_urls)
    scored = []
    for u in sitemap_urls:
        if u.rstrip("/") == current_url.rstrip("/"):
            continue
        path = urlparse(u).path.lower()
        scored.append((sum(1 for w in words if w in path), u))
    scored.sort(key=lambda x: -x[0])
    return scored


def _targets_from_page_links(scraped, keywords, max_targets):
    """Real pages this site itself links to (menu, footer, sidebar), scored by
    how many topic words appear in the link text or URL."""
    host = urlparse(scraped["url"]).netloc
    words = {w for k in keywords for w in re.findall(r"[a-z0-9]{4,}", k.lower())}
    scored, seen = [], {scraped["url"].rstrip("/")}
    for a in BeautifulSoup(scraped["site_html"] or scraped["html"], "html.parser").find_all("a", href=True):
        u = urljoin(scraped["url"], a["href"]).split("#")[0]
        text = a.get_text(" ", strip=True)
        if urlparse(u).netloc != host or u.rstrip("/") in seen or not text or len(text) > 80:
            continue
        if re.search(r"cart|checkout|account|login|wishlist|compare|javascript:", u, re.I):
            continue
        seen.add(u.rstrip("/"))
        hay = (text + " " + urlparse(u).path).lower()
        score = sum(1 for w in words if w in hay)
        if score:
            scored.append((score, u, text))
    scored.sort(key=lambda x: -x[0])
    return [{"url": u, "title": t} for _, u, t in scored[:max_targets]]


def _is_200_ok(r, url):  # noqa: note requests.Response is falsy for 4xx/5xx
    """True only for a page that answers 200 at its own address. A page that
    errors, or that redirects somewhere else, must not be linked to (a
    http->https, www or trailing-slash hop to the same page is fine)."""
    if r is None or r.status_code != 200:
        return False

    def norm(u):
        u = re.sub(r"^https?://(www\.)?", "", (u or "").lower()).split("#")[0]
        return u.rstrip("/")
    final = getattr(r, "url", "") or url
    return norm(final) == norm(url)


def find_related_posts(scraped, keywords, want=2):
    """Two real posts from the same blog that are closest to this post's
    topic, for the Previous / Next (read more) links. Read from the post page
    and the blog's own listing page - never invented."""
    url = scraped["url"]
    if "journal_blog_post_id=" in url:
        listing, marker = url.split("/post&")[0], "journal_blog_post_id="
    else:
        listing = url.rstrip("/").rsplit("/", 1)[0]
        marker = urlparse(listing).path.rstrip("/") + "/"
        if marker == "/":
            return []
    pages = [scraped["site_html"] or scraped["html"]]
    r = fetch(listing, timeout=15)
    if r and r.status_code < 400:
        pages.append(r.text)
    host, found = urlparse(url).netloc, {}
    for html in pages:
        for a in BeautifulSoup(html, "html.parser").find_all("a", href=True):
            u = urljoin(url, a["href"]).split("#")[0]
            text = a.get_text(" ", strip=True)
            if (urlparse(u).netloc != host or marker not in u or u.rstrip("/") == url.rstrip("/")
                    or u.rstrip("/") == listing.rstrip("/") or "/tagged/" in u
                    or not 15 <= len(text) <= 160):
                continue
            if len(text) > len(found.get(u, "")):
                found[u] = text
    words = {w for k in (list(keywords) or [scraped["existing_h1"]])
             for w in re.findall(r"\w{4,}", k.lower())}
    scored = sorted(((sum(1 for w in words if w in t.lower()), u, t) for u, t in found.items()),
                    key=lambda x: -x[0])
    return [{"url": u, "title": t} for sc, u, t in scored[:want] if sc > 0]


def find_link_targets(scraped, target_pages, target_keywords, topic_keywords, max_targets=2, fetch_cap=12):
    """Real internal-link / CTA-button targets:
    1. Team-supplied target pages, each checked live. Only pages answering
       200 OK are kept; an error or a redirect is skipped and logged.
    2. Plus pages found on the site itself (sitemap, or its menu/footer
       links) that match the post's topic, also 200 OK only.
    These are CANDIDATES, not orders: pick_link_sentences() links a page only
    where the post genuinely talks about that page's topic, so a team page
    with no relevant sentence is left out rather than forced in.
    Returns [{"url":..., "title":..., "from_team": bool}, ...] - never a
    fabricated URL."""
    out, seen = [], {scraped["url"].rstrip("/")}
    for u in (target_pages or [])[:15]:
        if u.rstrip("/") in seen:
            continue
        seen.add(u.rstrip("/"))
        r = fetch(u, timeout=12)
        if _is_200_ok(r, u):
            t = BeautifulSoup(r.text, "html.parser").find("title")
            out.append({"url": u, "title": t.get_text(strip=True) if t else u, "from_team": True})
        else:
            why = ("no response" if r is None else f"HTTP {r.status_code}" if r.status_code != 200
                   else f"redirects to {getattr(r, 'url', '')}")
            log(f"   [warn] target page is not 200 OK ({why}), not linking to it: {u}")
    if target_pages:
        log(f"   {len(out)} of {len(target_pages[:15])} team target page(s) are 200 OK.")

    keywords = target_keywords or topic_keywords or []
    sitemap_urls = crawl_sitemap(scraped["site_root"])
    if not sitemap_urls:
        log("   No sitemap found - using the pages linked from this site's own menu and footer.")
        found = _targets_from_page_links(scraped, keywords, max_targets)
        candidates, words = [f["url"] for f in found], None
    else:
        scored = score_candidates(sitemap_urls, scraped["url"], keywords)
        candidates = [u for score, u in scored if score > 0][:fetch_cap] or [u for _, u in scored[:fetch_cap]]
        words = _topic_words(keywords, sitemap_urls)

    auto = 0
    for u in candidates:
        if auto >= max_targets:
            break
        if u.rstrip("/") in seen:
            continue
        seen.add(u.rstrip("/"))
        r = fetch(u, timeout=12)
        if not _is_200_ok(r, u):
            continue
        t = BeautifulSoup(r.text, "html.parser").find("title")
        title = t.get_text(strip=True) if t else u
        if words is None or any(w in title.lower() or w in urlparse(u).path.lower() for w in words):
            out.append({"url": u, "title": title, "from_team": False})
            auto += 1
    return out


# --------------------------------------------------------------------------- #
# AI suggestions - OpenAI tried FIRST (per-team decision for this tool, since
# it's a single-URL report rather than a bulk run), then the same free-tier
# fallbacks generate_seo_onpage_phase2 already uses. Reuses that module's
# actual provider functions instead of duplicating the HTTP/retry logic.
# --------------------------------------------------------------------------- #
def ai_chain(prompt):
    for fn in (_onpage._ai_suggest_openai, _onpage._ai_suggest_gemini,
              _onpage._ai_suggest_groq, _onpage._ai_suggest_openrouter):
        result = fn(prompt)
        if result is not None:
            return result
    return None


def ai_suggestions(scraped):
    prompt = f"""You are an SEO editor. Read this real blog post and suggest improvements.
Return ONLY a JSON object with exactly these keys:
- "category": a single WordPress-style category name for this post
- "tags": array of 3-6 short keyword tags
- "title": either the string "No Change Required" (if the existing title is already good) or an improved SEO title
- "meta_description": an improved meta description, 140-160 characters
- "h1_change": either the string "No Change Required" or a new suggested H1 (if the existing H1 is missing, suggest one)
- "topic_keywords": array of 4-8 short keywords/phrases capturing what this post is really about (used to find relevant internal links and images - be specific, not generic)
- "image_alt": SEO-friendly alt text for the featured image: a short keyword phrase of 2-6 words for the post topic, not a sentence
- "image_title_attr": suggested title attribute for the featured image
- "cta_text": a short (3-6 word) call-to-action button label relevant to this post's topic
- "cta_placement": where on the post to place the CTA button - use a plain text description such as "End of post", "After heading '<exact heading text>'", or "Below the section about <topic>"; never use the word "screenshot"

Existing title: {scraped['existing_title']}
Existing meta description: {scraped['existing_meta']}
Existing H1: {scraped['existing_h1']}
Existing category: {scraped['existing_category']}

Post content (real, scraped from the live page):
{scraped['body_text'][:5000]}
"""
    result = ai_chain(prompt)
    if not result:
        # Every AI provider failed (bad/missing keys, rate limits, etc.) - this
        # used to hard-crash the whole run (exit code 1, no report at all).
        # Team ask: give them SOMETHING usable instead of nothing - fall back
        # to the page's own existing title/meta/H1/category rather than a
        # fabricated suggestion. Downstream callers already handle empty
        # tags/h2_list/topic_keywords gracefully (find_reusable_image and
        # pick_link_sentences both no-op on empty input and flag the gap in
        # red, same as the existing missing-image/missing-CTA fallback path).
        log("   [warn] Every AI provider failed - falling back to the page's existing "
            "title/meta/H1 instead of failing the whole report.")
        return {
            "category": "",
            "tags": [],
            "title": "No Change Required",
            "meta_description": "" if scraped.get("existing_meta") == MISSING else scraped.get("existing_meta") or "",
            "h1_change": "No Change Required",
            "topic_keywords": [],
            "image_alt": "",
            "image_title_attr": "",
            "cta_text": "",
            "cta_placement": "End of post",
        }
    return result


def _image_for_ai(scraped, src):
    """The real image as a small JPEG (base64), ready to show to a vision
    model. `src` is an image URL, a local file path, or a picture inside the
    content document. Returns "" if the image can't be loaded."""
    import base64
    try:
        from PIL import Image
        if src in scraped.get("doc_images", {}):
            raw = scraped["doc_images"][src]
        elif os.path.exists(src):
            raw = open(src, "rb").read()
        elif urlparse(src).netloc in _BROWSER_HOSTS:
            b64 = _get_driver().execute_async_script("""
                var done = arguments[arguments.length - 1];
                fetch(arguments[0]).then(function (r) { return r.blob(); }).then(function (b) {
                    var fr = new FileReader();
                    fr.onload = function () { done(fr.result.split(',')[1]); };
                    fr.readAsDataURL(b);
                }).catch(function () { done(''); });
            """, src)
            raw = base64.b64decode(b64 or "")
        else:
            raw = requests.get(src, headers={"User-Agent": UA}, timeout=25).content
        img = Image.open(io.BytesIO(raw)).convert("RGB")
        img.thumbnail((1024, 1024))
        buf = io.BytesIO()
        img.save(buf, "JPEG", quality=80)
        return base64.b64encode(buf.getvalue()).decode("ascii")
    except Exception as e:
        log(f"   [warn] could not load an image to look at it: {type(e).__name__}")
        return ""


ALT_STYLE = ("The alt text must be SEO friendly: a short keyword phrase of 2 to 6 words that people would "
             "search for, NOT a sentence describing the scene. For example \"Home Gym Equipment\" or "
             "\"Adjustable Weight Bench\", never \"Man training while seated with dumbbells on a bench\". "
             "It must still be true to what the image shows and to the post's topic. Never include a brand "
             "name, store name or product model name, even if one is visible in the image or mentioned in "
             "the post: use the generic product or topic keyword instead. Write it in the same "
             "language as the post. Reply with the alt text only.")


def ai_alt_from_image(scraped, src, section="", keywords=(), used=()):
    """Alt text written by a model that actually SEES the image, with the post
    as context. OpenAI first, then Gemini. Returns "" if neither could do it,
    and the caller falls back to the text-only guess (and says so)."""
    import urllib.request
    b64 = _image_for_ai(scraped, src)
    if not b64:
        return ""
    title = scraped["existing_h1"] if scraped["existing_h1"] != MISSING else scraped["existing_title"]
    prompt = (f'This image is used in a blog post titled "{title}"'
              + (f', in the section "{section}"' if section else "")
              + f'. The post begins: "{scraped["body_text"][:500]}"\n\n'
              "Look at the image and write its alt text. " + ALT_STYLE
              + (f"\nThe post's target keywords are: {', '.join(keywords)}. Use one of them if it fits "
                 "what the image shows, otherwise write a close keyword phrase that does." if keywords else "")
              + (f"\nAlready used for other images in this post, so pick a different phrase: "
                 f"{', '.join(used)}." if used else ""))

    def _clean(text):
        text = (text or "").strip().strip('"').strip()
        return text.splitlines()[0][:160] if text else ""

    key = os.environ.get("OPENAI_API_KEY", "").strip()
    if key:
        try:
            body = {"model": os.environ.get("OPENAI_MODEL", "gpt-5.6-luna"), "messages": [{"role": "user", "content": [
                {"type": "text", "text": prompt},
                {"type": "image_url", "image_url": {"url": "data:image/jpeg;base64," + b64, "detail": "low"}}]}]}
            req = urllib.request.Request("https://api.openai.com/v1/chat/completions",
                                         data=json.dumps(body).encode("utf-8"),
                                         headers={"Content-Type": "application/json",
                                                  "Authorization": f"Bearer {key}"})
            with urllib.request.urlopen(req, timeout=60) as r:
                alt = _clean(json.loads(r.read())["choices"][0]["message"]["content"])
            if alt:
                return alt
        except Exception as e:
            log(f"   [warn] OpenAI could not read the image: {type(e).__name__}: {str(e)[:120]}")
    key = os.environ.get("GEMINI_API_KEY", "").strip()
    if key:
        try:
            model = os.environ.get("GEMINI_MODEL", "gemini-2.5-flash")
            body = {"contents": [{"parts": [{"text": prompt},
                                            {"inline_data": {"mime_type": "image/jpeg", "data": b64}}]}]}
            req = urllib.request.Request(
                f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={key}",
                data=json.dumps(body).encode("utf-8"), headers={"Content-Type": "application/json"})
            with urllib.request.urlopen(req, timeout=60) as r:
                alt = _clean(json.loads(r.read())["candidates"][0]["content"]["parts"][0]["text"])
            if alt:
                return alt
        except Exception as e:
            log(f"   [warn] Gemini could not read the image: {type(e).__name__}")
    return ""


def suggest_image_alts(scraped, images, keywords=()):
    """Fills im["suggested"] for every image: by looking at the image where
    possible (im["seen"] = True), otherwise from its file name and context.
    Each image gets its own keyword phrase, not the same one repeated."""
    used = []
    for im in images:
        im["suggested"] = ai_alt_from_image(scraped, im["url"], im.get("near", ""), keywords, used)
        im["seen"] = bool(im["suggested"])
        if im["seen"]:
            used.append(im["suggested"])
    blind = [im for im in images if not im["seen"]]
    for im, alt in zip(blind, ai_image_alts(scraped, blind, keywords)):
        im["suggested"] = alt


def ai_image_alts(scraped, images, keywords=()):
    """One suggested alt text per existing image, in the same order. The AI
    only sees the file name, current alt and nearest heading (it cannot view
    the picture), so the report says to check each one against the image."""
    if not images:
        return []
    listed = "\n".join(f"[{i}] file: {os.path.basename(urlparse(im['url']).path)} | current alt: "
                       f"{im['alt'] or '(empty)'} | section: {im.get('near') or '(top of post)'}"
                       for i, im in enumerate(images))
    post_title = scraped["existing_h1"] if scraped["existing_h1"] != MISSING else scraped["existing_title"]
    prompt = f"""Write image alt text for a blog post titled "{post_title}".
For each image below give one alt text, based on its file name, current alt text and the section
it sits in. {ALT_STYLE.replace("Reply with the alt text only.", "")}
Target keywords of the post: {', '.join(keywords) or '(none given)'}. Give each image a different phrase.
Return ONLY a JSON array of strings, one per image, in the same order ({len(images)} items).

Images:
{listed}
"""
    result = ai_chain(prompt)
    if isinstance(result, dict):
        result = next((v for v in result.values() if isinstance(v, list)), None)
    if not isinstance(result, list) or len(result) != len(images):
        return [""] * len(images)
    return [str(x).strip() for x in result]


def pick_link_sentences(paragraphs, link_targets, max_pairs=None, keywords=(), pairs=()):
    """Ask the AI to pick which REAL paragraph(s) should carry an internal
    link to each target, by INDEX only - it never writes new anchor text,
    so "Find with" is always a verbatim sentence from the actual page."""
    if not paragraphs or not link_targets:
        return []
    numbered = "\n".join(f"[{i}] {p}" for i, p in enumerate(paragraphs[:40]))
    max_pairs = max_pairs or len(link_targets)
    wanted = {p["url"].rstrip("/"): p["keyword"] for p in pairs if p.get("keyword")}
    targets_desc = "\n".join(
        f"- {t['url']} (\"{t['title']}\")"
        + (f" [team's keyword for this page: \"{wanted[t['url'].rstrip('/')]}\"]"
           if t["url"].rstrip("/") in wanted else "")
        for t in link_targets)
    kw_line = (f"Target keywords: {', '.join(keywords)}. When one of these (or the team's keyword for a "
               "page) appears in a relevant paragraph, use it as the anchor text.\n" if keywords else "")
    prompt = f"""Given these real numbered paragraphs from a blog post, and these real target pages
on the same site, pick ONE paragraph index per target page that would be a natural
place to add an internal link to that target (based on topical relevance), and the exact
words inside that paragraph that should become the link.
Return ONLY a JSON array of objects:
[{{"index": <int>, "target_url": "<one of the target URLs above>", "anchor_text": "<2-5 words copied EXACTLY from that paragraph>"}}]
Pick at most {max_pairs} pairs, each with a DIFFERENT paragraph index and a DIFFERENT target page.
Do not force a link: only link a page where the paragraph is genuinely about that page's topic, and
leave a target page out completely if no paragraph fits it. Never invent anchor text: it must
appear character for character in the chosen paragraph.
{kw_line}
Paragraphs:
{numbered}

Target pages:
{targets_desc}
"""
    result = ai_chain(prompt)
    if not isinstance(result, list):
        return []
    out = []
    for item in result:
        try:
            idx = int(item.get("index"))
            target_url = item.get("target_url")
            if not (0 <= idx < len(paragraphs) and any(t["url"] == target_url for t in link_targets)):
                continue
            # The team's own keyword for this page wins as the anchor when the
            # paragraph really contains it.
            team_kw = wanted.get((target_url or "").rstrip("/"), "")
            anchor = team_kw if team_kw and team_kw.lower() in paragraphs[idx].lower() else item.get("anchor_text")
            pair = _sentence_with_anchor(paragraphs[idx], anchor or "")
            if pair and len(out) < max_pairs and not any(o["target_url"] == target_url for o in out):
                out.append({"sentence": pair[0], "anchor": pair[1], "target_url": target_url})
        except Exception:
            continue
    return out


def _sentence_with_anchor(paragraph, anchor):
    """(sentence, anchor) where the anchor is verified to sit inside that real
    sentence of the page. None if the AI's anchor text isn't really there."""
    anchor = anchor.strip().strip(".,;:!?")
    if not anchor:
        return None
    pos = paragraph.lower().find(anchor.lower())
    if pos < 0:
        return None
    anchor = paragraph[pos:pos + len(anchor)]
    for sentence in re.split(r"(?<=[.!?])\s+", paragraph):
        if anchor in sentence:
            return sentence.strip(), anchor
    return paragraph, anchor


# --------------------------------------------------------------------------- #
# Real image reuse (never a fabricated URL)
# --------------------------------------------------------------------------- #
def find_reusable_image(scraped, topic_keywords):
    """Look for a real image already on the site that matches the post's
    topic - checks WordPress's media library search first (if available),
    then falls back to scanning the post's own images for one with relevant
    alt text. Returns a real image URL, or None if nothing confident was
    found - the caller flags that in red rather than guessing a URL."""
    root = scraped["site_root"]
    for kw in (topic_keywords or [])[:3]:
        r = fetch(f"{root}/wp-json/wp/v2/media?search={requests.utils.quote(kw)}&per_page=3", timeout=12)
        if r and r.status_code == 200:
            try:
                items = r.json()
                if isinstance(items, list) and items:
                    src = (items[0].get("source_url") or "").strip()
                    if src:
                        return src
            except Exception:
                pass

    soup = BeautifulSoup(scraped["html"], "html.parser")
    for img in soup.find_all("img"):
        alt = (img.get("alt") or "").lower()
        src = img.get("src") or ""
        if not src or src.startswith("docimg:"):
            continue
        if any((kw or "").lower() in alt for kw in (topic_keywords or [])):
            return urljoin(scraped["url"], src)
    return None


# --------------------------------------------------------------------------- #
# Best-effort annotated screenshot (real screenshot, box drawn around the
# element in question) - approximation of a human's exact placement call,
# never a fabricated image.
# --------------------------------------------------------------------------- #
_driver = None
_driver_visible = False


def _get_driver(visible=None):
    """Hidden Chrome by default. `visible=True` switches to a normal Chrome
    window for the rest of the run - some protected sites (Cloudflare) answer a
    hidden browser with an error page but load normally in a real window."""
    global _driver, _driver_visible
    if visible is None:
        visible = _driver_visible
    if _driver is not None and visible == _driver_visible:
        try:
            _ = _driver.title
            return _driver
        except Exception:
            _driver = None
    close_driver()
    try:
        import tempfile
        from selenium.webdriver import Chrome, ChromeOptions
        opts = ChromeOptions()
        opts.add_argument(f"--user-data-dir={tempfile.mkdtemp(prefix='blogopt_')}")
        if not visible:
            opts.add_argument("--headless=new")
            opts.add_argument(f"--user-agent={UA}")
        opts.add_argument("--no-sandbox")
        opts.add_argument("--disable-dev-shm-usage")
        opts.add_argument("--disable-gpu")
        opts.add_argument("--window-size=1366,1400")
        opts.add_argument("--disable-blink-features=AutomationControlled")
        opts.add_experimental_option("excludeSwitches", ["enable-automation"])
        _driver = Chrome(options=opts)
        _driver.set_page_load_timeout(60)
        _driver.set_script_timeout(30)
        _driver_visible = visible
        return _driver
    except Exception as e:
        log(f"   [warn] could not launch Chrome: {type(e).__name__}: {e}")
        return None


_ERROR_TITLE = re.compile(r"\b(40[34]|5\d\d)\b|just a moment|attention required|access denied|"
                          r"not found|error code|unknown error", re.I)


def _load_real_page(driver, url):
    """Page HTML, or "" if the browser only got an error/challenge page."""
    try:
        driver.get(url)
    except Exception as e:
        # A slow third-party script can keep the page "loading" past the
        # timeout even though the post itself is already there - stop the
        # load and read what arrived.
        log(f"   Page was slow to finish loading ({type(e).__name__}) - reading what has loaded.")
        try:
            driver.execute_script("window.stop();")
        except Exception:
            pass
    html = ""
    for _ in range(5):  # a Cloudflare check can take a few seconds to pass
        time.sleep(4)
        html = driver.page_source or ""
        if not _ERROR_TITLE.search(driver.title or ""):
            break
    # An error or challenge page is not the post - confirmed real case: a
    # blocked request served Cloudflare's "520" page, which would have been
    # analysed as if it were the article.
    if _ERROR_TITLE.search(driver.title or ""):
        log(f"   [warn] the browser got an error page (\"{driver.title}\").")
        return ""
    real = [p for p in BeautifulSoup(html, "html.parser").find_all("p")
            if len(p.get_text(" ", strip=True).split()) >= 8]
    return html if len(real) >= 3 else ""


def fetch_in_browser(url):
    """Hidden browser first, then a normal visible window. Once a site needs
    the browser, every later request to it goes through the browser too."""
    for visible in (False, True):
        if visible:
            log("   Hidden browser was blocked too - opening a normal Chrome window "
                "(it closes by itself when the report is done)...")
        driver = _get_driver(visible=visible)
        if not driver:
            return ""
        try:
            html = _load_real_page(driver, url)
        except Exception as e:
            log(f"   [warn] browser fetch failed: {type(e).__name__}: {e}")
            html = ""
        if html:
            _BROWSER_HOSTS.add(urlparse(url).netloc)
            return html
    return ""


class _BrowserResponse:
    def __init__(self, status, text, final_url=""):
        self.status_code, self.text, self.content = status, text, text.encode("utf-8", "ignore")
        self.url = final_url

    def json(self):
        return json.loads(self.text)


def _browser_fetch(url):
    """A request made from inside the already-open browser (same cookies and
    clearance as the page), for sites that refuse plain requests - used for
    the sitemap, other pages' titles and the media search."""
    driver = _get_driver()
    if not driver:
        return None
    try:
        if urlparse(driver.current_url).netloc != urlparse(url).netloc:
            driver.get(site_root(url) + "/")
            time.sleep(3)
        res = driver.execute_async_script("""
            var done = arguments[arguments.length - 1];
            fetch(arguments[0], {credentials: 'include'})
              .then(function (r) { return r.text().then(function (t) { done([r.status, t, r.url]); }); })
              .catch(function (e) { done([0, String(e)]); });
        """, url)
        if not res or not res[0]:
            return None
        return _BrowserResponse(int(res[0]), res[1] or "", res[2] if len(res) > 2 else "")
    except Exception as e:
        log(f"   [warn] browser request failed for {url}: {type(e).__name__}: {e}")
        return None


def close_driver():
    global _driver
    if _driver:
        try:
            _driver.quit()
        except Exception:
            pass
        _driver = None


# Age gates, email/discount popups and cookie banners sit on top of the page
# and would be all a screenshot shows - confirmed real case: an "Are you over
# 21?" gate plus a "Get my 15%" email popup covering the whole post.
_DISMISS_POPUPS_JS = """
    if (window.__popupSweep) return;
    var sweep = function () {
        var clickable = document.querySelectorAll('button, a, [role="button"], input[type="button"], input[type="submit"]');
        for (var i = 0; i < clickable.length; i++) {
            var t = (clickable[i].innerText || clickable[i].value || '').trim().toLowerCase();
            if (/^(yes|yes,? i am.*|i am (over )?\\d+.*|enter|accept( all)?|accept cookies|agree|i agree|got it|allow all)$/.test(t)) {
                if (!clickable[i].__swept) { clickable[i].__swept = true; try { clickable[i].click(); } catch (e) {} }
            }
        }
        var vw = window.innerWidth, vh = window.innerHeight, all = document.querySelectorAll('body *');
        for (var j = 0; j < all.length; j++) {
            var cs = getComputedStyle(all[j]);
            if (cs.position !== 'fixed' && cs.position !== 'absolute') continue;
            if ((parseInt(cs.zIndex, 10) || 0) < 100) continue;
            var r = all[j].getBoundingClientRect();
            var covers = (Math.min(r.right, vw) - Math.max(r.left, 0)) * (Math.min(r.bottom, vh) - Math.max(r.top, 0));
            if (covers > vw * vh * 0.25) all[j].style.setProperty('display', 'none', 'important');
        }
        document.documentElement.style.overflow = 'auto';
        document.body.style.overflow = 'auto';
    };
    sweep();
    window.__popupSweep = setInterval(sweep, 250);
"""


def _dismiss_popups(driver):
    """Keeps sweeping until the screenshot is taken - confirmed real case: the
    email popup only appeared a few seconds after the age gate was closed, so
    a one-time cleanup missed it."""
    try:
        driver.execute_script(_DISMISS_POPUPS_JS)
    except Exception:
        pass
    time.sleep(6)


def capture_annotated_screenshot(url, out_path, highlight_selector=None, pick_last=False):
    """Real screenshot of the live page, with a red box drawn around the
    element matching highlight_selector (if found) - an approximation of a
    human's placement callout, not a guarantee of exact taste. Returns
    out_path on success, None if the capture failed entirely."""
    driver = _get_driver()
    if not driver:
        return None
    try:
        try:
            driver.get(url)
        except Exception:
            driver.execute_script("window.stop();")  # slow page: capture what has loaded
        time.sleep(3 if _driver_visible else 1.5)
        _dismiss_popups(driver)
        if highlight_selector:
            try:
                driver.execute_script("""
                    // Only elements a person can actually see: a hidden search
                    // form or an empty trailing paragraph matched first and
                    // left the screenshot at the top of the page. Selector
                    // alternatives separated by "||" are tried in order.
                    var el = null, wantText = arguments[1];
                    arguments[0].split('||').some(function (sel) {
                        var els = Array.prototype.filter.call(document.querySelectorAll(sel), function (e) {
                            var r = e.getBoundingClientRect();
                            return e.offsetParent !== null && r.height > 15 && r.width > 80 &&
                                   (!wantText || (e.innerText || '').trim().length > 60);
                        });
                        el = wantText ? els[els.length - 1] : els[0];
                        return !!el;
                    });
                    if (el) {
                        el.scrollIntoView({block: 'center'});
                        el.style.outline = '4px solid #ff3b30';
                        el.style.outlineOffset = '2px';
                    }
                """, highlight_selector, pick_last)
                time.sleep(0.4)
            except Exception:
                pass
        driver.save_screenshot(out_path)
        return out_path
    except Exception as e:
        log(f"   [warn] screenshot capture failed: {type(e).__name__}: {e}")
        return None


# --------------------------------------------------------------------------- #
# Structured Data (JSON-LD) - fully real, every field scraped or known
# --------------------------------------------------------------------------- #
def build_json_ld(scraped, suggestions, image_url=""):
    root = scraped["site_root"]
    homepage = fetch(root, timeout=12)
    site_name = urlparse(root).netloc
    logo_url = ""
    if homepage and homepage.status_code < 400:
        hsoup = BeautifulSoup(homepage.text, "html.parser")
        og_site = hsoup.find("meta", attrs={"property": "og:site_name"})
        if og_site and og_site.get("content"):
            site_name = og_site["content"].strip()
        logo_img = hsoup.select_one('img[class*="logo"], header img, .site-logo img')
        if logo_img and logo_img.get("src"):
            logo_url = urljoin(root, logo_img["src"])

    today = datetime.now().strftime("%Y-%m-%d")
    published = scraped.get("date_published") or today
    title = suggestions.get("title") or ""
    if not title or title.strip().lower().startswith("no change"):
        title = scraped["existing_title"]
    return {
        "@context": "https://schema.org",
        "@type": "BlogPosting",
        "mainEntityOfPage": {"@type": "WebPage", "@id": scraped["url"]},
        "headline": title,
        "description": suggestions.get("meta_description") or scraped["existing_meta"],
        "image": image_url or scraped["featured_image"] or "",
        "author": {"@type": "Organization", "name": site_name, "url": root + "/"},
        "publisher": {"@type": "Organization", "name": site_name,
                      "logo": {"@type": "ImageObject", "url": logo_url}},
        "datePublished": published,
        "dateModified": scraped.get("date_modified") or published,
    }


def build_faq_ld(faqs):
    return {
        "@context": "https://schema.org",
        "@type": "FAQPage",
        "mainEntity": [{"@type": "Question", "name": f["q"],
                        "acceptedAnswer": {"@type": "Answer", "text": f["a"]}} for f in faqs],
    }


# --------------------------------------------------------------------------- #
# .docx helpers
# --------------------------------------------------------------------------- #
RED = RGBColor(0xEE, 0x00, 0x00)
GREEN = RGBColor(0x00, 0xB0, 0x50)
STARS_47, STARS_44, STARS_43 = "*" * 47, "*" * 44, "*" * 43
NO_CHANGE = "No Changes Require!"  # spelled exactly as in the team's reference


def add_hyperlink(paragraph, url, text, color="0563C1", bold=False, underline=True):
    """python-docx has no native hyperlink support - builds the real
    w:hyperlink run by hand, pointing at a real external relationship."""
    part = paragraph.part
    r_id = part.relate_to(url, RT.HYPERLINK, is_external=True)
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)

    run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    if bold:
        rPr.append(OxmlElement('w:b'))
    if color:
        c = OxmlElement('w:color')
        c.set(qn('w:val'), color)
        rPr.append(c)
    if underline:
        u = OxmlElement('w:u')
        u.set(qn('w:val'), 'single')
        rPr.append(u)
    run.append(rPr)
    t = OxmlElement('w:t')
    t.set(qn('xml:space'), 'preserve')
    t.text = text
    run.append(t)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)
    return hyperlink


def _run(p, text, bold=True, color=None, highlight=False):
    r = p.add_run(text)
    r.bold = bold
    if color is not None:
        r.font.color.rgb = color
    if highlight:
        r.font.highlight_color = WD_COLOR_INDEX.GRAY_25
    return r


def _line(doc, label="", value="", color=None, value_bold=True, highlight=False):
    p = doc.add_paragraph()
    if label:
        _run(p, label, highlight=highlight)
    if value:
        _run(p, value, bold=value_bold, color=color)
    return p


def _picture(doc, path, width=5.5):
    """`path` is a file path, or the raw bytes of an image."""
    try:
        doc.add_picture(io.BytesIO(path) if isinstance(path, bytes) else path, width=Inches(width))
        return True
    except Exception as e:
        log(f"   [warn] could not embed an image: {type(e).__name__}: {e}")
        return False


def _schema_block(doc, data):
    _line(doc, "Start")
    doc.add_paragraph("")
    doc.add_paragraph('<script type="application/ld+json">')
    for line in json.dumps(data, indent=2, ensure_ascii=False).splitlines():
        doc.add_paragraph(line)
    doc.add_paragraph("</script>")
    doc.add_paragraph("")
    _line(doc, "End")


def _no_change(text):
    return (text or "").strip().lower().startswith("no change")


# --------------------------------------------------------------------------- #
# One plain labelled-line layout for every platform (no tables - the plugin
# and the admin extension read these lines). Section order, labels, bold/red/
# green marks and separators follow the team's WordPress ("james") reference
# documents; `profile` drops the sections a platform has no field for. The
# only deliberate difference from the reference: where it pastes a prnt.sc
# link, the real screenshot is embedded instead.
# --------------------------------------------------------------------------- #
def build_docx(out_path, profile, scraped, suggestions, link_pairs, image, cta_ref_url,
               button_link, screenshot_paths, json_ld, faq_ld, existing_images, related_posts):
    doc = docx.Document()
    title = doc.add_heading("", level=0)
    r = _run(title, "Blog Optimization Suggestions", highlight=True)
    r.font.color.rgb = RGBColor(0, 0, 0)
    doc.add_paragraph("")
    doc.add_paragraph("               ")
    _line(doc, f"Blog URL – {scraped['url']}", highlight=True)
    doc.add_paragraph("")
    _line(doc, STARS_47)

    # ---- Additional Suggestions
    _line(doc, "Additional Suggestions-")
    if profile["category"]:
        _line(doc, "Category-")
        cat = scraped["existing_category"]
        if cat in ("Missing", MISSING, ""):
            _line(doc, "Existing Category: ", MISSING, color=RED)
        else:
            _line(doc, f"Existing Category: {cat}")
        _line(doc, f"Suggested Category: {suggestions.get('category', '')}")
    _line(doc, "Tag Suggestions (Provide Keywords Based): " + ", ".join(suggestions.get("tags") or []))
    doc.add_paragraph("")
    if related_posts and not profile["new_image"]:
        _line(doc, "Read More Suggestions:")
        for label, post in zip(("Previous Blog Topic", "Next Blog Topic"), related_posts):
            p = _line(doc, f"{label}: ")
            _run(p, post["title"], bold=False)
            p = _line(doc, "Hyperlink with: ")
            _run(p, post["url"], bold=False)
    elif scraped["from_doc"]:
        p = _line(doc, "Read More Suggestions: ")
        _run(p, "Check after publishing.", bold=False)
    elif scraped["has_read_more"]:
        _line(doc, "Read More Suggestions: ", "Already!", color=GREEN)
    else:
        _line(doc, "Read More Suggestions: ", MISSING, color=RED)
        doc.add_paragraph("Add Previous / Next post links (or a related posts block) below the article.")
    doc.add_paragraph("")
    _line(doc, STARS_47)
    doc.add_paragraph("")

    # ---- Meta Suggestions
    _line(doc, "Meta Suggestions-")
    doc.add_paragraph("")
    if scraped["existing_title"] == NOT_PUBLISHED:
        _line(doc, "Existing Title: ", NOT_PUBLISHED, color=RED)
    else:
        _line(doc, f"Existing Title: {scraped['existing_title']}")
    sug_title = suggestions.get("title") or ""
    if scraped["existing_title"] == NOT_PUBLISHED and (_no_change(sug_title) or not sug_title):
        sug_title = scraped["existing_h1"] if scraped["existing_h1"] != MISSING else ""
    if _no_change(sug_title) or not sug_title or sug_title.strip() == scraped["existing_title"].strip():
        _line(doc, "Suggested Title: ", NO_CHANGE, color=GREEN)
    else:
        _line(doc, f"Suggested Title: {sug_title}")
    doc.add_paragraph("")
    if scraped["existing_meta"] == NOT_PUBLISHED:
        _line(doc, "Existing Meta Description: ", NOT_PUBLISHED, color=RED)
    elif scraped["existing_meta"] == MISSING:
        _line(doc, "Existing Meta Description: ", "Description is missing!", color=RED)
    else:
        _line(doc, f"Existing Meta Description: {scraped['existing_meta']}")
    _line(doc, f"Suggested Meta Description: {suggestions.get('meta_description', '')}")
    doc.add_paragraph("")
    h1_missing = scraped["existing_h1"] == MISSING
    if h1_missing:
        _line(doc, "Heading Tags-\nH1 Existing: ", MISSING, color=RED)
    else:
        _line(doc, f"Heading Tags-\nH1 Existing: {scraped['existing_h1']}")
    # Team rule: on a blog the H1 IS the blog topic, so an existing H1 is never
    # changed (only the meta title is). A missing H1 gets the blog topic,
    # taken from the post's own title rather than invented.
    if not h1_missing:
        _line(doc, "Suggested H1: ", NO_CHANGE, color=GREEN)
    else:
        topic = scraped["existing_title"]
        if topic in (MISSING, NOT_PUBLISHED):
            topic = suggestions.get("h1_change") or ""
        _line(doc, "Suggested H1: " + re.split(r"\s+[-|–]\s+", topic)[0])
    doc.add_paragraph("")
    _line(doc, STARS_47)
    doc.add_paragraph("")

    if profile["cta"]:
        # ---- Call to Action Button
        _line(doc, "Suggested Call to Action Button-")
        cta_btn = suggestions.get("cta_text", "")
        p = _line(doc, "Button Text: ")
        _run(p, cta_btn if cta_btn else MISSING, bold=False,
             color=None if cta_btn else RED)
        p = _line(doc, "Reference URL")
        _run(p, f": {cta_ref_url}", bold=False)
        if screenshot_paths.get("cta_reference"):
            _line(doc, "Screenshot: ")
            _picture(doc, screenshot_paths["cta_reference"])
        else:
            _line(doc, "Screenshot: ", "could not capture automatically", color=RED)
        cta_place = suggestions.get("cta_placement", "") or "End of post"
        p = _line(doc, "Call to Action Button Location: ")
        _run(p, cta_place, bold=False)
        doc.add_paragraph("")

        # ---- Contact Us Form
        if scraped["has_contact_form"]:
            p = _line(doc, "Suggested Contact Us Form-\n")
            _run(p, "Already!", color=GREEN)
        else:
            _line(doc, "Suggested Contact Us Form-")
            doc.add_paragraph("")
            doc.add_paragraph("Add this Contact Us form on the page. You will find the form on the "
                              "Contact page from there you can\nget the code and add it.")
            if screenshot_paths.get("contact_placement"):
                _line(doc, "Screenshot: ")
                _picture(doc, screenshot_paths["contact_placement"])
            else:
                _line(doc, "Screenshot: ", "could not capture automatically", color=RED)
            if scraped["contact_url"]:
                p = _line(doc, "URL: ")
                _run(p, scraped["contact_url"], bold=False)
            else:
                _line(doc, "URL: ", MISSING, color=RED)
            p = _line(doc, "Contact us placement suggestion: ")
            _run(p, "near the end of the post, below the main content.", bold=False)
        doc.add_paragraph(STARS_44)
        for _ in range(3):
            doc.add_paragraph("")

    # ---- Internal Linking
    _line(doc, "Internal Linking Suggestion:", highlight=True)
    doc.add_paragraph("")
    if link_pairs:
        for pair in link_pairs:
            p = _line(doc, "Find With: ")
            _run(p, pair["sentence"], bold=False)
            before, _, after = pair["sentence"].partition(pair["anchor"])
            p = _line(doc, "Replace With: ")
            if before:
                _run(p, before, bold=False)
            add_hyperlink(p, pair["target_url"], pair["anchor"])
            if after:
                _run(p, after, bold=False)
            doc.add_paragraph("")
    else:
        _line(doc, "", "No confident real internal-link target was found automatically "
                       "(no sitemap, or no page matched this post's topic closely enough). "
                       "Add manually or supply target pages/keywords next time.", color=RED)
        doc.add_paragraph("")
    _line(doc, STARS_43)
    doc.add_paragraph("")

    # ---- Image Optimization
    _line(doc, "Image Optimization Suggestions-  ")
    doc.add_paragraph("")
    team_image = bool(image["files"] or image["from_team"])  # any image the team gave
    if profile["new_image"] or team_image:
        if profile["new_image"]:
            if scraped["from_doc"]:
                p = _line(doc, "Featured Image Suggestion: ")
                _run(p, "Check after publishing.", bold=False)
            elif scraped["featured_image"]:
                _line(doc, "Featured Image Suggestion: Already Given!")
            else:
                _line(doc, "Featured Image Suggestion: ", MISSING, color=RED)
        many = len(image["items"]) > 1
        if not image["items"]:
            _line(doc, "Image URL: ", "Need to find new image and update", color=RED)
            _line(doc, f"Suggested Image Alt: {suggestions.get('image_alt', '')}")
            if profile["new_image"]:
                _line(doc, f"Suggested Title Attribute: {suggestions.get('image_title_attr', '')}")
        for n, item in enumerate(image["items"], 1):
            _line(doc, f"Image {n} URL: " if many else "Image URL: ")
            if item["file"]:
                doc.add_paragraph("Provided by team (attached below). Upload it to the site and add it "
                                  "to the post.")
                _picture(doc, item["src"], width=4.0)
            else:
                doc.add_paragraph(item["src"])
            _line(doc, f"Suggested Image Alt: {item['alt']}")
            if profile["new_image"]:
                _line(doc, f"Suggested Title Attribute: {item['alt'] or suggestions.get('image_title_attr', '')}")
        if profile["cta"]:
            cta = suggestions.get("cta_text", "")
            p = _line(doc, f"Clickable Button: Create a clickable button below the image with the text "
                           f"\u201c{cta}\u201d and link it with this page-\nLink: ")
            if button_link:
                _run(p, button_link)
            else:
                _run(p, "No confident internal page found automatically. Link manually "
                        "(e.g. to a relevant product/shop page)", color=RED)
        p = _line(doc, "Placement Suggestion: ")
        if image["location"]:
            _run(p, image["location"], bold=False)
        if screenshot_paths.get("image_placement"):
            _picture(doc, screenshot_paths["image_placement"])
        elif not image["location"]:
            _run(p, "could not capture automatically. Place near the top of the content.", bold=False)
        doc.add_paragraph("")
    for im in existing_images:
        _line(doc, im["label"])
        if im["url"] in scraped["doc_images"]:
            p = _line(doc, "Image: ")
            _run(p, "from the content document (shown below)", bold=False)
            _picture(doc, scraped["doc_images"][im["url"]], width=3.5)
        else:
            _line(doc, "Image URL: ")
            doc.add_paragraph(im["url"])
        if im["alt"]:
            p = _line(doc, "Existing Alt Text: ")
            _run(p, im["alt"], bold=False)
        else:
            _line(doc, "Existing Alt Text: ", MISSING, color=RED)
        if im.get("suggested"):
            _line(doc, f"Suggested Alt Text: {im['suggested']}")
        else:
            _line(doc, "Suggested Alt Text: ", "Could not generate automatically. Write manually.", color=RED)
        doc.add_paragraph("")
    if any(not im.get("seen") for im in existing_images):
        doc.add_paragraph("Check each suggested alt text against the image itself before saving.")
        doc.add_paragraph("")
    else:
        doc.add_paragraph("No images found in the content document."
                          if scraped["from_doc"] else "No images found inside the post.")
        doc.add_paragraph("")
    doc.add_paragraph(STARS_44)
    doc.add_paragraph("")

    # ---- Structured Data
    _line(doc, "Structured Data:")
    doc.add_paragraph("")
    _schema_block(doc, json_ld)
    doc.add_paragraph("")
    doc.add_paragraph("")
    _line(doc, "FAQ’s Schema:")
    doc.add_paragraph("")
    if faq_ld:
        _schema_block(doc, faq_ld)
    else:
        _line(doc, "", "No FAQ section found on the page. Add FAQs to the post first, "
                       "then add FAQ schema for them.", color=RED)

    doc.save(out_path)


# --------------------------------------------------------------------------- #
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("url", help="Full blog post URL to analyze")
    ap.add_argument("--out", required=True, help="Output directory")
    ap.add_argument("--format", default="james", choices=["james"])
    ap.add_argument("--platform", default="auto", choices=["auto"] + list(PLATFORMS))
    ap.add_argument("--targets", default=None,
                    help="Optional JSON file: target_pages, target_keywords, provided_images, "
                         "provided_image_url, image_location")
    ap.add_argument("--content-doc", default=None,
                    help="Blog content .docx for a post that is not published yet")
    ap.add_argument("--no-capture", action="store_true", help="skip live screenshots")
    args = ap.parse_args()

    os.makedirs(args.out, exist_ok=True)

    target_pages, target_keywords, provided_images = [], [], []
    provided_image_urls, image_location, target_pairs = [], "", []
    if args.targets and os.path.exists(args.targets):
        with open(args.targets, encoding="utf-8") as f:
            tdata = json.load(f)
        target_pages = tdata.get("target_pages") or []
        target_keywords = tdata.get("target_keywords") or []
        provided_images = [p for p in (tdata.get("provided_images") or []) if os.path.exists(p)]
        provided_image_urls = [u.strip() for u in re.split(r"[\n,]+", tdata.get("provided_image_url") or "")
                               if u.strip()]
        target_pairs = [p for p in (tdata.get("target_pairs") or []) if p.get("url")]
        target_pages = list(dict.fromkeys(target_pages + [p["url"] for p in target_pairs]))
        target_keywords = list(dict.fromkeys(target_keywords
                                             + [p["keyword"] for p in target_pairs if p.get("keyword")]))
        image_location = (tdata.get("image_location") or "").strip()

    content_doc = args.content_doc if args.content_doc and os.path.exists(args.content_doc) else None
    if content_doc and not (provided_images or provided_image_urls):
        log("[ERROR] For a post that is not published yet, add the image file(s) or an image URL to insert in the post.")
        sys.exit(1)
    if content_doc:
        log(f"[1/6] Reading the blog content document (site: {site_root(args.url)})...")
    else:
        log(f"[1/6] Fetching {args.url}...")
    try:
        scraped = scrape_blog(args.url, content_doc)
        if content_doc and len(scraped["paragraphs"]) < 3:
            raise Exception("the content document has almost no text in it")
    except Exception as e:
        # No real page content was ever fetched, so there's nothing real to
        # build suggestions from - never fabricate one. Still exits non-zero
        # (the run genuinely failed), but the error is now specific rather
        # than a bare traceback, so the team can tell "site unreachable" from
        # every other possible cause.
        log(f"[ERROR] Could not fetch the blog URL after retries: {e}")
        log("   Check the URL is correct and the post is published and reachable "
            "(a scheduled/draft post has no public page yet), then try again.")
        raise

    platform = scraped["platform"] if args.platform == "auto" else args.platform
    if platform not in PLATFORMS:
        log("   [warn] Could not tell which platform this site runs on - using WordPress. "
            "Pick the platform in the tool if that is wrong.")
        platform = "wordpress"
    profile = PLATFORMS[platform]
    log(f"   Platform: {profile['label']}" + (" (auto-detected)" if args.platform == "auto" else ""))

    log("[2/6] Generating AI suggestions from the real page content...")
    suggestions = ai_suggestions(scraped)

    log("[3/6] Finding real internal-link targets...")
    link_targets = find_link_targets(scraped, target_pages, target_keywords,
                                     suggestions.get("topic_keywords") or [],
                                     max_targets=profile["links"], fetch_cap=6 + 3 * profile["links"])
    link_pairs = (pick_link_sentences(scraped["paragraphs"], link_targets, profile["links"],
                                      target_keywords, target_pairs) if link_targets else [])
    if link_pairs:
        log(f"   Found {len(link_pairs)} real internal-link opportunity(ies).")
    else:
        log("   No confident internal-link target found.")
    related_posts = []
    if not profile["new_image"]:
        related_posts = find_related_posts(scraped, suggestions.get("topic_keywords") or [])
        log(f"   {len(related_posts)} related post(s) found for the read more links.")
    button_link = (link_pairs[0]["target_url"] if link_pairs
                   else link_targets[0]["url"] if link_targets else "")
    cta_ref_url = link_targets[0]["url"] if link_targets else scraped["site_root"] + "/"

    image = {"files": provided_images, "url": "", "location": image_location,
             "from_team": bool(provided_image_urls), "items": []}
    if provided_images or provided_image_urls:
        if provided_images:
            log(f"[4/6] Using {len(provided_images)} team-provided image file(s) instead of auto-detection.")
        if provided_image_urls and provided_images:
            log(f"   Plus {len(provided_image_urls)} team-provided image URL(s).")
        elif provided_image_urls:
            log(f"[4/6] Using {len(provided_image_urls)} team-provided image URL(s) instead of auto-detection.")
        image["url"] = provided_image_urls[0] if provided_image_urls else ""
    elif profile["new_image"]:
        log("[4/6] Looking for a reusable real image...")
        image["url"] = find_reusable_image(scraped, suggestions.get("topic_keywords") or []) or ""
        if image["url"]:
            log(f"   Found: {image['url']}")
        else:
            log("   No matching real image found on the site - flagged for manual sourcing.")
    else:
        log("[4/6] Checking the alt text of the post's existing images...")

    # Keywords the alt text should be built from: the team's own target
    # keywords first, then the tags and topic keywords found for this post.
    alt_keywords = list(dict.fromkeys(target_keywords + (suggestions.get("tags") or [])
                                      + (suggestions.get("topic_keywords") or [])))[:12]
    existing_images = []
    if profile["content_image_alts"]:
        if profile["featured_alt"] and scraped["featured_image"]:
            existing_images.append({"label": "Featured Image-", "url": scraped["featured_image"],
                                    "alt": scraped["featured_alt"]})
        for i, im in enumerate(scraped["content_images"], 1):
            existing_images.append({"label": f"Content Image {i}-", **im})
        suggest_image_alts(scraped, existing_images, alt_keywords)
        log(f"   {len(existing_images)} existing image(s) listed for alt text "
            f"({sum(1 for im in existing_images if im['seen'])} read by looking at the image).")

    # The image being added: alt text from the picture itself where possible.
    # Every image being added (files and URLs can both be given, and there
    # can be several): each gets its own alt text, written from the picture.
    sources = ([(f, True) for f in provided_images] + [(u, False) for u in provided_image_urls]
               or ([(image["url"], False)] if image["url"] else []))
    used_alts = [im["suggested"] for im in existing_images if im.get("suggested")]
    for src, is_file in sources:
        alt = ai_alt_from_image(scraped, src, "", alt_keywords, used_alts) or suggestions.get("image_alt", "")
        used_alts.append(alt)
        image["items"].append({"src": src, "file": is_file, "alt": alt})
    if image["items"]:
        log(f"   {len(image['items'])} image(s) to add, each with its own alt text.")

    screenshot_paths = {}
    if not args.no_capture:
        log("[5/6] Capturing best-effort placement screenshots...")
        content_p = ".entry-content p, article p, main p"
        shots = []
        if profile["cta"]:
            shots.append(("cta_reference", cta_ref_url, None, False))
            # cta_location screenshot removed: placement is now a text description from AI
        if not content_doc and (profile["new_image"] or provided_images or provided_image_urls):
            shots.append(("image_placement", args.url, "article img, .entry-content img, main img", False))
        if profile["contact_form"] and not scraped["has_contact_form"] and scraped["contact_url"]:
            shots.append(("contact_placement", scraped["contact_url"], "form:has(textarea)||form", False))
        for key, shot_url, selector, pick_last in shots:
            path = os.path.join(args.out, f"_{key}.png")
            if capture_annotated_screenshot(shot_url, path, selector, pick_last):
                screenshot_paths[key] = path
        close_driver()
    else:
        log("[5/6] Screenshot capture skipped (--no-capture).")

    json_ld = build_json_ld(scraped, suggestions, image["url"])
    faq_ld = build_faq_ld(scraped["faqs"]) if scraped["faqs"] else None
    log(f"   FAQ on page: {len(scraped['faqs'])} question(s).")

    _u = urlparse(args.url)
    _slug = _u.path.strip("/")
    if _slug.endswith(".php") or not _slug:
        _slug = f"{_u.netloc} {_u.query.split('=')[-1]}"
    safe_name = re.sub(r"[^a-zA-Z0-9]+", "_", _slug).strip("_")[:60] or "post"
    out_file = os.path.join(args.out, f"Blog_Optimization_Suggestions_-_{safe_name}.docx")
    log("[6/6] Writing .docx...")
    build_docx(out_file, profile, scraped, suggestions, link_pairs, image, cta_ref_url,
               button_link, screenshot_paths, json_ld, faq_ld, existing_images, related_posts)

    for p in list(screenshot_paths.values()) + provided_images + ([content_doc] if content_doc else []):
        try:
            os.remove(p)
        except Exception:
            pass

    log(f"[DONE] {out_file}")


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except Exception as e:
        log(f"[ERROR] {type(e).__name__}: {e}")
        sys.exit(1)
