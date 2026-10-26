"""Builds a "Reddit FAQ Recommendations" DOCX from Reddit Research results -
real question-shaped post titles found on a subreddit, with where each one
was found and why answering it on the website helps AI-engine citations.

Format mirrors a real client deliverable (Enfortra_Reddit_FAQ_Recommendations.docx):
a title, a short "what we did and why" intro, then one Qn. heading per
question followed by a 3-row key/value table (Found on / Answer to add on
website / Why it helps with AI results), and a closing "how to use these"
section. The intro/outro copy is generic GEO guidance, not client-specific,
so it's reused verbatim.

"Answer to add on website" is deliberately left as a fill-in placeholder,
not auto-written from scraped comments - the business's actual answer needs
real product/policy knowledge this tool has no way to know, and passing off
unedited Reddit replies as finished website copy would be actively
misleading. The real top comment(s) are instead included as reference
material right under the table so there's something to write from.
"""
import re
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

_QUESTION_STARTERS = (
    "what", "why", "how", "is", "are", "does", "do", "can", "could",
    "should", "will", "when", "where", "who", "which", "would",
)

COL0_WIDTH = Inches(1.25)
COL1_WIDTH = Inches(5.25)


def _looks_like_question(title):
    t = title.strip()
    if not t:
        return False
    if t.rstrip().endswith("?"):
        return True
    first_word = re.split(r"[\s:]+", t, 1)[0].lower().strip(".,!")
    return first_word in _QUESTION_STARTERS


def select_questions(results, max_questions=15):
    """Pick question-shaped posts from Reddit Research results, highest
    engagement (score) first, skipping near-duplicate titles (the same
    question cross-posted/re-asked across threads)."""
    candidates = [r for r in results if _looks_like_question(r["post"]["title"])]
    candidates.sort(key=lambda r: int(r["post"].get("score") or 0), reverse=True)
    seen_titles = set()
    deduped = []
    for r in candidates:
        norm = re.sub(r"[^a-z0-9 ]", "", r["post"]["title"].lower()).strip()
        if norm in seen_titles:
            continue
        seen_titles.add(norm)
        deduped.append(r)
    return deduped[:max_questions]


def _add_heading(doc, text, size_pt):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(size_pt)
    run.font.color.rgb = RGBColor(0, 0, 0)
    return p


def _add_body(doc, text):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.color.rgb = RGBColor(0, 0, 0)
    return p


def _set_cell_width(cell, width):
    cell.width = width


def _add_kv_table(doc, rows):
    table = doc.add_table(rows=len(rows), cols=2)
    table.style = "Normal Table"
    table.autofit = False
    for ri, (label, value) in enumerate(rows):
        row = table.rows[ri]
        label_cell, value_cell = row.cells
        _set_cell_width(label_cell, COL0_WIDTH)
        _set_cell_width(value_cell, COL1_WIDTH)

        lp = label_cell.paragraphs[0]
        lr = lp.add_run(label)
        lr.bold = True
        lr.font.color.rgb = RGBColor(0, 0, 0)

        vp = value_cell.paragraphs[0]
        vr = vp.add_run(value)
        vr.font.color.rgb = RGBColor(0, 0, 0)
    return table


def _found_on_text(post):
    return (f'r/{post["subreddit"]} - "{post["title"]}" '
            f'({post.get("score", 0)} upvotes, {post.get("num_comments", 0)} comments)')


def _why_text(post):
    return (f'This question is actively being asked on r/{post["subreddit"]} '
            f'({post.get("num_comments", 0)} comment(s) on this thread alone). If the '
            f'website does not already answer it directly, an AI engine answering the '
            f'same question is likely to cite a competitor\'s page instead of this one.')


def build_faq_docx(results, out_path, subreddit, keyword=None, max_questions=15):
    """results: the list Reddit Research's /api/reddit/search returns
    ([{"post": {...}, "comments": [...]}, ...])."""
    selected = select_questions(results, max_questions=max_questions)

    doc = Document()
    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.left_margin = section.right_margin = Inches(1)
    section.top_margin = section.bottom_margin = Inches(1)

    _add_heading(doc, "Reddit Research: Questions Found in Our Niche and "
                      "Recommended Answers to Add to Website", 12)

    _add_heading(doc, "What We Did and Why", 11)
    scope = f'r/{subreddit}' + (f', filtered by "{keyword}"' if keyword else '')
    _add_body(doc, f"Reddit was reviewed in {scope} to find the exact questions real "
                   f"people are already asking in this niche - questions they are also "
                   f"typing into AI engines like ChatGPT, Perplexity, and Google AI Overviews.")
    _add_body(doc, "When someone asks an AI engine a question, the AI pulls its answer "
                   "from whichever website already has a clear, relevant answer to that "
                   "question. If a website does not have that content, the AI cannot "
                   "reference it - and a competitor's page gets cited instead.")
    _add_body(doc, "The recommendation is simple: add these questions and their answers "
                   "to the website as a FAQ section or dedicated FAQ pages. This directly "
                   "increases the chances of the website appearing in AI-generated answers "
                   "when buyers and decision-makers ask these questions.")

    if not selected:
        _add_body(doc, "No clearly question-shaped posts were found in this search - try "
                       "a broader time range, a higher post count, or a different subreddit.")

    for i, item in enumerate(selected, 1):
        post = item["post"]
        _add_heading(doc, f"Q{i}. {post['title']}", 11)
        _add_kv_table(doc, [
            ("Found on", _found_on_text(post)),
            ("Answer to add on website", "[Write the answer here]"),
            ("Why it helps with AI results", _why_text(post)),
        ])

        comments = item.get("comments") or []
        if comments:
            ref = doc.add_paragraph()
            ref_run = ref.add_run("Reference - real Reddit replies on this thread (for writing the answer above, not for publishing as-is):")
            ref_run.italic = True
            ref_run.font.size = Pt(9)
            ref_run.font.color.rgb = RGBColor(0x59, 0x59, 0x59)
            for c in comments[:3]:
                cp = doc.add_paragraph()
                cr = cp.add_run(f'u/{c.get("author") or "[unknown]"}: {c.get("body", "")}')
                cr.italic = True
                cr.font.size = Pt(9)
                cr.font.color.rgb = RGBColor(0x59, 0x59, 0x59)

        doc.add_paragraph()

    _add_heading(doc, "How to Use These on the Website", 11)
    _add_body(doc, "Each question above can be added directly to an FAQ section on the "
                   "website, or as individual FAQ pages if the question is broad enough "
                   "to justify a full page. Write the answer in plain English, in the "
                   "website's existing tone.")
    _add_body(doc, "The question text itself should be used as the heading, exactly as "
                   "written above. AI engines match the heading language to what users "
                   "search - changing the wording significantly reduces the chance of "
                   "being cited.")
    _add_body(doc, "Adding FAQ schema markup to these questions and answers is still "
                   "recommended - not for Google, which no longer displays FAQ rich "
                   "results for most sites, but because AI engines and other bots "
                   "(Perplexity, Bing Copilot, ChatGPT plugins) read structured data "
                   "when crawling pages. Schema helps these engines understand that a "
                   "piece of content is a question and a direct answer, which improves "
                   "the chance of that content being pulled when a relevant query is made.")

    doc.save(out_path)
    return out_path
