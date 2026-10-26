"""Reddit Research - search real Reddit posts/comments by keyword and/or
subreddit via Arctic Shift, a free public Reddit archive API. No Reddit API
key, OAuth, or login needed, so it has none of the official API's
registration/verification requirements.
"""
import time
from datetime import datetime, timezone

import requests

_SESSION = requests.Session()
_SESSION.headers.update({"User-Agent": "SEOToolkitPro/1.0"})
_BASE = "https://arctic-shift.photon-reddit.com/api"

_TIME_FILTER_SECONDS = {
    "hour": 3600, "day": 86400, "week": 604800,
    "month": 2592000, "year": 31536000,
}


class RedditSearchError(Exception):
    """A real, user-facing error from Arctic Shift (bad params, etc.) -
    distinct from a transient network hiccup, which is retried instead."""


def _fetch(url, params=None, retries=5):
    last_exc = None
    for attempt in range(retries):
        try:
            r = _SESSION.get(url, params=params, timeout=15)
            # 429 is the standard rate-limit status; Arctic Shift also uses
            # its own 422 "Timeout. Maybe slow down a bit" for the same thing
            # (confirmed live - same request succeeds a few seconds later),
            # so both get a backoff-and-retry instead of being treated as a
            # real client error.
            if r.status_code in (429, 422):
                last_exc = RedditSearchError("Reddit's archive API is rate-limiting requests - try again in a moment.")
                time.sleep(int(r.headers.get("Retry-After", 5)))
                continue
            # Any other 4xx is a real "your request is invalid" error (e.g.
            # Arctic Shift rejects title-only search with no subreddit/
            # author) - surface it instead of silently retrying/giving up and
            # reporting "0 results found", which looks like a working search
            # that just found nothing rather than a broken request.
            if 400 <= r.status_code < 500:
                try:
                    msg = r.json().get("error") or r.text
                except Exception:
                    msg = r.text
                raise RedditSearchError(msg.strip() or f"HTTP {r.status_code}")
            r.raise_for_status()
            return r.json()
        except RedditSearchError:
            raise
        except Exception as e:
            last_exc = e
            if attempt < retries - 1:
                time.sleep(2)
    if isinstance(last_exc, RedditSearchError):
        raise last_exc
    if last_exc:
        raise RedditSearchError(f"Could not reach Reddit's archive API: {last_exc}")
    return None


def _ts(ts):
    try:
        return datetime.fromtimestamp(float(ts), tz=timezone.utc).strftime("%Y-%m-%d %H:%M")
    except Exception:
        return ""


def search(keyword=None, subreddit=None, time_filter="all", limit=25):
    """Search Reddit posts by keyword (title) within a subreddit.

    Arctic Shift's /posts/search rejects a 'title' filter unless 'subreddit'
    (or 'author') is also given - there's no true site-wide keyword search on
    this free API, confirmed live (returns HTTP 400 "'title' query parameter
    requires one of: author, subreddit"). So subreddit is required; keyword
    just narrows within it.
    """
    if keyword and not subreddit:
        raise RedditSearchError(
            "Enter a subreddit too - this free Reddit API can't search by "
            "keyword across all of Reddit, only within a specific subreddit.")

    params = {"limit": min(max(int(limit or 25), 1), 100), "sort": "desc"}
    if keyword:
        params["title"] = keyword
    if subreddit:
        params["subreddit"] = subreddit
    if time_filter in _TIME_FILTER_SECONDS:
        params["after"] = int(time.time()) - _TIME_FILTER_SECONDS[time_filter]

    data = _fetch(f"{_BASE}/posts/search", params)
    if not data:
        return []

    posts = []
    for p in data.get("data", []):
        permalink = p.get("permalink", "")
        if permalink and not permalink.startswith("http"):
            permalink = "https://www.reddit.com" + permalink
        posts.append({
            "post_id": p.get("id", ""),
            "title": p.get("title", ""),
            "selftext": p.get("selftext", ""),
            "subreddit": p.get("subreddit", ""),
            "author": p.get("author", ""),
            "score": p.get("score", 0),
            "num_comments": p.get("num_comments", 0),
            "permalink": permalink,
            "created_utc": _ts(p.get("created_utc", 0)),
            "flair": p.get("link_flair_text") or "",
        })

    posts.sort(key=lambda x: int(x.get("score") or 0), reverse=True)
    return posts


def get_comments(post_id, limit=5):
    """Fetch top-scored comments for a Reddit post."""
    if not limit:
        return []
    link_id = post_id if post_id.startswith("t3_") else f"t3_{post_id}"
    params = {"link_id": link_id, "limit": limit * 3, "sort": "desc"}
    try:
        data = _fetch(f"{_BASE}/comments/search", params)
    except RedditSearchError:
        # Comments are a secondary enrichment per post - one post's comment
        # fetch failing (rate limit, etc.) shouldn't kill the whole post list.
        return []
    if not data:
        return []

    comments = []
    for c in data.get("data", []):
        body = (c.get("body") or "").strip()
        if not body or body in ("[deleted]", "[removed]"):
            continue
        comments.append({
            "comment_id": c.get("id", ""),
            "author": c.get("author", ""),
            "body": body,
            "score": c.get("score", 0),
            "created_utc": _ts(c.get("created_utc", 0)),
        })

    comments.sort(key=lambda x: int(x.get("score") or 0), reverse=True)
    return comments[:limit]


def scrape_topic(keyword=None, subreddit=None, time_filter="all", limit=25, comments_per_post=5):
    """Search posts and fetch comments for each; returns [{"post": .., "comments": [..]}]."""
    posts = search(keyword=keyword, subreddit=subreddit, time_filter=time_filter, limit=limit)
    results = []
    for post in posts:
        comments = get_comments(post["post_id"], limit=comments_per_post)
        results.append({"post": post, "comments": comments})
        time.sleep(0.5)
    return results
