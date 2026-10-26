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


def _fetch(url, params=None, retries=3):
    for attempt in range(retries):
        try:
            r = _SESSION.get(url, params=params, timeout=15)
            if r.status_code == 429:
                time.sleep(int(r.headers.get("Retry-After", 10)))
                continue
            r.raise_for_status()
            return r.json()
        except Exception:
            if attempt < retries - 1:
                time.sleep(2)
    return None


def _ts(ts):
    try:
        return datetime.fromtimestamp(float(ts), tz=timezone.utc).strftime("%Y-%m-%d %H:%M")
    except Exception:
        return ""


def search(keyword=None, subreddit=None, time_filter="all", limit=25):
    """Search Reddit posts by keyword (title) and/or subreddit."""
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
    params = {"link_id": post_id, "limit": limit * 3, "sort": "desc"}
    data = _fetch(f"{_BASE}/comments/search", params)
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
