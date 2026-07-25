"""
SEO Toolkit Pro — OTA Update Module
Checks a remote manifest for updated files and downloads them on startup.
No reinstall needed — only Python scripts, templates, and static files are updated.
"""

import os
import json
import time
import hashlib
import urllib.request
import shutil
from datetime import datetime

BUNDLE_DIR = os.path.dirname(os.path.abspath(__file__))
UPDATE_LOG = os.path.join(BUNDLE_DIR, ".update_log")

# Remote manifest URL(s), tried in order until one responds - each entry is a
# complete, self-contained manifest (its files' own "url" fields already point
# at that same source), not a list of mirrors for one manifest, so whichever
# source answers first is used for every file in that update pass, never a mix.
# Manifest JSON format: {"version": "3.3", "files": [{"path": "health_audit.py", "hash": "sha256...", "url": "https://..."}]}
# VPS is tried first - raw.githubusercontent.com is a known target for
# corporate web-filter/antivirus HTTPS interception (confirmed real case: a
# machine's proxy substituted a block page for every file, every retry). A
# domain on our own VPS starts with no such filter-category reputation, so it
# sidesteps that specific failure mode; GitHub stays as the fallback in case
# the VPS itself is ever down, so there's still a working path either way.
GITHUB_FILE_BASE = "https://raw.githubusercontent.com/apex-on-pageautomationtools/seo-toolkit-updates/main/"
VPS_FILE_BASE = "https://indexing.weblinkbuzz.com/ota-updates/"


def _swap_source(url):
    """Return the same file's URL at the OTHER mirror (VPS <-> GitHub), or None if
    url doesn't match either known base. Used for per-file fallback when whichever
    source answered the manifest fetch turns out to be the one a network is
    intercepting - previously the whole update just gave up once any file came
    back intercepted, even though the other mirror (confirmed reachable moments
    earlier, or not yet tried at all) might work fine for the exact same files."""
    for base, other in ((GITHUB_FILE_BASE, VPS_FILE_BASE), (VPS_FILE_BASE, GITHUB_FILE_BASE)):
        if url.startswith(base):
            return other + url[len(base):]
    return None


UPDATE_MANIFEST_URLS = [
    "https://indexing.weblinkbuzz.com/ota-updates/update_manifest.json",
    "https://raw.githubusercontent.com/apex-on-pageautomationtools/seo-toolkit-updates/main/update_manifest.json",
]


def _file_hash(filepath):
    """SHA-256 hash of a local file, line-ending normalized (CRLF -> LF) - this
    matches exactly how the manifest's own hashes are generated. Some networks
    (corporate proxy / antivirus / SSL-inspection) silently rewrite line endings in
    transit, which produced a permanent false 'hash mismatch' on affected machines
    even though the downloaded content was functionally identical. Normalizing here
    makes the comparison immune to that regardless of what's altering bytes in
    transit."""
    try:
        with open(filepath, "rb") as f:
            data = f.read()
        return hashlib.sha256(data.replace(b"\r\n", b"\n")).hexdigest()
    except Exception:
        return ""


def _cache_busted(url):
    """Append a unique query param so no proxy/CDN/ISP cache between here and GitHub
    can ever serve a stale response - a stuck cache was confirmed to be why updater.py
    itself went nearly a week without ever picking up a real update: hash comparisons
    were silently running against a frozen-in-time manifest."""
    sep = "&" if "?" in url else "?"
    return f"{url}{sep}_cb={int(time.time() * 1000)}"


_NO_CACHE_HEADERS = {"User-Agent": "SEOToolkitPro-Updater/1.0",
                    "Cache-Control": "no-cache, no-store, must-revalidate",
                    "Pragma": "no-cache"}


def _fetch_json(url, timeout=30):
    """Fetch JSON from a URL."""
    try:
        req = urllib.request.Request(_cache_busted(url), headers=_NO_CACHE_HEADERS)
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return json.loads(r.read().decode("utf-8"))
    except Exception:
        return None


def _looks_intercepted(path):
    """A downloaded .py/.html/.json file that starts with an HTML doctype/tag (or
    a captive-portal/login redirect) means something between here and GitHub
    substituted a block/login page for the real content - a corporate proxy or
    antivirus web filter, not a real network failure. Confirmed live: a machine
    hit a 100% hash-mismatch rate on EVERY file, every retry, for every update
    cycle - a pattern ordinary flakiness doesn't produce, since flaky networks
    fail intermittently, not deterministically on everything. Detecting this
    specifically lets the log say what's actually wrong instead of just
    'hash mismatch', which looks identical to a transient blip."""
    try:
        with open(path, "rb") as f:
            head = f.read(512).lstrip().lower()
    except Exception:
        return False
    return head.startswith((b"<!doctype html", b"<html"))


def _download_file(url, dest, timeout=90):
    """Download a file from URL to dest path. Downloads to a temp file first, then
    atomically replaces dest - a failure/interruption partway through never leaves a
    truncated file sitting at dest for the hash check to trip over. Returns
    (ok, error_message) so callers can log WHY a download failed, not just that it
    did - a swallowed exception was making every failure look identical/unexplained."""
    tmp = dest + ".part"
    try:
        req = urllib.request.Request(_cache_busted(url), headers=_NO_CACHE_HEADERS)
        with urllib.request.urlopen(req, timeout=timeout) as r:
            os.makedirs(os.path.dirname(dest) or ".", exist_ok=True)
            with open(tmp, "wb") as f:
                shutil.copyfileobj(r, f)
        # Antivirus real-time scanning can lock a just-written file before the
        # rename - confirmed on a real machine to sometimes outlast a few seconds
        # of backoff (site-packages/*.zip paths are a common AV/EDR scan target).
        # 10 tries with growing backoff, capped at ~2s/try (~20s total worst case),
        # gives real headroom without hanging a launch indefinitely if the lock
        # truly never clears (in which case the exclusion is the actual fix).
        for _r in range(10):
            try:
                os.replace(tmp, dest)
                break
            except PermissionError:
                if _r == 9:
                    raise
                time.sleep(min(2.0, 0.4 * (_r + 1)))
        return True, ""
    except Exception as e:
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except Exception:
            pass
        return False, f"{type(e).__name__}: {e}"


def _log_update(msg):
    """Append to update log."""
    try:
        with open(UPDATE_LOG, "a", encoding="utf-8") as f:
            f.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}\n")
    except Exception:
        pass


def check_and_update(log_fn=None):
    """Check for updates and apply them. Returns dict with update info.
    Call this on app startup — it's fast (single HTTP call) if nothing changed."""
    if log_fn is None:
        log_fn = print

    if not UPDATE_MANIFEST_URLS:
        return {"updated": False, "reason": "No update URL configured"}

    # Prevent two updaters (launcher background run + in-app auto-check) from writing
    # the same files at once, which could corrupt them. Stale locks (>5 min) are ignored.
    _lock = os.path.join(BUNDLE_DIR, ".update_lock")
    _have_lock = False
    try:
        if os.path.exists(_lock) and (time.time() - os.path.getmtime(_lock)) < 300:
            return {"updated": False, "reason": "Update already in progress"}
        with open(_lock, "w") as _lf:
            _lf.write(str(int(time.time())))
        _have_lock = True
    except Exception:
        pass

    def _unlock():
        if _have_lock and os.path.exists(_lock):
            try:
                os.remove(_lock)
            except Exception:
                pass

    # A transient network blip used to kill the ENTIRE update check for that launch -
    # someone who only restarts occasionally could go a long time genuinely never
    # getting updates just from bad luck on one attempt. Retry the manifest fetch
    # itself a couple times per source before moving to the next source, and try
    # every configured source (VPS, then GitHub) before giving up for this launch.
    manifest = None
    for _source_url in UPDATE_MANIFEST_URLS:
        for _attempt in range(2):
            manifest = _fetch_json(_source_url)
            if manifest:
                break
            if _attempt < 1:
                time.sleep(1.5 * (_attempt + 1))
        if manifest:
            break
    if not manifest:
        _unlock()
        return {"updated": False, "reason": "Could not reach update server"}

    remote_version = manifest.get("version", "0")
    files = manifest.get("files", [])
    if not files:
        _unlock()
        return {"updated": False, "reason": "No files in manifest", "remote_version": remote_version}

    updated_files = []
    skipped = []
    failed = []
    failed_reasons = {}
    intercepted = False

    # Once a file confirms one mirror is intercepted and the other isn't, remember
    # that winning base and use it FIRST for every remaining file - otherwise each
    # subsequent file would independently rediscover the same interception before
    # falling back, burning a full retry cycle per file for no reason.
    switched_base = None
    both_sources_blocked = False

    for entry in files:
        rel_path = entry.get("path", "")
        remote_hash = entry.get("hash", "")
        download_url = entry.get("url", "")

        if not rel_path or not download_url:
            continue

        if switched_base:
            _alt = _swap_source(download_url)
            if _alt:
                download_url = _alt

        # Protected files that should never be auto-updated
        if rel_path in (".auth_token", "config.json", ".update_log"):
            skipped.append(rel_path)
            continue

        local_path = os.path.join(BUNDLE_DIR, rel_path.replace("/", os.sep))
        local_hash = _file_hash(local_path)

        if local_hash == remote_hash and remote_hash:
            skipped.append(rel_path)
            continue

        # Backup existing file
        if os.path.exists(local_path):
            backup = local_path + ".bak"
            try:
                shutil.copy2(local_path, backup)
            except Exception:
                pass

        log_fn(f"[update] Downloading {rel_path}...")
        # Retry on flaky/slow networks. A single failed or truncated download used to
        # leave that file stale (so a feature never updated on some machines). Verify
        # the hash each try so a corrupt/partial download is rejected and retried.
        ok = False
        last_err = ""
        file_intercepted = False
        for _attempt in range(3):
            dl_ok, err = _download_file(download_url, local_path)
            if dl_ok:
                got_hash = _file_hash(local_path)
                if not remote_hash or got_hash == remote_hash:
                    ok = True
                    break
                if _looks_intercepted(local_path):
                    _blocked_domain = download_url.split("/")[2] if "/" in download_url else download_url
                    _log_update(f"Hash mismatch (try {_attempt + 1}): {rel_path} - looks INTERCEPTED "
                                f"via {_blocked_domain} (downloaded content is an HTML page, not the "
                                f"real file) - expected {remote_hash[:8]}..., got {got_hash[:8]}...")
                    # Confirmed interception, not a fluke - retrying the SAME url won't
                    # help (the block page is a deterministic response). Before giving
                    # up on this file, try the other mirror (VPS <-> GitHub) for this
                    # exact file - the interception is specific to whichever domain
                    # the network filter has flagged, not the file itself, so the other
                    # mirror is often still reachable.
                    _alt_url = _swap_source(download_url)
                    if _alt_url:
                        _alt_domain = _alt_url.split("/")[2]
                        _log_update(f"{rel_path}: retrying via {_alt_domain} instead of {_blocked_domain}")
                        dl_ok2, _ = _download_file(_alt_url, local_path)
                        if dl_ok2:
                            got_hash2 = _file_hash(local_path)
                            if not remote_hash or got_hash2 == remote_hash:
                                ok = True
                                switched_base = _alt_url.rsplit(rel_path, 1)[0] if rel_path in _alt_url else None
                                _log_update(f"{rel_path}: succeeded via {_alt_domain} - using it for remaining files too")
                                break
                            if _looks_intercepted(local_path):
                                _log_update(f"{rel_path}: ALSO intercepted via {_alt_domain} - both mirrors blocked")
                                both_sources_blocked = True
                    last_err = (f"network is intercepting this download (a proxy/antivirus web filter is "
                                f"substituting a block or login page for the real file) - ask IT to allow "
                                f"{_blocked_domain}" + (f" and {_alt_url.split('/')[2]}" if _alt_url else "") +
                                ", a normal retry won't help")
                    file_intercepted = True
                    break
                else:
                    last_err = "hash mismatch after download"
                    _log_update(f"Hash mismatch (try {_attempt + 1}): {rel_path} - "
                                f"expected {remote_hash[:8]}..., got {got_hash[:8]}...")
            else:
                last_err = err
                _log_update(f"Download error (try {_attempt + 1}) for {rel_path}: {err}")
            if _attempt < 2:
                time.sleep(2 * (_attempt + 1))
        if ok:
            updated_files.append(rel_path)
            _log_update(f"Updated: {rel_path} ({local_hash[:8]}... -> {remote_hash[:8]}...)")
            # Remove backup on success
            backup = local_path + ".bak"
            if os.path.exists(backup):
                try:
                    os.remove(backup)
                except Exception:
                    pass
        else:
            failed.append(rel_path)
            failed_reasons[rel_path] = last_err
            _log_update(f"Failed: {rel_path} - {last_err}")
            # Roll back: restore the previous good version, or drop a corrupt new file
            backup = local_path + ".bak"
            if os.path.exists(backup):
                try:
                    shutil.copy2(backup, local_path)
                    os.remove(backup)
                except Exception:
                    pass
            else:
                try:
                    if os.path.exists(local_path):
                        os.remove(local_path)
                except Exception:
                    pass

        # Only abort the whole batch when BOTH mirrors are confirmed blocked - a
        # single-source interception that the per-file fallback above already
        # recovered from (ok is True) should just continue to the next file, now
        # using switched_base for the rest of this run.
        if both_sources_blocked:
            intercepted = True
            break

    result = {
        "updated": len(updated_files) > 0,
        "remote_version": remote_version,
        "updated_files": updated_files,
        "skipped": len(skipped),
        "failed": failed,
        "failed_reasons": failed_reasons,
    }
    if intercepted:
        result["reason"] = ("Network is intercepting the update download (a proxy/antivirus web filter is "
                            "substituting a block or login page for the real file) - both mirrors "
                            "(GitHub and our own VPS) came back blocked for at least one file, so this "
                            "isn't one specific domain to allowlist. Stopped early instead of retrying "
                            "every remaining file the same way.")

    if updated_files:
        log_fn(f"[update] {len(updated_files)} file(s) updated to v{remote_version}")
        _log_update(f"Update complete: {len(updated_files)} files, v{remote_version}")
    elif failed:
        log_fn(f"[update] {len(failed)} file(s) failed to update - NOT up to date")
    else:
        log_fn("[update] Everything up to date")

    _unlock()
    return result


def generate_manifest(directory=None, base_url=""):
    """Helper: generate a manifest.json for the current files.
    Run this locally to create the manifest you upload to GitHub/server.
    Usage: python updater.py --generate --base-url https://raw.githubusercontent.com/you/repo/main/"""
    if directory is None:
        directory = BUNDLE_DIR

    INCLUDE_PATTERNS = {
        "web_app_batch.py", "engine.py", "da_checker.py", "health_audit.py", "gsc_audit.py", "auth.py", "updater.py",
    }
    INCLUDE_DIRS = {"templates", "static", "scripts"}

    files = []

    # Single files
    for fname in INCLUDE_PATTERNS:
        fpath = os.path.join(directory, fname)
        if os.path.exists(fpath):
            files.append({
                "path": fname,
                "hash": _file_hash(fpath),
                "url": base_url + fname,
            })

    # Directory files
    for dname in INCLUDE_DIRS:
        dpath = os.path.join(directory, dname)
        if not os.path.isdir(dpath):
            continue
        for root, _, fnames in os.walk(dpath):
            for fn in fnames:
                full = os.path.join(root, fn)
                rel = os.path.relpath(full, directory).replace(os.sep, "/")
                files.append({
                    "path": rel,
                    "hash": _file_hash(full),
                    "url": base_url + rel,
                })

    from importlib.metadata import version as pkg_version
    try:
        app_ver = "3.9"
    except Exception:
        app_ver = "3.9"

    manifest = {"version": app_ver, "files": files}
    out = os.path.join(directory, "update_manifest.json")
    with open(out, "w", encoding="utf-8") as f:
        json.dump(manifest, f, indent=2)
    print(f"Manifest written: {out} ({len(files)} files)")
    return manifest


if __name__ == "__main__":
    import sys
    if "--generate" in sys.argv:
        base = ""
        for i, a in enumerate(sys.argv):
            if a == "--base-url" and i + 1 < len(sys.argv):
                base = sys.argv[i + 1]
        generate_manifest(base_url=base)
    else:
        result = check_and_update()
        print(json.dumps(result, indent=2))
