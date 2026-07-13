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

# Remote manifest URL — set to your GitHub raw URL or Apps Script endpoint
# Manifest JSON format: {"version": "3.3", "files": [{"path": "health_audit.py", "hash": "sha256...", "url": "https://..."}]}
UPDATE_MANIFEST_URL = "https://raw.githubusercontent.com/apex-on-pageautomationtools/seo-toolkit-updates/main/update_manifest.json"


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
        # Antivirus real-time scanning can briefly lock a just-written file before
        # the rename, especially for .zip/.exe-adjacent paths - retry the rename a
        # few times instead of failing outright on a transient PermissionError.
        for _r in range(5):
            try:
                os.replace(tmp, dest)
                break
            except PermissionError:
                if _r == 4:
                    raise
                time.sleep(0.5 * (_r + 1))
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

    if not UPDATE_MANIFEST_URL:
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

    manifest = _fetch_json(UPDATE_MANIFEST_URL)
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

    for entry in files:
        rel_path = entry.get("path", "")
        remote_hash = entry.get("hash", "")
        download_url = entry.get("url", "")

        if not rel_path or not download_url:
            continue

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
        for _attempt in range(3):
            dl_ok, err = _download_file(download_url, local_path)
            if dl_ok:
                if not remote_hash or _file_hash(local_path) == remote_hash:
                    ok = True
                    break
                last_err = "hash mismatch after download"
                _log_update(f"Hash mismatch (try {_attempt + 1}): {rel_path}")
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

    result = {
        "updated": len(updated_files) > 0,
        "remote_version": remote_version,
        "updated_files": updated_files,
        "skipped": len(skipped),
        "failed": failed,
        "failed_reasons": failed_reasons,
    }

    if updated_files:
        log_fn(f"[update] {len(updated_files)} file(s) updated to v{remote_version}")
        _log_update(f"Update complete: {len(updated_files)} files, v{remote_version}")
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
