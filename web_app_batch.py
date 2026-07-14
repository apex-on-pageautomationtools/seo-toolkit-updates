"""
SEO Toolkit Pro v3.2 - Merged edition
======================================
Combines v3 hardened engine (engine.py) with v2.3 features:
  - Country-specific Google domains + typed search + continuous scroll (v3)
  - Block recovery ladder: backoff -> proxy -> identity -> Buster -> manual (v3)
  - Browser bring-to-front on CAPTCHA (v2.3)
  - CSV export with domain name in filename (v2.3)
  - Profile 24h auto-wipe instead of fresh-every-launch (v2.3)
  - Neutral site visits between keywords (v2.3)
  - Search modes: stop_on_found / scan_all (v2.3)
  - VPN method dropdown with all options (v2.3)
  - PWA installable + dark mode (v3)
  - Config.json for proxy pool / settings (v3)
  - Autosave results (v3)
"""

import os
import sys
import io
import csv
import json
import time
import random
import shutil
import logging
import threading
import webbrowser
import subprocess
import tempfile
from datetime import datetime

from flask import (Flask, render_template, request, jsonify, send_file,
                   Response, send_from_directory)

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import engine
from engine import (BrowserClosedError, ProxyPool, clean_domain, match_domain,
                    extract_organic, find_domain_in_page, classify_page, human_search, warm_up,
                    load_more_results, is_alive, page_source, safe_get,
                    human_pause, google_domain, bring_browser_to_front,
                    human_visit_neutral, human_visit_neutral_bg,
                    check_ip_location, set_geolocation)
from da_checker import check_da_pa
import health_audit
import gsc_audit
import auth
import updater
import requests as http_requests
import brief_analysis

logging.getLogger("werkzeug").setLevel(logging.ERROR)
logging.basicConfig(level=logging.INFO, format="%(message)s")

APP_VERSION = "4.0"

# --------------------------------------------------------------------------- #
# Paths
# --------------------------------------------------------------------------- #
def _bundle_dir():
    return getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))

def _data_dir():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    # If installed in Program Files (read-only), use AppData for writable data
    if script_dir.lower().startswith(os.environ.get("ProgramFiles", "C:\\Program Files").lower()) or \
       script_dir.lower().startswith(os.environ.get("ProgramFiles(x86)", "C:\\Program Files (x86)").lower()):
        appdata = os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "SEO Toolkit Pro")
        os.makedirs(appdata, exist_ok=True)
        return appdata
    return script_dir

BUNDLE_DIR    = _bundle_dir()
DATA_DIR      = _data_dir()
UPLOADS_DIR   = os.path.join(DATA_DIR, "uploads")
_DEFAULT_DOWNLOADS = os.path.join(os.path.expanduser("~"), "Downloads")
DOWNLOADS_DIR = _DEFAULT_DOWNLOADS
TEMPLATE_DIR  = os.path.join(BUNDLE_DIR, "templates")
STATIC_DIR    = os.path.join(BUNDLE_DIR, "static")
PROFILES_DIR  = os.path.join(DATA_DIR, "profiles")
AUTOSAVE_FILE = os.path.join(DATA_DIR, "autosave_results.json")
CONFIG_FILE   = os.path.join(DATA_DIR, "config.json")
BUSTER_DIR    = os.path.join(BUNDLE_DIR, "extensions", "buster")
URBANVPN_DIR  = os.path.join(BUNDLE_DIR, "extensions", "urbanvpn")
WINDSCRIBE_DIR = os.path.join(BUNDLE_DIR, "extensions", "windscribe")
SERPCOUNTER_DIR = os.path.join(BUNDLE_DIR, "extensions", "serpcounter")

SCREENSHOTS_DIR = os.path.join(DATA_DIR, "screenshots")
# A durable copy of each report's output, kept OUTSIDE the user-facing Downloads
# folder - so "Download Report"/"Download ZIP" keeps working even if the user
# deletes/moves the folder those reports were saved into. Only the latest report per
# (category, domain) is kept - each new run overwrites the last, no manual cleanup
# needed.
REPORT_BACKUPS_DIR = os.path.join(DATA_DIR, "report_backups")

PROFILE_POOL_SIZE = 5
PROFILE_MAX_AGE_H = 168   # 7 days - keep (Google) cookies for human-like sessions; only wipe truly stale profiles

for d in (UPLOADS_DIR, DOWNLOADS_DIR, SCREENSHOTS_DIR, PROFILES_DIR, REPORT_BACKUPS_DIR):
    os.makedirs(d, exist_ok=True)


def _clear_profile_cache(profile_dir):
    """Remove cache but keep cookies, preferences, and extension state."""
    import shutil
    for sub in ("Cache", "Code Cache", "GPUCache", "Service Worker",
                "ShaderCache", "GrShaderCache", "blob_storage"):
        p = os.path.join(profile_dir, sub)
        if os.path.isdir(p):
            shutil.rmtree(p, ignore_errors=True)
        p2 = os.path.join(profile_dir, "Default", sub)
        if os.path.isdir(p2):
            shutil.rmtree(p2, ignore_errors=True)


def _prune_cookies_keep_google(profile_dir):
    """Keep Google's cookies (so the profile looks like a returning human and
    avoids CAPTCHAs/blocks) but delete every OTHER site's cookies. This keeps
    profiles small and stops visited/third-party site cookies from personalising
    and skewing the rank results. GSC logins live in a separate gsc_sessions
    folder and are never touched."""
    import sqlite3
    for rel in (os.path.join("Default", "Network", "Cookies"),
                os.path.join("Default", "Cookies"),
                os.path.join("Network", "Cookies"),
                "Cookies"):
        db = os.path.join(profile_dir, rel)
        if not os.path.isfile(db):
            continue
        con = None
        try:
            con = sqlite3.connect(db, timeout=2)
            con.execute("DELETE FROM cookies WHERE host_key NOT LIKE '%google%'")
            con.commit()
            con.execute("VACUUM")
        except Exception:
            pass
        finally:
            try:
                if con is not None:
                    con.close()
            except Exception:
                pass


def _profile_pool_init():
    """Ensure pool has PROFILE_POOL_SIZE profiles. Wipe stale ones, create missing."""
    import shutil
    for i in range(1, PROFILE_POOL_SIZE + 1):
        p = os.path.join(PROFILES_DIR, f"profile_{i}")
        if os.path.isdir(p):
            try:
                mtime = os.path.getmtime(p)
                age_h = (time.time() - mtime) / 3600
                if age_h > PROFILE_MAX_AGE_H:
                    shutil.rmtree(p, ignore_errors=True)
                    os.makedirs(p, exist_ok=True)
                else:
                    _clear_profile_cache(p)
                    _prune_cookies_keep_google(p)
            except Exception:
                pass
        else:
            os.makedirs(p, exist_ok=True)


def pick_profile():
    """Pick a random profile from the pool."""
    import random as _r
    profiles = [os.path.join(PROFILES_DIR, f"profile_{i}")
                for i in range(1, PROFILE_POOL_SIZE + 1)]
    chosen = _r.choice(profiles)
    os.makedirs(chosen, exist_ok=True)
    return chosen


def mode_profile(profile_key):
    """Return a DEDICATED browser profile for one tool (ranking/index/count/backlink).

    Each tool gets its own profile directory so the four tools can run in parallel —
    Chrome locks a --user-data-dir, so two tools sharing one profile would fail to
    launch the second browser. A dedicated profile per tool also keeps each tool's
    Google cookies/session separate. Falls back to a pooled profile for unknown keys."""
    if not profile_key or profile_key not in BATCH_MODES:
        return pick_profile()
    p = os.path.join(PROFILES_DIR, f"profile_{profile_key}")
    try:
        if os.path.isdir(p):
            age_h = (time.time() - os.path.getmtime(p)) / 3600
            if age_h > PROFILE_MAX_AGE_H:
                shutil.rmtree(p, ignore_errors=True)
                os.makedirs(p, exist_ok=True)
            else:
                _clear_profile_cache(p)
                _prune_cookies_keep_google(p)
        else:
            os.makedirs(p, exist_ok=True)
    except Exception:
        os.makedirs(p, exist_ok=True)
    return p


_profile_pool_init()
# Migrate old single profile if it exists
_old_profile = os.path.join(DATA_DIR, "chrome_profile")
if os.path.isdir(_old_profile):
    import shutil
    _dest = os.path.join(PROFILES_DIR, "profile_1")
    if not os.listdir(_dest):
        try:
            shutil.rmtree(_dest, ignore_errors=True)
            shutil.move(_old_profile, _dest)
        except Exception:
            pass
    else:
        shutil.rmtree(_old_profile, ignore_errors=True)

# For backward compat - PROFILE_DIR points to a random profile each launch
PROFILE_DIR = pick_profile()
print(f"[GRC v{APP_VERSION}] Profile pool: {PROFILES_DIR} ({PROFILE_POOL_SIZE} profiles)")
print(f"[GRC v{APP_VERSION}] Selected: {os.path.basename(PROFILE_DIR)}")

app = Flask(__name__, template_folder=TEMPLATE_DIR, static_folder=STATIC_DIR)
app.secret_key = "grc-v3-secret"
app.config["TEMPLATES_AUTO_RELOAD"] = True
app.jinja_env.auto_reload = True

# --------------------------------------------------------------------------- #
# Config
# --------------------------------------------------------------------------- #
DEFAULT_CONFIG = {
    "proxies": [],
    "use_buster": True,
    "manual_fallback": True,
    "max_block_retries": 3,
    "min_keyword_delay": 2,
    "default_country": "us",
    "default_pages": 5,
}

def load_config():
    cfg = dict(DEFAULT_CONFIG)
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                cfg.update(json.load(f) or {})
    except Exception as e:
        print(f"[config] load failed, using defaults: {e}")
    return cfg

def save_config(cfg):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2)
        return True
    except Exception as e:
        print(f"[config] save failed: {e}")
        return False

CONFIG = load_config()

# Load custom cities into engine dicts so they survive restarts
for _disp, _canon in CONFIG.get("custom_cities", {}).items():
    engine.CITY_CANONICAL[_disp] = _canon
    # Coordinates are stored separately if present
    _cc = CONFIG.get("custom_city_coords", {}).get(_disp)
    if _cc and isinstance(_cc, (list, tuple)) and len(_cc) == 2:
        try:
            engine.CITY_COORDS[_disp] = (float(_cc[0]), float(_cc[1]))
        except (ValueError, TypeError):
            pass

# Allow config to override downloads folder
if CONFIG.get("downloads_folder") and os.path.isdir(CONFIG["downloads_folder"]):
    DOWNLOADS_DIR = CONFIG["downloads_folder"]
SCREENSHOTS_DIR = DOWNLOADS_DIR


def _domain_folder(domain, mode="ranking"):
    """Create and return a domain-specific subfolder inside DOWNLOADS_DIR."""
    slug = domain.lower().replace("https://", "").replace("http://", "").replace("www.", "").strip("/").replace("/", "_")
    if not slug:
        slug = "unknown"
    suffix = {"ranking": "ranking", "index": "index_check", "backlink": "backlink_check"}.get(mode, mode)
    folder_name = f"{slug} {suffix}"
    folder = os.path.join(DOWNLOADS_DIR, folder_name)
    os.makedirs(folder, exist_ok=True)
    return folder


def _backup_report_path(category, domain, ext=".zip"):
    """Where a durable backup copy of a report lives (outside the user-facing
    Downloads folder) - see REPORT_BACKUPS_DIR."""
    import re as _re
    slug = _re.sub(r"[^a-z0-9._-]+", "_", (domain or "unknown").lower())
    folder = os.path.join(REPORT_BACKUPS_DIR, category)
    os.makedirs(folder, exist_ok=True)
    return os.path.join(folder, f"{slug}{ext}")


def _backup_report(category, domain, src_path):
    """Copy a finished report into REPORT_BACKUPS_DIR so 'Download' keeps working
    even if the user deletes/moves the folder it was originally saved into. Best
    effort - a failed backup never breaks the actual report generation."""
    try:
        dest = _backup_report_path(category, domain, os.path.splitext(src_path)[1] or ".zip")
        shutil.copy2(src_path, dest)
        return dest
    except Exception:
        return None


# --------------------------------------------------------------------------- #
# Domain / URL input normalization
# --------------------------------------------------------------------------- #
_NORM_UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")

def to_domain(value):
    """Reduce a URL (or already-bare domain) to just its host, keeping any subdomain
    exactly as entered. Strips scheme, path, query and fragment so a user pasting a full
    URL into a domain field (ranking, backlink, GSC, health, brief) becomes the plain
    domain/subdomain the tool expects. e.g. https://blog.example.com/a/b?x=1 -> blog.example.com"""
    import re as _r
    v = str(value or "").strip()
    v = _r.sub(r'^\s*[a-zA-Z][a-zA-Z0-9+.\-]*://', '', v)   # drop scheme
    v = v.split('/')[0].split('?')[0].split('#')[0]         # host only
    v = v.split('@')[-1]                                    # drop any user:pass@
    v = v.strip().strip('.').lower()
    return v

def resolve_working_url(value, log=None):
    """Given a domain OR a URL, follow the full redirect chain to the version that
    actually serves content (HTTP 200), and return (final_url, host). Tries the entered
    value first, then https/http x www/non-www variants. Used for URL-based tools
    (on-page) so the report is generated for the version that really works - e.g. a site
    that redirects http/non-www to https://www.x/ resolves to https://www.x. Returns
    (None, host) if nothing responds."""
    import re as _r
    import urllib.request as _u, urllib.error as _ue
    from urllib.parse import urlparse as _uparse
    raw = str(value or "").strip()
    host = to_domain(raw)
    if not host:
        return None, ""
    bare = host[4:] if host.startswith("www.") else host
    candidates = []
    if _r.match(r'^[a-zA-Z][a-zA-Z0-9+.\-]*://', raw):
        candidates.append(raw)                       # honour exactly what the user typed first
    for h in (host, f"www.{bare}", bare):
        for scheme in ("https", "http"):
            candidates.append(f"{scheme}://{h}/")
    seen = set()
    for url in candidates:
        if url in seen:
            continue
        seen.add(url)
        try:
            req = _u.Request(url, headers={"User-Agent": _NORM_UA})
            with _u.urlopen(req, timeout=12) as r:
                final = (r.geturl() or url).rstrip("/")
                if r.status == 200:
                    fh = _uparse(final).netloc
                    if log:
                        log(f"Working version detected: {final}")
                    return final, fh
        except Exception:
            continue
    if log:
        log(f"Could not reach any version of {host} - proceeding with {host} as entered.")
    return None, host

# --------------------------------------------------------------------------- #
# Per-mode run state (parallel jobs)
# --------------------------------------------------------------------------- #
# The four browser tools (ranking / index / count / backlink) each drive their own
# Chrome and used to share ONE global state + ONE stop/pause event + ONE browser, so
# only one could run at a time and starting a second clobbered the first. Each now has
# its own independent Job (state, events, lock, browser profile). The module-level
# `state`, `stop_event`, `pause_event`, `state_lock` names below are thread-aware
# proxies that transparently resolve to the CURRENT thread's job - so every existing
# `state[...]` / `stop_event.set()` call keeps working unchanged while operating on the
# right job. A worker thread pins its mode at the top of run_*(); Flask request handlers
# pin the mode from the request. See _ctx / _set_mode / _current_job below.
BATCH_MODES = ("ranking", "index", "count", "backlink")

def _fresh_state(mode):
    return {"status": "idle", "current_keyword": "", "current_index": 0, "total": 0,
            "results": [], "captcha_msg": "", "error_msg": "", "log": [],
            "driver": None, "mode": mode, "domain": ""}

class _Job:
    def __init__(self, mode):
        self.mode = mode
        self.state = _fresh_state(mode)
        self.pause = threading.Event(); self.pause.set()
        self.stop  = threading.Event()
        self.lock  = threading.RLock()   # re-entrant: safe even if a helper re-acquires

JOBS = {m: _Job(m) for m in BATCH_MODES}

_ctx = threading.local()

def _set_mode(mode):
    """Pin the mode for the current thread (worker thread or Flask request thread)."""
    _ctx.mode = mode if mode in JOBS else "ranking"

def _current_job():
    return JOBS.get(getattr(_ctx, "mode", None), JOBS["ranking"])

class _StateProxy:
    """Dict-like view of the current thread's job state (supports dict(state))."""
    def __getitem__(self, k):        return _current_job().state[k]
    def __setitem__(self, k, v):     _current_job().state[k] = v
    def __delitem__(self, k):        del _current_job().state[k]
    def __contains__(self, k):       return k in _current_job().state
    def __iter__(self):              return iter(_current_job().state)
    def __len__(self):               return len(_current_job().state)
    def get(self, k, d=None):        return _current_job().state.get(k, d)
    def update(self, *a, **kw):      _current_job().state.update(*a, **kw)
    def setdefault(self, k, d=None): return _current_job().state.setdefault(k, d)
    def pop(self, k, *a):            return _current_job().state.pop(k, *a)
    def keys(self):                  return _current_job().state.keys()
    def values(self):                return _current_job().state.values()
    def items(self):                 return _current_job().state.items()

class _EventProxy:
    def __init__(self, attr):        self._attr = attr
    def _e(self):                    return getattr(_current_job(), self._attr)
    def set(self):                   self._e().set()
    def clear(self):                 self._e().clear()
    def is_set(self):                return self._e().is_set()
    def wait(self, timeout=None):    return self._e().wait(timeout)

class _LockProxy:
    # __enter__ and __exit__ both resolve _current_job() in the SAME thread, so they
    # always acquire/release the same RLock. Never store the lock on the shared proxy —
    # that would race across threads.
    def __enter__(self):             _current_job().lock.acquire(); return _current_job().lock
    def __exit__(self, *a):          _current_job().lock.release(); return False
    def acquire(self, *a, **kw):     return _current_job().lock.acquire(*a, **kw)
    def release(self):               return _current_job().lock.release()

state       = _StateProxy()
pause_event = _EventProxy("pause")
stop_event  = _EventProxy("stop")
state_lock  = _LockProxy()

def add_log(msg, to_activity=False):
    ts = datetime.now().strftime("%H:%M:%S")
    line = f"[{ts}] {msg}"
    with state_lock:
        state["log"].append(line)
        if len(state["log"]) > 240:
            state["log"] = state["log"][-140:]
    print(line)
    if to_activity:
        activity(msg)

# --------------------------------------------------------------------------- #
# Auto-save
# --------------------------------------------------------------------------- #
def _autosave_file(mode=None):
    # Per-mode autosave so parallel tools don't overwrite each other's results.
    mode = mode or _current_job().mode
    return os.path.join(DATA_DIR, f"autosave_{mode}.json")

def autosave():
    try:
        with state_lock:
            mode = state.get("mode", "ranking")
            data = {"mode": mode,
                    "results": list(state["results"]),
                    "saved_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
        with open(_autosave_file(mode), "w", encoding="utf-8") as f:
            json.dump(data, f)   # compact (no indent) - this file is only ever read back by
                                  # load_autosave(), never by a human, and it's rewritten on
                                  # every item in a run so skipping pretty-printing keeps that
                                  # per-item write cheap without changing what's saved
    except Exception:
        pass

def load_autosave():
    try:
        f = _autosave_file()
        if os.path.exists(f):
            with open(f, "r", encoding="utf-8") as fh:
                return json.load(fh)
        # Fall back to the legacy single-file autosave (pre-parallel builds)
        if os.path.exists(AUTOSAVE_FILE):
            with open(AUTOSAVE_FILE, "r", encoding="utf-8") as fh:
                return json.load(fh)
    except Exception:
        pass
    return None

def clear_autosave():
    for f in (_autosave_file(), AUTOSAVE_FILE):
        try:
            if os.path.exists(f):
                os.remove(f)
        except Exception:
            pass

# --------------------------------------------------------------------------- #
# Buster captcha solver
# --------------------------------------------------------------------------- #
def _recaptcha_status(driver):
    from selenium.webdriver.common.by import By
    if not is_alive(driver):
        return "DEAD"
    try:
        for fr in driver.find_elements(By.CSS_SELECTOR, "iframe[src*='recaptcha']"):
            driver.switch_to.frame(fr)
            try:
                a = driver.find_element(By.ID, "recaptcha-anchor")
                cls = a.get_attribute("class") or ""
                driver.switch_to.default_content()
                return "CHECKED" if "recaptcha-checkbox-checked" in cls else "UNCHECKED"
            except Exception:
                driver.switch_to.default_content()
        return "NONE"
    except Exception:
        try:
            driver.switch_to.default_content()
        except Exception:
            pass
        return "DEAD" if not is_alive(driver) else "ERROR"

def _click_checkbox(driver):
    from selenium.webdriver.common.by import By
    try:
        for fr in driver.find_elements(By.CSS_SELECTOR, "iframe[src*='recaptcha']"):
            driver.switch_to.frame(fr)
            try:
                a = driver.find_element(By.ID, "recaptcha-anchor")
                driver.execute_script("arguments[0].click();", a)
                driver.switch_to.default_content()
                return True
            except Exception:
                driver.switch_to.default_content()
        return False
    except Exception:
        try:
            driver.switch_to.default_content()
        except Exception:
            pass
        return False

def _click_in_bframe(driver, selectors):
    from selenium.webdriver.common.by import By
    try:
        driver.switch_to.default_content()
        # Find the bframe (challenge iframe, not the anchor/checkbox iframe)
        fr = driver.find_elements(By.CSS_SELECTOR,
            "iframe[title*='recaptcha challenge']")
        if not fr:
            fr = driver.find_elements(By.CSS_SELECTOR,
                "iframe[src*='bframe']")
        if not fr:
            fr = [f for f in driver.find_elements(By.CSS_SELECTOR, "iframe[src*='recaptcha']")
                  if 'anchor' not in (f.get_attribute('src') or '')
                  and 'bframe' in (f.get_attribute('src') or '')]
        if not fr:
            add_log(f"bframe not found. iframes on page: "
                    f"{len(driver.find_elements(By.TAG_NAME, 'iframe'))}")
            return False
        add_log(f"Switching to bframe: {fr[0].get_attribute('src')[:60]}...")
        driver.switch_to.frame(fr[0])
        ok = False
        for sel in selectors:
            els = driver.find_elements(By.CSS_SELECTOR, sel)
            if els:
                add_log(f"Found {len(els)} elements for '{sel}'")
            for el in els:
                try:
                    el.click(); ok = True; break
                except Exception:
                    try:
                        driver.execute_script("arguments[0].click();", el); ok = True; break
                    except Exception:
                        continue
            if ok:
                break
        if not ok:
            add_log("No Buster button found inside bframe")
        driver.switch_to.default_content()
        return ok
    except Exception:
        try:
            driver.switch_to.default_content()
        except Exception:
            pass
        return False

def _click_buster_button(driver):
    """Find and click Buster's solver button - it may be outside the iframe."""
    from selenium.webdriver.common.by import By
    try:
        driver.switch_to.default_content()
    except Exception:
        pass
    # Buster adds its button outside iframes, on the main page
    buster_selectors = [
        "button[title*='Buster']", "button[title*='buster']",
        "button[title*='solver']", "button[title*='Solver']",
        "#solver-button", ".solver-button",
        "[id*='solver']", "[class*='solver-button']",
    ]
    for sel in buster_selectors:
        try:
            els = driver.find_elements(By.CSS_SELECTOR, sel)
            for el in els:
                try:
                    el.click()
                    return True
                except Exception:
                    try:
                        driver.execute_script("arguments[0].click();", el)
                        return True
                    except Exception:
                        continue
        except Exception:
            continue
    # Also try inside the bframe
    return _click_in_bframe(driver, ["#solver-button", ".help-button-holder button",
                                     "[id*='solver']", "[class*='solver']"])

def _click_audio_and_buster(driver):
    """Recovery when Buster/audio keeps failing: FIRST toggle to the IMAGE challenge
    (the image / 'eye' button) and reload it a couple of times - this clears the audio
    'automated queries' rate-limit - THEN switch to a fresh AUDIO challenge and click
    Buster (which solves audio). Image-reload → audio → Buster is what gets it unstuck."""
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.action_chains import ActionChains
    try:
        driver.switch_to.default_content()
        fr = driver.find_elements(By.CSS_SELECTOR,
            "iframe[title*='recaptcha challenge'], iframe[src*='bframe']")
        if not fr:
            return False
        driver.switch_to.frame(fr[0])

        def _click(selectors, label=None):
            for sel in selectors:
                els = driver.find_elements(By.CSS_SELECTOR, sel)
                if not els:
                    continue
                try:
                    els[0].click()
                except Exception:
                    try:
                        ActionChains(driver).move_to_element(els[0]).click().perform()
                    except Exception:
                        continue
                if label:
                    add_log(label)
                return True
            return False

        # 1) Switch to the IMAGE challenge (the image / "eye" button) and reload it 2x.
        #    Reloading the image challenge resets the audio rate-limit for the next step.
        _click(["#recaptcha-image-button", "button[id*='image']", "button[title*='image']"],
               "Switched to the image challenge (eye)")
        time.sleep(1.5)
        for _ in range(2):
            if _click(["#recaptcha-reload-button", ".reload-button-holder button",
                       "button[id*='reload']"], "Reloaded the challenge"):
                time.sleep(1.5)

        # 2) Switch to a FRESH audio challenge (the headphones button).
        _click(["#recaptcha-audio-button", "button[id*='audio']", "button[title*='audio']"],
               "Switched to a fresh audio challenge")
        time.sleep(2.5)

        # 3) Click Buster's solve button on the audio challenge.
        clicked = False
        host = driver.find_elements(By.CSS_SELECTOR, "div.help-button-holder")
        if host:
            try:
                ActionChains(driver).move_to_element(host[0]).click().perform()
                add_log("Buster clicked on the audio challenge")
                clicked = True
            except Exception:
                pass
        driver.switch_to.default_content()
        return clicked
    except Exception:
        try: driver.switch_to.default_content()
        except Exception: pass
        return False


def solve_with_buster(driver, max_attempts=1):
    from selenium.webdriver.common.by import By
    add_log("Trying Buster CAPTCHA solver...")
    found = False
    for _ in range(12):
        if not is_alive(driver):
            return False
        if driver.find_elements(By.CSS_SELECTOR, "iframe[src*='recaptcha']"):
            found = True
            break
        time.sleep(1)
    if not found:
        add_log("No reCAPTCHA iframe (text-only block).")
        return False

    for attempt in range(1, max_attempts + 1):
        if stop_event.is_set() or not is_alive(driver):
            return False
        add_log(f"Buster attempt {attempt}/{max_attempts}")
        if _recaptcha_status(driver) == "CHECKED":
            return True
        _click_checkbox(driver)
        time.sleep(3)
        if _recaptcha_status(driver) == "CHECKED":
            add_log("Solved by checkbox.")
            return True

        # Try the image-challenge-reset -> fresh-audio -> Buster sequence FIRST
        # (switching to the "eye"/image challenge briefly, then a fresh audio
        # challenge) - Buster solves audio far more reliably than image, and a
        # fresh audio challenge avoids the "automated queries" rate-limit that
        # makes repeated solves on the SAME challenge fail. This used to only run
        # as a late-stage fallback after 3 reload cycles; trying it immediately
        # solves most CAPTCHAs faster and with fewer DOM interactions (each one
        # is a chance for a stale-element error if the page changes mid-click).
        if _click_audio_and_buster(driver):
            for _ in range(10):
                time.sleep(2)
                if not is_alive(driver):
                    return False
                if _recaptcha_status(driver) == "CHECKED":
                    add_log("CAPTCHA solved by Buster on the audio challenge!")
                    return True
                try:
                    if classify_page(page_source(driver)) == "ok":
                        add_log("CAPTCHA solved manually!")
                        return True
                except Exception:
                    pass

        # Wait for Buster to inject its button in the bframe's help-button-holder
        # Buster uses a CLOSED shadow DOM - use CDP to pierce it
        clicked = False
        for _w in range(6):
            time.sleep(1.5)
            try:
                driver.switch_to.default_content()
                # Use CDP to find and click the solver-button inside closed shadow DOM
                try:
                    doc = driver.execute_cdp_cmd("DOM.getDocument", {"depth": -1, "pierce": True})
                    result = driver.execute_cdp_cmd("DOM.querySelectorAll", {
                        "nodeId": doc["root"]["nodeId"],
                        "selector": "#solver-button"
                    })
                    if result.get("nodeIds"):
                        for nid in result["nodeIds"]:
                            try:
                                remote = driver.execute_cdp_cmd("DOM.resolveNode", {"nodeId": nid})
                                obj_id = remote["object"]["objectId"]
                                driver.execute_cdp_cmd("Runtime.callFunctionOn", {
                                    "objectId": obj_id,
                                    "functionDeclaration": "function() { this.click(); }",
                                    "silent": True
                                })
                                add_log("Buster #solver-button clicked via CDP")
                                clicked = True
                                break
                            except Exception as e2:
                                add_log(f"CDP click attempt: {e2}")
                        if clicked:
                            break
                except Exception as e:
                    add_log(f"CDP method: {e}")

                # Fallback: switch to bframe and click help-button-holder
                fr = driver.find_elements(By.CSS_SELECTOR,
                    "iframe[title*='recaptcha challenge'], iframe[src*='bframe']")
                if not fr:
                    continue
                driver.switch_to.frame(fr[0])
                host = driver.find_elements(By.CSS_SELECTOR, "div.help-button-holder")
                if host:
                    add_log(f"help-button-holder found, size={host[0].size}")
                    from selenium.webdriver.common.action_chains import ActionChains
                    try:
                        ActionChains(driver).move_to_element(host[0]).click().perform()
                        add_log("Buster clicked via ActionChains on host")
                        clicked = True
                    except Exception:
                        pass
                driver.switch_to.default_content()
            except Exception:
                try: driver.switch_to.default_content()
                except Exception: pass
            if clicked:
                break
        if clicked:
            add_log("Buster button clicked, waiting up to 20s... (solve manually if you can)")
        else:
            add_log("Buster button not found - waiting for manual solve...")
        for _ in range(10):
            time.sleep(2)
            if not is_alive(driver):
                return False
            if _recaptcha_status(driver) == "CHECKED":
                add_log("CAPTCHA solved by Buster!")
                return True
            try:
                kind = classify_page(page_source(driver))
                if kind == "ok":
                    add_log("CAPTCHA solved manually!")
                    return True
            except Exception:
                pass

        # Buster didn't solve - try reloading CAPTCHA and retrying (up to 3 times)
        for reload_try in range(3):
            if stop_event.is_set() or not is_alive(driver):
                return False
            add_log(f"Reloading CAPTCHA (retry {reload_try + 1}/3)...")
            # Click reload button inside bframe
            reloaded = False
            try:
                driver.switch_to.default_content()
                fr = driver.find_elements(By.CSS_SELECTOR,
                    "iframe[title*='recaptcha challenge'], iframe[src*='bframe']")
                if fr:
                    driver.switch_to.frame(fr[0])
                    reload_btn = driver.find_elements(By.CSS_SELECTOR,
                        "#recaptcha-reload-button, .reload-button-holder button, "
                        "button[id*='reload']")
                    if reload_btn:
                        reload_btn[0].click()
                        reloaded = True
                        add_log("Clicked CAPTCHA reload button")
                    driver.switch_to.default_content()
            except Exception:
                try: driver.switch_to.default_content()
                except Exception: pass
            if not reloaded:
                # Reload button not found - wait and check for manual solve before giving up
                for _ in range(10):
                    time.sleep(2)
                    if not is_alive(driver):
                        return False
                    try:
                        if classify_page(page_source(driver)) == "ok":
                            add_log("CAPTCHA solved manually!")
                            return True
                    except Exception:
                        pass
                break
            time.sleep(3)
            if _recaptcha_status(driver) == "CHECKED":
                add_log("CAPTCHA solved after reload!")
                return True
            # On the reloaded challenge, force a FRESH AUDIO challenge (the "eye"/
            # headphones toggle) and re-click Buster. Buster solves AUDIO challenges best,
            # and switching to a fresh audio challenge clears the "automated queries"
            # audio rate-limit that makes repeated solves fail.
            clicked2 = _click_audio_and_buster(driver)
            if clicked2:
                for _ in range(10):
                    time.sleep(2)
                    if not is_alive(driver):
                        return False
                    if _recaptcha_status(driver) == "CHECKED":
                        add_log("CAPTCHA solved by Buster after reload!")
                        return True
                    try:
                        if classify_page(page_source(driver)) == "ok":
                            add_log("CAPTCHA solved manually!")
                            return True
                    except Exception:
                        pass

    # Final fallback: force a FRESH audio challenge (the "eye"/headphones button)
    # then let Buster try once more - often solves when reloads alone didn't.
    if is_alive(driver) and not stop_event.is_set() and _click_audio_and_buster(driver):
        for _ in range(12):
            time.sleep(2)
            if not is_alive(driver):
                return False
            if _recaptcha_status(driver) == "CHECKED":
                add_log("CAPTCHA solved by Buster on the audio challenge!")
                return True
            try:
                if classify_page(page_source(driver)) == "ok":
                    add_log("CAPTCHA solved!")
                    return True
            except Exception:
                pass

    add_log("Buster could not solve.")
    return False

# --------------------------------------------------------------------------- #
# Session holder (driver + identity rotation)
# --------------------------------------------------------------------------- #
class Session:
    def __init__(self, headless, country, proxy_pool, browser_pref="auto", vpn_method="none",
                 latitude=None, longitude=None, lang="en", profile_key=None):
        self.headless = headless
        self.country = country
        self.pool = proxy_pool
        self.browser_pref = browser_pref
        self.vpn_method = vpn_method
        self.latitude = latitude
        self.longitude = longitude
        self.lang = lang
        # Dedicated per-tool browser profile so the four tools can run in parallel.
        self.profile_key = profile_key
        self.driver = None

    def _extensions(self):
        exts = []
        if self.headless:
            return exts
        if CONFIG.get("use_buster", True) and os.path.isdir(BUSTER_DIR):
            exts.append(BUSTER_DIR)
            add_log("Buster CAPTCHA solver loaded")
        if self.vpn_method in ("windscribe",) and os.path.isdir(WINDSCRIBE_DIR):
            exts.append(WINDSCRIBE_DIR)
            add_log("Windscribe VPN loaded -- sign in once to activate")
        # "system" and "pause" rely on the user's OS-level VPN (Urban VPN desktop app,
        # ProtonVPN, Windscribe app, etc.) - no extension needed. Loading the Urban VPN
        # browser extension causes a registration error because it gets a different ID
        # when loaded as unpacked vs installed from the Chrome Web Store.
        if os.path.isdir(SERPCOUNTER_DIR):
            exts.append(SERPCOUNTER_DIR)
            add_log("SERP Counter loaded - results will show position numbers")
        return exts

    def start(self, rotate=False):
        self.quit()
        self.profile = mode_profile(self.profile_key) if self.profile_key else pick_profile()
        add_log(f"Using profile: {os.path.basename(self.profile)}")
        proxy = self.pool.next() if rotate else self.pool.current()
        if proxy is None and self.pool:
            proxy = self.pool.next()
        if proxy:
            add_log(f"Using proxy {proxy.get('host')}:{proxy.get('port')}")
        self.driver = engine.build_driver(
            self.profile, proxy=proxy, headless=self.headless,
            country=self.country, extra_extensions=self._extensions(),
            logger=add_log, browser_pref=self.browser_pref,
            latitude=self.latitude, longitude=self.longitude, lang=self.lang)
        with state_lock:
            state["driver"] = self.driver
        # Always fetch FRESH rankings: a rank check must reflect the live SERP, never a
        # page served from disk cache or personalised by a prior session's cookies.
        # Disable the HTTP cache for the whole session and wipe any saved cache + cookies
        # before warm_up (which re-does Google's consent). Keeps results current.
        try:
            self.driver.execute_cdp_cmd("Network.setCacheDisabled", {"cacheDisabled": True})
            self.driver.execute_cdp_cmd("Network.clearBrowserCache", {})
            self.driver.execute_cdp_cmd("Network.clearBrowserCookies", {})
            add_log("Cleared cache & cookies - fetching fresh results")
        except Exception:
            pass
        warm_up(self.driver, self.country, add_log, lang=self.lang)
        return self.driver

    def quit(self):
        try:
            if self.driver and is_alive(self.driver):
                self.driver.quit()
        except Exception:
            pass
        self.driver = None
        with state_lock:
            state["driver"] = None

# --------------------------------------------------------------------------- #
# Block recovery
# --------------------------------------------------------------------------- #
def _recover(sess, kind):
    add_log(f"Block detected ({kind}). Starting recovery...")

    if kind == "captcha" and not sess.headless:
        # Try Buster up to 3 times before falling back to manual
        if CONFIG.get("use_buster", True):
            if solve_with_buster(sess.driver, 3):
                return True

        # Buster failed - pause and auto-detect when solved (no Resume button needed)
        bring_browser_to_front()
        return _captcha_manual_wait(sess.driver)

    # When using a VPN (system/pause), the current VPN IP is likely hard-blocked by Google.
    # Cooling down on the same IP won't help - ask the user to switch VPN server instead.
    if sess.vpn_method in ("system", "pause") and not sess.headless:
        bring_browser_to_front()
        return _manual_pause(
            "Google has blocked this VPN IP. Switch to a different VPN server/location "
            "in your VPN app, then click Resume"
        )

    # For non-captcha blocks (soft_block, http_403), try cooldown + restart
    retries = CONFIG.get("max_block_retries", 3)
    for i in range(1, retries + 1):
        if stop_event.is_set():
            return False
        backoff = min(120, 15 * (2 ** (i - 1))) + random.uniform(0, 8)
        add_log(f"Recovery {i}/{retries}: cooldown {backoff:.0f}s"
                + (" + rotating proxy/identity" if sess.pool else " + fresh identity"))
        slept = 0
        while slept < backoff and not stop_event.is_set():
            time.sleep(1); slept += 1
        if stop_event.is_set():
            return False
        try:
            sess.start(rotate=bool(sess.pool))
        except BrowserClosedError:
            raise
        except Exception as e:
            add_log(f"Restart failed: {e}")
            continue
        try:
            kind2 = classify_page(page_source(sess.driver))
            if kind2 in ("ok", "empty", "consent"):
                add_log("Recovery looks clear.")
                return True
        except BrowserClosedError:
            raise
        except Exception:
            pass

    # All retries failed - last resort manual pause
    if CONFIG.get("manual_fallback", True) and not sess.headless:
        bring_browser_to_front()
        return _manual_pause("Automatic recovery failed")
    return False


def _reanchor_locale(sess, country, lang):
    """After a CAPTCHA/block is cleared, explicitly re-navigate to the Google
    homepage with the ORIGINAL hl=lang&gl=country - solving a CAPTCHA challenge
    (or a soft-block interstitial) can leave Google's session cookies drifted
    toward whatever locale that challenge page itself was served in (which
    follows the exit IP's GeoIP, not our hl/gl params) - so a keyword restart
    after recovery shouldn't just resume typing into whatever page state we're
    now on; it should re-establish the exact same starting conditions (language,
    country) as the very first attempt for that keyword."""
    try:
        if is_alive(sess.driver):
            dom = google_domain(country)
            safe_get(sess.driver, f"https://www.{dom}/?gl={country}&hl={lang}")
            human_pause(1.5, 2.5)
    except BrowserClosedError:
        raise
    except Exception:
        pass


def _recover_page(sess, page_num):
    """A pagination page came back with 0 organic links.

    That is almost always a block (CAPTCHA / soft-block) or a transient empty
    render - NOT the true end of results. Previously the loop trusted the empty
    page, found no 'Next' button on it, and reported the keyword as "not found",
    producing false negatives whenever the site actually ranked deeper.

    Try to clear the page *in place* (no browser restart, so we keep our spot in
    the pagination) and re-extract. Returns (links, reason):
        (links, "ok")      recovered - links found
        ([],    "empty")   page is genuinely empty (real end of results)
        ([],    "blocked") a block persisted and could not be cleared (incomplete)
    """
    reason = "empty"
    retries = CONFIG.get("max_block_retries", 3)
    for attempt in range(1, retries + 1):
        if stop_event.is_set():
            return [], reason
        try:
            kind = classify_page(page_source(sess.driver))
        except BrowserClosedError:
            raise
        except Exception:
            kind = "empty"

        if kind == "captcha":
            reason = "blocked"
            add_log(f"CAPTCHA on page {page_num} - recovering (attempt {attempt}/{retries})...")
            cur = None
            try:
                cur = sess.driver.current_url
            except Exception:
                pass
            if not _recover(sess, kind):
                add_log(f"Could not clear CAPTCHA on page {page_num} - stopping (results incomplete)")
                return [], "blocked"
            # Cool down before reloading - reloading immediately after a solve usually just
            # re-serves the challenge (Google is rate-limiting this IP). A short backoff
            # gives it a moment; the durable fix for repeated page-2 blocks is a proxy/VPN.
            backoff = min(45.0, 8.0 * attempt) + random.uniform(0, 4)
            slept = 0.0
            while slept < backoff and not stop_event.is_set():
                time.sleep(1); slept += 1
            # Reload the page we were on so its results render after the solve.
            try:
                if cur and is_alive(sess.driver):
                    safe_get(sess.driver, cur)
                    human_pause(1.5, 2.5)
            except BrowserClosedError:
                raise
            except Exception:
                pass
        elif kind in ("soft_block", "http_403"):
            reason = "blocked"
            add_log(f"Block ({kind}) on page {page_num} - cooling down and reloading "
                    f"(attempt {attempt}/{retries})...")
            # Reload the same page after a backoff. Deliberately avoid a full browser
            # restart here so we don't lose our place in the pagination.
            backoff = min(90.0, 10.0 * attempt) + random.uniform(0, 5)
            slept = 0.0
            while slept < backoff and not stop_event.is_set():
                time.sleep(1); slept += 1
            try:
                if is_alive(sess.driver):
                    sess.driver.refresh()
                    human_pause(1.5, 2.5)
            except BrowserClosedError:
                raise
            except Exception:
                pass
        elif kind == "consent":
            add_log(f"Consent wall on page {page_num} - accepting and reloading...")
            try:
                engine.accept_consent(sess.driver, add_log)
                engine.accept_google_consent(sess.driver, add_log)
                human_pause(1.0, 2.0)
            except BrowserClosedError:
                raise
            except Exception:
                pass
        else:
            # Not a detected block - likely a slow/partial render. Reload once.
            add_log(f"Page {page_num}: 0 links ({kind}) - reloading (attempt {attempt}/{retries})...")
            human_pause(2.5, 4.5)
            try:
                if is_alive(sess.driver):
                    sess.driver.refresh()
                    human_pause(1.5, 2.5)
            except BrowserClosedError:
                raise
            except Exception:
                pass

        try:
            links = extract_organic(sess.driver)
        except BrowserClosedError:
            raise
        except Exception:
            links = []
        if links:
            return links, "ok"

    return [], reason


def _captcha_manual_wait(driver):
    """Show manual CAPTCHA prompt and auto-resume once the CAPTCHA clears —
    no Resume button click needed. Also resumes if user clicks Resume manually."""
    msg = ("CAPTCHA detected - solve it in the browser window. "
           "The tool will continue automatically once solved.")
    add_log(msg)
    with state_lock:
        state["captcha_msg"] = msg
        state["status"] = "paused"
    pause_event.clear()

    # Poll every 2 seconds - resume as soon as CAPTCHA disappears from the page
    while not stop_event.is_set() and not pause_event.is_set():
        time.sleep(2)
        try:
            src = page_source(driver)
            kind = classify_page(src)
            if kind in ("ok", "empty", "consent"):
                add_log("CAPTCHA solved - resuming automatically.")
                break
        except Exception:
            pass

    with state_lock:
        state["captcha_msg"] = ""
        if not stop_event.is_set():
            state["status"] = "running"
    return not stop_event.is_set()


def _manual_pause(reason):
    add_log(f"{reason}. Pausing for manual solve - solve in the browser, then Resume.")
    bring_browser_to_front()
    with state_lock:
        state["captcha_msg"] = (f"{reason}. The browser window has been brought to the "
                                f"front - solve the CAPTCHA there, then click Resume.")
        state["status"] = "paused"
    pause_event.clear()
    pause_event.wait()
    with state_lock:
        state["captcha_msg"] = ""
        if not stop_event.is_set():
            state["status"] = "running"
    return not stop_event.is_set()

def _browser_closed_pause(sess):
    """Browser window was closed (by the user or a crash) mid-run. Instead of ending
    the whole run with an error, pause the same way a manual CAPTCHA solve does -
    already-checked items stay in the results (nothing is lost) - and on Resume,
    reopen a fresh browser so the caller can retry whatever item was in progress
    from the start. Returns False if the user stopped the run instead of resuming."""
    add_log("Browser window was closed. Click Resume to reopen it and continue - "
            "already-checked results are kept; the current item will be re-checked "
            "from the start.")
    with state_lock:
        state["captcha_msg"] = ("The browser window was closed. Click Resume to reopen it "
                                "and continue. Already-checked results are safe.")
        state["status"] = "paused"
    pause_event.clear()
    pause_event.wait()
    if stop_event.is_set():
        return False
    try:
        sess.start(rotate=bool(sess.pool))
        add_log("Browser reopened - resuming.")
    except Exception as e:
        add_log(f"Could not reopen browser: {e}")
        with state_lock:
            state["captcha_msg"] = ""
        return False
    with state_lock:
        state["captcha_msg"] = ""
        state["status"] = "running"
    return True


def _highlight_domain_in_serp(driver, domain_clean, first_only=False):
    """Highlight the target domain's organic result(s) in the LIVE SERP before the
    screenshot - like the SERP Highlighter extension: an orange outline, a soft
    background and a 'YOUR SITE' badge - so the client's position is obvious in the
    captured image.

    Matching is on the result's URL (the title link's href), NOT the title text.
    first_only=True highlights just the first match (for 'stop when domain found');
    otherwise every occurrence on the page is highlighted ('Scan All Pages').
    Returns how many results were highlighted."""
    js = r"""
    var dom = arguments[0], firstOnly = arguments[1];
    var count = 0;
    var anchors = document.querySelectorAll('a');
    for (var i = 0; i < anchors.length; i++) {
        if (firstOnly && count >= 1) break;
        var a = anchors[i];
        if (!a.querySelector || !a.querySelector('h3')) continue;   // organic title links
        // Match the result's HOST (or subdomain), NOT the title text or the domain
        // merely appearing in another site's URL path (e.g. trustpilot.com/review/
        // exactprint.co.uk must not highlight for exactprint.co.uk).
        var host = (a.hostname || '').toLowerCase().replace(/^www\./, '');
        if (host !== dom && !(host.length > dom.length && host.slice(-(dom.length + 1)) === '.' + dom)) continue;
        var box = a.closest('[data-hveid]') || a.closest('.g') || a.closest('.MjjYud') || a.parentElement;
        if (!box || box.getAttribute('data-stp-hl')) continue;
        box.setAttribute('data-stp-hl', '1');
        box.style.outline = '4px solid #ff5a00';
        box.style.outlineOffset = '3px';
        box.style.background = 'rgba(255,206,130,0.32)';
        box.style.borderRadius = '10px';
        box.style.position = 'relative';
        var badge = document.createElement('span');
        badge.textContent = '★ YOUR SITE';
        badge.style.cssText = 'display:inline-block;background:#ff5a00;color:#fff;' +
            'font:bold 12px Arial,sans-serif;padding:2px 9px;border-radius:6px;margin:0 0 6px 0;';
        box.insertBefore(badge, box.firstChild);
        count++;
    }
    return count;
    """
    try:
        n = driver.execute_script(js, domain_clean, bool(first_only))
        return int(n) if n else 0
    except Exception:
        return 0


def _save_full_page_screenshot(driver, path):
    """Capture a full-page screenshot via CDP (entire scrollable page, not just viewport)."""
    import base64 as b64
    try:
        metrics = driver.execute_cdp_cmd("Page.getLayoutMetrics", {})
        content = metrics.get("contentSize", metrics.get("cssContentSize", {}))
        width = int(content.get("width", 1280))
        height = int(content.get("height", 3000))
        height = min(height, 15000)
        driver.execute_cdp_cmd("Emulation.setDeviceMetricsOverride", {
            "mobile": False, "width": width, "height": height, "deviceScaleFactor": 1
        })
        result = driver.execute_cdp_cmd("Page.captureScreenshot", {
            "format": "png", "captureBeyondViewport": True
        })
        driver.execute_cdp_cmd("Emulation.clearDeviceMetricsOverride", {})
        with open(path, "wb") as f:
            f.write(b64.b64decode(result["data"]))
    except Exception:
        driver.save_screenshot(path)


def _upload_ranking_screenshot(path):
    """Upload a ranking SERP screenshot somewhere public so the exported CSV/sheet
    can carry a real shareable URL instead of the /api/screenshot/<file> link,
    which only resolves while this app's local server is running.

    ImgBB (real, purpose-built image host) is tried FIRST when a key is configured
    (Admin -> Sync API Keys) - it's the only option confirmed reliable, and checking
    it first avoids burning 30+ seconds per screenshot on dead hosts before reaching
    it. Pixeldrain is a no-key fallback for when no key is set. Dropped entirely:
    catbox.moe/0x0.st (unreachable/uploads disabled), transfer.sh (connection
    actively refused - service appears dead), storage.to (Cloudflare bot-challenge
    page instead of an upload response - can never work via a plain HTTP request,
    on any network) - all confirmed on a real machine, not just this sandbox."""
    imgbb_key = CONFIG.get("imgbb_api_key", "").strip()
    if imgbb_key:
        try:
            with open(path, "rb") as f:
                r = http_requests.post("https://api.imgbb.com/1/upload",
                                        params={"key": imgbb_key},
                                        files={"image": f}, timeout=25)
            data = r.json()
            url = (data.get("data") or {}).get("url", "")
            if r.status_code == 200 and data.get("success") and url:
                return url
            add_log(f"Screenshot upload (ImgBB) failed: {data.get('error', {}).get('message', r.text[:120])}")
        except Exception as e:
            add_log(f"Screenshot upload (ImgBB) failed: {type(e).__name__}: {e}")
        return ""

    # No ImgBB key configured - Pixeldrain as a no-key fallback (shorter timeout
    # since it's a fallback of last resort, not worth a long wait if it's blocked
    # on this network too).
    try:
        with open(path, "rb") as f:
            r = http_requests.post("https://pixeldrain.com/api/file", files={"file": f}, timeout=10)
        data = r.json()
        if r.status_code in (200, 201) and data.get("success") and data.get("id"):
            return "https://pixeldrain.com/u/" + data["id"]
        add_log(f"Screenshot upload (Pixeldrain) failed: {data.get('message', r.text[:120])}")
    except Exception as e:
        add_log(f"Screenshot upload (Pixeldrain) failed: {type(e).__name__}: {e}")
    return ""


# --------------------------------------------------------------------------- #
# Keyword ranking - hardened
# --------------------------------------------------------------------------- #
def rank_one(sess, keyword, domain, country, max_pages, search_mode="stop_on_found", city=None, lang="en"):
    domain_clean = clean_domain(domain)
    target = max_pages * 10

    for _try in range(CONFIG.get("max_block_retries", 3) + 1):
        if stop_event.is_set():
            return {"status": "stopped", "matches": []}
        pause_event.wait()
        if not is_alive(sess.driver):
            add_log("Browser lost - restarting for retry...")
            try:
                sess.start(rotate=bool(sess.pool))
            except Exception:
                raise BrowserClosedError("Browser closed and could not restart")

        human_search(sess.driver, keyword, country, add_log, city=city, lang=lang)

        src = page_source(sess.driver)
        kind = classify_page(src)

        if kind == "consent":
            add_log("Consent page detected - accepting...")
            engine.accept_consent(sess.driver, add_log)
            engine.accept_google_consent(sess.driver, add_log)
            human_search(sess.driver, keyword, country, add_log, city=city, lang=lang)
            src = page_source(sess.driver)
            kind = classify_page(src)

        if kind in ("captcha", "soft_block", "http_403"):
            add_log(f"Block detected ({kind}). Starting recovery...")
            if not _recover(sess, kind):
                return {"status": "captcha", "matches": match_domain([], domain_clean)}
            _reanchor_locale(sess, country, lang)
            continue

        # Small wait then log what's on page 1
        time.sleep(1.0)
        links_page1, dbg = extract_organic(sess.driver, debug=True)

        # If no links found, might be consent or empty - try accepting consent and retry
        if not links_page1 and _try < CONFIG.get("max_block_retries", 3):
            engine.accept_consent(sess.driver, add_log)
            engine.accept_google_consent(sess.driver, add_log)
            time.sleep(1)
            links_page1, dbg = extract_organic(sess.driver, debug=True)
        if links_page1:
            add_log(f"Page 1: {len(links_page1)} organic links:")
            for idx, l in enumerate(links_page1[:20], 1):
                try:
                    domain_part = l.split('//')[1].split('/')[0].replace('www.', '')
                except Exception:
                    domain_part = l[:50]
                add_log(f"  #{idx}: {domain_part}")
        else:
            add_log(f"Page 1: 0 links - URL={dbg.get('url','?')[:80]} "
                    f"h3={dbg.get('h3count','?')} jsname={dbg.get('jsname_count','?')} "
                    f"zReHs={dbg.get('zReHs_count','?')} rso={dbg.get('rso','?')}")

        # Ctrl+F search on page 1
        all_matches = find_domain_in_page(sess.driver, domain_clean, page_offset=0)

        # Capture a HIGHLIGHTED screenshot of every page where the domain is found.
        # stop_on_found -> highlight just the first match; Scan All Pages -> every
        # occurrence. Captured per-page (before pagination navigates away) so a
        # Scan-All run shows the domain wherever it ranks, not just the last page.
        _first_only = (search_mode == "stop_on_found")

        def _shot_serp_page(page_matches):
            if not page_matches or not is_alive(sess.driver):
                return
            try:
                hl = _highlight_domain_in_serp(sess.driver, domain_clean, first_only=_first_only)
                if hl:
                    add_log(f"Highlighted {hl} result(s) for {domain_clean}")
                    time.sleep(0.3)
                safe_kw = "".join(c if c.isalnum() or c in (' ', '-', '_') else '_' for c in keyword)[:50]
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                ss_name = f"{safe_kw}_{ts}.png"
                ss_path = os.path.join(_domain_folder(domain, "ranking"), ss_name)
                _save_full_page_screenshot(sess.driver, ss_path)
                add_log(f"SERP screenshot saved: {ss_name}")
                ss_url = _upload_ranking_screenshot(ss_path)
                if ss_url:
                    add_log(f"Screenshot URL: {ss_url}")
                for m in page_matches:
                    m["screenshot"] = ss_name
                    if ss_url:
                        m["screenshot_url"] = ss_url
            except Exception as e:
                add_log(f"Screenshot failed: {e}")

        if all_matches:
            _shot_serp_page(all_matches)

        # Paginate through ALL selected pages (respects max_pages regardless of search_mode)
        page_num = 1
        total_links = len(links_page1)
        blocked_incomplete = False
        from selenium.webdriver.common.by import By

        # stop_on_found: stop as soon as domain is found on any page
        # check_all / multiple: keep going through all pages to find every occurrence
        keep_going = (not all_matches) if search_mode == "stop_on_found" else True

        if max_pages > 1 and keep_going:
            add_log(f"Checking up to {max_pages} pages...")
            _NEXT_SELECTORS = (
                "#pnnext, a[id='pnnext'], "
                "a[aria-label='Next page'], a[aria-label='Pagina successiva'], "
                "a[aria-label='Siguiente'], a[aria-label='Próxima página'], "
                "a[aria-label='Nächste Seite'], a[aria-label='Page suivante'], "
                "td.navend a"
            )
            while page_num < max_pages:
                try:
                    nxt = sess.driver.find_elements(By.CSS_SELECTOR, _NEXT_SELECTORS)
                    if not nxt:
                        add_log(f"No 'Next' button on page {page_num} - stopping")
                        break
                    # Human-like paging: scroll to the Next button and pause before AND after
                    # the click. Rapid back-to-back page clicks from one IP are what trip
                    # Google's CAPTCHA on page 2+; slowing this down reduces (does not
                    # eliminate) blocks. The real fix for deep pagination is a proxy/VPN.
                    try:
                        sess.driver.execute_script(
                            "arguments[0].scrollIntoView({block:'center'});", nxt[0])
                    except Exception:
                        pass
                    human_pause(1.2, 2.4)
                    nxt[0].click()
                    human_pause(2.8, 5.5)
                    page_num += 1
                    page_links = extract_organic(sess.driver)
                    if not page_links:
                        # 0 links is almost always a block or a transient empty render,
                        # NOT the end of results - try to clear it before trusting it.
                        page_links, reason = _recover_page(sess, page_num)
                        if not page_links:
                            if reason == "blocked":
                                add_log(f"Page {page_num}: blocked and could not recover - "
                                        f"stopping (results may be incomplete)")
                                blocked_incomplete = True
                            else:
                                add_log(f"Page {page_num}: 0 links - reached end of results")
                            break
                    prev_total = total_links   # organic results counted on all prior pages
                    total_links += len(page_links)
                    add_log(f"Page {page_num}: {len(page_links)} links")
                    # Rank = results on prior pages + position on THIS page. Use the ACTUAL
                    # cumulative count, not page_num*10 - Google frequently shows fewer than
                    # 10 organic results per page (ads/snippets/local pack take slots; page 1
                    # here had 9), and assuming 10/page inflates the reported position.
                    page_matches = find_domain_in_page(sess.driver, domain_clean, page_offset=prev_total)
                    if page_matches:
                        all_matches.extend(page_matches)
                        add_log(f"  Found at position {page_matches[0]['position']}")
                        _shot_serp_page(page_matches)   # highlight + screenshot this page
                    # stop_on_found: exit pagination as soon as domain appears
                    if search_mode == "stop_on_found" and all_matches:
                        break
                except BrowserClosedError:
                    raise
                except Exception as e:
                    add_log(f"Pagination error: {e}")
                    break

        matches = all_matches
        if matches:
            if search_mode == "stop_on_found":
                matches = matches[:1]
            # Screenshots were already captured per page (with the domain
            # highlighted) in _shot_serp_page, so no end-of-run capture is needed.
            # serp_url is already the exact URL shown in the SERP result (captured
            # straight from the page source in find_domain_in_page) - no extra page
            # load needed, and none is done, so the driver stays on the SERP page.
            add_log(f"'{keyword}': found at #{matches[0]['position']} "
                    f"({total_links} results across {page_num} pages)")
            return {"status": "found", "matches": matches, "pages": page_num}
        if blocked_incomplete:
            max_retries = CONFIG.get("max_block_retries", 3)
            if _try < max_retries:
                add_log(f"'{keyword}': search cut short by a block at page {page_num} - "
                        f"restarting keyword from page 1 (retry {_try + 1}/{max_retries})")
                _reanchor_locale(sess, country, lang)
                continue
            add_log(f"'{keyword}': search cut short by a block at page {page_num} - "
                    f"ranking may exist deeper ({total_links} results seen)")
            return {"status": f"not_found (incomplete - blocked at page {page_num})",
                    "matches": [], "pages": page_num, "incomplete": True}
        add_log(f"'{keyword}': not in top {page_num} pages ({total_links} results)")
        return {"status": f"not_found in {page_num} pages", "matches": [], "pages": page_num}

    return {"status": f"not_found in {max_pages} pages", "matches": []}

def _vpn_pause_at_start(vpn_method, driver=None):
    # Check system IP (no browser needed)
    add_log("Checking your IP location...")
    check_ip_location(logger=add_log)

    msgs = {
        "windscribe": "Windscribe is loaded. Click its icon in the browser toolbar, "
                      "sign in (free), connect to your target country, then click Resume.",
        "edge_secure": "Turn on Edge Secure Network: click the shield icon in Edge's "
                       "address bar and enable it, then click Resume.",
        "system": "Connect your VPN app (Urban VPN / ProtonVPN etc.) to the target "
                  "country now, then click Resume.",
        "pause": "Connect your VPN now (Windscribe/Urban VPN/any), then click Resume.",
    }
    msg = msgs.get(vpn_method)
    if not msg:
        return True
    add_log("Paused for VPN setup.")
    with state_lock:
        state["captcha_msg"] = msg
        state["status"] = "paused"
    pause_event.clear()
    pause_event.wait()
    with state_lock:
        state["captcha_msg"] = ""
        if not stop_event.is_set():
            state["status"] = "running"
    if not stop_event.is_set():
        add_log("Checking IP after VPN connect (via browser)...")
        loc = check_ip_location(driver=driver, logger=add_log, use_browser=True)
        if loc:
            with state_lock:
                state["vpn_location"] = loc.get("location", "")
    return not stop_event.is_set()


def _vpn_disconnect_pause(vpn_method, driver=None):
    """After results are done, pause to let user disconnect VPN before browser closes."""
    if vpn_method in ("none", "proxy"):
        return
    loc = check_ip_location(logger=add_log)
    loc_str = loc.get("location", "VPN") if loc else "VPN"
    if driver and is_alive(driver):
        add_log(f"Analysis complete. VPN still connected ({loc_str}). Pausing to disconnect...")
        bring_browser_to_front()
        with state_lock:
            state["captcha_msg"] = (
                f"Analysis complete! Your VPN is still connected ({loc_str}).\n\n"
                f"Please disconnect your VPN now, then click Resume to close the browser.")
            state["status"] = "paused"
        pause_event.clear()
        pause_event.wait()
        with state_lock:
            state["captcha_msg"] = ""
    else:
        add_log("Reminder: Disconnect your VPN if it is still running.")

def run_rank_analysis(keywords, domain, country, delay, max_pages, headless, proxies,
                      browser_pref="auto", search_mode="stop_on_found", vpn_method="none",
                      latitude=None, longitude=None, city=None, lang="en", target_pages=None):
    _set_mode("ranking")
    sess = Session(headless, country, ProxyPool(proxies), browser_pref, vpn_method,
                   latitude, longitude, lang=lang, profile_key="ranking")
    try:
        with state_lock:
            state["status"] = "starting"
            state["domain"] = domain
        add_log("Launching hardened browser...")
        sess.start(rotate=bool(sess.pool))
        if not _vpn_pause_at_start(vpn_method, sess.driver):
            return
        with state_lock:
            state["status"] = "running"

        last_ip_check = time.time()
        initial_ip = None
        with state_lock:
            initial_ip = state.get("vpn_location", "")

        # Keywords whose search was cut short by a CAPTCHA/block - retried at the end.
        incomplete = []   # list of (kw, kw_domain, kw_target, result_row_index)

        def _make_row(kw, kw_target, result):
            matches = result.get("matches", [])
            if matches:
                row = {"keyword": kw, "status": "found"}
                if kw_target:
                    row["target_page"] = kw_target
                for idx, m in enumerate(matches):
                    suffix = "" if idx == 0 else f"_{idx+1}"
                    row[f"position{suffix}"] = m["position"]
                    row[f"serp_url{suffix}"] = m["serp_url"]
                if matches[0].get("screenshot"):
                    row["screenshot"] = matches[0]["screenshot"]
                if matches[0].get("screenshot_url"):
                    row["screenshot_url"] = matches[0]["screenshot_url"]
                return row
            row = {"keyword": kw, "position": "-", "serp_url": "",
                   "status": result.get("status", "not_found")}
            if kw_target:
                row["target_page"] = kw_target
            return row

        for i, kw in enumerate(keywords):
            if stop_event.is_set():
                break
            pause_event.wait()
            if stop_event.is_set():
                break
            if not is_alive(sess.driver):
                if not _browser_closed_pause(sess):
                    raise BrowserClosedError("Browser closed and the run was stopped")

            # Check VPN location every 10 keywords to detect drift
            if vpn_method not in ("none", "proxy") and i > 0 and i % 10 == 0:
                if time.time() - last_ip_check > 120:
                    loc = check_ip_location(logger=add_log)
                    last_ip_check = time.time()
                    if loc:
                        cur_loc = loc.get("location", "")
                        with state_lock:
                            state["vpn_location"] = cur_loc
                        if initial_ip and cur_loc != initial_ip:
                            add_log(f"[warn] VPN location changed: {initial_ip} -> {cur_loc}")

            with state_lock:
                state["current_keyword"] = kw
                state["current_index"] = i + 1
                state["status"] = "running"
            add_log(f"Checking '{kw}' ({i+1}/{len(keywords)})")

            # Auto-recover if browser died between keywords
            if not is_alive(sess.driver):
                if not _browser_closed_pause(sess):
                    raise BrowserClosedError("Browser closed and the run was stopped")

            if not target_pages:
                target_pages = {}
            kw_target = target_pages.get(kw, "")
            kw_domain = domain
            if kw_target:
                from urllib.parse import urlparse as _up2
                _parsed = _up2(kw_target if "://" in kw_target else "https://" + kw_target)
                kw_domain = _parsed.netloc.replace("www.", "") or domain

            # If the browser dies mid-check (not just between keywords), pause and
            # retry the SAME keyword from scratch once it's reopened, rather than
            # losing the whole run - everything checked before this point is safe
            # in state["results"] already.
            while True:
                try:
                    result = rank_one(sess, kw, kw_domain, country, max_pages, search_mode, city=city, lang=lang)
                    break
                except BrowserClosedError:
                    if not _browser_closed_pause(sess):
                        raise BrowserClosedError("Browser closed and the run was stopped")
                    add_log(f"Retrying '{kw}' from the start after the browser reopened...")

            with state_lock:
                row = _make_row(kw, kw_target, result)
                state["results"].append(row)
                if result.get("incomplete"):
                    incomplete.append((kw, kw_domain, kw_target, len(state["results"]) - 1))
            autosave()

            if i < len(keywords) - 1 and not stop_event.is_set():
                pause_event.wait()
                if not stop_event.is_set():
                    t0 = time.time()
                    # Quick neutral visit in a background tab (opens, loads, closes ~2s)
                    human_visit_neutral_bg(sess.driver, domain, add_log)
                    elapsed = time.time() - t0
                    wait = max(0, max(CONFIG.get("min_keyword_delay", 2), delay) + random.uniform(0.5, 1.5) - elapsed)
                    if wait > 0:
                        add_log(f"Waiting {wait:.0f}s before next keyword...")
                        t = 0
                        while t < wait and not stop_event.is_set():
                            time.sleep(1); t += 1

        # Retry pass: re-check keywords whose search was cut short by a CAPTCHA/block.
        # A fresh session (new profile, rotated proxy if a pool is set) resets the IP /
        # challenge state, giving the best chance to complete them - so a site that DOES
        # rank isn't left as a false "not found (incomplete)".
        if incomplete and not stop_event.is_set():
            add_log(f"Re-checking {len(incomplete)} keyword(s) that a block cut short...",
                    to_activity=True)
            for kw, kw_domain, kw_target, ridx in incomplete:
                if stop_event.is_set():
                    break
                pause_event.wait()
                if stop_event.is_set():
                    break
                try:
                    sess.start(rotate=True)   # fresh profile + rotate proxy for a clean IP
                except BrowserClosedError:
                    raise
                except Exception as re_err:
                    add_log(f"Could not restart browser for retry: {re_err}")
                    break
                cool = random.uniform(6, 12)
                t = 0
                while t < cool and not stop_event.is_set():
                    time.sleep(1); t += 1
                with state_lock:
                    state["current_keyword"] = f"{kw} (retry)"
                    state["status"] = "running"
                add_log(f"Re-checking '{kw}' (was cut short by a block)...")
                result = rank_one(sess, kw, kw_domain, country, max_pages, search_mode,
                                  city=city, lang=lang)
                with state_lock:
                    new_row = _make_row(kw, kw_target, result)
                    if 0 <= ridx < len(state["results"]):
                        state["results"][ridx] = new_row   # replace the incomplete row
                    else:
                        state["results"].append(new_row)
                if result.get("matches"):
                    add_log(f"'{kw}': retry found at #{result['matches'][0]['position']}",
                            to_activity=True)
                else:
                    add_log(f"'{kw}': retry -> {result.get('status', 'not_found')}")
                autosave()

        with state_lock:
            state["status"] = "stopped" if stop_event.is_set() else "completed"
        add_log("Rank analysis finished.", to_activity=True)
        autosave()
        _vpn_disconnect_pause(vpn_method, sess.driver)
    except BrowserClosedError as e:
        add_log(f"Browser closed: {e}")
        autosave()
        with state_lock:
            state["status"] = "error"
            state["error_msg"] = "Browser was closed. Results saved - Export CSV to download."
    except Exception as e:
        import traceback; traceback.print_exc()
        add_log(f"Fatal error: {e}")
        autosave()
        with state_lock:
            state["status"] = "error"; state["error_msg"] = str(e)
    finally:
        sess.quit()

# --------------------------------------------------------------------------- #
# Index checker - hardened
# --------------------------------------------------------------------------- #
def index_one(sess, raw_url, country, city=None, lang="en"):
    from urllib.parse import urlparse
    url = raw_url.strip()
    if not url.startswith("http"):
        url = "https://" + url
    p = urlparse(url)
    q = "site:" + p.netloc.replace("www.", "") + p.path.rstrip("/")

    for _try in range(CONFIG.get("max_block_retries", 3) + 1):
        if stop_event.is_set():
            return {"status": "stopped", "indexed": "Unknown", "found_url": ""}
        human_search(sess.driver, q, country, add_log, city=city, lang=lang)
        src = page_source(sess.driver)
        kind = classify_page(src)
        if kind in ("captcha", "soft_block", "http_403"):
            if not _recover(sess, kind):
                return {"status": "captcha", "indexed": "Unknown", "found_url": ""}
            _reanchor_locale(sess, country, lang)
            continue
        if "did not match any documents" in src.lower():
            return {"status": "ok", "indexed": "No", "found_url": ""}
        links = extract_organic(sess.driver)
        if links:
            return {"status": "ok", "indexed": "Yes", "found_url": links[0]}
        return {"status": "ok", "indexed": "No", "found_url": ""}
    return {"status": "ok", "indexed": "Unknown", "found_url": ""}

def run_index_analysis(urls, delay, headless, country, proxies,
                       browser_pref="auto", vpn_method="none",
                       latitude=None, longitude=None, city=None, lang="en"):
    _set_mode("index")
    sess = Session(headless, country, ProxyPool(proxies), browser_pref, vpn_method,
                   latitude, longitude, lang=lang, profile_key="index")
    try:
        with state_lock:
            state["status"] = "starting"
        add_log("Launching hardened browser...")
        sess.start(rotate=bool(sess.pool))
        if not _vpn_pause_at_start(vpn_method, sess.driver):
            return
        with state_lock:
            state["status"] = "running"

        for i, url in enumerate(urls):
            if stop_event.is_set():
                break
            pause_event.wait()
            if stop_event.is_set():
                break
            if not is_alive(sess.driver):
                if not _browser_closed_pause(sess):
                    raise BrowserClosedError("Browser closed and the run was stopped")
            with state_lock:
                state["current_keyword"] = url
                state["current_index"] = i + 1
                state["status"] = "running"
            add_log(f"Index check '{url}' ({i+1}/{len(urls)})")
            while True:
                try:
                    result = index_one(sess, url, country, city=city, lang=lang)
                    break
                except BrowserClosedError:
                    if not _browser_closed_pause(sess):
                        raise BrowserClosedError("Browser closed and the run was stopped")
                    add_log(f"Retrying '{url}' from the start after the browser reopened...")
            with state_lock:
                state["results"].append({
                    "url": url, "indexed": result.get("indexed", "Unknown"),
                    "found_url": result.get("found_url", ""),
                    "status": result.get("status", "unknown"),
                    "checked_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")})
            autosave()
            if i < len(urls) - 1 and not stop_event.is_set():
                pause_event.wait()
                if not stop_event.is_set():
                    t0 = time.time()
                    human_visit_neutral_bg(sess.driver, None, add_log)
                    elapsed = time.time() - t0
                    wait = max(0, max(CONFIG.get("min_keyword_delay", 2), delay) + random.uniform(1, 3) - elapsed)
                    if wait > 0:
                        add_log(f"Waiting {wait:.0f}s before next URL...")
                        t = 0
                        while t < wait and not stop_event.is_set():
                            time.sleep(1); t += 1

        with state_lock:
            state["status"] = "stopped" if stop_event.is_set() else "completed"
        add_log("Index check finished.", to_activity=True)
        autosave()
        _vpn_disconnect_pause(vpn_method, sess.driver)
    except BrowserClosedError as e:
        add_log(f"Browser closed: {e}")
        autosave()
        with state_lock:
            state["status"] = "error"
            state["error_msg"] = "Browser closed. Results saved - Export CSV to download."
    except Exception as e:
        import traceback; traceback.print_exc()
        add_log(f"Fatal error: {e}")
        autosave()
        with state_lock:
            state["status"] = "error"; state["error_msg"] = str(e)
    finally:
        sess.quit()

# --------------------------------------------------------------------------- #
# Search-results count - searches each keyword and reads Google's "About N results"
# --------------------------------------------------------------------------- #
def _capture_result_count(driver):
    """Read Google's 'About N results' total from the current SERP. Returns the
    number exactly as Google groups it (e.g. '16,500' or the Indian '20,20,00,000'),
    or '' if the page has no result-stats element."""
    import re as _rc
    def _extract(text):
        if not text:
            return ""
        # First number-like run (digits + grouping separators). Language-agnostic —
        # doesn't depend on the words "About"/"results", which are localised.
        m = _rc.search(r'\d[\d.,   ]*\d|\d', text)
        return m.group(0).strip("   ") if m else ""
    # 1) The count lives in #result-stats: "About 16,500 results (0.18 seconds)"
    try:
        from selenium.webdriver.common.by import By
        els = driver.find_elements(By.ID, "result-stats")
        if els:
            c = _extract(els[0].text or "")
            if c:
                return c
    except Exception:
        pass
    # 2) Fallback: pull #result-stats out of the page source and strip its tags
    try:
        src = page_source(driver) or ""
        m = _rc.search(r'id="result-stats"[^>]*>(.*?)</div>', src, _rc.S)
        if m:
            plain = _rc.sub(r'<[^>]+>', ' ', m.group(1))
            c = _extract(plain)
            if c:
                return c
        m = _rc.search(r'About\s+([\d.,   ]+?)\s+results', src)
        if m:
            return m.group(1).strip("   ")
    except Exception:
        pass
    return ""


def count_one(sess, keyword, country, city=None, lang="en"):
    """Search Google for the keyword and return its total 'About N results' count."""
    for _try in range(CONFIG.get("max_block_retries", 3) + 1):
        if stop_event.is_set():
            return {"status": "stopped", "results_count": ""}
        pause_event.wait()
        if not is_alive(sess.driver):
            raise BrowserClosedError("Browser closed before search")
        human_search(sess.driver, keyword, country, add_log, city=city, lang=lang)
        src = page_source(sess.driver)
        kind = classify_page(src)
        if kind == "consent":
            engine.accept_consent(sess.driver, add_log)
            engine.accept_google_consent(sess.driver, add_log)
            human_search(sess.driver, keyword, country, add_log, city=city, lang=lang)
            src = page_source(sess.driver)
            kind = classify_page(src)
        if kind in ("captcha", "soft_block", "http_403"):
            add_log(f"Block detected ({kind}). Starting recovery...")
            if not _recover(sess, kind):
                return {"status": "captcha", "results_count": ""}
            _reanchor_locale(sess, country, lang)
            continue
        time.sleep(0.8)
        count = _capture_result_count(sess.driver)
        if count:
            add_log(f"'{keyword}': about {count} results")
            return {"status": "ok", "results_count": count}
        low = src.lower()
        if "did not match any documents" in low or "no results found" in low:
            add_log(f"'{keyword}': 0 results")
            return {"status": "ok", "results_count": "0"}
        return {"status": "ok", "results_count": "N/A"}
    return {"status": "ok", "results_count": "N/A"}


def run_count_analysis(keywords, delay, headless, country, proxies,
                       browser_pref="auto", vpn_method="none",
                       latitude=None, longitude=None, city=None, lang="en"):
    _set_mode("count")
    sess = Session(headless, country, ProxyPool(proxies), browser_pref, vpn_method,
                   latitude, longitude, lang=lang, profile_key="count")
    try:
        with state_lock:
            state["status"] = "starting"
        add_log("Launching hardened browser...")
        sess.start(rotate=bool(sess.pool))
        if not _vpn_pause_at_start(vpn_method, sess.driver):
            return
        with state_lock:
            state["status"] = "running"

        for i, kw in enumerate(keywords):
            if stop_event.is_set():
                break
            pause_event.wait()
            if stop_event.is_set():
                break
            if not is_alive(sess.driver):
                if not _browser_closed_pause(sess):
                    raise BrowserClosedError("Browser closed and the run was stopped")
            with state_lock:
                state["current_keyword"] = kw
                state["current_index"] = i + 1
                state["status"] = "running"
            add_log(f"Search results count '{kw}' ({i+1}/{len(keywords)})")
            while True:
                try:
                    result = count_one(sess, kw, country, city=city, lang=lang)
                    break
                except BrowserClosedError:
                    if not _browser_closed_pause(sess):
                        raise BrowserClosedError("Browser closed and the run was stopped")
                    add_log(f"Retrying '{kw}' from the start after the browser reopened...")
            with state_lock:
                state["results"].append({
                    "keyword": kw,
                    "results_count": result.get("results_count", ""),
                    "status": result.get("status", "unknown"),
                    "checked_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")})
            autosave()
            if i < len(keywords) - 1 and not stop_event.is_set():
                pause_event.wait()
                if not stop_event.is_set():
                    t0 = time.time()
                    human_visit_neutral_bg(sess.driver, None, add_log)
                    elapsed = time.time() - t0
                    wait = max(0, max(CONFIG.get("min_keyword_delay", 2), delay) + random.uniform(1, 3) - elapsed)
                    if wait > 0:
                        add_log(f"Waiting {wait:.0f}s before next keyword...")
                        t = 0
                        while t < wait and not stop_event.is_set():
                            time.sleep(1); t += 1

        with state_lock:
            state["status"] = "stopped" if stop_event.is_set() else "completed"
        add_log("Search results count finished.", to_activity=True)
        autosave()
        _vpn_disconnect_pause(vpn_method, sess.driver)
    except BrowserClosedError as e:
        add_log(f"Browser closed: {e}")
        autosave()
        with state_lock:
            state["status"] = "error"
            state["error_msg"] = "Browser closed. Results saved - Export CSV to download."
    except Exception as e:
        import traceback; traceback.print_exc()
        add_log(f"Fatal error: {e}")
        autosave()
        with state_lock:
            state["status"] = "error"; state["error_msg"] = str(e)
    finally:
        sess.quit()

# --------------------------------------------------------------------------- #
# Backlink checker - visits each backlink URL, finds domain link, checks meta
# --------------------------------------------------------------------------- #
def _wait_page_ready(driver, timeout=10):
    """Block until the browser reports the document has finished loading, so anchors
    are actually present in the DOM before we scan for the backlink."""
    end = time.time() + timeout
    while time.time() < end:
        try:
            if driver.execute_script("return document.readyState") == "complete":
                return
        except Exception:
            return
        time.sleep(0.4)


def _page_blocked_or_empty(src):
    """True when the fetched page is a transient challenge / rate-limit / empty shell
    rather than the real content. On such pages the backlink would falsely read as
    'not found', so we retry once - this is what caused the '404 first, Ok on retry'
    inconsistency."""
    if not src or len(src) < 800:
        return True
    low = src.lower()
    markers = ("just a moment", "checking your browser", "enable javascript and cookies",
               "access denied", "attention required", "429 too many requests",
               "too many requests", "rate limited", "request blocked",
               "you have been blocked", "captcha-delivery", "cf-error-details")
    return any(m in low for m in markers)


def backlink_one(sess, backlink_url, target_domain, check_da=True):
    from urllib.parse import urlparse
    import re as _re
    url = backlink_url.strip()
    if not url.startswith("http"):
        url = "https://" + url
    target_clean = target_domain.lower().replace("www.", "").strip("/")

    # Load the page, waiting for it to finish and retrying ONCE if the first response
    # is a transient block / empty shell. Without this a slow or rate-limited first
    # load made the link read as missing ("404"), then present ("Ok") on a re-check.
    src = ""
    for attempt in range(2):
        try:
            safe_get(sess.driver, url)
        except Exception as e:
            if attempt == 0:
                human_pause(2, 4)
                continue
            return {"status": "error", "domain_found": "Error", "meta_robots": "N/A",
                    "link_type": "N/A", "link_url": "", "error": str(e)}
        if not is_alive(sess.driver):
            return {"status": "error", "domain_found": "Error", "meta_robots": "N/A",
                    "link_type": "N/A", "link_url": ""}
        _wait_page_ready(sess.driver, timeout=10)
        human_pause(2, 4)
        src = page_source(sess.driver)
        if not _page_blocked_or_empty(src):
            break
        if attempt == 0:
            human_pause(3, 6)  # settle, then one clean retry

    if not is_alive(sess.driver):
        return {"status": "error", "domain_found": "Error", "meta_robots": "N/A",
                "link_type": "N/A", "link_url": ""}

    # Capture final URL after any redirects
    try:
        final_url = sess.driver.current_url
    except Exception:
        final_url = url

    # Check meta robots
    meta_robots = "index"
    meta_match = _re.search(r'<meta\s[^>]*name=["\']robots["\'][^>]*content=["\']([^"\']+)["\']',
                            src, _re.IGNORECASE)
    if not meta_match:
        meta_match = _re.search(r'<meta\s[^>]*content=["\']([^"\']+)["\'][^>]*name=["\']robots["\']',
                                src, _re.IGNORECASE)
    if meta_match:
        content = meta_match.group(1).lower()
        if "noindex" in content:
            meta_robots = "noindex"
        else:
            meta_robots = "index"

    # Find target domain links. Profile/directory sites usually wrap the outbound
    # website link in a redirect/tracking URL (e.g. href="/outbound?url=https%3A%2F%2F
    # target.com" or a l.php?u=... wrapper), so the anchor's NETLOC is the host site, not
    # the target - matching only the netloc misses a link that's genuinely there. We match
    # the target anywhere in the DECODED href, and fall back to scanning the rendered HTML.
    from urllib.parse import unquote as _unquote
    domain_found = "No"
    link_url = ""
    link_type = "N/A"

    def _host_is_target(host):
        host = (host or "").lower().replace("www.", "").strip("/")
        return bool(host) and (host == target_clean
                               or host.endswith("." + target_clean)
                               or target_clean.endswith("." + host))

    try:
        from selenium.webdriver.common.by import By
        anchors = sess.driver.find_elements(By.TAG_NAME, "a")
        for a in anchors:
            try:
                href = (a.get_attribute("href") or "").strip()
                if not href or href.startswith(("javascript:", "#", "mailto:", "tel:")):
                    continue
                parsed = urlparse(href)
                decoded = _unquote(href).lower()
                if _host_is_target(parsed.netloc) or target_clean in decoded:
                    domain_found = "Yes"
                    link_url = href
                    rel = (a.get_attribute("rel") or "").lower()
                    link_type = "nofollow" if "nofollow" in rel else "dofollow"
                    break
            except Exception:
                continue
    except Exception:
        pass

    # Fallback: some sites inject the website link via JS or nest it in a way the anchor
    # scan misses. Look for the target inside any href in the rendered HTML.
    if domain_found == "No" and target_clean:
        try:
            m = _re.search(r'href=["\']([^"\']*' + _re.escape(target_clean) + r'[^"\']*)["\']',
                           src, _re.IGNORECASE)
            if m:
                domain_found = "Yes"
                link_url = m.group(1)
                seg = src[max(0, m.start() - 220): m.end() + 60]
                link_type = "nofollow" if _re.search(r'rel=["\'][^"\']*nofollow', seg, _re.IGNORECASE) else "dofollow"
        except Exception:
            pass

    da_val, dr_val, pa_val, da_src = "N/A", "N/A", "N/A", "N/A"
    if check_da:
        bl_domain = urlparse(url).netloc.replace("www.", "").lower()
        da_result = check_da_pa(sess.driver, bl_domain, log_fn=add_log)
        da_val = da_result.get("da", "N/A")
        dr_val = da_result.get("dr", "N/A")
        pa_val = da_result.get("pa", "N/A")
        da_src = da_result.get("source", "N/A")

    return {"status": "ok", "domain_found": domain_found, "meta_robots": meta_robots,
            "link_type": link_type, "link_url": link_url, "final_url": final_url,
            "da": da_val, "dr": dr_val, "pa": pa_val, "da_source": da_src}


def run_backlink_analysis(urls, domain, delay, headless, country, proxies,
                          browser_pref="auto", vpn_method="none", check_da=True,
                          latitude=None, longitude=None):
    _set_mode("backlink")
    sess = Session(headless, country, ProxyPool(proxies), browser_pref, vpn_method,
                   latitude, longitude, profile_key="backlink")
    try:
        with state_lock:
            state["status"] = "starting"
        add_log("Launching browser for backlink check...")
        sess.start(rotate=bool(sess.pool))
        if not _vpn_pause_at_start(vpn_method, sess.driver):
            return
        with state_lock:
            state["status"] = "running"

        for i, url in enumerate(urls):
            if stop_event.is_set():
                break
            pause_event.wait()
            if stop_event.is_set():
                break
            if not is_alive(sess.driver):
                if not _browser_closed_pause(sess):
                    raise BrowserClosedError("Browser closed and the run was stopped")
            with state_lock:
                state["current_keyword"] = url
                state["current_index"] = i + 1
                state["status"] = "running"
            add_log(f"Checking backlink '{url}' ({i+1}/{len(urls)})")
            while True:
                try:
                    result = backlink_one(sess, url, domain, check_da=check_da)
                    break
                except BrowserClosedError:
                    if not _browser_closed_pause(sess):
                        raise BrowserClosedError("Browser closed and the run was stopped")
                    add_log(f"Retrying '{url}' from the start after the browser reopened...")
            add_log(f"  domain {'found' if result['domain_found']=='Yes' else 'not found'}, "
                    f"meta: {result['meta_robots']}, link: {result['link_type']}")
            with state_lock:
                state["results"].append({
                    "url": url, "domain_found": result.get("domain_found", "No"),
                    "meta_robots": result.get("meta_robots", "N/A"),
                    "link_type": result.get("link_type", "N/A"),
                    "link_url": result.get("link_url", ""),
                    "da": result.get("da", "N/A"), "pa": result.get("pa", "N/A"),
                    "dr": result.get("dr", "N/A"),
                    "da_source": result.get("da_source", "N/A"),
                    "status": result.get("status", "unknown"),
                    "checked_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")})
            autosave()
            if i < len(urls) - 1 and not stop_event.is_set():
                pause_event.wait()
                if not stop_event.is_set():
                    wait = max(2, delay) + random.uniform(1, 3)
                    t = 0
                    while t < wait and not stop_event.is_set():
                        time.sleep(1); t += 1

        with state_lock:
            state["status"] = "stopped" if stop_event.is_set() else "completed"
        add_log("Backlink check finished.", to_activity=True)
        autosave()
        _vpn_disconnect_pause(vpn_method, sess.driver)
    except BrowserClosedError as e:
        add_log(f"Browser closed: {e}")
        autosave()
        with state_lock:
            state["status"] = "error"
            state["error_msg"] = "Browser closed. Results saved - Export CSV to download."
    except Exception as e:
        import traceback; traceback.print_exc()
        add_log(f"Fatal error: {e}")
        autosave()
        with state_lock:
            state["status"] = "error"; state["error_msg"] = str(e)
    finally:
        sess.quit()

# --------------------------------------------------------------------------- #
# Shared proxy pool (central sheet, admin-managed) - fetched server-side only,
# never returned to the browser. Used as an OCCASIONAL fallback so the office
# network's own IP doesn't take all the CAPTCHA heat - see _proxies_from_request.
# --------------------------------------------------------------------------- #
_shared_proxies_cache = {"data": [], "ts": 0}
_shared_proxies_lock = threading.Lock()
SHARED_PROXIES_TTL = 600
SHARED_PROXY_USE_CHANCE = 0.2         # baseline chance - no proxy in the pool matches the search's target country
SHARED_PROXY_USE_CHANCE_MATCHED = 0.7  # much higher chance when a pool proxy actually matches the target country -
                                        # this is the ideal case (an office-IP-hiding proxy that's ALSO geo-consistent
                                        # with the search), so use it far more often than the blind/random fallback.

# Common aliases so a human-entered "Region" value (e.g. "United Kingdom", "USA")
# still matches the 2-letter country codes the app uses internally (engine.GOOGLE_DOMAINS).
_REGION_COUNTRY_ALIASES = {
    "uk": "gb", "united kingdom": "gb", "great britain": "gb", "britain": "gb",
    "usa": "us", "united states": "us", "united states of america": "us", "america": "us",
    "australia": "au", "canada": "ca", "india": "in", "germany": "de", "france": "fr",
    "spain": "es", "italy": "it", "netherlands": "nl", "brazil": "br", "mexico": "mx",
    "japan": "jp", "korea": "kr", "south korea": "kr", "russia": "ru",
}


def _region_matches_country(region, country):
    """True if a proxy's free-text Region field (admin-entered) refers to the
    same country as the search's target `country` code. Tolerant of common
    naming variants (2-letter code, full name, "USA"/"UK") since Region isn't
    a validated/constrained field."""
    if not region or not country:
        return False
    r = region.strip().lower()
    c = country.strip().lower()
    if r == c:
        return True
    if _REGION_COUNTRY_ALIASES.get(r) == c:
        return True
    # Datacenter-proxy-style naming: "us-east", "us1", "US - New York"
    import re as _re_local
    if _re_local.match(rf"^{_re_local.escape(c)}([^a-z]|$)", r):
        return True
    return False


def _fetch_shared_proxies_now():
    email, password = auth.get_any_credentials()
    if not email:
        return
    try:
        result = auth._api_call({"action": "proxies_list", "email": email, "password": password})
        rows = result.get("proxies")
        if isinstance(rows, list):
            with _shared_proxies_lock:
                _shared_proxies_cache["data"] = [r for r in rows if r.get("active", True)
                                                  and r.get("host") and r.get("port")]
                _shared_proxies_cache["ts"] = time.time()
    except Exception:
        pass


def _shared_proxies():
    with _shared_proxies_lock:
        fresh = _shared_proxies_cache["data"] and (time.time() - _shared_proxies_cache["ts"] < SHARED_PROXIES_TTL)
        data = list(_shared_proxies_cache["data"])
    if not fresh:
        _fetch_shared_proxies_now()
        with _shared_proxies_lock:
            data = list(_shared_proxies_cache["data"])
    return data


# --------------------------------------------------------------------------- #
# Runtime API keys (central sheet, admin-managed; optional per-building
# overrides) - synced automatically in the background for every user, mirroring
# the shared-proxy pool above, instead of requiring a manual admin button press.
# --------------------------------------------------------------------------- #
RUNTIME_KEYS_TTL = 900


def _fetch_runtime_keys_now():
    """Merge central-sheet keys with this user's building override, if any, and
    write straight into CONFIG. Central keys come from the auth-gated
    runtime_keys action on the central gateway; the per-building override (if the
    building's own GSC-tracker sheet has a "Keys" tab) comes from that building's
    own GSC script - each building's admin manages their own override there
    without needing any central-sheet access."""
    global CONFIG
    email, password = auth.get_any_credentials()
    if not email:
        return
    keys = {}
    try:
        result = auth._api_call({"action": "runtime_keys", "email": email, "password": password})
        if result.get("success"):
            keys.update({k: v for k, v in result.get("keys", {}).items() if v})
    except Exception:
        pass
    try:
        webapp_url = _gsc_webapp_url()
        if webapp_url:
            resp = http_requests.post(webapp_url, json={"action": "get_keys"}, timeout=15)
            data = resp.json()
            if data.get("success"):
                keys.update({k: v for k, v in data.get("keys", {}).items() if v})
    except Exception:
        pass
    if keys:
        changed = any(CONFIG.get(k) != v for k, v in keys.items())
        for k, v in keys.items():
            CONFIG[k] = v.strip() if isinstance(v, str) else v
        if changed:
            save_config(CONFIG)


def _runtime_keys_sync_loop():
    """Waits for a logged-in user, fetches/merges runtime keys in the background,
    then refreshes every RUNTIME_KEYS_TTL seconds for the life of the app - so a
    key added centrally (or overridden per-building) reaches every machine
    automatically, without anyone needing to click 'Sync Keys'."""
    while True:
        email, _pw = auth.get_any_credentials()
        if email:
            _fetch_runtime_keys_now()
            time.sleep(RUNTIME_KEYS_TTL)
        else:
            time.sleep(15)


# --------------------------------------------------------------------------- #
# Helpers for request parsing
# --------------------------------------------------------------------------- #
def _least_recently_used(candidates):
    """Pick among the candidates that have gone longest without being used,
    with light randomization across the tied-oldest group instead of a single
    deterministic winner - the shared proxies cache only refreshes every
    SHARED_PROXIES_TTL seconds, so many independent local apps could otherwise
    all read the exact same 'last_used' snapshot and stampede onto the same
    'most idle' proxy at once. A proxy with no last_used value yet (never
    used, or the sheet predates the column) sorts first."""
    def _key(p):
        ts = (p.get("last_used") or "").strip()
        if not ts:
            return ""
        return ts
    ranked = sorted(candidates, key=_key)
    oldest_ts = _key(ranked[0])
    tied = [p for p in ranked if _key(p) == oldest_ts]
    return random.choice(tied)


def _mark_proxy_used_async(pick):
    """Fire-and-forget: tell the central sheet this proxy was just picked, so
    other users' independent local apps prefer a different one next time.
    Runs in a background thread with a short timeout - a slow/failed write
    here must never delay the actual run that's about to use this proxy."""
    def _do():
        try:
            email, password = auth.get_any_credentials()
            if not email:
                return
            auth._api_call({"action": "proxies_mark_used", "email": email, "password": password,
                            "host": pick.get("host", ""), "port": pick.get("port", "")}, timeout=8)
        except Exception:
            pass
    threading.Thread(target=_do, daemon=True).start()


def _proxies_from_request(data, country=None):
    proxies = list(CONFIG.get("proxies", []))
    ph = (data.get("proxy_host") or "").strip()
    pp = (data.get("proxy_port") or "").strip()
    if ph and pp:
        proxies.insert(0, {
            "type": data.get("proxy_type", "http"), "host": ph, "port": pp,
            "user": (data.get("proxy_user") or "").strip(),
            "pass": (data.get("proxy_pass") or "").strip()})
    if not proxies:
        shared = _shared_proxies()
        if shared:
            matched = [p for p in shared if _region_matches_country(p.get("region", ""), country)]
            pick = None
            if matched and random.random() < SHARED_PROXY_USE_CHANCE_MATCHED:
                pick = _least_recently_used(matched)
                add_log(f"Using a shared {(country or '').upper()}-region proxy for this run "
                        f"(geo-matched, least-recently-used - protects the office IP, keeps results "
                        f"consistent, and spreads load across the pool so no single static-IP proxy "
                        f"gets hit by everyone at once).")
            elif random.random() < SHARED_PROXY_USE_CHANCE:
                pick = _least_recently_used(shared)
                add_log("Using a shared proxy for this run (occasional rotation to protect the office IP).")
            if pick:
                proxies = [{"type": pick.get("type", "http"), "host": pick["host"], "port": pick["port"],
                            "user": pick.get("user", ""), "pass": pick.get("pass", "")}]
                _mark_proxy_used_async(pick)
    return proxies

# --------------------------------------------------------------------------- #
# Flask routes - Auth
# --------------------------------------------------------------------------- #
@app.route("/api/auth/check")
def api_auth_check():
    """Check if user is authenticated (saved token)."""
    result = auth.check_saved_auth()
    if result.get("status") == "approved":
        result["accounts"] = auth.list_logged_in()
        result["api_configured"] = bool(auth._get_api_url())
    elif not auth._get_api_url():
        result = {"status": "no_api_url", "message": "User Auth URL not configured. Ask your admin to add it in Admin Settings."}
    return jsonify(result)

@app.route("/api/auth/login", methods=["POST"])
def api_auth_login():
    data = request.get_json(force=True) or {}
    email = (data.get("email") or "").strip()
    password = (data.get("password") or "").strip()
    if not email or not password:
        return jsonify({"status": "error", "message": "Email and password required."}), 400
    result = auth.login(email, password)
    return jsonify(result)

@app.route("/api/auth/logout", methods=["POST"])
def api_auth_logout():
    data = request.get_json(silent=True) or {}
    email = (data.get("email") or "").strip() or None
    return jsonify(auth.logout(email))

@app.route("/api/auth/accounts")
def api_auth_accounts():
    """List all logged-in accounts on this device."""
    return jsonify({"accounts": auth.list_logged_in()})

@app.route("/api/auth/version")
def api_auth_version():
    return jsonify(auth.check_version())

@app.route("/api/auth/change-password", methods=["POST"])
def api_change_password():
    data = request.get_json(silent=True) or {}
    old_pw = (data.get("old_password") or "").strip()
    new_pw = (data.get("new_password") or "").strip()
    if not old_pw or not new_pw:
        return jsonify({"status": "error", "message": "Current and new password are required."}), 400
    if len(new_pw) < 6:
        return jsonify({"status": "error", "message": "New password must be at least 6 characters."}), 400
    logged = auth.list_logged_in()
    if not logged:
        return jsonify({"status": "error", "message": "Not logged in."}), 401
    return jsonify(auth.change_password(logged[0]["email"], old_pw, new_pw))

@app.route("/api/auth/update-name", methods=["POST"])
def api_update_name():
    data = request.get_json(silent=True) or {}
    name = (data.get("name") or "").strip()
    if not name:
        return jsonify({"status": "error", "message": "Name cannot be empty."}), 400
    logged = auth.list_logged_in()
    if not logged:
        return jsonify({"status": "error", "message": "Not logged in."}), 401
    return jsonify(auth.update_name(logged[0]["email"], name))

@app.route("/api/auth/mac")
def api_auth_mac():
    """Return MAC address for display to user (token for admin)."""
    return jsonify({"mac": auth.get_mac_address()})

@app.route("/api/auth/formats")
def api_auth_formats():
    """Return allowed report formats for the current user."""
    fmts = auth.get_allowed_formats()
    return jsonify({"formats": fmts})

@app.route("/api/auth/tools")
def api_auth_tools():
    """Return allowed tools for the current user."""
    return jsonify({"tools": auth.get_allowed_tools()})

# --------------------------------------------------------------------------- #
# Flask routes - API
# --------------------------------------------------------------------------- #
@app.route("/")
def index():
    return render_template("index.html", version=APP_VERSION)

def _parse_keywords(raw_text):
    """Parse keywords textarea. Supports plain list or tab/comma-separated with headers.
    If first line headers contain 'keyword' and 'page', parse as keyword+target pairs.
    Returns (keywords_list, target_pages_dict).
    """
    import re
    lines = [l for l in raw_text.splitlines() if l.strip()]
    if not lines:
        return [], {}

    header = lines[0].lower()
    has_header = "keyword" in header and "page" in header
    if has_header and len(lines) > 1:
        sep = "\t" if "\t" in lines[0] else ","
        cols = [c.strip().lower() for c in lines[0].split(sep)]
        kw_idx = next((i for i, c in enumerate(cols) if "keyword" in c), None)
        pg_idx = next((i for i, c in enumerate(cols) if "page" in c), None)
        if kw_idx is not None:
            keywords = []
            target_pages = {}
            seen = set()
            for line in lines[1:]:
                parts = line.split(sep)
                kw = parts[kw_idx].strip() if kw_idx < len(parts) else ""
                pg = parts[pg_idx].strip() if pg_idx is not None and pg_idx < len(parts) else ""
                if not kw or kw.lower() in seen:
                    continue
                seen.add(kw.lower())
                keywords.append(kw)
                if pg:
                    if not re.match(r'https?://', pg, re.I):
                        pg = "https://" + pg
                    target_pages[kw] = pg
            return keywords[:100], target_pages
    keywords = list(dict.fromkeys(k.strip() for k in lines if k.strip()))[:100]
    return keywords, {}

@app.route("/api/start", methods=["POST"])
def api_start():
    data = request.get_json(silent=True) or {}
    mode = data.get("mode", "ranking")
    # Pin this request to the requested tool's job so all state below is per-mode.
    _set_mode(mode)
    with state_lock:
        if state["status"] in ("running", "paused", "starting"):
            return jsonify({"error": f"{mode.title()} is already running. Stop or wait first."}), 400

    # Domain-based tools: if the user pasted a full URL into the domain field
    # (ranking domain, backlink "your domain to find"), trim it to just the host.
    domain = to_domain(data.get("domain") or "")
    keywords, target_pages = _parse_keywords(data.get("keywords") or "")
    if not domain and target_pages:
        from urllib.parse import urlparse as _up
        first_url = next(iter(target_pages.values()))
        domain = _up(first_url).netloc.replace("www.", "") or first_url
    country = (data.get("country") or CONFIG.get("default_country", "us")).strip().lower()
    delay = max(1, int(data.get("delay", 3)))
    max_pages = max(1, min(10, int(data.get("max_pages", CONFIG.get("default_pages", 5)))))
    headless = bool(data.get("headless", False))
    browser_pref = (data.get("browser", "auto") or "auto").strip().lower()
    search_mode = (data.get("search_mode", "stop_on_found") or "stop_on_found").strip()
    vpn_method = (data.get("vpn_method", "none") or "none").strip()
    city = (data.get("city") or "").strip() or None
    lang = (data.get("lang") or "en").strip().lower() or "en"
    latitude = None
    longitude = None
    if city and city in engine.CITY_COORDS:
        latitude, longitude = engine.CITY_COORDS[city]
    proxies = _proxies_from_request(data, country=country)

    if vpn_method == "proxy" and not proxies:
        add_log(f"Fetching a free proxy for {country.upper()}...")
        p = engine.fetch_free_proxy(country, add_log)
        if p:
            proxies = [p]

    if mode == "ranking" and (not domain or not keywords):
        return jsonify({"error": "Domain and at least one keyword required."}), 400
    if mode == "index" and not keywords:
        return jsonify({"error": "At least one URL required."}), 400
    if mode == "backlink" and (not domain or not keywords):
        return jsonify({"error": "Domain and at least one backlink URL required."}), 400
    if mode == "count" and not keywords:
        return jsonify({"error": "At least one keyword required."}), 400

    with state_lock:
        state.update({"status": "starting", "current_keyword": "", "current_index": 0,
                      "total": len(keywords), "results": [], "captcha_msg": "",
                      "error_msg": "", "mode": mode, "log": [], "domain": domain})
    pause_event.set(); stop_event.clear(); clear_autosave()

    if city and latitude is not None:
        add_log(f"City: {city} - Geolocation sensor: {latitude}, {longitude}")

    if mode == "ranking":
        t = threading.Thread(target=run_rank_analysis,
                             args=(keywords, domain, country, delay, max_pages, headless,
                                   proxies, browser_pref, search_mode, vpn_method,
                                   latitude, longitude, city, lang, target_pages),
                             daemon=True)
    elif mode == "backlink":
        check_da = bool(data.get("check_da", True))
        t = threading.Thread(target=run_backlink_analysis,
                             args=(keywords, domain, delay, headless, country, proxies,
                                   browser_pref, vpn_method, check_da,
                                   latitude, longitude),
                             daemon=True)
    elif mode == "count":
        t = threading.Thread(target=run_count_analysis,
                             args=(keywords, delay, headless, country, proxies,
                                   browser_pref, vpn_method,
                                   latitude, longitude, city, lang),
                             daemon=True)
    else:
        t = threading.Thread(target=run_index_analysis,
                             args=(keywords, delay, headless, country, proxies,
                                   browser_pref, vpn_method,
                                   latitude, longitude, city, lang),
                             daemon=True)
    t.start()
    activity(f"{mode.title()} started - {domain or 'multi-target'} ({len(keywords)} keywords)")
    return jsonify({"status": "started", "total": len(keywords), "mode": mode})

def _pin_request_mode():
    """Pin the current request to a specific tool's job. Mode comes from JSON body or
    ?mode= query; defaults to ranking for backward compatibility."""
    m = None
    data = request.get_json(silent=True) or {}
    m = data.get("mode") or request.args.get("mode")
    _set_mode(m or "ranking")
    return getattr(_ctx, "mode")

@app.route("/api/pause", methods=["POST"])
def api_pause():
    _pin_request_mode()
    with state_lock:
        if state["status"] == "running":
            state["status"] = "paused"
    pause_event.clear()
    return jsonify({"status": "paused"})

@app.route("/api/resume", methods=["POST"])
def api_resume():
    _pin_request_mode()
    with state_lock:
        if state["status"] == "paused":
            state["status"] = "running"; state["captcha_msg"] = ""
    pause_event.set()
    return jsonify({"status": "running"})

@app.route("/api/stop", methods=["POST"])
def api_stop():
    _pin_request_mode()
    stop_event.set(); pause_event.set(); autosave()
    with state_lock:
        state["status"] = "stopped"
    return jsonify({"status": "stopped"})

@app.route("/api/reset", methods=["POST"])
def api_reset():
    _pin_request_mode()
    stop_event.set(); pause_event.set()
    with state_lock:
        state.update({"status": "idle", "current_keyword": "", "current_index": 0,
                      "total": 0, "results": [], "captcha_msg": "", "error_msg": "",
                      "log": [], "domain": ""})
    stop_event.clear(); clear_autosave()
    return jsonify({"status": "idle"})

@app.route("/api/cities")
def api_cities():
    cities = [{"name": k} for k in engine.CITY_CANONICAL]
    # Merge custom cities from CONFIG
    custom = CONFIG.get("custom_cities", {})
    custom_names = set()
    for display_name in custom:
        custom_names.add(display_name)
        if display_name not in engine.CITY_CANONICAL:
            cities.append({"name": display_name})
    return jsonify(cities)

# --------------------------------------------------------------------------- #
# Admin - City Management
# --------------------------------------------------------------------------- #
@app.route("/api/admin/cities")
def api_admin_cities():
    custom = CONFIG.get("custom_cities", {})
    builtin = list(engine.CITY_CANONICAL.keys())
    return jsonify({"custom": custom, "builtin": builtin})

@app.route("/api/admin/add_city", methods=["POST"])
def api_admin_add_city():
    data = request.get_json(silent=True) or {}
    display = data.get("display", "").strip()
    canonical = data.get("canonical", "").strip()
    lat = data.get("lat")
    lng = data.get("lng")
    if not display or not canonical:
        return jsonify({"error": "Display name and canonical name are required"})
    custom = CONFIG.setdefault("custom_cities", {})
    # Never silently overwrite an existing location (built-in or custom) - a
    # re-add with different coordinates would corrupt the lat/lng already
    # in use for that city.
    existing = next((k for k in list(custom.keys()) + list(engine.CITY_CANONICAL.keys())
                     if k.strip().lower() == display.lower()), None)
    if existing:
        return jsonify({"error": f'"{existing}" is already added - skipped to avoid overwriting its coordinates.'})
    custom[display] = canonical
    # Register in engine dicts so ranking/UULE work immediately
    engine.CITY_CANONICAL[display] = canonical
    if lat is not None and lng is not None:
        try:
            lat_f, lng_f = float(lat), float(lng)
            engine.CITY_COORDS[display] = (lat_f, lng_f)
            CONFIG.setdefault("custom_city_coords", {})[display] = [lat_f, lng_f]
        except (ValueError, TypeError):
            pass
    save_config(CONFIG)
    return jsonify({"ok": True, "display": display})

@app.route("/api/admin/remove_city", methods=["POST"])
def api_admin_remove_city():
    data = request.get_json(silent=True) or {}
    display = data.get("display", "").strip()
    if not display:
        return jsonify({"error": "Display name required"})
    custom = CONFIG.get("custom_cities", {})
    if display in custom:
        del custom[display]
        coords = CONFIG.get("custom_city_coords", {})
        coords.pop(display, None)
        engine.CITY_CANONICAL.pop(display, None)
        engine.CITY_COORDS.pop(display, None)
        save_config(CONFIG)
    return jsonify({"ok": True})

@app.route("/api/admin/upload_cities", methods=["POST"])
def api_admin_upload_cities():
    """Bulk-add cities from CSV text.  Expected columns:
       City, Country, Latitude, Longitude [, Canonical]
    Canonical is optional - when omitted it is generated as "City,Country".
    Display name is built as "City, CC" (e.g. "Perth, AU")."""
    import csv, io
    data = request.get_json(silent=True) or {}
    csv_text = data.get("csv", "").strip()
    if not csv_text:
        return jsonify({"error": "No CSV data provided"})
    reader = csv.DictReader(io.StringIO(csv_text))
    # Normalise header names to lowercase for flexible matching
    if not reader.fieldnames:
        return jsonify({"error": "CSV appears empty or has no header row"})
    header_map = {h.strip().lower(): h for h in reader.fieldnames}
    city_col = header_map.get("city")
    country_col = header_map.get("country")
    lat_col = header_map.get("latitude") or header_map.get("lat")
    lng_col = header_map.get("longitude") or header_map.get("lng") or header_map.get("lon")
    canonical_col = header_map.get("canonical")
    if not city_col or not country_col:
        return jsonify({"error": "CSV must have at least 'City' and 'Country' columns"})
    custom = CONFIG.setdefault("custom_cities", {})
    added = []
    skipped = []
    for i, row in enumerate(reader, start=2):
        city_name = (row.get(city_col) or "").strip()
        country = (row.get(country_col) or "").strip()
        if not city_name or not country:
            skipped.append(f"Row {i}: missing city or country")
            continue
        display = f"{city_name}, {country}"
        # Canonical: use provided value or generate "City,Country"
        canonical = ""
        if canonical_col:
            canonical = (row.get(canonical_col) or "").strip()
        if not canonical:
            canonical = f"{city_name},{country}"
        custom[display] = canonical
        engine.CITY_CANONICAL[display] = canonical
        # Coordinates (optional)
        if lat_col and lng_col:
            lat_s = (row.get(lat_col) or "").strip()
            lng_s = (row.get(lng_col) or "").strip()
            if lat_s and lng_s:
                try:
                    lat_f, lng_f = float(lat_s), float(lng_s)
                    engine.CITY_COORDS[display] = (lat_f, lng_f)
                    CONFIG.setdefault("custom_city_coords", {})[display] = [lat_f, lng_f]
                except (ValueError, TypeError):
                    pass
        added.append(display)
    save_config(CONFIG)
    return jsonify({"ok": True, "added": added, "added_count": len(added),
                    "skipped": skipped, "skipped_count": len(skipped)})

@app.route("/api/admin/sample_cities_csv")
def api_admin_sample_cities_csv():
    """Return a sample CSV file for bulk city upload."""
    sample = (
        "City,Country,Latitude,Longitude,Canonical\r\n"
        "Perth,AU,-31.9505,115.8605,\"Perth,Western Australia,Australia\"\r\n"
        "Sydney,AU,-33.8688,151.2093,\"Sydney,New South Wales,Australia\"\r\n"
        "Melbourne,AU,-37.8136,144.9631,\"Melbourne,Victoria,Australia\"\r\n"
        "Toronto,CA,43.6532,-79.3832,\"Toronto,Ontario,Canada\"\r\n"
        "Vancouver,CA,49.2827,-123.1207,\"Vancouver,British Columbia,Canada\"\r\n"
    )
    return Response(sample, mimetype="text/csv",
                    headers={"Content-Disposition": "attachment; filename=sample_cities.csv"})

@app.route("/api/screenshot/<path:filename>")
def api_screenshot(filename):
    safe = os.path.basename(filename)
    # Search in domain subfolders
    for root, dirs, files in os.walk(DOWNLOADS_DIR):
        if safe in files:
            return send_from_directory(root, safe)
    return send_from_directory(SCREENSHOTS_DIR, safe)

@app.route("/api/status")
def api_status():
    _pin_request_mode()
    with state_lock:
        snap = dict(state)
    snap.pop("driver", None)
    return jsonify(snap)

@app.route("/api/load-autosave")
def api_load_autosave():
    _pin_request_mode()
    data = load_autosave()
    if data and data.get("results"):
        with state_lock:
            state["results"] = data["results"]
            state["mode"] = data.get("mode", "ranking")
            state["status"] = "completed"
        return jsonify({"loaded": True, "count": len(data["results"]),
                        "saved_at": data.get("saved_at", "")})
    return jsonify({"loaded": False, "count": 0})

@app.route("/api/config", methods=["GET", "POST"])
def api_config():
    global CONFIG
    if request.method == "POST":
        global DOWNLOADS_DIR, SCREENSHOTS_DIR
        data = request.get_json(silent=True) or {}
        for k in DEFAULT_CONFIG:
            if k in data:
                CONFIG[k] = data[k]
        if "downloads_folder" in data:
            folder = (data["downloads_folder"] or "").strip()
            if folder and os.path.isdir(folder):
                CONFIG["downloads_folder"] = folder
                DOWNLOADS_DIR = folder
                SCREENSHOTS_DIR = folder
            elif not folder:
                CONFIG.pop("downloads_folder", None)
                DOWNLOADS_DIR = _DEFAULT_DOWNLOADS
                SCREENSHOTS_DIR = _DEFAULT_DOWNLOADS
        save_config(CONFIG)
        safe = {k: v for k, v in CONFIG.items() if k not in SENSITIVE_KEYS}
        return jsonify({"saved": True, "config": safe})
    cfg = {k: v for k, v in CONFIG.items() if k not in SENSITIVE_KEYS}
    cfg["downloads_folder"] = DOWNLOADS_DIR
    cfg["auth_api_url"] = CONFIG.get("auth_api_url", "")
    cfg["user_auth_url"] = CONFIG.get("user_auth_url", "")
    return jsonify(cfg)

@app.route("/api/export/csv")
def api_export_csv():
    _pin_request_mode()
    with state_lock:
        results = list(state["results"]); m = state.get("mode", "ranking")
        domain = state.get("domain", "")
    out = io.StringIO()
    if m == "ranking":
        # Collect all column names across all rows (handles multiple matches)
        has_targets = any(r.get("target_page") for r in results)
        base = ["keyword", "target_page", "status", "position", "serp_url"] if has_targets else ["keyword", "status", "position", "serp_url"]
        extra = set()
        for r in results:
            for k in r:
                if k not in base and k not in extra:
                    extra.add(k)
        # Sort extra columns: position_2, serp_url_2, position_3, ...; screenshot
        # columns fixed at the end (filename then its shareable URL) instead of
        # relying on set-iteration order, which isn't deterministic.
        def _col_sort(c):
            for i, prefix in enumerate(["position_", "serp_url_"]):
                if c.startswith(prefix):
                    num = c[len(prefix):]
                    return (int(num) if num.isdigit() else 99, i)
            if c == "screenshot":
                return (100, 0)
            if c == "screenshot_url":
                return (100, 1)
            return (99, 99)
        fields = base + sorted(extra, key=_col_sort)
    elif m == "backlink":
        fields = ["url", "domain_found", "meta_robots", "link_type", "da", "dr", "pa", "da_source", "link_url", "status", "checked_at"]
    elif m == "count":
        fields = ["keyword", "results_count"]
    else:
        fields = ["url", "indexed", "found_url", "status", "checked_at"]
    w = csv.DictWriter(out, fieldnames=fields, extrasaction="ignore")
    w.writeheader(); w.writerows(results); out.seek(0)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    _dslug = (domain or "report").lower().replace("https://", "").replace("http://", "").replace("www.", "").strip("/").split("/")[0].replace(":", "_") or "report"
    if m == "ranking":
        name = f"{_dslug}_rankings_{ts}.csv"
    elif m == "backlink":
        name = f"{_dslug}_backlink_check_{ts}.csv"
    elif m == "count":
        name = f"search_results_{ts}.csv"
    else:
        name = f"index_check_{ts}.csv"
    # Save to domain folder. UTF-8 BOM prefix - without it, Excel opens a UTF-8 CSV
    # using the system ANSI codepage instead, garbling any non-ASCII character
    # (em-dashes, accented characters in domains/URLs) into mojibake like "a€"".
    csv_data = "﻿" + out.getvalue()
    try:
        folder = _domain_folder(domain, m)
        with open(os.path.join(folder, name), "w", encoding="utf-8", newline="") as f:
            f.write(csv_data)
    except Exception:
        pass
    return Response(csv_data, mimetype="text/csv",
                    headers={"Content-Disposition": f"attachment; filename={name}"})

@app.route("/api/health")
def api_health():
    return jsonify({"ok": True, "version": APP_VERSION})


@app.route("/api/check-updates", methods=["POST"])
def api_check_updates():
    """Run the OTA updater on demand - used by the Settings 'Check for updates'
    button and the auto-check on open, so users never touch the command line.
    Downloads any changed files to disk; a relaunch applies backend changes."""
    try:
        result = updater.check_and_update(log_fn=lambda _m: None)
    except Exception as e:
        return jsonify({"updated": False, "updated_files": [], "failed": [], "reason": str(e)})
    return jsonify(result)


@app.route("/api/restart", methods=["POST"])
def api_restart():
    """Relaunch the app from a button: start a fresh launcher, which kills this server
    and opens a new window. The current window closes itself client-side."""
    try:
        vbs = os.path.join(BUNDLE_DIR, "Start Tool.vbs")
        if os.path.exists(vbs):
            subprocess.Popen(["wscript.exe", vbs], cwd=BUNDLE_DIR,
                             creationflags=getattr(subprocess, "DETACHED_PROCESS", 0))
            return jsonify({"ok": True})
        return jsonify({"ok": False, "error": "launcher not found"})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)})

# --------------------------------------------------------------------------- #
# SEO On-Page Report - runs phase2 script as subprocess
# --------------------------------------------------------------------------- #
onpage_state = {
    "status": "idle",  # idle, running, completed, error
    "log": [],
    "domain": "",
    "output_zip": "",
    "error_msg": "",
    "progress": "",
}
onpage_lock = threading.Lock()
onpage_stop = threading.Event()

SCRIPTS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "scripts")


# --------------------------------------------------------------------------- #
# On-page "target pages & keywords" parsing - order-agnostic (paste + Excel)
# --------------------------------------------------------------------------- #
import re as _re


def _looks_like_url(tok):
    """True if this token looks like a page URL/path rather than a keyword."""
    t = (tok or "").strip().lower()
    if not t:
        return False
    if t.startswith(("http://", "https://", "www.")):
        return True
    if "/" in t:                                  # relative path e.g. 'gold-coast/'
        return True
    return bool(_re.match(r"^[a-z0-9-]+(\.[a-z0-9-]+)+$", t))   # bare domain


def _split_keyword_rank(tok):
    """'keyword:12' -> ('keyword', '12'); 'keyword' -> ('keyword', None). Only splits
    on the LAST ':' when what follows looks like a ranking value (a number, or a
    not-ranked marker like 'NR'/'Not Found'), so a keyword that happens to contain a
    colon isn't misparsed."""
    t = (tok or "").strip()
    if ":" not in t:
        return t, None
    kw, _, rank = t.rpartition(":")
    kw, rank = kw.strip(), rank.strip()
    if not kw:
        return t, None
    if _re.match(r"^\d+(\.\d+)?$", rank) or rank.lower() in ("nr", "not found", "n/a"):
        return kw, (rank or None)
    return t, None


def _looks_like_rank(tok):
    """True if this token, on its own, looks like a ranking value rather than a
    keyword (used when a Ranking column lands in its own field, e.g. from Excel)."""
    t = (tok or "").strip().lower()
    return bool(t) and (bool(_re.match(r"^\d+(\.\d+)?$", t)) or t in ("nr", "not found", "n/a"))


def _parse_target_line(line):
    """Parse one row into (page, [keywords], {keyword: rank}).

    URL and keyword(s) may appear in EITHER column order, separated by a tab
    (Excel paste), a '|' (documented manual format), or 2+ spaces. Keywords may be
    comma-separated, each optionally suffixed 'keyword:12' with its ranking (from an
    optional Ranking column in the team's sheet). Relative paths are kept as-is (the
    phase-2 script resolves them against the domain via normalize_url)."""
    s = (line or "").strip()
    if not s:
        return None, [], {}
    if "\t" in s:
        parts = s.split("\t")
    elif "|" in s:
        parts = s.split("|")
    elif _re.search(r"\s{2,}", s):
        parts = _re.split(r"\s{2,}", s)
    else:
        parts = [s]
    parts = [p.strip() for p in parts if p.strip()]
    if not parts:
        return None, [], {}

    page = None
    kw_fields = []
    for tok in parts:
        if page is None and _looks_like_url(tok):
            page = tok
        else:
            kw_fields.append(tok)
    if page is None:                              # nothing url-like -> first field is the page
        page, kw_fields = parts[0], parts[1:]

    # A trailing lone field that's purely a ranking value (e.g. a separate Ranking
    # column, not a "keyword:12" suffix) applies to the single preceding keyword field.
    if len(kw_fields) >= 2 and _looks_like_rank(kw_fields[-1]) and "," not in kw_fields[-2]:
        rank_tok = kw_fields.pop()
        kw_fields[-1] = f"{kw_fields[-1]}:{rank_tok}"

    keywords, ranks = [], {}
    for kf in kw_fields:
        for k in kf.split(","):
            kw, rank = _split_keyword_rank(k)
            if kw and kw not in keywords:
                keywords.append(kw)
                if rank is not None:
                    ranks[kw] = rank
    return page, keywords, ranks


def _parse_onpage_targets(raw_text):
    """Pasted text -> flat [{"keyword":k, "page":url, "rank":r}] rows. This is the
    phase-2 loader's native form, which (unlike the grouped form) correctly keeps
    URL-only pages that have no keyword. De-duplicated, first-seen order."""
    rows, seen = [], set()
    for line in (raw_text or "").splitlines():
        page, kws, ranks = _parse_target_line(line)
        if not page:
            continue
        for k in (kws or [""]):
            key = (page, k)
            if key not in seen:
                seen.add(key)
                rows.append({"keyword": k, "page": page, "rank": ranks.get(k)})
    return rows


def _targets_to_lines(rows):
    """Flat rows -> 'URL | kw1, kw2:12' textarea lines (grouped by page), carrying
    each keyword's ranking (if any) as a 'keyword:rank' suffix."""
    grouped, order = {}, []
    for r in rows:
        p = r["page"]
        if p not in grouped:
            grouped[p] = []
            order.append(p)
        kw = r["keyword"]
        disp = f"{kw}:{r['rank']}" if kw and r.get("rank") not in (None, "") else kw
        if kw and disp not in grouped[p]:
            grouped[p].append(disp)
    return [f"{p} | {', '.join(grouped[p])}" if grouped[p] else p for p in order]


def _lines_from_excel(file_storage):
    """Read an uploaded .xlsx into 'URL | kw1, kw2:rank' lines, reusing the same
    order-agnostic parsing as pasted text (header row / either column order; an
    optional Ranking column is matched by header name, not fixed position)."""
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(file_storage.read()), data_only=True, read_only=True)
    raw_lines = []
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            if not row:
                continue
            cells = [("" if c is None else str(c).strip()) for c in row]
            nonempty = [c for c in cells if c]
            if not nonempty:
                continue
            low = " ".join(nonempty).lower()
            # skip a header row (labels only, no link/path in it)
            if ("keyword" in low or "target" in low) and ("page" in low or "url" in low) \
               and not any(("http" in c or "/" in c) for c in nonempty):
                continue
            raw_lines.append("\t".join(nonempty))
    return _targets_to_lines(_parse_onpage_targets("\n".join(raw_lines)))


def _run_onpage_report(domain, targets_json, fmt, no_capture):
    """Run the on-page phase2 script as a subprocess, stream logs."""
    with onpage_lock:
        onpage_state.update({"status": "running", "log": [], "domain": to_domain(domain),
                             "output_zip": "", "output_zip_backup": "", "error_msg": "",
                             "progress": "Checking working version..."})
    onpage_stop.clear()
    activity(f"SEO On-Page report started for {to_domain(domain)} ({fmt})")

    def _log0(msg):
        with onpage_lock:
            onpage_state["log"].append(msg)
            onpage_state["progress"] = msg

    # On-page takes a Website URL: before generating, find which version actually works
    # (follow the redirect chain to the live HTTP 200 version) and build the report for
    # THAT version - e.g. an entry of example.com that redirects to https://www.example.com/
    # produces a report for www.example.com, not a dead non-www host.
    _log0("Checking which website version is running...")
    final_url, host = resolve_working_url(domain, log=_log0)
    # Pass the canonical HOST to the script (it builds https://<host>); keep the working
    # subdomain (www or not) so the report reflects the version that actually serves.
    domain = to_domain(host or domain)
    with onpage_lock:
        onpage_state["domain"] = domain
        onpage_state["progress"] = "Starting..."

    # Save into the user's configured Downloads folder (per-domain), same as every
    # other tool - not a hidden "onpage_output" folder inside the install dir.
    out_dir = _domain_folder(domain, "onpage")
    os.makedirs(out_dir, exist_ok=True)

    targets_file = None
    if targets_json:
        targets_file = os.path.join(out_dir, f"_targets_{domain}.json")
        with open(targets_file, "w", encoding="utf-8") as f:
            json.dump(targets_json, f, ensure_ascii=False)

    python_exe = os.path.join(os.path.dirname(os.path.abspath(__file__)), "python", "python.exe")
    script = os.path.join(SCRIPTS_DIR, "generate_seo_onpage_phase2.py")

    cmd = [python_exe, "-u", script, domain, "--out", out_dir, "--format", fmt]
    if targets_file:
        cmd.extend(["--targets", targets_file])
    if no_capture:
        cmd.append("--no-capture")
    # Without this, the script picks an arbitrary connected GSC account (confirmed
    # live: it silently returned a completely unrelated domain's GSC data - "no
    # match" was falling back to that account's first property instead of
    # reporting "not found"). Resolve the account that actually manages this
    # domain via the same cached mapping GSC Audit/Health Audit already use.
    gsc_account = (_gsc_mapping().get(domain.strip().lower()) or {}).get("email", "")
    if gsc_account:
        cmd.extend(["--account", gsc_account])

    def _log(msg):
        with onpage_lock:
            onpage_state["log"].append(msg)
            if msg.startswith("["):
                onpage_state["progress"] = msg

    _log(f"Running: {domain} (format={fmt})")
    _log(f"Output dir: {out_dir}")

    try:
        # GEMINI_API_KEY is synced centrally (Admin -> Sync API Keys, same as the PSI
        # key) into CONFIG, not a machine env var - pass it through to the subprocess so
        # suggest_meta()'s Gemini call works team-wide without per-machine setup.
        proc_env = {**os.environ, "PYTHONDONTWRITEBYTECODE": "1"}
        gemini_key = CONFIG.get("gemini_api_key", "").strip()
        if gemini_key:
            proc_env["GEMINI_API_KEY"] = gemini_key
        proc = subprocess.Popen(
            cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, bufsize=1, cwd=SCRIPTS_DIR,
            env=proc_env
        )

        for line in proc.stdout:
            if onpage_stop.is_set():
                proc.kill()
                _log("Stopped by user.")
                with onpage_lock:
                    onpage_state["status"] = "stopped"
                return
            _log(line.rstrip())

        proc.wait()

        if proc.returncode != 0:
            _log(f"Script exited with code {proc.returncode}")
            with onpage_lock:
                onpage_state["status"] = "error"
                onpage_state["error_msg"] = f"Script failed (exit code {proc.returncode})"
            return

        # Find the output ZIP
        import glob
        zips = sorted(glob.glob(os.path.join(out_dir, f"*{domain}*.zip")),
                      key=os.path.getmtime, reverse=True)
        if not zips:
            zips = sorted(glob.glob(os.path.join(out_dir, "*.zip")),
                          key=os.path.getmtime, reverse=True)

        if zips:
            backup_path = _backup_report("onpage", domain, zips[0])
            with onpage_lock:
                onpage_state["status"] = "completed"
                onpage_state["output_zip"] = zips[0]
                onpage_state["output_zip_backup"] = backup_path or ""
                onpage_state["progress"] = "Report ready for download"
            _log(f"Report generated: {os.path.basename(zips[0])}")
        else:
            with onpage_lock:
                onpage_state["status"] = "error"
                onpage_state["error_msg"] = "No ZIP file found in output"
            _log("No output ZIP found")

    except Exception as e:
        _log(f"Error: {e}")
        with onpage_lock:
            onpage_state["status"] = "error"
            onpage_state["error_msg"] = str(e)

    finally:
        if targets_file and os.path.exists(targets_file):
            try:
                os.remove(targets_file)
            except Exception:
                pass


@app.route("/api/onpage/start", methods=["POST"])
def api_onpage_start():
    with onpage_lock:
        if onpage_state["status"] == "running":
            return jsonify({"error": "On-page report already running."}), 400

    data = request.get_json(silent=True) or {}
    domain = (data.get("domain") or "").strip()
    if not domain:
        return jsonify({"error": "Domain is required."}), 400

    fmt = data.get("format", "james")
    no_capture = bool(data.get("no_capture", False))

    targets = None
    targets_raw = (data.get("targets") or "").strip()
    if targets_raw:
        try:
            parsed = json.loads(targets_raw)
            targets = parsed if isinstance(parsed, list) else None
        except json.JSONDecodeError:
            targets = None
        if targets is None:
            # Parse pasted "URL | kw", "URL <tab> kw" or "kw <tab> URL" rows
            # (either column order) into the flat form the phase-2 loader reads.
            targets = _parse_onpage_targets(targets_raw)

    t = threading.Thread(target=_run_onpage_report,
                         args=(domain, targets, fmt, no_capture), daemon=True)
    t.start()
    return jsonify({"status": "started", "domain": domain})


@app.route("/api/onpage/status")
def api_onpage_status():
    with onpage_lock:
        return jsonify({
            "status": onpage_state["status"],
            "log": onpage_state["log"][-200:],
            "domain": onpage_state["domain"],
            "progress": onpage_state["progress"],
            "error_msg": onpage_state["error_msg"],
            "has_zip": bool(onpage_state["output_zip"]),
        })


@app.route("/api/onpage/stop", methods=["POST"])
def api_onpage_stop():
    onpage_stop.set()
    return jsonify({"status": "stopping"})


@app.route("/api/onpage/download")
def api_onpage_download():
    with onpage_lock:
        zip_path = onpage_state.get("output_zip", "")
        backup_path = onpage_state.get("output_zip_backup", "")
        domain = onpage_state.get("domain", "")
    if not zip_path or not os.path.exists(zip_path):
        # Falls back to the durable backup copy (REPORT_BACKUPS_DIR) if the user
        # deleted/moved the folder the report was originally saved into.
        zip_path = backup_path if backup_path and os.path.exists(backup_path) else ""
        if not zip_path and domain:
            candidate = _backup_report_path("onpage", domain)
            if os.path.exists(candidate):
                zip_path = candidate
    if not zip_path:
        return jsonify({"error": "No report available for download."}), 404
    resp = send_file(zip_path, as_attachment=True,
                     download_name=os.path.basename(zip_path))
    resp.headers["Cache-Control"] = "no-store"
    return resp


# --------------------------------------------------------------------------- #
# Wayback Machine Submitter - standalone tool. Pure HTTP (no Selenium session
# needed), so it's a lightweight background job like Brief Analysis rather than
# a full checker Session.
# --------------------------------------------------------------------------- #
wayback_state = {"status": "idle", "log": [], "results": [], "error_msg": ""}
wayback_lock = threading.Lock()
wayback_stop = threading.Event()


def _wayback_proxy_url(p):
    auth = f"{p['user']}:{p['pass']}@" if p.get("user") else ""
    return f"{p.get('type', 'http')}://{auth}{p['host']}:{p['port']}"


def _submit_wayback_url(url, max_tries=3, timeout=45):
    """Submit `url` to the Wayback Machine's Save Page Now, rotating through a
    different proxy each attempt (archive.org blocks/limits by IP, so retrying on
    the SAME IP would just fail the same way). Capped at max_tries so a slow/blocked
    archive.org can never hang the job - returns None on exhausted retries."""
    import re as _re
    proxies_pool = list(CONFIG.get("proxies", [])) + _shared_proxies()
    attempts = (random.sample(proxies_pool, min(max_tries, len(proxies_pool)))
                if proxies_pool else [None] * max_tries)
    save_url = "https://web.archive.org/save/" + url
    for proxy in attempts:
        try:
            kwargs = {}
            if proxy:
                pu = _wayback_proxy_url(proxy)
                kwargs["proxies"] = {"http": pu, "https": pu}
            r = http_requests.get(save_url, headers={"User-Agent": "Mozilla/5.0 SEOToolkitPro"},
                                  timeout=timeout, allow_redirects=True, **kwargs)
            # Content-Location is Wayback's own redirect header for exactly the page
            # just requested - authoritative, prefer it over scraping the HTML.
            loc = r.headers.get("Content-Location", "")
            if loc and _re.search(r"/web/\d{10,}/https?://", loc):
                return "https://web.archive.org" + loc

            # Fallback: the results page's "Visit page: <a href=...>" line, NOT the
            # resource listing further down (a page pulls in dozens of images/CSS/JS,
            # each with its own /web/.../ link in that list - grabbing the first match
            # anywhere in the page picked up a random embedded image instead of the
            # page itself). Require the archived URL to actually match what was
            # submitted, not just look like a Wayback link.
            for m in _re.finditer(r'/web/\d{10,}/(https?://[^\s"\'<>]+)', r.text):
                archived_target = m.group(1).rstrip("/")
                if archived_target == url.rstrip("/"):
                    return "https://web.archive.org" + m.group(0)
        except Exception:
            continue
    return None


def _run_wayback_submit(urls):
    with wayback_lock:
        wayback_state.update({"status": "running", "log": [], "results": [], "error_msg": ""})
    wayback_stop.clear()
    activity(f"Wayback submission started for {len(urls)} URL(s)")
    for i, u in enumerate(urls):
        if wayback_stop.is_set():
            with wayback_lock:
                wayback_state["log"].append("Stopped by user.")
                wayback_state["status"] = "stopped"
            return
        with wayback_lock:
            wayback_state["log"].append(f"[{i+1}/{len(urls)}] Submitting {u}...")
        archived = _submit_wayback_url(u)
        with wayback_lock:
            wayback_state["results"].append({
                "url": u, "archived_url": archived or "",
                "status": "submitted" if archived else "failed",
                "checked_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")})
            wayback_state["log"].append(
                "  -> " + (f"Archived: {archived}" if archived else "Failed after retries"))
    with wayback_lock:
        ok = sum(1 for r in wayback_state["results"] if r["status"] == "submitted")
        wayback_state["log"].append(f"Completed -- {ok} ok, {len(wayback_state['results']) - ok} error(s).")
        wayback_state["status"] = "completed"


@app.route("/api/wayback/start", methods=["POST"])
def api_wayback_start():
    with wayback_lock:
        if wayback_state["status"] == "running":
            return jsonify({"error": "Wayback submission already running."}), 400
    data = request.get_json(silent=True) or {}
    raw = (data.get("urls") or "").strip()
    seen = set()
    urls = []
    for u in raw.splitlines():
        u = u.strip()
        if not u:
            continue
        key = u.rstrip("/").lower()
        if key not in seen:
            seen.add(key)
            urls.append(u)
    if not urls:
        return jsonify({"error": "At least one URL required."}), 400
    if len(urls) > 20:
        return jsonify({"error": "Max 20 URLs per batch - archive.org blocks bulk submissions. "
                                  "Split into smaller batches."}), 400
    t = threading.Thread(target=_run_wayback_submit, args=(urls,), daemon=True)
    t.start()
    return jsonify({"status": "started", "count": len(urls)})


@app.route("/api/wayback/status")
def api_wayback_status():
    with wayback_lock:
        return jsonify({
            "status": wayback_state["status"],
            "log": wayback_state["log"][-200:],
            "results": wayback_state["results"],
            "error_msg": wayback_state["error_msg"],
        })


@app.route("/api/wayback/stop", methods=["POST"])
def api_wayback_stop():
    wayback_stop.set()
    return jsonify({"status": "stopping"})


@app.route("/api/wayback/export")
def api_wayback_export():
    with wayback_lock:
        results = list(wayback_state["results"])
    out = io.StringIO()
    w = csv.DictWriter(out, fieldnames=["url", "archived_url", "status", "checked_at"])
    w.writeheader(); w.writerows(results)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    return Response(out.getvalue(), mimetype="text/csv",
                    headers={"Content-Disposition": f"attachment; filename=wayback_submissions_{ts}.csv"})


# --------------------------------------------------------------------------- #
# SEranking Audit - turns an uploaded SEranking Site Audit .xlsx export into a
# suggestion-filled final workbook (subprocess, same pattern as On-Page).
# --------------------------------------------------------------------------- #
sr_state = {"status": "idle", "log": [], "output_file": "", "output_file_backup": "", "error_msg": ""}
sr_lock = threading.Lock()
sr_stop = threading.Event()


def _run_seranking_audit(in_path, pdf_path, brand, zip_path=None):
    with sr_lock:
        sr_state.update({"status": "running", "log": [], "output_file": "",
                         "output_file_backup": "", "error_msg": ""})
    sr_stop.clear()
    src_name = os.path.basename(in_path or zip_path or pdf_path)
    activity(f"SEranking audit started ({src_name})")

    def _log(msg):
        with sr_lock:
            sr_state["log"].append(msg)

    slug = os.path.splitext(src_name)[0]
    out_dir = os.path.join(UPLOADS_DIR, "seranking_out")
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, f"Final Audit - {slug}.xlsx")

    python_exe = os.path.join(os.path.dirname(os.path.abspath(__file__)), "python", "python.exe")
    script = os.path.join(SCRIPTS_DIR, "generate_seranking_audit.py")
    cmd = [python_exe, "-u", script, "--out", out_path, "--brand", brand]
    if in_path:
        cmd += ["--in", in_path]
    if pdf_path:
        cmd += ["--pdf", pdf_path]
    if zip_path:
        cmd += ["--zip", zip_path]

    try:
        proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                text=True, bufsize=1, cwd=SCRIPTS_DIR,
                                env={**os.environ, "PYTHONDONTWRITEBYTECODE": "1"})
        for line in proc.stdout:
            if sr_stop.is_set():
                proc.kill()
                _log("Stopped by user.")
                with sr_lock:
                    sr_state["status"] = "stopped"
                return
            _log(line.rstrip())
        proc.wait()
        if proc.returncode != 0 or not os.path.exists(out_path):
            with sr_lock:
                sr_state["status"] = "error"
                sr_state["error_msg"] = f"Script failed (exit code {proc.returncode})"
            _log(f"Script exited with code {proc.returncode}")
            return
        backup_path = _backup_report("seranking", slug, out_path)
        with sr_lock:
            sr_state["status"] = "completed"
            sr_state["output_file"] = out_path
            sr_state["output_file_backup"] = backup_path or ""
        _log("Done.")
    except Exception as e:
        _log(f"Error: {e}")
        with sr_lock:
            sr_state["status"] = "error"
            sr_state["error_msg"] = str(e)
    finally:
        for _p in (in_path, pdf_path, zip_path):
            if _p:
                try:
                    os.remove(_p)
                except Exception:
                    pass


@app.route("/api/seranking/start", methods=["POST"])
def api_seranking_start():
    with sr_lock:
        if sr_state["status"] == "running":
            return jsonify({"error": "SEranking audit already running."}), 400
    xf = request.files.get("file")
    pf = request.files.get("pdf")
    zf = request.files.get("zip")
    if (not xf or not xf.filename) and (not pf or not pf.filename) and (not zf or not zf.filename):
        return jsonify({"error": "Upload a SEranking .xlsx export, a PDF audit export, "
                                  "and/or a zip of per-issue .xls exports."}), 400
    if xf and xf.filename and not xf.filename.lower().endswith((".xlsx", ".xls")):
        return jsonify({"error": "The Excel file must be .xlsx/.xls."}), 400
    if pf and pf.filename and not pf.filename.lower().endswith(".pdf"):
        return jsonify({"error": "The PDF file must be .pdf."}), 400
    if zf and zf.filename and not zf.filename.lower().endswith(".zip"):
        return jsonify({"error": "The zip file must be .zip."}), 400
    brand = (request.form.get("brand") or "").strip()
    os.makedirs(UPLOADS_DIR, exist_ok=True)
    ts = int(time.time())
    in_path = None
    pdf_path = None
    zip_path = None
    if xf and xf.filename:
        in_path = os.path.join(UPLOADS_DIR, f"seranking_{ts}_{os.path.basename(xf.filename)}")
        xf.save(in_path)
    if pf and pf.filename:
        pdf_path = os.path.join(UPLOADS_DIR, f"seranking_{ts}_{os.path.basename(pf.filename)}")
        pf.save(pdf_path)
    if zf and zf.filename:
        zip_path = os.path.join(UPLOADS_DIR, f"seranking_{ts}_{os.path.basename(zf.filename)}")
        zf.save(zip_path)
    t = threading.Thread(target=_run_seranking_audit, args=(in_path, pdf_path, brand, zip_path), daemon=True)
    t.start()
    return jsonify({"status": "started"})


@app.route("/api/seranking/status")
def api_seranking_status():
    with sr_lock:
        return jsonify({
            "status": sr_state["status"],
            "log": sr_state["log"][-200:],
            "error_msg": sr_state["error_msg"],
            "has_file": bool(sr_state["output_file"]),
        })


@app.route("/api/seranking/stop", methods=["POST"])
def api_seranking_stop():
    sr_stop.set()
    return jsonify({"status": "stopping"})


@app.route("/api/seranking/download")
def api_seranking_download():
    with sr_lock:
        out_path = sr_state.get("output_file", "")
        backup_path = sr_state.get("output_file_backup", "")
    if not out_path or not os.path.exists(out_path):
        out_path = backup_path if backup_path and os.path.exists(backup_path) else ""
    if not out_path:
        return jsonify({"error": "No report available for download."}), 404
    resp = send_file(out_path, as_attachment=True, download_name=os.path.basename(out_path))
    resp.headers["Cache-Control"] = "no-store"
    return resp


# --------------------------------------------------------------------------- #
# GEO / AI Optimization - runs generate_geo_report.py (subprocess, same pattern
# as SEranking above). NEW and UNTESTED - superadmin-only for now (server-side
# enforced below, not just hidden in the UI), per explicit instruction: promote
# to admin, then everyone, only once approved. Each stage is a one-line change
# to the role check in _require_geo_role().
# --------------------------------------------------------------------------- #
GEO_MIN_ROLE = "superadmin"  # -> "admin" once approved for admins, then remove the check entirely


def _current_user_role():
    """Same role resolution as /api/auth/is_admin, without the account-caching
    side effects - just what's needed to gate a route."""
    accounts = auth.list_logged_in()
    if not accounts:
        return "", ""
    email = accounts[0]["email"]
    try:
        result = auth._api_call({"action": "is_admin", "email": email})
    except Exception:
        return "", email
    return (result.get("role") or ""), email


def _require_geo_role():
    """Returns None if the current logged-in user is allowed to use GEO, else a
    (response, status_code) tuple to return immediately. Server-side check - a
    hidden UI tab alone doesn't stop someone calling the API directly."""
    role, email = _current_user_role()
    allowed = {"superadmin"} if GEO_MIN_ROLE == "superadmin" else {"superadmin", "admin"}
    if role not in allowed:
        return jsonify({"error": "GEO/AI Optimization is not available for your account yet."}), 403
    return None


geo_state = {"status": "idle", "log": [], "domain": "", "output_file": "", "error_msg": ""}
geo_lock = threading.Lock()
geo_stop = threading.Event()


def _run_geo_report(domain, targets_path, check_visibility, keywords):
    with geo_lock:
        geo_state.update({"status": "running", "log": [], "domain": domain,
                          "output_file": "", "error_msg": ""})
    geo_stop.clear()
    activity(f"GEO report started ({domain})")

    def _log(msg):
        with geo_lock:
            geo_state["log"].append(msg)

    out_dir = os.path.join(UPLOADS_DIR, "geo_out")
    os.makedirs(out_dir, exist_ok=True)

    python_exe = os.path.join(os.path.dirname(os.path.abspath(__file__)), "python", "python.exe")
    script = os.path.join(SCRIPTS_DIR, "generate_geo_report.py")
    cmd = [python_exe, "-u", script, domain, "--out", out_dir]
    if targets_path:
        cmd += ["--targets", targets_path]
    if not check_visibility:
        cmd += ["--no-visibility-check"]
    if keywords:
        cmd += ["--keywords", keywords]

    try:
        proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                text=True, bufsize=1, cwd=SCRIPTS_DIR,
                                env={**os.environ, "PYTHONDONTWRITEBYTECODE": "1"})
        for line in proc.stdout:
            if geo_stop.is_set():
                proc.kill()
                _log("Stopped by user.")
                with geo_lock:
                    geo_state["status"] = "stopped"
                return
            _log(line.rstrip())
        proc.wait()
        out_path = os.path.join(out_dir, f"GEO Report - {domain}.zip")
        if proc.returncode != 0 or not os.path.exists(out_path):
            with geo_lock:
                geo_state["status"] = "error"
                geo_state["error_msg"] = f"Script failed (exit code {proc.returncode})"
            _log(f"Script exited with code {proc.returncode}")
            return
        with geo_lock:
            geo_state["status"] = "completed"
            geo_state["output_file"] = out_path
        _log("Done.")
    except Exception as e:
        _log(f"Error: {e}")
        with geo_lock:
            geo_state["status"] = "error"
            geo_state["error_msg"] = str(e)
    finally:
        if targets_path:
            try:
                os.remove(targets_path)
            except Exception:
                pass


@app.route("/api/geo/start", methods=["POST"])
def api_geo_start():
    denied = _require_geo_role()
    if denied:
        return denied
    with geo_lock:
        if geo_state["status"] == "running":
            return jsonify({"error": "GEO report already running."}), 400
    domain = (request.form.get("domain") or "").strip()
    if not domain:
        return jsonify({"error": "Domain is required."}), 400
    check_visibility = request.form.get("check_visibility", "1") != "0"
    keywords = (request.form.get("keywords") or "").strip() or None
    targets_path = None
    tf = request.files.get("targets")
    if tf and tf.filename:
        os.makedirs(UPLOADS_DIR, exist_ok=True)
        targets_path = os.path.join(UPLOADS_DIR, f"geo_{int(time.time())}_{os.path.basename(tf.filename)}")
        tf.save(targets_path)
    t = threading.Thread(target=_run_geo_report, args=(domain, targets_path, check_visibility, keywords), daemon=True)
    t.start()
    return jsonify({"status": "started"})


@app.route("/api/geo/status")
def api_geo_status():
    denied = _require_geo_role()
    if denied:
        return denied
    with geo_lock:
        return jsonify({
            "status": geo_state["status"],
            "log": geo_state["log"][-200:],
            "error_msg": geo_state["error_msg"],
            "has_file": bool(geo_state["output_file"]),
        })


@app.route("/api/geo/stop", methods=["POST"])
def api_geo_stop():
    denied = _require_geo_role()
    if denied:
        return denied
    geo_stop.set()
    return jsonify({"status": "stopping"})


@app.route("/api/geo/download")
def api_geo_download():
    denied = _require_geo_role()
    if denied:
        return denied
    with geo_lock:
        out_path = geo_state.get("output_file", "")
    if not out_path or not os.path.exists(out_path):
        return jsonify({"error": "No report available for download."}), 404
    resp = send_file(out_path, as_attachment=True, download_name=os.path.basename(out_path))
    resp.headers["Cache-Control"] = "no-store"
    return resp


# --------------------------------------------------------------------------- #
# Health Audit - runs checks + optional Selenium screenshots, builds report
# --------------------------------------------------------------------------- #
ha_state = {
    "status": "idle",  # idle, running, completed, error
    "log": [],
    "domain": "",
    "output_file": "",
    "error_msg": "",
    "progress": "",
}
ha_lock = threading.Lock()
ha_stop = threading.Event()


def _run_health_audit(domain, fmt, target_pages, no_capture, headless, browser_name, psi_api_key=None):
    with ha_lock:
        ha_state.update({"status": "running", "log": [], "domain": domain,
                         "output_file": "", "error_msg": "", "progress": "Starting..."})
    ha_stop.clear()

    def _log(msg):
        with ha_lock:
            ha_state["log"].append(msg)
            if msg.startswith("["):
                ha_state["progress"] = msg

    out_folder = _domain_folder(domain, "health_audit")
    driver = None

    try:
        import concurrent.futures
        slow_executor = concurrent.futures.ThreadPoolExecutor(max_workers=2)
        if psi_api_key:
            _log("Starting Sucuri & PageSpeed in background...")
            fut_psi = slow_executor.submit(health_audit.check_pagespeed, domain, psi_api_key)
        else:
            _log("Starting Sucuri in background (PageSpeed skipped)...")
            fut_psi = None
        fut_sucuri = slow_executor.submit(health_audit.check_sucuri, domain)

        _log("[0/3] Launching browser for Google checks" + (" & screenshots..." if not no_capture else "..."))
        profile = pick_profile()
        driver = engine.build_driver(
            profile, proxy=None, headless=headless,
            country="us", extra_extensions=[],
            logger=_log, browser_pref=browser_name,
        )

        path = health_audit.run_health_audit(
            domain, fmt=fmt, target_pages=target_pages,
            out_dir=out_folder, driver=driver, no_capture=no_capture,
            log_fn=_log, psi_api_key=psi_api_key,
            prefetched_futures={"sucuri": fut_sucuri, **({"psi": fut_psi} if fut_psi else {})},
        )

        with ha_lock:
            ha_state["status"] = "completed"
            ha_state["output_file"] = path
            ha_state["progress"] = "Report ready for download"
        _log(f"Report saved: {os.path.basename(path)}")

    except Exception as e:
        _log(f"Error: {e}")
        with ha_lock:
            ha_state["status"] = "error"
            ha_state["error_msg"] = str(e)
    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass


@app.route("/api/health-audit/start", methods=["POST"])
def api_ha_start():
    with ha_lock:
        if ha_state["status"] == "running":
            return jsonify({"error": "Health audit already running."}), 400

    data = request.get_json(silent=True) or {}
    # Health audit is domain-based (it does its own www/non-www version detection),
    # so trim a pasted URL down to the bare domain.
    domain = to_domain(data.get("domain") or "")
    if not domain:
        return jsonify({"error": "Domain is required."}), 400

    fmt = data.get("format", "james")
    no_capture = bool(data.get("no_capture", False))
    headless = data.get("headless", True)
    browser_name = data.get("browser", "edge")

    target_pages = []
    targets_raw = (data.get("targets") or "").strip()
    if targets_raw:
        target_pages = [l.strip() for l in targets_raw.splitlines() if l.strip()]

    include_psi = bool(data.get("include_psi", False))
    psi_key = CONFIG.get("psi_api_key", "").strip() or None
    if not include_psi:
        psi_key = None
    t = threading.Thread(target=_run_health_audit,
                         args=(domain, fmt, target_pages, no_capture, headless, browser_name, psi_key),
                         daemon=True)
    t.start()
    activity(f"Health audit started - {domain} ({fmt})")
    return jsonify({"status": "started", "domain": domain})


@app.route("/api/health-audit/status")
def api_ha_status():
    with ha_lock:
        return jsonify({
            "status": ha_state["status"],
            "log": ha_state["log"][-200:],
            "domain": ha_state["domain"],
            "progress": ha_state["progress"],
            "error_msg": ha_state["error_msg"],
            "has_file": bool(ha_state["output_file"]),
        })


@app.route("/api/health-audit/stop", methods=["POST"])
def api_ha_stop():
    ha_stop.set()
    return jsonify({"status": "stopping"})


@app.route("/api/health-audit/download")
def api_ha_download():
    with ha_lock:
        fpath = ha_state.get("output_file", "")
    if not fpath or not os.path.exists(fpath):
        return jsonify({"error": "No report available."}), 404
    return send_file(fpath, as_attachment=True,
                     download_name=os.path.basename(fpath))


@app.route("/api/health-audit/formats")
def api_ha_formats():
    return jsonify([{"value": k, "label": v["label"], "ext": v["ext"]}
                    for k, v in health_audit.FORMAT_INFO.items()])


# --------------------------------------------------------------------------- #
# GSC Audit
# --------------------------------------------------------------------------- #
gsc_state = {
    "status": "idle",
    "log": [],
    "domain": "",
    "output_file": "",
    "error_msg": "",
    "progress": "",
}
gsc_lock = threading.Lock()
gsc_stop = threading.Event()

GSC_PROFILE_DIR = os.path.join(DATA_DIR, "gsc_profile")
os.makedirs(GSC_PROFILE_DIR, exist_ok=True)


def _gsc_webapp_url():
    """GSC / Crawl endpoint. Prefers the logged-in user's BUILDING GSC script (returned
    by the gateway at login) so a teammate only ever reaches their own building's GSC
    accounts. Falls back to the configured default (PTP) for super admins / if unset."""
    try:
        u = auth.get_gsc_url()
        if u:
            return u
    except Exception:
        pass
    return CONFIG.get("auth_api_url", "").strip()


_gsc_mapping_cache = {"data": {}, "ts": 0, "url": None}
_gsc_mapping_lock = threading.Lock()
GSC_MAPPING_TTL = 3600  # background refresh interval - keeps the cache warm so GSC
                        # Audit / Health Audit's domain lookup never waits on a live
                        # Sheets read, without hammering the sheet every few minutes


def _fetch_gsc_mapping_now(webapp_url):
    try:
        resp = http_requests.get(webapp_url, params={"action": "get_config"}, timeout=20)
        data = resp.json()
        if data.get("success"):
            with _gsc_mapping_lock:
                _gsc_mapping_cache["data"] = data.get("mapping", {})
                _gsc_mapping_cache["ts"] = time.time()
                _gsc_mapping_cache["url"] = webapp_url
            return True
    except Exception:
        pass
    return False


def _gsc_mapping():
    """Cached domain -> GSC-account mapping. Kept warm by _gsc_mapping_prefetch_loop
    (started at app launch); only makes a live Apps Script call itself if the
    background loop hasn't populated the cache for this user's building yet."""
    webapp_url = _gsc_webapp_url()
    if not webapp_url:
        return {}
    with _gsc_mapping_lock:
        have_current = _gsc_mapping_cache["url"] == webapp_url and _gsc_mapping_cache["data"]
        data = dict(_gsc_mapping_cache["data"]) if have_current else {}
    if not have_current:
        _fetch_gsc_mapping_now(webapp_url)
        with _gsc_mapping_lock:
            data = dict(_gsc_mapping_cache["data"]) if _gsc_mapping_cache["url"] == webapp_url else {}
    return data


def _gsc_mapping_prefetch_loop():
    """Waits for the logged-in user's building GSC URL to become available (a few
    seconds to a couple minutes after login), fetches the domain->account mapping in
    the background, then refreshes it every GSC_MAPPING_TTL seconds for the life of
    the app - so GSC Audit / Health Audit's domain lookup is instant by the time
    someone actually opens those tabs."""
    while True:
        webapp_url = _gsc_webapp_url()
        if webapp_url:
            _fetch_gsc_mapping_now(webapp_url)
            time.sleep(GSC_MAPPING_TTL)
        else:
            time.sleep(15)


def _run_gsc_audit(domain, email, fmt, headless, browser_name):
    with gsc_lock:
        gsc_state.update({"status": "running", "log": [], "domain": domain,
                          "output_file": "", "error_msg": "", "progress": "Starting..."})
    gsc_stop.clear()

    def _log(msg):
        with gsc_lock:
            gsc_state["log"].append(msg)
            if msg.startswith("[") or msg.startswith("  "):
                gsc_state["progress"] = msg

    out_folder = _domain_folder(domain, "gsc_audit")

    try:
        # Screenshots are captured from the per-account logged-in session inside
        # run_gsc_audit (each session opens its own browser), so we no longer launch a
        # shared-profile browser here - that shared profile isn't signed in and was
        # screenshotting the Google sign-in page instead of real GSC data.
        _log("[1/2] Running GSC audit (API data + session screenshots)...")
        # Same normalization getConfigForExtension() uses for its mapping keys
        # (domain.toString().trim().toLowerCase()) - just for the alert email's
        # "Access Level" display, not anything access-critical.
        access_level = (_gsc_mapping().get(domain.strip().lower()) or {}).get("accessLevel", "")
        path = gsc_audit.run_gsc_audit(
            domain, email, fmt=fmt, out_dir=out_folder, log_fn=_log,
            webapp_url=_gsc_webapp_url(), access_level=access_level,
        )

        with gsc_lock:
            gsc_state["status"] = "completed"
            gsc_state["output_file"] = path
            gsc_state["progress"] = "Report ready for download"
        _log(f"[2/2] Report saved: {os.path.basename(path)}")

    except Exception as e:
        _log(f"Error: {e}")
        with gsc_lock:
            gsc_state["status"] = "error"
            gsc_state["error_msg"] = str(e)


@app.route("/api/gsc/accounts")
def api_gsc_accounts():
    return jsonify(gsc_audit.list_accounts())


@app.route("/api/gsc/connect", methods=["POST"])
def api_gsc_connect():
    """Launch OAuth in a Selenium popup for the user to connect a Google account."""
    data = request.get_json(silent=True) or {}
    headless = data.get("headless", False)
    browser_name = data.get("browser", "edge")

    config = gsc_audit._load_gsc_config()
    client_id = config.get("gsc_client_id", "").strip()
    client_secret = config.get("gsc_client_secret", "").strip()
    if not client_id or not client_secret:
        return jsonify({"error": "GSC OAuth Client ID and Secret not configured. Set them in Settings."}), 400

    # Run OAuth inside a per-account browser SESSION profile so the very same
    # signed-in browser is saved for the Manual Action / Security screenshot
    # capture - one login covers both the API token AND the screenshots. Using a
    # fresh profile per account also avoids Google's ~10-accounts-per-browser cap.
    session = gsc_audit.create_session(label="Connecting...")
    driver = None
    try:
        driver = engine.build_driver(
            session["profile_dir"], proxy=None, headless=headless,
            country="us", extra_extensions=[],
            browser_pref=browser_name,
        )
        email = gsc_audit.oauth_login_selenium(driver, client_id, client_secret)
        # Tag this session with the account, and drop any older session for the
        # same email so sessions don't pile up on reconnect.
        for s in gsc_audit.list_sessions():
            if s["id"] != session["id"] and email.lower() in [a.lower() for a in s.get("accounts", [])]:
                gsc_audit.remove_session(s["id"])
        gsc_audit.set_session_account(session["id"], email)
        return jsonify({"status": "connected", "email": email})
    except Exception as e:
        gsc_audit.remove_session(session["id"])  # clean up the empty session on failure
        return jsonify({"error": str(e)}), 500
    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass


@app.route("/api/gsc/disconnect", methods=["POST"])
def api_gsc_disconnect():
    data = request.get_json(silent=True) or {}
    email = (data.get("email") or "").strip()
    if email:
        gsc_audit.remove_account(email)
    return jsonify({"status": "removed"})


@app.route("/api/gsc/properties")
def api_gsc_properties():
    email = request.args.get("email", "").strip()
    if not email:
        return jsonify({"error": "Email required"}), 400
    try:
        token = gsc_audit.get_access_token(email)
        props = gsc_audit.list_properties(token)
        return jsonify(props)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/gsc/start", methods=["POST"])
def api_gsc_start():
    with gsc_lock:
        if gsc_state["status"] == "running":
            return jsonify({"error": "GSC audit already running."}), 400

    data = request.get_json(silent=True) or {}
    # GSC audit is domain/property-based - trim a pasted URL to the bare domain.
    domain = to_domain(data.get("domain") or "")
    email = (data.get("email") or "").strip()
    if not domain:
        return jsonify({"error": "Domain is required."}), 400
    if not email:
        return jsonify({"error": "Please connect a Google account first."}), 400

    fmt = data.get("format", "james")
    headless = data.get("headless", True)
    browser_name = data.get("browser", "edge")

    t = threading.Thread(target=_run_gsc_audit,
                         args=(domain, email, fmt, headless, browser_name),
                         daemon=True)
    t.start()
    return jsonify({"status": "started", "domain": domain})


@app.route("/api/gsc/status")
def api_gsc_status():
    with gsc_lock:
        return jsonify({
            "status": gsc_state["status"],
            "log": gsc_state["log"][-200:],
            "domain": gsc_state["domain"],
            "progress": gsc_state["progress"],
            "error_msg": gsc_state["error_msg"],
            "has_file": bool(gsc_state["output_file"]),
        })


@app.route("/api/gsc/stop", methods=["POST"])
def api_gsc_stop():
    gsc_stop.set()
    return jsonify({"status": "stopping"})


@app.route("/api/gsc/download")
def api_gsc_download():
    with gsc_lock:
        fpath = gsc_state.get("output_file", "")
    if not fpath or not os.path.exists(fpath):
        return jsonify({"error": "No report available."}), 404
    return send_file(fpath, as_attachment=True,
                     download_name=os.path.basename(fpath))


@app.route("/api/gsc/formats")
def api_gsc_formats():
    return jsonify([{"value": k, "label": v["label"]}
                    for k, v in gsc_audit.GSC_FORMATS.items()])


# --------------------------------------------------------------------------- #
# GSC Browser Sessions
# --------------------------------------------------------------------------- #

_session_drivers = {}

@app.route("/api/gsc/sessions")
def api_gsc_sessions():
    return jsonify(gsc_audit.list_sessions())


@app.route("/api/gsc/sessions/create", methods=["POST"])
def api_gsc_session_create():
    data = request.get_json(silent=True) or {}
    label = (data.get("label") or "").strip() or None
    s = gsc_audit.create_session(label)
    return jsonify(s)


@app.route("/api/gsc/sessions/<sid>/launch", methods=["POST"])
def api_gsc_session_launch(sid):
    data = request.get_json(silent=True) or {}
    browser_name = data.get("browser", "edge")
    if sid in _session_drivers:
        try:
            _session_drivers[sid].current_url
            return jsonify({"status": "already_open", "session_id": sid})
        except Exception:
            _session_drivers.pop(sid, None)
    try:
        driver = gsc_audit.launch_session_browser(sid, browser_pref=browser_name)
        _session_drivers[sid] = driver
        return jsonify({"status": "launched", "session_id": sid,
                        "message": "Browser opened - log into your Google accounts, then click 'Done' when finished."})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/gsc/sessions/<sid>/done", methods=["POST"])
def api_gsc_session_done(sid):
    driver = _session_drivers.pop(sid, None)
    if not driver:
        return jsonify({"error": "No open browser for this session."}), 400
    try:
        result = gsc_audit.scan_session_cookies(sid, driver=driver)
        return jsonify({"status": "scanned", **result})
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        try:
            driver.quit()
        except Exception:
            pass


@app.route("/api/gsc/sessions/<sid>/remove", methods=["POST"])
def api_gsc_session_remove(sid):
    driver = _session_drivers.pop(sid, None)
    if driver:
        try:
            driver.quit()
        except Exception:
            pass
    gsc_audit.remove_session(sid)
    return jsonify({"status": "removed"})


@app.route("/api/gsc/sessions/<sid>/refresh", methods=["POST"])
def api_gsc_session_refresh(sid):
    """Re-scan cookies: launch headless, check myaccount page, detect emails."""
    data = request.get_json(silent=True) or {}
    browser_name = data.get("browser", "edge")
    driver = None
    profile_dir = os.path.join(gsc_audit._sessions_dir(), sid, "chrome_profile")
    # Share the per-session profile lock so this can't collide with a screenshot capture
    # that's using the same Google session's Chrome profile.
    lock = gsc_audit._profile_lock(profile_dir)
    if not lock.acquire(timeout=180):
        return jsonify({"error": "session_busy"}), 409
    try:
        import engine as _eng
        driver = _eng.build_driver(
            profile_dir, proxy=None, headless=True,
            country="us", extra_extensions=[],
            browser_pref=browser_name,
        )
        result = gsc_audit.scan_session_cookies(sid, driver=driver)
        return jsonify({"status": "refreshed", **result})
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass
        lock.release()


# --------------------------------------------------------------------------- #
# PWA routes
# --------------------------------------------------------------------------- #
@app.route("/manifest.webmanifest")
def manifest():
    return send_from_directory(STATIC_DIR, "manifest.webmanifest",
                               mimetype="application/manifest+json")

@app.route("/sw.js")
def service_worker():
    resp = send_from_directory(STATIC_DIR, "sw.js", mimetype="application/javascript")
    resp.headers["Service-Worker-Allowed"] = "/"
    resp.headers["Cache-Control"] = "no-cache"
    return resp

@app.route("/favicon.ico")
def favicon():
    return send_from_directory(os.path.join(STATIC_DIR, "icons"), "favicon.ico")

@app.route("/offline.html")
def offline():
    return send_from_directory(STATIC_DIR, "offline.html")

@app.errorhandler(500)
def err500(e):
    return jsonify({"error": "Internal server error", "detail": str(e)}), 500

# --------------------------------------------------------------------------- #
# Activity log
# --------------------------------------------------------------------------- #
_activity_log = []
_activity_lock = threading.Lock()
def activity(msg, level="info"):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    entry = {"ts": ts, "msg": msg, "level": level}
    with _activity_lock:
        _activity_log.append(entry)
        if len(_activity_log) > 500:
            _activity_log[:] = _activity_log[-500:]

@app.route("/api/activity-log")
def api_activity_log():
    with _activity_lock:
        return jsonify(list(_activity_log))

# --------------------------------------------------------------------------- #
# Auth - is_admin
# --------------------------------------------------------------------------- #
@app.route("/api/auth/is_admin")
def api_auth_is_admin():
    email = request.args.get("email", "").strip()
    accounts = auth.list_logged_in()
    if not accounts:
        return jsonify({"is_admin": False})
    if not email:
        email = accounts[0]["email"]
    result = auth._api_call({"action": "is_admin", "email": email})
    is_admin = result.get("is_admin", False)
    role = result.get("role", "")
    building = result.get("building", "")
    name = result.get("name", "")
    if is_admin:
        accts = auth._load_accounts()
        if email in accts:
            accts[email]["is_admin"] = True
            auth._save_accounts(accts)
    elif result.get("error"):
        accts = auth._load_accounts()
        is_admin = accts.get(email, {}).get("is_admin", False)
    if not name:
        name = auth._load_accounts().get(email, {}).get("name", "")
    return jsonify({"is_admin": is_admin, "role": role, "building": building, "email": email, "name": name})

# --------------------------------------------------------------------------- #
# Admin config (Apps Script URL, API key sync)
# --------------------------------------------------------------------------- #
SENSITIVE_KEYS = {"psi_api_key", "gemini_api_key", "imgbb_api_key", "gsc_client_id", "gsc_client_secret",
                  "auth_api_url", "user_auth_url", "gsc_projects"}

@app.route("/api/admin/save_config", methods=["POST"])
def api_admin_save_config():
    global CONFIG
    data = request.get_json(silent=True) or {}
    if "auth_api_url" in data:
        CONFIG["auth_api_url"] = data["auth_api_url"].strip()
    if "user_auth_url" in data:
        CONFIG["user_auth_url"] = data["user_auth_url"].strip()
    save_config(CONFIG)
    return jsonify({"saved": True})

@app.route("/api/admin/sync_keys", methods=["POST"])
def api_admin_sync_keys():
    """Manual trigger for the same central+per-building merge the background
    loop (_runtime_keys_sync_loop) already does automatically every
    RUNTIME_KEYS_TTL seconds - useful right after adding/changing a key in a
    sheet, so an admin doesn't have to wait for the next scheduled sync."""
    try:
        before = dict(CONFIG)
        _fetch_runtime_keys_now()
        changed_keys = {k: v for k, v in CONFIG.items() if k in SENSITIVE_KEYS and v and before.get(k) != v}
        masked_keys = {k: ("****" + v[-4:] if isinstance(v, str) and len(v) > 4 else "****")
                        for k, v in CONFIG.items() if k in SENSITIVE_KEYS and v}
        activity(f"API keys synced: {', '.join(masked_keys.keys())}")
        return jsonify({"ok": True, "keys": masked_keys, "changed": list(changed_keys.keys())})
    except Exception as e:
        activity(f"Key sync error: {e}", "error")
        return jsonify({"ok": False, "error": str(e)})

@app.route("/api/admin/keys_status")
def api_admin_keys_status():
    key_names = {"psi_api_key": "PageSpeed API Key", "gemini_api_key": "Gemini API Key",
                "imgbb_api_key": "ImgBB API Key (ranking screenshot URLs)"}
    status = {}
    for k, label in key_names.items():
        val = CONFIG.get(k, "")
        if val:
            status[label] = "****" + val[-4:] if len(val) > 4 else "****"
        else:
            status[label] = ""
    projects = CONFIG.get("gsc_projects", [])
    proj_display = []
    for p in projects:
        name = p.get("name", "?")
        cid = p.get("client_id", "")
        masked_id = "****" + cid[-4:] if len(cid) > 4 else "****" if cid else ""
        proj_display.append({"name": name, "client_id": masked_id,
                             "properties": p.get("properties", "")})
    return jsonify({"keys": status, "gsc_projects": proj_display})

# --------------------------------------------------------------------------- #
# Admin user management - proxies to central_gateway_apps_script.js (the live
# multi-tenant backend: a central "users" sheet for super admins + one users sheet
# per building, roles user/admin/superadmin). Every call is authorized as WHOEVER
# IS ALREADY LOGGED INTO THE EXE AS AN ADMIN, using their own email+password
# (already saved locally for periodic re-validation). No shared admin_key - a
# single bypass secret shared across every admin is a materially weaker model
# than per-account credentials, and this app's admin_key had already leaked once.
# --------------------------------------------------------------------------- #
def _admin_auth_params():
    """{'admin_email':..., 'admin_password':...} for the locally-logged-in admin,
    or None if no admin is logged in on this device."""
    email, password = auth.get_admin_credentials()
    if email:
        return {"admin_email": email, "admin_password": password}
    return None

def _admin_call(action, **params):
    auth_params = _admin_auth_params()
    if auth_params is None:
        return {"error": "You're not logged in as an approved admin on this device."}
    return auth._api_call({"action": action, **auth_params, **params})

@app.route("/api/admin/users")
def api_admin_users():
    return jsonify(_admin_call("admin_list"))

@app.route("/api/admin/buildings")
def api_admin_buildings():
    return jsonify(_admin_call("buildings_list"))

@app.route("/api/admin/proxies")
def api_admin_proxies():
    """Admin/superadmin-only viewer for the shared proxy pool. This is the ONLY
    place in the exe's UI that ever renders these passwords - regular users never
    reach this route (no menu entry, and the backend still gates it via admin
    credentials regardless)."""
    return jsonify(_admin_call("proxies_list"))

@app.route("/api/admin/save_proxy", methods=["POST"])
def api_admin_save_proxy():
    data = request.get_json(silent=True) or {}
    result = _admin_call(
        "proxies_upsert",
        type=(data.get("type") or "http").strip(),
        host=(data.get("host") or "").strip(),
        port=(data.get("port") or "").strip(),
        user=(data.get("user") or "").strip(),
        **{"pass": (data.get("pass") or "").strip()},
        region=(data.get("region") or "").strip(),
        active=data.get("active", True),
        notes=(data.get("notes") or "").strip(),
        orig_host=(data.get("orig_host") or "").strip(),
        orig_port=(data.get("orig_port") or "").strip(),
    )
    if result.get("success"):
        activity(f"Admin saved shared proxy: {data.get('host')}:{data.get('port')}")
        with _shared_proxies_lock:
            _shared_proxies_cache["ts"] = 0   # force a fresh fetch next use
    return jsonify(result)

@app.route("/api/admin/delete_proxy", methods=["POST"])
def api_admin_delete_proxy():
    data = request.get_json(silent=True) or {}
    result = _admin_call("proxies_delete", host=(data.get("host") or "").strip(),
                          port=(data.get("port") or "").strip())
    if result.get("success"):
        activity(f"Admin removed shared proxy: {data.get('host')}:{data.get('port')}")
        with _shared_proxies_lock:
            _shared_proxies_cache["ts"] = 0
    return jsonify(result)

@app.route("/api/admin/save_building", methods=["POST"])
def api_admin_save_building():
    data = request.get_json(silent=True) or {}
    result = _admin_call(
        "buildings_upsert",
        building=(data.get("building") or "").strip(),
        building_sheet_id=(data.get("building_sheet_id") or "").strip(),
        gsc_script_url=(data.get("gsc_script_url") or "").strip(),
    )
    if result.get("success"):
        activity(f"Admin saved building: {data.get('building')}")
    return jsonify(result)

@app.route("/api/admin/bulk_add_users", methods=["POST"])
def api_admin_bulk_add_users():
    """Add many users at once (one 'admin_upsert' call per row) - each row is
    {email, password, name, mac, role, building, notes}. Building-admin callers are
    always forced into their own building by the backend regardless of what's sent,
    same as the single-user path; a superadmin can target any building per row."""
    data = request.get_json(silent=True) or {}
    rows = data.get("rows") or []
    if not isinstance(rows, list) or not rows:
        return jsonify({"error": "No rows provided"})
    results = []
    for row in rows:
        email = (row.get("email") or "").strip()
        if not email:
            results.append({"email": "", "error": "Missing email"})
            continue
        upsert_params = {
            "email": email,
            "name": (row.get("name") or "").strip(),
            "role": (row.get("role") or "").strip(),
            "building": (row.get("building") or "").strip(),
            "approved": bool(row.get("approved", True)),
        }
        if row.get("password"):
            upsert_params["password"] = row["password"]
        if row.get("mac"):
            upsert_params["device"] = row["mac"]
        if row.get("notes"):
            upsert_params["notes"] = row["notes"]
        r = _admin_call("admin_upsert", **upsert_params)
        r["email"] = email
        results.append(r)
    ok_count = sum(1 for r in results if r.get("success"))
    activity(f"Admin bulk-added {ok_count}/{len(rows)} user(s)")
    return jsonify({"results": results, "ok_count": ok_count, "total": len(rows)})

@app.route("/api/admin/add_user", methods=["POST"])
def api_admin_add_user():
    """Create OR edit a user (admin_upsert decides by whether the email already
    exists) - matches the reference admin panel's single save-user action."""
    data = request.get_json(silent=True) or {}
    upsert_params = {
        "email": (data.get("email") or "").strip(),
        "name": (data.get("name") or "").strip(),
        "notes": (data.get("notes") or "").strip(),
        "device": (data.get("mac") or data.get("device") or "").strip(),
        "role": (data.get("role") or "").strip(),
        "building": (data.get("building") or "").strip(),
        "approved": bool(data.get("approved", True)),
    }
    if data.get("password"):
        upsert_params["password"] = data["password"]
    if data.get("clear_mac"):
        upsert_params["clear_mac"] = True
    if "formats" in data:
        upsert_params["formats"] = data.get("formats") or ""
    if "tools" in data:
        upsert_params["tools"] = data.get("tools") or ""
    result = _admin_call("admin_upsert", **upsert_params)
    if result.get("success"):
        activity(f"Admin {result.get('mode', 'saved')} user: {upsert_params['email']}")
    return jsonify(result)

@app.route("/api/admin/approve_user", methods=["POST"])
def api_admin_approve_user():
    data = request.get_json(silent=True) or {}
    email = (data.get("email") or "").strip()
    result = _admin_call("admin_approve", email=email)
    if result.get("success") or result.get("status") == "approved":
        activity(f"Admin approved user: {email}")
    return jsonify(result)

@app.route("/api/admin/reject_user", methods=["POST"])
def api_admin_reject_user():
    data = request.get_json(silent=True) or {}
    email = (data.get("email") or "").strip()
    result = _admin_call("admin_reject", email=email)
    if result.get("success") or result.get("status") == "rejected":
        activity(f"Admin rejected user: {email}")
    return jsonify(result)

@app.route("/api/admin/change_password", methods=["POST"])
def api_admin_change_password():
    """No separate admin_change_password action on the live backend - folded into
    admin_upsert (only the password field changes, every other field is preserved
    when omitted)."""
    data = request.get_json(silent=True) or {}
    email = (data.get("email") or "").strip()
    result = _admin_call("admin_upsert", email=email, password=data.get("new_password") or "")
    if result.get("success"):
        activity(f"Admin changed password for: {email}")
    return jsonify(result)

@app.route("/api/admin/update_mac", methods=["POST"])
def api_admin_update_mac():
    """Same fold-into-upsert approach as change_password above."""
    data = request.get_json(silent=True) or {}
    email = (data.get("email") or "").strip()
    mac = (data.get("mac") or "").strip()
    kwargs = {"clear_mac": True} if not mac else {"device": mac}
    result = _admin_call("admin_upsert", email=email, **kwargs)
    if result.get("success"):
        activity(f"Admin updated Device ID for: {email}")
    return jsonify(result)

@app.route("/api/admin/remove_user", methods=["POST"])
def api_admin_remove_user():
    data = request.get_json(silent=True) or {}
    email = (data.get("email") or "").strip()
    result = _admin_call("admin_delete", email=email)
    if result.get("success") or result.get("status") == "deleted":
        activity(f"Admin removed user: {email}")
    return jsonify(result)

@app.route("/api/gsc/projects")
def api_gsc_projects():
    projects = CONFIG.get("gsc_projects", [])
    return jsonify([{"name": p.get("name", ""), "properties": p.get("properties", "")}
                    for p in projects])

_gsc_account_props_cache = {}
GSC_ACCOUNT_PROPS_TTL = 300   # a connected account's property list rarely changes
                               # mid-session - avoids a live token refresh + API call
                               # per connected account on every domain checked in the UI


def _gsc_account_properties(email):
    cached = _gsc_account_props_cache.get(email)
    if cached and (time.time() - cached[0]) < GSC_ACCOUNT_PROPS_TTL:
        return cached[1]
    token = gsc_audit.get_access_token(email)
    props = gsc_audit.list_properties(token)
    _gsc_account_props_cache[email] = (time.time(), props)
    return props


@app.route("/api/gsc/check-for-domain")
def api_gsc_check_for_domain():
    """Check if GSC is connected locally for a domain; if not, check sheet mapping."""
    domain = request.args.get("domain", "").strip().lower().replace("www.", "")
    if not domain:
        return jsonify({"error": "No domain"})
    accounts = gsc_audit.list_accounts()
    connected = [a for a in accounts if a.get("has_refresh")]
    for acct in connected:
        try:
            props = _gsc_account_properties(acct["email"])
            for p in props:
                site = (p.get("siteUrl") or "").lower()
                if domain in site:
                    return jsonify({"connected": True, "email": acct["email"],
                                    "property": p.get("siteUrl"),
                                    "permission": p.get("permissionLevel", "")})
        except Exception:
            continue
    webapp_url = _gsc_webapp_url()
    if webapp_url:
        try:
            mapping = _gsc_mapping()
            info = mapping.get(domain)
            if info:
                return jsonify({"connected": False,
                    "available": True,
                    "email": info.get("email", ""),
                    "accountKey": info.get("accountKey", ""),
                    "accessLevel": info.get("accessLevel", ""),
                    "message": f"Domain found in GSC account {info.get('email', '')}. "
                               f"Connect this account to include GSC screenshots in reports."})
        except Exception:
            pass
    return jsonify({"connected": False, "available": False,
                    "message": "No GSC account found for this domain. Connect a Google account with GSC access to include screenshots."})


@app.route("/api/gsc/domain-check")
def api_gsc_domain_check():
    """Check which GSC account owns a domain via the cached Apps Script mapping."""
    domain = request.args.get("domain", "").strip().lower().replace("www.", "")
    if not domain:
        return jsonify({"error": "No domain"})
    webapp_url = _gsc_webapp_url()
    if not webapp_url:
        return jsonify({"found": False, "reason": "Apps Script URL not configured"})
    try:
        mapping = _gsc_mapping()
        info = mapping.get(domain)
        if info:
            return jsonify({
                "found": True,
                "domain": domain,
                "email": info.get("email", ""),
                "accountKey": info.get("accountKey", ""),
                "accessLevel": info.get("accessLevel", ""),
                "connected": True
            })
        return jsonify({"found": False, "domain": domain, "reason": "Domain not in GSC config"})
    except Exception as e:
        return jsonify({"found": False, "error": str(e)})

# --------------------------------------------------------------------------- #
# Crawl Tracker - proxy to Apps Script Web App
# --------------------------------------------------------------------------- #
def _crawl_apps_script_post(webapp_url, payload, timeout=180, retries=1):
    """POST to the Apps Script web app the same way the known-working
    last-gsc-crawl-date-check frontend does: Content-Type: text/plain
    (requests.post(json=...) sends application/json instead, which some
    doPost(e) handlers branch on) and a generous timeout - that frontend
    uses no timeout at all, since the URL Inspection API can genuinely take
    well over 30-60s per URL. One retry, since a slow/cold Apps Script
    execution is transient, not a hard failure."""
    import json as _json
    body = _json.dumps(payload).encode("utf-8")
    headers = {"Content-Type": "text/plain;charset=utf-8"}
    last_err = None
    for attempt in range(retries + 1):
        try:
            return http_requests.post(webapp_url, data=body, headers=headers, timeout=timeout)
        except http_requests.exceptions.Timeout as e:
            last_err = e
            continue
        except Exception as e:
            last_err = e
            break
    raise last_err

@app.route("/api/crawl/validate", methods=["POST"])
def api_crawl_validate():
    webapp_url = _gsc_webapp_url()
    if not webapp_url:
        return jsonify({"error": "GSC accounts not configured. Set Apps Script URL in Admin."}), 400
    data = request.get_json(silent=True) or {}
    try:
        resp = _crawl_apps_script_post(webapp_url, {"action": "validate_batch", "urls": data.get("urls", [])})
        return jsonify(resp.json())
    except http_requests.exceptions.Timeout:
        return jsonify({"error": "The Google Sheet is taking too long to respond. Please try again."}), 504
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/crawl/inspect", methods=["POST"])
def api_crawl_inspect():
    webapp_url = _gsc_webapp_url()
    if not webapp_url:
        return jsonify({"error": "GSC accounts not configured. Set Apps Script URL in Admin."}), 400
    data = request.get_json(silent=True) or {}
    try:
        resp = _crawl_apps_script_post(webapp_url, {
            "action": "inspect_single",
            "url": data.get("url", ""),
            "domain": data.get("domain", ""),
            "accountKey": data.get("accountKey", ""),
            "batchId": data.get("batchId", "")
        })
        return jsonify(resp.json())
    except http_requests.exceptions.Timeout:
        return jsonify({"error": "URL Inspection API is taking too long to respond. Please try again."}), 504
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# --------------------------------------------------------------------------- #
# Brief Analysis
# --------------------------------------------------------------------------- #
_brief_state = {"running": False, "progress": 0, "status": "idle", "result": None, "error": None, "error_msg": None, "stop": False, "log": []}
_brief_lock = threading.Lock()

@app.route("/api/brief/start", methods=["POST"])
def api_brief_start():
    with _brief_lock:
        if _brief_state["running"]:
            return jsonify({"error": "Brief analysis already running"}), 400
        _brief_state.update({"running": True, "progress": 0, "status": "starting", "result": None, "error": None, "error_msg": None, "stop": False, "log": []})
    data = request.get_json(silent=True) or {}
    # Brief analysis is domain-based - trim a pasted URL to the bare domain.
    domain = to_domain(data.get("domain") or "")
    if not domain:
        with _brief_lock:
            _brief_state["running"] = False
        return jsonify({"error": "Domain required"}), 400
    # Share the rank checker's Buster CAPTCHA solver + extension with the brief
    # report's indexing query, so the Google "site:" count can beat a CAPTCHA
    # (escalates to a visible browser) instead of coming back N/A.
    try:
        brief_analysis.configure_indexing(
            extensions=[BUSTER_DIR] if (CONFIG.get("use_buster", True) and os.path.isdir(BUSTER_DIR)) else None,
            solve_captcha=solve_with_buster)
    except Exception:
        pass
    # Honour the selected report format + optional target pages (one path per line).
    # Pass the format through as-is; the generator builds exactly it or errors.
    fmt = (data.get("format") or "james").strip().lower()
    _targets_raw = (data.get("targets") or "").strip()
    target_pages = [ln.strip() for ln in _targets_raw.splitlines() if ln.strip()] or None
    def _run():
        try:
            def prog(msg):
                with _brief_lock:
                    _brief_state["status"] = str(msg)
                    _brief_state["log"].append(str(msg))
            out_file = brief_analysis.run_brief_analysis(
                domain, fmt=fmt, target_pages=target_pages, log_fn=prog)
            with _brief_lock:
                _brief_state["result"] = {"file": out_file}
                _brief_state["status"] = "completed"
                _brief_state["progress"] = 100
        except Exception as e:
            import traceback
            tb = traceback.format_exc()
            print(f"[Brief Analysis ERROR] {tb}")
            with _brief_lock:
                _brief_state["error_msg"] = str(e)
                _brief_state["log"].append(f"ERROR: {e}")
                _brief_state["status"] = "error"
        finally:
            with _brief_lock:
                _brief_state["running"] = False
    t = threading.Thread(target=_run, daemon=True)
    t.start()
    activity(f"Brief analysis started for {domain}")
    return jsonify({"started": True})

@app.route("/api/brief/status")
def api_brief_status():
    with _brief_lock:
        return jsonify(dict(_brief_state))

@app.route("/api/brief/stop", methods=["POST"])
def api_brief_stop():
    with _brief_lock:
        _brief_state["stop"] = True
    return jsonify({"stopped": True})

@app.route("/api/brief/download")
def api_brief_download():
    with _brief_lock:
        result = _brief_state.get("result")
    if not result or not result.get("file"):
        return jsonify({"error": "No report available"}), 404
    fpath = result["file"]
    if not os.path.exists(fpath):
        return jsonify({"error": "Report file not found"}), 404
    return send_file(fpath, as_attachment=True, download_name=os.path.basename(fpath))

@app.route("/api/brief/formats")
def api_brief_formats():
    return jsonify([{"value": "pptx", "label": "PowerPoint (.pptx)"}])

# --------------------------------------------------------------------------- #
# On-Page parse targets
# --------------------------------------------------------------------------- #
@app.route("/api/onpage/parse-targets", methods=["POST"])
def api_onpage_parse_targets():
    # Excel upload (multipart) -> parse the sheet into "URL | kw1, kw2" lines.
    f = request.files.get("file")
    if f is not None and f.filename:
        try:
            lines = _lines_from_excel(f)
        except Exception as e:
            return jsonify({"error": f"Failed to parse Excel file: {e}"}), 400
        if not lines:
            return jsonify({"error": "No target pages found in the Excel file."}), 400
        return jsonify({"lines": lines, "count": len(lines)})
    # Text fallback (JSON body of pasted rows) -> same grouped lines.
    data = request.get_json(silent=True) or {}
    raw = data.get("urls", "") or data.get("targets", "")
    lines = _targets_to_lines(_parse_onpage_targets(raw))
    return jsonify({"lines": lines, "urls": lines, "count": len(lines)})

# --------------------------------------------------------------------------- #
# Entry point
# --------------------------------------------------------------------------- #
def _kill_previous():
    """Kill any previous GRC python processes on our port range."""
    import subprocess
    try:
        out = subprocess.check_output(
            'netstat -ano | findstr "507[0-9].*LISTENING"',
            shell=True, text=True, stderr=subprocess.DEVNULL)
        my_pid = os.getpid()
        killed = set()
        for line in out.strip().splitlines():
            parts = line.split()
            if len(parts) >= 5:
                pid = int(parts[-1])
                if pid != my_pid and pid not in killed:
                    try:
                        subprocess.call(f'taskkill /F /PID {pid}', shell=True,
                                        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                        killed.add(pid)
                    except Exception:
                        pass
        if killed:
            print(f"[GRC] Stopped previous instance(s): {killed}")
            import time as _t; _t.sleep(2)
    except Exception:
        pass

def main():
    # Skip the update check when launched by a start script (Start Tool.vbs / run.bat
    # already run the updater BEFORE the server). Re-checking here just adds a network
    # round-trip to startup, which on a slow connection pushes the port-file write past
    # the launcher's timeout -> "SEO Toolkit Pro could not start". GRC_NO_BROWSER is set
    # by the launcher, so it doubles as the "already updated" signal.
    if not os.environ.get("GRC_NO_BROWSER") and not os.environ.get("GRC_SKIP_UPDATE"):
        try:
            update_result = updater.check_and_update()
            if update_result.get("updated"):
                print(f"[GRC] Updated {len(update_result.get('updated_files', []))} file(s)")
        except Exception as e:
            print(f"[GRC] Update check skipped: {e}")

    port = int(os.environ.get("GRC_PORT", "5070"))
    url = f"http://127.0.0.1:{port}"

    _kill_previous()

    # Use fixed port since we killed previous instances
    import socket
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        s.bind(("127.0.0.1", port))
        s.close()
    except OSError:
        # Port still stuck, find a free one
        for try_port in range(port + 1, port + 20):
            try:
                s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 0)
                s.bind(("127.0.0.1", try_port))
                s.close()
                print(f"[GRC] Port {port} busy, using {try_port}")
                port = try_port
                break
            except OSError:
                continue
    url = f"http://127.0.0.1:{port}"

    # Write port file (used when launched via external script)
    port_file = os.environ.get("GRC_PORT_FILE", os.path.join(BUNDLE_DIR, ".grc_port"))
    with open(port_file, "w") as f:
        f.write(str(port))

    # Open browser unless suppressed by launcher
    if not os.environ.get("GRC_NO_BROWSER"):
        _open_app_window(url)

    # Warm the GSC domain-mapping cache in the background so GSC Audit / Health
    # Audit's domain lookup is instant by the time someone opens those tabs,
    # instead of a live Sheets read on first use.
    threading.Thread(target=_gsc_mapping_prefetch_loop, daemon=True).start()

    # Keep runtime API keys (ImgBB, PSI, Gemini, etc.) synced for every user
    # automatically - see _runtime_keys_sync_loop for the central+per-building merge.
    threading.Thread(target=_runtime_keys_sync_loop, daemon=True).start()

    try:
        from werkzeug.serving import make_server
        srv = make_server("127.0.0.1", port, app, threaded=True)
        srv.socket.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        srv.serve_forever()
    except OSError as e:
        pass


def _open_app_window(url):
    """Open the app in Edge app-mode window. Falls back to default browser."""
    import time as _t
    _t.sleep(1)
    edge_paths = [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    ]
    edge = next((p for p in edge_paths if os.path.exists(p)), None)
    if edge:
        profile = os.path.join(
            os.environ.get("LOCALAPPDATA", os.path.expanduser("~")),
            "SEO Toolkit Pro", "edge_app"
        )
        import glob as _glob
        for _p in _glob.glob(os.path.join(profile, "Default", "Favicons*")):
            try: os.remove(_p)
            except OSError: pass
        _icons_dir = os.path.join(profile, "Default", "Web Applications", "Manifest Resources")
        if os.path.isdir(_icons_dir):
            import shutil
            for _d in os.listdir(_icons_dir):
                _ic = os.path.join(_icons_dir, _d, "Icons")
                if os.path.isdir(_ic):
                    shutil.rmtree(_ic, ignore_errors=True)
        icon_path = os.path.join(BUNDLE_DIR, "rank-checker-search-bars.ico")
        cmd = [edge, f"--app={url}",
               f"--user-data-dir={profile}",
               "--window-size=1100,820"]
        if os.path.exists(icon_path):
            cmd.append(f"--app-icon={icon_path}")
        subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    else:
        import webbrowser
        webbrowser.open(url)

if __name__ == "__main__":
    try:
        main()
    except Exception as _e:
        import traceback, ctypes
        _log = os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")),
                            "SEO Toolkit Pro", "crash.log")
        os.makedirs(os.path.dirname(_log), exist_ok=True)
        with open(_log, "w") as _f:
            traceback.print_exc(file=_f)
        ctypes.windll.user32.MessageBoxW(
            0,
            f"SEO Toolkit Pro crashed.\n\nDetails saved to:\n{_log}\n\nSend this file to support.",
            "SEO Toolkit Pro - Error", 0x10
        )
