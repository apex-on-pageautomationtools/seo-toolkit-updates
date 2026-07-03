"""
SEO Toolkit Pro v3.2 — Merged edition
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

APP_VERSION = "3.2"

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

PROFILE_POOL_SIZE = 5
PROFILE_MAX_AGE_H = 72

for d in (UPLOADS_DIR, DOWNLOADS_DIR, SCREENSHOTS_DIR, PROFILES_DIR):
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

# For backward compat — PROFILE_DIR points to a random profile each launch
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

# --------------------------------------------------------------------------- #
# Global run state
# --------------------------------------------------------------------------- #
state = {
    "status": "idle", "current_keyword": "", "current_index": 0, "total": 0,
    "results": [], "captcha_msg": "", "error_msg": "", "log": [],
    "driver": None, "mode": "ranking", "domain": "",
}
pause_event = threading.Event(); pause_event.set()
stop_event  = threading.Event()
state_lock  = threading.Lock()

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
def autosave():
    try:
        with state_lock:
            data = {"mode": state.get("mode", "ranking"),
                    "results": list(state["results"]),
                    "saved_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
        with open(AUTOSAVE_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
    except Exception:
        pass

def load_autosave():
    try:
        if os.path.exists(AUTOSAVE_FILE):
            with open(AUTOSAVE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return None

def clear_autosave():
    try:
        if os.path.exists(AUTOSAVE_FILE):
            os.remove(AUTOSAVE_FILE)
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
    """Find and click Buster's solver button — it may be outside the iframe."""
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
        # Wait for Buster to inject its button in the bframe's help-button-holder
        # Buster uses a CLOSED shadow DOM — use CDP to pierce it
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
            add_log("Buster button not found — waiting for manual solve...")
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

        # Buster didn't solve — try reloading CAPTCHA and retrying (up to 3 times)
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
                # Reload button not found — wait and check for manual solve before giving up
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
            # Try Buster again on the new challenge
            clicked2 = False
            try:
                driver.switch_to.default_content()
                fr = driver.find_elements(By.CSS_SELECTOR,
                    "iframe[title*='recaptcha challenge'], iframe[src*='bframe']")
                if fr:
                    driver.switch_to.frame(fr[0])
                    host = driver.find_elements(By.CSS_SELECTOR, "div.help-button-holder")
                    if host:
                        from selenium.webdriver.common.action_chains import ActionChains
                        ActionChains(driver).move_to_element(host[0]).click().perform()
                        add_log("Buster clicked on reloaded CAPTCHA")
                        clicked2 = True
                    driver.switch_to.default_content()
            except Exception:
                try: driver.switch_to.default_content()
                except Exception: pass
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

    add_log("Buster could not solve.")
    return False

# --------------------------------------------------------------------------- #
# Session holder (driver + identity rotation)
# --------------------------------------------------------------------------- #
class Session:
    def __init__(self, headless, country, proxy_pool, browser_pref="auto", vpn_method="none",
                 latitude=None, longitude=None, lang="en"):
        self.headless = headless
        self.country = country
        self.pool = proxy_pool
        self.browser_pref = browser_pref
        self.vpn_method = vpn_method
        self.latitude = latitude
        self.longitude = longitude
        self.lang = lang
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
        # ProtonVPN, Windscribe app, etc.) — no extension needed. Loading the Urban VPN
        # browser extension causes a registration error because it gets a different ID
        # when loaded as unpacked vs installed from the Chrome Web Store.
        if os.path.isdir(SERPCOUNTER_DIR):
            exts.append(SERPCOUNTER_DIR)
            add_log("SERP Counter loaded — results will show position numbers")
        return exts

    def start(self, rotate=False):
        self.quit()
        self.profile = pick_profile()
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

        # Buster failed — pause and auto-detect when solved (no Resume button needed)
        bring_browser_to_front()
        return _captcha_manual_wait(sess.driver)

    # When using a VPN (system/pause), the current VPN IP is likely hard-blocked by Google.
    # Cooling down on the same IP won't help — ask the user to switch VPN server instead.
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

    # All retries failed — last resort manual pause
    if CONFIG.get("manual_fallback", True) and not sess.headless:
        bring_browser_to_front()
        return _manual_pause("Automatic recovery failed")
    return False

def _captcha_manual_wait(driver):
    """Show manual CAPTCHA prompt and auto-resume once the CAPTCHA clears —
    no Resume button click needed. Also resumes if user clicks Resume manually."""
    msg = ("CAPTCHA detected — solve it in the browser window. "
           "The tool will continue automatically once solved.")
    add_log(msg)
    with state_lock:
        state["captcha_msg"] = msg
        state["status"] = "paused"
    pause_event.clear()

    # Poll every 2 seconds — resume as soon as CAPTCHA disappears from the page
    while not stop_event.is_set() and not pause_event.is_set():
        time.sleep(2)
        try:
            src = page_source(driver)
            kind = classify_page(src)
            if kind in ("ok", "empty", "consent"):
                add_log("CAPTCHA solved — resuming automatically.")
                break
        except Exception:
            pass

    with state_lock:
        state["captcha_msg"] = ""
        if not stop_event.is_set():
            state["status"] = "running"
    return not stop_event.is_set()


def _manual_pause(reason):
    add_log(f"{reason}. Pausing for manual solve — solve in the browser, then Resume.")
    bring_browser_to_front()
    with state_lock:
        state["captcha_msg"] = (f"{reason}. The browser window has been brought to the "
                                f"front — solve the CAPTCHA there, then click Resume.")
        state["status"] = "paused"
    pause_event.clear()
    pause_event.wait()
    with state_lock:
        state["captcha_msg"] = ""
        if not stop_event.is_set():
            state["status"] = "running"
    return not stop_event.is_set()

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


# --------------------------------------------------------------------------- #
# Keyword ranking — hardened
# --------------------------------------------------------------------------- #
def rank_one(sess, keyword, domain, country, max_pages, search_mode="stop_on_found", city=None, lang="en"):
    domain_clean = clean_domain(domain)
    target = max_pages * 10

    for _try in range(CONFIG.get("max_block_retries", 3) + 1):
        if stop_event.is_set():
            return {"status": "stopped", "matches": []}
        pause_event.wait()
        if not is_alive(sess.driver):
            add_log("Browser lost — restarting for retry...")
            try:
                sess.start(rotate=bool(sess.pool))
            except Exception:
                raise BrowserClosedError("Browser closed and could not restart")

        human_search(sess.driver, keyword, country, add_log, city=city, lang=lang)

        src = page_source(sess.driver)
        kind = classify_page(src)

        if kind == "consent":
            add_log("Consent page detected — accepting...")
            engine.accept_consent(sess.driver, add_log)
            engine.accept_google_consent(sess.driver, add_log)
            human_search(sess.driver, keyword, country, add_log, city=city, lang=lang)
            src = page_source(sess.driver)
            kind = classify_page(src)

        if kind in ("captcha", "soft_block", "http_403"):
            add_log(f"Block detected ({kind}). Starting recovery...")
            if not _recover(sess, kind):
                return {"status": "captcha", "matches": match_domain([], domain_clean)}
            continue

        # Small wait then log what's on page 1
        time.sleep(1.0)
        links_page1, dbg = extract_organic(sess.driver, debug=True)

        # If no links found, might be consent or empty — try accepting consent and retry
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
            add_log(f"Page 1: 0 links — URL={dbg.get('url','?')[:80]} "
                    f"h3={dbg.get('h3count','?')} jsname={dbg.get('jsname_count','?')} "
                    f"zReHs={dbg.get('zReHs_count','?')} rso={dbg.get('rso','?')}")

        # Ctrl+F search on page 1
        all_matches = find_domain_in_page(sess.driver, domain_clean, page_offset=0)

        # Paginate through ALL selected pages (respects max_pages regardless of search_mode)
        page_num = 1
        total_links = len(links_page1)
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
                        add_log(f"No 'Next' button on page {page_num} — stopping")
                        break
                    nxt[0].click()
                    human_pause(1.5, 2.5)
                    page_num += 1
                    page_links = extract_organic(sess.driver)
                    if not page_links and page_num > 1:
                        human_pause(2.0, 3.5)
                        page_links = extract_organic(sess.driver)
                    total_links += len(page_links)
                    add_log(f"Page {page_num}: {len(page_links)} links")
                    offset = (page_num - 1) * 10
                    page_matches = find_domain_in_page(sess.driver, domain_clean, page_offset=offset)
                    if page_matches:
                        all_matches.extend(page_matches)
                        add_log(f"  Found at position {page_matches[0]['position']}")
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
            # Capture full-page SERP screenshot
            try:
                safe_kw = "".join(c if c.isalnum() or c in (' ', '-', '_') else '_' for c in keyword)[:60]
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                ss_name = f"{safe_kw}_{ts}.png"
                ss_folder = _domain_folder(domain, "ranking")
                ss_path = os.path.join(ss_folder, ss_name)
                if is_alive(sess.driver):
                    _save_full_page_screenshot(sess.driver, ss_path)
                    add_log(f"SERP screenshot saved: {ss_name}")
                    for m in matches:
                        m["screenshot"] = ss_name
            except Exception as e:
                add_log(f"Screenshot failed: {e}")
            # Visit each matched URL to capture final loaded URL (redirects)
            for m in matches:
                try:
                    if is_alive(sess.driver):
                        sess.driver.get(m["serp_url"])
                        time.sleep(random.uniform(2, 4))
                        m["loaded_url"] = sess.driver.current_url
                        add_log(f"Loaded URL: {m['loaded_url']}")
                except Exception:
                    m["loaded_url"] = ""
            # Navigate back to Google for next keyword
            try:
                if is_alive(sess.driver):
                    sess.driver.back()
                    time.sleep(1)
            except Exception:
                pass
            add_log(f"'{keyword}': found at #{matches[0]['position']} "
                    f"({total_links} results across {page_num} pages)")
            return {"status": "found", "matches": matches, "pages": page_num}
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
    sess = Session(headless, country, ProxyPool(proxies), browser_pref, vpn_method,
                   latitude, longitude, lang=lang)
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

        for i, kw in enumerate(keywords):
            if stop_event.is_set():
                break
            pause_event.wait()
            if stop_event.is_set():
                break
            if not is_alive(sess.driver):
                raise BrowserClosedError("Browser closed between keywords")

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
                add_log("Browser disconnected — restarting...")
                try:
                    sess.start(rotate=bool(sess.pool))
                    add_log("Browser restarted successfully.")
                except Exception as re_err:
                    raise BrowserClosedError(f"Could not restart browser: {re_err}")

            if not target_pages:
                target_pages = {}
            kw_target = target_pages.get(kw, "")
            kw_domain = domain
            if kw_target:
                from urllib.parse import urlparse as _up2
                _parsed = _up2(kw_target if "://" in kw_target else "https://" + kw_target)
                kw_domain = _parsed.netloc.replace("www.", "") or domain

            result = rank_one(sess, kw, kw_domain, country, max_pages, search_mode, city=city, lang=lang)

            with state_lock:
                matches = result.get("matches", [])
                if matches:
                    row = {"keyword": kw, "status": "found"}
                    if kw_target:
                        row["target_page"] = kw_target
                    for idx, m in enumerate(matches):
                        suffix = "" if idx == 0 else f"_{idx+1}"
                        row[f"position{suffix}"] = m["position"]
                        row[f"serp_url{suffix}"] = m["serp_url"]
                        row[f"loaded_url{suffix}"] = m.get("loaded_url", "")
                    if matches[0].get("screenshot"):
                        row["screenshot"] = matches[0]["screenshot"]
                    state["results"].append(row)
                else:
                    row = {"keyword": kw, "position": "-",
                           "serp_url": "", "loaded_url": "",
                           "status": result.get("status", "not_found")}
                    if kw_target:
                        row["target_page"] = kw_target
                    state["results"].append(row)
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
            state["error_msg"] = "Browser was closed. Results saved — Export CSV to download."
    except Exception as e:
        import traceback; traceback.print_exc()
        add_log(f"Fatal error: {e}")
        autosave()
        with state_lock:
            state["status"] = "error"; state["error_msg"] = str(e)
    finally:
        sess.quit()

# --------------------------------------------------------------------------- #
# Index checker — hardened
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
    sess = Session(headless, country, ProxyPool(proxies), browser_pref, vpn_method,
                   latitude, longitude, lang=lang)
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
                add_log("Browser disconnected — restarting...")
                try:
                    sess.start(rotate=bool(sess.pool))
                    add_log("Browser restarted successfully.")
                except Exception as re_err:
                    raise BrowserClosedError(f"Could not restart browser: {re_err}")
            with state_lock:
                state["current_keyword"] = url
                state["current_index"] = i + 1
                state["status"] = "running"
            add_log(f"Index check '{url}' ({i+1}/{len(urls)})")
            result = index_one(sess, url, country, city=city, lang=lang)
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
            state["error_msg"] = "Browser closed. Results saved — Export CSV to download."
    except Exception as e:
        import traceback; traceback.print_exc()
        add_log(f"Fatal error: {e}")
        autosave()
        with state_lock:
            state["status"] = "error"; state["error_msg"] = str(e)
    finally:
        sess.quit()

# --------------------------------------------------------------------------- #
# Backlink checker — visits each backlink URL, finds domain link, checks meta
# --------------------------------------------------------------------------- #
def backlink_one(sess, backlink_url, target_domain, check_da=True):
    from urllib.parse import urlparse
    import re as _re
    url = backlink_url.strip()
    if not url.startswith("http"):
        url = "https://" + url
    target_clean = target_domain.lower().replace("www.", "").strip("/")

    try:
        safe_get(sess.driver, url)
        human_pause(2, 4)
    except Exception as e:
        return {"status": "error", "domain_found": "Error", "meta_robots": "—",
                "link_type": "—", "link_url": "", "error": str(e)}

    if not is_alive(sess.driver):
        return {"status": "error", "domain_found": "Error", "meta_robots": "—",
                "link_type": "—", "link_url": ""}

    # Capture final URL after any redirects
    try:
        final_url = sess.driver.current_url
    except Exception:
        final_url = url

    src = page_source(sess.driver)

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

    # Find target domain links
    domain_found = "No"
    link_url = ""
    link_type = "—"
    try:
        from selenium.webdriver.common.by import By
        anchors = sess.driver.find_elements(By.TAG_NAME, "a")
        for a in anchors:
            try:
                href = (a.get_attribute("href") or "").strip()
                if not href or href.startswith("javascript:") or href == "#":
                    continue
                parsed = urlparse(href)
                link_domain = parsed.netloc.lower().replace("www.", "")
                if target_clean in link_domain or link_domain in target_clean:
                    domain_found = "Yes"
                    link_url = href
                    rel = (a.get_attribute("rel") or "").lower()
                    link_type = "nofollow" if "nofollow" in rel else "dofollow"
                    break
            except Exception:
                continue
    except Exception:
        pass

    da_val, pa_val, da_src = "—", "—", "—"
    if check_da:
        bl_domain = urlparse(url).netloc.replace("www.", "").lower()
        da_result = check_da_pa(sess.driver, bl_domain, log_fn=add_log)
        da_val = da_result.get("da", "—")
        pa_val = da_result.get("pa", "—")
        da_src = da_result.get("source", "—")

    return {"status": "ok", "domain_found": domain_found, "meta_robots": meta_robots,
            "link_type": link_type, "link_url": link_url, "final_url": final_url,
            "da": da_val, "pa": pa_val, "da_source": da_src}


def run_backlink_analysis(urls, domain, delay, headless, country, proxies,
                          browser_pref="auto", vpn_method="none", check_da=True,
                          latitude=None, longitude=None):
    sess = Session(headless, country, ProxyPool(proxies), browser_pref, vpn_method,
                   latitude, longitude)
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
                add_log("Browser disconnected — restarting...")
                try:
                    sess.start(rotate=bool(sess.pool))
                    add_log("Browser restarted successfully.")
                except Exception as re_err:
                    raise BrowserClosedError(f"Could not restart browser: {re_err}")
            with state_lock:
                state["current_keyword"] = url
                state["current_index"] = i + 1
                state["status"] = "running"
            add_log(f"Checking backlink '{url}' ({i+1}/{len(urls)})")
            result = backlink_one(sess, url, domain, check_da=check_da)
            add_log(f"  domain {'found' if result['domain_found']=='Yes' else 'not found'}, "
                    f"meta: {result['meta_robots']}, link: {result['link_type']}")
            with state_lock:
                state["results"].append({
                    "url": url, "domain_found": result.get("domain_found", "No"),
                    "meta_robots": result.get("meta_robots", "—"),
                    "link_type": result.get("link_type", "—"),
                    "link_url": result.get("link_url", ""),
                    "da": result.get("da", "—"), "pa": result.get("pa", "—"),
                    "da_source": result.get("da_source", "—"),
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
            state["error_msg"] = "Browser closed. Results saved — Export CSV to download."
    except Exception as e:
        import traceback; traceback.print_exc()
        add_log(f"Fatal error: {e}")
        autosave()
        with state_lock:
            state["status"] = "error"; state["error_msg"] = str(e)
    finally:
        sess.quit()

# --------------------------------------------------------------------------- #
# Helpers for request parsing
# --------------------------------------------------------------------------- #
def _proxies_from_request(data):
    proxies = list(CONFIG.get("proxies", []))
    ph = (data.get("proxy_host") or "").strip()
    pp = (data.get("proxy_port") or "").strip()
    if ph and pp:
        proxies.insert(0, {
            "type": data.get("proxy_type", "http"), "host": ph, "port": pp,
            "user": (data.get("proxy_user") or "").strip(),
            "pass": (data.get("proxy_pass") or "").strip()})
    return proxies

# --------------------------------------------------------------------------- #
# Flask routes — Auth
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

@app.route("/api/auth/mac")
def api_auth_mac():
    """Return MAC address for display to user (token for admin)."""
    return jsonify({"mac": auth.get_mac_address()})

@app.route("/api/auth/formats")
def api_auth_formats():
    """Return allowed report formats for the current user."""
    fmts = auth.get_allowed_formats()
    return jsonify({"formats": fmts})

# --------------------------------------------------------------------------- #
# Flask routes — API
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
    with state_lock:
        if state["status"] in ("running", "paused", "starting"):
            return jsonify({"error": "Already running. Stop or wait first."}), 400

    data = request.get_json(silent=True) or {}
    mode = data.get("mode", "ranking")
    domain = (data.get("domain") or "").strip()
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
    proxies = _proxies_from_request(data)

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

    with state_lock:
        state.update({"status": "starting", "current_keyword": "", "current_index": 0,
                      "total": len(keywords), "results": [], "captcha_msg": "",
                      "error_msg": "", "mode": mode, "log": [], "domain": domain})
    pause_event.set(); stop_event.clear(); clear_autosave()

    if city and latitude is not None:
        add_log(f"City: {city} — Geolocation sensor: {latitude}, {longitude}")

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
    else:
        t = threading.Thread(target=run_index_analysis,
                             args=(keywords, delay, headless, country, proxies,
                                   browser_pref, vpn_method,
                                   latitude, longitude, city, lang),
                             daemon=True)
    t.start()
    activity(f"{mode.title()} started — {domain or 'multi-target'} ({len(keywords)} keywords)")
    return jsonify({"status": "started", "total": len(keywords), "mode": mode})

@app.route("/api/pause", methods=["POST"])
def api_pause():
    with state_lock:
        if state["status"] == "running":
            state["status"] = "paused"
    pause_event.clear()
    return jsonify({"status": "paused"})

@app.route("/api/resume", methods=["POST"])
def api_resume():
    with state_lock:
        if state["status"] == "paused":
            state["status"] = "running"; state["captcha_msg"] = ""
    pause_event.set()
    return jsonify({"status": "running"})

@app.route("/api/stop", methods=["POST"])
def api_stop():
    stop_event.set(); pause_event.set(); autosave()
    with state_lock:
        state["status"] = "stopped"
    return jsonify({"status": "stopped"})

@app.route("/api/reset", methods=["POST"])
def api_reset():
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
# Admin — City Management
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
    Canonical is optional — when omitted it is generated as "City,Country".
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
    with state_lock:
        snap = dict(state)
    snap.pop("driver", None)
    return jsonify(snap)

@app.route("/api/load-autosave")
def api_load_autosave():
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
    with state_lock:
        results = list(state["results"]); m = state.get("mode", "ranking")
        domain = state.get("domain", "")
    out = io.StringIO()
    if m == "ranking":
        # Collect all column names across all rows (handles multiple matches)
        has_targets = any(r.get("target_page") for r in results)
        base = ["keyword", "target_page", "status", "position", "serp_url", "loaded_url"] if has_targets else ["keyword", "status", "position", "serp_url", "loaded_url"]
        extra = set()
        for r in results:
            for k in r:
                if k not in base and k not in extra:
                    extra.add(k)
        # Sort extra columns: position_2, serp_url_2, loaded_url_2, position_3, ...
        def _col_sort(c):
            for i, prefix in enumerate(["position_", "serp_url_", "loaded_url_"]):
                if c.startswith(prefix):
                    num = c[len(prefix):]
                    return (int(num) if num.isdigit() else 99, i)
            return (99, 99)
        fields = base + sorted(extra, key=_col_sort)
    elif m == "backlink":
        fields = ["url", "domain_found", "meta_robots", "link_type", "da", "pa", "da_source", "link_url", "status", "checked_at"]
    else:
        fields = ["url", "indexed", "found_url", "status", "checked_at"]
    w = csv.DictWriter(out, fieldnames=fields, extrasaction="ignore")
    w.writeheader(); w.writerows(results); out.seek(0)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    if m == "ranking":
        name = f"rankings_{ts}.csv"
    elif m == "backlink":
        name = f"backlink_check_{ts}.csv"
    else:
        name = f"index_check_{ts}.csv"
    # Save to domain folder
    csv_data = out.getvalue()
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

# --------------------------------------------------------------------------- #
# SEO On-Page Report — runs phase2 script as subprocess
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

def _run_onpage_report(domain, targets_json, fmt, no_capture):
    """Run the on-page phase2 script as a subprocess, stream logs."""
    with onpage_lock:
        onpage_state.update({"status": "running", "log": [], "domain": domain,
                             "output_zip": "", "error_msg": "", "progress": "Starting..."})
    onpage_stop.clear()

    out_dir = os.path.join(_data_dir(), "onpage_output")
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

    def _log(msg):
        with onpage_lock:
            onpage_state["log"].append(msg)
            if msg.startswith("["):
                onpage_state["progress"] = msg

    _log(f"Running: {domain} (format={fmt})")
    _log(f"Output dir: {out_dir}")

    try:
        proc = subprocess.Popen(
            cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, bufsize=1, cwd=SCRIPTS_DIR,
            env={**os.environ, "PYTHONDONTWRITEBYTECODE": "1"}
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
            with onpage_lock:
                onpage_state["status"] = "completed"
                onpage_state["output_zip"] = zips[0]
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
            targets = json.loads(targets_raw)
        except json.JSONDecodeError:
            lines = [l.strip() for l in targets_raw.splitlines() if l.strip()]
            targets = [{"page": l, "keywords": []} for l in lines]

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
    if not zip_path or not os.path.exists(zip_path):
        return jsonify({"error": "No report available for download."}), 404
    return send_file(zip_path, as_attachment=True,
                     download_name=os.path.basename(zip_path))


# --------------------------------------------------------------------------- #
# Health Audit — runs checks + optional Selenium screenshots, builds report
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
    domain = (data.get("domain") or "").strip()
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
    activity(f"Health audit started — {domain} ({fmt})")
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
    driver = None

    try:
        _log("[1/3] Launching GSC browser...")
        driver = engine.build_driver(
            GSC_PROFILE_DIR, proxy=None, headless=headless,
            country="us", extra_extensions=[],
            logger=_log, browser_pref=browser_name,
        )

        _log("[2/3] Running GSC audit...")
        path = gsc_audit.run_gsc_audit(
            domain, email, fmt=fmt, out_dir=out_folder,
            driver=driver, log_fn=_log,
        )

        with gsc_lock:
            gsc_state["status"] = "completed"
            gsc_state["output_file"] = path
            gsc_state["progress"] = "Report ready for download"
        _log(f"[3/3] Report saved: {os.path.basename(path)}")

    except Exception as e:
        _log(f"Error: {e}")
        with gsc_lock:
            gsc_state["status"] = "error"
            gsc_state["error_msg"] = str(e)
    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass


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

    driver = None
    try:
        driver = engine.build_driver(
            GSC_PROFILE_DIR, proxy=None, headless=headless,
            country="us", extra_extensions=[],
            browser_pref=browser_name,
        )
        email = gsc_audit.oauth_login_selenium(driver, client_id, client_secret)
        return jsonify({"status": "connected", "email": email})
    except Exception as e:
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
    domain = (data.get("domain") or "").strip()
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
                        "message": "Browser opened — log into your Google accounts, then click 'Done' when finished."})
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
    try:
        import engine as _eng
        profile_dir = os.path.join(gsc_audit._sessions_dir(), sid, "chrome_profile")
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
# Auth — is_admin
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
    if is_admin:
        accts = auth._load_accounts()
        if email in accts:
            accts[email]["is_admin"] = True
            auth._save_accounts(accts)
    elif result.get("error"):
        accts = auth._load_accounts()
        is_admin = accts.get(email, {}).get("is_admin", False)
    return jsonify({"is_admin": is_admin})

# --------------------------------------------------------------------------- #
# Admin config (Apps Script URL, admin key, API key sync)
# --------------------------------------------------------------------------- #
SENSITIVE_KEYS = {"psi_api_key", "gsc_client_id", "gsc_client_secret",
                  "admin_key", "auth_api_url", "user_auth_url", "gsc_projects"}

@app.route("/api/admin/save_config", methods=["POST"])
def api_admin_save_config():
    global CONFIG
    data = request.get_json(silent=True) or {}
    if "auth_api_url" in data:
        CONFIG["auth_api_url"] = data["auth_api_url"].strip()
    if "user_auth_url" in data:
        CONFIG["user_auth_url"] = data["user_auth_url"].strip()
    if "admin_key" in data:
        CONFIG["admin_key"] = data["admin_key"].strip()
    save_config(CONFIG)
    return jsonify({"saved": True})

@app.route("/api/admin/sync_keys", methods=["POST"])
def api_admin_sync_keys():
    global CONFIG
    webapp_url = CONFIG.get("auth_api_url", "").strip()
    if not webapp_url:
        return jsonify({"ok": False, "error": "Apps Script URL not configured"})
    try:
        resp = http_requests.post(webapp_url, json={"action": "get_keys"}, timeout=20)
        data = resp.json()
        if not data.get("success"):
            return jsonify({"ok": False, "error": data.get("error", "Sheet returned error")})
        keys = data.get("keys", {})
        for k, v in keys.items():
            if v and isinstance(v, str):
                CONFIG[k] = v.strip()
        projects = data.get("gsc_projects", [])
        if projects:
            CONFIG["gsc_projects"] = projects
        save_config(CONFIG)
        masked_keys = {k: ("****" + v[-4:] if len(v) > 4 else "****") for k, v in keys.items() if v}
        activity(f"API keys synced: {', '.join(masked_keys.keys())}; {len(projects)} GSC project(s)")
        return jsonify({"ok": True, "keys": masked_keys, "gsc_projects_count": len(projects)})
    except Exception as e:
        activity(f"Key sync error: {e}", "error")
        return jsonify({"ok": False, "error": str(e)})

@app.route("/api/admin/keys_status")
def api_admin_keys_status():
    key_names = {"psi_api_key": "PageSpeed API Key"}
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

@app.route("/api/gsc/projects")
def api_gsc_projects():
    projects = CONFIG.get("gsc_projects", [])
    return jsonify([{"name": p.get("name", ""), "properties": p.get("properties", "")}
                    for p in projects])

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
            token = gsc_audit.get_access_token(acct["email"])
            props = gsc_audit.list_properties(token)
            for p in props:
                site = (p.get("siteUrl") or "").lower()
                if domain in site:
                    return jsonify({"connected": True, "email": acct["email"],
                                    "property": p.get("siteUrl"),
                                    "permission": p.get("permissionLevel", "")})
        except Exception:
            continue
    webapp_url = CONFIG.get("auth_api_url", "").strip()
    if webapp_url:
        try:
            resp = http_requests.get(webapp_url, params={"action": "get_config"}, timeout=15)
            data = resp.json()
            mapping = data.get("mapping", {}) if data.get("success") else {}
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
    """Check which GSC account owns a domain via Apps Script config."""
    domain = request.args.get("domain", "").strip().lower().replace("www.", "")
    if not domain:
        return jsonify({"error": "No domain"})
    webapp_url = CONFIG.get("auth_api_url", "").strip()
    if not webapp_url:
        return jsonify({"found": False, "reason": "Apps Script URL not configured"})
    try:
        resp = http_requests.get(webapp_url, params={"action": "get_config"}, timeout=15)
        data = resp.json()
        if not data.get("success"):
            return jsonify({"found": False, "reason": data.get("error", "Sheet error")})
        mapping = data.get("mapping", {})
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
# Crawl Tracker — proxy to Apps Script Web App
# --------------------------------------------------------------------------- #
@app.route("/api/crawl/validate", methods=["POST"])
def api_crawl_validate():
    webapp_url = CONFIG.get("auth_api_url", "").strip()
    if not webapp_url:
        return jsonify({"error": "GSC accounts not configured. Set Apps Script URL in Admin."}), 400
    data = request.get_json(silent=True) or {}
    try:
        resp = http_requests.post(webapp_url, json={"action": "validate_batch", "urls": data.get("urls", [])}, timeout=30)
        return jsonify(resp.json())
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/crawl/inspect", methods=["POST"])
def api_crawl_inspect():
    webapp_url = CONFIG.get("auth_api_url", "").strip()
    if not webapp_url:
        return jsonify({"error": "GSC accounts not configured. Set Apps Script URL in Admin."}), 400
    data = request.get_json(silent=True) or {}
    try:
        resp = http_requests.post(webapp_url, json={
            "action": "inspect_single",
            "url": data.get("url", ""),
            "domain": data.get("domain", ""),
            "accountKey": data.get("accountKey", ""),
            "batchId": data.get("batchId", "")
        }, timeout=30)
        return jsonify(resp.json())
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
    domain = (data.get("domain") or "").strip()
    if not domain:
        with _brief_lock:
            _brief_state["running"] = False
        return jsonify({"error": "Domain required"}), 400
    def _run():
        try:
            def prog(msg):
                with _brief_lock:
                    _brief_state["status"] = str(msg)
                    _brief_state["log"].append(str(msg))
            out_file = brief_analysis.run_brief_analysis(domain, log_fn=prog)
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
    data = request.get_json(silent=True) or {}
    raw = data.get("urls", "")
    urls = [u.strip() for u in raw.replace(",", "\n").splitlines() if u.strip()]
    return jsonify({"urls": urls, "count": len(urls)})

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
    # OTA updates disabled — local customizations would be overwritten
    # To re-enable: uncomment the block below
    # try:
    #     update_result = updater.check_and_update()
    #     if update_result.get("updated"):
    #         print(f"[GRC] Updated {len(update_result.get('updated_files', []))} file(s)")
    # except Exception as e:
    #     print(f"[GRC] Update check skipped: {e}")

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
