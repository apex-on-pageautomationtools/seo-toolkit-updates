"""
DA/PA Checker - uses the Selenium browser to check DA/PA via free online tools.
Tries multiple tools in sequence, caches results per domain.
"""

import re
import json
import time
import random
import logging
import requests
from urllib.parse import urlparse

logger = logging.getLogger(__name__)

_da_cache = {}
_dr_cache = {}


def _extract_domain(url_or_domain):
    d = url_or_domain.strip()
    if not d.startswith("http"):
        d = "https://" + d
    parsed = urlparse(d)
    return parsed.netloc.replace("www.", "").lower()


_WPB_HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                              "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"}


def _try_wpblogging101(driver, domain):
    """wpblogging101.com's public DA/PA Checker tool (tools/da-pa-checker/). A plain
    HTTP flow - load the page for a fresh nonce, then submit via the page's own
    admin-ajax.php - the exact interaction a visitor performs; no bypassed access
    control, no browser needed, so it's both faster and less fragile than the
    Selenium scrapers below. Tried first since it returns clean JSON rather than a
    number that has to be regexed out of rendered HTML."""
    try:
        s = requests.Session()
        r = s.get("https://wpblogging101.com/tools/da-pa-checker/", headers=_WPB_HEADERS, timeout=15)
        # No generic _is_blocked() check here - the page always ships a hidden
        # (display:none) CAPTCHA modal container in its markup even on a normal
        # load, which false-positives that check; a captcha-gated state comes back
        # as an explicit "error" in the AJAX JSON response below instead, which IS
        # checked, so a real block still can't produce a fabricated DA/PA value.
        if r.status_code != 200:
            return None
        m = re.search(r"atools101DapaSecure\s*=\s*(\{[^}]*\})", r.text)
        if not m:
            return None
        cfg = json.loads(m.group(1))
        nonce, ajax = cfg.get("nonce"), cfg.get("ajax")
        if not nonce or not ajax:
            return None
        r2 = s.post(ajax, headers=_WPB_HEADERS, timeout=15,
                     data={"action": "atools101_dapa_check", "url": domain, "_ajax_nonce": nonce})
        if r2.status_code != 200:
            return None
        data = r2.json()
        if data.get("error") or str(data.get("domain_name", "")).lower() != domain.lower():
            return None
        da_raw, pa_raw = str(data.get("da", "")).strip(), str(data.get("pa", "")).strip()
        if not da_raw.isdigit():
            return None
        pa = int(pa_raw) if pa_raw.isdigit() else "N/A"
        return {"da": int(da_raw), "pa": pa, "source": "wpblogging101.com"}
    except Exception as e:
        logger.debug(f"wpblogging101 failed: {e}")
    return None


def _try_rhinorank(driver, domain):
    """rhinorank.io DA checker"""
    try:
        from selenium.webdriver.common.by import By
        from selenium.webdriver.common.keys import Keys
        driver.get("https://www.rhinorank.io/tools/da-dr-checker/")
        time.sleep(3)

        inp = driver.find_elements(By.CSS_SELECTOR, "input[type='text'], textarea")
        if not inp:
            return None
        inp[0].clear()
        inp[0].send_keys(domain)
        time.sleep(0.5)

        btns = driver.find_elements(By.CSS_SELECTOR, "button[type='submit'], input[type='submit'], button.btn-primary, button.submit")
        if not btns:
            btns = driver.find_elements(By.XPATH, "//button[contains(text(),'Check') or contains(text(),'Submit') or contains(text(),'Go')]")
        if btns:
            btns[0].click()
        else:
            inp[0].send_keys(Keys.RETURN)
        time.sleep(8)

        src = driver.page_source
        if _is_blocked(src):
            return None
        da = _find_metric(src, ["domain authority", "DA"], limit=100)
        pa = _find_metric(src, ["page authority", "PA"], limit=100)

        if da is not None:
            return {"da": da,
                    "pa": pa if pa is not None else "N/A",
                    "source": "rhinorank.io"}
    except Exception as e:
        logger.debug(f"rhinorank failed: {e}")
    return None


def _try_websiteseochecker(driver, domain):
    """websiteseochecker.com DA checker"""
    try:
        from selenium.webdriver.common.by import By
        from selenium.webdriver.common.keys import Keys
        driver.get("https://websiteseochecker.com/domain-authority-checker/")
        time.sleep(3)

        inp = driver.find_elements(By.CSS_SELECTOR, "textarea, input[name='domain'], input[type='text']")
        if not inp:
            return None
        inp[0].clear()
        inp[0].send_keys(domain)
        time.sleep(0.5)

        btns = driver.find_elements(By.CSS_SELECTOR, "button[type='submit'], input[type='submit']")
        if not btns:
            btns = driver.find_elements(By.XPATH, "//button[contains(text(),'Check') or contains(text(),'Submit')]")
        if btns:
            btns[0].click()
        else:
            inp[0].send_keys(Keys.RETURN)
        time.sleep(10)

        src = driver.page_source
        if _is_blocked(src):
            return None
        da = _find_metric(src, ["domain authority", "DA Score", "DA"], limit=100)
        pa = _find_metric(src, ["page authority", "PA Score", "PA"], limit=100)

        if da is not None:
            return {"da": da, "pa": pa if pa is not None else "N/A",
                    "source": "websiteseochecker.com"}
    except Exception as e:
        logger.debug(f"websiteseochecker failed: {e}")
    return None


def _try_dapa_checker(driver, domain):
    """dapa-checker.com"""
    try:
        from selenium.webdriver.common.by import By
        from selenium.webdriver.common.keys import Keys
        driver.get("https://dapa-checker.com/")
        time.sleep(3)

        inp = driver.find_elements(By.CSS_SELECTOR, "textarea, input[name='domain'], input[type='text'], input[type='url']")
        if not inp:
            return None
        inp[0].clear()
        inp[0].send_keys(domain)
        time.sleep(0.5)

        btns = driver.find_elements(By.CSS_SELECTOR, "button[type='submit'], input[type='submit']")
        if not btns:
            btns = driver.find_elements(By.XPATH, "//button[contains(text(),'Check') or contains(text(),'Submit')]")
        if btns:
            btns[0].click()
        else:
            inp[0].send_keys(Keys.RETURN)
        time.sleep(10)

        src = driver.page_source
        if _is_blocked(src):
            return None
        da = _find_metric(src, ["domain authority", "DA"], limit=100)
        pa = _find_metric(src, ["page authority", "PA"], limit=100)

        if da is not None:
            return {"da": da, "pa": pa if pa is not None else "N/A",
                    "source": "dapa-checker.com"}
    except Exception as e:
        logger.debug(f"dapa-checker failed: {e}")
    return None


def _try_da_checker_org(driver, domain):
    """da-checker.org"""
    try:
        from selenium.webdriver.common.by import By
        from selenium.webdriver.common.keys import Keys
        driver.get("https://www.da-checker.org/")
        time.sleep(3)

        inp = driver.find_elements(By.CSS_SELECTOR, "textarea, input[name='domain'], input[type='text']")
        if not inp:
            return None
        inp[0].clear()
        inp[0].send_keys(domain)
        time.sleep(0.5)

        btns = driver.find_elements(By.CSS_SELECTOR, "button[type='submit'], input[type='submit']")
        if not btns:
            btns = driver.find_elements(By.XPATH, "//button[contains(text(),'Check') or contains(text(),'Submit')]")
        if btns:
            btns[0].click()
        else:
            inp[0].send_keys(Keys.RETURN)
        time.sleep(10)

        src = driver.page_source
        if _is_blocked(src):
            return None
        da = _find_metric(src, ["domain authority", "DA"], limit=100)
        pa = _find_metric(src, ["page authority", "PA"], limit=100)

        if da is not None:
            return {"da": da, "pa": pa if pa is not None else "N/A",
                    "source": "da-checker.org"}
    except Exception as e:
        logger.debug(f"da-checker.org failed: {e}")
    return None


def _try_teqtop(driver, domain):
    """teqtop.com DA PA checker"""
    try:
        from selenium.webdriver.common.by import By
        from selenium.webdriver.common.keys import Keys
        driver.get("https://www.teqtop.com/da-pa-checker")
        time.sleep(3)

        inp = driver.find_elements(By.CSS_SELECTOR, "textarea, input[name='domain'], input[type='text'], input[type='url']")
        if not inp:
            return None
        inp[0].clear()
        inp[0].send_keys(domain)
        time.sleep(0.5)

        btns = driver.find_elements(By.CSS_SELECTOR, "button[type='submit'], input[type='submit']")
        if not btns:
            btns = driver.find_elements(By.XPATH, "//button[contains(text(),'Check') or contains(text(),'Submit')]")
        if btns:
            btns[0].click()
        else:
            inp[0].send_keys(Keys.RETURN)
        time.sleep(10)

        src = driver.page_source
        if _is_blocked(src):
            return None
        da = _find_metric(src, ["domain authority", "DA"], limit=100)
        pa = _find_metric(src, ["page authority", "PA"], limit=100)

        if da is not None:
            return {"da": da, "pa": pa if pa is not None else "N/A",
                    "source": "teqtop.com"}
    except Exception as e:
        logger.debug(f"teqtop failed: {e}")
    return None


_BLOCK_MARKERS = (
    "captcha", "recaptcha", "cf-turnstile", "cf-challenge", "verify you are human",
    "verify you're human", "checking your browser", "attention required",
    "unusual traffic", "security check", "access denied", "are you a robot",
    "hcaptcha", "just a moment", "ddos protection",
)


def _is_blocked(html):
    """True if this page is a CAPTCHA/bot-check/block wall rather than a real result.
    A number scraped off a block page is not a real DA/PA value - returning "N/A" is
    always better than silently returning wrong data."""
    low = (html or "").lower()
    return any(marker in low for marker in _BLOCK_MARKERS)


def _find_metric(html, labels, limit=100):
    """
    Find a numeric metric near a label in HTML.
    Looks for patterns like:
      - <span>DA</span><span>45</span>
      - Domain Authority: 45
      - DA 45
      - data-value="45" near a label
    Only returns values 1-100.
    """
    for label in labels:
        patterns = [
            # Label followed by number in nearby tag
            re.compile(
                r'(?:>|\b)' + re.escape(label) + r'[^0-9]{0,60}?(\d{1,3})(?:\D|$)',
                re.IGNORECASE
            ),
            # Number in tag right after label tag
            re.compile(
                r'(?:>|\b)' + re.escape(label) + r'</[^>]+>\s*<[^>]+>\s*(\d{1,3})\b',
                re.IGNORECASE
            ),
            # data-value or value attribute near label
            re.compile(
                r'' + re.escape(label) + r'[^>]{0,100}(?:data-value|value)\s*=\s*["\'](\d{1,3})["\']',
                re.IGNORECASE
            ),
        ]
        for pat in patterns:
            for m in pat.finditer(html):
                val = int(m.group(1))
                if 1 <= val <= limit:
                    return val
    return None


def check_domain_rating(domain, log_fn=None):
    """
    Fetch Ahrefs Domain Rating (DR) via Ahrefs' free public API - no browser needed.
    Cached per domain. Returns int (0-100) or "N/A" if unavailable.
    """
    domain = domain.strip().lower()
    if domain in _dr_cache:
        return _dr_cache[domain]

    result = "N/A"
    try:
        resp = requests.get(
            "https://api.ahrefs.com/v3/public/domain-rating-free",
            params={"target": domain, "output": "json"},
            headers={"User-Agent": "Mozilla/5.0"},
            timeout=10,
        )
        resp.raise_for_status()
        data = resp.json()
        dr = data.get("domain_rating", {}).get("domain_rating")
        if dr is not None:
            result = round(dr)
    except Exception as e:
        if log_fn:
            log_fn(f"  DR fetch failed for {domain}: {e}")
        logger.debug(f"DR fetch failed for {domain}: {e}")

    _dr_cache[domain] = result
    return result


_CHECKERS = [
    ("wpblogging101.com", _try_wpblogging101),
    ("rhinorank.io", _try_rhinorank),
    ("websiteseochecker.com", _try_websiteseochecker),
    ("dapa-checker.com", _try_dapa_checker),
    ("da-checker.org", _try_da_checker_org),
    ("teqtop.com", _try_teqtop),
]


def check_da_pa(driver, url_or_domain, log_fn=None):
    """
    Check DA/PA for a domain using the Selenium browser.
    Tries multiple free tools in sequence, returns first successful result.
    Results are cached per domain.

    Args:
        driver: Selenium WebDriver instance
        url_or_domain: URL or domain to check
        log_fn: optional logging function

    Returns: {"da": int|str, "dr": int|str, "pa": int|str, "source": str}
    """
    domain = _extract_domain(url_or_domain)

    if domain in _da_cache:
        cached = _da_cache[domain]
        if log_fn:
            log_fn(f"  DA/PA for {domain}: cached - DA={cached.get('da','N/A')}, DR={cached.get('dr','N/A')}, PA={cached.get('pa','N/A')}")
        return cached

    if log_fn:
        log_fn(f"  Checking DA/PA for {domain}...")

    # DR comes straight from Ahrefs' free public API - no browser needed.
    dr_val = check_domain_rating(domain, log_fn=log_fn)

    # Save current URL to go back after
    try:
        original_url = driver.current_url
    except Exception:
        original_url = None

    for name, checker in _CHECKERS:
        try:
            if log_fn:
                log_fn(f"  Trying {name}...")
            result = checker(driver, domain)
            if result and (isinstance(result.get("da"), int) and result["da"] > 0):
                result["dr"] = dr_val
                _da_cache[domain] = result
                if log_fn:
                    log_fn(f"  DA={result['da']}, DR={dr_val}, PA={result.get('pa','N/A')} (from {result['source']})")
                # Navigate back to avoid interference
                if original_url and original_url.startswith("http"):
                    try:
                        driver.get(original_url)
                        time.sleep(1)
                    except Exception:
                        pass
                return result
        except Exception as e:
            if log_fn:
                log_fn(f"  {name} failed: {e}")
        time.sleep(random.uniform(1, 2))

    fallback = {"da": "N/A", "pa": "N/A", "dr": dr_val, "source": "N/A"}
    _da_cache[domain] = fallback
    if log_fn:
        log_fn(f"  DA/PA: could not retrieve from any tool")
    # Navigate back
    if original_url and original_url.startswith("http"):
        try:
            driver.get(original_url)
            time.sleep(1)
        except Exception:
            pass
    return fallback


def clear_cache():
    """Clear the DA/PA/DR cache."""
    _da_cache.clear()
    _dr_cache.clear()
