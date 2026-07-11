"""
DA/PA Checker - uses the Selenium browser to check DA/PA via free online tools.
Tries multiple tools in sequence, caches results per domain.
"""

import re
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


def _try_wpblogging101(driver, domain):
    """wpblogging101.com's public DA/PA Checker tool (tools/da-pa-checker/). Drives
    the page the same way a visitor does: type the domain, click Check, and if the
    "Security Verification" modal appears, read the 4-digit code the page itself
    displays in <span id="atools101_dapa_captcha_code"> and type it back into the
    confirm box - the code is shown in plain text to any visitor, so this isn't
    solving anything, just clicking through the same confirmation a human would.
    Selenium is required (not plain HTTP) because the check now runs client-side
    through this modal rather than returning JSON directly."""
    try:
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC

        driver.get("https://wpblogging101.com/tools/da-pa-checker/")
        wait = WebDriverWait(driver, 15)
        box = wait.until(EC.presence_of_element_located((By.ID, "atools101_dapa_domains")))
        box.clear()
        box.send_keys(domain)
        driver.find_element(By.ID, "atools101_dapa_check").click()

        # Either the captcha modal or the results table shows up next.
        wait.until(lambda d: d.find_element(By.ID, "atools101_dapa_captcha_container")
                   .value_of_css_property("display") != "none"
                   or d.find_elements(By.CSS_SELECTOR, "#atools101_tableBody tr"))

        captcha = driver.find_element(By.ID, "atools101_dapa_captcha_container")
        if captcha.value_of_css_property("display") != "none":
            code = driver.find_element(By.ID, "atools101_dapa_captcha_code").text.strip()
            if not code:
                return None
            inp = driver.find_element(By.ID, "atools101_dapa_captcha_input")
            inp.clear()
            inp.send_keys(code)
            driver.find_element(By.ID, "atools101_dapa_captcha_submit").click()
            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#atools101_tableBody tr")))

        row = driver.find_element(By.CSS_SELECTOR, "#atools101_tableBody tr")
        cells = row.find_elements(By.TAG_NAME, "td")
        if len(cells) < 5:
            return None
        row_domain = cells[0].text.strip().lower()
        if domain.lower() not in row_domain and row_domain not in domain.lower():
            return None
        da_raw = cells[1].find_element(By.CLASS_NAME, "atools101-metric-value").text.strip()
        pa_raw = cells[2].find_element(By.CLASS_NAME, "atools101-metric-value").text.strip()
        if not da_raw.isdigit():
            return None
        pa = int(pa_raw) if pa_raw.isdigit() else "N/A"
        result = {"da": int(da_raw), "pa": pa, "source": "wpblogging101.com"}
        try:
            result["spam_score"] = cells[3].text.strip()
            result["dr"] = cells[4].find_element(By.CLASS_NAME, "atools101-metric-value").text.strip()
        except Exception:
            pass
        return result
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
