"""
DA/PA Checker — uses the Selenium browser to check DA/PA via free online tools.
Tries multiple tools in sequence, caches results per domain.
"""

import re
import time
import random
import logging
from urllib.parse import urlparse

logger = logging.getLogger(__name__)

_da_cache = {}


def _extract_domain(url_or_domain):
    d = url_or_domain.strip()
    if not d.startswith("http"):
        d = "https://" + d
    parsed = urlparse(d)
    return parsed.netloc.replace("www.", "").lower()


def _try_rhinorank(driver, domain):
    """rhinorank.io DA/DR checker"""
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
        da = _find_metric(src, ["domain authority", "DA"], limit=100)
        dr = _find_metric(src, ["domain rating", "DR"], limit=100)
        pa = _find_metric(src, ["page authority", "PA"], limit=100)

        if da is not None or dr is not None:
            return {"da": da if da is not None else "—",
                    "pa": pa if pa is not None else "—",
                    "dr": dr if dr is not None else "—",
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
        da = _find_metric(src, ["domain authority", "DA Score", "DA"], limit=100)
        pa = _find_metric(src, ["page authority", "PA Score", "PA"], limit=100)

        if da is not None:
            return {"da": da, "pa": pa if pa is not None else "—",
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
        da = _find_metric(src, ["domain authority", "DA"], limit=100)
        pa = _find_metric(src, ["page authority", "PA"], limit=100)

        if da is not None:
            return {"da": da, "pa": pa if pa is not None else "—",
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
        da = _find_metric(src, ["domain authority", "DA"], limit=100)
        pa = _find_metric(src, ["page authority", "PA"], limit=100)

        if da is not None:
            return {"da": da, "pa": pa if pa is not None else "—",
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
        da = _find_metric(src, ["domain authority", "DA"], limit=100)
        pa = _find_metric(src, ["page authority", "PA"], limit=100)

        if da is not None:
            return {"da": da, "pa": pa if pa is not None else "—",
                    "source": "teqtop.com"}
    except Exception as e:
        logger.debug(f"teqtop failed: {e}")
    return None


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


def _find_metric_in_text(text, labels, limit=100):
    """Find a numeric metric in rendered page text.

    This is best for browser extensions or toolbars that inject metrics into the
    page DOM instead of a dedicated checker page.
    """
    if not text:
        return None
    low = re.sub(r"\s+", " ", text)
    for label in labels:
        for m in re.finditer(re.escape(label), low, re.IGNORECASE):
            window = low[m.end(): m.end() + 90]
            m2 = re.search(r"[:=\-]?\s*(\d{1,3})\b", window)
            if not m2:
                # Also allow the number to appear just before the label on the same line.
                prefix = low[max(0, m.start() - 30): m.start()]
                m2 = re.search(r"(\d{1,3})\s*[:=\-]?\s*$", prefix)
            if m2:
                val = int(m2.group(1))
                if 1 <= val <= limit:
                    return val
    return None


def _try_rendered_page_metrics(driver, domain):
    """Best-effort scan of the current rendered page for DA/DR/PA values.

    This is intended for SEO toolbars/extensions that inject the metrics into
    the page DOM while the backlink page is open.
    """
    try:
        from selenium.webdriver.common.by import By

        body_text = ""
        try:
            body_text = driver.execute_script(
                "return document.body ? document.body.innerText : ''"
            ) or ""
        except Exception:
            pass

        try:
            body = driver.find_element(By.TAG_NAME, "body")
            body_text = (body_text or "") + "\n" + (body.text or "")
        except Exception:
            pass

        src = ""
        try:
            src = driver.page_source or ""
        except Exception:
            pass

        combined = f"{body_text}\n{src}"
        da = _find_metric_in_text(combined, ["domain authority", "da"], limit=100)
        dr = _find_metric_in_text(combined, ["domain rating", "dr"], limit=100)
        pa = _find_metric_in_text(combined, ["page authority", "pa"], limit=100)

        if da is not None or dr is not None or pa is not None:
            return {
                "da": da if da is not None else "—",
                "dr": dr if dr is not None else "—",
                "pa": pa if pa is not None else "—",
                "source": "browser-extension",
            }
    except Exception as e:
        logger.debug(f"rendered-page metrics failed: {e}")
    return None


_CHECKERS = [
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

    Returns: {"da": int|str, "pa": int|str, "source": str}
    """
    domain = _extract_domain(url_or_domain)

    if domain in _da_cache:
        cached = _da_cache[domain]
        if log_fn:
            log_fn(f"  DA/PA for {domain}: cached — DA={cached.get('da','—')}, PA={cached.get('pa','—')}")
        return cached

    if log_fn:
        log_fn(f"  Checking DA/PA for {domain}...")

    # Save current URL to go back after
    try:
        original_url = driver.current_url
    except Exception:
        original_url = None

    if False:
        # First try the current page itself. This is the path used when a loaded
        # browser extension injects DA/DR metrics into the rendered page.
        try:
            result = _try_rendered_page_metrics(driver, domain)
            if result and (isinstance(result.get("da"), int) and result["da"] > 0):
                _da_cache[domain] = result
                if log_fn:
                    log_fn(
                        f"  DA={result['da']}, DR={result.get('dr','—')}, "
                        f"PA={result.get('pa','—')} (from {result['source']})"
                    )
                if original_url and original_url.startswith("http"):
                    try:
                        driver.get(original_url)
                        time.sleep(1)
                    except Exception:
                        pass
                return result
        except Exception as e:
            if log_fn:
                log_fn(f"  rendered page scan failed: {e}")

    for name, checker in _CHECKERS:
        try:
            if log_fn:
                log_fn(f"  Trying {name}...")
            result = checker(driver, domain)
            if result and (isinstance(result.get("da"), int) and result["da"] > 0):
                _da_cache[domain] = result
                if log_fn:
                    log_fn(f"  DA={result['da']}, PA={result.get('pa','—')} (from {result['source']})")
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

    fallback = {"da": "—", "pa": "—", "source": "—"}
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
    """Clear the DA/PA cache."""
    _da_cache.clear()
