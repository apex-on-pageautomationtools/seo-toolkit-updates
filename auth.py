"""
SEO Toolkit Pro — Authentication & Licensing Module
Uses Google Apps Script as backend. Checks email/password + MAC address.
"""

import os
import json
import uuid
import hashlib
import urllib.request
import urllib.parse

BUNDLE_DIR = os.path.dirname(os.path.abspath(__file__))
AUTH_FILE = os.path.join(BUNDLE_DIR, ".auth_token")

# Replace with your deployed Apps Script web app URL
AUTH_API_URL = "https://script.google.com/macros/s/AKfycbzqCGuSSfGThDoI1N88BtRxNjwiLbD2RCVsKaDAV6J171WKvFHv574j3hEHbCwiFLpC/exec"  # <-- Paste your Apps Script URL here

APP_VERSION = "3.2"


def get_mac_address():
    """Get this machine's MAC address as a stable uppercase hex string."""
    mac = uuid.getnode()
    return ':'.join(f'{(mac >> (8 * i)) & 0xFF:02X}' for i in reversed(range(6)))


def _api_call(params, timeout=15):
    """Call the auth API (GET with query params for Apps Script compatibility)."""
    try:
        qs = urllib.parse.urlencode(params)
        url = f"{AUTH_API_URL}?{qs}"
        req = urllib.request.Request(url, headers={"User-Agent": "SEOToolkitPro-Auth/1.0"})
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return json.loads(r.read().decode("utf-8"))
    except Exception as e:
        return {"error": str(e)}


def check_version():
    """Check if this app version is allowed."""
    if not AUTH_API_URL:
        return {"allowed": True, "message": "Auth not configured"}
    return _api_call({"action": "version_check", "version": APP_VERSION})


def login(email, password):
    """Login with email + password. Returns status dict."""
    if not AUTH_API_URL:
        return {"status": "approved", "message": "Auth not configured"}
    mac = get_mac_address()
    result = _api_call({
        "action": "login",
        "email": email,
        "password": password,
        "mac": mac,
        "version": APP_VERSION,
    })
    if result.get("status") == "approved":
        _save_token(email, mac)
    return result


def register(email, password, name=""):
    """Register a new account. Returns status dict."""
    if not AUTH_API_URL:
        return {"status": "registered", "message": "Auth not configured"}
    mac = get_mac_address()
    return _api_call({
        "action": "register",
        "email": email,
        "password": password,
        "mac": mac,
        "name": name,
    })


def check_saved_auth():
    """Check if there's a valid saved auth token. Returns status dict."""
    if not AUTH_API_URL:
        return {"status": "approved", "message": "Auth not configured"}
    token = _load_token()
    if not token:
        return {"status": "no_token"}
    mac = get_mac_address()
    if token.get("mac") != mac:
        _clear_token()
        return {"status": "mac_changed", "message": "Device changed. Please login again."}
    result = _api_call({
        "action": "login",
        "email": token["email"],
        "password": token["password"],
        "mac": mac,
        "version": APP_VERSION,
    })
    if result.get("status") != "approved":
        _clear_token()
    return result


def logout():
    """Clear saved auth token."""
    _clear_token()
    return {"status": "logged_out"}


def _save_token(email, mac):
    """Save auth token locally (email + mac hash for quick re-auth)."""
    token = _load_token() or {}
    token.update({"email": email, "password": token.get("password", ""), "mac": mac})
    try:
        with open(AUTH_FILE, "w", encoding="utf-8") as f:
            json.dump(token, f)
    except Exception:
        pass


def _save_full_token(email, password, mac):
    """Save full credentials for auto-login."""
    try:
        with open(AUTH_FILE, "w", encoding="utf-8") as f:
            json.dump({"email": email, "password": password, "mac": mac}, f)
    except Exception:
        pass


def _load_token():
    """Load saved auth token."""
    try:
        if os.path.exists(AUTH_FILE):
            with open(AUTH_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return None


def _clear_token():
    """Remove saved auth token."""
    try:
        if os.path.exists(AUTH_FILE):
            os.remove(AUTH_FILE)
    except Exception:
        pass


def is_authenticated():
    """Quick check: is there a valid local token?"""
    token = _load_token()
    if not token:
        return False
    mac = get_mac_address()
    return token.get("mac") == mac
