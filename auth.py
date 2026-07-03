"""
SEO Toolkit Pro — Authentication & Licensing Module
Uses Google Apps Script as backend. Checks email/password + MAC address.
Supports multiple accounts per device and multiple devices per account.
"""

import os
import json
import uuid
import urllib.request
import urllib.parse

BUNDLE_DIR = os.path.dirname(os.path.abspath(__file__))

def _writable_dir():
    script_dir = BUNDLE_DIR
    pf = os.environ.get("ProgramFiles", "C:\\Program Files").lower()
    pfx86 = os.environ.get("ProgramFiles(x86)", "C:\\Program Files (x86)").lower()
    if script_dir.lower().startswith(pf) or script_dir.lower().startswith(pfx86):
        appdata = os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "SEO Toolkit Pro")
        os.makedirs(appdata, exist_ok=True)
        return appdata
    return script_dir

def _auth_file():
    return os.path.join(_writable_dir(), ".auth_token")

def _config_file():
    d = _writable_dir()
    cf = os.path.join(d, "config.json")
    if os.path.exists(cf):
        return cf
    cf2 = os.path.join(BUNDLE_DIR, "config.json")
    if os.path.exists(cf2):
        return cf2
    return cf

AUTH_FILE = _auth_file()
APP_VERSION = "3.8"


def _get_api_url():
    """Read user_auth_url from config.json. Returns empty string if not configured."""
    try:
        cf = _config_file()
        if os.path.exists(cf):
            with open(cf, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            return (cfg.get("user_auth_url") or "").strip()
    except Exception:
        pass
    return ""


def get_mac_address():
    """Get this machine's MAC address as a stable uppercase hex string."""
    mac = uuid.getnode()
    return ':'.join(f'{(mac >> (8 * i)) & 0xFF:02X}' for i in reversed(range(6)))


def _api_call(params, timeout=15):
    """Call the auth API (GET with query params for Apps Script compatibility)."""
    api_url = _get_api_url()
    if not api_url:
        return {"error": "Auth API URL not configured"}
    try:
        qs = urllib.parse.urlencode(params)
        url = f"{api_url}?{qs}"
        req = urllib.request.Request(url, headers={"User-Agent": "SEOToolkitPro-Auth/1.0"})
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return json.loads(r.read().decode("utf-8"))
    except Exception as e:
        return {"error": str(e)}


def check_version():
    """Check if this app version is allowed."""
    if not _get_api_url():
        return {"allowed": True, "message": "Auth not configured"}
    return _api_call({"action": "version_check", "version": APP_VERSION})


def login(email, password):
    """Login with email + password. Returns status dict."""
    if not _get_api_url():
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
        _save_account(email, password, mac, is_admin=result.get("is_admin", False),
                      allowed_formats=result.get("allowed_formats"))
    return result


def register(email, password, name=""):
    """Register a new account. Returns status dict."""
    if not _get_api_url():
        return {"status": "registered", "message": "Auth not configured"}
    mac = get_mac_address()
    return _api_call({
        "action": "register",
        "email": email,
        "password": password,
        "mac": mac,
        "name": name,
    })


def check_saved_auth(email=None):
    """Check if there's a valid saved auth token.
    If email is given, check that specific account; otherwise check any/all."""
    if not _get_api_url():
        return {"status": "approved", "message": "Auth not configured"}
    accounts = _load_accounts()
    if not accounts:
        return {"status": "no_token"}
    mac = get_mac_address()

    if email:
        acct = accounts.get(email)
        if not acct:
            return {"status": "no_token"}
        return _validate_account(email, acct, mac, accounts)

    for em, acct in list(accounts.items()):
        result = _validate_account(em, acct, mac, accounts)
        if result.get("status") == "approved":
            return result
    return {"status": "no_token"}


def _validate_account(email, acct, mac, accounts):
    """Validate a single stored account against the backend."""
    result = _api_call({
        "action": "login",
        "email": email,
        "password": acct.get("password", ""),
        "mac": mac,
        "version": APP_VERSION,
    })
    if result.get("error"):
        return {"status": "approved", "email": email, "offline": True,
                "is_admin": acct.get("is_admin", False)}
    if result.get("status") == "approved":
        result["email"] = email
        acct["is_admin"] = result.get("is_admin", False)
        if result.get("allowed_formats"):
            acct["allowed_formats"] = result["allowed_formats"]
        _save_accounts(accounts)
        return result
    accounts.pop(email, None)
    _save_accounts(accounts)
    return result


def logout(email=None):
    """Clear saved auth token for a specific email, or all if none given."""
    if email:
        accounts = _load_accounts()
        accounts.pop(email, None)
        _save_accounts(accounts)
    else:
        _clear_all_tokens()
    return {"status": "logged_out"}


def list_logged_in():
    """Return list of all logged-in accounts on this device."""
    accounts = _load_accounts()
    mac = get_mac_address()
    return [{"email": em, "is_admin": a.get("is_admin", False)}
            for em, a in accounts.items()]


def get_allowed_formats():
    """Return the allowed formats for the currently logged-in user, or None (all allowed)."""
    accounts = _load_accounts()
    for em, acct in accounts.items():
        fmts = acct.get("allowed_formats")
        if fmts:
            return fmts
    return None


def _save_account(email, password, mac, is_admin=False, allowed_formats=None):
    """Add or update one account in the multi-account token file."""
    accounts = _load_accounts()
    accounts[email] = {"password": password, "mac": mac, "is_admin": is_admin}
    if allowed_formats:
        accounts[email]["allowed_formats"] = allowed_formats
    _save_accounts(accounts)


def _load_accounts():
    """Load all accounts from token file. Handles old single-account format."""
    try:
        if os.path.exists(AUTH_FILE):
            with open(AUTH_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict) and "accounts" in data:
                return data["accounts"]
            if isinstance(data, dict) and "email" in data:
                em = data["email"]
                return {em: {"password": data.get("password", ""),
                             "mac": data.get("mac", ""),
                             "is_admin": data.get("is_admin", False)}}
    except Exception:
        pass
    return {}


def _save_accounts(accounts):
    """Save all accounts to token file."""
    try:
        with open(AUTH_FILE, "w", encoding="utf-8") as f:
            json.dump({"accounts": accounts}, f)
    except Exception:
        pass


def _clear_all_tokens():
    """Remove saved auth token file."""
    try:
        if os.path.exists(AUTH_FILE):
            os.remove(AUTH_FILE)
    except Exception:
        pass


# --- Backward-compatible aliases used by web_app_batch.py ---

def _load_token():
    """Load first saved account as a flat token dict (backward compat)."""
    accounts = _load_accounts()
    if not accounts:
        return None
    email = next(iter(accounts))
    acct = accounts[email]
    return {"email": email, "password": acct.get("password", ""),
            "mac": acct.get("mac", ""), "is_admin": acct.get("is_admin", False)}


def _save_full_token(email, password, mac):
    """Save full credentials (backward compat — wraps _save_account)."""
    _save_account(email, password, mac)


def _clear_token():
    """Remove all saved tokens (backward compat)."""
    _clear_all_tokens()


def is_authenticated():
    """Quick check: is there any valid local token?"""
    accounts = _load_accounts()
    return len(accounts) > 0
