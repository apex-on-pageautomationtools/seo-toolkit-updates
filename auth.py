"""
SEO Toolkit Pro — Authentication & Licensing Module
Uses Google Apps Script as backend. Checks email/password + MAC address.
Supports multiple accounts per device and multiple devices per account.
"""

import os
import json
import time
import uuid
import hashlib
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
APP_VERSION = "4.0"
SESSION_MAX_HOURS = 12   # a saved login stays valid this long, then re-login is required


# Central login gateway (multi-tenant). Delivered via OTA so the whole team's login
# can be pointed at a new backend with a single push. This TAKES PRECEDENCE over any
# user_auth_url in config.json — set it to "" to fall back to per-machine config.
# ROLLBACK: replace this with the previous PTP script /exec URL and re-push via OTA.
DEFAULT_AUTH_URL = "https://script.google.com/macros/s/AKfycbyYhBdjOinGEFgN_unzHXlSuhQWpqdGoipN4dB1iVTRxrqoq_5c_grAi1LAjCG80VTwmw/exec"


def _get_api_url():
    """Login endpoint. Prefers the OTA-managed DEFAULT_AUTH_URL; if that's blank,
    falls back to user_auth_url in config.json (legacy per-machine setting)."""
    if DEFAULT_AUTH_URL.strip():
        return DEFAULT_AUTH_URL.strip()
    try:
        cf = _config_file()
        if os.path.exists(cf):
            with open(cf, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            return (cfg.get("user_auth_url") or "").strip()
    except Exception:
        pass
    return ""


def _device_id_file():
    return os.path.join(_writable_dir(), ".device_id")


def _persisted_device_id():
    """A random id generated once and stored on disk, so it stays constant for
    this install even if the registry read below fails. Last-resort guarantee."""
    path = _device_id_file()
    try:
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                v = f.read().strip()
            if v:
                return v
    except Exception:
        pass
    v = uuid.uuid4().hex
    try:
        with open(path, "w", encoding="utf-8") as f:
            f.write(v)
    except Exception:
        pass
    return v


def _stable_device_key():
    """A durable, per-machine key. Prefers the Windows MachineGuid — unique per OS
    install; it survives reboots, PC renames and network-adapter changes — and
    falls back to a persisted random id so the value is ALWAYS stable here."""
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                            r"SOFTWARE\Microsoft\Cryptography") as k:
            guid, _ = winreg.QueryValueEx(k, "MachineGuid")
        if guid and str(guid).strip():
            return "MG:" + str(guid).strip()
    except Exception:
        pass
    return "PID:" + _persisted_device_id()


def get_mac_address():
    """Stable per-device ID (name kept for backend compatibility).

    This used to return uuid.getnode() — a NETWORK MAC that flips between the
    Hyper-V / VPN / random adapters run-to-run (e.g. 00:15:5D:… one day,
    01:01:01:01:00:00 the next) and kept tripping the device lock. It now derives
    from the durable Windows MachineGuid, so the same machine always yields the
    same id. Format: 8 hyphen-separated 4-char groups, e.g. A1B2-C3D4-…-CDEF."""
    key = _stable_device_key()
    digest = hashlib.sha256(key.encode("utf-8", "ignore")).hexdigest().upper()
    d = digest[:32]
    return "-".join(d[i:i + 4] for i in range(0, 32, 4))


def _api_call(params, timeout=15):
    """Call the auth API (GET with query params for Apps Script compatibility).

    Apps Script web apps can cold-start with a transient 404/timeout on the first hit,
    so retry a couple of times before giving up — otherwise a valid login would fail."""
    api_url = _get_api_url()
    if not api_url:
        return {"error": "Auth API URL not configured"}
    qs = urllib.parse.urlencode(params)
    url = f"{api_url}?{qs}"
    last_err = None
    for attempt in range(3):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "SEOToolkitPro-Auth/1.0"})
            with urllib.request.urlopen(req, timeout=timeout) as r:
                return json.loads(r.read().decode("utf-8"))
        except Exception as e:
            last_err = e
            if attempt < 2:
                time.sleep(1.2 * (attempt + 1))
    return {"error": str(last_err)}


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
                      allowed_formats=result.get("allowed_formats"),
                      allowed_tools=result.get("allowed_tools"),
                      gsc_url=result.get("gsc_url"),
                      building=result.get("building"),
                      role=result.get("role"))
    return result


def change_password(email, old_password, new_password):
    """Change the logged-in user's password via the backend, then update the
    saved token so auto-login keeps working and refresh the session window."""
    if not _get_api_url():
        return {"status": "error", "message": "Auth not configured"}
    result = _api_call({
        "action": "change_password",
        "email": email,
        "old_password": old_password,
        "new_password": new_password,
    })
    if result.get("status") == "password_changed":
        accounts = _load_accounts()
        if email in accounts:
            accounts[email]["password"] = new_password
            accounts[email]["login_time"] = time.time()
            _save_accounts(accounts)
    return result


def update_name(email, name):
    """Change the logged-in user's display name (self-service - email itself is
    never editable this way, only an admin can change it). Uses the password
    already saved locally for this account, same trust model as periodic
    re-validation - no need to make the user re-type it just to rename themselves."""
    if not _get_api_url():
        return {"status": "error", "message": "Auth not configured"}
    accounts = _load_accounts()
    acct = accounts.get(email)
    if not acct:
        return {"status": "error", "message": "Not logged in."}
    result = _api_call({
        "action": "update_name",
        "email": email,
        "password": acct.get("password", ""),
        "name": name,
    })
    if result.get("status") == "name_changed":
        acct["name"] = name
        _save_accounts(accounts)
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
    # Session expiry — force re-login after SESSION_MAX_HOURS since login.
    lt = acct.get("login_time")
    if lt is None:                       # backfill for tokens saved before this feature
        acct["login_time"] = time.time()
        _save_accounts(accounts)
        lt = acct["login_time"]
    if (time.time() - lt) > SESSION_MAX_HOURS * 3600:
        accounts.pop(email, None)
        _save_accounts(accounts)
        return {"status": "session_expired", "email": email,
                "message": "Session expired. Please log in again."}
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
        if result.get("allowed_tools"):
            acct["allowed_tools"] = result["allowed_tools"]
        # Keep the per-building GSC endpoint / building / role fresh on each re-validate.
        acct["gsc_url"] = result.get("gsc_url", acct.get("gsc_url", ""))
        acct["building"] = result.get("building", acct.get("building", ""))
        acct["role"] = result.get("role", acct.get("role", ""))
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


def get_admin_credentials():
    """(email, password) of the first logged-in account on this device that is an
    admin, or (None, None) if none. The backend's _callerContext() already accepts
    admin_email/admin_password as an alternative to a shared admin_key - since the
    password is already saved locally for periodic re-validation, a logged-in admin
    shouldn't need a SEPARATE admin_key configured just to use the Admin tab."""
    accounts = _load_accounts()
    for em, acct in accounts.items():
        if acct.get("is_admin"):
            return em, acct.get("password", "")
    return None, None


def get_any_credentials():
    """(email, password) of the first logged-in account on this device, admin or
    not. Used for server-side-only calls (e.g. fetching the shared proxy pool) that
    need any valid authenticated user rather than specifically an admin."""
    accounts = _load_accounts()
    for em, acct in accounts.items():
        return em, acct.get("password", "")
    return None, None


def get_allowed_formats():
    """Return the allowed formats for the currently logged-in user, or None (all allowed)."""
    accounts = _load_accounts()
    for em, acct in accounts.items():
        fmts = acct.get("allowed_formats")
        if fmts:
            return fmts
    return None


def get_allowed_tools():
    """Return the allowed tools for the currently logged-in user, or None (all allowed)."""
    accounts = _load_accounts()
    for em, acct in accounts.items():
        tools = acct.get("allowed_tools")
        if tools:
            return tools
    return None


def get_gsc_url():
    """Per-building GSC endpoint for the logged-in user (empty if none/super admin).
    The tool routes Crawl/GSC to this so a teammate only reaches their building's accounts."""
    accounts = _load_accounts()
    for em, acct in accounts.items():
        url = (acct.get("gsc_url") or "").strip()
        if url:
            return url
    return ""


def get_building():
    """Building of the logged-in user (empty if none/super admin)."""
    accounts = _load_accounts()
    for em, acct in accounts.items():
        b = (acct.get("building") or "").strip()
        if b:
            return b
    return ""


def _save_account(email, password, mac, is_admin=False, allowed_formats=None,
                  allowed_tools=None, gsc_url=None, building=None, role=None):
    """Add or update one account in the multi-account token file."""
    accounts = _load_accounts()
    accounts[email] = {"password": password, "mac": mac, "is_admin": is_admin,
                       "login_time": time.time()}
    if allowed_formats:
        accounts[email]["allowed_formats"] = allowed_formats
    if allowed_tools:
        accounts[email]["allowed_tools"] = allowed_tools
    if gsc_url:
        accounts[email]["gsc_url"] = gsc_url
    if building:
        accounts[email]["building"] = building
    if role:
        accounts[email]["role"] = role
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
