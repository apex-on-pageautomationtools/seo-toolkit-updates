"""
SEO Toolkit Pro - native window launcher (pywebview + WebView2).

Opens the already-running Flask server (see Start Tool.vbs, which starts the server
and waits for the port file before launching this) in a real native window instead of
an Edge browser window - own taskbar icon/grouping, no address bar, no "Edge" branding,
while reusing the exact same HTML/CSS/JS frontend and Windows' built-in WebView2
runtime (already installed alongside Edge on effectively every Windows 10/11 machine -
nothing new for the end user to install).

Usage: python native_window.py <url> [icon_path]
Exits non-zero on any failure to create the window, so the caller (Start Tool.vbs) can
fall back to the previous Edge --app launch - this never leaves someone unable to open
the tool just because one machine's WebView2/.NET setup is unusual.
"""
import sys
import os
import tempfile
import traceback
import subprocess
from datetime import datetime

# VBS launches this hidden (WindowStyle 0) so nothing captures stdout/stderr - log to
# a file instead so a failure (and WHY) is actually inspectable afterward, rather than
# silently falling back to Edge with no way to tell what went wrong.
LOG_PATH = os.path.join(tempfile.gettempdir(), "seotoolkitpro_native_window.log")


def _log(msg):
    try:
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}\n")
    except Exception:
        pass


def _try_import(module_name):
    try:
        __import__(module_name)
        return True
    except ImportError:
        return False


def _background_install(missing):
    """Fire-and-forget: kick off a DETACHED pip install for whatever's missing, so a
    future launch has it, WITHOUT this launch waiting on it. A synchronous pip install
    here (the previous approach) could hang for minutes with zero visible feedback if
    the machine has no PyPI access (corporate network/firewall/offline) - that looked
    exactly like "the tool isn't loading" since VBS runs this whole script hidden and
    waits for it. Never block the actual window/fallback decision on a network call."""
    try:
        pip_names = {"webview": "pywebview"}
        args = [sys.executable, "-m", "pip", "install", "--quiet", "--disable-pip-version-check"]
        args += [pip_names.get(m, m) for m in missing]
        subprocess.Popen(
            args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        _log(f"Kicked off background install (not waited on) for: {missing}")
    except Exception as e:
        _log(f"Could not start background install: {type(e).__name__}: {e}")


def main():
    if len(sys.argv) < 2:
        print("Usage: native_window.py <url> [icon_path]", file=sys.stderr)
        sys.exit(2)
    url = sys.argv[1]
    icon_path = sys.argv[2] if len(sys.argv) > 2 else None

    _log(f"Starting native window for {url} (python: {sys.executable})")

    # Fast, local, no network: just check what's importable right now. If anything's
    # missing, fail THIS launch immediately (falls back to Edge with no delay) and
    # kick off a non-blocking background install so next launch can use the native
    # window - never make the user wait on pip/network before they see anything.
    missing = [m for m in ("webview", "cffi") if not _try_import(m)]
    if missing:
        _log(f"Missing package(s): {missing} - falling back to Edge this launch.")
        _background_install(missing)
        sys.exit(1)

    import webview
    _log(f"webview module OK: {webview.__file__}")

    window_kwargs = dict(
        title="SEO Toolkit Pro",
        url=url,
        width=1100,
        height=820,
        min_size=(800, 600),
        confirm_close=False,
    )
    webview.create_window(**window_kwargs)

    start_kwargs = {"gui": "edgechromium"}
    if icon_path and os.path.exists(icon_path):
        # pywebview's Windows backend picks up the taskbar/window icon from the exe
        # icon by default; an explicit .ico here isn't universally supported across
        # pywebview versions, so this is best-effort and never fatal on its own.
        start_kwargs["icon"] = icon_path

    _log("Calling webview.start()...")
    webview.start(**start_kwargs)
    _log("webview.start() returned (window closed normally).")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        _log(f"FAILED: {type(e).__name__}: {e}\n{traceback.format_exc()}")
        print(f"[native_window] Failed to open native window: {type(e).__name__}: {e}", file=sys.stderr)
        sys.exit(1)
