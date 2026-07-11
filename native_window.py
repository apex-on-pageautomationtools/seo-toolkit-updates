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


def _ensure_package(module_name, pip_name=None):
    """Both pywebview and cffi ship in requirements.txt for fresh installs, but the
    embedded Python bundle isn't itself distributed via OTA (too large; only source
    files are) - so an existing install updated via OTA may predate one or both.
    Rather than require a full reinstall, install silently in the background on first
    run instead - small, self-contained packages, no user action, no prompt."""
    try:
        __import__(module_name)
        return True
    except ImportError:
        pass
    _log(f"{module_name} not found, attempting silent install of '{pip_name or module_name}'...")
    try:
        result = subprocess.run(
            [sys.executable, "-m", "pip", "install", "--quiet", "--disable-pip-version-check",
             pip_name or module_name],
            timeout=90, capture_output=True, text=True,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        if result.returncode != 0:
            _log(f"pip install {pip_name or module_name} failed (exit {result.returncode}): "
                 f"{result.stderr[-500:] if result.stderr else '(no stderr)'}")
            return False
        import importlib
        importlib.invalidate_caches()
        __import__(module_name)
        _log(f"{module_name} installed and imported successfully.")
        return True
    except Exception as e:
        _log(f"_ensure_package({module_name}) failed: {type(e).__name__}: {e}")
        return False


def main():
    if len(sys.argv) < 2:
        print("Usage: native_window.py <url> [icon_path]", file=sys.stderr)
        sys.exit(2)
    url = sys.argv[1]
    icon_path = sys.argv[2] if len(sys.argv) > 2 else None

    _log(f"Starting native window for {url} (python: {sys.executable})")

    # Best-effort; if either fails, the webview import below raises and the caller
    # falls back to the Edge launch, same as any other failure.
    _ensure_package("webview", pip_name="pywebview")
    _ensure_package("cffi")

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
