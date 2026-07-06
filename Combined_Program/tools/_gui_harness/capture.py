"""Tier 2: window screenshot via scrot/ImageMagick + Pillow downscale.

Captures the specific Tk window by its X11 window ID rather than the root
framebuffer. Without a window manager, root-screen capture is all black
because windows aren't composited onto the root.
"""

import os
import subprocess
import tempfile
import time
from pathlib import Path
from typing import Optional


MAX_WIDTH = 900  # px — keeps vision token cost ≤ ~1k


def _window_id_hex(root_tk) -> Optional[str]:
    """Return the X11 window ID as a 0x-prefixed hex string, or None."""
    try:
        wid = root_tk.winfo_id()
        return hex(wid)
    except Exception:
        return None


def _capture_scrot_window(display: str, win_id: str, out_path: str) -> bool:
    """Capture a specific window by ID with scrot."""
    env = {**os.environ, "DISPLAY": display}
    try:
        result = subprocess.run(
            ["scrot", "--window", win_id, out_path],
            env=env, capture_output=True, timeout=10
        )
        return result.returncode == 0 and Path(out_path).exists()
    except (OSError, subprocess.TimeoutExpired):
        return False


def _capture_import_window(display: str, win_id: str, out_path: str) -> bool:
    """ImageMagick `import -window <id>`."""
    env = {**os.environ, "DISPLAY": display}
    try:
        result = subprocess.run(
            ["import", "-window", win_id, out_path],
            env=env, capture_output=True, timeout=10
        )
        return result.returncode == 0 and Path(out_path).exists()
    except (OSError, subprocess.TimeoutExpired):
        return False


def _capture_xwd(display: str, win_id: str, out_path: str) -> bool:
    """xwd → convert to PNG (fallback if scrot/import unavailable)."""
    import shutil
    if not shutil.which("xwd") or not shutil.which("convert"):
        return False
    env = {**os.environ, "DISPLAY": display}
    xwd_path = out_path + ".xwd"
    try:
        r1 = subprocess.run(["xwd", "-id", win_id, "-out", xwd_path],
                            env=env, capture_output=True, timeout=10)
        if r1.returncode != 0:
            return False
        r2 = subprocess.run(["convert", xwd_path, out_path],
                            env=env, capture_output=True, timeout=10)
        return r2.returncode == 0 and Path(out_path).exists()
    except (OSError, subprocess.TimeoutExpired):
        return False
    finally:
        Path(xwd_path).unlink(missing_ok=True)


def _downscale(src: str, dst: str, max_width: int = MAX_WIDTH) -> None:
    """Downscale *src* PNG to ≤*max_width* px wide and save to *dst*."""
    from PIL import Image
    img = Image.open(src)
    w, h = img.size
    if w > max_width:
        ratio = max_width / w
        img = img.resize((max_width, int(h * ratio)), Image.LANCZOS)
    img.save(dst, format="PNG", optimize=True)


def screenshot(
    tab: str,
    out_dir: str,
    display: Optional[str] = None,
    max_width: int = MAX_WIDTH,
    backend: Optional[str] = None,
    root_tk=None,
) -> Optional[str]:
    """Capture the app window and write a downscaled PNG.

    *root_tk* is the Tk root window; its X11 ID is used for window-specific
    capture so the correct window is captured even without a WM.

    Returns the output path on success, None on failure.
    """
    if display is None:
        display = os.environ.get("DISPLAY", ":99")

    from .display import capture_backend as _detect_backend
    if backend is None:
        backend = _detect_backend()
    if backend is None:
        print("  No capture backend available (install scrot or imagemagick).")
        return None

    os.makedirs(out_dir, exist_ok=True)
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    raw_path = os.path.join(tempfile.gettempdir(), f"gui_raw_{tab}_{timestamp}.png")
    out_path = os.path.join(out_dir, f"{tab}_{timestamp}.png")

    # Let X process all pending expose/paint events before capture
    if root_tk is not None:
        try:
            root_tk.update()
            root_tk.update_idletasks()
        except Exception:
            pass
    time.sleep(0.3)

    win_id = _window_id_hex(root_tk) if root_tk is not None else None

    ok = False
    if win_id:
        if backend == "scrot":
            ok = _capture_scrot_window(display, win_id, raw_path)
        if not ok:
            ok = _capture_import_window(display, win_id, raw_path)
        if not ok:
            ok = _capture_xwd(display, win_id, raw_path)

    # Fallback: root-screen capture (may be black without WM)
    if not ok:
        if backend == "scrot":
            env = {**os.environ, "DISPLAY": display}
            try:
                r = subprocess.run(["scrot", raw_path], env=env,
                                   capture_output=True, timeout=10)
                ok = r.returncode == 0 and Path(raw_path).exists()
            except (OSError, subprocess.TimeoutExpired):
                pass
        if not ok:
            env = {**os.environ, "DISPLAY": display}
            try:
                r = subprocess.run(["import", "-window", "root", raw_path],
                                   env=env, capture_output=True, timeout=10)
                ok = r.returncode == 0 and Path(raw_path).exists()
            except (OSError, subprocess.TimeoutExpired):
                pass

    if not ok:
        print(f"  Capture failed (backend={backend}, display={display}, win={win_id}).")
        return None

    try:
        _downscale(raw_path, out_path, max_width)
    except Exception as e:
        print(f"  Downscale failed: {e}")
        import shutil
        shutil.copy(raw_path, out_path)

    return out_path
