"""Xvfb lifecycle management, DISPLAY env setup, and capture-backend detection."""

import os
import shutil
import subprocess
import time
from dataclasses import dataclass, field


XVFB_DISPLAY = ":99"
_xvfb_proc: subprocess.Popen | None = None


def xvfb_available() -> bool:
    return shutil.which("Xvfb") is not None or shutil.which("xvfb-run") is not None


def capture_backend() -> str | None:
    """Return the name of the first available capture backend, or None."""
    for cmd in ("scrot", "import"):
        if shutil.which(cmd):
            return cmd
    return None


def start_xvfb(display: str = XVFB_DISPLAY, geometry: str = "1100x800x24") -> bool:
    """Start Xvfb on *display*. Returns True if started (or already running)."""
    global _xvfb_proc

    if not shutil.which("Xvfb"):
        return False

    # Already set by caller (e.g. xvfb-run)
    if os.environ.get("DISPLAY") == display:
        return True

    try:
        _xvfb_proc = subprocess.Popen(
            ["Xvfb", display, "-screen", "0", geometry],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        time.sleep(0.5)  # allow X server to start
        if _xvfb_proc.poll() is not None:
            return False
        os.environ["DISPLAY"] = display
        return True
    except OSError:
        return False


def stop_xvfb() -> None:
    """Terminate the Xvfb process started by this module."""
    global _xvfb_proc
    if _xvfb_proc is not None and _xvfb_proc.poll() is None:
        _xvfb_proc.terminate()
        try:
            _xvfb_proc.wait(timeout=5)
        except subprocess.TimeoutExpired:
            _xvfb_proc.kill()
    _xvfb_proc = None


@dataclass
class DoctorResult:
    xvfb: bool = False
    capture_backend: str | None = None
    fonts: dict[str, bool] = field(default_factory=dict)
    pillow: bool = False

    @property
    def ok(self) -> bool:
        return self.xvfb and self.pillow

    @property
    def tier2_ok(self) -> bool:
        return self.xvfb and self.capture_backend is not None and self.pillow


def _check_font(family: str) -> bool:
    try:
        result = subprocess.run(
            ["fc-list", f":family={family}"],
            capture_output=True, text=True, timeout=5
        )
        return bool(result.stdout.strip())
    except (OSError, subprocess.TimeoutExpired):
        return False


def run_doctor() -> DoctorResult:
    """Check all prerequisites and return a DoctorResult."""
    r = DoctorResult()

    r.xvfb = xvfb_available()
    r.capture_backend = capture_backend()

    try:
        import PIL  # noqa: F401
        r.pillow = True
    except ImportError:
        r.pillow = False

    for family in ("Inter", "JetBrains Mono", "Rajdhani", "Orbitron"):
        r.fonts[family] = _check_font(family)

    return r


def print_doctor() -> int:
    """Print a human-readable doctor report. Returns 0 if ready for Tier 1, 1 otherwise."""
    r = run_doctor()
    print("── GUI Eyes Doctor ──────────────────────────────────")

    mark = lambda ok: "  OK  " if ok else " MISS "

    print(f"[{mark(r.xvfb)}] Xvfb (virtual display for headless rendering)")
    if not r.xvfb:
        print("        → sudo apt-get install -y xvfb")

    print(f"[{mark(r.capture_backend is not None)}] Capture backend: {r.capture_backend or 'none found'}")
    if r.capture_backend is None:
        print("        → sudo apt-get install -y scrot")

    print(f"[{mark(r.pillow)}] Pillow (image downscaling)")
    if not r.pillow:
        print("        → python3 $TMPDIR/pip-install/pip install --target=/home/sheamcg/RLM/venv-linux Pillow")

    print()
    print("  Fonts:")
    for family, ok in r.fonts.items():
        print(f"  [{mark(ok)}] {family}")
    missing_fonts = [f for f, ok in r.fonts.items() if not ok]
    if missing_fonts:
        print()
        print("  To install fonts:")
        print("    sudo apt-get install -y fonts-inter fonts-jetbrains-mono")
        print("    # Rajdhani + Orbitron: see docs/specs/gui-eyes-spec.md → Resolved Decisions")

    print()
    if r.tier2_ok:
        print("  STATUS: Ready for Tier 1 (report) and Tier 2 (screenshot)")
    elif r.xvfb:
        print("  STATUS: Ready for Tier 1 (report). Tier 2 needs a capture backend.")
    else:
        print("  STATUS: Not ready — install Xvfb first.")
    print("─────────────────────────────────────────────────────")

    return 0 if r.ok else 1
