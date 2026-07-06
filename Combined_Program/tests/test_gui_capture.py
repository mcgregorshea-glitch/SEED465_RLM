"""Tests for the GUI Eyes harness (gui_capture / _gui_harness package).

Levels:
  - checks-fire (pure): four deterministic detectors against synthetic WidgetInfo — no display
  - report-schema (pure): Report.format() structure
  - boot (display): SEEDApplication boots under Xvfb; all tabs navigate without exception
  - capture smoke (display + backend): Tier 2 writes a non-empty downscaled PNG
"""

import os
import sys
import pytest

# ── Path setup ───────────────────────────────────────────────────────────────
_tests_dir = os.path.dirname(os.path.abspath(__file__))
_sc_root = os.path.dirname(_tests_dir)
_tools_dir = os.path.join(_sc_root, "tools")
if _tools_dir not in sys.path:
    sys.path.insert(0, _tools_dir)

# ── Display availability ──────────────────────────────────────────────────────
def _xvfb_reachable() -> bool:
    """True only if we have a working X server (not just DISPLAY set)."""
    # Delegate to the same logic used by launch.py
    sys.path.insert(0, _tools_dir)
    try:
        from _gui_harness.launch import _has_real_display
        return _has_real_display()
    except ImportError:
        return False


_XVFB_OK = _xvfb_reachable()

SKIP_NO_XVFB = pytest.mark.skipif(
    not _XVFB_OK,
    reason=(
        "No working X server found. "
        "Install and start Xvfb: sudo apt-get install -y xvfb && "
        "Xvfb :99 -screen 0 1100x800x24 &  DISPLAY=:99"
    )
)


# ═══════════════════════════════════════════════════════════════════════════
# Pure unit tests — detector check_* functions (no display required)
# ═══════════════════════════════════════════════════════════════════════════

from _gui_harness.structural import WidgetInfo, check_offscreen, check_zero_size, check_overlaps, check_clipped


def _wi(path, x, y, w, h, parent="root", cls="Label", manager="place"):
    return WidgetInfo(path=path, cls=cls, x=x, y=y, width=w, height=h,
                      manager=manager, parent_path=parent)


def test_check_offscreen_detects_far_widget():
    widgets = [
        _wi("root", 0, 0, 200, 200, parent="", cls="Tk"),
        _wi("root.lbl", 9000, 9000, 50, 20),
    ]
    findings = check_offscreen(widgets, win_w=200, win_h=200, root_path="root")
    assert any(f.widget == "root.lbl" for f in findings), f"Expected finding: {findings}"


def test_check_offscreen_ok_for_in_bounds():
    widgets = [
        _wi("root", 0, 0, 200, 200, parent="", cls="Tk"),
        _wi("root.lbl", 10, 10, 50, 20),
    ]
    findings = check_offscreen(widgets, 200, 200, root_path="root")
    assert not findings


def test_check_zero_size_detects_zero_width():
    widgets = [_wi("root.frame", 10, 10, 0, 50)]
    findings = check_zero_size(widgets)
    assert any(f.widget == "root.frame" for f in findings)


def test_check_zero_size_ok_for_normal():
    widgets = [_wi("root.frame", 10, 10, 50, 50)]
    assert not check_zero_size(widgets)


def test_check_overlaps_detects_sibling_overlap():
    # Two siblings sharing parent="root", overlapping horizontally
    widgets = [
        _wi("root.a", 5,  5, 60, 30),   # x:5-65
        _wi("root.b", 40, 5, 60, 30),   # x:40-100 → 25px overlap with a
    ]
    findings = check_overlaps(widgets)
    assert findings, f"Expected overlap finding, got: {findings}"
    assert any("px" in f.detail for f in findings)


def test_check_overlaps_ok_for_non_overlapping():
    widgets = [
        _wi("root.a", 0,  0, 50, 30),
        _wi("root.b", 60, 0, 50, 30),
    ]
    assert not check_overlaps(widgets)


def test_check_clipped_detects_child_outside_parent():
    widgets = [
        _wi("container", 10, 10, 50, 50, parent="root", cls="Frame"),
        # child starts at same coords as container but is 200px wide
        _wi("container.child", 10, 10, 200, 20, parent="container"),
    ]
    findings = check_clipped(widgets)
    assert any(f.widget == "container.child" for f in findings), f"{findings}"


def test_check_clipped_ok_for_child_inside_parent():
    widgets = [
        _wi("container", 10, 10, 100, 100, parent="root", cls="Frame"),
        _wi("container.child", 15, 15, 30, 30, parent="container"),
    ]
    assert not check_clipped(widgets)


# ═══════════════════════════════════════════════════════════════════════════
# Report schema (pure — no display)
# ═══════════════════════════════════════════════════════════════════════════

from _gui_harness.structural import Report, Finding


def test_report_format_schema():
    widgets = [_wi("root.a", 0, 0, 50, 50), _wi("root.b", 0, 0, 50, 50)]
    findings = [
        Finding(kind="offscreen", widget="root.a", detail="bbox=(...)"),
        Finding(kind="overlap",   widget="root.b", detail="..."),
    ]
    report = Report(tab="sender", window_w=1100, window_h=800,
                    widgets=widgets, findings=findings)
    text = report.format()

    assert "GUI REPORT" in text
    assert "tab=sender" in text
    assert "window=1100x800" in text
    assert "widgets=2" in text
    for category in ("offscreen", "zero-size", "overlap", "clipped"):
        assert category in text, f"Missing {category!r} in report"
    assert "FAIL" in text
    assert "WARN" in text
    assert "-> " in text


def test_report_uninspectable_window_is_not_green():
    """A 0x0 / single-widget window means nothing rendered — must NOT report a pass.

    Regression: a broken display-probe dropped the harness into mocked Tk and
    produced 'window=0x0 widgets=1 -> 0 FAIL, 0 WARN', a false green.
    """
    empty = Report(tab="sender", window_w=0, window_h=0,
                   widgets=[_wi("root", 0, 0, 0, 0, parent="", cls="Tk")],
                   findings=[])
    assert empty.inspectable is False
    text = empty.format()
    assert "UNINSPECTABLE" in text
    assert "0 FAIL, 0 WARN" not in text

    real = Report(tab="sender", window_w=1100, window_h=800,
                  widgets=[_wi("root", 0, 0, 1100, 800, parent="", cls="Tk"),
                           _wi("root.a", 0, 0, 50, 50)],
                  findings=[])
    assert real.inspectable is True


# ═══════════════════════════════════════════════════════════════════════════
# Boot test — requires a working X server (Xvfb)
# ═══════════════════════════════════════════════════════════════════════════

@SKIP_NO_XVFB
def test_seed_app_boots_and_all_tabs_navigate():
    from _gui_harness.launch import HarnessSession, ALL_TAB_KEYS

    with HarnessSession() as session:
        for key in ALL_TAB_KEYS:
            session.select_tab(key)
            session.settle(100)
        assert session.root is not None
        assert session.app is not None


# ═══════════════════════════════════════════════════════════════════════════
# Capture smoke test — requires Xvfb + a capture backend
# ═══════════════════════════════════════════════════════════════════════════

@SKIP_NO_XVFB
def test_capture_smoke(tmp_path):
    from _gui_harness.display import capture_backend
    backend = capture_backend()
    if backend is None:
        pytest.skip("No capture backend (install scrot: sudo apt-get install -y scrot)")

    from _gui_harness.launch import HarnessSession
    from _gui_harness.capture import screenshot

    display = os.environ.get("DISPLAY", ":99")
    with HarnessSession() as session:
        session.select_tab("sender")
        session.settle(300)
        path = screenshot(tab="sender", out_dir=str(tmp_path),
                          display=display, backend=backend)

    assert path is not None, "screenshot() returned None"
    assert os.path.exists(path), f"Output file not found: {path}"
    assert os.path.getsize(path) > 0, "Screenshot file is empty"

    from PIL import Image
    img = Image.open(path)
    assert img.width <= 900, f"Image too wide: {img.width}px"
