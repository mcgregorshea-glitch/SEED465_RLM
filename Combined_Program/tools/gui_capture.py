#!/usr/bin/env python3
"""GUI Eyes — headless GUI verification harness for the SEED Control Center.

Usage examples:
    python3 seed-control-center/tools/gui_capture.py --doctor
    python3 seed-control-center/tools/gui_capture.py --tab sender --report
    python3 seed-control-center/tools/gui_capture.py --tab sender --shot
    python3 seed-control-center/tools/gui_capture.py --all --report
    python3 seed-control-center/tools/gui_capture.py --all --shot
"""

import argparse
import os
import sys

# Ensure the tools package is importable
_tools_dir = os.path.dirname(os.path.abspath(__file__))
if _tools_dir not in sys.path:
    sys.path.insert(0, _tools_dir)


ALL_TABS = ["pattern", "manual", "sender", "vivigo-controls", "pop-viz"]
DEFAULT_OUT_DIR = os.path.join(
    os.path.dirname(os.path.dirname(_tools_dir)), "out", "gui"
)


def _ensure_xvfb(geometry: str) -> bool:
    """Start Xvfb if available and DISPLAY isn't set already. Returns True if display ready."""
    from _gui_harness.display import start_xvfb, xvfb_available
    if os.environ.get("DISPLAY"):
        return True
    if xvfb_available():
        return start_xvfb(geometry=f"{geometry}x24")
    return False


def cmd_doctor(args: argparse.Namespace) -> int:
    from _gui_harness.display import print_doctor
    return print_doctor()


def cmd_report(tabs: list[str], geometry: str) -> int:
    from _gui_harness.launch import HarnessSession, _has_real_display
    from _gui_harness.structural import run_report

    _ensure_xvfb(geometry)
    if not _has_real_display():
        print(
            "ERROR: No usable X display. Tier-1 reports require a real (or virtual)\n"
            "       X server; under mocked Tk the geometry is meaningless and a\n"
            "       report would be falsely green. Refusing to emit one.\n"
            "       Fix: install + start Xvfb (see --doctor), then retry."
        )
        return 2

    exit_code = 0
    with HarnessSession(geometry=geometry) as session:
        for tab in tabs:
            session.select_tab(tab)
            session.settle(300)
            report = run_report(session.root, tab)
            print(report.format())
            print()
            if not report.inspectable:
                exit_code = max(exit_code, 2)
            elif any(f.severity == "FAIL" for f in report.findings):
                exit_code = max(exit_code, 1)
    return exit_code


def cmd_shot(tabs: list[str], geometry: str, out_dir: str) -> int:
    from _gui_harness.launch import HarnessSession
    from _gui_harness.capture import screenshot
    from _gui_harness.display import capture_backend

    backend = capture_backend()
    if not backend:
        print("No capture backend available. Install scrot: sudo apt-get install -y scrot")
        return 1

    display_ok = _ensure_xvfb(geometry)
    if not display_ok:
        print("No virtual display. Install Xvfb: sudo apt-get install -y xvfb")
        return 1

    display = os.environ.get("DISPLAY", ":99")
    with HarnessSession(geometry=geometry) as session:
        for tab in tabs:
            session.select_tab(tab)
            session.settle(500)
            path = screenshot(tab=tab, out_dir=out_dir, display=display,
                          backend=backend, root_tk=session.root)
            if path:
                print(f"  Saved: {path}")
            else:
                print(f"  FAILED capture for tab={tab}")
                return 1
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(
        description="GUI Eyes — headless SEED Control Center verifier"
    )
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--doctor", action="store_true",
                       help="Check prerequisites and print actionable fixes")
    group.add_argument("--tab", metavar="TAB",
                       help=f"Tab key to inspect. One of: {', '.join(ALL_TABS)}")
    group.add_argument("--all", action="store_true",
                       help="Inspect all tabs")

    parser.add_argument("--report", action="store_true",
                        help="Tier 1: emit structural text report (default when no --shot)")
    parser.add_argument("--shot", action="store_true",
                        help="Tier 2: capture a downscaled PNG screenshot")
    parser.add_argument("--geometry", default="1100x800",
                        help="Window size for rendering (default: 1100x800)")
    parser.add_argument("--out-dir", default=DEFAULT_OUT_DIR,
                        help=f"Output directory for screenshots (default: {DEFAULT_OUT_DIR})")

    args = parser.parse_args()

    if args.doctor:
        return cmd_doctor(args)

    tabs = ALL_TABS if args.all else [args.tab]
    if args.tab and args.tab not in ALL_TABS:
        parser.error(f"Unknown tab {args.tab!r}. Valid: {', '.join(ALL_TABS)}")

    # Default to --report if neither flag given
    do_report = args.report or not args.shot
    do_shot = args.shot

    code = 0
    if do_report:
        code = max(code, cmd_report(tabs, args.geometry))
    if do_shot:
        code = max(code, cmd_shot(tabs, args.geometry, args.out_dir))
    return code


if __name__ == "__main__":
    sys.exit(main())
