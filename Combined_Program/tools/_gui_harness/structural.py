"""Tier 1: widget-tree walk and deterministic geometry checks.

Each detector has two forms:
  - find_X(root)        — reads geometry live from a Tk widget tree (needs real display)
  - check_X(widgets, …) — pure function over WidgetInfo objects (no Tk needed; unit-testable)
"""

import tkinter as tk
from dataclasses import dataclass
from typing import Iterator


@dataclass
class Finding:
    kind: str   # "offscreen" | "overlap" | "zero-size" | "clipped"
    widget: str
    detail: str

    @property
    def severity(self) -> str:
        return "FAIL" if self.kind in ("offscreen", "clipped") else "WARN"


# ── Widget walk ──────────────────────────────────────────────────────────────

def walk(widget: tk.Misc) -> Iterator[tk.Misc]:
    """Yield *widget* and all descendants depth-first."""
    yield widget
    for child in widget.winfo_children():
        yield from walk(child)


def widget_path(w: tk.Misc) -> str:
    return str(w)


# ── Geometry helpers ─────────────────────────────────────────────────────────

def _abs_bbox(w: tk.Misc, root: tk.Misc) -> tuple[int, int, int, int]:
    """Return (x, y, w, h) relative to *root*'s top-left."""
    rx = root.winfo_rootx()
    ry = root.winfo_rooty()
    x = w.winfo_rootx() - rx
    y = w.winfo_rooty() - ry
    return x, y, w.winfo_width(), w.winfo_height()


# ── Widget tree snapshot ─────────────────────────────────────────────────────

@dataclass
class WidgetInfo:
    path: str
    cls: str
    x: int
    y: int
    width: int
    height: int
    manager: str
    parent_path: str = ""


def widget_tree(root: tk.Misc) -> list[WidgetInfo]:
    """Return a flat list of WidgetInfo for every widget in the tree."""
    result = []
    for w in walk(root):
        x, y, ww, wh = _abs_bbox(w, root)
        parent = getattr(w, "master", None)
        result.append(WidgetInfo(
            path=widget_path(w),
            cls=w.winfo_class(),
            x=x, y=y,
            width=ww, height=wh,
            manager=w.winfo_manager(),
            parent_path=widget_path(parent) if parent else "",
        ))
    return result


# ── Pure geometry detectors (no Tk required) ─────────────────────────────────

def check_offscreen(
    widgets: list[WidgetInfo],
    win_w: int,
    win_h: int,
    root_path: str = "",
) -> list[Finding]:
    """Flag widgets whose bounding box falls entirely outside the window."""
    findings: list[Finding] = []
    for wi in widgets:
        if wi.path == root_path:
            continue
        if wi.width <= 0 or wi.height <= 0:
            continue
        x, y, ww, wh = wi.x, wi.y, wi.width, wi.height
        if (x + ww <= 0) or (y + wh <= 0) or (x >= win_w) or (y >= win_h):
            findings.append(Finding(
                kind="offscreen",
                widget=wi.path,
                detail=f"bbox=({x},{y},{ww},{wh}) window=({win_w},{win_h})",
            ))
    return findings


def check_zero_size(
    widgets: list[WidgetInfo],
    root_path: str = "",
) -> list[Finding]:
    """Flag widgets with zero width or height."""
    findings: list[Finding] = []
    for wi in widgets:
        if wi.path == root_path:
            continue
        if wi.width == 0 or wi.height == 0:
            findings.append(Finding(
                kind="zero-size",
                widget=wi.path,
                detail=f"size=({wi.width},{wi.height})",
            ))
    return findings


def check_overlaps(
    widgets: list[WidgetInfo],
    root_path: str = "",
) -> list[Finding]:
    """Flag pairs of sibling widgets whose bounding boxes overlap."""
    findings: list[Finding] = []
    # Group by parent
    by_parent: dict[str, list[WidgetInfo]] = {}
    for wi in widgets:
        if wi.path == root_path:
            continue
        by_parent.setdefault(wi.parent_path, []).append(wi)

    for siblings in by_parent.values():
        visible = [wi for wi in siblings if wi.width > 0 and wi.height > 0]
        for i in range(len(visible)):
            for j in range(i + 1, len(visible)):
                a, b = visible[i], visible[j]
                ox = min(a.x + a.width, b.x + b.width) - max(a.x, b.x)
                oy = min(a.y + a.height, b.y + b.height) - max(a.y, b.y)
                if ox > 0 and oy > 0:
                    findings.append(Finding(
                        kind="overlap",
                        widget=a.path,
                        detail=f"overlaps {b.path}  ({ox}px x {oy}px)",
                    ))
    return findings


def check_clipped(
    widgets: list[WidgetInfo],
    root_path: str = "",
) -> list[Finding]:
    """Flag widgets that extend beyond their parent's bounding box."""
    by_path = {wi.path: wi for wi in widgets}
    findings: list[Finding] = []
    for wi in widgets:
        if wi.path == root_path or not wi.parent_path:
            continue
        parent = by_path.get(wi.parent_path)
        if parent is None or parent.width <= 0 or parent.height <= 0:
            continue

        over_right  = max(0, (wi.x + wi.width)  - (parent.x + parent.width))
        over_bottom = max(0, (wi.y + wi.height) - (parent.y + parent.height))
        over_left   = max(0, parent.x - wi.x)
        over_top    = max(0, parent.y - wi.y)

        if any([over_right, over_bottom, over_left, over_top]):
            parts = []
            if over_right:  parts.append(f"{over_right}px right")
            if over_bottom: parts.append(f"{over_bottom}px below")
            if over_left:   parts.append(f"{over_left}px left")
            if over_top:    parts.append(f"{over_top}px above")
            findings.append(Finding(
                kind="clipped",
                widget=wi.path,
                detail=f"extends {', '.join(parts)} beyond {wi.parent_path}",
            ))
    return findings


# ── Live Tk-backed wrappers (need a real display) ─────────────────────────────

def find_offscreen(root: tk.Misc) -> list[Finding]:
    """Read geometry live from Tk and delegate to check_offscreen."""
    tree = widget_tree(root)
    root_path = widget_path(root)
    return check_offscreen(tree, root.winfo_width(), root.winfo_height(), root_path)


def find_zero_size(root: tk.Misc) -> list[Finding]:
    tree = widget_tree(root)
    return check_zero_size(tree, widget_path(root))


def find_overlaps(root: tk.Misc) -> list[Finding]:
    tree = widget_tree(root)
    return check_overlaps(tree, widget_path(root))


def find_clipped(root: tk.Misc) -> list[Finding]:
    tree = widget_tree(root)
    return check_clipped(tree, widget_path(root))


# ── Report ───────────────────────────────────────────────────────────────────

@dataclass
class Report:
    tab: str
    window_w: int
    window_h: int
    widgets: list[WidgetInfo]
    findings: list[Finding]

    @property
    def inspectable(self) -> bool:
        """False when nothing real was rendered (no usable display / mocked Tk).

        A 0-size window or a tree of one widget means the app never actually
        rendered, so an "all OK" result would be a false pass. Callers must
        treat ``not inspectable`` as an error, never as success.
        """
        return self.window_w > 1 and self.window_h > 1 and len(self.widgets) > 1

    def format(self) -> str:
        header = (
            f"GUI REPORT — tab={self.tab}  "
            f"window={self.window_w}x{self.window_h}  "
            f"widgets={len(self.widgets)}"
        )
        if not self.inspectable:
            return (
                f"{header}\n"
                f"  ERROR  window not rendered (no usable X display — running mocked Tk).\n"
                f"         Tier-1 geometry is meaningless here; this is NOT a pass.\n"
                f"         Start a real X server (e.g. Xvfb) and retry. See --doctor.\n"
                f"  -> UNINSPECTABLE"
            )
        lines = [header]

        checks = {
            "offscreen": [f for f in self.findings if f.kind == "offscreen"],
            "zero-size": [f for f in self.findings if f.kind == "zero-size"],
            "overlap":   [f for f in self.findings if f.kind == "overlap"],
            "clipped":   [f for f in self.findings if f.kind == "clipped"],
        }

        for kind, items in checks.items():
            if not items:
                lines.append(f"  OK     no {kind} widgets")
            else:
                for f in items:
                    lines.append(f"  {f.severity:<6} {f.kind}: {f.widget}  {f.detail}")

        fails = sum(1 for f in self.findings if f.severity == "FAIL")
        warns = sum(1 for f in self.findings if f.severity == "WARN")
        lines.append(f"  -> {fails} FAIL, {warns} WARN")
        return "\n".join(lines)


def run_report(root: tk.Misc, tab: str) -> Report:
    """Build and return a Report for the current state of *root*."""
    root.update_idletasks()
    win_w = root.winfo_width()
    win_h = root.winfo_height()
    root_path = widget_path(root)
    tree = widget_tree(root)

    findings: list[Finding] = []
    findings.extend(check_offscreen(tree, win_w, win_h, root_path))
    findings.extend(check_zero_size(tree, root_path))
    findings.extend(check_overlaps(tree, root_path))
    findings.extend(check_clipped(tree, root_path))

    return Report(
        tab=tab,
        window_w=win_w,
        window_h=win_h,
        widgets=tree,
        findings=findings,
    )
