"""Boot SEEDApplication headless with mocks and drive tab navigation."""

import os
import sys
import time
import threading
from typing import Optional
from unittest.mock import MagicMock, patch


def _apply_tk_headless_patches() -> None:
    """Apply the same Tk display patches as conftest.py when no real X server is available.

    Must be called before any tkinter import that would try to connect to a display.
    """
    import _tkinter
    import tkinter as tk
    import tkinter.ttk as ttk

    _TK_VER = _tkinter.TK_VERSION
    _TCL_VER = _tkinter.TCL_VERSION

    def _make_interp():
        m = MagicMock()
        m.wantobjects.return_value = True
        m.getint.return_value = 0
        m.getdouble.return_value = 0.0
        m.getboolean.return_value = False
        m.splitlist.return_value = ()
        def _getvar(name=None, *a):
            if name == 'tk_version':  return _TK_VER
            if name == 'tcl_version': return _TCL_VER
            return ''
        m.getvar.side_effect = _getvar
        return m

    if not hasattr(_tkinter, '_headless_patched'):
        patch.object(_tkinter, 'create', side_effect=lambda *a, **kw: _make_interp()).start()
        _tkinter._headless_patched = True

    # matplotlib TkAgg
    try:
        import matplotlib.backends.backend_tkagg as _mpl_tk
        if not hasattr(_mpl_tk, '_headless_patched'):
            class _FakeFigureCanvas:
                class _W:
                    def pack(self, *a, **kw): pass
                    def grid(self, *a, **kw): pass
                    def configure(self, *a, **kw): pass
                    def bind(self, *a, **kw): return ''
                    def destroy(self): pass
                    def winfo_width(self): return 500
                    def winfo_height(self): return 400
                def __init__(self, figure=None, master=None, *a, **kw):
                    self._widget = self._W()
                def get_tk_widget(self): return self._widget
                def draw(self): pass
                def draw_idle(self): pass
                def mpl_connect(self, event, cb): return 0
                def mpl_disconnect(self, cid): pass
                def flush_events(self): pass
                def resize(self, *a, **kw): pass
            class _FakeToolbar:
                def __init__(self, *a, **kw): pass
                def update(self): pass
                def pack(self, *a, **kw): pass
                def grid(self, *a, **kw): pass
            _mpl_tk.FigureCanvasTkAgg    = _FakeFigureCanvas
            _mpl_tk.NavigationToolbar2Tk = _FakeToolbar
            _mpl_tk._headless_patched = True
    except ImportError:
        pass

    class _StringVar:
        def __init__(self, master=None, value='', name=None): self._val = str(value or '')
        def get(self): return self._val
        def set(self, v): self._val = str(v or '')
        def trace_add(self, *a, **kw): return 'id'
        def trace_remove(self, *a, **kw): pass
        def trace_info(self): return []
        def trace(self, *a, **kw): return 'id'
        def __str__(self): return self._val

    class _IntVar:
        def __init__(self, master=None, value=0, name=None): self._val = int(value or 0)
        def get(self): return self._val
        def set(self, v): self._val = int(v)
        def trace_add(self, *a, **kw): return 'id'
        def trace_remove(self, *a, **kw): pass
        def trace_info(self): return []
        def trace(self, *a, **kw): return 'id'

    class _BoolVar:
        def __init__(self, master=None, value=False, name=None): self._val = bool(value)
        def get(self): return self._val
        def set(self, v): self._val = bool(v)
        def trace_add(self, *a, **kw): return 'id'
        def trace_remove(self, *a, **kw): pass
        def trace_info(self): return []
        def trace(self, *a, **kw): return 'id'

    class _DoubleVar:
        def __init__(self, master=None, value=0.0, name=None): self._val = float(value or 0.0)
        def get(self): return self._val
        def set(self, v): self._val = float(v)
        def trace_add(self, *a, **kw): return 'id'
        def trace_remove(self, *a, **kw): pass
        def trace_info(self): return []
        def trace(self, *a, **kw): return 'id'

    if not getattr(tk, '_headless_patched', False):
        tk.StringVar  = _StringVar
        tk.IntVar     = _IntVar
        tk.BooleanVar = _BoolVar
        tk.DoubleVar  = _DoubleVar

        def _winfo_toplevel(self):
            w = self
            while getattr(w, 'master', None) is not None:
                w = w.master
            return w
        tk.Misc.winfo_toplevel = _winfo_toplevel

        _canvas_counter = [0]
        def _canvas_create(*a, **kw):
            _canvas_counter[0] += 1
            return _canvas_counter[0]
        def _noop(*a, **kw): return None
        def _noop_str(*a, **kw): return ''

        for _m in ('create_polygon','create_text','create_window','create_rectangle',
                   'create_line','create_oval','create_arc','create_image'):
            setattr(tk.Canvas, _m, _canvas_create)
        tk.Canvas.itemconfig     = _noop
        tk.Canvas.itemconfigure  = _noop
        tk.Canvas.tag_bind       = _noop_str
        tk.Canvas.bbox           = lambda *a, **kw: (0,0,100,100)
        tk.Canvas.yview          = _noop
        tk.Canvas.xview          = _noop
        tk.Canvas.delete         = _noop
        tk.PanedWindow.add       = _noop
        tk.PanedWindow.remove    = _noop

        _entry_store: dict = {}
        _orig_entry_init = tk.Entry.__init__
        def _entry_init(self, master=None, cnf=None, **kw):
            _orig_entry_init(self, master, cnf or {}, **kw)
            _entry_store[id(self)] = ''
        tk.Entry.__init__  = _entry_init
        tk.Entry.delete    = lambda self, *a: _entry_store.__setitem__(id(self), '')
        tk.Entry.insert    = lambda self, i, s: _entry_store.__setitem__(id(self), str(s))
        tk.Entry.get       = lambda self: _entry_store.get(id(self), '')

        _nb_store: dict = {}
        _orig_nb_init = ttk.Notebook.__init__
        def _nb_init(self, master=None, **kw):
            _orig_nb_init(self, master, **kw)
            _nb_store[id(self)] = {'tabs': [], 'texts': {}}
        def _nb_add(self, child, **kw):
            d = _nb_store.setdefault(id(self), {'tabs': [], 'texts': {}})
            idx = len(d['tabs'])
            d['tabs'].append(child)
            d['texts'][idx] = kw.get('text', '')
        def _nb_tabs(self):
            return list(range(len(_nb_store.get(id(self), {'tabs': []})['tabs'])))
        def _nb_tab(self, index, option=None, **kw):
            d = _nb_store.get(id(self), {'tabs': [], 'texts': {}})
            if option == 'text':
                return d['texts'].get(index, '')
            return {}
        def _nb_select(self, tab_id=None):
            d = _nb_store.get(id(self), {'tabs': [], 'texts': {}})
            if tab_id is None:
                first = d['tabs'][0] if d['tabs'] else MagicMock()
                return getattr(first, '_w', '')
            return None
        ttk.Notebook.__init__ = _nb_init
        ttk.Notebook.add      = _nb_add
        ttk.Notebook.tabs     = _nb_tabs
        ttk.Notebook.tab      = _nb_tab
        ttk.Notebook.select   = _nb_select

        _combo_store: dict = {}
        _orig_combo_init = ttk.Combobox.__init__
        def _combo_init(self, master=None, **kw):
            _orig_combo_init(self, master, **kw)
            vals = list(kw.get('values', ()))
            _combo_store[id(self)] = {'val': vals[0] if vals else '', 'values': vals}
        def _combo_current(self, index=None):
            if index is not None:
                d = _combo_store.setdefault(id(self), {'val': '', 'values': []})
                if 0 <= index < len(d['values']):
                    d['val'] = d['values'][index]
                return None
            return 0
        ttk.Combobox.__init__   = _combo_init
        ttk.Combobox.get        = lambda self: _combo_store.get(id(self), {}).get('val', '')
        ttk.Combobox.set        = lambda self, v: _combo_store.setdefault(id(self), {'val':'','values':[]}).update({'val': str(v)})
        ttk.Combobox.current    = _combo_current

        tk._headless_patched = True


def _has_real_display() -> bool:
    """Return True only if a usable X server is reachable on DISPLAY.

    Definitive probe: actually try to create (and destroy) a real Tk root.
    Heuristics like xdpyinfo/socket-poking gave false negatives (e.g. when
    xdpyinfo is not installed or Xvfb uses a different socket path), which
    silently dropped the harness into mock-Tk mode and produced meaningless
    "all-green" reports. Creating a real Tk is the only reliable test, and it
    must happen before any mock patches replace ``_tkinter.create``.
    """
    if not os.environ.get("DISPLAY"):
        return False
    import tkinter as tk
    try:
        probe = tk.Tk()
    except Exception:
        return False
    try:
        # If _tkinter.create has been patched (by our headless patches OR by the
        # test suite's conftest.py), probe.tk is a MagicMock, not a real Tcl
        # interpreter — that is NOT a real display and must not be treated as one.
        is_real = type(probe.tk).__module__ != "unittest.mock"
    except Exception:
        is_real = False
    try:
        probe.destroy()
    except Exception:
        pass
    return is_real

# ── Path setup ──────────────────────────────────────────────────────────────
_tools_dir = os.path.dirname(os.path.abspath(__file__))
_sc_root = os.path.dirname(os.path.dirname(_tools_dir))  # seed-control-center/
_repo_root = os.path.dirname(_sc_root)

for _p in [
    os.path.join(_sc_root, "src"),
    os.path.join(_repo_root, "legacy", "gui", "src"),
    os.path.join(_repo_root, "legacy"),
]:
    if _p not in sys.path:
        sys.path.insert(0, _p)

# Tab key → notebook index (0-based top-level tabs in SEEDApplication order)
TAB_INDEX = {
    "pattern":         0,
    "manual":          1,
    "sender":          2,
    "vivigo":          3,
    "vivigo-controls": 3,  # Vivigo tab, default (first) subtab
    "pop-viz":         3,  # Vivigo tab, second subtab
}

# For Vivigo subtabs: key → subtab index within VivigoPanel
VIVIGO_SUBTAB = {
    "vivigo-controls": 0,
    "pop-viz":         1,
}

ALL_TAB_KEYS = ["pattern", "manual", "sender", "vivigo-controls", "pop-viz"]


def _apply_hardware_mocks() -> list:
    """Patch serial, VISA, and other hardware entry-points. Returns active patches."""
    patches = []

    def _p(target, **kw):
        p = patch(target, **kw)
        p.start()
        patches.append(p)

    # matplotlib: standalone plt.figure/subplots calls (legacy gui.py) fail under Xvfb
    # because FigureManagerTk.wm_frame() returns a non-hex string with no window manager.
    # FigureCanvasTkAgg is also mocked so embedded canvas init doesn't fail on the fake fig.
    try:
        import matplotlib.pyplot as _plt
        _fake_line = MagicMock()
        _fake_ax = MagicMock()
        _fake_ax.plot.return_value = [_fake_line]   # supports `x, = ax.plot(...)`
        _fake_ax.twinx.return_value = _fake_ax
        _fake_fig = MagicMock()
        _fake_fig.add_subplot.return_value = _fake_ax
        _p("matplotlib.pyplot.figure", return_value=_fake_fig)
        _p("matplotlib.pyplot.subplots", return_value=(_fake_fig, _fake_ax))
    except Exception:
        pass
    try:
        class _FakeCanvas:
            class _W:
                def pack(self, *a, **kw): pass
                def grid(self, *a, **kw): pass
                def configure(self, *a, **kw): pass
                def bind(self, *a, **kw): return ''
                def destroy(self): pass
                def winfo_width(self): return 500
                def winfo_height(self): return 400
            def __init__(self, figure=None, master=None, *a, **kw):
                self._widget = self._W()
            def get_tk_widget(self): return self._widget
            def draw(self): pass
            def draw_idle(self): pass
            def mpl_connect(self, event, cb): return 0
            def mpl_disconnect(self, cid): pass
            def flush_events(self): pass
            def resize(self, *a, **kw): pass
            def get_width_height(self, physical=False): return (500, 400)
        class _FakeToolbar:
            def __init__(self, *a, **kw): pass
            def update(self): pass
            def pack(self, *a, **kw): pass
            def grid(self, *a, **kw): pass
        _p("matplotlib.backends.backend_tkagg.FigureCanvasTkAgg", new=_FakeCanvas)
        _p("matplotlib.backends.backend_tkagg.NavigationToolbar2Tk", new=_FakeToolbar)
        _p("matplotlib.backends._backend_tk.FigureCanvasTk", new=_FakeCanvas)
    except Exception:
        pass

    # Serial
    _p("serial.Serial", new_callable=lambda: lambda *a, **kw: MagicMock())
    _p("serial.tools.list_ports.comports", return_value=[])

    # VISA / pyvisa
    try:
        _p("pyvisa.ResourceManager", new_callable=lambda: lambda *a, **kw: MagicMock())
    except Exception:
        pass

    # usb / hid (Vivigo hub)
    try:
        _p("hid.device", new_callable=lambda: lambda *a, **kw: MagicMock())
    except Exception:
        pass
    try:
        _p("usb.core.find", return_value=None)
    except Exception:
        pass

    return patches


class HarnessSession:
    """A running SEEDApplication under Xvfb.

    Lifecycle::

        with HarnessSession() as s:
            s.select_tab("sender")
            s.settle()

    The app is destroyed on ``__exit__``.
    """

    def __init__(self, geometry: str = "1100x800"):
        self.geometry = geometry
        self.root = None
        self.app = None
        self._patches: list = []
        self._thread: Optional[threading.Thread] = None
        self._ready = threading.Event()
        self._error: Optional[Exception] = None

    def __enter__(self) -> "HarnessSession":
        self._start()
        return self

    def __exit__(self, *_) -> None:
        self.close()

    # ── Public API ───────────────────────────────────────────────────────────

    def select_tab(self, key: str) -> None:
        """Switch the notebook to the tab identified by *key*."""
        idx = TAB_INDEX.get(key)
        if idx is None:
            raise ValueError(f"Unknown tab key: {key!r}. Valid keys: {list(TAB_INDEX)}")

        notebook = self.app.notebook
        tabs = notebook.tabs()
        if idx >= len(tabs):
            raise RuntimeError(f"Tab index {idx} out of range (only {len(tabs)} tabs)")

        notebook.select(tabs[idx])
        self.root.event_generate("<<NotebookTabChanged>>")
        self.settle()

        # Navigate to Vivigo subtab if requested
        if key in VIVIGO_SUBTAB:
            subtab_idx = VIVIGO_SUBTAB[key]
            vivigo = self.app.vivigo_panel
            if hasattr(vivigo, "notebook"):
                sub_tabs = vivigo.notebook.tabs()
                if subtab_idx < len(sub_tabs):
                    vivigo.notebook.select(sub_tabs[subtab_idx])
                    vivigo.notebook.event_generate("<<NotebookTabChanged>>")
                    self.settle()

    def settle(self, ms: int = 200) -> None:
        """Pump the event loop and let deferred renders complete."""
        self.root.update_idletasks()
        self.root.update()
        time.sleep(ms / 1000)
        self.root.update_idletasks()

    def window_size(self) -> tuple[int, int]:
        return self.root.winfo_width(), self.root.winfo_height()

    # ── Internal ─────────────────────────────────────────────────────────────

    def _start(self) -> None:
        # Apply Tk headless patches before any Tk imports if no real display
        if not _has_real_display():
            _apply_tk_headless_patches()

        self._patches = _apply_hardware_mocks()

        import customtkinter as ctk
        # Import from seed-control-center/src explicitly to avoid legacy/gui/src/main.py
        import importlib.util as _ilu
        _spec = _ilu.spec_from_file_location(
            "seed_main",
            os.path.join(_sc_root, "src", "main.py"),
        )
        _mod = _ilu.module_from_spec(_spec)
        _spec.loader.exec_module(_mod)
        SEEDApplication = _mod.SEEDApplication

        self.root = ctk.CTk()
        try:
            self.root.geometry(self.geometry)
        except Exception:
            pass

        self.app = SEEDApplication(self.root)
        self.settle(300)

    def close(self) -> None:
        for p in reversed(self._patches):
            try:
                p.stop()
            except Exception:
                pass
        self._patches = []

        if self.root is not None:
            try:
                self.root.destroy()
            except Exception:
                pass
            self.root = None
        self.app = None


def boot(geometry: str = "1100x800") -> HarnessSession:
    """Create and start a HarnessSession. Caller must call `.close()` when done."""
    session = HarnessSession(geometry=geometry)
    session._start()
    return session
