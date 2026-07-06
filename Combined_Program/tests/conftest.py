import sys
import os
import pytest
from unittest.mock import MagicMock

# ── Path setup ─────────────────────────────────────────────────────────────────
_tests_dir = os.path.dirname(os.path.abspath(__file__))
_sc_root = os.path.dirname(_tests_dir)
_repo_root = os.path.dirname(_sc_root)

for _p in [
    os.path.join(_sc_root, "src"),
    os.path.join(_repo_root, "legacy", "gui", "src"),
    os.path.join(_repo_root, "legacy"),
]:
    if _p not in sys.path:
        sys.path.insert(0, _p)


# ── Headless-display detection ─────────────────────────────────────────────────
def _has_display() -> bool:
    try:
        import _tkinter
        # AF_UNIX is blocked in this sandbox; real check is socket creation
        import socket
        socket.socket(socket.AF_UNIX, socket.SOCK_STREAM).close()
        return True
    except Exception:
        return False


_DISPLAY_AVAILABLE = _has_display()

if not _DISPLAY_AVAILABLE:
    # ── Step 1: Patch _tkinter.create so tk.Tk() doesn't need X11 ─────────────
    import _tkinter
    from unittest.mock import patch

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

    patch.object(_tkinter, 'create', side_effect=lambda *a, **kw: _make_interp()).start()

    # ── Step 2: Patch matplotlib Tk backend BEFORE manual_movement_panel imports ─
    try:
        import matplotlib.backends.backend_tkagg as _mpl_tk

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
    except ImportError:
        pass

    # ── Step 3: Now import tkinter (Tk() works via the patched _tkinter.create) ──
    import tkinter as tk
    import tkinter.ttk as ttk

    # ── Step 4: Pure-Python Var replacements ────────────────────────────────────
    class _StringVar:
        def __init__(self, master=None, value='', name=None):
            self._val = str(value) if value is not None else ''
        def get(self): return self._val
        def set(self, v): self._val = str(v) if v is not None else ''
        def trace_add(self, *a, **kw): return 'id'
        def trace_remove(self, *a, **kw): pass
        def trace_info(self): return []
        def trace(self, *a, **kw): return 'id'
        def __str__(self): return self._val

    class _IntVar:
        def __init__(self, master=None, value=0, name=None):
            self._val = int(value) if value is not None else 0
        def get(self): return self._val
        def set(self, v): self._val = int(v)
        def trace_add(self, *a, **kw): return 'id'
        def trace_remove(self, *a, **kw): pass
        def trace_info(self): return []
        def trace(self, *a, **kw): return 'id'

    class _BoolVar:
        def __init__(self, master=None, value=False, name=None):
            self._val = bool(value)
        def get(self): return self._val
        def set(self, v): self._val = bool(v)
        def trace_add(self, *a, **kw): return 'id'
        def trace_remove(self, *a, **kw): pass
        def trace_info(self): return []
        def trace(self, *a, **kw): return 'id'

    class _DoubleVar:
        def __init__(self, master=None, value=0.0, name=None):
            self._val = float(value) if value is not None else 0.0
        def get(self): return self._val
        def set(self, v): self._val = float(v)
        def trace_add(self, *a, **kw): return 'id'
        def trace_remove(self, *a, **kw): pass
        def trace_info(self): return []
        def trace(self, *a, **kw): return 'id'

    tk.StringVar  = _StringVar
    tk.IntVar     = _IntVar
    tk.BooleanVar = _BoolVar
    tk.DoubleVar  = _DoubleVar

    # ── Step 5: Fix winfo_toplevel to walk master chain, not call Tcl ───────────
    def _winfo_toplevel(self):
        w = self
        while getattr(w, 'master', None) is not None:
            w = w.master
        return w

    tk.Misc.winfo_toplevel = _winfo_toplevel

    # ── Step 6: Canvas create_* return integer IDs; fix other canvas ops ────────
    _canvas_counter = [0]
    def _canvas_create(*a, **kw):
        _canvas_counter[0] += 1
        return _canvas_counter[0]

    def _noop(*a, **kw): return None
    def _noop_str(*a, **kw): return ''

    tk.Canvas.create_polygon   = _canvas_create
    tk.Canvas.create_text      = _canvas_create
    tk.Canvas.create_window    = _canvas_create
    tk.Canvas.create_rectangle = _canvas_create
    tk.Canvas.create_line      = _canvas_create
    tk.Canvas.create_oval      = _canvas_create
    tk.Canvas.create_arc       = _canvas_create
    tk.Canvas.create_image     = _canvas_create
    tk.Canvas.itemconfig       = _noop
    tk.Canvas.itemconfigure    = _noop
    tk.Canvas.tag_bind         = _noop_str
    tk.Canvas.bbox             = lambda *a, **kw: (0, 0, 100, 100)
    tk.Canvas.yview            = _noop
    tk.Canvas.xview            = _noop
    tk.Canvas.delete           = _noop

    # PanedWindow.add doesn't need to do real geometry tracking
    tk.PanedWindow.add    = _noop
    tk.PanedWindow.remove = _noop

    # ── Step 7: Entry — real Python-backed get/insert/delete ────────────────────
    _entry_store: dict = {}

    _orig_entry_init = tk.Entry.__init__

    def _entry_init(self, master=None, cnf=None, **kw):
        _orig_entry_init(self, master, cnf or {}, **kw)
        _entry_store[id(self)] = ''

    def _entry_delete(self, first=0, last=None):
        _entry_store[id(self)] = ''

    def _entry_insert(self, index, string):
        _entry_store[id(self)] = str(string)

    def _entry_get(self):
        return _entry_store.get(id(self), '')

    tk.Entry.__init__ = _entry_init
    tk.Entry.delete   = _entry_delete
    tk.Entry.insert   = _entry_insert
    tk.Entry.get      = _entry_get

    # ── Step 8: ttk.Notebook — track add() calls for proper tabs()/tab() ────────
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
        d = _nb_store.get(id(self), {'tabs': [], 'texts': {}})
        return list(range(len(d['tabs'])))

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

    # ── Step 9: ttk.Combobox — real get/set ─────────────────────────────────────
    _combo_store: dict = {}

    _orig_combo_init = ttk.Combobox.__init__

    def _combo_init(self, master=None, **kw):
        _orig_combo_init(self, master, **kw)
        vals = list(kw.get('values', ()))
        _combo_store[id(self)] = {'val': vals[0] if vals else '', 'values': vals}

    def _combo_get(self):
        return _combo_store.get(id(self), {}).get('val', '')

    def _combo_set(self, value):
        _combo_store.setdefault(id(self), {'val': '', 'values': []})['val'] = str(value)

    def _combo_current(self, index=None):
        if index is not None:
            d = _combo_store.setdefault(id(self), {'val': '', 'values': []})
            if 0 <= index < len(d['values']):
                d['val'] = d['values'][index]
            return None
        return 0

    ttk.Combobox.__init__ = _combo_init
    ttk.Combobox.get      = _combo_get
    ttk.Combobox.set      = _combo_set
    ttk.Combobox.current  = _combo_current
