import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import matplotlib
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from mpl_toolkits.axes_grid1 import make_axes_locatable
import numpy as np
import utils
import json
import os
import sys
import pandas as pd
import csv

from tkinter import simpledialog
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.colors import ListedColormap
import datetime
import itertools

class StatusIndicator(tk.Canvas):
    """
    A circular visual status indicator (LED-style) for real-time state feedback.
    Supports three states: 'idle' (dim), 'busy' (blue), and 'on' (green).
    """
    def __init__(self, parent, bg_color, size=20):
        super().__init__(parent, width=size, height=size, bg=bg_color, highlightthickness=0)
        self.size = size
        self.padding = 4
        self.circle = self.create_oval(self.padding, self.padding, 
                                      size-self.padding, size-self.padding, 
                                      fill="#333333", outline="#111111")
        
    def set_status(self, state):
        colors = {
            "idle": "#333333", "busy": "#00d2ff", "on": "#00ff00",
            "error": "#ff3333", "warn": "#ffcc00", "off": "#333333"
        }
        color = colors.get(state, "#333333")
        self.itemconfig(self.circle, fill=color)


class SliceSelectionDialog:
    """
    Modal dialog for selecting which slicer variable values to include in a
    multi-slice PDF report.  Returns {field: [selected_float_values]} or None.
    """
    def __init__(self, parent, slider_vars):
        self.result = None
        self._parent = parent

        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Select Export Slices")
        self.dialog.configure(bg=utils.COLOR_PANEL_BG)
        self.dialog.resizable(True, True)
        self.dialog.transient(parent.winfo_toplevel())
        self.dialog.grab_set()

        # Center on parent
        parent.winfo_toplevel().update_idletasks()
        px, py = parent.winfo_toplevel().winfo_x(), parent.winfo_toplevel().winfo_y()
        pw, ph = parent.winfo_toplevel().winfo_width(), parent.winfo_toplevel().winfo_height()
        w, h = 480, min(80 + len(slider_vars) * 220, 700)
        self.dialog.geometry(f"{w}x{h}+{max(0, px + (pw - w)//2)}+{max(0, py + (ph - h)//2)}")

        tk.Label(self.dialog, text="SELECT EXPORT SLICES",
                 font=('Rajdhani', 14, 'bold'),
                 fg=utils.COLOR_ACCENT_CYAN, bg=utils.COLOR_PANEL_BG).pack(pady=(16, 4))
        tk.Label(self.dialog, text="Choose which values to include for each slicer variable.",
                 font=('Inter', 9), fg=utils.COLOR_TEXT_SECONDARY,
                 bg=utils.COLOR_PANEL_BG).pack(pady=(0, 10))

        scroll_outer = tk.Frame(self.dialog, bg=utils.COLOR_PANEL_BG)
        scroll_outer.pack(fill=tk.BOTH, expand=True, padx=16, pady=(0, 8))

        canvas = tk.Canvas(scroll_outer, bg=utils.COLOR_PANEL_BG, highlightthickness=0)
        sb = ttk.Scrollbar(scroll_outer, orient=tk.VERTICAL, command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        inner = tk.Frame(canvas, bg=utils.COLOR_PANEL_BG)
        canvas.create_window((0, 0), window=inner, anchor='nw')
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        self._check_vars = {}  # {field: [BooleanVar, ...]}

        unit_map = {'x': 'mm', 'y': 'mm', 'z': 'mm', 'rot': '°', 'e': '°', 'rotation': '°'}

        for field, (var, vals, _, _) in slider_vars.items():
            unit = unit_map.get(field.lower(), '')
            frame = tk.LabelFrame(inner, text=f"  {field.upper()}  ",
                                  font=('Rajdhani', 10, 'bold'),
                                  fg=utils.COLOR_ACCENT_CYAN, bg=utils.COLOR_PANEL_BG,
                                  padx=8, pady=6)
            frame.pack(fill=tk.X, padx=4, pady=6)

            btn_row = tk.Frame(frame, bg=utils.COLOR_PANEL_BG)
            btn_row.pack(fill=tk.X, pady=(0, 4))

            # Smart Selection: Default to ONLY checking what the user is currently observing
            current_obs_val = var.get()
            bvars = [tk.BooleanVar(value=(abs(v - current_obs_val) < 1e-4)) for v in vals]
            self._check_vars[field] = (bvars, vals)

            def _all(bvs=bvars): [v.set(True) for v in bvs]
            def _none(bvs=bvars): [v.set(False) for v in bvs]

            tk.Button(btn_row, text="All", width=5, font=('Inter', 8),
                      bg=utils.COLOR_BORDER, fg=utils.COLOR_TEXT_PRIMARY,
                      relief=tk.FLAT, command=_all).pack(side=tk.LEFT, padx=(0, 4))
            tk.Button(btn_row, text="None", width=5, font=('Inter', 8),
                      bg=utils.COLOR_BORDER, fg=utils.COLOR_TEXT_PRIMARY,
                      relief=tk.FLAT, command=_none).pack(side=tk.LEFT)

            # Sub-frame for checkboxes to avoid pack/grid conflict on 'frame'
            cb_frame = tk.Frame(frame, bg=utils.COLOR_PANEL_BG)
            cb_frame.pack(fill=tk.X, pady=2)

            cols = 4
            for i, (bv, val) in enumerate(zip(bvars, vals)):
                label = f"{val:.4g}{unit}"
                cb = tk.Checkbutton(cb_frame, text=label, variable=bv,
                                    font=('JetBrains Mono', 9),
                                    fg=utils.COLOR_TEXT_PRIMARY, bg=utils.COLOR_PANEL_BG,
                                    selectcolor=utils.COLOR_BG, activebackground=utils.COLOR_PANEL_BG,
                                    activeforeground=utils.COLOR_ACCENT_CYAN)
                cb.grid(row=i // cols, column=i % cols, sticky='w', padx=6, pady=1)

        # OK / Cancel
        btn_frame = tk.Frame(self.dialog, bg=utils.COLOR_PANEL_BG)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="Generate", command=self._on_ok, width=15).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="Cancel", command=self._on_cancel, width=15).pack(side=tk.LEFT, padx=10)

        self.dialog.bind("<Return>", lambda e: self._on_ok())
        self.dialog.bind("<Escape>", lambda e: self._on_cancel())

    def _on_ok(self):
        selected = {}
        for field, (bvars, vals) in self._check_vars.items():
            chosen = [v for bv, v in zip(bvars, vals) if bv.get()]
            if chosen:
                selected[field] = chosen
        if not selected:
            messagebox.showwarning("No Selection",
                                   "Select at least one value per slicer variable.",
                                   parent=self.dialog)
            return
        self.result = selected
        self.dialog.destroy()

    def _on_cancel(self):
        self.dialog.destroy()

    def show(self):
        self._parent.winfo_toplevel().wait_window(self.dialog)
        return self.result


class POPVisualizationPanel(ttk.Frame):
    """
    A GUI panel for visualizing Proof of Power (POP) data in 2D contour plots.
    Source-Driven Architecture: Commanded positions (IVs) and Hub telemetry (DVs)
    are treated as distinct data sources.
    """
    # Comprehensive Key Map for HUD
    # (Maps raw RLM API keys, DMM channels, and user mode names to HUD labels)
    HUD_KEY_MAP = {
        'rotation': 'rot', 'e': 'rot', 'rot': 'rot',
        'battery': 'v_batt', 'v_batt': 'v_batt',
        'v_rec': 'v_rec', 'v_sys': 'v_sys', 'v_bst': 'v_bst',
        'i_sys': 'i_sys', 'i_bst': 'i_bst',
        't_bst': 't_bst', 'temp_bst': 't_bst', 'temp': 't_bst',
        'dmm_1': 'dmm_1', 'dmm_2': 'dmm_2', 'dmm_3': 'dmm_3', 'dmm_4': 'dmm_4'
    }

    def __init__(self, parent, controller=None):
        super().__init__(parent, style='Dark.TFrame')
        self.c = controller # controller is the VivigoPanel
        
        # Access EventBus via the sender panel if available
        self.event_bus = getattr(self.c, 'event_bus', None)
        if self.event_bus:
            self.event_bus.subscribe("scan_position", self._on_scan_position)
            self.event_bus.subscribe("hub_telemetry", self._on_hub_telemetry)
            self.event_bus.subscribe("gated_measurement", self._on_gated_measurement)
            self.event_bus.subscribe("scan_finished", self._on_scan_finished)
            self.event_bus.subscribe("scan_started", self._on_scan_started)
        
        self.data = []
        self.live_session_data = [] # Persistent background buffer for current hardware session
        self.fields = set() # Use set for efficient superset checks
        self.iv_fields = set()
        self.dv_fields = set()
        self.current_iv_data = {}
        
        self.slider_vars = {}  # {field: (tk.DoubleVar, [vals], scale_widget, combo_widget)}
        self.dv_vars = {}      # {field: tk.BooleanVar}
        self.plot_widgets = {} # {field: canvas_widget}
        self.numeric_fields = []
        self.current_filepath = None
        self.distance_vars = ['x', 'y', 'z']
        self.current_slice_str = ""
        self.fixed_bounds = {} # {x: (min, max), y: (min, max)}
        self._live_axis_extents = {} # Running min/max of IV axes seen in current live scan
        
        # Custom Colormap: Slice bottom 20% of inferno (starts at mid-purple)
        self.custom_inferno = ListedColormap(matplotlib.colormaps['inferno'](np.linspace(0.2, 1.0, 256)))

        # Live Mode State
        self.live_mode = tk.BooleanVar(value=False)
        self.follow_probe = tk.BooleanVar(value=True)
        self.live_tick_id = None
        self.needs_redraw = False
        self.last_plot_params = {} # {field: (data_len, x_col, y_col, z_col, slice_str)}
        self._in_snap = False # Flag to prevent feedback loops during slider snapping
        self.render_mode = 'scatter'  # 'scatter' | 'gradient'
        self.render_style = 'contourf'  # 'contourf' | 'gouraud' | 'flat'
        self.show_topo_overlay = tk.BooleanVar(value=False)
        self.topo_levels_var = tk.IntVar(value=10)
        self.cmap_name = 'jet'
        self.topo_color = '#ffffff'
        self.topo_linewidth = 0.6

        self._setup_ui()

    def _create_hud(self):
        """Creates a high-visibility HUD for real-time telemetry - 2 Row Layout."""
        self.hud_frame = tk.Frame(self, bg=utils.COLOR_BLACK)
        self.hud_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        
        self.hud_labels = {} # {field: (label_widget, unit)}
        
        # Row 1: Prime IVs and basic health
        row1 = tk.Frame(self.hud_frame, bg=utils.COLOR_BLACK)
        row1.pack(fill=tk.X, expand=True)
        
        # Define HUD fields: (Label, Unit, Key)
        hud1_configs = [
            ('X-POS', 'mm', 'x'),
            ('Y-POS', 'mm', 'y'),
            ('Z-POS', 'mm', 'z'),
            ('ROT', '°', 'rot'),
            ('RSSI', ' dBm', 'rssi'),
            ('BATT', 'V', 'v_batt')
        ]
        
        for label, unit, key in hud1_configs:
            f = tk.Frame(row1, bg=utils.COLOR_BLACK)
            f.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
            tk.Label(f, text=label, font=('Rajdhani', 9, 'bold'), fg=utils.COLOR_TEXT_SECONDARY, bg=utils.COLOR_BLACK).pack()
            val_label = tk.Label(f, text="---", font=('JetBrains Mono', 12, 'bold'), fg=utils.COLOR_ACCENT_CYAN, bg=utils.COLOR_BLACK)
            val_label.pack()
            self.hud_labels[key] = (val_label, unit)

        # Row 2: Hub Telemetry (DVs)
        row2 = tk.Frame(self.hud_frame, bg=utils.COLOR_BLACK)
        row2.pack(fill=tk.X, expand=True, pady=(5,0))
        
        hud2_configs = [
            ('V-REC', 'V', 'v_rec'),
            ('V-SYS', 'V', 'v_sys'),
            ('V-BST', 'V', 'v_bst'),
            ('I-SYS', 'A', 'i_sys'),
            ('I-BST', 'A', 'i_bst'),
            ('T-BST', '°C', 't_bst')
        ]
        
        for label, unit, key in hud2_configs:
            f = tk.Frame(row2, bg=utils.COLOR_BLACK)
            f.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
            tk.Label(f, text=label, font=('Rajdhani', 9, 'bold'), fg=utils.COLOR_TEXT_SECONDARY, bg=utils.COLOR_BLACK).pack()
            val_label = tk.Label(f, text="---", font=('JetBrains Mono', 11, 'bold'), fg=utils.COLOR_ACCENT_GREEN, bg=utils.COLOR_BLACK)
            val_label.pack()
            self.hud_labels[key] = (val_label, unit)

        # Row 3: DMM / External Telemetry
        row3 = tk.Frame(self.hud_frame, bg=utils.COLOR_BLACK)
        row3.pack(fill=tk.X, expand=True, pady=(2,0))
        
        hud3_configs = [
            ('DMM-1', 'V', 'dmm_1'),
            ('DMM-2', 'V', 'dmm_2'),
            ('DMM-3', 'V', 'dmm_3'),
            ('DMM-4', 'V', 'dmm_4')
        ]
        
        for label, unit, key in hud3_configs:
            f = tk.Frame(row3, bg=utils.COLOR_BLACK)
            f.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
            tk.Label(f, text=label, font=('Rajdhani', 8, 'bold'), fg=utils.COLOR_TEXT_SECONDARY, bg=utils.COLOR_BLACK).pack()
            val_label = tk.Label(f, text="---", font=('JetBrains Mono', 10, 'bold'), fg=utils.COLOR_ACCENT_CYAN, bg=utils.COLOR_BLACK)
            val_label.pack()
            self.hud_labels[key] = (val_label, unit)

    def _update_hud(self, data):
        """Updates HUD labels with new data."""
        if not hasattr(self, 'hud_labels'): return
        
        for field, value in data.items():
            key = self.HUD_KEY_MAP.get(field.lower(), field.lower())
            
            if key in self.hud_labels:
                label_widget, unit = self.hud_labels[key]
                try:
                    if isinstance(value, (int, float)):
                        text = f"{value:.1f}{unit}" if abs(value) > 100 else f"{value:.2f}{unit}"
                    else:
                        text = f"{value}{unit}"
                    label_widget.config(text=text)
                except:
                    label_widget.config(text=f"{value}{unit}")

    def _update_live_status_ui(self):
        """Updates the visual indicator for live vs paused state."""
        if not hasattr(self, 'live_status_indicator'): return
        
        is_live = self.live_mode.get() and self.current_filepath == "LIVE_SESSION"
        if is_live:
            self.live_status_indicator.set_status("on")
            self.live_status_label.config(text="LIVE", foreground=utils.COLOR_ACCENT_GREEN)
            self.sync_toggle_btn.config(text="🛑 Stop Live Sync", style='Accent.TButton')
        else:
            self.live_status_indicator.set_status("idle")
            self.live_status_label.config(text="PAUSED", foreground=utils.COLOR_TEXT_SECONDARY)
            self.sync_toggle_btn.config(text="📡 Sync Live", style='TButton')

    def update_fixed_bounds(self, profile_data):
        """Sets plot limits based on G-code scan profile metadata for all axes."""
        self.fixed_bounds = {}
        try:
            for ax in ['x', 'y', 'z', 'rot']:
                min_key, max_key = f'{ax}_min', f'{ax}_max'
                if min_key in profile_data and max_key in profile_data:
                    v_min, v_max = float(profile_data[min_key]), float(profile_data[max_key])
                    
                    # Add a small margin (5%) for visual clarity
                    v_range = v_max - v_min
                    margin = v_range * 0.05 if v_range > 0 else 5
                    self.fixed_bounds[ax] = (v_min - margin, v_max + margin)
            
            # Map aliases for flexible axis selection in UI
            if 'rot' in self.fixed_bounds: 
                self.fixed_bounds['rotation'] = self.fixed_bounds['rot']
                self.fixed_bounds['e'] = self.fixed_bounds['rot']
            
            self.needs_redraw = True
        except (ValueError, TypeError):
            self.fixed_bounds = None

    def _apply_viewport_lock(self, ax, x_col, y_col):
        """Standardized logic to lock axis limits and aspect ratio."""
        xl, yl = x_col.lower(), y_col.lower()
        if self.fixed_bounds:
            if xl in self.fixed_bounds: ax.set_xlim(self.fixed_bounds[xl])
            if yl in self.fixed_bounds: ax.set_ylim(self.fixed_bounds[yl])
            ax.set_autoscale_on(False)
            ax.autoscale(False)
        
        # Use 'adjustable=box' to force axes to maintain limits when aspect=equal
        ax.set_aspect('equal' if xl in self.distance_vars and yl in self.distance_vars else 'auto', 
                      adjustable='box')

    def _get_filtered_data(self):
        """Returns the current dataset filtered by the slicer variables."""
        filtered, slice_info = self.data, []
        for field, (var, _, _, _) in self.slider_vars.items():
            val = var.get()
            filtered = [d for d in filtered if d.get(field) is not None and abs(d[field] - val) < 1e-4]
            unit = "mm" if field.lower() in self.distance_vars else ("°" if field.upper() == 'ROT' else "")
            slice_info.append(f"{field}={val:.4g}{unit}")
        return filtered, ", ".join(slice_info)

    def _setup_ui(self):
        # Header area
        header = ttk.Frame(self, style='Dark.TFrame')
        header.pack(fill=tk.X, padx=10, pady=10)

        ttk.Label(header, text="POP 4D SPATIAL ANALYSIS",
                  font=('Rajdhani', 14, 'bold'), foreground=utils.COLOR_ACCENT_CYAN).pack(side=tk.LEFT)

        # Live/Paused Indicator
        self.live_status_indicator = StatusIndicator(header, utils.COLOR_BG, size=18)
        self.live_status_indicator.pack(side=tk.LEFT, padx=(20, 5))
        
        self.live_status_label = ttk.Label(header, text="PAUSED", font=('Rajdhani', 10, 'bold'), foreground=utils.COLOR_TEXT_SECONDARY)
        self.live_status_label.pack(side=tk.LEFT)

        # HUD Area
        self._create_hud()

        # Right-aligned buttons
        self.save_btn = ttk.Button(header, text="📊 Generate Report", command=self._save_snapshot)
        self.save_btn.pack(side=tk.RIGHT, padx=2)

        # Render mode toggle (dots ↔ gradient) — hidden during live streaming
        self.render_toggle_btn = ttk.Button(header, text="· Dot View", command=self._toggle_render_mode)
        # Not packed here — shown/hidden by _update_render_toggle_visibility()

        # Topo overlay toggle — only shown in gradient mode
        self.topo_btn = ttk.Button(header, text="⛰ Topo: OFF", command=self._toggle_topo_overlay)
        # Not packed here — shown/hidden by _update_render_toggle_visibility()
        # Plot settings popup (colormap + contour style) — only shown in gradient mode
        self.plot_settings_btn = ttk.Button(header, text="⚙ Settings", command=self._open_plot_settings)
        # Not packed here — shown/hidden by _update_render_toggle_visibility()

        # Consolidated Sync Toggle
        self.sync_toggle_btn = ttk.Button(header, text="📡 Sync Live", command=self._toggle_live_sync)
        self.sync_toggle_btn.pack(side=tk.RIGHT, padx=2)

        self.load_btn = ttk.Button(header, text="📁 Load File", command=self._load_data)
        self.load_btn.pack(side=tk.RIGHT, padx=2)

        # Main Layout
        self.main_container = ttk.Frame(self, style='Dark.TFrame')
        self.main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Sidebar
        self.sidebar = ttk.Frame(self.main_container, style='Panel.TFrame', width=300)
        self.sidebar.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        self.sidebar.pack_propagate(False)

        # 1. Axis Selection
        axis_frame = ttk.LabelFrame(self.sidebar, text="PLANE SELECTION (IVs)", style='Card.TLabelframe')
        axis_frame.pack(fill=tk.X, padx=5, pady=5)

        ttk.Label(axis_frame, text="X-Axis:", font=utils.FONT_BODY_SMALL).grid(row=0, column=0, sticky='w')
        self.x_axis_var = tk.StringVar()
        self.x_axis_combo = ttk.Combobox(axis_frame, textvariable=self.x_axis_var, state="readonly")
        self.x_axis_combo.grid(row=0, column=1, sticky='ew', pady=2)
        self.x_axis_var.trace_add("write", lambda *args: self._on_axis_change())

        ttk.Label(axis_frame, text="Y-Axis:", font=utils.FONT_BODY_SMALL).grid(row=1, column=0, sticky='w')
        self.y_axis_var = tk.StringVar()
        self.y_axis_combo = ttk.Combobox(axis_frame, textvariable=self.y_axis_var, state="readonly")
        self.y_axis_combo.grid(row=1, column=1, sticky='ew', pady=2)
        self.y_axis_var.trace_add("write", lambda *args: self._on_axis_change())
        
        axis_frame.columnconfigure(1, weight=1)

        # 2. Variable Slicing
        self.slicer_frame = ttk.LabelFrame(self.sidebar, text="DATA SLICING (IVs)", style='Card.TLabelframe')
        self.slicer_frame.pack(fill=tk.X, padx=5, pady=5)
        self.slider_container = ttk.Frame(self.slicer_frame, style='Panel.TFrame')
        self.slider_container.pack(fill=tk.X, pady=5)

        # 3. Dependent Variables
        dv_frame = ttk.LabelFrame(self.sidebar, text="HEATMAP FIELDS (DVs)", style='Card.TLabelframe')
        dv_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        dv_ctrl_frame = ttk.Frame(dv_frame, style='Panel.TFrame')
        dv_ctrl_frame.pack(fill=tk.X, padx=5, pady=2)
        ttk.Button(dv_ctrl_frame, text="All", width=5, command=self._select_all_dvs).pack(side=tk.LEFT, padx=2)
        ttk.Button(dv_ctrl_frame, text="None", width=5, command=self._select_none_dvs).pack(side=tk.LEFT, padx=2)

        self.dv_container = ttk.Frame(dv_frame, style='Panel.TFrame')
        self.dv_container.pack(fill=tk.BOTH, expand=True, pady=5)

        # Plot Area
        self.plot_canvas = tk.Canvas(self.main_container, bg=utils.COLOR_BG, highlightthickness=0)
        self.plot_scrollbar = ttk.Scrollbar(self.main_container, orient=tk.VERTICAL, command=self.plot_canvas.yview)
        self.plot_canvas.configure(yscrollcommand=self.plot_scrollbar.set)
        self.plot_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.plot_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.grid_container = ttk.Frame(self.plot_canvas, style='Dark.TFrame')
        self.plot_canvas.create_window((0,0), window=self.grid_container, anchor='nw', tags="grid_window")
        self.grid_container.bind("<Configure>", self._on_frame_configure)
        self.plot_canvas.bind("<Configure>", self._on_canvas_configure)
        self.bind_all("<MouseWheel>", self._on_mousewheel)
        
        # Initialize Indicator State
        self._update_live_status_ui()

    def _on_frame_configure(self, event):
        self.plot_canvas.configure(scrollregion=self.plot_canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        self.plot_canvas.itemconfig("grid_window", width=event.width)

    def _on_mousewheel(self, event):
        try:
            self.plot_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        except (KeyError, tk.TclError): pass

    def _on_scan_started(self, data):
        """Event: New scan beginning. Clear the background session buffer."""
        self._live_axis_extents = {}
        self.live_session_data = []
        if self.current_filepath == "LIVE_SESSION":
            self.clear_data()
            self.current_filepath = "LIVE_SESSION" # Restore after clear
        # Re-apply scan profile bounds after any clear (data carries relative min/max from sender)
        if data:
            self.update_fixed_bounds(data)

    def _on_scan_position(self, data):
        """IV Source: Commanded positions from the scan engine."""
        if not data: return
        self.current_iv_data.update(data)
        self.iv_fields.update(data.keys())
        
        if self.live_mode.get() and self.follow_probe.get() and self.current_filepath == "LIVE_SESSION":
            self._snap_sliders_to_telemetry(data)
            self.needs_redraw = True
        
        self._update_hud(data)

    def _on_hub_telemetry(self, data):
        """DV Source: Background telemetry polling. Used for HUD only."""
        if not data: return
        self._update_hud(data)
        pass

    def _on_gated_measurement(self, point):
        """Atomic Source: High-fidelity (Coords + DVs) packet from gated scan stops."""
        if not point: return

        # 1. Clean and Normalize Data Point immediately
        def _to_num(v):
            try: return float(v)
            except (ValueError, TypeError): return v
        
        clean_point = {k: _to_num(v) for k, v in point.items()}

        # 2. Always store in Background Session Buffer
        self.live_session_data.append(clean_point)
        if len(self.live_session_data) > 5000: self.live_session_data = self.live_session_data[-5000:]

        # 3. Only process for UI if currently in Live Mode
        if self.current_filepath != "LIVE_SESSION": return

        # Update Field Sets
        known_ivs = set(self.distance_vars) | {'rot', 'e', 'rotation', 'timestamp', 'Timestamp'}
        hub_dvs = {k for k in clean_point.keys() if k not in known_ivs and k.lower() not in known_ivs}
        self.dv_fields.update(hub_dvs)
        
        self.data.append(clean_point)
        if len(self.data) > 5000: self.data = self.data[-5000:]
        self._update_live_extents(clean_point)

        # Optimized UI Rebuild Check (Case-Insensitive Subset)
        point_keys = set(clean_point.keys())
        rebuild_needed = not point_keys.issubset(self.fields)
        
        # Check for new unique values in existing sliders (e.g. new Z-layer)
        if not rebuild_needed and self.live_mode.get():
            for f, (_, vals, _, _) in self.slider_vars.items():
                if f in clean_point:
                    val = clean_point[f]
                    # If this is a numeric field and the value is new, trigger rebuild
                    if isinstance(val, (int, float)) and not any(abs(v - val) < 1e-4 for v in vals):
                        rebuild_needed = True
                        break

        if rebuild_needed:
            self.fields.update(point_keys)
            self.fields.update(self.iv_fields)
            self.fields.update(self.dv_fields)
            # Rebuild controls only when new physical metrics or new layers are detected
            self._update_ui_fields()
        
        # Atomic Snap: Ensure sliders match the point we just recorded
        if self.live_mode.get() and self.follow_probe.get():
            self._snap_sliders_to_telemetry(clean_point)

        self.needs_redraw = True

    def _snap_sliders_to_telemetry(self, packet):
        """Programmatically move sliders to match machine position without triggering redraws."""
        self._in_snap = True # Block feedback loop
        try:
            for field, (var, vals, scale, combo) in self.slider_vars.items():
                if field in packet:
                    val = packet[field]
                    if vals:
                        idx = (np.abs(np.array(vals) - val)).argmin()
                        closest = vals[idx]
                        if abs(var.get() - closest) > 1e-4:
                            var.set(closest)
                            if combo: combo.set(closest)
        finally:
            self._in_snap = False
            self.needs_redraw = True # Ensure the new slice is drawn

    def _on_slider(self, field, val):
        """User interaction handler for sliders."""
        if self._in_snap: return # Ignore programmatic updates
        
        _, vals, _, combo = self.slider_vars[field]
        snapped = vals[(np.abs(np.array(vals) - float(val))).argmin()]
        self.slider_vars[field][0].set(snapped)
        if combo: combo.set(snapped)
        self._update_all_plots()

    def _update_ui_fields(self):
        """Rebuilds axis and checkbox lists when data structure changes."""
        if not self.data: return
        
        # Determine numeric fields based on available data
        self.numeric_fields = sorted([f for f in self.fields if any(isinstance(d.get(f), (int, float)) for d in self.data)])

        # Axis defaults
        curr_x, curr_y = self.x_axis_var.get(), self.y_axis_var.get()
        if not curr_x or curr_x not in self.iv_fields:
            for f in ['x', 'y', 'z']:
                if f in self.iv_fields: self.x_axis_var.set(f); break
        if not curr_y or curr_y not in self.iv_fields:
            for f in ['y', 'x', 'z']:
                if f in self.iv_fields and f != self.x_axis_var.get(): self.y_axis_var.set(f); break

        self._rebuild_controls()
        self._sync_plots()

    def _rebuild_controls(self):
        for child in self.slider_container.winfo_children(): child.destroy()
        for child in self.dv_container.winfo_children(): child.destroy()
        
        # Define hardcoded IV set to strictly exclude from DVs
        core_ivs = {'x', 'y', 'z', 'rot', 'e', 'rotation', 'timestamp', 'Timestamp'}
        all_iv_set = self.iv_fields | core_ivs
        
        present_ivs = sorted(list(self.iv_fields.intersection(set(self.numeric_fields))))
        self.x_axis_combo['values'] = tuple(present_ivs)
        self.y_axis_combo['values'] = tuple(present_ivs)
        
        curr_x, curr_y = self.x_axis_var.get(), self.y_axis_var.get()
        slice_ivs = [f for f in present_ivs if f != curr_x and f != curr_y]
        
        # Filter DVs to ensure no IVs leak into the checkbox list
        dv_fields = sorted(list(self.dv_fields.difference(all_iv_set).intersection(set(self.numeric_fields))))
        
        self.slider_vars = {}
        for field in slice_ivs:
            vals = sorted(list(set([d[field] for d in self.data if field in d])))
            if not vals: continue
            var = tk.DoubleVar(value=vals[0])
            f_row = ttk.Frame(self.slider_container, style='Panel.TFrame')
            f_row.pack(fill=tk.X, pady=2)
            ttk.Label(f_row, text=f"{field}:", font=utils.FONT_BODY_SMALL, width=6).pack(side=tk.LEFT)
            scale = ttk.Scale(f_row, from_=min(vals), to=max(vals), variable=var, command=lambda v, f=field: self._on_slider(f, v))
            scale.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
            combo = ttk.Combobox(f_row, values=vals, width=8, state="readonly"); combo.set(vals[0]); combo.pack(side=tk.RIGHT)
            combo.bind("<<ComboboxSelected>>", lambda e, f=field, c=combo, v=var: [v.set(float(c.get())), self._sync_plots()])
            self.slider_vars[field] = (var, vals, scale, combo)

        for field in dv_fields:
            if field not in self.dv_vars: self.dv_vars[field] = tk.BooleanVar(value=True)
            cb = ttk.Checkbutton(self.dv_container, text=field, variable=self.dv_vars[field], command=self._sync_plots)
            cb.pack(anchor='w', padx=10)

    def clear_data(self):
        self.data = []; self.fields = set(); self.iv_fields.clear(); self.dv_fields.clear()
        self.dv_vars.clear()
        self.last_plot_params.clear()
        self.current_slice_str = ""
        self.fixed_bounds = {} # Clear spatial limits
        self._live_axis_extents = {}
        for field, (container, canvas, fig, ax) in self.plot_widgets.items():
            ax.clear(); ax.set_facecolor(utils.COLOR_BLACK); canvas.draw()
        self.plot_widgets.clear()
        self._rebuild_controls()

    def _update_live_extents(self, point):
        """Expand running axis extents to include new point (never shrinks)."""
        for key in ['x', 'y', 'z', 'rot', 'e', 'rotation']:
            if key in point and isinstance(point[key], (int, float)):
                val = point[key]
                lo, hi = self._live_axis_extents.get(key, (val, val))
                self._live_axis_extents[key] = (min(lo, val), max(hi, val))
        # Keep rotation aliases in sync
        if 'rot' in self._live_axis_extents:
            self._live_axis_extents['rotation'] = self._live_axis_extents['rot']
            self._live_axis_extents['e'] = self._live_axis_extents['rot']

    def start_live_mode(self):
        """Activates Live Mode and imports all data collected in the current session."""
        was_live = (self.current_filepath == "LIVE_SESSION")
        self.current_filepath = "LIVE_SESSION"; self.live_mode.set(True)
        
        # Import historical data from the session buffer
        if not was_live and self.live_session_data:
            self.data = list(self.live_session_data)
            self.fields = set(self.data[0].keys())
            # Regenerate IV/DV sets based on imported data
            known = ['x', 'y', 'z', 'rot', 'height', 'elevation', 'rotation', 'r', 'e']
            self.iv_fields = {f for f in self.fields if f.lower() in known}
            self.dv_fields = {f for f in self.fields if f.lower() not in known and f != 'Timestamp'}
            # Rebuild axis extents from the full imported buffer
            self._live_axis_extents = {}
            for pt in self.data:
                self._update_live_extents(pt)
            self._update_ui_fields()
        elif was_live:
            # Resuming after pause: clear the draw cache so the first tick always
            # redraws (without this, _draw_heatmap's param-equality check skips
            # the redraw when the slice/data-count haven't changed since the pause).
            self.last_plot_params.clear()
            self.needs_redraw = True
            if self.current_iv_data and self.slider_vars:
                self._snap_sliders_to_telemetry(self.current_iv_data)

        self.follow_probe.set(True) # Auto-enable tracking for Live view
        self._update_live_status_ui()
        self.render_mode = 'scatter'
        self._update_render_toggle_visibility()
        if self.event_bus: self.event_bus.publish("live_sync_changed", {"active": True})
        if not self.live_tick_id: self._live_tick()

    def stop_live_mode(self):
        self.live_mode.set(False)
        self._update_live_status_ui()
        if self.live_tick_id: self.after_cancel(self.live_tick_id); self.live_tick_id = None
        self.render_mode = 'gradient'
        self._update_render_toggle_visibility()
        if self.event_bus: self.event_bus.publish("live_sync_changed", {"active": False})

    def _on_scan_finished(self, _):
        """Called when the serial engine signals scan complete. Stops live mode and
        upgrades all plots to gradient now that the full dataset is available."""
        self.stop_live_mode()
        if self.data:
            # Rebuild plot grid from the complete final dataset. During the scan,
            # the last _rebuild_controls() call resets slider vars to vals[0], and
            # any _sync_plots() race that sees empty dv_fields destroys plot_widgets.
            # Rebuilding here guarantees plot_widgets are present before redraw.
            self._update_ui_fields()
            # Restore slicers to the last scanned position so the final slice is shown.
            if self.current_iv_data and self.slider_vars:
                self._snap_sliders_to_telemetry(self.current_iv_data)
        self._redraw_all_plots()

    def _live_tick(self):
        """Throttled refresh loop (1Hz)."""
        if not self.live_mode.get() or self.current_filepath != "LIVE_SESSION": 
            self.live_tick_id = None; return
            
        # GATED RENDERING: Only draw if the panel is actually visible to the user
        if self.winfo_viewable() and self.needs_redraw:
            self._update_all_plots(); self.needs_redraw = False
            
        self.live_tick_id = self.after(1000, self._live_tick)

    def _load_data(self):
        path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv"), ("JSON Files", "*.json"), ("All Files", "*.*")])
        if path: self._process_file(path)

    def _refresh_current_file(self):
        if hasattr(self, 'current_filepath') and self.current_filepath:
            if self.current_filepath == "LIVE_SESSION": self.start_live_mode()
            else: self._process_file(self.current_filepath)
        else: messagebox.showinfo("Info", "No file loaded.")

    def _toggle_render_mode(self):
        self.render_mode = 'scatter' if self.render_mode == 'gradient' else 'gradient'
        self.last_plot_params.clear()
        self._update_render_toggle_visibility()
        self._redraw_all_plots()

    def _update_render_toggle_visibility(self):
        if not hasattr(self, 'render_toggle_btn'): return
        is_live = self.live_mode.get() and self.current_filepath == "LIVE_SESSION"
        in_gradient = (self.render_mode == 'gradient') and not is_live

        # Render mode toggle (always visible when not live)
        if is_live:
            self.render_toggle_btn.pack_forget()
        else:
            self.render_toggle_btn.pack(side=tk.RIGHT, padx=2)
            self.render_toggle_btn.config(text="≋ Gradient View" if self.render_mode == 'scatter' else "· Dot View")

        # Gradient-only controls: topo toggle + settings popup
        for widget in (self.topo_btn, self.plot_settings_btn):
            widget.pack_forget()
        if in_gradient:
            topo_text = "⛰ Topo: ON" if self.show_topo_overlay.get() else "⛰ Topo: OFF"
            self.topo_btn.config(text=topo_text)
            self.topo_btn.pack(side=tk.RIGHT, padx=2)
            self.plot_settings_btn.pack(side=tk.RIGHT, padx=2)

    def _cycle_render_style(self):
        order = ['contourf', 'gouraud', 'flat']
        self.render_style = order[(order.index(self.render_style) + 1) % len(order)]
        self._update_render_toggle_visibility()
        self._redraw_all_plots()

    def _toggle_topo_overlay(self):
        self.show_topo_overlay.set(not self.show_topo_overlay.get())
        self._update_render_toggle_visibility()
        self._redraw_all_plots()

    def _get_active_cmap(self):
        if self.cmap_name == 'custom_inferno':
            return self.custom_inferno
        return matplotlib.colormaps[self.cmap_name]

    def _open_plot_settings(self):
        popup = tk.Toplevel(self)
        popup.title("Plot Settings")
        popup.configure(bg=utils.COLOR_PANEL_BG)
        popup.resizable(False, False)
        popup.transient(self.winfo_toplevel())
        popup.grab_set()

        self.winfo_toplevel().update_idletasks()
        px = self.winfo_toplevel().winfo_x()
        py = self.winfo_toplevel().winfo_y()
        pw = self.winfo_toplevel().winfo_width()
        ph = self.winfo_toplevel().winfo_height()
        w, h = 320, 440
        popup.geometry(f"{w}x{h}+{max(0, px + (pw - w)//2)}+{max(0, py + (ph - h)//2)}")

        tk.Label(popup, text="PLOT SETTINGS", font=('Rajdhani', 13, 'bold'),
                 fg=utils.COLOR_ACCENT_CYAN, bg=utils.COLOR_PANEL_BG).pack(pady=(14, 6))

        # --- Colormap section ---
        cmap_frame = tk.LabelFrame(popup, text="  COLORMAP  ",
                                   font=('Rajdhani', 10, 'bold'),
                                   fg=utils.COLOR_ACCENT_CYAN, bg=utils.COLOR_PANEL_BG,
                                   padx=10, pady=6)
        cmap_frame.pack(fill=tk.X, padx=16, pady=(0, 8))

        CMAPS = [
            ('custom_inferno', 'Custom Inferno (default)'),
            ('inferno',        'Inferno'),
            ('viridis',        'Viridis'),
            ('plasma',         'Plasma'),
            ('magma',          'Magma'),
            ('coolwarm',       'Coolwarm'),
            ('jet',            'Jet'),
            ('RdBu_r',         'Red–Blue (diverging)'),
        ]
        cmap_var = tk.StringVar(value=self.cmap_name)
        for val, label in CMAPS:
            tk.Radiobutton(cmap_frame, text=label, variable=cmap_var, value=val,
                           font=('Inter', 9), fg=utils.COLOR_TEXT_PRIMARY,
                           bg=utils.COLOR_PANEL_BG, selectcolor=utils.COLOR_BG,
                           activebackground=utils.COLOR_PANEL_BG).pack(anchor='w')

        # --- Contour line style section ---
        topo_frame = tk.LabelFrame(popup, text="  CONTOUR LINE STYLE  ",
                                   font=('Rajdhani', 10, 'bold'),
                                   fg=utils.COLOR_ACCENT_CYAN, bg=utils.COLOR_PANEL_BG,
                                   padx=10, pady=8)
        topo_frame.pack(fill=tk.X, padx=16, pady=(0, 8))

        color_row = tk.Frame(topo_frame, bg=utils.COLOR_PANEL_BG)
        color_row.pack(fill=tk.X, pady=(0, 4))
        tk.Label(color_row, text="Color:", font=('Inter', 9),
                 fg=utils.COLOR_TEXT_PRIMARY, bg=utils.COLOR_PANEL_BG).pack(side=tk.LEFT)
        TOPO_COLORS = [
            ('#ffffff', 'White'),
            ('#000000', 'Black'),
            ('#ffff00', 'Yellow'),
            ('#00ffff', 'Cyan'),
            ('#ff6600', 'Orange'),
        ]
        topo_color_var = tk.StringVar(value=self.topo_color)
        color_combo = ttk.Combobox(color_row, textvariable=topo_color_var, state='readonly',
                                   values=[lbl for _, lbl in TOPO_COLORS], width=10)
        color_lbl_to_hex = {lbl: hx for hx, lbl in TOPO_COLORS}
        hex_to_lbl = {hx: lbl for hx, lbl in TOPO_COLORS}
        color_combo.set(hex_to_lbl.get(self.topo_color, 'White'))
        color_combo.pack(side=tk.LEFT, padx=6)

        thick_row = tk.Frame(topo_frame, bg=utils.COLOR_PANEL_BG)
        thick_row.pack(fill=tk.X, pady=(0, 2))
        tk.Label(thick_row, text="Thickness:", font=('Inter', 9),
                 fg=utils.COLOR_TEXT_PRIMARY, bg=utils.COLOR_PANEL_BG).pack(side=tk.LEFT)
        topo_lw_var = tk.DoubleVar(value=self.topo_linewidth)
        ttk.Spinbox(thick_row, from_=0.2, to=4.0, increment=0.2, width=5,
                    textvariable=topo_lw_var, format="%.1f").pack(side=tk.LEFT, padx=6)

        levels_row = tk.Frame(topo_frame, bg=utils.COLOR_PANEL_BG)
        levels_row.pack(fill=tk.X, pady=(0, 2))
        tk.Label(levels_row, text="Levels:", font=('Inter', 9),
                 fg=utils.COLOR_TEXT_PRIMARY, bg=utils.COLOR_PANEL_BG).pack(side=tk.LEFT)
        topo_lvl_var = tk.IntVar(value=self.topo_levels_var.get())
        ttk.Spinbox(levels_row, from_=2, to=30, increment=1, width=5,
                    textvariable=topo_lvl_var).pack(side=tk.LEFT, padx=6)

        # --- Buttons ---
        btn_row = tk.Frame(popup, bg=utils.COLOR_PANEL_BG)
        btn_row.pack(pady=10)

        def _apply():
            self.cmap_name = cmap_var.get()
            selected_lbl = topo_color_var.get()
            self.topo_color = color_lbl_to_hex.get(selected_lbl, selected_lbl)
            try:
                self.topo_linewidth = float(topo_lw_var.get())
            except (tk.TclError, ValueError):
                pass
            try:
                self.topo_levels_var.set(int(topo_lvl_var.get()))
            except (tk.TclError, ValueError):
                pass
            popup.destroy()
            self._redraw_all_plots()

        def _cancel():
            popup.destroy()

        ttk.Button(btn_row, text="Apply", command=_apply, width=12).pack(side=tk.LEFT, padx=8)
        ttk.Button(btn_row, text="Cancel", command=_cancel, width=12).pack(side=tk.LEFT, padx=8)
        popup.bind("<Return>", lambda e: _apply())
        popup.bind("<Escape>", lambda e: _cancel())

    def _redraw_all_plots(self):
        """Force-redraws every active plot by clearing the param cache and refreshing."""
        self.last_plot_params.clear()
        self.needs_redraw = True
        filtered, self.current_slice_str = self._get_filtered_data()
        x_col = self.x_axis_var.get() if hasattr(self, 'x_axis_var') else 'x'
        y_col = self.y_axis_var.get() if hasattr(self, 'y_axis_var') else 'y'
        for field, (container, canvas, fig, ax) in self.plot_widgets.items():
            self._draw_heatmap(ax, canvas, fig, filtered, x_col, y_col, field)

    def _toggle_live_sync(self):
        """Consolidated toggle for Sync Live / Stop Sync."""
        if self.live_mode.get() and self.current_filepath == "LIVE_SESSION":
            self.stop_live_mode()
        else:
            self.start_live_mode()

    def _process_file(self, path):
        try:
            ivs, dvs = set(), set()
            # Scan for metadata comments
            with open(path, 'r') as f:
                header_json = None
                for _ in range(20):
                    line = f.readline()
                    if not line: break
                    if line.startswith("# SOURCE_IVS:"):
                        ivs = set(item.strip() for item in line.split(',')[1:] if item.strip())
                    if line.startswith("# SOURCE_DVS:"):
                        dvs = set(item.strip() for item in line.split(',')[1:] if item.strip())
                    if line.startswith("; PARAMS:"):
                        header_json = line.replace("; PARAMS:", "").strip()
                
                if header_json:
                    try: self.update_fixed_bounds(json.loads(header_json))
                    except: pass
            
            # Use comment='#' to automatically skip all metadata rows, making it robust to header changes
            df = pd.read_csv(path, comment='#')
            self.data = df.to_dict(orient='records')
            if not self.data: messagebox.showwarning("Warning", "The file is empty."); return
            
            self.fields = set(self.data[0].keys())
            if not ivs:
                known = ['x', 'y', 'z', 'rot', 'height', 'elevation', 'rotation', 'r', 'e']
                ivs = {f for f in self.fields if f.lower() in known}
                dvs = {f for f in self.fields if f.lower() not in known and f != 'Timestamp'}
            
            self.iv_fields, self.dv_fields = ivs, dvs
            self.dv_vars.clear()
            self.last_plot_params.clear()
            self.current_filepath = path; self.stop_live_mode(); self._update_ui_fields()
        except Exception as e: messagebox.showerror("Error", f"Failed to process file: {e}")

    def _on_axis_change(self):
        if self.data: self._rebuild_controls()

    def _select_all_dvs(self):
        for v in self.dv_vars.values(): v.set(True)
        self._sync_plots()

    def _select_none_dvs(self):
        for v in self.dv_vars.values(): v.set(False)
        self._sync_plots()

    def _sync_plots(self):
        # Only process fields that are actually in the current dataset's DVs
        selected_dvs = [f for f in self.dv_fields if f in self.dv_vars and self.dv_vars[f].get()]
        
        # 1. Remove widgets for DVs that are no longer selected or present
        to_remove = [f for f in self.plot_widgets.keys() if f not in selected_dvs]
        for field in to_remove:
            container, canvas, fig, ax = self.plot_widgets.pop(field)
            container.destroy()

        if not selected_dvs:
            # Clear everything and show placeholder
            for widget in self.grid_container.winfo_children(): widget.destroy()
            ttk.Label(self.grid_container, text="Select fields in sidebar to visualize", 
                      font=utils.FONT_BODY, foreground=utils.COLOR_TEXT_SECONDARY).pack(pady=100)
            return

        # Remove placeholder if it exists
        for widget in self.grid_container.winfo_children():
            if isinstance(widget, ttk.Label) and widget.cget("text") == "Select fields in sidebar to visualize":
                widget.destroy()

        # 2. Create widgets for newly selected DVs
        cols = 2
        for field in selected_dvs:
            if field not in self.plot_widgets:
                container = ttk.Frame(self.grid_container, style='Card.TLabelframe')
                title_bar = ttk.Frame(container, style='Panel.TFrame'); title_bar.pack(fill=tk.X)
                ttk.Label(title_bar, text=field.upper(), font=('Inter', 10, 'bold'), 
                          foreground=utils.COLOR_ACCENT_CYAN).pack(side=tk.LEFT, padx=5)
                ttk.Button(title_bar, text="💾", width=3, style='Small.TButton', 
                           command=lambda f=field: self._export_single_plot(f)).pack(side=tk.RIGHT)
                
                fig = Figure(figsize=(5, 4), dpi=100, facecolor=utils.COLOR_PANEL_BG)
                ax = fig.add_subplot(111)
                canvas = FigureCanvasTkAgg(fig, master=container)
                canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
                
                self.plot_widgets[field] = (container, canvas, fig, ax)

        # 3. Re-grid existing and new widgets to ensure correct order/layout
        for i, field in enumerate(selected_dvs):
            container, canvas, fig, ax = self.plot_widgets[field]
            r, c = divmod(i, cols)
            container.grid(row=r, column=c, padx=10, pady=10, sticky='nsew')
            self.grid_container.rowconfigure(r, weight=0, minsize=380)
        
        for i in range(cols):
            self.grid_container.columnconfigure(i, weight=1, minsize=450)

        self._update_all_plots()

    def _update_all_plots(self):
        if not self.data: return
        x_col, y_col = self.x_axis_var.get(), self.y_axis_var.get()
        if not (x_col and y_col): return
        
        filtered, self.current_slice_str = self._get_filtered_data()
        for field, (container, canvas, fig, ax) in self.plot_widgets.items():
            self._draw_heatmap(ax, canvas, fig, filtered, x_col, y_col, field)

    def _export_single_plot(self, field):
        """Exports a single plot as PNG or PDF with white background."""
        if field not in self.plot_widgets: return
        _, _, old_fig, _ = self.plot_widgets[field]
        
        path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG Image", "*.png"), ("PDF Document", "*.pdf")],
            initialfile=f"POP_{field}_{datetime.datetime.now().strftime('%y%m%d_%H%M')}"
        )
        if not path: return

        # Ensure slice info is current
        self._update_all_plots()

        # Re-render on a separate white-background figure for export
        x_col, y_col = self.x_axis_var.get(), self.y_axis_var.get()
        filtered, _ = self._get_filtered_data()

        export_fig = Figure(figsize=(8, 6), dpi=150, facecolor='white')
        export_ax = export_fig.add_subplot(111)
        self._draw_export_heatmap(export_ax, export_fig, filtered, x_col, y_col, field)
        
        # Add slice metadata to individual export - positioned below X-axis, left aligned
        export_fig.text(0.1, 0.02, self.current_slice_str, 
                        ha='left', fontsize=9, fontstyle='italic', color='black')
        
        try:
            export_fig.savefig(path, bbox_inches='tight')
            messagebox.showinfo("Success", f"Exported {os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export: {e}")

    def _get_render_strategy(self):
        if self.live_mode.get() and self.current_filepath == "LIVE_SESSION":
            return "scatter"
        return self.render_mode

    def _draw_heatmap(self, ax, canvas, fig, filtered_data, x_col, y_col, z_col):
        current_params = (len(filtered_data), x_col, y_col, z_col, self.current_slice_str,
                          self.render_style, self.show_topo_overlay.get(), self.topo_levels_var.get(),
                          self.cmap_name, self.topo_color, self.topo_linewidth)
        if self.last_plot_params.get(z_col) == current_params: return
        self.last_plot_params[z_col] = current_params

        # Remove any previously registered draw callbacks for this axes
        cb_attr = f"_draw_cb_{z_col}"
        old_cb = getattr(fig, cb_attr, None)
        if old_cb is not None:
            try: fig.canvas.mpl_disconnect(old_cb)
            except Exception: pass
        setattr(fig, cb_attr, None)

        ax.clear()
        ax.set_facecolor(utils.COLOR_BLACK)
        ax.tick_params(colors=utils.COLOR_TEXT_SECONDARY, labelsize=7)

        xl, yl = x_col.lower(), y_col.lower()
        use_equal_aspect = (xl in self.distance_vars and yl in self.distance_vars)

        # Capture the bounds we want to enforce into local variables so the
        # callback closure holds them by value, not by reference to self.fixed_bounds
        # (which could change between draws).
        def _get_axis_lim(key):
            if self.fixed_bounds and key in self.fixed_bounds:
                return self.fixed_bounds[key]
            if self._live_axis_extents and key in self._live_axis_extents:
                lo, hi = self._live_axis_extents[key]
                v_range = hi - lo
                margin = v_range * 0.05 if v_range > 0 else 5.0
                return (lo - margin, hi + margin)
            return None

        x_lim = _get_axis_lim(xl)
        y_lim = _get_axis_lim(yl)

        # Set aspect ratio once as a persistent axes property.
        # Keeping this outside _lock_viewport avoids triggering a redraw
        # loop when the callback fires (set_aspect invalidates the canvas,
        # which would re-fire the draw_event callback indefinitely).
        ax.set_aspect('equal' if use_equal_aspect else 'auto', adjustable='box')

        def _lock_viewport():
            """Re-enforce axis limits — called as a post-draw callback to
            override the autoscale reset that contourf/tricontourf applies."""
            if x_lim: ax.set_xlim(x_lim)
            if y_lim: ax.set_ylim(y_lim)
            if x_lim or y_lim:
                ax.set_autoscale_on(False)

        _lock_viewport()

        if not filtered_data:
            ax.text(0.5, 0.5, "No Data for Slice", ha='center', va='center',
                    color='white', transform=ax.transAxes)
            if canvas: canvas.draw_idle()
            return

        valid_points = [d for d in filtered_data if all(k in d for k in [x_col, y_col, z_col])]
        if not valid_points:
            ax.text(0.5, 0.5, f"No Data for {z_col}\n(Check IV/DV Mapping)",
                    ha='center', va='center', color='white', transform=ax.transAxes)
            if canvas: canvas.draw_idle()
            return

        xs = np.array([d[x_col] for d in valid_points])
        ys = np.array([d[y_col] for d in valid_points])
        zs = np.array([d[z_col] for d in valid_points])

        ax.set_title(f"{z_col.upper()}", color=utils.COLOR_ACCENT_CYAN,
                    fontdict={'weight': 'bold', 'size': 9}, pad=10)
        ax.set_xlabel(x_col, color=utils.COLOR_TEXT_SECONDARY, size=7)
        ax.set_ylabel(y_col, color=utils.COLOR_TEXT_SECONDARY, size=7)

        cp = None
        try:
            strategy = self._get_render_strategy()
            ux, uy = sorted(set(xs)), sorted(set(ys))

            z_min, z_max = float(np.min(zs)), float(np.max(zs))
            if z_min == z_max: z_max += 1.0

            scatter_kw = dict(cmap=self._get_active_cmap(), s=25, alpha=0.95,
                            edgecolors='white', linewidths=0.1,
                            vmin=z_min, vmax=z_max)

            if strategy == "scatter" or len(valid_points) < 10:
                cp = ax.scatter(xs, ys, c=zs, **scatter_kw)
            elif len(ux) > 1 and len(uy) > 1:
                if len(ux) * len(uy) == len(valid_points):
                    X, Y = np.meshgrid(ux, uy)
                    Z = np.zeros((len(uy), len(ux)))
                    for d in valid_points:
                        Z[uy.index(d[y_col]), ux.index(d[x_col])] = d[z_col]
                    if self.render_style == 'gouraud':
                        cp = ax.pcolormesh(X, Y, Z, shading='gouraud',
                                           cmap=self._get_active_cmap(), vmin=z_min, vmax=z_max)
                    elif self.render_style == 'flat':
                        cp = ax.pcolormesh(X, Y, Z, shading='flat',
                                           cmap=self._get_active_cmap(), vmin=z_min, vmax=z_max)
                    else:  # contourf
                        cp = ax.contourf(X, Y, Z, levels=20, cmap=self._get_active_cmap(),
                                         vmin=z_min, vmax=z_max)
                    if self.show_topo_overlay.get():
                        n_lvls = self.topo_levels_var.get()
                        if n_lvls > 0:
                            try:
                                ax.contour(X, Y, Z, levels=n_lvls,
                                           colors=self.topo_color, alpha=0.35,
                                           linewidths=self.topo_linewidth)
                            except Exception: pass
                else:
                    try:
                        cp = ax.tricontourf(xs, ys, zs, levels=20, cmap=self._get_active_cmap(),
                                            vmin=z_min, vmax=z_max)
                        if self.show_topo_overlay.get():
                            n_lvls = self.topo_levels_var.get()
                            if n_lvls > 0:
                                try:
                                    ax.tricontour(xs, ys, zs, levels=n_lvls,
                                                  colors=self.topo_color, alpha=0.35,
                                                  linewidths=self.topo_linewidth)
                                except Exception: pass
                    except Exception:
                        cp = ax.scatter(xs, ys, c=zs, **scatter_kw)
            else:
                cp = ax.scatter(xs, ys, c=zs, **scatter_kw)

            if cp:
                cax_tag = f"_cax_{z_col}"
                cax = getattr(fig, cax_tag, None)
                if cax is not None and cax not in fig.axes:
                    cax = None
                if cax is None:
                    pos = ax.get_position()
                    cax = fig.add_axes([
                        pos.x1 + 0.01,
                        pos.y0,
                        pos.width * 0.04,
                        pos.height
                    ])
                    setattr(fig, cax_tag, cax)
                else:
                    cax.clear()
                fig.colorbar(cp, cax=cax).ax.tick_params(
                    labelsize=6, colors=utils.COLOR_TEXT_SECONDARY)

        except Exception as e:
            ax.text(0.5, 0.5, f"Plot Error: {e}", ha='center', va='center',
                    color=utils.COLOR_ACCENT_RED, transform=ax.transAxes)

        # Re-enforce immediately (catches scatter/contourf autoscale reset)
        _lock_viewport()

        # Register as a draw callback — this fires AFTER Matplotlib's full render
        # pipeline, so it wins against anything the layout engine does internally.
        # This is the only reliable way to freeze limits against contourf/tricontourf.
        if canvas:
            def _on_draw(event, _ax=ax, _fn=_lock_viewport):
                _fn()
            cid = canvas.mpl_connect('draw_event', _on_draw)
            setattr(fig, cb_attr, cid)

        if canvas: canvas.draw_idle()

    def _get_report_title(self):
        """Custom centered dialog for entering the report title - Oversized for visibility."""
        dialog = tk.Toplevel(self)
        dialog.title("Report Configuration")
        dialog.geometry("600x250")
        dialog.resizable(False, False)
        dialog.configure(bg=utils.COLOR_PANEL_BG)
        dialog.transient(self.winfo_toplevel())
        dialog.grab_set()
        
        # Center on parent window
        parent = self.winfo_toplevel()
        parent.update_idletasks()
        px, py = parent.winfo_x(), parent.winfo_y()
        pw, ph = parent.winfo_width(), parent.winfo_height()
        
        x = px + (pw // 2) - 300
        y = py + (ph // 2) - 125
        dialog.geometry(f"+{max(0, x)}+{max(0, y)}")

        ttk.Label(dialog, text="ENTER REPORT TITLE", 
                  font=('Rajdhani', 16, 'bold'), foreground=utils.COLOR_ACCENT_CYAN).pack(pady=(30, 10))
        
        title_var = tk.StringVar(value="SEED POP Analysis")
        # Oversized entry block
        entry = ttk.Entry(dialog, textvariable=title_var, font=('Inter', 14), width=45)
        entry.pack(pady=15, padx=30)
        entry.select_range(0, tk.END)
        entry.focus_set()

        result = {"title": None}

        def on_ok(event=None):
            val = title_var.get().strip()
            if val:
                result["title"] = val
                dialog.destroy()
            else:
                messagebox.showwarning("Warning", "Title cannot be empty.")

        def on_cancel(event=None):
            dialog.destroy()

        btn_frame = ttk.Frame(dialog, style='Panel.TFrame')
        btn_frame.pack(pady=20)
        
        ttk.Button(btn_frame, text="Generate", command=on_ok, width=15).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="Cancel", command=on_cancel, width=15).pack(side=tk.LEFT, padx=10)

        dialog.bind("<Return>", on_ok)
        dialog.bind("<Escape>", on_cancel)
        
        self.wait_window(dialog)
        return result["title"]

    def _save_snapshot(self):
        """Generates a professional multi-page PDF report in a 2x3 grid with branding and custom title."""
        if not self.data or not self.plot_widgets:
            messagebox.showwarning("Warning", "No active plots to export.")
            return
            
        # 1. Prompt for Report Name (Centered Custom Dialog)
        report_title = self._get_report_title()
        if not report_title: return # User cancelled

        # 2. Slice selection dialog (skipped when no slicer variables exist)
        if self.slider_vars:
            selected = SliceSelectionDialog(self, self.slider_vars).show()
            if selected is None: return  # user cancelled
            fields = list(selected.keys())
            slice_combos = [dict(zip(fields, combo))
                            for combo in itertools.product(*[selected[f] for f in fields])]
        else:
            slice_combos = [{}]  # no slicers — single pass, backwards compatible

        # Sanitize for filename - Ensure automated naming works
        safe_title = "".join([c if c.isalnum() or c in "._-" else "_" for c in report_title])
        default_name = f"{safe_title}_{datetime.datetime.now().strftime('%y%m%d_%H%M')}.pdf"

        path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF Report", "*.pdf")],
            initialfile=default_name,
            title="Save POP Report"
        )
        if not path: return

        # Ensure slice info is current
        self._update_all_plots()

        try:
            active_fields = [f for f, var in self.dv_vars.items() if var.get()]
            plots_per_page = 6
            cols_per_page = 2
            pages_per_combo = max(1, (len(active_fields) + plots_per_page - 1) // plots_per_page)
            total_pages = pages_per_combo * len(slice_combos)

            x_col, y_col = self.x_axis_var.get(), self.y_axis_var.get()

            # Logo Handling - Optimizing for clarity
            if getattr(sys, "frozen", False):
                project_root = sys._MEIPASS
            else:
                script_dir = os.path.dirname(os.path.abspath(__file__))
                project_root = os.path.dirname(script_dir)
            logo_path = os.path.join(project_root, "assets", "RLM_logo.png")
            logo_img = None
            if os.path.exists(logo_path):
                from matplotlib.image import imread
                logo_img = imread(logo_path)

            _unit = {'x': 'mm', 'y': 'mm', 'z': 'mm', 'rot': '°', 'e': '°', 'rotation': '°'}

            with PdfPages(path) as pdf:
                global_page = 0
                for combo in slice_combos:
                    # Filter data for this slice combination
                    filtered = self.data
                    if combo:
                        for f, v in combo.items():
                            filtered = [d for d in filtered if d.get(f) is not None and abs(d[f] - v) < 1e-4]
                    else:
                        for f, (var, *_) in self.slider_vars.items():
                            filtered = [d for d in filtered if d.get(f) is not None and abs(d[f] - var.get()) < 1e-4]

                    # Build slice label for page sub-header
                    if combo:
                        parts = [f"{f.upper()} = {v:.4g}{_unit.get(f.lower(), '')}"
                                 for f, v in combo.items()]
                        slice_label = "  |  ".join(parts)
                    else:
                        slice_label = self.current_slice_str

                    for p_idx in range(pages_per_combo):
                        global_page += 1
                        # Ultra-high DPI (300) for professional print quality and logo clarity
                        fig = Figure(figsize=(8.27, 11.69), dpi=300, facecolor='white')

                        # 1. Header (Updated Structure)
                        if logo_img is not None:
                            # Spanned Logo on Right
                            ax_logo = fig.add_axes([0.7, 0.90, 0.22, 0.08], anchor='NE', zorder=1)
                            ax_logo.imshow(logo_img, interpolation='bicubic')
                            ax_logo.axis('off')

                        # Left Side Header Text
                        fig.text(0.1, 0.96, report_title.upper(), fontsize=13, fontweight='bold', color='black')
                        fig.text(0.1, 0.942, "Proof of Power Report", fontsize=10, color='black')
                        fig.text(0.1, 0.928, datetime.date.today().strftime('%Y-%m-%d'), fontsize=9, color='black')

                        # Header Divider
                        fig.add_artist(matplotlib.lines.Line2D((0.1, 0.9), (0.91, 0.91), color='black', linewidth=1.2, transform=fig.transFigure))

                        # Meta Info (Below Divider) — slice label centered on its own line
                        fig.text(0.1, 0.892, f"Source: {os.path.basename(self.current_filepath or 'Live Session')}", fontsize=9, color='black')
                        fig.text(0.5, 0.878, slice_label, ha='center', fontsize=9, fontstyle='italic', color='black')

                        # 3. Plot Grid (2x3)
                        for i in range(plots_per_page):
                            f_idx = p_idx * plots_per_page + i
                            if f_idx >= len(active_fields): break

                            field = active_fields[f_idx]
                            row, col = divmod(i, cols_per_page)

                            left = 0.1 + col * 0.42
                            bottom = 0.61 - row * 0.25
                            ax = fig.add_axes([left, bottom, 0.38, 0.18])
                            self._draw_export_heatmap(ax, fig, filtered, x_col, y_col, field, is_grid=True)

                        # 4. Footer
                        fig.add_artist(matplotlib.lines.Line2D((0.1, 0.9), (0.08, 0.08), color='black', linewidth=0.8, transform=fig.transFigure))
                        fig.text(0.9, 0.05, f"Page {global_page} of {total_pages}", ha='right', fontsize=9, color='black')

                        pdf.savefig(fig)
                
            messagebox.showinfo("Success", f"Report generated: {os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate report: {e}")

    def _draw_export_heatmap(self, ax, fig, filtered_data, x_col, y_col, z_col, is_grid=False):
        """Helper to draw heatmap with printer-friendly colors for export."""
        ax.set_facecolor('white')
        font_scale = 0.7 if is_grid else 1.0
        ax.tick_params(colors='black', labelsize=8 * font_scale)
        
        # Physical 1:1 scaling
        if x_col.lower() in self.distance_vars and y_col.lower() in self.distance_vars:
            ax.set_aspect('equal', adjustable='datalim')
        
        # Increase title padding to avoid scientific notation overlap
        title_size = 11 if is_grid else 14
        ax.set_title(f"{z_col.upper()}", color='black', fontdict={'weight': 'bold', 'size': title_size}, pad=20*font_scale)
        
        # Adjust offset text position if present
        ax.yaxis.get_offset_text().set_fontsize(7 * font_scale)
        ax.xaxis.get_offset_text().set_fontsize(7 * font_scale)

        ax.set_xlabel(f"{x_col} (mm)", color='black', size=9*font_scale)
        ax.set_ylabel(f"{y_col} (mm)", color='black', size=9*font_scale)
        
        xs = np.array([d[x_col] for d in filtered_data])
        ys = np.array([d[y_col] for d in filtered_data])
        zs = np.array([d[z_col] for d in filtered_data])
        
        try:
            ux, uy = sorted(list(set(xs))), sorted(list(set(ys)))
            if len(ux) < 2: ax.set_xlim(xs[0]-5, xs[0]+5)
            if len(uy) < 2: ax.set_ylim(ys[0]-5, ys[0]+5)

            if len(ux) > 1 and len(uy) > 1:
                if len(ux) * len(uy) == len(filtered_data):
                    X, Y = np.meshgrid(ux, uy)
                    Z = np.zeros((len(uy), len(ux)))
                    lookup = {(d[x_col], d[y_col]): d[z_col] for d in filtered_data}
                    for i, y_v in enumerate(uy):
                        for j, x_v in enumerate(ux):
                            Z[i, j] = lookup.get((x_v, y_v), np.nan)
                    if self.render_style == 'gouraud':
                        cp = ax.pcolormesh(X, Y, Z, shading='gouraud', cmap=self._get_active_cmap())
                    elif self.render_style == 'flat':
                        cp = ax.pcolormesh(X, Y, Z, shading='flat', cmap=self._get_active_cmap())
                    else:
                        cp = ax.contourf(X, Y, Z, cmap=self._get_active_cmap(), levels=20)
                    if self.show_topo_overlay.get():
                        n_lvls = self.topo_levels_var.get()
                        if n_lvls > 0:
                            try:
                                ax.contour(X, Y, Z, levels=n_lvls,
                                           colors='black', alpha=0.35,
                                           linewidths=self.topo_linewidth)
                            except Exception: pass
                else:
                    cp = ax.tricontourf(xs, ys, zs, cmap=self._get_active_cmap(), levels=20)
                    if self.show_topo_overlay.get():
                        n_lvls = self.topo_levels_var.get()
                        if n_lvls > 0:
                            try:
                                ax.tricontour(xs, ys, zs, levels=n_lvls,
                                              colors='black', alpha=0.35,
                                              linewidths=self.topo_linewidth)
                            except Exception: pass
            else:
                cp = ax.scatter(xs, ys, c=zs, cmap=self._get_active_cmap(), edgecolors='black', s=40*font_scale)
            
            if cp:
                divider = make_axes_locatable(ax)
                cax = divider.append_axes("right", size="5%", pad=0.05)
                cbar = fig.colorbar(cp, cax=cax)
                cax.tick_params(colors='black', labelsize=7 * font_scale)
                cax.yaxis.get_offset_text().set_fontsize(6 * font_scale)
        except: pass


