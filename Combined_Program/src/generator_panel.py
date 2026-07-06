import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import math
import os
import re
import webbrowser
from datetime import datetime
import json
from typing import Any, Dict, List, Optional
import utils
from utils import PRINTER_LIMITS, PRINTER_BOUNDS

# Component Imports
from generator_components.pattern_input import PatternInput
from generator_components.pattern_preview import PatternPreview
from generator_components.command_injection import CommandInjection

# Optional Numpy
try:
    import numpy as np
    HAS_NUMPY = True
except ImportError:
    HAS_NUMPY = False

class PatternGeneratorGUI:
    """
    Main Container for the Pattern Generator tool.
    Orchestrates the design, visualization, and Hub command injection tabs.
    """
    def __init__(self, parent_frame):
        self.parent = parent_frame
        self.root = parent_frame.winfo_toplevel()
        
        # --- UI CONSTANTS & STYLES ---
        self._setup_appearance()
        self._setup_ttk_styles()

        # --- SHARED STATE ---
        self._initialize_state()

        # --- UI CONSTRUCTION ---
        self.notebook = ttk.Notebook(self.parent, style='TNotebook')
        self.notebook.pack(fill=tk.BOTH, expand=True)
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

        # Tab 1: Design
        self.design_tab = ttk.Frame(self.notebook, style='Dark.TFrame')
        self.notebook.add(self.design_tab, text=" 1. DESIGN PATTERN ")
        
        design_container = ttk.Frame(self.design_tab, style='Dark.TFrame')
        design_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.input_panel = PatternInput(design_container, self)
        self.input_panel.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        
        self.preview_panel = PatternPreview(design_container, self)
        self.preview_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # Tab 2: Injection
        self.injection_tab = ttk.Frame(self.notebook, style='Dark.TFrame')
        self.notebook.add(self.injection_tab, text=" 2. COMMAND INJECTION ")
        
        self.injection_panel = CommandInjection(self.injection_tab, self)
        self.injection_panel.pack(fill=tk.BOTH, expand=True)

        self._load_last_parameters()

        # Sync UI states and trigger initial preview
        self._on_x_symmetric_toggle(derive_values=False)
        self._on_y_symmetric_toggle(derive_values=False)
        self._on_z_symmetric_toggle(derive_values=False)
        self._on_rot_symmetric_toggle(derive_values=False)
        self._auto_update_preview()

    def _setup_appearance(self):
        self.COLOR_BG = "#0a0e14"
        self.COLOR_PANEL_BG = "#161b22"
        self.COLOR_BORDER = "#30363d"
        self.COLOR_TEXT_PRIMARY = "#e6edf3"
        self.COLOR_TEXT_SECONDARY = "#7d8590"
        self.COLOR_ACCENT_CYAN = "#00d4ff"
        self.COLOR_ACCENT_PURPLE = "#a371f7"
        self.COLOR_ACCENT_GREEN = "#3fb950"
        self.COLOR_ACCENT_AMBER = "#ffa657"
        self.COLOR_ACCENT_RED = "#ff4444"
        self.COLOR_BLACK = "#000000"
        self.FONT_HEADER = ("Orbitron", 13)
        self.FONT_BODY = ("Inter", 10)
        self.FONT_BODY_SMALL = ("Inter", 9)
        self.FONT_BODY_BOLD = ("Inter", 10, "bold")
        self.FONT_MONO = ("JetBrains Mono", 9)
        self.FONT_MONO_LARGE = ('JetBrains Mono', 11, 'bold')

    def _setup_ttk_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('.', background=self.COLOR_PANEL_BG, foreground=self.COLOR_TEXT_PRIMARY, 
                        fieldbackground=self.COLOR_BLACK, bordercolor=self.COLOR_BORDER, font=self.FONT_BODY)
        style.configure('Dark.TFrame', background=self.COLOR_BG)
        style.configure('TNotebook', background=self.COLOR_BG, borderwidth=0)
        style.configure('TNotebook.Tab', background=self.COLOR_PANEL_BG, foreground=self.COLOR_TEXT_SECONDARY, padding=(15, 5), font=self.FONT_BODY_BOLD)
        style.map('TNotebook.Tab', background=[('selected', self.COLOR_ACCENT_CYAN), ('active', '#2c333e')],
                  foreground=[('selected', self.COLOR_BLACK), ('active', self.COLOR_ACCENT_CYAN)])
        style.configure('Card.TLabelframe', background=self.COLOR_PANEL_BG, bordercolor=self.COLOR_BORDER, borderwidth=1, relief=tk.SOLID, padding=12)
        style.configure('Card.TLabelframe.Label', background=self.COLOR_PANEL_BG, foreground=self.COLOR_ACCENT_CYAN, font=('Inter', 10, 'bold'))
        style.configure('Primary.TButton', background=self.COLOR_ACCENT_CYAN, foreground=self.COLOR_BLACK, padding=(12, 10), font=self.FONT_BODY_BOLD)
        style.map('Primary.TButton', background=[('active', '#00eaff')])
        # Rotation Button Style
        style.configure('ViewCube.TButton', background=self.COLOR_PANEL_BG, foreground=self.COLOR_ACCENT_CYAN, font=("Inter", 8))

    def _initialize_state(self):
        self.profile_name_var = tk.StringVar(value="GCODE")
        self.x_symmetric = tk.BooleanVar(value=True); self.y_symmetric = tk.BooleanVar(value=True)
        self.z_symmetric = tk.BooleanVar(value=False); self.rot_symmetric = tk.BooleanVar(value=True)
        self.export_format = tk.StringVar(value="gcode"); self.include_timestamp = tk.BooleanVar(value=True)
        self.hub_actions = []; self.hub_action_types = ["WPT Start", "WPT Stop", "DCDC Enable", "DCDC Disable", "WAIT"]
        self.repeated_hub_actions = []
        self.matplotlib_imported = False; self.is_3d_plot_enabled = tk.BooleanVar(value=True); self.toolpath_3d_opacity_var = tk.DoubleVar(value=0.8)
        self.on_send_to_sender = None
        
        # 3D Plot references
        self.ax_3d = None
        self.canvas_3d = None
        self.fig_3d = None

        # UI Component tracking
        self.filename_previews = []
        self._s_popup = None

        # Color Palette for Hub Actions
        self.ACTION_COLORS = [
            "#3fb950", # Green
            "#ffa657", # Amber
            "#a371f7", # Purple
            "#f287aa", # Pink
            "#ffeb3b", # Yellow
            "#00d4ff", # Cyan
            "#ffffff", # White
            "#7d8590"  # Gray
        ]

        # Autocomplete Caching
        self._axis_cache = {} 
        self._cache_params_key = None

    def _create_naming_ui(self, parent):
        """Creates a synchronized naming UI block - Explicitly top of footer."""
        profile_frame = ttk.LabelFrame(parent, text="Test Profile", style='Card.TLabelframe', padding=12)
        profile_frame.pack(side=tk.TOP, fill=tk.X, pady=(0, 0), padx=2)
        profile_frame.columnconfigure(1, weight=1)

        ttk.Label(profile_frame, text="Profile Name:", foreground=self.COLOR_TEXT_SECONDARY).grid(row=0, column=0, sticky=tk.W, pady=2)
        ent = ttk.Entry(profile_frame, textvariable=self.profile_name_var)
        ent.grid(row=0, column=1, pady=2, padx=5, sticky=tk.EW)
        ent.bind('<KeyRelease>', self.update_filename_preview)

        ttk.Checkbutton(profile_frame, text="Include timestamp", variable=self.include_timestamp, command=self.update_filename_preview).grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=2)

        ttk.Label(profile_frame, text="Filename:", foreground=self.COLOR_TEXT_SECONDARY).grid(row=2, column=0, sticky=tk.W, pady=2)
        preview = ttk.Label(profile_frame, text="", foreground=self.COLOR_ACCENT_CYAN, wraplength=300)
        preview.grid(row=2, column=1, sticky=tk.W, pady=2, padx=5)
        self.filename_previews.append(preview)
        return profile_frame

    def _create_action_buttons(self, parent):
        """Creates action buttons - Explicitly bottom of footer."""
        btn_frame = ttk.Frame(parent, style='Dark.TFrame')
        btn_frame.pack(side=tk.TOP, fill=tk.X, pady=(10, 0), padx=2)

        load_btn = ttk.Button(btn_frame, text="📂 LOAD G-CODE", command=self._on_load_parameters_click)
        load_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        save_btn = ttk.Button(btn_frame, text="💾 EXPORT G-CODE", command=self._start_generation_process)
        save_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        send_btn = ttk.Button(btn_frame, text="Load to Sender", style='Primary.TButton',
                               command=lambda: self._start_generation_process(send_to_sender=True))
        send_btn.pack(side=tk.LEFT, fill=tk.X, expand=True)
        return btn_frame

    def generate_step_values(self, min_val, max_val, step):
        if step <= 0: return [min_val]
        vals = []; curr = min_val
        while curr <= max_val + 1e-9:
            vals.append(round(curr, 6))
            curr += step
        return vals

    def create_pattern(self, params):
        xs = self.generate_step_values(params['x_min'], params['x_max'], params['x_step'])
        ys = self.generate_step_values(params['y_min'], params['y_max'], params['y_step'])
        zs = self.generate_step_values(params['z_min'], params['z_max'], params['z_step'])
        rs = self.generate_step_values(params['rot_min'], params['rot_max'], params['rot_step'])
        for r in rs:
            for z in zs:
                fwd = True
                for y in ys:
                    for x in (xs if fwd else reversed(xs)): yield {'x': x, 'y': y, 'z': z, 'rotation': r}
                    fwd = not fwd

    def _get_params_silently(self):
        try:
            p = {}
            for ax in ['x', 'y', 'z', 'rot']:
                if getattr(self, f"{ax}_symmetric").get():
                    off = abs(float(getattr(self, f"{ax}_offset").get())); p[f'{ax}_min'], p[f'{ax}_max'] = -off, off
                else:
                    p[f'{ax}_min'], p[f'{ax}_max'] = float(getattr(self, f"{ax}_min").get()), float(getattr(self, f"{ax}_max").get())
                p[f'{ax}_step'] = float(getattr(self, f"{ax}_step").get())
            p['travelspeed'], p['pause_time'] = float(self.travelspeed.get()), float(self.pause_time.get())
            return p
        except: return None

    def _check_printer_bounds(self, params):
        if params is None: return ([], 0)
        warns, level = [], 0
        pl = PRINTER_LIMITS
        for ax in ['x', 'y']:
            ext = max(abs(params[f'{ax}_min']), abs(params[f'{ax}_max']))
            if ext > pl[ax]: warns.append(f"{ax.upper()} exceeded"); level = 2
            elif ext > pl[ax] - 10: warns.append(f"{ax.upper()} near limit"); level = max(level, 1)
        if params['z_max'] > pl['z_max']: warns.append("Z exceeded"); level = 2
        return warns, level

    def update_statistics(self, params, total_points, warns, level):
        if not hasattr(self, 'stats_text'): return
        self.stats_text.config(state=tk.NORMAL); self.stats_text.delete(1.0, tk.END)
        self.stats_text.insert(tk.END, f"Total Points: {total_points:,}\n", 'header')
        if warns:
            for w in warns: self.stats_text.insert(tk.END, f"  {w}\n", 'warning' if level==2 else 'amber_warning')
        self.stats_text.config(state=tk.DISABLED)

    def _on_tab_changed(self, event=None):
        self._close_suggestion_popup()
        if self.notebook.select() == str(self.injection_tab):
            # Defer 3D plot creation until first click
            if not self.ax_3d:
                self._create_3d_plot_widgets(self.injection_panel.plot_container)
            self._draw_3d_toolpath()

    def _auto_update_preview(self, event=None):
        params = self._get_params_silently()
        if params:
            tp = self._calculate_total_points(params)
            warns, level = self._check_printer_bounds(params)
            self.update_statistics(params, tp, warns, level)
            self.preview_panel.draw_preview_diagram(params, warns, level)
            self.update_filename_preview()
            
            # Refresh action table to show/hide infeasible flagging
            if hasattr(self, 'injection_panel'):
                self.injection_panel._refresh_action_table()

    def _calculate_total_points(self, p):
        def c(mi, ma, s): return int(math.floor((ma-mi)/s + 1e-9)) + 1 if s > 0 and ma > mi else 1
        return c(p['x_min'], p['x_max'], p['x_step']) * c(p['y_min'], p['y_max'], p['y_step']) * \
               c(p['z_min'], p['z_max'], p['z_step']) * c(p['rot_min'], p['rot_max'], p['rot_step'])

    def update_filename_preview(self, event=None):
        name = self.profile_name_var.get()
        ts = datetime.now().strftime("%Y%m%d-%H%M%S") if self.include_timestamp.get() else ""
        ext = ".csv" if self.export_format.get() == 'csv' else ".gcode"
        fname = f"{name}_{ts}{ext}"
        for p in self.filename_previews:
            try: p.config(text=fname)
            except: pass

    def _on_x_symmetric_toggle(self, derive_values=True): self._toggle_symmetric_widgets(self.x_symmetric.get(), self.x_min_label, self.x_min, self.x_max_label, self.x_max, self.x_offset_label, self.x_offset, self.x_min, self.x_max, self.x_offset, "-50", "50", derive_values)
    def _on_y_symmetric_toggle(self, derive_values=True): self._toggle_symmetric_widgets(self.y_symmetric.get(), self.y_min_label, self.y_min, self.y_max_label, self.y_max, self.y_offset_label, self.y_offset, self.y_min, self.y_max, self.y_offset, "-50", "50", derive_values)
    def _on_z_symmetric_toggle(self, derive_values=True): self._toggle_symmetric_widgets(self.z_symmetric.get(), self.z_min_label, self.z_min, self.z_max_label, self.z_max, self.z_offset_label, self.z_offset, self.z_min, self.z_max, self.z_offset, "0", "100", derive_values)
    def _on_rot_symmetric_toggle(self, derive_values=True): self._toggle_symmetric_widgets(self.rot_symmetric.get(), self.rot_min_label, self.rot_min, self.rot_max_label, self.rot_max, self.rot_offset_label, self.rot_offset, self.rot_min, self.rot_max, self.rot_offset, "0", "0", derive_values)

    def _toggle_symmetric_widgets(self, is_sym, min_l, min_e, max_l, max_e, off_l, off_e, m_ent, mx_ent, o_ent, d_min, d_max, derive):
        if is_sym:
            if derive:
                try: v = max(abs(float(m_ent.get())), abs(float(mx_ent.get()))); o_ent.delete(0, tk.END); o_ent.insert(0, str(v))
                except: pass
            min_l.grid_remove(); min_e.grid_remove(); max_l.grid_remove(); max_e.grid_remove(); off_l.grid(); off_e.grid()
        else:
            if derive:
                try: v = abs(float(o_ent.get())); m_ent.delete(0, tk.END); m_ent.insert(0, str(-v)); mx_ent.delete(0, tk.END); mx_ent.insert(0, str(v))
                except: pass
            min_l.grid(); min_e.grid(); max_l.grid(); max_e.grid(); off_l.grid_remove(); off_e.grid_remove()
        self._auto_update_preview()

    def _create_3d_plot_widgets(self, parent):
        try:
            from matplotlib.figure import Figure
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
            self.matplotlib_imported = True
            
            # Rotation Controls Frame (mirrored from sender_panel)
            controls_frame = ttk.Frame(parent, style='Dark.TFrame')
            controls_frame.pack(side=tk.TOP, fill=tk.X, pady=2)
            
            ttk.Label(controls_frame, text="3D VIEW: ", font=self.FONT_BODY_SMALL, foreground=self.COLOR_TEXT_SECONDARY).pack(side=tk.LEFT, padx=5)
            ttk.Button(controls_frame, text="↑", command=lambda: self._rotate_view(elev_change=15), style='ViewCube.TButton', width=2).pack(side=tk.LEFT, padx=1)
            ttk.Button(controls_frame, text="↓", command=lambda: self._rotate_view(elev_change=-15), style='ViewCube.TButton', width=2).pack(side=tk.LEFT, padx=1)
            ttk.Button(controls_frame, text="←", command=lambda: self._rotate_view(azim_change=15), style='ViewCube.TButton', width=2).pack(side=tk.LEFT, padx=1)
            ttk.Button(controls_frame, text="→", command=lambda: self._rotate_view(azim_change=-15), style='ViewCube.TButton', width=2).pack(side=tk.LEFT, padx=1)
            
            ttk.Separator(controls_frame, orient='vertical').pack(side=tk.LEFT, fill=tk.Y, padx=4, pady=2)
            
            ttk.Button(controls_frame, text="TOP", command=lambda: self._set_view(90, -90), style='ViewCube.TButton', width=4).pack(side=tk.LEFT, padx=1)
            ttk.Button(controls_frame, text="FRONT", command=lambda: self._set_view(0, -90), style='ViewCube.TButton', width=6).pack(side=tk.LEFT, padx=1)
            ttk.Button(controls_frame, text="ISO", command=lambda: self._set_view(30, -60), style='ViewCube.TButton', width=4).pack(side=tk.LEFT, padx=1)

            self.fig_3d = Figure(figsize=(5, 4), dpi=100, facecolor=self.COLOR_BG)
            self.ax_3d = self.fig_3d.add_subplot(111, projection='3d')
            self.canvas_3d = FigureCanvasTkAgg(self.fig_3d, parent)
            self.canvas_3d.get_tk_widget().pack(fill=tk.BOTH, expand=True)
            self._style_3d_plot()
            self._set_view(30, -60)
            self._draw_3d_toolpath()
        except ImportError: pass

    def _set_view(self, elev, azim):
        if not self.ax_3d: return
        self.ax_3d.view_init(elev=elev, azim=azim)
        self.canvas_3d.draw()

    def _rotate_view(self, elev_change=0, azim_change=0):
        if not self.ax_3d: return
        new_elev = max(-90, min(90, self.ax_3d.elev + elev_change))
        new_azim = self.ax_3d.azim + azim_change
        self.ax_3d.view_init(elev=new_elev, azim=new_azim)
        self.canvas_3d.draw()

    def _style_3d_plot(self):
        if not self.ax_3d: return
        self.ax_3d.set_facecolor(self.COLOR_BLACK); self.fig_3d.patch.set_facecolor(self.COLOR_BG)
        self.ax_3d.xaxis.set_pane_color((0,0,0,0)); self.ax_3d.yaxis.set_pane_color((0,0,0,0)); self.ax_3d.zaxis.set_pane_color((0,0,0,0))
        self.ax_3d.grid(False)
        self.ax_3d.xaxis.pane.set_edgecolor('none'); self.ax_3d.yaxis.pane.set_edgecolor('none'); self.ax_3d.zaxis.pane.set_edgecolor('none')
        self.ax_3d.tick_params(colors=self.COLOR_TEXT_PRIMARY); self.fig_3d.tight_layout(pad=0)

    def _draw_3d_toolpath(self):
        if not self.ax_3d: return
        
        # Only render if we are on the Command Injection tab
        if self.notebook.select() != str(self.injection_tab):
            return

        # Store current view angles to restore them
        curr_elev, curr_azim = self.ax_3d.elev, self.ax_3d.azim
        
        self.ax_3d.clear(); self._style_3d_plot()
        self.ax_3d.view_init(elev=curr_elev, azim=curr_azim)
        
        p = self._get_params_silently()
        if not p: self.canvas_3d.draw(); return
        
        pts = list(self.create_pattern(p))
        infeasible_mask = self._get_infeasible_actions_mask()

        if pts:
            xs_all = [pt['x'] for pt in pts]; ys_all = [pt['y'] for pt in pts]; zs_all = [pt['z'] for pt in pts]
            padding = 5
            self.ax_3d.set_xlim(min(xs_all)-padding, max(xs_all)+padding)
            self.ax_3d.set_ylim(min(ys_all)-padding, max(ys_all)+padding)
            self.ax_3d.set_zlim(min(zs_all)-padding, max(zs_all)+padding)
            first_angle = pts[0]['rotation']
            single_angle_pts = [pt for pt in pts if pt['rotation'] == first_angle]
            scrub_limit = getattr(self.injection_panel, 'scrub_var', None)
            if scrub_limit:
                limit_idx = max(1, int(len(single_angle_pts) * scrub_limit.get()))
                visible_pts = single_angle_pts[:limit_idx]
            else: visible_pts = single_angle_pts
            if visible_pts:
                self.ax_3d.plot([pt['x'] for pt in visible_pts], [pt['y'] for pt in visible_pts], [pt['z'] for pt in visible_pts], 
                                color=self.COLOR_ACCENT_CYAN, alpha=0.6, linewidth=1.5)
            
            self._update_scrubber_bookmarks(single_angle_pts)

        # Draw all feasible trigger points
        for i, a in enumerate(self.hub_actions):
            # Skip if infeasible (not on path or empty)
            if i < len(infeasible_mask) and infeasible_mask[i]:
                continue

            try:
                tx_s, ty_s, tz_s = a['x'].get().strip(), a['y'].get().strip(), a['z'].get().strip()
                tr_s = a['rot'].get().strip()
                
                # Double check empty strings (infeasible_mask should have caught them but just in case)
                if not all([tx_s, ty_s, tz_s, tr_s]): continue
                
                tx, ty, tz = float(tx_s), float(ty_s), float(tz_s)
                
                # Show all feasible dots at full color/opacity
                self.ax_3d.scatter([tx], [ty], [tz], color=a['color'], s=150, 
                                   edgecolors='white', linewidths=1.5,
                                   depthshade=False, alpha=1.0, zorder=20)
            except: pass
        self.canvas_3d.draw()

    def _update_scrubber_bookmarks(self, pts):
        """Draws colored markers on the scrubber bar indicating trigger points."""
        panel = getattr(self, 'injection_panel', None)
        if not panel or not hasattr(panel, 'bookmark_canvas'): return
        
        canv = panel.bookmark_canvas
        canv.delete("all")
        
        w = canv.winfo_width()
        if w <= 1: return 
        
        pts_coords = [(round(p['x'], 4), round(p['y'], 4), round(p['z'], 4)) for p in pts]
        total = len(pts_coords)
        if total == 0: return

        for a in self.hub_actions:
            try:
                tx, ty, tz = float(a['x'].get()), float(a['y'].get()), float(a['z'].get())
                t_rounded = (round(tx, 4), round(ty, 4), round(tz, 4))
                
                if t_rounded in pts_coords:
                    idx = pts_coords.index(t_rounded)
                    rel_x = (idx / total) * w
                    # Draw "v" shape
                    points = [rel_x-4, 0, rel_x+4, 0, rel_x, 8]
                    canv.create_polygon(points, fill=a['color'], outline="white", width=1)
            except: continue

    def _load_last_parameters(self):
        script_dir = os.path.dirname(os.path.abspath(__file__))
        profile_path = os.path.join(script_dir, 'last_scan_profile.json')
        if os.path.exists(profile_path):
            try:
                with open(profile_path, 'r') as f: s = json.load(f)
                if 'profile_name' in s: self.profile_name_var.set(s['profile_name'])
            except: pass

    def _show_tutorial_popup(self):
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            tutorial_path = os.path.join(script_dir, "HINTS_TUTORIAL.txt")
            with open(tutorial_path, "r") as f:
                content = f.read()
        except Exception:
            messagebox.showinfo("Tutorial", "Define scan area in tab 1, add triggers in tab 2.")
            return

        BG_DARK = "#000000"
        ACCENT  = utils.COLOR_ACCENT_CYAN

        popup = tk.Toplevel(self.root)
        popup.title("Operating Guide")
        popup.geometry("700x520")
        popup.grab_set()
        popup.configure(bg=utils.COLOR_BG)

        header = tk.Frame(popup, bg=BG_DARK, height=60)
        header.pack(side=tk.TOP, fill=tk.X)
        tk.Label(header, text="4-AXIS SCANNING SYSTEM — OPERATING GUIDE",
                 font=('Rajdhani', 15, 'bold'), bg=BG_DARK, fg=ACCENT).pack(pady=15)

        content_frame = tk.Frame(popup, bg=utils.COLOR_BG)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(10, 0))

        scrollbar = ttk.Scrollbar(content_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        txt_area = tk.Text(
            content_frame, font=("Consolas", 11), bg=BG_DARK, fg="#ffffff",
            insertbackground=ACCENT, padx=15, pady=15, bd=0,
            highlightthickness=1, highlightbackground=utils.COLOR_BORDER,
            yscrollcommand=scrollbar.set, wrap=tk.WORD
        )
        txt_area.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=txt_area.yview)
        txt_area.insert(tk.END, content)
        txt_area.configure(state=tk.DISABLED)

        footer = tk.Frame(popup, bg=utils.COLOR_BG)
        footer.pack(side=tk.BOTTOM, fill=tk.X)
        btn = tk.Button(footer, text="CLOSE", font=("Segoe UI", 10, "bold"),
                        bg=ACCENT, fg=BG_DARK, activebackground="#ffffff", activeforeground=BG_DARK,
                        bd=0, padx=20, pady=8, command=popup.destroy)
        btn.pack(pady=10)
        btn.bind("<Return>", lambda _e: popup.destroy())

    def _start_generation_process(self, send_to_sender=False):
        p = self._get_params_silently()
        if not p: messagebox.showerror("Error", "Invalid parameters"); return
        
        # Check for infeasible actions and warn user
        infeasible_mask = self._get_infeasible_actions_mask()
        if any(infeasible_mask):
            count = sum(infeasible_mask)
            msg = f"WARNING: {count} hub action(s) are NOT on the current scan path.\n\n" \
                  "Infeasible actions will be EXCLUDED from the generated G-code.\n\n" \
                  "Do you want to proceed anyway?"
            if not messagebox.askyesno("Infeasible Actions", msg):
                return

        total_p = self._calculate_total_points(p)
        gcode = list(self.create_gcode(self.create_pattern(p), p, total_p))
        if send_to_sender:
            import tempfile; tf = os.path.join(tempfile.gettempdir(), "last_pattern.gcode")
            with open(tf, 'w') as f:
                for line in gcode: f.write(line + "\n")
            if self.on_send_to_sender: self.on_send_to_sender(tf)
        else:
            # Suggest filename based on profile settings
            name = self.profile_name_var.get()
            ts = datetime.now().strftime("%Y%m%d-%H%M%S") if self.include_timestamp.get() else ""
            ext = ".csv" if self.export_format.get() == 'csv' else ".gcode"
            suggested_fname = f"{name}_{ts}{ext}" if ts else f"{name}{ext}"

            f = filedialog.asksaveasfilename(
                initialfile=suggested_fname,
                defaultextension=ext,
                filetypes=[("G-code files", "*.gcode"), ("CSV files", "*.csv")]
            )
            if f:
                _, ext = os.path.splitext(f)
                if not ext: f += ".gcode"
                with open(f, 'w') as out:
                    if f.endswith('.csv'):
                        # Assuming G-code generator logic here needs to adapt to CSV if necessary.
                        # Based on request, simple extension append is requested first.
                        for line in gcode: out.write(line + "\n")
                    else:
                        for line in gcode: out.write(line + "\n")

    def _on_load_parameters_click(self):
        f = filedialog.askopenfilename(filetypes=[("G-code files", "*.gcode"), ("All files", "*.*")])
        if f: self._load_parameters_from_gcode_file(f)

    def _load_parameters_from_gcode_file(self, file_path):
        try:
            with open(file_path, 'r') as f:
                header_json = None
                for _ in range(20): # Only scan first 20 lines
                    line = f.readline()
                    if not line: break
                    if line.startswith("; PARAMS:"):
                        header_json = line.replace("; PARAMS:", "").strip()
                        break
                
                if not header_json:
                    messagebox.showerror("Error", "No parameter header found in this G-code file.")
                    return

                data = json.loads(header_json)
                
                # 1. Update basic state variables
                if 'profile_name' in data: self.profile_name_var.set(data['profile_name'])
                if 'x_symmetric' in data: self.x_symmetric.set(data['x_symmetric'])
                if 'y_symmetric' in data: self.y_symmetric.set(data['y_symmetric'])
                if 'z_symmetric' in data: self.z_symmetric.set(data['z_symmetric'])
                if 'rot_symmetric' in data: self.rot_symmetric.set(data['rot_symmetric'])
                
                # 2. Update Entry widgets (numeric parameters)
                def set_val(attr, val):
                    if hasattr(self, attr) and val is not None:
                        ent = getattr(self, attr)
                        ent.delete(0, tk.END)
                        ent.insert(0, str(val))

                for ax in ['x', 'y', 'z', 'rot']:
                    set_val(f"{ax}_min", data.get(f"{ax}_min"))
                    set_val(f"{ax}_max", data.get(f"{ax}_max"))
                    set_val(f"{ax}_step", data.get(f"{ax}_step"))
                    set_val(f"{ax}_offset", data.get(f"{ax}_offset"))
                
                set_val("travelspeed", data.get("travelspeed"))
                set_val("pause_time", data.get("pause_time"))

                # 3. Update Hub Actions
                if 'hub_actions' in data:
                    self.hub_actions = []
                    for a_data in data['hub_actions']:
                        color_idx = len(self.hub_actions) % len(self.ACTION_COLORS)
                        new_action = {
                            'x': tk.StringVar(value=a_data.get('x', "")),
                            'y': tk.StringVar(value=a_data.get('y', "")),
                            'z': tk.StringVar(value=a_data.get('z', "")),
                            'rot': tk.StringVar(value=a_data.get('rot', "")),
                            'type': tk.StringVar(value=a_data.get('type', self.hub_action_types[0])),
                            'color': self.ACTION_COLORS[color_idx]
                        }
                        self.hub_actions.append(new_action)

                # 3b. Update Repeated Hub Actions
                if 'repeated_hub_actions' in data:
                    self.repeated_hub_actions = []
                    for a_data in data['repeated_hub_actions']:
                        self.repeated_hub_actions.append({
                            'timing': tk.StringVar(value=a_data.get('timing', 'Before Measurement')),
                            'type':   tk.StringVar(value=a_data.get('type', self.hub_action_types[0])),
                        })

                # 4. Refresh UI and Sync Symmetry
                self._on_x_symmetric_toggle(derive_values=False)
                self._on_y_symmetric_toggle(derive_values=False)
                self._on_z_symmetric_toggle(derive_values=False)
                self._on_rot_symmetric_toggle(derive_values=False)
                
                self._sort_hub_actions()

                if hasattr(self, 'injection_panel'):
                    self.injection_panel._refresh_action_table()
                
                self._auto_update_preview()
                self._draw_3d_toolpath()
                
                messagebox.showinfo("Success", f"Parameters loaded from {os.path.basename(file_path)}")

        except Exception as e:
            messagebox.showerror("Load Error", f"Failed to parse G-code header: {str(e)}")

    def create_gcode(self, pattern_gen, params, total_p):
        # 0. Generate Header JSON for round-tripping
        header_data = {
            "profile_name": self.profile_name_var.get(),
            "x_symmetric": self.x_symmetric.get(),
            "y_symmetric": self.y_symmetric.get(),
            "z_symmetric": self.z_symmetric.get(),
            "rot_symmetric": self.rot_symmetric.get(),
            "travelspeed": self.travelspeed.get(),
            "pause_time": self.pause_time.get(),
            "hub_actions": [
                {
                    "x": a['x'].get(), "y": a['y'].get(), "z": a['z'].get(),
                    "rot": a['rot'].get(), "type": a['type'].get()
                } for a in self.hub_actions
            ],
            "repeated_hub_actions": [
                {"timing": a['timing'].get(), "type": a['type'].get()}
                for a in self.repeated_hub_actions
            ]
        }
        # Add all axis fields
        for ax in ['x', 'y', 'z', 'rot']:
            for suffix in ['min', 'max', 'step', 'offset']:
                attr = f"{ax}_{suffix}"
                if hasattr(self, attr):
                    header_data[attr] = getattr(self, attr).get()
        
        yield "; SEED Pattern"
        yield f"; PARAMS: {json.dumps(header_data)}"
        yield ""
        
        # 1. Prepare triggers for fast lookup (round to 4 decimal places for matching)
        triggers = {} # (x, y, z, r) -> [actions]
        for a in self.hub_actions:
            try:
                def get_v(var):
                    s = var.get().strip()
                    return float(s if s else "0")
                
                x, y, z = get_v(a['x']), get_v(a['y']), get_v(a['z'])
                r = get_v(a['rot'])
                key = (round(x, 4), round(y, 4), round(z, 4), round(r, 4))
                if key not in triggers: triggers[key] = []
                triggers[key].append({'type': a['type'].get()})
            except ValueError: continue

        # Shared map used for both repeated and coordinate-based hub command serialization
        # param names must match rlm_proj_vivigo.py exactly
        hub_cmd_map = {
            "WPT Start":    "param=WPT Start value=true",
            "WPT Stop":     "param=WPT Stop value=true",
            "DCDC Enable":  "param=DCDC Enable value=true",
            "DCDC Disable": "param=DCDC Enable value=false",
        }

        # Build repeated action comment lines once (before/after every point)
        before_lines = []
        after_lines = []
        for a in self.repeated_hub_actions:
            t = a['timing'].get()
            atype = a['type'].get()
            if atype == "WAIT":
                comment = f"; {'WAIT_PRE' if t == 'Before Measurement' else 'WAIT'} seconds={params.get('pause_time', 1.0)}"
            else:
                prefix = "HUB_CMD_PRE" if t == "Before Measurement" else "HUB_CMD"
                comment = f"; {prefix} {hub_cmd_map.get(atype, '')}"
            if t == "Before Measurement":
                before_lines.append(comment)
            else:
                after_lines.append(comment)

        matched_count = 0
        # Iterate through pattern; before_lines appear after the G1 so the
        # serial_engine look-ahead (cmd_index+1) can execute them before the measurement.
        for pt in pattern_gen:
            yield f"G1 X{pt['x']:.3f} Y{pt['y']:.3f} Z{pt['z']:.3f} E{pt['rotation']:.3f} F{params['travelspeed']:.0f}"
            for line in before_lines:
                yield line

            # Coordinate-based triggers (injected after G1, treated as after-measurement)
            key = (round(pt['x'], 4), round(pt['y'], 4), round(pt['z'], 4), round(pt['rotation'], 4))
            if key in triggers:
                for action in triggers[key]:
                    matched_count += 1
                    atype = action['type']
                    if atype == "WAIT":
                        yield f"; WAIT seconds={params.get('pause_time', 1.0)}"
                    else:
                        yield f"; HUB_CMD {hub_cmd_map.get(atype, '')}"

            for line in after_lines:
                yield line

        yield f"; TOTAL_MATCHED_TRIGGERS: {matched_count}"

    def _on_canvas_resize(self, event): self._auto_update_preview()

    def _add_repeated_action_row(self):
        self.repeated_hub_actions.append({
            'timing': tk.StringVar(value='Before Measurement'),
            'type':   tk.StringVar(value=self.hub_action_types[0]),
        })
        if hasattr(self, 'injection_panel'):
            self.injection_panel._refresh_action_table()

    def _remove_repeated_action_row(self, index):
        if 0 <= index < len(self.repeated_hub_actions):
            self.repeated_hub_actions.pop(index)
            if hasattr(self, 'injection_panel'):
                self.injection_panel._refresh_action_table()

    def _sort_hub_actions(self):
        """
        Sorts self.hub_actions based on the order they will be encountered during the scan.
        Uses the generated pattern as the reference for chronological order.
        """
        if not self.hub_actions:
            return

        params = self._get_params_silently()
        if not params:
            return

        # 1. Generate the sequence of coordinates for the entire scan
        pattern = list(self.create_pattern(params))
        
        # 2. Create a lookup map: (x,y,z,r) -> scan_index
        # We round to 4 decimal places to match the logic used in G-code generation
        coord_map = {
            (round(pt['x'], 4), round(pt['y'], 4), round(pt['z'], 4), round(pt['rotation'], 4)): i 
            for i, pt in enumerate(pattern)
        }

        def get_trigger_index(action):
            try:
                def get_v(var):
                    s = var.get().strip()
                    return float(s if s else "0")
                
                x, y, z = get_v(action['x']), get_v(action['y']), get_v(action['z'])
                r = get_v(action['rot'])
                key = (round(x, 4), round(y, 4), round(z, 4), round(r, 4))
                
                # Return the index in scan, or a very high number if not in scan (moves to end)
                return coord_map.get(key, float('inf'))
            except (ValueError, AttributeError):
                return float('inf')

        # 3. Perform the sort
        self.hub_actions.sort(key=get_trigger_index)

    def _get_infeasible_actions_mask(self):
        """
        Returns a list of booleans indicating which hub_actions are infeasible.
        True = Infeasible (not on current scan path).
        """
        if not self.hub_actions:
            return []

        params = self._get_params_silently()
        if not params:
            return [False] * len(self.hub_actions)

        pattern = list(self.create_pattern(params))
        pts_coords = {(round(pt['x'], 4), round(pt['y'], 4), round(pt['z'], 4), round(pt['rotation'], 4)) for pt in pattern}

        mask = []
        for a in self.hub_actions:
            try:
                tx_s, ty_s, tz_s = a['x'].get().strip(), a['y'].get().strip(), a['z'].get().strip()
                tr_s = a['rot'].get().strip()
                
                # If any field is empty, the action is infeasible
                if not all([tx_s, ty_s, tz_s, tr_s]):
                    mask.append(True)
                    continue
                
                x, y, z = float(tx_s), float(ty_s), float(tz_s)
                r = float(tr_s)
                key = (round(x, 4), round(y, 4), round(z, 4), round(r, 4))
                mask.append(key not in pts_coords)
            except (ValueError, AttributeError):
                mask.append(True)
        return mask

    def _get_unique_axis_values(self, axis):
        p = self._get_params_silently()
        if not p: return []
        p_key = (p['x_min'], p['x_max'], p['x_step'], p['y_min'], p['y_max'], p['y_step'], p['z_min'], p['z_max'], p['z_step'], p['rot_min'], p['rot_max'], p['rot_step'])
        if self._cache_params_key != p_key: self._axis_cache = {}; self._cache_params_key = p_key
        if axis not in self._axis_cache:
            vals = self.generate_step_values(p[f'{axis}_min'], p[f'{axis}_max'], p[f'{axis}_step'])
            self._axis_cache[axis] = {'numeric': vals, 'strings': [f"{v:g}" for v in vals]}
        return self._axis_cache[axis]

    def _show_coordinate_suggestions(self, entry, axis, var):
        if self.root.focus_get() != entry: return
        val_str = var.get().strip()
        cache = self._get_unique_axis_values(axis)
        if not cache: return
        matches = [v for v in cache['strings'] if val_str in v]
        is_prox = False
        if not matches:
            try:
                target = float(val_str)
                closest_nums = sorted(cache['numeric'], key=lambda x: abs(x - target))[:2]
                matches = [f"{v:g}" for v in sorted(closest_nums)]; is_prox = True; entry.configure(foreground=self.COLOR_ACCENT_AMBER)
            except: 
                if val_str == "": matches = cache['strings'][:12]
                else: return
        if not matches: return
        if not is_prox: entry.configure(foreground=self.COLOR_ACCENT_CYAN)
        self._display_suggestion_popup(entry, matches, var, is_prox)

    def _close_suggestion_popup(self):
        """Safely destroys the suggestion popup and clears references."""
        if hasattr(self, '_s_popup') and self._s_popup:
            try:
                if self._s_popup.winfo_exists():
                    self._s_popup.destroy()
            except: pass
            self._s_popup = None

    def _display_suggestion_popup(self, entry, matches, var, is_prox):
        self._close_suggestion_popup()
        
        x, y = entry.winfo_rootx(), entry.winfo_rooty() + entry.winfo_height()
        self._s_popup = tk.Toplevel(self.root)
        self._s_popup.wm_overrideredirect(True)
        self._s_popup.geometry(f"+{x}+{y}")
        self._s_popup.configure(bg=self.COLOR_BLACK)
        self._s_popup.attributes("-topmost", True)

        lb = tk.Listbox(self._s_popup, bg=self.COLOR_BLACK, fg=self.COLOR_ACCENT_AMBER if is_prox else self.COLOR_ACCENT_CYAN,
                         font=self.FONT_MONO, borderwidth=1, highlightthickness=0, height=min(len(matches), 10), selectmode=tk.SINGLE)
        lb.pack()
        for m in matches: lb.insert(tk.END, f"{'≈ ' if is_prox else ''}{m}")

        def on_select(e):
            if lb.curselection():
                val = matches[lb.curselection()[0]]
                var.set(val)
                self._close_suggestion_popup()
                self._draw_3d_toolpath()

        lb.bind("<Button-1>", on_select)
        
        def check_click(event):
            # Check if popup still exists before querying it
            if not self._s_popup or not self._s_popup.winfo_exists():
                return
            
            # Identify the clicked widget
            clicked_widget = self.root.winfo_containing(event.x_root, event.y_root)
            if clicked_widget == entry:
                return # Ignore clicks on the owner entry (keeps dropdown open when clicking focused entry)

            try:
                # If click was not inside popup, close it
                px, py = self._s_popup.winfo_rootx(), self._s_popup.winfo_rooty()
                pw, ph = self._s_popup.winfo_width(), self._s_popup.winfo_height()
                if not (px < event.x_root < px + pw and py < event.y_root < py + ph):
                    self._close_suggestion_popup()
            except tk.TclError:
                self._s_popup = None # Already gone

        # Use a non-additive binding or manage it carefully to avoid orphaned callbacks
        self.root.bind_all("<Button-1>", check_click, add="+")
