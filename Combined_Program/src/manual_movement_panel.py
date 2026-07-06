import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import os
import math
import re
from datetime import datetime
import utils

# Conditional imports for 3D plotting
try:
    from matplotlib.figure import Figure
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    from mpl_toolkits.mplot3d import Axes3D
    MATPLOTLIB_AVAILABLE = True
except ImportError:
    MATPLOTLIB_AVAILABLE = False

class ManualMovementPanel(ttk.Frame):
    """
    A GUI panel for manually creating and visualizing test sequences.
    Features a dynamic table for action sequencing and a 3D preview with an interactive rotation gauge.
    """
    def __init__(self, parent, controller=None):
        super().__init__(parent, style='Dark.TFrame')
        self.c = controller
        
        # Data Model
        self.actions = []
        self.selected_row_idx = tk.IntVar(value=-1)
        self.hub_action_types = ["WPT Start", "WPT Stop", "DCDC Enable", "DCDC Disable", "WAIT"]
        self.action_types = ["Move", "Hub Action", "Pause", "Macro: Move & Measure", "Unrecognized Command"]

        self.interval_modes = ["Num Steps", "Step Size (mm)"]
        self.on_send_to_sender = None
        self.entry_refs = {}
        
        # Animation and Scrubber State
        self.is_animating = False
        self._animate_job = None
        self.scrub_var = tk.DoubleVar(value=1.0)
        
        # Color Palette for markers
        self.ACTION_COLORS = [
            utils.COLOR_ACCENT_GREEN,  # Move
            utils.COLOR_ACCENT_GREEN,  # Hub
            utils.COLOR_ACCENT_AMBER,  # Pause
            utils.COLOR_ACCENT_PURPLE, # Macro
            utils.COLOR_ACCENT_RED     # Unrecognized
        ]
        
        self._setup_ui()
        
    def _setup_ui(self):
        # Outer PanedWindow for adjustable Table vs Plot layout
        self.paned = tk.PanedWindow(self, orient=tk.HORIZONTAL, bg=utils.COLOR_BG, sashwidth=4, sashpad=0, borderwidth=0)
        self.paned.pack(fill=tk.BOTH, expand=True)

        # --- LEFT PANEL: Table Area ---
        self.left_panel = ttk.Frame(self.paned, style='Dark.TFrame', width=750)
        self.paned.add(self.left_panel)

        # Title and Hints
        header = tk.Frame(self.left_panel, bg=utils.COLOR_BG)
        header.pack(fill=tk.X, pady=(10, 5), padx=10)
        title_lbl = ttk.Label(header, text="MANUAL SEQUENCE BUILDER", 
                              font=('Rajdhani', 14, 'bold'), foreground=utils.COLOR_ACCENT_CYAN)
        title_lbl.pack(side=tk.LEFT)
        
        hint_btn = tk.Button(header, text="?", font=("Segoe UI", 12, "bold"),
                             bg=utils.COLOR_BG, fg=utils.COLOR_ACCENT_CYAN, bd=0, 
                             activebackground=utils.COLOR_BG, activeforeground="white",
                             cursor="hand2", command=self._show_hints_popup)
        hint_btn.pack(side=tk.LEFT, padx=10)

        note_lbl = ttk.Label(self.left_panel, text="Measurements automatically recorded at every defined point.",
                             font=('Segoe UI', 9), foreground=utils.COLOR_TEXT_SECONDARY)
        note_lbl.pack(anchor=tk.W, padx=12, pady=(0, 4))

        # Table Container
        self.table_frame = ttk.Frame(self.left_panel, style='Dark.TFrame')
        self.table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Scrollable Area for Table
        self.canvas = tk.Canvas(self.table_frame, bg=utils.COLOR_BG, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self.table_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas, style='Dark.TFrame')

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfig(self.canvas_window, width=e.width))
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Controls Footer
        self.footer = ttk.Frame(self.left_panel, style='Dark.TFrame')
        self.footer.pack(fill=tk.X, padx=10, pady=10)

        self.load_btn = ttk.Button(self.footer, text="📂 LOAD G-CODE", command=self._load_gcode)
        self.load_btn.pack(side=tk.LEFT, padx=5)

        self.export_btn = ttk.Button(self.footer, text="💾 EXPORT G-CODE", command=self._export_gcode)
        self.export_btn.pack(side=tk.LEFT, padx=5)

        self.add_btn = ttk.Button(self.footer, text="+ ADD ACTION (Shift+Enter)", style='Amber.TButton', command=self._add_action_row)
        self.add_btn.pack(side=tk.LEFT, padx=5)

        self.send_btn = ttk.Button(self.footer, text="❱ SEND TO SENDER", style='Primary.TButton', command=self._send_to_sender)
        self.send_btn.pack(side=tk.LEFT, padx=5)
        
        self.bind("<Map>", self._on_map)
        self.bind("<Unmap>", self._on_unmap)

        # --- RIGHT PANEL: Notebook (Visualizer + GCode Editor) ---
        self.right_notebook = ttk.Notebook(self.paned)
        self.paned.add(self.right_notebook)

        # =========================================================
        # TAB 1: Visualizer
        # =========================================================
        self.viz_tab = ttk.Frame(self.right_notebook, style='Dark.TFrame')
        self.right_notebook.add(self.viz_tab, text="Visualizer")
        
        # Split viz_tab internally to hold the 2D tools and 3D preview
        self.viz_paned = tk.PanedWindow(self.viz_tab, orient=tk.HORIZONTAL, bg=utils.COLOR_BG, sashwidth=4, sashpad=0, borderwidth=0)
        self.viz_paned.pack(fill=tk.BOTH, expand=True)

        # --- Sub-panel: Graphical Input (Left half of Visualizer) ---
        self.mid_panel = ttk.Frame(self.viz_paned, style='Dark.TFrame', width=300)
        self.viz_paned.add(self.mid_panel)

        mid_title = ttk.Label(self.mid_panel, text="GRAPHICAL INSERT", 
                              font=('Rajdhani', 14, 'bold'), foreground=utils.COLOR_ACCENT_CYAN)
        mid_title.pack(pady=(10, 5), padx=10, anchor="w")

        # 2D XY Canvas & Z Canvas Container
        canvas_container = ttk.Frame(self.mid_panel, style='Dark.TFrame')
        canvas_container.pack(fill=tk.X, padx=10, pady=5)
        
        xy_container = ttk.LabelFrame(canvas_container, text="XY Plane", style='TLabelframe')
        xy_container.pack(side=tk.LEFT, padx=(0, 5))
        
        self.xy_canvas_size = 200
        self.xy_canvas = tk.Canvas(xy_container, width=self.xy_canvas_size, height=self.xy_canvas_size, 
                                   bg=utils.COLOR_BLACK, highlightthickness=1, highlightbackground=utils.COLOR_BORDER, cursor="crosshair")
        self.xy_canvas.pack(pady=5, padx=5)
        self.xy_canvas.bind("<Button-1>", self._on_xy_click)
        self.xy_canvas.bind("<B1-Motion>", self._on_xy_click)
        self._draw_xy_grid()

        z_container = ttk.LabelFrame(canvas_container, text="Z", style='TLabelframe')
        z_container.pack(side=tk.LEFT, fill=tk.Y)
        
        self.z_canvas = tk.Canvas(z_container, width=30, height=self.xy_canvas_size, 
                                  bg=utils.COLOR_BLACK, highlightthickness=1, highlightbackground=utils.COLOR_BORDER, cursor="sb_v_double_arrow")
        self.z_canvas.pack(pady=5, padx=5, fill=tk.Y, expand=True)
        self.z_canvas.bind("<Button-1>", self._on_z_canvas_click)
        self.z_canvas.bind("<B1-Motion>", self._on_z_canvas_click)
        self.z_var = tk.DoubleVar(value=0.0)
        self._draw_z_canvas_marker()

        # Rotation Gauge
        self.gauge_container = ttk.Frame(self.mid_panel, style='Dark.TFrame', height=250)
        self.gauge_container.pack(fill=tk.X, padx=10, pady=10)
        self._setup_rotation_gauge()

        # --- Sub-panel: Visualization Area (Right half of Visualizer) ---
        self.right_panel = ttk.Frame(self.viz_paned, style='Dark.TFrame', width=300)
        self.viz_paned.add(self.right_panel)

        # Visualization Title and Controls
        viz_header = ttk.Frame(self.right_panel, style='Dark.TFrame')
        viz_header.pack(fill=tk.X, pady=(10, 5), padx=10)
        
        viz_title = ttk.Label(viz_header, text="PATH PREVIEW", 
                              font=('Rajdhani', 14, 'bold'), foreground=utils.COLOR_ACCENT_CYAN)
        viz_title.pack(side=tk.LEFT)

        controls_frame = ttk.Frame(viz_header, style='Dark.TFrame')
        controls_frame.pack(side=tk.LEFT, padx=(20, 0))
        ttk.Button(controls_frame, text="↑", command=lambda: self._rotate_view(elev_change=15), width=2).pack(side=tk.LEFT, padx=1)
        ttk.Button(controls_frame, text="↓", command=lambda: self._rotate_view(elev_change=-15), width=2).pack(side=tk.LEFT, padx=1)
        ttk.Button(controls_frame, text="←", command=lambda: self._rotate_view(azim_change=15), width=2).pack(side=tk.LEFT, padx=1)
        ttk.Button(controls_frame, text="→", command=lambda: self._rotate_view(azim_change=-15), width=2).pack(side=tk.LEFT, padx=1)
        ttk.Button(controls_frame, text="Top", command=lambda: self._set_view(elev=90, azim=-90), width=4).pack(side=tk.LEFT, padx=1)
        ttk.Button(controls_frame, text="Iso", command=lambda: self._set_view(elev=30, azim=-60), width=3).pack(side=tk.LEFT, padx=1)

        # 3D Plot Area
        self.plot_container = ttk.Frame(self.right_panel, style='Dark.TFrame')
        self.plot_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        if MATPLOTLIB_AVAILABLE:
            self._setup_3d_plot()
        else:
            ttk.Label(self.plot_container, text="Matplotlib not found.\n3D Preview unavailable.", 
                      font=utils.FONT_BODY, foreground=utils.COLOR_ACCENT_RED).pack(expand=True)

        # Scrubber UI
        self.scrubber_outer = ttk.Frame(self.right_panel, style='Dark.TFrame')
        self.scrubber_outer.pack(fill=tk.X, padx=10, pady=5)

        self.play_btn = tk.Button(self.scrubber_outer, text="▶", font=("Inter", 12),
                                  bg=utils.COLOR_BG, fg=utils.COLOR_ACCENT_GREEN,
                                  activebackground=utils.COLOR_BLACK, activeforeground="#ffffff",
                                  bd=0, command=self._toggle_animation)
        self.play_btn.pack(side=tk.LEFT, padx=5)

        self.scrub_stack = ttk.Frame(self.scrubber_outer, style='Dark.TFrame')
        self.scrub_stack.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.bookmark_canvas = tk.Canvas(self.scrub_stack, height=12, bg=utils.COLOR_BG, highlightthickness=0, cursor="hand2")
        self.bookmark_canvas.pack(fill=tk.X, pady=(0, 2))
        self.bookmark_canvas.bind("<Button-1>", self._on_bookmark_click)

        self.scrubber = ttk.Scale(self.scrub_stack, from_=0.0, to=1.0, variable=self.scrub_var, 
                                  command=lambda v: self._draw_3d_toolpath())
        self.scrubber.pack(fill=tk.X)

        # =========================================================
        # TAB 2: G-Code Editor
        # =========================================================
        self.gcode_tab = ttk.Frame(self.right_notebook, style='Dark.TFrame')
        self.right_notebook.add(self.gcode_tab, text="G-Code Editor")
        
        gcode_header = ttk.Frame(self.gcode_tab, style='Dark.TFrame')
        gcode_header.pack(fill=tk.X, pady=(10, 5), padx=10)
        
        ttk.Label(gcode_header, text="LIVE G-CODE", font=('Rajdhani', 14, 'bold'), 
                  foreground=utils.COLOR_ACCENT_CYAN).pack(side=tk.LEFT)
                  
        self.sync_btn = ttk.Button(gcode_header, text="⟳ Sync to Table", command=self._sync_from_editor)
        self.sync_btn.pack(side=tk.RIGHT)

        editor_frame = ttk.Frame(self.gcode_tab, style='Dark.TFrame')
        editor_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        
        self.gcode_editor = tk.Text(editor_frame, font=("Consolas", 11), bg="#000000", fg="#ffffff", 
                                    insertbackground=utils.COLOR_ACCENT_CYAN, bd=0, 
                                    highlightthickness=1, highlightbackground=utils.COLOR_BORDER,
                                    undo=True)
        self.gcode_editor.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        ed_scroll = ttk.Scrollbar(editor_frame, command=self.gcode_editor.yview)
        ed_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.gcode_editor.config(yscrollcommand=ed_scroll.set)
        
        self._editor_timer = None
        
        # Sync editor when switching to the GCode tab
        self.right_notebook.bind("<<NotebookTabChanged>>", self._on_notebook_tab_changed)

        # Initial Row
        self._refresh_table()

    def _on_notebook_tab_changed(self, event):
        selected_tab_id = self.right_notebook.select()
        tab_text = self.right_notebook.tab(selected_tab_id, "text")
        if tab_text == "G-Code Editor":
            self._sync_to_editor()

    def _on_map(self, event):
        if event.widget == self:
            top = self.winfo_toplevel()
            top.bind("<KeyPress>", self._handle_shortcut, add="+")

    def _on_unmap(self, event):
        if event.widget == self:
            top = self.winfo_toplevel()
            # It's cleaner to handle this by unbinding or checking mapped state in the handler
            pass

    def _show_hints_popup(self):
        """Displays a modal popup with keyboard shortcuts and instructions."""
        popup = tk.Toplevel(self)
        popup.title("Manual Builder Hints & Shortcuts")
        popup.geometry("600x400")
        popup.grab_set()
        popup.configure(bg=utils.COLOR_BG)

        # Style colors
        BG_DARK = "#000000"
        ACCENT  = utils.COLOR_ACCENT_CYAN
        TEXT    = "#ffffff"

        # --- Header ---
        header = tk.Frame(popup, bg=BG_DARK, height=60)
        header.pack(side=tk.TOP, fill=tk.X)
        lbl = tk.Label(header, text="KEYBOARD SHORTCUTS", font=('Rajdhani', 16, 'bold'), bg=BG_DARK, fg=ACCENT)
        lbl.pack(pady=15)

        # --- Content ---
        content_frame = tk.Frame(popup, bg=utils.COLOR_BG)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        scrollbar = ttk.Scrollbar(content_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        txt_area = tk.Text(
            content_frame, font=("Consolas", 11), bg=BG_DARK, fg=TEXT,
            insertbackground=ACCENT, padx=15, pady=15, bd=0,
            highlightthickness=1, highlightbackground=utils.COLOR_BORDER,
            yscrollcommand=scrollbar.set
        )
        txt_area.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=txt_area.yview)

        hints = (
            "GLOBAL SHORTCUTS:\n"
            "These shortcuts work anytime the Manual Movement tab is active,\n"
            "provided you are not actively typing inside a text box.\n\n"
            "  Shift+Enter : Add a new Action row to the end of the sequence.\n"
            "  Delete      : Remove the currently selected row.\n"
            "  Up Arrow    : Move the selected row up one position.\n"
            "  Down Arrow  : Move the selected row down one position.\n\n"
            "FIELD FOCUS SHORTCUTS:\n"
            "Instantly select and edit fields in the highlighted Move/Macro row:\n"
            "  X : Focus the X (mm) coordinate field.\n"
            "  Y : Focus the Y (mm) coordinate field.\n"
            "  Z : Focus the Z (mm) coordinate field.\n"
            "  R : Focus the ROT (°) field.\n"
            "  S : Focus the SPEED (F) field.\n\n"
            "GRAPHICAL INSERT:\n"
            "Clicking on the 2D XY plane, the Z slider, or the Rotation gauge\n"
            "will instantly update the target coordinates of the currently\n"
            "selected Move or Macro row. If no valid row is selected, a new\n"
            "Move row will be inserted automatically.\n"
        )
        txt_area.insert(tk.END, hints)
        txt_area.configure(state=tk.DISABLED)

        # --- Footer ---
        footer = tk.Frame(popup, bg=utils.COLOR_BG, height=60)
        footer.pack(side=tk.BOTTOM, fill=tk.X)
        btn = tk.Button(footer, text="CLOSE", font=("Segoe UI", 10, "bold"),
                        bg=ACCENT, fg=BG_DARK, activebackground="#ffffff", activeforeground=BG_DARK,
                        bd=0, padx=20, pady=8, command=popup.destroy)
        btn.pack(pady=10)
        btn.bind("<Return>", lambda _e: popup.destroy())
        
        popup.update_idletasks()
        pw, ph = popup.winfo_width(), popup.winfo_height()
        rx = self.winfo_toplevel().winfo_x() + (self.winfo_toplevel().winfo_width()  - pw) // 2
        ry = self.winfo_toplevel().winfo_y() + (self.winfo_toplevel().winfo_height() - ph) // 2
        popup.geometry(f"+{rx}+{ry}")
        popup.focus_force()
        btn.focus_set()

    def _handle_shortcut(self, event):
        if not self.winfo_ismapped(): return
        
        widg = self.focus_get()
        if widg == self.gcode_editor:
            return
            
        # Don't steal focus if user is already typing in an entry (except maybe arrows/enter if we want)
        # But actually, 'x', 'y', 'z' shouldn't be trapped if they are typing.
        is_typing = isinstance(widg, (tk.Entry, ttk.Entry, ttk.Combobox))
        
        char = event.char.lower()
        keysym = event.keysym

        if keysym == "Return" and (event.state & 0x0001): # Shift+Enter
            self._add_action_row()
            return "break"
            
        if is_typing and char in ['x', 'y', 'z', 'r', 's']:
            return # Let them type normal letters
            
        idx = self.selected_row_idx.get()
        if idx < 0 or idx >= len(self.actions): return
        
        if keysym == "Delete":
            self._remove_action_row(idx)
            return "break"
            
        # Avoid stealing up/down from comboboxes, but handle if not typing
        if not is_typing:
            if keysym == "Up":
                self._move_action(idx, -1)
                return "break"
            elif keysym == "Down":
                self._move_action(idx, 1)
                return "break"

            if char in ['x', 'y', 'z']:
                ent = self.entry_refs.get((idx, char))
                if ent: 
                    ent.focus_set()
                    ent.select_range(0, tk.END)
                return "break"
            elif char == 'r':
                ent = self.entry_refs.get((idx, 'rot'))
                if ent: 
                    ent.focus_set()
                    ent.select_range(0, tk.END)
                return "break"
            elif char == 's':
                ent = self.entry_refs.get((idx, 'speed'))
                if ent: 
                    ent.focus_set()
                    ent.select_range(0, tk.END)
                return "break"

    def _draw_xy_grid(self):
        c = self.xy_canvas
        w = self.xy_canvas_size
        c.delete("all")
        # Draw grid lines (every 10% of travel)
        for i in range(1, 10):
            pos = i * (w / 10)
            c.create_line(pos, 0, pos, w, fill=utils.COLOR_BORDER, dash=(2, 4))
            c.create_line(0, pos, w, pos, fill=utils.COLOR_BORDER, dash=(2, 4))
        
        # Crosshair at center
        c.create_line(w/2, 0, w/2, w, fill="#444444")
        c.create_line(0, w/2, w, w/2, fill="#444444")

    def _on_xy_click(self, event):
        w = self.xy_canvas_size
        x_max = utils.PRINTER_BOUNDS['x_max']
        y_max = utils.PRINTER_BOUNDS['y_max']

        # Calculate real coordinates
        x_mm = (event.x / w) * x_max
        # Y axis on canvas is inverted (0 is top, max is bottom)
        y_mm = ((w - event.y) / w) * y_max

        x_mm = max(0.0, min(x_max, x_mm))
        y_mm = max(0.0, min(y_max, y_mm))
        
        z_mm = self.z_var.get()
        # Get rotation from gauge needle (stored in a hidden variable or read from current row)
        # Since gauge updates the selected row directly, we need to extract it or use a default
        rot = "0.0"
        
        idx = self.selected_row_idx.get()
        if 0 <= idx < len(self.actions) and self.actions[idx]['type'].get() in ["Move", "Macro: Move & Measure"]:
            # Update existing move row
            self.actions[idx]['x'].set(f"{x_mm:.1f}")
            self.actions[idx]['y'].set(f"{y_mm:.1f}")
            self.actions[idx]['z'].set(f"{z_mm:.1f}")
            rot = self.actions[idx]['rot'].get()
            if self.actions[idx]['type'].get() == "Macro: Move & Measure":
                self._schedule_macro_regen(idx)
        else:
            # Insert new move row
            if 0 <= idx < len(self.actions):
                rot = self.actions[idx]['rot'].get() if 'rot' in self.actions[idx] else "0.0"
                
            self._insert_action(idx) # This handles index increment and inheritance
            new_idx = self.selected_row_idx.get()
            self.actions[new_idx]['type'].set("Move")
            self.actions[new_idx]['x'].set(f"{x_mm:.1f}")
            self.actions[new_idx]['y'].set(f"{y_mm:.1f}")
            self.actions[new_idx]['z'].set(f"{z_mm:.1f}")
            
        self._refresh_table()
        self._draw_3d_toolpath()
        self._draw_xy_marker(event.x, event.y)

    def _draw_xy_marker(self, cx, cy):
        self._draw_xy_grid()
        r = 4
        self.xy_canvas.create_oval(cx-r, cy-r, cx+r, cy+r, fill=utils.COLOR_ACCENT_CYAN, outline="white")

    def _on_z_canvas_click(self, event):
        h = self.xy_canvas_size
        if h <= 1: return
        click_y = max(0, min(h, event.y))
        
        max_z = utils.PRINTER_BOUNDS['z_max']
        z_val = ((h - click_y) / h) * max_z
        z_val = max(0.0, min(max_z, z_val))
        
        self.z_var.set(z_val)
        self._draw_z_canvas_marker()
        
        idx = self.selected_row_idx.get()
        if 0 <= idx < len(self.actions) and self.actions[idx]['type'].get() in ["Move", "Macro: Move & Measure"]:
            self.actions[idx]['z'].set(f"{z_val:.1f}")
            self._draw_3d_toolpath()
            if self.actions[idx]['type'].get() == "Macro: Move & Measure":
                self._schedule_macro_regen(idx)

    def _draw_z_canvas_marker(self, event=None):
        c = self.z_canvas
        w = 30
        h = self.xy_canvas_size
        c.delete("all")
        
        max_z = 128.0
        z_val = self.z_var.get()
        
        # Calculate Y pos
        y = h - ((z_val / max_z) * h) if max_z > 0 else h
        y = max(1, min(h - 1, y))
        
        # Draw level line
        c.create_line(0, y, w, y, fill=utils.COLOR_ACCENT_CYAN, width=2)
        # Fill below
        c.create_rectangle(0, y, w, h, fill="#2a3d4d", outline="")
        
        # Label
        c.create_text(w/2, 10 if y > 20 else h - 10, text=f"{z_val:.0f}", fill="white", font=("Inter", 8))

    def _update_selected_z(self):
        idx = self.selected_row_idx.get()
        if 0 <= idx < len(self.actions) and self.actions[idx]['type'].get() in ["Move", "Macro: Move & Measure"]:
            self.actions[idx]['z'].set(f"{self.z_var.get():.1f}")
            self._draw_3d_toolpath()
            if self.actions[idx]['type'].get() == "Macro: Move & Measure":
                self._schedule_macro_regen(idx)

    def _rotate_view(self, elev_change=0, azim_change=0):
        if not MATPLOTLIB_AVAILABLE: return
        new_elev = max(-90, min(90, self.ax.elev + elev_change))
        new_azim = self.ax.azim + azim_change
        self.ax.view_init(elev=new_elev, azim=new_azim)
        self.canvas_plt.draw()

    def _set_view(self, elev, azim):
        if not MATPLOTLIB_AVAILABLE: return
        self.ax.view_init(elev=elev, azim=azim)
        self.canvas_plt.draw()

    def _toggle_animation(self):
        """Starts or stops the scrubber animation."""
        if self.is_animating:
            self.is_animating = False
            self.play_btn.config(text="▶", fg=utils.COLOR_ACCENT_GREEN)
            if self._animate_job:
                self.after_cancel(self._animate_job)
                self._animate_job = None
        else:
            if self.scrub_var.get() >= 0.99:
                self.scrub_var.set(0.0)
            self.is_animating = True
            self.play_btn.config(text="⏸", fg=utils.COLOR_ACCENT_AMBER)
            self._animate_step()

    def _animate_step(self):
        """Increments the scrubber and redraws the path."""
        if not self.is_animating: return
        curr = self.scrub_var.get()
        if curr < 1.0:
            new_val = min(1.0, curr + 0.005)
            self.scrub_var.set(new_val)
            self._draw_3d_toolpath()
            self._animate_job = self.after(30, self._animate_step)
        else:
            self.is_animating = False
            self.play_btn.config(text="▶", fg=utils.COLOR_ACCENT_GREEN)
            self._animate_job = None

    def _on_bookmark_click(self, event):
        """Moves the scrubber to the clicked position on the canvas."""
        w = self.bookmark_canvas.winfo_width()
        if w > 0:
            rel_pos = max(0.0, min(1.0, event.x / w))
            self.scrub_var.set(rel_pos)
            self._draw_3d_toolpath()

    def _setup_3d_plot(self):
        self.fig = Figure(figsize=(5, 5), dpi=100, facecolor=utils.COLOR_BG)
        self.ax = self.fig.add_subplot(111, projection='3d')
        self.canvas_plt = FigureCanvasTkAgg(self.fig, self.plot_container)
        self.canvas_plt.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        self._style_plot()

    def _style_plot(self):
        self.ax.set_facecolor(utils.COLOR_BG)
        self.ax.tick_params(colors=utils.COLOR_TEXT_SECONDARY, labelsize=8)
        self.ax.xaxis.label.set_color(utils.COLOR_TEXT_SECONDARY)
        self.ax.yaxis.label.set_color(utils.COLOR_TEXT_SECONDARY)
        self.ax.zaxis.label.set_color(utils.COLOR_TEXT_SECONDARY)
        
        self.ax.set_xlabel("X (mm)")
        self.ax.set_ylabel("Y (mm)")
        self.ax.set_zlabel("Z (mm)")
        self.ax.grid(False)
        self.ax.xaxis.pane.fill = False
        self.ax.yaxis.pane.fill = False
        self.ax.zaxis.pane.fill = False

    def _setup_rotation_gauge(self):
        canvas_size = 200
        # Half-gauge geometry (mirrors the G-Code Sender rotation gauge):
        # center near the top, bottom semicircle is the valid -90..+90 arc,
        # angle 0 points straight down. cy/r budgeted so the arc and the
        # numeric label both fit inside the fixed canvas without clipping.
        self._gauge_cx = canvas_size / 2
        self._gauge_cy = 30
        self._gauge_r = 80
        self.e_canvas = tk.Canvas(self.gauge_container, width=canvas_size, height=canvas_size,
                                  bg=utils.COLOR_BLACK, highlightthickness=1,
                                  highlightbackground=utils.COLOR_BORDER)
        self.e_canvas.pack(pady=5)
        self.e_canvas.bind("<Button-1>", self._on_gauge_click)
        self.e_canvas.bind("<B1-Motion>", self._on_gauge_click)
        self._draw_gauge()

    def _draw_gauge(self, angle=0):
        c = self.e_canvas
        c.delete("all")
        cx, cy, r = self._gauge_cx, self._gauge_cy, self._gauge_r

        # Bottom semicircle outline (-90 to +90)
        c.create_arc(cx-r, cy-r, cx+r, cy+r, start=0, extent=-180,
                     outline=utils.COLOR_BORDER, style=tk.ARC, width=2)

        # Tick marks every 30 degrees across the arc
        for a in range(-90, 91, 30):
            rad = math.radians(a)
            x1, y1 = cx + (r-5)*math.sin(rad), cy + (r-5)*math.cos(rad)
            x2, y2 = cx + (r+5)*math.sin(rad), cy + (r+5)*math.cos(rad)
            c.create_line(x1, y1, x2, y2, fill=utils.COLOR_TEXT_SECONDARY)

        # Needle (single cyan target needle)
        rad_target = math.radians(angle)
        nx, ny = cx + r*math.sin(rad_target), cy + r*math.cos(rad_target)
        c.create_line(cx, cy, nx, ny, fill=utils.COLOR_ACCENT_CYAN, width=3, arrow=tk.LAST)

        c.create_text(cx, cy+r+25, text=f"{angle:.1f}°", fill=utils.COLOR_TEXT_PRIMARY, font=utils.FONT_BODY_BOLD)

    def _on_gauge_click(self, event):
        cx, cy = self._gauge_cx, self._gauge_cy
        dx, dy = event.x - cx, event.y - cy
        # 0 points down, +90 right, -90 left (matches the sender gauge)
        angle = math.degrees(math.atan2(dx, dy))

        # Above the horizon is outside the valid arc: hold the last value
        if angle < -90 or angle > 90:
            return

        self._draw_gauge(angle)
        
        # Update selected row
        idx = self.selected_row_idx.get()
        if 0 <= idx < len(self.actions):
            action = self.actions[idx]
            if action['type'].get() in ["Move", "Macro: Move & Measure"]:
                action['rot'].set(f"{angle:.1f}")
                self._draw_3d_toolpath()
                if action['type'].get() == "Macro: Move & Measure":
                    self._schedule_macro_regen(idx)

    def _add_action_row(self):
        """Appends a new action to the end of the sequence."""
        new_action = {
            'type': tk.StringVar(value="Move"),
            'x': tk.StringVar(value="0"),
            'y': tk.StringVar(value="0"),
            'z': tk.StringVar(value="0"),
            'rot': tk.StringVar(value="0"),
            'speed': tk.StringVar(value="3000"),
            'sub_type': tk.StringVar(value=self.hub_action_types[0]),
            'duration': tk.StringVar(value="1000"),
            # Macro specific
            'interval_mode': tk.StringVar(value=self.interval_modes[0]),
            'interval_value': tk.StringVar(value="5"),
            'expanded': tk.BooleanVar(value=True),
            'sub_actions': []
        }
        
        # Inherit parameters from previous move
        for a in reversed(self.actions):
            if a['type'].get() in ["Move", "Macro: Move & Measure"]:
                for key in ['x', 'y', 'z', 'rot', 'speed']:
                    new_action[key].set(a[key].get())
                break
                
        self.actions.append(new_action)
        self.selected_row_idx.set(len(self.actions) - 1)
        self._refresh_table()
        self._draw_3d_toolpath()
        self._sync_to_editor()

    def _insert_action(self, idx):
        """Inserts a new action after the specified index, inheriting its values."""
        old = self.actions[idx]
        new_action = {}
        for k, v in old.items():
            if isinstance(v, tk.Variable):
                new_action[k] = type(v)(value=v.get())
            elif k == 'sub_actions':
                new_action[k] = [] # Don't clone generated sub-actions directly
            else:
                new_action[k] = v
                
        self.actions.insert(idx + 1, new_action)
        self.selected_row_idx.set(idx + 1)
        self._refresh_table()
        self._draw_3d_toolpath()
        self._sync_to_editor()

    def _move_action(self, idx, direction):
        """Moves an action up (-1) or down (1) in the sequence."""
        new_idx = idx + direction
        if 0 <= new_idx < len(self.actions):
            self.actions[idx], self.actions[new_idx] = self.actions[new_idx], self.actions[idx]
            self.selected_row_idx.set(new_idx)
            self._refresh_table()
            self._draw_3d_toolpath()
            self._sync_to_editor()

    def _duplicate_action(self, idx):
        """Creates a deep copy of the action at the given index."""
        if 0 <= idx < len(self.actions):
            self._insert_action(idx)

    def _remove_action_row(self, idx):
        if 0 <= idx < len(self.actions):
            self.actions.pop(idx)
            if self.selected_row_idx.get() >= len(self.actions):
                self.selected_row_idx.set(len(self.actions) - 1)
            self._refresh_table()
            self._draw_3d_toolpath()
            self._sync_to_editor()

    def _refresh_table(self):
        self.entry_refs.clear()
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
            
        # Configure columns: 0=Reorder, 1=Insert, 2=Type, 3-7=Params, 8=Clone, 9=Delete
        self.scrollable_frame.columnconfigure(0, weight=0, minsize=50) # Controls
        self.scrollable_frame.columnconfigure(1, weight=0, minsize=30) # Insert
        self.scrollable_frame.columnconfigure(2, weight=0, minsize=140) # Type
        for i in range(3, 8): self.scrollable_frame.columnconfigure(i, weight=1)
        self.scrollable_frame.columnconfigure(8, weight=0)
        self.scrollable_frame.columnconfigure(9, weight=0)

        # Header definitions for each action type
        header_map = {
            "Move": ["", "", "ACTION TYPE", "X (mm)", "Y (mm)", "Z (mm)", "ROT (°)", "SPEED (F)", "", ""],
            "Hub Action": ["", "", "ACTION TYPE", "COMMAND", "", "COLOR", "", "", "", ""],
            "Pause": ["", "", "ACTION TYPE", "", "DURATION (ms)", "", "", "", "", ""],
            "Macro: Move & Measure": ["", "", "MACRO TYPE", "X (mm)", "Y (mm)", "Z (mm)", "ROT (°)", "SPEED (F)", "", ""]
        }

        current_grid_row = 0
        last_type = None

        for i, action in enumerate(self.actions):
            atype = action['type'].get()
            
            # If type changes, insert a header row
            if atype != last_type:
                headers = header_map.get(atype, ["", "", "ACTION TYPE", "P1", "P2", "P3", "P4", "P5", "", ""])
                for j, h in enumerate(headers):
                    if h:
                        lbl = ttk.Label(self.scrollable_frame, text=h, font=utils.FONT_BODY_SMALL, 
                                        foreground=utils.COLOR_ACCENT_CYAN if j > 2 else utils.COLOR_TEXT_SECONDARY)
                        lbl.grid(row=current_grid_row, column=j, padx=5, pady=(10, 2), sticky="w" if j==2 else "n")
                current_grid_row += 1
                last_type = atype

            rows_used = self._create_row_widgets(i, action, current_grid_row)
            current_grid_row += rows_used
            
            # Render sub-actions if it's an expanded macro
            if atype == "Macro: Move & Measure" and action.get('expanded', tk.BooleanVar()).get():
                sub_actions = action.get('sub_actions', [])
                if sub_actions:
                    for sub_idx, sub_act in enumerate(sub_actions):
                        self._create_sub_row_widgets(i, sub_idx, sub_act, current_grid_row)
                        current_grid_row += 1

    def _create_row_widgets(self, idx, action, grid_row):
        is_selected = (self.selected_row_idx.get() == idx)
        bg_color = "#1c2128" if is_selected else utils.COLOR_BG
        
        # Background frame for row selection (lowered so it doesn't block inputs)
        bg_frame = tk.Frame(self.scrollable_frame, bg=bg_color)
        bg_frame.grid(row=grid_row, column=0, columnspan=10, sticky="nsew", pady=1)
        bg_frame.lower()
        bg_frame.bind("<Button-1>", lambda e, i=idx: self._select_row(i))
        
        # --- Column 0: Reorder Controls (Low Profile) ---
        ctrl_frame = tk.Frame(self.scrollable_frame, bg=bg_color)
        ctrl_frame.grid(row=grid_row, column=0, sticky="nsew", padx=2)
        ctrl_frame.bind("<Button-1>", lambda e, i=idx: self._select_row(i))
        
        # Selection Bar (Left edge)
        if is_selected:
            tk.Frame(ctrl_frame, bg=utils.COLOR_ACCENT_CYAN, width=3).pack(side=tk.LEFT, fill=tk.Y, padx=(0, 5))

        btn_up = tk.Button(ctrl_frame, text="▲", font=("Inter", 8), bg=bg_color, fg=utils.COLOR_TEXT_SECONDARY,
                           bd=0, activebackground=utils.COLOR_BLACK, activeforeground="white",
                           command=lambda: self._move_action(idx, -1))
        btn_up.pack(side=tk.LEFT)
        
        btn_down = tk.Button(ctrl_frame, text="▼", font=("Inter", 8), bg=bg_color, fg=utils.COLOR_TEXT_SECONDARY,
                             bd=0, activebackground=utils.COLOR_BLACK, activeforeground="white",
                             command=lambda: self._move_action(idx, 1))
        btn_down.pack(side=tk.LEFT)

        # --- Column 1: Insert Mid-Sequence (＋) ---
        btn_ins = tk.Button(self.scrollable_frame, text="＋", font=("Inter", 10, "bold"), bg=bg_color, 
                            fg=utils.COLOR_ACCENT_GREEN, bd=0, activebackground=utils.COLOR_BLACK, 
                            activeforeground="white", command=lambda: self._insert_action(idx))
        btn_ins.grid(row=grid_row, column=1, padx=2)

        # --- Column 2: Action Type ---
        display_types = [t for t in self.action_types if t != "Unrecognized Command"]
        type_cb = ttk.Combobox(self.scrollable_frame, textvariable=action['type'], 
                               values=display_types, width=20, state="readonly")
        type_cb.grid(row=grid_row, column=2, padx=5, pady=4)
        type_cb.bind("<<ComboboxSelected>>", lambda e, i=idx: self._on_type_change(i))

        atype = action['type'].get()
        if atype == "Unrecognized Command":
            type_cb.configure(state="disabled")
        
        # Contextual Columns (3-7)
        if atype == "Move":
            for j, key in enumerate(['x', 'y', 'z', 'rot', 'speed']):
                ent = ttk.Entry(self.scrollable_frame, textvariable=action[key], width=8)
                ent.grid(row=grid_row, column=j+3, padx=2, pady=4)
                ent.bind("<KeyRelease>", lambda e, i=idx: self._on_move_keyrelease(i))
                ent.bind("<FocusIn>", lambda e, i=idx, k=key: self._select_row(i, focus_key=k))
                self.entry_refs[(idx, key)] = ent
        
        elif atype == "Hub Action":
            sub_cb = ttk.Combobox(self.scrollable_frame, textvariable=action['sub_type'], 
                                  values=self.hub_action_types, width=15, state="readonly")
            sub_cb.grid(row=grid_row, column=3, columnspan=2, padx=5, pady=4, sticky="w")
            
            color_canvas = tk.Canvas(self.scrollable_frame, width=20, height=20, bg=bg_color, highlightthickness=0)
            color_canvas.grid(row=grid_row, column=5, padx=5, pady=4)
            hub_color = self.ACTION_COLORS[self.action_types.index("Hub Action")]
            color_canvas.create_oval(4, 4, 16, 16, fill=hub_color, outline="white")
            color_canvas.bind("<Button-1>", lambda e, i=idx: self._select_row(i))
            
        elif atype == "Pause":
            ent = ttk.Entry(self.scrollable_frame, textvariable=action['duration'], width=10)
            ent.grid(row=grid_row, column=4, padx=2, pady=4, sticky="w")
            ent.bind("<FocusIn>", lambda e, i=idx: self._select_row(i, focus_key='duration'))
            self.entry_refs[(idx, 'duration')] = ent
            
        elif atype == "Unrecognized Command":
            ent = ttk.Entry(self.scrollable_frame, textvariable=action.get('raw_text', tk.StringVar(value="")), width=40, state="disabled")
            ent.grid(row=grid_row, column=3, columnspan=5, padx=2, pady=4, sticky="w")
            ent.bind("<Button-1>", lambda e, i=idx: self._select_row(i))
            
        elif atype == "Macro: Move & Measure":
            # --- Line 1: Standard Move Parameters ---
            for j, key in enumerate(['x', 'y', 'z', 'rot', 'speed']):
                ent = ttk.Entry(self.scrollable_frame, textvariable=action[key], width=8)
                ent.grid(row=grid_row, column=j+3, padx=2, pady=4)
                ent.bind("<KeyRelease>", lambda e, i=idx: self._on_move_keyrelease(i))
                ent.bind("<FocusIn>", lambda e, i=idx, k=key: self._select_row(i, focus_key=k))
                self.entry_refs[(idx, key)] = ent

            # --- Line 2: Macro specific controls (Indented) ---
            macro_row = grid_row + 1
            
            # Background frame for line 2
            bg_frame2 = tk.Frame(self.scrollable_frame, bg=bg_color)
            bg_frame2.grid(row=macro_row, column=0, columnspan=10, sticky="nsew", pady=1)
            bg_frame2.lower()
            bg_frame2.bind("<Button-1>", lambda e, i=idx: self._select_row(i))
            
            # Indent indicator
            tk.Frame(self.scrollable_frame, bg=utils.COLOR_BORDER, width=2).grid(row=macro_row, column=2, sticky="nse", pady=2, padx=(0,5))
            
            # Container for all line 2 config to prevent grid distortion
            config_frame = tk.Frame(self.scrollable_frame, bg=bg_color)
            config_frame.grid(row=macro_row, column=3, columnspan=6, sticky="w", pady=2)
            config_frame.bind("<Button-1>", lambda e, i=idx: self._select_row(i))
            
            ttk.Label(config_frame, text="↳ Config:", background=bg_color, foreground=utils.COLOR_TEXT_SECONDARY, font=utils.FONT_BODY_SMALL).pack(side=tk.LEFT, padx=(0, 5))

            interval_cb = ttk.Combobox(config_frame, textvariable=action['interval_mode'], values=self.interval_modes, width=12, state="readonly")
            interval_cb.pack(side=tk.LEFT, padx=2)
            interval_cb.bind("<<ComboboxSelected>>", lambda e, i=idx: self._schedule_macro_regen(i))
            
            ent_val = ttk.Entry(config_frame, textvariable=action['interval_value'], width=5)
            ent_val.pack(side=tk.LEFT, padx=2)
            ent_val.bind("<KeyRelease>", lambda e, i=idx: self._on_move_keyrelease(i))
            ent_val.bind("<FocusIn>", lambda e, i=idx: self._select_row(i, focus_key='interval_value'))
            self.entry_refs[(idx, 'interval_value')] = ent_val
            
            exp_txt = "▼" if action['expanded'].get() else "▶"
            tk.Button(config_frame, text=exp_txt, bg=bg_color, fg=utils.COLOR_TEXT_SECONDARY, bd=0, 
                      command=lambda i=idx: self._toggle_macro_expand(i)).pack(side=tk.LEFT, padx=(10, 2))
                      
            ttk.Button(config_frame, text="⟳ Gen", width=6, command=lambda i=idx: self._generate_macro_steps(i)).pack(side=tk.LEFT, padx=2)


        # --- Column 8: Clone (❐) ---
        clone_btn = tk.Button(self.scrollable_frame, text="❐", fg=utils.COLOR_ACCENT_CYAN, 
                              bg=bg_color, bd=0, activebackground=utils.COLOR_BLACK, 
                              activeforeground="white", font=("Inter", 11),
                              command=lambda i=idx: self._duplicate_action(i))
        clone_btn.grid(row=grid_row, column=8, padx=2)

        # --- Column 9: Delete (✕) ---
        del_btn = tk.Button(self.scrollable_frame, text="✕", fg=utils.COLOR_ACCENT_RED, 
                            bg=bg_color, bd=0, activebackground=utils.COLOR_BLACK, 
                            activeforeground="white", font=("Inter", 10, "bold"),
                            command=lambda i=idx: self._remove_action_row(i))
        del_btn.grid(row=grid_row, column=9, padx=5)

        return 2 if atype == "Macro: Move & Measure" else 1

    def _create_sub_row_widgets(self, parent_idx, sub_idx, sub_action, grid_row):
        """Renders an indented row for a macro's sub-action."""
        # Visual indent and indicator
        tk.Frame(self.scrollable_frame, bg=utils.COLOR_BORDER, width=2).grid(row=grid_row, column=1, sticky="nse", pady=2, padx=5)
        
        # Action Type (Readonly label-like combobox)
        type_cb = ttk.Combobox(self.scrollable_frame, textvariable=sub_action['type'], 
                               values=self.action_types, width=12, state="disabled")
        type_cb.grid(row=grid_row, column=2, padx=(20, 5), pady=2, sticky="w")

        atype = sub_action['type'].get()
        
        if atype == "Move":
            f = ttk.Frame(self.scrollable_frame, style='Dark.TFrame')
            f.grid(row=grid_row, column=3, columnspan=4, sticky="w", padx=2)
            for key in ['x', 'y', 'z', 'rot', 'speed']:
                ent = ttk.Entry(f, textvariable=sub_action[key], width=6)
                ent.pack(side=tk.LEFT, padx=1)
                ent.bind("<KeyRelease>", lambda e: self._draw_3d_toolpath())
                
        # Add a local delete button for the sub-action
        del_btn = tk.Button(self.scrollable_frame, text="✕", fg=utils.COLOR_TEXT_SECONDARY, 
                            bg=utils.COLOR_BG, bd=0, activebackground=utils.COLOR_BLACK, 
                            activeforeground=utils.COLOR_ACCENT_RED, font=("Inter", 8),
                            command=lambda pi=parent_idx, si=sub_idx: self._remove_sub_action(pi, si))
        del_btn.grid(row=grid_row, column=9, padx=5)

    def _toggle_macro_expand(self, idx):
        self.actions[idx]['expanded'].set(not self.actions[idx]['expanded'].get())
        self._refresh_table()

    def _remove_sub_action(self, parent_idx, sub_idx):
        self.actions[parent_idx]['sub_actions'].pop(sub_idx)
        self._refresh_table()
        self._draw_3d_toolpath()

    def _on_move_keyrelease(self, idx):
        self._draw_3d_toolpath()
        if 0 <= idx < len(self.actions) and self.actions[idx]['type'].get() == "Macro: Move & Measure":
            self._schedule_macro_regen(idx)

    def _schedule_macro_regen(self, idx):
        if hasattr(self, '_macro_regen_timer') and self._macro_regen_timer:
            self.after_cancel(self._macro_regen_timer)
        self._macro_regen_timer = self.after(600, lambda: self._generate_macro_steps(idx))

    def _generate_macro_steps(self, idx):
        action = self.actions[idx]
        
        # Save focus
        focused = self.focus_get()
        focus_key = None
        cursor_pos = None
        for (i, k), w in self.entry_refs.items():
            if w == focused and i == idx:
                focus_key = k
                if isinstance(w, (tk.Entry, ttk.Entry)):
                    cursor_pos = w.index(tk.INSERT)
                break
        
        # 1. Find starting position
        start_x, start_y, start_z, start_rot = 0.0, 0.0, 0.0, 0.0
        for i in range(idx - 1, -1, -1):
            if self.actions[i]['type'].get() in ["Move", "Macro: Move & Measure"]:
                # If the previous is a macro, we should technically look at its last sub-action,
                # but to keep it simple, we'll look at the macro's target.
                try: start_x = float(self.actions[i]['x'].get())
                except: pass
                try: start_y = float(self.actions[i]['y'].get())
                except: pass
                try: start_z = float(self.actions[i]['z'].get())
                except: pass
                try: start_rot = float(self.actions[i]['rot'].get())
                except: pass
                break
                
        # 2. Get target position
        try: target_x = float(action['x'].get())
        except: target_x = start_x
        try: target_y = float(action['y'].get())
        except: target_y = start_y
        try: target_z = float(action['z'].get())
        except: target_z = start_z
        try: target_rot = float(action['rot'].get())
        except: target_rot = start_rot
        
        speed = action['speed'].get()

        # 3. Determine number of steps
        mode = action['interval_mode'].get()
        try: val = float(action['interval_value'].get())
        except: val = 1.0
        
        num_steps = 1
        if mode == "Num Steps":
            num_steps = max(1, int(val))
        elif mode == "Step Size (mm)":
            dist = math.sqrt((target_x - start_x)**2 + (target_y - start_y)**2 + (target_z - start_z)**2)
            if val > 0:
                num_steps = max(1, math.ceil(dist / val))
                
        # 4. Generate interpolation
        sub_actions = []
        for i in range(1, num_steps + 1):
            t = i / num_steps
            cx = start_x + (target_x - start_x) * t
            cy = start_y + (target_y - start_y) * t
            cz = start_z + (target_z - start_z) * t
            cr = start_rot + (target_rot - start_rot) * t
            
            sub_actions.append({
                'type': tk.StringVar(value="Move"),
                'x': tk.StringVar(value=f"{cx:.2f}"),
                'y': tk.StringVar(value=f"{cy:.2f}"),
                'z': tk.StringVar(value=f"{cz:.2f}"),
                'rot': tk.StringVar(value=f"{cr:.2f}"),
                'speed': tk.StringVar(value=speed)
            })
            
        action['sub_actions'] = sub_actions
        action['expanded'].set(True)
        self._refresh_table()
        self._draw_3d_toolpath()

    def _get_flattened_actions(self):
        """Yields a flat list of actions, injecting sub-actions where appropriate."""
        flat = []
        for i, action in enumerate(self.actions):
            if action['type'].get() == "Macro: Move & Measure":
                for sub in action.get('sub_actions', []):
                    sub['_parent_idx'] = i
                    flat.append(sub)
            else:
                action['_parent_idx'] = i
                flat.append(action)
        return flat

    def _on_type_change(self, idx):
        # When type changes, we might want to default some values
        atype = self.actions[idx]['type'].get()
        if atype in ["Move", "Macro: Move & Measure"]:
            # Inherit previous position if possible
            prev_x, prev_y, prev_z, prev_rot = "0", "0", "0", "0"
            for i in range(idx-1, -1, -1):
                if self.actions[i]['type'].get() in ["Move", "Macro: Move & Measure"]:
                    prev_x = self.actions[i]['x'].get()
                    prev_y = self.actions[i]['y'].get()
                    prev_z = self.actions[i]['z'].get()
                    prev_rot = self.actions[i]['rot'].get()
                    break
            self.actions[idx]['x'].set(prev_x)
            self.actions[idx]['y'].set(prev_y)
            self.actions[idx]['z'].set(prev_z)
            self.actions[idx]['rot'].set(prev_rot)

        self._refresh_table()
        self._draw_3d_toolpath()

    def _select_row(self, idx, focus_key=None):
        if self.selected_row_idx.get() == idx:
            return
            
        self.selected_row_idx.set(idx)
        self._refresh_table()
        
        # Update gauge if it's a move row
        if 0 <= idx < len(self.actions):
            action = self.actions[idx]
            if action['type'].get() in ["Move", "Macro: Move & Measure"]:
                try:
                    ang = float(action['rot'].get())
                    self._draw_gauge(ang)
                except: pass

        # Restore focus to the specific entry if it was clicked
        if focus_key:
            ent = self.entry_refs.get((idx, focus_key))
            if ent:
                ent.focus_set()

    def _draw_3d_toolpath(self):
        if not MATPLOTLIB_AVAILABLE: return
        self.ax.clear()
        self._style_plot()
        
        flat_actions = self._get_flattened_actions()
        if not flat_actions:
            self.ax.set_xlim(-110, 110)
            self.ax.set_ylim(-110, 110)
            self.ax.set_zlim(0, 128)
            self.canvas_plt.draw()
            return

        # Calculate chronological progress
        num_actions = len(flat_actions)
        limit_idx = int(self.scrub_var.get() * (num_actions - 1)) if num_actions > 0 else 0
        
        # Calculate global bounds first so scale doesn't jitter during scrubbing
        max_xy = 10.0
        max_z = 10.0
        for action in flat_actions:
            if action['type'].get() == "Move":
                try:
                    x, y, z = float(action['x'].get()), float(action['y'].get()), float(action['z'].get())
                    max_xy = max(max_xy, abs(x), abs(y))
                    max_z = max(max_z, z)
                except: pass

        pts = []
        highlight_segments = [] # To store (p1, p2) for green highlighting
        curr_x, curr_y, curr_z = 0, 0, 0
        selected_parent = self.selected_row_idx.get()
        
        for i, action in enumerate(flat_actions):
            if i > limit_idx: break
            
            atype = action['type'].get()
            if atype == "Move":
                try:
                    prev_p = (curr_x, curr_y, curr_z)
                    curr_x, curr_y, curr_z = float(action['x'].get()), float(action['y'].get()), float(action['z'].get())
                    pts.append((curr_x, curr_y, curr_z))
                    
                    if action.get('_parent_idx') == selected_parent:
                        highlight_segments.append((prev_p, (curr_x, curr_y, curr_z)))
                except: pass
            else:
                # Plot a marker at the current toolhead position
                color = self.ACTION_COLORS[self.action_types.index(atype)]
                self.ax.scatter([curr_x], [curr_y], [curr_z], color=color, s=50, edgecolors='white', depthshade=False)

        if len(pts) > 1:
            xs, ys, zs = zip(*pts)
            self.ax.plot(xs, ys, zs, color=utils.COLOR_ACCENT_CYAN, linewidth=2, alpha=0.8)

        # Draw highlighted segments in thicker green
        for p1, p2 in highlight_segments:
            self.ax.plot([p1[0], p2[0]], [p1[1], p2[1]], [p1[2], p2[2]], 
                         color=utils.COLOR_ACCENT_GREEN, linewidth=4, alpha=1.0)
            
        # Draw current toolhead position
        self.ax.scatter([curr_x], [curr_y], [curr_z], color="white", s=80, marker='+', depthshade=False)
        
        # Apply autoscaled limits (padded by 10%)
        padding_xy = max_xy * 0.1
        self.ax.set_xlim(-(max_xy + padding_xy), max_xy + padding_xy)
        self.ax.set_ylim(-(max_xy + padding_xy), max_xy + padding_xy)
        self.ax.set_zlim(0, max_z + (max_z * 0.1))
            
        self.canvas_plt.draw()
        self._draw_bookmarks()

    def _draw_bookmarks(self):
        """Draws small ticks on the bookmark canvas for each non-movement action."""
        self.bookmark_canvas.delete("all")
        w = self.bookmark_canvas.winfo_width()
        if w <= 1: return
        
        flat_actions = self._get_flattened_actions()
        num_actions = len(flat_actions)
        if num_actions <= 1: return
        
        for i, action in enumerate(flat_actions):
            atype = action['type'].get()
            if atype != "Move":
                rel_pos = i / (num_actions - 1)
                x = rel_pos * w
                color = self.ACTION_COLORS[self.action_types.index(atype)]
                self.bookmark_canvas.create_line(x, 0, x, 12, fill=color, width=2)

    def _load_gcode(self):
        """Loads a G-code file and attempts to recover actions and movements."""
        f = filedialog.askopenfilename(filetypes=[("G-code files", "*.gcode"), ("All files", "*.*")])
        if not f: return
        
        try:
            with open(f, 'r') as file:
                text = file.read()
                
            new_actions = self._parse_gcode_text(text)
            
            if new_actions:
                self.actions = new_actions
                self.selected_row_idx.set(-1)
                self.scrub_var.set(1.0)
                self._refresh_table()
                self._draw_3d_toolpath()
                
                # Update editor
                self.gcode_editor.delete("1.0", tk.END)
                self.gcode_editor.insert("1.0", text)
                
                messagebox.showinfo("Success", f"Recovered {len(new_actions)} actions from G-code.")
            else:
                messagebox.showwarning("Warning", "No recognizable actions found in file.")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load G-code: {e}")
            
    def _parse_gcode_text(self, text):
        new_actions = []
        curr_x, curr_y, curr_z, curr_rot = "0", "0", "0", "0"
        curr_speed = "3000"
        
        lines = text.split('\n')
        for line in lines:
            line = line.strip()
            if not line: continue
            
            # Skip pure comments that are just metadata
            if line == "; SEED Manual Sequence" or line.startswith("; Generated:") or line == "; End of Sequence":
                continue
                
            matched = False
            
            # 1. Recover Movements (G1)
            if line.startswith("G1"):
                def get_coord(l, char, default):
                    match = re.search(fr"{char}([-+]?\d*\.?\d+)", l)
                    return match.group(1) if match else default
                
                curr_x = get_coord(line, "X", curr_x)
                curr_y = get_coord(line, "Y", curr_y)
                curr_z = get_coord(line, "Z", curr_z)
                curr_rot = get_coord(line, "E", curr_rot)
                curr_speed = get_coord(line, "F", curr_speed)
                
                new_actions.append({
                    'type': tk.StringVar(value="Move"),
                    'x': tk.StringVar(value=curr_x),
                    'y': tk.StringVar(value=curr_y),
                    'z': tk.StringVar(value=curr_z),
                    'rot': tk.StringVar(value=curr_rot),
                    'speed': tk.StringVar(value=curr_speed),
                    'sub_type': tk.StringVar(value=self.hub_action_types[0]),
                    'duration': tk.StringVar(value="1000"),
                })
                matched = True

            # 2. Recover Hub Actions
            elif "; HUB_CMD" in line:
                param_match = re.search(r"param=(.+?) value=", line)
                value_match = re.search(r"value=(\S+)", line)
                if param_match and value_match:
                    p, v = param_match.group(1), value_match.group(1).lower()
                    if p == "WPT Start": ui_type = "WPT Start"
                    elif p == "WPT Stop": ui_type = "WPT Stop"
                    elif p == "DCDC Enable" and v == "true": ui_type = "DCDC Enable"
                    elif p == "DCDC Enable" and v == "false": ui_type = "DCDC Disable"
                    else: ui_type = "WPT Start"

                    new_actions.append({
                        'type': tk.StringVar(value="Hub Action"),
                        'x': tk.StringVar(value=curr_x),
                        'y': tk.StringVar(value=curr_y),
                        'z': tk.StringVar(value=curr_z),
                        'rot': tk.StringVar(value=curr_rot),
                        'speed': tk.StringVar(value=curr_speed),
                        'sub_type': tk.StringVar(value=ui_type),
                        'duration': tk.StringVar(value="1000"),
                    })
                    matched = True

            # 3. Recover Pauses
            elif "; WAIT" in line:
                secs_match = re.search(r"seconds=([-+]?\d*\.?\d+)", line)
                if secs_match:
                    ms = str(int(float(secs_match.group(1)) * 1000))
                    new_actions.append({
                        'type': tk.StringVar(value="Pause"),
                        'x': tk.StringVar(value=curr_x),
                        'y': tk.StringVar(value=curr_y),
                        'z': tk.StringVar(value=curr_z),
                        'rot': tk.StringVar(value=curr_rot),
                        'speed': tk.StringVar(value=curr_speed),
                        'sub_type': tk.StringVar(value=self.hub_action_types[0]),
                        'duration': tk.StringVar(value=ms),
                    })
                    matched = True

            # 4. Unrecognized Command
            if not matched:
                new_actions.append({
                    'type': tk.StringVar(value="Unrecognized Command"),
                    'x': tk.StringVar(value=curr_x),
                    'y': tk.StringVar(value=curr_y),
                    'z': tk.StringVar(value=curr_z),
                    'rot': tk.StringVar(value=curr_rot),
                    'speed': tk.StringVar(value=curr_speed),
                    'raw_text': tk.StringVar(value=line),
                    # Provide defaults for others to avoid key errors
                    'sub_type': tk.StringVar(value=self.hub_action_types[0]),
                    'duration': tk.StringVar(value="1000"),
                    'interval_mode': tk.StringVar(value=self.interval_modes[0]),
                    'interval_value': tk.StringVar(value="5"),
                    'expanded': tk.BooleanVar(value=True),
                    'sub_actions': []
                })
                
        return new_actions

    def _on_editor_key_release(self, event):
        if self._editor_timer:
            self.after_cancel(self._editor_timer)
        self._editor_timer = self.after(1000, self._sync_from_editor)

    def _sync_from_editor(self):
        text = self.gcode_editor.get("1.0", tk.END)
        new_actions = self._parse_gcode_text(text)
        
        # Only update if we successfully parsed things or if they literally cleared the whole editor
        if new_actions or not text.strip():
            self.actions = new_actions
            # Try to preserve selected index if possible, else reset
            if self.selected_row_idx.get() >= len(self.actions):
                self.selected_row_idx.set(len(self.actions) - 1)
            self._refresh_table()
            self._draw_3d_toolpath()

    def _send_to_sender(self):
        """Generates a temporary G-Code file and sends it to the Sender panel."""
        if not self.actions:
            messagebox.showwarning("Empty Sequence", "Please add at least one action before sending.")
            return

        import tempfile
        tf = os.path.join(tempfile.gettempdir(), f"manual_sequence_{datetime.now().strftime('%H%M%S')}.gcode")
        
        try:
            self._write_gcode_file(tf)
            if self.c and hasattr(self.c, '_handle_send_to_sender'):
                self.c._handle_send_to_sender(tf)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to send sequence: {e}")

    def _export_gcode(self):
        """Prompts user to save the manual sequence as a G-Code file."""
        if not self.actions:
            messagebox.showwarning("Empty Sequence", "Nothing to export.")
            return

        f = filedialog.asksaveasfilename(defaultextension=".gcode", 
                                         filetypes=[("G-code files", "*.gcode"), ("All files", "*.*")])
        if f:
            try:
                self._write_gcode_file(f)
                messagebox.showinfo("Success", f"Sequence exported to {os.path.basename(f)}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export: {e}")

    def _write_gcode_file(self, filepath):
        """Core logic to translate the internal action list into G-Code format."""
        with open(filepath, 'w') as out:
            self._write_gcode_stream(out)

    def _write_gcode_stream(self, out):
        out.write("; SEED Manual Sequence\n")
        out.write(f"; Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
        
        # Map UI types to HUB_CMD format — param names must match rlm_proj_vivigo.py exactly
        hub_cmd_map = {
            "WPT Start":    "param=WPT Start value=true",
            "WPT Stop":     "param=WPT Stop value=true",
            "DCDC Enable":  "param=DCDC Enable value=true",
            "DCDC Disable": "param=DCDC Enable value=false",
        }

        flat_actions = self._get_flattened_actions()
        for action in flat_actions:
            atype = action['type'].get()
            
            if atype == "Move":
                x, y, z = action.get('x', tk.StringVar(value="0")).get(), action.get('y', tk.StringVar(value="0")).get(), action.get('z', tk.StringVar(value="0")).get()
                rot, speed = action.get('rot', tk.StringVar(value="0")).get(), action.get('speed', tk.StringVar(value="3000")).get()
                out.write(f"G1 X{x} Y{y} Z{z} E{rot} F{speed}\n")
            
            elif atype == "Hub Action":
                sub = action.get('sub_type', tk.StringVar(value="")).get()
                if sub == "WAIT":
                    out.write("; WAIT seconds=1.0\n")
                else:
                    out.write(f"; HUB_CMD {hub_cmd_map.get(sub, '')}\n")
            
            elif atype == "Pause":
                try: secs = float(action.get('duration', tk.StringVar(value="1000")).get()) / 1000.0
                except: secs = 1.0
                out.write(f"; WAIT seconds={secs}\n")
            
            elif atype == "Unrecognized Command":
                raw = action.get('raw_text', tk.StringVar(value="")).get()
                out.write(f"{raw}\n")

        out.write("\n; End of Sequence\n")

    def _sync_to_editor(self):
        """Generates G-code from current table actions and populates the editor."""
        import io
        out = io.StringIO()
        self._write_gcode_stream(out)
        self.gcode_editor.delete("1.0", tk.END)
        self.gcode_editor.insert("1.0", out.getvalue())
