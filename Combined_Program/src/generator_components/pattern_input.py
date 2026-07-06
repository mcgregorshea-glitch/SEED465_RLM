import tkinter as tk
from tkinter import ttk, messagebox
import os
from datetime import datetime

class PatternInput(ttk.Frame):
    """
    Component for handling all user inputs for scan pattern generation.
    Encapsulates X, Y, Z, Rot settings and symmetry logic.
    """
    def __init__(self, parent, controller):
        super().__init__(parent, style='Dark.TFrame')
        self.c = controller
        self._setup_ui()

    def _setup_ui(self):
        # 1. Header section (TOP)
        title_row = ttk.Frame(self, style='Dark.TFrame')
        title_row.pack(side=tk.TOP, fill=tk.X, pady=(0, 10))

        title = ttk.Label(title_row, text="PATTERN GENERATOR",
                            font=('Rajdhani', 16, 'bold'),
                            foreground=self.c.COLOR_ACCENT_CYAN, background=self.c.COLOR_BG)
        title.pack(side=tk.LEFT, anchor=tk.W)

        self.info_button = tk.Button(title_row, text="?", font=("Segoe UI", 12, "bold"),
            bg=self.c.COLOR_BG, fg=self.c.COLOR_ACCENT_CYAN, activebackground=self.c.COLOR_BG,
            activeforeground="#ffffff", bd=0, highlightthickness=0, takefocus=0, cursor="hand2", padx=10,
            command=self.c._show_tutorial_popup)
        self.info_button.pack(side=tk.RIGHT, padx=(5, 0), pady=8)

        # 2. FIXED FOOTER (BOTTOM) - SYNCED ORDER
        footer_frame = ttk.Frame(self, style='Dark.TFrame')
        footer_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 0))

        # Test Profile Section (ON TOP of buttons)
        self.c._create_naming_ui(footer_frame)

        # Action Buttons (ON BOTTOM of footer)
        self.c._create_action_buttons(footer_frame)

        # 3. CENTER CONTENT (Scrollable)
        scroll_container = ttk.Frame(self, style='Dark.TFrame')
        scroll_container.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(scroll_container, highlightthickness=0, bg=self.c.COLOR_BG)
        scrollbar = ttk.Scrollbar(scroll_container, orient="vertical", command=canvas.yview, style='Vertical.TScrollbar')
        scrollable_frame = ttk.Frame(canvas, style='Dark.TFrame')

        def _on_mousewheel(event):
            scroll_val = 0
            if event.num == 4: scroll_val = -1
            elif event.num == 5: scroll_val = 1
            elif event.delta: scroll_val = int(-1 * (event.delta / 120)) if abs(event.delta) >= 120 else -1 * event.delta
            canvas.yview_scroll(scroll_val, "units")

        def _bind_mousewheel_recursive(widget):
            widget.bind("<MouseWheel>", _on_mousewheel)
            widget.bind("<Button-4>", _on_mousewheel)
            widget.bind("<Button-5>", _on_mousewheel)
            for child in widget.winfo_children(): _bind_mousewheel_recursive(child)

        self.c.scrollable_window_id = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(self.c.scrollable_window_id, width=e.width))
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.configure(yscrollcommand=scrollbar.set)

        # --- Axis Helper ---
        def create_axis_group(label, attr_prefix, default_min, default_max, symmetric_var, toggle_cmd):
            frame = ttk.LabelFrame(scrollable_frame, text=label, style='Card.TLabelframe', padding=12)
            frame.pack(fill=tk.X, pady=(0, 8), padx=2)
            frame.columnconfigure(1, weight=1); frame.columnconfigure(3, weight=1)
            
            min_lbl = ttk.Label(frame, text="Min:", foreground=self.c.COLOR_TEXT_SECONDARY)
            min_lbl.grid(row=0, column=0, sticky=tk.E, padx=(0,5))
            min_ent = ttk.Entry(frame, width=8); min_ent.insert(0, default_min)
            min_ent.grid(row=0, column=1, padx=3, sticky=tk.EW)
            setattr(self.c, f"{attr_prefix}_min_label", min_lbl); setattr(self.c, f"{attr_prefix}_min", min_ent)

            max_lbl = ttk.Label(frame, text="Max:", foreground=self.c.COLOR_TEXT_SECONDARY)
            max_lbl.grid(row=0, column=2, sticky=tk.E, padx=(8, 5))
            max_ent = ttk.Entry(frame, width=8); max_ent.insert(0, default_max)
            max_ent.grid(row=0, column=3, padx=3, sticky=tk.EW)
            setattr(self.c, f"{attr_prefix}_max_label", max_lbl); setattr(self.c, f"{attr_prefix}_max", max_ent)

            off_lbl = ttk.Label(frame, text="±Offset:", foreground=self.c.COLOR_TEXT_SECONDARY)
            off_lbl.grid(row=0, column=0, sticky=tk.E, padx=(0,5))
            off_ent = ttk.Entry(frame, width=10)
            off_ent.grid(row=0, column=1, columnspan=3, sticky="ew", padx=3)
            off_lbl.grid_remove(); off_ent.grid_remove()
            setattr(self.c, f"{attr_prefix}_offset_label", off_lbl); setattr(self.c, f"{attr_prefix}_offset", off_ent)

            ttk.Label(frame, text="Step:", foreground=self.c.COLOR_TEXT_SECONDARY).grid(row=1, column=0, sticky=tk.E, pady=(8,0), padx=(0,5))
            step_ent = ttk.Entry(frame, width=8); step_ent.insert(0, "5")
            step_ent.grid(row=1, column=1, padx=3, pady=(8,0), sticky=tk.EW)
            setattr(self.c, f"{attr_prefix}_step", step_ent)

            ttk.Checkbutton(frame, text="Symmetric", variable=symmetric_var, command=toggle_cmd).grid(row=1, column=2, columnspan=2, sticky=tk.W, padx=8, pady=(8,0))

            for ent in [min_ent, max_ent, off_ent, step_ent]:
                ent.bind('<FocusOut>', self.c._auto_update_preview); ent.bind('<Return>', self.c._auto_update_preview)

        create_axis_group("X Axis (mm)", "x", "-50", "50", self.c.x_symmetric, self.c._on_x_symmetric_toggle)
        create_axis_group("Y Axis (mm)", "y", "-50", "50", self.c.y_symmetric, self.c._on_y_symmetric_toggle)
        create_axis_group("Z Axis (mm)", "z", "0", "100", self.c.z_symmetric, self.c._on_z_symmetric_toggle)
        create_axis_group("Rotation (deg)", "rot", "0", "0", self.c.rot_symmetric, self.c._on_rot_symmetric_toggle)

        # --- Movement Settings ---
        mov_frame = ttk.LabelFrame(scrollable_frame, text="Movement Settings", style='Card.TLabelframe', padding=12)
        mov_frame.pack(fill=tk.X, pady=(0, 8), padx=2)
        
        ttk.Label(mov_frame, text="Travel Speed (mm/min):", foreground=self.c.COLOR_TEXT_SECONDARY).grid(row=0, column=0, sticky=tk.E, padx=(0,5))
        self.c.travelspeed = ttk.Entry(mov_frame, width=10); self.c.travelspeed.insert(0, "3000")
        self.c.travelspeed.grid(row=0, column=1, padx=3, sticky=tk.W)

        ttk.Label(mov_frame, text="Pause (seconds):", foreground=self.c.COLOR_TEXT_SECONDARY).grid(row=1, column=0, sticky=tk.E, pady=(8,0), padx=(0,5))
        self.c.pause_time = ttk.Entry(mov_frame, width=10); self.c.pause_time.insert(0, "1")
        self.c.pause_time.grid(row=1, column=1, padx=3, pady=(8,0), sticky=tk.W)

        self.c.travelspeed.bind('<FocusOut>', self.c._auto_update_preview); self.c.travelspeed.bind('<Return>', self.c._auto_update_preview)
        self.c.pause_time.bind('<FocusOut>', self.c._auto_update_preview); self.c.pause_time.bind('<Return>', self.c._auto_update_preview)
        
        # --- Export Format ---
        exp_frame = ttk.LabelFrame(scrollable_frame, text="Export Format", style='Card.TLabelframe', padding=12)
        exp_frame.pack(fill=tk.X, pady=(0, 8), padx=2)
        ttk.Radiobutton(exp_frame, text="G-code (.gcode)", variable=self.c.export_format, value="gcode", command=self.c._auto_update_preview).pack(anchor=tk.W)
        ttk.Radiobutton(exp_frame, text="CSV Coordinates (.csv)", variable=self.c.export_format, value="csv", command=self.c._auto_update_preview).pack(anchor=tk.W)

        _bind_mousewheel_recursive(scrollable_frame)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.bind("<MouseWheel>", _on_mousewheel); canvas.bind("<Button-4>", _on_mousewheel); canvas.bind("<Button-5>", _on_mousewheel)
