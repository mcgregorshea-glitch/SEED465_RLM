import tkinter as tk
from tkinter import ttk, messagebox
from utils import PRINTER_BOUNDS
import math

class CommandInjection(ttk.Frame):
    """
    Component for coordinate-based Hub command insertion.
    Encapsulates the Action Table and Matplotlib 3D preview with sequence scrubbing.
    """
    def __init__(self, parent, controller):
        super().__init__(parent, style='Dark.TFrame')
        self.c = controller
        self.is_animating = False
        self._animate_job = None
        self._setup_ui()

    def _setup_ui(self):
        # Outer PanedWindow for adjustable Table vs Plot layout
        self.injection_paned = tk.PanedWindow(self, orient=tk.HORIZONTAL, bg=self.c.COLOR_BG, sashwidth=4, sashpad=0, borderwidth=0)
        self.injection_paned.pack(fill=tk.BOTH, expand=True)

        # LEFT SIDE: Action Table Frame
        # Increase min width to 450 to accommodate all columns
        self.action_panel = ttk.Frame(self.injection_paned, style='Dark.TFrame', width=350)
        self.injection_paned.add(self.action_panel, minsize=350)

        # Table Header/Title
        table_title = ttk.Label(self.action_panel, text="ACTION TRIGGER TABLE",
                                font=('Rajdhani', 14, 'bold'),
                                foreground=self.c.COLOR_ACCENT_CYAN, background=self.c.COLOR_BG)
        table_title.pack(side=tk.TOP, fill=tk.X, pady=(10, 5), padx=10)

        # --- FIXED FOOTER (BOTTOM) ---
        footer_frame = ttk.Frame(self.action_panel, style='Dark.TFrame')
        footer_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 0), padx=10)

        # 1. Naming Section (TOP of footer)
        self.c._create_naming_ui(footer_frame)

        # 2. Action Buttons (BOTTOM of footer)
        self.c._create_action_buttons(footer_frame)

        # Table Controls (Add Button)
        self.table_controls = ttk.Frame(self.action_panel, style='Dark.TFrame')
        self.table_controls.pack(side=tk.BOTTOM, fill=tk.X, pady=10, padx=10)

        self.add_action_btn = ttk.Button(self.table_controls, text="+ ADD HUB ACTION (Non-Repeated)", style='Amber.TButton', command=self._add_action_row)
        self.add_action_btn.pack(fill=tk.X)

        self.add_repeated_btn = ttk.Button(self.table_controls, text="+ ADD HUB ACTION (Repeated - Every Point)", style='Purple.TButton', command=self.c._add_repeated_action_row)
        self.add_repeated_btn.pack(fill=tk.X, pady=(4, 0))

        # Scrollable table area (CENTER)
        self.table_canvas = tk.Canvas(self.action_panel, highlightthickness=0, bg=self.c.COLOR_BG)
        self.table_scrollbar = ttk.Scrollbar(self.action_panel, orient="vertical", command=self.table_canvas.yview, style='Vertical.TScrollbar')
        self.table_scroll_inner = ttk.Frame(self.table_canvas, style='Dark.TFrame')

        self.table_canvas_window = self.table_canvas.create_window((0, 0), window=self.table_scroll_inner, anchor="nw")
        
        self.table_scroll_inner.bind("<Configure>", lambda e: self.table_canvas.configure(scrollregion=self.table_canvas.bbox("all")))
        self.table_canvas.bind("<Configure>", lambda e: self.table_canvas.itemconfig(self.table_canvas_window, width=e.width))
        self.table_canvas.configure(yscrollcommand=self.table_scrollbar.set)

        self.table_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.table_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(10, 0))

        # Define column widths (compacted)
        # Old: [4, 12, 12, 12, 12, 18, 5]
        self.col_widths = [2, 8, 8, 8, 8, 14, 2]

        # Initialization
        self._refresh_action_table()

        # RIGHT SIDE: 3D Visualization
        self.plot_container = ttk.Frame(self.injection_paned, style='Dark.TFrame')
        self.injection_paned.add(self.plot_container)

        # --- SCRUBBER WITH INTERACTIVE BOOKMARKS ---
        self.scrubber_outer = ttk.Frame(self.plot_container, style='Dark.TFrame')
        self.scrubber_outer.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=5)
        
        ttk.Label(self.scrubber_outer, text="PATH PROGRESS", font=self.c.FONT_BODY_SMALL, 
                  foreground=self.c.COLOR_TEXT_SECONDARY).pack(side=tk.LEFT, padx=(0, 5))

        self.play_btn = tk.Button(self.scrubber_outer, text="▶", font=("Inter", 12),
                                  bg=self.c.COLOR_BG, fg=self.c.COLOR_ACCENT_GREEN,
                                  activebackground=self.c.COLOR_BLACK, activeforeground="#ffffff",
                                  bd=0, command=self._toggle_animation)
        self.play_btn.pack(side=tk.LEFT, padx=5)
        
        self.scrub_stack = ttk.Frame(self.scrubber_outer, style='Dark.TFrame')
        self.scrub_stack.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # 1. Bookmark Canvas (Clickable)
        self.bookmark_canvas = tk.Canvas(self.scrub_stack, height=12, bg=self.c.COLOR_BG, highlightthickness=0, cursor="hand2")
        self.bookmark_canvas.pack(fill=tk.X, pady=(0, 2))
        self.bookmark_canvas.bind("<Button-1>", self._on_bookmark_click)

        # 2. Scrubber Scale
        self.scrub_var = tk.DoubleVar(value=1.0)
        self.scrubber = ttk.Scale(self.scrub_stack, from_=0.0, to=1.0, variable=self.scrub_var, 
                                  command=lambda v: self.c._draw_3d_toolpath())
        self.scrubber.pack(fill=tk.X)

        self.c._create_3d_plot_widgets(self.plot_container)

    def _toggle_animation(self):
        """Starts or stops the scrubber animation."""
        if self.is_animating:
            self.is_animating = False
            self.play_btn.config(text="▶", fg=self.c.COLOR_ACCENT_GREEN)
            if self._animate_job:
                self.after_cancel(self._animate_job)
                self._animate_job = None
        else:
            # If at the end, reset to start
            if self.scrub_var.get() >= 0.99:
                self.scrub_var.set(0.0)
            
            self.is_animating = True
            self.play_btn.config(text="⏸", fg=self.c.COLOR_ACCENT_AMBER)
            self._animate_step()

    def _animate_step(self):
        """Increments the scrubber and redraws the path."""
        if not self.is_animating:
            return
            
        curr = self.scrub_var.get()
        if curr < 1.0:
            # Increment by 0.5% each step for smoothness
            new_val = min(1.0, curr + 0.005)
            self.scrub_var.set(new_val)
            self.c._draw_3d_toolpath()
            # 30ms delay provides ~33fps
            self._animate_job = self.after(30, self._animate_step)
        else:
            self.is_animating = False
            self.play_btn.config(text="▶", fg=self.c.COLOR_ACCENT_GREEN)
            self._animate_job = None

    def _on_bookmark_click(self, event):
        """Moves the scrubber to the clicked position on the canvas."""
        w = self.bookmark_canvas.winfo_width()
        if w > 0:
            rel_pos = max(0.0, min(1.0, event.x / w))
            self.scrub_var.set(rel_pos)
            self.c._draw_3d_toolpath()

    def _handle_coordinate_change(self):
        """Sorts the table and re-draws the preview when a coordinate is modified."""
        self.c._sort_hub_actions()
        self._refresh_action_table()
        self.c._draw_3d_toolpath()

    def _add_action_row(self):
        
        color_idx = len(self.c.hub_actions) % len(self.c.ACTION_COLORS)
        assigned_color = self.c.ACTION_COLORS[color_idx]

        new_action = {
            'x': tk.StringVar(value=""),
            'y': tk.StringVar(value=""),
            'z': tk.StringVar(value=""),
            'rot': tk.StringVar(value=""),
            'type': tk.StringVar(value=self.c.hub_action_types[0]),
            'color': assigned_color
        }
        self.c.hub_actions.append(new_action)
        self.c._sort_hub_actions()
        self._refresh_action_table()
        self.c._draw_3d_toolpath()

    def _refresh_action_table(self):
        for widget in self.table_scroll_inner.winfo_children():
            widget.destroy()

        for i in range(7):
            self.table_scroll_inner.columnconfigure(i, weight=1 if 0 < i < 6 else 0)

        # Header Row (Row 0)
        headers = ["", "X", "Y", "Z", "ROT", "ACTION", ""]
        for i, text in enumerate(headers):
            if text: # Skip blank headers for cleaner alignment
                lbl = ttk.Label(self.table_scroll_inner, text=text, font=self.c.FONT_BODY_SMALL, 
                                foreground=self.c.COLOR_TEXT_SECONDARY, anchor=tk.CENTER)
                lbl.grid(row=0, column=i, padx=2, pady=(5, 5), sticky=tk.EW)

        infeasible_mask = self.c._get_infeasible_actions_mask()

        if not self.c.hub_actions:
            ttk.Label(self.table_scroll_inner,
                      text="\n\nNO HUB COMMANDS CONFIGURED\n\nClick '+ ADD HUB ACTION (Non-Repeated)' to begin spatial triggers.",
                      font=self.c.FONT_BODY, foreground=self.c.COLOR_TEXT_SECONDARY,
                      justify=tk.CENTER).grid(row=1, column=0, columnspan=7, pady=40, sticky="ew")

        for i, action in enumerate(self.c.hub_actions):
            grid_row = i + 1
            is_infeasible = infeasible_mask[i] if i < len(infeasible_mask) else False
            dot_color = "white" if not is_infeasible else self.c.COLOR_ACCENT_RED

            dot_canvas = tk.Canvas(self.table_scroll_inner, width=20, height=20, bg=self.c.COLOR_BG, highlightthickness=0)
            dot_canvas.grid(row=grid_row, column=0, padx=(5, 0), pady=1)
            dot_canvas.create_oval(4, 4, 16, 16, fill=action['color'], outline=dot_color, width=1)

            for j, key in enumerate(['x', 'y', 'z', 'rot']):
                ent = ttk.Entry(self.table_scroll_inner, textvariable=action[key], width=self.col_widths[j+1], font=self.c.FONT_MONO)
                ent.grid(row=grid_row, column=j+1, padx=2, pady=2, sticky=tk.EW)
                if is_infeasible:
                    ent.configure(foreground=self.c.COLOR_ACCENT_RED)
                ent.bind('<FocusOut>', lambda e: self._handle_coordinate_change())
                ent.bind('<Return>', lambda e: self._handle_coordinate_change())
                ent.bind('<KeyRelease>', lambda e, w=ent, ax=key, v=action[key]: self.c._show_coordinate_suggestions(w, ax, v))
                ent.bind('<FocusIn>', lambda e, w=ent, ax=key, v=action[key]: self.c._show_coordinate_suggestions(w, ax, v))

            cb = ttk.Combobox(self.table_scroll_inner, textvariable=action['type'], values=self.c.hub_action_types, width=self.col_widths[5], state="readonly")
            cb.grid(row=grid_row, column=5, padx=2, pady=2, sticky=tk.EW)

            btn = tk.Button(self.table_scroll_inner, text="✕", bg=self.c.COLOR_BG, fg=self.c.COLOR_ACCENT_RED,
                            bd=0, activebackground=self.c.COLOR_BLACK, activeforeground="#ffffff",
                            font=("Inter", 10, "bold"),
                            command=lambda idx=i: self._remove_action_row(idx))
            btn.grid(row=grid_row, column=6, padx=5, pady=2, sticky=tk.W)

        # --- Repeated Actions Section ---
        next_row = len(self.c.hub_actions) + 1

        sep = ttk.Label(self.table_scroll_inner, text="── REPEATED ACTIONS (EVERY POINT) ──",
                        font=self.c.FONT_BODY_SMALL, foreground=self.c.COLOR_TEXT_SECONDARY,
                        anchor=tk.CENTER)
        sep.grid(row=next_row, column=0, columnspan=7, pady=(12, 4), sticky=tk.EW)
        next_row += 1

        repeated = getattr(self.c, 'repeated_hub_actions', [])
        if not repeated:
            empty_rep = ttk.Label(self.table_scroll_inner, text="No repeated actions configured.",
                                  font=self.c.FONT_BODY_SMALL, foreground=self.c.COLOR_TEXT_SECONDARY,
                                  anchor=tk.CENTER)
            empty_rep.grid(row=next_row, column=0, columnspan=7, pady=4, sticky=tk.EW)
        else:
            for i, action in enumerate(repeated):
                row_container = tk.Frame(self.table_scroll_inner, bg=self.c.COLOR_BG)
                row_container.grid(row=next_row + i, column=0, columnspan=7, sticky="nsew", pady=1)
                row_container.columnconfigure(0, weight=1)
                row_container.columnconfigure(1, weight=1)

                cb_timing = ttk.Combobox(row_container, textvariable=action['timing'],
                                         values=["Before Measurement", "After Measurement"],
                                         width=22, state="readonly", font=self.c.FONT_MONO)
                cb_timing.grid(row=0, column=0, padx=2, pady=2, sticky=tk.EW)

                cb_action = ttk.Combobox(row_container, textvariable=action['type'],
                                          values=self.c.hub_action_types, width=14,
                                          state="readonly", font=self.c.FONT_MONO)
                cb_action.grid(row=0, column=1, padx=2, pady=2, sticky=tk.EW)

                del_btn = tk.Button(row_container, text="✕", bg=self.c.COLOR_BG,
                                    fg=self.c.COLOR_ACCENT_RED, bd=0,
                                    activebackground=self.c.COLOR_BLACK, activeforeground="#ffffff",
                                    font=("Inter", 10, "bold"),
                                    command=lambda idx=i: self.c._remove_repeated_action_row(idx))
                del_btn.grid(row=0, column=2, padx=5, pady=2, sticky=tk.W)

    def _remove_action_row(self, index):
        if 0 <= index < len(self.c.hub_actions):
            self.c.hub_actions.pop(index)
            self._refresh_action_table()
            self.c._draw_3d_toolpath()
