import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import math  # Needed for step calculation
import os
from datetime import datetime
import json  # For profile saving/loading
from typing import Any
from utils import PRINTER_LIMITS  # Single source of truth for hardware limits


class PatternGeneratorGUI:
    """
    A GUI-based tool for designing and generating scan patterns for a 3D printer.
    
    This class handles the user interface for inputting scan volume parameters (X, Y, Z, 
    and Rotation), provides a real-time 3D wireframe preview of the scan volume 
    relative to printer limits, and generates G-code or CSV files for execution.
    
    The generator uses a boustrophedon (snake-like) path optimization to minimize 
    travel time between points. It also includes safety checks to ensure the 
    requested pattern fits within the physical boundaries of the target printer.
    """
    
    # Type hints for IDE support and clarity
    generate_button: ttk.Button
    scrollable_window_id: int
    profile_name: ttk.Entry
    filename_preview: ttk.Label

    x_min_label: ttk.Label; x_min: ttk.Entry
    x_max_label: ttk.Label; x_max: ttk.Entry
    x_offset_label: ttk.Label; x_offset: ttk.Entry
    x_step: ttk.Entry

    y_min_label: ttk.Label; y_min: ttk.Entry
    y_max_label: ttk.Label; y_max: ttk.Entry
    y_offset_label: ttk.Label; y_offset: ttk.Entry
    y_step: ttk.Entry

    z_min_label: ttk.Label; z_min: ttk.Entry
    z_max_label: ttk.Label; z_max: ttk.Entry
    z_offset_label: ttk.Label; z_offset: ttk.Entry
    z_step: ttk.Entry

    rot_min_label: ttk.Label; rot_min: ttk.Entry
    rot_max_label: ttk.Label; rot_max: ttk.Entry
    rot_offset_label: ttk.Label; rot_offset: ttk.Entry
    rot_step: ttk.Entry

    travelspeed: ttk.Entry
    pause_time: ttk.Entry

    stats_text: tk.Text
    preview_canvas: tk.Canvas
    
    on_send_to_sender: Any
    send_button: ttk.Button
    generate_button: ttk.Button
    _spinner_index: int
    _spinner_chars: list
    _spinner_after_id: str

    def __init__(self, parent_frame):
        """
        Initializes the Pattern Generator interface within the provided frame.
        
        Args:
            parent_frame: The tkinter frame (or root) where this panel will be placed.
        """
        self.parent = parent_frame                      # Store reference to parent container
        self.root = parent_frame.winfo_toplevel()       # Get top-level window for modals and timers
        
        # UI Appearance: Color Palette
        self.COLOR_BG = "#0a0e14"                       # Main background (deep navy/black)
        self.COLOR_PANEL_BG = "#161b22"                 # Card/Widget background (dark gray)
        self.COLOR_BORDER = "#30363d"                   # Subdued border color
        self.COLOR_TEXT_PRIMARY = "#e6edf3"             # High-contrast text
        self.COLOR_TEXT_SECONDARY = "#7d8590"           # Dimmed metadata text
        self.COLOR_ACCENT_CYAN = "#00d4ff"              # Primary highlight color
        self.COLOR_ACCENT_PURPLE = "#a371f7"            # Secondary highlight (Rotation)
        self.COLOR_ACCENT_GREEN = "#3fb950"             # Success / Valid state
        self.COLOR_ACCENT_AMBER = "#ffa657"             # Warning / Proximity state
        self.COLOR_ACCENT_RED = "#ff4444"               # Danger / Error state
        self.COLOR_BLACK = "#000000"                    # Deep black for input fields

        # UI Appearance: Typography
        self.FONT_HEADER = ("Orbitron", 13)             # Modern futuristic font for titles
        self.FONT_BODY = ("Inter", 10)                  # Clean sans-serif for general UI
        self.FONT_BODY_SMALL = ("Inter", 9)             # Small UI labels
        self.FONT_BODY_BOLD = ("Inter", 10, "bold")     # Emphasized UI text
        self.FONT_MONO = ("JetBrains Mono", 9)          # Technical/Numeric data
        self.FONT_MONO_LARGE = ('JetBrains Mono', 11, 'bold') # Emphasized data
        self.FONT_TERMINAL = ("JetBrains Mono", 10)     # Code-like displays
        
        self._canvas_resize_timer = None                # Timer ID for debouncing UI redraws

        # Initial Configuration for ttk Styles
        style = ttk.Style()
        style.theme_use('clam')

        # Configure Global Application Styles
        style.configure('.',
                            background=self.COLOR_PANEL_BG,
                            foreground=self.COLOR_TEXT_PRIMARY,
                            fieldbackground=self.COLOR_BLACK,
                            bordercolor=self.COLOR_BORDER,
                            lightcolor=self.COLOR_BORDER,
                            darkcolor=self.COLOR_BORDER,
                            font=self.FONT_BODY)
        style.map('.',
                  background=[('disabled', self.COLOR_PANEL_BG), ('active', self.COLOR_PANEL_BG)],
                  foreground=[('disabled', self.COLOR_TEXT_SECONDARY)],
                  bordercolor=[('focus', self.COLOR_ACCENT_CYAN), ('active', self.COLOR_BORDER)],
                  fieldbackground=[('disabled', self.COLOR_PANEL_BG)])

        # Setup Specific Widget Styles
        style.configure('TFrame', background=self.COLOR_PANEL_BG)
        style.configure('Dark.TFrame', background=self.COLOR_BG)
        
        style.configure('TLabel', background=self.COLOR_PANEL_BG, foreground=self.COLOR_TEXT_PRIMARY, font=self.FONT_BODY)
        style.configure('Secondary.TLabel', background=self.COLOR_PANEL_BG, foreground=self.COLOR_TEXT_SECONDARY, font=self.FONT_BODY_SMALL)
        style.configure('Filename.TLabel', background=self.COLOR_PANEL_BG, foreground=self.COLOR_ACCENT_CYAN, font=self.FONT_MONO)

        style.configure('Card.TLabelframe',
            background=self.COLOR_PANEL_BG,
            bordercolor=self.COLOR_BORDER,
            borderwidth=1,
            relief=tk.SOLID,
            padding=12)
        style.configure('Card.TLabelframe.Label',
            background=self.COLOR_PANEL_BG,
            foreground=self.COLOR_ACCENT_CYAN,
            font=('Inter', 10, 'bold'))

        style.configure('TButton',
                            background=self.COLOR_PANEL_BG,
                            foreground=self.COLOR_TEXT_PRIMARY,
                            bordercolor=self.COLOR_BORDER,
                            borderwidth=1,
                            relief=tk.SOLID,
                            padding=(12, 8),
                            font=self.FONT_BODY)
        style.map('TButton',
                  background=[('active', '#2c333e'), ('pressed', self.COLOR_BLACK)],
                  foreground=[('active', self.COLOR_ACCENT_CYAN)],
                  bordercolor=[('active', self.COLOR_ACCENT_CYAN)])

        style.configure('Primary.TButton',
                            background=self.COLOR_ACCENT_CYAN,
                            foreground=self.COLOR_BLACK,
                            padding=(12, 10),
                            font=self.FONT_BODY_BOLD)
        style.map('Primary.TButton',
                  background=[('active', '#00eaff'), ('pressed', self.COLOR_ACCENT_CYAN)],
                  foreground=[('active', self.COLOR_BLACK), ('pressed', self.COLOR_BLACK)],
                  bordercolor=[('active', self.COLOR_ACCENT_CYAN)])

        style.configure('TEntry',
                            fieldbackground=self.COLOR_BLACK,
                            foreground=self.COLOR_ACCENT_CYAN,
                            bordercolor=self.COLOR_BORDER,
                            borderwidth=1,
                            relief=tk.SOLID,
                            padding=6,
                            font=self.FONT_MONO)
        style.map('TEntry',
                  fieldbackground=[('focus', self.COLOR_BLACK)],
                  foreground=[('focus', self.COLOR_ACCENT_CYAN)],
                  bordercolor=[('focus', self.COLOR_ACCENT_CYAN)],
                  insertcolor=[('focus', self.COLOR_ACCENT_CYAN)])
        
        style.configure('TCheckbutton',
            background=self.COLOR_PANEL_BG,
            foreground=self.COLOR_TEXT_SECONDARY,
            font=self.FONT_BODY_SMALL)
        style.map('TCheckbutton',
            background=[('active', self.COLOR_PANEL_BG)],
            foreground=[('active', self.COLOR_ACCENT_CYAN), ('selected', self.COLOR_ACCENT_CYAN)])

        style.configure('TRadiobutton',
            background=self.COLOR_PANEL_BG,
            foreground=self.COLOR_TEXT_SECONDARY,
            font=self.FONT_BODY_SMALL)
        style.map('TRadiobutton',
            background=[('active', self.COLOR_PANEL_BG)],
            foreground=[('active', self.COLOR_ACCENT_CYAN), ('selected', self.COLOR_ACCENT_CYAN)])

        style.configure('Vertical.TScrollbar',
                            background=self.COLOR_BORDER,
                            troughcolor=self.COLOR_BG,
                            bordercolor=self.COLOR_BG,
                            arrowcolor=self.COLOR_TEXT_PRIMARY,
                            relief=tk.FLAT,
                            arrowsize=14)
        style.map('Vertical.TScrollbar',
                  background=[('active', self.COLOR_ACCENT_CYAN), ('!active', self.COLOR_BORDER)],
                  troughcolor=[('active', self.COLOR_BG), ('!active', self.COLOR_BG)])

        # UI State Variables
        self.x_symmetric = tk.BooleanVar(value=True)    # If True, uses ±Offset for X instead of Min/Max
        self.y_symmetric = tk.BooleanVar(value=True)    # If True, uses ±Offset for Y instead of Min/Max
        self.z_symmetric = tk.BooleanVar(value=False)   # If True, uses ±Offset for Z instead of Min/Max
        self.rot_symmetric = tk.BooleanVar(value=True)  # If True, uses ±Offset for Rot instead of Min/Max
        self.export_format = tk.StringVar(value="gcode") # Chosen file format ('gcode' or 'csv')
        self.include_timestamp = tk.BooleanVar(value=True) # If True, appends timestamp to generated filenames

        # Layout Construction
        main_container = ttk.Frame(self.parent, style='Dark.TFrame')
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Left Column: User Inputs
        left_panel = ttk.Frame(main_container, width=350, style='Dark.TFrame')
        left_panel.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        left_panel.pack_propagate(False)

        # Right Column: Live Preview and Statistics
        right_panel = ttk.Frame(main_container, style='Dark.TFrame')
        right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        self.create_input_panel(left_panel)
        self.create_preview_panel(right_panel)

        self._load_last_parameters()

        # Sync UI states and trigger initial preview
        self._on_x_symmetric_toggle(derive_values=False)
        self._on_y_symmetric_toggle(derive_values=False)
        self._on_z_symmetric_toggle(derive_values=False)
        self._on_rot_symmetric_toggle(derive_values=False)
        self._auto_update_preview()

    # ===== NEW: WIDGET CREATION HELPERS (REMOVED) =====
    # Removed create_entry_with_glow and create_primary_button
    # Now using ttk styles

    def create_input_panel(self, parent):
        """
        Constructs the scrollable left-hand panel containing all user input controls.
        
        This includes axis range settings (Min/Max or ±Offset), step sizes, 
        movement speeds, and export format selection.
        """

        # Header section with title and tutorial access
        title_row = ttk.Frame(parent, style='Dark.TFrame')
        title_row.pack(side=tk.TOP, fill=tk.X, pady=(0, 10))

        title = ttk.Label(title_row, text="PATTERN GENERATOR",
                            font=('Rajdhani', 16, 'bold'),
                            foreground=self.COLOR_ACCENT_CYAN, background=self.COLOR_BG)
        title.pack(side=tk.LEFT, anchor=tk.W)

        self.info_button = tk.Button(
            title_row, 
            text="?", 
            font=("Segoe UI", 12, "bold"),
            bg=self.COLOR_BG, 
            fg=self.COLOR_ACCENT_CYAN,
            activebackground=self.COLOR_BG,
            activeforeground="#ffffff",
            bd=0, 
            highlightthickness=0,
            takefocus=0,
            cursor="hand2",
            padx=10,
            command=self._show_tutorial_popup
        )
        self.info_button.pack(side=tk.RIGHT, padx=(5, 95), pady=8)

        # Primary action buttons at the bottom of the input area
        button_frame = ttk.Frame(parent, style='Dark.TFrame')
        button_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 0), padx=2)

        self.generate_button = ttk.Button(button_frame, 
            text="Save File",
            command=self._start_generation_process,
            style='TButton')
        self.generate_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        self.send_button = ttk.Button(button_frame, 
            text="Load to Sender",
            command=lambda: self._start_generation_process(send_to_sender=True),
            style='Primary.TButton')
        self.send_button.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.on_send_to_sender = None
        
        # Scrollable container for the main input fields
        scroll_container = ttk.Frame(parent, style='Dark.TFrame')
        scroll_container.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=(0, 0))

        canvas = tk.Canvas(scroll_container, highlightthickness=0, bg=self.COLOR_BG)
        scrollbar = ttk.Scrollbar(scroll_container, orient="vertical", command=canvas.yview, style='Vertical.TScrollbar')
        scrollable_frame = ttk.Frame(canvas, style='Dark.TFrame')

        def _on_mousewheel(event):
            """Internal handler for mouse wheel scrolling across platforms."""
            scroll_val = 0
            if event.num == 4: # Linux Up
                scroll_val = -1
            elif event.num == 5: # Linux Down
                scroll_val = 1
            elif event.delta: # Windows/macOS
                if abs(event.delta) >= 120:
                    scroll_val = int(-1 * (event.delta / 120))
                else:
                    scroll_val = -1 * event.delta
            canvas.yview_scroll(scroll_val, "units")

        def _bind_mousewheel_recursive(widget):
            """Ensures mousewheel events work regardless of which child widget is hovered."""
            widget.bind("<MouseWheel>", _on_mousewheel)
            widget.bind("<Button-4>", _on_mousewheel)
            widget.bind("<Button-5>", _on_mousewheel)
            for child in widget.winfo_children():
                _bind_mousewheel_recursive(child)

        def set_frame_width(event):
            """Forces the scrollable frame to match the canvas width."""
            canvas_width = event.width
            canvas.itemconfig(self.scrollable_window_id, width=canvas_width)

        self.scrollable_window_id = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.bind("<Configure>", set_frame_width)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.configure(yscrollcommand=scrollbar.set)

        # Profile Naming Section
        profile_frame = ttk.LabelFrame(scrollable_frame, text="Test Profile", style='Card.TLabelframe', padding=12)
        profile_frame.pack(fill=tk.X, pady=(0, 8), padx=2)
        profile_frame.columnconfigure(1, weight=1)

        ttk.Label(profile_frame, text="Profile Name:", style='Secondary.TLabel').grid(row=0, column=0, sticky=tk.W, pady=2)
        self.profile_name = ttk.Entry(profile_frame, width=20)
        self.profile_name.insert(0, "GCODE")
        self.profile_name.grid(row=0, column=1, pady=2, padx=5, sticky=tk.EW)
        self.profile_name.bind('<KeyRelease>', self.update_filename_preview)

        ttk.Checkbutton(profile_frame, text="Include timestamp",
                            variable=self.include_timestamp,
                            command=self.update_filename_preview).grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=2)

        ttk.Label(profile_frame, text="Filename:", style='Secondary.TLabel').grid(row=2, column=0, sticky=tk.W, pady=2)
        self.filename_preview = ttk.Label(profile_frame, text="", style='Filename.TLabel', wraplength=300)
        self.filename_preview.grid(row=2, column=1, sticky=tk.W, pady=2, padx=5)
        self.update_filename_preview()

        # X-Axis Settings
        x_frame = ttk.LabelFrame(scrollable_frame, text="X Axis (mm)", style='Card.TLabelframe', padding=12)
        x_frame.pack(fill=tk.X, pady=(0, 8), padx=2)
        x_frame.columnconfigure(1, weight=1)
        x_frame.columnconfigure(3, weight=1)
        
        self.x_min_label = ttk.Label(x_frame, text="Min:", style='Secondary.TLabel')
        self.x_min_label.grid(row=0, column=0, sticky=tk.E, padx=(0,5))
        self.x_min = ttk.Entry(x_frame, width=8)
        self.x_min.insert(0, "-50")
        self.x_min.grid(row=0, column=1, padx=3, sticky=tk.EW)

        self.x_max_label = ttk.Label(x_frame, text="Max:", style='Secondary.TLabel')
        self.x_max_label.grid(row=0, column=2, sticky=tk.E, padx=(8, 5))
        self.x_max = ttk.Entry(x_frame, width=8)
        self.x_max.insert(0, "50")
        self.x_max.grid(row=0, column=3, padx=3, sticky=tk.EW)

        self.x_offset_label = ttk.Label(x_frame, text="±Offset:", style='Secondary.TLabel')
        self.x_offset_label.grid(row=0, column=0, sticky=tk.E, padx=(0,5))
        self.x_offset = ttk.Entry(x_frame, width=10)
        self.x_offset.grid(row=0, column=1, columnspan=3, sticky="ew", padx=3)
        self.x_offset_label.grid_remove()
        self.x_offset.grid_remove()

        ttk.Label(x_frame, text="Step:", style='Secondary.TLabel').grid(row=1, column=0, sticky=tk.E, pady=(8,0), padx=(0,5))
        self.x_step = ttk.Entry(x_frame, width=8)
        self.x_step.insert(0, "5")
        self.x_step.grid(row=1, column=1, padx=3, pady=(8,0), sticky=tk.EW)

        ttk.Checkbutton(x_frame, text="Symmetric", variable=self.x_symmetric, command=self._on_x_symmetric_toggle).grid(row=1, column=2, columnspan=2, sticky=tk.W, padx=8, pady=(8,0))

        self.x_min.bind('<FocusOut>', self._auto_update_preview); self.x_min.bind('<Return>', self._auto_update_preview)
        self.x_max.bind('<FocusOut>', self._auto_update_preview); self.x_max.bind('<Return>', self._auto_update_preview)
        self.x_step.bind('<FocusOut>', self._auto_update_preview); self.x_step.bind('<Return>', self._auto_update_preview)
        self.x_offset.bind('<FocusOut>', self._auto_update_preview); self.x_offset.bind('<Return>', self._auto_update_preview)

        # Y-Axis Settings
        y_frame = ttk.LabelFrame(scrollable_frame, text="Y Axis (mm)", style='Card.TLabelframe', padding=12)
        y_frame.pack(fill=tk.X, pady=(0, 8), padx=2)
        y_frame.columnconfigure(1, weight=1)
        y_frame.columnconfigure(3, weight=1)
        
        self.y_min_label = ttk.Label(y_frame, text="Min:", style='Secondary.TLabel')
        self.y_min_label.grid(row=0, column=0, sticky=tk.E, padx=(0,5))
        self.y_min = ttk.Entry(y_frame, width=8)
        self.y_min.insert(0, "-50")
        self.y_min.grid(row=0, column=1, padx=3, sticky=tk.EW)

        self.y_max_label = ttk.Label(y_frame, text="Max:", style='Secondary.TLabel')
        self.y_max_label.grid(row=0, column=2, sticky=tk.E, padx=(8, 5))
        self.y_max = ttk.Entry(y_frame, width=8)
        self.y_max.insert(0, "50")
        self.y_max.grid(row=0, column=3, padx=3, sticky=tk.EW)

        self.y_offset_label = ttk.Label(y_frame, text="±Offset:", style='Secondary.TLabel')
        self.y_offset_label.grid(row=0, column=0, sticky=tk.E, padx=(0,5))
        self.y_offset = ttk.Entry(y_frame, width=10)
        self.y_offset.grid(row=0, column=1, columnspan=3, sticky="ew", padx=3)
        self.y_offset_label.grid_remove()
        self.y_offset.grid_remove()
        
        ttk.Label(y_frame, text="Step:", style='Secondary.TLabel').grid(row=1, column=0, sticky=tk.E, pady=(8,0), padx=(0,5))
        self.y_step = ttk.Entry(y_frame, width=8)
        self.y_step.insert(0, "5")
        self.y_step.grid(row=1, column=1, padx=3, pady=(8,0), sticky=tk.EW)
        
        ttk.Checkbutton(y_frame, text="Symmetric", variable=self.y_symmetric, command=self._on_y_symmetric_toggle).grid(row=1, column=2, columnspan=2, sticky=tk.W, padx=8, pady=(8,0))
        
        self.y_min.bind('<FocusOut>', self._auto_update_preview); self.y_min.bind('<Return>', self._auto_update_preview)
        self.y_max.bind('<FocusOut>', self._auto_update_preview); self.y_max.bind('<Return>', self._auto_update_preview)
        self.y_step.bind('<FocusOut>', self._auto_update_preview); self.y_step.bind('<Return>', self._auto_update_preview)
        self.y_offset.bind('<FocusOut>', self._auto_update_preview); self.y_offset.bind('<Return>', self._auto_update_preview)

        # Z-Axis Settings
        z_frame = ttk.LabelFrame(scrollable_frame, text="Z Axis (mm)", style='Card.TLabelframe', padding=12)
        z_frame.pack(fill=tk.X, pady=(0, 8), padx=2)
        z_frame.columnconfigure(1, weight=1)
        z_frame.columnconfigure(3, weight=1)

        self.z_min_label = ttk.Label(z_frame, text="Min:", style='Secondary.TLabel')
        self.z_min_label.grid(row=0, column=0, sticky=tk.E, padx=(0,5))
        self.z_min = ttk.Entry(z_frame, width=8)
        self.z_min.insert(0, "0")
        self.z_min.grid(row=0, column=1, padx=3, sticky=tk.EW)
        
        self.z_max_label = ttk.Label(z_frame, text="Max:", style='Secondary.TLabel')
        self.z_max_label.grid(row=0, column=2, sticky=tk.E, padx=(8, 5))
        self.z_max = ttk.Entry(z_frame, width=8)
        self.z_max.insert(0, "100")
        self.z_max.grid(row=0, column=3, padx=3, sticky=tk.EW)
        
        self.z_offset_label = ttk.Label(z_frame, text="±Offset:", style='Secondary.TLabel')
        self.z_offset_label.grid(row=0, column=0, sticky=tk.E, padx=(0,5))
        self.z_offset = ttk.Entry(z_frame, width=10)
        self.z_offset.grid(row=0, column=1, columnspan=3, sticky="ew", padx=3)
        self.z_offset_label.grid_remove()
        self.z_offset.grid_remove()
        
        ttk.Label(z_frame, text="Step:", style='Secondary.TLabel').grid(row=1, column=0, sticky=tk.E, pady=(8,0), padx=(0,5))
        self.z_step = ttk.Entry(z_frame, width=8)
        self.z_step.insert(0, "5")
        self.z_step.grid(row=1, column=1, padx=3, pady=(8,0), sticky=tk.EW)
        
        ttk.Checkbutton(z_frame, text="Symmetric", variable=self.z_symmetric, command=self._on_z_symmetric_toggle).grid(row=1, column=2, columnspan=2, sticky=tk.W, padx=8, pady=(8,0))
        
        self.z_min.bind('<FocusOut>', self._auto_update_preview); self.z_min.bind('<Return>', self._auto_update_preview)
        self.z_max.bind('<FocusOut>', self._auto_update_preview); self.z_max.bind('<Return>', self._auto_update_preview)
        self.z_step.bind('<FocusOut>', self._auto_update_preview); self.z_step.bind('<Return>', self._auto_update_preview)
        self.z_offset.bind('<FocusOut>', self._auto_update_preview); self.z_offset.bind('<Return>', self._auto_update_preview)

        # Rotation Axis Settings
        rot_frame = ttk.LabelFrame(scrollable_frame, text="Rotation (degrees)", style='Card.TLabelframe', padding=12)
        rot_frame.pack(fill=tk.X, pady=(0, 8), padx=2)
        rot_frame.columnconfigure(1, weight=1)
        rot_frame.columnconfigure(3, weight=1)

        self.rot_min_label = ttk.Label(rot_frame, text="Min:", style='Secondary.TLabel')
        self.rot_min_label.grid(row=0, column=0, sticky=tk.E, padx=(0,5))
        self.rot_min = ttk.Entry(rot_frame, width=8)
        self.rot_min.insert(0, "0")
        self.rot_min.grid(row=0, column=1, padx=3, sticky=tk.EW)
        
        self.rot_max_label = ttk.Label(rot_frame, text="Max:", style='Secondary.TLabel')
        self.rot_max_label.grid(row=0, column=2, sticky=tk.E, padx=(8, 5))
        self.rot_max = ttk.Entry(rot_frame, width=8)
        self.rot_max.insert(0, "0")
        self.rot_max.grid(row=0, column=3, padx=3, sticky=tk.EW)
        
        self.rot_offset_label = ttk.Label(rot_frame, text="±Offset:", style='Secondary.TLabel')
        self.rot_offset_label.grid(row=0, column=0, sticky=tk.E, padx=(0,5))
        self.rot_offset = ttk.Entry(rot_frame, width=10)
        self.rot_offset.grid(row=0, column=1, columnspan=3, sticky="ew", padx=3)
        self.rot_offset_label.grid_remove()
        self.rot_offset.grid_remove()
        
        ttk.Label(rot_frame, text="Step:", style='Secondary.TLabel').grid(row=1, column=0, sticky=tk.E, pady=(8,0), padx=(0,5))
        self.rot_step = ttk.Entry(rot_frame, width=8)
        self.rot_step.insert(0, "5")
        self.rot_step.grid(row=1, column=1, padx=3, pady=(8,0), sticky=tk.EW)
        
        ttk.Checkbutton(rot_frame, text="Symmetric", variable=self.rot_symmetric, command=self._on_rot_symmetric_toggle).grid(row=1, column=2, columnspan=2, sticky=tk.W, padx=8, pady=(8,0))
        
        self.rot_min.bind('<FocusOut>', self._auto_update_preview); self.rot_min.bind('<Return>', self._auto_update_preview)
        self.rot_max.bind('<FocusOut>', self._auto_update_preview); self.rot_max.bind('<Return>', self._auto_update_preview)
        self.rot_step.bind('<FocusOut>', self._auto_update_preview); self.rot_step.bind('<Return>', self._auto_update_preview)
        self.rot_offset.bind('<FocusOut>', self._auto_update_preview); self.rot_offset.bind('<Return>', self._auto_update_preview)

        # Global Movement Settings
        movement_frame = ttk.LabelFrame(scrollable_frame, text="Movement Settings", style='Card.TLabelframe', padding=12)
        movement_frame.pack(fill=tk.X, pady=(0, 8), padx=2)
        movement_frame.columnconfigure(1, weight=1)

        ttk.Label(movement_frame, text="Travel Speed (mm/min):", style='Secondary.TLabel').grid(row=0, column=0, sticky=tk.E, padx=(0,5))
        self.travelspeed = ttk.Entry(movement_frame, width=10)
        self.travelspeed.insert(0, "3000")
        self.travelspeed.grid(row=0, column=1, padx=3, sticky=tk.W)

        ttk.Label(movement_frame, text="Pause (seconds):", style='Secondary.TLabel').grid(row=1, column=0, sticky=tk.E, pady=(8,0), padx=(0,5))
        self.pause_time = ttk.Entry(movement_frame, width=10)
        self.pause_time.insert(0, "1")
        self.pause_time.grid(row=1, column=1, padx=3, pady=(8,0), sticky=tk.W)

        self.travelspeed.bind('<FocusOut>', self._auto_update_preview); self.travelspeed.bind('<Return>', self._auto_update_preview)
        self.pause_time.bind('<FocusOut>', self._auto_update_preview); self.pause_time.bind('<Return>', self._auto_update_preview)
        
        # Export Selection
        export_frame = ttk.LabelFrame(scrollable_frame, text="Export Format", style='Card.TLabelframe', padding=12)
        export_frame.pack(fill=tk.X, pady=(0, 8), padx=2)

        ttk.Radiobutton(export_frame, text="G-code (.gcode)", variable=self.export_format, value="gcode", command=self._auto_update_preview).pack(anchor=tk.W)
        ttk.Radiobutton(export_frame, text="CSV Coordinates (.csv)", variable=self.export_format, value="csv", command=self._auto_update_preview).pack(anchor=tk.W)

        _bind_mousewheel_recursive(scrollable_frame)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        canvas.bind("<MouseWheel>", _on_mousewheel)
        canvas.bind("<Button-4>", _on_mousewheel)
        canvas.bind("<Button-5>", _on_mousewheel)


    def create_preview_panel(self, parent):
        """
        Constructs the right-hand panel for volume visualization and statistics.
        
        This panel includes the 3D wireframe canvas and a detailed statistics 
        readout that updates in real-time as the user modifies parameters.
        """

        title = ttk.Label(parent, text="SCAN VOLUME PREVIEW",
                            font=('Rajdhani', 16, 'bold'),
                            foreground=self.COLOR_ACCENT_CYAN, background=self.COLOR_BG)
        title.pack(side=tk.TOP, pady=(0, 10), anchor=tk.W)

        # Bottom Section: Statistics Readout
        stats_frame = ttk.LabelFrame(parent, text="Scan Statistics", style='Card.TLabelframe', padding=0)
        stats_frame.pack(fill=tk.X, side=tk.BOTTOM, pady=(10, 0))

        self.stats_text = tk.Text(stats_frame, height=9, width=50,
                                    state='disabled', wrap=tk.WORD,
                                    bg=self.COLOR_BLACK,
                                    fg=self.COLOR_ACCENT_GREEN,
                                    font=self.FONT_MONO,
                                    relief=tk.FLAT,
                                    bd=0,
                                    padx=12,
                                    pady=10
                                    )
        self.stats_text.pack(fill=tk.BOTH, expand=True)

        # Style tags for statistics formatting
        self.stats_text.tag_configure('header', foreground=self.COLOR_ACCENT_CYAN, font=self.FONT_MONO_LARGE)
        self.stats_text.tag_configure('value', foreground=self.COLOR_ACCENT_GREEN, font=self.FONT_MONO_LARGE)
        self.stats_text.tag_configure('warning', foreground=self.COLOR_ACCENT_RED, font=(self.FONT_MONO[0], self.FONT_MONO[1], 'bold'))
        self.stats_text.tag_configure('amber_warning', foreground=self.COLOR_ACCENT_AMBER, font=(self.FONT_MONO[0], self.FONT_MONO[1], 'bold'))
        self.stats_text.tag_configure('success', foreground=self.COLOR_ACCENT_GREEN, font=(self.FONT_MONO[0], self.FONT_MONO[1], 'bold'))
        self.stats_text.tag_configure('label', foreground=self.COLOR_TEXT_SECONDARY, font=self.FONT_MONO)
        self.stats_text.tag_configure('separator', foreground=self.COLOR_BORDER, font=self.FONT_MONO)

        # Top Section: 3D Visualization Canvas
        canvas_frame = ttk.Frame(parent, style='Dark.TFrame')
        canvas_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=(0, 0))

        self.preview_canvas = tk.Canvas(canvas_frame, bg=self.COLOR_BLACK, highlightthickness=1,
                                            highlightbackground=self.COLOR_BORDER)
        self.preview_canvas.pack(fill=tk.BOTH, expand=True)
        self.preview_canvas.bind("<Configure>", self._on_canvas_resize)

        self.draw_preview_diagram(None, [], 0)

    # ===== NEW: SYMMETRIC TOGGLE HANDLERS =====
    # *** STYLE UPDATE: Changed from _wrapper to the entry widget itself ***

    def _on_x_symmetric_toggle(self, derive_values=True):
        """Toggles X-axis UI between Min/Max inputs and a single ±Offset input."""
        self._toggle_symmetric_widgets(
            self.x_symmetric.get(),
            self.x_min_label, self.x_min,
            self.x_max_label, self.x_max,
            self.x_offset_label, self.x_offset,
            self.x_min, self.x_max, self.x_offset, "-50", "50",
            derive_values=derive_values
        )
    
    def _on_y_symmetric_toggle(self, derive_values=True):
        """Toggles Y-axis UI between Min/Max inputs and a single ±Offset input."""
        self._toggle_symmetric_widgets(
            self.y_symmetric.get(),
            self.y_min_label, self.y_min,
            self.y_max_label, self.y_max,
            self.y_offset_label, self.y_offset,
            self.y_min, self.y_max, self.y_offset, "-50", "50",
            derive_values=derive_values
        )
        
    def _on_z_symmetric_toggle(self, derive_values=True):
        """Toggles Z-axis UI between Min/Max inputs and a single ±Offset input."""
        self._toggle_symmetric_widgets(
            self.z_symmetric.get(),
            self.z_min_label, self.z_min,
            self.z_max_label, self.z_max,
            self.z_offset_label, self.z_offset,
            self.z_min, self.z_max, self.z_offset, "0", "100",
            derive_values=derive_values
        )
        
    def _on_rot_symmetric_toggle(self, derive_values=True):
        """Toggles Rotation UI between Min/Max inputs and a single ±Offset input."""
        self._toggle_symmetric_widgets(
            self.rot_symmetric.get(),
            self.rot_min_label, self.rot_min,
            self.rot_max_label, self.rot_max,
            self.rot_offset_label, self.rot_offset,
            self.rot_min, self.rot_max, self.rot_offset, "0", "0",
            derive_values=derive_values
        )

    def _save_last_parameters(self):
        """Persists current UI settings to a local JSON file for session recovery."""
        try:
            settings = {
                'profile_name': self.profile_name.get(),
                'include_timestamp': self.include_timestamp.get(),
                'export_format': self.export_format.get(),
                
                'x_symmetric': self.x_symmetric.get(),
                'x_min': self.x_min.get(),
                'x_max': self.x_max.get(),
                'x_offset': self.x_offset.get(),
                'x_step': self.x_step.get(),
                
                'y_symmetric': self.y_symmetric.get(),
                'y_min': self.y_min.get(),
                'y_max': self.y_max.get(),
                'y_offset': self.y_offset.get(),
                'y_step': self.y_step.get(),
                
                'z_symmetric': self.z_symmetric.get(),
                'z_min': self.z_min.get(),
                'z_max': self.z_max.get(),
                'z_offset': self.z_offset.get(),
                'z_step': self.z_step.get(),
                
                'rot_symmetric': self.rot_symmetric.get(),
                'rot_min': self.rot_min.get(),
                'rot_max': self.rot_max.get(),
                'rot_offset': self.rot_offset.get(),
                'rot_step': self.rot_step.get(),
                
                'travelspeed': self.travelspeed.get(),
                'pause_time': self.pause_time.get()
            }
            with open('last_scan_profile.json', 'w') as f:
                json.dump(settings, f)
        except Exception as e:
            print(f"Failed to auto-save parameters: {e}")

    def _load_last_parameters(self):
        """Restores UI settings from the local JSON file if it exists."""
        import os
        try:
            if os.path.exists('last_scan_profile.json'):
                with open('last_scan_profile.json', 'r') as f:
                    settings = json.load(f)
                
                def set_entry(widget, value):
                    widget.delete(0, 'end')
                    widget.insert(0, str(value))
                    
                if 'profile_name' in settings: set_entry(self.profile_name, settings['profile_name'])
                if 'include_timestamp' in settings: self.include_timestamp.set(settings['include_timestamp'])
                if 'export_format' in settings: self.export_format.set(settings['export_format'])
                
                if 'x_symmetric' in settings: self.x_symmetric.set(settings['x_symmetric'])
                if 'x_min' in settings: set_entry(self.x_min, settings['x_min'])
                if 'x_max' in settings: set_entry(self.x_max, settings['x_max'])
                if 'x_offset' in settings: set_entry(self.x_offset, settings['x_offset'])
                if 'x_step' in settings: set_entry(self.x_step, settings['x_step'])
                
                if 'y_symmetric' in settings: self.y_symmetric.set(settings['y_symmetric'])
                if 'y_min' in settings: set_entry(self.y_min, settings['y_min'])
                if 'y_max' in settings: set_entry(self.y_max, settings['y_max'])
                if 'y_offset' in settings: set_entry(self.y_offset, settings['y_offset'])
                if 'y_step' in settings: set_entry(self.y_step, settings['y_step'])
                
                if 'z_symmetric' in settings: self.z_symmetric.set(settings['z_symmetric'])
                if 'z_min' in settings: set_entry(self.z_min, settings['z_min'])
                if 'z_max' in settings: set_entry(self.z_max, settings['z_max'])
                if 'z_offset' in settings: set_entry(self.z_offset, settings['z_offset'])
                if 'z_step' in settings: set_entry(self.z_step, settings['z_step'])
                
                if 'rot_symmetric' in settings: self.rot_symmetric.set(settings['rot_symmetric'])
                if 'rot_min' in settings: set_entry(self.rot_min, settings['rot_min'])
                if 'rot_max' in settings: set_entry(self.rot_max, settings['rot_max'])
                if 'rot_offset' in settings: set_entry(self.rot_offset, settings['rot_offset'])
                if 'rot_step' in settings: set_entry(self.rot_step, settings['rot_step'])
                
                if 'travelspeed' in settings: set_entry(self.travelspeed, settings['travelspeed'])
                if 'pause_time' in settings: set_entry(self.pause_time, settings['pause_time'])
        except Exception as e:
            print(f"Failed to auto-load parameters: {e}")

    def _toggle_symmetric_widgets(self, is_symmetric, 
                                    min_lbl, min_entry_widget, max_lbl, max_entry_widget,
                                    off_lbl, off_entry_widget,
                                    min_entry, max_entry, off_entry,
                                    default_min, default_max,
                                    derive_values=True):
        """
        Generic UI helper to toggle between asymmetric (Min/Max) and 
        symmetric (±Offset) input fields for any given axis.
        """
        
        if is_symmetric:
            # SYMMETRIC MODE: Hide min/max, show offset
            if derive_values:
                try:
                    current_min = float(min_entry.get())
                    current_max = float(max_entry.get())
                    # Offset is the max of the absolute values
                    offset = max(abs(current_min), abs(current_max))
                    off_entry.delete(0, tk.END)
                    off_entry.insert(0, f"{offset:g}")
                except ValueError:
                    # Fallback to defaults on invalid numeric input
                    try:
                        offset = max(abs(float(default_min)), abs(float(default_max)))
                        off_entry.delete(0, tk.END)
                        off_entry.insert(0, f"{offset:g}")
                    except ValueError:
                        off_entry.insert(0, "50")
            
            min_lbl.grid_remove(); min_entry_widget.grid_remove()
            max_lbl.grid_remove(); max_entry_widget.grid_remove()
            off_lbl.grid(); off_entry_widget.grid()
            
        else:
            # ASYMMETRIC MODE: Show min/max, hide offset
            if derive_values:
                try:
                    offset = abs(float(off_entry.get()))
                    min_entry.delete(0, tk.END)
                    min_entry.insert(0, f"{-offset:g}")
                    max_entry.delete(0, tk.END)
                    max_entry.insert(0, f"{offset:g}")
                except ValueError:
                    # Restore defaults on invalid numeric input
                    min_entry.delete(0, tk.END); min_entry.insert(0, default_min)
                    max_entry.delete(0, tk.END); max_entry.insert(0, default_max)
            
            min_lbl.grid(); min_entry_widget.grid()
            max_lbl.grid(); max_entry_widget.grid()
            off_lbl.grid_remove(); off_entry_widget.grid_remove()
        
        self._auto_update_preview()
        
    # ===== END SYMMETRIC HANDLERS =====


    def draw_preview_diagram(self, params, bounds_warnings, warning_level=0):
        """
        Renders a 3D wireframe visualization of the scan volume on the 2D canvas.
        
        Args:
            params: Dictionary of scan parameters.
            bounds_warnings: List of strings describing which limits are violated.
            warning_level: 0 for OK, 1 for proximity warning (amber), 2 for exceeded (red).
        """
        self.preview_canvas.delete("all")

        canvas_w = self.preview_canvas.winfo_width()
        canvas_h = self.preview_canvas.winfo_height()

        # Handle early calls before canvas is fully rendered
        if canvas_w <= 1 or canvas_h <= 1:
            self.preview_canvas.after(50, self.draw_preview_diagram, params, bounds_warnings, warning_level)
            return

        if params is None:
            self.preview_canvas.create_text(
                canvas_w / 2, canvas_h / 2,
                text="Waiting for valid parameters...",
                fill=self.COLOR_TEXT_SECONDARY,
                anchor=tk.CENTER, font=self.FONT_BODY
            )
            # Default to zero-volume centered box for initial render
            params = {'x_min': 0, 'x_max': 0, 'y_min': 0, 'y_max': 0, 'z_min': 0, 'z_max': 0}
        
        x_range = params['x_max'] - params['x_min']
        y_range = params['y_max'] - params['y_min']
        z_range = params['z_max'] - params['z_min']

        # Determine effective scene boundaries (Union of pattern and printer limits)
        total_min_x = min(params['x_min'], -PRINTER_LIMITS['x'])
        total_max_x = max(params['x_max'], PRINTER_LIMITS['x'])
        total_min_y = min(params['y_min'], -PRINTER_LIMITS['y'])
        total_max_y = max(params['y_max'], PRINTER_LIMITS['y'])
        total_min_z = min(params['z_min'], PRINTER_LIMITS['z_min'])
        total_max_z = max(params['z_max'], PRINTER_LIMITS['z_max'])

        total_x_rng = total_max_x - total_min_x
        total_y_rng = total_max_y - total_min_y
        total_z_rng = total_max_z - total_min_z
        
        eff_total_x_rng = total_x_rng if total_x_rng != 0 else 1
        eff_total_y_rng = total_y_rng if total_y_rng != 0 else 1
        eff_total_z_rng = total_z_rng if total_z_rng != 0 else 1

        # Calculation constants for oblique projection
        pad = 40
        oblique_factor = 0.4

        total_w_units = eff_total_x_rng + (eff_total_y_rng * oblique_factor)
        total_h_units = eff_total_z_rng + (eff_total_y_rng * oblique_factor)

        scale_x = (canvas_w - 2 * pad) / total_w_units
        scale_y = (canvas_h - 2 * pad) / total_h_units
        scale = max(0, min(scale_x, scale_y))

        def project(x, y, z):
            """Projects 3D (x,y,z) coordinates to 2D (canvas_x, canvas_y)."""
            x_pct = (x - total_min_x) / eff_total_x_rng if eff_total_x_rng != 0 else 0.5
            y_pct = (y - total_min_y) / eff_total_y_rng if eff_total_y_rng != 0 else 0.5
            z_pct = (z - total_min_z) / eff_total_z_rng if eff_total_z_rng != 0 else 0.5
            
            scaled_drawing_w = total_x_rng * scale
            scaled_drawing_h = total_z_rng * scale
            scaled_drawing_d = total_y_rng * scale * oblique_factor

            x_start = (canvas_w - (scaled_drawing_w + scaled_drawing_d)) / 2
            y_start = (canvas_h - (scaled_drawing_h + scaled_drawing_d)) / 2

            screen_w = x_pct * scaled_drawing_w
            screen_h = (1 - z_pct) * scaled_drawing_h
            screen_d = y_pct * scaled_drawing_d
            
            return (x_start + screen_w + screen_d, y_start + screen_h + screen_d)

        # Draw Printer Boundaries (Safety Box)
        pl = PRINTER_LIMITS
        pb1 = project(-pl['x'], -pl['y'], pl['z_min'])
        pb2 = project( pl['x'], -pl['y'], pl['z_min'])
        pb3 = project( pl['x'], -pl['y'], pl['z_max'])
        pb4 = project(-pl['x'], -pl['y'], pl['z_max'])
        pb5 = project(-pl['x'],  pl['y'], pl['z_min'])
        pb6 = project( pl['x'],  pl['y'], pl['z_min'])
        pb7 = project( pl['x'],  pl['y'], pl['z_max'])
        pb8 = project(-pl['x'],  pl['y'], pl['z_max'])
        
        def draw_warning_line(start, end):
            self.preview_canvas.create_line(start, end, fill=self.COLOR_ACCENT_RED, dash=(4, 4), width=2)

        draw_warning_line(pb1, pb2); draw_warning_line(pb2, pb3); draw_warning_line(pb3, pb4); draw_warning_line(pb4, pb1)
        draw_warning_line(pb5, pb6); draw_warning_line(pb6, pb7); draw_warning_line(pb7, pb8); draw_warning_line(pb8, pb5)
        draw_warning_line(pb1, pb5); draw_warning_line(pb2, pb6); draw_warning_line(pb3, pb7); draw_warning_line(pb4, pb8)
        
        self.preview_canvas.create_text((pb4[0] + pb8[0])/2, pb8[1] - 5,
            text="Printer Limits", fill=self.COLOR_ACCENT_RED, font=self.FONT_BODY_SMALL, anchor=tk.S)


        # Draw Pattern Volume (The requested scan area)
        if x_range != 0 or y_range != 0 or z_range != 0:
            p1 = project(params['x_min'], params['y_min'], params['z_min'])
            p2 = project(params['x_max'], params['y_min'], params['z_min'])
            p3 = project(params['x_max'], params['y_min'], params['z_max'])
            p4 = project(params['x_min'], params['y_min'], params['z_max'])
            p5 = project(params['x_min'], params['y_max'], params['z_min'])
            p6 = project(params['x_max'], params['y_max'], params['z_min'])
            p7 = project(params['x_max'], params['y_max'], params['z_max'])
            p8 = project(params['x_min'], params['y_max'], params['z_max'])
            
            def draw_hidden_line(start, end):
                self.preview_canvas.create_line(start, end, fill=self.COLOR_ACCENT_CYAN, dash=(2, 4), width=1)
                
            def draw_visible_line(start, end):
                self.preview_canvas.create_line(start, end, fill=self.COLOR_ACCENT_CYAN, width=2)
            
            # Use dashed lines for back-facing edges
            draw_hidden_line(p2, p6); draw_hidden_line(p1, p2); draw_hidden_line(p2, p3)
            
            # Use solid lines for front-facing edges
            draw_visible_line(p5, p6); draw_visible_line(p6, p7); draw_visible_line(p3, p4)
            draw_visible_line(p4, p1); draw_visible_line(p3, p7); draw_visible_line(p4, p8)
            draw_visible_line(p7, p8); draw_visible_line(p8, p5); draw_visible_line(p1, p5)


        # Draw Origin (Home) Marker if it's within the current view
        origin_in_bounds = (total_min_x <= 0 <= total_max_x and
                            total_min_y <= 0 <= total_max_y and
                            total_min_z <= 0 <= total_max_z)

        if origin_in_bounds:
            (ox, oy) = project(0, 0, 0)
            self.preview_canvas.create_oval(ox - 8, oy - 8, ox + 8, oy + 8, fill='', outline='#ff6666', width=1)
            self.preview_canvas.create_oval(ox - 5, oy - 5, ox + 5, oy + 5,
                                            fill=self.COLOR_ACCENT_RED, outline=self.COLOR_ACCENT_RED, width=2)


        # Visual Warning Labels
        if warning_level == 2:
            text_color, warn_text = self.COLOR_ACCENT_RED, "⚠️ BOUNDS EXCEEDED"
        elif warning_level == 1:
            text_color, warn_text = self.COLOR_ACCENT_AMBER, "⚠️ PROXIMITY WARNING"
        else:
            warn_text = None
        
        if warn_text:
            self.preview_canvas.create_text(canvas_w - 20, 40,
                text=warn_text, fill=text_color, font=self.FONT_MONO_LARGE, anchor=tk.NE)

            # Identify which specific axes are problematic
            problem_axes = set()
            for warning_string in bounds_warnings:
                if warning_string.startswith("X"): problem_axes.add("X-AXIS")
                elif warning_string.startswith("Y"): problem_axes.add("Y-AXIS")
                elif warning_string.startswith("Z"): problem_axes.add("Z-AXIS")
            
            axes_text = " / ".join(sorted(list(problem_axes)))
            if axes_text:
                self.preview_canvas.create_text(canvas_w - 20, 65,
                    text=axes_text, fill=text_color, font=self.FONT_MONO_LARGE, anchor=tk.NE)
            
        # Rotation Preview (Compass-like arc)
        rot_r = 35
        rot_cx, rot_cy = canvas_w - rot_r - 20, canvas_h - rot_r - 20
        
        # Background arc representing the available rotation field
        self.preview_canvas.create_arc(
            rot_cx - rot_r, rot_cy - rot_r, rot_cx + rot_r, rot_cy + rot_r,
            start=0, extent=180, fill=self.COLOR_BORDER, outline='', style=tk.PIESLICE
        )
        self.preview_canvas.create_arc(
            rot_cx - rot_r, rot_cy - rot_r, rot_cx + rot_r, rot_cy + rot_r,
            start=180, extent=180, fill=self.COLOR_PANEL_BG, outline=self.COLOR_TEXT_SECONDARY, style=tk.CHORD, width=1
        )
        
        if params is not None and 'rot_min' in params and 'rot_max' in params:
            rmin, rmax = max(-90, min(90, params['rot_min'])), max(-90, min(90, params['rot_max']))
            # Map degrees to tkinter arc angles (270 is South/0°)
            self.preview_canvas.create_arc(
                rot_cx - rot_r, rot_cy - rot_r, rot_cx + rot_r, rot_cy + rot_r,
                start=270 + rmin, extent=rmax - rmin, fill=self.COLOR_ACCENT_PURPLE, outline=self.COLOR_ACCENT_CYAN, style=tk.PIESLICE
            )
        
        # Labels for rotation compass
        self.preview_canvas.create_line(rot_cx, rot_cy, rot_cx, rot_cy + rot_r + 5, fill=self.COLOR_TEXT_SECONDARY)
        self.preview_canvas.create_text(rot_cx, rot_cy + rot_r + 12, text="0°", fill=self.COLOR_TEXT_SECONDARY, font=("Inter", 8))
        self.preview_canvas.create_text(rot_cx - rot_r - 15, rot_cy, text="-90°", fill=self.COLOR_TEXT_SECONDARY, font=("Inter", 8))
        self.preview_canvas.create_text(rot_cx + rot_r + 15, rot_cy, text="+90°", fill=self.COLOR_TEXT_SECONDARY, font=("Inter", 8))

    def update_filename_preview(self, event=None):
        """Generates and displays a preview of the filename based on the profile name and settings."""
        name = self.profile_name.get()
        # Sanitize name to remove filesystem-illegal characters
        name = "".join(c for c in name if c.isalnum() or c in ('-', '_'))

        if self.include_timestamp.get():
            timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
            filename = f"{name}_{timestamp}"
        else:
            filename = f"{name}"
        
        # Append extension based on current export format
        ext = ".csv" if self.export_format.get() == 'csv' else ".gcode"
        self.filename_preview.config(text=filename + ext)

    def _get_params_silently(self):
        """
        Attempts to read and parse all UI parameters. 
        Returns a dictionary of floats if successful, or None if any input is invalid.
        This method suppresses all UI popups/errors.
        """
        params = {}
        try:
            # Parse X-Axis
            if self.x_symmetric.get():
                offset = abs(float(self.x_offset.get()))
                params['x_min'], params['x_max'] = -offset, offset
            else:
                params['x_min'], params['x_max'] = float(self.x_min.get()), float(self.x_max.get())
            params['x_step'] = float(self.x_step.get())
            
            # Parse Y-Axis
            if self.y_symmetric.get():
                offset = abs(float(self.y_offset.get()))
                params['y_min'], params['y_max'] = -offset, offset
            else:
                params['y_min'], params['y_max'] = float(self.y_min.get()), float(self.y_max.get())
            params['y_step'] = float(self.y_step.get())
            
            # Parse Z-Axis
            if self.z_symmetric.get():
                offset = abs(float(self.z_offset.get()))
                params['z_min'], params['z_max'] = -offset, offset
            else:
                params['z_min'], params['z_max'] = float(self.z_min.get()), float(self.z_max.get())
            params['z_step'] = float(self.z_step.get())
            
            # Parse Rotation
            if self.rot_symmetric.get():
                offset = abs(float(self.rot_offset.get()))
                params['rot_min'], params['rot_max'] = -offset, offset
            else:
                params['rot_min'], params['rot_max'] = float(self.rot_min.get()), float(self.rot_max.get())
            params['rot_step'] = float(self.rot_step.get())
            
            # Parse Movement Settings
            params['travelspeed'] = float(self.travelspeed.get())
            params['pause_time'] = float(self.pause_time.get())

            # Validation Logic: Ensure ranges and steps are logical
            if params['x_min'] > params['x_max']: return None
            if params['y_min'] > params['y_max']: return None
            if params['z_min'] > params['z_max']: return None
            if params['rot_min'] > params['rot_max']: return None
            
            if any(params[f'{ax}_step'] < 0 for ax in ['x', 'y', 'z', 'rot']): return None
            
            # Ensure non-zero steps for non-zero ranges
            for ax in ['x', 'y', 'z', 'rot']:
                if params[f'{ax}_min'] != params[f'{ax}_max'] and params[f'{ax}_step'] == 0:
                    return None
            
            if params['travelspeed'] <= 0 or params['pause_time'] < 0:
                return None
                
            return params
        except ValueError:
            return None

    def get_parameters(self):
        """
        Reads and parses all UI parameters, showing error dialogs for invalid inputs.
        Returns a dictionary of parameters or None if validation fails.
        """
        params = {}
        try:
            # Re-use the same parsing logic as silent read
            if self.x_symmetric.get():
                offset = abs(float(self.x_offset.get())); params['x_min'], params['x_max'] = -offset, offset
            else:
                params['x_min'], params['x_max'] = float(self.x_min.get()), float(self.x_max.get())
            params['x_step'] = float(self.x_step.get())
            
            if self.y_symmetric.get():
                offset = abs(float(self.y_offset.get())); params['y_min'], params['y_max'] = -offset, offset
            else:
                params['y_min'], params['y_max'] = float(self.y_min.get()), float(self.y_max.get())
            params['y_step'] = float(self.y_step.get())
            
            if self.z_symmetric.get():
                offset = abs(float(self.z_offset.get())); params['z_min'], params['z_max'] = -offset, offset
            else:
                params['z_min'], params['z_max'] = float(self.z_min.get()), float(self.z_max.get())
            params['z_step'] = float(self.z_step.get())
            
            if self.rot_symmetric.get():
                offset = abs(float(self.rot_offset.get())); params['rot_min'], params['rot_max'] = -offset, offset
            else:
                params['rot_min'], params['rot_max'] = float(self.rot_min.get()), float(self.rot_max.get())
            params['rot_step'] = float(self.rot_step.get())
            
            params['travelspeed'] = float(self.travelspeed.get())
            params['pause_time'] = float(self.pause_time.get())

            # Detailed validation with user feedback
            if params['x_min'] > params['x_max']: messagebox.showerror("Error", "X Min must be <= X Max"); return None
            if params['y_min'] > params['y_max']: messagebox.showerror("Error", "Y Min must be <= Y Max"); return None
            if params['z_min'] > params['z_max']: messagebox.showerror("Error", "Z Min must be <= Z Max"); return None
            if params['rot_min'] > params['rot_max']: messagebox.showerror("Error", "Rot Min must be <= Rot Max"); return None
            
            if any(params[f'{ax}_step'] < 0 for ax in ['x', 'y', 'z', 'rot']): 
                messagebox.showerror("Error", "Step sizes must be positive numbers."); return None
                
            for ax in ['x', 'y', 'z', 'rot']:
                if params[f'{ax}_min'] != params[f'{ax}_max'] and params[f'{ax}_step'] == 0:
                    messagebox.showerror("Error", f"{ax.upper()} Step must be > 0 if range is non-zero."); return None
            
            if params['travelspeed'] <= 0: messagebox.showerror("Error", "Travel Speed must be > 0"); return None
            if params['pause_time'] < 0: messagebox.showerror("Error", "Pause Time cannot be negative"); return None
            
            return params
        except ValueError:
            messagebox.showerror("Error", "All input fields must contain valid numbers."); return None

    def generate_step_values(self, min_val, max_val, step):
        """
        Creates a list of discrete coordinates from min to max using the specified step.
        Uses a small epsilon to handle floating point precision issues at the max boundary.
        """
        if step == 0: return [min_val]
        values = []
        current = min_val
        while current <= max_val + 1e-9:
            values.append(round(current, 6))
            current += step
        return values

    def _calculate_total_points(self, params):
        """Calculates the total number of positions in the scan without generating the list."""
        def count_steps(min_val, max_val, step):
            if min_val > max_val: return 0
            if step == 0 or min_val == max_val: return 1
            return int(math.floor((max_val - min_val) / step + 1e-9)) + 1
        try:
            nx = count_steps(params['x_min'], params['x_max'], params['x_step'])
            ny = count_steps(params['y_min'], params['y_max'], params['y_step'])
            nz = count_steps(params['z_min'], params['z_max'], params['z_step'])
            nr = count_steps(params['rot_min'], params['rot_max'], params['rot_step'])
            return nx * ny * nz * nr
        except Exception:
            return 0

    def create_pattern(self, params):
        """
        A generator that yields individual (x, y, z, rot) scan coordinates.
        
        Uses a serpentine (boustrophedon) path for the X-axis to minimize 
        travel distances. The sequence of loops is: 
        Rotation -> Z-Plane -> Y-Line -> X-Points.
        """
        x_values = self.generate_step_values(params['x_min'], params['x_max'], params['x_step'])
        y_values = self.generate_step_values(params['y_min'], params['y_max'], params['y_step'])
        z_values = self.generate_step_values(params['z_min'], params['z_max'], params['z_step'])
        rot_values = self.generate_step_values(params['rot_min'], params['rot_max'], params['rot_step'])

        for rot in rot_values:
            for z in z_values:
                x_direction_forward = True
                for y in y_values:
                    # Alternate X direction for every Y-line to create snake path
                    x_iterator = x_values if x_direction_forward else reversed(x_values)
                    for x in x_iterator:
                        yield {'x': x, 'y': y, 'z': z, 'rotation': rot}
                    x_direction_forward = not x_direction_forward

    def _format_time(self, total_seconds):
        """Converts raw seconds into a human-readable string (e.g., '1d 2h 3m 4s')."""
        if total_seconds < 0: return "0s"
        total_seconds = int(total_seconds)
        d, rem = divmod(total_seconds, 86400)
        h, rem = divmod(rem, 3600)
        m, s = divmod(rem, 60)
        parts = []
        if d > 0: parts.append(f"{d}d")
        if h > 0: parts.append(f"{h}h")
        if m > 0: parts.append(f"{m}m")
        if s > 0 or not parts: parts.append(f"{s}s")
        return " ".join(parts)

    def _calculate_estimated_time(self, params, total_points):
        """
        Estimates total execution time by modeling printer kinematics.
        Calculates acceleration/deceleration ramps for every move to provide 
        an accurate prediction of real-world hardware performance.
        """
        if total_points == 0 or params is None: return 0
        
        # Calculate total static overhead from pauses
        total_pause_s = max(0, (total_points - 1) * params['pause_time'])
        
        # v_max converted to mm/s
        v_max = params['travelspeed'] / 60.0
        if v_max <= 0: return total_pause_s
        
        # Acceleration model (based on typical i3-style firmware defaults)
        accel = 500.0 # mm/s^2 
        t_accel = v_max / accel
        d_accel = 0.5 * accel * (t_accel ** 2)
        
        def calculate_move_time(distance):
            """Calculates time for a single axis move using a trapezoidal velocity profile."""
            if distance <= 0: return 0
            if distance >= 2 * d_accel:
                # Trapezoidal: can reach full cruise speed
                return 2 * t_accel + (distance - 2 * d_accel) / v_max
            else:
                # Triangular: move is too short to reach cruise speed
                return 2 * math.sqrt(distance / accel)

        def count_steps(min_v, max_v, stp):
            if min_v > max_v: return 0
            if stp == 0 or min_v == max_v: return 1
            return int(math.floor((max_v - min_v) / stp + 1e-9)) + 1
            
        nx = count_steps(params['x_min'], params['x_max'], params['x_step'])
        ny = count_steps(params['y_min'], params['y_max'], params['y_step'])
        nz = count_steps(params['z_min'], params['z_max'], params['z_step'])
        nr = count_steps(params['rot_min'], params['rot_max'], params['rot_step'])
        
        total_travel_s = 0.0

        # Calculate time for X-axis increments (serpentine primary axis)
        if nx > 1:
            total_travel_s += (nx - 1) * ny * nz * nr * calculate_move_time(params['x_step'])
            
        # Calculate time for Y-axis increments (line steps)
        if ny > 1:
            total_travel_s += (ny - 1) * nz * nr * calculate_move_time(params['y_step'])
            
        # Calculate time for Z-axis increments (layer steps)
        if nz > 1:
            total_travel_s += (nz - 1) * nr * calculate_move_time(params['z_step'])
            
        # Calculate time for Rotational increments
        if nr > 1:
            total_travel_s += (nr - 1) * calculate_move_time(params['rot_step'])
            
        return total_travel_s + total_pause_s

    def update_statistics(self, params, total_points, bounds_warnings, warning_level=0):        
        """Updates the textual scan statistics display and applies color-coded warning tags."""
        self.stats_text.config(state=tk.NORMAL)
        self.stats_text.delete(1.0, tk.END)
        
        if params is None: 
            self.stats_text.insert(1.0, "Waiting for valid parameters...", 'warning')
            self.stats_text.config(state=tk.DISABLED)
            return

        # Descriptive metrics
        xr, yr, zr = params['x_max'] - params['x_min'], params['y_max'] - params['y_min'], params['z_max'] - params['z_min']
        
        def count_steps(min_v, max_v, stp):
            if min_v > max_v: return 0
            if stp == 0 or min_v == max_v: return 1
            return int(math.floor((max_v - min_v) / stp + 1e-9)) + 1
            
        nx, ny, nz, nr = [count_steps(params[f'{ax}_min'], params[f'{ax}_max'], params[f'{ax}_step']) for ax in ['x', 'y', 'z', 'rot']]
        
        est_secs = self._calculate_estimated_time(params, total_points)
        time_str = self._format_time(est_secs)

        sep = "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
        self.stats_text.insert(tk.END, sep, 'separator')
        
        self.stats_text.insert(tk.END, "Total Points:   ", 'label')
        self.stats_text.insert(tk.END, f"{total_points:,}\n", 'value')
        
        self.stats_text.insert(tk.END, "Grid (X,Y,Z,R): ", 'label')
        self.stats_text.insert(tk.END, f"{nx}×{ny}×{nz}×{nr}\n", 'value')

        self.stats_text.insert(tk.END, "Volume:          ", 'label')
        self.stats_text.insert(tk.END, f"{xr:.1f}×{yr:.1f}×{zr:.1f} mm³\n", 'value')

        self.stats_text.insert(tk.END, "Est. Runtime:   ", 'label')
        self.stats_text.insert(tk.END, f"{time_str}\n", 'value')
        
        self.stats_text.insert(tk.END, sep, 'separator')

        # Warning Status Headers
        if warning_level == 2:
            self.stats_text.insert(tk.END, "⚠️ BOUNDS EXCEEDED!\n", 'warning')
            tag = 'warning'
        elif warning_level == 1:
            self.stats_text.insert(tk.END, "⚠️ PROXIMITY WARNING\n", 'amber_warning')
            tag = 'amber_warning'
        else:
            self.stats_text.insert(tk.END, "✓ Pattern fits printer bounds\n", 'success')
            tag = 'success'

        # Specific Violation Details
        if bounds_warnings:
            for warning in bounds_warnings:
                self.stats_text.insert(tk.END, f"  {warning}\n", tag)

        self.stats_text.config(state=tk.DISABLED)

    def _check_printer_bounds(self, params):
        """
        Validates the requested scan pattern against physical printer limits.
        
        Returns:
            A tuple of (warning_list, warning_level).
            warning_level: 0=OK, 1=Near limit (Amber), 2=Exceeds limit (Red).
        """
        if params is None: return ([], 0)
            
        warnings = []
        warning_level = 0
        pl = PRINTER_LIMITS
        proximity = 10.0 # Define a 10mm safety buffer for proximity warnings
        
        # Check X-Axis Extents
        pattern_x_extent = max(abs(params['x_min']), abs(params['x_max']))
        if pattern_x_extent > pl['x']:
            warnings.append(f"X extent ({pattern_x_extent:.1f}mm) > limit ({pl['x']:.1f}mm)")
            warning_level = 2
        elif pattern_x_extent > pl['x'] - proximity:
             warnings.append(f"X extent ({pattern_x_extent:.1f}mm) near limit ({pl['x']:.1f}mm)")
             warning_level = max(warning_level, 1)
            
        # Check Y-Axis Extents
        pattern_y_extent = max(abs(params['y_min']), abs(params['y_max']))
        if pattern_y_extent > pl['y']:
            warnings.append(f"Y extent ({pattern_y_extent:.1f}mm) > limit ({pl['y']:.1f}mm)")
            warning_level = 2
        elif pattern_y_extent > pl['y'] - proximity:
             warnings.append(f"Y extent ({pattern_y_extent:.1f}mm) near limit ({pl['y']:.1f}mm)")
             warning_level = max(warning_level, 1)
            
        # Check Z-Axis Ceiling
        if params['z_max'] > pl['z_max']:
            warnings.append(f"Z max ({params['z_max']:.1f}mm) > limit ({pl['z_max']:.1f}mm)")
            warning_level = 2
        elif params['z_max'] > pl['z_max'] - proximity:
            warnings.append(f"Z max ({params['z_max']:.1f}mm) near limit ({pl['z_max']:.1f}mm)")
            warning_level = max(warning_level, 1)

        # Check Z-Axis Floor (Collision risk)
        if params['z_min'] < pl['z_min']:
             warnings.append(f"Z min ({params['z_min']:.1f}mm) < limit ({pl['z_min']:.1f}mm)")
             warning_level = 2
             
        # Check Rotational Range
        if params.get('rot_max', 0) > pl.get('rot_max', 90.0):
            warnings.append(f"Rot max ({params.get('rot_max', 0):.1f}°) > limit ({pl.get('rot_max', 90.0):.1f}°)")
            warning_level = 2
        elif params.get('rot_max', 0) > pl.get('rot_max', 90.0) - proximity:
            warnings.append(f"Rot max ({params.get('rot_max', 0):.1f}°) near limit ({pl.get('rot_max', 90.0):.1f}°)")
            warning_level = max(warning_level, 1)

        if params.get('rot_min', 0) < pl.get('rot_min', -90.0):
            warnings.append(f"Rot min ({params.get('rot_min', 0):.1f}°) < limit ({pl.get('rot_min', -90.0):.1f}°)")
            warning_level = 2
        elif params.get('rot_min', 0) < pl.get('rot_min', -90.0) + proximity:
            warnings.append(f"Rot min ({params.get('rot_min', 0):.1f}°) near limit ({pl.get('rot_min', -90.0):.1f}°)")
            warning_level = max(warning_level, 1)

        return (warnings, warning_level)

    def _auto_update_preview(self, event=None):
        """Triggered on UI events to refresh the 3D diagram, statistics, and filename preview."""
        params = self._get_params_silently()
        total_points = self._calculate_total_points(params) if params else 0
        
        bounds_warnings, warning_level = self._check_printer_bounds(params)
        
        self.update_statistics(params, total_points, bounds_warnings, warning_level)
        self.draw_preview_diagram(params, bounds_warnings, warning_level)
        self.update_filename_preview()

    def _animate_spinner(self):
        """Internal helper to cycle the 'Working...' animation characters on the Load button."""
        if not hasattr(self, '_spinner_index'): return
        chars = self._spinner_chars
        c = chars[self._spinner_index % len(chars)]
        try:
            self.send_button.config(text=f"{c}  Working...")
            self._spinner_index += 1
            self._spinner_after_id = self.root.after(150, self._animate_spinner)
        except Exception:
            pass # Handle widget destruction during animation

    def _restore_send_button(self):
        """Stops spinner animations and resets buttons to their interactive state."""
        if hasattr(self, '_spinner_after_id'):
            try: self.root.after_cancel(self._spinner_after_id)
            except Exception: pass
        if hasattr(self, '_spinner_index'): del self._spinner_index
        try:
            self.send_button.configure(text="Load to Sender", state=tk.NORMAL)
            self.generate_button.configure(state=tk.NORMAL)
            self.root.update_idletasks()
        except Exception:
            pass

    def _start_generation_process(self, send_to_sender=False):
        """
        Coordinates the logic for file creation. 
        Validates parameters, checks safety limits, and routes to either 
        direct saving or loading into the Sender panel.
        """
        params = self.get_parameters()
        if params is None: return
        self._save_last_parameters()
        
        total_points = self._calculate_total_points(params)
        if total_points == 0: 
            messagebox.showerror("Error", "The current pattern contains 0 points."); return
            
        # Perform Physical Bounds Check (Safety Override)
        bounds_warnings, warning_level = self._check_printer_bounds(params)
        if warning_level == 2:
            warning_msg = "SAFETY WARNING:\n\nThe pattern exceeds the physical printer limits:\n"
            for w in bounds_warnings: warning_msg += f"- {w}\n"
            warning_msg += "\nThis may cause hardware damage or crashes.\nAre you absolutely sure you want to proceed?"
            if not messagebox.askyesno("Safety Override", warning_msg, icon='warning'):
                return
        
        # Confirm for extremely large datasets
        if total_points > 1_000_000:
            if not messagebox.askokcancel("Large Dataset", f"Pattern contains {total_points:,} points.\nGeneration may be slow. Continue?"): 
                return
        
        # Sanitize profile name for filename
        name = "".join(c for c in self.profile_name.get() if c.isalnum() or c in ('-', '_'))
        if not name: 
            messagebox.showerror("Error", "Please enter a valid profile name."); return
        
        ts = datetime.now().strftime("_%Y%m%d-%H%M%S") if self.include_timestamp.get() else ""
        format_choice = self.export_format.get()
        
        # Scenario A: User clicked 'Load to Sender'
        if send_to_sender:
            import tempfile, os, threading
            temp_dir = tempfile.gettempdir()
            fname = os.path.join(temp_dir, f"{name}{ts}.gcode")
            
            # Switch to 'Working' state
            self.send_button.configure(text="Working...", state=tk.DISABLED)
            self.generate_button.configure(state=tk.DISABLED)
            self.root.update_idletasks()
            self._spinner_index = 0
            self._spinner_chars = ['\u25dc', '\u25dd', '\u25de', '\u25df']
            self._animate_spinner()
            
            def _do_generate():
                try:
                    self._generate_gcode_file(params, total_points, fname, send_to_sender=True)
                finally:
                    # Reset UI state on completion
                    self.root.after(0, self._restore_send_button)
            
            threading.Thread(target=_do_generate, daemon=True).start()
            return
        
        # Scenario B: User clicked 'Save File' (Manual export)
        if format_choice == "gcode":
            fname = filedialog.asksaveasfilename(
                title="Save G-code", defaultextension=".gcode",
                initialdir='Data Files',
                initialfile=f"{name}{ts}.gcode",
                filetypes=[("G-code", "*.gcode"), ("All Files", "*.*")]
            )
            if fname: self._generate_gcode_file(params, total_points, fname, send_to_sender=False)

        elif format_choice == "csv":
            fname = filedialog.asksaveasfilename(
                title="Save CSV", defaultextension=".csv",
                initialdir='Data Files',
                initialfile=f"{name}{ts}.csv",
                filetypes=[("CSV (Comma-separated)", "*.csv"), ("All Files", "*.*")]
            )
            if fname: self._generate_csv_file(params, total_points, fname)
    def _generate_gcode_file(self, params, total_points, fname, send_to_sender=False):
        """Generates a full G-code file with header metadata and pattern commands."""
        try:
            pattern_gen = self.create_pattern(params)
            gcode_gen = self.create_gcode(pattern_gen, params, total_points)
            
            with open(fname, 'w') as f:
                for line in gcode_gen: f.write(line + "\n")
                    
            if send_to_sender and self.on_send_to_sender:
                self.on_send_to_sender(fname)
            else:
                messagebox.showinfo("Success", f"G-code saved to:\n{fname}")
        except Exception as e: 
            messagebox.showerror("Error", f"Failed to write G-code file:\n{e}")

    def _generate_csv_file(self, params, total_points, fname):
        """Generates a CSV file containing the raw coordinate list."""
        try:
            pattern_gen = self.create_pattern(params)
            csv_gen = self.create_csv_data(pattern_gen, params, total_points)
            
            with open(fname, 'w') as f:
                for line in csv_gen: f.write(line + "\n")
                    
            messagebox.showinfo("Success", f"CSV saved to:\n{fname}")
        except Exception as e: 
            messagebox.showerror("Error", f"Failed to write CSV file:\n{e}")
    
    # ===== END FILE GENERATION =====

    # ===== PROFILE HANDLING METHODS =====

    def _get_profile_data(self):
        """
        Compiles all current UI settings into a serializable dictionary.
        This includes axis symmetries, ranges, speeds, and chosen export formats.
        """
        profile_data = {
            'profile_name': self.profile_name.get(),
            'include_timestamp': self.include_timestamp.get(),
            'x_symmetric': self.x_symmetric.get(),
            'y_symmetric': self.y_symmetric.get(),
            'z_symmetric': self.z_symmetric.get(),
            'rot_symmetric': self.rot_symmetric.get(),
            'x_step': self.x_step.get(),
            'y_step': self.y_step.get(),
            'z_step': self.z_step.get(),
            'rot_step': self.rot_step.get(),
            'travelspeed': self.travelspeed.get(),
            'pause_time': self.pause_time.get(),
            'export_format': self.export_format.get()
        }
        
        # Save specific axis extents based on their current symmetry mode
        if self.x_symmetric.get(): profile_data['x_offset'] = self.x_offset.get()
        else: profile_data['x_min'], profile_data['x_max'] = self.x_min.get(), self.x_max.get()
            
        if self.y_symmetric.get(): profile_data['y_offset'] = self.y_offset.get()
        else: profile_data['y_min'], profile_data['y_max'] = self.y_min.get(), self.y_max.get()
            
        if self.z_symmetric.get(): profile_data['z_offset'] = self.z_offset.get()
        else: profile_data['z_min'], profile_data['z_max'] = self.z_min.get(), self.z_max.get()
        
        if self.rot_symmetric.get(): profile_data['rot_offset'] = self.rot_offset.get()
        else: profile_data['rot_min'], profile_data['rot_max'] = self.rot_min.get(), self.rot_max.get()
            
        return profile_data

    def load_profile(self):
        """
        Opens a file dialog to select a previously generated G-code file and 
        attempts to extract and restore the scan profile embedded in its header.
        """
        fname = filedialog.askopenfilename(
            title="Load Profile from G-code",
            initialdir='Data Files',
            filetypes=[("G-code Files", "*.gcode"), ("All Files", "*.*")]
        )
        if not fname: return

        MAGIC_PREFIX = "; PROFILE_JSON: "
        profile_data = None

        try:
            with open(fname, 'r') as f:
                for line in f:
                    if line.startswith(MAGIC_PREFIX):
                        profile_data = json.loads(line.removeprefix(MAGIC_PREFIX))
                        break

            if not isinstance(profile_data, dict):
                messagebox.showerror("Error", "No valid profile metadata found in the selected file.")
                return

            # UI Update Helpers
            def set_widget(widget, key):
                if profile_data and key in profile_data:
                    widget.delete(0, tk.END)
                    widget.insert(0, str(profile_data[key]))
            
            def set_var(var, key, default=None):
                var.set(profile_data.get(key, default) if profile_data else default)

            # Restore basic parameters
            set_widget(self.profile_name, 'profile_name')
            set_var(self.include_timestamp, 'include_timestamp', default=True)
            set_widget(self.x_step, 'x_step'); set_widget(self.y_step, 'y_step')
            set_widget(self.z_step, 'z_step'); set_widget(self.rot_step, 'rot_step')
            set_widget(self.travelspeed, 'travelspeed'); set_widget(self.pause_time, 'pause_time')
            set_var(self.export_format, 'export_format', default='gcode')

            # Restore Axis-Specific Symmetry States and Values
            set_var(self.x_symmetric, 'x_symmetric', default=True)
            if profile_data.get('x_symmetric'): set_widget(self.x_offset, 'x_offset')
            else: set_widget(self.x_min, 'x_min'); set_widget(self.x_max, 'x_max')
            
            set_var(self.y_symmetric, 'y_symmetric', default=True)
            if profile_data.get('y_symmetric'): set_widget(self.y_offset, 'y_offset')
            else: set_widget(self.y_min, 'y_min'); set_widget(self.y_max, 'y_max')

            set_var(self.z_symmetric, 'z_symmetric', default=False)
            if profile_data.get('z_symmetric'): set_widget(self.z_offset, 'z_offset')
            else: set_widget(self.z_min, 'z_min'); set_widget(self.z_max, 'z_max')

            set_var(self.rot_symmetric, 'rot_symmetric', default=True)
            if profile_data.get('rot_symmetric'): set_widget(self.rot_offset, 'rot_offset')
            else: set_widget(self.rot_min, 'rot_min'); set_widget(self.rot_max, 'rot_max')

            # Synchronize the visibility of UI fields
            self._on_x_symmetric_toggle(derive_values=False)
            self._on_y_symmetric_toggle(derive_values=False)
            self._on_z_symmetric_toggle(derive_values=False)
            self._on_rot_symmetric_toggle(derive_values=False)

            self._auto_update_preview()
            messagebox.showinfo("Success", "Scan profile successfully restored.")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to parse profile data: {e}")

    def _show_tutorial_popup(self):
        """
        Launches a modal tutorial window. 
        Loads instructions from 'HINTS_TUTORIAL.txt' to assist new users.
        """
        popup = tk.Toplevel(self.root)
        popup.title("SEED Control Center - Tutorial")
        popup.geometry("850x700")
        popup.grab_set() 
        popup.configure(bg=self.COLOR_BG)

        header = tk.Frame(popup, bg=self.COLOR_BLACK, height=60)
        header.pack(side=tk.TOP, fill=tk.X)
        header.pack_propagate(False)

        tk.Label(header, text="ⓘ  MACHINE OPERATING GUIDE", font=("Segoe UI", 16, "bold"),
                 fg=self.COLOR_ACCENT_CYAN, bg=self.COLOR_BLACK).pack(side=tk.LEFT, padx=20)

        content_frame = tk.Frame(popup, bg=self.COLOR_BG)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Standard Text area with Scrollbar for better color control
        scrollbar = tk.Scrollbar(content_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        txt_area = tk.Text(content_frame, wrap=tk.WORD, font=("Consolas", 11),
                           bg="#000000", fg="#ffffff",
                           insertbackground=self.COLOR_ACCENT_CYAN,
                           padx=15, pady=15, bd=0, highlightthickness=1,
                           highlightbackground=self.COLOR_BORDER,
                           yscrollcommand=scrollbar.set)
        txt_area.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=txt_area.yview)

        tutorial_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "HINTS_TUTORIAL.txt")
        content = "Guide content not found."
        
        if os.path.exists(tutorial_file):
            try:
                with open(tutorial_file, "r", encoding="utf-8-sig") as f:
                    content = f.read()
            except Exception as e:
                content = f"Error reading guide: {e}"
        else:
            content = f"Critical Error: HINTS_TUTORIAL.txt not found at:\n{tutorial_file}"

        txt_area.delete("1.0", tk.END)
        txt_area.insert("1.0", content)
        txt_area.configure(state=tk.DISABLED)

        footer = tk.Frame(popup, bg=self.COLOR_BG, height=70)
        footer.pack(side=tk.BOTTOM, fill=tk.X)

        tk.Button(footer, text="Got it!", font=("Segoe UI", 12, "bold"),
                  bg=self.COLOR_ACCENT_GREEN, fg="#000000", activebackground="#ffffff",
                  relief=tk.FLAT, padx=30, pady=10, cursor="hand2",
                  command=popup.destroy).pack(pady=15)

        popup.update_idletasks()
        rx = self.root.winfo_x() + (self.root.winfo_width() - popup.winfo_width()) // 2
        ry = self.root.winfo_y() + (self.root.winfo_height() - popup.winfo_height()) // 2
        popup.geometry(f"+{rx}+{ry}")

    def create_gcode(self, pattern_generator, params, total_points):
        """
        A generator that yields lines of G-code for the scan pattern.
        Embeds a JSON representation of the scan profile in the header for 
        future restoration and auditability.
        """
        profile_json = json.dumps(self._get_profile_data())

        # Header Metadata
        yield f"; Pattern: {self.profile_name.get()}"
        yield f"; Generated: {datetime.now():%Y-%m-%d %H:%M:%S}"
        yield f"; Total points: {total_points:,}"
        yield f"; Speed: {params['travelspeed']} mm/min"
        yield f"; Pause: {params['pause_time']} s"
        yield ";"
        yield f"; PROFILE_JSON: {profile_json}"
        yield ";"
        yield "; NOTE: Coordinates are relative to the scan center (0,0,0)."
        yield ""

        # Initialization Block
        yield "; === INIT ==="
        yield "G28 ; Home all axes"
        yield "G90 ; Absolute positioning"
        yield "M82 ; Absolute extruder mode"
        yield "G1 X0 Y0 Z0 ; Define starting state for parser (relative to center)"
        yield ""

        # Pattern Execution Loop
        yield "; === PATTERN ==="
        last_z = None
        first_point = True
        for i, pos in enumerate(pattern_generator, 1):
            if first_point:
                yield "; Initial Safe Move (Z first)"
                yield f"G1 Z{pos['z']:.3f} F{params['travelspeed']:.0f}"
                yield f"G1 X{pos['x']:.3f} Y{pos['y']:.3f} E{pos['rotation']:.3f} F{params['travelspeed']:.0f}"
                if params['pause_time'] > 0 and i < total_points:
                    yield f"G4 P{int(params['pause_time'] * 1000)}"
                first_point, last_z = False, pos['z']
                continue

            # Re-home X/Y on layer change to mitigate cumulative stepper drift
            if last_z is not None and pos['z'] != last_z:
                yield "; --- Layer Sync & Re-home ---"
                yield "G28 X Y"
                yield f"G1 Z{pos['z']:.3f} F{params['travelspeed']:.0f}"
            
            if i % 10000 == 0: yield f"; --- Progress: {i:,}/{total_points:,} ---"
            yield f"G1 X{pos['x']:.3f} Y{pos['y']:.3f} Z{pos['z']:.3f} E{pos['rotation']:.3f} F{params['travelspeed']:.0f}"
            if params['pause_time'] > 0 and i < total_points:
                yield f"G4 P{int(params['pause_time'] * 1000)}"
            last_z = pos['z']

        # Shutdown Block
        yield ""
        yield "; === END ==="
        yield f"G1 X0 Y0 Z{params['z_max']} E0 F{params['travelspeed']:.0f} ; Safe height return"
        yield "G90"
        yield "; Scan complete"

    def create_csv_data(self, pattern_generator, params, total_points):
        """A generator that yields lines of CSV data representing the scan points."""
        yield "Point,X,Y,Z,Rotation"
        for i, pos in enumerate(pattern_generator, 1):
            yield f"{i},{pos['x']:.3f},{pos['y']:.3f},{pos['z']:.3f},{pos['rotation']:.1f}"

    def _on_canvas_resize(self, event):
        """Handles window resize events by debouncing the redraw logic to maintain performance."""
        if self._canvas_resize_timer:
            self.root.after_cancel(self._canvas_resize_timer)
        self._canvas_resize_timer = self.root.after(25, self._perform_delayed_redraw)

    def _perform_delayed_redraw(self):
        """Executes the actual UI refresh after the debounce timeout has passed."""
        self._canvas_resize_timer = None
        self._auto_update_preview()


# --- Main Execution ---
if __name__ == "__main__":
    root = tk.Tk()
    
    # --- A simple fix for blurry fonts on Windows ---
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1) # type: ignore
    except Exception:
        pass 
        
    app = PatternGeneratorGUI(root)
    root.mainloop()