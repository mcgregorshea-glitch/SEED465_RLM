import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import math  # Needed for step calculation
from datetime import datetime
import json  # For profile saving/loading

# ===== NEW: PRINTER BOUNDARY CONSTANTS =====
PRINTER_LIMITS = {
    'x': 110.0,  # 220mm / 2
    'y': 110.0,  # 220mm / 2
    'z_max': 250.0,
    'z_min': 0.0
}


class PatternGeneratorGUI:
    """
    GUI application for generating 3D printer scan pattern G-code.
    Creates a file with G-code commands based on user parameters.
    
    --- MODIFIED ---
    - Matched G-Code Sender styling (ttk.Entry, Primary.TButton, colors, fonts).
    - Replaced tk.Entry glow wrappers with styled ttk.Entry.
    - Replaced tk.Button with styled ttk.Button (Primary.TButton).
    - Fixed left panel width by disabling pack_propagate and adding wraplength
      to the filename label.
    - Fixed stats panel layout bug (packed to bottom).
    - Fixed canvas hidden line visibility (dashed cyan).
    - Fixed asymmetric entry field widths (Min/Max) using sticky=tk.E.
    - Updated menu item label to be more descriptive.
    - Fixed scrollbar styling.
    """

    def __init__(self, root):
        """Initialize the GUI window and all widgets"""
        self.root = root
        
        self.root.title("◢ PATTERN GENERATOR")
        self.root.geometry("850x700")
        self.root.minsize(750, 600)
        
        # --- Color Palette (from G-Code Sender) ---
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

        # --- Fonts (from G-Code Sender) ---
        self.FONT_HEADER = ("Orbitron", 13)
        self.FONT_BODY = ("Inter", 10) # Changed from 11
        self.FONT_BODY_SMALL = ("Inter", 9)
        self.FONT_BODY_BOLD = ("Inter", 10, "bold") # Changed from 11
        self.FONT_MONO = ("JetBrains Mono", 9) # Changed from 10
        self.FONT_MONO_LARGE = ('JetBrains Mono', 11, 'bold')
        self.FONT_TERMINAL = ("JetBrains Mono", 10)
        
        self.root.configure(bg=self.COLOR_BG)
        
        # Timer ID for debouncing canvas resizes
        self._canvas_resize_timer = None

        # --- NEW: Setup ttk Styling (Merged from G-Code Sender) ---
        style = ttk.Style()
        style.theme_use('clam')

        # --- Global Style ---
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

        # --- Frame Styles ---
        style.configure('TFrame', background=self.COLOR_PANEL_BG)
        style.configure('Dark.TFrame', background=self.COLOR_BG) # Use self.COLOR_BG
        
        # --- Label Styles ---
        style.configure('TLabel', background=self.COLOR_PANEL_BG, foreground=self.COLOR_TEXT_PRIMARY, font=self.FONT_BODY)
        style.configure('Secondary.TLabel', background=self.COLOR_PANEL_BG, foreground=self.COLOR_TEXT_SECONDARY, font=self.FONT_BODY_SMALL) # Use BODY_SMALL
        style.configure('Filename.TLabel', background=self.COLOR_PANEL_BG, foreground=self.COLOR_ACCENT_CYAN, font=self.FONT_MONO)

        # --- LabelFrame (Panel) Style ---
        style.configure('Card.TLabelframe',
            background=self.COLOR_PANEL_BG,
            bordercolor=self.COLOR_BORDER,
            borderwidth=1,
            relief=tk.SOLID,
            padding=12)
        style.configure('Card.TLabelframe.Label',
            background=self.COLOR_PANEL_BG,
            foreground=self.COLOR_ACCENT_CYAN, # Use Cyan for title
            font=('Inter', 10, 'bold'))

        # --- Button Styles (from G-Code Sender) ---
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

        # --- Primary Action Button (from G-Code Sender) ---
        style.configure('Primary.TButton',
                            background=self.COLOR_ACCENT_CYAN,
                            foreground=self.COLOR_BLACK,
                            padding=(12, 10), # Added more padding
                            font=self.FONT_BODY_BOLD)
        style.map('Primary.TButton',
                  background=[('active', '#00eaff'), ('pressed', self.COLOR_ACCENT_CYAN)],
                  foreground=[('active', self.COLOR_BLACK), ('pressed', self.COLOR_BLACK)],
                  bordercolor=[('active', self.COLOR_ACCENT_CYAN)])

        # --- Entry (Input) Style (from G-Code Sender) ---
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
                  insertcolor=[('focus', self.COLOR_ACCENT_CYAN)]) # Add cursor color
        
        # Checkbutton style
        style.configure('TCheckbutton',
            background=self.COLOR_PANEL_BG,
            foreground=self.COLOR_TEXT_SECONDARY,
            font=self.FONT_BODY_SMALL) # Use BODY_SMALL
        style.map('TCheckbutton',
            background=[('active', self.COLOR_PANEL_BG)],
            foreground=[('active', self.COLOR_ACCENT_CYAN), ('selected', self.COLOR_ACCENT_CYAN)])

        # Radiobutton style
        style.configure('TRadiobutton',
            background=self.COLOR_PANEL_BG,
            foreground=self.COLOR_TEXT_SECONDARY,
            font=self.FONT_BODY_SMALL) # Use BODY_SMALL
        style.map('TRadiobutton',
            background=[('active', self.COLOR_PANEL_BG)],
            foreground=[('active', self.COLOR_ACCENT_CYAN), ('selected', self.COLOR_ACCENT_CYAN)])

        # Scrollbar style (from G-Code Sender)
        # *** FIX: Use Vertical.TScrollbar ***
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
        # --- End of Styling ---

        # Variables for symmetric checkboxes
        self.x_symmetric = tk.BooleanVar(value=True)
        self.y_symmetric = tk.BooleanVar(value=True)
        self.z_symmetric = tk.BooleanVar(value=False)
        self.rot_symmetric = tk.BooleanVar(value=True)
        
        # NEW: Variable for export format
        self.export_format = tk.StringVar(value="gcode")
        
        # *** FIX ***: Initialize the missing attribute
        self.include_timestamp = tk.BooleanVar(value=True)


        # ===== MENU BAR =====
        self.menu_bar = tk.Menu(root,
            bg=self.COLOR_PANEL_BG, 
            fg=self.COLOR_TEXT_PRIMARY,
            activebackground=self.COLOR_ACCENT_CYAN,
            activeforeground=self.COLOR_BLACK,
            font=self.FONT_BODY,
            relief=tk.FLAT,
            bd=0
        )
        root.config(menu=self.menu_bar)

        file_menu = tk.Menu(self.menu_bar, 
            tearoff=0,
            bg=self.COLOR_PANEL_BG,
            fg=self.COLOR_TEXT_PRIMARY,
            activebackground=self.COLOR_ACCENT_CYAN,
            activeforeground=self.COLOR_BLACK,
            font=self.FONT_BODY
        )
        self.menu_bar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Load Profile from G-code...", command=self.load_profile)
        file_menu.add_separator()
        # *** FIX: Updated menu label ***
        file_menu.add_command(label="Generate File (.gcode or .csv)...", command=self._start_generation_process)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=root.quit)

        # Create main container
        main_container = ttk.Frame(root, style='Dark.TFrame')
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Left panel for inputs (fixed width)
        # *** FIX ***: Enforce fixed width
        left_panel = ttk.Frame(main_container, width=350, style='Dark.TFrame')
        left_panel.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        left_panel.pack_propagate(False) # *** IMPORTANT ***

        # Right panel for visualization
        right_panel = ttk.Frame(main_container, style='Dark.TFrame')
        right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # ===== LEFT PANEL - INPUT FIELDS =====
        self.create_input_panel(left_panel)

        # ===== RIGHT PANEL - VISUALIZATION =====
        self.create_preview_panel(right_panel)

        # Trigger initial symmetric UI state & preview
        self._on_x_symmetric_toggle()
        self._on_y_symmetric_toggle()
        self._on_z_symmetric_toggle()
        self._on_rot_symmetric_toggle()
        self._auto_update_preview()

    # ===== NEW: WIDGET CREATION HELPERS (REMOVED) =====
    # Removed create_entry_with_glow and create_primary_button
    # Now using ttk styles

    def create_input_panel(self, parent):
        """Create all input fields and controls"""

        # Title
        title = ttk.Label(parent, text="PATTERN GENERATOR",
                            font=('Rajdhani', 16, 'bold'),
                            foreground=self.COLOR_ACCENT_CYAN, background=self.COLOR_BG)
        title.pack(side=tk.TOP, pady=(0, 10), anchor=tk.W)

        # --- MODIFIED: ACTION BUTTONS (Moved to bottom of panel) ---
        button_frame = ttk.Frame(parent, style='Dark.TFrame')
        button_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 0), padx=2)

        self.generate_button = ttk.Button(button_frame, 
            text="Generate G-Code", # <-- MODIFIED: Text changed
            command=self._start_generation_process,
            style='Primary.TButton')
        self.generate_button.pack(fill=tk.X)
        
        # --- MODIFIED: New container for scrollable area ---
        scroll_container = ttk.Frame(parent, style='Dark.TFrame')
        scroll_container.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=(0, 0))

        # Scrollable frame for inputs
        canvas = tk.Canvas(scroll_container, highlightthickness=0, bg=self.COLOR_BG)
        # *** STYLE UPDATE ***
        scrollbar = ttk.Scrollbar(scroll_container, orient="vertical", command=canvas.yview, style='Vertical.TScrollbar')
        scrollable_frame = ttk.Frame(canvas, style='Dark.TFrame')

        # --- NEW: Mousewheel Scrolling ---
        def _on_mousewheel(event):
            """Handles cross-platform mouse wheel scrolling."""
            scroll_val = 0
            if event.num == 4: # Linux scroll up
                scroll_val = -1
            elif event.num == 5: # Linux scroll down
                scroll_val = 1
            elif event.delta: # Windows/macOS
                # Normalize delta
                if abs(event.delta) >= 120:
                    scroll_val = int(-1 * (event.delta / 120))
                else:
                    scroll_val = -1 * event.delta
            
            canvas.yview_scroll(scroll_val, "units")

        def _bind_mousewheel_recursive(widget):
            """Recursively binds mousewheel events to all child widgets."""
            widget.bind("<MouseWheel>", _on_mousewheel)
            widget.bind("<Button-4>", _on_mousewheel) # Linux
            widget.bind("<Button-5>", _on_mousewheel) # Linux
            for child in widget.winfo_children():
                _bind_mousewheel_recursive(child)
        # --- End Mousewheel Scrolling ---

        # --- FIX: Bind canvas <Configure> to set scrollable_frame width ---
        def set_frame_width(event):
            canvas_width = event.width
            canvas.itemconfig(self.scrollable_window_id, width=canvas_width)

        self.scrollable_window_id = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.bind("<Configure>", set_frame_width)
        # --- END FIX ---
        
        # This bind is for the scrollregion, it is also necessary
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.configure(yscrollcommand=scrollbar.set)

        # --- Profile Name ---
        profile_frame = ttk.LabelFrame(scrollable_frame, text="Test Profile", style='Card.TLabelframe', padding=12)
        profile_frame.pack(fill=tk.X, pady=(0, 8), padx=2)
        profile_frame.columnconfigure(1, weight=1)

        ttk.Label(profile_frame, text="Profile Name:", style='Secondary.TLabel').grid(row=0, column=0, sticky=tk.W, pady=2)
        # *** STYLE UPDATE: Replaced create_entry_with_glow with ttk.Entry ***
        self.profile_name = ttk.Entry(profile_frame, width=20)
        self.profile_name.insert(0, "New_Test_Pattern")
        self.profile_name.grid(row=0, column=1, pady=2, padx=5, sticky=tk.EW)
        self.profile_name.bind('<KeyRelease>', self.update_filename_preview)

        ttk.Checkbutton(profile_frame, text="Include timestamp",
                            variable=self.include_timestamp,
                            command=self.update_filename_preview).grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=2)

        ttk.Label(profile_frame, text="Filename:", style='Secondary.TLabel').grid(row=2, column=0, sticky=tk.W, pady=2)
        # *** FIX: Added wraplength=300 ***
        self.filename_preview = ttk.Label(profile_frame, text="", style='Filename.TLabel', wraplength=300)
        self.filename_preview.grid(row=2, column=1, sticky=tk.W, pady=2, padx=5)
        self.update_filename_preview()

        # --- X AXIS ---
        x_frame = ttk.LabelFrame(scrollable_frame, text="X Axis (mm)", style='Card.TLabelframe', padding=12)
        x_frame.pack(fill=tk.X, pady=(0, 8), padx=2)
        x_frame.columnconfigure(1, weight=1) # Make entry widgets expand
        x_frame.columnconfigure(3, weight=1) # *** FIX: Make Max expand too ***
        
        # Min/Max (Asymmetric)
        # *** FIX: Use sticky=tk.E for right-alignment ***
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

        # Offset (Symmetric)
        self.x_offset_label = ttk.Label(x_frame, text="±Offset:", style='Secondary.TLabel')
        self.x_offset_label.grid(row=0, column=0, sticky=tk.E, padx=(0,5))
        self.x_offset = ttk.Entry(x_frame, width=10)
        self.x_offset.grid(row=0, column=1, columnspan=3, sticky="ew", padx=3)
        self.x_offset_label.grid_remove()
        self.x_offset.grid_remove()

        # Step/Symmetric
        ttk.Label(x_frame, text="Step:", style='Secondary.TLabel').grid(row=1, column=0, sticky=tk.E, pady=(8,0), padx=(0,5))
        self.x_step = ttk.Entry(x_frame, width=8)
        self.x_step.insert(0, "5")
        self.x_step.grid(row=1, column=1, padx=3, pady=(8,0), sticky=tk.EW)

        ttk.Checkbutton(x_frame, text="Symmetric", variable=self.x_symmetric, command=self._on_x_symmetric_toggle).grid(row=1, column=2, columnspan=2, sticky=tk.W, padx=8, pady=(8,0))

        # Bind X events
        self.x_min.bind('<FocusOut>', self._auto_update_preview); self.x_min.bind('<Return>', self._auto_update_preview)
        self.x_max.bind('<FocusOut>', self._auto_update_preview); self.x_max.bind('<Return>', self._auto_update_preview)
        self.x_step.bind('<FocusOut>', self._auto_update_preview); self.x_step.bind('<Return>', self._auto_update_preview)
        self.x_offset.bind('<FocusOut>', self._auto_update_preview); self.x_offset.bind('<Return>', self._auto_update_preview)


        # --- Y AXIS ---
        y_frame = ttk.LabelFrame(scrollable_frame, text="Y Axis (mm)", style='Card.TLabelframe', padding=12)
        y_frame.pack(fill=tk.X, pady=(0, 8), padx=2)
        y_frame.columnconfigure(1, weight=1)
        y_frame.columnconfigure(3, weight=1) # *** FIX: Make Max expand too ***
        
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

        # --- Z AXIS ---
        z_frame = ttk.LabelFrame(scrollable_frame, text="Z Axis (mm)", style='Card.TLabelframe', padding=12)
        z_frame.pack(fill=tk.X, pady=(0, 8), padx=2)
        z_frame.columnconfigure(1, weight=1)
        z_frame.columnconfigure(3, weight=1) # *** FIX: Make Max expand too ***

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

        # --- ROTATION AXIS ---
        rot_frame = ttk.LabelFrame(scrollable_frame, text="Rotation (degrees)", style='Card.TLabelframe', padding=12)
        rot_frame.pack(fill=tk.X, pady=(0, 8), padx=2)
        rot_frame.columnconfigure(1, weight=1)
        rot_frame.columnconfigure(3, weight=1) # *** FIX: Make Max expand too ***

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

        # --- MOVEMENT PARAMETERS ---
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
        
        # --- NEW: EXPORT FORMAT ---
        export_frame = ttk.LabelFrame(scrollable_frame, text="Export Format", style='Card.TLabelframe', padding=12)
        export_frame.pack(fill=tk.X, pady=(0, 8), padx=2)

        ttk.Radiobutton(export_frame, text="G-code (.gcode)", variable=self.export_format, value="gcode", command=self._auto_update_preview).pack(anchor=tk.W)
        ttk.Radiobutton(export_frame, text="CSV Coordinates (.csv)", variable=self.export_format, value="csv", command=self._auto_update_preview).pack(anchor=tk.W)


        # --- ACTION BUTTONS (REMOVED from here) ---

        # --- MODIFIED: Bind after all widgets in scrollable_frame are created
        _bind_mousewheel_recursive(scrollable_frame)

        # Pack canvas and scrollbar
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # --- NEW: Bind the canvas itself too (for empty space) ---
        canvas.bind("<MouseWheel>", _on_mousewheel)
        canvas.bind("<Button-4>", _on_mousewheel)
        canvas.bind("<Button-5>", _on_mousewheel)


    def create_preview_panel(self, parent):
        """Create the non-interactive preview and stats area"""

        # Title
        title = ttk.Label(parent, text="SCAN VOLUME PREVIEW",
                            font=('Rajdhani', 16, 'bold'),
                            foreground=self.COLOR_ACCENT_CYAN, background=self.COLOR_BG)
        title.pack(side=tk.TOP, pady=(0, 10), anchor=tk.W) # *** FIX: Explicitly pack TOP ***

        # *** LAYOUT FIX: Pack stats panel to BOTTOM first ***
        
        # Statistics frame - Fixed height
        stats_frame = ttk.LabelFrame(parent, text="Scan Statistics", style='Card.TLabelframe', padding=0)
        stats_frame.pack(fill=tk.X, side=tk.BOTTOM, pady=(10, 0)) # Pack at bottom
        
        # --- MODIFIED: REMOVED stats_frame.pack_propagate(False) ---
        # This was causing the frame to collapse.

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
        # --- MODIFIED: Changed fill=tk.X to fill=tk.BOTH ---
        self.stats_text.pack(fill=tk.BOTH, expand=True)

        # --- NEW: Configure text tags for stats ---
        self.stats_text.tag_configure('header', foreground=self.COLOR_ACCENT_CYAN, font=self.FONT_MONO_LARGE)
        self.stats_text.tag_configure('value', foreground=self.COLOR_ACCENT_GREEN, font=self.FONT_MONO_LARGE)
        self.stats_text.tag_configure('warning', foreground=self.COLOR_ACCENT_RED, font=(self.FONT_MONO[0], self.FONT_MONO[1], 'bold'))
        # --- NEW: Amber warning tag ---
        self.stats_text.tag_configure('amber_warning', foreground=self.COLOR_ACCENT_AMBER, font=(self.FONT_MONO[0], self.FONT_MONO[1], 'bold'))
        self.stats_text.tag_configure('success', foreground=self.COLOR_ACCENT_GREEN, font=(self.FONT_MONO[0], self.FONT_MONO[1], 'bold'))
        self.stats_text.tag_configure('label', foreground=self.COLOR_TEXT_SECONDARY, font=self.FONT_MONO)
        self.stats_text.tag_configure('separator', foreground=self.COLOR_BORDER, font=self.FONT_MONO)

        # 2D Preview Canvas (Pack this second so it fills remaining space)
        canvas_frame = ttk.Frame(parent, style='Dark.TFrame')
        canvas_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=(0, 0)) # *** FIX: Explicitly pack TOP ***

        self.preview_canvas = tk.Canvas(canvas_frame, bg=self.COLOR_BLACK, highlightthickness=1,
                                            highlightbackground=self.COLOR_BORDER)
        self.preview_canvas.pack(fill=tk.BOTH, expand=True)
        self.preview_canvas.bind("<Configure>", self._on_canvas_resize)

        # Initial empty state
        # --- MODIFIED: Pass warning_level=0 ---
        self.draw_preview_diagram(None, [], 0)

    # ===== NEW: SYMMETRIC TOGGLE HANDLERS =====
    # *** STYLE UPDATE: Changed from _wrapper to the entry widget itself ***

    def _on_x_symmetric_toggle(self):
        self._toggle_symmetric_widgets(
            self.x_symmetric.get(),
            self.x_min_label, self.x_min,
            self.x_max_label, self.x_max,
            self.x_offset_label, self.x_offset,
            self.x_min, self.x_max, self.x_offset, "-50", "50"
        )
    
    def _on_y_symmetric_toggle(self):
        self._toggle_symmetric_widgets(
            self.y_symmetric.get(),
            self.y_min_label, self.y_min,
            self.y_max_label, self.y_max,
            self.y_offset_label, self.y_offset,
            self.y_min, self.y_max, self.y_offset, "-50", "50"
        )
        
    def _on_z_symmetric_toggle(self):
        self._toggle_symmetric_widgets(
            self.z_symmetric.get(),
            self.z_min_label, self.z_min,
            self.z_max_label, self.z_max,
            self.z_offset_label, self.z_offset,
            self.z_min, self.z_max, self.z_offset, "0", "100"
        )
        
    def _on_rot_symmetric_toggle(self):
        self._toggle_symmetric_widgets(
            self.rot_symmetric.get(),
            self.rot_min_label, self.rot_min,
            self.rot_max_label, self.rot_max,
            self.rot_offset_label, self.rot_offset,
            self.rot_min, self.rot_max, self.rot_offset, "0", "0"
        )

    def _toggle_symmetric_widgets(self, is_symmetric, 
                                    min_lbl, min_entry_widget, max_lbl, max_entry_widget,
                                    off_lbl, off_entry_widget,
                                    min_entry, max_entry, off_entry,
                                    default_min, default_max):
        """Generic helper to toggle UI for symmetric mode"""
        
        if is_symmetric:
            # SYMMETRIC MODE: Hide min/max, show offset
            try:
                current_min = float(min_entry.get())
                current_max = float(max_entry.get())
                # Offset is the max of the absolute values
                offset = max(abs(current_min), abs(current_max))
                off_entry.delete(0, tk.END)
                off_entry.insert(0, f"{offset:g}")
            except ValueError:
                # Use max of default values
                try:
                    offset = max(abs(float(default_min)), abs(float(default_max)))
                    off_entry.delete(0, tk.END)
                    off_entry.insert(0, f"{offset:g}")
                except ValueError:
                    off_entry.insert(0, "50") # Fallback
            
            min_lbl.grid_remove(); min_entry_widget.grid_remove()
            max_lbl.grid_remove(); max_entry_widget.grid_remove()
            off_lbl.grid(); off_entry_widget.grid()
            
        else:
            # ASYMMETRIC MODE: Show min/max, hide offset
            try:
                offset = abs(float(off_entry.get()))
                min_entry.delete(0, tk.END)
                min_entry.insert(0, f"{-offset:g}")
                max_entry.delete(0, tk.END)
                max_entry.insert(0, f"{offset:g}")
            except ValueError:
                # On error, restore defaults
                min_entry.delete(0, tk.END); min_entry.insert(0, default_min)
                max_entry.delete(0, tk.END); max_entry.insert(0, default_max)
            
            min_lbl.grid(); min_entry_widget.grid()
            max_lbl.grid(); max_entry_widget.grid()
            off_lbl.grid_remove(); off_entry_widget.grid_remove()
        
        # Update preview
        self._auto_update_preview()
        
    # ===== END SYMMETRIC HANDLERS =====


    # --- MODIFIED: Added warning_level parameter ---
    def draw_preview_diagram(self, params, bounds_warnings, warning_level=0):
        """
        (MODIFIED) Draws a transparent wireframe sketch with
        custom colors, origin marker, and printer boundary warnings.
        """
        self.preview_canvas.delete("all")

        canvas_w = self.preview_canvas.winfo_width()
        canvas_h = self.preview_canvas.winfo_height()

        if canvas_w <= 1 or canvas_h <= 1:
            # --- MODIFIED: Pass warning_level to recursive call ---
            self.preview_canvas.after(50, lambda: self.draw_preview_diagram(params, bounds_warnings, warning_level))
            return

        if params is None:
            self.preview_canvas.create_text(
                canvas_w / 2, canvas_h / 2,
                text="Waiting for valid parameters...",
                fill=self.COLOR_TEXT_SECONDARY,
                anchor=tk.CENTER, font=self.FONT_BODY
            )
            # --- MODIFIED: Draw printer bounds even with no params ---
            params = {'x_min': 0, 'x_max': 0, 'y_min': 0, 'y_max': 0, 'z_min': 0, 'z_max': 0}
            # Fall through to draw the printer box
        
        # Get ranges
        x_range = params['x_max'] - params['x_min']
        y_range = params['y_max'] - params['y_min']
        z_range = params['z_max'] - params['z_min']

        # Handle zero ranges to prevent division by zero
        effective_x_range = x_range if x_range != 0 else 1
        effective_y_range = y_range if y_range != 0 else 1
        effective_z_range = z_range if z_range != 0 else 1

        # --- Calculate Scaling ---
        pad = 40
        oblique_factor = 0.4

        # NEW: Include printer limits in scaling calculation
        # --- MODIFIED: Removed 'if bounds_warnings:' ---
        total_min_x = min(params['x_min'], -PRINTER_LIMITS['x'])
        total_max_x = max(params['x_max'], PRINTER_LIMITS['x'])
        total_min_y = min(params['y_min'], -PRINTER_LIMITS['y'])
        total_max_y = max(params['y_max'], PRINTER_LIMITS['y'])
        total_min_z = min(params['z_min'], PRINTER_LIMITS['z_min'])
        total_max_z = max(params['z_max'], PRINTER_LIMITS['z_max'])


        # Total ranges for scaling
        total_x_rng = total_max_x - total_min_x
        total_y_rng = total_max_y - total_min_y
        total_z_rng = total_max_z - total_min_z
        
        eff_total_x_rng = total_x_rng if total_x_rng != 0 else 1
        eff_total_y_rng = total_y_rng if total_y_rng != 0 else 1
        eff_total_z_rng = total_z_rng if total_z_rng != 0 else 1

        total_w_units = eff_total_x_rng + (eff_total_y_rng * oblique_factor)
        total_h_units = eff_total_z_rng + (eff_total_y_rng * oblique_factor)

        if total_w_units == 0: total_w_units = 1
        if total_h_units == 0: total_h_units = 1

        scale_x = (canvas_w - 2 * pad) / total_w_units
        scale_y = (canvas_h - 2 * pad) / total_h_units
        scale = min(scale_x, scale_y)
        if scale < 0: scale = 0

        # --- Define 3D -> 2D Projection Helper ---
        # *** FIX: Simplified redundant logic ***
        def project(x, y, z):
            x_pct = (x - total_min_x) / eff_total_x_rng if eff_total_x_rng != 0 else 0.5
            y_pct = (y - total_min_y) / eff_total_y_rng if eff_total_y_rng != 0 else 0.5
            z_pct = (z - total_min_z) / eff_total_z_rng if eff_total_z_rng != 0 else 0.5
            
            scaled_drawing_w = total_x_rng * scale
            scaled_drawing_h = total_z_rng * scale
            scaled_drawing_d = total_y_rng * scale * oblique_factor

            # Centering offset
            x_start = (canvas_w - (scaled_drawing_w + scaled_drawing_d)) / 2
            y_start = (canvas_h - (scaled_drawing_h + scaled_drawing_d)) / 2

            screen_w = x_pct * scaled_drawing_w
            screen_h = (1 - z_pct) * scaled_drawing_h # Invert Z pct for canvas Y
            screen_d = y_pct * scaled_drawing_d
            
            final_x = x_start + screen_w + screen_d
            final_y = y_start + screen_h + screen_d
            
            return (final_x, final_y)

        # --- Define styles ---
        visible_style = {'fill': self.COLOR_ACCENT_CYAN, 'width': 2}
        # *** FIX: Use a dashed cyan for hidden lines ***
        hidden_style = {'fill': self.COLOR_ACCENT_CYAN, 'dash': (2, 4), 'width': 1}
        # *** BUG FIX: Changed 'outline' to 'fill' ***
        warning_style = {'fill': self.COLOR_ACCENT_RED, 'dash': (4, 4), 'width': 2}

        # --- NEW: Draw Printer Bounds ---
        # --- MODIFIED: Removed 'if bounds_warnings:' ---
        pl = PRINTER_LIMITS # short alias
        # Project 8 corners of the printer box
        pb1 = project(-pl['x'], -pl['y'], pl['z_min']) # Front-bottom-left
        pb2 = project( pl['x'], -pl['y'], pl['z_min']) # Front-bottom-right        
        pb3 = project( pl['x'], -pl['y'], pl['z_max']) # Front-top-right
        pb4 = project(-pl['x'], -pl['y'], pl['z_max']) # Front-top-left
        pb5 = project(-pl['x'],  pl['y'], pl['z_min']) # Back-bottom-left
        pb6 = project( pl['x'],  pl['y'], pl['z_min']) # Back-bottom-right
        pb7 = project( pl['x'],  pl['y'], pl['z_max']) # Back-top-right
        pb8 = project(-pl['x'],  pl['y'], pl['z_max']) # Back-top-left
        
        # Draw all 12 edges of printer box
        self.preview_canvas.create_line(pb1, pb2, **warning_style) # front-bottom
        self.preview_canvas.create_line(pb2, pb3, **warning_style) # front-right
        self.preview_canvas.create_line(pb3, pb4, **warning_style) # front-top
        self.preview_canvas.create_line(pb4, pb1, **warning_style) # front-left
        
        self.preview_canvas.create_line(pb5, pb6, **warning_style) # back-bottom
        self.preview_canvas.create_line(pb6, pb7, **warning_style) # back-right
        self.preview_canvas.create_line(pb7, pb8, **warning_style) # back-top
        self.preview_canvas.create_line(pb8, pb5, **warning_style) # back-left
        
        self.preview_canvas.create_line(pb1, pb5, **warning_style) # connect-bottom-left
        self.preview_canvas.create_line(pb2, pb6, **warning_style) # connect-bottom-right
        self.preview_canvas.create_line(pb3, pb7, **warning_style) # connect-top-right
        self.preview_canvas.create_line(pb4, pb8, **warning_style) # connect-top-left
        
        # Label
        self.preview_canvas.create_text((pb4[0] + pb8[0])/2, pb8[1] - 5,
            text="Printer Limits", fill=self.COLOR_ACCENT_RED, font=self.FONT_BODY_SMALL, anchor=tk.S)


        # --- Draw Pattern Wireframe Edges (only if ranges are not zero) ---
        if x_range != 0 or y_range != 0 or z_range != 0:
            # Project 8 corners of the *pattern* box
            p1 = project(params['x_min'], params['y_min'], params['z_min']) # Front-bottom-left
            p2 = project(params['x_max'], params['y_min'], params['z_min']) # Front-bottom-right
            p3 = project(params['x_max'], params['y_min'], params['z_max']) # Front-top-right
            p4 = project(params['x_min'], params['y_min'], params['z_max']) # Front-top-left
            p5 = project(params['x_min'], params['y_max'], params['z_min']) # Back-bottom-left
            p6 = project(params['x_max'], params['y_max'], params['z_min']) # Back-bottom-right
            p7 = project(params['x_max'], params['y_max'], params['z_max']) # Back-top-right
            p8 = project(params['x_min'], params['y_max'], params['z_max']) # Back-top-left
            
            # Hidden edges
            self.preview_canvas.create_line(p2, p6, **visible_style)
            self.preview_canvas.create_line(p1, p2, **visible_style)
            self.preview_canvas.create_line(p2, p3, **visible_style)
            self.preview_canvas.create_line(p3, p4, **visible_style)
            self.preview_canvas.create_line(p4, p1, **visible_style)

            # Visible edges
            self.preview_canvas.create_line(p3, p7, **visible_style)
            self.preview_canvas.create_line(p4, p8, **visible_style)
            self.preview_canvas.create_line(p7, p8, **visible_style)
            self.preview_canvas.create_line(p6, p7, **visible_style)
            self.preview_canvas.create_line(p5, p6, **visible_style)
            self.preview_canvas.create_line(p8, p5, **visible_style)
            self.preview_canvas.create_line(p1, p5, **visible_style)


        # --- Draw Origin Marker ---
        origin_in_bounds = (total_min_x <= 0 <= total_max_x and
                            total_min_y <= 0 <= total_max_y and
                            total_min_z <= 0 <= total_max_z)

        if origin_in_bounds:
            (ox, oy) = project(0, 0, 0)
            # Glowing effect
            self.preview_canvas.create_oval(ox - 8, oy - 8, ox + 8, oy + 8,
                                            fill='', outline='#ff6666', width=1)
            self.preview_canvas.create_oval(ox - 5, oy - 5, ox + 5, oy + 5,
                                            fill=self.COLOR_ACCENT_RED, outline=self.COLOR_ACCENT_RED, width=2)


        # --- NEW: Draw Warning Text ---
        if warning_level == 2: # Red
            text_color = self.COLOR_ACCENT_RED
            warn_text = "⚠️ BOUNDS EXCEEDED"
        elif warning_level == 1: # Amber
            text_color = self.COLOR_ACCENT_AMBER
            warn_text = "⚠️ PROXIMITY WARNING"
        else:
            warn_text = None
        
        if warn_text:
            # Draw the main warning text (e.g., "BOUNDS EXCEEDED")
            self.preview_canvas.create_text(canvas_w - 20, 40,
                text=warn_text,
                fill=text_color,
                font=self.FONT_MONO_LARGE,
                anchor=tk.NE) # Top-right corner

            # --- NEW: Parse and draw the specific axes ---
            problem_axes = set()
            for warning_string in bounds_warnings:
                if warning_string.startswith("X"):
                    problem_axes.add("X-AXIS")
                elif warning_string.startswith("Y"):
                    problem_axes.add("Y X-AXIS")
                elif warning_string.startswith("Z"):
                    problem_axes.add("Z-AXIS")
            
            # Format the axes string, e.g., "X / Y / Z"
            axes_text = " / ".join(sorted(list(problem_axes)))
            
            if axes_text:
                # Draw the axes text just below the main warning
                self.preview_canvas.create_text(canvas_w - 20, 65, # Positioned below (40 + 25)
                    text=axes_text,
                    fill=text_color, # Use the same color
                    font=self.FONT_MONO_LARGE, # Use the same font
                    anchor=tk.NE)
            # --- END NEW ---
            
        # --- End Warning Text ---

    def update_filename_preview(self, event=None):
        """Update the filename preview label"""
        name = self.profile_name.get()
        name = "".join(c for c in name if c.isalnum() or c in ('-', '_'))

        if self.include_timestamp.get():
            timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            filename = f"{name}_{timestamp}"
        else:
            filename = f"{name}"
        
        # NEW: Update based on export format
        format = self.export_format.get()
        if format == 'csv':
            filename += ".csv"
        else:
            filename += ".gcode"

        self.filename_preview.config(text=filename)

    def _get_params_silently(self):
        """ Read parameters, return None if invalid, no popups. """
        params = {}
        try:
            # --- MODIFIED: Read from offset or min/max ---
            # X Axis
            if self.x_symmetric.get():
                offset = abs(float(self.x_offset.get()))
                params['x_min'], params['x_max'] = -offset, offset
            else:
                params['x_min'] = float(self.x_min.get()); params['x_max'] = float(self.x_max.get())
            params['x_step'] = float(self.x_step.get())
            
            # Y Axis
            if self.y_symmetric.get():
                offset = abs(float(self.y_offset.get()))
                params['y_min'], params['y_max'] = -offset, offset
            else:
                params['y_min'] = float(self.y_min.get()); params['y_max'] = float(self.y_max.get())
            params['y_step'] = float(self.y_step.get())
            
            # Z Axis
            if self.z_symmetric.get():
                offset = abs(float(self.z_offset.get()))
                params['z_min'], params['z_max'] = -offset, offset
            else:
                params['z_min'] = float(self.z_min.get()); params['z_max'] = float(self.z_max.get())
            params['z_step'] = float(self.z_step.get())
            
            # Rotation Axis
            if self.rot_symmetric.get():
                offset = abs(float(self.rot_offset.get()))
                params['rot_min'], params['rot_max'] = -offset, offset
            else:
                params['rot_min'] = float(self.rot_min.get()); params['rot_max'] = float(self.rot_max.get())
            params['rot_step'] = float(self.rot_step.get())
            
            # Movement
            params['travelspeed'] = float(self.travelspeed.get())
            params['pause_time'] = float(self.pause_time.get())
            # --- END MODIFIED READ ---

            if params['x_min'] > params['x_max']: return None
            if params['y_min'] > params['y_max']: return None
            if params['z_min'] > params['z_max']: return None
            if params['rot_min'] > params['rot_max']: return None
            if params['x_step'] < 0 or params['y_step'] < 0 or params['z_step'] < 0 or params['rot_step'] < 0: return None
            if params['x_min'] != params['x_max'] and params['x_step'] == 0: return None
            if params['y_min'] != params['y_max'] and params['y_step'] == 0: return None
            if params['z_min'] != params['z_max'] and params['z_step'] == 0: return None
            if params['rot_min'] != params['rot_max'] and params['rot_step'] == 0: return None
            if params['travelspeed'] <= 0: return None
            if params['pause_time'] < 0: return None
            return params
        except ValueError:
            return None

    def get_parameters(self):
        """ Read parameters, show error popups if invalid. """
        params = {}
        try:
            # --- MODIFIED: Read from offset or min/max ---
            # X Axis
            if self.x_symmetric.get():
                offset = abs(float(self.x_offset.get()))
                params['x_min'], params['x_max'] = -offset, offset
            else:
                params['x_min'] = float(self.x_min.get()); params['x_max'] = float(self.x_max.get())
            params['x_step'] = float(self.x_step.get())
            
            # Y Axis
            if self.y_symmetric.get():
                offset = abs(float(self.y_offset.get()))
                params['y_min'], params['y_max'] = -offset, offset
            else:
                params['y_min'] = float(self.y_min.get()); params['y_max'] = float(self.y_max.get())
            params['y_step'] = float(self.y_step.get())
            
            # Z Axis
            if self.z_symmetric.get():
                offset = abs(float(self.z_offset.get()))
                params['z_min'], params['z_max'] = -offset, offset
            else:
                params['z_min'] = float(self.z_min.get()); params['z_max'] = float(self.z_max.get())
            params['z_step'] = float(self.z_step.get())
            
            # Rotation Axis
            if self.rot_symmetric.get():
                offset = abs(float(self.rot_offset.get()))
                params['rot_min'], params['rot_max'] = -offset, offset
            else:
                params['rot_min'] = float(self.rot_min.get()); params['rot_max'] = float(self.rot_max.get())
            params['rot_step'] = float(self.rot_step.get())
            
            # Movement
            params['travelspeed'] = float(self.travelspeed.get())
            params['pause_time'] = float(self.pause_time.get())
            # --- END MODIFIED READ ---

            if params['x_min'] > params['x_max']: messagebox.showerror("Error", "X Min <= X Max"); return None
            if params['y_min'] > params['y_max']: messagebox.showerror("Error", "Y Min <= Y Max"); return None
            if params['z_min'] > params['z_max']: messagebox.showerror("Error", "Z Min <= Z Max"); return None
            if params['rot_min'] > params['rot_max']: messagebox.showerror("Error", "Rot Min <= Rot Max"); return None
            if params['x_step'] < 0 or params['y_step'] < 0 or params['z_step'] < 0 or params['rot_step'] < 0: messagebox.showerror("Error", "Steps >= 0"); return None
            if params['x_min'] != params['x_max'] and params['x_step'] == 0: messagebox.showerror("Error", "X Step > 0 if X range > 0"); return None
            if params['y_min'] != params['y_max'] and params['y_step'] == 0: messagebox.showerror("Error", "Y Step > 0 if Y range > 0"); return None
            if params['z_min'] != params['z_max'] and params['z_step'] == 0: messagebox.showerror("Error", "Z Step > 0 if Z range > 0"); return None
            if params['rot_min'] != params['rot_max'] and params['rot_step'] == 0: messagebox.showerror("Error", "Rot Step > 0 if Rot range > 0"); return None
            if params['travelspeed'] <= 0: messagebox.showerror("Error", "Travel Speed > 0"); return None
            if params['pause_time'] < 0: messagebox.showerror("Error", "Pause Time >= 0"); return None
            return params
        except ValueError:
            messagebox.showerror("Error", "All fields must be valid numbers"); return None

    def generate_step_values(self, min_val, max_val, step):
        """Generate list of values from min to max with given step"""
        if step == 0: return [min_val] if min_val == max_val else [min_val] # Return min_val even if not equal, for 0-step
        values = []
        current = min_val
        while current <= max_val + 1e-9:  # Epsilon for float comparison
            values.append(round(current, 6))
            current += step
        return values

    def _calculate_total_points(self, params):
        """Helper function to calculate total points without generating them."""
        def count_steps(min_val, max_val, step):
            if min_val > max_val: return 0
            if step == 0: return 1 # A single point
            num_steps = math.floor((max_val - min_val) / step + 1e-9)
            return int(num_steps) + 1
        try:
            nx = count_steps(params['x_min'], params['x_max'], params['x_step'])
            ny = count_steps(params['y_min'], params['y_max'], params['y_step'])
            nz = count_steps(params['z_min'], params['z_max'], params['z_step'])
            n_rot = count_steps(params['rot_min'], params['rot_max'], params['rot_step'])
            return nx * ny * nz * n_rot
        except Exception:
            return 0

    def create_pattern(self, params):
        """
        Creates the scan pattern as a GENERATOR using an efficient
        boustrophedon (snake-like) path for the X-axis.
        """
        x_values = self.generate_step_values(params['x_min'], params['x_max'], params['x_step'])
        y_values = self.generate_step_values(params['y_min'], params['y_max'], params['y_step'])
        z_values = self.generate_step_values(params['z_min'], params['z_max'], params['z_step'])
        rot_values = self.generate_step_values(params['rot_min'], params['rot_max'], params['rot_step'])

        # Optimize for 0-step case
        if not x_values: x_values = [params['x_min']]
        if not y_values: y_values = [params['y_min']]
        if not z_values: z_values = [params['z_min']]
        if not rot_values: rot_values = [params['rot_min']]

        for rot in rot_values:
            for z in z_values:
                # For each Z-plane, we alternate the X-direction for each line in Y
                x_direction_forward = True
                for y in y_values:
                    # Use regular or reversed X values based on the alternating direction
                    x_iterator = x_values if x_direction_forward else reversed(x_values)

                    for x in x_iterator:
                        yield {'x': x, 'y': y, 'z': z, 'rotation': rot}

                    # Flip the direction for the next Y line
                    x_direction_forward = not x_direction_forward

    def _format_time(self, total_seconds):
        """Formats seconds into d h m s string."""
        if total_seconds < 0: return "0s"
        total_seconds = int(total_seconds)
        d, rem = divmod(total_seconds, 86400); h, rem = divmod(rem, 3600); m, s = divmod(rem, 60)
        parts = []
        if d > 0: parts.append(f"{d}d")
        if h > 0: parts.append(f"{h}h")
        if m > 0: parts.append(f"{m}m")
        if s > 0 or not parts: parts.append(f"{s}s")
        return " ".join(parts)

    def _calculate_estimated_time(self, params, total_points):
        """ Calculates estimated time with high accuracy, including inter-rotation moves. """
        if total_points == 0 or params is None: return 0
        total_pause_s = max(0, (total_points - 1) * params['pause_time'])
        travelspeed_mms = params['travelspeed'] / 60.0
        if travelspeed_mms <= 0: return total_pause_s
        def count_steps(min_v, max_v, stp):
            if min_v > max_v: return 0
            if stp == 0: return 1
            return int(math.floor((max_v - min_v) / stp + 1e-9)) + 1
        nx, ny, nz, n_rot = [count_steps(params[f'{ax}_min'], params[f'{ax}_max'], params[f'{ax}_step']) for ax in ['x', 'y', 'z', 'rot']]
        total_dist = 0
        xr, yr, zr = params['x_max'] - params['x_min'], params['y_max'] - params['y_min'], params['z_max'] - params['z_min']
        if nx > 1: total_dist += (nx - 1) * ny * nz * n_rot * params['x_step']
        if ny > 1: total_dist += (ny - 1) * nz * n_rot * math.sqrt(xr ** 2 + params['y_step'] ** 2)
        if nz > 1: total_dist += (nz - 1) * n_rot * math.sqrt(xr ** 2 + yr ** 2 + params['z_step'] ** 2)
        if n_rot > 1: total_dist += (n_rot - 1) * math.sqrt(xr ** 2 + yr ** 2 + zr ** 2)
        total_travel_s = total_dist / travelspeed_mms if travelspeed_mms > 0 else 0
        return total_travel_s + total_pause_s

        # --- MODIFIED: bounds_warnings is now just the list of strings ---
    def update_statistics(self, params, total_points, bounds_warnings, warning_level=0):        
        """
        (MODIFIED) Update statistics text display using
        the new custom text widget and tags.
        """
        self.stats_text.config(state=tk.NORMAL)
        self.stats_text.delete(1.0, tk.END)
        
        if params is None: 
            self.stats_text.insert(1.0, "Waiting for valid parameters...", 'warning')
            self.stats_text.config(state=tk.DISABLED)
            return

        xr, yr, zr = params['x_max'] - params['x_min'], params['y_max'] - params['y_min'], params['z_max'] - params['z_min']
        def count_steps(min_v, max_v, stp):
            if min_v > max_v: return 0
            if stp == 0: return 1
            return int(math.floor((max_v - min_v) / stp + 1e-9)) + 1
        nx, ny, nz, n_rot = [count_steps(params[f'{ax}_min'], params[f'{ax}_max'], params[f'{ax}_step']) for ax in ['x', 'y', 'z', 'rot']]
        est_secs = self._calculate_estimated_time(params, total_points); time_str = self._format_time(est_secs)

        sep = "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"

        self.stats_text.insert(tk.END, sep, 'separator')
        
        self.stats_text.insert(tk.END, "Total Points:   ", 'label')
        self.stats_text.insert(tk.END, f"{total_points:,}\n", 'value')
        
        self.stats_text.insert(tk.END, "Grid (X,Y,Z,R): ", 'label')
        self.stats_text.insert(tk.END, f"{nx}×{ny}×{nz}×{n_rot}\n", 'value')

        self.stats_text.insert(tk.END, "Volume:          ", 'label')
        self.stats_text.insert(tk.END, f"{xr:.1f}×{yr:.1f}×{zr:.1f} mm³\n", 'value')

        self.stats_text.insert(tk.END, "Est. Runtime:   ", 'label')
        self.stats_text.insert(tk.END, f"{time_str}\n", 'value')
        
        self.stats_text.insert(tk.END, sep, 'separator')

        # Add warnings or success
        if warning_level == 2: # Red warning (Exceeded)
            self.stats_text.insert(tk.END, "⚠️ BOUNDS EXCEEDED!\n", 'warning')
            tag = 'warning'
        elif warning_level == 1: # Amber warning (Proximity)
            self.stats_text.insert(tk.END, "⚠️ PROXIMITY WARNING\n", 'amber_warning')
            tag = 'amber_warning'
        else: # warning_level == 0 (OK)
            self.stats_text.insert(tk.END, "✓ Pattern fits printer bounds\n", 'success')
            tag = 'success' # Set a default tag

        # Now, if there are warnings, list them under the header.
        if bounds_warnings:
            for warning in bounds_warnings:
                self.stats_text.insert(tk.END, f"  {warning}\n", tag)

        self.stats_text.config(state=tk.DISABLED)

    # --- MODIFIED: Function now returns (warning_list, warning_level) ---
    # warning_level: 0=OK, 1=Amber (proximity), 2=Red (exceeded)
    def _check_printer_bounds(self, params):
        """
        NEW: Checks pattern parameters against printer limits.
        Returns a (list of warning strings, warning_level).
        """
        if params is None:
            return ([], 0)
            
        warnings = []
        warning_level = 0 # 0=OK, 1=Amber, 2=Red
        pl = PRINTER_LIMITS
        proximity = 10.0 # 10mm proximity warning
        
        # Check X
        pattern_x_extent = max(abs(params['x_min']), abs(params['x_max']))
        if pattern_x_extent > pl['x']:
            warnings.append(f"X extent ({pattern_x_extent:.1f}mm) > limit ({pl['x']:.1f}mm)")
            warning_level = 2 # Red
        elif pattern_x_extent > pl['x'] - proximity:
             warnings.append(f"X extent ({pattern_x_extent:.1f}mm) near limit ({pl['x']:.1f}mm)")
             warning_level = max(warning_level, 1) # Amber
            
        # Check Y
        pattern_y_extent = max(abs(params['y_min']), abs(params['y_max']))
        if pattern_y_extent > pl['y']:
            warnings.append(f"Y extent ({pattern_y_extent:.1f}mm) > limit ({pl['y']:.1f}mm)")
            warning_level = 2 # Red
        elif pattern_y_extent > pl['y'] - proximity:
             warnings.append(f"Y extent ({pattern_y_extent:.1f}mm) near limit ({pl['y']:.1f}mm)")
             warning_level = max(warning_level, 1) # Amber
            
        # Check Z (Max)
        if params['z_max'] > pl['z_max']:
            warnings.append(f"Z max ({params['z_max']:.1f}mm) > limit ({pl['z_max']:.1f}mm)")
            warning_level = 2 # Red
        elif params['z_max'] > pl['z_max'] - proximity:
            warnings.append(f"Z max ({params['z_max']:.1f}mm) near limit ({pl['z_max']:.1f}mm)")
            warning_level = max(warning_level, 1) # Amber

        # Check Z (Min)
        if params['z_min'] < pl['z_min']:
             warnings.append(f"Z min ({params['z_min']:.1f}mm) < limit ({pl['z_min']:.1f}mm)")
             warning_level = 2 # Red
        # Proximity check for Z-min intentionally removed.
             
        return (warnings, warning_level)

    def _auto_update_preview(self, event=None):
        """ Called on <FocusOut> or <Return> to update stats/diagram. """
        params = self._get_params_silently()
        total_points = self._calculate_total_points(params) if params else 0
        
        # --- MODIFIED: Unpack new return values ---
        bounds_warnings, warning_level = self._check_printer_bounds(params)
        
        # --- MODIFIED: Pass new args ---
        self.update_statistics(params, total_points, bounds_warnings, warning_level)
        self.draw_preview_diagram(params, bounds_warnings, warning_level)
                
        # NEW: Update filename preview on any change
        self.update_filename_preview()

    # ===== NEW: FILE GENERATION ROUTER =====
    
    def _start_generation_process(self):
        """
        NEW: Called by the "Generate File" button.
        Checks export format and calls the correct method.
        """
        params = self.get_parameters()
        if params is None: return
        
        total_points = self._calculate_total_points(params)
        if total_points == 0: 
            messagebox.showerror("Error", "Pattern has 0 points."); return
        
        if total_points > 1_000_000:
            if not messagebox.askokcancel("Warning: Large File", f"This pattern contains {total_points:,} points.\nGenerating this file may take some time.\n\nContinue?"): 
                return
        
        # Get base filename
        name = self.profile_name.get()
        name = "".join(c for c in name if c.isalnum() or c in ('-', '_'))
        if not name: 
            messagebox.showerror("Error", "Please enter a valid profile name."); return
        
        ts = datetime.now().strftime("_%Y-%m-%d_%H-%M-%S") if self.include_timestamp.get() else ""
        
        
        # Route based on format
        format_choice = self.export_format.get()
        
        if format_choice == "gcode":
            def_fname = f"{name}{ts}.gcode"
            fname = filedialog.asksaveasfilename(
                title="Save G-code File",
                defaultextension=".gcode", 
                initialfile=def_fname, 
                filetypes=[("G-code", "*.gcode"), ("All Files", "*.*")]
            )
            if not fname: return
            self._generate_gcode_file(params, total_points, fname)
            
        elif format_choice == "csv":
            def_fname = f"{name}{ts}.csv"
            fname = filedialog.asksaveasfilename(
                title="Save CSV File",
                defaultextension=".csv", 
                initialfile=def_fname, 
                filetypes=[("CSV (Comma-separated)", "*.csv"), ("All Files", "*.*")]
            )
            if not fname: return
            self._generate_csv_file(params, total_points, fname)

    def _generate_gcode_file(self, params, total_points, fname):
        """Generate G-code file using the generator"""
        try:
            pattern_gen = self.create_pattern(params)
            gcode_gen = self.create_gcode(pattern_gen, params, total_points)
            
            with open(fname, 'w') as f:
                for line in gcode_gen: 
                    f.write(line + "\n")
                    
            messagebox.showinfo("Success", f"G-code generated!\nFile: {fname}\nPositions: {total_points:,}")
        except Exception as e: 
            messagebox.showerror("Error", f"Failed to generate G-code:\n{e}")

    def _generate_csv_file(self, params, total_points, fname):
        """NEW: Generate CSV file using the generator"""
        try:
            pattern_gen = self.create_pattern(params)
            csv_gen = self.create_csv_data(pattern_gen, params, total_points)
            
            with open(fname, 'w') as f:
                for line in csv_gen: 
                    f.write(line + "\n")
                    
            messagebox.showinfo("Success", f"CSV exported!\nFile: {fname}\nPoints: {total_points:,}")
        except Exception as e: 
            messagebox.showerror("Error", f"Failed to export CSV:\n{e}")
    
    # ===== END FILE GENERATION =====

    # ===== PROFILE HANDLING METHODS =====

    def _get_profile_data(self):
        """
        (MODIFIED) Gathers all UI settings, including symmetric state and offset.
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
            
            'export_format': self.export_format.get() # Save export choice
        }
        
        # Save min/max or offset based on symmetric state
        if self.x_symmetric.get(): profile_data['x_offset'] = self.x_offset.get()
        else: profile_data['x_min'] = self.x_min.get(); profile_data['x_max'] = self.x_max.get()
            
        if self.y_symmetric.get(): profile_data['y_offset'] = self.y_offset.get()
        else: profile_data['y_min'] = self.y_min.get(); profile_data['y_max'] = self.y_max.get()
            
        if self.z_symmetric.get(): profile_data['z_offset'] = self.z_offset.get()
        else: profile_data['z_min'] = self.z_min.get(); profile_data['z_max'] = self.z_max.get()
        
        if self.rot_symmetric.get(): profile_data['rot_offset'] = self.rot_offset.get()
        else: profile_data['rot_min'] = self.rot_min.get(); profile_data['rot_max'] = self.rot_max.get()
            
        return profile_data

    def load_profile(self):
        """
        (MODIFIED) Loads parameters, including new symmetric offset values.
        """
        fname = filedialog.askopenfilename(
            title="Load Profile from G-code File",
            filetypes=[("G-code Files", "*.gcode"), ("All Files", "*.*")]
        )
        if not fname:
            return

        MAGIC_PREFIX = "; PROFILE_JSON: "
        profile_data = None

        try:
            with open(fname, 'r') as f:
                for line in f:
                    if line.startswith(MAGIC_PREFIX):
                        json_string = line[len(MAGIC_PREFIX):]
                        profile_data = json.loads(json_string)
                        break

            if profile_data is None:
                messagebox.showerror("Error", "No profile data found in this G-code file.")
                return

            # --- Populate the UI from the loaded data ---
            def set_widget(widget, key):
                if key in profile_data:
                    widget.delete(0, tk.END)
                    widget.insert(0, str(profile_data[key]))
            
            def set_var(var, key, default=None):
                if key in profile_data:
                    var.set(profile_data[key])
                elif default is not None:
                    var.set(default)


            set_widget(self.profile_name, 'profile_name')
            set_var(self.include_timestamp, 'include_timestamp', default=True)
            
            set_widget(self.x_step, 'x_step'); set_widget(self.y_step, 'y_step')
            set_widget(self.z_step, 'z_step'); set_widget(self.rot_step, 'rot_step')
            
            set_widget(self.travelspeed, 'travelspeed'); set_widget(self.pause_time, 'pause_time')
            
            set_var(self.export_format, 'export_format', default='gcode')

            # --- Load Symmetric State AND Values ---
            # Must set the checkbox *first*, then the values
            
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

            # --- NEW: Manually trigger UI toggle ---
            # This shows/hides the correct min/max/offset fields
            self._on_x_symmetric_toggle()
            self._on_y_symmetric_toggle()
            self._on_z_symmetric_toggle()
            self._on_rot_symmetric_toggle()

            # Update previews
            self.update_filename_preview()
            self._auto_update_preview()
            messagebox.showinfo("Success", "Profile loaded successfully.")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load profile:\n{e}")

    # ===== GENERATOR METHODS (G-CODE, CSV) =====

    def create_gcode(self, pattern_generator, params, total_points):
        """ (MODIFIED) Creates G-code string as a GENERATOR. """
        profile_data = self._get_profile_data()
        profile_json = json.dumps(profile_data)
        magic_profile_line = f"; PROFILE_JSON: {profile_json}"

        yield f"; Pattern: {self.profile_name.get()}"
        yield f"; Generated: {datetime.now():%Y-%m-%d %H:%M:%S}"
        yield f"; Total points: {total_points:,}"
        yield f"; Speed: {params['travelspeed']} mm/min"
        yield f"; Pause: {params['pause_time']} s"
        yield ";"
        yield magic_profile_line
        yield ";"
        yield "; NOTE: File generated relative to (0,0,0)."
        yield "; Use G-Code Sender to apply absolute center offset."
        yield ";"
        yield ""; yield "; === INIT ==="
        yield "G28 ; Home"
        yield "G90 ; Absolute pos"
        yield "M82 ; Absolute extruder"
        yield ""
        yield "; === PATTERN ==="
        yield f"; {total_points:,} points (relative to center)"
        yield ""
        for i, pos in enumerate(pattern_generator, 1):
            if i % 10000 == 0: yield f"; --- Progress: {i:,}/{total_points:,} ---"
            yield f"; Pos {i}"
            yield f"G1 X{pos['x']:.3f} Y{pos['y']:.3f} Z{pos['z']:.3f} F{params['travelspeed']:.0f}"
            if pos['rotation'] != 0: yield f"; Rotation: {pos['rotation']:.1f} deg (4th axis?)"
            if params['pause_time'] > 0 and i < total_points: yield f"G4 P{int(params['pause_time'] * 1000)} ; Pause"
            yield ""
        yield "; === END ==="
        yield f"G1 X0 Y0 Z{params['z_max']} F3000 ; Move Z up (relative to center)"
        yield "G90 ; Absolute pos"
        yield "M84 ; Disable steppers"
        yield "; Complete!"

    def create_csv_data(self, pattern_generator, params, total_points):
        """ NEW: Creates CSV data as a GENERATOR. """
        # Yield header
        yield "Point,X,Y,Z,Rotation"
        
        # Yield data rows
        for i, pos in enumerate(pattern_generator, 1):
            yield f"{i},{pos['x']:.3f},{pos['y']:.3f},{pos['z']:.3f},{pos['rotation']:.1f}"

    # ===== CANVAS RESIZE HANDLERS =====

    def _on_canvas_resize(self, event):
        """ Debounces the redraw to prevent lag during window resizing. """
        if self._canvas_resize_timer:
            self.root.after_cancel(self._canvas_resize_timer)
        self._canvas_resize_timer = self.root.after(25, self._perform_delayed_redraw)

    def _perform_delayed_redraw(self):
        """ Called by the 'after' timer to run the actual redraw. """
        self._canvas_resize_timer = None
        self._auto_update_preview()


# --- Main Execution ---
if __name__ == "__main__":
    root = tk.Tk()
    
    # --- A simple fix for blurry fonts on Windows ---
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass 
        
    app = PatternGeneratorGUI(root)
    root.mainloop()