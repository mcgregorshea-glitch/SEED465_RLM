import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import serial
import time
import re
import threading
import queue
import serial.tools.list_ports # For port scanning
import math

# SCI-FI: Main Application Class
class GCodeSenderGUI:
    """
    GUI Application to send G-code files to a 3D printer via serial.
    - Translates relative G-code to absolute coordinates.
    - Enforces hardcoded printer bounds.
    - Features "Mark Center" and "Absolute/Relative" coordinate modes.
    - All internal logic is absolute; GUI display can be toggled.
    """
    def __init__(self, root):
        self.root = root
        self.root.title("G-Code Sender")
        self.root.geometry("900x800") # SCI-FI: Increased size
        self.root.minsize(700, 600)  # SCI-FI: Increased minsize

        # --- SCI-FI: Color Palette ---
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
        
        # --- SCI-FI: Fonts ---
        self.FONT_HEADER = ("Orbitron", 13)
        self.FONT_BODY = ("Inter", 11)
        self.FONT_BODY_SMALL = ("Inter", 9)
        self.FONT_BODY_BOLD = ("Inter", 11, "bold")
        self.FONT_BODY_BOLD_LARGE = ("Inter", 20, "bold")
        self.FONT_MONO = ("JetBrains Mono", 10)
        self.FONT_DRO = ("Space Mono", 16, "bold")
        self.FONT_TERMINAL = ("JetBrains Mono", 10)

        # SCI-FI: Set root background
        self.root.configure(bg=self.COLOR_BG)

        # --- SCI-FI: Define Styles ---
        style = ttk.Style()
        style.theme_use('clam') # Use 'clam' as it's more stylable

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
        style.configure('TFrame',
                        background=self.COLOR_BG)
        style.configure('Panel.TFrame',
                        background=self.COLOR_PANEL_BG)
        style.configure('Header.TFrame',
                        background=self.COLOR_PANEL_BG,
                        bordercolor=self.COLOR_BORDER,
                        borderwidth=1,
                        relief='solid')
        style.configure('Footer.TFrame',
                        background=self.COLOR_BLACK,
                        bordercolor=self.COLOR_BORDER,
                        borderwidth=1,
                        relief='solid')
        
        # Style for ttk.Frames that need a black background (like the DRO panel)
        style.configure('Black.TFrame',
                        background=self.COLOR_BLACK)

        # --- LabelFrame (Panel) Style ---
        style.configure('TLabelframe',
                        background=self.COLOR_PANEL_BG,
                        bordercolor=self.COLOR_BORDER,
                        borderwidth=1,
                        relief=tk.SOLID,
                        padding=16)
        style.configure('TLabelframe.Label',
                        background=self.COLOR_PANEL_BG,
                        foreground=self.COLOR_TEXT_SECONDARY,
                        font=("Rajdhani", 13, "bold"),
                        padding=(10, 5))

        # --- Label Styles ---
        style.configure('TLabel',
                        background=self.COLOR_PANEL_BG,
                        foreground=self.COLOR_TEXT_PRIMARY,
                        font=self.FONT_BODY)
        style.configure('Header.TLabel',
                        background=self.COLOR_PANEL_BG,
                        font=self.FONT_BODY)
        style.configure('Footer.TLabel',
                        background=self.COLOR_BLACK,
                        foreground=self.COLOR_TEXT_SECONDARY,
                        font=self.FONT_MONO)
        style.configure('Filepath.TLabel',
                        background=self.COLOR_PANEL_BG,
                        foreground=self.COLOR_TEXT_SECONDARY,
                        font=self.FONT_BODY_SMALL)
        
        # --- DRO Label Styles ---
        style.configure('DRO.TLabel', 
                        font=self.FONT_MONO, 
                        padding=(5, 5), 
                        background=self.COLOR_BLACK,
                        foreground=self.COLOR_TEXT_SECONDARY, 
                        borderwidth=1, 
                        relief='sunken',
                        anchor='w')
        style.configure('Red.DRO.TLabel', 
                        font=self.FONT_DRO, 
                        padding=(5, 5), 
                        background=self.COLOR_BLACK, 
                        foreground=self.COLOR_ACCENT_RED, 
                        borderwidth=0, 
                        relief='flat',
                        anchor='e')
        style.configure('Blue.DRO.TLabel', 
                        font=self.FONT_DRO, 
                        padding=(5, 5), 
                        background=self.COLOR_BLACK, 
                        foreground=self.COLOR_ACCENT_AMBER, 
                        borderwidth=0, 
                        relief='flat',
                        anchor='e')

        # --- Button Styles ---
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

        # --- Primary Action Button (Connect, Start) ---
        style.configure('Primary.TButton',
                        background=self.COLOR_ACCENT_CYAN,
                        foreground=self.COLOR_BLACK,
                        font=self.FONT_BODY_BOLD)
        style.map('Primary.TButton',
                  background=[('active', '#00eaff'), ('pressed', self.COLOR_ACCENT_CYAN)],
                  foreground=[('active', self.COLOR_BLACK), ('pressed', self.COLOR_BLACK)],
                  bordercolor=[('active', self.COLOR_ACCENT_CYAN)])

        # --- Danger Button (STOP) ---
        style.configure('Danger.TButton',
                        background=self.COLOR_ACCENT_RED,
                        foreground=self.COLOR_TEXT_PRIMARY,
                        font=self.FONT_BODY_BOLD)
        style.map('Danger.TButton',
                  background=[('active', '#ff6666'), ('pressed', self.COLOR_ACCENT_RED)],
                  bordercolor=[('active', self.COLOR_ACCENT_RED)])

        # --- Segmented Control Buttons ---
        style.configure('Segment.TButton',
                        background=self.COLOR_PANEL_BG,
                        foreground=self.COLOR_TEXT_SECONDARY,
                        padding=(10, 5),
                        font=self.FONT_BODY_SMALL)
        style.map('Segment.TButton',
                  background=[('active', '#2c333e'), ('pressed', self.COLOR_BLACK)],
                  foreground=[('active', self.COLOR_ACCENT_CYAN)])
        
        style.configure('Segment.Active.TButton',
                        background=self.COLOR_ACCENT_CYAN,
                        foreground=self.COLOR_BLACK,
                        padding=(10, 5),
                        font=self.FONT_BODY_SMALL)
        style.map('Segment.Active.TButton',
                  background=[('active', self.COLOR_ACCENT_CYAN), ('pressed', self.COLOR_ACCENT_CYAN)],
                  foreground=[('active', self.COLOR_BLACK), ('pressed', self.COLOR_BLACK)])

        # --- Jog Buttons ---
        style.configure('Jog.TButton', font=self.FONT_BODY_BOLD, width=5, padding=(10, 10))
        style.configure('Home.TButton', font=self.FONT_BODY_BOLD_LARGE, width=5, padding=(4, 4))


        # --- Entry (Input) Style ---
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
                  bordercolor=[('focus', self.COLOR_ACCENT_CYAN)])

        # --- Combobox (Dropdown) Style ---
        style.configure('TCombobox',
                        fieldbackground=self.COLOR_BLACK,
                        foreground=self.COLOR_ACCENT_CYAN,
                        bordercolor=self.COLOR_BORDER,
                        arrowcolor=self.COLOR_ACCENT_CYAN,
                        background=self.COLOR_BLACK,
                        padding=6,
                        font=self.FONT_MONO)
        style.map('TCombobox',
                  bordercolor=[('focus', self.COLOR_ACCENT_CYAN)])

        # --- NEW: Style the Combobox dropdown list ---
        # This is often OS-dependent and may not work perfectly
        self.root.option_add('*TCombobox*Listbox.background', self.COLOR_BLACK)
        self.root.option_add('*TCombobox*Listbox.foreground', self.COLOR_ACCENT_CYAN)
        self.root.option_add('*TCombobox*Listbox.selectBackground', self.COLOR_ACCENT_CYAN)
        self.root.option_add('*TCombobox*Listbox.selectForeground', self.COLOR_BLACK)
        self.root.option_add('*TCombobox*Listbox.font', self.FONT_MONO)
        self.root.option_add('*TCombobox*Listbox.borderWidth', 0)
        # --- END NEW ---

        # --- Progress Bar Style ---
        style.configure('TProgressbar',
                        troughcolor=self.COLOR_BLACK,
                        background=self.COLOR_ACCENT_CYAN,
                        bordercolor=self.COLOR_BORDER,
                        borderwidth=1,
                        relief=tk.SOLID)

        # --- PanedWindow Sash Style ---
        style.configure('Sash',
                        background=self.COLOR_BG,
                        sashthickness=6,
                        relief=tk.FLAT)
        style.map('Sash',
                  background=[('active', self.COLOR_ACCENT_CYAN)])
        
        # --- NEW: Scrollbar Style ---
        style.configure('TScrollbar',
                        background=self.COLOR_BORDER,           # The slider bar
                        troughcolor=self.COLOR_BG,              # The background trough
                        bordercolor=self.COLOR_BG,
                        arrowcolor=self.COLOR_TEXT_PRIMARY,
                        relief=tk.FLAT,
                        arrowsize=14)
        style.map('TScrollbar',
                  background=[('active', self.COLOR_ACCENT_CYAN), ('!active', self.COLOR_BORDER)],
                  troughcolor=[('active', self.COLOR_BG), ('!active', self.COLOR_BG)])
        # --- END NEW ---


        # --- Core Attributes ---
        self.serial_connection = None
        self.gcode_lines = []; self.processed_gcode = []
        self.is_sending = False; self.is_paused = False; self.is_manual_command_running = False
        self.stop_event = threading.Event(); self.pause_event = threading.Event(); self.pause_event.set(); self.cancel_connect_event = threading.Event()
        self.message_queue = queue.Queue()

        # --- Printer Bounds ---
        self.PRINTER_BOUNDS = { 'x_min': 0, 'x_max': 220, 'y_min': 0, 'y_max': 220, 'z_min': 0, 'z_max': 250 }

        # --- StringVars (VIEW) ---
        self.file_path_var = tk.StringVar(value="No file selected")
        self.center_x_var = tk.StringVar(value="110.0"); self.center_y_var = tk.StringVar(value="110.0"); self.center_z_var = tk.StringVar(value="50.0")
        self.available_ports = ["Auto-detect"] + self._get_available_ports(); self.port_var = tk.StringVar(value=self.available_ports[0] if self.available_ports else ""); self.baud_var = tk.StringVar(value="115200")
        self.connection_status_var = tk.StringVar(value="Status: Disconnected")
        self.jog_step_var = tk.StringVar(value="10"); self.jog_feedrate_var = tk.StringVar(value="1000")
        self.progress_var = tk.DoubleVar(value=0.0); self.progress_label_var = tk.StringVar(value="Progress: Idle"); self.total_lines_to_send = 0
        
        # Display StringVars for Coordinates (View)
        self.goto_x_display_var = tk.StringVar(value="0.00")
        self.goto_y_display_var = tk.StringVar(value="0.00")
        self.goto_z_display_var = tk.StringVar(value="0.00")
        self.last_cmd_x_display_var = tk.StringVar(value="N/A")
        self.last_cmd_y_display_var = tk.StringVar(value="N/A")
        self.last_cmd_z_display_var = tk.StringVar(value="N/A")
        
        # SCI-FI: StringVars for Header/Footer
        self.header_file_var = tk.StringVar(value="NO FILE")
        self.footer_coords_var = tk.StringVar(value="X: N/A  Y: N/A  Z: N/A")
        self.footer_status_var = tk.StringVar(value="COM: -- @ --")

        # --- Internal Floats (MODEL) ---
        self.target_abs_x = self.PRINTER_BOUNDS['x_max'] / 2
        self.target_abs_y = self.PRINTER_BOUNDS['y_max'] / 2
        self.target_abs_z = self.PRINTER_BOUNDS['z_max'] / 4
        self.last_cmd_abs_x = None # Use None to indicate uninitialized
        self.last_cmd_abs_y = None
        self.last_cmd_abs_z = None
        
        self.coord_mode = tk.StringVar(value="absolute") # 'absolute' or 'relative'

        self.z_canvas_marker_id_blue = None; self.z_canvas_marker_id_red = None
        self.xy_canvas_marker_id_blue = None; self.xy_canvas_marker_id_red = None # SCI-FI: Changed from list

        # --- SCI-FI: Create Header Bar ---
        self.create_header_bar(self.root)

        # --- Main Layout using PanedWindow ---
        main_container = ttk.Frame(root, padding=5, style='TFrame'); 
        main_container.pack(fill=tk.BOTH, expand=True)
        
        self.paned_window = tk.PanedWindow(main_container, orient=tk.HORIZONTAL, sashrelief=tk.FLAT, sashwidth=6, bg=self.COLOR_BG, showhandle=False)
        self.paned_window.pack(fill=tk.BOTH, expand=True)

        # --- Scrollable Left Panel setup ---
        self.left_canvas_frame = ttk.Frame(self.paned_window, style='TFrame')
        self.left_canvas_frame.rowconfigure(0, weight=1); self.left_canvas_frame.columnconfigure(0, weight=1)
        
        # SCI-FI: Configure canvas colors
        self.left_canvas = tk.Canvas(self.left_canvas_frame, highlightthickness=0, bg=self.COLOR_BG)
        
        left_scrollbar = ttk.Scrollbar(self.left_canvas_frame, orient="vertical", command=self.left_canvas.yview)
        
        self.left_panel_scrollable = ttk.Frame(self.left_canvas, style='TFrame')
        self.left_panel_scrollable.bind("<Configure>", lambda e: self.left_canvas.configure(scrollregion=self.left_canvas.bbox("all")))
        
        self.left_canvas.create_window((0, 0), window=self.left_panel_scrollable, anchor="nw")
        self.left_canvas.configure(yscrollcommand=left_scrollbar.set)
        
        self.left_canvas.grid(row=0, column=0, sticky="nsew"); left_scrollbar.grid(row=0, column=1, sticky="ns")
        
        # --- NEW: Bind mouse wheel scrolling ---
        # We bind_all, but only when the mouse is over the left panel
        def _bind_scrolling(event):
            self.root.bind_all("<MouseWheel>", self._on_mousewheel_scroll) # Windows/macOS
            self.root.bind_all("<Button-4>", self._on_mousewheel_scroll)   # Linux (scroll up)
            self.root.bind_all("<Button-5>", self._on_mousewheel_scroll)   # Linux (scroll down)
        
        def _unbind_scrolling(event):
            self.root.unbind_all("<MouseWheel>")
            self.root.unbind_all("<Button-4>")
            self.root.unbind_all("<Button-5>")

        # Bind/unbind when mouse enters/leaves the *entire left panel*
        self.left_canvas_frame.bind("<Enter>", _bind_scrolling)
        self.left_canvas_frame.bind("<Leave>", _unbind_scrolling)
        # --- END NEW ---

        self.paned_window.add(self.left_canvas_frame, minsize=350) # SCI-FI: new minsize

        # --- Right Panel (Log Area) ---
        right_panel = ttk.Frame(self.paned_window, style='TFrame')
        right_panel.rowconfigure(0, weight=1) # Log area expands
        right_panel.rowconfigure(1, weight=0) # Terminal does not expand
        right_panel.columnconfigure(0, weight=1)
        
        self.paned_window.add(right_panel, minsize=300)

        # ===== Build GUI Sections into the SCROLLABLE LEFT PANEL =====
        self.create_connection_frame(self.left_panel_scrollable) # SCI-FI: Moved up
        self.create_file_center_frame(self.left_panel_scrollable)
        self.create_control_frame(self.left_panel_scrollable)
        self.create_progress_frame(self.left_panel_scrollable)
        self.create_position_control_frame(self.left_panel_scrollable) # Combined GoTo/DRO/Visuals
        self.create_manual_control_frame(self.left_panel_scrollable)

        # ===== Place Log Area into RIGHT PANEL =====
        self.create_log_panel(right_panel)
        
        # --- SCI-FI: Create Footer Bar ---
        self.create_footer_bar(self.root)

        # --- Auto-size left panel ---
        self.left_panel_scrollable.update_idletasks()
        required_width = self.left_panel_scrollable.winfo_reqwidth() + 20 
        self.left_canvas_frame.config(width=required_width)
        self.left_canvas.config(width=required_width)
        self.paned_window.paneconfigure(self.left_canvas_frame, width=required_width, minsize=required_width)
        
        # SCI-FI: Set initial pane position
        self.root.update_idletasks()
        self.paned_window.sash_place(0, required_width + 5, 0)


        # --- Final Setup ---
        self.root.after(100, self.check_message_queue)
        self.root.after(300, self.rescan_ports)
        self.root.after(150, self._update_all_displays) # Initial display update

    # --- GUI Creation Methods ---

    # SCI-FI: New method for Header
    def create_header_bar(self, parent):
        """Creates the top header bar."""
        header_bar = ttk.Frame(parent, style='Header.TFrame', padding=(10, 5))
        header_bar.pack(side=tk.TOP, fill=tk.X)
        
        title_label = ttk.Label(header_bar, text="⚡ G-CODE SENDER", style='Header.TLabel', font=self.FONT_HEADER, foreground=self.COLOR_ACCENT_CYAN)
        title_label.pack(side=tk.LEFT)
        
        file_label = ttk.Label(header_bar, textvariable=self.header_file_var, style='Header.TLabel', font=self.FONT_MONO, foreground=self.COLOR_TEXT_SECONDARY)
        file_label.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=20)
        
        # This is a *second* status indicator for the header
        self.header_status_indicator = StatusIndicator(header_bar, self.COLOR_BG)
        self.header_status_indicator.pack(side=tk.RIGHT, padx=10)
        
    # SCI-FI: New method for Footer
    def create_footer_bar(self, parent):
        """Creates the bottom footer bar."""
        footer_bar = ttk.Frame(parent, style='Footer.TFrame', padding=(10, 8))
        footer_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        port_label = ttk.Label(footer_bar, textvariable=self.footer_status_var, style='Footer.TLabel')
        port_label.pack(side=tk.LEFT)
        
        coord_label = ttk.Label(footer_bar, textvariable=self.footer_coords_var, style='Footer.TLabel')
        coord_label.pack(side=tk.RIGHT)

    def create_file_center_frame(self, parent):
        """Creates the Setup frame with file, center coords, and 'Mark Current' button."""
        # SCI-FI: Uppercase text for panel
        frame = ttk.LabelFrame(parent, text="SETUP", padding="10"); frame.pack(fill=tk.X, pady=(0, 10), padx=5)
        # Configure grid for even spacing
        frame.columnconfigure(1, weight=1); frame.columnconfigure(3, weight=1); frame.columnconfigure(5, weight=1); frame.columnconfigure(6, weight=1)

        ttk.Button(frame, text="Select G-Code File", command=self.select_file).grid(row=0, column=0, columnspan=2, padx=(0, 5), sticky="ew")
        # SCI-FI: Use new Filepath.TLabel style
        ttk.Label(frame, textvariable=self.file_path_var, wraplength=300, style='Filepath.TLabel').grid(row=0, column=2, columnspan=5, sticky="ew")
        
        ttk.Label(frame, text="Center X:").grid(row=1, column=0, sticky="w", pady=(5,0));
        self.center_x_entry = ttk.Entry(frame, textvariable=self.center_x_var, width=8)
        self.center_x_entry.grid(row=1, column=1, sticky="ew", pady=(5,0), padx=(0, 5))
        
        ttk.Label(frame, text="Center Y:").grid(row=1, column=2, sticky="w", padx=(5, 0), pady=(5,0));
        self.center_y_entry = ttk.Entry(frame, textvariable=self.center_y_var, width=8)
        self.center_y_entry.grid(row=1, column=3, sticky="ew", pady=(5,0), padx=(0, 5))

        ttk.Label(frame, text="Center Z:").grid(row=1, column=4, sticky="w", padx=(5, 0), pady=(5,0));
        self.center_z_entry = ttk.Entry(frame, textvariable=self.center_z_var, width=8)
        self.center_z_entry.grid(row=1, column=5, sticky="ew", pady=(5,0), padx=(0, 10))

        self.mark_center_button = ttk.Button(frame, text="Mark Current\nas Center", command=self._mark_current_as_center, state=tk.DISABLED)
        self.mark_center_button.grid(row=0, column=6, rowspan=2, sticky="nsew", pady=(0,0), padx=(5,0))
        
        # Bind center entries to update displays on change
        self.center_x_entry.bind('<FocusOut>', self._on_center_change); self.center_x_entry.bind('<Return>', self._on_center_change)
        self.center_y_entry.bind('<FocusOut>', self._on_center_change); self.center_y_entry.bind('<Return>', self._on_center_change)
        self.center_z_entry.bind('<FocusOut>', self._on_center_change); self.center_z_entry.bind('<Return>', self._on_center_change)


    def create_connection_frame(self, parent):
        # SCI-FI: Uppercase text for panel
        frame = ttk.LabelFrame(parent, text="CONNECTION", padding="10"); frame.pack(fill=tk.X, pady=(0, 10), padx=5); frame.columnconfigure(4, weight=1)
        
        ttk.Label(frame, text="Port:").grid(row=0, column=0, sticky="w"); 
        self.port_combobox = ttk.Combobox(frame, textvariable=self.port_var, values=self.available_ports, width=15, state="readonly", font=self.FONT_MONO)
        self.port_combobox.grid(row=0, column=1, padx=(0, 5))
        
        ttk.Button(frame, text="Rescan", command=self.rescan_ports, width=7).grid(row=0, column=2, padx=(0, 10))
        
        ttk.Label(frame, text="Baud Rate:").grid(row=1, column=0, sticky="w", pady=(5,0)); 
        self.baud_entry = ttk.Entry(frame, textvariable=self.baud_var, width=10); 
        self.baud_entry.grid(row=1, column=1, padx=(0, 10), sticky="w")
        
        # SCI-FI: Use Primary.TButton style
        self.connect_button = ttk.Button(frame, text="Connect", command=self.toggle_connection, style='Primary.TButton'); 
        self.connect_button.grid(row=0, column=3, rowspan=2, sticky="ns", padx=(5,0))
        
        self.cancel_connect_button = ttk.Button(frame, text="Cancel", command=self._cancel_connection_attempt)
        
        # SCI-FI: Use new StatusIndicator widget
        self.status_indicator = StatusIndicator(frame, self.COLOR_PANEL_BG)
        self.status_indicator.grid(row=0, column=4, rowspan=2, padx=(10, 0), sticky="w")
        self.status_label = ttk.Label(frame, textvariable=self.connection_status_var, font=self.FONT_BODY_SMALL, style='Filepath.TLabel'); 
        self.status_label.grid(row=0, column=5, rowspan=2, padx=(0, 0), sticky="w")
        

    def create_control_frame(self, parent):
        # SCI-FI: Uppercase text for panel
        frame = ttk.LabelFrame(parent, text="EXECUTION CONTROL", padding=10); frame.pack(fill=tk.X, pady=(0, 10), padx=5)
        
        # SCI-FI: Use Primary.TButton style
        self.start_button = ttk.Button(frame, text="Start Sending", command=self.start_sending, state=tk.DISABLED, style='Primary.TButton'); 
        self.start_button.pack(side=tk.LEFT, padx=(0, 5))
        
        self.pause_resume_button = ttk.Button(frame, text="Pause", command=self.toggle_pause_resume, state=tk.DISABLED); 
        self.pause_resume_button.pack(side=tk.LEFT, padx=(0, 10))
        
        # SCI-FI: Use Danger.TButton style
        self.stop_button = ttk.Button(frame, text="EMERGENCY STOP", command=self.emergency_stop, state=tk.NORMAL, style='Danger.TButton'); 
        self.stop_button.pack(side=tk.LEFT)

    def create_progress_frame(self, parent):
        # SCI-FI: Uppercase text for panel
        frame = ttk.LabelFrame(parent, text="PROGRESS", padding="10"); frame.pack(fill=tk.X, pady=(0, 10), padx=5); frame.columnconfigure(0, weight=1)
        
        # SCI-FI: Style progress label
        self.progress_label = ttk.Label(frame, textvariable=self.progress_label_var, font=self.FONT_MONO, foreground=self.COLOR_TEXT_SECONDARY); 
        self.progress_label.grid(row=0, column=0, sticky="ew", padx=5)
        
        self.progress_bar = ttk.Progressbar(frame, orient=tk.HORIZONTAL, length=300, mode='determinate', variable=self.progress_var); 
        self.progress_bar.grid(row=1, column=0, sticky="ew", padx=5, pady=(5,0))

    def create_position_control_frame(self, parent):
        """Creates the DROs, Go To inputs, and visualization canvases all in one frame."""
        # SCI-FI: Uppercase text for panel
        frame = ttk.LabelFrame(parent, text="POSITION CONTROL & STATUS", padding="10")
        frame.pack(fill=tk.X, expand=True, pady=(0, 10), padx=5)
        frame.columnconfigure(1, weight=1) # Column for canvases
        
        # --- ROW 0: Coordinate Mode Buttons ---
        mode_frame = ttk.Frame(frame, style='Panel.TFrame'); mode_frame.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 10))
        
        # SCI-FI: Use Segment.TButton styles
        self.abs_button = ttk.Button(mode_frame, text="ABSOLUTE", command=lambda: self._set_coord_mode("absolute"), style='Segment.TButton')
        self.abs_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 0))
        self.rel_button = ttk.Button(mode_frame, text="RELATIVE", command=lambda: self._set_coord_mode("relative"), style='Segment.TButton')
        self.rel_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 0))

        # --- ROW 1: DRO (Digital Readout) Frame ---
        # SCI-FI: Create black background frame for DRO
        dro_bg_frame = tk.Frame(frame, bg=self.COLOR_BLACK, relief='sunken', borderwidth=1)
        dro_bg_frame.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(5,5))
        
        # --- FIX: Use the new 'Black.TFrame' style and remove the .configure line ---
        dro_frame = ttk.Frame(dro_bg_frame, padding=(10, 5), style='Black.TFrame')
        dro_frame.pack(fill=tk.BOTH, expand=True)
        # (The failing .configure() line is now deleted)
        
        dro_frame.columnconfigure(1, weight=1); dro_frame.columnconfigure(2, weight=1); dro_frame.columnconfigure(3, weight=1)
        
        ttk.Label(dro_frame, text="CURRENT:", style='DRO.TLabel').grid(row=0, column=0, sticky="w"); 
        ttk.Label(dro_frame, textvariable=self.last_cmd_x_display_var, style='Red.DRO.TLabel').grid(row=0, column=1, sticky="ew")
        ttk.Label(dro_frame, textvariable=self.last_cmd_y_display_var, style='Red.DRO.TLabel').grid(row=0, column=2, sticky="ew", padx=5)
        ttk.Label(dro_frame, textvariable=self.last_cmd_z_display_var, style='Red.DRO.TLabel').grid(row=0, column=3, sticky="ew")
        
        ttk.Label(dro_frame, text=" TARGET:", style='DRO.TLabel').grid(row=1, column=0, sticky="w"); 
        ttk.Label(dro_frame, textvariable=self.goto_x_display_var, style='Blue.DRO.TLabel').grid(row=1, column=1, sticky="ew")
        ttk.Label(dro_frame, textvariable=self.goto_y_display_var, style='Blue.DRO.TLabel').grid(row=1, column=2, sticky="ew", padx=5)
        ttk.Label(dro_frame, textvariable=self.goto_z_display_var, style='Blue.DRO.TLabel').grid(row=1, column=3, sticky="ew")
        
        # --- ROW 2: Canvas Frame (to hold both canvases) ---
        self.canvas_frame = ttk.Frame(frame, height=150, style='Panel.TFrame'); self.canvas_frame.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(10,0))
        self.canvas_frame.rowconfigure(0, weight=1); self.canvas_frame.columnconfigure(1, weight=1) # Let XY expand

        # --- Go To Position Inputs (Left of Canvases) ---
        input_frame = ttk.Frame(self.canvas_frame, style='Panel.TFrame')
        input_frame.grid(row=0, column=0, sticky="nsw", padx=(0, 10)) # SCI-FI: Increased padding
        
        ttk.Label(input_frame, text="Set X:").grid(row=0, column=0, sticky="w"); self.goto_x_entry = ttk.Entry(input_frame, width=8, state=tk.DISABLED); self.goto_x_entry.grid(row=0, column=1, sticky="w", pady=2)
        self.goto_x_entry.bind('<Return>', self._on_goto_entry_commit); self.goto_x_entry.bind('<FocusOut>', self._on_goto_entry_commit)
        
        ttk.Label(input_frame, text="Set Y:").grid(row=1, column=0, sticky="w"); self.goto_y_entry = ttk.Entry(input_frame, width=8, state=tk.DISABLED); self.goto_y_entry.grid(row=1, column=1, sticky="w", pady=2)
        self.goto_y_entry.bind('<Return>', self._on_goto_entry_commit); self.goto_y_entry.bind('<FocusOut>', self._on_goto_entry_commit)

        ttk.Label(input_frame, text="Set Z:").grid(row=2, column=0, sticky="w"); self.goto_z_entry = ttk.Entry(input_frame, width=8, state=tk.DISABLED); self.goto_z_entry.grid(row=2, column=1, sticky="w", pady=2)
        self.goto_z_entry.bind('<Return>', self._on_goto_entry_commit); self.goto_z_entry.bind('<FocusOut>', self._on_goto_entry_commit)
        
        # SCI-FI: Use Primary.TButton style
        self.go_button = ttk.Button(input_frame, text="Go", command=self._go_to_position, state=tk.DISABLED, style='Primary.TButton'); self.go_button.grid(row=3, column=0, columnspan=2, pady=(10, 0), sticky="ew")

        # --- "Go to Center" Button ---
        self.go_to_center_button = ttk.Button(input_frame, text="Go to Center", command=self._go_to_center, state=tk.DISABLED)
        self.go_to_center_button.grid(row=4, column=0, columnspan=2, pady=(5, 0), sticky="ew")

        # --- Canvases ---
        canvas_size = 215
        # SCI-FI: Styled canvas
        self.xy_canvas = tk.Canvas(self.canvas_frame, width=canvas_size, height=canvas_size, bg=self.COLOR_BLACK, highlightthickness=1, highlightbackground=self.COLOR_BORDER)
        self.xy_canvas.grid(row=0, column=1, sticky="n", padx=2) # Centered
        self.xy_canvas.bind("<Button-1>", self._on_xy_canvas_click); self.xy_canvas.bind("<B1-Motion>", self._on_xy_canvas_click); self.xy_canvas.bind("<Configure>", self._draw_xy_canvas_guides)
        
        # SCI-FI: Styled canvas
        self.z_canvas = tk.Canvas(self.canvas_frame, width=25, height=canvas_size, bg=self.COLOR_BLACK, highlightthickness=1, highlightbackground=self.COLOR_BORDER); self.z_canvas.grid(row=0, column=2, sticky="ns", padx=(2, 0))
        self.z_canvas.bind("<Button-1>", self._on_z_canvas_click); self.z_canvas.bind("<B1-Motion>", self._on_z_canvas_click); self.z_canvas.bind("<Configure>", self._draw_z_canvas_marker)

        # Store controls for easy enable/disable
        self.goto_controls = [ self.goto_x_entry, self.goto_y_entry, self.goto_z_entry, self.go_button, 
                                 self.go_to_center_button, self.xy_canvas, self.z_canvas, self.abs_button, self.rel_button ]
        self._set_coord_mode("absolute") # Set initial mode and style
        

    def create_manual_control_frame(self, parent):
        # SCI-FI: Uppercase text for panel
        manual_frame = ttk.LabelFrame(parent, text="MANUAL JOG CONTROL", padding="10"); manual_frame.pack(fill=tk.X, pady=(0, 10), padx=5); manual_frame.columnconfigure(0, weight=1); manual_frame.columnconfigure(4, weight=1)
        
        # SCI-FI: Use Panel.TFrame style
        z_control_left_frame = ttk.Frame(manual_frame, style='Panel.TFrame'); z_control_left_frame.grid(row=0, column=1, rowspan=3, sticky="ns", padx=(0,20)); 
        ttk.Label(z_control_left_frame, text="Z-AXIS", font=self.FONT_BODY_BOLD).pack(pady=(0,5))
        
        self.jog_z_pos = ttk.Button(z_control_left_frame, text="Z+", command=lambda: self._jog('Z', 1), state=tk.DISABLED, style='Jog.TButton'); self.jog_z_pos.pack(pady=(2, 10), fill=tk.X)
        self.jog_z_neg = ttk.Button(z_control_left_frame, text="Z-", command=lambda: self._jog('Z', -1), state=tk.DISABLED, style='Jog.TButton'); self.jog_z_neg.pack(pady=(10, 2), fill=tk.X)
        
        # SCI-FI: Use Panel.TFrame style
        jog_grid_frame = ttk.Frame(manual_frame, style='Panel.TFrame'); jog_grid_frame.grid(row=0, column=2, rowspan=3); 

        jog_grid_frame.rowconfigure(0, weight=1, uniform="jog_grid")
        jog_grid_frame.rowconfigure(1, weight=1, uniform="jog_grid")
        jog_grid_frame.rowconfigure(2, weight=1, uniform="jog_grid")
        jog_grid_frame.columnconfigure(0, weight=1, uniform="jog_grid")
        jog_grid_frame.columnconfigure(1, weight=1, uniform="jog_grid")
        jog_grid_frame.columnconfigure(2, weight=1, uniform="jog_grid")
        
        self.jog_y_pos = ttk.Button(jog_grid_frame, text="Y+", style='Jog.TButton', command=lambda: self._jog('Y', 1), state=tk.DISABLED); self.jog_y_pos.grid(row=0, column=1, padx=2, pady=2, sticky="nsew")
        self.jog_x_neg = ttk.Button(jog_grid_frame, text="X-", style='Jog.TButton', command=lambda: self._jog('X', -1), state=tk.DISABLED); self.jog_x_neg.grid(row=1, column=0, padx=2, pady=2, sticky="nsew")
        self.home_button = ttk.Button(jog_grid_frame, text="⌂", style='Home.TButton', command=self._home_all, state=tk.DISABLED); self.home_button.grid(row=1, column=1, padx=2, pady=2, sticky="nsew")
        self.jog_x_pos = ttk.Button(jog_grid_frame, text="X+", style='Jog.TButton', command=lambda: self._jog('X', 1), state=tk.DISABLED); self.jog_x_pos.grid(row=1, column=2, padx=2, pady=2, sticky="nsew")
        self.jog_y_neg = ttk.Button(jog_grid_frame, text="Y-", style='Jog.TButton', command=lambda: self._jog('Y', -1), state=tk.DISABLED); self.jog_y_neg.grid(row=2, column=1, padx=2, pady=2, sticky="nsew")
        
        # SCI-FI: Use Panel.TFrame style
        jog_params_frame = ttk.Frame(manual_frame, style='Panel.TFrame'); jog_params_frame.grid(row=3, column=1, columnspan=3, pady=(10,0)); 
        ttk.Label(jog_params_frame, text="Step Size (mm):").pack(side=tk.LEFT, padx=(0, 5)); 
        self.jog_step_entry = ttk.Entry(jog_params_frame, textvariable=self.jog_step_var, width=6); self.jog_step_entry.pack(side=tk.LEFT, padx=(0, 15))
        ttk.Label(jog_params_frame, text="Travel Speed (mm/min):").pack(side=tk.LEFT, padx=(0, 5));  # SCI-FI: Shortened label
        self.jog_feedrate_entry = ttk.Entry(jog_params_frame, textvariable=self.jog_feedrate_var, width=8); self.jog_feedrate_entry.pack(side=tk.LEFT)
        
        self.manual_buttons = [self.home_button, self.jog_x_neg, self.jog_x_pos, self.jog_y_neg, self.jog_y_pos, self.jog_z_neg, self.jog_z_pos]; self._set_manual_controls_state(tk.DISABLED)

    def create_log_panel(self, parent):
        """Creates the log area and the new terminal input box."""
        # Row 0: Log area (expandable)
        # SCI-FI: Style log area
        self.log_area = scrolledtext.ScrolledText(parent, height=10, wrap=tk.WORD, state=tk.DISABLED,
                                                    font=self.FONT_TERMINAL,
                                                    bg=self.COLOR_BLACK,
                                                    fg=self.COLOR_ACCENT_GREEN,
                                                    bd=0,
                                                    padx=10,
                                                    pady=10)
        self.log_area.grid(row=0, column=0, sticky="nsew", padx=5, pady=(5,0))
        
        # --- NEW: Row 1: Terminal Input Frame (not expandable) ---
        # SCI-FI: Use Panel.TFrame style
        terminal_frame = ttk.Frame(parent, padding=(5, 5), style='Panel.TFrame')
        terminal_frame.grid(row=1, column=0, sticky="ew", padx=5, pady=(0, 5))
        terminal_frame.columnconfigure(0, weight=1) # Let the entry expand
        
        self.terminal_input = ttk.Entry(terminal_frame, state=tk.DISABLED, font=self.FONT_MONO)
        self.terminal_input.grid(row=0, column=0, sticky="ew", padx=(0, 5))
        
        # SCI-FI: Use Primary.TButton style
        self.terminal_send_button = ttk.Button(terminal_frame, text="Send", state=tk.DISABLED, command=self._send_from_terminal, style='Primary.TButton')
        self.terminal_send_button.grid(row=0, column=1, sticky="w")
        
        # Bind Enter key
        self.terminal_input.bind('<Return>', self._send_from_terminal)


    # --- Utility Methods ---
    def _get_available_ports(self):
        ports = serial.tools.list_ports.comports()
        filtered_ports = [p for p in ports if 'CH340' in p.description or 'USB Serial' in p.description or 'Arduino' in p.description or 'Serial Port' in p.description or not p.description]
        return sorted([port.device for port in filtered_ports])


    def rescan_ports(self):
        current_selection = self.port_var.get()
        self.available_ports = ["Auto-detect"] + self._get_available_ports()
        try:
            if self.serial_connection: self.port_combobox['values'] = self.available_ports;
            else: self.port_combobox['values'] = self.available_ports; self.port_combobox.config(state="readonly")
            if current_selection in self.available_ports: self.port_var.set(current_selection)
            else: self.port_var.set(self.available_ports[0] if self.available_ports else "")
            self.log_message(f"Ports updated: {', '.join(self.available_ports)}")
        except tk.TclError: self.log_message("Warn: Could not update port list.", "WARN")


    def _set_manual_controls_state(self, state):
        for button in self.manual_buttons: button.config(state=state)
        entry_state = tk.NORMAL if state == tk.NORMAL else tk.DISABLED
        self.jog_step_entry.config(state=entry_state); self.jog_feedrate_entry.config(state=entry_state)
        if hasattr(self, 'mark_center_button'):
             self.mark_center_button.config(state=state)


    def _set_goto_controls_state(self, state):
        tk_state = tk.NORMAL if state == tk.NORMAL else tk.DISABLED
        
        for control in self.goto_controls:
            if isinstance(control, (ttk.Entry, ttk.Button)):
                control.config(state=tk_state)
            elif isinstance(control, tk.Canvas):
                # SCI-FI: Set canvas disabled state
                canvas_bg = self.COLOR_BLACK if state == tk.NORMAL else '#111'
                control.config(bg=canvas_bg)
        
        self._update_all_displays() # Call update which handles redraws

    def _set_terminal_controls_state(self, state):
        """Enables or disables the terminal input controls."""
        tk_state = tk.NORMAL if state == tk.NORMAL else tk.DISABLED
        
        if hasattr(self, 'terminal_input'):
            self.terminal_input.config(state=tk_state)
        if hasattr(self, 'terminal_send_button'):
            self.terminal_send_button.config(state=tk_state)


    def log_message(self, message, level="INFO"):
        if not hasattr(self, 'log_area') or not self.log_area: print(f"[LOG_EARLY {level}] {message}"); return
        
        # SCI-FI: Use new color palette
        color_map = { 
            "INFO": self.COLOR_ACCENT_GREEN, 
            "SUCCESS": self.COLOR_ACCENT_CYAN, 
            "WARN": self.COLOR_ACCENT_AMBER, 
            "ERROR": self.COLOR_ACCENT_RED, 
            "CRITICAL": self.COLOR_ACCENT_RED 
        }
        timestamp_color = self.COLOR_TEXT_SECONDARY

        tag_name = level.upper(); timestamp = time.strftime("%H:%M:%S"); 
        
        try:
             self.log_area.config(state=tk.NORMAL)
             
             # Configure timestamp tag
             ts_tag = "timestamp"
             if ts_tag not in self.log_area.tag_names():
                 self.log_area.tag_configure(ts_tag, foreground=timestamp_color)
                 
             # Configure level tag
             if tag_name not in self.log_area.tag_names():
                  self.log_area.tag_configure(tag_name, foreground=color_map.get(level, self.COLOR_ACCENT_GREEN))
                  if level in ["ERROR", "CRITICAL"]: 
                      self.log_area.tag_configure(tag_name, font=(self.FONT_TERMINAL[0], self.FONT_TERMINAL[1], 'bold'))
             
             # Insert message
             self.log_area.insert(tk.END, f"[{timestamp}] ", (ts_tag,))
             self.log_area.insert(tk.END, f"{message}\n", (tag_name,))
             
             self.log_area.see(tk.END); 
             self.log_area.config(state=tk.DISABLED)
        except tk.TclError as e: print(f"[LOG_ERROR {level}] {message} (TclError: {e})")

    # --- File Handling ---

    def select_file(self):
        filepath = filedialog.askopenfilename(title="Select G-Code File", filetypes=[("G-code", "*.gcode"), ("Text", "*.txt"), ("All", "*.*")])
        if filepath: 
            self.file_path_var.set(filepath)
            # SCI-FI: Update header file var
            filename = filepath.split('/')[-1]
            self.header_file_var.set(filename.upper())
            self.log_message(f"Selected file: {filepath}"); 
            self.load_gcode_file(filepath)
            
        if self.serial_connection and self.processed_gcode: self.start_button.config(state=tk.NORMAL)

    def load_gcode_file(self, filepath):
        self.start_button.config(state=tk.DISABLED); self.progress_var.set(0.0); self.progress_label_var.set("Progress: Idle")
        try:
            with open(filepath, 'r') as f: self.gcode_lines = f.readlines()
            self.log_message(f"Loaded {len(self.gcode_lines)} lines."); self.process_gcode() # Process immediately
        except FileNotFoundError: self.log_message(f"Error: File not found '{filepath}'", "ERROR"); messagebox.showerror("Error", f"File not found:\n{filepath}"); self.gcode_lines, self.processed_gcode = [], []
        except Exception as e: self.log_message(f"Error loading file: {e}", "ERROR"); messagebox.showerror("Error", f"Could not read file:\n{e}"); self.gcode_lines, self.processed_gcode = [], []

    def process_gcode(self):
        """
        Processes loaded G-code.
        Translates all G0/G1 moves by the Center offset.
        Blocks G92 commands.
        Checks against hardcoded PRINTER_BOUNDS.
        """
        if not self.gcode_lines: self.processed_gcode = []; self.start_button.config(state=tk.DISABLED if self.serial_connection else tk.DISABLED); return
        
        try:
            center_x = float(self.center_x_var.get())
            center_y = float(self.center_y_var.get())
            center_z = float(self.center_z_var.get())
        except ValueError:
            self.log_message("Error: Invalid center coords. Must be numbers.", "ERROR"); messagebox.showerror("Error", "Center coords must be numbers."); self.processed_gcode = []; self.start_button.config(state=tk.DISABLED); return
        
        temp_processed = []
        lines_processed = 0
        lines_translated = 0
        current_pos = {'x': None, 'y': None, 'z': None} 

        for line_number, line in enumerate(self.gcode_lines, 1):
            stripped_line = line.strip()
            
            if not stripped_line or stripped_line.startswith(';'):
                temp_processed.append(line)
                continue

            if "G92" in stripped_line.upper():
                self.log_message(f"Warning: Blocked G92 command on line {line_number}: '{stripped_line}'", "WARN")
                temp_processed.append(f"; Original line {line_number} blocked (G92 conflicts with absolute coords): {stripped_line}\n")
                continue 
            
            if "G28" in stripped_line.upper():
                temp_processed.append(line)
                if 'X' not in stripped_line.upper() and 'Y' not in stripped_line.upper() and 'Z' not in stripped_line.upper():
                        current_pos = {'x': self.PRINTER_BOUNDS['x_min'], 'y': self.PRINTER_BOUNDS['y_min'], 'z': self.PRINTER_BOUNDS['z_min']}
                else: 
                        if 'X' in stripped_line.upper(): current_pos['x'] = self.PRINTER_BOUNDS['x_min']
                        if 'Y' in stripped_line.upper(): current_pos['y'] = self.PRINTER_BOUNDS['y_min']
                        if 'Z' in stripped_line.upper(): current_pos['z'] = self.PRINTER_BOUNDS['z_min']
                continue 
            
            if stripped_line.upper().startswith("G0") or stripped_line.upper().startswith("G1"):
                parsed_coords = self._parse_gcode_coords(stripped_line)
                
                if not parsed_coords: 
                    temp_processed.append(line)
                    continue
                
                rel_x = parsed_coords.get('x')
                rel_y = parsed_coords.get('y')
                rel_z = parsed_coords.get('z')

                # G-code file is assumed to be absolute *relative to its own 0,0,0*
                abs_x = rel_x + center_x if rel_x is not None else current_pos.get('x') # Use last known if not specified
                abs_y = rel_y + center_y if rel_y is not None else current_pos.get('y')
                abs_z = rel_z + center_z if rel_z is not None else current_pos.get('z')
                
                if rel_x is not None: current_pos['x'] = abs_x
                if rel_y is not None: current_pos['y'] = abs_y
                if rel_z is not None: current_pos['z'] = abs_z
                
                if any(v is None for v in [current_pos['x'], current_pos['y'], current_pos['z']] if v is not None):
                    # Check if any *commanded* axis is None when current pos is also None
                    if (rel_x is not None and current_pos['x'] is None) or \
                         (rel_y is not None and current_pos['y'] is None) or \
                         (rel_z is not None and current_pos['z'] is None):
                        err_msg = f"G-code line {line_number} ({stripped_line}) has unknown coordinate. Ensure G28 or G0/G1 with full X,Y,Z is near start."
                        self.log_message(err_msg, "ERROR"); messagebox.showerror("Processing Error", err_msg); self.processed_gcode = []; self.start_button.config(state=tk.DISABLED); return

                # Bounds Check
                if (abs_x is not None and not (self.PRINTER_BOUNDS['x_min'] <= abs_x <= self.PRINTER_BOUNDS['x_max'])):
                    err_msg = f"G-code line {line_number} ({stripped_line}) results in X out-of-bounds ({abs_x:.2f}). Aborting."
                    self.log_message(err_msg, "ERROR"); messagebox.showerror("Processing Error", err_msg); self.processed_gcode = []; self.start_button.config(state=tk.DISABLED); return
                if (abs_y is not None and not (self.PRINTER_BOUNDS['y_min'] <= abs_y <= self.PRINTER_BOUNDS['y_max'])):
                    err_msg = f"G-code line {line_number} ({stripped_line}) results in Y out-of-bounds ({abs_y:.2f}). Aborting."
                    self.log_message(err_msg, "ERROR"); messagebox.showerror("Processing Error", err_msg); self.processed_gcode = []; self.start_button.config(state=tk.DISABLED); return
                if (abs_z is not None and not (self.PRINTER_BOUNDS['z_min'] <= abs_z <= self.PRINTER_BOUNDS['z_max'])):
                    err_msg = f"G-code line {line_number} ({stripped_line}) results in Z out-of-bounds ({abs_z:.2f}). Aborting."
                    self.log_message(err_msg, "ERROR"); messagebox.showerror("Processing Error", err_msg); self.processed_gcode = []; self.start_button.config(state=tk.DISABLED); return

                f_match = re.search(r"F(\d+(\.\d+)?)", stripped_line)
                e_match = re.search(r"E([-+]?\d*\.?\d+)", stripped_line) 
                
                new_line_parts = [stripped_line.split()[0]] 
                if rel_x is not None: new_line_parts.append(f"X{abs_x:.3f}")
                if rel_y is not None: new_line_parts.append(f"Y{abs_y:.3f}")
                if rel_z is not None: new_line_parts.append(f"Z{abs_z:.3f}")
                if e_match: new_line_parts.append(e_match.group(0))
                if f_match: new_line_parts.append(f_match.group(0))
                
                new_line = " ".join(new_line_parts) + "\n"
                temp_processed.append(new_line)
                lines_translated += 1
            
            else:
                temp_processed.append(line)

        self.processed_gcode = temp_processed
        self.log_message(f"G-code processed. {lines_translated} moves translated to absolute.", "SUCCESS")
        
        if self.serial_connection:
             self.start_button.config(state=tk.NORMAL)


    # --- Connection Handling ---

    def toggle_connection(self):
        if self.serial_connection: self.disconnect_printer()
        else: self.connect_printer()

    def connect_printer(self):
        selected_port = self.port_var.get();
        try: baudrate = int(self.baud_var.get())
        except ValueError: self.log_message("Error: Invalid Baud Rate.", "ERROR"); messagebox.showerror("Error", "Baud Rate must be valid."); return
        self.log_message(f"Connecting... Port: {selected_port}, Baud: {baudrate}...")
        
        self.connect_button.config(state=tk.DISABLED); 
        self.cancel_connect_button.grid(row=0, column=6, rowspan=2, sticky="ns", padx=(5,0)); # SCI-FI: Col 6
        self.cancel_connect_button.config(state=tk.NORMAL)
        self.port_combobox.config(state=tk.DISABLED); self.baud_entry.config(state=tk.DISABLED)
        
        # SCI-FI: Update status indicators
        self.connection_status_var.set("Connecting...")
        self.status_indicator.set_status("busy")
        self.header_status_indicator.set_status("busy")
        self.footer_status_var.set(f"Connecting...")
        
        # --- NEW: Disable terminal on connect attempt ---
        self._set_terminal_controls_state(tk.DISABLED)

        self.cancel_connect_event.clear(); threading.Thread(target=self._connect_thread, args=(selected_port, baudrate), daemon=True).start()

    def _cancel_connection_attempt(self):
        self.log_message("Connection attempt cancelled by user.", "WARN"); self.cancel_connect_event.set()
        self.cancel_connect_button.grid_remove(); self.cancel_connect_button.config(state=tk.DISABLED)
        
        # SCI-FI: Update status indicators
        self.connection_status_var.set("Cancelling..."); 
        self.status_indicator.set_status("busy")
        self.header_status_indicator.set_status("busy")

    def _connect_thread(self, selected_port, baudrate):
        """Worker thread for establishing serial connection, using robust checks."""
        ports_to_try = []
        if selected_port == "Auto-detect":
            self.queue_message("Auto-detecting printer port...")
            ports_to_try = [p for p in self.available_ports if p != "Auto-detect"]
            if not ports_to_try: self.queue_message("No serial ports found.", "ERROR"); self.message_queue.put(("CONNECT_FAIL", "No serial ports found.")); return
        else: ports_to_try = [selected_port]
        serial_conn, found_port, connection_cancelled = None, None, False
        try:
            for port in ports_to_try:
                if self.cancel_connect_event.is_set(): connection_cancelled = True; break
                self.queue_message(f"Trying port: {port}..."); temp_ser = None
                try:
                    temp_ser = serial.Serial(port=port, baudrate=baudrate, timeout=5, write_timeout=10, dsrdtr=False, rtscts=False)
                    try: temp_ser.setDTR(False); time.sleep(0.02); temp_ser.setRTS(False)
                    except Exception as set_err: self.queue_message(f"Note: Could not set DTR/RTS on {port}: {set_err}", "INFO")
                    self.queue_message(f"Port {port} opened. Waiting..."); time.sleep(5)
                    self.queue_message(f"Clearing buffer for {port}..."); temp_ser.reset_input_buffer(); time.sleep(1.5)
                    startup_lines = []; start_read_time = time.time()
                    while time.time() - start_read_time < 2:
                        if self.cancel_connect_event.is_set(): raise InterruptedError("Cancelled")
                        if temp_ser.in_waiting > 0:
                            line = temp_ser.readline().decode('utf-8', errors='ignore').strip()
                            if line: startup_lines.append(line)
                            else: break
                        else: time.sleep(0.1)
                    if startup_lines: self.queue_message(f"Initial: {' | '.join(startup_lines)}")
                    temp_ser.reset_input_buffer()
                    responsive = False; keywords = ['ok', 't:', 'temp', 'echo:', 'marlin', 'start', 'wait']
                    for attempt in range(3):
                        if self.cancel_connect_event.is_set(): raise InterruptedError("Cancelled")
                        self.queue_message(f"Sending M105 to {port} ({attempt + 1})..."); time.sleep(0.5)
                        try: temp_ser.write(b'M105\n'); temp_ser.flush()
                        except serial.SerialTimeoutException as write_e: self.queue_message(f"Write timeout M105 on {port}: {write_e}", "WARN"); break
                        responses = []; start_time = time.time()
                        while time.time() - start_time < 5.0:
                            if self.cancel_connect_event.is_set(): raise InterruptedError("Cancelled")
                            try:
                                if temp_ser.in_waiting > 0:
                                    line = temp_ser.readline().decode('utf-8', errors='ignore').strip()
                                    if line: responses.append(line); self.queue_message(f"Resp on {port}: {line}");
                                    if any(k in line.lower() for k in keywords): responsive = True; break
                                else: time.sleep(0.1)
                            except serial.SerialException as read_e: self.queue_message(f"Read error {port}: {read_e}", "WARN"); time.sleep(0.1); break
                            except Exception as read_e_gen: self.queue_message(f"Unexpected read error {port}: {read_e_gen}", "WARN"); time.sleep(0.1); break
                        if responsive: break
                        time.sleep(1.0)
                    if responsive: serial_conn, found_port = temp_ser, port; break
                    else:
                        if 'write_e' not in locals(): self.queue_message(f"'{port}' no response. Closing.", "WARN")
                        if temp_ser and temp_ser.is_open: temp_ser.close()
                except InterruptedError: self.queue_message("Connection cancelled.", "WARN"); connection_cancelled = True; break
                except serial.SerialException as e:
                    if self.cancel_connect_event.is_set(): connection_cancelled = True; break
                    self.queue_message(f"Comm error '{port}': {e}", "WARN")
                except Exception as e:
                    if self.cancel_connect_event.is_set(): connection_cancelled = True; break
                    self.queue_message(f"Unexpected error '{port}': {e}", "WARN")
                finally:
                    if temp_ser and temp_ser.is_open and serial_conn != temp_ser: temp_ser.close()
            if connection_cancelled: self.message_queue.put(("CONNECT_CANCELLED", None))
            elif serial_conn and found_port: self.message_queue.put(("CONNECTED", (serial_conn, found_port, baudrate)))
            else:
                if not self.cancel_connect_event.is_set(): self.message_queue.put(("CONNECT_FAIL", "No responsive printer found."))
                else: self.message_queue.put(("CONNECT_CANCELLED", None))
        finally: self.message_queue.put(("CONNECT_ATTEMPT_FINISHED", None))

    def disconnect_printer(self, silent=False):
        if self.connect_button['state'] == tk.DISABLED and not self.serial_connection and hasattr(self, 'cancel_connect_button') and self.cancel_connect_button.winfo_ismapped():
             self.log_message("Disconnect during connect - Cancelling.", "WARN"); self.cancel_connect_event.set(); return
        if self.is_sending or self.is_manual_command_running: self.log_message("Cannot disconnect while busy.", "WARN"); messagebox.showwarning("Busy", "Please stop first."); return
        if self.serial_connection:
            try: self.serial_connection.close();
            except Exception as e: self.log_message(f"Disconnect error: {e}", "ERROR") if not silent else None
            
        self.serial_connection = None; 
        
        # SCI-FI: Update status indicators
        self.connection_status_var.set("Disconnected"); 
        self.status_indicator.set_status("off")
        self.header_status_indicator.set_status("off")
        self.footer_status_var.set("COM: -- @ --")
        
        self.connect_button.config(text="Connect", state=tk.NORMAL); self.port_combobox.config(state="readonly"); self.baud_entry.config(state=tk.NORMAL)
        self.start_button.config(state=tk.DISABLED); self._set_manual_controls_state(tk.DISABLED)
        
        if hasattr(self, 'cancel_connect_button'): self.cancel_connect_button.grid_remove(); self.cancel_connect_button.config(state=tk.DISABLED)
        
        self.progress_var.set(0.0); self.progress_label_var.set("Progress: Idle")
        self._set_goto_controls_state(tk.DISABLED)
        
        # --- NEW: Disable terminal on disconnect ---
        self._set_terminal_controls_state(tk.DISABLED)

        self.last_cmd_abs_x, self.last_cmd_abs_y, self.last_cmd_abs_z = None, None, None
        self._update_all_displays() # Redraw markers


    # --- G-Code Sending & Control ---

    def start_sending(self):
        if not self.serial_connection: messagebox.showerror("Error", "Not connected."); return
        self.process_gcode() 
        if not self.processed_gcode: messagebox.showerror("Error", "No valid G-code to send (check file, center coords, and bounds)."); return
        if self.is_sending or self.is_manual_command_running: messagebox.showwarning("Warning", "Printer busy."); return
        self.total_lines_to_send = len([line for line in self.processed_gcode if line.strip() and not line.strip().startswith(';')])
        if self.total_lines_to_send == 0: messagebox.showwarning("Warning", "G-code file has no sendable commands."); return
        self.progress_var.set(0.0); self.progress_label_var.set(f"0/{self.total_lines_to_send} lines")
        self.is_sending = True; self.is_paused = False
        self.stop_event.clear(); self.pause_event.clear()
        self.start_button.config(state=tk.DISABLED); self.pause_resume_button.config(text="Pause", state=tk.NORMAL)
        
        self._set_manual_controls_state(tk.DISABLED); self._set_goto_controls_state(tk.DISABLED)
        
        # --- NEW: Disable terminal on send start ---
        self._set_terminal_controls_state(tk.DISABLED)

        self.log_message("Starting G-code stream...")
        threading.Thread(target=self.gcode_sender_thread, args=(list(self.processed_gcode),), daemon=True).start()

    def toggle_pause_resume(self):
         # Allow pause if sending file OR running manual command
         if not self.is_sending and not self.is_manual_command_running: 
             return
         
         if self.is_paused:
               self.pause_event.set(); self.is_paused = False; self.pause_resume_button.config(text="Pause"); self.log_message("Resumed.", "INFO")
               # When resuming, disable manual controls (since a command is now active again)
               self._set_manual_controls_state(tk.DISABLED); self._set_goto_controls_state(tk.DISABLED)
               self._set_terminal_controls_state(tk.DISABLED)
         else:
               self.pause_event.clear(); self.is_paused = True; self.pause_resume_button.config(text="Resume"); self.log_message("Pausing...", "INFO")
               # When pausing, enable manual controls (to allow jogging)
               self._set_manual_controls_state(tk.NORMAL); self._set_goto_controls_state(tk.NORMAL)
               self._set_terminal_controls_state(tk.NORMAL)
               
               # SCI-FI: Update status indicators to "busy" (amber)
               self.status_indicator.set_status("busy")
               self.header_status_indicator.set_status("busy")

    def emergency_stop(self):
        if not self.serial_connection and not (self.connect_button['state'] == tk.DISABLED and hasattr(self, 'cancel_connect_button') and self.cancel_connect_button.winfo_ismapped()):
             self.log_message("Not connected.", "WARN"); self._reset_gui_after_stop(); return
        if self.connect_button['state'] == tk.DISABLED and not self.serial_connection and hasattr(self, 'cancel_connect_button') and self.cancel_connect_button.winfo_ismapped():
             self.log_message("Stop during connection - Cancelling.", "WARN"); self._cancel_connection_attempt(); return
        
        self.log_message("!!! EMERGENCY STOP triggered !!!", "CRITICAL"); 
        self.pause_event.clear(); self.stop_event.set()
        
        # SCI-FI: Update status indicators to "error" (red)
        self.status_indicator.set_status("error")
        self.header_status_indicator.set_status("error")

        if self.serial_connection:
            try: self.log_message("Sending M112..."); self.serial_connection.write(b'M112\n'); time.sleep(0.5); self.serial_connection.reset_input_buffer(); self.log_message("M112 sent.")
            except Exception as e: self.log_message(f"Error sending M112: {e}", "ERROR")
            finally: self.disconnect_printer(silent=True); messagebox.showwarning("Emergency Stop", "M112 sent.\nPrinter requires reset.\nConnection closed.")
        self._reset_gui_after_stop()


    def _reset_gui_after_stop(self):
        self.is_sending, self.is_paused, self.is_manual_command_running = False, False, False
        self.start_button.config(state=tk.DISABLED); self.pause_resume_button.config(text="Pause", state=tk.DISABLED)
        
        current_state = tk.NORMAL if self.serial_connection else tk.DISABLED
        self._set_manual_controls_state(current_state); self._set_goto_controls_state(current_state)
        
        # --- NEW: Set terminal state on stop/reset ---
        self._set_terminal_controls_state(current_state)

        if hasattr(self, 'cancel_connect_button'): self.cancel_connect_button.grid_remove(); self.cancel_connect_button.config(state=tk.DISABLED)
        self.progress_var.set(0.0); self.progress_label_var.set("Progress: Stopped")
        # Reset last pos to 0,0,0 after stop
        self.last_cmd_abs_x, self.last_cmd_abs_y, self.last_cmd_abs_z = self.PRINTER_BOUNDS['x_min'], self.PRINTER_BOUNDS['y_min'], self.PRINTER_BOUNDS['z_min']
        self._update_all_displays() # Update display


    # --- Thread Workers ---
    def _parse_gcode_coords(self, gcode_line):
        """Extracts X, Y, Z coordinates from a G0/G1/G92 command line."""
        coords = {}
        match = re.search(r"^[Gg](?:[01]|92)\s+(?:[Xx]([-+]?\d*\.?\d+)\s*)?(?:[Yy]([-+]?\d*\.?\d+)\s*)?(?:[Zz]([-+]?\d*\.?\d+)\s*)?", gcode_line)
        if match:
            x, y, z = match.groups()
            if x is not None: coords['x'] = float(x)
            if y is not None: coords['y'] = float(y)
            if z is not None: coords['z'] = float(z)
        return coords


    def _send_manual_command_thread(self, command):
        self.is_manual_command_running = True; self.queue_message(f"Sending: {command.replace(chr(10), '; ')}"); success = False
        target_pos = {'x': self.last_cmd_abs_x, 'y': self.last_cmd_abs_y, 'z': self.last_cmd_abs_z}
        try:
            lines = [line.strip() for line in command.splitlines() if line.strip()]
            in_relative_mode = False 
            try: current_target = {'x': self.last_cmd_abs_x, 'y': self.last_cmd_abs_y, 'z': self.last_cmd_abs_z}
            except (ValueError, TypeError): current_target = {'x': 0.0, 'y': 0.0, 'z': 0.0} # Default to 0 if N/A or None

            for line in lines:
                if self.stop_event.is_set(): raise InterruptedError("Stop before send")
                if "G90" in line.upper(): in_relative_mode = False
                elif "G91" in line.upper(): in_relative_mode = True
                
                if line.upper().startswith("G0") or line.upper().startswith("G1"):
                        parsed = self._parse_gcode_coords(line)
                        if in_relative_mode:
                                pass # Jogging logic now sends absolute G90
                        else: # G90
                            if 'x' in parsed: current_target['x'] = parsed.get('x')
                            if 'y' in parsed: current_target['y'] = parsed.get('y')
                            if 'z' in parsed: current_target['z'] = parsed.get('z')
                elif "G28" in line.upper():
                    if 'X' not in line.upper() and 'Y' not in line.upper() and 'Z' not in line.upper():
                        current_target = {'x': self.PRINTER_BOUNDS['x_min'], 'y': self.PRINTER_BOUNDS['y_min'], 'z': self.PRINTER_BOUNDS['z_min']}
                    else: 
                        if 'X' in line.upper(): current_target['x'] = self.PRINTER_BOUNDS['x_min']
                        if 'Y' in line.upper(): current_target['y'] = self.PRINTER_BOUNDS['y_min']
                        if 'Z' in line.upper(): current_target['z'] = self.PRINTER_BOUNDS['z_min']
                
                self.serial_connection.write(line.encode('utf-8') + b'\n'); self.queue_message(f"Sent: {line}"); ok_received = False; response_buffer = ""
                timeout = 90.0 if "G28" in line.upper() else 20.0; start_time = time.time()
                while time.time() - start_time < timeout:
                    if self.stop_event.is_set(): raise InterruptedError("Stop during wait")
                    
                    # --- NEW: Pause logic for manual commands ---
                    if self.pause_event.is_set():
                        self.queue_message("Manual command paused...", "INFO")
                        self.message_queue.put(("SET_STATUS", "busy")) # SCI-FI
                        self.pause_event.wait() # Block here
                        self.queue_message("Manual command resumed.", "INFO")
                        self.message_queue.put(("SET_STATUS", "on")) # SCI-FI
                    # --- END NEW ---

                    try:
                        if self.serial_connection.in_waiting > 0:
                            response_buffer += self.serial_connection.read(self.serial_connection.in_waiting).decode('utf-8', errors='ignore')
                            while '\n' in response_buffer:
                                full_line, response_buffer = response_buffer.split('\n', 1); full_line = full_line.strip()
                                if full_line: self.queue_message(f"Received: {full_line}");
                                if 'ok' in full_line.lower(): ok_received = True; break
                        if ok_received: break
                        time.sleep(0.05)
                    except serial.SerialException as read_err: self.queue_message(f"Read error: {read_err}", "ERROR"); raise
                if ok_received: self.queue_message(f"Line '{line}' confirmed.", "SUCCESS")
                else: self.queue_message(f"Warning: No 'ok' for '{line}' (timeout: {timeout:.1f}s).", "WARN"); success = False; break
            
            if ok_received or not lines:
                 self.queue_message("Manual command completed.", "SUCCESS"); success = True
                 if any(v is not None for v in current_target.values()):
                       target_pos = current_target
        except InterruptedError: self.queue_message("Manual command interrupted.", "WARN"); success = False
        except serial.SerialException as e: self.queue_message(f"Serial error: {e}", "ERROR"); success = False; self.message_queue.put(("CONNECTION_LOST", None))
        except Exception as e: self.queue_message(f"Error sending command '{command}': {e}", "ERROR"); success = False
        finally:
             if success and any(v is not None for v in target_pos.values()):
                 self.message_queue.put(("POSITION_UPDATE", target_pos))
             self.message_queue.put(("MANUAL_FINISHED", success))


    def _send_manual_command(self, command):
        if not self.serial_connection: messagebox.showerror("Error", "Not connected."); return
        if self.is_sending or self.is_manual_command_running: messagebox.showwarning("Busy", "Printer busy."); return
        
        self.is_manual_command_running = True; self.stop_event.clear()
        self.pause_event.set() # Ensure the thread doesn't start in a paused state
        self._set_manual_controls_state(tk.DISABLED); self._set_goto_controls_state(tk.DISABLED)
        
        # --- NEW: Disable terminal when manual command starts ---
        self._set_terminal_controls_state(tk.DISABLED)

        self.start_button.config(state=tk.DISABLED)
        
        # --- MODIFIED: Enable Pause button if a manual command is running ---
        self.pause_resume_button.config(text="Pause", state=tk.NORMAL)
        
        threading.Thread(target=self._send_manual_command_thread, args=(command,), daemon=True).start()

    def _send_from_terminal(self, event=None):
        """Sends the command from the terminal input entry."""
        command = self.terminal_input.get().strip()
        if not command:
            return
            
        self.log_message(f"Terminal > {command}", "INFO") # Log the command
        self.terminal_input.delete(0, tk.END) # Clear the input
        
        # _send_manual_command will handle connection/busy checks
        self._send_manual_command(command) 

    def _jog(self, axis, direction):
        """Updates target vars (blue) and sends an ABSOLUTE G1 command based on last_cmd (red)."""
        try: step, feedrate = float(self.jog_step_var.get()), float(self.jog_feedrate_var.get())
        except ValueError: messagebox.showerror("Error", "Invalid jog step/feedrate."); return
        if step <= 0 or feedrate <= 0: messagebox.showerror("Error", "Step/Feedrate must be positive."); return
        try:
            # Jog from the LAST COMMANDED (red) position
            current_x = self.last_cmd_abs_x if self.last_cmd_abs_x is not None else self.PRINTER_BOUNDS['x_min']
            current_y = self.last_cmd_abs_y if self.last_cmd_abs_y is not None else self.PRINTER_BOUNDS['y_min']
            current_z = self.last_cmd_abs_z if self.last_cmd_abs_z is not None else self.PRINTER_BOUNDS['z_min']
            
            new_x, new_y, new_z = current_x, current_y, current_z
            if axis == 'X': new_x += direction * step
            elif axis == 'Y': new_y += direction * step
            elif axis == 'Z': new_z += direction * step
            
            # Clamp new target position to printer bounds
            new_x = max(self.PRINTER_BOUNDS['x_min'], min(self.PRINTER_BOUNDS['x_max'], new_x))
            new_y = max(self.PRINTER_BOUNDS['y_min'], min(self.PRINTER_BOUNDS['y_max'], new_y))
            new_z = max(self.PRINTER_BOUNDS['z_min'], min(self.PRINTER_BOUNDS['z_max'], new_z))
            
            # Update the internal TARGET (blue marker)
            self.target_abs_x = new_x
            self.target_abs_y = new_y
            self.target_abs_z = new_z
            self._update_all_displays() # Update all GUI fields and markers
            
            # Send the new ABSOLUTE position command
            self._send_manual_command(f"G90\nG1 X{new_x:.3f} Y{new_y:.3f} Z{new_z:.3f} F{feedrate:.0f}")

        except ValueError as e:
             self.log_message(f"Could not parse position for jog update: {e}", "WARN")

    def _home_all(self): 
        # G28 will update last_cmd_pos to (0,0,0) via POSITION_UPDATE
        self._send_manual_command("G28")
        # Also reset GoTo target vars (blue marker)
        self.target_abs_x = self.PRINTER_BOUNDS['x_min']
        self.target_abs_y = self.PRINTER_BOUNDS['y_min']
        self.target_abs_z = self.PRINTER_BOUNDS['z_min']
        self._update_all_displays()


    def _go_to_position(self):
        """Sends printer to the absolute coordinates stored in target_abs_x/y/z."""
        if not self.serial_connection: messagebox.showerror("Error", "Not connected."); return
        if self.is_sending or self.is_manual_command_running: messagebox.showwarning("Busy", "Printer busy."); return
        try:
            # Read from the internal model, not the GUI
            x, y, z = self.target_abs_x, self.target_abs_y, self.target_abs_z
            feedrate = float(self.jog_feedrate_var.get())
            
            # Bounds Check (should be redundant if inputs are clamped, but good safety)
            if not (self.PRINTER_BOUNDS['x_min'] <= x <= self.PRINTER_BOUNDS['x_max']): raise ValueError(f"X ({x:.2f}) out of bounds")
            if not (self.PRINTER_BOUNDS['y_min'] <= y <= self.PRINTER_BOUNDS['y_max']): raise ValueError(f"Y ({y:.2f}) out of bounds")
            if not (self.PRINTER_BOUNDS['z_min'] <= z <= self.PRINTER_BOUNDS['z_max']): raise ValueError(f"Z ({z:.2f}) out of bounds")
            if feedrate <= 0: raise ValueError("Feedrate must be positive")
                 
        except ValueError as e: self.log_message(f"Go To Error: {e}", "ERROR"); messagebox.showerror("Invalid Input", f"Cannot Go To:\n{e}"); return
        command = f"G90\nG1 X{x:.3f} Y{y:.3f} Z{z:.3f} F{feedrate:.0f}"
        self._send_manual_command(command)

    def _go_to_center(self):
        """Sets the target (blue) coordinates to the stored Center coordinates."""
        try:
            # 1. Read center coordinates
            center_x = float(self.center_x_var.get())
            center_y = float(self.center_y_var.get())
            center_z = float(self.center_z_var.get())

            # 2. Set internal target model
            self.target_abs_x = center_x
            self.target_abs_y = center_y
            self.target_abs_z = center_z

            # 3. Update all displays (labels and canvas markers)
            self._update_all_displays()
            self.log_message(f"Target set to center: X={center_x:.2f}, Y={center_y:.2f}, Z={center_z:.2f}", "INFO")

            # --- ADD THESE LINES ---
            mode = self.coord_mode.get()
            # Determine coordinate to display based on current mode
            display_x = f"{center_x:.2f}" if mode == "absolute" else "0.00"
            display_y = f"{center_y:.2f}" if mode == "absolute" else "0.00"
            display_z = f"{center_z:.2f}" if mode == "absolute" else "0.00"

            # Update the Entry widgets
            self.goto_x_entry.delete(0, tk.END)
            self.goto_x_entry.insert(0, display_x)
            self.goto_y_entry.delete(0, tk.END)
            self.goto_y_entry.insert(0, display_y)
            self.goto_z_entry.delete(0, tk.END)
            self.goto_z_entry.insert(0, display_z)
            # --- END ADDED LINES ---

        except ValueError:
            self.log_message("Cannot Go to Center: Invalid center coordinates.", "ERROR")
            messagebox.showerror("Error", "Center coordinates are invalid. Cannot set target.")

    # --- Canvas Click Handlers ---
    
    def _on_xy_canvas_click(self, event):
        """Handle clicks/drags on the XY canvas to set Target X/Y."""
        if self.go_button['state'] == tk.DISABLED: return
        bounds = self.PRINTER_BOUNDS
        canvas_w = self.xy_canvas.winfo_width(); canvas_h = self.xy_canvas.winfo_height()
        if canvas_w <= 1 or canvas_h <= 1: return
        x_range = bounds['x_max'] - bounds['x_min']; y_range = bounds['y_max'] - bounds['y_min']
        click_x_rel = max(0, min(canvas_w, event.x)); click_y_rel = max(0, min(canvas_h, event.y))
        world_x = bounds['x_min'] + (click_x_rel / canvas_w) * x_range if x_range != 0 else bounds['x_min']
        world_y = bounds['y_min'] + ((canvas_h - click_y_rel) / canvas_h) * y_range if y_range != 0 else bounds['y_min'] # Inverted Y
        
        # Update internal model
        self.target_abs_x = max(bounds['x_min'], min(bounds['x_max'], world_x))
        self.target_abs_y = max(bounds['y_min'], min(bounds['y_max'], world_y))
        
        # Update all GUI displays
        self._update_all_displays()


    def _draw_xy_canvas_guides(self, event=None):
        """Draw origin lines and markers based on printer bounds."""
        self.xy_canvas.delete("all")
        bounds = self.PRINTER_BOUNDS
        w = self.xy_canvas.winfo_width(); h = self.xy_canvas.winfo_height()
        if w <= 1 or h <= 1: return
        
        # SCI-FI: Draw background rect
        self.xy_canvas.create_rectangle(0, 0, w, h, fill=self.COLOR_BLACK, outline=self.COLOR_BORDER, width=1)
        
        x_range = bounds['x_max'] - bounds['x_min']; y_range = bounds['y_max'] - bounds['y_min']
        if x_range == 0 or y_range == 0: return 

        def world_to_canvas(wx, wy):
            cx = w * (wx - bounds['x_min']) / x_range
            cy = h - (h * (wy - bounds['y_min']) / y_range) # Inverted Y
            return cx, cy

        # SCI-FI: Draw grid
        grid_color = "#1a2c3a" # Faint cyan
        for i in range(int(bounds['x_min']), int(bounds['x_max']), 10):
            if i != 0:
                cx, _ = world_to_canvas(i, 0)
                self.xy_canvas.create_line(cx, 0, cx, h, fill=grid_color, tags="guides", dash=(2, 4))
        for i in range(int(bounds['y_min']), int(bounds['y_max']), 10):
            if i != 0:
                _, cy = world_to_canvas(0, i)
                self.xy_canvas.create_line(0, cy, w, cy, fill=grid_color, tags="guides", dash=(2, 4))

        # SCI-FI: Draw origin lines
        if bounds['x_min'] <= 0 <= bounds['x_max'] and bounds['y_min'] <= 0 <= bounds['y_max']:
            canvas_x0, canvas_y0 = world_to_canvas(0, 0)
            self.xy_canvas.create_line(canvas_x0, 0, canvas_x0, h, fill=self.COLOR_BORDER, tags="guides")
            self.xy_canvas.create_line(0, canvas_y0, w, canvas_y0, fill=self.COLOR_BORDER, tags="guides")

        m_size = 5 # SCI-FI: Size of marker

        # --- Draw Blue Marker (Go To Target) ---
        try:
            target_x = self.target_abs_x; target_y = self.target_abs_y # Read from internal model
            marker_cx, marker_cy = world_to_canvas(target_x, target_y)
            marker_cx = max(2, min(w - 2, marker_cx)); marker_cy = max(2, min(h - 2, marker_cy))
            # SCI-FI: Draw oval
            self.xy_canvas_marker_id_blue = self.xy_canvas.create_oval(marker_cx - m_size, marker_cy - m_size, marker_cx + m_size, marker_cy + m_size, fill=self.COLOR_ACCENT_CYAN, outline=self.COLOR_ACCENT_CYAN, tags="marker_blue")
        except Exception: self.xy_canvas_marker_id_blue = None

        # --- Draw Red Marker (Last Commanded Position) ---
        try:
            if self.last_cmd_abs_x is not None and self.last_cmd_abs_y is not None:
                 last_x = self.last_cmd_abs_x; last_y = self.last_cmd_abs_y
                 marker_cx, marker_cy = world_to_canvas(last_x, last_y)
                 marker_cx = max(2, min(w - 2, marker_cx)); marker_cy = max(2, min(h - 2, marker_cy))
                 # SCI-FI: Draw oval
                 self.xy_canvas_marker_id_red = self.xy_canvas.create_oval(marker_cx - m_size, marker_cy - m_size, marker_cx + m_size, marker_cy + m_size, fill=self.COLOR_ACCENT_RED, outline=self.COLOR_ACCENT_RED, tags="marker_red")
            else: self.xy_canvas_marker_id_red = None
        except Exception: self.xy_canvas_marker_id_red = None


    def _on_z_canvas_click(self, event):
        """Handle clicks/drags on the Z canvas to set Go To Z, using printer bounds."""
        if self.go_button['state'] == tk.DISABLED: return
        bounds = self.PRINTER_BOUNDS; canvas_h = self.z_canvas.winfo_height()
        if canvas_h <= 1: return
        z_range = bounds['z_max'] - bounds['z_min']; click_y = max(0, min(canvas_h, event.y))
        world_z = bounds['z_min'] + ((canvas_h - click_y) / canvas_h) * z_range if z_range != 0 else bounds['z_min']
        
        self.target_abs_z = max(bounds['z_min'], min(bounds['z_max'], world_z)) # Update internal model
        self._update_all_displays() # Update GUI

    def _draw_z_canvas_marker(self, event=None):
         """Draws blue (target) and red (last cmd) markers on the Z canvas."""
         self.z_canvas.delete("all")
         bounds = self.PRINTER_BOUNDS; canvas_w = self.z_canvas.winfo_width(); canvas_h = self.z_canvas.winfo_height()
         if canvas_h <= 1: return
         
         # SCI-FI: Draw background
         self.z_canvas.create_rectangle(0, 0, canvas_w, canvas_h, fill=self.COLOR_BLACK, outline=self.COLOR_BORDER, width=1)
         
         z_range = bounds['z_max'] - bounds['z_min']
         if z_range == 0: return

         def z_to_canvas_y(world_z):
             canvas_y = canvas_h - ( (world_z - bounds['z_min']) / z_range * canvas_h )
             return max(1, min(canvas_h - 1, canvas_y))

         # --- Draw Blue Marker (Go To Target) ---
         try:
              target_z = self.target_abs_z # Read from internal model
              canvas_y = z_to_canvas_y(target_z)
              # SCI-FI: Draw horizontal line marker
              self.z_canvas_marker_id_blue = self.z_canvas.create_line(2, canvas_y, canvas_w - 2, canvas_y, fill=self.COLOR_ACCENT_CYAN, width=3, tags="marker_blue")
         except Exception: self.z_canvas_marker_id_blue = None

         # --- Draw Red Marker (Last Commanded Position) ---
         try:
              if self.last_cmd_abs_z is not None:
                  last_z = self.last_cmd_abs_z # Read from internal model
                  canvas_y = z_to_canvas_y(last_z)
                  # SCI-FI: Draw dashed horizontal line
                  self.z_canvas_marker_id_red = self.z_canvas.create_line(1, canvas_y, canvas_w - 1, canvas_y, fill=self.COLOR_ACCENT_RED, width=2, dash=(4, 2), tags="marker_red")
              else: self.z_canvas_marker_id_red = None
         except Exception: self.z_canvas_marker_id_red = None


    # --- G-Code Sending Thread ---
    def gcode_sender_thread(self, gcode_to_send):
        # SCI-FI: Set status to "on" (green)
        self.message_queue.put(("SET_STATUS", "on"))

        sendable_lines = [(i, line) for i, line in enumerate(gcode_to_send) if line.strip() and not line.strip().startswith(';')]
        total_sendable_lines = len(sendable_lines); sent_line_count = 0; success = True
        # Use internal float model as the source of truth
        last_pos = {'x': self.last_cmd_abs_x, 'y': self.last_cmd_abs_y, 'z': self.last_cmd_abs_z}
        if last_pos['x'] is None: last_pos['x'] = self.PRINTER_BOUNDS['x_min']
        if last_pos['y'] is None: last_pos['y'] = self.PRINTER_BOUNDS['y_min']
        if last_pos['z'] is None: last_pos['z'] = self.PRINTER_BOUNDS['z_min']
        
        self.queue_message("Waiting 5s before start..."); time.sleep(5)
        original_line_num = -1
        for original_line_num, line in enumerate(gcode_to_send):
            if self.pause_event.is_set(): 
                self.queue_message("Stream paused...", "INFO")
                self.message_queue.put(("SET_STATUS", "busy")) # SCI-FI
                self.pause_event.wait()
                self.queue_message("Stream resumed.", "INFO")
                self.message_queue.put(("SET_STATUS", "on")) # SCI-FI
            if self.stop_event.is_set(): success = False; break
            if self.stop_event.is_set(): success = False; break

            gcode_line = line.strip(); target_pos_for_line = None
            if not gcode_line or gcode_line.startswith(';'): continue

            sent_line_count += 1; status_update = f"[{sent_line_count}/{total_sendable_lines}]"
            percentage = (sent_line_count / total_sendable_lines) * 100 if total_sendable_lines > 0 else 0
            self.message_queue.put(("PROGRESS", (sent_line_count, total_sendable_lines, percentage)))
            # self.queue_message(f"{status_update} Sending: {gcode_line}") # SCI-FI: Too spammy, remove

            if gcode_line.upper().startswith("G0") or gcode_line.upper().startswith("G1"):
                parsed = self._parse_gcode_coords(gcode_line)
                if parsed:
                    target_pos_for_line = last_pos.copy() 
                    target_pos_for_line.update(parsed) # Update with new absolute coords
            elif "G28" in gcode_line.upper():
                 target_pos_for_line = {'x': self.PRINTER_BOUNDS['x_min'], 'y': self.PRINTER_BOUNDS['y_min'], 'z': self.PRINTER_BOUNDS['z_min']} 
                 if 'X' not in gcode_line.upper() and 'Y' not in gcode_line.upper() and 'Z' not in gcode_line.upper(): last_pos = target_pos_for_line.copy() # Full home
                 else: 
                     if 'X' in gcode_line.upper(): last_pos['x'] = self.PRINTER_BOUNDS['x_min']
                     if 'Y' in gcode_line.upper(): last_pos['y'] = self.PRINTER_BOUNDS['y_min']
                     if 'Z' in gcode_line.upper(): last_pos['z'] = self.PRINTER_BOUNDS['z_min']
                     target_pos_for_line = last_pos.copy()

            try:
                self.serial_connection.write(gcode_line.encode('utf-8') + b'\n'); ok_received = False; response_buffer = ""
                timeout = self.serial_connection.timeout if self.serial_connection.timeout else 10.0; start_time = time.time()
                while time.time() - start_time < timeout + 2:
                    if self.stop_event.is_set(): raise InterruptedError("Stop during wait")
                    if self.pause_event.is_set(): 
                        self.queue_message("Pause while waiting 'ok'...", "INFO")
                        self.message_queue.put(("SET_STATUS", "busy")) # SCI-FI
                        self.pause_event.wait(); 
                        self.queue_message("Resuming wait.", "INFO")
                        self.message_queue.put(("SET_STATUS", "on")) # SCI-FI
                    if self.stop_event.is_set(): raise InterruptedError("Stop after pause")
                    try:
                        if self.serial_connection.in_waiting > 0:
                            response_buffer += self.serial_connection.read(self.serial_connection.in_waiting).decode('utf-8', errors='ignore')
                            while '\n' in response_buffer:
                                full_line, response_buffer = response_buffer.split('\n', 1); full_line = full_line.strip()
                                # SCI-FI: Don't log every 'ok'
                                if full_line and 'ok' not in full_line.lower(): self.queue_message(f"Received: {full_line}");
                                if 'ok' in full_line.lower(): ok_received = True; break
                        if ok_received: break
                        time.sleep(0.02)
                    except serial.SerialException as read_err: self.queue_message(f"Read error: {read_err}", "ERROR"); raise
                
                if ok_received:
                    if target_pos_for_line is not None:
                         last_pos = target_pos_for_line 
                         valid_pos = {k: v for k, v in last_pos.items() if v is not None}
                         if valid_pos: self.message_queue.put(("POSITION_UPDATE", valid_pos))
                else: self.queue_message(f"Warning: No 'ok' for '{gcode_line}' (timeout: {timeout:.1f}s).", "WARN")

            except InterruptedError: self.queue_message("Stream interrupted.", "WARN"); success = False; break
            except serial.SerialException as e: self.queue_message(f"Serial Error: {e}", "ERROR"); success = False; self.message_queue.put(("CONNECTION_LOST", None)); break
            except Exception as e: self.queue_message(f"Error line {sent_line_count} ('{gcode_line}'): {e}", "ERROR"); success = False; break
        final_msg = "Stream finished." if success and not self.stop_event.is_set() and not self.pause_event.is_set() else "Stream stopped."
        self.queue_message(final_msg, "SUCCESS" if success and not self.stop_event.is_set() and not self.pause_event.is_set() else "INFO")
        
        # SCI-FI: Set status based on final state
        if success and not self.stop_event.is_set() and not self.pause_event.is_set():
            self.message_queue.put(("SET_STATUS", "on"))
        elif self.stop_event.is_set():
            self.message_queue.put(("SET_STATUS", "error"))
        
        if not success or self.stop_event.is_set(): self.message_queue.put(("PROGRESS_RESET", None))
        self.message_queue.put(("FILE_SEND_FINISHED", None))

    # --- GUI Update & Event Handling ---

    def check_message_queue(self):
        try:
            while True:
                msg_type, msg_content = self.message_queue.get_nowait()
                if msg_type == "PROGRESS": current, total, percentage = msg_content; self.progress_var.set(percentage); self.progress_label_var.set(f"{current}/{total} lines")
                elif msg_type == "PROGRESS_RESET": self.progress_var.set(0.0); self.progress_label_var.set("Progress: Stopped/Idle")
                elif msg_type == "LOG": self.log_message(msg_content[1], msg_content[0])
                
                # SCI-FI: New message type for status
                elif msg_type == "SET_STATUS":
                    if hasattr(self, 'status_indicator'):
                        self.status_indicator.set_status(msg_content)
                    if hasattr(self, 'header_status_indicator'):
                        self.header_status_indicator.set_status(msg_content)

                elif msg_type == "POSITION_UPDATE":
                    pos_dict = msg_content
                    # Update internal model
                    if 'x' in pos_dict and pos_dict['x'] is not None: self.last_cmd_abs_x = pos_dict['x']
                    if 'y' in pos_dict and pos_dict['y'] is not None: self.last_cmd_abs_y = pos_dict['y']
                    if 'z' in pos_dict and pos_dict['z'] is not None: self.last_cmd_abs_z = pos_dict['z']
                    self._update_all_displays() # Update all GUI labels and markers
                
                elif msg_type == "CONNECTED":
                    self.serial_connection, found_port, baudrate = msg_content; self.log_message(f"Connected on {found_port}!", "SUCCESS"); 
                    
                    # SCI-FI: Update status indicators
                    self.connection_status_var.set(f"Connected to {found_port}"); 
                    self.status_indicator.set_status("on")
                    self.header_status_indicator.set_status("on")
                    self.footer_status_var.set(f"{found_port} @ {baudrate}")
                    
                    self.connect_button.config(text="Disconnect", state=tk.NORMAL); self.port_combobox.config(state=tk.DISABLED); self.baud_entry.config(state=tk.DISABLED);
                    self._set_manual_controls_state(tk.NORMAL); self._set_goto_controls_state(tk.NORMAL)
                    
                    # --- NEW: Enable terminal on connect ---
                    self._set_terminal_controls_state(tk.NORMAL)

                    if self.processed_gcode: self.start_button.config(state=tk.NORMAL)
                    if hasattr(self, 'cancel_connect_button'): self.cancel_connect_button.grid_remove(); self.cancel_connect_button.config(state=tk.DISABLED)
                    
                    self.progress_var.set(0.0); self.progress_label_var.set("Progress: Idle")
                    # Set last_cmd_pos to 0,0,0 on connect
                    self.last_cmd_abs_x, self.last_cmd_abs_y, self.last_cmd_abs_z = self.PRINTER_BOUNDS['x_min'], self.PRINTER_BOUNDS['y_min'], self.PRINTER_BOUNDS['z_min']
                    self._update_all_displays()
                
                elif msg_type == "CONNECT_FAIL":
                    err_msg = msg_content; self.log_message(f"Connect failed: {err_msg}", "ERROR");
                    if "No responsive printer found" not in err_msg: messagebox.showerror("Connection Failed", err_msg)
                    self.serial_connection = None; 
                    
                    # SCI-FI: Update status indicators
                    self.connection_status_var.set("Failed"); 
                    self.status_indicator.set_status("error")
                    self.header_status_indicator.set_status("error")
                    self.footer_status_var.set("COM: -- @ --")
                    
                    self.connect_button.config(text="Connect", state=tk.NORMAL); self.port_combobox.config(state="readonly"); self.baud_entry.config(state=tk.NORMAL)
                    self._set_manual_controls_state(tk.DISABLED); self._set_goto_controls_state(tk.DISABLED);
                    
                    # --- NEW: Disable terminal on connect fail ---
                    self._set_terminal_controls_state(tk.DISABLED)

                    self.start_button.config(state=tk.DISABLED); self.pause_resume_button.config(state=tk.DISABLED)
                    
                    if hasattr(self, 'cancel_connect_button'): self.cancel_connect_button.grid_remove(); self.cancel_connect_button.config(state=tk.DISABLED)
                    
                    self.progress_var.set(0.0); self.progress_label_var.set("Progress: Idle")
                    self.last_cmd_abs_x, self.last_cmd_abs_y, self.last_cmd_abs_z = None, None, None; self._update_all_displays()
                
                elif msg_type == "CONNECT_CANCELLED":
                    self.log_message("Connection cancelled.", "INFO"); self.serial_connection = None; 
                    
                    # SCI-FI: Update status indicators
                    self.connection_status_var.set("Disconnected"); 
                    self.status_indicator.set_status("off")
                    self.header_status_indicator.set_status("off")
                    self.footer_status_var.set("COM: -- @ --")
                    
                    self.connect_button.config(text="Connect", state=tk.NORMAL); self.port_combobox.config(state="readonly"); self.baud_entry.config(state=tk.NORMAL); self.pause_resume_button.config(state=tk.DISABLED)
                    self._set_goto_controls_state(tk.DISABLED);
                    
                    # --- NEW: Disable terminal on connect cancel ---
                    self._set_terminal_controls_state(tk.DISABLED)

                    if hasattr(self, 'cancel_connect_button'): self.cancel_connect_button.grid_remove(); self.cancel_connect_button.config(state=tk.DISABLED)
                    
                    self.progress_var.set(0.0); self.progress_label_var.set("Progress: Idle")
                    self.last_cmd_abs_x, self.last_cmd_abs_y, self.last_cmd_abs_z = None, None, None; self._update_all_displays()
                
                elif msg_type == "CONNECT_ATTEMPT_FINISHED":
                    if not self.serial_connection: self.connect_button.config(state=tk.NORMAL); self.port_combobox.config(state="readonly"); self.baud_entry.config(state=tk.NORMAL)
                    if hasattr(self, 'cancel_connect_button'): self.cancel_connect_button.grid_remove(); self.cancel_connect_button.config(state=tk.DISABLED)
                
                elif msg_type == "FILE_SEND_FINISHED":
                    self.is_sending = False; self.is_paused = False
                    if self.serial_connection and not self.stop_event.is_set():
                        self.start_button.config(state=tk.NORMAL if self.processed_gcode else tk.DISABLED); self._set_manual_controls_state(tk.NORMAL); self._set_goto_controls_state(tk.NORMAL)
                        
                        # --- NEW: Enable terminal on send finish ---
                        self._set_terminal_controls_state(tk.NORMAL)
                        
                        # SCI-FI: Set status back to "on"
                        self.status_indicator.set_status("on")
                        self.header_status_indicator.set_status("on")

                        self.progress_var.set(100.0); self.progress_label_var.set(f"Finished: {self.total_lines_to_send}/{self.total_lines_to_send} lines")
                    else: self.progress_var.set(0.0); self.progress_label_var.set("Progress: Stopped/Idle")
                    self.pause_resume_button.config(text="Pause", state=tk.DISABLED)
                    
                
                elif msg_type == "MANUAL_FINISHED":
                    self.is_manual_command_running = False; self.is_paused = False
                    if self.serial_connection and not self.stop_event.is_set():
                        self._set_manual_controls_state(tk.NORMAL); self._set_goto_controls_state(tk.NORMAL)
                        
                        # --- NEW: Enable terminal on manual cmd finish ---
                        self._set_terminal_controls_state(tk.NORMAL)
                        
                        self.start_button.config(state=tk.NORMAL if self.processed_gcode else tk.DISABLED)
                        
                        # SCI-FI: Set status back to "on"
                        self.status_indicator.set_status("on")
                        self.header_status_indicator.set_status("on")
                        
                    if not self.is_sending: self.pause_resume_button.config(text="Pause", state=tk.DISABLED)
                
                elif msg_type == "CONNECTION_LOST":
                    self.log_message("Connection lost.", "ERROR"); messagebox.showerror("Connection Lost", "Serial connection lost.\nPlease reconnect."); self.disconnect_printer(silent=True)
                    self.progress_var.set(0.0); self.progress_label_var.set("Progress: Idle")
                    
                    # SCI-FI: Update status
                    self.status_indicator.set_status("error")
                    self.header_status_indicator.set_status("error")
                    self.footer_status_var.set("COM: -- @ --")
                    
                    # --- NEW: Disable terminal on connection loss ---
                    self._set_terminal_controls_state(tk.DISABLED)

                    self.last_cmd_abs_x, self.last_cmd_abs_y, self.last_cmd_abs_z = None, None, None; self._update_all_displays()
        except queue.Empty: pass
        finally: self.root.after(100, self.check_message_queue)

    def queue_message(self, message, level="INFO"):
        self.message_queue.put(("LOG", (level, message)))

    def on_closing(self):
        if self.is_sending or self.is_manual_command_running:
            if messagebox.askyesno("Confirm Exit", "Operation in progress. Abort and exit?"): self.pause_event.clear(); self.emergency_stop(); time.sleep(1); self.root.destroy()
            else: return
        else: self.disconnect_printer(silent=True); self.root.destroy()
    
    # --- NEW: Mouse Wheel Scroll Handler ---
    def _on_mousewheel_scroll(self, event):
        """Scrolls the left canvas with the mouse wheel."""
        # On Windows/macOS, event.delta is usually +/- 120.
        # On Linux, event.num is 4 (up) or 5 (down).
        if event.num == 5 or event.delta < 0:
            self.left_canvas.yview_scroll(1, "units")
        elif event.num == 4 or event.delta > 0:
            self.left_canvas.yview_scroll(-1, "units")
        
    # --- New Coordinate Control Methods ---
    
    def _mark_current_as_center(self):
        """Sets the 'Center' coordinates to the 'Last Commanded Position'."""
        if self.last_cmd_abs_x is None:
             self.log_message("Cannot mark center: No known last position.", "WARN")
             messagebox.showwarning("Mark Center Failed", "Cannot mark center.\nNo known printer position.\nTry homing first.")
             return
             
        try:
            self.center_x_var.set(f"{self.last_cmd_abs_x:.2f}")
            self.center_y_var.set(f"{self.last_cmd_abs_y:.2f}")
            self.center_z_var.set(f"{self.last_cmd_abs_z:.2f}")
            self.log_message(f"New center set to: X={self.last_cmd_abs_x:.2f}, Y={self.last_cmd_abs_y:.2f}, Z={self.last_cmd_abs_z:.2f}", "SUCCESS")
            self._update_all_displays()
        except Exception as e:
             self.log_message(f"Error marking center: {e}", "ERROR")

    def _on_center_change(self, event=None):
        """Callback when center coordinates are manually changed."""
        self._update_all_displays() # Update all relative displays

    def _set_coord_mode(self, mode):
        """Sets the coordinate display mode to 'absolute' or 'relative'."""
        if mode == "absolute":
            self.coord_mode.set("absolute")
            # SCI-FI: Use Segment.Active.TButton style
            self.abs_button.config(style="Segment.Active.TButton") 
            self.rel_button.config(style="Segment.TButton") # Default style
        else: # relative
            self.coord_mode.set("relative")
            self.abs_button.config(style="Segment.TButton")
            self.rel_button.config(style="Segment.Active.TButton") # Green text
        
        self.goto_x_entry.delete(0, tk.END); self.goto_y_entry.delete(0, tk.END); self.goto_z_entry.delete(0, tk.END)
        self._update_all_displays()

    def _update_all_displays(self, event=None):
        """Central function to update all coordinate labels and markers based on mode."""
        try:
            center_x = float(self.center_x_var.get())
            center_y = float(self.center_y_var.get())
            center_z = float(self.center_z_var.get())
        except ValueError:
             # Don't log spam, just skip update
             return 

        mode = self.coord_mode.get()

        # --- Update "Current" (Red) Display Labels ---
        if self.last_cmd_abs_x is not None:
            if mode == "absolute":
                self.last_cmd_x_display_var.set(f"{self.last_cmd_abs_x:.2f}")
                self.last_cmd_y_display_var.set(f"{self.last_cmd_abs_y:.2f}")
                self.last_cmd_z_display_var.set(f"{self.last_cmd_abs_z:.2f}")
            else: # relative
                self.last_cmd_x_display_var.set(f"{self.last_cmd_abs_x - center_x:.2f}")
                self.last_cmd_y_display_var.set(f"{self.last_cmd_abs_y - center_y:.2f}")
                self.last_cmd_z_display_var.set(f"{self.last_cmd_abs_z - center_z:.2f}")
            
            # SCI-FI: Update footer
            if hasattr(self, 'footer_coords_var'):
                self.footer_coords_var.set(f"X: {self.last_cmd_abs_x:.2f}  Y: {self.last_cmd_abs_y:.2f}  Z: {self.last_cmd_abs_z:.2f}")
            
        else: # Not yet initialized
            self.last_cmd_x_display_var.set("N/A"); self.last_cmd_y_display_var.set("N/A"); self.last_cmd_z_display_var.set("N/A")
            if hasattr(self, 'footer_coords_var'):
                self.footer_coords_var.set("X: N/A  Y: N/A  Z: N/A")

        # --- Update "Target" (Blue) Display Labels ---
        if mode == "absolute":
            self.goto_x_display_var.set(f"{self.target_abs_x:.2f}")
            self.goto_y_display_var.set(f"{self.target_abs_y:.2f}")
            self.goto_z_display_var.set(f"{self.target_abs_z:.2f}")
        else: # relative
            self.goto_x_display_var.set(f"{self.target_abs_x - center_x:.2f}")
            self.goto_y_display_var.set(f"{self.target_abs_y - center_y:.2f}")
            self.goto_z_display_var.set(f"{self.target_abs_z - center_z:.2f}")

        # --- Redraw Canvas Markers ---
        if hasattr(self, 'xy_canvas') and self.xy_canvas.winfo_width() > 1: self._draw_xy_canvas_guides()
        if hasattr(self, 'z_canvas') and self.z_canvas.winfo_height() > 1: self._draw_z_canvas_marker()
        
    def _on_goto_entry_commit(self, event=None):
        """Called when user presses Enter or FocusOut on a Go To entry."""
        try:
            val_x_str = self.goto_x_entry.get(); val_y_str = self.goto_y_entry.get(); val_z_str = self.goto_z_entry.get()
            new_abs_x, new_abs_y, new_abs_z = self.target_abs_x, self.target_abs_y, self.target_abs_z
            center_x = float(self.center_x_var.get()); center_y = float(self.center_y_var.get()); center_z = float(self.center_z_var.get())
            mode = self.coord_mode.get()

            if val_x_str:
                val_x = float(val_x_str); new_abs_x = val_x + center_x if mode == "relative" else val_x
                self.target_abs_x = max(self.PRINTER_BOUNDS['x_min'], min(self.PRINTER_BOUNDS['x_max'], new_abs_x))
            if val_y_str:
                val_y = float(val_y_str); new_abs_y = val_y + center_y if mode == "relative" else val_y
                self.target_abs_y = max(self.PRINTER_BOUNDS['y_min'], min(self.PRINTER_BOUNDS['y_max'], new_abs_y))
            if val_z_str:
                val_z = float(val_z_str); new_abs_z = val_z + center_z if mode == "relative" else val_z
                self.target_abs_z = max(self.PRINTER_BOUNDS['z_min'], min(self.PRINTER_BOUNDS['z_max'], new_abs_z))
        except ValueError: self.log_message("Invalid coordinate entered in Go To field.", "WARN")
        finally:
            self._update_all_displays()

# --- SCI-FI: New Status Indicator Widget ---
class StatusIndicator(tk.Canvas):
    """A glowing LED-like status indicator."""
    def __init__(self, parent, bg, size=12):
        super().__init__(parent, width=size, height=size, bg=bg, highlightthickness=0)
        
        self.size = size
        self.colors = {
            "off": ("#444", "#555"),
            "on": ("#2a843d", "#3fb950"),  # Green
            "busy": ("#b37400", "#ffa657"), # Amber
            "error": ("#990000", "#ff4444") # Red
        }
        self.current_state = "off"
        self.pulse_on = False
        self.pulse_job = None

        self.led = self.create_oval(2, 2, size-2, size-2, fill=self.colors["off"][0], outline=self.colors["off"][1])
        self.set_status("off")

    def set_status(self, state="off"):
        if state not in self.colors:
            state = "off"
        if state == self.current_state:
            return
            
        self.current_state = state
        
        if self.pulse_job:
            self.after_cancel(self.pulse_job)
            self.pulse_job = None
            
        if state == "on" or state == "busy":
            self._pulse_animation()
        else:
            self.itemconfig(self.led, fill=self.colors[state][1], outline=self.colors[state][1])

    def _pulse_animation(self):
        if self.current_state not in ("on", "busy"):
            return
            
        color1, color2 = self.colors[self.current_state]
        
        if self.pulse_on:
            self.itemconfig(self.led, fill=color1, outline=color1)
            self.pulse_on = False
        else:
            self.itemconfig(self.led, fill=color2, outline=color2)
            self.pulse_on = True
            
        self.pulse_job = self.after(800, self._pulse_animation)

# --- Main Execution ---
if __name__ == "__main__":
    root = tk.Tk()
    
    # SCI-FI: Pre-load fonts if possible (helps with rendering)
    try:
        from tkinter import font
        font.Font(family="Orbitron", size=13).metrics()
        font.Font(family="Inter", size=11).metrics()
        font.Font(family="JetBrains Mono", size=10).metrics()
        font.Font(family="Space Mono", size=16, weight="bold").metrics()
        font.Font(family="Rajdhani", size=13, weight="bold").metrics()
    except Exception as e:
        print(f"Font loading note: {e}") # Non-critical
        
    app = GCodeSenderGUI(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()