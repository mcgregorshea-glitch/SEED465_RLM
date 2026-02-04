import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import serial
import time
import re
import threading
import queue
import serial.tools.list_ports # For port scanning
import csv
from datetime import datetime
import os

# Optional PyVISA import for DMM control
try:
    from pyvisa import ResourceManager
    HAS_PYVISA = True
except ImportError:
    ResourceManager = None
    HAS_PYVISA = False

# Optional RPi.GPIO import for limit switches
try:
    import RPi.GPIO as GPIO
    HAS_GPIO = True
except ImportError:
    GPIO = None
    HAS_GPIO = False

# Optional numpy import for memory optimization
try:
    import numpy as np
    HAS_NUMPY = True
except ImportError:
    HAS_NUMPY = False

# --- DMM Integration Classes ---
DMM_CONFIG = [
    # [120, 100, 'VINP'],
    # [104, 100, 'IINP', 1e3],
    # [107, 100, 'VSYS'],
    # [103, 100, 'SAUX', 1e3],
    [102, 1, 'MyDMM'], # Updated for 10.123.210.102
    # [109, 100, 'SINV', 1e3],
]

# --- GPIO CONFIGURATION ---
# The GPIO pin (BCM numbering) connected to the Z-Max limit switch.
Z_MAX_LIMIT_PIN = 4
Z_PROBE_SPEED = 150 # mm/min


class DmmInst:
    def __init__(self, id: int, samples: int, name: str, scale: float = 1) -> None:
        self.id = id
        self.samples = samples
        self.name = name
        self.scale = scale
        self.pv = None

    def connect(self, pvrmgr) -> None:
        if not pvrmgr: return
        # Using the IP schema from dmm-example.py
        id_str = f'TCPIP0::10.123.210.{self.id}::inst0::INSTR' 
        print(f"Attempting to connect to DMM: {id_str}")
        try:
            self.pv = pvrmgr.open_resource(id_str)
            print(f"Connected to DMM {self.id}")
        except Exception as e:
            print(f"Failed to connect to DMM {self.id}: {e}")

    def setup(self) -> None:
        if self.pv:
            self.pv.write(f'SAMP:COUN {self.samples}')
            self.pv.write(f'CALC:AVER:STAT ON')

    def trigger(self) -> None:
        if self.pv:
            self.pv.write('INIT')

    def ready(self) -> bool:
        if self.pv:
            return int(self.pv.query_ascii_values('CALC:AVER:COUN?')[0]) >= self.samples
        return False

    def read(self) -> float:
        if self.pv:
            return self.pv.query_ascii_values('CALC:AVER:ALL?')[0] * self.scale
        return 0.0


class DmmGroup:
    def __init__(self, config: list) -> None:
        self.dmms: list[DmmInst] = []
        for info in config:
            self.dmms.append(DmmInst(*info))
        self.pvrmgr = None

    def initialize(self) -> None:
        if not HAS_PYVISA or not ResourceManager:
            raise ImportError("pyvisa not installed")
        self.pvrmgr = ResourceManager()
        for dmm in self.dmms:
            dmm.connect(self.pvrmgr)
        for dmm in self.dmms:
            dmm.setup()

    def trigger(self) -> None:
        for dmm in self.dmms:
            dmm.trigger()

    def read(self) -> list:
        # Block until all DMMs are ready
        ready = False
        start_time = time.time()
        timeout = 5.0 # 5 seconds timeout

        while not ready:
            if time.time() - start_time > timeout:
                print(f"TIMEOUT: DMM measurement took longer than {timeout}s.")
                print("LIMITATION: Measurement timed out. Returning 0.0. Ensure DMM is not waiting for an external trigger.")
                return [0.0] * len(self.dmms)

            ready = True
            for dmm in self.dmms:
                ready &= dmm.ready()
            if not ready:
                time.sleep(0.1) 
                
        values = []
        for dmm in self.dmms:
            values.append(dmm.read())
        return values
        
    def close(self):
        if self.pvrmgr:
            for dmm in self.dmms:
                if dmm.pv:
                    try: 
                        dmm.pv.close()
                    except Exception: 
                        pass
            try: 
                self.pvrmgr.close()
            except Exception: 
                pass


class GCodeSenderGUI:
    """
    A graphical user interface for sending G-code to a 3D printer or CNC machine.
    """
    def __init__(self, root):
        self.root = root
        self.root.title("G-Code Sender")
        self.root.geometry("900x800")
        self.root.minsize(700, 600)

        # --- Color Palette ---
        # Defines the color scheme used throughout the application for a consistent look.
        self.COLOR_BG = "#0a0e14"              # Dark background for the main window
        self.COLOR_PANEL_BG = "#161b22"         # Background for frames and panels
        self.COLOR_BORDER = "#30363d"          # Border color for widgets and panels
        self.COLOR_TEXT_PRIMARY = "#e6edf3"     # Primary text color
        self.COLOR_TEXT_SECONDARY = "#7d8590"  # Secondary/dimmed text color
        self.COLOR_ACCENT_CYAN = "#00d4ff"      # Main accent color for buttons, highlights
        self.COLOR_ACCENT_GREEN = "#3fb950"     # Color for success messages, 'on' status
        self.COLOR_ACCENT_AMBER = "#ffa657"     # Color for warnings
        self.COLOR_ACCENT_RED = "#ff4444"       # Color for errors and stop buttons
        self.COLOR_BLACK = "#000000"           # Used for input fields and canvas backgrounds
        self.COLOR_GREY_COMPLETED = "#484f58"   # Color for completed segments of the toolpath

        # --- Fonts ---
        # Defines the font families, sizes, and weights used in the application.
        self.FONT_HEADER = ("Orbitron", 13)
        self.FONT_BODY = ("Inter", 11)
        self.FONT_BODY_SMALL = ("Inter", 9)
        self.FONT_BODY_BOLD = ("Inter", 11, "bold")
        self.FONT_BODY_BOLD_LARGE = ("Inter", 20, "bold")
        self.FONT_MONO = ("JetBrains Mono", 10)
        self.FONT_DRO = ("Space Mono", 16, "bold") # Digital Read-Out font
        self.FONT_TERMINAL = ("JetBrains Mono", 10)

        # Set the main window background color
        self.root.configure(bg=self.COLOR_BG)

        # --- GUI Styling ---
        # This section configures the visual appearance of all ttk widgets (buttons, labels, etc.)
        # It uses a 'clam' theme as a base because it's highly customizable.
        style = ttk.Style()
        style.theme_use('clam')

        # --- Global Style (applies to all ttk widgets) ---
        style.configure('.',
                        background=self.COLOR_PANEL_BG,
                        foreground=self.COLOR_TEXT_PRIMARY,
                        fieldbackground=self.COLOR_BLACK,
                        bordercolor=self.COLOR_BORDER,
                        lightcolor=self.COLOR_BORDER,
                        darkcolor=self.COLOR_BORDER,
                        font=self.FONT_BODY)
        
        # Maps widget states (e.g., 'active', 'disabled') to specific visual properties.
        style.map('.',
                  background=[('disabled', self.COLOR_PANEL_BG), ('active', self.COLOR_PANEL_BG)],
                  foreground=[('disabled', self.COLOR_TEXT_SECONDARY)],
                  bordercolor=[('focus', self.COLOR_ACCENT_CYAN), ('active', self.COLOR_BORDER)],
                  fieldbackground=[('disabled', self.COLOR_PANEL_BG)])

        # --- Frame Styles ---
        style.configure('TFrame', background=self.COLOR_BG)
        style.configure('Panel.TFrame', background=self.COLOR_PANEL_BG)
        style.configure('Header.TFrame', background=self.COLOR_PANEL_BG, bordercolor=self.COLOR_BORDER, borderwidth=1, relief='solid')
        style.configure('Footer.TFrame', background=self.COLOR_BLACK, bordercolor=self.COLOR_BORDER, borderwidth=1, relief='solid')
        style.configure('Black.TFrame', background=self.COLOR_BLACK) # For the DRO panel background

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
        style.configure('TLabel', background=self.COLOR_PANEL_BG, foreground=self.COLOR_TEXT_PRIMARY, font=self.FONT_BODY)
        style.configure('Header.TLabel', background=self.COLOR_PANEL_BG, font=self.FONT_BODY)
        style.configure('Footer.TLabel', background=self.COLOR_BLACK, foreground=self.COLOR_TEXT_SECONDARY, font=self.FONT_MONO)
        style.configure('Filepath.TLabel', background=self.COLOR_PANEL_BG, foreground=self.COLOR_TEXT_SECONDARY, font=self.FONT_BODY_SMALL)
        
        # --- DRO (Digital Read-Out) Label Styles ---
        style.configure('DRO.TLabel', font=self.FONT_MONO, padding=(5, 5), background=self.COLOR_BLACK, foreground=self.COLOR_TEXT_SECONDARY, borderwidth=1, relief='sunken', anchor='w')
        style.configure('Red.DRO.TLabel', font=self.FONT_DRO, width=8, padding=(5, 5), background=self.COLOR_BLACK, foreground=self.COLOR_ACCENT_RED, borderwidth=0, relief='flat', anchor='e')
        style.configure('Blue.DRO.TLabel', font=self.FONT_DRO, width=8, padding=(5, 5), background=self.COLOR_BLACK, foreground=self.COLOR_ACCENT_AMBER, borderwidth=0, relief='flat', anchor='e')

        # --- Button Styles ---
        style.configure('TButton', background=self.COLOR_PANEL_BG, foreground=self.COLOR_TEXT_PRIMARY, bordercolor=self.COLOR_BORDER, borderwidth=1, relief=tk.SOLID, padding=(12, 8), font=self.FONT_BODY)
        style.map('TButton', background=[('active', '#2c333e'), ('pressed', self.COLOR_BLACK)], foreground=[('active', self.COLOR_ACCENT_CYAN)], bordercolor=[('active', self.COLOR_ACCENT_CYAN)])

        # --- Primary Action Button (e.g., Connect, Start) ---
        style.configure('Primary.TButton', background=self.COLOR_ACCENT_CYAN, foreground=self.COLOR_BLACK, font=self.FONT_BODY_BOLD)
        style.map('Primary.TButton', background=[('active', '#00eaff'), ('pressed', self.COLOR_ACCENT_CYAN)], foreground=[('active', self.COLOR_BLACK), ('pressed', self.COLOR_BLACK)], bordercolor=[('active', self.COLOR_ACCENT_CYAN)])

        # --- Danger Button (e.g., STOP) ---
        style.configure('Danger.TButton', background=self.COLOR_ACCENT_RED, foreground=self.COLOR_TEXT_PRIMARY, font=self.FONT_BODY_BOLD)
        style.map('Danger.TButton', background=[('active', '#ff6666'), ('pressed', self.COLOR_ACCENT_RED)], bordercolor=[('active', self.COLOR_ACCENT_RED)])

        # --- Amber Button (e.g., Quick Stop) ---
        style.configure('Amber.TButton', background=self.COLOR_ACCENT_AMBER, foreground=self.COLOR_BLACK, font=self.FONT_BODY_BOLD)
        style.map('Amber.TButton', background=[('active', '#ffc080'), ('pressed', self.COLOR_ACCENT_AMBER)], foreground=[('active', self.COLOR_BLACK), ('pressed', self.COLOR_BLACK)], bordercolor=[('active', self.COLOR_ACCENT_AMBER)])

        # --- Segmented Control Buttons (e.g., Absolute/Relative) ---
        style.configure('Segment.TButton', background=self.COLOR_PANEL_BG, foreground=self.COLOR_TEXT_SECONDARY, padding=(10, 5), font=self.FONT_BODY_SMALL)
        style.map('Segment.TButton', background=[('active', '#2c333e'), ('pressed', self.COLOR_BLACK)], foreground=[('active', self.COLOR_ACCENT_CYAN)])
        style.configure('Segment.Active.TButton', background=self.COLOR_ACCENT_CYAN, foreground=self.COLOR_BLACK, padding=(10, 5), font=self.FONT_BODY_SMALL)
        style.map('Segment.Active.TButton', background=[('active', self.COLOR_ACCENT_CYAN), ('pressed', self.COLOR_ACCENT_CYAN)], foreground=[('active', self.COLOR_BLACK), ('pressed', self.COLOR_BLACK)])

        # --- Jog Buttons ---
        style.configure('Jog.TButton', font=self.FONT_BODY_BOLD, width=5, padding=(10, 10))
        style.configure('JogIcon.TButton', font=("Inter", 18, "bold"), width=5, padding=(4, 4))
        style.configure('Home.TButton', font=self.FONT_BODY_BOLD_LARGE, width=5, padding=(4, 4))

        style.configure('ViewCube.TButton', padding=(2, 2), font=("Inter", 12), width=2)
        style.map('ViewCube.TButton',
            background=[('active', '#2c333e'), ('pressed', self.COLOR_BLACK)],
            foreground=[('active', self.COLOR_ACCENT_CYAN)])

        # --- Toggle Button Styles ---
        custom_font = ("Rajdhani", 9, "bold")
        padding = (5, 3)
        # Style for the "Off" (disabled) state
        style.configure('Custom.Toggle.Off.TButton', background=self.COLOR_PANEL_BG, foreground=self.COLOR_TEXT_SECONDARY, bordercolor=self.COLOR_BORDER, font=custom_font, padding=padding)
        style.map('Custom.Toggle.Off.TButton', bordercolor=[('active', self.COLOR_ACCENT_CYAN)], foreground=[('active', self.COLOR_ACCENT_CYAN)])
        # Style for the "On" (enabled) state - dark with illuminated border/text
        style.configure('Custom.Toggle.On.TButton', background=self.COLOR_PANEL_BG, foreground=self.COLOR_ACCENT_CYAN, bordercolor=self.COLOR_ACCENT_CYAN, font=custom_font, padding=padding)
        style.map('Custom.Toggle.On.TButton', bordercolor=[('active', self.COLOR_ACCENT_CYAN)])

        # --- Entry (Input) Style ---
        style.configure('TEntry',
                        fieldbackground=self.COLOR_BLACK,
                        foreground=self.COLOR_ACCENT_CYAN,
                        bordercolor=self.COLOR_BORDER,
                        insertcolor=self.COLOR_ACCENT_CYAN, # Fix for invisible cursor in some themes
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
        style.map('TCombobox', bordercolor=[('focus', self.COLOR_ACCENT_CYAN)])

        # Style the Combobox dropdown list (OS-dependent).
        self.root.option_add('*TCombobox*Listbox.background', self.COLOR_BLACK)
        self.root.option_add('*TCombobox*Listbox.foreground', self.COLOR_ACCENT_CYAN)
        self.root.option_add('*TCombobox*Listbox.selectBackground', self.COLOR_ACCENT_CYAN)
        self.root.option_add('*TCombobox*Listbox.selectForeground', self.COLOR_BLACK)
        self.root.option_add('*TCombobox*Listbox.font', self.FONT_MONO)
        self.root.option_add('*TCombobox*Listbox.borderWidth', 0)

        # --- Progress Bar Style ---
        style.configure('TProgressbar',
                        troughcolor=self.COLOR_BLACK,
                        background=self.COLOR_ACCENT_CYAN,
                        bordercolor=self.COLOR_BORDER,
                        borderwidth=1,
                        relief=tk.SOLID)

        # --- PanedWindow Sash Style ---
        style.configure('Sash', background=self.COLOR_BG, sashthickness=6, relief=tk.FLAT)
        style.map('Sash', background=[('active', self.COLOR_ACCENT_CYAN)])
        
        # --- Scrollbar Style ---
        style.configure('TScrollbar',
                        background=self.COLOR_BORDER,
                        troughcolor=self.COLOR_BG,
                        bordercolor=self.COLOR_BG,
                        arrowcolor=self.COLOR_TEXT_PRIMARY,
                        relief=tk.FLAT,
                        arrowsize=14)
        style.map('TScrollbar',
                  background=[('active', self.COLOR_ACCENT_CYAN), ('!active', self.COLOR_BORDER)],
                  troughcolor=[('active', self.COLOR_BG), ('!active', self.COLOR_BG)])

        # --- Core Application Attributes ---
        self.serial_connection = None
        self.gcode_filepath = None # Store path to G-code file instead of contents
        self.processed_gcode = []
        self.is_sending = False
        self.is_paused = False
        self.is_manual_command_running = False
        self.is_calibrating = False
        
        # Threading events for controlling background tasks (sending G-code, connecting)
        self.stop_event = threading.Event()
        self.pause_event = threading.Event()
        self.pause_event.set() # Initialize in the "go" state (not paused)
        self.cancel_connect_event = threading.Event()
        
        # A queue for passing messages from background threads to the main GUI thread
        self.message_queue = queue.Queue()
        
        # For the terminal's command history
        self.command_history = []
        self.history_index = 0
        
        # Data structures for visualizing the G-code toolpath
        self.toolpath_by_layer = {}   # {z_level: [((x1,y1),(x2,y2)), ...], ...}
        self.move_to_layer_map = []   # [(z_level, index_on_layer), ...]
        self.ordered_z_values = []    # [z1, z2, z3, ...]
        self.completed_move_count = 0 # How many moves have been completed so far
        self.after_id = None          # To store the ID of the recurring 'after' job

        # Cache for the 3D plot coordinates to avoid recalculating them on every redraw
        self._plot_coords_cache = None
        self._plot_cache_valid = False

        # Control variable for enabling/disabling the 3D plot for performance
        self.is_3d_plot_enabled = tk.BooleanVar(value=True)

        # Control variable for enabling/disabling the 2D plots for performance
        self.is_2d_plot_enabled = tk.BooleanVar(value=True)

        # --- Printer Physical Bounds (in mm) ---
        # E is 'Rotation' (repurposed extruder), units depend on steps/mm config, assuming linear mm for now as requested.
        self.PRINTER_BOUNDS = { 'x_min': 0, 'x_max': 220, 'y_min': 0, 'y_max': 220, 'z_min': 0, 'z_max': 250, 'e_min': -10000, 'e_max': 10000 }

        # --- Tkinter StringVars (for dynamically updating GUI labels) ---
        self.file_path_var = tk.StringVar(value="No file selected")
        self.center_x_var = tk.StringVar(value="110.0")
        self.center_y_var = tk.StringVar(value="110.0")
        self.center_z_var = tk.StringVar(value="50.0")
        self.available_ports = ["Auto-detect"] + self._get_available_ports()
        self.port_var = tk.StringVar(value=self.available_ports[0] if self.available_ports else "")
        self.baud_var = tk.StringVar(value="115200")
        self.connection_status_var = tk.StringVar(value="Status: Disconnected")

        # --- DMM / Measurement State ---
        self.dmm_group = None
        self.is_dmm_connected = False
        self.auto_measure_enabled = tk.BooleanVar(value=False)
        self.log_measurements_enabled = tk.BooleanVar(value=True)
        self.measurement_log_file = None # Internal file handle/flag
        self.log_filepath_var = tk.StringVar(value="") # Initialize empty, set on G-code load
        self.dmm_status_var = tk.StringVar(value="DMMs: Disconnected")
        self.last_measurement_var = tk.StringVar(value="Last: --")
        self.stability_threshold_var = tk.DoubleVar(value=1.0) # 1.0%
        self.max_retries_var = tk.IntVar(value=10)
        self.measurements_per_point_var = tk.IntVar(value=3)
        self.pre_measure_delay_var = tk.DoubleVar(value=0.5) # seconds


        self.jog_step_var = tk.StringVar(value="10")
        self.jog_feedrate_var = tk.StringVar(value="1000")
        self.rotation_step_var = tk.StringVar(value="5")
        self.rotation_feedrate_var = tk.StringVar(value="3000")
        
        self.progress_var = tk.DoubleVar(value=0.0)
        self.progress_label_var = tk.StringVar(value="Progress: Idle")
        self.total_lines_to_send = 0
        self.toolpath_3d_opacity_var = tk.DoubleVar(value=0.8)
        
        # StringVars for the DRO (Digital Read-Out) display
        self.goto_x_display_var = tk.StringVar(value="0.00")
        self.goto_y_display_var = tk.StringVar(value="0.00")
        self.goto_z_display_var = tk.StringVar(value="0.00")
        self.goto_e_display_var = tk.StringVar(value="0.00")
        
        self.last_cmd_x_display_var = tk.StringVar(value="N/A")
        self.last_cmd_y_display_var = tk.StringVar(value="N/A")
        self.last_cmd_z_display_var = tk.StringVar(value="N/A")
        self.last_cmd_e_display_var = tk.StringVar(value="N/A")
        
        # StringVars for the header and footer bars
        self.header_file_var = tk.StringVar(value="NO FILE")
        self.footer_coords_var = tk.StringVar(value="X: N/A  Y: N/A  Z: N/A")
        self.footer_status_var = tk.StringVar(value="COM: -- @ --")

        # --- Internal State (The "Model" in a Model-View-Controller sense) ---
        # These store the actual floating-point numbers for the printer's position.
        # The StringVars above are just for display.
        
        # The 'target' position (blue marker on canvas), where we WANT the printer to go.
        self.target_abs_x = self.PRINTER_BOUNDS['x_max'] / 2
        self.target_abs_y = self.PRINTER_BOUNDS['y_max'] / 2
        self.target_abs_z = self.PRINTER_BOUNDS['z_max'] / 4
        self.target_abs_e = 0.0
        
        # The 'last commanded' position (red marker on canvas), where the printer SHOULD be.
        self.last_cmd_abs_x = None # Use None to indicate the position is not yet known.
        self.last_cmd_abs_y = None
        self.last_cmd_abs_z = None
        self.last_cmd_abs_e = None
        
        # The current coordinate display mode ('absolute' or 'relative' to the center point).
        self.coord_mode = tk.StringVar(value="absolute")

        # --- Build the main GUI layout ---
        # Container for the standard view (Header, Panels, Footer)
        self.main_view_frame = ttk.Frame(root, style='TFrame')
        self.main_view_frame.pack(fill=tk.BOTH, expand=True)

        self.create_header_bar(self.main_view_frame)

        # The main layout is a PanedWindow, which allows the user to resize the left and right panels.
        # Note: This is now packed into self.main_view_frame
        main_container = ttk.Frame(self.main_view_frame, padding=5, style='TFrame')
        main_container.pack(fill=tk.BOTH, expand=True)
        
        self.paned_window = tk.PanedWindow(main_container, orient=tk.HORIZONTAL, sashrelief=tk.FLAT, sashwidth=6, bg=self.COLOR_BG, showhandle=False)
        self.paned_window.pack(fill=tk.BOTH, expand=True)

        # --- Left Panel (Scrollable) ---
        # The left panel contains all the controls and is scrollable in case the window is too short.
        self.left_canvas_frame = ttk.Frame(self.paned_window, style='TFrame')
        self.left_canvas_frame.rowconfigure(0, weight=1)
        self.left_canvas_frame.columnconfigure(0, weight=1)
        
        self.left_canvas = tk.Canvas(self.left_canvas_frame, highlightthickness=0, bg=self.COLOR_BG)
        
        left_scrollbar = ttk.Scrollbar(self.left_canvas_frame, orient="vertical", command=self.left_canvas.yview)
        
        self.left_panel_scrollable = ttk.Frame(self.left_canvas, style='TFrame')
        self.left_panel_scrollable.bind("<Configure>", lambda e: self.left_canvas.configure(scrollregion=self.left_canvas.bbox("all")))
        
        self.left_canvas.create_window((0, 0), window=self.left_panel_scrollable, anchor="nw")
        self.left_canvas.configure(yscrollcommand=left_scrollbar.set)
        
        self.left_canvas.grid(row=0, column=0, sticky="nsew")
        left_scrollbar.grid(row=0, column=1, sticky="ns")
        
        # Bind mouse wheel scrolling for the left panel.
        def _bind_scrolling(event):
            self.root.bind_all("<MouseWheel>", self._on_mousewheel_scroll) # Windows/macOS
            self.root.bind_all("<Button-4>", self._on_mousewheel_scroll)   # Linux (scroll up)
            self.root.bind_all("<Button-5>", self._on_mousewheel_scroll)   # Linux (scroll down)
        
        def _unbind_scrolling(event):
            self.root.unbind_all("<MouseWheel>")
            self.root.unbind_all("<Button-4>")
            self.root.unbind_all("<Button-5>")

        self.left_canvas_frame.bind("<Enter>", _bind_scrolling)
        self.left_canvas_frame.bind("<Leave>", _unbind_scrolling)

        self.paned_window.add(self.left_canvas_frame, minsize=350)

        # --- Right Panel (Tabbed View) ---
        right_panel = ttk.Frame(self.paned_window, style='TFrame', padding=(5, 0, 0, 0))
        right_panel.rowconfigure(0, weight=1)
        right_panel.columnconfigure(0, weight=1)
        
        self.notebook = ttk.Notebook(right_panel, style='TNotebook')
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # Create the frames for each tab
        cli_tab = ttk.Frame(self.notebook, style='Panel.TFrame')
        display_tab = ttk.Frame(self.notebook, style='Panel.TFrame')
        
        self.notebook.add(cli_tab, text="Command Line")
        self.notebook.add(display_tab, text="3D View")
        
        # --- Style for Notebook ---
        style.configure('TNotebook', tabposition='n', borderwidth=0)
        style.configure('TNotebook.Tab', background=self.COLOR_PANEL_BG, foreground=self.COLOR_TEXT_SECONDARY, padding=[12, 6], font=self.FONT_BODY)
        style.map('TNotebook.Tab', 
                  background=[('selected', self.COLOR_BG), ('active', self.COLOR_BORDER)], 
                  foreground=[('selected', self.COLOR_ACCENT_CYAN), ('active', self.COLOR_TEXT_PRIMARY)])

        self.paned_window.add(right_panel, minsize=300)

        # --- Matplotlib 3D Plot Attributes (Lazy Loaded) ---
        self.matplotlib_imported = False
        self.fig_3d = None
        self.ax_3d = None
        self.canvas_3d = None
        self.marker_3d = None

        # --- Populate the GUI panels ---
        self.create_connection_frame(self.left_panel_scrollable)
        self.create_measurement_frame(self.left_panel_scrollable)
        self.create_file_center_frame(self.left_panel_scrollable)
        self.create_control_frame(self.left_panel_scrollable)
        self.create_progress_frame(self.left_panel_scrollable)
        self.create_position_control_frame(self.left_panel_scrollable)
        self.create_manual_control_frame(self.left_panel_scrollable)
        
        # Populate the tabs
        self.create_log_panel(cli_tab)
        self.create_3d_display_panel(display_tab)
        
        self.create_footer_bar(self.main_view_frame)

        # --- Final Layout Adjustments ---
        # This section ensures the left panel is sized correctly on startup.
        self.left_panel_scrollable.update_idletasks()
        required_width = self.left_panel_scrollable.winfo_reqwidth() + 20 
        self.left_canvas_frame.config(width=required_width)
        self.left_canvas.config(width=required_width)
        self.paned_window.paneconfigure(self.left_canvas_frame, width=required_width, minsize=required_width)
        
        # Set the initial position of the sash that divides the panels.
        self.root.update_idletasks()
        self.paned_window.sash_place(0, required_width + 5, 0)

        # --- Final Application Setup ---
        # Start the recurring check of the message queue from background threads.
        self.after_id = self.root.after(100, self.check_message_queue)
        # Periodically rescan for available serial ports.
        self.root.after(300, self.rescan_ports)
        # Perform an initial update of all display labels and canvases.
        self.root.after(150, self._update_all_displays)
        
        # Initialize GPIO
        self._setup_gpio()

    def _setup_gpio(self):
        """Configures the GPIO pins for the limit switches."""
        if not HAS_GPIO:
            self.log_message("GPIO (RPi.GPIO) not available. Limit switch detection DISABLED.", "WARN")
            return

        try:
            GPIO.setmode(GPIO.BCM)
            # Assuming Normally Closed (Switch connects to GND normally, Opens when hit)
            # Normal State: LOW (GND)
            # Trigger State: HIGH (Pulled Up)
            GPIO.setup(Z_MAX_LIMIT_PIN, GPIO.IN, pull_up_down=GPIO.PUD_UP)
            
            # Safety Interrupt: Watchdog for Z-Max
            # Triggers _on_z_max_trigger when switch opens (voltage rises)
            GPIO.add_event_detect(Z_MAX_LIMIT_PIN, GPIO.RISING, callback=self._on_z_max_trigger, bouncetime=200)
            
            # Initial Check: If switch is already open (HIGH), trigger immediately!
            if GPIO.input(Z_MAX_LIMIT_PIN) == 1:
                self.log_message("Startup Safety Check: Z-Max is OPEN (Triggered)!", "CRITICAL")
                self._on_z_max_trigger(Z_MAX_LIMIT_PIN)
            
            self.log_message(f"GPIO initialized. Z-Max on Pin {Z_MAX_LIMIT_PIN} (Safety Interrupt Active).", "INFO")
        except Exception as e:
            self.log_message(f"GPIO Setup Error: {e}", "WARN")

    def _on_z_max_trigger(self, channel):
        """Hardware Interrupt: Immediately stops printer if Z-Max is hit."""
        # Print to system terminal immediately for debugging/confirmation
        print(f"\n[HARDWARE] !!! Z-MAX LIMIT TRIGGERED on Channel {channel} !!!\n")
        
        # Send M112 immediately to the serial port (bypass queue for speed)
        if self.serial_connection:
            try:
                self.serial_connection.write(b'M112\n')
            except Exception:
                pass
        # Notify main thread to handle GUI/Logic cleanup
        self.message_queue.put(("EMERGENCY_STOP_TRIGGERED", "Z-Max Limit Switch Hit!"))

    # --- GUI Creation Methods ---

    def create_header_bar(self, parent):
        """Creates the top header bar containing the title and current filename."""
        header_bar = ttk.Frame(parent, style='Header.TFrame', padding=(10, 5))
        header_bar.pack(side=tk.TOP, fill=tk.X)
        
        title_label = ttk.Label(header_bar, text="⚡ G-CODE SENDER", style='Header.TLabel', font=self.FONT_HEADER, foreground=self.COLOR_ACCENT_CYAN)
        title_label.pack(side=tk.LEFT)
        
        file_label = ttk.Label(header_bar, textvariable=self.header_file_var, style='Header.TLabel', font=self.FONT_MONO, foreground=self.COLOR_TEXT_SECONDARY)
        file_label.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=20)
        
        # A small status indicator LED in the header.
        self.header_status_indicator = StatusIndicator(header_bar, self.COLOR_BG)
        self.header_status_indicator.pack(side=tk.RIGHT, padx=10)
        
    def create_footer_bar(self, parent):
        """Creates the bottom footer bar for status and coordinates."""
        footer_bar = ttk.Frame(parent, style='Footer.TFrame', padding=(10, 8))
        footer_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        port_label = ttk.Label(footer_bar, textvariable=self.footer_status_var, style='Footer.TLabel')
        port_label.pack(side=tk.LEFT)
        
        coord_label = ttk.Label(footer_bar, textvariable=self.footer_coords_var, style='Footer.TLabel')
        coord_label.pack(side=tk.RIGHT)

    def create_file_center_frame(self, parent):
        """Creates the 'SETUP' panel for file selection and defining the center point."""
        frame = ttk.LabelFrame(parent, text="SETUP", padding="10")
        frame.pack(fill=tk.X, pady=(0, 10), padx=5)
        frame.columnconfigure(1, weight=1); frame.columnconfigure(3, weight=1); frame.columnconfigure(5, weight=1); frame.columnconfigure(6, weight=1)

        ttk.Button(frame, text="Select G-Code File", command=self.select_file).grid(row=0, column=0, columnspan=2, padx=(0, 5), sticky="ew")
        ttk.Label(frame, textvariable=self.file_path_var, wraplength=300, style='Filepath.TLabel').grid(row=0, column=2, columnspan=5, sticky="ew")
        
        # Entry fields for the user to define the "center" of their workpiece in the printer's coordinate system.
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
        self.mark_center_button.grid(row=0, column=6, rowspan=1, sticky="nsew", pady=(0,2), padx=(5,0))

        self.collision_test_button = ttk.Button(frame, text="Collision\nAvoidance Test", command=self._open_collision_test_screen, state=tk.DISABLED) # Initially disabled until connected
        self.collision_test_button.grid(row=1, column=6, rowspan=1, sticky="nsew", pady=(2,0), padx=(5,0))
        
        # Bind changes in the center entries to update the coordinate displays.
        self.center_x_entry.bind('<FocusOut>', self._on_center_change); self.center_x_entry.bind('<Return>', self._on_center_change)
        self.center_y_entry.bind('<FocusOut>', self._on_center_change); self.center_y_entry.bind('<Return>', self._on_center_change)
        self.center_z_entry.bind('<FocusOut>', self._on_center_change); self.center_z_entry.bind('<Return>', self._on_center_change)


    def create_connection_frame(self, parent):
        """Creates the 'CONNECTION' panel for managing the serial connection."""
        frame = ttk.LabelFrame(parent, text="CONNECTION", padding="10")
        frame.pack(fill=tk.X, pady=(0, 10), padx=5)
        frame.columnconfigure(4, weight=1)
        
        ttk.Label(frame, text="Port:").grid(row=0, column=0, sticky="w"); 
        self.port_combobox = ttk.Combobox(frame, textvariable=self.port_var, values=self.available_ports, width=15, state="readonly", font=self.FONT_MONO)
        self.port_combobox.grid(row=0, column=1, padx=(0, 5))
        
        ttk.Button(frame, text="Rescan", command=self.rescan_ports, width=7).grid(row=0, column=2, padx=(0, 10))
        
        ttk.Label(frame, text="Baud Rate:").grid(row=1, column=0, sticky="w", pady=(5,0)); 
        self.baud_entry = ttk.Entry(frame, textvariable=self.baud_var, width=10); 
        self.baud_entry.grid(row=1, column=1, padx=(0, 10), sticky="w")
        
        self.connect_button = ttk.Button(frame, text="Connect", command=self.toggle_connection, style='Primary.TButton'); 
        self.connect_button.grid(row=0, column=3, rowspan=2, sticky="ns", padx=(5,0))
        
        # The cancel button is only shown during a connection attempt.
        self.cancel_connect_button = ttk.Button(frame, text="Cancel", command=self._cancel_connection_attempt)
        
        # The status indicator "LED" and its text label.
        self.status_indicator = StatusIndicator(frame, self.COLOR_PANEL_BG)
        self.status_indicator.grid(row=0, column=4, rowspan=2, padx=(10, 0), sticky="w")
        self.status_label = ttk.Label(frame, textvariable=self.connection_status_var, font=self.FONT_BODY_SMALL, style='Filepath.TLabel'); 
        self.status_label.grid(row=0, column=5, rowspan=2, padx=(0, 0), sticky="w")
        

    def create_measurement_frame(self, parent):
        """Creates the 'MEASUREMENT' panel."""
        frame = ttk.LabelFrame(parent, text="MEASUREMENT", padding="10")
        frame.pack(fill=tk.X, pady=(0, 10), padx=5)
        
        # Connect / Status
        conn_frame = ttk.Frame(frame, style='Panel.TFrame')
        conn_frame.pack(fill=tk.X, pady=(0, 5))
        
        self.dmm_connect_button = ttk.Button(conn_frame, text="Connect DMMs", command=self.toggle_dmm_connection)
        self.dmm_connect_button.pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Label(conn_frame, textvariable=self.dmm_status_var, style='Filepath.TLabel').pack(side=tk.LEFT)

        # Auto Measure Toggle
        opts_frame = ttk.Frame(frame, style='Panel.TFrame')
        opts_frame.pack(fill=tk.X, pady=(0, 5))
        
        self.auto_measure_check = ttk.Checkbutton(opts_frame, text="Auto-Measure on Move", variable=self.auto_measure_enabled, command=self._on_auto_measure_toggle)
        self.auto_measure_check.pack(side=tk.LEFT, padx=(0, 15))
        
        # Logging Frame
        log_frame = ttk.Frame(frame, style='Panel.TFrame')
        log_frame.pack(fill=tk.X, pady=(0, 5))
        
        self.log_check = ttk.Checkbutton(log_frame, text="Log to CSV", variable=self.log_measurements_enabled)
        self.log_check.pack(side=tk.LEFT)
        
        self.log_path_entry = ttk.Entry(log_frame, textvariable=self.log_filepath_var, width=15, font=self.FONT_BODY_SMALL)
        self.log_path_entry.pack(side=tk.LEFT, padx=(5, 5), fill=tk.X, expand=True)
        
        self.browse_log_btn = ttk.Button(log_frame, text="...", width=3, command=self.select_log_file)
        self.browse_log_btn.pack(side=tk.LEFT)

        # Settings Frame (Stability)
        settings_frame = ttk.Frame(frame, style='Panel.TFrame')
        settings_frame.pack(fill=tk.X, pady=(0, 5))
        
        # Row 1: Threshold & Retries
        ttk.Label(settings_frame, text="Stability (%):", font=self.FONT_BODY_SMALL).grid(row=0, column=0, sticky="w")
        ttk.Entry(settings_frame, textvariable=self.stability_threshold_var, width=4, font=self.FONT_MONO).grid(row=0, column=1, sticky="w", padx=5)
        
        ttk.Label(settings_frame, text="Max Retries:", font=self.FONT_BODY_SMALL).grid(row=0, column=2, sticky="w", padx=(10, 0))
        ttk.Entry(settings_frame, textvariable=self.max_retries_var, width=4, font=self.FONT_MONO).grid(row=0, column=3, sticky="w", padx=5)

        # Row 2: Window & Delay
        ttk.Label(settings_frame, text="Measurements per Point:", font=self.FONT_BODY_SMALL).grid(row=1, column=0, sticky="w", pady=(5,0))
        ttk.Entry(settings_frame, textvariable=self.measurements_per_point_var, width=4, font=self.FONT_MONO).grid(row=1, column=1, sticky="w", padx=5, pady=(5,0))

        ttk.Label(settings_frame, text="Stabilizing Delay (s):", font=self.FONT_BODY_SMALL).grid(row=1, column=2, sticky="w", padx=(10, 0), pady=(5,0))
        ttk.Entry(settings_frame, textvariable=self.pre_measure_delay_var, width=4, font=self.FONT_MONO).grid(row=1, column=3, sticky="w", padx=5, pady=(5,0))

        # Manual Trigger & Last Reading
        ctrl_frame = ttk.Frame(frame, style='Panel.TFrame')
        ctrl_frame.pack(fill=tk.X, pady=(5, 0))
        
        self.measure_button = ttk.Button(ctrl_frame, text="Measure Now", command=self.trigger_manual_measurement, state=tk.DISABLED)
        self.measure_button.pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Label(ctrl_frame, textvariable=self.last_measurement_var, font=self.FONT_MONO).pack(side=tk.LEFT)

    def _on_auto_measure_toggle(self):
        if self.auto_measure_enabled.get():
            self.log_message("Auto-measurement ENABLED. DMMs will trigger after every move.", "INFO")
        else:
            self.log_message("Auto-measurement DISABLED.", "INFO")


    def create_control_frame(self, parent):
        """Creates the 'EXECUTION CONTROL' panel with Start, Pause, and Stop buttons."""
        frame = ttk.LabelFrame(parent, text="EXECUTION CONTROL", padding=10)
        frame.pack(fill=tk.X, pady=(0, 10), padx=5)
        
        self.start_button = ttk.Button(frame, text="Start Sending", command=self.start_sending, state=tk.DISABLED, style='Primary.TButton'); 
        self.start_button.pack(side=tk.LEFT, padx=(0, 5))
        
        self.pause_resume_button = ttk.Button(frame, text="Pause", command=self.toggle_pause_resume, state=tk.DISABLED); 
        self.pause_resume_button.pack(side=tk.LEFT, padx=(0, 10))
        
        # M112: A hard stop that requires a printer reset.
        self.stop_button = ttk.Button(frame, text="EMERGENCY STOP", command=self.emergency_stop, state=tk.NORMAL, style='Danger.TButton'); 
        self.stop_button.pack(side=tk.LEFT)
        
        # M410: A soft stop that finishes the current move then halts.
        self.quick_stop_button = ttk.Button(frame, text="QUICK STOP", command=self.quick_stop, state=tk.NORMAL, style='Amber.TButton')
        self.quick_stop_button.pack(side=tk.LEFT, padx=(10, 0))

    def create_progress_frame(self, parent):
        """Creates the 'PROGRESS' panel with a progress bar and label."""
        frame = ttk.LabelFrame(parent, text="PROGRESS", padding="10")
        frame.pack(fill=tk.X, pady=(0, 10), padx=5)
        frame.columnconfigure(0, weight=1)
        
        self.progress_label = ttk.Label(frame, textvariable=self.progress_label_var, font=self.FONT_MONO, foreground=self.COLOR_TEXT_SECONDARY); 
        self.progress_label.grid(row=0, column=0, sticky="ew", padx=5)
        
        self.progress_bar = ttk.Progressbar(frame, orient=tk.HORIZONTAL, length=300, mode='determinate', variable=self.progress_var); 
        self.progress_bar.grid(row=1, column=0, sticky="ew", padx=5, pady=(5,0))

    def _toggle_2d_plot_button(self):
        """Toggles the state of the 2D plot and updates the button style."""
        self.is_2d_plot_enabled.set(not self.is_2d_plot_enabled.get())
        self._update_2d_plot_button_style()
        self._update_all_displays()

    def _update_2d_plot_button_style(self):
        """Applies the 'On' or 'Off' style to the 2D plot toggle button."""
        if not hasattr(self, 'toggle_2d_button'): return
        if self.is_2d_plot_enabled.get():
            self.toggle_2d_button.config(style='Custom.Toggle.On.TButton')
        else:
            self.toggle_2d_button.config(style='Custom.Toggle.Off.TButton')

    def create_position_control_frame(self, parent):
        """Creates the main panel for position control, including DROs, inputs, and visual canvases."""
        frame = ttk.LabelFrame(parent, text="POSITION CONTROL & STATUS", padding="10")
        frame.pack(fill=tk.X, expand=True, pady=(0, 10), padx=5)
        frame.columnconfigure(1, weight=1)
        
        # --- Coordinate Mode Buttons (Absolute/Relative) ---
        mode_frame = ttk.Frame(frame, style='Panel.TFrame')
        mode_frame.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 10))
        
        self.abs_button = ttk.Button(mode_frame, text="ABSOLUTE", command=lambda: self._set_coord_mode("absolute"), style='Segment.TButton')
        self.abs_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 0))
        self.rel_button = ttk.Button(mode_frame, text="RELATIVE", command=lambda: self._set_coord_mode("relative"), style='Segment.TButton')
        self.rel_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 0))

        # --- DRO (Digital Readout) Frame ---
        dro_bg_frame = tk.Frame(frame, bg=self.COLOR_BLACK, relief='sunken', borderwidth=1)
        dro_bg_frame.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(5,5))
        
        dro_frame = ttk.Frame(dro_bg_frame, padding=(10, 5), style='Black.TFrame')
        dro_frame.pack(fill=tk.BOTH, expand=True)
        
        dro_frame.columnconfigure(1, weight=1); dro_frame.columnconfigure(2, weight=1); dro_frame.columnconfigure(3, weight=1); dro_frame.columnconfigure(4, weight=1)
        
        # Labels for the 'CURRENT' (last commanded) position.
        ttk.Label(dro_frame, text="CURRENT:", style='DRO.TLabel').grid(row=0, column=0, sticky="w"); 
        ttk.Label(dro_frame, textvariable=self.last_cmd_x_display_var, style='Red.DRO.TLabel').grid(row=0, column=1, sticky="ew")
        ttk.Label(dro_frame, textvariable=self.last_cmd_y_display_var, style='Red.DRO.TLabel').grid(row=0, column=2, sticky="ew", padx=5)
        ttk.Label(dro_frame, textvariable=self.last_cmd_z_display_var, style='Red.DRO.TLabel').grid(row=0, column=3, sticky="ew")
        ttk.Label(dro_frame, textvariable=self.last_cmd_e_display_var, style='Red.DRO.TLabel').grid(row=0, column=4, sticky="ew", padx=5)
        
        # Labels for the 'TARGET' (Go To) position.
        ttk.Label(dro_frame, text=" TARGET:", style='DRO.TLabel').grid(row=1, column=0, sticky="w"); 
        ttk.Label(dro_frame, textvariable=self.goto_x_display_var, style='Blue.DRO.TLabel').grid(row=1, column=1, sticky="ew")
        ttk.Label(dro_frame, textvariable=self.goto_y_display_var, style='Blue.DRO.TLabel').grid(row=1, column=2, sticky="ew", padx=5)
        ttk.Label(dro_frame, textvariable=self.goto_z_display_var, style='Blue.DRO.TLabel').grid(row=1, column=3, sticky="ew")
        ttk.Label(dro_frame, textvariable=self.goto_e_display_var, style='Blue.DRO.TLabel').grid(row=1, column=4, sticky="ew", padx=5)
        
        # --- Frame to hold the XY, Z, and E canvases ---
        self.canvas_frame = ttk.Frame(frame, height=150, style='Panel.TFrame')
        self.canvas_frame.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(10,0))
        self.canvas_frame.rowconfigure(0, weight=1)
        self.canvas_frame.columnconfigure(1, weight=1)

        # --- 'Go To' Position Input Fields ---
        input_frame = ttk.Frame(self.canvas_frame, style='Panel.TFrame')
        input_frame.grid(row=0, column=0, sticky="nsw", padx=(0, 10))
        
        ttk.Label(input_frame, text="Set X:").grid(row=0, column=0, sticky="w"); self.goto_x_entry = ttk.Entry(input_frame, width=8, state=tk.DISABLED); self.goto_x_entry.grid(row=0, column=1, sticky="w", pady=2)
        self.goto_x_entry.bind('<Return>', self._on_goto_entry_commit); self.goto_x_entry.bind('<FocusOut>', self._on_goto_entry_commit)
        
        ttk.Label(input_frame, text="Set Y:").grid(row=1, column=0, sticky="w"); self.goto_y_entry = ttk.Entry(input_frame, width=8, state=tk.DISABLED); self.goto_y_entry.grid(row=1, column=1, sticky="w", pady=2)
        self.goto_y_entry.bind('<Return>', self._on_goto_entry_commit); self.goto_y_entry.bind('<FocusOut>', self._on_goto_entry_commit)

        ttk.Label(input_frame, text="Set Z:").grid(row=2, column=0, sticky="w"); self.goto_z_entry = ttk.Entry(input_frame, width=8, state=tk.DISABLED); self.goto_z_entry.grid(row=2, column=1, sticky="w", pady=2)
        self.goto_z_entry.bind('<Return>', self._on_goto_entry_commit); self.goto_z_entry.bind('<FocusOut>', self._on_goto_entry_commit)

        ttk.Label(input_frame, text="Set R (°):").grid(row=3, column=0, sticky="w"); self.goto_e_entry = ttk.Entry(input_frame, width=8, state=tk.DISABLED); self.goto_e_entry.grid(row=3, column=1, sticky="w", pady=2)
        self.goto_e_entry.bind('<Return>', self._on_goto_entry_commit); self.goto_e_entry.bind('<FocusOut>', self._on_goto_entry_commit)
        
        self.go_button = ttk.Button(input_frame, text="Go", command=self._go_to_position, state=tk.DISABLED, style='Primary.TButton'); self.go_button.grid(row=4, column=0, columnspan=2, pady=(10, 0), sticky="ew")

        self.go_to_center_button = ttk.Button(input_frame, text="Go to Center", command=self._go_to_center, state=tk.DISABLED)
        self.go_to_center_button.grid(row=5, column=0, columnspan=2, pady=(5, 0), sticky="ew")

        # --- Visualization Canvases ---
        canvas_size = 215
        # XY (top-down) view
        self.xy_canvas = tk.Canvas(self.canvas_frame, width=canvas_size, height=canvas_size, bg=self.COLOR_BLACK, highlightthickness=1, highlightbackground=self.COLOR_BORDER)
        self.xy_canvas.grid(row=0, column=1, sticky="n", padx=2)
        # Draw the background once and tag it so we don't have to delete it.
        self.xy_canvas.create_rectangle(0, 0, 10000, 10000, 
                                     fill=self.COLOR_BLACK, 
                                     outline=self.COLOR_BORDER, 
                                     width=1, tags="background")
        self.xy_canvas.bind("<Button-1>", self._on_xy_canvas_click); self.xy_canvas.bind("<B1-Motion>", self._on_xy_canvas_click); self.xy_canvas.bind("<Configure>", self._draw_xy_canvas_guides)
        
        # Z (side) view
        self.z_canvas = tk.Canvas(self.canvas_frame, width=25, height=canvas_size, bg=self.COLOR_BLACK, highlightthickness=1, highlightbackground=self.COLOR_BORDER); self.z_canvas.grid(row=0, column=2, sticky="ns", padx=(2, 0))
        self.z_canvas.bind("<Button-1>", self._on_z_canvas_click); self.z_canvas.bind("<B1-Motion>", self._on_z_canvas_click); self.z_canvas.bind("<Configure>", self._draw_z_canvas_marker)

        # E (Rotation) view - Circular Gauge with Snap Buttons
        # Container for the canvas and floating buttons
        self.e_container = ttk.Frame(self.canvas_frame, style='Panel.TFrame', width=canvas_size, height=canvas_size)
        self.e_container.grid(row=0, column=3, sticky="n", padx=(2, 0))
        self.e_container.grid_propagate(False) # Don't shrink
        
        self.e_canvas = tk.Canvas(self.e_container, width=canvas_size, height=canvas_size, bg=self.COLOR_BLACK, highlightthickness=1, highlightbackground=self.COLOR_BORDER)
        self.e_canvas.place(x=0, y=0, relwidth=1, relheight=1)
        self.e_canvas.bind("<Button-1>", self._on_e_canvas_click); self.e_canvas.bind("<B1-Motion>", self._on_e_canvas_click); self.e_canvas.bind("<Configure>", self._draw_e_canvas_gauge)

        # Snap Buttons (N, E, S, W)
        btn_style = 'Segment.TButton'
        # Helper to set E target
        def set_e(val):
            self.target_abs_e = float(val)
            self._update_all_displays()

        # Top (0)
        ttk.Button(self.e_container, text="0°", width=3, style=btn_style, command=lambda: set_e(0)).place(relx=0.5, rely=0.02, anchor='n')
        # Right (90)
        ttk.Button(self.e_container, text="90°", width=3, style=btn_style, command=lambda: set_e(90)).place(relx=0.98, rely=0.5, anchor='e')
        # Bottom (180)
        ttk.Button(self.e_container, text="180°", width=4, style=btn_style, command=lambda: set_e(180)).place(relx=0.5, rely=0.98, anchor='s')
        # Left (270/-90)
        ttk.Button(self.e_container, text="270°", width=4, style=btn_style, command=lambda: set_e(270)).place(relx=0.02, rely=0.5, anchor='w')

        # --- 2D Plot Toggle ---
        self.toggle_2d_button = ttk.Button(
            frame,
            text="2D TOOLPATH",
            command=self._toggle_2d_plot_button
        )
        self.toggle_2d_button.grid(row=3, column=0, columnspan=3, sticky="w", pady=(10, 0), padx=5)

        # Store controls for easy enabling/disabling based on connection state.
        self.goto_controls = [ self.goto_x_entry, self.goto_y_entry, self.goto_z_entry, self.goto_e_entry, self.go_button, 
                                 self.go_to_center_button, self.xy_canvas, self.z_canvas, self.e_canvas, self.abs_button, self.rel_button ]
        self._set_coord_mode("absolute") # Set initial mode and style
        self._update_2d_plot_button_style() # Set initial button style
        

    def create_manual_control_frame(self, parent):
        """Creates the 'MANUAL JOG CONTROL' panel for moving the printer."""
        manual_frame = ttk.LabelFrame(parent, text="MANUAL JOG CONTROL", padding="10")
        manual_frame.pack(fill=tk.X, pady=(0, 10), padx=5)
        # 5 columns: Spacer, Z-Axis, XY-Grid, E-Axis, Spacer
        manual_frame.columnconfigure(0, weight=1); manual_frame.columnconfigure(4, weight=1)
        
        # --- Z-axis jog buttons (Left Wing) ---
        z_control_frame = ttk.Frame(manual_frame, style='Panel.TFrame')
        z_control_frame.grid(row=0, column=1, rowspan=3, sticky="ns", padx=(0,15)); 
        ttk.Label(z_control_frame, text="Z-AXIS", font=self.FONT_BODY_BOLD).pack(pady=(0,5))
        
        self.jog_z_pos = ttk.Button(z_control_frame, text="Z+", command=lambda: self._jog('Z', 1), state=tk.DISABLED, style='Jog.TButton'); self.jog_z_pos.pack(pady=(2, 10), fill=tk.X)
        self.jog_z_neg = ttk.Button(z_control_frame, text="Z-", command=lambda: self._jog('Z', -1), state=tk.DISABLED, style='Jog.TButton'); self.jog_z_neg.pack(pady=(10, 2), fill=tk.X)
        
        # --- XY-axis jog buttons (Center) ---
        jog_grid_frame = ttk.Frame(manual_frame, style='Panel.TFrame')
        jog_grid_frame.grid(row=0, column=2, rowspan=3); 

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
        
        # --- E-axis jog buttons (Right Wing) ---
        e_control_frame = ttk.Frame(manual_frame, style='Panel.TFrame')
        e_control_frame.grid(row=0, column=3, rowspan=3, sticky="ns", padx=(15,0)); 
        ttk.Label(e_control_frame, text="R-AXIS", font=self.FONT_BODY_BOLD).pack(pady=(0,5))
        
        self.jog_e_pos = ttk.Button(e_control_frame, text="↻", command=lambda: self._jog('E', 1), state=tk.DISABLED, style='JogIcon.TButton'); self.jog_e_pos.pack(pady=(2, 10), fill=tk.X)
        self.jog_e_neg = ttk.Button(e_control_frame, text="↺", command=lambda: self._jog('E', -1), state=tk.DISABLED, style='JogIcon.TButton'); self.jog_e_neg.pack(pady=(10, 2), fill=tk.X)

        # --- Entry fields for jog parameters ---
        jog_params_frame = ttk.Frame(manual_frame, style='Panel.TFrame')
        jog_params_frame.grid(row=3, column=0, columnspan=5, pady=(10,0))
        
        # XYZ Params
        ttk.Label(jog_params_frame, text="XYZ Step (mm):").pack(side=tk.LEFT, padx=(0, 5))
        self.jog_step_entry = ttk.Entry(jog_params_frame, textvariable=self.jog_step_var, width=5); self.jog_step_entry.pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(jog_params_frame, text="XYZ Speed (mm/min):").pack(side=tk.LEFT, padx=(0, 5))
        self.jog_feedrate_entry = ttk.Entry(jog_params_frame, textvariable=self.jog_feedrate_var, width=6); self.jog_feedrate_entry.pack(side=tk.LEFT, padx=(0, 20))

        # E Params (Rotation)
        ttk.Label(jog_params_frame, text="Rot Step (deg):").pack(side=tk.LEFT, padx=(0, 5))
        self.rot_step_entry = ttk.Entry(jog_params_frame, textvariable=self.rotation_step_var, width=5); self.rot_step_entry.pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(jog_params_frame, text="Rot Speed (deg/min):").pack(side=tk.LEFT, padx=(0, 5))
        self.rot_feedrate_entry = ttk.Entry(jog_params_frame, textvariable=self.rotation_feedrate_var, width=6); self.rot_feedrate_entry.pack(side=tk.LEFT)
        
        self.manual_buttons = [self.home_button, self.jog_x_neg, self.jog_x_pos, self.jog_y_neg, self.jog_y_pos, self.jog_z_neg, self.jog_z_pos, self.jog_e_pos, self.jog_e_neg]
        self.manual_entries = [self.jog_step_entry, self.jog_feedrate_entry, self.rot_step_entry, self.rot_feedrate_entry]
        self._set_manual_controls_state(tk.DISABLED)

    def create_log_panel(self, parent):
        """Creates the serial log display and the manual command terminal."""
        parent.rowconfigure(0, weight=1)
        parent.columnconfigure(0, weight=1)
        # The scrolled text widget for displaying messages from the printer and application.
        self.log_area = scrolledtext.ScrolledText(parent, height=10, wrap=tk.WORD, state=tk.DISABLED,
                                                    font=self.FONT_TERMINAL,
                                                    bg=self.COLOR_BLACK,
                                                    fg=self.COLOR_ACCENT_GREEN,
                                                    bd=0,
                                                    padx=10,
                                                    pady=10)
        self.log_area.grid(row=0, column=0, sticky="nsew", padx=5, pady=(5,0))
        
        # The input area for sending single G-code commands.
        terminal_frame = ttk.Frame(parent, padding=(5, 5), style='Panel.TFrame')
        terminal_frame.grid(row=1, column=0, sticky="ew", padx=5, pady=(0, 5))
        terminal_frame.columnconfigure(0, weight=1)
        
        self.terminal_input = ttk.Entry(terminal_frame, state=tk.DISABLED, font=self.FONT_MONO)
        self.terminal_input.grid(row=0, column=0, sticky="ew", padx=(0, 5))
        
        self.terminal_send_button = ttk.Button(terminal_frame, text="Send", state=tk.DISABLED, command=self._send_from_terminal, style='Primary.TButton')
        self.terminal_send_button.grid(row=0, column=1, sticky="w")
        
        # Bind keyboard shortcuts for the terminal.
        self.terminal_input.bind('<Return>', self._send_from_terminal)
        self.terminal_input.bind('<Up>', self._history_up)
        self.terminal_input.bind('<Down>', self._history_down)

    def create_view_controls(self, parent):
        """Creates a small panel with buttons to control the 3D view."""
        if not self.matplotlib_imported:
            return

        # Use a standard tk.Frame for placing on top of the canvas, as it supports place() better
        controls_frame = tk.Frame(parent, bg=self.COLOR_PANEL_BG, bd=1, relief=tk.SOLID, highlightbackground=self.COLOR_BORDER, highlightthickness=1)
        controls_frame.place(relx=1.0, rely=0.0, x=-10, y=10, anchor='ne')

        # --- View Cube Buttons ---
        # Using a grid to simulate a view cube
        up_btn = ttk.Button(controls_frame, text="↑", command=lambda: self._rotate_view(elev_change=15), style='ViewCube.TButton')
        up_btn.grid(row=0, column=1, pady=(5, 0))

        left_btn = ttk.Button(controls_frame, text="←", command=lambda: self._rotate_view(azim_change=15), style='ViewCube.TButton')
        left_btn.grid(row=1, column=0, padx=(5, 0))
        
        top_btn = ttk.Button(controls_frame, text="Top", command=lambda: self._set_view(elev=90, azim=-90), style="Segment.TButton", width=4)
        top_btn.grid(row=1, column=1, padx=2, pady=2)
        
        right_btn = ttk.Button(controls_frame, text="→", command=lambda: self._rotate_view(azim_change=-15), style='ViewCube.TButton')
        right_btn.grid(row=1, column=2, padx=(0, 5))

        down_btn = ttk.Button(controls_frame, text="↓", command=lambda: self._rotate_view(elev_change=-15), style='ViewCube.TButton')
        down_btn.grid(row=2, column=1, pady=(0, 5))
        
        # Separator and additional view buttons
        sep = ttk.Separator(controls_frame, orient='horizontal')
        sep.grid(row=3, column=0, columnspan=3, sticky='ew', pady=2)

        btn_frame = ttk.Frame(controls_frame, style="Panel.TFrame")
        btn_frame.grid(row=4, column=0, columnspan=3, pady=(0,5), padx=5)
        btn_frame.columnconfigure((0,1), weight=1)

        front_btn = ttk.Button(btn_frame, text="Front", command=lambda: self._set_view(elev=0, azim=-90), style="Segment.TButton")
        front_btn.grid(row=0, column=0, sticky="ew")
        iso_btn = ttk.Button(btn_frame, text="Iso", command=lambda: self._set_view(elev=30, azim=-60), style="Segment.TButton")
        iso_btn.grid(row=0, column=1, sticky="ew", padx=(2,0))

    def _create_3d_plot_widgets(self, parent):
        """Creates and packs the matplotlib 3D plot widgets."""
        # This frame will hold all the 3D plot widgets
        plot_frame = ttk.Frame(parent, style="Panel.TFrame")
        plot_frame.pack(fill=tk.BOTH, expand=True)

        def deferred_load():
            """The actual import and creation logic, run after a short delay."""
            try:
                from matplotlib.figure import Figure
                from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
                from mpl_toolkits.mplot3d import Axes3D
                
                self.matplotlib_imported = True
                
                # Clear the loading message
                for widget in plot_frame.winfo_children():
                    widget.destroy()

                self.fig_3d = Figure(figsize=(5, 4), dpi=100, facecolor=self.COLOR_PANEL_BG)
                self.ax_3d = self.fig_3d.add_subplot(111, projection='3d')
                
                self.canvas_3d = FigureCanvasTkAgg(self.fig_3d, plot_frame)
                self.canvas_3d.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)
                
                self.create_view_controls(plot_frame)
                self._draw_3d_toolpath()

                opacity_frame = ttk.Frame(plot_frame, style='Panel.TFrame', padding=10)
                opacity_frame.pack(side=tk.BOTTOM, fill=tk.X)
                ttk.Label(opacity_frame, text="3D Toolpath Opacity:", font=self.FONT_BODY_SMALL).pack(side=tk.LEFT, padx=(0, 10))
                opacity_slider = ttk.Scale(opacity_frame, from_=0.0, to=1.0, variable=self.toolpath_3d_opacity_var, orient=tk.HORIZONTAL, command=lambda e: self._draw_3d_toolpath())
                opacity_slider.pack(side=tk.LEFT, fill=tk.X, expand=True)

            except ImportError:
                # If import fails, show a helpful error message.
                error_label = ttk.Label(plot_frame, text="3D view requires 'matplotlib'.\n\nPlease install it by running:\npip install matplotlib", foreground=self.COLOR_ACCENT_AMBER, justify=tk.CENTER, font=self.FONT_MONO)
                error_label.pack(fill=tk.BOTH, expand=True)

        if not self.matplotlib_imported:
            loading_label = ttk.Label(plot_frame, text="Loading 3D visualization...", justify=tk.CENTER, font=self.FONT_MONO, foreground=self.COLOR_TEXT_SECONDARY)
            loading_label.pack(fill=tk.BOTH, expand=True)
            self.root.after(50, deferred_load)
        else:
            deferred_load() # If already imported, just run it directly

    def _update_3d_plot_visibility(self):
        """Shows or hides the 3D plot based on the toggle state."""
        # Clear the container frame of any previous widgets
        for widget in self.plot_container_frame.winfo_children():
            widget.destroy()

        if self.is_3d_plot_enabled.get():
            # If enabled, create the plot widgets (which handles lazy loading)
            self._create_3d_plot_widgets(self.plot_container_frame)
        else:
            # If disabled, show a message and clear the cache to save memory
            disabled_label = ttk.Label(
                self.plot_container_frame,
                text="3D Plot disabled for performance.",
                justify=tk.CENTER,
                font=self.FONT_MONO,
                foreground=self.COLOR_TEXT_SECONDARY
            )
            disabled_label.pack(fill=tk.BOTH, expand=True, pady=20)
            self._plot_coords_cache = None
            self._plot_cache_valid = False

    def create_3d_display_panel(self, parent):
        """
        Creates the 3D toolpath visualization panel with an enable/disable toggle.
        """
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)

        # --- Top control bar with the toggle ---
        control_bar = ttk.Frame(parent, style="Panel.TFrame", padding=(5, 5))
        control_bar.grid(row=0, column=0, sticky="ew")
        
        self.toggle_3d_button = ttk.Button(
            control_bar, 
            text="3D PLOT", 
            command=self._toggle_3d_plot_button
        )
        self.toggle_3d_button.pack(side=tk.LEFT)

        # --- Container for the plot or the "disabled" message ---
        self.plot_container_frame = ttk.Frame(parent, style="Panel.TFrame")
        self.plot_container_frame.grid(row=1, column=0, sticky="nsew")
        self.plot_container_frame.rowconfigure(0, weight=1)
        self.plot_container_frame.columnconfigure(0, weight=1)

        # Set the initial state
        self._update_3d_plot_button_style()
        self._update_3d_plot_visibility()

    def _toggle_3d_plot_button(self):
        """Toggles the state of the 3D plot and updates the button style."""
        self.is_3d_plot_enabled.set(not self.is_3d_plot_enabled.get())
        self._update_3d_plot_button_style()
        self._update_3d_plot_visibility()

    def _update_3d_plot_button_style(self):
        """Applies the 'On' or 'Off' style to the 3D plot toggle button."""
        if self.is_3d_plot_enabled.get():
            self.toggle_3d_button.config(style='Custom.Toggle.On.TButton')
        else:
            self.toggle_3d_button.config(style='Custom.Toggle.Off.TButton')

    def _style_3d_plot(self):
        """Applies the custom dark theme to the 3D plot."""
        if not self.matplotlib_imported or not self.ax_3d:
            return

        # Set background colors
        self.fig_3d.patch.set_facecolor(self.COLOR_PANEL_BG)
        self.ax_3d.set_facecolor(self.COLOR_BLACK)

        # Style the grid and axis panes
        self.ax_3d.xaxis.set_pane_color((0.0, 0.0, 0.0, 0.0))
        self.ax_3d.yaxis.set_pane_color((0.0, 0.0, 0.0, 0.0))
        self.ax_3d.zaxis.set_pane_color((0.0, 0.0, 0.0, 0.0))
        
        self.ax_3d.grid(True, color=self.COLOR_ACCENT_CYAN, linestyle='-', linewidth=0.5, alpha=0.3)

        # Style the axis spines and labels
        for axis in [self.ax_3d.xaxis, self.ax_3d.yaxis, self.ax_3d.zaxis]:
            axis.line.set_visible(False) # Hide the axis lines
            axis.label.set_color(self.COLOR_TEXT_PRIMARY) # Brighter labels
            axis.set_tick_params(colors=self.COLOR_TEXT_PRIMARY) # Brighter ticks

        self.ax_3d.set_xlabel("X", color=self.COLOR_TEXT_PRIMARY, labelpad=15)
        self.ax_3d.set_ylabel("Y", color=self.COLOR_TEXT_PRIMARY, labelpad=15)
        self.ax_3d.set_zlabel("Z", color=self.COLOR_TEXT_PRIMARY, labelpad=15)
        
        self.fig_3d.tight_layout(pad=0)
        
        # Set a better initial viewing angle for less disorientation
        self.ax_3d.view_init(elev=30, azim=-60)

    def _are_collinear(self, p1, p2, p3, epsilon=1e-6):
        """Checks if three 3D points are collinear using the cross-product method."""
        v1 = (p2[0] - p1[0], p2[1] - p1[1], p2[2] - p1[2])
        v2 = (p3[0] - p1[0], p3[1] - p1[1], p3[2] - p1[2])
        
        cross_product_x = v1[1] * v2[2] - v1[2] * v2[1]
        cross_product_y = v1[2] * v2[0] - v1[0] * v2[2]
        cross_product_z = v1[0] * v2[1] - v1[1] * v2[0]
        
        # Check if the magnitude of the cross product is close to zero
        return abs(cross_product_x) < epsilon and abs(cross_product_y) < epsilon and abs(cross_product_z) < epsilon

    def _build_plot_coordinates(self):
        """
        Builds and caches the coordinate arrays for the 3D plot.
        This version simplifies the toolpath by removing redundant collinear points.
        """
        if not self.is_3d_plot_enabled.get():
            return ([], [], [])
            
        if self._plot_cache_valid and self._plot_coords_cache:
            return self._plot_coords_cache
        
        all_points = []
        if self.ordered_z_values:
            try:
                first_z = next(iter(sorted(self.toolpath_by_layer.keys())))
                first_segment = self.toolpath_by_layer[first_z][0]
                all_points.append( (first_segment[0][0], first_segment[0][1], first_z) )

                for z, index_on_layer in self.move_to_layer_map:
                    segment = self.toolpath_by_layer[z][index_on_layer]
                    all_points.append( (segment[1][0], segment[1][1], z) )
            except (IndexError, StopIteration):
                all_points = []

        if len(all_points) < 3:
            # Not enough points to simplify, proceed as before
            simplified_points = all_points
            self._simplified_indices_cache = list(range(len(all_points)))
        else:
            # --- Line Simplification using Collinearity Check ---
            # We track the original indices to map 'completed moves' to 'simplified segments'.
            simplified_points = [all_points[0]]
            self._simplified_indices_cache = [0]
            
            for i in range(1, len(all_points) - 1):
                p1 = simplified_points[-1]
                p2 = all_points[i]
                p3 = all_points[i+1]
                
                # If the current point p2 is NOT collinear with the last saved point
                # and the next point, then it's a necessary vertex.
                if not self._are_collinear(p1, p2, p3):
                    simplified_points.append(p2)
                    self._simplified_indices_cache.append(i)
            
            # Always add the very last point to complete the path
            simplified_points.append(all_points[-1])
            self._simplified_indices_cache.append(len(all_points) - 1)
            
            original_count = len(all_points)
            simplified_count = len(simplified_points)
            if original_count > simplified_count:
                self.log_message(f"3D plot optimized: Reduced {original_count} vertices to {simplified_count} ({(original_count - simplified_count) / original_count:.1%} reduction).")

        if not simplified_points:
            self._plot_coords_cache = ([], [], [])
            self._plot_cache_valid = True
            return self._plot_coords_cache

        # Use numpy for memory efficiency if available
        if HAS_NUMPY:
            coords_array = np.array(simplified_points, dtype=np.float32)
            x_coords = coords_array[:, 0]
            y_coords = coords_array[:, 1]
            z_coords = coords_array[:, 2]
        else:
            x_coords = [p[0] for p in simplified_points]
            y_coords = [p[1] for p in simplified_points]
            z_coords = [p[2] for p in simplified_points]
        
        self._plot_coords_cache = (x_coords, y_coords, z_coords)
        self._plot_cache_valid = True
        return self._plot_coords_cache

    def _draw_3d_toolpath(self):
        """Draws the full G-code toolpath on the 3D plot."""
        if not self.is_3d_plot_enabled.get() or not self.matplotlib_imported or not self.ax_3d:
            return
            
        self.ax_3d.clear()
        self.marker_3d = None # Clear the reference to the old marker artist
        self._style_3d_plot() # Re-apply style after clearing

        x_coords, y_coords, z_coords = self._build_plot_coordinates()

        if not len(x_coords):
            self.canvas_3d.draw()
            return

        # Determine where to split the line colors based on simplification map
        split_index = 0
        if self.completed_move_count > 0 and hasattr(self, '_simplified_indices_cache'):
            # Find the last simplified vertex that corresponds to a completed raw move.
            # self.completed_move_count is the number of raw moves done.
            # Corresponds to index in all_points.
            for i in range(1, len(self._simplified_indices_cache)):
                end_raw_idx = self._simplified_indices_cache[i]
                if self.completed_move_count >= end_raw_idx:
                    split_index = i
                else:
                    break

        # Plot completed segments in grey
        if split_index > 0:
            # We include point at split_index to ensure connectivity
            self.ax_3d.plot(x_coords[:split_index+1], y_coords[:split_index+1], z_coords[:split_index+1], color=self.COLOR_GREY_COMPLETED, linewidth=1, alpha=0.4)

        # Plot pending segments in cyan
        if split_index < len(x_coords) - 1:
            self.ax_3d.plot(x_coords[split_index:], y_coords[split_index:], z_coords[split_index:], color=self.COLOR_ACCENT_CYAN, linewidth=1.2, alpha=self.toolpath_3d_opacity_var.get())

        # Set fixed plot limits from PRINTER_BOUNDS
        self.ax_3d.set_xlim(self.PRINTER_BOUNDS['x_min'], self.PRINTER_BOUNDS['x_max'])
        self.ax_3d.set_ylim(self.PRINTER_BOUNDS['y_min'], self.PRINTER_BOUNDS['y_max'])
        self.ax_3d.set_zlim(self.PRINTER_BOUNDS['z_min'], self.PRINTER_BOUNDS['z_max'])

        self._update_3d_position_marker()
        self.canvas_3d.draw()

    def _update_3d_position_marker(self):
        """Updates the red dot indicating the current printer position on the 3D plot."""
        if not self.is_3d_plot_enabled.get() or not self.matplotlib_imported or not self.ax_3d:
            return

        # Remove the previous marker
        if self.marker_3d:
            self.marker_3d.remove()
            self.marker_3d = None

        # Draw the new marker if the position is known
        if self.last_cmd_abs_x is not None:
            self.marker_3d = self.ax_3d.scatter(
                [self.last_cmd_abs_x], [self.last_cmd_abs_y], [self.last_cmd_abs_z],
                color=self.COLOR_ACCENT_RED,
                s=40, # size
                depthshade=False,
                label="Current Position"
            )
        
        self.canvas_3d.draw()

    def _set_view(self, elev, azim):
        """Sets the 3D plot to a specific viewing angle."""
        if not self.matplotlib_imported or not self.ax_3d:
            return
        self.ax_3d.view_init(elev=elev, azim=azim)
        self.canvas_3d.draw()

    def _rotate_view(self, elev_change=0, azim_change=0):
        """Incrementally rotates the 3D plot."""
        if not self.matplotlib_imported or not self.ax_3d:
            return
        # Clamp elevation to avoid flipping the view too far
        new_elev = max(-90, min(90, self.ax_3d.elev + elev_change))
        new_azim = self.ax_3d.azim + azim_change
        self._set_view(new_elev, new_azim)

    def _color_blend(self, color1_hex, color2_hex, alpha):
        """
        Blends two hex colors together by an alpha value.
        
        NOTE: This method is currently unused. It was intended for future UI enhancements
        like blending toolpath colors based on speed or other parameters.
        If integrated, update its usage throughout the drawing functions (e.g., _draw_xy_canvas_guides, _draw_z_canvas_marker).
        """
        color1_rgb = tuple(int(color1_hex.lstrip('#')[i:i+2], 16) for i in (0, 2, 4))
        color2_rgb = tuple(int(color2_hex.lstrip('#')[i:i+2], 16) for i in (0, 2, 4))
        
        blended_rgb = [int(color1_rgb[i] * alpha + color2_rgb[i] * (1 - alpha)) for i in range(3)]
        
        return f"#{blended_rgb[0]:02x}{blended_rgb[1]:02x}{blended_rgb[2]:02x}"


    # --- Utility Methods ---
    def _get_available_ports(self):
        """
        Scans for and returns a list of available serial ports.
        
        It filters the list to include common serial-to-USB chipsets found
        on 3D printers and controller boards (e.g., CH340, Arduino).
        """
        ports = serial.tools.list_ports.comports()
        # Filter for common printer controller names or generic serial ports
        filtered_ports = [p for p in ports if 'CH340' in p.description or 'USB Serial' in p.description or 'Arduino' in p.description or 'Serial Port' in p.description or not p.description]
        return sorted([port.device for port in filtered_ports])


    def rescan_ports(self):
        """
        Rescans for available serial ports and updates the dropdown menu.
        
        It preserves the user's current selection if it's still available after
        the scan. It also logs the updated list of ports.
        """
        current_selection = self.port_var.get()
        self.available_ports = ["Auto-detect"] + self._get_available_ports()
        try:
            # Update the Combobox's list of values
            if self.serial_connection:
                self.port_combobox['values'] = self.available_ports
            else:
                self.port_combobox['values'] = self.available_ports
                self.port_combobox.config(state="readonly")

            # If the previously selected port is still in the list, keep it selected.
            if current_selection in self.available_ports:
                self.port_var.set(current_selection)
            else:
                # Otherwise, default to the first port in the list.
                self.port_var.set(self.available_ports[0] if self.available_ports else "")
            
            self.log_message(f"Ports updated: {', '.join(self.available_ports)}")
        except tk.TclError:
            # This can happen if the widget is destroyed during the update.
            self.log_message("Warn: Could not update port list.", "WARN")


    def _set_manual_controls_state(self, state):
        """
        Enables or disables all manual jog and homing controls.

        Args:
            state (str): The state to set the widgets to, either 'normal' or 'disabled'.
        """
        for button in self.manual_buttons:
            button.config(state=state)
        
        entry_state = tk.NORMAL if state == tk.NORMAL else tk.DISABLED
        if hasattr(self, 'manual_entries'):
            for entry in self.manual_entries:
                entry.config(state=entry_state)
        else:
            # Fallback for older initialization or if attribute missing
            self.jog_step_entry.config(state=entry_state)
            self.jog_feedrate_entry.config(state=entry_state)
        
        # Also enable/disable the 'Mark Current as Center' button
        if hasattr(self, 'mark_center_button'):
             self.mark_center_button.config(state=state)
        
        if hasattr(self, 'collision_test_button'):
             self.collision_test_button.config(state=state)


    def _set_goto_controls_state(self, state):
        """
        Enables or disables all 'Go To' position controls, including canvases.

        Args:
            state (str): The state to set the widgets to, either 'normal' or 'disabled'.
        """
        tk_state = tk.NORMAL if state == tk.NORMAL else tk.DISABLED
        
        for control in self.goto_controls:
            if isinstance(control, (ttk.Entry, ttk.Button)):
                control.config(state=tk_state)
            elif isinstance(control, tk.Canvas):
                # Canvases are not disabled but their background is changed to indicate state.
                canvas_bg = self.COLOR_BLACK if state == tk.NORMAL else '#111'
                control.config(bg=canvas_bg)
        
        self._update_all_displays() # Redraw canvases to reflect the state change.

    def _set_terminal_controls_state(self, state):
        """
        Enables or disables the manual command terminal input and send button.

        Args:
            state (str): The state to set the widgets to, either 'normal' or 'disabled'.
        """
        tk_state = tk.NORMAL if state == tk.NORMAL else tk.DISABLED
        
        if hasattr(self, 'terminal_input'):
            self.terminal_input.config(state=tk_state)
        if hasattr(self, 'terminal_send_button'):
            self.terminal_send_button.config(state=tk_state)


    def log_message(self, message, level="INFO"):
        """
        Displays a message in the log area with appropriate color-coding.

        This method is thread-safe and can be called from background threads
        via the `queue_message` helper.

        Args:
            message (str): The message string to display.
            level (str): The severity level ('INFO', 'SUCCESS', 'WARN', 'ERROR', 'CRITICAL').
                         Determines the color of the message.
        """
        # If the log area isn't created yet, print to the console as a fallback.
        if not hasattr(self, 'log_area') or not self.log_area:
            print(f"[LOG_EARLY {level}] {message}")
            return
        
        color_map = { 
            "INFO": self.COLOR_ACCENT_GREEN, 
            "SUCCESS": self.COLOR_ACCENT_CYAN, 
            "WARN": self.COLOR_ACCENT_AMBER, 
            "ERROR": self.COLOR_ACCENT_RED, 
            "CRITICAL": self.COLOR_ACCENT_RED 
        }
        timestamp_color = self.COLOR_TEXT_SECONDARY

        tag_name = level.upper()
        timestamp = time.strftime("%H:%M:%S")
        
        try:
             # The ScrolledText widget must be temporarily unlocked to add text.
             self.log_area.config(state=tk.NORMAL)
             
             # --- Log Rotation: Keep the log from growing indefinitely ---
             line_count = int(self.log_area.index('end-1c').split('.')[0])
             if line_count > 500:
                 # Delete the oldest lines, keeping the last 500
                 self.log_area.delete('1.0', f'{line_count - 499}.0')

             # Configure a tag for the timestamp with its unique color.
             ts_tag = "timestamp"
             if ts_tag not in self.log_area.tag_names():
                 self.log_area.tag_configure(ts_tag, foreground=timestamp_color)
                 
             # Configure a tag for the message level (e.g., 'ERROR') with its color.
             if tag_name not in self.log_area.tag_names():
                  self.log_area.tag_configure(tag_name, foreground=color_map.get(level, self.COLOR_ACCENT_GREEN))
                  # Make critical errors bold.
                  if level in ["ERROR", "CRITICAL"]: 
                      self.log_area.tag_configure(tag_name, font=(self.FONT_TERMINAL[0], self.FONT_TERMINAL[1], 'bold'))
             
             # Insert the timestamp and message with their respective tags.
             self.log_area.insert(tk.END, f"[{timestamp}] ", (ts_tag,))
             self.log_area.insert(tk.END, f"{message}\n", (tag_name,))
             
             self.log_area.see(tk.END) # Scroll to the bottom.
             self.log_area.config(state=tk.DISABLED) # Lock the widget again.
        except tk.TclError as e:
            # This can happen if the widget is destroyed during the update.
            print(f"[LOG_ERROR {level}] {message} (TclError: {e})")

    # --- File Handling ---

    def select_file(self):
        """
        Opens a file dialog for the user to select a G-code file.

        If a file is selected, it updates the GUI, logs the selection, and
        triggers the loading and processing of the file.
        """
        filepath = filedialog.askopenfilename(
            title="Select G-Code File",
            filetypes=[("G-code", "*.gcode"), ("Text", "*.txt"), ("All", "*.*")]
        )
        if filepath: 
            self.file_path_var.set(filepath)
            
            # Update the header to show just the filename.
            filename = filepath.split('/')[-1]
            self.header_file_var.set(filename.upper())
            
            self.log_message(f"Selected file: {filepath}")
            self.load_gcode_file(filepath)
            
            # Set default log filename based on G-code filename
            base_name = os.path.splitext(filename)[0]
            self.log_filepath_var.set(f"{base_name}_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")
            
        # Enable the start button if we are connected and have a valid file.
        if self.serial_connection and self.processed_gcode:
            self.start_button.config(state=tk.NORMAL)

    def load_gcode_file(self, filepath):
        """
        Stores the path to a G-code file and triggers processing.
        The file is read line-by-line during processing to save memory.

        Args:
            filepath (str): The path to the G-code file.
        """
        # Reset progress and disable the start button while loading.
        self.start_button.config(state=tk.DISABLED)
        self.progress_var.set(0.0)
        self.progress_label_var.set("Progress: Idle")
        
        try:
            # Store path instead of reading the whole file into memory
            self.gcode_filepath = filepath
            self.log_message(f"Loading from {filepath}...")
            self.process_gcode() # Process the file immediately.
        except Exception as e:
            self.log_message(f"Error preparing file: {e}", "ERROR")
            messagebox.showerror("Error", f"Could not prepare file:\n{e}")
            self.gcode_filepath = None
            self.processed_gcode = []

    def process_gcode(self):
        """
        Processes loaded G-code to be suitable for the printer and GUI.
        - Translates all G0/G1 moves by the Center offset to be in the printer's absolute coordinates.
        - Blocks G92 (Set Position) commands to avoid conflicting with the app's coordinate system.
        - Checks all moves against the hardcoded PRINTER_BOUNDS.
        - Generates structured toolpath data for visualization, separated by Z-level (layers).
        """
        # Reset all data structures related to the G-code file
        self.processed_gcode = []
        self.toolpath_by_layer = {}
        self.move_to_layer_map = []
        self.ordered_z_values = []
        self.completed_move_count = 0
        self._plot_cache_valid = False # Invalidate the plot cache

        if not hasattr(self, 'gcode_filepath') or not self.gcode_filepath:
            self.start_button.config(state=tk.DISABLED)
            self._update_all_displays() # Redraw canvas to clear old path
            return

        try:
            center_x = float(self.center_x_var.get())
            center_y = float(self.center_y_var.get())
            center_z = float(self.center_z_var.get())
        except ValueError:
            messagebox.showerror("Error", "Center coords must be numbers.")
            self.log_message("Error: Invalid center coords. Must be numbers.", "ERROR")
            self.start_button.config(state=tk.DISABLED)
            return

        temp_processed = []
        lines_translated = 0
        
        # --- Toolpath Processing Variables ---
        # Tracks the absolute XYZ position as we parse the file
        current_pos = {'x': None, 'y': None, 'z': None}
        # A dictionary to count moves on each layer, e.g., {10.2: 5, 10.4: 12}
        layer_move_counters = {} 

        line_number = 0
        with open(self.gcode_filepath, 'r') as f:
            for line in f:
                line_number += 1
                stripped_line = line.strip()

                if not stripped_line or stripped_line.startswith(';'):
                    temp_processed.append(line)
                    continue

                if "G92" in stripped_line.upper():
                    self.log_message(f"Warning: Blocked G92 command on line {line_number}: '{stripped_line}'", "WARN")
                    temp_processed.append(f"; Original line {line_number} blocked (G92 conflicts with app coords): {stripped_line}\n")
                    continue

                # Block heating commands (M104, M109) to ensure cold extrusion safety
                if "M104" in stripped_line.upper() or "M109" in stripped_line.upper():
                    self.log_message(f"Warning: Blocked heating command on line {line_number}: '{stripped_line}'", "WARN")
                    temp_processed.append(f"; Original line {line_number} blocked (Heating disabled): {stripped_line}\n")
                    continue
                
                # Check if the line is a move command that we need to process for the toolpath
                is_move = "G0" in stripped_line.upper() or "G1" in stripped_line.upper() or "G28" in stripped_line.upper()
                
                if is_move:
                    start_pos = current_pos.copy()
                    new_line = line # Default to original line for G28
                    
                    # --- 1. Determine the new `current_pos` based on the command ---
                    if "G28" in stripped_line.upper():
                        # A full G28 homes all axes
                        if 'X' not in stripped_line.upper() and 'Y' not in stripped_line.upper() and 'Z' not in stripped_line.upper():
                            current_pos = {'x': self.PRINTER_BOUNDS['x_min'], 'y': self.PRINTER_BOUNDS['y_min'], 'z': self.PRINTER_BOUNDS['z_min']}
                        else: # A partial G28 homes only specified axes
                            if 'X' in stripped_line.upper(): current_pos['x'] = self.PRINTER_BOUNDS['x_min']
                            if 'Y' in stripped_line.upper(): current_pos['y'] = self.PRINTER_BOUNDS['y_min']
                            if 'Z' in stripped_line.upper(): current_pos['z'] = self.PRINTER_BOUNDS['z_min']

                    else: # G0 or G1 move
                        parsed_coords = self._parse_gcode_coords(stripped_line)
                        if not parsed_coords:
                            temp_processed.append(line)
                            continue

                        # The G-code file's coordinates are relative to its own origin (0,0,0).
                        # We translate them into the printer's absolute coordinate space.
                        # If an axis is not specified in the command, we assume it stays at its last known position.
                        rel_x = parsed_coords.get('x')
                        rel_y = parsed_coords.get('y')
                        rel_z = parsed_coords.get('z')
                        
                        abs_x = rel_x + center_x if rel_x is not None else current_pos.get('x')
                        abs_y = rel_y + center_y if rel_y is not None else current_pos.get('y')
                        abs_z = rel_z + center_z if rel_z is not None else current_pos.get('z')

                        # On the very first move, the start position might be unknown.
                        # We establish the position but cannot draw a path segment yet.
                        if any(v is None for v in [abs_x, abs_y, abs_z]):
                            if start_pos['x'] is None: # This is the very first move, it's ok if it's partial
                                current_pos.update({k:v for k,v in {'x':abs_x, 'y':abs_y, 'z':abs_z}.items() if v is not None})
                            else: # This is a subsequent move with missing axes, which is an error
                                err_msg = f"G-code line {line_number} ({stripped_line}) has an unknown coordinate. Ensure G28 or a G0/G1 move with all X,Y,Z axes is near the start of the file."
                                self.log_message(err_msg, "ERROR"); messagebox.showerror("Processing Error", err_msg)
                                # Clear all data and abort
                                self.processed_gcode, self.toolpath_by_layer, self.move_to_layer_map, self.ordered_z_values = [], {}, [], []
                                self.start_button.config(state=tk.DISABLED)
                                return
                        else:
                            current_pos = {'x': abs_x, 'y': abs_y, 'z': abs_z}

                        # --- Bounds Check ---
                        if not (self.PRINTER_BOUNDS['x_min'] <= abs_x <= self.PRINTER_BOUNDS['x_max']):
                            err_msg = f"G-code line {line_number} results in X out-of-bounds ({abs_x:.2f}). Aborting."
                            self.log_message(err_msg, "ERROR"); messagebox.showerror("Processing Error", err_msg); return
                        if not (self.PRINTER_BOUNDS['y_min'] <= abs_y <= self.PRINTER_BOUNDS['y_max']):
                            err_msg = f"G-code line {line_number} results in Y out-of-bounds ({abs_y:.2f}). Aborting."
                            self.log_message(err_msg, "ERROR"); messagebox.showerror("Processing Error", err_msg); return
                        if not (self.PRINTER_BOUNDS['z_min'] <= abs_z <= self.PRINTER_BOUNDS['z_max']):
                            err_msg = f"G-code line {line_number} results in Z out-of-bounds ({abs_z:.2f}). Aborting."
                            self.log_message(err_msg, "ERROR"); messagebox.showerror("Processing Error", err_msg); return

                        # Reconstruct the G-code line with the new absolute coordinates
                        f_match = re.search(r"F(\d+(\.\d+)?)", stripped_line)
                        e_match = re.search(r"E([-+]?\d*\.?\d+)", stripped_line)
                        new_line_parts = [stripped_line.split()[0]]
                        if rel_x is not None: new_line_parts.append(f"X{abs_x:.3f}")
                        if rel_y is not None: new_line_parts.append(f"Y{abs_y:.3f}")
                        if rel_z is not None: new_line_parts.append(f"Z{abs_z:.3f}")
                        if e_match: new_line_parts.append(e_match.group(0))
                        if f_match: new_line_parts.append(f_match.group(0))
                        new_line = " ".join(new_line_parts) + "\n"
                        lines_translated += 1

                    # --- 2. Add segment to our toolpath data structures (if we have a start and end) ---
                    if start_pos['x'] is not None:
                        z_level = current_pos['z']
                        segment = ((start_pos['x'], start_pos['y']), (current_pos['x'], current_pos['y']))

                        # Add the segment to the correct layer in the dictionary
                        if z_level not in self.toolpath_by_layer:
                            self.toolpath_by_layer[z_level] = []
                        self.toolpath_by_layer[z_level].append(segment)
                        
                        # Get the index of this move *within its layer*
                        index_on_layer = layer_move_counters.get(z_level, 0)
                        
                        # Map the global move index to its layer and its index on that layer
                        self.move_to_layer_map.append((z_level, index_on_layer))
                        
                        # Add the Z value to the ordered list for the Z-canvas graph
                        self.ordered_z_values.append(z_level)

                        # Increment the counter for this specific layer
                        layer_move_counters[z_level] = index_on_layer + 1

                    temp_processed.append(new_line)
                
                else: # Not a move command, just add it to the processed list
                    temp_processed.append(line)

        self.processed_gcode = temp_processed
        self.log_message(f"G-code processed. {lines_translated} moves translated. Toolpath data generated for {len(self.toolpath_by_layer)} layers.", "SUCCESS")
        
        if self.serial_connection:
            self.start_button.config(state=tk.NORMAL)
        
        self._update_all_displays() # Redraw canvas with new path
        self._draw_3d_toolpath() # Draw the new 3D toolpath


    # --- Connection Handling ---

    def toggle_connection(self):
        """Connects or disconnects the printer based on the current state."""
        if self.serial_connection:
            self.disconnect_printer()
        else:
            self.connect_printer()

    def connect_printer(self):
        """
        Initiates the serial connection process.

        It reads the port and baud rate from the GUI, updates the UI to a
        'connecting' state, and starts a background thread to handle the
        actual connection attempt.
        """
        selected_port = self.port_var.get()
        try:
            baudrate = int(self.baud_var.get())
        except ValueError:
            self.log_message("Error: Invalid Baud Rate.", "ERROR")
            messagebox.showerror("Error", "Baud Rate must be valid.")
            return
            
        self.log_message(f"Connecting... Port: {selected_port}, Baud: {baudrate}...")
        
        # --- Update GUI to "Connecting" state ---
        self.connect_button.config(state=tk.DISABLED)
        self.cancel_connect_button.grid(row=0, column=6, rowspan=2, sticky="ns", padx=(5,0))
        self.cancel_connect_button.config(state=tk.NORMAL)
        self.port_combobox.config(state=tk.DISABLED)
        self.baud_entry.config(state=tk.DISABLED)
        
        self.connection_status_var.set("Connecting...")
        self.status_indicator.set_status("busy")
        self.header_status_indicator.set_status("busy")
        self.footer_status_var.set(f"Connecting...")
        
        self._set_terminal_controls_state(tk.DISABLED)

        # Start the connection thread
        self.cancel_connect_event.clear()
        threading.Thread(target=self._connect_thread, args=(selected_port, baudrate), daemon=True).start()

    def _cancel_connection_attempt(self):
        """
        Cancels an in-progress connection attempt.

        This sets an event that the background connection thread checks,
        causing it to abort its process.
        """
        self.log_message("Connection attempt cancelled by user.", "WARN")
        self.cancel_connect_event.set()
        
        # Hide the cancel button and update the status display.
        self.cancel_connect_button.grid_remove()
        self.cancel_connect_button.config(state=tk.DISABLED)
        
        self.connection_status_var.set("Cancelling...")
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
            elif serial_conn and found_port: 
                # Send M302 P1 S0 to explicitly allow cold extrusion regardless of temp
                self.queue_message("Sending M302 P1 S0 (Allow Cold Extrusion)...")
                try:
                    serial_conn.write(b'M302 P1 S0\n')
                    serial_conn.flush()
                    m302_ok = False
                    m302_start_time = time.time()
                    m302_response_buffer = ""
                    while time.time() - m302_start_time < 5.0: # Short timeout for M302
                        if serial_conn.in_waiting > 0:
                            m302_response_buffer += serial_conn.read(serial_conn.in_waiting).decode('utf-8', errors='ignore')
                            if 'ok' in m302_response_buffer.lower():
                                m302_ok = True
                                break
                        time.sleep(0.05)
                    if not m302_ok:
                        self.queue_message("Warning: M302 P1 S0 'ok' not received. Cold extrusion might still be enabled.", "WARN")
                    else:
                        self.queue_message("M302 P1 S0 confirmed.", "SUCCESS")
                except Exception as e:
                    self.queue_message(f"Error sending M302 P1 S0: {e}", "ERROR")

                self.message_queue.put(("CONNECTED", (serial_conn, found_port, baudrate)))
            else:
                if not self.cancel_connect_event.is_set(): self.message_queue.put(("CONNECT_FAIL", "No responsive printer found."))
                else: self.message_queue.put(("CONNECT_CANCELLED", None))
        finally: self.message_queue.put(("CONNECT_ATTEMPT_FINISHED", None))

    def disconnect_printer(self, silent=False):
        """
        Closes the serial connection and resets the GUI to a disconnected state.

        Args:
            silent (bool): If True, suppresses logging of disconnection errors.
                           Useful when called during a forced shutdown.
        """
        # If a connection is in progress, cancel it instead.
        if self.connect_button['state'] == tk.DISABLED and not self.serial_connection and hasattr(self, 'cancel_connect_button') and self.cancel_connect_button.winfo_ismapped():
             self.log_message("Disconnect during connect - Cancelling.", "WARN")
             self.cancel_connect_event.set()
             return
             
        # Prevent disconnection while a job is running.
        if self.is_sending or self.is_manual_command_running:
            self.log_message("Cannot disconnect while busy.", "WARN")
            messagebox.showwarning("Busy", "Please stop the current operation before disconnecting.")
            return
            
        if self.serial_connection:
            try:
                self.serial_connection.close()
            except Exception as e:
                if not silent:
                    self.log_message(f"Disconnect error: {e}", "ERROR")
            
        self.serial_connection = None
        
        # --- Reset GUI to "Disconnected" state ---
        self.connection_status_var.set("Disconnected")
        self.status_indicator.set_status("off")
        self.header_status_indicator.set_status("off")
        self.footer_status_var.set("COM: -- @ --")
        
        self.connect_button.config(text="Connect", state=tk.NORMAL)
        self.port_combobox.config(state="readonly")
        self.baud_entry.config(state=tk.NORMAL)
        
        self.start_button.config(state=tk.DISABLED)
        self._set_manual_controls_state(tk.DISABLED)
        self._set_goto_controls_state(tk.DISABLED)
        self._set_terminal_controls_state(tk.DISABLED)

        if hasattr(self, 'cancel_connect_button'):
            self.cancel_connect_button.grid_remove()
            self.cancel_connect_button.config(state=tk.DISABLED)
        
        self.progress_var.set(0.0)
        self.progress_label_var.set("Progress: Idle")

        # Reset the known position and update the display.
        self.last_cmd_abs_x, self.last_cmd_abs_y, self.last_cmd_abs_z = None, None, None
        self._update_all_displays()


    # --- G-Code Sending & Control ---

    def start_sending(self):
        """
        Starts sending the processed G-code file to the printer.

        It performs several pre-flight checks, updates the GUI to a 'sending'
        state, and starts the background thread that handles the line-by-line
        sending of G-code.
        """
        if not self.serial_connection:
            messagebox.showerror("Error", "Not connected to a printer.")
            return
            
        # Reprocess the G-code to ensure it's up-to-date with any center changes.
        self.process_gcode() 
        
        if not self.processed_gcode:
            messagebox.showerror("Error", "No valid G-code to send. Check file, center coordinates, and printer bounds.")
            return
            
        if self.is_sending or self.is_manual_command_running:
            messagebox.showwarning("Warning", "Printer is already busy with another operation.")
            return
            
        # Count only the lines that will actually be sent.
        self.total_lines_to_send = len([line for line in self.processed_gcode if line.strip() and not line.strip().startswith(';')])
        if self.total_lines_to_send == 0:
            messagebox.showwarning("Warning", "The G-code file contains no sendable commands.")
            return
            
        # --- Update GUI to "Sending" state ---
        self.progress_var.set(0.0)
        self.progress_label_var.set(f"0/{self.total_lines_to_send} lines")
        
        self.is_sending = True
        self.is_paused = False
        self.stop_event.clear()
        self.pause_event.set() # The pause_event is set to allow sending to start.
        
        self.start_button.config(state=tk.DISABLED)
        self.pause_resume_button.config(text="Pause", state=tk.NORMAL)
        
        self._set_manual_controls_state(tk.DISABLED)
        self._set_goto_controls_state(tk.DISABLED)
        self._set_terminal_controls_state(tk.DISABLED)

        self.log_message("Starting G-code stream...")
        
        # Start the sender thread.
        threading.Thread(target=self.gcode_sender_thread, args=(list(self.processed_gcode),), daemon=True).start()

    def toggle_pause_resume(self):
        """
        Pauses or resumes the current G-code sending operation.

        This works for both file sending and manual commands. When paused,
        it enables manual controls to allow for jogging and other adjustments.
        When resumed, it disables them again.
        """
        # This button is only active when a file is sending or a manual command is running.
        if not self.is_sending and not self.is_manual_command_running: 
             return
         
        if self.is_paused:
            # --- RESUME ---
            self.pause_event.set() # This unblocks the sender thread.
            self.is_paused = False
            self.pause_resume_button.config(text="Pause")
            self.log_message("Resumed.", "INFO")
            
            # Re-disable manual controls now that the automated process is running again.
            self._set_manual_controls_state(tk.DISABLED)
            self._set_goto_controls_state(tk.DISABLED)
            self._set_terminal_controls_state(tk.DISABLED)
            
            # Set status indicators back to "on" (green)
            self.status_indicator.set_status("on")
            self.header_status_indicator.set_status("on")
        else:
            # --- PAUSE ---
            self.pause_event.clear() # This blocks the sender thread.
            self.is_paused = True
            self.pause_resume_button.config(text="Resume")
            self.log_message("Pausing... Manual controls enabled.", "INFO")
            
            # Enable manual controls to allow for jogging, etc., during the pause.
            self._set_manual_controls_state(tk.NORMAL)
            self._set_goto_controls_state(tk.NORMAL)
            self._set_terminal_controls_state(tk.NORMAL)
               
            # Set status indicators to "busy" (amber) to show it's paused.
            self.status_indicator.set_status("busy")
            self.header_status_indicator.set_status("busy")

    def emergency_stop(self):
        """
        Triggers an immediate, hard stop of the printer (M112).

        This sets the stop event to halt any running threads, sends M112 to the
        printer (which usually requires a physical reset), and disconnects the
        application from the serial port.
        """
        # If trying to stop while a connection is being attempted, cancel the connection.
        if self.connect_button['state'] == tk.DISABLED and not self.serial_connection and hasattr(self, 'cancel_connect_button') and self.cancel_connect_button.winfo_ismapped():
             self.log_message("Stop during connection - Cancelling.", "WARN")
             self._cancel_connection_attempt()
             return
             
        if not self.serial_connection:
             self.log_message("Emergency Stop called but not connected.", "WARN")
             self._reset_gui_after_stop()
             return
        
        self.log_message("!!! EMERGENCY STOP triggered !!!", "CRITICAL")
        
        # Un-pause and then stop all running threads.
        self.pause_event.set()
        self.stop_event.set()
        
        # Update status indicators to "error" (red).
        self.status_indicator.set_status("error")
        self.header_status_indicator.set_status("error")

        if self.serial_connection:
            try:
                self.log_message("Sending M112...")
                self.serial_connection.write(b'M112\n')
                time.sleep(0.5) # Give the command a moment to send
                self.serial_connection.reset_input_buffer()
                self.log_message("M112 sent.")
            except Exception as e:
                self.log_message(f"Error sending M112: {e}", "ERROR")
            finally:
                # M112 requires a printer reset, so we must disconnect.
                self.disconnect_printer(silent=True)
                messagebox.showwarning("Emergency Stop", "M112 sent.\nPrinter requires reset.\nConnection closed.")
                
        self._reset_gui_after_stop(reset_position=True)

    def quick_stop(self):
        """
        Triggers a soft stop of the printer (M410).

        This stops the application's sender thread and sends M410 to the printer,
        which should cause it to finish its last buffered move and then halt.
        Manual controls are enabled after this.
        """
        # If trying to stop while a connection is being attempted, cancel the connection.
        if self.connect_button['state'] == tk.DISABLED and not self.serial_connection and hasattr(self, 'cancel_connect_button') and self.cancel_connect_button.winfo_ismapped():
             self.log_message("Stop during connection - Cancelling.", "WARN")
             self._cancel_connection_attempt()
             return

        if not self.serial_connection:
             self.log_message("Quick Stop called but not connected.", "WARN")
             self._reset_gui_after_stop(reset_position=False)
             return
        
        self.log_message("Quick Stop requested.", "WARN")
        
        # Un-pause and then stop all running threads.
        self.pause_event.set()
        self.stop_event.set()
        
        # Update status indicators to "busy" (amber).
        self.status_indicator.set_status("busy")
        self.header_status_indicator.set_status("busy")

        if self.serial_connection:
            try:
                self.log_message("Sending M410 (Quick Stop)...")
                self.serial_connection.write(b'M410\n')
                time.sleep(0.1)
                self.serial_connection.reset_input_buffer()
                self.log_message("M410 sent.")
            except Exception as e:
                self.log_message(f"Error sending M410: {e}", "ERROR")
            
            messagebox.showinfo("Quick Stop", "M410 sent. Printer will stop after clearing its buffer.\nManual controls are now enabled.")
            
        self._reset_gui_after_stop(reset_position=False)


    def _reset_gui_after_stop(self, reset_position=True):
        """
        Resets the GUI to a safe, idle state after a stop or finished job.

        This re-enables controls, resets progress bars, and optionally resets
        the last known position to the origin.

        Args:
            reset_position (bool): If True (e.g., after an M112), the last known
                                   position is reset to the printer's origin.
                                   If False (e.g., after M410), the last known
                                   position is preserved.
        """
        self.is_sending, self.is_paused, self.is_manual_command_running = False, False, False
        
        # Reset button states
        self.start_button.config(state=tk.DISABLED)
        self.pause_resume_button.config(text="Pause", state=tk.DISABLED)
        
        # Determine the state of controls based on whether we are still connected.
        current_state = tk.NORMAL if self.serial_connection else tk.DISABLED
        self._set_manual_controls_state(current_state)
        self._set_goto_controls_state(current_state)
        self._set_terminal_controls_state(current_state)

        # Hide the cancel connection button if it's visible.
        if hasattr(self, 'cancel_connect_button'):
            self.cancel_connect_button.grid_remove()
            self.cancel_connect_button.config(state=tk.DISABLED)
            
        self.progress_var.set(0.0)
        self.progress_label_var.set("Progress: Stopped")
        
        if reset_position:
            # After a hard stop, the printer's position is lost. Reset it to origin.
            self.last_cmd_abs_x = self.PRINTER_BOUNDS['x_min']
            self.last_cmd_abs_y = self.PRINTER_BOUNDS['y_min']
            self.last_cmd_abs_z = self.PRINTER_BOUNDS['z_min']
            
        self._update_all_displays()


    # --- Thread Workers ---
    def _parse_gcode_coords(self, gcode_line):
        """Extracts X, Y, Z, and E coordinates from a G0/G1/G92 command line."""
        coords = {}
        # Regex to capture X, Y, Z, and E (case insensitive)
        # Group 1: X value, Group 2: Y value, Group 3: Z value, Group 4: E value
        match = re.search(r"^[Gg](?:[01]|92).*?(?:[Xx]([-+]?\d*\.?\d+))?.*?(?:[Yy]([-+]?\d*\.?\d+))?.*?(?:[Zz]([-+]?\d*\.?\d+))?.*?(?:[Ee]([-+]?\d*\.?\d+))?", gcode_line)
        
        # Note: The above regex is position dependent if not careful. 
        # A safer way is to find all matches or separate searches.
        # Let's use individual searches for robustness against parameter order (e.g. G1 E10 X5).
        
        x_match = re.search(r"[Xx]([-+]?\d*\.?\d+)", gcode_line)
        if x_match: coords['x'] = float(x_match.group(1))
        
        y_match = re.search(r"[Yy]([-+]?\d*\.?\d+)", gcode_line)
        if y_match: coords['y'] = float(y_match.group(1))
        
        z_match = re.search(r"[Zz]([-+]?\d*\.?\d+)", gcode_line)
        if z_match: coords['z'] = float(z_match.group(1))
        
        e_match = re.search(r"[Ee]([-+]?\d*\.?\d+)", gcode_line)
        if e_match: coords['e'] = float(e_match.group(1))
            
        return coords


    def _send_manual_command_thread(self, command):
        """
        The background worker thread for sending a manual G-code command.

        This handles sending the command, waiting for an 'ok' response, and
        updating the GUI via the message queue. It supports multi-line commands
        and respects the pause/stop events.

        Args:
            command (str): The G-code command string to send. Can contain newlines.
        """
        self.is_manual_command_running = True
        self.queue_message(f"Sending: {command.replace(chr(10), '; ')}")
        success = False
        
        # Keep track of the target position to update the GUI accurately after the move.
        target_pos = {'x': self.last_cmd_abs_x, 'y': self.last_cmd_abs_y, 'z': self.last_cmd_abs_z}
        
        try:
            lines = [line.strip() for line in command.splitlines() if line.strip()]
            in_relative_mode = False # Assume absolute mode unless G91 is seen
            
            try:
                current_target = {'x': self.last_cmd_abs_x, 'y': self.last_cmd_abs_y, 'z': self.last_cmd_abs_z}
            except (ValueError, TypeError):
                current_target = {'x': 0.0, 'y': 0.0, 'z': 0.0} # Default if position is unknown

            for line in lines:
                # This will block the thread if the pause event is cleared (paused).
                self.pause_event.wait()
                
                if self.stop_event.is_set():
                    raise InterruptedError("Stop event set before sending line.")
                    
                # Track G90/G91 for correct coordinate handling, though jogs now send absolute.
                if "G90" in line.upper(): in_relative_mode = False
                elif "G91" in line.upper(): in_relative_mode = True
                
                # Predict the destination of the move command.
                if line.upper().startswith("G0") or line.upper().startswith("G1"):
                        parsed = self._parse_gcode_coords(line)
                        if not in_relative_mode: # G90 Absolute
                            if 'x' in parsed: current_target['x'] = parsed.get('x')
                            if 'y' in parsed: current_target['y'] = parsed.get('y')
                            if 'z' in parsed: current_target['z'] = parsed.get('z')
                            if 'e' in parsed: current_target['e'] = parsed.get('e')
                        # Note: Relative (G91) jogging is handled by the _jog method which sends absolute commands.
                elif "G28" in line.upper(): # Homing command
                    if 'X' not in line.upper() and 'Y' not in line.upper() and 'Z' not in line.upper():
                        current_target = {'x': self.PRINTER_BOUNDS['x_min'], 'y': self.PRINTER_BOUNDS['y_min'], 'z': self.PRINTER_BOUNDS['z_min']}
                    else: 
                        if 'X' in line.upper(): current_target['x'] = self.PRINTER_BOUNDS['x_min']
                        if 'Y' in line.upper(): current_target['y'] = self.PRINTER_BOUNDS['y_min']
                        if 'Z' in line.upper(): current_target['z'] = self.PRINTER_BOUNDS['z_min']
                
                # --- Send the line and wait for 'ok' ---
                self.serial_connection.write(line.encode('utf-8') + b'\n')
                self.queue_message(f"Sent: {line}")
                ok_received = False
                response_buffer = ""
                timeout = 90.0 if "G28" in line.upper() else 20.0 # Homing can take a long time
                start_time = time.time()
                
                while time.time() - start_time < timeout:
                    if self.stop_event.is_set():
                        raise InterruptedError("Stop event set while waiting for 'ok'.")

                    try:
                        if self.serial_connection.in_waiting > 0:
                            response_buffer += self.serial_connection.read(self.serial_connection.in_waiting).decode('utf-8', errors='ignore')
                            # Process all complete lines in the buffer.
                            while '\n' in response_buffer:
                                full_line, response_buffer = response_buffer.split('\n', 1)
                                full_line = full_line.strip()
                                if full_line:
                                    self.queue_message(f"Received: {full_line}")
                                if 'ok' in full_line.lower():
                                    ok_received = True
                                    break
                        if ok_received:
                            break
                        time.sleep(0.05) # Small delay to prevent busy-waiting
                    except serial.SerialException as read_err:
                        self.queue_message(f"Serial read error: {read_err}", "ERROR")
                        raise
                        
                if ok_received:
                    self.queue_message(f"Line '{line}' confirmed.", "SUCCESS")
                else:
                    self.queue_message(f"Warning: No 'ok' received for '{line}' (timeout: {timeout:.1f}s).", "WARN")
                    success = False
                    break
            
            if ok_received or not lines:
                 self.queue_message("Manual command completed.", "SUCCESS")
                 success = True
                 if any(v is not None for v in current_target.values()):
                       target_pos = current_target
                       
        except InterruptedError:
            self.queue_message("Manual command interrupted by user.", "WARN")
            success = False
        except serial.SerialException as e:
            self.queue_message(f"Serial error during manual command: {e}", "ERROR")
            success = False
            self.message_queue.put(("CONNECTION_LOST", None))
        except Exception as e:
            self.queue_message(f"Unexpected error sending command '{command}': {e}", "ERROR")
            success = False
        finally:
             # If the command was successful and resulted in a new position, update the GUI.
             if success and any(v is not None for v in target_pos.values()):
                 self.message_queue.put(("POSITION_UPDATE", target_pos))
             # Signal that the manual command is finished.
             self.message_queue.put(("MANUAL_FINISHED", success))


    def _send_manual_command(self, command):
        """
        Prepares and initiates the sending of a manual G-code command.

        This method acts as a gatekeeper, checking if the printer is connected
        and not busy. It then updates the GUI to a 'busy' state and starts the
        background thread that does the actual sending.

        Args:
            command (str): The G-code command to send.
        """
        if not self.serial_connection:
            messagebox.showerror("Error", "Not connected to a printer.")
            return
        if self.is_sending or self.is_manual_command_running:
            messagebox.showwarning("Busy", "Printer is busy with another operation.")
            return
        
        # --- Update GUI to "Busy" state for a manual command ---
        self.is_manual_command_running = True
        self.stop_event.clear()
        self.pause_event.set() # Ensure the thread doesn't start in a paused state.

        # Disable other controls while the manual command is running.
        self._set_manual_controls_state(tk.DISABLED)
        self._set_goto_controls_state(tk.DISABLED)
        self._set_terminal_controls_state(tk.DISABLED)
        self.start_button.config(state=tk.DISABLED)
        
        # The pause button is enabled to allow pausing the manual command.
        self.pause_resume_button.config(text="Pause", state=tk.NORMAL)
        
        # Start the sender thread.
        threading.Thread(target=self._send_manual_command_thread, args=(command,), daemon=True).start()

    def _send_from_terminal(self, event=None):
        """
        Handles sending a command from the manual terminal input.

        It retrieves the text from the input box, adds it to the command
        history, clears the input, and then calls the `_send_manual_command`
        method to dispatch it.
        """
        command = self.terminal_input.get().strip()
        if not command:
            return

        # Add the command to history if it's new.
        if not self.command_history or self.command_history[-1] != command:
            self.command_history.append(command)
        self.history_index = len(self.command_history) # Reset history navigation
            
        self.log_message(f"Terminal > {command}", "INFO")
        self.terminal_input.delete(0, tk.END)
        
        # The _send_manual_command method handles all state checks.
        self._send_manual_command(command)

    def _jog(self, axis, direction):
        """
        Executes a manual jog move in a given direction.

        It calculates the new target position based on the current 'last commanded'
        position and the user-defined step size. It then sends an absolute G1
        move command to the printer.

        Args:
            axis (str): The axis to move ('X', 'Y', 'Z', or 'E').
            direction (int): The direction to move (-1 or 1).
        """
        try:
            if axis == 'E':
                step = float(self.rotation_step_var.get())
                feedrate = float(self.rotation_feedrate_var.get())
            else:
                step = float(self.jog_step_var.get())
                feedrate = float(self.jog_feedrate_var.get())
        except ValueError:
            messagebox.showerror("Error", "Invalid step or feedrate. Must be numbers.")
            return
            
        if step <= 0 or feedrate <= 0:
            messagebox.showerror("Error", "Step and feedrate must be positive.")
            return
            
        try:
            # Jogging is always relative to the last *commanded* position.
            current_x = self.last_cmd_abs_x if self.last_cmd_abs_x is not None else self.PRINTER_BOUNDS['x_min']
            current_y = self.last_cmd_abs_y if self.last_cmd_abs_y is not None else self.PRINTER_BOUNDS['y_min']
            current_z = self.last_cmd_abs_z if self.last_cmd_abs_z is not None else self.PRINTER_BOUNDS['z_min']
            current_e = self.last_cmd_abs_e if self.last_cmd_abs_e is not None else 0.0
            
            new_x, new_y, new_z, new_e = current_x, current_y, current_z, current_e
            
            if axis == 'X': new_x += direction * step
            elif axis == 'Y': new_y += direction * step
            elif axis == 'Z': new_z += direction * step
            elif axis == 'E': new_e += direction * step
            
            # Clamp the new target position to within the printer's physical bounds.
            new_x = max(self.PRINTER_BOUNDS['x_min'], min(self.PRINTER_BOUNDS['x_max'], new_x))
            new_y = max(self.PRINTER_BOUNDS['y_min'], min(self.PRINTER_BOUNDS['y_max'], new_y))
            new_z = max(self.PRINTER_BOUNDS['z_min'], min(self.PRINTER_BOUNDS['z_max'], new_z))
            new_e = max(self.PRINTER_BOUNDS['e_min'], min(self.PRINTER_BOUNDS['e_max'], new_e))
            
            # Update the internal 'target' model (the blue marker).
            self.target_abs_x = new_x
            self.target_abs_y = new_y
            self.target_abs_z = new_z
            self.target_abs_e = new_e
            self._update_all_displays() # Update GUI fields and canvas markers.
            
            # Send the jog move as an absolute G90 command.
            command = f"G90\nG1 X{new_x:.3f} Y{new_y:.3f} Z{new_z:.3f} E{new_e:.3f} F{feedrate:.0f}"
            self._send_manual_command(command)

        except ValueError as e:
             self.log_message(f"Could not parse current position for jog: {e}", "WARN")

    def _home_all(self):
        """
        Starts the automated homing sequence:
        1. G28 (Standard Home)
        2. Z-Max Probing (Calibrate Max Height)
        """
        if not self.serial_connection:
            messagebox.showerror("Error", "Not connected.")
            return

        if self.is_sending or self.is_manual_command_running:
            messagebox.showwarning("Busy", "Printer is busy.")
            return

        # Disable controls during the sequence
        self.is_manual_command_running = True
        self._set_manual_controls_state(tk.DISABLED)
        self._set_goto_controls_state(tk.DISABLED)
        self.start_button.config(state=tk.DISABLED)
        
        self.log_message("Starting Auto-Homing & Calibration...", "INFO")
        threading.Thread(target=self._homing_sequence_worker, daemon=True).start()

    def _homing_sequence_worker(self):
        """
        Background thread for the full homing + calibration sequence.
        1. G28 (Home Z-Min to 0)
        2. Step-and-Check probe to find Z-Max
        3. Update PRINTER_BOUNDS['z_max']
        """
        self.is_calibrating = True
        try:
            self.message_queue.put(("SET_STATUS", "busy"))
            
            # --- Step 1: Standard G28 Homing ---
            self.queue_message("Homing (G28)...")
            if self.serial_connection:
                self.serial_connection.reset_input_buffer()
                self.serial_connection.write(b'G28\n')
            else:
                raise Exception("Serial connection lost.")
            
            # Wait for 'ok' (G28 can take time)
            if not self._wait_for_ok(timeout=120):
                raise Exception("G28 Homing timeout.")
            
            # Update internal position to Origin
            self.message_queue.put(("POSITION_UPDATE", {'x': 0.0, 'y': 0.0, 'z': 0.0, 'e': 0.0}))
            
            # --- Step 2: Z-Max Probing (Step-and-Check) ---
            if not HAS_GPIO: 
                # If no GPIO, we can't probe. Set a safe default?
                # For now, just warn and keep the default.
                self.queue_message("No GPIO detected. Z-Max calibration skipped.", "WARN")
                self.message_queue.put(("SET_STATUS", "on"))
                return

            self.queue_message("Calibrating Z-Max (Step-and-Check)...")
            
            # Disable Safety Interrupt during intentional probing
            try:
                GPIO.remove_event_detect(Z_MAX_LIMIT_PIN)
            except Exception:
                pass
            
            try:
                # Switch to Relative Mode
                self.serial_connection.write(b'G91\n')
                self._wait_for_ok()

                z_limit_hit = False
                
                # --- Phase A: Coarse Seek (1mm steps) ---
                # Max travel: 300mm (safety limit)
                for _ in range(300):
                    if self.stop_event.is_set(): raise InterruptedError("Stopped")
                    
                    # Check switch BEFORE move
                    if GPIO.input(Z_MAX_LIMIT_PIN) == 0:
                        z_limit_hit = True
                        # STOP IMMEDIATELY to kill buffered moves
                        self.serial_connection.write(b'M410\n')
                        time.sleep(0.2)
                        self.serial_connection.reset_input_buffer()
                        break
                    
                    # Move 1mm up
                    self.serial_connection.write(b'G1 Z1 F600\n') 
                    self._wait_for_ok()
                    
                    # Sync: Wait for move to physically finish
                    self.serial_connection.write(b'M400\n')
                    self._wait_for_ok()
                
                if not z_limit_hit:
                    # One last check
                    if GPIO.input(Z_MAX_LIMIT_PIN) == 0: 
                        z_limit_hit = True
                        self.serial_connection.write(b'M410\n') # Stop if hit at end
                        time.sleep(0.1)
                        self.serial_connection.reset_input_buffer()

                if not z_limit_hit:
                    # Failed to find top
                    self.serial_connection.write(b'G90\n') # Safety restore
                    raise Exception("Z-Max switch not found within 300mm.")
                
                # --- Phase B: Back-off ---
                self.queue_message("Switch hit. Backing off...")
                self.serial_connection.write(b'G1 Z-5 F600\n')
                self._wait_for_ok()
                time.sleep(0.5)

                # --- Phase C: Fine Seek (0.1mm steps) ---
                z_limit_hit = False
                for _ in range(60): # 6mm range
                    if self.stop_event.is_set(): raise InterruptedError("Stopped")
                    
                    if GPIO.input(Z_MAX_LIMIT_PIN) == 0:
                        z_limit_hit = True
                        # STOP IMMEDIATELY
                        self.serial_connection.write(b'M410\n')
                        time.sleep(0.1)
                        self.serial_connection.reset_input_buffer()
                        break
                        
                    self.serial_connection.write(b'G1 Z0.1 F60\n')
                    self._wait_for_ok()
                    
                    # Sync: Wait for move to physically finish
                    self.serial_connection.write(b'M400\n')
                    self._wait_for_ok()
                
                if not z_limit_hit:
                     # Check one last time
                    if GPIO.input(Z_MAX_LIMIT_PIN) == 0: 
                        z_limit_hit = True
                        self.serial_connection.write(b'M410\n')
                        time.sleep(0.1)
                        self.serial_connection.reset_input_buffer()
                
                if not z_limit_hit:
                     self.serial_connection.write(b'G90\n')
                     raise Exception("Z-Max fine seek failed.")

                # --- Step 3: Read Position & Update Bounds ---
                self.queue_message("Reading Z-Max position...")
                
                # Back to Absolute Mode to read correct position
                self.serial_connection.write(b'G90\n')
                self._wait_for_ok()
                
                # Send M114
                self.serial_connection.reset_input_buffer()
                self.serial_connection.write(b'M114\n')
                
                measured_z = None
                read_start = time.time()
                while time.time() - read_start < 5.0:
                    if self.serial_connection.in_waiting > 0:
                        line = self.serial_connection.readline().decode('utf-8', errors='ignore').strip()
                        if "Z:" in line:
                            # Parse Z value
                            import re as local_re
                            match = local_re.search(r"Z:?\s*([-+]?\d*\.?\d+)", line)
                            if match:
                                measured_z = float(match.group(1))
                                break
                    time.sleep(0.05)
                
                if measured_z is not None:
                    # Add safety margin (e.g., 1mm)
                    safe_z_max = max(0.0, measured_z - 1.0)
                    
                    # Back off to safe Z
                    self.serial_connection.write(f"G1 Z{safe_z_max:.2f} F1000\n".encode('utf-8'))
                    self._wait_for_ok()
                    
                    self.message_queue.put(("CALIBRATION_COMPLETE", safe_z_max))
                else:
                    raise Exception("Failed to read Z position from M114.")
            
            finally:
                # Re-enable Safety Interrupt
                try:
                    GPIO.add_event_detect(Z_MAX_LIMIT_PIN, GPIO.FALLING, callback=self._on_z_max_trigger, bouncetime=200)
                except Exception:
                    pass

        except Exception as exc:
            self.queue_message(f"Homing Error: {exc}", "ERROR")
            self.message_queue.put(("CALIBRATION_FAILED", str(exc)))
            
            # Attempt to restore G90 in case of error
            if self.serial_connection:
                try: self.serial_connection.write(b'G90\n')
                except: pass
        finally:
            self.is_calibrating = False

    def _wait_for_ok(self, timeout=10.0):
        """Helper to block until 'ok' is received."""
        start = time.time()
        buffer = ""
        while time.time() - start < timeout:
             if self.stop_event.is_set(): return False
             if self.serial_connection.in_waiting > 0:
                 buffer += self.serial_connection.read(self.serial_connection.in_waiting).decode('utf-8', errors='ignore')
                 if 'ok' in buffer.lower(): return True
             time.sleep(0.05)
        return False


    def _go_to_position(self):
        """
        Sends the printer to the currently set 'target' position.

        This is the position indicated by the blue marker on the canvases and
        the 'TARGET' DRO display. It reads the internal floating-point values,
        not the GUI labels, for maximum precision.
        """
        if not self.serial_connection:
            messagebox.showerror("Error", "Not connected to a printer.")
            return
        if self.is_sending or self.is_manual_command_running:
            messagebox.showwarning("Busy", "Printer is busy with another operation.")
            return
            
        try:
            # Read the target position from the internal model variables.
            x, y, z, e_val = self.target_abs_x, self.target_abs_y, self.target_abs_z, self.target_abs_e
            
            # Determine feedrate. If E is moving significantly but XYZ are not, use rotation feedrate?
            # Simplest strategy: Use jog feedrate for XYZ moves, Rotation feedrate for pure E moves?
            # For combined moves, the printer limits by the slowest axis anyway. Let's use the Jog Feedrate as primary.
            feedrate = float(self.jog_feedrate_var.get())
            
            # Safety check against invalid values, though this should be prevented by other logic.
            if not (self.PRINTER_BOUNDS['x_min'] <= x <= self.PRINTER_BOUNDS['x_max']): raise ValueError(f"X ({x:.2f}) is out of bounds")
            if not (self.PRINTER_BOUNDS['y_min'] <= y <= self.PRINTER_BOUNDS['y_max']): raise ValueError(f"Y ({y:.2f}) is out of bounds")
            if not (self.PRINTER_BOUNDS['z_min'] <= z <= self.PRINTER_BOUNDS['z_max']): raise ValueError(f"Z ({z:.2f}) is out of bounds")
            if not (self.PRINTER_BOUNDS['e_min'] <= e_val <= self.PRINTER_BOUNDS['e_max']): raise ValueError(f"E ({e_val:.2f}) is out of bounds")
            if feedrate <= 0: raise ValueError("Feedrate must be a positive number")
                 
        except ValueError as e:
            self.log_message(f"Go To Position Error: {e}", "ERROR")
            messagebox.showerror("Invalid Input", f"Cannot execute 'Go To' command:\n{e}")
            return
            
        command = f"G90\nG1 X{x:.3f} Y{y:.3f} Z{z:.3f} E{e_val:.3f} F{feedrate:.0f}"
        self._send_manual_command(command)

    def _go_to_center(self):
        """
        Sets the 'target' position to the user-defined 'center' coordinates.

        This updates the internal target model and all GUI displays, including
        the input fields and canvas markers. It does *not* send a command to
        the printer; it only stages the coordinates for a subsequent 'Go' command.
        """
        try:
            # 1. Read the center coordinates from their StringVars.
            center_x = float(self.center_x_var.get())
            center_y = float(self.center_y_var.get())
            center_z = float(self.center_z_var.get())

            # 2. Set the internal 'target' position model.
            self.target_abs_x = center_x
            self.target_abs_y = center_y
            self.target_abs_z = center_z

            # 3. Update all displays (DRO labels and canvas markers).
            self._update_all_displays()
            self.log_message(f"Target set to center: X={center_x:.2f}, Y={center_y:.2f}, Z={center_z:.2f}", "INFO")

            # 4. Also update the 'Go To' entry boxes to reflect this change.
            mode = self.coord_mode.get()
            display_x = f"{center_x:.2f}" if mode == "absolute" else "0.00"
            display_y = f"{center_y:.2f}" if mode == "absolute" else "0.00"
            display_z = f"{center_z:.2f}" if mode == "absolute" else "0.00"

            self.goto_x_entry.delete(0, tk.END); self.goto_x_entry.insert(0, display_x)
            self.goto_y_entry.delete(0, tk.END); self.goto_y_entry.insert(0, display_y)
            self.goto_z_entry.delete(0, tk.END); self.goto_z_entry.insert(0, display_z)

        except ValueError:
            self.log_message("Cannot 'Go to Center': Invalid center coordinates.", "ERROR")
            messagebox.showerror("Error", "Center coordinates are invalid. Cannot set target.")

    # --- Canvas Click Handlers ---
    
    def _on_xy_canvas_click(self, event):
        """
        Handles clicks and drags on the XY canvas to set the Target X/Y position.
        """
        # pylint: disable=unused-argument
        if self.go_button['state'] == tk.DISABLED:
            return
            
        bounds = self.PRINTER_BOUNDS
        canvas_w = self.xy_canvas.winfo_width()
        canvas_h = self.xy_canvas.winfo_height()
        if canvas_w <= 1 or canvas_h <= 1:
            return
            
        x_range = bounds['x_max'] - bounds['x_min']
        y_range = bounds['y_max'] - bounds['y_min']
        
        # Clamp click coordinates to be within the canvas bounds.
        click_x_rel = max(0, min(canvas_w, event.x))
        click_y_rel = max(0, min(canvas_h, event.y))
        
        # Convert canvas pixel coordinates to world coordinates.
        world_x = bounds['x_min'] + (click_x_rel / canvas_w) * x_range if x_range != 0 else bounds['x_min']
        world_y = bounds['y_min'] + ((canvas_h - click_y_rel) / canvas_h) * y_range if y_range != 0 else bounds['y_min'] # Y is inverted
        
        # Update the internal model for the target position.
        self.target_abs_x = max(bounds['x_min'], min(bounds['x_max'], world_x))
        self.target_abs_y = max(bounds['y_min'], min(bounds['y_max'], world_y))
        
        # Update all GUI displays to reflect the new target.
        self._update_all_displays()


    def _draw_xy_canvas_guides(self, event=None):
        """Draws the toolpath for the CURRENT Z-LEVEL ONLY, plus grid, origin, and markers."""
        # pylint: disable=unused-argument
        # Only delete tagged items, not the persistent background
        self.xy_canvas.delete("toolpath", "guides", "marker_blue", "marker_red", "crosshair")
        bounds = self.PRINTER_BOUNDS
        w = self.xy_canvas.winfo_width(); h = self.xy_canvas.winfo_height()
        if w <= 1 or h <= 1: return
        
        x_range = bounds['x_max'] - bounds['x_min']; y_range = bounds['y_max'] - bounds['y_min']
        if x_range == 0 or y_range == 0: return 

        def world_to_canvas(wx, wy):
            cx = w * (wx - bounds['x_min']) / x_range
            cy = h - (h * (wy - bounds['y_min']) / y_range) # Inverted Y
            return cx, cy

        # --- Draw Grid & Origin Lines ---
        grid_color = "#1a2c3a"
        for i in range(int(bounds['x_min']), int(bounds['x_max']), 10):
            if i != 0:
                cx, _ = world_to_canvas(i, 0)
                self.xy_canvas.create_line(cx, 0, cx, h, fill=grid_color, tags="guides", dash=(2, 4))
        for i in range(int(bounds['y_min']), int(bounds['y_max']), 10):
            if i != 0:
                _, cy = world_to_canvas(0, i)
                self.xy_canvas.create_line(0, cy, w, cy, fill=grid_color, tags="guides", dash=(2, 4))

        if bounds['x_min'] <= 0 <= bounds['x_max'] and bounds['y_min'] <= 0 <= bounds['y_max']:
            canvas_x0, canvas_y0 = world_to_canvas(0, 0)
            self.xy_canvas.create_line(canvas_x0, 0, canvas_x0, h, fill=self.COLOR_BORDER, tags="guides")
            self.xy_canvas.create_line(0, canvas_y0, w, canvas_y0, fill=self.COLOR_BORDER, tags="guides")

        # --- Draw Center Crosshair ---
        try:
            c_x = float(self.center_x_var.get())
            c_y = float(self.center_y_var.get())
            center_cx, center_cy = world_to_canvas(c_x, c_y)
            
            # Dotted purple crosshair
            crosshair_color = "purple"
            self.xy_canvas.create_line(center_cx, 0, center_cx, h, fill=crosshair_color, dash=(2, 4), tags="crosshair")
            self.xy_canvas.create_line(0, center_cy, w, center_cy, fill=crosshair_color, dash=(2, 4), tags="crosshair")
        except ValueError:
            pass # Ignore if center coords are not valid numbers

        # --- Draw Toolpath for Current Layer ONLY (if enabled) ---
        if self.is_2d_plot_enabled.get() and self.toolpath_by_layer:
            # Determine the current Z-level to display. Prioritize the actual printer position.
            current_z = self.last_cmd_abs_z
            if current_z is None:
                # If position is unknown, fall back to the Z of the first move in the file
                if self.ordered_z_values:
                    current_z = self.ordered_z_values[0]
                else:
                    current_z = 0.0 # Absolute fallback

            # Get the path segments for this layer
            segments_for_current_layer = self.toolpath_by_layer.get(current_z, [])
            
            # Determine how many moves on THIS layer are completed
            completed_on_this_layer = 0
            if self.completed_move_count > 0 and self.completed_move_count <= len(self.move_to_layer_map):
                # Get the layer info for the last completed move
                last_completed_move_z, _ = self.move_to_layer_map[self.completed_move_count - 1]
                
                if last_completed_move_z > current_z:
                    # If we are printing a layer *above* the one being displayed, the displayed one is fully complete.
                    completed_on_this_layer = len(segments_for_current_layer)
                elif last_completed_move_z == current_z:
                    # If we are printing on the same layer, find out how many moves are done.
                    # Find the first global move index that belongs to this layer
                    first_move_idx_on_layer = -1
                    for i, (z, _) in enumerate(self.move_to_layer_map):
                        if z == current_z:
                            first_move_idx_on_layer = i
                            break
                    if first_move_idx_on_layer != -1:
                        completed_on_this_layer = self.completed_move_count - first_move_idx_on_layer
            
            # Now, draw the segments for the current layer with the correct colors
            for idx, (start_point, end_point) in enumerate(segments_for_current_layer):
                if start_point is None or end_point is None: continue
                
                start_cx, start_cy = world_to_canvas(start_point[0], start_point[1])
                end_cx, end_cy = world_to_canvas(end_point[0], end_point[1])
                
                path_color = self.COLOR_GREY_COMPLETED if idx < completed_on_this_layer else self.COLOR_ACCENT_CYAN
                
                self.xy_canvas.create_line(start_cx, start_cy, end_cx, end_cy, 
                                           fill=path_color, width=1, tags="toolpath")

        m_size = 5

        # --- Draw Blue Marker (Go To Target) ---
        try:
            target_x = self.target_abs_x; target_y = self.target_abs_y
            marker_cx, marker_cy = world_to_canvas(target_x, target_y)
            marker_cx = max(2, min(w - 2, marker_cx)); marker_cy = max(2, min(h - 2, marker_cy))
            self.xy_canvas.create_oval(marker_cx - m_size, marker_cy - m_size, marker_cx + m_size, marker_cy + m_size, fill=self.COLOR_ACCENT_CYAN, outline=self.COLOR_ACCENT_CYAN, tags="marker_blue")
        except Exception: pass

        # --- Draw Red Marker (Last Commanded Position) ---
        try:
            if self.last_cmd_abs_x is not None and self.last_cmd_abs_y is not None:
                 last_x = self.last_cmd_abs_x; last_y = self.last_cmd_abs_y
                 marker_cx, marker_cy = world_to_canvas(last_x, last_y)
                 marker_cx = max(2, min(w - 2, marker_cx)); marker_cy = max(2, min(h - 2, marker_cy))
                 self.xy_canvas.create_oval(marker_cx - m_size, marker_cy - m_size, marker_cx + m_size, marker_cy + m_size, fill=self.COLOR_ACCENT_RED, outline=self.COLOR_ACCENT_RED, tags="marker_red")
            else: pass
        except Exception: pass


    def _on_z_canvas_click(self, event):
        """
        Handles clicks and drags on the Z canvas to set the Target Z position.
        """
        # pylint: disable=unused-argument
        if self.go_button['state'] == tk.DISABLED:
            return
            
        bounds = self.PRINTER_BOUNDS
        canvas_h = self.z_canvas.winfo_height()
        if canvas_h <= 1:
            return
            
        z_range = bounds['z_max'] - bounds['z_min']
        click_y = max(0, min(canvas_h, event.y))
        
        # Convert canvas pixel Y coordinate to world Z coordinate.
        world_z = bounds['z_min'] + ((canvas_h - click_y) / canvas_h) * z_range if z_range != 0 else bounds['z_min']
        
        # Update the internal model for the target position.
        self.target_abs_z = max(bounds['z_min'], min(bounds['z_max'], world_z))
        
        # Update all GUI displays to reflect the new target.
        self._update_all_displays()

    def _draw_z_canvas_marker(self, event=None):
         """Draws the Z-axis toolpath graph and position markers."""
         # pylint: disable=unused-argument
         self.z_canvas.delete("all")
         bounds = self.PRINTER_BOUNDS; canvas_w = self.z_canvas.winfo_width(); canvas_h = self.z_canvas.winfo_height()
         if canvas_h <= 1: return
         
         self.z_canvas.create_rectangle(0, 0, canvas_w, canvas_h, fill=self.COLOR_BLACK, outline=self.COLOR_BORDER, width=1)
         
         z_range = bounds['z_max'] - bounds['z_min']
         if z_range == 0: return

         def z_to_canvas_y(world_z):
             canvas_y = canvas_h - ( (world_z - bounds['z_min']) / z_range * canvas_h )
             return max(1, min(canvas_h - 1, canvas_y))

         # --- Draw Z-Path Preview Graph (if enabled) ---
         if self.is_2d_plot_enabled.get() and self.ordered_z_values and len(self.ordered_z_values) > 1:
            num_points = len(self.ordered_z_values)
            # The graph is drawn horizontally across the narrow canvas
            x_step = (canvas_w - 1) / (num_points - 1) if num_points > 1 else 0

            for i in range(1, num_points):
                start_x = (i - 1) * x_step
                end_x = i * x_step
                
                start_y = z_to_canvas_y(self.ordered_z_values[i-1])
                end_y = z_to_canvas_y(self.ordered_z_values[i])
                
                # The index `i` corresponds to the move number.
                path_color = self.COLOR_GREY_COMPLETED if i < self.completed_move_count else self.COLOR_ACCENT_CYAN
                
                self.z_canvas.create_line(start_x, start_y, end_x, end_y, fill=path_color, width=1, tags="z_toolpath")
         
         # --- Draw Scale Labels ---
         # Max Z Label at top
         self.z_canvas.create_text(canvas_w/2, 10, text=f"{bounds['z_max']:.0f}", fill=self.COLOR_TEXT_SECONDARY, font=("Inter", 8), tags="labels")
         # Min Z Label at bottom
         self.z_canvas.create_text(canvas_w/2, canvas_h - 10, text="0", fill=self.COLOR_TEXT_SECONDARY, font=("Inter", 8), tags="labels")

         # --- Draw Blue Marker (Go To Target) ---
         try:
              target_z = self.target_abs_z
              canvas_y = z_to_canvas_y(target_z)
              self.z_canvas.create_line(2, canvas_y, canvas_w - 2, canvas_y, fill=self.COLOR_ACCENT_CYAN, width=3, tags="marker_blue")
         except Exception: pass

         # --- Draw Red Marker (Last Commanded Position) ---
         try:
              if self.last_cmd_abs_z is not None:
                  last_z = self.last_cmd_abs_z
                  canvas_y = z_to_canvas_y(last_z)
                  self.z_canvas.create_line(1, canvas_y, canvas_w - 1, canvas_y, fill=self.COLOR_ACCENT_RED, width=2, dash=(4, 2), tags="marker_red")
              else: pass
         except Exception: pass

    def _on_e_canvas_click(self, event):
        """Sets the E (Rotation) target based on click angle."""
        # pylint: disable=unused-argument
        if self.go_button['state'] == tk.DISABLED: return
        w = self.e_canvas.winfo_width()
        h = self.e_canvas.winfo_height()
        if w <= 1 or h <= 1: return
        
        cx, cy = w / 2, h / 2
        dx, dy = event.x - cx, event.y - cy
        
        # Calculate angle in degrees (0 is East/Right, 90 is South/Down)
        # We want 0 at North/Up, clockwise.
        # math.atan2(y, x) returns radians. East=0, South=PI/2, West=PI, North=-PI/2
        import math
        angle_rad = math.atan2(dy, dx)
        angle_deg = math.degrees(angle_rad)
        
        # Transform to 0=Up, Clockwise
        # Standard: East=0, CW=+
        # Target: Up=0, CW=+
        # Relation: Target = Standard + 90
        mapped_angle = angle_deg + 90
        
        # Normalize to 0-360
        if mapped_angle < 0: mapped_angle += 360
        if mapped_angle >= 360: mapped_angle -= 360
        
        # Snap to nearest 5 degrees for cleaner UI interaction
        mapped_angle = round(mapped_angle / 5) * 5
        
        self.target_abs_e = mapped_angle
        self._update_all_displays()

    def _draw_e_canvas_gauge(self, event=None):
        """Draws the circular gauge for the E-axis."""
        # pylint: disable=unused-argument
        self.e_canvas.delete("all")
        w = self.e_canvas.winfo_width()
        h = self.e_canvas.winfo_height()
        if w <= 1 or h <= 1: return
        
        cx, cy = w / 2, h / 2
        radius = min(w, h) / 2 - 10
        
        # Background Circle
        self.e_canvas.create_oval(cx - radius, cy - radius, cx + radius, cy + radius, outline=self.COLOR_BORDER, width=2)
        
        import math
        
        # Ticks
        for i in range(0, 360, 45):
            rad = math.radians(i - 90) # 0 is up
            r_in = radius - 8 if i % 90 == 0 else radius - 5
            x1 = cx + radius * math.cos(rad)
            y1 = cy + radius * math.sin(rad)
            x2 = cx + r_in * math.cos(rad)
            y2 = cy + r_in * math.sin(rad)
            color = self.COLOR_ACCENT_CYAN if i % 90 == 0 else self.COLOR_TEXT_SECONDARY
            self.e_canvas.create_line(x1, y1, x2, y2, fill=color, width=2 if i % 90 == 0 else 1)

        # Helper to draw a needle/marker
        def draw_needle(angle_deg, color, length_pct, width, tag):
            rad_n = math.radians(angle_deg - 90)
            nx = cx + (radius * length_pct) * math.cos(rad_n)
            ny = cy + (radius * length_pct) * math.sin(rad_n)
            self.e_canvas.create_line(cx, cy, nx, ny, fill=color, width=width, capstyle=tk.ROUND, tags=tag)
            # Small circle at tip
            r_tip = 3
            self.e_canvas.create_oval(nx - r_tip, ny - r_tip, nx + r_tip, ny + r_tip, fill=color, outline=color, tags=tag)

        # Target Needle (Blue)
        draw_needle(self.target_abs_e, self.COLOR_ACCENT_CYAN, 0.9, 3, "marker_blue")
        
        # Current Position Marker (Red)
        # Since E is multi-turn (can be > 360), we modulate it for the display
        if self.last_cmd_abs_e is not None:
            current_angle = self.last_cmd_abs_e % 360
            draw_needle(current_angle, self.COLOR_ACCENT_RED, 0.7, 2, "marker_red")
            
        # Center Cap
        self.e_canvas.create_oval(cx - 5, cy - 5, cx + 5, cy + 5, fill=self.COLOR_PANEL_BG, outline=self.COLOR_BORDER)




    # --- G-Code Sending Thread ---
    def gcode_sender_thread(self, gcode_to_send):
        """
        The main background worker thread for sending a G-code file.

        It iterates through the processed G-code lines, sending them one by one
        and waiting for an 'ok' from the printer. It handles pause/stop events,
        updates progress, and sends position updates to the main GUI thread
        via the message queue.

        Args:
            gcode_to_send (list): A list of the processed G-code strings to send.
        """
        self.message_queue.put(("SET_STATUS", "on")) # Green light for sending

        total_sendable_lines = len([line for line in gcode_to_send if line.strip() and not line.strip().startswith(';')])
        sent_line_count = 0
        success = True
        moves_sent_in_thread = 0
        
        # Initialize the position tracker for this thread.
        last_pos = {
            'x': self.last_cmd_abs_x, 
            'y': self.last_cmd_abs_y, 
            'z': self.last_cmd_abs_z,
            'e': self.last_cmd_abs_e
        }
        if last_pos['x'] is None: last_pos['x'] = self.PRINTER_BOUNDS['x_min']
        if last_pos['y'] is None: last_pos['y'] = self.PRINTER_BOUNDS['y_min']
        if last_pos['z'] is None: last_pos['z'] = self.PRINTER_BOUNDS['z_min']
        if last_pos['e'] is None: last_pos['e'] = 0.0
        
        # A brief delay to allow the user to see the "Starting..." message.
        self.queue_message("Waiting 5s before start...")
        time.sleep(5)
        
        for line in gcode_to_send:
            # This is the core of the pause functionality. It will block the thread
            # until the pause_event is set (i.e., when the user clicks "Resume").
            self.pause_event.wait() 

            if self.stop_event.is_set():
                success = False
                break

            gcode_line = line.strip()
            target_pos_for_line = None
            
            # Skip empty lines and comments.
            if not gcode_line or gcode_line.startswith(';'):
                continue

            is_move_command = gcode_line.upper().startswith("G0") or gcode_line.upper().startswith("G1")

            # --- Update Progress ---
            sent_line_count += 1
            percentage = (sent_line_count / total_sendable_lines) * 100 if total_sendable_lines > 0 else 0
            self.message_queue.put(("PROGRESS", (sent_line_count, total_sendable_lines, percentage)))

            # --- Predict Position for this Line ---
            if is_move_command:
                parsed = self._parse_gcode_coords(gcode_line)
                if parsed:
                    target_pos_for_line = last_pos.copy() 
                    target_pos_for_line.update(parsed) # The processed G-code is already absolute.
            elif "G28" in gcode_line.upper():
                target_pos_for_line = {
                    'x': self.PRINTER_BOUNDS['x_min'], 
                    'y': self.PRINTER_BOUNDS['y_min'], 
                    'z': self.PRINTER_BOUNDS['z_min'],
                    'e': 0.0 # Assuming G28 homes E too? Usually not for extrusion, but for 4-axis maybe.
                } 
                if 'X' not in gcode_line.upper() and 'Y' not in gcode_line.upper() and 'Z' not in gcode_line.upper():
                    # Full home. Usually resets XYZ. E behavior depends on firmware.
                    # Let's assume E is NOT homed by default G28 unless specified, or keep it as is.
                    # For safety in 4-axis, let's update XYZ and keep E unless E is in command.
                    target_pos_for_line['e'] = last_pos['e'] 
                    last_pos.update(target_pos_for_line)
                else: # Partial home
                    if 'X' in gcode_line.upper(): last_pos['x'] = self.PRINTER_BOUNDS['x_min']
                    if 'Y' in gcode_line.upper(): last_pos['y'] = self.PRINTER_BOUNDS['y_min']
                    if 'Z' in gcode_line.upper(): last_pos['z'] = self.PRINTER_BOUNDS['z_min']
                    # If E is ever homed via G28 E
                    if 'E' in gcode_line.upper(): last_pos['e'] = 0.0 
                    target_pos_for_line = last_pos.copy()

            # --- Send Line and Wait for 'ok' ---
            try:
                self.serial_connection.write(gcode_line.encode('utf-8') + b'\n')
                ok_received = False
                response_buffer = ""
                timeout = self.serial_connection.timeout if self.serial_connection.timeout else 10.0
                start_time = time.time()
                
                while time.time() - start_time < timeout + 2: # Add a small grace period
                    if self.stop_event.is_set():
                        raise InterruptedError("Stop event set while waiting for 'ok'.")

                    try:
                        if self.serial_connection.in_waiting > 0:
                            response_buffer += self.serial_connection.read(self.serial_connection.in_waiting).decode('utf-8', errors='ignore')
                            while '\n' in response_buffer:
                                full_line, response_buffer = response_buffer.split('\n', 1)
                                full_line = full_line.strip()
                                if full_line and 'ok' not in full_line.lower():
                                    self.queue_message(f"Received: {full_line}")
                                if 'ok' in full_line.lower():
                                    ok_received = True
                                    break
                        if ok_received:
                            break
                        time.sleep(0.02)
                    except serial.SerialException as read_err:
                        self.queue_message(f"Serial read error: {read_err}", "ERROR")
                        raise
                
                if ok_received:
                    if target_pos_for_line is not None:
                        last_pos = target_pos_for_line 
                        valid_pos = {k: v for k, v in last_pos.items() if v is not None}
                        if valid_pos:
                            self.message_queue.put(("POSITION_UPDATE", valid_pos))
                    
                    # If the line was a move, increment the progress counter for the toolpath display.
                    if is_move_command or "G28" in gcode_line.upper():
                        moves_sent_in_thread += 1
                        self.message_queue.put(("PATH_PROGRESS_UPDATE", moves_sent_in_thread))

                        # --- Auto-Measurement Logic ---
                        # Check if DMMs are connected AND Auto-Measure is ON
                        if self.is_dmm_connected and self.auto_measure_enabled.get():
                            self.queue_message("Auto-Measure: Stabilizing (M400)...")
                            
                            # 1. Send M400 to wait for move completion
                            m400_ok = False
                            try:
                                self.serial_connection.write(b'M400\n')
                                # M400 blocks until moves are done.
                                m400_start = time.time()
                                m400_buffer = ""
                                while time.time() - m400_start < 60.0: # 60s max wait
                                    if self.stop_event.is_set(): raise InterruptedError("Stop")
                                    if self.serial_connection.in_waiting > 0:
                                        m400_buffer += self.serial_connection.read(self.serial_connection.in_waiting).decode('utf-8', errors='ignore')
                                        if 'ok' in m400_buffer:
                                            m400_ok = True
                                            break
                                    time.sleep(0.1)
                            except Exception as e:
                                self.queue_message(f"M400 Error: {e}", "WARN")

                            if m400_ok:
                                self.queue_message("Auto-Measure: Measuring...")
                                # 2. Trigger Measurement with Stability Check
                                if self.dmm_group:
                                    try:
                                        # Use the new stability method
                                        vals = self._measure_with_stability()
                                        
                                        self.message_queue.put(("MEASUREMENT_RESULT", vals))
                                        if self.log_measurements_enabled.get():
                                            self._log_measurement_to_file(vals, coords=last_pos)
                                    except Exception as e:
                                        self.queue_message(f"Auto-Measure Error: {e}", "ERROR")
                            else:
                                self.queue_message("Auto-Measure skipped (M400 timeout)", "WARN")

                else:
                    self.queue_message(f"Warning: No 'ok' received for '{gcode_line}' (timeout: {timeout:.1f}s).", "WARN")

            except InterruptedError:
                self.queue_message("G-code stream interrupted by user.", "WARN")
                success = False
                break
            except serial.SerialException as e:
                self.queue_message(f"Serial Error during stream: {e}", "ERROR")
                success = False
                self.message_queue.put(("CONNECTION_LOST", None))
                break
            except Exception as e:
                self.queue_message(f"Unexpected error on line {sent_line_count} ('{gcode_line}'): {e}", "ERROR")
                success = False
                break
                
        # --- Finalize ---
        final_msg = "Stream finished successfully." if success and not self.stop_event.is_set() else "Stream stopped."
        self.queue_message(final_msg, "SUCCESS" if success and not self.stop_event.is_set() else "INFO")
        
        if success and not self.stop_event.is_set():
            self.message_queue.put(("SET_STATUS", "on"))
        elif self.stop_event.is_set():
            self.message_queue.put(("SET_STATUS", "error"))
        
        if not success or self.stop_event.is_set():
            self.message_queue.put(("PROGRESS_RESET", None))
            
        self.message_queue.put(("FILE_SEND_FINISHED", None))

    # --- GUI Update & Event Handling ---

    def check_message_queue(self):
        """
        Periodically checks a queue for messages from background threads.

        This is the primary mechanism for thread-safe communication from the
        serial and sender threads back to the main GUI thread. It processes
        messages to update progress bars, logs, connection status, and more.
        This method reschedules itself to run continuously.
        """
        try:
            while True: # Process all messages currently in the queue.
                msg_type, msg_content = self.message_queue.get_nowait()

                if msg_type == "LOG":
                    self.log_message(msg_content[1], msg_content[0])
                
                elif msg_type == "PROGRESS":
                    current, total, percentage = msg_content
                    self.progress_var.set(percentage)
                    self.progress_label_var.set(f"{current}/{total} lines")
                
                elif msg_type == "PROGRESS_RESET":
                    self.progress_var.set(0.0)
                    self.progress_label_var.set("Progress: Stopped/Idle")

                elif msg_type == "SET_STATUS":
                    if hasattr(self, 'status_indicator'): self.status_indicator.set_status(msg_content)
                    if hasattr(self, 'header_status_indicator'): self.header_status_indicator.set_status(msg_content)

                elif msg_type == "POSITION_UPDATE":
                    pos_dict = msg_content
                    if 'x' in pos_dict and pos_dict['x'] is not None: self.last_cmd_abs_x = pos_dict['x']
                    if 'y' in pos_dict and pos_dict['y'] is not None: self.last_cmd_abs_y = pos_dict['y']
                    if 'z' in pos_dict and pos_dict['z'] is not None: self.last_cmd_abs_z = pos_dict['z']
                    if 'e' in pos_dict and pos_dict['e'] is not None: self.last_cmd_abs_e = pos_dict['e']
                    self._update_all_displays()
                
                elif msg_type == "PATH_PROGRESS_UPDATE":
                    self.completed_move_count = msg_content
                    self._draw_xy_canvas_guides()
                    self._draw_z_canvas_marker()
                    if self.is_3d_plot_enabled.get() and hasattr(self, 'canvas_3d') and self.canvas_3d:
                        self._draw_3d_toolpath()

                elif msg_type == "CONNECTED":
                    self.serial_connection, found_port, baudrate = msg_content
                    self.log_message(f"Connected on {found_port}!", "SUCCESS")
                    
                    self.connection_status_var.set(f"Connected to {found_port}")
                    self.status_indicator.set_status("on")
                    self.header_status_indicator.set_status("on")
                    self.footer_status_var.set(f"{found_port} @ {baudrate}")
                    
                    self.connect_button.config(text="Disconnect", state=tk.NORMAL)
                    self.port_combobox.config(state=tk.DISABLED)
                    self.baud_entry.config(state=tk.DISABLED)
                    
                    self._set_manual_controls_state(tk.NORMAL)
                    self._set_goto_controls_state(tk.NORMAL)
                    self._set_terminal_controls_state(tk.NORMAL)

                    if self.processed_gcode: self.start_button.config(state=tk.NORMAL)
                    if hasattr(self, 'cancel_connect_button'): self.cancel_connect_button.grid_remove(); self.cancel_connect_button.config(state=tk.DISABLED)
                    
                    self.progress_var.set(0.0)
                    self.progress_label_var.set("Progress: Idle")
                    
                    # On successful connection, assume printer is at origin until told otherwise.
                    self.last_cmd_abs_x, self.last_cmd_abs_y, self.last_cmd_abs_z = self.PRINTER_BOUNDS['x_min'], self.PRINTER_BOUNDS['y_min'], self.PRINTER_BOUNDS['z_min']
                    self._update_all_displays()
                
                elif msg_type == "CONNECT_FAIL":
                    err_msg = msg_content
                    self.log_message(f"Connect failed: {err_msg}", "ERROR")
                    if "No responsive printer found" not in err_msg:
                        messagebox.showerror("Connection Failed", err_msg)
                    self.disconnect_printer(silent=True) # Reset GUI to disconnected state
                
                elif msg_type == "CONNECT_CANCELLED":
                    self.log_message("Connection cancelled.", "INFO")
                    self.disconnect_printer(silent=True) # Reset GUI to disconnected state
                
                elif msg_type == "CONNECT_ATTEMPT_FINISHED":
                    # This message ensures controls are re-enabled even if connection fails or is cancelled.
                    if not self.serial_connection:
                        self.connect_button.config(state=tk.NORMAL)
                        self.port_combobox.config(state="readonly")
                        self.baud_entry.config(state=tk.NORMAL)
                    if hasattr(self, 'cancel_connect_button'):
                        self.cancel_connect_button.grid_remove()
                        self.cancel_connect_button.config(state=tk.DISABLED)
                
                elif msg_type == "FILE_SEND_FINISHED":
                    self.is_sending = False
                    self.is_paused = False
                    self.pause_resume_button.config(text="Pause", state=tk.DISABLED)
                    if self.serial_connection and not self.stop_event.is_set():
                        # Re-enable controls if the job finished normally.
                        self.start_button.config(state=tk.NORMAL if self.processed_gcode else tk.DISABLED)
                        self._set_manual_controls_state(tk.NORMAL)
                        self._set_goto_controls_state(tk.NORMAL)
                        self._set_terminal_controls_state(tk.NORMAL)
                        self.status_indicator.set_status("on")
                        self.header_status_indicator.set_status("on")
                        self.progress_var.set(100.0)
                        self.progress_label_var.set(f"Finished: {self.total_lines_to_send}/{self.total_lines_to_send} lines")
                    else:
                        # If stopped or disconnected, keep controls disabled.
                        self.progress_var.set(0.0)
                        self.progress_label_var.set("Progress: Stopped/Idle")
                
                elif msg_type == "MANUAL_FINISHED":
                    self.is_manual_command_running = False
                    self.is_paused = False
                    if not self.is_sending: # Don't mess with the button if a file send is also happening
                        self.pause_resume_button.config(text="Pause", state=tk.DISABLED)
                    if self.serial_connection and not self.stop_event.is_set():
                        # Re-enable controls.
                        self._set_manual_controls_state(tk.NORMAL)
                        self._set_goto_controls_state(tk.NORMAL)
                        self._set_terminal_controls_state(tk.NORMAL)
                        self.start_button.config(state=tk.NORMAL if self.processed_gcode else tk.DISABLED)
                        self.status_indicator.set_status("on")
                        self.header_status_indicator.set_status("on")
                
                elif msg_type == "CONNECTION_LOST":
                    self.log_message("Connection lost.", "ERROR")
                    messagebox.showerror("Connection Lost", "Serial connection lost.\nPlease reconnect.")
                    self.disconnect_printer(silent=True)

                elif msg_type == "EMERGENCY_STOP_TRIGGERED":
                    reason = msg_content
                    self.log_message(f"!!! STOP: {reason} !!!", "CRITICAL")
                    self.emergency_stop()

                elif msg_type == "DMM_CONNECTED":
                    self.is_dmm_connected = True
                    self.dmm_status_var.set("DMMs: Connected")
                    self.dmm_connect_button.config(text="Disconnect DMMs", state=tk.NORMAL)
                    self.measure_button.config(state=tk.NORMAL)
                    self.log_message("DMMs connected successfully.", "SUCCESS")

                elif msg_type == "DMM_FAIL":
                    self.is_dmm_connected = False
                    self.dmm_status_var.set("DMMs: Error")
                    self.dmm_connect_button.config(text="Connect DMMs", state=tk.NORMAL)
                    self.measure_button.config(state=tk.DISABLED)
                    self.log_message(f"DMM Error: {msg_content}", "ERROR")

                elif msg_type == "MEASUREMENT_RESULT":
                    vals = msg_content 
                    if vals:
                         # Config: [Vin, Iin, Vsys, Saux, Vinv, Sinv]
                         display_str = f"Vin: {vals[0]:.2f}V"
                         if len(vals) > 1: display_str += f" | Iin: {vals[1]:.2f}"
                         self.last_measurement_var.set(display_str)
                         self.log_message(f"Measured: {', '.join([f'{v:.4f}' for v in vals])}", "INFO")

                elif msg_type == "CALIBRATION_COMPLETE":
                    safe_z_max = msg_content
                    self.PRINTER_BOUNDS['z_max'] = safe_z_max
                    self.log_message(f"Auto-Homing Complete. New Z-Max: {safe_z_max:.2f}mm", "SUCCESS")
                    
                    # Reset controls
                    self.is_manual_command_running = False
                    self._set_manual_controls_state(tk.NORMAL)
                    self._set_goto_controls_state(tk.NORMAL)
                    self.start_button.config(state=tk.NORMAL if self.processed_gcode else tk.DISABLED)
                    self.status_indicator.set_status("on")
                    
                    # Re-validate file
                    if self.gcode_filepath:
                        self.process_gcode()

                elif msg_type == "CALIBRATION_FAILED":
                    self.log_message(f"Auto-Homing Failed: {msg_content}", "ERROR")
                    messagebox.showerror("Homing Failed", msg_content)
                    
                    # Reset controls
                    self.is_manual_command_running = False
                    self._set_manual_controls_state(tk.NORMAL)
                    self._set_goto_controls_state(tk.NORMAL)
                    self.start_button.config(state=tk.NORMAL if self.processed_gcode else tk.DISABLED)
                    self.status_indicator.set_status("error")


        except queue.Empty:
            pass # No messages to process.
        finally:
            # Reschedule this method to run again after a short delay.
            self.after_id = self.root.after(100, self.check_message_queue)

    def queue_message(self, message, level="INFO"):
        """
        A thread-safe way to send a log message to the GUI.

        This should be used by background threads instead of calling `log_message`
        directly, as Tkinter is not thread-safe.

        Args:
            message (str): The message to log.
            level (str): The severity level of the message.
        """
        self.message_queue.put(("LOG", (level, message)))

    def on_closing(self):
        """
        Handles the application window being closed.

        If an operation is in progress, it asks the user for confirmation before
        stopping the operation and closing the application. Otherwise, it closes
        gracefully.
        """
        if self.is_sending or self.is_manual_command_running:
            if messagebox.askyesno("Confirm Exit", "An operation is in progress. Abort and exit?"):
                if self.after_id:
                    self.root.after_cancel(self.after_id)
                self.pause_event.set() # Unblock any paused threads
                self.emergency_stop()
                self.disconnect_dmms()
                time.sleep(1) # Give time for stop command to process
                self.root.destroy()
            else:
                return # Do not close the window
        else:
            if self.after_id:
                self.root.after_cancel(self.after_id)
            self.disconnect_printer(silent=True)
            self.disconnect_dmms()
            if HAS_GPIO and GPIO:
                try: GPIO.cleanup()
                except: pass
            self.root.destroy()

    def _history_up(self, event):
        """
        Navigates up through the terminal command history.
        
        Bound to the <Up> arrow key in the terminal input.
        """
        if not self.command_history:
            return "break"
        if self.history_index > 0:
            self.history_index -= 1
            self.terminal_input.delete(0, tk.END)
            self.terminal_input.insert(0, self.command_history[self.history_index])
        return "break" # Prevents the default key binding from firing.

    def _history_down(self, event):
        """
        Navigates down through the terminal command history.

        Bound to the <Down> arrow key in the terminal input.
        """
        if not self.command_history:
            return "break"
        if self.history_index < len(self.command_history):
            self.history_index += 1
            self.terminal_input.delete(0, tk.END)
            # If we're not at the end of history, show the next command.
            if self.history_index < len(self.command_history):
                self.terminal_input.insert(0, self.command_history[self.history_index])
            # If we are at the end, the input is cleared, allowing a new command.
        return "break" # Prevents the default key binding from firing.
    
    # --- NEW: Mouse Wheel Scroll Handler ---
    def _on_mousewheel_scroll(self, event):
        """
        Scrolls the left-hand control panel canvas with the mouse wheel.
        
        This is bound when the mouse enters the left panel and unbound when it leaves,
        preventing accidental scrolling when interacting with other widgets.
        It handles different scroll events for Windows/macOS and Linux.
        """
        # On Windows/macOS, event.delta is typically +/- 120.
        # On Linux, event.num is 4 for scroll up and 5 for scroll down.
        if event.num == 5 or event.delta < 0:
            self.left_canvas.yview_scroll(1, "units")
        elif event.num == 4 or event.delta > 0:
            self.left_canvas.yview_scroll(-1, "units")
        
    # --- New Coordinate Control Methods ---
    
    def _mark_current_as_center(self):
        """
        Sets the user-defined 'Center' coordinates to the printer's last known position.

        This is useful for establishing a work coordinate system (WCS) origin
        on the workpiece without manually typing in the coordinates.
        """
        if self.last_cmd_abs_x is None:
             self.log_message("Cannot mark center: No known last position.", "WARN")
             messagebox.showwarning("Mark Center Failed", "Cannot mark center.\nNo known printer position.\nTry homing or moving the printer first.")
             return
             
        try:
            self.center_x_var.set(f"{self.last_cmd_abs_x:.2f}")
            self.center_y_var.set(f"{self.last_cmd_abs_y:.2f}")
            self.center_z_var.set(f"{self.last_cmd_abs_z:.2f}")
            self.log_message(f"New center marked at: X={self.last_cmd_abs_x:.2f}, Y={self.last_cmd_abs_y:.2f}, Z={self.last_cmd_abs_z:.2f}", "SUCCESS")
            # Trigger the same logic as if the user changed the entry fields manually.
            self._on_center_change()
        except Exception as e:
             self.log_message(f"Error marking center: {e}", "ERROR")

    def _on_center_change(self, event=None):
        """
        Callback for when the user manually changes the Center coordinate entries.
        
        This triggers a full GUI update and re-processes the G-code file if one is loaded.
        """
        if self.gcode_filepath:
            # Re-process the file, which will use the new center values from the StringVars.
            # This function also handles logging and calls _update_all_displays internally.
            self.process_gcode()
        else:
            # If no file is loaded, we still need to update displays (e.g., crosshair, DROs).
            self._update_all_displays()

    def _set_coord_mode(self, mode):
        """
        Sets the coordinate display mode for the 'Go To' controls.

        This changes the appearance of the 'Absolute'/'Relative' buttons and
        updates all coordinate displays to reflect the chosen mode.

        Args:
            mode (str): The mode to switch to, either "absolute" or "relative".
        """
        if mode == "absolute":
            self.coord_mode.set("absolute")
            self.abs_button.config(style="Segment.Active.TButton") 
            self.rel_button.config(style="Segment.TButton")
        else: # relative
            self.coord_mode.set("relative")
            self.abs_button.config(style="Segment.TButton")
            self.rel_button.config(style="Segment.Active.TButton")
        
        # Clear the entry boxes and update all displays to reflect the new mode.
        self.goto_x_entry.delete(0, tk.END)
        self.goto_y_entry.delete(0, tk.END)
        self.goto_z_entry.delete(0, tk.END)
        self._update_all_displays()

    def _update_all_displays(self, event=None):
        """
        Central function to refresh all coordinate-based GUI elements.

        This should be called whenever a position (last commanded or target) or
        the coordinate mode changes. It recalculates all values for the DROs,
        footer, and redraws the canvas markers.
        """
        try:
            center_x = float(self.center_x_var.get())
            center_y = float(self.center_y_var.get())
            center_z = float(self.center_z_var.get())
        except ValueError:
             # If center coords are invalid, just skip the update to avoid errors.
             return 

        mode = self.coord_mode.get()

        # --- Update "Current" (Red) Display Labels ---
        if self.last_cmd_abs_x is not None:
            if mode == "absolute":
                self.last_cmd_x_display_var.set(f"{self.last_cmd_abs_x:.2f}")
                self.last_cmd_y_display_var.set(f"{self.last_cmd_abs_y:.2f}")
                self.last_cmd_z_display_var.set(f"{self.last_cmd_abs_z:.2f}")
            else: # "relative"
                self.last_cmd_x_display_var.set(f"{self.last_cmd_abs_x - center_x:.2f}")
                self.last_cmd_y_display_var.set(f"{self.last_cmd_abs_y - center_y:.2f}")
                self.last_cmd_z_display_var.set(f"{self.last_cmd_abs_z - center_z:.2f}")
            
            # The footer always displays the absolute machine coordinates.
            if hasattr(self, 'footer_coords_var'):
                self.footer_coords_var.set(f"X: {self.last_cmd_abs_x:.2f}  Y: {self.last_cmd_abs_y:.2f}  Z: {self.last_cmd_abs_z:.2f}")
            
        else: # Position is not yet known.
            self.last_cmd_x_display_var.set("N/A")
            self.last_cmd_y_display_var.set("N/A")
            self.last_cmd_z_display_var.set("N/A")
            if hasattr(self, 'footer_coords_var'):
                self.footer_coords_var.set("X: N/A  Y: N/A  Z: N/A")

        if self.last_cmd_abs_e is not None:
            self.last_cmd_e_display_var.set(f"{self.last_cmd_abs_e:.2f}")
        else:
            self.last_cmd_e_display_var.set("N/A")

        # --- Update "Target" (Blue) Display Labels ---
        if mode == "absolute":
            self.goto_x_display_var.set(f"{self.target_abs_x:.2f}")
            self.goto_y_display_var.set(f"{self.target_abs_y:.2f}")
            self.goto_z_display_var.set(f"{self.target_abs_z:.2f}")
        else: # "relative"
            self.goto_x_display_var.set(f"{self.target_abs_x - center_x:.2f}")
            self.goto_y_display_var.set(f"{self.target_abs_y - center_y:.2f}")
            self.goto_z_display_var.set(f"{self.target_abs_z - center_z:.2f}")
        
        self.goto_e_display_var.set(f"{self.target_abs_e:.2f}")

        # --- Redraw Canvases ---
        # This is necessary to move the position markers.
        if hasattr(self, 'xy_canvas') and self.xy_canvas.winfo_width() > 1:
            self._draw_xy_canvas_guides()
        if hasattr(self, 'z_canvas') and self.z_canvas.winfo_height() > 1:
            self._draw_z_canvas_marker()
        if hasattr(self, 'e_canvas') and self.e_canvas.winfo_height() > 1:
            self._draw_e_canvas_gauge()
        if hasattr(self, 'canvas_3d') and self.canvas_3d:
            self._update_3d_position_marker()
        
    def _on_goto_entry_commit(self, event=None):
        """
        Callback for when the user enters a value in the 'Go To' entry boxes.

        This is triggered by pressing Enter or focus leaving the widget. It parses
        the input, converts it to an absolute coordinate if necessary, clamps it
        to the printer bounds, and updates the internal 'target' position model.
        """
        try:
            val_x_str = self.goto_x_entry.get()
            val_y_str = self.goto_y_entry.get()
            val_z_str = self.goto_z_entry.get()
            val_e_str = self.goto_e_entry.get()
            
            new_abs_x, new_abs_y, new_abs_z = self.target_abs_x, self.target_abs_y, self.target_abs_z
            new_abs_e = self.target_abs_e
            
            center_x = float(self.center_x_var.get())
            center_y = float(self.center_y_var.get())
            center_z = float(self.center_z_var.get())
            mode = self.coord_mode.get()

            # If the entry has a value, parse it and update the target position.
            if val_x_str:
                val_x = float(val_x_str)
                new_abs_x = val_x + center_x if mode == "relative" else val_x
                self.target_abs_x = max(self.PRINTER_BOUNDS['x_min'], min(self.PRINTER_BOUNDS['x_max'], new_abs_x))
            if val_y_str:
                val_y = float(val_y_str)
                new_abs_y = val_y + center_y if mode == "relative" else val_y
                self.target_abs_y = max(self.PRINTER_BOUNDS['y_min'], min(self.PRINTER_BOUNDS['y_max'], new_abs_y))
            if val_z_str:
                val_z = float(val_z_str)
                new_abs_z = val_z + center_z if mode == "relative" else val_z
                self.target_abs_z = max(self.PRINTER_BOUNDS['z_min'], min(self.PRINTER_BOUNDS['z_max'], new_abs_z))
            if val_e_str:
                val_e = float(val_e_str)
                # E is Rotation, Center usually doesn't apply, but logic allows relative E if needed.
                # Assuming E is always absolute angle for "Set E" in this context, or relative if mode is relative.
                # Let's treat E relative to 0 if mode is relative, effectively additive.
                new_abs_e = val_e if mode == "absolute" else (val_e + self.target_abs_e) # Relative adds to current target? Or current pos?
                # Standard behavior for UI input is usually "Set Target to X".
                # If relative mode is on, input "10" usually means "Current + 10".
                # BUT here, this function sets the TARGET.
                # Let's stick to the pattern: Absolute input sets Absolute Target. Relative input adds to Center offset.
                # Since E doesn't have a "Center" defined in UI, Relative Mode for E is ambiguous.
                # Let's assume "Set E" is always absolute angle for simplicity, or add to current target if relative?
                # Re-reading X/Y logic: "val_x + center_x". This implies the input is a coordinate relative to WCS origin.
                # For E, WCS origin is 0. So Relative Input = Absolute Input.
                self.target_abs_e = max(self.PRINTER_BOUNDS['e_min'], min(self.PRINTER_BOUNDS['e_max'], new_abs_e))
                
        except ValueError:
            self.log_message("Invalid coordinate entered in 'Go To' field.", "WARN")
        finally:
            # Refresh all displays to show the clamped/updated target position.
            self._update_all_displays()

    # --- DMM Methods ---

    def toggle_dmm_connection(self):
        if self.is_dmm_connected:
            self.disconnect_dmms()
        else:
            self.dmm_connect_button.config(state=tk.DISABLED)
            self.dmm_status_var.set("Connecting...")
            threading.Thread(target=self._connect_dmm_thread, daemon=True).start()

    def _connect_dmm_thread(self):
        try:
            if not HAS_PYVISA:
                self.queue_message("PyVISA not found. Cannot connect.", "ERROR")
                self.message_queue.put(("DMM_FAIL", "PyVISA Missing"))
                return

            self.queue_message("Initializing DMMs...")
            self.dmm_group = DmmGroup(DMM_CONFIG)
            self.dmm_group.initialize()
            self.message_queue.put(("DMM_CONNECTED", None))
        except Exception as e:
            self.queue_message(f"DMM Connect Error: {e}", "ERROR")
            self.message_queue.put(("DMM_FAIL", str(e)))

    def disconnect_dmms(self):
        if self.dmm_group:
            try:
                self.dmm_group.close()
            except: pass
        self.dmm_group = None
        self.is_dmm_connected = False
        self.dmm_status_var.set("DMMs: Disconnected")
        self.dmm_connect_button.config(text="Connect DMMs")
        self.measure_button.config(state=tk.DISABLED)

    def trigger_manual_measurement(self):
        if not self.is_dmm_connected: return
        self.log_message("Triggering manual measurement...", "INFO")
        threading.Thread(target=self._measure_thread, args=(True,), daemon=True).start()

    def _measure_thread(self, is_manual=False):
        """
        Performs the measurement.
        is_manual: if True, logs result to GUI console.
        """
        if not self.dmm_group: return
        try:
            self.dmm_group.trigger()
            values = self.dmm_group.read()
            
            self.message_queue.put(("MEASUREMENT_RESULT", values))
            
            # Log to file if enabled
            if self.log_measurements_enabled.get():
                self._log_measurement_to_file(values)
                
        except Exception as e:
            self.queue_message(f"Measurement Error: {e}", "ERROR")

    def _measure_with_stability(self):
        """
        Takes readings until stability criteria are met or retries exhausted.
        Returns the final averaged list of values.
        """
        if not self.dmm_group: return []

        threshold_pct = self.stability_threshold_var.get()
        max_retries = self.max_retries_var.get()
        measurements_per_point = self.measurements_per_point_var.get()
        pre_delay = self.pre_measure_delay_var.get()
        
        if measurements_per_point < 2: measurements_per_point = 2 # Minimum for std dev

        # 1. Pre-Measure Delay
        if pre_delay > 0:
            self.queue_message(f"Stabilizing... ({pre_delay}s)")
            time.sleep(pre_delay)

        # 2. Fill initial window
        history = []
        for _ in range(measurements_per_point):
             self.dmm_group.trigger()
             history.append(self.dmm_group.read())
        
        attempts = 0
        stable = False
        
        while attempts < max_retries:
            # Transpose history to get lists of values per DMM channel
            channels = list(zip(*history))
            
            max_dev = 0.0
            for channel_data in channels:
                if not channel_data: continue
                
                mean = sum(channel_data) / len(channel_data)
                if mean == 0: continue 
                
                variance = sum([((x - mean) ** 2) for x in channel_data]) / len(channel_data)
                std_dev = variance ** 0.5
                
                pct_dev = (std_dev / abs(mean)) * 100
                if pct_dev > max_dev:
                    max_dev = pct_dev

            self.queue_message(f"Stability Check: Max Dev = {max_dev:.3f}% (Threshold: {threshold_pct}%)")
            
            if max_dev <= threshold_pct:
                stable = True
                break
            
            # Not stable, take another reading
            self.dmm_group.trigger()
            new_reading = self.dmm_group.read()
            
            # Maintain sliding window size
            history.append(new_reading)
            while len(history) > measurements_per_point:
                history.pop(0)
            
            attempts += 1
            time.sleep(0.2) 

        if not stable:
            self.queue_message("Warning: Stability threshold not met. Proceeding with last reading.", "WARN")
        else:
             self.queue_message("Readings stabilized.", "SUCCESS")

        # Return the average of the current history window
        final_avg = []
        if history:
            final_channels = list(zip(*history))
            for ch in final_channels:
                final_avg.append(sum(ch) / len(ch))
        return final_avg

    def select_log_file(self):
        """Opens a file dialog to choose the CSV log file."""
        filename = filedialog.asksaveasfilename(
            title="Save Measurement Log",
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")],
            initialfile=self.log_filepath_var.get()
        )
        if filename:
            self.log_filepath_var.set(filename)

    def _log_measurement_to_file(self, values, coords=None):
        filepath = self.log_filepath_var.get()
        if not filepath:
            self.queue_message("Log Error: No filename specified.", "ERROR")
            return

        # Check if file exists to write header
        file_exists = False
        try:
            with open(filepath, 'r', encoding='utf-8'): 
                file_exists = True
        except FileNotFoundError: 
            pass
        except Exception: 
            pass # Permission error or other
        
        try:
            with open(filepath, 'a', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                
                # Write header if new file
                if not file_exists:
                    headers = ["Timestamp", "X", "Y", "Z"] + [d[2] for d in DMM_CONFIG]
                    writer.writerow(headers)
                    self.queue_message(f"Created log file: {filepath}", "SUCCESS")

                if coords:
                    log_x = coords.get('x', 0.0)
                    log_y = coords.get('y', 0.0)
                    log_z = coords.get('z', 0.0)
                else:
                    # Get current position (use last known)
                    log_x = self.last_cmd_abs_x if self.last_cmd_abs_x is not None else 0.0
                    log_y = self.last_cmd_abs_y if self.last_cmd_abs_y is not None else 0.0
                    log_z = self.last_cmd_abs_z if self.last_cmd_abs_z is not None else 0.0
                
                row = [datetime.now().isoformat(), log_x, log_y, log_z] + values
                writer.writerow(row)
        except Exception as log_err:
            self.queue_message(f"Log Write Error: {log_err}", "ERROR")


    # --- Collision Avoidance Test Screen ---

    def _open_collision_test_screen(self):
        """Swaps the main UI for the Collision Avoidance Test screen."""
        if not self.serial_connection:
            messagebox.showerror("Error", "Not connected to printer.")
            return

        # Hide the main view
        self.main_view_frame.pack_forget()

        # Create the test view frame
        self.test_view_frame = tk.Frame(self.root, bg=self.COLOR_BG)
        self.test_view_frame.pack(fill=tk.BOTH, expand=True)

        # --- Content ---
        # 1. Warning Label
        warning_frame = tk.Frame(self.test_view_frame, bg=self.COLOR_ACCENT_AMBER, padx=20, pady=20)
        warning_frame.pack(fill=tk.X, padx=50, pady=(50, 20))
        
        lbl_warn = tk.Label(warning_frame, text="⚠ CAUTION ⚠", 
                            font=("Inter", 24, "bold"), bg=self.COLOR_ACCENT_AMBER, fg=self.COLOR_BLACK)
        lbl_warn.pack()
        
        lbl_msg = tk.Label(warning_frame, text="Remove hub and implant to avoid damage to equipment.", 
                           font=("Inter", 16), bg=self.COLOR_ACCENT_AMBER, fg=self.COLOR_BLACK)
        lbl_msg.pack(pady=(10, 0))

        # 2. Controls Frame
        ctrl_frame = tk.Frame(self.test_view_frame, bg=self.COLOR_BG)
        ctrl_frame.pack(expand=True)

        # Begin Test Button
        self.btn_begin_test = tk.Button(ctrl_frame, text="BEGIN TEST", 
                                        font=("Orbitron", 18, "bold"), 
                                        bg=self.COLOR_ACCENT_CYAN, fg=self.COLOR_BLACK,
                                        activebackground="#00eaff", activeforeground=self.COLOR_BLACK,
                                        relief=tk.RAISED, bd=3, padx=30, pady=15,
                                        command=self._start_collision_test)
        self.btn_begin_test.pack(pady=20)

        # STOP MOTION Button (Big Red)
        self.btn_stop_test = tk.Button(ctrl_frame, text="STOP MOTION", 
                                       font=("Orbitron", 24, "bold"), 
                                       bg=self.COLOR_ACCENT_RED, fg="white",
                                       activebackground="#ff6666", activeforeground="white",
                                       relief=tk.RAISED, bd=5, padx=50, pady=30,
                                       command=self.emergency_stop)
        self.btn_stop_test.pack(pady=20)

        # Status Label for Test
        self.lbl_test_status = tk.Label(ctrl_frame, text="Ready", font=("Inter", 14), bg=self.COLOR_BG, fg=self.COLOR_TEXT_SECONDARY)
        self.lbl_test_status.pack(pady=10)

        # Exit Button (Small)
        self.btn_exit_test = tk.Button(self.test_view_frame, text="Exit Test", 
                                       font=("Inter", 12), 
                                       bg=self.COLOR_PANEL_BG, fg=self.COLOR_TEXT_PRIMARY,
                                       relief=tk.FLAT, padx=20, pady=10,
                                       command=self._close_collision_test_screen)
        self.btn_exit_test.pack(side=tk.BOTTOM, pady=30)

    def _close_collision_test_screen(self):
        """Restores the main UI."""
        if hasattr(self, 'test_view_frame') and self.test_view_frame:
            self.test_view_frame.destroy()
        
        self.main_view_frame.pack(fill=tk.BOTH, expand=True)

    def _start_collision_test(self):
        """Starts the collision test sequence in a background thread."""
        if self.is_sending or self.is_manual_command_running:
            return
        
        self.btn_begin_test.config(state=tk.DISABLED)
        self.btn_exit_test.config(state=tk.DISABLED)
        self.lbl_test_status.config(text="Moving to Center...", fg=self.COLOR_ACCENT_CYAN)
        
        threading.Thread(target=self._collision_test_worker, daemon=True).start()

    def _collision_test_worker(self):
        """
        Background worker for the collision test.
        1. Move to Center (X, Y).
        2. Sweep Tilt (E) through the full range found in the loaded G-code.
        """
        try:
            self.is_manual_command_running = True
            
            # --- 1. Validation & Range Calculation ---
            if not self.processed_gcode:
                raise Exception("Error: Must load test profile prior to collision test.")

            min_e = 0.0
            max_e = 0.0
            found_e = False

            # Scan processed G-code for E limits
            # processed_gcode contains strings like "G1 X10 Y10 E45 F1000"
            for line in self.processed_gcode:
                if 'E' in line.upper():
                    import re as local_re
                    match = local_re.search(r"E([-+]?\d*\.?\d+)", line)
                    if match:
                        val = float(match.group(1))
                        if not found_e:
                            min_e = val
                            max_e = val
                            found_e = True
                        else:
                            if val < min_e: min_e = val
                            if val > max_e: max_e = val
            
            if not found_e:
                # Fallback if no E moves found (e.g. flat print)
                self.queue_message("No rotation (E) found in G-code. Using default +/- 10°.", "WARN")
                min_e = -10.0
                max_e = 10.0
            
            # Add a small safety buffer or use exact? 
            # Request said "full range of rotation requested". Let's stick to exact per file.
            self.queue_message(f"Test Range: E{min_e:.1f}° to E{max_e:.1f}°", "INFO")

            
            # --- 2. Execution ---

            # Get Center
            try:
                cx = float(self.center_x_var.get())
                cy = float(self.center_y_var.get())
            except ValueError:
                self.queue_message("Invalid Center Coordinates!", "ERROR")
                return

            speed = 1000 # mm/min or deg/min
            
            # Move to Center
            self.queue_message(f"Moving to Center ({cx}, {cy})...")
            cmd = f"G90\nG1 X{cx} Y{cy} F{speed}\nM400\n"
            self.serial_connection.write(cmd.encode('utf-8'))
            if not self._wait_for_ok(timeout=30): raise Exception("Move to Center timeout")

            # Tilt Sweep
            # Move to Min E
            self.queue_message(f"Tilting to Min ({min_e:.1f}°)...")
            self.serial_connection.write(f"G1 E{min_e:.2f} F{speed}\nM400\n".encode('utf-8'))
            if not self._wait_for_ok(timeout=30): raise Exception("Tilt Min timeout")

            # Move to Max E
            self.queue_message(f"Tilting to Max ({max_e:.1f}°)...")
            self.serial_connection.write(f"G1 E{max_e:.2f} F{speed}\nM400\n".encode('utf-8'))
            if not self._wait_for_ok(timeout=60): raise Exception("Tilt Max timeout")

            # Return to 0
            self.queue_message("Returning to 0°...")
            self.serial_connection.write(f"G1 E0 F{speed}\nM400\n".encode('utf-8'))
            if not self._wait_for_ok(timeout=30): raise Exception("Return 0 timeout")

            self.queue_message("Collision Test Complete.", "SUCCESS")
            
        except Exception as e:
            self.queue_message(f"Test Failed: {e}", "ERROR")
            if "Error: Must load" in str(e):
                messagebox.showerror("Setup Error", str(e))

        finally:
            self.is_manual_command_running = False
            # Update UI via queue or invoke
            self.root.after(0, self._reset_test_ui)

    def _reset_test_ui(self):
        """Re-enables the test screen buttons."""
        if hasattr(self, 'btn_begin_test'):
            self.btn_begin_test.config(state=tk.NORMAL)
            self.btn_exit_test.config(state=tk.NORMAL)
            self.lbl_test_status.config(text="Test Complete", fg=self.COLOR_TEXT_PRIMARY)


# --- Status Indicator Widget ---
class StatusIndicator(tk.Canvas):
    """
    A custom widget that displays a glowing, colored "LED" to indicate status.
    It supports off, on (green), busy (amber), and error (red) states, with
    a pulsing animation for the 'on' and 'busy' states.
    """
    def __init__(self, parent, bg, size=12):
        super().__init__(parent, width=size, height=size, bg=bg, highlightthickness=0)
        
        self.size = size
        self.colors = {
            "off": ("#444", "#555"),
            "on": ("#2a843d", "#3fb950"),      # Green
            "busy": ("#b37400", "#ffa657"),   # Amber
            "error": ("#990000", "#ff4444")   # Red
        }
        self.current_state = "off"
        self.pulse_on = False
        self.pulse_job = None # To store the ID of the 'after' job for pulsing

        # Create the circle that represents the LED.
        self.led = self.create_oval(2, 2, size-2, size-2, fill=self.colors["off"][0], outline=self.colors["off"][1])
        self.set_status("off")

    def set_status(self, state="off"):
        """
        Changes the color and animation of the LED based on the desired state.

        Args:
            state (str): The new state. Can be "off", "on", "busy", or "error".
        """
        if state not in self.colors:
            state = "off"
        if state == self.current_state:
            return # No change needed
            
        self.current_state = state
        
        # Cancel any existing pulse animation before starting a new one.
        if self.pulse_job:
            self.after_cancel(self.pulse_job)
            self.pulse_job = None
            
        if state == "on" or state == "busy":
            self._pulse_animation()
        else:
            # For solid states, just set the color directly.
            self.itemconfig(self.led, fill=self.colors[state][1], outline=self.colors[state][1])

    def _pulse_animation(self):
        """
        The internal method that creates the pulsing effect for the LED.
        
        It alternates between two shades of the current color and reschedules
        itself to create a continuous animation.
        """
        if self.current_state not in ("on", "busy"):
            return
            
        color1, color2 = self.colors[self.current_state]
        
        # Alternate between the two colors.
        if self.pulse_on:
            self.itemconfig(self.led, fill=color1, outline=color1)
            self.pulse_on = False
        else:
            self.itemconfig(self.led, fill=color2, outline=color2)
            self.pulse_on = True
            
        # Schedule the next frame of the animation.
        self.pulse_job = self.after(800, self._pulse_animation)

# --- Main Execution ---
if __name__ == "__main__":
    root = tk.Tk()
    
    # Pre-load fonts if possible to prevent a flash of unstyled text on startup.
    try:
        from tkinter import font
        font.Font(family="Orbitron", size=13).metrics()
        font.Font(family="Inter", size=11).metrics()
        font.Font(family="JetBrains Mono", size=10).metrics()
        font.Font(family="Space Mono", size=16, weight="bold").metrics()
        font.Font(family="Rajdhani", size=13, weight="bold").metrics()
    except Exception as e:
        # This is not critical, so we just print a note if it fails.
        print(f"Font loading note: {e}")
        
    app = GCodeSenderGUI(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()