import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import serial # type: ignore
import time
import re
import threading
import queue
import serial.tools.list_ports # For port scanning
import csv
from datetime import datetime
import os
import webbrowser
from typing import Any, Optional, Dict, List, Tuple

DMM_IP = "10.247.103"

# Optional PyVISA import for DMM control
try:
    from pyvisa import ResourceManager # type: ignore
    HAS_PYVISA = True
except ImportError:
    ResourceManager = None
    HAS_PYVISA = False

# Optional numpy import for memory optimization
try:
    import numpy as np # type: ignore
    HAS_NUMPY = True
except ImportError:
    HAS_NUMPY = False

# --- DMM Integration Classes ---
DMM_IP_PREFIX = '10.247.103'

DMM_CONFIG = [
    [102, 100, 'VINV'], # Moved to top as it's the primary device
    [120, 100, 'VINP'],
    [104, 100, 'IINP', 1e3],
    [107, 100, 'VSYS'],
    [103, 100, 'SAUX', 1e3],
    [109, 100, 'SINV', 1e3],
]

# Map of user-friendly names to SCPI mode strings
DMM_MODES = {
    "DC Voltage": "VOLT:DC",
    "AC Voltage": "VOLT:AC",
    "DC Current": "CURR:DC",
    "AC Current": "CURR:AC",
    "2W Resistance": "RES",
    "4W Resistance": "FRES",
    "Frequency": "FREQ",
    "Continuity": "CONT",
    "Diode": "DIOD",
}


class DmmInst:
    def __init__(self, id: int, samples: int, name: str, scale: float = 1) -> None:
        self.id = id
        self.samples = samples
        self.name = name
        self.scale = scale
        self.pv: Any = None

    def connect(self, pvrmgr: Any, ip_prefix: str = DMM_IP_PREFIX) -> None:
        if not pvrmgr: return
        
        # Try three common resource string formats
        resource_formats = [
            f'TCPIP0::{ip_prefix}.{self.id}::inst0::INSTR',
            f'TCPIP::{ip_prefix}.{self.id}::INSTR',
            f'TCPIP::{ip_prefix}.{self.id}::5025::SOCKET' # Added port-specific socket fallback
        ]
        
        errors = []
        for id_str in resource_formats:
            try:
                self.pv = pvrmgr.open_resource(id_str)
                self.pv.timeout = 60000  # 60 second VISA timeout
                # Test connectivity with an IDN query
                self.pv.query("*IDN?")
                return # Success!
            except Exception as e:
                if self.pv:
                    try: self.pv.close()
                    except: pass
                self.pv = None
                errors.append(f"{id_str}: {e}")
        
        # If we get here, both formats failed
        raise ConnectionError(
            f"Failed to connect to DMM '{self.name}' at {ip_prefix}.{self.id}.\n"
            f"Tried formats: {', '.join(resource_formats)}\n"
            f"Errors: {'; '.join(errors)}"
        )

    def setup(self, mode: str = 'VOLT:DC') -> None:
        if self.pv:
            self.pv.write(f'CONF:{mode}')
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
            return float(self.pv.query_ascii_values('CALC:AVER:ALL?')[0]) * self.scale
        return 0.0


class DmmGroup:
    def __init__(self, config: list) -> None:
        self.all_dmms: list[DmmInst] = []
        for info in config:
            self.all_dmms.append(DmmInst(*info)) # type: ignore
        self.dmms: list[DmmInst] = [] # Only successfully connected DMMs go here
        self.pvrmgr: Any = None

    def initialize(self, mode: str = 'VOLT:DC', ip_prefix: str = DMM_IP_PREFIX) -> None:
        if not HAS_PYVISA or not ResourceManager:
            raise ImportError(
                "PyVISA is not installed. Install with:\n"
                "  pip install pyvisa pyvisa-py"
            )
        try:
            # Explicitly force the pyvisa-py backend for Linux/Raspberry Pi
            self.pvrmgr = ResourceManager('@py')
        except Exception:
            # Fallback for Windows where default NI-VISA is used
            self.pvrmgr = ResourceManager()
            
        self.dmms = []
        errors = []
        for dmm in self.all_dmms:
            try:
                # Attempt to connect to this DMM
                dmm.connect(self.pvrmgr, ip_prefix)
                dmm.setup(mode)
                self.dmms.append(dmm)
                # User preference: Only connect to one multimeter at a time.
                # Stop as soon as the first successful connection is made.
                break
            except Exception as e:
                # Suppress individual errors, just collect them in case all fail
                errors.append(f"{dmm.name} (ID {dmm.id}): {e}")
                continue

        if not self.dmms:
            # Only raise an error if NO multimeters were connectable
            raise ConnectionError(
                "No DMMs could be connected. Please check network/power.\n" + 
                "\n".join(errors)
            )

    def trigger(self) -> None:
        for dmm in self.dmms:
            dmm.trigger()

    def read(self) -> list:
        # Block until all connected DMMs are ready
        if not self.dmms: return []
        
        ready = False
        start_time = time.time()
        timeout = 60.0 # 60 seconds timeout

        while not ready:
            if time.time() - start_time > timeout:
                print(f"TIMEOUT: DMM measurement took longer than {timeout}s.")
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
    # --- Type Annotations for Dynamic Attributes ---
    header_status_indicator: Any
    center_x_entry: ttk.Entry
    center_y_entry: ttk.Entry
    center_z_entry: ttk.Entry
    mark_center_button: ttk.Button
    collision_test_button: ttk.Button
    port_combobox: ttk.Combobox
    baud_entry: ttk.Entry
    connect_button: ttk.Button
    cancel_connect_button: ttk.Button
    status_indicator: Any
    status_label: ttk.Label
    dmm_connect_button: ttk.Button
    auto_measure_check: ttk.Checkbutton
    log_check: ttk.Checkbutton
    log_path_entry: ttk.Entry
    browse_log_btn: ttk.Button
    measure_button: ttk.Button
    start_button: ttk.Button
    pause_resume_button: ttk.Button
    stop_button: ttk.Button
    quick_stop_button: ttk.Button
    progress_label: ttk.Label
    progress_bar: ttk.Progressbar
    toggle_2d_button: ttk.Button
    abs_button: ttk.Button
    rel_button: ttk.Button
    canvas_frame: ttk.Frame
    goto_x_entry: ttk.Entry
    goto_y_entry: ttk.Entry
    goto_z_entry: ttk.Entry
    goto_e_entry: ttk.Entry
    go_button: ttk.Button
    go_to_center_button: ttk.Button
    xy_canvas: tk.Canvas
    z_canvas: tk.Canvas
    e_container: ttk.Frame
    e_canvas: tk.Canvas
    goto_controls: List[Any]
    jog_z_pos: ttk.Button
    jog_z_neg: ttk.Button
    jog_y_pos: ttk.Button
    jog_x_neg: ttk.Button
    home_button: ttk.Button
    jog_x_pos: ttk.Button
    jog_y_neg: ttk.Button
    jog_e_pos: ttk.Button
    jog_e_neg: ttk.Button
    jog_step_entry: ttk.Entry
    jog_feedrate_entry: ttk.Entry
    rot_step_entry: ttk.Entry
    rot_feedrate_entry: ttk.Entry
    manual_buttons: List[ttk.Button]
    manual_entries: List[ttk.Entry]
    log_area: scrolledtext.ScrolledText
    terminal_input: ttk.Entry
    terminal_send_button: ttk.Button
    plot_container_frame: ttk.Frame
    toggle_3d_button: ttk.Button
    _simplified_indices_cache: list

    def __init__(self, parent_frame):
        self.parent = parent_frame
        self.root = parent_frame.winfo_toplevel()

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
        self.COLOR_PENDING_RING = "#c4c1ff"     # Color for incomplete frame borders and pending buttons
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
        style.configure('Subtle.DRO.TLabel', font=self.FONT_BODY_SMALL, padding=(5, 0), background=self.COLOR_BLACK, foreground=self.COLOR_TEXT_SECONDARY, borderwidth=0, relief='flat', anchor='e')

        # --- Button Styles ---
        style.configure('TButton', background=self.COLOR_PANEL_BG, foreground=self.COLOR_TEXT_PRIMARY, bordercolor=self.COLOR_BORDER, borderwidth=1, relief=tk.SOLID, padding=(12, 8), font=self.FONT_BODY)
        style.map('TButton', background=[('active', '#2c333e'), ('pressed', self.COLOR_BLACK)], foreground=[('active', self.COLOR_ACCENT_CYAN)], bordercolor=[('active', self.COLOR_ACCENT_CYAN)])

        # --- Ringed Button Styles ---
        style.configure('YellowRing.TButton', background=self.COLOR_PANEL_BG, foreground=self.COLOR_TEXT_PRIMARY, bordercolor=self.COLOR_PENDING_RING, borderwidth=1, relief=tk.SOLID, padding=(12, 8), font=self.FONT_BODY)
        style.map('YellowRing.TButton', background=[('active', '#2c333e'), ('pressed', self.COLOR_BLACK)], foreground=[('active', self.COLOR_PENDING_RING)], bordercolor=[('active', self.COLOR_PENDING_RING)])
        
        style.configure('GreenRing.TButton', background=self.COLOR_PANEL_BG, foreground=self.COLOR_TEXT_PRIMARY, bordercolor=self.COLOR_ACCENT_GREEN, borderwidth=1, relief=tk.SOLID, padding=(12, 8), font=self.FONT_BODY)
        style.map('GreenRing.TButton', background=[('active', '#2c333e'), ('pressed', self.COLOR_BLACK)], foreground=[('active', self.COLOR_ACCENT_GREEN)], bordercolor=[('active', self.COLOR_ACCENT_GREEN)])

        style.configure('PurpleRing.TButton', background=self.COLOR_PANEL_BG, foreground=self.COLOR_TEXT_PRIMARY, bordercolor='#BA55D3', borderwidth=1, relief=tk.SOLID, padding=(12, 8), font=self.FONT_BODY)
        style.map('PurpleRing.TButton', background=[('active', '#2c333e'), ('pressed', self.COLOR_BLACK)], foreground=[('active', '#BA55D3')], bordercolor=[('active', '#BA55D3')])

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
        style.map('Custom.Toggle.On.TButton', bordercolor=[('active', self.COLOR_ACCENT_CYAN)])

        # --- Small Button Style for secondary tools ---
        style.configure('Small.TButton', background=self.COLOR_PANEL_BG, foreground=self.COLOR_TEXT_PRIMARY, bordercolor=self.COLOR_BORDER, font=custom_font, padding=padding)
        style.map('Small.TButton', bordercolor=[('active', self.COLOR_ACCENT_CYAN)], foreground=[('active', self.COLOR_ACCENT_CYAN)])

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
                        selectbackground=self.COLOR_BLACK,
                        selectforeground=self.COLOR_ACCENT_CYAN,
                        padding=6,
                        font=self.FONT_MONO)
        style.map('TCombobox', 
                  bordercolor=[('focus', self.COLOR_ACCENT_CYAN)],
                  fieldbackground=[('readonly', self.COLOR_BLACK)],
                  background=[('readonly', self.COLOR_BLACK)],
                  selectbackground=[('readonly', self.COLOR_BLACK)],
                  selectforeground=[('readonly', self.COLOR_ACCENT_CYAN)])

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

        # --- Green-border LabelFrame (section complete) ---
        style.configure('Green.TLabelframe',
                        background=self.COLOR_PANEL_BG,
                        bordercolor=self.COLOR_ACCENT_GREEN,
                        relief=tk.SOLID, borderwidth=2)
        style.configure('Green.TLabelframe.Label',
                        background=self.COLOR_PANEL_BG,
                        foreground=self.COLOR_ACCENT_GREEN,
                        font=self.FONT_BODY_BOLD)
        style.configure('Grey.TLabelframe',
                        background=self.COLOR_PANEL_BG,
                        bordercolor=self.COLOR_BORDER,
                        relief=tk.SOLID, borderwidth=1)
        style.configure('Grey.TLabelframe.Label',
                        background=self.COLOR_PANEL_BG,
                        foreground=self.COLOR_TEXT_SECONDARY,
                        font=self.FONT_BODY_BOLD)

        style.configure('Yellow.TLabelframe',
                        background=self.COLOR_PANEL_BG,
                        bordercolor=self.COLOR_PENDING_RING,
                        relief=tk.SOLID, borderwidth=2)
        style.configure('Yellow.TLabelframe.Label',
                        background=self.COLOR_PANEL_BG,
                        foreground=self.COLOR_PENDING_RING,
                        font=self.FONT_BODY_BOLD)

        style.configure('Purple.TLabelframe',
                        background=self.COLOR_PANEL_BG,
                        bordercolor='#BA55D3',
                        relief=tk.SOLID, borderwidth=3)
        style.configure('Purple.TLabelframe.Label',
                        background=self.COLOR_PANEL_BG,
                        foreground='#BA55D3',
                        font=self.FONT_BODY_BOLD)
        
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
        self.hardware_fault = False
        self.gcode_filepath = None # Store path to G-code file instead of contents
        self.processed_gcode = []
        self.is_sending = False
        self.is_paused = False
        self.is_manual_command_running = False
        self.is_collision_test_running = False
        self.is_calibrating = False
        self.rotation_crash_test_complete = False
        
        # Threading events for controlling background tasks (sending G-code, connecting)
        self.stop_event = threading.Event()
        self.pause_event = threading.Event()
        self.pause_event.set() # Initialize in the "go" state (not paused)
        self.cancel_connect_event = threading.Event()
        
        # Lock to serialize all serial port writes across threads.
        # emergency_stop() acquires this to guarantee M112 is never interleaved
        # with a G-code command that a background thread is mid-write.
        self.serial_lock = threading.Lock()
        
        # A queue for passing messages from background threads to the main GUI thread
        self.message_queue = queue.Queue()
        
        # For the terminal's command history
        self.command_history = []
        self.history_index = 0
        
        # Data structures for visualizing the G-code toolpath
        self.toolpath_by_layer: Dict[float, List[Any]] = {}   # {z_level: [((x1,y1),(x2,y2)), ...], ...}
        self.move_to_layer_map: List[Any] = []   # [(z_level, index_on_layer), ...]
        self.ordered_z_values: List[float] = []    # [z1, z2, z3, ...]
        self.completed_move_count: int = 0 # How many moves have been completed so far
        self.after_id: Any = None          # To store the ID of the recurring 'after' job

        # Cache for the 3D plot coordinates to avoid recalculating them on every redraw
        self._plot_coords_cache = None
        self._plot_cache_valid = False

        # Control variable for enabling/disabling the 3D plot for performance
        self.is_3d_plot_enabled = tk.BooleanVar(value=True)
        self._3d_control_bar: Optional[ttk.Frame] = None  # Will be set in create_3d_display_panel

        # Control variable for enabling/disabling the 2D plots for performance
        self.is_2d_plot_enabled = tk.BooleanVar(value=True)

        # --- Printer Physical Bounds (in mm) ---
        # E is 'Rotation' (repurposed extruder), units are now DEGREES. Firmware must be configured accordingly (e.g., M92 E8.888).
        self.PRINTER_BOUNDS: Dict[str, float] = { 'x_min': 0, 'x_max': 220, 'y_min': 0, 'y_max': 220, 'z_min': 0, 'z_max': 130, 'e_min': -90, 'e_max': 90 }

        # --- Tkinter StringVars (for dynamically updating GUI labels) ---
        self.file_path_var = tk.StringVar(value="No file selected")
        self.center_x_var = tk.StringVar(value="110.0")
        self.center_y_var = tk.StringVar(value="110.0")
        self.center_z_var = tk.StringVar(value="0.0")
        self.center_e_var = tk.StringVar(value="0.0")
        self.available_ports = ["Auto-detect"] + self._get_available_ports()
        self.port_var = tk.StringVar(value=self.available_ports[0] if self.available_ports else "")
        self.baud_var = tk.StringVar(value="115200")
        self.connection_status_var = tk.StringVar(value="Status: Disconnected")

        # --- DMM / Measurement State ---
        self.dmm_ip_prefix_var = tk.StringVar(value=DMM_IP_PREFIX)
        self.dmm_group = None
        self.is_dmm_connected = False
        self.auto_measure_enabled = tk.BooleanVar(value=True)
        self.log_measurements_enabled = tk.BooleanVar(value=True)
        self.measurement_log_file = None # Internal file handle/flag
        self.log_filepath_var = tk.StringVar(value="") # Initialize empty, set on G-code load
        self.dmm_status_var = tk.StringVar(value="DMMs: Disconnected")
        self.dmm_mode_var = tk.StringVar(value="DC Voltage")
        self.last_measurement_var = tk.StringVar(value="Last: --")
        self.pre_measure_delay_var = tk.DoubleVar(value=0.2) # seconds


        self.jog_step_var = tk.StringVar(value="10")
        self.jog_feedrate_var = tk.StringVar(value="1000")
        self.rotation_step_var = tk.StringVar(value="5")
        self.rotation_feedrate_var = tk.StringVar(value="3000")
        self.mm_per_degree_var = tk.DoubleVar(value=8.888) # Calibration: MM extrusion per Degree of tilt
        
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
        self.last_cmd_abs_x: Optional[float] = None # Use None to indicate the position is not yet known.
        self.last_cmd_abs_y: Optional[float] = None
        self.last_cmd_abs_z: Optional[float] = None
        self.last_cmd_abs_e = None
        
        # The current coordinate display mode ('absolute' or 'relative' to the center point).
        self.coord_mode = tk.StringVar(value="absolute")

        # --- Build the main GUI layout ---
        # Container for the standard view (Header, Panels, Footer)
        self.main_view_frame = ttk.Frame(self.parent, style='TFrame')
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
        
        self.display_tab = display_tab  # Keep reference for tab-active checks
        self.notebook.add(cli_tab, text="Command Line")
        self.notebook.add(display_tab, text="3D View")
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)
        
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

        # --- Populate the GUI panels (in setup-flow order) ---
        self.create_connection_frame(self.left_panel_scrollable)
        self.create_measurement_frame(self.left_panel_scrollable)
        self.create_manual_control_frame(self.left_panel_scrollable)
        self.create_file_center_frame(self.left_panel_scrollable)   # SETUP (center XYZ)
        self.create_control_frame(self.left_panel_scrollable)        # EXECUTION CONTROL (file + progress merged in)
        self.create_position_control_frame(self.left_panel_scrollable)

        # Variable traces for green-border feedback
        self.log_measurements_enabled.trace_add('write', lambda varname, index, mode: self._update_section_borders())
        self.log_filepath_var.trace_add('write', lambda varname, index, mode: self._update_section_borders())
        self.file_path_var.trace_add('write', lambda varname, index, mode: self._update_section_borders())

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
        
        # --- Global Key Bindings ---
        self.root.bind('<Escape>', lambda e: self.emergency_stop())
        # Globally remove focus from comboboxes after selection to prevent text highlighting
        self.root.bind_class('TCombobox', '<<ComboboxSelected>>', lambda e: self.root.focus_set())
        
        # Bind global keyboard shortcuts for jogging
        self.root.bind('<Key>', self._handle_key_press)
        


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
        """Creates the 'SETUP' panel for defining the center point and running calibration."""
        self._setup_frame = ttk.LabelFrame(parent, text="SETUP", padding="10", style='Yellow.TLabelframe')
        frame = self._setup_frame
        frame.pack(fill=tk.X, pady=(0, 10), padx=5)
        frame.columnconfigure(0, weight=1); frame.columnconfigure(1, weight=1); frame.columnconfigure(2, weight=1); frame.columnconfigure(3, weight=1); frame.columnconfigure(4, weight=1)

        # Row 0: Setup Action Buttons
        self.mark_center_button = ttk.Button(frame, text="Mark Current as Center", command=self._mark_current_as_center, state=tk.DISABLED, style='YellowRing.TButton')
        self.mark_center_button.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0,10), padx=(0,5))

        self.collision_test_button = ttk.Button(frame, text="Collision Avoidance Test", command=self._open_collision_test_screen, state=tk.DISABLED, style='YellowRing.TButton')
        self.collision_test_button.grid(row=0, column=2, columnspan=3, sticky="ew", pady=(0,10), padx=(5,0))

        # Row 1: Labels
        ttk.Label(frame, text="Center X:").grid(row=1, column=0, sticky="w", padx=(0, 5))
        ttk.Label(frame, text="Center Y:").grid(row=1, column=1, sticky="w", padx=(5, 5))
        ttk.Label(frame, text="Center Z:").grid(row=1, column=2, sticky="w", padx=(5, 5))
        ttk.Label(frame, text="Center Tilt:").grid(row=1, column=3, sticky="w", padx=(5, 5))
        ttk.Label(frame, text="E-Cal (mm/°):").grid(row=1, column=4, sticky="w", padx=(5, 5))

        # Row 2: Entries and Lock Toggle
        self.center_x_entry = ttk.Entry(frame, textvariable=self.center_x_var, width=8)
        self.center_x_entry.grid(row=2, column=0, sticky="ew", pady=(2,0), padx=(0, 5))

        self.center_y_entry = ttk.Entry(frame, textvariable=self.center_y_var, width=8)
        self.center_y_entry.grid(row=2, column=1, sticky="ew", pady=(2,0), padx=(5, 5))

        self.center_z_entry = ttk.Entry(frame, textvariable=self.center_z_var, width=8)
        self.center_z_entry.grid(row=2, column=2, sticky="ew", pady=(2,0), padx=(5, 5))
        
        self.center_e_entry = ttk.Entry(frame, textvariable=self.center_e_var, width=8)
        self.center_e_entry.grid(row=2, column=3, sticky="ew", pady=(2,0), padx=(5, 5))

        ecal_container = ttk.Frame(frame, style='Panel.TFrame')
        ecal_container.grid(row=2, column=4, sticky="ew", pady=(2,0), padx=(5, 5), columnspan=2)

        self.e_cal_entry = ttk.Entry(ecal_container, textvariable=self.mm_per_degree_var, width=8, state='readonly')
        self.e_cal_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.e_cal_lock_var = tk.BooleanVar(value=True)
        
        def toggle_e_cal():
            is_locked = not self.e_cal_lock_var.get()
            
            if not is_locked:
                if not messagebox.askyesno("Confirm Unlock", "Are you sure you want to change the rotation ratio? This will affect tilt axis movements and could cause collisions."):
                    return
                self.e_cal_lock_var.set(False)
                self.e_cal_entry.config(state='normal')
                self.e_cal_lock_button.config(text="🔓")
            else:
                self.e_cal_lock_var.set(True)
                self.e_cal_entry.config(state='readonly')
                self.e_cal_lock_button.config(text="🔒")

        self.e_cal_lock_button = ttk.Button(ecal_container, text="🔒", width=3, style='ViewCube.TButton', command=toggle_e_cal)
        self.e_cal_lock_button.pack(side=tk.LEFT, padx=(2,0))

        # Bind changes in the center entries to update the coordinate displays.
        self.center_x_entry.bind('<FocusOut>', self._on_center_change); self.center_x_entry.bind('<Return>', self._on_center_change)
        self.center_y_entry.bind('<FocusOut>', self._on_center_change); self.center_y_entry.bind('<Return>', self._on_center_change)
        self.center_z_entry.bind('<FocusOut>', self._on_center_change); self.center_z_entry.bind('<Return>', self._on_center_change)
        self.center_e_entry.bind('<FocusOut>', self._on_center_change); self.center_e_entry.bind('<Return>', self._on_center_change)

    def create_connection_frame(self, parent):
        """Creates the 'CONNECTION' panel for managing the serial connection."""
        self._conn_frame = ttk.LabelFrame(parent, text="CONNECTION", padding="10", style='Yellow.TLabelframe')
        frame = self._conn_frame
        frame.pack(fill=tk.X, pady=(0, 10), padx=5)
        
        # Spacer column - pushes the status label to the far right
        frame.columnconfigure(7, weight=1)
        
        # Row 0
        self.connect_button = ttk.Button(frame, text="Connect", command=self.toggle_connection, width=13, style='YellowRing.TButton')
        self.connect_button.grid(row=0, column=0, sticky="w", padx=(0, 10))

        ttk.Label(frame, text="Port:").grid(row=0, column=1, sticky="w")
        self.port_combobox = ttk.Combobox(frame, textvariable=self.port_var, values=self.available_ports, width=15, state="readonly", font=self.FONT_MONO)
        self.port_combobox.grid(row=0, column=2, padx=(0, 5))
        
        ttk.Button(frame, text="Rescan", command=self.rescan_ports, width=7).grid(row=0, column=3, padx=(0, 10))

        ttk.Label(frame, text="Baud Rate:").grid(row=0, column=4, sticky="w", padx=(5,0))
        self.baud_entry = ttk.Entry(frame, textvariable=self.baud_var, width=10) 
        self.baud_entry.grid(row=0, column=5, padx=(0, 10), sticky="w")
        
        # The cancel button appears in column 6 when needed
        self.cancel_connect_button = ttk.Button(frame, text="Cancel", command=self._cancel_connection_attempt)
        # We grid it here but use grid_remove to hide it initially without shifting columns 0-5
        self.cancel_connect_button.grid(row=0, column=6, sticky="w", padx=(0, 10))
        self.cancel_connect_button.grid_remove()
        
        # The status text label.
        self.status_label = ttk.Label(frame, textvariable=self.connection_status_var, font=self.FONT_BODY_SMALL, style='Filepath.TLabel') 
        self.status_label.grid(row=0, column=8, padx=(0, 0), sticky="e")
        

    def create_measurement_frame(self, parent):
        """Creates the 'MEASUREMENT' panel."""
        self._meas_frame = ttk.LabelFrame(parent, text="MEASUREMENT", padding="10", style='Yellow.TLabelframe')
        frame = self._meas_frame
        frame.pack(fill=tk.X, pady=(0, 10), padx=5)
        
        # Top Row: Connect, IP, Mode, Settling Time
        top_frame = ttk.Frame(frame, style='Panel.TFrame')
        top_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.dmm_connect_button = ttk.Button(top_frame, text="Connect DMMs", command=self.toggle_dmm_connection, style='YellowRing.TButton')
        self.dmm_connect_button.pack(side=tk.LEFT, padx=(0, 15))
        
        ttk.Label(top_frame, text="IP Prefix:", font=self.FONT_BODY_SMALL).pack(side=tk.LEFT)
        self.dmm_ip_prefix_entry = ttk.Entry(top_frame, textvariable=self.dmm_ip_prefix_var, width=12, font=self.FONT_MONO)
        self.dmm_ip_prefix_entry.pack(side=tk.LEFT, padx=(2, 15))

        ttk.Label(top_frame, text="Mode:", font=self.FONT_BODY_SMALL).pack(side=tk.LEFT)
        self.dmm_mode_combo = ttk.Combobox(
            top_frame, textvariable=self.dmm_mode_var,
            values=list(DMM_MODES.keys()), state="readonly",
            width=12, font=self.FONT_MONO
        )
        self.dmm_mode_combo.pack(side=tk.LEFT, padx=(2, 15))
        self.dmm_mode_combo.bind('<<ComboboxSelected>>', self._on_dmm_mode_change)

        ttk.Label(top_frame, text="Settling Time (s):", font=self.FONT_BODY_SMALL).pack(side=tk.LEFT)
        ttk.Entry(top_frame, textvariable=self.pre_measure_delay_var, width=4, font=self.FONT_MONO).pack(side=tk.LEFT, padx=(2, 0))

        self.auto_measure_check = ttk.Checkbutton(top_frame, text="Auto-Measure on Move", variable=self.auto_measure_enabled, command=self._on_auto_measure_toggle)
        self.auto_measure_check.pack(side=tk.LEFT, padx=(15, 0))

        # Mid Row: Logging
        mid_frame = ttk.Frame(frame, style='Panel.TFrame')
        mid_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.log_check = ttk.Checkbutton(mid_frame, text="Log to CSV:", variable=self.log_measurements_enabled)
        self.log_check.grid(row=0, column=0, sticky="w")
        
        self.log_path_entry = ttk.Entry(mid_frame, textvariable=self.log_filepath_var, font=self.FONT_BODY_SMALL, state='readonly')
        self.log_path_entry.grid(row=0, column=1, sticky="ew", padx=(5, 5))
        
        self.browse_log_btn = ttk.Button(mid_frame, text="Select Data File...", command=self.select_log_file, style='YellowRing.TButton')
        self.browse_log_btn.grid(row=0, column=2, sticky="e")
        
        mid_frame.columnconfigure(1, weight=1)

        # Control Row: Manual Trigger & Last Reading
        ctrl_frame = ttk.Frame(frame, style='Panel.TFrame')
        ctrl_frame.pack(fill=tk.X)
        
        self.measure_button = ttk.Button(ctrl_frame, text="▶ Measure Now", command=self.trigger_manual_measurement, state=tk.DISABLED, width=25)
        self.measure_button.pack(side=tk.LEFT, padx=(0, 15))
        
        ttk.Label(ctrl_frame, textvariable=self.last_measurement_var, font=self.FONT_MONO).pack(side=tk.LEFT)

    def _on_auto_measure_toggle(self):
        if self.auto_measure_enabled.get():
            self.log_message("Auto-measurement ENABLED. DMMs will trigger after every move.", "INFO")
        else:
            self.log_message("Auto-measurement DISABLED.", "INFO")

    def _on_dmm_mode_change(self, event=None):
        """Called when the user selects a new DMM measurement mode from the dropdown."""
        mode_name = self.dmm_mode_var.get()
        scpi_mode = DMM_MODES.get(mode_name, 'VOLT:DC')
        self.log_message(f"DMM Mode changed to: {mode_name} ({scpi_mode})", "INFO")
        
        # Check if log exists and has data
        filepath = self.log_filepath_var.get()
        if filepath and os.path.exists(filepath):
            try:
                if os.path.getsize(filepath) > 0:
                    if messagebox.askyesno("Change Log File?", f"The DMM mode has been changed to '{mode_name}', but the current log file already contains data from a previous mode.\n\nWould you like to select a new log file to prevent conflicting data units?"):
                        self.select_log_file()
            except Exception as e:
                self.log_message(f"Could not check log file size: {e}", "WARN")
                
        # If DMMs are already connected, reconfigure them live
        if self.is_dmm_connected and self.dmm_group:
            try:
                for dmm in self.dmm_group.dmms:
                    if dmm.pv:
                        dmm.pv.write(f'CONF:{scpi_mode}')
                        dmm.pv.write(f'SAMP:COUN {dmm.samples}')
                        dmm.pv.write(f'CALC:AVER:STAT ON')
                self.log_message(f"All DMMs reconfigured to {mode_name}.", "SUCCESS")
            except Exception as e:
                self.log_message(f"Error changing DMM mode: {e}", "ERROR")


    def create_control_frame(self, parent):
        """Creates the 'EXECUTION CONTROL' panel with file picker, Start/Pause/Stop, and progress."""
        self._ctrl_frame = ttk.LabelFrame(parent, text="EXECUTION CONTROL", padding=10, style='Yellow.TLabelframe')
        frame = self._ctrl_frame
        frame.pack(fill=tk.X, pady=(0, 10), padx=5)
        frame.columnconfigure(1, weight=1)

        # --- File Picker Row (moved from SETUP) ---
        file_row = ttk.Frame(frame, style='Panel.TFrame')
        file_row.pack(fill=tk.X, pady=(0, 8))
        self.select_gcode_button = ttk.Button(file_row, text="Select G-Code File", command=self.select_file, style='YellowRing.TButton')
        self.select_gcode_button.pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(file_row, text="Clear", command=self.clear_file).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Label(file_row, textvariable=self.file_path_var, wraplength=250, style='Filepath.TLabel').pack(side=tk.LEFT, fill=tk.X, expand=True)

        # --- Action Buttons Row ---
        btn_row = ttk.Frame(frame, style='Panel.TFrame')
        btn_row.pack(fill=tk.X)

        self.start_button = ttk.Button(btn_row, text="Start Sending", command=self.start_sending, state=tk.DISABLED, style='Primary.TButton')
        self.start_button.pack(side=tk.LEFT, padx=(0, 5))

        self.pause_resume_button = ttk.Button(btn_row, text="Pause", command=self.toggle_pause_resume, state=tk.DISABLED)
        self.pause_resume_button.pack(side=tk.LEFT, padx=(0, 10))

        # M112: A hard stop that requires a printer reset.
        self.stop_button = ttk.Button(btn_row, text="EMERGENCY STOP", command=self.emergency_stop, state=tk.NORMAL, style='Danger.TButton')
        self.stop_button.pack(side=tk.LEFT)

        # M410: A soft stop that finishes the current move then halts.
        self.quick_stop_button = ttk.Button(btn_row, text="QUICK STOP", command=self.quick_stop, state=tk.NORMAL, style='Amber.TButton')
        self.quick_stop_button.pack(side=tk.LEFT, padx=(10, 0))

        # --- Progress (merged from standalone PROGRESS section) ---
        self.progress_label = ttk.Label(frame, textvariable=self.progress_label_var, font=self.FONT_MONO, foreground=self.COLOR_TEXT_SECONDARY)
        self.progress_label.pack(fill=tk.X, padx=2, pady=(8, 2))

        self.progress_bar = ttk.Progressbar(frame, orient=tk.HORIZONTAL, mode='determinate', variable=self.progress_var)
        self.progress_bar.pack(fill=tk.X, padx=2)

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
        
        # Column Headers
        ttk.Label(dro_frame, text="", style='Subtle.DRO.TLabel').grid(row=0, column=0, sticky="ew") 
        ttk.Label(dro_frame, text="X", style='Subtle.DRO.TLabel').grid(row=0, column=1, sticky="ew")
        ttk.Label(dro_frame, text="Y", style='Subtle.DRO.TLabel').grid(row=0, column=2, sticky="ew", padx=5)
        ttk.Label(dro_frame, text="Z", style='Subtle.DRO.TLabel').grid(row=0, column=3, sticky="ew")
        ttk.Label(dro_frame, text="Tilt", style='Subtle.DRO.TLabel').grid(row=0, column=4, sticky="ew", padx=5)

        # Labels for the 'CURRENT' (last commanded) position.
        ttk.Label(dro_frame, text="CURRENT:", style='DRO.TLabel').grid(row=1, column=0, sticky="w"); 
        ttk.Label(dro_frame, textvariable=self.last_cmd_x_display_var, style='Red.DRO.TLabel').grid(row=1, column=1, sticky="ew")
        ttk.Label(dro_frame, textvariable=self.last_cmd_y_display_var, style='Red.DRO.TLabel').grid(row=1, column=2, sticky="ew", padx=5)
        ttk.Label(dro_frame, textvariable=self.last_cmd_z_display_var, style='Red.DRO.TLabel').grid(row=1, column=3, sticky="ew")
        ttk.Label(dro_frame, textvariable=self.last_cmd_e_display_var, style='Red.DRO.TLabel').grid(row=1, column=4, sticky="ew", padx=5)
        
        # Labels for the 'TARGET' (Go To) position.
        ttk.Label(dro_frame, text=" TARGET:", style='DRO.TLabel').grid(row=2, column=0, sticky="w"); 
        ttk.Label(dro_frame, textvariable=self.goto_x_display_var, style='Blue.DRO.TLabel').grid(row=2, column=1, sticky="ew")
        ttk.Label(dro_frame, textvariable=self.goto_y_display_var, style='Blue.DRO.TLabel').grid(row=2, column=2, sticky="ew", padx=5)
        ttk.Label(dro_frame, textvariable=self.goto_z_display_var, style='Blue.DRO.TLabel').grid(row=2, column=3, sticky="ew")
        ttk.Label(dro_frame, textvariable=self.goto_e_display_var, style='Blue.DRO.TLabel').grid(row=2, column=4, sticky="ew", padx=5)
        
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
        self.z_canvas = tk.Canvas(self.canvas_frame, width=25, height=canvas_size, bg=self.COLOR_BLACK, highlightthickness=1, highlightbackground=self.COLOR_BORDER)
        self.z_canvas.grid(row=0, column=2, sticky="ns", padx=30)
        self.z_canvas.bind("<Button-1>", self._on_z_canvas_click); self.z_canvas.bind("<B1-Motion>", self._on_z_canvas_click); self.z_canvas.bind("<Configure>", self._draw_z_canvas_marker)

        # E (Rotation) view - Circular Gauge with Snap Buttons
        # Container for the canvas and floating buttons
        self.e_container = ttk.Frame(self.canvas_frame, style='Panel.TFrame', width=canvas_size, height=canvas_size)
        self.e_container.grid(row=0, column=3, sticky="n", padx=(0, 10))
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

        # Bottom (0)
        ttk.Button(self.e_container, text=" 0° ", width=4, style=btn_style, command=lambda: set_e(0)).place(relx=0.5, rely=0.98, anchor='s')
        # Right (90)
        ttk.Button(self.e_container, text="+90°", width=4, style=btn_style, command=lambda: set_e(90)).place(relx=0.98, rely=0.5, anchor='e')
        # Left (-90)
        ttk.Button(self.e_container, text="-90°", width=4, style=btn_style, command=lambda: set_e(-90)).place(relx=0.02, rely=0.5, anchor='w')

        # --- "Mark Tilt as Level" button below the E gauge ---
        ttk.Button(
            self.canvas_frame, text="Mark Tilt as Level (0°)",
            command=self._mark_tilt_as_level
        ).grid(row=1, column=3, pady=(4, 0), sticky="ew")

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
        
        self.jog_y_neg = ttk.Button(jog_grid_frame, text="Y-", style='Jog.TButton', command=lambda: self._jog('Y', -1), state=tk.DISABLED); self.jog_y_neg.grid(row=0, column=1, padx=2, pady=2, sticky="nsew")
        self.jog_x_neg = ttk.Button(jog_grid_frame, text="X-", style='Jog.TButton', command=lambda: self._jog('X', -1), state=tk.DISABLED); self.jog_x_neg.grid(row=1, column=0, padx=2, pady=2, sticky="nsew")
        self.home_button = ttk.Button(jog_grid_frame, text="⌂", style='Home.TButton', command=self._home_all, state=tk.DISABLED); self.home_button.grid(row=1, column=1, padx=2, pady=2, sticky="nsew")
        self.jog_x_pos = ttk.Button(jog_grid_frame, text="X+", style='Jog.TButton', command=lambda: self._jog('X', 1), state=tk.DISABLED); self.jog_x_pos.grid(row=1, column=2, padx=2, pady=2, sticky="nsew")
        self.jog_y_pos = ttk.Button(jog_grid_frame, text="Y+", style='Jog.TButton', command=lambda: self._jog('Y', 1), state=tk.DISABLED); self.jog_y_pos.grid(row=2, column=1, padx=2, pady=2, sticky="nsew")
        
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
        """Creates a compact horizontal toolbar for 3D view controls, placed inline."""
        if not self.matplotlib_imported:
            return

        # Compact horizontal strip — all buttons in a single row
        controls_frame = ttk.Frame(parent, style="Panel.TFrame")
        controls_frame.pack(side=tk.RIGHT, padx=(10, 0))

        # Rotation arrows
        ttk.Button(controls_frame, text="↑", command=lambda: self._rotate_view(elev_change=15), style='ViewCube.TButton', width=2).pack(side=tk.LEFT, padx=1)
        ttk.Button(controls_frame, text="↓", command=lambda: self._rotate_view(elev_change=-15), style='ViewCube.TButton', width=2).pack(side=tk.LEFT, padx=1)
        ttk.Button(controls_frame, text="←", command=lambda: self._rotate_view(azim_change=15), style='ViewCube.TButton', width=2).pack(side=tk.LEFT, padx=1)
        ttk.Button(controls_frame, text="→", command=lambda: self._rotate_view(azim_change=-15), style='ViewCube.TButton', width=2).pack(side=tk.LEFT, padx=1)

        # Small separator
        ttk.Separator(controls_frame, orient='vertical').pack(side=tk.LEFT, fill=tk.Y, padx=4, pady=2)

        # Preset view buttons
        ttk.Button(controls_frame, text="Top", command=lambda: self._set_view(elev=90, azim=-90), style="Segment.TButton", width=4).pack(side=tk.LEFT, padx=1)
        ttk.Button(controls_frame, text="Front", command=lambda: self._set_view(elev=0, azim=-90), style="Segment.TButton", width=5).pack(side=tk.LEFT, padx=1)
        ttk.Button(controls_frame, text="Iso", command=lambda: self._set_view(elev=30, azim=-60), style="Segment.TButton", width=3).pack(side=tk.LEFT, padx=1)

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
                
                self.create_view_controls(self._3d_control_bar)
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
        self._3d_control_bar = ttk.Frame(parent, style="Panel.TFrame", padding=(5, 5))
        self._3d_control_bar.grid(row=0, column=0, sticky="ew")
        control_bar = self._3d_control_bar
        
        self.toggle_3d_button = ttk.Button(
            control_bar, 
            text="3D PLOT", 
            command=self._toggle_3d_plot_button
        )
        self.toggle_3d_button.pack(side=tk.LEFT)

        self.launch_visualizer_button = ttk.Button(
            control_bar,
            text="Launch Data Visualizer",
            style='Small.TButton',
            command=self.launch_visualizer
        )
        self.launch_visualizer_button.pack(side=tk.LEFT, padx=(10, 0))

        # --- Container for the plot or the "disabled" message ---
        self.plot_container_frame = ttk.Frame(parent, style="Panel.TFrame")
        self.plot_container_frame.grid(row=1, column=0, sticky="nsew")
        self.plot_container_frame.rowconfigure(0, weight=1)
        self.plot_container_frame.columnconfigure(0, weight=1)

        # Set the initial state
        self._update_3d_plot_button_style()
        self._update_3d_plot_visibility()

    def launch_visualizer(self):
        """Opens visualizer.html in the system default web browser."""
        script_dir = os.path.dirname(os.path.abspath(__file__))
        visualizer_path = os.path.join(script_dir, "visualizer.html")
        if os.path.exists(visualizer_path):
            webbrowser.open_new_tab(f"file:///{visualizer_path}")
        else:
            messagebox.showerror(
                "File Not Found",
                f"Could not find visualizer.html at:\n{visualizer_path}"
            )

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

    def _is_3d_tab_active(self):
        """Returns True only when the '3D View' notebook tab is currently selected."""
        try:
            return self.notebook.select() == str(self.display_tab)
        except Exception:
            return False

    def _on_tab_changed(self, event=None):
        """Triggered when the user switches tabs. Refreshes the 3D plot if needed."""
        if self._is_3d_tab_active():
            # Defer slightly so the tab's canvas has finished drawing before matplotlib renders.
            self.root.after(50, self._draw_3d_toolpath)

    def _draw_3d_toolpath(self):
        """Draws the full G-code toolpath on the 3D plot."""
        if not self.is_3d_plot_enabled.get() or not self.matplotlib_imported or not self.ax_3d:
            return
        if not self._is_3d_tab_active():
            return  # Skip render when 3D tab is hidden
            
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

        # Plot all segments in cyan, regardless of completion status
        if len(x_coords) > 1:
            self.ax_3d.plot(x_coords, y_coords, z_coords, color=self.COLOR_ACCENT_CYAN, linewidth=1.2, alpha=self.toolpath_3d_opacity_var.get())

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
        if not self._is_3d_tab_active():
            return  # Skip render when 3D tab is hidden

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

    def clear_file(self):
        """Clears the currently loaded G-code program."""
        self.gcode_filepath = None
        self.processed_gcode = []
        self.toolpath_by_layer = {}
        self.move_to_layer_map = []
        self.ordered_z_values = []
        self.completed_move_count = 0
        self._plot_cache_valid = False
        self.rotation_crash_test_complete = False
        
        self.file_path_var.set("No file selected.")
        self.header_file_var.set("NONE")
        self.progress_var.set(0.0)
        self.progress_label_var.set("0 / 0 (0%)")
        self.start_button.config(state=tk.DISABLED)
        
        # Invalidate any cached plots
        self._invalidate_all_plot_caches()
        
        # Clear the matplotlib 3D plotting canvas
        if hasattr(self, 'ax_3d') and self.ax_3d is not None:
            self.ax_3d.clear()
            self.ax_3d.set_axis_off()
            if hasattr(self, 'canvas_3d') and self.canvas_3d is not None:
                self.canvas_3d.draw()
                
        # The XY/Z/E canvases are refreshed automatically by _update_all_displays calling their draw methods, 
        # which will wipe lines if `processed_gcode` is empty.
            
        self.log_message("Cleared loaded G-code program.")
        self._update_all_displays()

    def _apply_e_conversion(self, command):
        """
        Scales 'E' values in the G-code command by the configured mm/degree ratio.
        Used to translate logical degrees (UI/File) to physical mm (Printer).
        """
        try:
            ratio = float(self.mm_per_degree_var.get())
            if abs(ratio - 1.0) < 0.000001: return command
            
            # Regex to find E<number>
            # Matches E followed by optional whitespace, optional sign, digits, optional decimal
            def replace_e(match):
                val = float(match.group(1))
                new_val = val * ratio
                return f"E{new_val:.4f}"
            
            import re
            # Only match E if it's a command parameter (not a comment, though stripped lines help)
            return re.sub(r"([Ee])\s*([-+]?\d*\.?\d+)", replace_e, command)
        except Exception:
            return command

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
            import os
            from datetime import datetime

            # Store path instead of reading the whole file into memory
            self.gcode_filepath = filepath
            self.file_path_var.set(filepath)
            
            # Update the header to show just the filename.
            filename = os.path.basename(filepath)
            self.header_file_var.set(filename.upper())
            
            self.select_gcode_button.config(style='GreenRing.TButton')

            self.progress_var.set(0.0)
            
            self.log_message(f"Loading from {filepath}...")
            
            # Set default log filename based on G-code filename
            base_name = os.path.splitext(filename)[0]
            self.log_filepath_var.set(f"{base_name}_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")
            
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
        self.rotation_crash_test_complete = False

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
                        
                        abs_x = current_pos.get('x') if rel_x is None else float(rel_x) + center_x
                        abs_y = current_pos.get('y') if rel_y is None else float(rel_y) + center_y
                        abs_z = current_pos.get('z') if rel_z is None else float(rel_z) + center_z

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
        
        # Check for rotation components to determine if collision test is mandatory
        has_rotation = any('E' in line.upper() for line in self.processed_gcode if not line.strip().startswith(';'))
        if not has_rotation:
            self.rotation_crash_test_complete = True
            self.log_message("No rotation detected in G-code. Collision test marked as complete.")
        
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
        self.connect_button.config(text="Connecting...", state=tk.DISABLED)
        self.cancel_connect_button.grid() # Reveal pre-gridded button in column 6
        self.cancel_connect_button.config(state=tk.NORMAL)
        self.port_combobox.config(state=tk.DISABLED)
        self.baud_entry.config(state=tk.DISABLED)
        
        self.connection_status_var.set("Connecting...")
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
        
        # Hide the cancel button, restore connect button text, and update the status display.
        self.cancel_connect_button.grid_remove()
        self.connect_button.config(text="Connect", state=tk.NORMAL)
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
                self.queue_message("Sending configuration commands (Cold Extrusion, Idle Timeout, Endstops)...")
                try:
                    # M302 P1 S0: Allow cold extrusion
                    # M84 S900: Set motor idle timeout to 15 minutes (900 seconds)
                    # M120: Enable endstop detection during all moves (not just homing)
                    serial_conn.write(b'M302 P1 S0\nM84 S900\nM120\n')
                    serial_conn.flush()
                    
                    # We expect three 'ok' responses (M302, M84, M120), wait for them
                    config_oks = 0
                    config_start_time = time.time()
                    config_response_buffer = ""
                    while time.time() - config_start_time < 5.0:
                        if serial_conn.in_waiting > 0:
                            config_response_buffer += serial_conn.read(serial_conn.in_waiting).decode('utf-8', errors='ignore')
                            config_oks = config_response_buffer.lower().count('ok')
                            if config_oks >= 3:
                                break
                        time.sleep(0.05)
                        
                    if config_oks < 3:
                        self.queue_message("Warning: Not all setup 'ok's received. Configuration might be incomplete.", "WARN")
                    else:
                        self.queue_message("Initial printer configuration confirmed.", "SUCCESS")
                except Exception as e:
                    self.queue_message(f"Error during initial configuration: {e}", "ERROR")

                # --- Poll Current Position via M114 ---
                initial_position = None
                self.queue_message("Polling current position (M114)...")
                try:
                    serial_conn.reset_input_buffer()
                    serial_conn.write(b'M114\n')
                    serial_conn.flush()
                    m114_buffer = ""
                    m114_start = time.time()
                    while time.time() - m114_start < 5.0:
                        if serial_conn.in_waiting > 0:
                            m114_buffer += serial_conn.read(serial_conn.in_waiting).decode('utf-8', errors='ignore') # type: ignore
                            if 'ok' in m114_buffer.lower():
                                break
                        time.sleep(0.05)
                    
                    # Parse M114 response: "X:110.00 Y:110.00 Z:50.00 E:0.00 Count X:..."
                    import re as _re
                    x_match = _re.search(r'X:\s*([-+]?\d*\.?\d+)', m114_buffer)
                    y_match = _re.search(r'Y:\s*([-+]?\d*\.?\d+)', m114_buffer)
                    z_match = _re.search(r'Z:\s*([-+]?\d*\.?\d+)', m114_buffer)
                    e_match = _re.search(r'E:\s*([-+]?\d*\.?\d+)', m114_buffer)
                    
                    if x_match and y_match and z_match:
                        initial_position = {
                            'x': float(x_match.group(1)),
                            'y': float(y_match.group(1)),
                            'z': float(z_match.group(1)),
                            'e': float(e_match.group(1)) if e_match else 0.0
                        }
                        self.message_queue.put(("POSITION_UPDATE", initial_position))
                        self.queue_message(
                            f"Position: X={initial_position['x']:.2f} "
                            f"Y={initial_position['y']:.2f} "
                            f"Z={initial_position['z']:.2f} "
                            f"E={initial_position['e']:.2f}", "SUCCESS"
                        )
                    else:
                        self.queue_message("M114 response could not be parsed. Assuming origin.", "WARN")
                        self.queue_message(f"Raw M114 response: {m114_buffer.strip()}", "INFO")
                except Exception as e:
                    self.queue_message(f"Error polling M114: {e}. Assuming origin.", "WARN")

                self.message_queue.put(("CONNECTED", (serial_conn, found_port, baudrate, initial_position)))
            else:
                self.hardware_fault = False # Clear fault on attempt
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
        self.log_message(f"[DIAG] disconnect_printer called. silent={silent}, is_sending={self.is_sending}, is_manual={self.is_manual_command_running}, serial_conn={self.serial_connection is not None}", "WARN")

        # If a connection is in progress, cancel it instead.
        if self.connect_button['state'] == tk.DISABLED and not self.serial_connection and hasattr(self, 'cancel_connect_button') and self.cancel_connect_button.winfo_ismapped():
             self.log_message("Disconnect during connect - Cancelling.", "WARN")
             self._cancel_connection_attempt()
             return
             
        # For user-initiated disconnects, refuse while a job is actively running.
        # Forced/silent disconnects (e.g. CONNECTION_LOST) must bypass this guard
        # so the GUI is always reset regardless of runtime state.
        if not silent and (self.is_sending or self.is_manual_command_running):
            self.log_message("Cannot disconnect while busy.", "WARN")
            messagebox.showwarning("Busy", "Please stop the current operation before disconnecting.")
            return

        # For forced disconnects, stop any active operations before resetting.
        if silent:
            self.stop_event.set()
            self.is_sending = False
            self.is_manual_command_running = False
            self.is_paused = False
            
        if self.serial_connection and self.serial_connection.is_open:
            try:
                self.serial_connection.close()
            except Exception as e:
                    self.log_message(f"Disconnect error: {e}", "ERROR")
            
        self.serial_connection = None
        self.log_message("[DIAG] serial_connection cleared. Draining queue and resetting GUI.", "WARN")

        # Drain any queued CONNECTED messages that the background thread may have
        # posted just before we closed the port. Without this, check_message_queue()
        # could process a stale CONNECTED message 100ms later and re-assert the green state.
        try:
            import queue as _queue
            preserved = []
            while True:
                try:
                    item = self.message_queue.get_nowait()
                    if item[0] != 'CONNECTED':
                        preserved.append(item)
                    else:
                        self.log_message("[DIAG] Drained a stale CONNECTED message.", "WARN")
                except _queue.Empty:
                    break
            for item in preserved:
                self.message_queue.put(item)
        except Exception:
            pass

        # --- Reset GUI to "Disconnected" state ---
        self.connection_status_var.set("Disconnected")
        self.status_indicator.set_status("off")
        self.header_status_indicator.set_status("off")
        self.footer_status_var.set("COM: -- @ --")

        # Use .configure() for ttk widgets — .config() is unreliable for style/text on ttk
        self.connect_button.configure(text="Connect", style='YellowRing.TButton')
        self.connect_button.state(['!disabled'])
        if hasattr(self, '_conn_frame'):
            self._conn_frame.configure(style='Yellow.TLabelframe')
        self.port_combobox.configure(state="readonly")
        self.baud_entry.configure(state=tk.NORMAL)
        
        self.start_button.config(state=tk.DISABLED)
        self.pause_resume_button.config(text="Pause", state=tk.DISABLED)
        self._set_manual_controls_state(tk.DISABLED)
        self._set_goto_controls_state(tk.DISABLED)
        self._set_terminal_controls_state(tk.DISABLED)

        if hasattr(self, 'cancel_connect_button'):
            self.cancel_connect_button.grid_remove()
            self.cancel_connect_button.config(state=tk.DISABLED)
        
        self.progress_var.set(0.0)
        self.progress_label_var.set("Progress: Idle")

        # Reset the known position, setup flags, and update the display.

        # Drain any queued CONNECTED messages that the background thread may have
        # posted just before we closed the port. Without this, check_message_queue()
        # could process a stale CONNECTED message 100ms later and re-assert the green state.
        try:
            import queue as _queue
            preserved = []
            while True:
                try:
                    item = self.message_queue.get_nowait()
                    if item[0] != 'CONNECTED':
                        preserved.append(item)
                    else:
                        self.log_message("[DIAG] Drained a stale CONNECTED message.", "WARN")
                except _queue.Empty:
                    break
            for item in preserved:
                self.message_queue.put(item)
        except Exception:
            pass

        # --- Reset GUI to "Disconnected" state ---
        self.connection_status_var.set("Disconnected")
        self.status_indicator.set_status("off")
        self.header_status_indicator.set_status("off")
        self.footer_status_var.set("COM: -- @ --")

        # Use .configure() for ttk widgets — .config() is unreliable for style/text on ttk
        self.connect_button.configure(text="Connect", style='YellowRing.TButton')
        self.connect_button.state(['!disabled'])
        if hasattr(self, '_conn_frame'):
            self._conn_frame.configure(style='Yellow.TLabelframe')
        self.port_combobox.configure(state="readonly")
        self.baud_entry.configure(state=tk.NORMAL)
        
        self.start_button.config(state=tk.DISABLED)
        self.pause_resume_button.config(text="Pause", state=tk.DISABLED)
        self._set_manual_controls_state(tk.DISABLED)
        self._set_goto_controls_state(tk.DISABLED)
        self._set_terminal_controls_state(tk.DISABLED)

        if hasattr(self, 'cancel_connect_button'):
            self.cancel_connect_button.grid_remove()
            self.cancel_connect_button.config(state=tk.DISABLED)
        
        self.progress_var.set(0.0)
        self.progress_label_var.set("Progress: Idle")

        # Reset the known position, setup flags, and update the display.
        self.last_cmd_abs_x, self.last_cmd_abs_y, self.last_cmd_abs_z = None, None, None
        self.center_marked = False
        self.rotation_crash_test_complete = False
        self._update_all_displays()
        self._update_section_borders()

        # Schedule a delayed re-enforcement of the disconnected state as a final
        # failsafe against any stale message that slips through the drain above.
        self.root.after(350, self._enforce_disconnect_state)

    def _enforce_disconnect_state(self):
        """
        Failsafe called 350ms after disconnect_printer() to ensure the GUI
        is showing the correct disconnected state. Guards against any stale
        CONNECTED message that may have slipped through the queue drain.
        """
        if self.serial_connection is not None:
            return  # We've reconnected since — don't override the connected state.
        self.log_message("[DIAG] _enforce_disconnect_state firing.", "WARN")
        self._update_section_borders()
        if hasattr(self, '_conn_frame'):
            self._conn_frame.configure(style='Yellow.TLabelframe')
        self.connect_button.configure(text="Connect", style='YellowRing.TButton')
        self.root.update_idletasks()


    # --- G-Code Sending & Control ---

    def start_sending(self):
        """
        Starts sending the processed G-code file to the printer.

        It performs several pre-flight checks, updates the GUI to a
        'sending' state, and starts the background thread that handles the line-by-line
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

        # --- Crash Avoidance Check ---
        if not self.rotation_crash_test_complete:
            ans = messagebox.askyesno("Safety Check", 
                                      "The Rotation Collision Avoidance Test has not been completed for this profile.\n\n"
                                      "Are you sure you want to proceed with the scan?")
            if not ans:
                return
                
        # --- CSV Logging Check ---
        if not self.log_measurements_enabled.get() or not self.log_filepath_var.get().strip():
            ans = messagebox.askyesno("Logging Not Active", 
                                      "You are about to start a run, but 'Log to CSV' is not enabled or no file is selected.\n\n"
                                      "Measurement data from this run will NOT be saved.\n\n"
                                      "Are you sure you want to proceed?")
            if not ans:
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
            self.header_status_indicator.set_status("busy")

    def emergency_stop(self):
        """
        Triggers an immediate, hard stop of the printer (M112).

        This sets the stop event to halt any running threads, sends M112 to the
        printer (which usually requires a physical reset), and disconnects the
        application from the serial port.

        Thread-safety: stop_event is set FIRST so background threads stop queuing
        new writes. We then acquire serial_lock to guarantee any in-progress write
        finishes before we flush the output buffer and write M410/M112. This means
        the E-stop bytes reach the printer immediately, regardless of what was
        previously queued.
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
        
        # Invalidate collision test status if we had to stop motion
        self.rotation_crash_test_complete = False
        
        # Step 1: Signal all background threads to stop immediately.
        # They check stop_event between writes, so they will not queue any further commands.
        self.pause_event.set()   # Unblock any paused thread so it can see stop_event.
        self.stop_event.set()
        
        # Update status indicators to "error" (red).
        self.status_indicator.set_status("error")
        self.header_status_indicator.set_status("error")

        if self.serial_connection:
            try:
                self.log_message("Sending M410 + M112 (Emergency Stop)...")
                # Step 2: Acquire the serial lock.
                # Any background thread that is mid-write will finish its current
                # write() call before we proceed. Once we hold the lock, no other
                # thread can write to the port until we release it.
                with self.serial_lock:
                    # Clear the OS-level output buffer so any buffered G-code bytes
                    # that haven't left the PC yet are discarded.
                    self.serial_connection.reset_output_buffer()
                    # M410: Quickstop — halts motion immediately without requiring reset.
                    self.serial_connection.write(b'M410\n')
                    # M112: Full emergency stop — kills all motion/heaters, requires reset.
                    self.serial_connection.write(b'M112\n')
                    # Force the bytes out of the PC's serial driver immediately.
                    self.serial_connection.flush()
                time.sleep(0.2)
                self.serial_connection.reset_input_buffer()
                self.log_message("M410 + M112 sent.")
            except Exception as e:
                self.log_message(f"Error sending emergency stop: {e}", "ERROR")
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
        
        # Invalidate collision test status if we had to stop motion
        self.rotation_crash_test_complete = False

        # Un-pause and then stop all running threads.
        self.pause_event.set()
        self.stop_event.set()
        
        # Update status indicators to "busy" (amber).
        self.status_indicator.set_status("busy")
        self.header_status_indicator.set_status("busy")

        if self.serial_connection:
            try:
                self.log_message("Sending M410 (Quick Stop)...")
                # Acquire the serial lock so M410 is not interleaved with a
                # background write. Unlike emergency_stop, we do NOT call
                # reset_output_buffer() here -- quick stop intentionally lets
                # the printer drain its existing move buffer before halting.
                with self.serial_lock:
                    self.serial_connection.write(b'M410\n')
                    self.serial_connection.flush()
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
    def _parse_gcode_coords(self, gcode_line: str) -> Dict[str, float]:
        """Extracts X, Y, Z, and E coordinates from a G0/G1/G92 command line."""
        coords: Dict[str, float] = {}
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


    def _parse_m119_response(self, response_text):
        """
        Parses the output of M119 (Endstop States).
        Returns a dict: {'x_min': 'open'/'triggered', ...}
        """
        states = {}
        lines = response_text.splitlines()
        for line in lines:
            if ':' in line:
                key, val = line.split(':', 1)
                states[key.strip().lower()] = val.strip().lower()
        return states

    def _homing_verification_routine(self):
        """
        Executes a multi-phase verification to detect X/Y drift before re-homing.
        Must be called from a background thread (sender or manual).
        Raises InterruptedError if verification fails or stop is requested.
        """
        self.queue_message("Starting Homing Verification...", "INFO")
        
        def send_and_wait_m119():
            with self.serial_lock:
                self.serial_connection.write(b'M119\n')
            buffer = ""
            start = time.time()
            while time.time() - start < 5.0:
                if self.stop_event.is_set(): raise InterruptedError("Stop")
                if self.serial_connection.in_waiting > 0:
                    buffer += self.serial_connection.read(self.serial_connection.in_waiting).decode('utf-8', errors='ignore') # type: ignore
                    if 'ok' in buffer.lower():
                        return self._parse_m119_response(buffer)
                time.sleep(0.05)
            raise TimeoutError("M119 timeout")

        def send_move_and_wait(cmd):
            with self.serial_lock:
                self.serial_connection.write(cmd.encode('utf-8') + b'\n')
            buffer = ""
            start = time.time()
            # Homing can take longer, but G1 moves should be within 30s
            timeout = 90.0 if "G28" in cmd.upper() else 30.0
            while time.time() - start < timeout:
                if self.stop_event.is_set(): raise InterruptedError("Stop")
                if self.serial_connection.in_waiting > 0:
                    buffer += self.serial_connection.read(self.serial_connection.in_waiting).decode('utf-8', errors='ignore') # type: ignore
                    if 'ok' in buffer.lower():
                        return
                time.sleep(0.05)
            raise TimeoutError(f"Command timeout: {cmd}")

        try:
            # 0. Wait for all previous moves to finish
            send_move_and_wait("M400")
            
            # 1. Phase A: Move to Safe Zone (0.5, 0.5)
            # This detects "Negative Drift" (printer thinks it's further from switch than it is)
            self.queue_message("Verification Phase A: Moving to Safe Zone (0.5, 0.5)...")
            send_move_and_wait("G1 X0.5 Y0.5 F3000")
            send_move_and_wait("M400")
            states = send_and_wait_m119()
            
            if states.get('x_min') == 'triggered' or states.get('y_min') == 'triggered':
                axis = 'X' if states.get('x_min') == 'triggered' else 'Y'
                diag = f"CRITICAL ERROR: {axis}-axis triggered early at Safe Zone (0.5mm). Negative drift/skipped steps detected."
                self._handle_homing_failure(diag)
                return

            # 2. Phase B: Move to Home Target (0, 0)
            # This detects "Positive Drift" (printer thinks it's at switch but it's not)
            self.queue_message("Verification Phase B: Moving to Home Target (0, 0)...")
            send_move_and_wait("G1 X0 Y0 F500") 
            send_move_and_wait("M400")
            states = send_and_wait_m119()
            
            if states.get('x_min') == 'open' or states.get('y_min') == 'open':
                axis = 'X' if states.get('x_min') == 'open' else 'Y'
                diag = f"CRITICAL ERROR: {axis}-axis failed to trigger at Home Target (0.0mm). Positive drift/skipped steps detected."
                self._handle_homing_failure(diag)
                return

            # 3. Phase C: Actual Re-Home to sync coordinate system
            self.queue_message("Verification Passed. Performing Sync Home (G28 X Y)...", "SUCCESS")
            send_move_and_wait("G28 X Y")
            self.queue_message("Sync Home Complete. Resuming scan.")

        except Exception as e:
            if isinstance(e, InterruptedError): raise
            diag = f"Verification Aborted: {str(e)}"
            self._handle_homing_failure(diag)

    def _handle_homing_failure(self, diagnosis):
        """Helper to halt motion and notify GUI on verification failure."""
        self.queue_message(diagnosis, "CRITICAL")
        if self.serial_connection:
            try: 
                self.serial_connection.write(b'M410\n') # Quick Stop
                self.serial_connection.flush()
            except: pass
        
        self.stop_event.set()
        self.message_queue.put(("HOMING_FAILURE", diagnosis))
        raise InterruptedError(diagnosis)


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
                line_to_send = self._apply_e_conversion(line)
                with self.serial_lock:
                    self.serial_connection.write(line_to_send.encode('utf-8') + b'\n')
                
                # Log modification if it happened
                if line != line_to_send:
                     self.queue_message(f"Sent: {line_to_send} (Orig: {line})")
                else:
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

    def _handle_key_press(self, event):
        """Processes global keyboard inputs for jogging and step size adjustments."""
        # Prevent jogging when typing in input fields
        if isinstance(event.widget, (tk.Entry, ttk.Entry, tk.Text, ttk.Combobox, ttk.Spinbox)):
            return

        # Handle priority interrupts (Stop / Pause) first, as these should work
        # even if a command is currently running.
        keysym = event.keysym.lower()
        
        if keysym == 'space':
            self.emergency_stop()
            return
        elif keysym == 'p':
            self._toggle_pause()
            return

        # Check if connected and controls are actually enabled for motion commands
        if not self.serial_connection or self.is_manual_command_running or self.is_sending:
            return

        # Handle arrow keys via keysym
        if keysym == 'left':
            self._jog('E', 1)
            return
        elif keysym == 'right':
            self._jog('E', -1)
            return
        elif keysym in ('up', 'down'):
            self._cycle_step_size('ROT', 1 if keysym == 'up' else -1)
            return
        
        # Handle motion and setup shortcuts
        char = event.char.lower()
        if not char:
            return

        if char == 'h':
            self._home_printer()
        elif char == 'c':
            self._go_to_center()
        elif char == 'w':
            self._jog('Y', -1)
        elif char == 's':
            self._jog('Y', 1)
        elif char == 'a':
            self._jog('X', -1)
        elif char == 'd':
            self._jog('X', 1)
        elif char == 'q':
            self._jog('Z', -1)
        elif char == 'e':
            self._jog('Z', 1)
        elif char == 'r':
            self._cycle_step_size('XYZ', 1)
        elif char == 'f':
            self._cycle_step_size('XYZ', -1)

    def _cycle_step_size(self, axis_type, direction):
        """Cycles through predefined step sizes for manual jogging."""
        xyz_steps = [0.1, 1.0, 5.0, 10.0, 50.0, 100.0]
        rot_steps = [1.0, 5.0, 10.0, 45.0, 90.0]
        
        if axis_type == 'XYZ':
            try:
                current = float(self.jog_step_var.get())
            except ValueError:
                current = 10.0
            steps = xyz_steps
            var = self.jog_step_var
        else:
            try:
                current = float(self.rotation_step_var.get())
            except ValueError:
                current = 5.0
            steps = rot_steps
            var = self.rotation_step_var
            
        # Find closest matching step size
        closest_idx = min(range(len(steps)), key=lambda i: abs(steps[i] - current))
        new_idx = max(0, min(len(steps) - 1, closest_idx + direction))
        var.set(str(steps[new_idx]))

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
        self.stop_event.clear()
        self.pause_event.set()
        self._set_manual_controls_state(tk.DISABLED)
        self._set_goto_controls_state(tk.DISABLED)
        self.start_button.config(state=tk.DISABLED)
        
        self.log_message("Starting Auto-Homing & Calibration...", "INFO")
        threading.Thread(target=self._homing_sequence_worker, daemon=True).start()

    def _homing_sequence_worker(self):
        """
        Background thread for the homing sequence.
        Sends G28 (Home All Axes) and then polls M114 to sync position.
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
            
            # --- Step 2: Poll position via M114 ---
            self.queue_message("Reading post-home position (M114)...")
            if self.serial_connection:
                self.serial_connection.reset_input_buffer()
                self.serial_connection.write(b'M114\n')
                self.serial_connection.flush()
                
                import re as _re
                m114_buffer = ""
                m114_start = time.time()
                while time.time() - m114_start < 5.0:
                    if self.serial_connection.in_waiting > 0:
                        m114_buffer += self.serial_connection.read(self.serial_connection.in_waiting).decode('utf-8', errors='ignore') # type: ignore
                        if 'ok' in m114_buffer.lower():
                            break
                    time.sleep(0.05)
                
                x_match = _re.search(r'X:\s*([-+]?\d*\.?\d+)', m114_buffer)
                y_match = _re.search(r'Y:\s*([-+]?\d*\.?\d+)', m114_buffer)
                z_match = _re.search(r'Z:\s*([-+]?\d*\.?\d+)', m114_buffer)
                e_match = _re.search(r'E:\s*([-+]?\d*\.?\d+)', m114_buffer)
                
                if x_match and y_match and z_match:
                    pos = {
                        'x': float(x_match.group(1)),
                        'y': float(y_match.group(1)),
                        'z': float(z_match.group(1)),
                        'e': float(e_match.group(1)) if e_match else 0.0
                    }
                    self.message_queue.put(("POSITION_UPDATE", pos))
                    self.queue_message(
                        f"Homed to: X={pos['x']:.2f} Y={pos['y']:.2f} Z={pos['z']:.2f}", "SUCCESS"
                    )
                else:
                    # Fallback: assume origin
                    self.message_queue.put(("POSITION_UPDATE", {'x': 0.0, 'y': 0.0, 'z': 0.0, 'e': 0.0}))
                    self.queue_message("M114 parse failed after G28. Assuming origin.", "WARN")
            else:
                self.message_queue.put(("POSITION_UPDATE", {'x': 0.0, 'y': 0.0, 'z': 0.0, 'e': 0.0}))
            
            self.queue_message("Homing complete.", "SUCCESS")
            self.message_queue.put(("SET_STATUS", "on"))
            
            # Re-enable controls
            self.is_manual_command_running = False
            self.message_queue.put(("MANUAL_FINISHED", True))

        except Exception as exc:
            self.queue_message(f"Homing Error: {exc}", "ERROR")
            self.message_queue.put(("SET_STATUS", "error"))
            
            # Attempt to restore G90 in case of error
            if self.serial_connection:
                try: self.serial_connection.write(b'G90\n')
                except: pass
            
            # Re-enable controls
            self.is_manual_command_running = False
            self.message_queue.put(("MANUAL_FINISHED", False))
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
        the input fields and canvas markers. It then automatically triggers
        the 'Go To' movement command.
        """
        try:
            # 1. Read the center coordinates from their StringVars.
            center_x = float(self.center_x_var.get())
            center_y = float(self.center_y_var.get())
            center_z = float(self.center_z_var.get())
            center_e = float(self.center_e_var.get())

            # 2. Set the internal 'target' position model.
            self.target_abs_x = center_x
            self.target_abs_y = center_y
            self.target_abs_z = center_z
            self.target_abs_e = center_e

            # 3. Update all displays (DRO labels and canvas markers).
            self._update_all_displays()
            self.log_message(f"Target set to center: X={center_x:.2f}, Y={center_y:.2f}, Z={center_z:.2f}, E={center_e:.2f}", "INFO")

            # 4. Also update the 'Go To' entry boxes to reflect this change.
            mode = self.coord_mode.get()
            display_x = f"{center_x:.2f}" if mode == "absolute" else "0.00"
            display_y = f"{center_y:.2f}" if mode == "absolute" else "0.00"
            display_z = f"{center_z:.2f}" if mode == "absolute" else "0.00"
            display_e = f"{center_e:.2f}" if mode == "absolute" else "0.00"

            self.goto_x_entry.delete(0, tk.END); self.goto_x_entry.insert(0, display_x)
            self.goto_y_entry.delete(0, tk.END); self.goto_y_entry.insert(0, display_y)
            self.goto_z_entry.delete(0, tk.END); self.goto_z_entry.insert(0, display_z)
            self.goto_e_entry.delete(0, tk.END); self.goto_e_entry.insert(0, display_e)

            # 5. Trigger the move
            self._go_to_position()

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
        world_y = bounds['y_min'] + (click_y_rel / canvas_h) * y_range if y_range != 0 else bounds['y_min'] # y=0 at top
        
        # Update the internal model for the target position.
        self.target_abs_x = max(bounds['x_min'], min(bounds['x_max'], world_x))
        self.target_abs_y = max(bounds['y_min'], min(bounds['y_max'], world_y))
        
        # Update all GUI displays to reflect the new target.
        self._update_all_displays()


    def _draw_xy_canvas_guides(self, event=None):
        """Draws the toolpath for the CURRENT Z-LEVEL ONLY, plus grid, origin, and markers."""
        # pylint: disable=unused-argument
        # Only delete tagged items, not the persistent background
        self.xy_canvas.delete("toolpath", "guides", "marker_blue", "marker_red", "crosshair", "margin")
        bounds = self.PRINTER_BOUNDS
        w = self.xy_canvas.winfo_width(); h = self.xy_canvas.winfo_height()
        if w <= 1 or h <= 1: return
        
        x_range = bounds['x_max'] - bounds['x_min']; y_range = bounds['y_max'] - bounds['y_min']
        if x_range == 0 or y_range == 0: return 

        def world_to_canvas(wx, wy):
            cx = w * (wx - bounds['x_min']) / x_range
            cy = h * (wy - bounds['y_min']) / y_range # y=0 at top
            return cx, cy
        # --- Draw Hash Grid & Axis Labels ---
        grid_minor = "#1a2c3a"
        grid_major = "#24404f"
        label_color = "#3a5570"
        font_small = ("Helvetica", 7)

        for i in range(int(bounds['x_min']), int(bounds['x_max']) + 1, 10):
            cx, _ = world_to_canvas(i, 0)
            is_major = (i % 50 == 0)
            color = grid_major if is_major else grid_minor
            dash = () if is_major else (1, 3)
            self.xy_canvas.create_line(cx, 0, cx, h, fill=color, tags="guides", dash=dash)
            if i != 0 and i % 20 == 0:
                self.xy_canvas.create_text(cx, h - 2, text=f"{i}", fill=label_color, font=font_small, anchor="s", tags="guides")

        for i in range(int(bounds['y_min']), int(bounds['y_max']) + 1, 10):
            _, cy = world_to_canvas(0, i)
            is_major = (i % 50 == 0)
            color = grid_major if is_major else grid_minor
            dash = () if is_major else (1, 3)
            self.xy_canvas.create_line(0, cy, w, cy, fill=color, tags="guides", dash=dash)
            if i != 0 and i % 20 == 0:
                self.xy_canvas.create_text(3, cy, text=f"{i}", fill=label_color, font=font_small, anchor="w", tags="guides")

        # --- Origin Axes (brighter solid lines) ---
        if bounds['x_min'] <= 0 <= bounds['x_max'] and bounds['y_min'] <= 0 <= bounds['y_max']:
            canvas_x0, canvas_y0 = world_to_canvas(0, 0)
            self.xy_canvas.create_line(canvas_x0, 0, canvas_x0, h, fill=self.COLOR_BORDER, tags="guides")
            self.xy_canvas.create_line(0, canvas_y0, w, canvas_y0, fill=self.COLOR_BORDER, tags="guides")

        # --- Margin Indicators from Current Position to Each Wall ---
        try:
            pos_x = self.last_cmd_abs_x
            pos_y = self.last_cmd_abs_y
            if pos_x is not None and pos_y is not None:
                marker_cx, marker_cy = world_to_canvas(pos_x, pos_y)
                edge_left_cx, _ = world_to_canvas(bounds['x_min'], 0)
                edge_right_cx, _ = world_to_canvas(bounds['x_max'], 0)
                _, edge_top_cy = world_to_canvas(0, bounds['y_max'])
                _, edge_bot_cy = world_to_canvas(0, bounds['y_min'])

                margin_color = "#4a6070"
                margin_font = ("Helvetica", 7)

                margin_left  = pos_x - bounds['x_min']
                margin_right = bounds['x_max'] - pos_x
                margin_front = pos_y - bounds['y_min']   # in printer space Y-min is "front"
                margin_back  = bounds['y_max'] - pos_y

                # Left margin line & label
                self.xy_canvas.create_line(edge_left_cx, marker_cy, marker_cx, marker_cy,
                                           fill=margin_color, dash=(2, 3), tags="margin")
                self.xy_canvas.create_text(edge_left_cx + 2, marker_cy - 2,
                                           text=f"{margin_left:.0f}", fill=margin_color,
                                           font=margin_font, anchor="w", tags="margin")
                # Right margin line & label
                self.xy_canvas.create_line(marker_cx, marker_cy, edge_right_cx, marker_cy,
                                           fill=margin_color, dash=(2, 3), tags="margin")
                self.xy_canvas.create_text(edge_right_cx - 2, marker_cy - 2,
                                           text=f"{margin_right:.0f}", fill=margin_color,
                                           font=margin_font, anchor="e", tags="margin")
                # Bottom (front) margin line & label
                self.xy_canvas.create_line(marker_cx, marker_cy, marker_cx, edge_bot_cy,
                                           fill=margin_color, dash=(2, 3), tags="margin")
                self.xy_canvas.create_text(marker_cx + 2, edge_bot_cy - 2,
                                           text=f"{margin_front:.0f}", fill=margin_color,
                                           font=margin_font, anchor="sw", tags="margin")
                # Top (back) margin line & label
                self.xy_canvas.create_line(marker_cx, edge_top_cy, marker_cx, marker_cy,
                                           fill=margin_color, dash=(2, 3), tags="margin")
                self.xy_canvas.create_text(marker_cx + 2, edge_top_cy + 2,
                                           text=f"{margin_back:.0f}", fill=margin_color,
                                           font=margin_font, anchor="nw", tags="margin")
        except Exception:
            pass


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
            
            # Now, draw all segments for the current layer in cyan
            for idx, (start_point, end_point) in enumerate(segments_for_current_layer):
                if start_point is None or end_point is None: continue
                
                start_cx, start_cy = world_to_canvas(start_point[0], start_point[1])
                end_cx, end_cy = world_to_canvas(end_point[0], end_point[1])
                
                self.xy_canvas.create_line(start_cx, start_cy, end_cx, end_cy, 
                                           fill=self.COLOR_ACCENT_CYAN, width=1, tags="toolpath")

        m_size = 5

        # --- Draw Blue Marker (Go To Target) ---
        if not getattr(self, 'is_collision_test_running', False):
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
            # Deduplicate across the entire toolpath sequence to only show each uniquely visited Z layer once overall, preventing repeat zigzags on multi-pass prints
            unique_z_values = sorted(list(set(self.ordered_z_values)))
                    
            num_points = len(unique_z_values)
            # Use a slightly less dense width to ensure the lines don't stack up visually
            if num_points > 1:
                # The graph is drawn horizontally across the narrow canvas
                x_step = (canvas_w - 1) / (num_points - 1)

                for i in range(1, num_points):
                    start_x = (i - 1) * x_step
                    end_x = i * x_step
                    
                    start_y = z_to_canvas_y(unique_z_values[i-1])
                    end_y = z_to_canvas_y(unique_z_values[i])
                    
                    self.z_canvas.create_line(start_x, start_y, end_x, end_y, fill=self.COLOR_ACCENT_CYAN, width=1, tags="z_toolpath")
         
         # --- Draw Scale Labels ---
         # Max Z Label at top
         self.z_canvas.create_text(canvas_w/2, 10, text=f"{bounds['z_max']:.0f}", fill=self.COLOR_TEXT_SECONDARY, font=("Inter", 8), tags="labels")
         # Min Z Label at bottom
         self.z_canvas.create_text(canvas_w/2, canvas_h - 10, text="0", fill=self.COLOR_TEXT_SECONDARY, font=("Inter", 8), tags="labels")

         # --- Draw Blue Marker (Go To Target) ---
         if not getattr(self, 'is_collision_test_running', False):
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
        """Sets the E (Rotation) target based on click angle. 0° at bottom, range [-90, +90]."""
        # pylint: disable=unused-argument
        if self.go_button['state'] == tk.DISABLED: return
        w = self.e_canvas.winfo_width()
        h = self.e_canvas.winfo_height()
        if w <= 1 or h <= 1: return
        
        cx, cy = w / 2, h / 2
        dx, dy = event.x - cx, event.y - cy
        
        import math
        angle_rad = math.atan2(dy, dx)
        angle_deg = math.degrees(angle_rad)  # East=0, CW positive
        # We want 0° at bottom (South). South is +90 in standard math.
        # Our coordinate: positive = counterclockwise from South (towards East/right side of gauge)
        # Negative = clockwise from South (towards West/left side of gauge)
        # mapped_angle: 0 when clicking due South, +90 at East, -90 at West
        
        # In screen coords, dy is positive downward, so South is positive dy.
        # East is positive dx. West is negative dx.
        # Therefore, South corresponds to atan2(+dy, 0) -> +90 deg.
        # To make South 0, East -90, West +90:
        # Our angle = atan2_angle - 90
        mapped_angle = angle_deg - 90        
        # Normalize to [-180, 180]
        if mapped_angle > 180: mapped_angle -= 360
        if mapped_angle < -180: mapped_angle += 360
        
        # Reject clicks in the forbidden zone (top half: < -90 or > +90)
        if mapped_angle > 90 or mapped_angle < -90:
            return
        
        # Snap to nearest 5 degrees for cleaner UI interaction
        mapped_angle = round(mapped_angle / 5) * 5
        
        # Clamp to bounds
        mapped_angle = max(self.PRINTER_BOUNDS['e_min'], min(self.PRINTER_BOUNDS['e_max'], mapped_angle))
        
        self.target_abs_e = mapped_angle
        self._update_all_displays()

    def _draw_e_canvas_gauge(self, event=None):
        """Draws the circular gauge for the E-axis. 0° at bottom, range [-90, +90]."""
        # pylint: disable=unused-argument
        self.e_canvas.delete("all")
        w = self.e_canvas.winfo_width()
        h = self.e_canvas.winfo_height()
        if w <= 1 or h <= 1: return
        
        cx, cy = w / 2, h / 2
        radius = min(w, h) / 2 - 10
        
        import math
        
        # --- Draw allowed zone (bottom half) ---
        # tkinter arc: start angle is measured from East(0), CCW positive.
        # Bottom semicircle is from 180 (West) to 360 (East). So start=180, extent=180.
        self.e_canvas.create_arc(
            cx - radius, cy - radius, cx + radius, cy + radius,
            start=180, extent=180,
            fill='', outline=self.COLOR_BORDER, width=2, style=tk.ARC
        )
        # We no longer draw the full outer circle or the red forbidden zone.

        # Draw a horizontal line to cap the semi-circle
        self.e_canvas.create_line(cx - radius, cy, cx + radius, cy, fill=self.COLOR_BORDER, width=2)
        
        # --- Helper: convert our E-angle to screen radians ---
        # Our 0° = South (bottom). +angle = toward West (+90 is West), -angle = toward East (-90 is East).
        # Screen: radians measured from East=0, CW positive (because Y is inverted).
        # South is PI/2 in screen coords. East is 0. West is PI.
        # screen_rad = PI/2 + radians(our_angle)
        def e_to_screen_rad(e_deg: float) -> float:
            return math.pi / 2 + math.radians(e_deg)        
        # --- Ticks at key positions ---
        # To match the image, don't show + signs, and just label limits and zero, but ticks every 45
        tick_angles = [
            (-90, "-90"), (-45, ""), (0, "0"), (45, ""), (90, "90")
        ]
        for angle, label in tick_angles:
            rad = e_to_screen_rad(angle)
            is_major = (angle % 90 == 0)
            r_in = radius - 10 if is_major else radius - 6
            x1 = cx + radius * math.cos(rad)
            y1 = cy + radius * math.sin(rad)
            x2 = cx + r_in * math.cos(rad)
            y2 = cy + r_in * math.sin(rad)
            color = self.COLOR_ACCENT_CYAN if is_major else self.COLOR_TEXT_SECONDARY
            self.e_canvas.create_line(x1, y1, x2, y2, fill=color, width=2 if is_major else 1)
            
            # Labels for major ticks
            if is_major and label:
                lx = cx + (radius + 14) * math.cos(rad)
                # Adjust label heights slightly
                if angle == -90: # East
                    ly = cy - 2
                    lx += 4 # push right slightly
                elif angle == 90: # West
                    ly = cy - 2
                    lx -= 4 # pull left slightly
                else:
                    ly = cy + (radius + 14) * math.sin(rad)
                self.e_canvas.create_text(lx, ly, text=label, fill=self.COLOR_TEXT_SECONDARY, font=("Inter", 8))

        # --- Helper to draw a needle/marker ---
        def draw_needle(e_deg: float, color: str, length_pct: float, width: int, tag: str):
            rad_n = e_to_screen_rad(e_deg)
            nx = cx + (radius * length_pct) * math.cos(rad_n)
            ny = cy + (radius * length_pct) * math.sin(rad_n)
            self.e_canvas.create_line(cx, cy, nx, ny, fill=color, width=width, capstyle=tk.ROUND, tags=tag)
            r_tip = 3
            self.e_canvas.create_oval(nx - r_tip, ny - r_tip, nx + r_tip, ny + r_tip, fill=color, outline=color, tags=tag)

        # Target Needle (Cyan)
        if not getattr(self, 'is_collision_test_running', False):
            target_clamped = max(-90.0, min(90.0, self.target_abs_e))
            draw_needle(target_clamped, self.COLOR_ACCENT_CYAN, 0.9, 3, "marker_blue")
        
        # Current Position Marker (Red)
        if self.last_cmd_abs_e is not None:
            current_clamped = max(-90.0, min(90.0, self.last_cmd_abs_e))
            draw_needle(current_clamped, self.COLOR_ACCENT_RED, 0.7, 2, "marker_red")
            
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

                    # --- Homing Verification Logic ---
                    # Detect if this move command changes the Z layer.
                    # If it does, we rehome before executing this command.
                    if target_pos_for_line['z'] != last_pos['z'] and last_pos['z'] is not None:
                        self.queue_message(f"Layer change detected ({last_pos['z']:.2f} -> {target_pos_for_line['z']:.2f}). Running homing verification...")
                        try:
                            self._homing_verification_routine()
                        except InterruptedError:
                            # Re-raise to be caught by the outer loop's error handler
                            raise

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
            
            # --- EXPLICIT HOMING VERIFICATION INTERCEPT ---
            # If this is a specific G28 X Y command (not the very first initializing G28),
            # intercept it and run the drift verification routine before homing.
            if 'X' in gcode_line.upper() and 'Y' in gcode_line.upper() and 'Z' not in gcode_line.upper():
                if sent_line_count > 1: # Let first initial home pass through normally
                    self.queue_message("Explicit G28 X Y intercepted. Running Per-Layer Homing Verification...")
                    try:
                        self._homing_verification_routine()
                        # _homing_verification_routine physically issues G28 X Y itself.
                        # So we skip sending this line to avoid duplicate homing commands.
                        last_pos.update(target_pos_for_line)
                        continue 
                    except InterruptedError:
                        raise

            # --- Send Line and Wait for 'ok' ---
            try:
                line_to_send = self._apply_e_conversion(gcode_line)
                with self.serial_lock:
                    self.serial_connection.write(line_to_send.encode('utf-8') + b'\n')
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
                                with self.serial_lock:
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
                                        vals = self._take_measurement()
                                        
                                        self.message_queue.put(("MEASUREMENT_RESULT", vals))
                                        if self.log_measurements_enabled.get():
                                            self._log_measurement_to_file(vals, coords=last_pos)
                                    except Exception as e:
                                        self.queue_message(f"Auto-Measure Error: {e}", "ERROR")
                            else:
                                self.queue_message("Auto-Measure skipped (M400 timeout)", "WARN")

                else:
                    self.queue_message(f"Warning: No 'ok' received for '{gcode_line}' (timeout: {timeout:.1f}s).", "WARN")

            except InterruptedError as e:
                msg = str(e) if str(e) else "G-code stream interrupted by user."
                self.queue_message(msg, "WARN" if not str(e) else "ERROR")
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
                    self.serial_connection, found_port, baudrate, initial_position = msg_content
                    self.log_message(f"Connected on {found_port}!", "SUCCESS")
                    
                    self.connection_status_var.set(f"Connected to {found_port}")
                    self.header_status_indicator.set_status("on")
                    self.footer_status_var.set(f"{found_port} @ {baudrate}")
                    
                    self.connect_button.config(text="Disconnect", state=tk.NORMAL, style='GreenRing.TButton')
                    self.port_combobox.config(state=tk.DISABLED)
                    self.baud_entry.config(state=tk.DISABLED)
                    
                    self._set_manual_controls_state(tk.NORMAL)
                    self._set_goto_controls_state(tk.NORMAL)
                    self._set_terminal_controls_state(tk.NORMAL)

                    if self.processed_gcode: self.start_button.config(state=tk.NORMAL)
                    if hasattr(self, 'cancel_connect_button'): self.cancel_connect_button.grid_remove(); self.cancel_connect_button.config(state=tk.DISABLED)
                    
                    self.progress_var.set(0.0)
                    self.progress_label_var.set("Progress: Idle")
                    
                    # Use polled position from M114 if available, otherwise fall back to origin.
                    if initial_position:
                        self.last_cmd_abs_x = initial_position['x']
                        self.last_cmd_abs_y = initial_position['y']
                        self.last_cmd_abs_z = initial_position['z']
                        self.last_cmd_abs_e = initial_position.get('e', 0.0)
                        # Also move the target (cyan) marker to match the synced position
                        self.target_abs_x = initial_position['x']
                        self.target_abs_y = initial_position['y']
                        self.target_abs_z = initial_position['z']
                        self.target_abs_e = initial_position.get('e', 0.0)
                        self.log_message(f"Position synced from printer: X={initial_position['x']:.2f} Y={initial_position['y']:.2f} Z={initial_position['z']:.2f}", "SUCCESS")
                    else:
                        self.last_cmd_abs_x, self.last_cmd_abs_y, self.last_cmd_abs_z = self.PRINTER_BOUNDS['x_min'], self.PRINTER_BOUNDS['y_min'], self.PRINTER_BOUNDS['z_min']
                        self.target_abs_x, self.target_abs_y, self.target_abs_z = self.PRINTER_BOUNDS['x_min'], self.PRINTER_BOUNDS['y_min'], self.PRINTER_BOUNDS['z_min']
                        self.log_message("Position unknown, assuming origin (0, 0, 0).", "WARN")
                    self._update_all_displays()
                    self._update_section_borders()
                
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
                        self.connect_button.config(text="Connect", state=tk.NORMAL)
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
                        self.header_status_indicator.set_status("on")
                
                elif msg_type == "CONNECTION_LOST":
                    self.log_message("Connection lost.", "ERROR")
                    messagebox.showerror("Connection Lost", "Serial connection lost.\nPlease reconnect.")
                    self.disconnect_printer(silent=True)



                elif msg_type == "HOMING_FAILURE":
                    diagnosis = msg_content
                    self.log_message(f"Homing Verification Failed: {diagnosis}", "CRITICAL")
                    messagebox.showwarning("Homing Verification Failure", 
                                          f"The printer has drifted from its expected position.\n\n"
                                          f"DIAGNOSIS:\n{diagnosis}\n\n"
                                          f"The scan has been halted for safety. Please check for mechanical obstructions, "
                                          f"loose belts, or missed steps. Re-home and restart the program when ready.")

                elif msg_type == "DMM_CONNECTED":
                    self.is_dmm_connected = True
                    self.dmm_status_var.set("DMMs: Connected")
                    self.dmm_connect_button.config(text="Disconnect DMMs", state=tk.NORMAL, style='GreenRing.TButton')
                    self.measure_button.config(state=tk.NORMAL)
                    self.log_message("DMMs connected successfully.", "SUCCESS")
                    self._update_section_borders()

                elif msg_type == "DMM_FAIL":
                    self.is_dmm_connected = False
                    self.dmm_status_var.set("DMMs: Error")
                    self.dmm_connect_button.config(text="Connect DMMs", state=tk.NORMAL)
                    self.measure_button.config(state=tk.DISABLED)
                    self.log_message(f"DMM Error: {msg_content}", "ERROR")
                    self._update_section_borders()

                elif msg_type == "MEASUREMENT_RESULT":
                    vals = msg_content 
                    if vals:
                         # Dynamic display based on connected DMMs
                         if self.dmm_group and self.dmm_group.dmms:
                             name = self.dmm_group.dmms[0].name
                             display_str = f"{name}: {vals[0]:.4f}"
                             if len(vals) > 1: 
                                 display_str += f" | {self.dmm_group.dmms[1].name}: {vals[1]:.4f}"
                         else:
                             display_str = f"Val: {vals[0]:.4f}"
                             
                         self.last_measurement_var.set(display_str)
                         self.log_message(f"Measured: {', '.join([f'{v:.4f}' for v in vals])}", "INFO")




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
    
    def _update_section_borders(self):
        """Updates the LabelFrame border colour to green when each section's setup condition is met."""
        GREEN = 'Green.TLabelframe'
        YELLOW = 'Yellow.TLabelframe'
        PURPLE = 'Purple.TLabelframe'

        # CONNECTION — green once printer is connected
        if hasattr(self, '_conn_frame'):
            if getattr(self, 'hardware_fault', False):
                self._conn_frame.configure(style=PURPLE)
                if hasattr(self, 'connect_button'):
                    self.connect_button.configure(style='PurpleRing.TButton')
            else:
                connected = bool(self.serial_connection)
                self._conn_frame.configure(style=GREEN if connected else YELLOW)
                if hasattr(self, 'connect_button'):
                    if connected:
                        self.connect_button.configure(style='GreenRing.TButton', text='Disconnect')
                    else:
                        self.connect_button.configure(style='YellowRing.TButton', text='Connect')

        # MEASUREMENT — green when DMM connected AND (Log to CSV off OR file path set)
        if hasattr(self, '_meas_frame'):
            log_ok = not self.log_measurements_enabled.get() or bool(self.log_filepath_var.get().strip())
            self._meas_frame.configure(style=GREEN if (self.is_dmm_connected and log_ok) else YELLOW)

        # SETUP — green when center has been marked AND collision test completed
        if hasattr(self, '_setup_frame'):
            setup_ok = self.center_marked and self.rotation_crash_test_complete
            if not self.center_marked:
                self._setup_frame.configure(style=PURPLE)
            else:
                self._setup_frame.configure(style=GREEN if setup_ok else YELLOW)
            
            if hasattr(self, 'mark_center_button'):
                self.mark_center_button.configure(style='GreenRing.TButton' if self.center_marked else 'PurpleRing.TButton')
            
        # EXECUTION CONTROL — green when gcode is loaded
        if hasattr(self, '_ctrl_frame'):
            ctrl_ok = bool(self.gcode_filepath)
            self._ctrl_frame.configure(style=GREEN if ctrl_ok else YELLOW)

        if hasattr(self, 'collision_test_button'):
            self.collision_test_button.configure(style='GreenRing.TButton' if self.rotation_crash_test_complete else 'YellowRing.TButton')


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
             
        # Auto-trigger "set tilt as level" before proceeding
        self._mark_tilt_as_level()
             
        try:
            self.center_x_var.set(f"{self.last_cmd_abs_x:.2f}")
            self.center_y_var.set(f"{self.last_cmd_abs_y:.2f}")
            self.center_z_var.set(f"{self.last_cmd_abs_z:.2f}")
            if self.last_cmd_abs_e is not None:
                self.center_e_var.set(f"{self.last_cmd_abs_e:.2f}")
            
            e_str = f", E={self.last_cmd_abs_e:.2f}" if self.last_cmd_abs_e is not None else ""
            self.log_message(f"New center marked at: X={self.last_cmd_abs_x:.2f}, Y={self.last_cmd_abs_y:.2f}, Z={self.last_cmd_abs_z:.2f}{e_str}", "SUCCESS")
            self.center_marked = True
            self.mark_center_button.config(style='GreenRing.TButton')
            try:
                self.shortcut_mark_center_button.config(style='GreenRing.TButton')
            except AttributeError:
                pass
                
            # Trigger the same logic as if the user changed the entry fields manually.
            self._on_center_change()
            self._update_section_borders()
        except Exception as e:
             self.log_message(f"Error marking center: {e}", "ERROR")

    def _mark_tilt_as_level(self):
        """
        Sends G92 E0 to the printer to set the current physical tilt as absolute 0.
        Updates internal models to reflect this new origin.
        """
        if not self.serial_connection or not self.serial_connection.is_open:
            self.log_message("Cannot mark tilt level: Printer not connected.", "WARN")
            messagebox.showwarning("Mark Level Failed", "Printer is not connected.")
            return

        if self.last_cmd_abs_e is None:
            self.log_message("Cannot mark tilt level: No known E position.", "WARN")
            messagebox.showwarning("Mark Level Failed", "Cannot mark tilt level.\nNo known tilt position.\nTry moving the printer first.")
            return
            
        try:
            # Tell printer to set current E to 0
            self._send_manual_command("G92 E0")
            
            # Update internal absolute tracking
            self.last_cmd_abs_e = 0.0
            self.target_abs_e = 0.0
            
            # Set the user's defined "Center E" point to 0.00 to match
            self.center_e_var.set("0.00")
            
            self.log_message("Absolute tilt position reset to 0°", "SUCCESS")
            self._on_center_change()
            self._update_all_displays()
        except Exception as e:
            self.log_message(f"Error marking tilt level: {e}", "ERROR")

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
            center_e = float(self.center_e_var.get())
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
            if mode == "absolute":
                self.last_cmd_e_display_var.set(f"{self.last_cmd_abs_e:.2f}")
            else:
                self.last_cmd_e_display_var.set(f"{self.last_cmd_abs_e - center_e:.2f}")
        else:
            self.last_cmd_e_display_var.set("N/A")

        # --- Update "Target" (Blue) Display Labels ---
        if mode == "absolute":
            self.goto_x_display_var.set(f"{self.target_abs_x:.2f}")
            self.goto_y_display_var.set(f"{self.target_abs_y:.2f}")
            self.goto_z_display_var.set(f"{self.target_abs_z:.2f}")
            self.goto_e_display_var.set(f"{self.target_abs_e:.2f}")
        else: # "relative"
            self.goto_x_display_var.set(f"{self.target_abs_x - center_x:.2f}")
            self.goto_y_display_var.set(f"{self.target_abs_y - center_y:.2f}")
            self.goto_z_display_var.set(f"{self.target_abs_z - center_z:.2f}")
            self.goto_e_display_var.set(f"{self.target_abs_e - center_e:.2f}")

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
                center_e = float(self.center_e_var.get())
                new_abs_e = val_e + center_e if mode == "relative" else val_e
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

            selected_mode = DMM_MODES.get(self.dmm_mode_var.get(), 'VOLT:DC')
            self.queue_message(f"Initializing DMMs (Mode: {self.dmm_mode_var.get()})...")
            self.dmm_group = DmmGroup(DMM_CONFIG)
            self.dmm_group.initialize(mode=selected_mode, ip_prefix=self.dmm_ip_prefix_var.get())
            self.message_queue.put(("DMM_CONNECTED", None))
        except Exception as e:
            self.queue_message(f"DMM Connect Error:\n{e}", "ERROR")
            self.message_queue.put(("DMM_FAIL", str(e)))

    def disconnect_dmms(self):
        if self.dmm_group:
            try:
                self.dmm_group.close()
            except: pass
        self.dmm_group = None
        self.is_dmm_connected = False
        self.dmm_status_var.set("DMMs: Disconnected")
        self.dmm_connect_button.config(text="Connect DMMs", state=tk.NORMAL, style='YellowRing.TButton')
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

    def _take_measurement(self):
        """
        Takes a reading after the specified settling time delay.
        Returns the final averaged list of values returned directly by the DMM.
        """
        if not self.dmm_group: return []

        pre_delay = self.pre_measure_delay_var.get()

        # 1. Pre-Measure Delay
        if pre_delay > 0:
            self.queue_message(f"Settling... ({pre_delay}s)")
            time.sleep(pre_delay)

        # 2. Trigger and Read
        self.dmm_group.trigger()
        return self.dmm_group.read()

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
            self.log_message(f"Selected log file: {filename}", "INFO")
            self.browse_log_btn.config(style='GreenRing.TButton')

    def _log_measurement_to_file(self, values, coords=None):
        filepath = self.log_filepath_var.get()
        if not filepath:
            self.queue_message("Log Error: No filename specified.", "ERROR")
            return

        # Check if file exists to write header
        file_exists = False
        try:
            if os.path.exists(filepath) and os.path.getsize(filepath) > 0:
                file_exists = True
        except Exception: 
            pass # Permission error or other
        
        try:
            with open(filepath, 'a', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                
                # Write header if new file
                if not file_exists:
                    mode = self.dmm_mode_var.get()
                    headers = ["Timestamp", "X", "Y", "Z", "E"] + [f"{d[2]} ({mode})" for d in DMM_CONFIG]
                    writer.writerow(headers)
                    self.queue_message(f"Created log file: {filepath}", "SUCCESS")

                if coords:
                    log_x = coords.get('x', 0.0)
                    log_y = coords.get('y', 0.0)
                    log_z = coords.get('z', 0.0)
                    log_e = coords.get('e', 0.0)
                else:
                    # Get current position (use last known)
                    log_x = self.last_cmd_abs_x if self.last_cmd_abs_x is not None else 0.0
                    log_y = self.last_cmd_abs_y if self.last_cmd_abs_y is not None else 0.0
                    log_z = self.last_cmd_abs_z if self.last_cmd_abs_z is not None else 0.0
                    log_e = self.last_cmd_abs_e if self.last_cmd_abs_e is not None else 0.0
                
                row = [datetime.now().isoformat(), log_x, log_y, log_z, log_e] + values
                writer.writerow(row)
        except Exception as log_err:
            self.queue_message(f"Log Write Error: {log_err}", "ERROR")


    # --- Collision Avoidance Test Screen ---

    def _open_collision_test_screen(self):
        """Swaps the main UI for the Collision Avoidance Test screen."""
        if not self.serial_connection:
            messagebox.showerror("Error", "Not connected to printer.")
            return

        if not self.processed_gcode:
            messagebox.showerror("Setup Error", "Error: Must load test profile before the test")
            return

        # Hide the main view
        self.main_view_frame.pack_forget()

        # Create the test view frame
        self.test_view_frame = tk.Frame(self.parent, bg=self.COLOR_BG)
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
                                       command=self._stop_collision_test)
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

    def _stop_collision_test(self):
        """
        Emergency stop specific to the collision avoidance test.
        Sends M112 immediately to kill all motion, fully disconnects,
        and paints UI purple to indicate hardware fault requirement.
        """
        self.log_message("!!! COLLISION TEST ABORTED — Emergency Stop !!!", "CRITICAL")
        self.hardware_fault = True

        # Signal the worker thread to abort immediately
        self.is_collision_test_running = False
        self.stop_event.set()
        self.pause_event.set()  # Unblock any paused state

        # Invalidate the test result
        self.rotation_crash_test_complete = False

        if hasattr(self, 'lbl_test_status') and self.lbl_test_status.winfo_exists():
            self.lbl_test_status.config(
                text="Test Failed - please cycle power on the test machine, unplug the USB from the Raspberry Pi, and reconnect.", 
                fg=self.COLOR_ACCENT_RED,
                font=("Rajdhani", 12, "bold"),
                wraplength=400
            )

        # Send M112 hard-stop directly
        if self.serial_connection:
            try:
                self.serial_connection.write(b'M112\n')
                import time as _t
                _t.sleep(0.3)
                self.serial_connection.reset_input_buffer()
                self.serial_connection.reset_output_buffer()
                self.log_message("M112 sent. Printer stopped.", "WARN")
            except Exception as e:
                self.log_message(f"Error sending M112: {e}", "ERROR")
        else:
            self.log_message("Not connected — no M112 sent.", "WARN")

        # Automatically disconnect the printer
        self.disconnect_printer(silent=True)

        # Update status indicators
        self.status_indicator.set_status("error")
        self.header_status_indicator.set_status("error")

        # Reset the worker flags and update display colors
        self.is_manual_command_running = False
        self.root.after(0, self._update_all_displays)
        self.root.after(0, self._update_section_borders)

        # Clear the stop event so subsequent manual commands can run
        self.stop_event.clear()

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
            def update_status(msg):
                self.queue_message(msg, "INFO")
                if hasattr(self, 'lbl_test_status') and self.lbl_test_status.winfo_exists():
                    self.lbl_test_status.config(text=msg)

            self.is_manual_command_running = True
            self.is_collision_test_running = True
            self.root.after(0, self._update_all_displays)
            
            # --- 1. Range Calculation ---
            min_e = 0.0
            max_e = 0.0
            pause_time_ms = 0
            found_e = False

            # Scan processed G-code for E limits and Pause Time (G4 Pxxx)
            # processed_gcode contains strings like "G1 X10 Y10 E45 F1000"
            for line in self.processed_gcode:
                import re as local_re
                if 'E' in line.upper():
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
                if 'G4' in line.upper() and 'P' in line.upper() and pause_time_ms == 0:
                    match = local_re.search(r"P(\d+)", line)
                    if match:
                        pause_time_ms = int(match.group(1))
            
            if not found_e:
                update_status("No rotation (E) found in G-code. Skipping rotation test.")
                self.rotation_crash_test_complete = True  # Trivially passed — no rotation in file
            else:
                update_status(f"Test Range: E{min_e:.1f}° to E{max_e:.1f}°")

            
            # --- 2. Execution ---

            # Get Center & Min Z
            try:
                cx = float(self.center_x_var.get())
                cy = float(self.center_y_var.get())
                cz = float(self.center_z_var.get())
            except ValueError:
                self.queue_message("Invalid Center Coordinates!", "ERROR")
                return

            if not self.ordered_z_values:
                min_z = cz
            else:
                min_z = min(self.ordered_z_values)

            speed_xy = 1000 # mm/min
            speed_e = 500 # deg/min slowly
            
            def send_wait(cmd_str, timeout_s=30):
                lines = [c.strip() for c in cmd_str.split('\n') if c.strip()]
                for c in lines:
                    with self.serial_lock:
                        self.serial_connection.write((c + '\n').encode('utf-8'))
                for _ in lines:
                    if not self._wait_for_ok(timeout=timeout_s):
                        raise Exception(f"Timeout waiting for OK")

            # Step 1: Move Z axis to scan's minimum Z value first
            update_status(f"Moving to Min Z ({min_z})...")
            send_wait(f"G90\nG1 Z{min_z} F{speed_xy}\nM400", timeout_s=30)

            # Step 2: Move to X,Y center
            update_status(f"Moving to Center ({cx}, {cy})...")
            send_wait(f"G1 X{cx} Y{cy} F{speed_xy}\nM400", timeout_s=30)

            # Step 2.5: Home X, then Y, then Z sequentially
            update_status("Homing X axis...")
            send_wait("G28 X\nM400", timeout_s=60)

            update_status("Homing Y axis...")
            send_wait("G28 Y\nM400", timeout_s=60)

            update_status("Homing Z axis...")
            send_wait("G28 Z\nM400", timeout_s=60)

            if found_e:
                # Step 3: Move to Center X, Y, Z before tilting
                update_status("moving to center position")
                send_wait(f"G90\nG1 X{cx} Y{cy} Z{cz} F{speed_xy}\nM400", timeout_s=30)

                # Step 4: Tilt to Max E slowly
                update_status(f"Moving to max tilt angle ({max_e:.1f}°)...")
                cmd = self._apply_e_conversion(f"G1 E{max_e:.2f} F{speed_e}\nM400")
                send_wait(cmd, timeout_s=60)
                
                # Step 4: Move to Min E slowly
                update_status(f"Moving to min tilt angle ({min_e:.1f}°)...")
                cmd = self._apply_e_conversion(f"G1 E{min_e:.2f} F{speed_e}\nM400")
                send_wait(cmd, timeout_s=60)

            # Step 5: Hold for Pause Time OR Multimeter Reading at Min E
            if getattr(self, 'is_dmm_connected', False) and self.auto_measure_enabled.get():
                update_status("Auto-Measure: Measuring at Min Tilt...")
                if self.dmm_group:
                    try:
                        vals = self._take_measurement()
                        self.message_queue.put(("MEASUREMENT_RESULT", vals))
                        if self.log_measurements_enabled.get():
                            self._log_measurement_to_file(vals, coords={'x': cx, 'y': cy, 'z': min_z, 'e': min_e})
                    except Exception as e:
                        update_status(f"Auto-Measure Error: {e}")
            elif pause_time_ms > 0:
                sec = pause_time_ms / 1000.0
                update_status(f"Holding at Min Tilt for {sec}s...")
                import time
                time.sleep(sec)

            # Step 7: Final Completion Sequence
            update_status("Test sequence finished. Holding 5s before returning to center...")
            import time as _time
            _time.sleep(5)

            update_status("Returning to center position...")
            # Move X, Y, Z, and E back to center/home
            cmd = self._apply_e_conversion(f"G1 Z{cz} X{cx} Y{cy} E0 F{speed_xy}\nM400")
            send_wait(cmd, timeout_s=30)

            self.rotation_crash_test_complete = True
            self._update_section_borders()
            
            # Request explicit position polling now that printer is back.
            if self.serial_connection:
                update_status("Polling for current position...")
                self.serial_connection.write(b"M114\n")
                
            update_status("Collision Test Complete!")
            
            # Show completed popup
            self.root.after(0, lambda: messagebox.showinfo("Test Complete", "Test Complete - Verify current location is at center position."))
            
        except Exception as e:
            self.queue_message(f"Test Failed: {e}", "ERROR")
            if "Error: Must load" in str(e):
                messagebox.showerror("Setup Error", str(e))

        finally:
            self.is_manual_command_running = False
            self.is_collision_test_running = False
            # Update UI via queue or invoke
            self.root.after(0, self._update_all_displays)
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
