import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import serial # type: ignore
import time
import re
import threading
import queue
import math
import serial.tools.list_ports
import csv
import json
from datetime import datetime
import os
import webbrowser
import utils
from utils import PRINTER_BOUNDS, EventBus
from sender_components.telemetry_service import DMMTelemetryProvider, VivigoTelemetryProvider
from sender_components.dmm_manager import DMM_IP_PREFIX, DMM_CONFIG, DMM_MODES, DmmGroup, HAS_PYVISA
from sender_components.serial_engine import SerialEngine
from sender_components.plot_manager import PlotManager
from sender_components.motion_controls import MotionControls
from sender_components.ui_layout import (
    HeaderBar, FooterBar, ConnectionFrame, MeasurementFrame,
    ExecutionControlFrame, FileCenterFrame, TerminalTab,
    ManualJogFrame, PositionControlFrame, DisplayTab
)
from sequence_manager import SequenceManager

# Local-only theme constant (no equivalent in utils.py)
COLOR_PENDING_RING = "#c4c1ff"

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_PROJECT_ROOT = os.path.dirname(_SCRIPT_DIR)
_DATA_DIR = os.path.join(_PROJECT_ROOT, "data")

class GCodeSenderGUI:
    """
    Orchestrator for the G-Code Sender application.
    Delegates hardware comms to SerialEngine, motion to MotionControls,
    and plotting to PlotManager.
    """
    def __init__(self, parent_frame, vivigo_panel=None, pop_viz_panel=None, event_bus=None):
        self.parent = parent_frame
        self.root = parent_frame.winfo_toplevel()
        self.vivigo_panel = vivigo_panel
        self.pop_viz_panel = pop_viz_panel
        self.event_bus = event_bus if event_bus else EventBus()
        self.telemetry_provider = None

        self._setup_theme()
        self._setup_state()
        self._setup_variables()
        
        # Initialize Core Engines
        self.serial_engine = SerialEngine(self.message_queue, self.stop_event, self.pause_event)
        self.plot_manager = PlotManager(self)
        self.motion_controls = MotionControls(self)
        self.dmm_telemetry = DMMTelemetryProvider(self)
        self.vivigo_telemetry = VivigoTelemetryProvider(self)
        
        self._setup_serial_callbacks()
        self._build_ui()

        # Start lifecycle loops
        self.after_id = self.root.after(100, self.check_message_queue)
        self.root.after(300, self.rescan_ports)
        self.root.after(150, self._update_all_displays)
        
        self._bind_globals()
        
        # --- Final Layout Adjustments ---
        self.left_panel_scrollable.update_idletasks()
        required_width = self.left_panel_scrollable.winfo_reqwidth() + 20 
        self.left_canvas_frame.config(width=required_width)
        self.left_canvas.config(width=required_width)
        self.paned_window.paneconfigure(self.left_canvas_frame, width=required_width, minsize=required_width)
        self.root.update_idletasks()
        self.paned_window.sash_place(0, required_width + 5, 0)


    def _setup_theme(self):
        # Backward-compatible theme aliases for external consumers (ui_layout.py,
        # plot_manager.py) that read gui.COLOR_*/gui.FONT_*. These alias the
        # canonical constants in utils.py rather than redefining literals, so
        # there is a single source of truth for the palette.
        self.COLOR_BG = utils.COLOR_BG
        self.COLOR_PANEL_BG = utils.COLOR_PANEL_BG
        self.COLOR_BORDER = utils.COLOR_BORDER
        self.COLOR_TEXT_PRIMARY = utils.COLOR_TEXT_PRIMARY
        self.COLOR_TEXT_SECONDARY = utils.COLOR_TEXT_SECONDARY
        self.COLOR_ACCENT_CYAN = utils.COLOR_ACCENT_CYAN
        self.COLOR_ACCENT_GREEN = utils.COLOR_ACCENT_GREEN
        self.COLOR_ACCENT_AMBER = utils.COLOR_ACCENT_AMBER
        self.COLOR_ACCENT_RED = utils.COLOR_ACCENT_RED
        self.COLOR_BLACK = utils.COLOR_BLACK
        self.COLOR_PENDING_RING = COLOR_PENDING_RING

        self.FONT_HEADER = utils.FONT_HEADER
        self.FONT_BODY = utils.FONT_BODY
        self.FONT_BODY_SMALL = utils.FONT_BODY_SMALL
        self.FONT_BODY_BOLD = utils.FONT_BODY_BOLD
        self.FONT_MONO = utils.FONT_MONO
        self.FONT_DRO = utils.FONT_DRO
        self.FONT_TERMINAL = utils.FONT_TERMINAL

        self.root.configure(bg=utils.COLOR_BG)
        style = ttk.Style()
        style.theme_use('clam')
        self._configure_styles(style)

    def _configure_styles(self, style):
        style.configure('.', background=utils.COLOR_PANEL_BG, foreground=utils.COLOR_TEXT_PRIMARY, fieldbackground=utils.COLOR_BLACK, bordercolor=utils.COLOR_BORDER, lightcolor=utils.COLOR_BORDER, darkcolor=utils.COLOR_BORDER, font=utils.FONT_BODY)
        style.map('.', background=[('disabled', utils.COLOR_PANEL_BG), ('active', utils.COLOR_PANEL_BG)], foreground=[('disabled', utils.COLOR_TEXT_SECONDARY)], bordercolor=[('focus', utils.COLOR_ACCENT_CYAN), ('active', utils.COLOR_BORDER)])
        style.configure('TFrame', background=utils.COLOR_BG)
        style.configure('Panel.TFrame', background=utils.COLOR_PANEL_BG)
        style.configure('Header.TFrame', background=utils.COLOR_PANEL_BG, bordercolor=utils.COLOR_BORDER, borderwidth=1, relief='solid')
        style.configure('Footer.TFrame', background=utils.COLOR_BLACK, bordercolor=utils.COLOR_BORDER, borderwidth=1, relief='solid')
        style.configure('Black.TFrame', background=utils.COLOR_BLACK)
        style.configure('TLabelframe', background=utils.COLOR_PANEL_BG, bordercolor=utils.COLOR_BORDER, borderwidth=1, relief=tk.SOLID, padding=16)
        style.configure('TLabelframe.Label', background=utils.COLOR_PANEL_BG, foreground=utils.COLOR_TEXT_SECONDARY, font=("Rajdhani", 13, "bold"), padding=(10, 5))
        style.configure('TLabel', background=utils.COLOR_PANEL_BG, foreground=utils.COLOR_TEXT_PRIMARY, font=utils.FONT_BODY)
        style.configure('DRO.TLabel', font=utils.FONT_MONO, padding=(5, 5), background=utils.COLOR_BLACK, foreground=utils.COLOR_TEXT_SECONDARY, borderwidth=1, relief='sunken', anchor='w')
        style.configure('Red.DRO.TLabel', font=utils.FONT_DRO, width=8, padding=(5, 5), background=utils.COLOR_BLACK, foreground=utils.COLOR_ACCENT_RED, borderwidth=0, relief='flat', anchor='e')
        style.configure('Blue.DRO.TLabel', font=utils.FONT_DRO, width=8, padding=(5, 5), background=utils.COLOR_BLACK, foreground=utils.COLOR_ACCENT_AMBER, borderwidth=0, relief='flat', anchor='e')
        style.configure('Subtle.DRO.TLabel', font=utils.FONT_BODY_SMALL, padding=(5, 0), background=utils.COLOR_BLACK, foreground=utils.COLOR_TEXT_SECONDARY, borderwidth=0, relief='flat', anchor='e')
        style.configure('TButton', background=utils.COLOR_PANEL_BG, foreground=utils.COLOR_TEXT_PRIMARY, bordercolor=utils.COLOR_BORDER, borderwidth=1, relief=tk.SOLID, padding=(12, 8), font=utils.FONT_BODY)
        style.map('TButton', background=[('active', '#2c333e'), ('pressed', utils.COLOR_BLACK)], foreground=[('active', utils.COLOR_ACCENT_CYAN)], bordercolor=[('active', utils.COLOR_ACCENT_CYAN)])
        
        style.configure('Primary.TButton', background=utils.COLOR_ACCENT_CYAN, foreground=utils.COLOR_BLACK, font=utils.FONT_BODY_BOLD)
        style.configure('Danger.TButton', background=utils.COLOR_ACCENT_RED, foreground=utils.COLOR_TEXT_PRIMARY, font=utils.FONT_BODY_BOLD)
        style.configure('Amber.TButton', background=utils.COLOR_ACCENT_AMBER, foreground=utils.COLOR_BLACK, font=utils.FONT_BODY_BOLD)
        
        style.configure('GreenRing.TButton', background=utils.COLOR_PANEL_BG, foreground=utils.COLOR_TEXT_PRIMARY, bordercolor=utils.COLOR_ACCENT_GREEN, borderwidth=1, relief=tk.SOLID, padding=(12, 8))
        style.configure('PurpleRing.TButton', background=utils.COLOR_PANEL_BG, foreground=utils.COLOR_TEXT_PRIMARY, bordercolor=COLOR_PENDING_RING, borderwidth=1, relief=tk.SOLID, padding=(12, 8))
        
        style.configure('Segment.TButton', background=utils.COLOR_PANEL_BG, foreground=utils.COLOR_TEXT_SECONDARY, padding=(10, 5), font=utils.FONT_BODY_SMALL)
        style.configure('Segment.Active.TButton', background=utils.COLOR_ACCENT_CYAN, foreground=utils.COLOR_BLACK, padding=(10, 5), font=utils.FONT_BODY_SMALL)
        style.configure('ViewCube.TButton', padding=(2, 2), font=("Inter", 12), width=2)
        style.configure('Custom.Toggle.Off.TButton', background=utils.COLOR_PANEL_BG, foreground=utils.COLOR_TEXT_SECONDARY, bordercolor=utils.COLOR_BORDER, font=("Rajdhani", 9, "bold"), padding=(5, 3))
        style.configure('Custom.Toggle.On.TButton', background=utils.COLOR_ACCENT_CYAN, foreground=utils.COLOR_BLACK, bordercolor=utils.COLOR_ACCENT_CYAN, font=("Rajdhani", 9, "bold"), padding=(5, 3))
        style.configure('Small.TButton', background=utils.COLOR_PANEL_BG, foreground=utils.COLOR_TEXT_PRIMARY, bordercolor=utils.COLOR_BORDER, font=("Rajdhani", 9, "bold"), padding=(5, 3))
        
        style.configure('Green.TLabelframe', background=utils.COLOR_PANEL_BG, bordercolor=utils.COLOR_ACCENT_GREEN, relief=tk.SOLID, borderwidth=2)
        style.configure('Purple.TLabelframe', background=utils.COLOR_PANEL_BG, bordercolor=COLOR_PENDING_RING, relief=tk.SOLID, borderwidth=3)
        style.configure('Cyan.TLabelframe', background=utils.COLOR_PANEL_BG, bordercolor=utils.COLOR_ACCENT_CYAN, relief=tk.SOLID, borderwidth=2)

    def _setup_state(self):
        self.serial_connection = None
        self.is_disconnecting = False
        self.is_closing = False
        self.homing_performed = False
        self.gcode_filepath = None
        self.processed_gcode = []
        self.is_sending = False
        self.is_paused = False
        self.is_manual_command_running = False
        self.is_collision_test_running = False
        self.center_marked = False
        self.rotation_crash_test_complete = False
        self.last_hub_data = {} # Shared state for latest hub readings
        self.stop_event = threading.Event()
        self.pause_event = threading.Event()
        self.pause_event.set()
        self.cancel_connect_event = threading.Event()
        self.message_queue = queue.Queue()

    def _setup_variables(self):
        self.PRINTER_BOUNDS = dict(PRINTER_BOUNDS)
        self.file_path_var = tk.StringVar(value="No file selected")
        self.header_file_var = tk.StringVar(value="NO FILE")
        self.center_x_var = tk.StringVar(value="110.0")
        self.center_y_var = tk.StringVar(value="110.0")
        self.center_z_var = tk.StringVar(value="0.0")
        self.center_rot_var = tk.StringVar(value="0.0")
        self.available_ports = ["Auto-detect"] + self._get_available_ports()
        _sim_port = os.environ.get("SEED_SIM_PORT")
        if _sim_port and _sim_port in self.available_ports:
            _default_port = _sim_port
        else:
            _default_port = self.available_ports[0] if self.available_ports else ""
        self.port_var = tk.StringVar(value=_default_port)
        self.baud_var = tk.StringVar(value="115200")
        self.connection_status_var = tk.StringVar(value="Status: Disconnected")
        self.footer_status_var = tk.StringVar(value="COM: -- @ --")
        self.footer_coords_var = tk.StringVar(value="X: N/A  Y: N/A  Z: N/A")
        self.dmm_ip_prefix_var = tk.StringVar(value=DMM_IP_PREFIX)
        self.dmm_group = None
        self.is_dmm_connected = False
        self.auto_measure_enabled = tk.BooleanVar(value=True)
        self.log_measurements_enabled = tk.BooleanVar(value=True)
        self.log_filepath_var = tk.StringVar(value="")
        self.dmm_status_var = tk.StringVar(value="DMMs: Disconnected")
        self.dmm_mode_var = tk.StringVar(value="DC Voltage")
        self.active_dmm_mode = "DC Voltage"
        self.hub_modes_list = ["V_rec", "V_sys", "V_bst", "I_sys", "I_bst", "Temp_bst", "Temp_inv", "Temp_rec"]
        self.hub_mode_vars = {m: tk.BooleanVar(value=(m == "V_rec")) for m in self.hub_modes_list}
        self.active_hub_modes = ["V_rec"]
        self.data_source_var = tk.StringVar(value="DMM")
        self.active_data_source = "DMM"
        self.data_source_is_hub = tk.BooleanVar(value=False)
        self.last_measurement_var = tk.StringVar(value="Last: --")
        self.pre_measure_delay_var = tk.DoubleVar(value=0.2)
        self.jog_step_var = tk.StringVar(value="10")
        self.jog_feedrate_var = tk.StringVar(value="1000")
        self.rotation_step_var = tk.StringVar(value="5")
        self.rotation_feedrate_var = tk.StringVar(value="3000")
        self.mm_per_degree_var = tk.DoubleVar(value=8.888)
        self.progress_var = tk.DoubleVar(value=0.0)
        self.progress_label_var = tk.StringVar(value="Progress: Idle")
        self.toolpath_3d_opacity_var = tk.DoubleVar(value=0.8)
        self.hub_connected = tk.BooleanVar(value=False)
        self.last_cmd_x_display_var = tk.StringVar(value="N/A")
        self.last_cmd_y_display_var = tk.StringVar(value="N/A")
        self.last_cmd_z_display_var = tk.StringVar(value="N/A")
        self.last_cmd_rot_display_var = tk.StringVar(value="N/A")
        self.goto_x_display_var = tk.StringVar(value="0.00")
        self.goto_y_display_var = tk.StringVar(value="0.00")
        self.goto_z_display_var = tk.StringVar(value="0.00")
        self.goto_rot_display_var = tk.StringVar(value="0.00")
        self.coord_mode = tk.StringVar(value="absolute")
        self.is_3d_plot_enabled = tk.BooleanVar(value=True)
        self.is_2d_plot_enabled = tk.BooleanVar(value=True)
        self.target_abs_x = self.PRINTER_BOUNDS['x_max'] / 2
        self.target_abs_y = self.PRINTER_BOUNDS['y_max'] / 2
        self.target_abs_z = self.PRINTER_BOUNDS['z_max'] / 4
        self.target_abs_rot = 0.0
        self.last_cmd_abs_x = self.last_cmd_abs_y = self.last_cmd_abs_z = self.last_cmd_abs_rot = None
        self.toolpath_by_layer, self.move_to_layer_map, self.ordered_z_values = {}, [], []
        self.completed_move_count = 0
        self.command_history = []; self.history_index = 0

    def _setup_serial_callbacks(self):
        self.serial_engine.set_callback('take_measurement', self._perform_automatic_capture)
        self.serial_engine.set_callback('homing_verification', self._homing_verification_routine)
        self.serial_engine.set_callback('handle_homing_failure', self._handle_homing_failure)
        self.serial_engine.set_callback('get_hub_panel', lambda: self.vivigo_panel)
        self.serial_engine.set_callback('apply_e_conversion', self._apply_rot_conversion)
        self.serial_engine.set_callback('parse_coords', self._parse_gcode_coords)
        self.serial_engine.set_callback('get_speed_cap', lambda: getattr(self, '_active_speed_cap', 99999))
        self.serial_engine.set_callback('set_speed_cap', lambda x: setattr(self, '_active_speed_cap', x))

    def _build_ui(self):
        self.header_bar = HeaderBar(self.parent, self)
        self.footer_bar = FooterBar(self.parent, self)
        self.main_view_frame = ttk.Frame(self.parent, style='TFrame')
        self.main_view_frame.pack(fill=tk.BOTH, expand=True)
        container = ttk.Frame(self.main_view_frame, padding=5, style='TFrame')
        container.pack(fill=tk.BOTH, expand=True)
        self.paned_window = tk.PanedWindow(container, orient=tk.HORIZONTAL, bg=utils.COLOR_BG, sashwidth=6)
        self.paned_window.pack(fill=tk.BOTH, expand=True)
        self._build_left_panel()
        self._build_right_panel()

    def _build_left_panel(self):
        self.left_canvas_frame = ttk.Frame(self.paned_window, style='TFrame')
        self.left_canvas = tk.Canvas(self.left_canvas_frame, highlightthickness=0, bg=utils.COLOR_BG)
        scrollbar = ttk.Scrollbar(self.left_canvas_frame, orient="vertical", command=self.left_canvas.yview)
        self.left_panel_scrollable = ttk.Frame(self.left_canvas, style='TFrame')
        self.left_panel_scrollable.bind("<Configure>", lambda e: self.left_canvas.configure(scrollregion=self.left_canvas.bbox("all")))
        self.canvas_window = self.left_canvas.create_window((0, 0), window=self.left_panel_scrollable, anchor="nw")
        self.left_canvas.bind('<Configure>', lambda e: self.left_canvas.itemconfig(self.canvas_window, width=e.width))
        self.left_canvas.configure(yscrollcommand=scrollbar.set)
        self.left_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self._conn_frame = ConnectionFrame(self.left_panel_scrollable, self)
        self._meas_frame = MeasurementFrame(self.left_panel_scrollable, self)
        self.manual_jog_frame = ManualJogFrame(self.left_panel_scrollable, self)
        self._setup_frame = FileCenterFrame(self.left_panel_scrollable, self)
        self._ctrl_frame = ExecutionControlFrame(self.left_panel_scrollable, self)
        self.pos_ctrl_frame = PositionControlFrame(self.left_panel_scrollable, self)
        self.paned_window.add(self.left_canvas_frame, minsize=420)

        # Bind mousewheel to canvas and all scrollable panel children
        def _bind_mousewheel_recursive(widget):
            widget.bind("<MouseWheel>", self._on_mousewheel_scroll)
            widget.bind("<Button-4>", self._on_mousewheel_scroll)
            widget.bind("<Button-5>", self._on_mousewheel_scroll)
            for child in widget.winfo_children():
                _bind_mousewheel_recursive(child)
        
        self.left_canvas.bind("<MouseWheel>", self._on_mousewheel_scroll)
        self.left_canvas.bind("<Button-4>", self._on_mousewheel_scroll)
        self.left_canvas.bind("<Button-5>", self._on_mousewheel_scroll)
        _bind_mousewheel_recursive(self.left_panel_scrollable)

    def _build_right_panel(self):
        right_panel = ttk.Frame(self.paned_window, style='TFrame', padding=(5, 0, 0, 0))
        self.notebook = ttk.Notebook(right_panel, style='TNotebook')
        self.notebook.pack(fill=tk.BOTH, expand=True)
        self.cli_tab = TerminalTab(self.notebook, self)
        self.display_tab = DisplayTab(self.notebook, self)
        self.notebook.add(self.cli_tab, text="Command Line")
        self.notebook.add(self.display_tab, text="3D View")
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)
        self.paned_window.add(right_panel, minsize=300)

    def _bind_globals(self):
        self.log_measurements_enabled.trace_add('write', lambda *a: self._update_section_borders())
        self.log_filepath_var.trace_add('write', lambda *a: self._update_section_borders())
        self.file_path_var.trace_add('write', lambda *a: self._update_section_borders())
        
        # Thread-safe variable caching for background access
        def update_cache(*a):
            self.active_dmm_mode = self.dmm_mode_var.get()
            self.active_data_source = self.data_source_var.get()
            self.active_hub_modes = [m for m in self.hub_modes_list if self.hub_mode_vars[m].get()]
            
        self.dmm_mode_var.trace_add('write', update_cache)
        self.data_source_var.trace_add('write', update_cache)
        for var in self.hub_mode_vars.values():
            var.trace_add('write', update_cache)

    # --- DELEGATION METHODS ---
    def _jog(self, axis, direction): self.motion_controls.jog(axis, direction)
    def _home_all(self): self.motion_controls.home_all()
    def _go_to_position(self): self.motion_controls.go_to_position(self.target_abs_x, self.target_abs_y, self.target_abs_z, self.target_abs_rot)
    def _go_to_center(self): self.motion_controls.go_to_center()

    # --- MISSING CALLBACKS ---
    def _perform_automatic_capture(self):
        """Fetches telemetry data and returns it to the calling thread (SerialEngine)."""
        if self.active_data_source == "DMM":
            return self.dmm_telemetry.fetch_data()
        else:
            return self.vivigo_telemetry.fetch_data()

    def _measure_thread(self):
        """Background thread for performing measurements."""
        try:
            if self.active_data_source == "DMM":
                vals = self.dmm_telemetry.fetch_data()
            else:
                vals = self.vivigo_telemetry.fetch_data()
            
            coords = {
                'x': self.last_cmd_abs_x, 
                'y': self.last_cmd_abs_y, 
                'z': self.last_cmd_abs_z,
                'rot': self.last_cmd_abs_rot
            }
            self.message_queue.put(("MEASUREMENT_RESULT", (vals, coords)))
        except Exception as e:
            self.queue_message(f"Measurement Error: {e}", "ERROR")

    def _homing_verification_routine(self, auto_restart=False):
        """
        Executes a multi-phase verification to detect X/Y drift before re-homing.
        Must be called from a background thread (sender or manual).
        Raises InterruptedError if verification fails or stop is requested.
        """
        self.log_message("Starting Homing Verification...", "INFO")
        
        def send_and_wait_m119():
            with self.serial_engine.lock:
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
            with self.serial_engine.lock:
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
            # LOGIC: If the printer's coordinate system is accurate, it should NOT 
            # trigger the limit switches at X=0.5, Y=0.5.
            # DETECTS: "Negative Drift" (skipped steps toward home). If a switch is 
            # triggered at 0.5mm, the printer is physically closer to the switch 
            # than its internal counter believes.
            self.log_message("Verification Phase A: Moving to Safe Zone (0.5, 0.5)...")
            send_move_and_wait("G1 X0.5 Y0.5 F3000")
            send_move_and_wait("M400")
            states = send_and_wait_m119()
            
            if states.get('x_min') == 'triggered' or states.get('y_min') == 'triggered':
                axis = 'X' if states.get('x_min') == 'triggered' else 'Y'
                diag = f"CRITICAL ERROR: {axis}-axis triggered early at Safe Zone (0.5mm). Negative drift detected (printer is physically closer to home than it thinks)."
                if auto_restart: raise ValueError(diag)
                self._handle_homing_failure(diag)
                return

            # 2. Phase B: Move to Home Target (0, 0)
            # LOGIC: At X=0, Y=0, the limit switches MUST be triggered.
            # DETECTS: "Positive Drift" (skipped steps away from home). If the printer 
            # reaches its internal 0.0 but the switches are still 'open', the 
            # printer is physically further from home than its internal counter believes.
            self.log_message("Verification Phase B: Moving to Home Target (0, 0)...")
            send_move_and_wait("G1 X0 Y0 F500") 
            send_move_and_wait("M400")
            states = send_and_wait_m119()
            
            if states.get('x_min') == 'open' or states.get('y_min') == 'open':
                axis = 'X' if states.get('x_min') == 'open' else 'Y'
                diag = f"CRITICAL ERROR: {axis}-axis failed to trigger at Home Target (0.0mm). Positive drift detected (printer is physically further from home than it thinks)."
                if auto_restart: raise ValueError(diag)
                self._handle_homing_failure(diag)
                return

            # 3. Phase C: Actual Re-Home to sync coordinate system
            self.log_message("Verification Passed. Performing Sync Home (G28 X Y)...", "SUCCESS")
            send_move_and_wait("G28 X Y")
            self.log_message("Sync Home Complete. Resuming scan.")

        except Exception as e:
            if isinstance(e, InterruptedError): raise
            if auto_restart and isinstance(e, ValueError): raise
            diag = f"Verification Aborted: {str(e)}"
            if auto_restart: raise ValueError(diag)
            self._handle_homing_failure(diag)

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
    def _handle_homing_failure(self, error=None): self.log_message(f"Homing failed: {error}", "ERROR")
    def _apply_rot_conversion(self, cmd): return cmd
    def _parse_gcode_coords(self, line):
        coords = {}
        for axis in ['X', 'Y', 'Z', 'E']:
            match = re.search(rf"{axis}([-+]?\d*\.?\d+)", line, re.I)
            if match:
                key = 'rot' if axis == 'E' else axis.lower()
                coords[key] = float(match.group(1))
        return coords

    def _get_center(self):
        """Returns the marked center as (cx, cy, cz, crot), falling back to zeros."""
        try:
            return (float(self.center_x_var.get()), float(self.center_y_var.get()),
                    float(self.center_z_var.get()), float(self.center_rot_var.get()))
        except ValueError:
            return (0.0, 0.0, 0.0, 0.0)

    def _to_relative(self, coords):
        """Returns a copy of coords with the marked center subtracted from each axis present."""
        cx, cy, cz, crot = self._get_center()
        relative = coords.copy()
        if 'x' in relative: relative['x'] -= cx
        if 'y' in relative: relative['y'] -= cy
        if 'z' in relative: relative['z'] -= cz
        if 'rot' in relative: relative['rot'] -= crot
        return relative

    def _update_all_displays(self):
        """Central function to refresh all coordinate-based GUI elements."""
        mode = self.coord_mode.get()
        cx, cy, cz, crot = self._get_center()

        # Update TARGET DROs
        if mode == "relative":
            self.goto_x_display_var.set(f"{self.target_abs_x - cx:.2f}")
            self.goto_y_display_var.set(f"{self.target_abs_y - cy:.2f}")
            self.goto_z_display_var.set(f"{self.target_abs_z - cz:.2f}")
            self.goto_rot_display_var.set(f"{self.target_abs_rot - crot:.2f}")
        else:
            self.goto_x_display_var.set(f"{self.target_abs_x:.2f}")
            self.goto_y_display_var.set(f"{self.target_abs_y:.2f}")
            self.goto_z_display_var.set(f"{self.target_abs_z:.2f}")
            self.goto_rot_display_var.set(f"{self.target_abs_rot:.2f}")

        # Update CURRENT DROs
        def set_dro(var, abs_val, center):
            if abs_val is not None:
                val = abs_val - center if mode == "relative" else abs_val
                var.set(f"{val:.2f}")
            else: var.set("N/A")

        set_dro(self.last_cmd_x_display_var, self.last_cmd_abs_x, cx)
        set_dro(self.last_cmd_y_display_var, self.last_cmd_abs_y, cy)
        set_dro(self.last_cmd_z_display_var, self.last_cmd_abs_z, cz)
        set_dro(self.last_cmd_rot_display_var, self.last_cmd_abs_rot, crot)

        # Redraw canvases
        self._draw_xy_canvas_guides()
        self._draw_z_canvas_marker()
        self._draw_e_canvas_gauge()
        self._update_3d_position_marker()
        
        if self.last_cmd_abs_x is not None:
            self.footer_coords_var.set(f"X: {self.last_cmd_abs_x:.1f}  Y: {self.last_cmd_abs_y:.1f}  Z: {self.last_cmd_abs_z:.1f}")

    def _3d_control_bar(self): pass  # Assigned by ui_layout.py; not implemented here
    def _history_up(self, event=None): pass   # No-op: terminal history not implemented; bound by ui_layout.py
    def _history_down(self, event=None): pass  # No-op: terminal history not implemented; bound by ui_layout.py
    def _show_tutorial_popup(self):
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            tutorial_path = os.path.join(script_dir, "HINTS_TUTORIAL.txt")
            with open(tutorial_path, "r") as f:
                content = f.read()
        except Exception as e:
            messagebox.showerror("Error", f"Could not load tutorial: {e}")
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
    
    # --- INTERACTIVE CANVAS HANDLERS ---
    def _on_xy_canvas_click(self, event=None):
        if not event: return
        canvas_size = 215
        x_max, y_max = self.PRINTER_BOUNDS['x_max'], self.PRINTER_BOUNDS['y_max']
        machine_x = (event.x / canvas_size) * x_max
        machine_y = (1 - (event.y / canvas_size)) * y_max
        self.target_abs_x = max(0, min(x_max, machine_x))
        self.target_abs_y = max(0, min(y_max, machine_y))
        self._update_all_displays()

    def _on_z_canvas_click(self, event=None):
        if not event: return
        canvas_size = 215
        z_max = self.PRINTER_BOUNDS['z_max']
        machine_z = (1 - (event.y / canvas_size)) * z_max
        self.target_abs_z = max(0, min(z_max, machine_z))
        self._update_all_displays()

    def _on_e_canvas_click(self, event=None):
        if not event: return
        c = event.widget
        w, h = c.winfo_width(), c.winfo_height()
        pivot_y = 20  # matches cy in _draw_e_canvas_gauge
        dx = event.x - (w / 2)
        dy = event.y - pivot_y
        target_angle = math.degrees(math.atan2(dx, dy))
        self.target_abs_rot = max(-90, min(90, target_angle))
        self._update_all_displays()

    # --- MISSING DRAW CALLBACKS ---
    def _draw_xy_canvas_guides(self, event=None):
        c = getattr(self, 'xy_canvas', None)
        if not c: return
        c.delete("all")
        
        # Draw 2D toolpath
        self.plot_manager.draw_2d_toolpath(c)
        
        w, h = c.winfo_width(), c.winfo_height()
        if w < 10: w, h = 215, 215
        x_max, y_max = self.PRINTER_BOUNDS['x_max'], self.PRINTER_BOUNDS['y_max']
        c.create_rectangle(0, 0, w, h, outline=utils.COLOR_BORDER, dash=(2, 4))
        try:
            cx, cy = float(self.center_x_var.get()), float(self.center_y_var.get())
            cvx, cvy = (cx / x_max) * w, (1 - (cy / y_max)) * h
            c.create_line(cvx, 0, cvx, h, fill=utils.COLOR_TEXT_SECONDARY, dash=(4, 4))
            c.create_line(0, cvy, w, cvy, fill=utils.COLOR_TEXT_SECONDARY, dash=(4, 4))
        except (ValueError, tk.TclError): pass
        tvx, tvy = (self.target_abs_x / x_max) * w, (1 - (self.target_abs_y / y_max)) * h
        c.create_oval(tvx-5, tvy-5, tvx+5, tvy+5, outline=utils.COLOR_ACCENT_CYAN, width=2)
        c.create_line(tvx-8, tvy, tvx+8, tvy, fill=utils.COLOR_ACCENT_CYAN)
        c.create_line(tvx, tvy-8, tvx, tvy+8, fill=utils.COLOR_ACCENT_CYAN)
        if self.last_cmd_abs_x is not None and self.last_cmd_abs_y is not None:
            avx, avy = (self.last_cmd_abs_x / x_max) * w, (1 - (self.last_cmd_abs_y / y_max)) * h
            c.create_oval(avx-4, avy-4, avx+4, avy+4, fill=utils.COLOR_ACCENT_RED, outline="white")

    def _draw_z_canvas_marker(self, event=None):
        c = getattr(self, 'z_canvas', None)
        if not c: return
        c.delete("all")
        w, h = c.winfo_width(), c.winfo_height()
        if w < 10: w, h = 25, 215
        z_max = self.PRINTER_BOUNDS['z_max']
        for i in range(0, int(z_max) + 1, 20):
            if i == 0: continue
            y = (1 - (i / z_max)) * h
            c.create_line(0, y, 5, y, fill=utils.COLOR_TEXT_SECONDARY)
            c.create_text(10, y, text=str(i), fill=utils.COLOR_TEXT_SECONDARY, font=utils.FONT_BODY_SMALL, anchor="w")
        tvy = (1 - (self.target_abs_z / z_max)) * h
        c.create_polygon(0, tvy, 10, tvy-5, 10, tvy+5, fill=utils.COLOR_ACCENT_CYAN)
        if self.last_cmd_abs_z is not None:
            avy = (1 - (self.last_cmd_abs_z / z_max)) * h
            c.create_rectangle(0, avy-2, w, avy+2, fill=utils.COLOR_ACCENT_RED, outline="white")

    def _draw_e_canvas_gauge(self, event=None):
        c = getattr(self, 'e_canvas', None)
        if not c: return
        c.delete("all")
        w, h = c.winfo_width(), c.winfo_height()
        if w < 10: w, h = 215, 215
        cx = w / 2
        cy = 20
        radius = min(w / 2 - 15, h - cy - 10)
        c.create_arc(cx-radius, cy-radius, cx+radius, cy+radius, start=0, extent=-180, outline=utils.COLOR_BORDER, style=tk.ARC, width=2)
        for angle in range(-90, 91, 30):
            rad = math.radians(angle)
            x1, y1 = cx + (radius-5)*math.sin(rad), cy + (radius-5)*math.cos(rad)
            x2, y2 = cx + (radius+5)*math.sin(rad), cy + (radius+5)*math.cos(rad)
            c.create_line(x1, y1, x2, y2, fill=utils.COLOR_TEXT_SECONDARY)
        t_rad = math.radians(self.target_abs_rot)
        tx, ty = cx + radius*math.sin(t_rad), cy + radius*math.cos(t_rad)
        c.create_line(cx, cy, tx, ty, fill=utils.COLOR_ACCENT_CYAN, width=3, arrow=tk.LAST)
        if self.last_cmd_abs_rot is not None:
            a_rad = math.radians(self.last_cmd_abs_rot)
            ax, ay = cx + (radius-10)*math.sin(a_rad), cy + (radius-10)*math.cos(a_rad)
            c.create_line(cx, cy, ax, ay, fill=utils.COLOR_ACCENT_RED, width=2)

    def _is_3d_tab_active(self):
        try: return self.notebook.select() == str(self.display_tab)
        except Exception: return False
    def _on_tab_changed(self, event=None):
        if self._is_3d_tab_active(): self.root.after(50, self._draw_3d_toolpath)
    def _draw_3d_toolpath(self):
        if self._is_3d_tab_active(): self.plot_manager.draw_3d_toolpath()
    def _update_3d_position_marker(self):
        if self._is_3d_tab_active(): self.plot_manager.update_3d_position_marker()
    def _set_view(self, elev, azim): self.plot_manager.set_view(elev, azim)
    def _rotate_view(self, elev_change=0, azim_change=0): self.plot_manager.rotate_view(elev_change, azim_change)
    def _toggle_3d_plot_button(self): self.plot_manager.toggle_3d_plot()
    def _update_3d_plot_button_style(self): self.plot_manager.update_3d_plot_button_style()
    def _update_3d_plot_visibility(self): self.plot_manager.update_3d_plot_visibility()
    def _toggle_2d_plot_button(self): 
        self.is_2d_plot_enabled.set(not self.is_2d_plot_enabled.get())
        self._update_all_displays()

    def _update_section_borders(self):
        """Updates the LabelFrame border colour based on whether each section's setup is complete."""
        # 1. Connection Section
        if self.serial_engine.is_connected():
            self._conn_frame.configure(style='Green.TLabelframe')
        else:
            self._conn_frame.configure(style='Purple.TLabelframe')

        # Update Connection Buttons
        if hasattr(self, 'conn_home_button'):
            self.conn_home_button.config(style='GreenRing.TButton' if self.homing_performed else 'PurpleRing.TButton')

        # 2. Setup Section (File + Center)
        file_ok = bool(self.file_path_var.get() and self.file_path_var.get() != "No file selected")
        center_ok = self.center_marked
        if file_ok and center_ok:
            self._setup_frame.configure(style='Green.TLabelframe')
        else:
            self._setup_frame.configure(style='Purple.TLabelframe')

        # Update Setup Buttons
        if hasattr(self, 'select_gcode_button'):
            self.select_gcode_button.config(style='GreenRing.TButton' if file_ok else 'PurpleRing.TButton')
        if hasattr(self, 'mark_center_button'):
            self.mark_center_button.config(style='GreenRing.TButton' if center_ok else 'PurpleRing.TButton')
        if hasattr(self, 'collision_test_button'):
            self.collision_test_button.config(style='GreenRing.TButton' if self.rotation_crash_test_complete else 'PurpleRing.TButton')

        # 3. Measurement Section
        if hasattr(self, '_meas_frame'):
            self._meas_frame._validate_state()

        # 4. Execution Control Section
        if self.is_sending:
            self._ctrl_frame.configure(style='Cyan.TLabelframe')
        elif file_ok and center_ok:
            self._ctrl_frame.configure(style='Green.TLabelframe')
            self.start_button.config(state=tk.NORMAL)
        else:
            self._ctrl_frame.configure(style='Purple.TLabelframe')
            self.start_button.config(state=tk.DISABLED)

    def _send_from_terminal(self, event=None):
        """Sends a custom G-code command from the terminal input field."""
        cmd = self.terminal_input.get().strip()
        if not cmd: return
        self.terminal_input.delete(0, tk.END)
        self.log_message(f"> {cmd}", "SENT")
        self._send_manual_command(cmd)

    def queue_message(self, message, level="INFO"):
        self.message_queue.put(("LOG", (message, level)))

    # --- CORE LOGIC & MESSAGE POLLING ---

    def check_message_queue(self):
        try:
            while True:
                msg_type, content = self.message_queue.get_nowait()
                self._handle_queue_message(msg_type, content)
        except queue.Empty: pass
        finally: self.after_id = self.root.after(100, self.check_message_queue)

    def _handle_queue_message(self, msg_type, content):
        """Dispatches incoming queue messages to their respective handlers."""
        dispatch = {
            "LOG": lambda c: self.log_message(c[0], c[1]),
            "PROGRESS": self._handle_progress_msg,
            "POSITION_UPDATE": self._handle_position_update,
            "PATH_PROGRESS_UPDATE": self._handle_path_progress,
            "CONNECTED": self._on_connected,
            "CONNECT_FAIL": self._handle_connect_fail,
            "FILE_SEND_FINISHED": lambda _: self._on_send_finished(),
            "SET_STATUS": lambda c: self.header_status_indicator.set_status(c),
            "MEASUREMENT_RESULT": lambda c: self._on_measurement(*c),
            "HOMING_FAILURE": self._handle_homing_failure,
            "DMM_CONNECTED": lambda _: self._handle_dmm_connected(),
            "DMM_FAIL": self._handle_dmm_fail,
            "MANUAL_FINISHED": self._handle_manual_finished,
            "FORCE_PAUSE": lambda _: self._handle_force_pause()
        }
        
        handler = dispatch.get(msg_type)
        if handler:
            try:
                handler(content)
            except Exception as e:
                self.log_message(f"Handler Error ({msg_type}): {e}", "ERROR")

    def _handle_progress_msg(self, content):
        self.progress_var.set(content[2])
        self.progress_label_var.set(f"{content[0]}/{content[1]} lines")

    def _handle_position_update(self, content):
        self.last_cmd_abs_x = content.get('x', self.last_cmd_abs_x)
        self.last_cmd_abs_y = content.get('y', self.last_cmd_abs_y)
        self.last_cmd_abs_z = content.get('z', self.last_cmd_abs_z)
        self.last_cmd_abs_rot = content.get('rot', self.last_cmd_abs_rot)
        self._update_all_displays()
        
        if self.event_bus:
            # Calculate relative position for PoP visualization
            relative_pos = self._to_relative(content)
            self.event_bus.publish("scan_position", relative_pos)

    def _handle_path_progress(self, content):
        self.completed_move_count = content
        self._update_all_displays()

    def _handle_connect_fail(self, content):
        self.disconnect_printer(silent=True)
        messagebox.showerror("Error", content)

    def _handle_dmm_connected(self):
        self.is_dmm_connected = True
        self.dmm_status_var.set("DMMs: Connected")
        self.log_message("DMMs connected successfully.", "SUCCESS")
        self._update_section_borders()

    def _handle_dmm_fail(self, content):
        self.is_dmm_connected = False
        self.dmm_status_var.set(f"DMMs: Error ({content})")
        if hasattr(self, 'dmm_connect_button'):
            self.dmm_connect_button.config(state=tk.NORMAL)
        self._update_section_borders()

    def _handle_manual_finished(self, is_homing):
        if is_homing:
            self.homing_performed = True
        self._update_section_borders()

    def _handle_force_pause(self):
        """Forces the UI into a paused state when triggered by the hardware engine."""
        self.is_paused = True
        self.pause_resume_button.config(text="Resume")
        self.restart_scan_button.config(state=tk.NORMAL if self.processed_gcode else tk.DISABLED)
        self._set_manual_controls_state(tk.NORMAL)
        self._set_goto_controls_state(tk.NORMAL)
        self._set_terminal_controls_state(tk.NORMAL)
        self.header_status_indicator.set_status("busy")
        self.log_message("Scan execution has been gated and paused due to measurement failure.", "ERROR")
        self._update_section_borders()

    def _on_connected(self, data):
        self.serial_connection, port, baud, pos = data
        self.log_message(f"Connected to {port}!", "SUCCESS")
        self.header_status_indicator.set_status("on")
        self.connection_status_var.set(f"Connected: {port}")
        self.footer_status_var.set(f"{port} @ {baud}")
        self.connect_button.config(text="Disconnect", style='GreenRing.TButton')
        if pos: 
            self.last_cmd_abs_x, self.last_cmd_abs_y, self.last_cmd_abs_z = pos['x'], pos['y'], pos['z']
            self.last_cmd_abs_rot = pos.get('rot', 0.0)
        self._set_manual_controls_state(tk.NORMAL); self._set_terminal_controls_state(tk.NORMAL); self._set_goto_controls_state(tk.NORMAL)
        if self.processed_gcode: self.start_button.config(state=tk.NORMAL)
        self._update_all_displays(); self._update_section_borders()

    def _on_send_finished(self):
        self.is_sending = False
        self.start_button.config(state=tk.NORMAL if self.processed_gcode else tk.DISABLED)
        self.restart_scan_button.config(state=tk.NORMAL if self.processed_gcode else tk.DISABLED)
        self.header_status_indicator.set_status("on")
        self.progress_var.set(100.0)
        if self.event_bus:
            self.event_bus.publish("scan_finished", {})
        self._show_scan_complete_popup()

    def _on_measurement(self, vals, coords):
        if not vals:
            self.last_measurement_var.set("Val: --")
            return

        display = f"Val: {vals[0]:.4f}"
        if len(vals) > 1: display += f" | {vals[1]:.4f}"
        self.last_measurement_var.set(display)
        self.log_message(f"Measured: {', '.join([f'{v:.4f}' for v in vals])}")

        if self.log_measurements_enabled.get():
            # Transform to relative coordinates for logging
            rel_coords = self._to_relative(coords)

            # Atomic Pinning: Publish high-fidelity gated data for Visualizers
            if self.event_bus:
                # Merge vals into coords with correct keys for the visualizer
                gated_data = rel_coords.copy()
                if self.active_data_source == "DMM":
                    for i, v in enumerate(vals): gated_data[f"DMM_{i+1}"] = v
                else:
                    for i, mode in enumerate(self.active_hub_modes):
                        if i < len(vals): gated_data[mode] = vals[i]
                
                self.event_bus.publish("gated_measurement", gated_data)

            self._log_measurement_to_file(vals, rel_coords)

    def _current_dv_keys(self):
        """Returns the dependent-variable column keys for the active data source."""
        if self.active_data_source == "DMM":
            return [f"DMM_{i+1}" for i in range(len(DMM_CONFIG))]
        dv_keys = self.active_hub_modes.copy()
        if not dv_keys: dv_keys = ["V_rec"]
        return dv_keys

    def _log_measurement_to_file(self, vals, coords):
        filepath = self.log_filepath_var.get()
        if not filepath:
            if self.log_measurements_enabled.get():
                # Default to a file in data/
                if not os.path.exists(_DATA_DIR): os.makedirs(_DATA_DIR)
                filepath = os.path.join(_DATA_DIR, f"log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")
                self.log_filepath_var.set(filepath)
                self._initialize_log_file(filepath)
            else:
                return
        
        try:
            # Prepare fieldnames - use consistent IVs
            iv_keys = ['x', 'y', 'z', 'rot']
            dv_keys = self._current_dv_keys()
            fieldnames = ['Timestamp'] + iv_keys + dv_keys
            
            if not os.path.exists(filepath):
                self._initialize_log_file(filepath)
            
            # Retry loop for Permission denied (Errno 13) - usually Excel or rapid hits
            max_retries = 3
            for attempt in range(max_retries):
                try:
                    with open(filepath, 'a', newline='') as f:
                        writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction='ignore')
                        
                        row = {
                            'Timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
                        }
                        # Ensure coords only contains x, y, z, rot for consistency with header
                        row.update({k: coords.get(k, 0.0) for k in iv_keys})
                        for i, k in enumerate(dv_keys):
                            row[k] = vals[i] if i < len(vals) else 0.0
                        writer.writerow(row)
                    
                    # Log success to GUI to confirm writing is active
                    self.log_message(f"Data point logged: {', '.join([f'{v:.3f}' for v in vals])}", "SUCCESS")
                    return # Success
                except PermissionError as e:
                    if attempt < max_retries - 1:
                        time.sleep(0.2 * (attempt + 1)) # More aggressive backoff
                        continue
                    raise e
        except Exception as e:
            self.log_message(f"Logging Error: {e}", "ERROR")

    def log_message(self, message, level="INFO"):
        if not hasattr(self, 'log_area'): return
        self.log_area.config(state=tk.NORMAL)
        ts = time.strftime("%H:%M:%S")
        tag = level.upper()
        self.log_area.insert(tk.END, f"[{ts}] ", "timestamp")
        self.log_area.insert(tk.END, f"{message}\n", tag)
        self.log_area.tag_configure(tag, foreground=self._get_level_color(level))
        if level in ["ERROR", "CRITICAL"]: self.log_area.tag_configure(tag, font=(utils.FONT_TERMINAL[0], utils.FONT_TERMINAL[1], 'bold'))
        self.log_area.see(tk.END); self.log_area.config(state=tk.DISABLED)

    def _get_level_color(self, level):
        return {"INFO": utils.COLOR_ACCENT_GREEN, "SUCCESS": utils.COLOR_ACCENT_CYAN, 
                "WARN": utils.COLOR_ACCENT_AMBER, "ERROR": utils.COLOR_ACCENT_RED,
                "CRITICAL": utils.COLOR_ACCENT_RED}.get(level, "white")

    # --- FILE HANDLING ---
    def select_file(self):
        path = filedialog.askopenfilename(initialdir=_DATA_DIR, filetypes=[("G-code", "*.gcode"), ("All", "*.*")])
        if path:
            self.load_gcode_file(path)
            if self.serial_connection and self.processed_gcode: self.start_button.config(state=tk.NORMAL)

    def clear_file(self):
        self.gcode_filepath = None; self.processed_gcode = []
        self.toolpath_by_layer = {}; self.move_to_layer_map = []; self.ordered_z_values = []; self.completed_move_count = 0
        self.file_path_var.set("No file selected"); self.header_file_var.set("NONE")
        self.start_button.config(state=tk.DISABLED); self.plot_manager.invalidate_cache()
        self.rotation_crash_test_complete = False; self._update_all_displays(); self._update_section_borders()
        self.log_message("Cleared loaded G-code program.")

    def load_gcode_file(self, path):
        self.gcode_filepath = path; self.log_message(f"Loading from {path}...")
        if self.process_gcode() and self.processed_gcode:
            # Populate the SETUP pane labels here (not just in select_file) so the
            # Pattern Generator's "Load to Sender" path, which calls this directly,
            # also reflects the loaded file instead of showing "No file selected".
            self.file_path_var.set(path); self.header_file_var.set(os.path.basename(path).upper())
            self.log_message(f"File loaded: {len(self.processed_gcode)} total steps identified.", "SUCCESS")
        else:
            # Load produced no usable program (out of bounds, empty, or parse error).
            # Don't claim success: reset the SETUP pane and surface the failure.
            self.file_path_var.set("No file selected"); self.header_file_var.set("NO FILE")
            self.log_message("Load failed: no valid moves found (see errors above).", "ERROR")
            messagebox.showerror("Load Failed",
                                 "The G-code could not be loaded.\n\n"
                                 "No valid moves were found — the pattern may be out of "
                                 "printer bounds. See the log for the offending line.")
        self._update_section_borders()

    def process_gcode(self):
        """Parse the loaded file into processed_gcode. Returns True on success,
        False if the file is missing, out of bounds, or raises while parsing."""
        self.processed_gcode = []; self.toolpath_by_layer = {}; self.move_to_layer_map = []; self.ordered_z_values = []; self.completed_move_count = 0
        self.plot_manager.invalidate_cache()
        if not self.gcode_filepath: return False
        try:
            with open(self.gcode_filepath, 'r') as f:
                lines = f.readlines()
            if not self._build_toolpath_map(lines):
                return False
            if self._detect_rotation_test_skip(lines):
                self.rotation_crash_test_complete = True

            mgr = SequenceManager()
            parsed_steps = mgr.parse_gcode_file(self.gcode_filepath)
            self.processed_gcode = self._process_mixed_sequence(parsed_steps)

            self._load_profile_metadata(lines)

            self._update_all_displays(); self.plot_manager.draw_3d_toolpath()
            return True
        except Exception as e:
            self.log_message(f"Process Error: {e}", "ERROR")
            return False

    def _build_toolpath_map(self, lines):
        """Walks the G-code lines building the layer toolpath map with bounds validation.

        Returns False if a bounds error is hit (caller should abort), True otherwise.
        """
        cx, cy, cz = float(self.center_x_var.get()), float(self.center_y_var.get()), float(self.center_z_var.get())
        crot = float(self.center_rot_var.get())
        current_pos = {'x': None, 'y': None, 'z': None, 'rot': None}; layer_move_counters = {}
        for line in lines:
            stripped = line.split(';')[0].strip()
            if not stripped: continue
            if "G92" in stripped.upper(): continue
            if "M104" in stripped.upper() or "M109" in stripped.upper(): continue
            is_move = "G0" in stripped.upper() or "G1" in stripped.upper() or "G28" in stripped.upper()
            if is_move:
                start_pos = current_pos.copy()
                if "G28" in stripped.upper():
                    if 'X' not in stripped.upper() and 'Y' not in stripped.upper() and 'Z' not in stripped.upper():
                        current_pos = {'x': self.PRINTER_BOUNDS['x_min'], 'y': self.PRINTER_BOUNDS['y_min'], 'z': self.PRINTER_BOUNDS['z_min'], 'rot': 0.0}
                    else:
                        if 'X' in stripped.upper(): current_pos['x'] = self.PRINTER_BOUNDS['x_min']
                        if 'Y' in stripped.upper(): current_pos['y'] = self.PRINTER_BOUNDS['y_min']
                        if 'Z' in stripped.upper(): current_pos['z'] = self.PRINTER_BOUNDS['z_min']
                else:
                    parsed = self._parse_gcode_coords(stripped)
                    if not parsed: continue
                    rel_x, rel_y, rel_z, rel_rot = parsed.get('x'), parsed.get('y'), parsed.get('z'), parsed.get('rot')
                    current_pos = {
                        'x': (rel_x + cx) if rel_x is not None else current_pos.get('x'),
                        'y': (rel_y + cy) if rel_y is not None else current_pos.get('y'),
                        'z': (rel_z + cz) if rel_z is not None else current_pos.get('z'),
                        'rot': (rel_rot + crot) if rel_rot is not None else current_pos.get('rot')
                    }
                    if any(v is not None and not (self.PRINTER_BOUNDS[k+'_min'] <= v <= self.PRINTER_BOUNDS[k+'_max']) for k,v in current_pos.items() if v is not None):
                        self.log_message(f"Bounds Error on line: {stripped}", "ERROR"); return False
                if start_pos['x'] is not None and current_pos['x'] is not None:
                    z = current_pos['z']; self.toolpath_by_layer.setdefault(z, []).append(((start_pos['x'], start_pos['y']), (current_pos['x'], current_pos['y'])))
                    idx = layer_move_counters.get(z, 0); self.move_to_layer_map.append((z, idx)); self.ordered_z_values.append(z); layer_move_counters[z] = idx + 1
        return True

    def _detect_rotation_test_skip(self, lines):
        """Returns True when no extrusion (E) commands are present, allowing the rotation crash test to be skipped."""
        return not any('E' in l.upper() for l in lines)

    def _load_profile_metadata(self, lines):
        """Extracts the PARAMS profile metadata from the header and updates the PoP panel's fixed bounds."""
        try:
            for line in lines[:20]:
                if line.startswith("; PARAMS:"):
                    profile_json = line.replace("; PARAMS:", "").strip()
                    profile_data = json.loads(profile_json)
                    self.current_profile_data = profile_data # Persist for log file initialization

                    if hasattr(self, 'pop_viz_panel'):
                        try:
                            self.pop_viz_panel.update_fixed_bounds(
                                self._profile_to_relative_bounds(profile_data))
                        except Exception as e:
                            self.log_message(f"Scaling Meta Error: {e}", "WARN")
                    break
        except Exception: pass

    def _process_mixed_sequence(self, steps):
        translated = []
        try:
            cx, cy, cz = float(self.center_x_var.get()), float(self.center_y_var.get()), float(self.center_z_var.get())
            current_pos = {'x': None, 'y': None, 'z': None}
            for s in steps:
                if isinstance(s, dict): translated.append(s); continue
                line = s.strip()
                if not line or line.startswith(';'): translated.append(line); continue
                if "G0" in line.upper() or "G1" in line.upper():
                    p = self._parse_gcode_coords(line)
                    rel_x, rel_y, rel_z = p.get('x'), p.get('y'), p.get('z')
                    abs_x = current_pos.get('x') if rel_x is None else rel_x + cx
                    abs_y = current_pos.get('y') if rel_y is None else rel_y + cy
                    abs_z = current_pos.get('z') if rel_z is None else rel_z + cz
                    current_pos.update({k:v for k,v in {'x':abs_x, 'y':abs_y, 'z':abs_z}.items() if v is not None})
                    parts = [line.split()[0]]
                    if rel_x is not None: parts.append(f"X{abs_x:.3f}")
                    if rel_y is not None: parts.append(f"Y{abs_y:.3f}")
                    if rel_z is not None: parts.append(f"Z{abs_z:.3f}")
                    e_m = re.search(r"E([-+]?\d*\.?\d+)", line, re.I); f_m = re.search(r"F(\d+)", line, re.I)
                    if e_m: parts.append(e_m.group(0))
                    if f_m: parts.append(f_m.group(0))
                    translated.append(" ".join(parts))
                else: translated.append(line)
        except Exception: return steps
        return translated

    # --- CONNECTION HANDLING ---
    def toggle_connection(self):
        if self.is_disconnecting: return
        if self.serial_connection:
            self.disconnect_printer()
        elif self.connect_button.cget("text") == "Connecting...":
            self._cancel_connection_attempt()
        else:
            self.connect_printer()

    def connect_printer(self):
        selected_port = self.port_var.get()
        try: baudrate = int(self.baud_var.get())
        except (ValueError, TypeError): messagebox.showerror("Error", "Invalid Baud Rate"); return
        self.log_message(f"Connecting... Port: {selected_port}, Baud: {baudrate}...")
        self.connect_button.config(text="Connecting...", state=tk.DISABLED)
        self.cancel_connect_button.grid(); self.cancel_connect_button.config(state=tk.NORMAL)
        self.port_combobox.config(state=tk.DISABLED); self.baud_entry.config(state=tk.DISABLED)
        self.connection_status_var.set("Connecting..."); self.header_status_indicator.set_status("busy")
        self._set_terminal_controls_state(tk.DISABLED); self.cancel_connect_event.clear()
        self.serial_engine.connect(selected_port, baudrate, self.cancel_connect_event)

    def _cancel_connection_attempt(self):
        self.log_message("Connection attempt cancelled.", "WARN"); self.cancel_connect_event.set()
        self.cancel_connect_button.grid_remove(); self.connect_button.config(text="Connect", state=tk.NORMAL)
        self.port_combobox.config(state="readonly"); self.baud_entry.config(state=tk.NORMAL)
        self.connection_status_var.set("Cancelling..."); self.header_status_indicator.set_status("busy")

    def disconnect_printer(self, silent=False):
        self.is_disconnecting = True
        try:
            if not silent and (self.is_sending or self.is_manual_command_running): messagebox.showwarning("Busy", "Please stop operation before disconnecting"); return
            if silent: self.stop_event.set(); self.is_sending = self.is_manual_command_running = False
            if self.serial_connection:
                try: self.serial_connection.close()
                except Exception: pass
            self.serial_connection = None; self.connection_status_var.set("Disconnected"); self.header_status_indicator.set_status("off")
            self.connect_button.config(text="Connect", style='PurpleRing.TButton', state=tk.NORMAL)
            self.port_combobox.config(state="readonly"); self.baud_entry.config(state=tk.NORMAL)
            self._set_manual_controls_state(tk.DISABLED); self._set_goto_controls_state(tk.DISABLED); self._set_terminal_controls_state(tk.DISABLED)
            self.progress_var.set(0.0); self.progress_label_var.set("Progress: Idle"); self.last_cmd_abs_x = self.last_cmd_abs_y = self.last_cmd_abs_z = None
            self.homing_performed = self.center_marked = self.rotation_crash_test_complete = False
            self._update_all_displays(); self._update_section_borders()
            self._enforce_id = self.root.after(350, self._enforce_disconnect_state)
        finally:
            self.is_disconnecting = False

    def _enforce_disconnect_state(self):
        if self.serial_connection: return
        self.header_status_indicator.set_status("off"); self._update_section_borders()
        self.connect_button.config(text="Connect", style='PurpleRing.TButton')

    def rescan_ports(self):
        curr = self.port_var.get()
        found_ports = self._get_available_ports()
        self.available_ports = ["Auto-detect"] + found_ports
        
        self.port_combobox['values'] = self.available_ports
        
        if curr in self.available_ports:
            self.port_var.set(curr)
        elif self.available_ports:
            self.port_var.set(self.available_ports[0])
        else:
            self.port_var.set("") # Nothing to set

    def _get_available_ports(self):
        ports = serial.tools.list_ports.comports()
        filtered = [p.device for p in ports if any(x in p.description for x in ['CH340', 'USB Serial', 'Arduino', 'Serial Port']) or not p.description]
        sim_port = os.environ.get("SEED_SIM_PORT")
        if sim_port and sim_port not in filtered:
            filtered.append(sim_port)
        return sorted(filtered)

    # --- EXECUTION CONTROL ---
    def start_sending(self):
        if not self.serial_connection: messagebox.showerror("Error", "Not connected"); return
        self.process_gcode() 
        if not self.processed_gcode: messagebox.showerror("Error", "No G-code to send"); return
        if not self.rotation_crash_test_complete:
            if not messagebox.askyesno("Safety Check", "Collision test not completed. Proceed?"): return
        if not self.log_measurements_enabled.get() or not self.log_filepath_var.get().strip():
            if not messagebox.askyesno("Logging Not Active", "Measurement data will NOT be saved. Proceed?"): return
        if self.is_sending or self.is_manual_command_running: return
        
        # Calculate total lines, handling both strings and dictionaries
        self.total_lines_to_send = 0
        for item in self.processed_gcode:
            if isinstance(item, str):
                if item.strip() and not item.strip().startswith(';'):
                    self.total_lines_to_send += 1
            else:
                self.total_lines_to_send += 1

        self.progress_var.set(0.0); self.progress_label_var.set(f"0/{self.total_lines_to_send} lines")
        self.is_sending = True; self.is_paused = False; self.stop_event.clear(); self.pause_event.set()
        self.start_button.config(state=tk.DISABLED); self.pause_resume_button.config(text="Pause", state=tk.NORMAL)
        self.restart_scan_button.config(state=tk.DISABLED)
        self._set_manual_controls_state(tk.DISABLED); self._set_goto_controls_state(tk.DISABLED); self._set_terminal_controls_state(tk.DISABLED)
        self.log_message("Starting G-code stream...")
        if self.event_bus:
            profile = getattr(self, 'current_profile_data', None) or {}
            self.event_bus.publish("scan_started", self._profile_to_relative_bounds(profile))
        self.serial_engine.send_gcode(list(self.processed_gcode), {'x':self.last_cmd_abs_x, 'y':self.last_cmd_abs_y, 'z':self.last_cmd_abs_z, 'rot':self.last_cmd_abs_rot}, self.PRINTER_BOUNDS, self.auto_measure_enabled.get())

    def restart_scan(self):
        self.progress_var.set(0.0)
        self.progress_label_var.set("Progress: Idle")
        filepath = self.log_filepath_var.get().strip()
        if self.log_measurements_enabled.get() and filepath:
            select_new = messagebox.askyesno(
                "Restart Scan — Select New Data File",
                "Select a new output file to preserve your existing data.\n\n"
                "Proceeding without a new file will permanently clear the current file.\n\n"
                "Select a new output file now?",
                icon='warning'
            )
            if select_new:
                old_path = filepath
                self.select_log_file()
                if self.log_filepath_var.get().strip() == old_path:
                    self.progress_label_var.set("Progress: Idle")
                    return
            else:
                self._initialize_log_file(filepath)
        self.start_sending()

    def _profile_to_relative_bounds(self, profile):
        """Return {ax_min, ax_max, ...} in relative-to-center coords from a PARAMS profile.
        When symmetric mode is enabled the actual scan extents come from *_offset, not
        from the x_min/x_max widget values (which stay at their default -50/50 when the
        symmetric toggle hides them)."""
        bounds = {}
        for ax in ('x', 'y', 'z', 'rot'):
            try:
                if profile.get(f'{ax}_symmetric'):
                    off = abs(float(profile[f'{ax}_offset']))
                    bounds[f'{ax}_min'] = -off
                    bounds[f'{ax}_max'] = off
                else:
                    lo = profile.get(f'{ax}_min')
                    hi = profile.get(f'{ax}_max')
                    if lo is not None and hi is not None:
                        bounds[f'{ax}_min'] = float(lo)
                        bounds[f'{ax}_max'] = float(hi)
            except (KeyError, TypeError, ValueError):
                pass
        return bounds

    def toggle_pause_resume(self):
        if not self.is_sending and not self.is_manual_command_running: return
        if self.is_paused:
            self.pause_event.set(); self.is_paused = False; self.pause_resume_button.config(text="Pause")
            self.restart_scan_button.config(state=tk.DISABLED)
            self._set_manual_controls_state(tk.DISABLED); self._set_goto_controls_state(tk.DISABLED); self._set_terminal_controls_state(tk.DISABLED)
            self.header_status_indicator.set_status("on")
        else:
            self.pause_event.clear(); self.is_paused = True; self.pause_resume_button.config(text="Resume")
            self.restart_scan_button.config(state=tk.NORMAL if self.processed_gcode else tk.DISABLED)
            self._set_manual_controls_state(tk.NORMAL); self._set_goto_controls_state(tk.NORMAL); self._set_terminal_controls_state(tk.NORMAL)
            self.header_status_indicator.set_status("busy")

    def emergency_stop(self):
        if not self.serial_connection: self._reset_gui_after_stop(); return
        self.log_message("!!! EMERGENCY STOP !!!", "CRITICAL")
        self.rotation_crash_test_complete = False; self.pause_event.set(); self.stop_event.set()
        try:
            if self.serial_connection:
                self.serial_connection.reset_output_buffer()
            self.serial_engine.write_command('M112')
        except Exception: pass
        self.header_status_indicator.set_status("error"); self.disconnect_printer(silent=True)
        messagebox.showwarning("Emergency Stop", "M112 sent.\nPrinter requires reset.\nConnection closed."); self._reset_gui_after_stop(reset_position=True)

    def quick_stop(self):
        if not self.serial_connection: self._reset_gui_after_stop(False); return
        self.log_message("Quick Stop requested.", "WARN"); self.rotation_crash_test_complete = False
        self.pause_event.set(); self.stop_event.set()
        try:
            self.serial_connection.write(b'\nM410\n'); self.serial_connection.flush()
        except Exception: pass
        self.header_status_indicator.set_status("busy"); messagebox.showinfo("Quick Stop", "M410 sent. Manual controls enabled."); self._reset_gui_after_stop(reset_position=False)

    def _reset_gui_after_stop(self, reset_position=True):
        self.is_sending = self.is_paused = self.is_manual_command_running = False
        state = tk.NORMAL if self.serial_connection else tk.DISABLED
        self._set_manual_controls_state(state); self._set_goto_controls_state(state); self._set_terminal_controls_state(state)
        self.start_button.config(state=tk.DISABLED); self.pause_resume_button.config(text="Pause", state=tk.DISABLED)
        self.restart_scan_button.config(state=tk.NORMAL if self.processed_gcode else tk.DISABLED)
        self.progress_var.set(0.0); self.progress_label_var.set("Progress: Stopped")
        if reset_position: self.last_cmd_abs_x = self.last_cmd_abs_y = self.last_cmd_abs_z = self.last_cmd_abs_rot = 0.0
        self._update_all_displays()

    # --- COLLISION TEST ---
    def _open_collision_test_screen(self):
        """Swaps the main UI for the Collision Avoidance Test screen."""
        if not self.serial_connection:
            messagebox.showerror("Error", "Not connected to printer.")
            return
        if not self.processed_gcode:
            messagebox.showerror("Setup Error", "Error: Must load test profile before the test")
            return
        self.main_view_frame.pack_forget()
        self.test_view_frame = tk.Frame(self.parent, bg=utils.COLOR_BG)
        self.test_view_frame.pack(fill=tk.BOTH, expand=True)
        warning_frame = tk.Frame(self.test_view_frame, bg=utils.COLOR_ACCENT_AMBER, padx=20, pady=20)
        warning_frame.pack(fill=tk.X, padx=50, pady=(50, 20))
        tk.Label(warning_frame, text="⚠ CAUTION ⚠", font=("Inter", 24, "bold"), bg=utils.COLOR_ACCENT_AMBER, fg=utils.COLOR_BLACK).pack()
        tk.Label(warning_frame, text="Remove hub and implant to avoid damage to equipment.", font=("Inter", 16), bg=utils.COLOR_ACCENT_AMBER, fg=utils.COLOR_BLACK).pack(pady=(10, 0))
        ctrl_frame = tk.Frame(self.test_view_frame, bg=utils.COLOR_BG)
        ctrl_frame.pack(expand=True)
        self.btn_begin_test = tk.Button(ctrl_frame, text="BEGIN TEST", font=("Orbitron", 18, "bold"), bg=utils.COLOR_ACCENT_CYAN, fg=utils.COLOR_BLACK, activebackground="#00eaff", activeforeground=utils.COLOR_BLACK, relief=tk.RAISED, bd=3, padx=30, pady=15, command=self._start_collision_test)
        self.btn_begin_test.pack(pady=20)
        self.btn_stop_test = tk.Button(ctrl_frame, text="STOP MOTION", font=("Orbitron", 24, "bold"), bg=utils.COLOR_ACCENT_RED, fg="white", activebackground="#ff6666", activeforeground="white", relief=tk.RAISED, bd=5, padx=50, pady=30, command=self._stop_collision_test)
        self.btn_stop_test.pack(pady=20)
        self.lbl_test_status = tk.Label(ctrl_frame, text="Ready", font=("Inter", 14), bg=utils.COLOR_BG, fg=utils.COLOR_TEXT_SECONDARY)
        self.lbl_test_status.pack(pady=10)
        self.btn_exit_test = tk.Button(self.test_view_frame, text="Exit Test", font=("Inter", 12), bg=utils.COLOR_PANEL_BG, fg=utils.COLOR_TEXT_PRIMARY, relief=tk.FLAT, padx=20, pady=10, command=self._close_collision_test_screen)
        self.btn_exit_test.pack(side=tk.BOTTOM, pady=30)

    def _close_collision_test_screen(self):
        if hasattr(self, 'test_view_frame') and self.test_view_frame: self.test_view_frame.destroy()
        self.main_view_frame.pack(fill=tk.BOTH, expand=True)

    def _stop_collision_test(self):
        self.log_message("!!! COLLISION TEST ABORTED — Quick Stop !!!", "WARN")
        self.is_collision_test_running = False
        self.rotation_crash_test_complete = False
        if hasattr(self, 'lbl_test_status') and self.lbl_test_status.winfo_exists():
            self.lbl_test_status.config(text="Test Aborted - Re-center the axis and try again.", fg=utils.COLOR_ACCENT_AMBER, font=("Rajdhani", 12, "bold"), wraplength=400)
        self.quick_stop()
        self.is_manual_command_running = False
        self.root.after(0, self._update_all_displays)
        self.root.after(0, self._update_section_borders)
        self.stop_event.clear()

    def _start_collision_test(self):
        if self.is_sending or self.is_manual_command_running: return
        self.btn_begin_test.config(state=tk.DISABLED); self.btn_exit_test.config(state=tk.DISABLED)
        self.lbl_test_status.config(text="Moving to Center...", fg=utils.COLOR_ACCENT_CYAN)
        self.stop_event.clear(); self.pause_event.set()
        threading.Thread(target=self._collision_test_worker, daemon=True).start()

    def _collision_test_worker(self):
        try:
            def update_status(msg):
                self.queue_message(msg, "INFO")
                if hasattr(self, 'lbl_test_status') and self.lbl_test_status.winfo_exists():
                    self.root.after(0, lambda: self.lbl_test_status.config(text=msg))
            self.is_manual_command_running = True; self.is_collision_test_running = True; self.root.after(0, self._update_all_displays)
            time.sleep(1.0)
            if self.serial_engine.is_connected():
                with self.serial_engine.lock:
                    self.serial_connection.reset_input_buffer()
            min_rot = max_rot = 0.0; pause_time_ms = 0; found_rot = False
            for line in self.processed_gcode:
                if 'E' in line.upper():
                    match = re.search(r"E([-+]?\d*\.?\d+)", line)
                    if match:
                        val = float(match.group(1))
                        if not found_rot: min_rot = max_rot = val; found_rot = True
                        else:
                            if val < min_rot: min_rot = val
                            if val > max_rot: max_rot = val
                if 'G4' in line.upper() and 'P' in line.upper() and pause_time_ms == 0:
                    match = re.search(r"P(\d+)", line)
                    if match: pause_time_ms = int(match.group(1))
            if not found_rot: update_status("No rotation (R) found in G-code. Skipping rotation test."); self.rotation_crash_test_complete = True
            else: update_status(f"Test Range: R{min_rot:.1f}° to R{max_rot:.1f}°")
            try:
                cx, cy, cz = float(self.center_x_var.get()), float(self.center_y_var.get()), float(self.center_z_var.get())
            except ValueError:
                self.queue_message("Invalid Center Coordinates!", "ERROR")
                if hasattr(self, 'lbl_test_status') and self.lbl_test_status.winfo_exists():
                    self.root.after(0, lambda: self.lbl_test_status.config(text="Test Failed: Center coordinates not marked.", fg=utils.COLOR_ACCENT_RED))
                return
            min_z = min(self.ordered_z_values) if self.ordered_z_values else cz
            speed_xy, speed_rot = 1000, 500
            def send_wait(cmd_str, timeout_s=30):
                lines = [c.strip() for c in cmd_str.split('\n') if c.strip()]
                for c in lines:
                    self.serial_engine.write_command(c)
                    start = time.time(); buffer = ""; ok_received = False
                    while time.time() - start < timeout_s:
                        if self.stop_event.is_set(): raise InterruptedError("Stop")
                        if self.serial_connection.in_waiting > 0:
                            buffer += self.serial_connection.read(self.serial_connection.in_waiting).decode('utf-8', errors='ignore')
                            while '\n' in buffer:
                                full_line, buffer = buffer.split('\n', 1)
                                if 'ok' in full_line.lower(): ok_received = True; break
                        if ok_received: break
                        time.sleep(0.05)
                    if not ok_received: raise Exception(f"Timeout waiting for OK for command: {c}")
            update_status(f"Moving to Min Z ({min_z})..."); send_wait(f"G90\nG1 Z{min_z} F{speed_xy}\nM400")
            update_status(f"Moving to Center ({cx}, {cy})..."); send_wait(f"G1 X{cx} Y{cy} F{speed_xy}\nM400")
            update_status("Resetting tilt axis to 0°..."); cmd = self._apply_rot_conversion(f"G90\nG1 E0 F{speed_rot}\nM400"); send_wait(cmd)
            update_status("Homing X axis..."); send_wait("G28 X\nM400", timeout_s=60)
            update_status("Homing Y axis..."); send_wait("G28 Y\nM400", timeout_s=60)
            update_status("Homing Z axis..."); send_wait("G28 Z\nM400", timeout_s=60)
            if found_rot:
                update_status("moving to center position"); send_wait(f"G90\nG1 X{cx} Y{cy} Z{cz} F{speed_xy}\nM400")
                update_status(f"Moving to max tilt angle ({max_rot:.1f}°)..."); cmd = self._apply_rot_conversion(f"G1 E{max_rot:.2f} F{speed_rot}\nM400"); send_wait(cmd, timeout_s=60)
                update_status(f"Moving to min tilt angle ({min_rot:.1f}°)..."); cmd = self._apply_rot_conversion(f"G1 E{min_rot:.2f} F{speed_rot}\nM400"); send_wait(cmd, timeout_s=60)
            elif pause_time_ms > 0:
                sec = pause_time_ms / 1000.0; update_status(f"Holding at Min Tilt for {sec}s..."); time.sleep(sec)
            update_status("Test sequence finished. Returning to center...")
            update_status("Returning to center position...")
            cmd = self._apply_rot_conversion(f"G1 Z{cz} X{cx} Y{cy} E0 F{speed_xy}\nM400"); send_wait(cmd)
            self.rotation_crash_test_complete = True; self._update_section_borders()
            if self.serial_engine.is_connected():
                update_status("Polling for current position...")
                self.serial_engine.write_command("M114")
                update_status("Collision Test Complete!")
                self.root.after(0, lambda: self.lbl_test_status.config(text="Test Complete", fg=utils.COLOR_TEXT_PRIMARY))
            self.root.after(0, self._show_collision_test_complete_popup)
        except Exception as e:
            self.queue_message(f"Test Failed: {e}", "ERROR")
            if "Error: Must load" in str(e): messagebox.showerror("Setup Error", str(e))
            if hasattr(self, 'lbl_test_status') and self.lbl_test_status.winfo_exists():
                self.root.after(0, lambda e=e: self.lbl_test_status.config(text=f"Test Failed: {e}", fg=utils.COLOR_ACCENT_RED))
        finally:
            self.is_manual_command_running = False; self.is_collision_test_running = False
            self.root.after(0, self._update_all_displays); self.root.after(0, self._reset_test_ui)

    def _show_completion_popup(self, *, title, heading, subtext, accent,
                               btn_bg, btn_fg, btn_active_bg, btn_active_fg):
        """Displays a modal completion popup with the given heading, subtext, and accent styling."""
        popup = tk.Toplevel(self.root)
        popup.title(title); popup.resizable(False, False); popup.grab_set(); popup.attributes("-topmost", True)
        BG = utils.COLOR_PANEL_BG
        FG_SUB = utils.COLOR_TEXT_SECONDARY
        popup.configure(bg=BG, highlightthickness=3, highlightbackground=accent)
        icon_frame = tk.Frame(popup, bg=BG); icon_frame.pack(pady=(30, 0), padx=50)
        tk.Label(icon_frame, text="\u2714", font=("Segoe UI", 72, "bold"), fg=accent, bg=BG).pack()
        tk.Label(popup, text=heading, font=("Segoe UI", 22, "bold"), fg=accent, bg=BG).pack(pady=(8, 0))
        tk.Label(popup, text=subtext, font=("Segoe UI", 11), fg=FG_SUB, bg=BG, justify="center").pack(pady=(6, 20))
        def _close(): popup.grab_release(); popup.destroy()
        btn = tk.Button(popup, text="  OK  ", font=("Segoe UI", 13, "bold"), fg=btn_fg, bg=btn_bg,
                        activebackground=btn_active_bg, activeforeground=btn_active_fg, relief=tk.FLAT,
                        cursor="hand2", bd=0, padx=24, pady=10, command=_close)
        btn.pack(pady=(0, 30)); btn.bind("<Return>", lambda _e: _close())

    def _show_collision_test_complete_popup(self):
        self._show_completion_popup(
            title="Collision Test Complete",
            heading="TEST COMPLETE",
            subtext="The collision avoidance test finished successfully.\n"
                    "The printer has returned to the center position.",
            accent='#00e676',
            btn_bg='#1b5e20', btn_fg="#ffffff",
            btn_active_bg='#2e7d32', btn_active_fg="#ffffff")

    def _reset_test_ui(self):
        if hasattr(self, 'btn_begin_test') and self.btn_begin_test.winfo_exists():
            self.btn_begin_test.config(state=tk.NORMAL); self.btn_exit_test.config(state=tk.NORMAL)

    # --- UI LOGIC ---
    def _on_center_change(self, event=None):
        if self.gcode_filepath: self.process_gcode()
        else: self._update_all_displays()

    def _on_goto_entry_commit(self, event=None):
        """Validates and updates target coordinates from entry fields."""
        try:
            mode = self.coord_mode.get()
            cx, cy, cz = float(self.center_x_var.get()), float(self.center_y_var.get()), float(self.center_z_var.get())
            crot = float(self.center_rot_var.get())
            
            # Read from entries if they have content, otherwise use current target
            def get_val(entry, current_abs, center):
                raw = entry.get().strip()
                if not raw:
                    return current_abs - center if mode == "relative" else current_abs
                return float(raw)

            new_x = get_val(self.goto_x_entry, self.target_abs_x, cx)
            new_y = get_val(self.goto_y_entry, self.target_abs_y, cy)
            new_z = get_val(self.goto_z_entry, self.target_abs_z, cz)
            new_rot = get_val(self.goto_rot_entry, self.target_abs_rot, crot)

            self.target_abs_x = new_x + cx if mode == "relative" else new_x
            self.target_abs_y = new_y + cy if mode == "relative" else new_y
            self.target_abs_z = new_z + cz if mode == "relative" else new_z
            self.target_abs_rot = new_rot + crot if mode == "relative" else new_rot

            # Clamp to bounds
            self.target_abs_x = max(self.PRINTER_BOUNDS['x_min'], min(self.PRINTER_BOUNDS['x_max'], self.target_abs_x))
            self.target_abs_y = max(self.PRINTER_BOUNDS['y_min'], min(self.PRINTER_BOUNDS['y_max'], self.target_abs_y))
            self.target_abs_z = max(self.PRINTER_BOUNDS['z_min'], min(self.PRINTER_BOUNDS['z_max'], self.target_abs_z))

            self.log_message(f"Target updated: X{self.target_abs_x:.2f} Y{self.target_abs_y:.2f} Z{self.target_abs_z:.2f} R{self.target_abs_rot:.2f}")
            self._update_all_displays()
        except ValueError:
            self.log_message("Invalid coordinate input", "ERROR")
            self._update_all_displays()

    def _mark_current_as_center(self):
        if self.last_cmd_abs_x is None: return
        self._mark_tilt_as_level(); self.center_x_var.set(f"{self.last_cmd_abs_x:.2f}"); self.center_y_var.set(f"{self.last_cmd_abs_y:.2f}")
        self.center_z_var.set(f"{self.last_cmd_abs_z:.2f}")
        if self.last_cmd_abs_rot is not None: self.center_rot_var.set(f"{self.last_cmd_abs_rot:.2f}")
        self.center_marked = True; self._on_center_change(); self._update_section_borders()

    def _mark_tilt_as_level(self):
        if not self.serial_connection or self.last_cmd_abs_rot is None: return
        self._send_manual_command("G92 E0"); self.last_cmd_abs_rot = self.target_abs_rot = 0.0
        self.center_rot_var.set("0.00"); self._on_center_change(); self._update_all_displays()

    def _on_data_source_change(self):
        source = self.data_source_var.get()
        if source == "DMM":
            if hasattr(self, 'dmm_mode_combo'):
                self.dmm_mode_combo['values'] = list(DMM_MODES.keys())
            self.dmm_connect_button.config(state=tk.NORMAL)
            if not self.is_dmm_connected:
                self.dmm_status_var.set("DMMs: Disconnected")
                if hasattr(self, 'measure_button'): self.measure_button.config(state=tk.DISABLED)
        else:
            self.dmm_connect_button.config(state=tk.DISABLED)
            # Check if Hub is connected
            hub_api = getattr(self.vivigo_panel.gui.backend, 'api', None) if getattr(self, 'vivigo_panel', None) else None
            if hub_api:
                self.dmm_status_var.set("Hub Connected")
                if hasattr(self, 'measure_button'): self.measure_button.config(state=tk.NORMAL)
            else:
                self.dmm_status_var.set("Hub Disconnected")
                if hasattr(self, 'measure_button'): self.measure_button.config(state=tk.DISABLED)
        self._update_section_borders()

    def _on_dmm_mode_change(self, event=None):
        """
        Callback for when the user selects a new DMM measurement mode from the dropdown.
        Reconfigures connected DMMs with the new SCPI mode.
        """
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

    def _on_hub_mode_change(self, *args):
        """
        Callback for when the user selects new Hub telemetry targets via checkboxes.
        Updates the log file header if it exists.
        """
        # Ensure the cache is updated first
        self.active_hub_modes = [m for m in self.hub_modes_list if self.hub_mode_vars[m].get()]
        
        filepath = self.log_filepath_var.get()
        if filepath and os.path.exists(filepath):
            try:
                has_data = False
                with open(filepath, 'r') as f:
                    for i, _ in enumerate(f):
                        if i >= 3: # 3 header rows
                            has_data = True
                            break
                
                if has_data:
                    if messagebox.askyesno("Change Log File?", "The telemetry targets have changed, but the current log file already contains data.\n\nWould you like to select a new log file to prevent corrupted columns?"):
                        self.select_log_file()
                else:
                    # File is just headers or empty, safe to overwrite with new headers
                    self._initialize_log_file(filepath)
            except Exception as e:
                self.log_message(f"Could not update log file headers: {e}", "WARN")

    def _on_auto_measure_toggle(self):
        """Callback for toggling automated DMM measurements after each G-code move."""
        if self.auto_measure_enabled.get():
            self.log_message("Auto-measurement ENABLED. DMMs will trigger after every move.", "INFO")
        else:
            self.log_message("Auto-measurement DISABLED.", "INFO")

    # --- DMM Methods ---

    def toggle_dmm_connection(self):
        """Connects or disconnects the group of Digital Multimeters."""
        if self.is_dmm_connected:
            self.disconnect_dmms()
        else:
            self.dmm_connect_button.config(state=tk.DISABLED)
            self.dmm_status_var.set("Connecting...")
            threading.Thread(target=self._connect_dmm_thread, daemon=True).start()

    def _connect_dmm_thread(self):
        """Background thread that attempts to initialize communication with DMMs via VISA."""
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

    def toggle_hub_connection(self):
        """Toggles the Vivigo Hub connection."""
        if not self.vivigo_panel:
            self.log_message("Vivigo Hub panel not initialized.", "ERROR")
            return
        
        # Assume VivigoPanel has a 'toggle_connection' or similar method
        # and a 'hub_connected' property/variable
        if hasattr(self.vivigo_panel, 'toggle_connection'):
            self.vivigo_panel.toggle_connection()
            # Assuming vivigo_panel has a way to check connection
            # We need to trigger a state update in MeasurementFrame
            self.root.after(100, self._update_measurement_frame_state)
        else:
            self.log_message("Vivigo Hub panel does not support toggling.", "ERROR")

    def _update_measurement_frame_state(self):
        if hasattr(self, '_meas_frame'):
            self._meas_frame._validate_state()

    def trigger_manual_measurement(self):
        if not self.is_dmm_connected: return
        self.log_message("Triggering manual measurement...", "INFO")
        threading.Thread(target=self._measure_thread, daemon=True).start()

    def select_log_file(self):
        """Opens a file dialog to choose the CSV log file."""
        filename = filedialog.asksaveasfilename(
            title="Save Measurement Log",
            initialdir=_DATA_DIR,
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")],
            initialfile=self.log_filepath_var.get()
        )
        if filename:
            self.log_filepath_var.set(filename)
            self.log_message(f"Selected log file: {filename}", "INFO")
            btn = getattr(self, 'browse_log_btn', None)
            if btn: btn.config(style='GreenRing.TButton')
            
            # Immediate creation and header writing
            self._initialize_log_file(filename)

    def _initialize_log_file(self, filepath):
        """Creates the log file and writes metadata/headers immediately."""
        try:
            # Ensure directory exists
            os.makedirs(os.path.dirname(filepath), exist_ok=True)
            
            # Determine keys
            iv_keys = ['x', 'y', 'z', 'rot'] # Standard IVs
            dv_keys = self._current_dv_keys()

            with open(filepath, 'w', newline='') as f:
                # 1. Metadata rows as comments for standard CSV compatibility
                f.write(f"# SOURCE_IVS:,{','.join(iv_keys)}\n")
                f.write(f"# SOURCE_DVS:,{','.join(dv_keys)}\n")
                
                # 2. Scan Profile Metadata (if available) - Critical for historical viewport locking
                if hasattr(self, 'current_profile_data') and self.current_profile_data:
                    f.write(f"# PARAMS: {json.dumps(self.current_profile_data)}\n")
                
                # 3. Standard Data Headers Row
                writer = csv.writer(f)
                header_row = ['Timestamp'] + iv_keys + dv_keys
                writer.writerow(header_row)

            self.log_message(f"Log file initialized: {os.path.basename(filepath)}", "SUCCESS")
        except Exception as e:
            self.log_message(f"Failed to initialize log file: {e}", "ERROR")

    def launch_visualizer(self):
        path = os.path.join(os.path.dirname(__file__), "visualizer.html")
        if os.path.exists(path): webbrowser.open(f"file:///{path}")

    def on_closing(self):
        """Cleanly shuts down the application."""
        self.is_closing = True
        if self.is_sending or self.is_manual_command_running:
            if not messagebox.askyesno("Confirm Exit", "Operation in progress. Exit?"): return
        
        # 1. Stop all threads/loops
        self.stop_event.set()
        self.is_sending = False
        
        # 2. Cancel all pending after() callbacks
        for _attr in ('after_id', '_enforce_id'):
            _job = getattr(self, _attr, None)
            if _job:
                try: self.root.after_cancel(_job)
                except Exception: pass
        if hasattr(self, 'measurement_frame') and hasattr(self.measurement_frame, '_poll_id'):
            try: self.root.after_cancel(self.measurement_frame._poll_id)
            except Exception: pass
            
        # 3. Disconnect
        self.disconnect_printer(silent=True)
        if self.dmm_group:
            try: self.dmm_group.close()
            except Exception: pass
        
        # 4. Destroy root
        self.root.destroy()

    def _set_coord_mode(self, mode):
        self.coord_mode.set(mode)
        self.abs_button.config(style="Segment.Active.TButton" if mode=="absolute" else "Segment.TButton")
        self.rel_button.config(style="Segment.Active.TButton" if mode=="relative" else "Segment.TButton")
        self.goto_x_entry.delete(0, tk.END); self.goto_y_entry.delete(0, tk.END); self.goto_z_entry.delete(0, tk.END); self._update_all_displays()

    def _safe_set_state(self, widgets, state):
        """Sets state on each widget, ignoring None entries and widgets without a 'state' option."""
        for widget in widgets:
            if widget is None: continue
            try: widget.config(state=state)
            except Exception: pass

    def _set_manual_controls_state(self, state):
        """Enables or disables manual jog buttons and entries."""
        self._safe_set_state(getattr(self, 'manual_buttons', []), state)
        self._safe_set_state(getattr(self, 'manual_entries', []), state)
        # Additional setup/home buttons
        extra = [getattr(self, attr, None) for attr in
                 ('mark_center_button', 'collision_test_button', 'conn_home_button')]
        self._safe_set_state(extra, state)

    def _set_goto_controls_state(self, state):
        """Enables or disables 'Go To' position controls."""
        # Canvases might not have 'state'; _safe_set_state skips those.
        self._safe_set_state(getattr(self, 'goto_controls', []), state)

    def _set_terminal_controls_state(self, state):
        """Enables or disables G-code terminal input."""
        self._safe_set_state(
            [getattr(self, 'terminal_input', None), getattr(self, 'terminal_send_button', None)],
            state)

    def _send_manual_command(self, cmd):
        if self.serial_engine.is_connected():
            threading.Thread(target=self._manual_worker, args=(cmd,), daemon=True).start()

    def _manual_worker(self, cmd):
        # Automatically append M114 if it's a movement command to trigger a position refresh
        is_homing = 'G28' in cmd.upper()
        if any(c in cmd.upper() for c in ['G0', 'G1', 'G28']):
            cmd += "\nM114"
            
        for line in cmd.splitlines():
            line = line.strip()
            if line:
                self.serial_engine.write_command(line)
                time.sleep(0.05)
        self.message_queue.put(("MANUAL_FINISHED", is_homing))

    def _on_mousewheel_scroll(self, event):
        """Handles mousewheel scrolling for the left panel canvas."""
        if event.num == 4:
            self.left_canvas.yview_scroll(-1, "units")
        elif event.num == 5:
            self.left_canvas.yview_scroll(1, "units")
        else:
            # Windows/MacOS
            self.left_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _show_scan_complete_popup(self):
        """Displays a prominent modal popup to notify the user that the scan has completed successfully."""
        self._show_completion_popup(
            title="Scan Complete",
            heading="SCAN COMPLETE",
            subtext="The G-code program has finished executing.",
            accent=utils.COLOR_ACCENT_CYAN,
            btn_bg=utils.COLOR_ACCENT_CYAN, btn_fg=utils.COLOR_BLACK,
            btn_active_bg=utils.COLOR_ACCENT_CYAN, btn_active_fg=utils.COLOR_BLACK)

if __name__ == "__main__":
    root = tk.Tk(); app = GCodeSenderGUI(root); root.protocol("WM_DELETE_WINDOW", app.on_closing); root.mainloop()
