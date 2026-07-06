import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import os
import webbrowser
import re
import math

class StatusIndicator(tk.Canvas):
    """
    A circular visual status indicator (LED-style) for the GUI header.
    Supports three states: 'idle' (dim), 'busy' (blue), and 'on' (green).
    """
    def __init__(self, parent, bg_color, size=20):
        super().__init__(parent, width=size, height=size, bg=bg_color, highlightthickness=0)
        self.size = size
        self.padding = 4
        self.circle = self.create_oval(self.padding, self.padding, 
                                      size-self.padding, size-self.padding, 
                                      fill="#333333", outline="#111111")
        
    def set_status(self, state):
        colors = {
            "idle": "#333333", "busy": "#00d2ff", "on": "#00ff00",
            "error": "#ff3333", "warn": "#ffcc00", "off": "#333333"
        }
        color = colors.get(state, "#333333")
        self.itemconfig(self.circle, fill=color)

class HeaderBar(ttk.Frame):
    def __init__(self, parent, gui):
        super().__init__(parent, style='Header.TFrame', height=50)
        self.gui = gui
        self.pack(side=tk.TOP, fill=tk.X)
        self.pack_propagate(False) # Force the frame to keep its defined height
        self._build_ui()

    def _build_ui(self):
        # 1. First: header_status_indicator
        self.gui.header_status_indicator = StatusIndicator(self, self.gui.COLOR_PANEL_BG)
        self.gui.header_status_indicator.pack(side=tk.RIGHT, padx=(5, 20), pady=10)

        # 2. Second: info_button
        # NOTE: padx right=95 clears the 80px E-STOP canvas overlay placed at top-right of root window
        self.gui.info_button = tk.Button(
            self, text="?", font=("Segoe UI", 12, "bold"),
            bg=self.gui.COLOR_PANEL_BG, fg=self.gui.COLOR_ACCENT_CYAN,
            activebackground=self.gui.COLOR_PANEL_BG, activeforeground="#ffffff",
            bd=0, highlightthickness=0,
            takefocus=0, cursor="hand2", padx=10,
            command=self.gui._show_tutorial_popup
        )
        self.gui.info_button.pack(side=tk.RIGHT, padx=(5, 95), pady=8)

        # 3. Third: title_label
        title_label = ttk.Label(self, text="\u26a1 G-CODE SENDER", style='Header.TLabel', 
                                font=self.gui.FONT_HEADER, foreground=self.gui.COLOR_ACCENT_CYAN)
        title_label.pack(side=tk.LEFT, padx=15)

        # 4. Fourth: file_label
        file_label = ttk.Label(self, textvariable=self.gui.header_file_var, style='Header.TLabel', 
                               font=self.gui.FONT_MONO, foreground=self.gui.COLOR_TEXT_SECONDARY)
        file_label.pack(side=tk.LEFT, expand=True, fill=tk.X)

class FooterBar(ttk.Frame):
    def __init__(self, parent, gui):
        super().__init__(parent, style='Footer.TFrame', padding=(10, 8))
        self.gui = gui
        self.pack(side=tk.BOTTOM, fill=tk.X)
        self._build_ui()

    def _build_ui(self):
        port_label = ttk.Label(self, textvariable=self.gui.footer_status_var, style='Footer.TLabel')
        port_label.pack(side=tk.LEFT)
        coord_label = ttk.Label(self, textvariable=self.gui.footer_coords_var, style='Footer.TLabel')
        coord_label.pack(side=tk.RIGHT)

class ConnectionFrame(ttk.LabelFrame):
    def __init__(self, parent, gui):
        super().__init__(parent, text="CONNECTION", padding="10", style='Purple.TLabelframe')
        self.gui = gui
        self.pack(fill=tk.X, pady=(0, 10), padx=5)
        self._build_ui()

    def _build_ui(self):
        # Spacer column - pushes the status label to the far right
        self.columnconfigure(7, weight=1)
        
        # Row 0
        self.gui.connect_button = ttk.Button(self, text="Connect", command=self.gui.toggle_connection, width=13, style='PurpleRing.TButton')
        self.gui.connect_button.grid(row=0, column=0, sticky="w", padx=(0, 10))

        ttk.Label(self, text="Port:").grid(row=0, column=1, sticky="w")
        self.gui.port_combobox = ttk.Combobox(self, textvariable=self.gui.port_var, values=self.gui.available_ports, width=15, state="readonly", font=self.gui.FONT_MONO)
        self.gui.port_combobox.grid(row=0, column=2, padx=(0, 5))
        
        ttk.Button(self, text="Rescan", command=self.gui.rescan_ports, width=7).grid(row=0, column=3, padx=(0, 10))

        ttk.Label(self, text="Baud Rate:").grid(row=0, column=4, sticky="w", padx=(5,0))
        self.gui.baud_entry = ttk.Entry(self, textvariable=self.gui.baud_var, width=10) 
        self.gui.baud_entry.grid(row=0, column=5, padx=(0, 10), sticky="w")
        
        # Homing button placed directly to the right of the Baud Rate input
        self.gui.conn_home_button = ttk.Button(self, text="\u2302", style='PurpleRing.TButton', command=self.gui._home_all, width=3)
        self.gui.conn_home_button.grid(row=0, column=6, padx=(0, 10), sticky="w")
        
        # Add cancel button
        self.gui.cancel_connect_button = ttk.Button(self, text="Cancel", command=self.gui.toggle_connection, width=7, style='Danger.TButton')
        self.gui.cancel_connect_button.grid(row=0, column=7, padx=(0, 10), sticky="w")
        self.gui.cancel_connect_button.grid_remove() # Hide initially
        
        # The status text label moved to column 8
        self.gui.status_label = ttk.Label(self, textvariable=self.gui.connection_status_var, font=self.gui.FONT_BODY_SMALL, style='Filepath.TLabel') 
        self.gui.status_label.grid(row=0, column=8, padx=(0, 0), sticky="e")

class MeasurementFrame(ttk.LabelFrame):
    def __init__(self, parent, gui):
        super().__init__(parent, text="MEASUREMENT", padding="10", style='Purple.TLabelframe')
        self.gui = gui
        self.pack(fill=tk.X, pady=(0, 10), padx=5)
        
        # State variables for UI toggling
        self.dmm_active = tk.BooleanVar(value=False)
        self.hub_active = tk.BooleanVar(value=False)
        
        self._build_ui()

    def _build_ui(self):
        # 1. Selection Toggles
        toggle_frame = ttk.Frame(self, style='Panel.TFrame')
        toggle_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Checkbutton(toggle_frame, text="DMM", variable=self.dmm_active, command=self._update_visibility).pack(side=tk.LEFT, padx=10)
        ttk.Checkbutton(toggle_frame, text="Vivigo Hub", variable=self.hub_active, command=self._update_visibility).pack(side=tk.LEFT, padx=10)

        # 2. Sub-panel Containers
        self.dmm_container = ttk.Frame(self, style='Panel.TFrame')
        self.hub_container = ttk.Frame(self, style='Panel.TFrame')
        
        # 3. CSV Logger (Always present)
        csv_frame = ttk.LabelFrame(self, text="CSV LOGGING", style='Panel.TFrame')
        csv_frame.pack(fill=tk.X, pady=10)
        self.gui.log_path_entry = ttk.Entry(csv_frame, textvariable=self.gui.log_filepath_var, state='readonly')
        self.gui.log_path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.gui.browse_log_btn = ttk.Button(csv_frame, text="Browse", command=self.gui.select_log_file)
        self.gui.browse_log_btn.pack(side=tk.LEFT, padx=5)
        
        # New Settling Delay
        delay_frame = ttk.Frame(self, style='Panel.TFrame')
        delay_frame.pack(fill=tk.X, pady=5)
        ttk.Label(delay_frame, text="Settling Delay (s):").pack(side=tk.LEFT, padx=5)
        ttk.Entry(delay_frame, textvariable=self.gui.pre_measure_delay_var, width=10).pack(side=tk.LEFT, padx=5)

        # 4. Initialize Sub-panels (but don't pack them yet)
        self._build_dmm_ui()
        self._build_hub_ui()

    def _build_dmm_ui(self):
        # DMM Settings
        self.dmm_connect_btn = self.gui.dmm_connect_button = ttk.Button(self.dmm_container, text="Connect DMMs", command=self.gui.toggle_dmm_connection, style='PurpleRing.TButton')
        self.dmm_connect_btn.pack(fill=tk.X, pady=2)
        ttk.Label(self.dmm_container, text="IP Prefix:").pack(anchor='w')
        ttk.Entry(self.dmm_container, textvariable=self.gui.dmm_ip_prefix_var).pack(fill=tk.X)
        
    def _build_hub_ui(self):
        # Instructions for the user
        instructions = "1. Go to 'Vivigo Controls' tab.\n2. Connect Hub in that tab.\n3. Return here and press 'Initialize'."
        self.hub_instr_label = ttk.Label(self.hub_container, text=instructions, justify=tk.LEFT, foreground=self.gui.COLOR_TEXT_SECONDARY)
        self.hub_instr_label.pack(anchor='w', pady=5)
        
        # Initialize/Reset API button
        self.hub_init_btn = ttk.Button(self.hub_container, text="Initialize RLM API", command=self._toggle_hub_api, style='PurpleRing.TButton')
        self.hub_init_btn.pack(fill=tk.X, pady=5)
        
        # Target Log Values (Checkboxes)
        ttk.Label(self.hub_container, text="Target Log Values:").pack(anchor='w', pady=(5,0))
        chk_frame = ttk.Frame(self.hub_container, style='Panel.TFrame')
        chk_frame.pack(fill=tk.X, pady=(0, 5))
        
        for i, mode in enumerate(self.gui.hub_modes_list):
            cb = ttk.Checkbutton(chk_frame, text=mode, variable=self.gui.hub_mode_vars[mode], command=getattr(self.gui, '_on_hub_mode_change', None))
            cb.grid(row=i//3, column=i%3, sticky='w', padx=5, pady=2)
            
        # Dynamic Telemetry Grid Container
        self.hub_readout_frame = ttk.Frame(self.hub_container, style='Panel.TFrame')
        self.hub_readout_frame.pack(anchor='w', pady=5, fill=tk.X)
        self.hub_telemetry_labels = {} # Cache for dynamic labels

    def _toggle_hub_api(self):
        if self.gui.hub_connected.get():
            self._reset_rlm_api()
        else:
            self._initialize_rlm_api()

    def _reset_rlm_api(self):
        """Resets the RLM API connection state."""
        self.gui.hub_connected.set(False)
        self.hub_init_btn.config(text="Initialize RLM API")
        instructions = "1. Go to 'Vivigo Controls' tab.\n2. Connect Hub in that tab.\n3. Return here and press 'Initialize'."
        self.hub_instr_label.config(text=instructions)
        self.gui.log_message("Vivigo Hub telemetry link disconnected.", "INFO")
        self._validate_state()

    def _initialize_rlm_api(self):
        """Initializes the RLM API backend and updates UI."""
        try:
            if hasattr(self.gui, 'vivigo_panel') and self.gui.vivigo_panel:
                vivigo_gui = self.gui.vivigo_panel.gui
                # Ensure the hub is connected in the GUI tab first
                if not hasattr(vivigo_gui, 'backend') or not vivigo_gui.backend:
                     raise Exception("Vivigo backend not ready. Please connect Hub in Vivigo tab.")
                
                # Check that the backend is actually running
                if not vivigo_gui.backend.running:
                    raise Exception("Vivigo backend not running. Please connect Hub in Vivigo tab.")

                self.gui.hub_connected.set(True)
                
                # Default selection: Select all Hub modes upon connection
                if hasattr(self.gui, 'hub_mode_vars'):
                    for var in self.gui.hub_mode_vars.values():
                        var.set(True)
                    # Force synchronization of the active_hub_modes cache in the main GUI
                    if hasattr(self.gui, '_on_hub_mode_change'):
                        self.gui._on_hub_mode_change()
                
                self.hub_init_btn.config(text="Reset API")
                self.hub_instr_label.config(text="RLM API is active and streaming.")
                self.gui.log_message("Vivigo Hub telemetry link confirmed.", "SUCCESS")
                self._start_telemetry_polling()
                self._validate_state()
            else:
                self.gui.log_message("Vivigo Hub panel not found.", "ERROR")
        except Exception as e:
            self.gui.log_message(f"Initialization Failed: {e}", "ERROR")
            self.gui.hub_connected.set(False)
            self._validate_state()

    def _start_telemetry_polling(self):
        """Periodically fetches data and updates the UI."""
        if getattr(self.gui, 'is_closing', False):
            return
        if not self.gui.hub_connected.get():
            return

        data = self._fetch_telemetry_data()
        self._update_telemetry_ui(data)

        self._poll_id = self.gui.root.after(1000, self._start_telemetry_polling)

    def _fetch_telemetry_data(self):
        """Fetches and aggregates the latest telemetry and positional data."""
        try:
            backend = self.gui.vivigo_panel.gui.backend
            # Get the board system data
            sys_data = backend.board.get_system_data() if hasattr(backend.board, 'get_system_data') else None
            
            # Extract latest data dictionary from the deque
            data = {}
            if sys_data and hasattr(sys_data, '_data_deque') and len(sys_data._data_deque) > 0:
                data = sys_data._data_deque[-1].copy()
            
            if data:
                # Store in shared state for telemetry providers
                self.gui.last_hub_data = data.copy()
                
                # Helper to get float from StringVar safely
                def get_center(var):
                    try: return float(var.get())
                    except: return 0.0

                # Inject spatial coordinates from sender panel for UI display
                # Subtract center offsets to provide relative coordinates to the EventBus/Visualizers
                # Safely default None to 0.0 if printer hasn't moved yet
                data['x'] = (self.gui.last_cmd_abs_x or 0.0) - get_center(self.gui.center_x_var)
                data['y'] = (self.gui.last_cmd_abs_y or 0.0) - get_center(self.gui.center_y_var)
                data['z'] = (self.gui.last_cmd_abs_z or 0.0) - get_center(self.gui.center_z_var)
                data['rot'] = (self.gui.last_cmd_abs_rot or 0.0) - get_center(self.gui.center_rot_var)
            return data
        except Exception as e:
            self.gui.log_message(f"Telemetry Fetch Error: {e}", "ERROR")
            return {}

    def _update_telemetry_ui(self, data):
        """Updates the telemetry grid labels with the provided data dictionary."""
        # Broadcast hub telemetry as strictly DVs to other panels
        if hasattr(self.gui, 'event_bus') and self.gui.event_bus:
            self.gui.root.after(0, lambda: self.gui.event_bus.publish("hub_telemetry", data))

        if not data: return
        for key, val in data.items():
            label_text = f"{key}: {val:.3f}" if isinstance(val, (int, float)) else f"{key}: {str(val)}"
            if key not in self.hub_telemetry_labels:
                lbl = ttk.Label(self.hub_readout_frame, text=label_text, font=self.gui.FONT_MONO, foreground=self.gui.COLOR_ACCENT_CYAN)
                lbl.grid(row=len(self.hub_telemetry_labels) // 2, column=len(self.hub_telemetry_labels) % 2, sticky='w', padx=5, pady=2)
                self.hub_telemetry_labels[key] = lbl
            else:
                self.hub_telemetry_labels[key].config(text=label_text)

    def _update_visibility(self):
        if self.dmm_active.get(): self.dmm_container.pack(fill=tk.X, pady=5)
        else: self.dmm_container.pack_forget()
        
        if self.hub_active.get(): self.hub_container.pack(fill=tk.X, pady=5)
        else: self.hub_container.pack_forget()
        
        # Determine the active data source based on checkboxes
        if self.hub_active.get():
            self.gui.data_source_var.set("Hub")
        elif self.dmm_active.get():
            self.gui.data_source_var.set("DMM")
                
        self._validate_state()

    def _validate_state(self):
        # Check connection status from self.gui
        dmm_connected = getattr(self.gui, 'is_dmm_connected', False)
        hub_connected = getattr(self.gui, 'hub_connected', tk.BooleanVar(value=False)).get()
        
        # Update button styles
        if hasattr(self, 'hub_init_btn'):
            self.hub_init_btn.config(style='GreenRing.TButton' if hub_connected else 'PurpleRing.TButton')
        if hasattr(self, 'dmm_connect_btn'):
            self.dmm_connect_btn.config(style='GreenRing.TButton' if dmm_connected else 'PurpleRing.TButton')

        dmm_ok = not self.dmm_active.get() or dmm_connected
        hub_ok = not self.hub_active.get() or hub_connected
        csv_ok = bool(self.gui.log_filepath_var.get())
        
        if dmm_ok and hub_ok and csv_ok:
            self.configure(style='Green.TLabelframe')
        else:
            self.configure(style='Purple.TLabelframe')


class ExecutionControlFrame(ttk.LabelFrame):
    def __init__(self, parent, gui):
        super().__init__(parent, text="EXECUTION CONTROL", padding=10, style='Purple.TLabelframe')
        self.gui = gui
        self.pack(fill=tk.X, pady=(0, 10), padx=5)
        self._build_ui()

    def _build_ui(self):
        self.columnconfigure(1, weight=1)

        # --- Action Buttons Row ---
        btn_row = ttk.Frame(self, style='Panel.TFrame')
        btn_row.pack(fill=tk.X)

        self.gui.start_button = ttk.Button(btn_row, text="Start Sending", command=self.gui.start_sending, state=tk.DISABLED, style='Primary.TButton')
        self.gui.start_button.pack(side=tk.LEFT, padx=(0, 5))

        self.gui.pause_resume_button = ttk.Button(btn_row, text="Pause", command=self.gui.toggle_pause_resume, state=tk.DISABLED)
        self.gui.pause_resume_button.pack(side=tk.LEFT, padx=(0, 10))

        # M112: A hard stop that requires a printer reset.
        self.gui.stop_button = ttk.Button(btn_row, text="EMERGENCY STOP", command=self.gui.emergency_stop, state=tk.NORMAL, style='Danger.TButton')
        self.gui.stop_button.pack(side=tk.LEFT)

        # M410: A soft stop that finishes the current move then halts.
        self.gui.quick_stop_button = ttk.Button(btn_row, text="QUICK STOP", command=self.gui.quick_stop, state=tk.NORMAL, style='Amber.TButton')
        self.gui.quick_stop_button.pack(side=tk.LEFT, padx=(10, 0))

        self.gui.restart_scan_button = ttk.Button(btn_row, text="↺ Restart Scan", command=self.gui.restart_scan, state=tk.DISABLED)
        self.gui.restart_scan_button.pack(side=tk.LEFT, padx=(10, 0))

        # --- Progress Display ---
        self.gui.progress_label = ttk.Label(self, textvariable=self.gui.progress_label_var, font=self.gui.FONT_MONO, foreground=self.gui.COLOR_TEXT_SECONDARY)
        self.gui.progress_label.pack(fill=tk.X, padx=2, pady=(8, 2))

        self.gui.progress_bar = ttk.Progressbar(self, orient=tk.HORIZONTAL, mode='determinate', variable=self.gui.progress_var)
        self.gui.progress_bar.pack(fill=tk.X, padx=2)

class FileCenterFrame(ttk.LabelFrame):
    def __init__(self, parent, gui):
        super().__init__(parent, text="SETUP", padding="10", style='Purple.TLabelframe')
        self.gui = gui
        self.pack(fill=tk.X, pady=(0, 10), padx=5)
        self._build_ui()

    def _build_ui(self):
        self.columnconfigure(0, weight=1); self.columnconfigure(1, weight=1); self.columnconfigure(2, weight=1); self.columnconfigure(3, weight=1); self.columnconfigure(4, weight=1)

        # Row 0: G-Code File Selection
        self.gui.select_gcode_button = ttk.Button(self, text="Select G-Code File", command=self.gui.select_file, style='PurpleRing.TButton')
        self.gui.select_gcode_button.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0,10), padx=(0,5))
        ttk.Button(self, text="Clear", command=self.gui.clear_file).grid(row=0, column=2, sticky="ew", pady=(0,10), padx=(5,5))
        ttk.Label(self, textvariable=self.gui.file_path_var, wraplength=400, style='Filepath.TLabel').grid(row=0, column=3, columnspan=2, sticky="ew", pady=(0,10), padx=(5,0))

        # Row 1: Setup Action Buttons
        self.gui.mark_center_button = ttk.Button(self, text="Mark Current as Center", command=self.gui._mark_current_as_center, state=tk.DISABLED, style='PurpleRing.TButton')
        self.gui.mark_center_button.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(0,10), padx=(0,5))

        self.gui.collision_test_button = ttk.Button(self, text="Collision Avoidance Test", command=self.gui._open_collision_test_screen, state=tk.DISABLED, style='PurpleRing.TButton')
        self.gui.collision_test_button.grid(row=1, column=3, columnspan=2, sticky="ew", pady=(0,10), padx=(5,0))

        # Row 2: Labels
        ttk.Label(self, text="Center X:").grid(row=2, column=0, sticky="w", padx=(0, 5))
        ttk.Label(self, text="Center Y:").grid(row=2, column=1, sticky="w", padx=(5, 5))
        ttk.Label(self, text="Center Z:").grid(row=2, column=2, sticky="w", padx=(5, 5))
        ttk.Label(self, text="Center Tilt:").grid(row=2, column=3, sticky="w", padx=(5, 5))
        ttk.Label(self, text="E-Cal (mm/\u00b0):").grid(row=2, column=4, sticky="w", padx=(5, 5))

        # Row 3: Entries and Lock Toggle
        self.gui.center_x_entry = ttk.Entry(self, textvariable=self.gui.center_x_var, width=8)
        self.gui.center_x_entry.grid(row=3, column=0, sticky="ew", pady=(2,0), padx=(0, 5))

        self.gui.center_y_entry = ttk.Entry(self, textvariable=self.gui.center_y_var, width=8)
        self.gui.center_y_entry.grid(row=3, column=1, sticky="ew", pady=(2,0), padx=(5, 5))

        self.gui.center_z_entry = ttk.Entry(self, textvariable=self.gui.center_z_var, width=8)
        self.gui.center_z_entry.grid(row=3, column=2, sticky="ew", pady=(2,0), padx=(5, 5))
        
        self.gui.center_rot_entry = ttk.Entry(self, textvariable=self.gui.center_rot_var, width=8)
        self.gui.center_rot_entry.grid(row=3, column=3, sticky="ew", pady=(2,0), padx=(5, 5))

        ecal_container = ttk.Frame(self, style='Panel.TFrame')
        ecal_container.grid(row=3, column=4, sticky="ew", pady=(2,0), padx=(5, 5), columnspan=2)

        self.gui.e_cal_entry = ttk.Entry(ecal_container, textvariable=self.gui.mm_per_degree_var, width=8, state='readonly')
        self.gui.e_cal_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.gui.e_cal_lock_var = tk.BooleanVar(value=True)
        
        def toggle_e_cal():
            is_locked = not self.gui.e_cal_lock_var.get()
            
            if not is_locked:
                if not messagebox.askyesno("Confirm Unlock", "Are you sure you want to change the rotation ratio? This will affect tilt axis movements and could cause collisions."):
                    return
                self.gui.e_cal_lock_var.set(False)
                self.gui.e_cal_entry.config(state='normal')
                self.gui.e_cal_lock_button.config(text="\u1f513") # Encoding workaround
                self.gui.e_cal_lock_button.config(text="\U0001f513")
            else:
                self.gui.e_cal_lock_var.set(True)
                self.gui.e_cal_entry.config(state='readonly')
                self.gui.e_cal_lock_button.config(text="\U0001f512")

        self.gui.e_cal_lock_button = ttk.Button(ecal_container, text="\U0001f512", width=3, style='ViewCube.TButton', command=toggle_e_cal)
        self.gui.e_cal_lock_button.pack(side=tk.LEFT, padx=(2,0))

        # Bind changes in the center entries to update the coordinate displays.
        for entry in [self.gui.center_x_entry, self.gui.center_y_entry, self.gui.center_z_entry, self.gui.center_rot_entry]:
            entry.bind('<FocusOut>', self.gui._on_center_change)
            entry.bind('<Return>', self.gui._on_center_change)

class TerminalTab(ttk.Frame):
    def __init__(self, parent, gui):
        super().__init__(parent, style='Panel.TFrame')
        self.gui = gui
        self._build_ui()

    def _build_ui(self):
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)
        self.gui.log_area = scrolledtext.ScrolledText(self, height=10, wrap=tk.WORD, state=tk.DISABLED,
                                                    font=self.gui.FONT_TERMINAL,
                                                    bg=self.gui.COLOR_BLACK,
                                                    fg=self.gui.COLOR_ACCENT_GREEN,
                                                    bd=0,
                                                    padx=10,
                                                    pady=10)
        self.gui.log_area.grid(row=0, column=0, sticky="nsew", padx=5, pady=(5,0))
        
        terminal_frame = ttk.Frame(self, padding=(5, 5), style='Panel.TFrame')
        terminal_frame.grid(row=1, column=0, sticky="ew", padx=5, pady=(0, 5))
        terminal_frame.columnconfigure(0, weight=1)
        
        self.gui.terminal_input = ttk.Entry(terminal_frame, state=tk.DISABLED, font=self.gui.FONT_MONO)
        self.gui.terminal_input.grid(row=0, column=0, sticky="ew", padx=(0, 5))
        
        self.gui.terminal_send_button = ttk.Button(terminal_frame, text="Send", state=tk.DISABLED, command=self.gui._send_from_terminal, style='Primary.TButton')
        self.gui.terminal_send_button.grid(row=0, column=1, sticky="w")
        
        self.gui.terminal_input.bind('<Return>', self.gui._send_from_terminal)
        self.gui.terminal_input.bind('<Up>', self.gui._history_up)
        self.gui.terminal_input.bind('<Down>', self.gui._history_down)

class DisplayTab(ttk.Frame):
    def __init__(self, parent, gui):
        super().__init__(parent, style='Panel.TFrame')
        self.gui = gui
        self._build_ui()

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        self.gui._3d_control_bar = ttk.Frame(self, style="Panel.TFrame", padding=(5, 5))
        self.gui._3d_control_bar.grid(row=0, column=0, sticky="ew")
        control_bar = self.gui._3d_control_bar
        
        self.gui.toggle_3d_button = ttk.Button(
            control_bar, 
            text="3D PLOT", 
            command=self.gui.plot_manager.toggle_3d_plot
        )
        self.gui.toggle_3d_button.pack(side=tk.LEFT)

        self.gui.launch_visualizer_button = ttk.Button(
            control_bar,
            text="Launch Data Visualizer",
            style='Small.TButton',
            command=self.launch_visualizer
        )
        self.gui.launch_visualizer_button.pack(side=tk.LEFT, padx=(10, 0))

        self.gui.dmm_reading_label = tk.Label(
            control_bar,
            textvariable=self.gui.last_measurement_var,
            font=self.gui.FONT_MONO,
            foreground=self.gui.COLOR_ACCENT_CYAN,
            background=self.gui.COLOR_PANEL_BG,
            padx=15
        )
        self.gui.dmm_reading_label.pack(side=tk.LEFT, padx=(20, 0))

        self.gui.plot_container_frame = ttk.Frame(self, style="Panel.TFrame")
        self.gui.plot_container_frame.grid(row=1, column=0, sticky="nsew")
        self.gui.plot_container_frame.rowconfigure(0, weight=1)
        self.gui.plot_container_frame.columnconfigure(0, weight=1)

        self.gui.plot_manager.update_3d_plot_button_style()
        self.gui.plot_manager.update_3d_plot_visibility()

    def launch_visualizer(self):
        """Opens the 'visualizer.html' data analysis tool in the system's default web browser."""
        script_dir = os.path.dirname(os.path.abspath(__file__))
        visualizer_path = os.path.abspath(os.path.join(script_dir, "..", "visualizer.html"))
        if os.path.exists(visualizer_path):
            webbrowser.open_new_tab(f"file:///{visualizer_path}")
        else:
            messagebox.showerror(
                "File Not Found",
                f"Could not find visualizer.html at:\n{visualizer_path}"
            )

class PositionControlFrame(ttk.LabelFrame):
    def __init__(self, parent, gui):
        super().__init__(parent, text="POSITION CONTROL & STATUS", padding="10")
        self.gui = gui
        self.pack(fill=tk.X, expand=True, pady=(0, 10), padx=5)
        self._build_ui()

    def _build_ui(self):
        self.columnconfigure(1, weight=1)
        
        # --- Coordinate Mode Buttons (Absolute/Relative) ---
        mode_frame = ttk.Frame(self, style='Panel.TFrame')
        mode_frame.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 10))
        
        self.gui.abs_button = ttk.Button(mode_frame, text="ABSOLUTE", command=lambda: self.gui._set_coord_mode("absolute"), style='Segment.TButton')
        self.gui.abs_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 0))
        self.gui.rel_button = ttk.Button(mode_frame, text="RELATIVE", command=lambda: self.gui._set_coord_mode("relative"), style='Segment.TButton')
        self.gui.rel_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 0))

        # --- DRO (Digital Readout) Frame ---
        dro_bg_frame = tk.Frame(self, bg=self.gui.COLOR_BLACK, relief='sunken', borderwidth=1)
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
        ttk.Label(dro_frame, textvariable=self.gui.last_cmd_x_display_var, style='Red.DRO.TLabel').grid(row=1, column=1, sticky="ew")
        ttk.Label(dro_frame, textvariable=self.gui.last_cmd_y_display_var, style='Red.DRO.TLabel').grid(row=1, column=2, sticky="ew", padx=5)
        ttk.Label(dro_frame, textvariable=self.gui.last_cmd_z_display_var, style='Red.DRO.TLabel').grid(row=1, column=3, sticky="ew")
        ttk.Label(dro_frame, textvariable=self.gui.last_cmd_rot_display_var, style='Red.DRO.TLabel').grid(row=1, column=4, sticky="ew", padx=5)
        
        # Labels for the 'TARGET' (Go To) position.
        ttk.Label(dro_frame, text=" TARGET:", style='DRO.TLabel').grid(row=2, column=0, sticky="w"); 
        ttk.Label(dro_frame, textvariable=self.gui.goto_x_display_var, style='Blue.DRO.TLabel').grid(row=2, column=1, sticky="ew")
        ttk.Label(dro_frame, textvariable=self.gui.goto_y_display_var, style='Blue.DRO.TLabel').grid(row=2, column=2, sticky="ew", padx=5)
        ttk.Label(dro_frame, textvariable=self.gui.goto_z_display_var, style='Blue.DRO.TLabel').grid(row=2, column=3, sticky="ew")
        ttk.Label(dro_frame, textvariable=self.gui.goto_rot_display_var, style='Blue.DRO.TLabel').grid(row=2, column=4, sticky="ew", padx=5)
        
        # --- Frame to hold the XY, Z, and E canvases ---
        self.gui.canvas_frame = ttk.Frame(self, height=150, style='Panel.TFrame')
        self.gui.canvas_frame.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(10,0))
        self.gui.canvas_frame.rowconfigure(0, weight=1)
        self.gui.canvas_frame.columnconfigure(1, weight=1)

        # --- 'Go To' Position Input Fields ---
        input_frame = ttk.Frame(self.gui.canvas_frame, style='Panel.TFrame')
        input_frame.grid(row=0, column=0, sticky="nsw", padx=(0, 10))
        
        ttk.Label(input_frame, text="Set X:").grid(row=0, column=0, sticky="w"); self.gui.goto_x_entry = ttk.Entry(input_frame, width=8, state=tk.DISABLED); self.gui.goto_x_entry.grid(row=0, column=1, sticky="w", pady=2)
        self.gui.goto_x_entry.bind('<Return>', self.gui._on_goto_entry_commit); self.gui.goto_x_entry.bind('<FocusOut>', self.gui._on_goto_entry_commit)
        
        ttk.Label(input_frame, text="Set Y:").grid(row=1, column=0, sticky="w"); self.gui.goto_y_entry = ttk.Entry(input_frame, width=8, state=tk.DISABLED); self.gui.goto_y_entry.grid(row=1, column=1, sticky="w", pady=2)
        self.gui.goto_y_entry.bind('<Return>', self.gui._on_goto_entry_commit); self.gui.goto_y_entry.bind('<FocusOut>', self.gui._on_goto_entry_commit)

        ttk.Label(input_frame, text="Set Z:").grid(row=2, column=0, sticky="w"); self.gui.goto_z_entry = ttk.Entry(input_frame, width=8, state=tk.DISABLED); self.gui.goto_z_entry.grid(row=2, column=1, sticky="w", pady=2)
        self.gui.goto_z_entry.bind('<Return>', self.gui._on_goto_entry_commit); self.gui.goto_z_entry.bind('<FocusOut>', self.gui._on_goto_entry_commit)

        ttk.Label(input_frame, text="Set R (\u00b0):").grid(row=3, column=0, sticky="w"); self.gui.goto_rot_entry = ttk.Entry(input_frame, width=8, state=tk.DISABLED); self.gui.goto_rot_entry.grid(row=3, column=1, sticky="w", pady=2)
        self.gui.goto_rot_entry.bind('<Return>', self.gui._on_goto_entry_commit); self.gui.goto_rot_entry.bind('<FocusOut>', self.gui._on_goto_entry_commit)
        
        self.gui.go_button = ttk.Button(input_frame, text="Go", command=self.gui._go_to_position, state=tk.DISABLED, style='Primary.TButton'); self.gui.go_button.grid(row=4, column=0, columnspan=2, pady=(10, 0), sticky="ew")

        self.gui.go_to_center_button = ttk.Button(input_frame, text="Go to Center", command=self.gui._go_to_center, state=tk.DISABLED)
        self.gui.go_to_center_button.grid(row=5, column=0, columnspan=2, pady=(5, 0), sticky="ew")

        # --- Visualization Canvases ---
        canvas_size = 215
        # XY (top-down) view
        self.gui.xy_canvas = tk.Canvas(self.gui.canvas_frame, width=canvas_size, height=canvas_size, bg=self.gui.COLOR_BLACK, highlightthickness=1, highlightbackground=self.gui.COLOR_BORDER)
        self.gui.xy_canvas.grid(row=0, column=1, sticky="n", padx=8)
        self.gui.xy_canvas.bind("<Button-1>", self.gui._on_xy_canvas_click); self.gui.xy_canvas.bind("<B1-Motion>", self.gui._on_xy_canvas_click); self.gui.xy_canvas.bind("<Configure>", self.gui._draw_xy_canvas_guides)
        
        # Z (side) view
        self.gui.z_canvas = tk.Canvas(self.gui.canvas_frame, width=50, height=canvas_size, bg=self.gui.COLOR_BLACK, highlightthickness=1, highlightbackground=self.gui.COLOR_BORDER)
        self.gui.z_canvas.grid(row=0, column=2, sticky="ns", padx=8)
        self.gui.z_canvas.bind("<Button-1>", self.gui._on_z_canvas_click); self.gui.z_canvas.bind("<B1-Motion>", self.gui._on_z_canvas_click); self.gui.z_canvas.bind("<Configure>", self.gui._draw_z_canvas_marker)

        # E (Rotation) view \u2014 wrapper spans both canvas rows so the group can be vertically centred
        tilt_group = ttk.Frame(self.gui.canvas_frame)
        tilt_group.grid(row=0, column=3, rowspan=2, sticky="ns", padx=8)

        tilt_inner = ttk.Frame(tilt_group)
        tilt_inner.pack(expand=True)  # centres tilt_inner vertically inside tilt_group

        self.gui.e_container = ttk.Frame(tilt_inner, style='Panel.TFrame', width=canvas_size, height=canvas_size)
        self.gui.e_container.pack()
        self.gui.e_container.grid_propagate(False)

        self.gui.e_canvas = tk.Canvas(self.gui.e_container, width=canvas_size, height=canvas_size, bg=self.gui.COLOR_BLACK, highlightthickness=1, highlightbackground=self.gui.COLOR_BORDER)
        self.gui.e_canvas.place(x=0, y=0, relwidth=1, relheight=1)
        self.gui.e_canvas.bind("<Button-1>", self.gui._on_e_canvas_click); self.gui.e_canvas.bind("<B1-Motion>", self.gui._on_e_canvas_click); self.gui.e_canvas.bind("<Configure>", self.gui._draw_e_canvas_gauge)

        # Snap Buttons
        btn_style = 'Segment.TButton'
        def set_rot(val):
            self.gui.target_abs_rot = float(val)
            self.gui._update_all_displays()

        ttk.Button(self.gui.e_container, text=" 0\u00b0 ", width=4, style=btn_style, command=lambda: set_rot(0)).place(relx=0.5, rely=0.54, anchor='n')
        ttk.Button(self.gui.e_container, text="+90\u00b0", width=4, style=btn_style, command=lambda: set_rot(90)).place(relx=0.98, rely=0.09, anchor='e')
        ttk.Button(self.gui.e_container, text="-90\u00b0", width=4, style=btn_style, command=lambda: set_rot(-90)).place(relx=0.02, rely=0.09, anchor='w')

        # "Mark Tilt as Level" packed directly below the gauge in the same inner group
        ttk.Button(tilt_inner, text="Mark Tilt as Level (0\u00b0)", command=self.gui._mark_tilt_as_level).pack(pady=(4, 0), fill='x')

        # --- 2D Plot Toggle ---
        self.gui.toggle_2d_button = ttk.Button(
            self,
            text="2D TOOLPATH",
            command=self.gui._toggle_2d_plot_button
        )
        self.gui.toggle_2d_button.grid(row=3, column=0, columnspan=3, sticky="w", pady=(10, 0), padx=5)

        # Store controls for easy enabling/disabling based on connection state.
        self.gui.goto_controls = [ self.gui.goto_x_entry, self.gui.goto_y_entry, self.gui.goto_z_entry, self.gui.goto_rot_entry, self.gui.go_button,
                                 self.gui.go_to_center_button, self.gui.xy_canvas, self.gui.z_canvas, self.gui.e_canvas ]
        self.gui._set_coord_mode("absolute") # Set initial mode and style

class ManualJogFrame(ttk.LabelFrame):
    def __init__(self, parent, gui):
        super().__init__(parent, text="MANUAL JOG CONTROL", padding="10")
        self.gui = gui
        self.pack(fill=tk.X, pady=(0, 10), padx=5)
        self._build_ui()

    def _build_ui(self):
        # 5 columns: Spacer, Z-Axis, XY-Grid, E-Axis, Spacer
        self.columnconfigure(0, weight=1); self.columnconfigure(4, weight=1)
        
        # --- Z-axis jog buttons (Left Wing) ---
        z_control_frame = ttk.Frame(self, style='Panel.TFrame')
        z_control_frame.grid(row=0, column=1, rowspan=3, sticky="ns", padx=(0,15)); 
        ttk.Label(z_control_frame, text="Z-AXIS", font=self.gui.FONT_BODY_BOLD).pack(pady=(0,5))
        
        self.gui.jog_z_pos = ttk.Button(z_control_frame, text="Z+", command=lambda: self.gui._jog('Z', 1), state=tk.DISABLED, style='Jog.TButton'); self.gui.jog_z_pos.pack(pady=(2, 10), fill=tk.X)
        self.gui.jog_z_neg = ttk.Button(z_control_frame, text="Z-", command=lambda: self.gui._jog('Z', -1), state=tk.DISABLED, style='Jog.TButton'); self.gui.jog_z_neg.pack(pady=(10, 2), fill=tk.X)
        
        # --- XY-axis jog buttons (Center) ---
        jog_grid_frame = ttk.Frame(self, style='Panel.TFrame')
        jog_grid_frame.grid(row=0, column=2, rowspan=3); 

        jog_grid_frame.rowconfigure(0, weight=1, uniform="jog_grid")
        jog_grid_frame.rowconfigure(1, weight=1, uniform="jog_grid")
        jog_grid_frame.rowconfigure(2, weight=1, uniform="jog_grid")
        jog_grid_frame.columnconfigure(0, weight=1, uniform="jog_grid")
        jog_grid_frame.columnconfigure(1, weight=1, uniform="jog_grid")
        jog_grid_frame.columnconfigure(2, weight=1, uniform="jog_grid")
        
        self.gui.jog_y_neg = ttk.Button(jog_grid_frame, text="Y-", style='Jog.TButton', command=lambda: self.gui._jog('Y', -1), state=tk.DISABLED); self.gui.jog_y_neg.grid(row=0, column=1, padx=2, pady=2, sticky="nsew")
        self.gui.jog_x_neg = ttk.Button(jog_grid_frame, text="X-", style='Jog.TButton', command=lambda: self.gui._jog('X', -1), state=tk.DISABLED); self.gui.jog_x_neg.grid(row=1, column=0, padx=2, pady=2, sticky="nsew")
        self.gui.home_button = ttk.Button(jog_grid_frame, text="\u2302", style='Home.TButton', command=self.gui._home_all, state=tk.DISABLED); self.gui.home_button.grid(row=1, column=1, padx=2, pady=2, sticky="nsew")
        self.gui.jog_x_pos = ttk.Button(jog_grid_frame, text="X+", style='Jog.TButton', command=lambda: self.gui._jog('X', 1), state=tk.DISABLED); self.gui.jog_x_pos.grid(row=1, column=2, padx=2, pady=2, sticky="nsew")
        self.gui.jog_y_pos = ttk.Button(jog_grid_frame, text="Y+", style='Jog.TButton', command=lambda: self.gui._jog('Y', 1), state=tk.DISABLED); self.gui.jog_y_pos.grid(row=2, column=1, padx=2, pady=2, sticky="nsew")
        
        # --- E-axis jog buttons (Right Wing) ---
        e_control_frame = ttk.Frame(self, style='Panel.TFrame')
        e_control_frame.grid(row=0, column=3, rowspan=3, sticky="ns", padx=(15,0)); 
        ttk.Label(e_control_frame, text="R-AXIS", font=self.gui.FONT_BODY_BOLD).pack(pady=(0,5))
        
        self.gui.jog_rot_pos = ttk.Button(e_control_frame, text="\u21bb", command=lambda: self.gui._jog('E', 1), state=tk.DISABLED, style='JogIcon.TButton'); self.gui.jog_rot_pos.pack(pady=(2, 10), fill=tk.X)
        self.gui.jog_rot_neg = ttk.Button(e_control_frame, text="\u21ba", command=lambda: self.gui._jog('E', -1), state=tk.DISABLED, style='JogIcon.TButton'); self.gui.jog_rot_neg.pack(pady=(10, 2), fill=tk.X)

        # --- Entry fields for jog parameters ---
        jog_params_frame = ttk.Frame(self, style='Panel.TFrame')
        jog_params_frame.grid(row=3, column=0, columnspan=5, pady=(10,0))
        
        # XYZ Params
        ttk.Label(jog_params_frame, text="XYZ Step (mm):").pack(side=tk.LEFT, padx=(0, 5))
        self.gui.jog_step_entry = ttk.Entry(jog_params_frame, textvariable=self.gui.jog_step_var, width=5); self.gui.jog_step_entry.pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(jog_params_frame, text="XYZ Speed (mm/min):").pack(side=tk.LEFT, padx=(0, 5))
        self.gui.jog_feedrate_entry = ttk.Entry(jog_params_frame, textvariable=self.gui.jog_feedrate_var, width=6); self.gui.jog_feedrate_entry.pack(side=tk.LEFT, padx=(0, 20))

        # E Params (Rotation)
        ttk.Label(jog_params_frame, text="Rot Step (deg):").pack(side=tk.LEFT, padx=(0, 5))
        self.gui.rot_step_entry = ttk.Entry(jog_params_frame, textvariable=self.gui.rotation_step_var, width=5); self.gui.rot_step_entry.pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(jog_params_frame, text="Rot Speed (deg/min):").pack(side=tk.LEFT, padx=(0, 5))
        self.gui.rot_feedrate_entry = ttk.Entry(jog_params_frame, textvariable=self.gui.rotation_feedrate_var, width=6); self.gui.rot_feedrate_entry.pack(side=tk.LEFT)

        # Store for state management
        self.gui.manual_buttons = [self.gui.home_button, self.gui.jog_x_neg, self.gui.jog_x_pos, self.gui.jog_y_neg, self.gui.jog_y_pos, self.gui.jog_z_neg, self.gui.jog_z_pos, self.gui.jog_rot_pos, self.gui.jog_rot_neg]
        self.gui.manual_entries = [self.gui.jog_step_entry, self.gui.jog_feedrate_entry, self.gui.rot_step_entry, self.gui.rot_feedrate_entry]
