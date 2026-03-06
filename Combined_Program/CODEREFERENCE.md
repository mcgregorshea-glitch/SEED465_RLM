# Code Reference

This document explains the internal working of the `generator_panel.py` and `sender_panel.py` programs.

## 1. Pattern Generator (`generator_panel.py`)

This program generates G-code or CSV files for a 3D printer scan pattern. It calculates a "snake-like" (boustrophedon) path within a defined volume.

### Constants
*   **`PRINTER_LIMITS`**: A dictionary defining the physical safety limits of the printer.
    *   `x`: 110.0 (Implies -110 to +110 range)
    *   `y`: 110.0 (Implies -110 to +110 range)
    *   `z_max`: 250.0
    *   `z_min`: 0.0
*   **Colors & Fonts**: Standardized styling constants (e.g., `COLOR_BG`, `FONT_HEADER`) used for the dark theme.

### Key Variables
*   **`x_symmetric`**, **`y_symmetric`**, **`z_symmetric`**, **`rot_symmetric`** (`tk.BooleanVar`): Control whether inputs are "Min/Max" or "±Offset".
*   **`export_format`** (`tk.StringVar`): Selected output format ("gcode" or "csv").
*   **`include_timestamp`** (`tk.BooleanVar`): Whether to append a timestamp to the filename.
*   **Input Fields** (`ttk.Entry`):
    *   `x_min`, `x_max`, `x_step` (and corresponding Y, Z, Rot fields).
    *   `travelspeed`: Movement speed in mm/min.
    *   `pause_time`: Dwell time in seconds at each point.

### Key Functions

#### Data & Logic
*   **`_get_params_silently()`**
    *   **Input**: None (Reads from UI widgets).
    *   **Output**: Dictionary of parameters (e.g., `{'x_min': -50, ...}`) or `None` if invalid.
    *   **Description**: Reads all input fields, handles symmetric/asymmetric logic conversion. Returns `None` instead of raising errors to allow silent validation for previews.
*   **`create_pattern(params)`**
    *   **Input**: Parameter dictionary.
    *   **Output**: Generator yielding dictionaries `{'x': float, 'y': float, 'z': float, 'rotation': float}`.
    *   **Description**: Implements the boustrophedon logic. Iterates Z, then Y, then X (alternating direction).
*   **`create_gcode(pattern_generator, params, total_points)`**
    *   **Input**: The pattern generator, parameters, and point count.
    *   **Output**: Generator yielding G-code strings.
    *   **Description**: Formats the point data into G-code commands (`G1 X...`, `G4 P...`). Adds metadata headers and a "Magic Profile" JSON comment for reloading settings later.
*   **`_check_printer_bounds(params)`**
    *   **Input**: Parameter dictionary.
    *   **Output**: Tuple `(warning_list, warning_level)`.
    *   **Description**: Compares the generated volume against `PRINTER_LIMITS`. Returns warnings if the pattern exceeds or approaches limits.

#### Visualization
*   **`draw_preview_diagram(params, bounds_warnings, warning_level)`**
    *   **Description**: Draws a 3D wireframe projection on the 2D canvas. Uses an oblique projection algorithm. Draws the printer's safe bounding box and the pattern's bounding box.
*   **`update_statistics(...)`**
    *   **Description**: Calculates and displays total points, volume dimensions, and estimated runtime.

---

## 2. G-Code Sender (`sender_panel.py`)

This program controls the 3D printer via serial USB. It handles connection, manual jogging, file streaming, and DMM data logging.

### Constants
*   **`PRINTER_BOUNDS`**: Defines the working volume for the sender's visualization and safety checks.
    *   Values: `x_min`: 0, `x_max`: 220, `y_min`: 0, `y_max`: 220, `z_min`: 0, `z_max`: 250.
    *   **Note**: Unlike the Generator (which centers on 0,0), the Sender assumes the printer's firmware uses positive coordinates (0-220).
*   **`DMM_CONFIG`**: List defining DMM addresses and scaling factors (e.g., `[120, 100, 'VINP']`).
*   **`Z_MAX_LIMIT_PIN`**: GPIO pin (BCM 17) for the Z-max limit switch.
*   **`Z_PROBE_SPEED`**: Speed (150 mm/min) for the calibration probing move.

### Key Variables (State)
*   **`serial_connection`**: The active `serial.Serial` object. `None` if disconnected.
*   **`last_cmd_abs_x/y/z/e`**: Floats tracking the *last known commanded position* of the printer. Used as the "current position" for relative moves and display.
*   **`target_abs_x/y/z/e`**: Floats tracking the "Go To" target position (Blue marker).
*   **`processed_gcode`**: List of strings. The G-code file loaded into memory, translated to absolute coordinates.
*   **`toolpath_by_layer`**: Dictionary `{z_height: [((x1,y1), (x2,y2)), ...]}`. Stores line segments for 2D/3D visualization.
*   **`dmm_group`**: Instance of `DmmGroup` managing connected DMMs.

### Key Functions

#### Connection & Threading
*   **`connect_printer()`**
    *   **Description**: Starts the `_connect_thread`. Scans ports, attempts handshake (`M105`), and enables "Cold Extrusion" (`M302`).
*   **`check_message_queue()`**
    *   **Description**: The "heartbeat" of the GUI. Runs every 100ms. Pops messages (`LOG`, `PROGRESS`, `POSITION_UPDATE`, etc.) from background threads and updates the UI (Tkinter is not thread-safe, so this bridge is required).

#### G-Code Processing & Sending
*   **`process_gcode()`**
    *   **Description**: Reads the loaded file. Translates coordinates based on the user-defined "Center X/Y/Z". Checks every move against `PRINTER_BOUNDS`. Populates `processed_gcode` and `toolpath_by_layer`.
*   **`gcode_sender_thread(gcode_to_send)`**
    *   **Description**: Background thread. Loops through lines, sends to serial, blocks until `ok` is received. Handles Pause/Stop events. Triggers DMM measurement if "Auto-Measure" is enabled.
*   **`_send_manual_command(command)`**
    *   **Description**: Wrapper for manual commands (Jog, Terminal). Checks busy state, then starts `_send_manual_command_thread`.

#### Motion Control
*   **`_jog(axis, direction)`**
    *   **Input**: Axis ('X', 'Y', 'Z', 'E'), Direction (1 or -1).
    *   **Description**: Calculates absolute target based on `last_cmd_abs` + `step_size`. Sends `G1` move.
*   **`_home_all()`**
    *   **Description**: Starts `_homing_sequence_worker`.
        1.  Sends `G28` (Standard Home).
        2.  Moves Z up searching for `Z_MAX_LIMIT_PIN` (GPIO).
        3.  If hit, reads `M114`, sets `PRINTER_BOUNDS['z_max']`, and backs off.

#### Visualization
*   **`_draw_xy_canvas_guides()`**: Draws the top-down 2D view. Renders grid, toolpath for the *current* Z-layer, and position markers.
*   **`_draw_3d_toolpath()`**: Uses Matplotlib to render the full 3D path. Optimized to redraw only when necessary.
*   **`_draw_e_canvas_gauge()`**: Draws the rotary axis state as a circular gauge.

#### Measurement
*   **`_measure_with_stability()`**
    *   **Description**: Reads DMMs repeatedly until the standard deviation of the window is within `stability_threshold_var`. Returns averaged values.
*   **`_log_measurement_to_file(values, coords)`**
    *   **Description**: Appends timestamp, coordinates, and DMM values to the selected CSV log file.
