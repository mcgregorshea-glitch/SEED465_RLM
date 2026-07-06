# SEED Control Center — Structural Reference

Quick lookup for classes, files, and how the application fits together.
For behavioral details, read the source. For safety restrictions, see `../IMMUTABLE_CONSTRAINTS.md`.

---

## Entry Point

| File | Class | Role |
|---|---|---|
| `main.py` | `SEEDApplication` | Constructs the root `ctk.CTk` window, `ttk.Notebook`, and all four top-level panel instances. Owns the E-STOP canvas overlay and global key bindings. **Do not modify.** |

---

## Top-Level Panels (each is a `ttk.Frame` added to the main notebook)

| Tab label | File | Class | Notes |
|---|---|---|---|
| ∿ Pattern Generator | `generator_panel.py` | `PatternGeneratorGUI` | Boustrophedon path design, live preview, G-code export. **Blanket restriction — do not modify.** |
| ✎ Manual Movement | `manual_movement_panel.py` | `ManualMovementPanel` | Visual sequence builder, waveform editor, G-code exporter/sender. |
| ❱ G-Code Sender | `sender_panel.py` | `GCodeSenderGUI` | Hardware orchestration: serial connect/send, DMM measurement, CSV logging. |
| ⚡ Vivigo Hub | `vivigo_panel.py` | `VivigoPanel` | Hosts two subtabs: VIVIGO CONTROLS (`GUI` from legacy) and POP Analysis (`POPVisualizationPanel`). |

> **The PoP Analysis panel is a subtab inside `VivigoPanel`, not a fifth top-level tab. This is an immutable constraint.**

---

## Vivigo Hub Subtabs

| Subtab label | Class | File |
|---|---|---|
| ⚡ VIVIGO CONTROLS | `GUI` | `legacy/gui/src/gui.py` |
| 📊 POP Analysis | `POPVisualizationPanel` | `pop_visualization_panel.py` |

---

## Pattern Generator Subcomponents (`generator_components/`)

| File | Class | Role |
|---|---|---|
| `pattern_input.py` | `PatternInput` | Axis range / step / offset input widgets. |
| `pattern_preview.py` | `PatternPreview` | Canvas diagram preview of the scan path. |
| `command_injection.py` | `CommandInjection` | Per-layer Hub command and Wait injection table. |

---

## G-Code Sender Subcomponents (`sender_components/`)

| File | Class(es) | Role |
|---|---|---|
| `serial_engine.py` | `SerialEngine` | Serial connection lifecycle; `_sender_thread` streams G-code, waits for `ok`, fires layer-change XY homing verification, and gates measurement after each move. **Do not modify.** |
| `ui_layout.py` | `HeaderBar`, `FooterBar`, `ConnectionFrame`, `MeasurementFrame`, `ExecutionControlFrame`, `FileCenterFrame`, `TerminalTab`, `DisplayTab`, `PositionControlFrame`, `StatusIndicator` | All widget layout for the Sender tab. `MeasurementFrame` owns the telemetry polling loop that distributes data to the UI. **Do not modify.** |
| `telemetry_service.py` | `TelemetryProvider` (ABC), `DMMTelemetryProvider`, `VivigoTelemetryProvider` | Data-source providers called during measurement. `DMMTelemetryProvider` reads VISA/SCPI; `VivigoTelemetryProvider` subscribes to the `"hub_telemetry"` EventBus event and caches the latest packet. **Do not modify.** |
| `dmm_manager.py` | `DmmInst`, `DmmGroup` | VISA/SCPI lifecycle for one or more digital multimeters over TCP/IP. `DmmGroup` coordinates simultaneous trigger and averaged read. |
| `gcode_processor.py` | _(module-level functions)_ | `parse_gcode_coords`, `translate_gcode_line`, `process_mixed_sequence` — convert relative-coordinate G-code to validated absolute lines. |
| `motion_controls.py` | `MotionControls` | Jog (`G91` relative move), full home (`G28` all axes), go-to-position, go-to-center for the manual sender controls. |
| `plot_manager.py` | `PlotManager` | Lazy-loaded matplotlib 3D toolpath plot and 2D canvas overlay for the Sender tab. |

---

## Shared Infrastructure

| File | What it provides |
|---|---|
| `utils.py` | `EventBus` (pub/sub), `PRINTER_LIMITS` (design-time, center-relative), `PRINTER_BOUNDS` (runtime, absolute), full color palette constants, font constants, `setup_global_styling()`. **Do not modify `PRINTER_BOUNDS` or the style palette.** |
| `sequence_manager.py` | `SequenceManager` — loads JSON sequences or parses G-code files for embedded `; HUB_CMD` / `; WAIT` / `; HUB_CMD_PRE` / `; WAIT_PRE` tags into a mixed step list. |
| `pop_visualization_panel.py` | `POPVisualizationPanel`, `SliceSelectionDialog`, `StatusIndicator` — live and post-scan PoP data visualization, topographic overlays, PDF/PNG report export. **Report dialogs and export logic are immutable.** |

---

## Tools (not imported by the app, run standalone)

| File | Role |
|---|---|
| `tools/pop_utility.py` | Combines multi-file PoP CSVs and converts to JSON for the web visualizer. **Do not modify.** |
| `tools/debug_dmm.py` | Standalone DMM diagnostic script. |

---

## Threading Model

| Thread | Owner | Communicates via |
|---|---|---|
| GUI / main | Tkinter event loop | — |
| Serial sender | `SerialEngine._sender_thread` | `message_queue` → `GCodeSenderGUI.check_message_queue()` (drains every 100 ms) |
| Connection attempt | `SerialEngine._connect_thread` | `message_queue` |
| Manual command | `GCodeSenderGUI._manual_worker` | `message_queue` |
| Telemetry poll | `MeasurementFrame._start_telemetry_polling` | direct UI calls on the GUI thread |

`serial_lock` (`threading.Lock` inside `SerialEngine`) prevents command interleaving between the sender thread and manual commands.

---

## Scan Cycle Data Flow

1. **Design** — `PatternGeneratorGUI` produces a boustrophedon `.gcode` file.
2. **Load** — `GCodeSenderGUI.load_gcode_file()` → `process_gcode()` translates relative coords to absolute via `gcode_processor.py`, validates against `PRINTER_BOUNDS`.
3. **Send** — `SerialEngine._sender_thread` streams lines, waits for `ok` after each.
4. **Layer change** — on Z change, `_homing_verification_routine` in `sender_panel.py` runs a 3-phase XY drift-detection check (Phase A: safe-zone endstop test; Phase B: home-target endstop test; Phase C: `G28 X Y` sync-home). Drift triggers layer retry and progressive speed-cap reduction.
5. **Measure** — after each move, `M400` (buffer drain) then `take_measurement` callback; retries up to 60× if telemetry is absent.
6. **Log** — `_log_measurement_to_file()` writes a 3-row CSV: metadata header, column header, data row.

---

## Key Callback Registry (`SerialEngine.callbacks`)

| Key | Registered from | Purpose |
|---|---|---|
| `take_measurement` | `sender_panel.py` | Fires after each move during auto-scan |
| `homing_verification` | `sender_panel.py` | Called on each layer change; raises `ValueError` on drift |
| `handle_homing_failure` | `sender_panel.py` | Logs fatal homing errors |
| `get_hub_panel` | `sender_panel.py` | Returns `vivigo_panel` for Hub command injection |
| `apply_e_conversion` | `sender_panel.py` | Translates E-axis values before sending |
| `parse_coords` | `sender_panel.py` | Parses X/Y/Z/E from a G-code line |
| `get_speed_cap` / `set_speed_cap` | `sender_panel.py` | Runtime feedrate ceiling; auto-lowered after repeated drift |
