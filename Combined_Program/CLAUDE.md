# SEED Control Center

Tkinter/CTk hardware orchestrator. Uses a modified Ender 3 as a precision motion fixture to automate electrical measurements over a 3D scan volume.

## Commands

```bash
# Tests
MPLCONFIGDIR=$TMPDIR PYTHONPATH=/home/sheamcg/RLM/venv-linux python3 -m pytest tests/ -v
MPLCONFIGDIR=$TMPDIR PYTHONPATH=/home/sheamcg/RLM/venv-linux python3 -m pytest tests/test_<name>.py -v

# Lint
PYTHONPATH=/home/sheamcg/RLM/venv-linux python3 -m ruff check src/
```

```bat
REM Windows launchers (developer runs manually)
launchers\windows\run_seed_control_center.bat
launchers\windows\run_fake_printer_simulator.bat
launchers\windows\run_fake_hub_simulator.bat
launchers\windows\run_fake_dmm_simulator.bat
.\venv\Scripts\python -m unittest discover tests
```

> Always run the full test suite after edits. ~10s, all hardware is mocked.

## Running/Testing Headlessly (applies to the main app too, not just sims)

- **`TMPDIR` must be exported** before running any launcher script by hand — every
  `launchers/linux/*.sh`, including `run_seed_control_center.sh` itself, does
  `MPLCONFIGDIR="$TMPDIR"` under `set -euo pipefail`, which dies on an unset var.
  The Claude sandbox exports it automatically; a plain interactive shell may not.
- **Screenshotting any Tk/CTk window** (the main app or a sim): use
  `/to-discord capture-gui <script.py>`, which runs it under a throwaway `Xvfb` and
  crops to the window. WSLg's live `:0` display produces solid-black captures for
  Tkinter — always let `capture-gui` spin up its own `Xvfb`, don't point it at `:0`.
- X11, raw sockets, and any `.git` write need `dangerouslyDisableSandbox: true` in
  the sandboxed Bash tool.

## Dev Simulator Stack (Linux)

`launchers/linux/run_dev_sim_stack.sh` wires the fake printer and fake hub sims to a
real running app via `socat` PTY pairs — no physical hardware or Windows virtual COM
ports needed. One command starts everything (socat pairs, both sims, the app) and
returns immediately; `run_dev_sim_stack.sh --stop` tears it down. Force-restarts on
every run. `SEED_SIM_PORT` / `SEED_SIM_HUB_PORT` env vars (read in `sender_panel.py`
and `legacy/gui/src/backend.py`) let the app see the sim ports even though they don't
appear in a normal `list_ports.comports()` scan.

The fake DMM (`sim/fake_dmm.py`) is different: real DMMs connect over network SCPI
(PyVISA TCPIP resource, falling back to `TCPIP::<ip>::5025::SOCKET`), not a serial
port, so there's no PTY pair for it — it binds to a loopback address whose last
octet stands in for the port number (default `127.0.0.50`). Point the existing
"DMM IP Prefix" field (Sender panel → DMM section) at `127.0.0` to reach it.

### Validating a new sim's protocol

If the real app-side connect code has an unrelated bug blocking the handshake (e.g.
`dmm_manager.py`'s `*IDN?` probe timing out — see `sender_components/dmm_manager.py`
and the `DmmInst.connect()` gap), don't let that block verifying the simulator. Open
the resource/port manually and drive the rest of the real client class's methods
(`setup()`/`trigger()`/`ready()`/`read()`, or the printer/hub equivalents) directly —
proves the wire protocol is correct independent of the unrelated bug.

## Structure

```
src/                  — Application source
  main.py             — Root window + E-STOP overlay
  utils.py            — PRINTER_BOUNDS, PRINTER_LIMITS, EventBus, UI theme
  generator_panel.py  — Pattern Generator tab (blanket restriction — see Safety)
  sender_panel.py     — G-Code Sender tab
  manual_movement_panel.py
  vivigo_panel.py     — Vivigo Hub tab (hosts POP Analysis subtab)
  pop_visualization_panel.py
  sequence_manager.py
  sender_components/  — Serial, telemetry, DMM, G-Code, motion, plot submodules
  generator_components/
sim/                  — Simulators (fake_printer.py, fake_hub.py, fake_dmm.py + backends)
tests/                — pytest suite (all hardware mocked)
tools/                — gui_capture.py for headless GUI screenshots (/screenshot-gui)
launchers/linux/      — Bash launch scripts
launchers/windows/    — .bat launch scripts
```

## Application Architecture

### Tabs — exactly four, never add more

| Label | Class |
|---|---|
| ∿ Pattern Generator | `PatternGeneratorGUI` |
| ✎ Manual Movement | `ManualMovementPanel` |
| ❱ G-Code Sender | `GCodeSenderGUI` |
| ⚡ Vivigo Hub | `VivigoPanel` |

**`POPVisualizationPanel` is a subtab of `VivigoPanel` — not a top-level tab.** This mistake has been reintroduced multiple times; do not add it to `main.py`'s notebook.

### `VivigoPanel` subtabs

| Label | Class |
|---|---|
| ⚡ VIVIGO CONTROLS | legacy `GUI` |
| 📊 POP Analysis | `POPVisualizationPanel` |

### `sender_components/`

| File | Role |
|---|---|
| `serial_engine.py` | Serial lifecycle, G-Code streaming, retry/rehome. **Do not modify.** |
| `telemetry_service.py` | EventBus subscriber → UI updates. **Do not modify.** |
| `ui_layout.py` | EventBus publication loop. **Do not modify.** |
| `dmm_manager.py` | VISA/SCPI for DMMs (TCP/IP) |
| `gcode_processor.py` | Validates + converts relative → absolute coords |
| `motion_controls.py` | Jog, homing, go-to-position |
| `plot_manager.py` | 3D matplotlib toolpath viz; lazy-loaded on tab activation |

### Threading Model

- **GUI thread** — only thread allowed to touch widgets
- **Serial thread** — `serial_engine.py`; communicates back via `message_queue`
- **DMM thread** — VISA polling ~10 Hz; same queue
- `check_message_queue()` in `sender_panel.py` drains queue every 100 ms on the GUI thread
- `serial_lock` prevents command interleaving between threads

### Data Flow (Scan Cycle)

1. `PatternGeneratorGUI` computes boustrophedon path → exports `.gcode`
2. `GCodeSenderGUI.load_gcode_file()` → `process_gcode()`: relative → absolute, validates against `PRINTER_BOUNDS`
3. `gcode_sender_thread()` streams line-by-line, waits for `ok`, re-homes on layer change (`G28 X Y` only — **never Z**)
4. `_take_measurement()` triggers all DMMs via `DmmGroup`, waits for stability, reads averaged values
5. `_log_measurement_to_file()` writes 3-row CSV: metadata header, column header, data

## Safety-Critical — Stop and Confirm Before Modifying

- `main.py` — E-STOP overlay and tab management
- `generator_panel.py` — entire panel (blanket restriction)
- `pop_visualization_panel.py` — report dialogs, PDF/PNG export
- `sender_panel.py` — connection state logic, CSV metadata format, hub logging
- `sender_components/telemetry_service.py` — EventBus subscriptions
- `sender_components/serial_engine.py` — retry loop, layer-change rehoming, bounds validation
- `src/utils.py` — `PRINTER_BOUNDS` and style palette

## GUI Conventions

- Panels inherit from Frame; support `embedded=True` to prevent window-level overrides
- Fonts: `Inter`/`Rajdhani` for body, `JetBrains Mono` for coordinates/numbers (width 6–8), `Orbitron` for headers
- Defer `matplotlib` rendering until tab is active (`<<NotebookTabChanged>>`)
- New functional modules → separate panel class → registered as tab in `SEEDApplication`

## Testing Conventions

- All tests mock serial/VISA — no physical hardware required
- Changes to `sender_panel.py` or `utils.py` must pass existing integration tests before merge
- After editing ≥3 lines of any launchable script, verify the app still launches
