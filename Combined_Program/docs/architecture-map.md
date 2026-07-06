# SEED Control Center — Architecture Map

Seven diagrams covering the full system: application structure, hardware integration, threading, scan lifecycle, event bus, safety systems, and data output.

---

## 1 · Application Layer

Every panel is a `ttk.Frame` registered as a tab in the root `ttk.Notebook`. `SEEDApplication` owns the global **E-STOP overlay** — a canvas that floats above all tabs at all times.

```mermaid
flowchart TD
    ROOT["main.py\nSEEDApplication"]
    ESTOP["⬡ E-STOP Overlay\nglobal Canvas — always interactive\nclick → M112"]
    NB["ttk.Notebook"]

    ROOT --> ESTOP
    ROOT --> NB

    NB --> T1["Tab 1 · Pattern Generator\ngenerator_panel.py\nPatternGeneratorGUI"]
    NB --> T2["Tab 2 · G-Code Sender\nsender_panel.py\nGCodeSenderGUI"]
    NB --> T3["Tab 3 · Vivigo Hub\nvivigo_panel.py\nVivigoPanel"]
    NB --> T4["Tab 4 · Manual Movement\nmanual_movement_panel.py\nManualMovementPanel"]

    T3 --> VNBK["Internal ttk.Notebook"]
    VNBK --> VS1["Subtab · Vivigo Controls\nlegacy GUI  ←  rlm_proj_vivigo.py"]
    VNBK --> VS2["Subtab · PoP Data Viz\npop_visualization_panel.py\nPOPVisualizationPanel"]

    T2 --> SC["sender_components/"]
    SC --> SE["serial_engine.py\nSerial connection + G-code stream\nconnection thread / sender thread"]
    SC --> DM["dmm_manager.py\nDmmInst · DmmGroup\npyvisa SCPI over TCP/IP"]
    SC --> GP["gcode_processor.py\nrelative→absolute translation\nbounds pre-validation"]
    SC --> MC["motion_controls.py\njog · home · go-to-position"]
    SC --> PM["plot_manager.py\nmatplotlib 3D toolpath\nlazy-loaded on tab activate"]
    SC --> TS["telemetry_service.py\nEventBus subscriber\nDMMTelemetryProvider\nVivigoTelemetryProvider"]
    SC --> UL["ui_layout.py\nEventBus publication loop\nWidget layout classes\nHeaderBar · FooterBar · etc."]

    UTILS["utils.py  ·  HAL\nEventBus · PRINTER_BOUNDS\nPRINTER_LIMITS · UI Theme"]
    T1 -. uses .-> UTILS
    T2 -. uses .-> UTILS
    T3 -. uses .-> UTILS
```

---

## 2 · Hardware Integration

```mermaid
flowchart LR
    subgraph SRC["Application"]
        SE2["SerialEngine"]
        DM2["DmmGroup"]
        VP["VivigoPanel\nlegacy GUI"]
    end

    subgraph HW["Physical Hardware"]
        PR["3D Printer\nEnder 3\nMarlin firmware"]
        D1["DMM 1\n192.168.0.x"]
        D2["DMM 2\n192.168.0.x"]
        D3["DMM 3\n192.168.0.x"]
        D4["DMM 4\n192.168.0.x"]
        HUB["Vivigo Hub\n(Device Under Test)"]
    end

    subgraph SIM["Simulator  sim/"]
        FP["fake_printer.py\nTkinter state display"]
        SS["serial_server.py\nVirtual COM port"]
        ML["marlin_logic.py\nMarlin response emulator"]
        FP --> ML --> SS
    end

    SE2 -->|"G-code lines  115200 baud  USB/Serial"| PR
    PR -->|"ok · T:xx · M119 endstop states"| SE2

    DM2 -->|"SCPI: CONF:VOLT:DC\nSAMP:COUN N\nINIT / READ?"| D1
    DM2 -->|"SCPI"| D2
    DM2 -->|"SCPI"| D3
    DM2 -->|"SCPI"| D4
    D1 & D2 & D3 & D4 -->|"averaged float readings"| DM2

    VP -->|"RLM API  COM port"| HUB
    HUB -->|"v_rec · v_sys · v_bst\ni_sys · i_bst · t_bst..."| VP

    SS -->|"Marlin responses"| SE2
```

---

## 3 · Threading Model

Only the **main thread** may touch Tkinter widgets. All background threads communicate exclusively through `message_queue`. The `serial_lock` serializes every byte written to the physical COM port.

```mermaid
flowchart TD
    subgraph MAIN["Main Thread — Tkinter event loop"]
        TK["Tkinter / CTk mainloop()"]
        CMQ["check_message_queue()\nroot.after(100ms) loop\ndrains queue, applies UI updates"]
        TK --> CMQ
    end

    subgraph BG["Background Threads  all daemon=True"]
        CT["Connection Thread\n_connect_thread()\nM105 × 3 handshake"]
        ST["G-Code Sender Thread\ngcode_sender_thread()\nmain scan loop"]
        HT["Homing Thread\n_homing_sequence_worker()"]
        MT["Measure Thread\n_measure_thread()\nmanual trigger only"]
    end

    MQ["queue.Queue\nmessage_queue\nthread-safe bridge"]
    SL["threading.Lock\nserial_lock\none writer at a time"]
    EVS["threading.Event\nstop_event  ·  pause_event\ncancel_connect_event"]

    CT -->|"CONNECT_OK / CONNECT_FAIL / LOG"| MQ
    ST -->|"LOG / MEASUREMENT / SCAN_COMPLETE\nSCAN_POSITION / GATED_MEASUREMENT"| MQ
    HT -->|"LOG / HOMING_DONE"| MQ
    MT -->|"MEASUREMENT"| MQ

    MQ --> CMQ

    CT & ST & HT & MT -->|"acquires before every serial write"| SL
    EVS -->|"stop_event.set() — aborts sender loop"| ST
    EVS -->|"pause_event.wait() — suspends sender loop"| ST
    EVS -->|"cancel_connect_event — aborts handshake"| CT
```

---

## 4 · Live Scan Lifecycle (Sequence)

One full automated scan from button-press to completion.

```mermaid
sequenceDiagram
    participant U  as User
    participant GUI as GCodeSenderGUI
    participant GP2 as GcodeProcessor
    participant SE3 as SerialEngine
    participant PR2 as Printer
    participant DM3 as DmmGroup
    participant EB  as EventBus
    participant POP as POPVisualizationPanel

    U->>GUI: click "Start Sending"
    GUI->>GP2: process_gcode() — translate + validate bounds
    GUI->>EB: publish("scan_started", {})
    EB-->>POP: _on_scan_started() — clear live buffer + extents
    GUI->>SE3: start gcode_sender_thread()

    loop Every G0/G1 line in file
        SE3->>PR2: G1 X___ Y___ Z___ F___
        PR2-->>SE3: ok
        SE3->>EB: publish("scan_position", {x, y, z, rot})
        EB-->>POP: _on_scan_position() — snap sliders, set needs_redraw

        alt Z value changed from previous line
            SE3->>PR2: G28 X Y  (XY rehome only — never Z)
            PR2-->>SE3: ok
            SE3->>SE3: _homing_verification_routine()
        end

        SE3->>SE3: sleep(settling_delay_s)
        SE3->>DM3: trigger() all connected DMMs simultaneously
        DM3-->>SE3: read() — blocks until all stable, returns averaged values
        SE3->>SE3: _log_measurement_to_file() → append row to CSV
        SE3->>EB: publish("gated_measurement", {x,y,z,rot,dmm_1,...,hub_fields...})
        EB-->>POP: _on_gated_measurement() — append data, expand _live_axis_extents
    end

    SE3->>EB: publish("scan_finished", {})
    EB-->>POP: stop_live_mode()
    GUI->>U: scan complete popup
```

---

## 5 · EventBus Pub / Sub Map

`EventBus` lives in `utils.py`. It is a simple synchronous pub/sub — subscribers are called directly on the publishing thread.

```mermaid
flowchart LR
    subgraph PUB["Publishers"]
        P1["GCodeSenderGUI\ngcode_sender_thread()"]
        P2["MeasurementFrame\n_update_telemetry_ui()\ncalled via root.after()"]
    end

    subgraph EVT["Event Channels"]
        E1(["scan_started\npayload: {}"])
        E2(["scan_position\npayload: {x, y, z, rot}"])
        E3(["gated_measurement\npayload: {coords + all DV fields}"])
        E4(["scan_finished\npayload: {}"])
        E5(["hub_telemetry\npayload: {v_rec, v_sys, v_bst,\ni_sys, i_bst, t_bst, x, y, z, rot}"])
    end

    subgraph SUB["Subscribers"]
        S1["POPVisualizationPanel\n_on_scan_started()\nclear live_session_data\nclear _live_axis_extents"]
        S2["POPVisualizationPanel\n_on_scan_position()\nsnap sliders to probe position"]
        S3["POPVisualizationPanel\n_on_gated_measurement()\nappend data point\nexpand axis extents\ntrigger 1Hz redraw"]
        S4["POPVisualizationPanel\nstop_live_mode()\ncancel live_tick loop"]
        S5["POPVisualizationPanel\n_on_hub_telemetry()\nupdate HUD labels only"]
        S6["VivigoTelemetryProvider\n_on_hub_data()\ncache packet for fetch_data()"]
    end

    P1 --> E1 --> S1
    P1 --> E2 --> S2
    P1 --> E3 --> S3
    P1 --> E4 --> S4
    P2 --> E5 --> S5
    P2 --> E5 --> S6
```

---

## 6 · Safety Architecture

```mermaid
flowchart TD
    subgraph ESTOP2["Emergency Stop"]
        HEX2["⬡ Hexagonal Canvas Overlay\nfloats above all tabs\ncreate_polygon + tag_bind\nplace() at top-right of root window"]
        M112_2["M112  Hard Stop\nfirmware kills all motors + heaters\nrequires power cycle to resume\n< 200 ms response target"]
        HEX2 -->|"click"| M112_2
    end

    subgraph QS["Quick Stop"]
        M410["M410  Soft Stop\nfinishes current move, then halts\nno power cycle required\nguard: checks winfo_ismapped()"]
    end

    subgraph BOUNDS2["Coordinate Guarding — two-stage"]
        PB2["PRINTER_BOUNDS  utils.py\nx_min 0  x_max 220 mm\ny_min 0  y_max 220 mm\nz_min 0  z_max 128 mm\nrot_min −90°  rot_max +90°"]
        STAGE1["Stage 1  Pre-flight\nGcodeProcessor validates\nentire file before first line sent"]
        STAGE2["Stage 2  Runtime\nSerialEngine re-validates\nevery G-code line on send"]
        PB2 --> STAGE1
        PB2 --> STAGE2
    end

    subgraph ZHOME["Z-Axis Protection"]
        ZXY["G28 X Y only\nnever G28 Z during a scan\nserial_engine.py:293  sender_panel.py:411\nNo Z max limit switch on this fixture"]
    end

    subgraph TLOCK["Thread Safety"]
        SERLOCK["serial_lock  threading.Lock\nprevents serial port contention\nduring HALT from any thread"]
        STOPEV["stop_event  threading.Event\naborting gcode_sender_thread\nimmediately on stop"]
        M112_2 -->|"acquires"| SERLOCK
        M112_2 -->|"sets"| STOPEV
    end
```

---

## 7 · Data Logging & Output Pipeline

```mermaid
flowchart TD
    subgraph SRC2["Sources at each scan stop"]
        IV["Independent Variables\ncommanded X · Y · Z · Rot\nfrom G-code line"]
        DMMSRC["DMM Readings\naveraged over N samples\nstability gate: StdDev < threshold\nvia pyvisa SCPI"]
        HUBSRC["Vivigo Hub Telemetry\nv_rec · v_sys · v_bst\ni_sys · i_bst · t_bst\nvia RLM API → EventBus"]
    end

    subgraph CSV2["CSV Log File  3-Row Format"]
        R1["Row 1: Metadata comment\n# PARAMS: {json scan profile}\n# SOURCE_IVS: x, y, z, rot\n# SOURCE_DVS: dmm_1, v_rec, ..."]
        R2["Row 2: Column header\nx, y, z, rot, Timestamp, dmm_1, ..."]
        R3["Row 3+: One data row per scan stop\nall IVs and DVs in a single flat record"]
        R1 --- R2 --- R3
    end

    subgraph VIZ2["Visualization"]
        LIVE["POPVisualizationPanel\nLive heatmaps during scan\n1 Hz gated refresh\nscatter → contourf auto-select\naxis bounds from _live_axis_extents"]
        WEB["visualizer.html\nPlotly.js  post-scan\nopen via Launch Data Visualizer btn"]
        PDF2["PDF Report\n_save_snapshot()\nA4 · 300 DPI\n2×3 plot grid · branded header\ncustom title dialog"]
        PNG["Single Plot PNG Export\n_export_single_plot()\nwhite background · 150 DPI"]
    end

    IV & DMMSRC & HUBSRC --> CSV2
    CSV2 -->|"_process_file()  load from disk"| LIVE
    CSV2 --> WEB
    LIVE -->|"Generate Report btn"| PDF2
    LIVE -->|"💾 per-plot btn"| PNG
```
