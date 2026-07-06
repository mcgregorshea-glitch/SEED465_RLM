---
title: "Relative Coordinate System for PoP Data Implementation Plan"
design_ref: "docs/maestro/plans/2026-05-31-live-pop-graphing-optimization-design.md"
created: "2026-05-31T21:00:00Z"
status: "approved"
total_phases: 2
estimated_files: 2
task_complexity: "simple"
---

# Relative Coordinate System for PoP Data Implementation Plan

## Objective
Offset all spatial data (X, Y, Z, Rot) by the "Mark Center" coordinates for both live Proof of Power (PoP) visualizations and CSV recording. This ensures that plots and logs are centered around the test subject rather than the printer's absolute origin.

## Key Files
- `Combined_Program/sender_panel.py`: Core logic for position updates and CSV logging.
- `Combined_Program/sender_components/ui_layout.py`: Telemetry polling and data injection.

## Implementation Details

### Phase 1: Offsetting in GCodeSenderGUI (`sender_panel.py`)
1.  **_handle_position_update**:
    *   Continue updating `self.last_cmd_abs_*` with absolute coordinates for DRO consistency.
    *   Create a `relative_pos` dictionary by subtracting the current center variables (`center_x_var`, etc.) from the incoming position.
    *   Publish the `relative_pos` to the `scan_position` EventBus topic.
2.  **_on_measurement**:
    *   Before calling `_log_measurement_to_file`, transform the `coords` dictionary to relative coordinates by subtracting the center values.

### Phase 2: Offsetting in MeasurementFrame (`ui_layout.py`)
1.  **_fetch_telemetry_data**:
    *   When injecting spatial coordinates (`x`, `y`, `z`, `rot`) into the telemetry packet, subtract the `gui.center_x_var` etc. values.
    *   This ensures the live HUD and plots receive relative data during 1Hz polling.

## Verification
1.  **CSV Check**: Perform a manual measurement or scan. Verify that the recorded X/Y/Z values in the CSV are relative to the "Mark Center" location.
2.  **HUD/Plot Check**: Open the PoP Analysis tab. Verify that the "X-POS", "Y-POS", etc. in the HUD show values relative to center (e.g., 0.00 when at center).
3.  **DRO Check**: Verify the main Sender DRO still works correctly in both Absolute and Relative modes.
