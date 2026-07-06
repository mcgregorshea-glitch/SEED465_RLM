---
title: "Live Proof of Power (PoP) Graphing Optimization"
created: "2026-05-31T19:45:00Z"
status: "approved"
authors: ["TechLead", "User"]
type: "design"
design_depth: "standard"
task_complexity: "medium"
---

# Live Proof of Power (PoP) Graphing Optimization Design Document

## Problem Statement
The Proof of Power (PoP) live graphing system in the SEED Control Center currently suffers from significant latency and UI responsiveness issues. A hardcoded 5-second refresh interval creates a "clunky" user experience, while heavy Matplotlib rendering on the main GUI thread causes intermittent freezes. Additionally, the lack of a dedicated live readout for specific data points forces users to rely on the visualizations to monitor test progress, making it difficult to verify precise values in real-time.

## Requirements
### Functional Requirements
1. **REQ-1 (Smart Refresh)**: Implement a ~1Hz live visualization update loop that only redraws plots if new data has arrived for that specific variable and the plot is currently selected for display.
2. **REQ-2 (Live Readout)**: Add a high-visibility Telemetry HUD at the top of the PoP Analysis tab displaying real-time coordinates (X, Y, Z, Rot) and all active DVs.
3. **REQ-3 (Adaptive Rendering)**: Automatically switch from `tricontourf` to fast scatter plotting when the cumulative points across all active plots exceed 1,000.
4. **REQ-4 (Terminology)**: Use "Proof of Power" and "PoP" interchangeably in all UI and code.

### Non-Functional Requirements
1. **UI Responsiveness**: Redraw operations must not block the main thread, ensuring E-STOP remains responsive.
2. **Efficiency**: Reinstate rendering cache and eliminate redundant code paths.

## Approach
### Selected Approach: Smart-Gated Rendering & Live Telemetry
We will refactor the `POPVisualizationPanel` to act as an efficient consumer of the `EventBus` using a dirty-flag system.
- **Selection Check**: Only render plots chosen via sidebar checkboxes.
- **Dirty Flags**: Redraw only when new data for a specific DV is received.
- **Adaptive Plotting**: Cumulative threshold of 1,000 points triggers high-speed scatter fallback.
- **HUD Layer**: Direct text updates for machine coordinates and telemetry values for zero-latency feedback.

## Architecture
The new architecture moves from a synchronous timer to an event-driven model.
- **Ingestion**: `MeasurementFrame` publishes to `EventBus: hub_telemetry`.
- **HUD**: Direct label updates bypass graphing logic.
- **Gated Renderer**: 1Hz loop checks `needs_redraw` flag.
- **Geometry Switcher**: Swaps rendering strategy based on data density.

## Agent Team
| Phase | Agent | Parallel | Deliverables |
|-------|-------|----------|--------------|
| 1 | `ux_designer` | No | Telemetry HUD Implementation |
| 2 | `performance_engineer` | No | Gated Rendering & Adaptive Plotting Refactor |
| 3 | `code_reviewer` | No | Final Quality & Compliance Audit |

## Risk Assessment
- **UI Deadlock**: Mitigated by strictly using `after()` for all UI dispatches.
- **Large Data Lag**: Mitigated by Adaptive Geometry fallback.
- **Constraint Violation**: Mitigated by explicit audit against `IMMUTABLE_CONSTRAINTS.md`.

## Success Criteria
1. Graph refresh < 500ms from data arrival.
2. Responsive E-STOP during active updates.
3. HUD values match CSV logs exactly.
4. Compliance with "Proof of Power" terminology.
