---
title: "Live Proof of Power (PoP) Graphing Optimization Implementation Plan"
design_ref: "docs/maestro/plans/2026-05-31-live-pop-graphing-optimization-design.md"
created: "2026-05-31T20:00:00Z"
status: "draft"
total_phases: 3
estimated_files: 2
task_complexity: "medium"
---

# Live Proof of Power (PoP) Graphing Optimization Implementation Plan

## Plan Overview
This plan implements a "Smart Refresh" architecture for the PoP Analysis panel, reducing UI latency and improving responsiveness through gated rendering and adaptive geometry. It also adds a high-visibility Telemetry HUD for real-time monitoring.

- **Total phases**: 3
- **Agents involved**: `ux_designer`, `performance_engineer`, `code_reviewer`
- **Estimated effort**: Medium. Refactoring core rendering loops and adding UI components.

## Dependency Graph
```text
Phase 1: Telemetry HUD (ux_designer)
       |
       v
Phase 2: Gated Rendering & Adaptive Plotting (performance_engineer)
       |
       v
Phase 3: Final Audit & Terminology Compliance (code_reviewer)
```

## Execution Strategy
| Stage | Phases  | Execution | Agent Count | Notes |
|-------|---------|-----------|-------------|-------|
| 1     | Phase 1 | Sequential| 1           | UI Foundation |
| 2     | Phase 2 | Sequential| 1           | Logic Refactor |
| 3     | Phase 3 | Sequential| 1           | Quality Gate |

---

## Phase 1: Live Telemetry HUD
### Objective
Implement a high-visibility "Heads-Up Display" (HUD) at the top of the PoP Analysis tab to display real-time machine coordinates and telemetry values.

### Agent: ux_designer
### Parallel: No

### Files to Modify
- `Combined_Program/pop_visualization_panel.py`: 
    - Implement `_create_hud()` in `_setup_ui()`.
    - Implement `_update_hud(data)` to be called from `_on_hub_telemetry`.
    - Style the HUD using `Rajdhani` and `JetBrains Mono` for a high-fidelity look.

### Implementation Details
- Create a horizontal frame at the top of the panel (below the header).
- Add labels for X, Y, Z, Rot and all active DVs.
- Ensure the HUD updates immediately upon receipt of `hub_telemetry` events, bypassing any graphing delays.

### Validation
- Launch the application.
- Open PoP Analysis tab.
- Verify HUD appears and updates in sync with the live telemetry from the Vivigo tab.

### Dependencies
- Blocked by: None
- Blocks: Phase 2

---

## Phase 2: Gated Rendering & Adaptive Plotting
### Objective
Refactor the PoP rendering loop to use dirty-flags for efficiency and switch to scatter plots for large datasets.

### Agent: performance_engineer
### Parallel: No

### Files to Modify
- `Combined_Program/pop_visualization_panel.py`:
    - Refactor `_on_hub_telemetry` to set a `needs_redraw` flag.
    - Update `_live_tick` to run at 1Hz and only call `_update_all_plots` if `needs_redraw` is True.
    - Implement `_get_render_strategy()` based on cumulative point counts.
    - Update `_draw_heatmap` to support `ax.scatter` fallback above the 1,000-point threshold.
    - Remove the duplicate `_draw_heatmap` method currently located at L668.

### Implementation Details
- Add `self.needs_redraw = False` to `__init__`.
- Use `sum(len(data) for field in active_plots)` to determine the rendering strategy.
- Ensure `ax.scatter` uses the same colormap ('inferno') as the contour plots for visual consistency.

### Validation
- Run a scan simulation.
- Verify the CPU usage is low when data is stationary.
- Verify the transition to scatter points occurs at exactly 1,000 cumulative points.
- Confirm E-STOP remains perfectly responsive during 1Hz redraws.

### Dependencies
- Blocked by: Phase 1
- Blocks: Phase 3

---

## Phase 3: Final Audit & Quality Gate
### Objective
Perform a final code review for thread safety, terminology compliance, and adherence to immutable constraints.

### Agent: code_reviewer
### Parallel: No

### Files to Modify
- None. (Review and Audit only)

### Implementation Details
- Verify all labels use "Proof of Power" or "PoP" interchangeably as per REQ-4.
- Confirm that NO changes were made to `generator_panel.py` (Pattern Generator blanket restriction).
- Audit all `after()` and EventBus calls for thread safety.

### Validation
- All success criteria from the design document are met.
- `IMMUTABLE_CONSTRAINTS.md` is strictly followed.

### Dependencies
- Blocked by: Phase 2
- Blocks: None

---

## File Inventory
| # | File | Phase | Purpose |
|---|------|-------|---------|
| 1 | `Combined_Program/pop_visualization_panel.py` | 1, 2 | Main implementation surface for HUD and Rendering. |

## Risk Classification
| Phase | Risk | Rationale |
|-------|------|-----------|
| 1 | LOW | UI-only change; low impact on core logic. |
| 2 | MEDIUM | Touches core rendering loop; potential for regression in plot accuracy. |
| 3 | LOW | Non-destructive audit. |

## Execution Profile
```text
Execution Profile:
- Total phases: 3
- Parallelizable phases: 0
- Sequential-only phases: 3
- Estimated parallel wall time: N/A
- Estimated sequential wall time: 10 mins

Note: Native subagents currently run without user approval gates.
All tool calls are auto-approved without user confirmation.
```
