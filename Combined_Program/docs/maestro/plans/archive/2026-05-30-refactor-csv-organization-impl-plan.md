---
title: "Refactor CSV Data Organization Implementation Plan"
design_ref: "docs/maestro/plans/2026-05-30-refactor-csv-organization-design.md"
created: "2026-05-31T00:00:00Z"
status: "draft"
total_phases: 3
estimated_files: 2
task_complexity: "medium"
---

# Refactor CSV Data Organization Implementation Plan

## Plan Overview

- **Total phases**: 3
- **Agents involved**: coder, tester
- **Estimated effort**: Moderate. Requires careful formatting of the CSV writer and parser to ensure alignment.

## Dependency Graph

```
[Phase 1: Update CSV Writer]
          |
          v
[Phase 2: Update CSV Parser]
          |
          v
[Phase 3: Verification]
```

## Execution Strategy

| Stage | Phases | Execution | Agent Count | Notes |
|-------|--------|-----------|-------------|-------|
| 1     | Phase 1 | Sequential | 1 | Foundation (Data generation) |
| 2     | Phase 2 | Sequential | 1 | Data consumption |
| 3     | Phase 3 | Sequential | 1 | Verification |

## Phase 1: Update CSV Writer

### Objective
Update `sender_panel.py` to write the new 3-row metadata/header structure and align data rows.

### Agent: coder
### Parallel: false

### Files to Modify

- `Combined_Program/sender_panel.py` — Update `_initialize_log_file` and `_log_measurement_to_file`.
  - In `_initialize_log_file`:
    - Row 1: `# SOURCE_IVS:,` followed by the IV keys.
    - Row 2: `# SOURCE_DVS:,` followed by the DV keys.
    - Row 3 (Headers): `Timestamp,#(this is the end of the header, below this is the actual data values),` followed by IV and DV keys.
    - Instantiate the `csv.DictWriter` using these explicit headers.
  - In `_log_measurement_to_file`:
    - Ensure the row dict includes an empty string for the comment column.

### Implementation Details
The `csv.DictWriter` needs a fieldnames array like: `['Timestamp', 'Comment'] + iv_keys + dv_keys`. The actual data rows will just set `'Comment': ''`.

### Validation
- `python -m py_compile Combined_Program/sender_panel.py`

### Dependencies
- Blocked by: None
- Blocks: [2]

---

## Phase 2: Update CSV Parser

### Objective
Update `pop_visualization_panel.py` to parse the new 3-row structure without failing on the `#` character in the header.

### Agent: coder
### Parallel: false

### Files to Modify

- `Combined_Program/pop_visualization_panel.py` — Update `_process_file`.
  - Change metadata reading: explicitly read line 1 (IVs) and line 2 (DVs). The format is now `# SOURCE_IVS:,x,y...` instead of `# SOURCE_IVS: x`.
  - Change `pd.read_csv` call: remove `comment='#'` to prevent it from ignoring the third row's comment. Instead, use `skiprows=2` to skip the metadata rows and treat row 3 as the header.

### Implementation Details
The parser must extract IVs and DVs by splitting by commas after skipping the first column. For example, `ivs = set(line.strip().split(',')[1:])`.

### Validation
- `python -m py_compile Combined_Program/pop_visualization_panel.py`

### Dependencies
- Blocked by: [1]
- Blocks: [3]

---

## Phase 3: Verification

### Objective
Verify the new CSV format is written and parsed correctly.

### Agent: tester
### Parallel: false

### Implementation Details
Run the integration tests and perform a quick test script to verify `sender_panel.py` and `pop_visualization_panel.py` interoperate correctly.

### Validation
- `python -m unittest tests/test_main_integration.py`

### Dependencies
- Blocked by: [2]
- Blocks: None

---

## File Inventory

| # | File | Phase | Purpose |
|---|------|-------|---------|
| 1 | `Combined_Program/sender_panel.py` | 1 | CSV Writer update |
| 2 | `Combined_Program/pop_visualization_panel.py` | 2 | CSV Parser update |

## Risk Classification

| Phase | Risk | Rationale |
|-------|------|-----------|
| 1     | LOW | Simple format change. |
| 2     | MEDIUM | Pandas parsing can be brittle if the CSV format deviates. |
| 3     | LOW | Standard verification. |

## Execution Profile

```
Execution Profile:
- Total phases: 3
- Parallelizable phases: 0
- Sequential-only phases: 3
- Estimated parallel wall time: N/A
- Estimated sequential wall time: 3 mins

Note: Native subagents currently run without user approval gates.
All tool calls are auto-approved without user confirmation.
```
