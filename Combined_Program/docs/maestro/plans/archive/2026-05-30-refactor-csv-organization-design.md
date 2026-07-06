---
title: "Refactor CSV Data Organization"
created: "2026-05-31T00:00:00Z"
status: "draft"
authors: ["TechLead", "User"]
type: "design"
design_depth: "standard"
task_complexity: "medium"
---

# Refactor CSV Data Organization Design Document

## Problem Statement

The current CSV data logging format is visually disorganized and problematic for users to read manually in spreadsheet applications. It mixes metadata (IV/DV definitions) with data headers in a way that creates misaligned key-value pairs (e.g., `# SOURCE_IVS: x` mapping to `# SOURCE_DVS: DMM_1`). A new, visually clean format is required that satisfies manual readability requirements while maintaining robust programmatic parsing for the PoP Visualization Panel.

## Requirements

### Functional Requirements

1. **REQ-1**: The CSV logging format must be refactored to prioritize visual readability in spreadsheet applications.
2. **REQ-2**: The `sender_panel.py` must write the new format correctly during live scans.
3. **REQ-3**: The `pop_visualization_panel.py` must be updated to correctly parse the new format, distinguishing between metadata and actual data rows.

### Non-Functional Requirements

1. **REQ-N1**: The parsing logic must be robust against the presence of comment characters (`#`) within cell data.

### Constraints

- The new format must support an arbitrary number of Dependent Variables (DVs).

## Approach

### Selected Approach

**Pivoted Metadata Layout**

The CSV will be structured with metadata definitions occupying dedicated rows at the top, cleanly separated from the data block.

**Format Structure:**
```csv
# SOURCE_IVS:,x,y,z,rot
# SOURCE_DVS:,HUB_VAL_7
Timestamp,#(this is the end of the header, below this is the actual data values),x,y,z,rot,HUB_VAL_7
2026-05-30 12:00:00,,10.0,20.0,5.0,90.0,1.23
```

- **Row 1**: Explicitly lists all Independent Variables (IVs).
- **Row 2**: Explicitly lists all Dependent Variables (DVs).
- **Row 3**: The actual data header row, including the requested user comment.
- **Row 4+**: The measurement data.

This approach separates metadata from the data table, making it highly readable in Excel while providing clear anchor points for programmatic parsing.

### Alternatives Considered

#### Strict JSON/Dictionary Row Mapping
- **Description**: Storing each row as a flattened dictionary representation.
- **Pros**: Matches the exact JSON structure provided in the prompt.
- **Cons**: Extremely difficult to parse programmatically with `pandas`, creates misaligned columns.
- **Rejected Because**: Fails the requirement for robust programmatic parsing by the PoP panel.

### Decision Matrix

| Criterion | Weight | Pivoted Metadata | Strict JSON Mapping |
|-----------|--------|------------------|---------------------|
| Visual Readability | 50% | 5: Clean separation of metadata and data. | 2: Misaligned columns and empty headers. |
| Parsing Robustness | 50% | 4: Easy to skip initial rows and read headers. | 1: Breaks standard CSV parsing libraries. |
| **Weighted Total** | | **4.5** | **1.5** |

## Architecture

### Component Diagram

```
[Sender Panel] --(Writes new CSV format)--> [CSV File]
                                                |
                                                v
[PoP Panel] <--(Parses with custom header skip)--
```

### Data Flow

1. `sender_panel.py` initializes the CSV with the new 3-row metadata/header structure.
2. `sender_panel.py` appends data rows matching the column order.
3. `pop_visualization_panel.py` opens the CSV, manually reads the first two rows to extract `ivs` and `dvs`.
4. `pop_visualization_panel.py` uses `pandas.read_csv` with `skiprows=2` to load the actual data table.

### Key Interfaces

```python
# sender_panel.py -> _initialize_log_file
# Writer must use the new 3-row header format.

# pop_visualization_panel.py -> _process_file
# Parser must read lines 1 and 2 for metadata, then load dataframe from line 3.
```

## Agent Team

| Phase | Agent(s) | Parallel | Deliverables |
|-------|----------|----------|--------------|
| 1     | coder  | No       | Updated `sender_panel.py` logging logic. |
| 2     | coder  | No       | Updated `pop_visualization_panel.py` parsing logic. |

## Risk Assessment

| Risk | Severity | Likelihood | Mitigation |
|------|----------|------------|------------|
| Data loss during transition | HIGH | LOW | Ensure backward compatibility in parser or clear documentation that old logs are incompatible. |
| Pandas parsing failure due to `#` | HIGH | MEDIUM | Do not use `comment='#'` in pandas; instead, use `skiprows` to explicitly bypass the metadata block. |

## Success Criteria

1. New CSV files generated during a scan match the exact visual layout requested.
2. The PoP panel successfully loads and visualizes data from these new CSV files without errors.
