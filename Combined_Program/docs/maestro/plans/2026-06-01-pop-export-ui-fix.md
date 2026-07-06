# Implementation Plan: PoP Export UI Fix and Smart Selection

The `SliceSelectionDialog` in the PoP panel is currently broken (checkboxes don't render) and lacks "smart" defaults for slice selection.

## 1. Problem Diagnosis
- **Rendering Bug:** The `tk.LabelFrame` (individual slicer section) uses `btn_row.pack()` for the All/None buttons and `cb.grid()` for the checkboxes. Tkinter does not allow mixing `pack` and `grid` in the same parent; this causes the parent to hang or child widgets to disappear.
- **Selection Logic:** Currently, all checkboxes default to `True`. The user wants the report to default to what they are "currently observing", which means checking only the values currently selected by the panel's sliders.

## 2. Proposed Changes

### `src/pop_visualization_panel.py`

#### A. Fix `SliceSelectionDialog` Layout
- Inside the loop that creates `LabelFrame` for each field:
    - Keep `btn_row.pack()` for the top buttons.
    - Create a new `cb_frame = tk.Frame(frame, ...)` and `pack` it below `btn_row`.
    - Grid all checkboxes into `cb_frame` instead of `frame`.

#### B. Implement Smart Default Selection
- Update the `BooleanVar` initialization:
    - Retrieve the current value of the slicer from the `var` (first element of the `slider_vars` tuple).
    - Set the `BooleanVar` value to `True` only if the checkbox value matches the current slicer value (within a small tolerance).

## 3. Verification Plan
- **Visual Check:** Open the "Generate Report" dialog and confirm that checkboxes for all unique slicer values are visible and gridded correctly (4 columns).
- **Functional Check:** 
    - Confirm that only the checkbox matching the *current* PoP panel slider value is initially checked.
    - Confirm that "All" and "None" buttons still work.
    - Generate a small report to ensure the filtering logic correctly uses the selected slices.
