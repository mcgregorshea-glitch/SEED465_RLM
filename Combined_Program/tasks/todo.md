# Task List: POP HTML Analyzer
Output: `tools/pop_analyzer.html`
Spec: `docs/spec-pop-html-analyzer.md`

---

## Task 1: HTML scaffold — libraries, skeleton, dark CSS, state, constants

**Description:** Create `tools/pop_analyzer.html` with the full static structure. Inline minified
Papa Parse, Plotly.js, and jsPDF as `<script>` blocks. Add the dark-theme CSS. Define the central
`state` object, the known-IV list, and all colormap definitions including the custom inferno-sliced
array. Embed the RLM logo as a base64 constant. Wire the `Load CSV` file input.

**Acceptance criteria:**
- [ ] File opens in Chrome/Firefox/Edge with no console errors and no network requests
- [ ] All three libraries are accessible (`Papa`, `Plotly`, `window.jspdf.jsPDF`)
- [ ] Dark theme renders correctly (sidebar, header bar, plot area visible but empty)
- [ ] `LOGO_B64` constant contains valid base64 PNG data URI
- [ ] `COLORMAPS` constant defines all 7 options; `'inferno_sliced'` is the default key
- [ ] `state` object defined with all fields initialized to empty/null

**Verification:**
- [ ] Open file in browser → no red console errors
- [ ] `typeof Plotly` → `'object'` in console
- [ ] `typeof Papa` → `'object'` in console
- [ ] `typeof window.jspdf.jsPDF` → `'function'` in console
- [ ] Network tab shows zero requests after load

**Dependencies:** None

**Files:** `tools/pop_analyzer.html` (create)

**Estimated scope:** M

---

## Task 2: CSV parsing — file load, metadata extraction, state population, axis/DV controls

**Description:** Implement the full CSV load pipeline. `FileReader` reads the file text; a pre-parse
pass extracts `# SOURCE_IVS`, `# SOURCE_DVS`, and `; PARAMS:` lines. Papa Parse ingests the remaining
data rows. Rows with non-finite values in any IV or DV column are dropped. State is populated with
`data`, `ivFields`, `dvFields`, `bounds`. Axis `<select>` elements and DV checkboxes are rebuilt from
state. Filename shown in the source label.

**Acceptance criteria:**
- [ ] Annotated CSV (with `# SOURCE_IVS/DVS` headers): IVs and DVs classified exactly per header
- [ ] Plain CSV (no metadata headers): IVs auto-detected from known-IV list; DVs are the rest (minus Timestamp)
- [ ] `; PARAMS:` JSON parsed and stored in `state.bounds` with keys `{x_min, x_max, ...}`
- [ ] Rows with any non-finite IV or DV value silently dropped
- [ ] X-Axis defaults to first IV; Y-Axis defaults to second IV
- [ ] DV checkboxes all checked on load
- [ ] Source filename label updated

**Verification:**
- [ ] Load `data/main.260530_13-38-48.csv` (hub telemetry, no spatial IVs) → auto-classifies, no crash
- [ ] Load an annotated scan log CSV → axis selects show `x, y, z, rot`; DVs show measurement columns
- [ ] Load CSV with a row containing a non-numeric cell → that row absent from `state.data`

**Dependencies:** Task 1

**Files:** `tools/pop_analyzer.html`

**Estimated scope:** M

---

## ✅ Checkpoint 1

- [ ] File opens with no console errors, zero network requests
- [ ] Loading an annotated scan CSV populates X/Y selects and DV checkboxes correctly
- [ ] Loading a plain CSV falls back to auto-classification without crashing

---

## Task 3: IV slicer controls — snap-to-discrete, axis↔slicer swap

**Description:** Build the DATA SLICING section of the sidebar. For each IV not currently assigned
as X or Y axis, render a slicer row: label, range input (snapped to discrete values), and a linked
`<select>`. The two controls are bidirectionally synced — moving the slider updates the select and
vice versa. When the X or Y axis dropdown changes, rebuild the slicer list: the vacated IV becomes
a slicer, the newly claimed IV is removed from slicers. Slicer change triggers `updateAllPlots()`.

**Acceptance criteria:**
- [ ] One slicer row per IV not in X/Y axes
- [ ] Slider snaps: releasing between two discrete values picks the nearest
- [ ] Moving slider updates select; picking from select updates slider
- [ ] Changing X-Axis dropdown rebuilds slicer list correctly
- [ ] `getFilteredData(state)` returns only rows matching all current slicer values (±1e-4 tolerance)
- [ ] `getSliceLabel(state)` returns `"z=5mm, rot=0°"` style string

**Verification:**
- [ ] Load 4-IV scan CSV → 2 slicers appear (z and rot)
- [ ] Move z slider to a new layer → select updates → `getFilteredData` returns only rows at that z
- [ ] Change X-Axis from x → z → z slicer disappears; x slicer appears

**Dependencies:** Task 2

**Files:** `tools/pop_analyzer.html`

**Estimated scope:** M

---

## Task 4: Scatter rendering — 2-column plot grid, change detection, bounds locking

**Description:** Implement `updateAllPlots()` and per-DV `renderPlot(dvField)`. For each checked DV,
create a plot card in the 2-column grid (if it doesn't exist) or update it. Use Plotly scatter trace
(`mode: 'markers'`, `marker.color = zValues`, default colormap). Apply bounds locking: if
`state.bounds` has ranges for the selected axes, pass them as `layout.xaxis.range` and
`layout.yaxis.range`. Set aspect ratio `equal` when both axes are spatial (x/y/z), `auto` otherwise.
Implement change detection: skip re-render if `(dataLen, xCol, yCol, sliceStr, renderMode, colormap)`
matches the cached params for that DV.

**Acceptance criteria:**
- [ ] One plot card per checked DV, arranged in a 2-column grid
- [ ] Each card has a title bar with the DV name and a `💾 PNG` button placeholder
- [ ] Scatter points colored by DV value using the current colormap
- [ ] Colorbar present on each plot
- [ ] If `state.bounds` present: axis ranges locked; slicing through layers does not rescale axes
- [ ] If `state.bounds` absent: Plotly auto-ranges
- [ ] Aspect ratio `equal` for x-vs-y; `auto` for x-vs-rot
- [ ] Unchecking a DV checkbox removes its card; re-checking adds it back

**Verification:**
- [ ] Load scan CSV → check two DVs → two scatter plots appear
- [ ] Move a slicer → plots update; axes do not rescale when PARAMS present
- [ ] Uncheck DV → card gone; check again → card re-appears in same column position

**Dependencies:** Task 3

**Files:** `tools/pop_analyzer.html`

**Estimated scope:** M

---

## Task 5: Gradient rendering — heatmap/contour/fallback, render mode toggle, aspect ratio

**Description:** Add the `· Dot View` ↔ `≋ Gradient View` toggle button to the header. In Gradient
mode, apply the three-tier rendering strategy: regular grid → `heatmap` trace; irregular → `contour`
trace (try/catch); <10 points or exception → fall back to scatter. All plots switch mode together.
Change detection key includes `renderMode`.

**Acceptance criteria:**
- [ ] Toggle button present in header; label changes on click
- [ ] Gradient mode with regular-grid data → Plotly heatmap trace (smooth interpolation visible)
- [ ] Gradient mode with irregular data → Plotly contour trace
- [ ] <10 filtered points → falls back to scatter in both modes
- [ ] Exception in contour → falls back to scatter, no console error (warn only)
- [ ] Switching modes triggers re-render of all active plots

**Verification:**
- [ ] Load a regular-grid scan CSV → switch to Gradient → heatmap trace appears (verify via Plotly trace type)
- [ ] Filter to a slice with <10 points → no crash; scatter appears with message or sparse dots
- [ ] Toggle Dot → Gradient → Dot → plots update correctly each time

**Dependencies:** Task 4

**Files:** `tools/pop_analyzer.html`

**Estimated scope:** S

---

## ✅ Checkpoint 2

- [ ] Full workflow: load CSV → adjust slicers → view scatter and gradient heatmaps
- [ ] Bounds locked when PARAMS header present
- [ ] Axis↔slicer swap works correctly
- [ ] No console errors during normal operation

---

## Task 6: Colormap selector — 7 options, custom inferno-sliced default

**Description:** Add the Colormap `<select>` to the header bar. Populate from `COLORMAPS` constant.
Default to `inferno_sliced`. On change, update `state.colormap`, clear change-detection cache, and
re-render all active plots. The custom inferno-sliced scale is a 256-stop Plotly colorscale array
computed from `linspace(0.2, 1.0, 256)` mapped through the Inferno palette RGB values, matching the
Python `ListedColormap(inferno(linspace(0.2, 1.0, 256)))` output.

**Acceptance criteria:**
- [ ] Dropdown shows all 7 options; Inferno (sliced) selected by default
- [ ] Switching colormap re-renders all active plots immediately
- [ ] Custom inferno-sliced scale: darkest stop corresponds to ~20% of raw inferno (mid-purple, not black)
- [ ] All 7 options produce visually distinct, correctly-scaled colorbars

**Verification:**
- [ ] Load CSV → switch through all 7 options → all plots re-render, no errors
- [ ] Side-by-side comparison: inferno-sliced HTML plot vs Python screenshot — dark end is purple, not black
- [ ] RdYlBu produces a diverging scale (red → yellow → blue)

**Dependencies:** Task 4

**Files:** `tools/pop_analyzer.html`

**Estimated scope:** S

---

## Task 7: PNG export — per-plot 💾 button, white background

**Description:** Wire the `💾 PNG` button on each plot card. On click, call
`Plotly.downloadImage(plotDiv, { format: 'png', filename: 'POP_{dvField}_{timestamp}',
width: 1200, height: 900 })` with a white paper and plot background override. The downloaded PNG
should have a white background regardless of the dark UI theme.

**Acceptance criteria:**
- [ ] Clicking `💾` on a plot card triggers a PNG download
- [ ] Downloaded PNG has white background (not dark)
- [ ] Filename: `POP_{DV_NAME}_{YYMMDD_HHMM}.png`
- [ ] Colorbar and axis labels visible in export

**Verification:**
- [ ] Click 💾 on a loaded plot → file downloaded to system
- [ ] Open PNG → white background, colorbar visible, plot title correct

**Dependencies:** Task 4

**Files:** `tools/pop_analyzer.html`

**Estimated scope:** XS

---

## ✅ Checkpoint 3

- [ ] All 7 colormaps render correctly; inferno-sliced is visually correct
- [ ] PNG export produces white-background images with correct filenames
- [ ] No regressions in scatter/gradient rendering

---

## Task 8: PDF modals — title prompt modal + slice selection dialog

**Description:** Implement two inline HTML modals (not `window.prompt`):

**Title modal:** Text input pre-filled with `"SEED POP Analysis"`, OK/Cancel buttons.
Returns a Promise resolving to the entered string or null on cancel.

**Slice selection dialog:** One collapsible section per slicer variable. Checkboxes for each
discrete value (label format: `{val:.4g}{unit}`). Default: only the current slider value checked
(smart selection). `[All]` / `[None]` per section. `Generate` / `Cancel` buttons. Returns a
Promise resolving to `{ field: [selectedValues] }` or null on cancel.

**Acceptance criteria:**
- [ ] Title modal appears centered; Enter key confirms; Escape cancels
- [ ] Slice dialog sections labeled by field name (uppercase), with unit hints
- [ ] Default selection = only current slider value for each field
- [ ] `[All]` selects all values for that field; `[None]` deselects all
- [ ] `Generate` disabled if no values selected for any field
- [ ] Both modals return Promises; `null` on cancel halts PDF generation

**Verification:**
- [ ] Click Generate Report → title modal appears → press Enter → slice dialog appears
- [ ] Slice dialog: default checks match current slider positions
- [ ] Cancel in either modal → no PDF generated, no error

**Dependencies:** Task 3

**Files:** `tools/pop_analyzer.html`

**Estimated scope:** M

---

## Task 9: PDF page compositor — A4 layout, header/logo/grid/footer, jsPDF output

**Description:** Implement `generateReport()`. After modals resolve, iterate over all slice
combinations (`itertools.product` equivalent in JS). For each combination:

1. Filter data to that slice
2. For each page (6 DVs per page): create a jsPDF page
3. Draw header: report title (bold), "Proof of Power Report", date (left); RLM logo (right, from `LOGO_B64`)
4. Draw horizontal rule below header
5. Draw source filename and centered slice label
6. Render each DV plot via `Plotly.toImage({format:'png', width:500, height:380, ...})` with white background
7. Place plot images in 2×3 grid using `doc.addImage()`
8. Draw footer rule + right-aligned `Page X of N`
9. `doc.save('{SafeTitle}_{YYMMDD_HHMM}.pdf')`

All `Plotly.toImage` calls are async — collect with `Promise.all` before assembling pages.
Show a loading indicator during generation (disable Generate button, show spinner text).

**Acceptance criteria:**
- [ ] PDF downloaded with correct filename format
- [ ] Header: title bold top-left; logo top-right; rule below
- [ ] Source filename on its own line; slice label centered below
- [ ] Up to 6 plots per page in 2×3 grid; overflow onto additional pages
- [ ] Each plot: white background, DV name as title, colorbar visible
- [ ] Footer: horizontal rule + `Page X of N` right-aligned
- [ ] Multiple slice combinations → multiple sets of pages
- [ ] Generate button re-enabled after download; loading state cleared

**Verification:**
- [ ] Generate report with 2 DVs, 1 slice → 1-page PDF, 2 plots in row 1
- [ ] Generate report with 7 DVs, 1 slice → 2-page PDF (6 + 1)
- [ ] Generate report with 2 slicer variables, 2 values each → 4 slice combinations → 4 pages (assuming 1 DV)
- [ ] Open PDF → visually compare header/grid/footer layout against Python version screenshot

**Dependencies:** Tasks 7, 8

**Files:** `tools/pop_analyzer.html`

**Estimated scope:** L

---

## ✅ Checkpoint 4 — Final

- [ ] All spec success criteria checked off:
  - [ ] Opens offline in Chrome/Firefox/Edge, no console errors
  - [ ] Annotated and plain CSVs load correctly
  - [ ] Sliders snap to discrete values; plots update
  - [ ] All 7 colormaps work; inferno-sliced is default
  - [ ] Dot and Gradient modes both produce correct plots
  - [ ] PNG export: white background, correct filename
  - [ ] PDF report: header/logo/2×3 grid/slice label/footer/page count matches Python layout
  - [ ] PDF slice dialog pre-selects current slider values
  - [ ] Fully offline (verified with network disabled)
- [ ] Human review of PDF output against Python reference
