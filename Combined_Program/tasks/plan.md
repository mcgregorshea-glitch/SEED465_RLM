# Implementation Plan: POP HTML Analyzer

## Overview

Build `tools/pop_analyzer.html` — a single self-contained offline HTML file that replicates
the post-scan analysis workflow of `POPVisualizationPanel`. All JS libraries (Plotly.js,
Papa Parse, jsPDF) are inlined. No server, no runtime network requests, no build tooling.

---

## Architecture Decisions

| Decision | Rationale |
|---|---|
| Plotly.js for all plot types | Handles scatter, heatmap, and contour natively; built-in zoom/pan/hover; `toImage()` for PDF export |
| Papa Parse for CSV | Handles comment-line skipping, type coercion, and edge cases better than hand-rolled parsing |
| jsPDF for PDF | Pure JS, no server round-trip; addImage() accepts PNG data URLs from Plotly.toImage() |
| Central `state` object | All rendering functions read from one source of truth; avoids scattered DOM queries |
| Pure filtering functions | `getFilteredData(state)` and `getSliceLabel(state)` are stateless — easy to call from both UI renders and PDF export |
| Change detection per DV | `lastRenderParams[dvField]` string-keyed cache skips expensive Plotly re-renders when nothing changed |
| Logo as inline base64 constant | Keeps portability guarantee; logo is stable; ~34KB is negligible |

---

## Dependency Graph

```
[Inlined libraries: Plotly, PapaParse, jsPDF]
    │
    ├── [Constants: logo b64, colormap defs, known-IV list]
    │
    └── [HTML skeleton + dark CSS]
            │
            └── [State object]
                    │
                    ├── [CSV parser] ──────────────────────────────┐
                    │                                              ▼
                    ├── [Axis selectors + DV checkboxes] ←── populated from state
                    │
                    ├── [IV slicer controls] ←──────────────────── depends on axis selectors
                    │
                    ├── [Scatter rendering] ←──────────────────── depends on slicers + state
                    │
                    ├── [Gradient rendering + mode toggle] ←───── extends scatter renderer
                    │
                    ├── [Colormap selector] ←─────────────────── affects all renders
                    │
                    ├── [PNG export] ←────────────────────────── depends on rendered plots
                    │
                    ├── [PDF modals: title + slice dialog] ←───── depends on slicer state
                    │
                    └── [PDF page compositor] ←────────────────── depends on modals + renderer
```

Implementation order follows this graph bottom-up.

---

## Task List

### Phase 1: Foundation

- [ ] Task 1: HTML scaffold — libraries, skeleton, dark CSS, state, constants
- [ ] Task 2: CSV parsing — file load, metadata, state population, axis/DV controls

### ✅ Checkpoint 1
- File opens in browser with no console errors
- Dropping a CSV populates X/Y selects and DV checkboxes

---

### Phase 2: Core Visualization

- [ ] Task 3: IV slicer controls — snap-to-discrete, axis↔slicer swap
- [ ] Task 4: Scatter rendering — 2-column plot grid, change detection, bounds locking
- [ ] Task 5: Gradient rendering — heatmap/contour/fallback, render mode toggle, aspect ratio

### ✅ Checkpoint 2
- Full IV slice → DV heatmap workflow works end-to-end in both render modes
- Bounds locked correctly when PARAMS present

---

### Phase 3: Colormaps + PNG Export

- [ ] Task 6: Colormap selector — 7 options, custom inferno-sliced default, applies to all renders
- [ ] Task 7: PNG export — per-plot 💾 button, white background, Plotly.downloadImage

### ✅ Checkpoint 3
- All 7 colormaps render correctly; PNG downloads are white-background
- Inferno-sliced matches Python visual output

---

### Phase 4: PDF Report

- [ ] Task 8: PDF modals — title prompt modal + slice selection dialog
- [ ] Task 9: PDF page compositor — A4 layout, header/logo/grid/footer, jsPDF output

### ✅ Checkpoint 4 — Final
- All spec success criteria satisfied
- Full offline verification (network disabled)
- PDF layout matches Python version

---

## Risks and Mitigations

| Risk | Impact | Mitigation |
|---|---|---|
| Plotly.js inlined size (~3.5 MB) makes file slow to open | Medium | Minified build; acceptable per spec (size not a concern) |
| `Plotly.toImage()` is async — PDF must await all plots | High | `Promise.all` over all DV × slice combinations before assembling pages |
| Contour trace fails on sparse/irregular data | Medium | Explicit try/catch; fallback to scatter on error |
| jsPDF `addImage` coordinate math for 2×3 grid | Medium | Hard-code A4 mm constants; test with 1, 2, 6, 7 active DVs |
| Custom inferno-sliced colorscale must visually match Python | Low | Generate from same `linspace(0.2, 1.0, 256)` formula in JS |

---

## Open Questions

None — all resolved in spec.
