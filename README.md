# SheetForge: Web-based Excel Clone (Single-Page App)

SheetForge is a browser-based spreadsheet app that emulates a practical subset of Excel behavior while staying lightweight.

## What this clone supports

### Spreadsheet interaction
- 100 rows × 26 columns (`A` → `Z`) per sheet.
- Direct in-cell editing and formula bar editing.
- Keyboard navigation with arrows + `Enter`.

### Workbook + sheets
- Multiple sheets with bottom tab UI.
- Add new sheet (`+ Sheet`) and rename via double-click.
- Save/load workbook to `localStorage`.

### Full in-app function catalog
- **Functions** toolbar button opens searchable function catalog panel.
- Catalog includes syntax + short description for supported functions.
- Current supported families:
  - Aggregation/statistics: `SUM`, `AVG`/`AVERAGE`, `MIN`, `MAX`, `COUNT`, `COUNTA`, `MEDIAN`, `PRODUCT`
  - Math/trig/log: `ABS`, `ROUND`, `ROUNDUP`, `ROUNDDOWN`, `INT`, `CEILING`, `FLOOR`, `MOD`, `POW`, `SQRT`, `EXP`, `LN`, `LOG`, `LOG10`, `SIN`, `COS`, `TAN`, `ASIN`, `ACOS`, `ATAN`, `PI`, `RAND`, `RANDBETWEEN`
  - Logical: `IF`, `AND`, `OR`, `NOT`
  - Text: `LEN`, `UPPER`, `LOWER`, `TRIM`, `LEFT`, `RIGHT`, `MID`, `CONCAT`
  - Date/time: `TODAY`, `NOW`, `YEAR`, `MONTH`, `DAY`, `HOUR`, `MINUTE`, `SECOND`, `DATE`

### Formula engine (tokenizer + parser)
- Formula tokenizer + recursive-descent parser (replaces regex substitution logic).
- Arithmetic with precedence and parentheses.
- Cell references (`A1`) and range references (`A1:B10`).
- Circular guard fallback and `#ERR` for invalid formulas.

### Range-aware selection + fill handle
- Drag across cells to create rectangular range selections.
- Status bar shows numeric aggregate metrics over selected range.
- Fill handle appears at bottom-right of selection.
- Drag fill handle up/down/left/right to propagate formulas/values/styles.

### Advanced grid features
- Frozen panes: freeze top N rows and left N columns.
- Column resizing by dragging header edge.
- Undo/Redo stack via toolbar and keyboard shortcuts:
  - `Ctrl/Cmd+Z`
  - `Ctrl/Cmd+Y`
  - `Ctrl/Cmd+Shift+Z`

### Data interchange
- CSV import/export (active sheet).
- XLSX import/export via SheetJS (all workbook tabs for export).

---

## Quick start

```bash
python3 -m http.server 4173
```

Open `http://localhost:4173`.

> XLSX support depends on CDN loading of SheetJS. If offline, CSV still works.

---

## Current limitations

- Not full Excel parity (no pivot tables/charts/full function catalog parity with desktop Excel).
- No collaborative editing.
- XLSX roundtrip currently focuses on values/formulas, not full style fidelity.
