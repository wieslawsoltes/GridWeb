# Reusable editing contracts · 0.6.0

All mutations use the existing workbook model, validation/protection rules and one atomic undoable transaction. No DOM is needed for the core functions. Operations enforce the 250,000-cell bound; special-cell queries can scan sparse full-column references when their query does not require visiting every blank.

```js
import {Workbook, captureRange, pasteSpecial, fillSeries, specialCells} from '@wieslawsoltes/gridweb';
const book = new Workbook();
const sheet = book.ActiveWorksheet;
sheet.GetRange('A1:B2').Values = [[1,2],[3,4]];
const copy = captureRange(sheet.GetRange('A1:B2')); // or range.Capture()
const target = pasteSpecial(sheet.GetRange('D1'), copy, {mode:'values', transpose:true});
console.log(target.Address, target.Values); // D1:E2, [[1,3],[2,4]]
fillSeries(sheet.GetRange('G1:G6'), {type:'growth', start:2, step:2});
console.log(specialCells(sheet.UsedRange, 'constants').map(cell => cell.Address));
book.Undo(); // undo the complete growth series
```

`CellRange.Capture`, `PasteSpecial`, `FillSeries` and `SpecialCells` expose the same implementation. The named functions are also exported from the package root. PascalCase control methods are available through a React/native web-control ref.

## Paste special

`destination.PasteSpecial(sourceRangeOrSnapshot, options)` returns the actual destination. A one-cell destination expands to the copied shape. A larger selection must be an exact multiple of the source shape and repeats whole tiles. Overlap is safe because the source is captured before mutation.

| Mode | Content affected |
|---|---|
| `all` | Cell inputs/formulas, exact cell styles, comments and supported validation rules |
| `values` | Captured calculated values; errors and literal formula-looking text remain typed values |
| `formulas` | Shifted formulas or original literal inputs, without replacing styles/comments |
| `formats` | Exact cell-style replacement, without changing inputs or comments |
| `comments` | Cell comments only |
| `validation` | Clipped/rebased supported validation rules only |
| `columnWidths` | Source column widths |
| `valuesAndNumberFormats` | Captured values and number formats only |
| `formulasAndNumberFormats` | Shifted formulas/literal inputs and number formats only |

`operation` is `none`, `add`, `subtract`, `multiply` or `divide`. Arithmetic requires a values mode and computes **destination op source** using captured operands. Division by zero produces a cell `#DIV/0!`; invalid numeric operands produce `#VALUE!`. This increment materializes arithmetic results rather than rewriting formula expression trees.

`transpose` exchanges source rows and columns. Relative formula offsets are based on each actual source/destination cell; absolute references remain absolute. `skipBlanks` skips genuinely blank source cells; a formula returning empty text is not an empty input. Both flags require booleans.

Snapshots are immutable, process-local and intentionally opaque. They retain values from the moment of Copy even if the source later changes. Do not deserialize arbitrary clipboard JSON as a snapshot. External clipboard text remains handled by normal Paste/TSV import.

“All” is the profile above, **not every Excel range-associated object**: conditional-format rules, table membership, merges, drawings and arbitrary OOXML metadata are not transferred. Copying merged sources or a whole spilled array requires values mode. Merged destinations and edits into spill children are rejected. Column-width transpose and transposed custom-validation formulas are rejected. Validation paste cannot be combined with blank skipping when source/destination validation intersects; paste values/formats separately. Validation replacement clips existing rules, shifts relative custom origins and enforces a 1,000-rule result bound. Content and validation failures roll back the entire operation.

The legacy `CopyFrom` now also supports `formats` and correctly copies calculated error values. Its existing shape behavior is retained. Office-style `Range.copyFrom` adds `Formats`, transpose and skip-blank flags, with the same unsupported-combination errors; reads/writes remain deferred to `sync()`.

## Fill series

`range.FillSeries({type, direction, start, step, stop, dateUnit})` returns the number of cells written. Types: `linear`, `growth`, `date`. Directions: `down`, `right`, `up`, `left`. Each row/column is an independent series. Without `start`, its first cell in the selected direction is the seed. Existing styles remain untouched.

Linear uses `seed + index * step`; growth uses `seed * step ** index` with a positive finite multiplier. An optional stop value leaves subsequent cells unchanged. Dates accept numeric workbook serials or JavaScript Dates, an integer step and `day`, `weekday`, `month` or `year`. Month/year series stay anchored to the original day: January 31 → February 28 → March 31. Time fractions and the workbook's 1900 leap-day compatibility are retained. Dates outside the supported serial interval, invalid seeds, invalid options or non-finite results are explicit failures. Weekdays exclude Saturday and Sunday; holidays/custom calendars are not series options.

```js
sheet.GetRange('J1:J12').FillSeries({
  type:'date', dateUnit:'month', start:new Date('2026-01-31T00:00:00Z'), step:1
});
sheet.GetRange('J1:J12').Format.NumberFormat = 'yyyy-mm-dd';
```

## Go to special

`range.SpecialCells(type, {valueTypes})` returns individual cell ranges in row-major order. Types: `formulas`, `constants`, `blanks`, `errors`, `comments`, `validation`, `visible`. Formula/constant queries optionally restrict value types to `numbers`, `text`, `logical`, `errors`.

Error searches include calculated errors in spill children. Blanks exclude formulas returning empty text. Visible excludes manually hidden rows/columns and filtered rows. The API does not claim a multi-area Range object or noncontiguous selection. The control's dialog reports the count and shows up to 500 navigable results. Selecting one cell searches the used range; selecting a rectangle searches that rectangle.

## Control and host integration

`grid.ShowPasteSpecial()`, `ShowFillSeries()` and `ShowGoToSpecial()` return native dialogs. They are responsive, theme-aware, keyboard-isolated and close safely on disconnect. Apply errors stay visible; successful mutations restore grid focus. `grid.PasteSpecial(options)` and `FillSeries(options)` enforce read-only state at apply time. `FindSpecialCells` is read-only. Copy is required before Paste special; a cut clipboard or external text is not accepted as an internal copy snapshot.

Ctrl/Command+D fills down; Ctrl/Command+R fills right; Ctrl/Command+Alt+V opens Paste special. The context menu and sample Home/Data/Formulas ribbons expose all three dialogs. The Calculation tools command creates executable examples and a copy/paste practice area.

Generic host and worker clients can use these allowlisted JSON requests:

```json
{"id":1,"method":"range.fillSeries","sheet":"Sheet1","address":"A1:A10","options":{"start":1,"step":1}}
{"id":2,"method":"range.pasteSpecial","sheet":"Sheet1","address":"D1","sourceSheet":"Sheet1","sourceAddress":"A1:A10","options":{"mode":"values","transpose":true}}
{"id":3,"method":"range.specialCells","sheet":"Sheet1","address":"A1:M10","type":"formulas","options":{"valueTypes":["numbers"]}}
```

The results are a write count, actual destination address and address list, respectively. `capabilities` advertises the methods. The bridge still accepts only its allowlist; it does not evaluate arbitrary scripts/member paths. Host authorization is separate from the presentation control's read-only option.
