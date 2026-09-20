# Reference semantics and worksheet organization · 0.7.0

This is the implemented GridWeb profile, not a claim of complete Excel compatibility. A reference retains worksheet identity and rectangular coordinates until a value consumer reads it. Formula unions do **not** turn the grid selection, Range API or clipboard into multi-area editing controls.

## Formula forms

| Form | Example | Implemented behavior |
|---|---|---|
| 3-D worksheet span | `=SUM(Jan:Mar!B2:B10)` | Includes each worksheet between the endpoints in workbook order, inclusive. Quoted names and full row/column spans are accepted. |
| Union | `=SUM((A1:A3,C1:C3))` | Preserves ordered areas, including duplicate/overlapping contributions. Parentheses distinguish a union from function-argument commas. |
| Intersection | `=SUM(A1:C4 B2:D5)` | Uses the actual overlap on the same worksheet; disjoint references return `#NULL!`. |
| Dynamic range | `=SUM(A1:INDEX(A1:A10,5))` | Colon combines reference-valued endpoints on one worksheet. It never guesses cell coordinates from calculated numbers. |
| INDEX reference form | `=SUM(INDEX((A1:A3,C1:C3),0,1,2))` | Selects the second area and its entire first column. Areas must belong to one worksheet. |
| Spill reference | `=SUM(INDEX(A1:A10,1)#)` | Requires a single anchor cell with a current spill. |
| Structured columns | `=SUM(Sales[[Units]:[Revenue]])` | Selects the contiguous data-column span. `#Data`, `#Headers`, `#All`, `#This Row` and simple `@` row selectors are supported. |

The parser implements colon before intersection before union. Arithmetic, comparisons, concatenation, array literals and callable LAMBDA expressions continue through the same parser. Whitespace around arithmetic or before a function's parenthesis is not treated as intersection. Formula strings remain data, never JavaScript.

### 3-D consumers

Supported statistical consumers are SUM, AVERAGE/AVERAGEA, COUNT/COUNTA, MIN/MINA, MAX/MAXA, PRODUCT, STDEV/STDEV.S/STDEVP/STDEV.P/STDEVA/STDEVPA, and VAR/VAR.S/VARP/VAR.P/VARA/VARPA. HSTACK/VSTACK expand a 3-D input into one array per worksheet under the existing operation limit; SHEETS counts span members. SHEET returns the earliest referenced worksheet position.

Other consumers, including INDEX, INDIRECT, SUBTOTAL and AGGREGATE, reject 3-D references in this profile. A 3-D reference cannot be combined with union/intersection. Returning an unsupported multi-area value as a regular rectangular array reports an error instead of silently dropping areas. These explicit rules prevent accidental value flattening from being presented as universal 3-D support.

### INDEX, identity and errors

INDEX has separate array and reference paths. Zero selects a whole row/column; omitted coordinates retain the corresponding dimension. A two-argument single-row reference uses its second argument as the column coordinate. The optional fourth argument selects a one-based area, and all reference areas must be on the same sheet. Negative coordinates return `#VALUE!`; out-of-range coordinates/areas return `#REF!`. Array form has no area argument. Invalid arity is rejected even when INDEX/OFFSET/INDIRECT appears inside another reference expression.

`AREAS(reference)` counts areas. `ISREF(value)` reports actual reference identity, not whether a value resembles an address. `SHEET()` returns the current/active worksheet's one-based position; it also accepts a worksheet name or supported reference. `SHEETS()` returns worksheet count, while `SHEETS(reference)` counts distinct referenced worksheets. Error values remain worksheet errors. The public Evaluate API retains its existing distinction: invalid address syntax may throw a FormulaError while evaluation errors become values; stored cell formulas expose caught errors in the grid.

### A-suffixed statistics

AVERAGEA, MINA, MAXA, VARA, VARPA, STDEVA and STDEVPA include referenced/array logical values as 1/0 and text as zero. Actual empty cells are skipped; a formula returning empty text is still text. Directly supplied numeric text is coerced numerically, while nonnumeric direct text is an error. Errors propagate. The sample/population forms use their respective denominator, with explicit empty/insufficient-input and nonfinite-result errors.

```text
=AVERAGEA({2,TRUE,"text"})  -> 1
=AVERAGEA("2",TRUE)         -> 1.5
=AREAS((A1:A3,C1:C3))       -> 2
=ISREF(INDEX(A1:C3,2,2))    -> TRUE
```

## Dependencies and resource limits

Reference-aware aggregations record dependencies on every input area, even when currently empty. Editing a newly populated cell, resizing a source spill, or changing a named reference invalidates dependent formulas. Sparse full-axis aggregation visits populated inputs and spill results without allocating a worksheet-sized matrix. SUBTOTAL/AGGREGATE keep filtered/hidden-row and nested-total provenance through unions and reference-form INDEX; computed arrays do not have that provenance.

Existing 250,000-cell/value operation bounds remain. Multi-area references are limited to 256 areas; intersections allow at most 4,096 area pairs. The formula parser retains length/token/depth limits. Reference resolution has a depth bound, and HSTACK/VSTACK preflight the combined rectangle before reading large spans. Limits produce explicit errors, not partial worksheet writes or nested limit-error arrays. These are GridWeb safety limits, not reproductions of every Excel resource ceiling.

## Worksheet APIs

```js
import {Workbook} from '@wieslawsoltes/gridweb';
const book = new Workbook();
const jan = book.Worksheets.Add('Jan');
const mar = book.Worksheets.Add('Mar');
const feb = book.Worksheets.Add('Feb', 2); // Insert before Mar in this new workbook.
jan.GetCell('A1').Value = 10;
feb.GetCell('A1').Value = 20;
mar.GetCell('A1').Value = 30;
const total = book.ActiveWorksheet.GetCell('D1');
total.Formula = '=SUM(Jan:Mar!A1)';
console.log(total.Value); // 60
book.Worksheets.Move(feb, 3); // Final index: [Sheet1, Jan, Mar, Feb].
console.log(total.Value); // 40
book.Undo();
console.log(total.Value); // 60
```

`Worksheets.Add(name, index)` inserts at a zero-based index; omission appends. `Worksheets.Move(sheetOrName, index)` moves an attached worksheet to its final zero-based index. Invalid indexes/foreign worksheets are rejected, and a no-op creates no history. Moves preserve worksheet object/active-sheet identity, notify the model and recalculate reference membership. One undo restores order and formula repairs together.

Deleting an ordinary referenced worksheet replaces its reference with permanent `#REF!`; creating another sheet with the old name does not reconnect it. Deleting a 3-D endpoint moves the endpoint inward, collapses a one-sheet span to an ordinary reference, or returns REF if no member survives. Moving an endpoint past the opposite endpoint shrinks the old span to the remaining members. Interior moves change membership by current order. Reversed endpoint notation is supported.

Copy/rename/deletion/move scans share lexical handling for escaped sheet names, mixed absolute references, string literals and table selectors. Cell formulas, defined names, and validation/conditional-format formula fields participate in reference repair. Literal text, including strings passed to INDIRECT, is not rewritten. Reference maintenance is internal and can repair formulas on protected worksheets; user edits remain subject to protection.

**Important structural boundary:** inserting/deleting rows or columns on one worksheet when that operation would alter a live multi-sheet 3-D reference is rejected before mutation. It is not silently applied to every worksheet or to only one endpoint. Exact grouped-sheet structural transformations remain unimplemented. Normal single-sheet references, unrelated axes and edits outside the affected extent continue to use the existing rewrite path.

## Reusable control, adapters and sample

`grid.MoveWorksheet(index, sheet = grid.Sheet)` performs an apply-time read-only check and commits/cancels editing through the existing control path. `grid.ShowMoveWorksheet()` opens the control-owned dialog. Its source uses stable worksheet IDs. Applying after workbook replacement or order changes reports a stale-state error rather than moving an unintended worksheet. Native controls supply keyboard isolation, inline error handling, Escape, dark/light responsive layout and focus restoration; a late close event cannot steal focus from a newly opened editing dialog.

Home/Formulas → **Reference tools** creates three source sheets and a summary as one undoable transaction. The unchanged startup workbook still has five worksheets. The example demonstrates 3-D aggregation, unions/intersections, INDEX, area/sheet identity and inclusive statistics with actual editable inputs. **Move worksheet** is also exposed in Home/Formulas and the context menu. Other applications can use the dialog without importing sample code.

Host/worker RPC exposes `worksheets.add` with optional `index` and `worksheets.move` with `sheet,index`. The Office-style adapter's `worksheet.position` setter queues the same move for `context.sync()`, with normal batch rollback/history. C# clients have `AddWorksheetAtAsync` and `MoveWorksheetAsync`; see [native guide](../dotnet/README.md). The separately pinned Blazor runtime is unchanged and does not automatically include these root-engine additions.

## Remaining boundaries and qualification

Reference-valued LET/LAMBDA variables, sheet-local/external names, full table-selector grammar (including complex escaping/totals/multiple special items), every reference consumer and coercion rule, discontiguous UI editing, iterative calculation, grouped 3-D structural editing and lossless OOXML metadata remain incomplete. JSON/XLSX self-round-trips check GridWeb preservation, not desktop Excel open/AUDIT/save equivalence. No desktop Excel oracle or physical-printer qualification was run.

Tests: `reference-operators.test.mjs`, `reference-operators-types.ts`, `reference-browser.py`, installed package consumers, native protocol tests and the extended native runtime smoke. See [local verification](verification-0.7.md) and CI on the exact PR/released commit.

## Primary target documentation

- Microsoft [3-D references and worksheet-order effects](https://support.microsoft.com/en-us/excel/create-a-3-d-reference-to-the-same-cell-range-on-multiple-worksheets).
- Microsoft [INDEX](https://support.microsoft.com/en-us/excel/functions/index-function), [AREAS](https://support.microsoft.com/en-us/excel/functions/areas-function), [SHEET](https://support.microsoft.com/en-us/excel/functions/sheet-function), [SHEETS](https://support.microsoft.com/en-us/excel/functions/sheets-function).
- Microsoft [AVERAGEA](https://support.microsoft.com/en-us/excel/functions/averagea-function) and [structured references](https://support.microsoft.com/en-us/excel/using-structured-references-with-excel-tables).

These describe the compatibility target. Only the repository's explicit implementation/tests establish the supported GridWeb profile.

## 0.8 additions

[Scoped names and reference bindings](defined-names.md) add local definitions and preserve direct reference-valued LET/LAMBDA aliases, arguments and returns. Their supported subset supersedes those two earlier gaps. Existing 3-D consumers, grouped structural limitations and remaining special-form restrictions still apply.
