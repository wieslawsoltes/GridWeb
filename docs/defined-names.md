# Scoped names and reference bindings · 0.8.0

This is the implemented GridWeb profile, not full Microsoft Excel or Office.js equivalence. The engine, controls, workers and native hosts use the same name collections. Formulas never execute as JavaScript.

## Scope and public API

`Workbook.Names` contains workbook-scoped definitions. `Worksheet.Names` contains only that worksheet's definitions. Matching is case-insensitive; display casing, comments and hidden state are preserved. Unqualified formula names resolve lexical LET/LAMBDA variables first, then local names, then global names. `'North region'!TaxRate` explicitly selects a local name. A missing local name does not fall back to a global of the same spelling.

```js
import {Workbook} from '@wieslawsoltes/gridweb';
const book = new Workbook(), north = book.ActiveWorksheet;
north.Name = 'North';
const south = book.Worksheets.Add('South');
book.Names.Create('TaxRate', '=0.05', {comment: 'Default'});
north.Names.Create('TaxRate', '=0.1');
south.Names.Create('TaxRate', '=0.2');
north.GetRange('B2:B4').Values = [[100], [200], [300]];
north.Names.Create('Revenue', '=B2:B4');
north.Names.Create('TaxDue', '=LAMBDA(amount,amount*TaxRate)');
south.GetCell('D1').Formula = '=North!TaxDue(SUM(North!Revenue))';
console.log(south.GetCell('D1').Value); // 60, using North's local rate.
console.log(book.Names.Evaluate('TaxRate')); // 0.05, bypassing local shadowing.
console.log(north.Names.GetRange('Revenue').FullAddress); // 'North'!B2:B4
```

`Add` retains the legacy `DefineName` upsert contract; `Create` rejects duplicates, and `Update` requires an existing name. `Get` and iteration return stored inputs, not calculated values. `GetDefinition`, `Items`, and `Workbook.GetDefinedNames()` return detached snapshots containing name, value, comment, hidden state, scope label, scope worksheet ID, definition-context ID and optional origin. `Has`, `Count`, `Remove`, `Rename`, `Evaluate` and `GetRange` complete the collection API. Names are limited to 10,000 across a workbook, 255 identifier/comment characters and the existing input limits. Identifiers use the supported ASCII/underscore/backslash grammar; reserved A1/R1C1-shaped names and table collisions are rejected. Full Unicode and localized naming remain work.

`Evaluate(name, {sheet, row, col})` returns the calculated value. An uninvoked named LAMBDA reports `#CALC!`. Invalid syntax retains the existing Evaluate exception boundary. `GetRange` requires one rectangular reference, not a value or discontiguous/3-D reference. Workbook/worksheet `GetRange` and grid `Select` also accept ordinary named-range strings. Collection methods address their own scope directly, unlike formula lookup with local shadowing.

New global definitions use the active worksheet as their definition context. Local definitions use their owner worksheet. Contexts retain stable IDs through renaming/reordering; deletion leaves a tombstone rather than binding to a new sheet with the old name. Explicit global `contextSheetId:null` selects caller context. Legacy v1 documents without metadata retain caller-context behavior. Named formulas do not capture a caller's LET/LAMBDA variables.

## Create from Selection

`Names.CreateFromSelection(sourceRange, options)` creates absolute, sheet-qualified references from top/bottom rows and left/right columns. No options means `{topRow:true}`. Selected boundaries are excluded from the data. The source may be another worksheet in the same workbook; the destination collection determines scope.

```js
north.GetRange('F1:H3').Values = [
  ['Sales', 'Costs', 'Net margin'], [100, 60, 40], [200, 110, 90]
];
north.Names.CreateFromSelection(north.GetRange('F1:H3'));
console.log(north.Names.GetRange('Net_margin').Address); // H2:H3
book.Undo(); // Removes the three definitions, preserving source cells.
```

Boolean options are `topRow`, `bottomRow`, `leftColumn`, `rightColumn`, `overwrite`. At least one boundary and nonempty remaining data are required. Blank labels are skipped. Unsupported characters become underscores; leading digits and reserved/reference-shaped labels receive an underscore prefix; labels are capped at 255 characters. This deterministic normalization is a GridWeb profile, not all Excel locale behavior.

Duplicate normalized labels, erroneous labels, table collisions, foreign sources and resource limits reject the operation before writing. Existing names reject unless `overwrite:true`; overwrite preserves metadata. Two boundaries generating the same name reject even with overwrite, since their target references are ambiguous. Only labels are scanned, not a dense full-sheet matrix. All changes form one transaction and undo step.

## Reference-valued LET/LAMBDA

LET aliases, closures, ordinary LAMBDA arguments and directly returned references retain coordinates and worksheet identity. Supported INDEX, dynamic colon, union/intersection, ROW/COLUMN, ISREF and reference-aware aggregates consume those references without losing dependencies or row-visibility/nested-total information.

```text
=LET(data,North!Revenue,ISREF(data))
=LAMBDA(data,ISREF(data))(North!Revenue)
=LET(data,North!Revenue,SUM(INDEX(data,1):INDEX(data,2)))
=LET(data,North!Revenue,SUBTOTAL(109,data))
=SUM(LAMBDA(data,data)(North!Revenue))
```

Arithmetic such as `data+0` produces values rather than references. Reference probing and subsequent value reading do not execute the same LET/callable/named expression twice in the same evaluation environment. Earlier closures retain their original lexical environment. Scoped name cycles report `#CIRC!`; existing recursion, reference-area, intersection and materialization bounds remain. `ISOMITTED` distinguishes omitted lexical parameters from explicitly qualified names of the same spelling.

Not every special form or array helper preserves reference returns: IF/CHOOSE and all higher-order combinations are not qualified. Helper elements remain values. Reference wrappers are internal and never enter stored inputs, JSON or host responses.

## Rename, history and structure

`Names.Rename(oldName, newName)` repairs free uses resolving to that scope in cell formulas, other names, validation and conditional-format formulas. Literal strings, INDIRECT strings, table selectors and lexical binders remain unchanged. Invalid existing formula syntax is retained; repair is not promised for an expression that could not already be parsed.

Rename cannot change scope. Duplicate/table collisions, lexical capture, accidental built-in calls and local/global shadowing that would alter meaning reject atomically. Case-only renames and renames at the name-count limit are supported. Name changes and formula repairs share one undo operation. Name edits conservatively invalidate calculation; fine-grained name scheduling is not implemented.

Worksheet rename repairs qualified local-name uses. Deletion removes that sheet and its local names, converts qualified uses to permanent `#REF!`, and preserves undo identity. Row/column transforms use each definition's context. Missing contexts are not replaced by the active worksheet. The earlier grouped 3-D structural-edit restrictions remain.

Optional `baseAddress` is a GridWeb extension: `=A1+$A$1` defined at origin `B2`, evaluated from `B3`, reads `A2+$A$1` on the definition worksheet. With no base address, fixed A1-coordinate behavior is retained. JSON preserves explicit origins, but XLSX export and ambiguous structural edits reject them until native relative-name translation is qualified.

## Reusable controls and sample

`grid.ShowDefineName()`, `grid.ShowNameManager()`, and `grid.ShowCreateNamesFromSelection()` are control-owned dialogs, with native labeled inputs, keyboard isolation, inline errors, Escape, focus restoration, stale-state guards, and apply-time read-only checks. Ctrl+F3 and the context menu open Name Manager. `grid.DefineName(name, value, options, sheet)` is the guarded programmatic creation path. Direct engine and host APIs remain application-owned and are not blocked by a view's read-only flag.

Name Manager searches name/scope/comment, filters scopes, edits/renames, confirms deletion, or navigates to a rectangular named range. It shows at most 200 matching visible names; hidden definitions remain available through the API. Value/error filtering, multi-delete, range picking, table-name management and complete Excel dialog equivalence remain work. Editing formula and name together repairs recursive self-references within the same transaction rather than restoring the old name afterwards.

Create from Selection captures source and workbook revision; intervening edits require reopening instead of using stale labels. Home/Formulas expose the naming dialogs and **Named formulas**, an executable example with two region sheets and one summary. The default five-sheet startup workbook is unchanged. Initial revenue totals are 600 and 1,200; taxes are 60 and 240; input edits recalculate through the engine.

## JSON, OOXML, Office and native integration

Optional v1 JSON `nameMetadata` and per-sheet `names`/`nameMetadata` retain scope without flattening. Imports reject duplicate worksheet identities, duplicate names within a scope, orphan metadata and table collisions. Collaboration treats definition/metadata changes as structural edits and preserves collection ownership when applying snapshots.

OOXML ordinary `definedName` elements preserve `localSheetId`, `comment` and `hidden`. Export computes current sheet indexes and qualifies references/free local uses in definition context. Import uses original sheet indexes even after skipping unsupported sheets. Missing/unsupported scopes warn rather than become globals. Primitive names normalize to formula strings on OOXML import; compare evaluated behavior, not primitive storage identity. Explicit relative origins, null-valued names, and deleted definition contexts currently reject XLSX export rather than silently changing semantics. Unknown parts, future-function metadata and arbitrary lossless workbook fidelity are not provided by this increment.

The Office-style adapter adds `context.workbook.names`, `worksheet.names`, `NamedItemCollection.add/getItem/load`, and NamedItem `formula/comment/visible`, loaded `name/scope/type/value`, `worksheet`, `getRange` and `delete`. Add accepts an invariant formula or a Range from the same context, not arbitrary primitive arguments. Writes are deferred to `sync`, loads publish only on success, and failed batches roll back. Range references use absolute quoted coordinates; foreign proxies, unsupported properties, missing sheets and inappropriate range requests reject. Full Office.js NamedItem/requirement-set equivalence is not claimed.

JSON host/worker routes: `names.define`, `names.list`, `names.get`, `names.rename`, `names.remove`, `names.createFromSelection`. `sheet` chooses definition scope; null means global. Selection creation uses source `sheet,address`, optional destination `scopeSheet`, and options. C# extensions provide `DefineScopedNameAsync`, `ListDefinedNamesAsync`, `GetDefinedNameAsync`, `RenameDefinedNameAsync`, `RemoveDefinedNameAsync` with cancellation. The separately pinned Blazor runtime is unchanged.

## Qualification and target references

Core, Office, selection-label, type, package, browser and native smoke tests exercise this profile. A standard scoped-name XLSX fixture is included in OpenXML schema checks. Self-round-trips do not prove native Excel open/save/reopen equivalence. See [verification](verification-0.8.md) and [feature audit](excel-feature-audit.md).

Primary targets: [Microsoft names and Create from Selection](https://support.microsoft.com/en-us/excel/names-in-formulas), [Name Manager](https://support.microsoft.com/en-us/excel/use-the-name-manager-in-excel), [ISOMITTED](https://support.microsoft.com/en-us/excel/functions/isomitted-function), [Office NamedItem](https://learn.microsoft.com/en-us/javascript/api/excel/excel.nameditem), [OOXML DefinedName](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.definedname). Target documentation is not evidence of complete implementation.
