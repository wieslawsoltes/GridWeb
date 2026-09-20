# GridWeb API

## Entry points

The package root exports the headless model, calculation, address, data, chart and print primitives. `/controls` registers `<grid-web>`; `/react` has `GridWeb`, `useWorkbook`, `useWorkbookRevision`; `/io` has JSON/CSV/XLSX helpers; `/host` provides `createHostBridge`; `/worker` provides `WorkbookWorkerClient` and `createWorkerHandler`. Public TypeScript declarations in `types/` describe arguments and results.

## Workbook and ranges

```js
import {Workbook} from '@wieslawsoltes/gridweb';
const book = new Workbook();
const sheet = book.ActiveWorksheet;
sheet.GetRange('A1:B2').Values = [[2, 3], [4, 5]];
sheet.GetCell('C1').Formula = '=SUM(A1:B2)';
console.log(sheet.GetCell('C1').Value); // 14
sheet.GetRange('A1:C1').SetStyle({font:{bold:true},fill:'#dceee3'});
const subscription = book.Changed.Subscribe(event => console.log(event));
// When the owner is disposed:
subscription.Dispose();
```

Rows and columns in index-based APIs are zero-based. A1 strings use the familiar one-based row labels and `$` absolute markers. Whole-axis ranges are represented sparsely; dense operations are bounded to prevent accidental huge allocations. Formulas begin with `=`; the Formula setter normalizes this. Literal strings beginning with `=` can be escaped with a leading apostrophe through Input.

Transactions group edits into one history entry and roll back on failure. Undo and redo apply the engine's model mutations; they are not Office automation. A workbook may be presented by several independent controls. Its calculation engine is independent of the DOM and can run in a worker.

## Web control

```js
import '@wieslawsoltes/gridweb/controls';
const grid = document.querySelector('grid-web');
grid.Workbook = book;
grid.Selection = 'B2:C4';
grid.Theme = 'dark';
grid.Zoom = 1.25;
grid.ViewMode = 'normal';
grid.addEventListener('selection-change', e => console.log(e.detail.address));
grid.addEventListener('cell-edit', e => console.log(e.detail));
```

The control owns viewport and input state, not a second workbook. `ReadOnly` restricts UI editing. The core remains application-owned. `Refresh()` requests a repaint. Disconnect/dispose subscriptions when a host is removed. React uses the same properties and custom events through refs and effects.

## Worker

```js
import {WorkbookWorkerClient} from '@wieslawsoltes/gridweb/worker';
const worker = new Worker(new URL('../src/worker.js', import.meta.url), {type:'module'});
const client = new WorkbookWorkerClient(worker);
await client.Call('range.values.set', {address:'A1:B1',values:[[2,3]]});
await client.Call('range.formulas.set', {address:'C1',formulas:[['=SUM(A1:B1)']]});
console.log(await client.Call('range.values.get', {address:'C1'}));
client.Dispose({terminate:true});
```

Bundlers should resolve their worker URL to the packaged worker entry. `Call` accepts an AbortSignal; cancellation abandons a response and does not imply interrupting a synchronous calculation already executing. A native .NET host uses the same allowlisted operation names through SpreadsheetClient.

## Files

`exportXlsx(book)` returns ZIP bytes. `await importXlsx(bytes)` returns `{workbook, warnings}`. Read warnings; recovered baseline export is not a lossless arbitrary workbook editor. `toDelimited`, `parseDelimited` and `importDelimited` are reusable CSV/TSV primitives. GridWeb JSON is its versioned model representation, not a Microsoft file format.

## Extension and safety

Use the calculation engine's explicit function registration rather than dynamic JavaScript evaluation. Imported workbooks cannot register arbitrary scripts. See declarations for event, table, pivot, chart, print and host data shapes, and the compatibility matrix for unsupported Excel semantics.

## Scoped names (0.8)

`Workbook.Names` and `Worksheet.Names` expose Create/Add/Update/Rename/Remove/Has/Get/GetDefinition/Items/Count/Evaluate/GetRange/CreateFromSelection. Scope-safe mutations share model transactions and undo. Grid methods ShowNameManager/ShowDefineName/ShowCreateNamesFromSelection and guarded DefineName use the same engine. See [contract and examples](defined-names.md).
