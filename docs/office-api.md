# Office-style batched API

Import `createExcelApi` from `@wieslawsoltes/gridweb/office`. The returned API operates on the exact Workbook instance used by the control, worker or native host. It does not require Office.js or install a global Excel object.

```js
import { Workbook } from '@wieslawsoltes/gridweb';
import { createExcelApi } from '@wieslawsoltes/gridweb/office';
const book = new Workbook();
const Excel = createExcelApi(book);
await Excel.run(async context => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange('A1:B2');
  range.values = [[1, 2], [3, 4]];
  range.format.font.bold = true;
  range.load('values,address');
  await context.sync();
  console.log(range.values, range.address);
});
```

Assignments enqueue operations. Reads require `load()` followed by `sync()`; otherwise PropertyNotLoaded is thrown. Loads are ordered relative to writes, return isolated snapshots, and refresh pending formula results. Each successful sync is one undoable engine transaction. A failed sync rolls back its writes and does not publish partial load results. This atomic rollback is an explicit GridWeb guarantee, not a claim that every Office host has identical failure behavior.

`Excel.run` performs a final sync when its callback returns with pending operations. A thrown callback discards unsynchronized operations. Contexts and proxies cannot be reused for new requests after disposal. No request is sent to Microsoft or a network service.

## Implemented surfaces

Worksheet collections: getItem, getItemAt, getActiveWorksheet, add, items loading. Worksheets: name/id/position, range/indexed/used-range access, activate and delete. Ranges: values, formulas, text, numberFormat, address/index/count loading, offset/resize/row/column/cell access, clear, merge/unmerge, formula-aware copy. Formatting: font/fill, alignment, wrap and column autofit. Matching declarations are shipped.

In matrix assignments, `null` skips an individual cell; empty text clears cell content or resets the number format. Whole-matrix null and dimension mismatches are rejected. Mixed font/fill values load as null. Unknown load properties, unsupported copy modes, row autofit and foreign/disposed contexts produce explicit errors.

The studio exposes a real example under View → Office API example. It creates a new worksheet through this API, calculates totals, reads them via load/sync, and is undoable as one transaction.

## Boundaries

This is not the full Office.js object model. Tables, charts, names, bindings, add-in permissions, events, host requirement sets, trackedObjects and COM/VBA APIs are not exposed through this adapter. Core equivalents may exist separately. Proxy tracking across structural edits is not Office-equivalent; use fresh ranges after those operations. Skip-blanks and transposed copying are explicitly unsupported. The adapter is single-process; it is not a coauthoring service.

Interoperability references:
- https://learn.microsoft.com/en-us/office/dev/add-ins/develop/application-specific-api-model
- https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-blank-null-values
- https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-ranges-set-format

These references explain target patterns. The executable tests establish only the contracts implemented here.
