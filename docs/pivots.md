# Managed pivot reports · 0.5.0

`Workbook.PivotTables` owns persistent, refreshable report definitions and the cells they materialize. This is separate from the legacy `PivotTable` one-shot summarizer. All web components, React views and native WebView hosts use the same model.

## Studio

Select a rectangular source including a unique text header row. Choose **Data → Pivot reports** (also under Insert). Supply the row/column fields, measure, aggregation and destination. The destination sheet is created if absent. Additional measures and filters are available through the dialog’s Advanced options JSON; they override the simple measure fields. Use Edit pivot to reconfigure the selected report, Refresh pivots after source edits, and Show details on a value cell to create a worksheet of its source records. Source headers with commas can be selected through the API rather than the comma-separated UI.

Normal JSON save/load retains definitions. Normal XLSX export writes native pivot/cache parts for the supported layouts below. Cached result cells remain ordinary editable cells: refreshing replaces edits inside the report’s owned output range. Growth into other occupied cells, tables, merged ranges, other pivots or the source is rejected without partial changes. Creating, reconfiguring, removing and refreshing reports are undoable transactions.

## JavaScript

```js
import { Workbook } from '@wieslawsoltes/gridweb';
import { exportXlsx } from '@wieslawsoltes/gridweb/io';
const book = new Workbook();
const data = book.ActiveWorksheet;
data.GetRange('A1:C5').Values = [
  ['Region','Product','Revenue'],
  ['North','A',10], ['North','B',20],
  ['South','A',30], ['South','B',40]
];
const output = book.Worksheets.Add('Report');
const pivot = book.PivotTables.Add('Sales', data.UsedRange, output.GetRange('A1'), {
  rows: ['Region'], values: [{column:'Revenue',aggregate:'sum',name:'Revenue'}],
  columnGrandTotals: true
});
data.GetCell('C2').Value = 15;
pivot.Refresh();
pivot.SetFilter('Region', ['North']);
const details = pivot.DrillDown(1, 1); // zero-based output row and column
const bytes = exportXlsx(book); // native pivot + cache definition + cache records
```

`Get(nameOrIndex)`, `Count`, iteration, `Remove(name,{clear})` and `RefreshAll()` manage reports. `Update({source,destination,options,name})` rebinds or reconfigures. `SetFilter(field,null)` removes a filter. `IsStale` compares a bounded source fingerprint; refresh always recalculates, never trusts that hint. Stale drill-down requires refresh before resolving displayed row identities. After JSON or remote loading, refresh clears the advisory stale state.

Fields accept names or zero-based source indexes. The managed engine supports up to 8 row fields, 4 column fields, 16 measures, 100 reports and the existing bounded dense cell limit. It includes sum, count, countNumbers, average, min, max, product, distinct, sample/population variance and standard deviation. Text grouping/filtering is case-insensitive. Grand averages and variances aggregate the original source rows, not aggregated summaries. Distinct count is a local aggregate, not a DAX/data-model implementation.

The source is a fixed range; extending data beyond it requires Update. Row/column insertion and deletion translate source/output references within the same undo transaction. Sheet renames retain stable IDs. Deleted references become explicitly invalid and require reconfiguration. Hierarchical subtotals, date grouping, calculated items, show-as percentages and slicers are not implemented.

## Native XLSX contract

Native export currently supports range-backed row-axis reports with standard numeric/count measures, optional grand totals, multiple measures and row-field value filters. It generates native `pivotTableDefinition`, `pivotCacheDefinition`, typed shared items, `pivotCacheRecords`, content types and workbook/worksheet/cache relationships. It refreshes a cloned snapshot, without modifying the open workbook, its history or revision.

Managed column-axis layouts, distinct count, filters on non-row fields and measures reusing a row field are not supported by native export. They throw an actionable error unless the caller explicitly chooses `exportXlsx(book,{pivots:'flatten',onWarning})`, which exports refreshed result cells without pivot definitions. The default does not silently flatten a report. The current studio’s default XLSX command uses native mode; use the library option for explicit flattening or retain the full definition in JSON.

Import discovers local range-backed row-layout native pivots, standard measures and supported hidden-item filters. Imported layout formatting is normalized on refresh. Unsupported advanced layouts remain ordinary imported cells with warnings rather than a falsely editable full native pivot. This is not lossless preservation of arbitrary pivot parts, slicers, styles, OLAP caches or data models. Four generated fixture workbooks are independently validated with Microsoft Open XML SDK in CI; schema validation is not an Excel desktop runtime test.

## Workers and .NET

The existing allowlisted host/worker protocol adds `pivots.list`, `.add`, `.update`, `.refresh`, `.refreshAll`, `.filter`, `.drillDown` and `.remove`. No arbitrary member path or script execution is added. Range descriptors are `{sheet,address}`. The typed C# client adds ListPivotsAsync, AddPivotAsync, RefreshPivotAsync, RefreshAllPivotsAsync, SetPivotFilterAsync, GetPivotDetailsAsync and RemovePivotAsync.

```csharp
await control.Client.AddPivotAsync("Sales", "A1:C5", "E1",
    new PivotOptions {
        Rows = ["Region"],
        Values = [new PivotValue("Revenue", "sum", "Revenue")],
        ColumnGrandTotals = true
    });
await control.Client.RefreshPivotAsync("Sales");
var records = await control.Client.GetPivotDetailsAsync("Sales", 1, 1);
```

Native Windows smoke tests exercise these methods through WPF, WinUI and Avalonia WebViews. Workbook JSON and coauthoring preserve pivot definitions; changing a definition is a structural collaboration edit. Concurrent structural changes still require explicit whole-document conflict resolution. The limited Office-style adapter does not yet expose the complete Office.js PivotTable object hierarchy.

## Primary interoperability references

- https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-pivottables
- https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.pivottabledefinition
- https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.pivotcachedefinition
- https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.datafield
