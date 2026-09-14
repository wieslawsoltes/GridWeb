# GridWeb.Blazor

[![NuGet](https://img.shields.io/nuget/v/GridWeb.Blazor)](https://www.nuget.org/packages/GridWeb.Blazor)
[![NuGet downloads](https://img.shields.io/nuget/dt/GridWeb.Blazor)](https://www.nuget.org/packages/GridWeb.Blazor)
[![Blazor CI](https://github.com/wieslawsoltes/GridWeb/actions/workflows/blazor.yml/badge.svg)](https://github.com/wieslawsoltes/GridWeb/actions/workflows/blazor.yml)

A .NET 8/.NET 10 Razor class library wrapping the **real GridWeb workbook engine and virtualized native browser editor**. Supports interactive WebAssembly and Server, full workbook JSON binding, EditForm field notifications, selection/worksheet binding, typed range/pivot APIs, CSV/XLSX interchange, charts/printing, collaboration and advanced native object/function interop. It is not a mock spreadsheet or a second calculation engine.

```sh
dotnet add package GridWeb.Blazor --version 0.1.0
```

No npm install, CDN, standalone script tag or global host bridge is required by NuGet consumers. The package carries its own static assets under `_content/GridWeb.Blazor/`, including a separately bundled worker. Fonts are not bundled. The native grid's shadow-root styles are automatic.

## Interactive hosting

Place the component in an interactive route/component, not static-only server rendering. WebAssembly and Interactive Server are qualified on both target frameworks. Interactive Auto can use those hosts; transitions across renderer instances recreate browser state and require application persistence. A browser workbook is never a Server singleton. Each component, provider or manually created `GridWebModule` owns its own lifetime.

```razor
@using GridWeb.Blazor
<Spreadsheet @ref="grid" @bind-Value="json" @bind-Selection="selection"
             @bind-Worksheet="sheet" Theme="light" Style="display:block;height:520px" />
@code {
    private Spreadsheet? grid;
    private string? json; // null initially creates a blank workbook
    private string? selection = "A1", sheet;
}
```

`Value` is GridWeb workbook JSON, not XLSX/base64. Browser edits carry monotonic revisions; stale echoed values do not overwrite newer local edits. Change `ValueRevision` when explicitly resetting to a previously supplied value. `SelectionRevision` provides the same explicit reset mechanism for selections. `DebounceMilliseconds` controls binding notifications (default 150). Full values are pulled as streams rather than truncated diagnostic snapshots. Call `FlushChangesAsync` before saving/submitting to commit an active cell edit and obtain the latest bound value; failed native validation is reported, not silently accepted.

`Theme`, `Zoom`, `ViewMode` (`normal`, `pageBreak`, `pageLayout`), `ReadOnly`, `ShowGridLines`, `AriaLabel`, CSS class/style and unmatched HTML attributes are component parameters. `ReadOnly` gates interactive edits; it intentionally does not deny trusted application code from changing the engine. `Ready` fires after the actual native control and `Client` are initialized. `CellEdited`, `Error`, `Changed`, `ZoomChanged`, selection and worksheet callbacks are available. Use `Options` for additional supported native presentation properties, not private members or replacing the workbook.

## Typed workbook API

```razor
@using GridWeb.Blazor
@using Microsoft.JSInterop
<Spreadsheet @ref="grid" Ready="Initialize" />
@code {
    private Spreadsheet grid = default!;
    private async Task Initialize(IJSObjectReference control)
    {
        var book = grid.Client!;
        await book.GetRange("A1:B2").SetValuesAsync([[10,20],[30,40]]);
        await book.GetRange("C1").SetFormulasAsync([["=SUM(A1:B2)"]]);
        await book.GetRange("C1").SetStyleAsync(new { font = new { bold = true }, numberFormat = "0.00" });
        var result = await book.GetRange("C1").GetValuesAsync(); // 100, calculated in the browser
    }
}
```

`GridWorkbook` provides worksheet creation/removal/rename/activation, range values/formulas/text/styles, merge/unmerge/clear, fill/copy/sort, row/column structure, freeze panes, names, find/replace, undo/redo and calculation. `BatchAsync` runs its documented synchronous mutations atomically; a failed operation rolls back the whole batch. Workbook replacement, collaboration and asynchronous I/O are deliberately not transaction-batch operations.

Managed pivots use `AddPivotAsync`, `UpdatePivotAsync`, `RefreshPivotAsync`, `RefreshPivotsAsync`, `SetPivotFilterAsync`, `GetPivotDetailsAsync` and `RemovePivotAsync`. `SheetRange` distinguishes source and destination worksheets. Tables, validation, conditional formatting and charts use typed entry methods; `CallWorksheetAsync` and `GridRange.InvokeAsync` expose remaining public native methods returning JSON values.

## Shared and nonvisual workbooks

```razor
<WorkbookProvider>
    <ChildContent Context="book">
        <Spreadsheet Workbook="book" Style="height:300px" />
        <Spreadsheet Workbook="book" ReadOnly="true" Style="height:200px" />
    </ChildContent>
    <Loading><p>Loading the engine…</p></Loading>
</WorkbookProvider>
```

The provider owns the workbook; child grids borrow it and release their own view/listener resources on removal. Do not combine `Workbook` and non-null `Value` binding on one grid. `WorkbookProvider.InitialValue` is initialization-only; call `book.LoadAsync` for later replacements. For services, create a `GridWebModule` using the circuit's `IJSRuntime` after interactive rendering, then `CreateWorkbookAsync`. Dispose the workbook before the module. `GridProvider` is a nonvisual `RenderFragment<BrowserModule>` owner exposing every engine export.

`SpreadsheetToolbar` is an optional Razor command/formula bar: pass a ready `Spreadsheet` through `Editor`, and place application controls in `ChildContent`. The native canvas deliberately does not claim arbitrary Razor cell-template support. Use Razor around the grid or native canvas/editor customization; the shared `BrowserTemplate<T>` facility is available for native DOM factories that accept elements.

## Files, printing, workers and collaboration

`ExportCsvAsync`/`ImportCsvAsync` preserve native CSV contracts. Imported formulas are literal by default; enabling `allowFormulas` must be an intentional trust decision. `ExportXlsxAsync` returns real ZIP bytes; `ImportXlsxAsync` returns all native compatibility warnings and refuses to apply a result when the workbook changed during its asynchronous import. Set application `InputFile` size limits explicitly. Large results use streamed JSON/binary reads (64 MiB default); native workbook/operation limits still apply.

`GetPrintHtmlAsync` and `GetChartSvgAsync` generate native inert output. `GetNativeWorkbookAsync`, `GetNativeRangeAsync`, `GetOfficeApiAsync` and `GetNativeControlAsync` return JS references for the existing complete public API. Dispose returned handles. Use `Module.CallAsync`, `GetAsync`, `SetAsync`, `CallReferenceAsync`, `CallFunctionAsync`, `CallFunctionJsonAsync` and `CallBatchAsync` as appropriate. Native live objects should stay references, not be serialized as JSON. `BrowserValue.Literal` protects application DTOs from callback-descriptor interpretation. Synchronous calculation callbacks must be browser functions (`BrowserFunction.Module`); a Server round trip cannot synchronously return a calculation result.

Create a module worker at `new URL('./_content/GridWeb.Blazor/worker.js', document.baseURI)` with `{type:'module'}`. `WorkbookWorkerClient` and `createWorkerHandler` are included in native exports. Worker sessions are separate engines and use the native allowlisted message protocol, not shared in-memory objects.

`ConnectAsync(new CollaborationOptions(url, room, token), "join"/"create")`, `SyncAsync`, `ResolveConflictAsync`, `SetPresenceAsync` and `DisconnectAsync` use the existing authenticated GridWeb service. Subscribe to `CollaborationChanged`, `CollaborationConflict`, `PresenceChanged` and `Error` on the workbook for UI updates. `PersistPending` opts into browser session storage without storing tokens. Disconnect before loading/replacing the workbook. The NuGet package and GitHub Pages do **not** host the Node service. Configure HTTPS, CORS, durable storage and authorization using [the service guide](https://github.com/wieslawsoltes/GridWeb/blob/main/docs/collaboration.md). Do not put credentials in logs, URLs or committed sample files.

## Forms

```razor
<EditForm Model="model" OnSubmit="Submit">
    <SpreadsheetInput @ref="input" @bind-Value="model.Workbook" />
    <button type="submit">Save</button>
</EditForm>
@code {
    private Model model = new();
    private SpreadsheetInput input = default!;
    private async Task Submit() { await input.FlushAsync(); /* persist model.Workbook */ }
    private sealed class Model { public string? Workbook { get; set; } }
}
```

The input participates in EditContext modification/validation through `InputBase<string?>`; it does not invent workbook-wide validation policies. Consumers can add DataAnnotations or custom validators. Native cell validation still applies.

## Source builds and release

```sh
git submodule update --init --recursive
npm ci
npm run build
node blazor/build.mjs
dotnet pack blazor/src/GridWeb.Blazor.csproj -c Release -o artifacts/nuget
dotnet run --project blazor/sample/Sample.csproj
# Or: dotnet run --project blazor/server/Server.csproj
```

Use the .NET 10 SDK plus .NET 8 targeting support. Generated common C#/JS, project files and sample host infrastructure come from the exact Dockyard source commit in the gitlink. This is build-time source reuse, **not a Dockyard NuGet/runtime dependency**. `Spreadsheet.cs`, `GridWorkbook.cs`, models, provider, toolbar, native adapters, sample and GridWeb-specific tests are maintained here.

CI builds the actual multi-target `.nupkg`, verifies payloads, runs native/interop tests, restores package consumers and exercises both hosts in Chromium at non-root paths. The sample covers keyboard edits, EditForm notifications, formula calculation, native pivots/charts/printing, safe CSV, large Unicode transfers, XLSX round trips and remounting, in addition to the shared lifecycle/template suite.

`blazor/Version.props` versions this package independently of npm and the existing `dotnet/` desktop adapters. Version-changing main merges publish using `NUGET_API_KEY` (`NUGET_TOKEN`/`NUGET_KEY` aliases), compare every public NuGet payload member, then create `blazor-v*` releases with packages, symbols, runnable samples and SHA-256 checksums. Retry only the failed original publish job after a delayed upload; do not rebuild/reuse an immutable version. Missing secrets or mismatching payloads fail the release.

## Boundaries

The wrapper preserves [GridWeb's native compatibility limits](https://github.com/wieslawsoltes/GridWeb/blob/main/docs/compatibility.md). It does not establish full Excel/Office.js/VBA equivalence, arbitrary lossless XLSX, enterprise coauthoring, printer-identical output or hardware-GPU/all-browser certification. The native grid is Canvas 2D, not a claimed WebGPU implementation. Static-only SSR renders a host placeholder; JavaScript starts only after interactivity. Native WPF/WinUI/Avalonia adapters remain separate packages.
