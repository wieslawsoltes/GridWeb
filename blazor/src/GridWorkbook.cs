using System.Text.Json;
using Microsoft.JSInterop;
namespace GridWeb.Blazor;

/// <summary>Typed asynchronous access to one real browser workbook. No global JS or server singleton.</summary>
public sealed class GridWorkbook : IAsyncDisposable
{
    public BrowserModule Module { get; }
    public IJSObjectReference Reference { get; }
    private readonly bool _ownsSession;
    private Task? _disposal;
    private readonly object _sync = new();
    public bool IsDisposed { get; private set; }
    internal GridWorkbook(BrowserModule module, IJSObjectReference reference, bool ownsSession = true) { Module = module; Reference = reference; _ownsSession = ownsSession; }
    private void Check() => ObjectDisposedException.ThrowIf(IsDisposed, this);
    public ValueTask<T> CallAsync<T>(string method, object?[]? arguments = null, CancellationToken token = default) { Check(); return Module.CallJsonAsync<T>(Reference, method, arguments, token); }
    public ValueTask CallVoidAsync(string method, object?[]? arguments = null, CancellationToken token = default) { Check(); return Module.CallVoidAsync(Reference, method, arguments, token); }
    public ValueTask<JsonElement> InvokeAsync(string method, object? arguments = null, CancellationToken token = default) => CallAsync<JsonElement>("Invoke", [method, BrowserValue.Literal(arguments ?? new { })], token);
    public ValueTask<T> InvokeAsync<T>(string method, object? arguments = null, CancellationToken token = default) => CallAsync<T>("Invoke", [method, BrowserValue.Literal(arguments ?? new { })], token);
    public GridRange GetRange(string address, string? worksheet = null) => new(this, address, worksheet);
    public ValueTask<WorkbookInfo> GetInfoAsync(CancellationToken token = default) => CallAsync<WorkbookInfo>("Info", token: token);
    public ValueTask<JsonElement> SaveAsync(CancellationToken token = default) => CallAsync<JsonElement>("Save", token: token);
    public ValueTask LoadAsync(string json, CancellationToken token = default) => CallVoidAsync("Load", [json], token);
    public ValueTask LoadAsync(JsonElement document, CancellationToken token = default) => CallVoidAsync("Load", [BrowserValue.Literal(document)], token);
    public ValueTask NewAsync(CancellationToken token = default) => CallVoidAsync("New", token: token);
    public ValueTask<JsonElement[]> BatchAsync(IReadOnlyList<WorkbookOperation> operations, string label = "Blazor batch", CancellationToken token = default) => CallAsync<JsonElement[]>("Batch", [BrowserValue.Literal(operations), label], token);
    public ValueTask<WorksheetInfo[]> GetWorksheetsAsync(CancellationToken token = default) => InvokeAsync<WorksheetInfo[]>("worksheets.list", token: token);
    public ValueTask<WorksheetInfo> AddWorksheetAsync(string name, CancellationToken token = default) => InvokeAsync<WorksheetInfo>("worksheets.add", new { name }, token);
    public ValueTask<JsonElement> RemoveWorksheetAsync(string sheet, CancellationToken token = default) => InvokeAsync("worksheets.remove", new { sheet }, token);
    public ValueTask<JsonElement> RenameWorksheetAsync(string sheet, string name, CancellationToken token = default) => InvokeAsync("worksheets.rename", new { sheet, name }, token);
    public ValueTask ActivateWorksheetAsync(string sheet, CancellationToken token = default) => CallVoidAsync("ActivateWorksheet", [sheet], token);
    public ValueTask SelectAsync(string address, string? sheet = null, CancellationToken token = default) => CallVoidAsync("Select", [address, sheet], token);
    public ValueTask<JsonElement> CalculateAsync(bool full = false, CancellationToken token = default) => InvokeAsync("workbook.calculate", new { full }, token);
    public ValueTask<bool> UndoAsync(CancellationToken token = default) => InvokeAsync<bool>("workbook.undo", token: token);
    public ValueTask<bool> RedoAsync(CancellationToken token = default) => InvokeAsync<bool>("workbook.redo", token: token);
    public ValueTask<JsonElement> FreezeAsync(int rows, int columns, string? sheet = null, CancellationToken token = default) => InvokeAsync("sheet.freeze", new { rows, columns, sheet }, token);
    public ValueTask<JsonElement> InsertRowsAsync(int index, int count = 1, string? sheet = null, CancellationToken token = default) => InvokeAsync("sheet.insertRows", new { index, count, sheet }, token);
    public ValueTask<JsonElement> DeleteRowsAsync(int index, int count = 1, string? sheet = null, CancellationToken token = default) => InvokeAsync("sheet.deleteRows", new { index, count, sheet }, token);
    public ValueTask<JsonElement> InsertColumnsAsync(int index, int count = 1, string? sheet = null, CancellationToken token = default) => InvokeAsync("sheet.insertColumns", new { index, count, sheet }, token);
    public ValueTask<JsonElement> DeleteColumnsAsync(int index, int count = 1, string? sheet = null, CancellationToken token = default) => InvokeAsync("sheet.deleteColumns", new { index, count, sheet }, token);
    public ValueTask<FindResult[]> FindAsync(string query, object? options = null, CancellationToken token = default) => CallAsync<FindResult[]>("Find", [query, BrowserValue.Literal(options ?? new { })], token);
    public ValueTask<int> ReplaceAsync(string query, string replacement, object? options = null, CancellationToken token = default) => CallAsync<int>("Replace", [query, replacement, BrowserValue.Literal(options ?? new { })], token);
    public ValueTask<JsonElement> DefineNameAsync(string name, object value, CancellationToken token = default) => InvokeAsync("names.define", new { name, value }, token);
    public ValueTask<JsonElement> CallWorksheetAsync(string? sheet, string method, object?[]? arguments = null, CancellationToken token = default) => CallAsync<JsonElement>("CallWorksheet", [sheet, method, BrowserValue.Literal(arguments ?? [])], token);
    public ValueTask<JsonElement> AddTableAsync(string address, string name, string? sheet = null, CancellationToken token = default) => CallWorksheetAsync(sheet, "AddTable", [address, name], token);
    public ValueTask<JsonElement> AddChartAsync(string address, object? options = null, string? sheet = null, CancellationToken token = default) => CallWorksheetAsync(sheet, "AddChart", [address, options ?? new { }], token);
    public ValueTask<JsonElement> AddValidationAsync(string address, object rule, string? sheet = null, CancellationToken token = default) => CallWorksheetAsync(sheet, "AddValidation", [address, rule], token);
    public ValueTask<JsonElement> AddConditionalFormatAsync(string address, object rule, string? sheet = null, CancellationToken token = default) => CallWorksheetAsync(sheet, "AddConditionalFormat", [address, rule], token);
    public ValueTask<JsonElement> AddPivotAsync(string name, SheetRange source, SheetRange destination, PivotOptions options, CancellationToken token = default) => InvokeAsync("pivots.add", new { name, source, destination, options }, token);
    public ValueTask<JsonElement> GetPivotsAsync(CancellationToken token = default) => InvokeAsync("pivots.list", token: token);
    public ValueTask<JsonElement> RefreshPivotAsync(string name, CancellationToken token = default) => InvokeAsync("pivots.refresh", new { name }, token);
    public ValueTask<JsonElement> RefreshPivotsAsync(CancellationToken token = default) => InvokeAsync("pivots.refreshAll", token: token);
    public ValueTask<JsonElement> UpdatePivotAsync(string name, PivotOptions options, CancellationToken token = default) => InvokeAsync("pivots.update", new { name, options }, token);
    public ValueTask<JsonElement> SetPivotFilterAsync(string name, string field, object?[]? values, CancellationToken token = default) => InvokeAsync("pivots.filter", new { name, field, values }, token);
    public ValueTask<JsonElement> GetPivotDetailsAsync(string name, int row, int column, int limit = 10000, CancellationToken token = default) => InvokeAsync("pivots.drillDown", new { name, row, column, options = new { limit } }, token);
    public ValueTask<JsonElement> RemovePivotAsync(string name, bool clear = false, CancellationToken token = default) => InvokeAsync("pivots.remove", new { name, clear }, token);
    public ValueTask<string> ExportCsvAsync(string? sheet = null, string? address = null, string delimiter = ",", CancellationToken token = default) => CallAsync<string>("ExportCsv", [sheet, address, delimiter], token);
    public ValueTask<string?> ImportCsvAsync(string text, string? sheet = null, string address = "A1", object? options = null, CancellationToken token = default) => CallAsync<string?>("ImportCsv", [text, sheet, address, BrowserValue.Literal(options ?? new { })], token);
    public ValueTask<byte[]> ExportXlsxAsync(object? options = null, CancellationToken token = default) { Check(); return Module.CallBytesAsync(Reference, "ExportXlsx", [BrowserValue.Literal(options ?? new { })], token); }
    public ValueTask<string[]> ImportXlsxAsync(byte[] bytes, CancellationToken token = default) => CallAsync<string[]>("ImportXlsx", [bytes], token);
    public ValueTask<string> GetPrintHtmlAsync(string? sheet = null, object? options = null, CancellationToken token = default) => CallAsync<string>("PrintHtml", [sheet, BrowserValue.Literal(options ?? new { })], token);
    public ValueTask<string> GetChartSvgAsync(string? sheet, string chartId, object? options = null, CancellationToken token = default) => CallAsync<string>("ChartSvg", [sheet, chartId, BrowserValue.Literal(options ?? new { })], token);
    public ValueTask<JsonElement> ConnectAsync(CollaborationOptions options, string mode = "join", CancellationToken token = default) => CallAsync<JsonElement>("Connect", [BrowserValue.Literal(options), mode], token);
    public ValueTask DisconnectAsync(CancellationToken token = default) => CallVoidAsync("Disconnect", token: token);
    public ValueTask<bool> SyncAsync(CancellationToken token = default) => CallAsync<bool>("Sync", token: token);
    public ValueTask<bool> ResolveConflictAsync(string strategy, CancellationToken token = default) => CallAsync<bool>("ResolveConflict", [strategy], token);
    public ValueTask<JsonElement> SetPresenceAsync(string sheet, string selection, CancellationToken token = default) => CallAsync<JsonElement>("SetPresence", [sheet, selection], token);
    public ValueTask<BrowserSubscription> SubscribeAsync<T>(string name, Func<T, Task> callback) { Check(); return Module.SubscribeJsonAsync(Reference, name, callback); }
    public ValueTask<IJSObjectReference> GetNativeWorkbookAsync(CancellationToken token = default) { Check(); return Module.CallAsync<IJSObjectReference>(Reference, "GetWorkbook", cancellationToken: token); }
    public ValueTask<IJSObjectReference> GetNativeRangeAsync(string address, string? sheet = null, CancellationToken token = default) { Check(); return Module.CallAsync<IJSObjectReference>(Reference, "GetRange", [address, sheet], token); }
    public ValueTask<IJSObjectReference> GetOfficeApiAsync(CancellationToken token = default) { Check(); return Module.CallAsync<IJSObjectReference>(Reference, "GetOfficeApi", cancellationToken: token); }
    public ValueTask DisposeAsync() { lock (_sync) { IsDisposed = true; return new(_disposal ??= DisposeCoreAsync()); } }
    private async Task DisposeCoreAsync()
    {
        try { if (_ownsSession) await Module.ReleaseAsync(Reference); else await Reference.DisposeAsync(); }
        catch (JSDisconnectedException) { }
        catch (ObjectDisposedException) { try { await Reference.DisposeAsync(); } catch (JSDisconnectedException) { } }
    }
}

public sealed class GridRange(GridWorkbook workbook, string address, string? worksheet)
{
    public GridWorkbook Workbook { get; } = workbook;
    public string Address { get; } = address;
    public string? Worksheet { get; } = worksheet;
    public ValueTask<JsonElement[][]> GetValuesAsync(CancellationToken token = default) => workbook.InvokeAsync<JsonElement[][]>("range.values.get", new { address = Address, sheet = Worksheet }, token);
    public ValueTask<JsonElement> SetValuesAsync(object?[][] values, CancellationToken token = default) => workbook.InvokeAsync("range.values.set", new { address = Address, sheet = Worksheet, values }, token);
    public ValueTask<string?[][]> GetFormulasAsync(CancellationToken token = default) => workbook.InvokeAsync<string?[][]>("range.formulas.get", new { address = Address, sheet = Worksheet }, token);
    public ValueTask<JsonElement> SetFormulasAsync(string?[][] formulas, CancellationToken token = default) => workbook.InvokeAsync("range.formulas.set", new { address = Address, sheet = Worksheet, formulas }, token);
    public ValueTask<string[][]> GetTextAsync(CancellationToken token = default) => workbook.InvokeAsync<string[][]>("range.text.get", new { address = Address, sheet = Worksheet }, token);
    public ValueTask<JsonElement> SetStyleAsync(object style, CancellationToken token = default) => workbook.InvokeAsync("range.style.set", new { address = Address, sheet = Worksheet, style }, token);
    public ValueTask<JsonElement> ClearAsync(string mode = "all", CancellationToken token = default) => workbook.InvokeAsync("range.clear", new { address = Address, sheet = Worksheet, mode }, token);
    public ValueTask<JsonElement> MergeAsync(CancellationToken token = default) => workbook.InvokeAsync("range.merge", new { address = Address, sheet = Worksheet }, token);
    public ValueTask<JsonElement> UnmergeAsync(CancellationToken token = default) => workbook.InvokeAsync("range.unmerge", new { address = Address, sheet = Worksheet }, token);
    public ValueTask SortAsync(IReadOnlyList<SortKey> keys, bool hasHeaders = true, CancellationToken token = default) => workbook.CallVoidAsync("Sort", [Address, BrowserValue.Literal(keys), Worksheet, hasHeaders], token);
    public ValueTask FillDownAsync(CancellationToken token = default) => workbook.CallVoidAsync("Fill", [Address, "down", Worksheet], token);
    public ValueTask FillRightAsync(CancellationToken token = default) => workbook.CallVoidAsync("Fill", [Address, "right", Worksheet], token);
    public ValueTask CopyFromAsync(GridRange source, string mode = "all", CancellationToken token = default)
    {
        ArgumentNullException.ThrowIfNull(source);
        if (!ReferenceEquals(Workbook, source.Workbook)) throw new ArgumentException("CopyFrom requires ranges from the same GridWorkbook handle. Use native references for an explicit cross-workbook copy.", nameof(source));
        return workbook.CallVoidAsync("Copy", [source.Address, Address, source.Worksheet, Worksheet, mode], token);
    }
    public ValueTask SetPropertyAsync(string property, object? value, CancellationToken token = default) => workbook.CallVoidAsync("SetRangeProperty", [Address, property, BrowserValue.Literal(value), Worksheet], token);
    public ValueTask<JsonElement> InvokeAsync(string method, object?[]? arguments = null, CancellationToken token = default) => workbook.CallAsync<JsonElement>("CallRange", [Address, method, BrowserValue.Literal(arguments ?? []), Worksheet], token);
}

public sealed class GridWebModule(IJSRuntime js) : BrowserModule(js)
{
    public async ValueTask<GridWorkbook> CreateWorkbookAsync(object? document = null, CancellationToken token = default) => new(this, await CreateAsync("SpreadsheetSession", [BrowserValue.Literal(document)], token));
}
