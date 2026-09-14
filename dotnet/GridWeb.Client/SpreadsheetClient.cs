using System.Text.Json;
using System.Text.Json.Nodes;
namespace GridWeb.Client;

/// <summary>The host must execute scripts on its UI dispatcher. Only trusted bundled HTML may be loaded.</summary>
public interface IJavaScriptTransport
{
    Task<string?> InvokeAsync(string script, CancellationToken cancellationToken = default);
}

public sealed class DelegateJavaScriptTransport(Func<string, CancellationToken, Task<string?>> invoke) : IJavaScriptTransport
{
    public Task<string?> InvokeAsync(string script, CancellationToken cancellationToken = default) => invoke(script, cancellationToken);
}
public sealed class SpreadsheetException(string message) : Exception(message) { }

/// <summary>Typed asynchronous client for the same JavaScript Workbook used by web controls.</summary>
public sealed class SpreadsheetClient(IJavaScriptTransport transport)
{
    private long _nextId;
    public RangeClient GetRange(string address, string? worksheet = null) => new(this, address, worksheet);
    public async Task<JsonElement> InvokeAsync(string method, object? arguments = null, CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(method);
        var request = arguments is null ? new JsonObject() : JsonSerializer.SerializeToNode(arguments) as JsonObject
            ?? throw new ArgumentException("RPC arguments must be an object", nameof(arguments));
        var id = Interlocked.Increment(ref _nextId); request["id"] = id; request["method"] = method;
        // Serialize twice: the inner JSON is data, the outer JSON is a safe JavaScript string literal.
        var raw = await transport.InvokeAsync("globalThis.gridWebHost.dispatch(" + JsonSerializer.Serialize(request.ToJsonString()) + ")", cancellationToken);
        var response = Decode(raw);
        if (response.ValueKind != JsonValueKind.Object) throw new SpreadsheetException("Invalid host response");
        if (!response.TryGetProperty("id", out var returnedId) || returnedId.GetInt64() != id) throw new SpreadsheetException("Host response ID mismatch");
        if (response.TryGetProperty("error", out var error)) throw new SpreadsheetException(error.GetProperty("message").GetString() ?? "Spreadsheet error");
        return response.TryGetProperty("result", out var result) ? result.Clone() : JsonSerializer.SerializeToElement<object?>(null);
    }
    public Task<JsonElement> SaveAsync(CancellationToken token = default) => InvokeAsync("workbook.get", cancellationToken: token);
    public Task<JsonElement> LoadAsync(JsonElement workbook, CancellationToken token = default) => InvokeAsync("workbook.load", new { workbook }, token);
    public Task<JsonElement> CalculateAsync(bool full = false, CancellationToken token = default) => InvokeAsync("workbook.calculate", new { full }, token);
    public Task<JsonElement> UndoAsync(CancellationToken token = default) => InvokeAsync("workbook.undo", cancellationToken: token);
    public Task<JsonElement> RedoAsync(CancellationToken token = default) => InvokeAsync("workbook.redo", cancellationToken: token);
    public Task<JsonElement> ListWorksheetsAsync(CancellationToken token = default) => InvokeAsync("worksheets.list", cancellationToken: token);
    public Task<JsonElement> AddWorksheetAsync(string name, CancellationToken token = default) => InvokeAsync("worksheets.add", new { name }, token);
    public Task<JsonElement> SelectAsync(string address, string? sheet = null, CancellationToken token = default) => InvokeAsync("view.selection.set", new { address, sheet }, token);
    public Task<JsonElement> SetViewOptionsAsync(string theme, bool readOnly, double zoom = 1, string viewMode = "normal", CancellationToken token = default)
        => InvokeAsync("view.options", new { Theme = theme, ReadOnly = readOnly, Zoom = zoom, ViewMode = viewMode }, token);
    public Task<JsonElement> DefineNameAsync(string name, object value, CancellationToken token = default) => InvokeAsync("names.define", new { name, value }, token);
    public static JsonElement Decode(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) throw new SpreadsheetException("Empty WebView response");
        var result = JsonSerializer.Deserialize<JsonElement>(raw);
        // WebView2 JSON-encodes JS return strings. Other transports can return them directly.
        if (result.ValueKind == JsonValueKind.String) result = JsonSerializer.Deserialize<JsonElement>(result.GetString()!);
        return result;
    }
}
public sealed class RangeClient(SpreadsheetClient client, string address, string? sheet)
{
    public string Address { get; } = address;
    public string? Worksheet { get; } = sheet;
    public Task<JsonElement> GetValuesAsync(CancellationToken token = default) => client.InvokeAsync("range.values.get", new { address = Address, sheet = Worksheet }, token);
    public Task<JsonElement> SetValuesAsync(object?[][] values, CancellationToken token = default) => client.InvokeAsync("range.values.set", new { address = Address, sheet = Worksheet, values }, token);
    public Task<JsonElement> GetFormulasAsync(CancellationToken token = default) => client.InvokeAsync("range.formulas.get", new { address = Address, sheet = Worksheet }, token);
    public Task<JsonElement> SetFormulasAsync(string?[][] formulas, CancellationToken token = default) => client.InvokeAsync("range.formulas.set", new { address = Address, sheet = Worksheet, formulas }, token);
    public Task<JsonElement> GetTextAsync(CancellationToken token = default) => client.InvokeAsync("range.text.get", new { address = Address, sheet = Worksheet }, token);
    public Task<JsonElement> SetStyleAsync(object style, CancellationToken token = default) => client.InvokeAsync("range.style.set", new { address = Address, sheet = Worksheet, style }, token);
    public Task<JsonElement> ClearAsync(string mode = "all", CancellationToken token = default) => client.InvokeAsync("range.clear", new { address = Address, sheet = Worksheet, mode }, token);
    public Task<JsonElement> MergeAsync(CancellationToken token = default) => client.InvokeAsync("range.merge", new { address = Address, sheet = Worksheet }, token);
    public Task<JsonElement> UnmergeAsync(CancellationToken token = default) => client.InvokeAsync("range.unmerge", new { address = Address, sheet = Worksheet }, token);
}
