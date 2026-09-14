using System.Text.Json;
using System.Text.Json.Serialization;
namespace GridWeb.Client;
public sealed record PivotValue(
    [property: JsonPropertyName("column")] string Column,
    [property: JsonPropertyName("aggregate")] string Aggregate = "sum",
    [property: JsonPropertyName("name"), JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)] string? Name = null);
public sealed record PivotFilter(
    [property: JsonPropertyName("column")] string Column,
    [property: JsonPropertyName("values")] object?[] Values);
public sealed record PivotOptions
{
    [JsonPropertyName("rows")] public string[] Rows { get; init; } = [];
    [JsonPropertyName("columns")] public string[] Columns { get; init; } = [];
    [JsonPropertyName("values")] public PivotValue[] Values { get; init; } = [];
    [JsonPropertyName("filters")] public PivotFilter[] Filters { get; init; } = [];
    [JsonPropertyName("rowGrandTotals")] public bool RowGrandTotals { get; init; }
    [JsonPropertyName("columnGrandTotals")] public bool ColumnGrandTotals { get; init; }
}
/// <summary>Managed pivot operations executed by the exact shared JavaScript engine.</summary>
public static class PivotClientExtensions
{
    public static Task<JsonElement> ListPivotsAsync(this SpreadsheetClient client, CancellationToken token = default)
        => client.InvokeAsync("pivots.list", cancellationToken: token);
    public static Task<JsonElement> AddPivotAsync(this SpreadsheetClient client, string name, string sourceAddress, string destinationAddress, PivotOptions options, string? sourceSheet = null, string? destinationSheet = null, CancellationToken token = default)
        => client.InvokeAsync("pivots.add", new { name, source = new { sheet = sourceSheet, address = sourceAddress }, destination = new { sheet = destinationSheet, address = destinationAddress }, options }, token);
    public static Task<JsonElement> RefreshPivotAsync(this SpreadsheetClient client, string name, CancellationToken token = default)
        => client.InvokeAsync("pivots.refresh", new { name }, token);
    public static Task<JsonElement> RefreshAllPivotsAsync(this SpreadsheetClient client, CancellationToken token = default)
        => client.InvokeAsync("pivots.refreshAll", cancellationToken: token);
    public static Task<JsonElement> SetPivotFilterAsync(this SpreadsheetClient client, string name, string field, object?[]? values, CancellationToken token = default)
        => client.InvokeAsync("pivots.filter", new { name, field, values }, token);
    public static Task<JsonElement> GetPivotDetailsAsync(this SpreadsheetClient client, string name, int row, int column, int limit = 10000, CancellationToken token = default)
        => client.InvokeAsync("pivots.drillDown", new { name, row, column, options = new { limit } }, token);
    public static Task<JsonElement> RemovePivotAsync(this SpreadsheetClient client, string name, bool clear = false, CancellationToken token = default)
        => client.InvokeAsync("pivots.remove", new { name, clear }, token);
}
