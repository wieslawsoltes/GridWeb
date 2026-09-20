using System.Text.Json;
using System.Text.Json.Serialization;
namespace GridWeb.Client;

/// <summary>Serializable options for a shared-engine defined name. Null options are omitted.</summary>
public sealed record DefinedNameOptions
{
    [JsonPropertyName("comment"), JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Comment { get; init; }
    [JsonPropertyName("hidden"), JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public bool? Hidden { get; init; }
    [JsonPropertyName("contextSheetId"), JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ContextSheetId { get; init; }
    [JsonPropertyName("baseAddress"), JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? BaseAddress { get; init; }
}

/// <summary>Names are scoped by worksheet name/ID, or by the workbook when sheet is null.</summary>
public static class DefinedNamesClientExtensions
{
    public static Task<JsonElement> DefineScopedNameAsync(this SpreadsheetClient client, string name, object? value, string? sheet = null, DefinedNameOptions? options = null, CancellationToken token = default)
    {
        ArgumentNullException.ThrowIfNull(client); ArgumentException.ThrowIfNullOrWhiteSpace(name);
        if (options?.Comment?.Length > 255) throw new ArgumentOutOfRangeException(nameof(options), "Comments are limited to 255 characters");
        return client.InvokeAsync("names.define", new { name, value, sheet, options }, token);
    }
    public static Task<JsonElement> ListDefinedNamesAsync(this SpreadsheetClient client, string? sheet = null, bool all = false, CancellationToken token = default)
    {
        ArgumentNullException.ThrowIfNull(client);
        return client.InvokeAsync("names.list", new { sheet, all }, token);
    }
    public static Task<JsonElement> GetDefinedNameAsync(this SpreadsheetClient client, string name, string? sheet = null, CancellationToken token = default)
    {
        ArgumentNullException.ThrowIfNull(client); ArgumentException.ThrowIfNullOrWhiteSpace(name);
        return client.InvokeAsync("names.get", new { name, sheet }, token);
    }
    public static Task<JsonElement> RenameDefinedNameAsync(this SpreadsheetClient client, string name, string newName, string? sheet = null, CancellationToken token = default)
    {
        ArgumentNullException.ThrowIfNull(client); ArgumentException.ThrowIfNullOrWhiteSpace(name); ArgumentException.ThrowIfNullOrWhiteSpace(newName);
        return client.InvokeAsync("names.rename", new { name, newName, sheet }, token);
    }
    public static Task<JsonElement> RemoveDefinedNameAsync(this SpreadsheetClient client, string name, string? sheet = null, CancellationToken token = default)
    {
        ArgumentNullException.ThrowIfNull(client); ArgumentException.ThrowIfNullOrWhiteSpace(name);
        return client.InvokeAsync("names.remove", new { name, sheet }, token);
    }
}
