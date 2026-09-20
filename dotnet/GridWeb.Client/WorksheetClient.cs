using System.Text.Json;
namespace GridWeb.Client;

/// <summary>Worksheet structure commands for the shared engine. Indexes are zero-based.</summary>
public static class WorksheetClientExtensions
{
    /// <summary>Insert before the worksheet at index, or append at the current worksheet count.</summary>
    public static Task<JsonElement> AddWorksheetAtAsync(this SpreadsheetClient client, string name, int index, CancellationToken token = default)
    {
        ArgumentNullException.ThrowIfNull(client);
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        ArgumentOutOfRangeException.ThrowIfNegative(index);
        // The engine checks the current count atomically when this command executes.
        return client.InvokeAsync("worksheets.add", new { name, index }, token);
    }

    /// <summary>Move an existing worksheet by name or ID to its final index. Formula repair is one undoable engine transaction.</summary>
    public static Task<JsonElement> MoveWorksheetAsync(this SpreadsheetClient client, string sheet, int index, CancellationToken token = default)
    {
        ArgumentNullException.ThrowIfNull(client);
        ArgumentException.ThrowIfNullOrWhiteSpace(sheet);
        ArgumentOutOfRangeException.ThrowIfNegative(index);
        return client.InvokeAsync("worksheets.move", new { sheet, index }, token);
    }
}
