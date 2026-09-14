using System.Text.Json;
using System.Text.Json.Serialization;
namespace GridWeb.Blazor;

public sealed record WorkbookInfo(long Revision, long EngineRevision, string Name, string ActiveWorksheet, string Selection, bool CanUndo, bool CanRedo, string Label);
public sealed record WorksheetInfo(string Id, string Name, int CellCount = 0);
public sealed record WorkbookValue(long Revision, string Value);
public sealed record SelectionInfo(string Address, string Sheet, int Row = 0, int Column = 0, long Revision = 0);
public sealed record CellEditInfo(string Sheet, int Row, int Column, string Input, JsonElement Value);
public sealed record GridError(string Message, string? Code = null);
public sealed record GridMetrics(long Frames, int VisibleCells, double RenderMs);
public sealed record FindResult(string Sheet, string Address, int Row, int Column, JsonElement Value);
public sealed record SortKey(int Column, bool Ascending = true);
public sealed record WorkbookOperation(string Method, object? Arguments = null);
public sealed record SheetRange(string Address, string? Sheet = null);
public sealed record PivotValue(string Column, string Aggregate = "sum", string? Name = null);
public sealed record PivotFilter(string Column, object?[] Values);
public sealed record PivotOptions
{
    public string[] Rows { get; init; } = [];
    public string[] Columns { get; init; } = [];
    public PivotValue[] Values { get; init; } = [];
    public PivotFilter[] Filters { get; init; } = [];
    public bool RowGrandTotals { get; init; }
    public bool ColumnGrandTotals { get; init; }
}
public sealed record CollaborationOptions(string Url, string Room, string Token)
{
    public bool AutoSync { get; init; } = true;
    public bool PersistPending { get; init; }
    public string? StorageKey { get; init; }
}
