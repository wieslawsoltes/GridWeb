using System.Text.Json;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Forms;
using Microsoft.AspNetCore.Components.Rendering;
using Microsoft.JSInterop;
namespace GridWeb.Blazor;

/// <summary>The real GridWeb canvas/editor, with owned or shared workbooks and full-value two-way binding.</summary>
public class Spreadsheet : BrowserComponent
{
    [Parameter] public GridWorkbook? Workbook { get; set; }
    [Parameter] public string? Value { get; set; }
    [Parameter] public EventCallback<string?> ValueChanged { get; set; }
    [Parameter] public long ValueRevision { get; set; }
    [Parameter] public string? Selection { get; set; }
    [Parameter] public EventCallback<string?> SelectionChanged { get; set; }
    [Parameter] public long SelectionRevision { get; set; }
    [Parameter] public string? Worksheet { get; set; }
    [Parameter] public EventCallback<string?> WorksheetChanged { get; set; }
    [Parameter] public string Theme { get; set; } = "light";
    [Parameter] public double Zoom { get; set; } = 1;
    [Parameter] public EventCallback<double> ZoomChanged { get; set; }
    [Parameter] public string ViewMode { get; set; } = "normal";
    [Parameter] public bool ReadOnly { get; set; }
    [Parameter] public bool ShowGridLines { get; set; } = true;
    [Parameter] public int DebounceMilliseconds { get; set; } = 150;
    [Parameter] public string AriaLabel { get; set; } = "GridWeb spreadsheet";
    [Parameter] public EventCallback<CellEditInfo> CellEdited { get; set; }
    [Parameter] public EventCallback<GridError> Error { get; set; }
    public GridWorkbook? Client { get; private set; }
    private GridWorkbook? _clientWorkbook;
    private string? _lastValue, _lastSelection;
    private long _valueVersion = -1, _selectionVersion = -1;
    protected override IReadOnlyList<string> RequiredEvents => ["BindingChanged", "SelectionChanged", "CellEdited", "ZoomChanged", "Error"];
    protected override IReadOnlyList<string> DefaultEvents => ["Changed"];
    protected override bool IsJsonEvent(string name) => true;
    protected override Dictionary<string, object?> BuildOptions() => new()
    {
        ["session"] = Workbook?.Reference, ["value"] = Value, ["valueRevision"] = ValueRevision,
        ["ackRevision"] = _lastValue == Value ? _valueVersion : -1, ["enableBinding"] = ValueChanged.HasDelegate,
        ["selection"] = Selection, ["selectionRevision"] = SelectionRevision, ["sheet"] = Worksheet,
        ["ackSelectionRevision"] = _lastSelection == Selection ? _selectionVersion : -1,
        ["theme"] = Theme, ["zoom"] = Zoom, ["viewMode"] = ViewMode, ["readOnly"] = ReadOnly,
        ["showGridLines"] = ShowGridLines, ["debounceMilliseconds"] = DebounceMilliseconds, ["ariaLabel"] = AriaLabel, ["native"] = base.BuildOptions()
    };
    protected override async Task OnBrowserReadyAsync(IJSObjectReference control)
    {
        var reference = await InvokeAsync<IJSObjectReference>("GetSession");
        Client = new(Module!, reference, false); _clientWorkbook = Workbook;
        if (IsDisposed) { await Client.DisposeAsync(); return; }
        await base.OnBrowserReadyAsync(control);
    }
    protected override async Task OnAfterRenderAsync(bool firstRender)
    {
        await base.OnAfterRenderAsync(firstRender);
        if (!IsReady || ReferenceEquals(_clientWorkbook, Workbook)) return;
        var reference = await InvokeAsync<IJSObjectReference>("GetSession");
        if (IsDisposed) { await reference.DisposeAsync(); return; }
        var previous = Client; Client = new(Module!, reference, false); _clientWorkbook = Workbook;
        if (previous is not null) await previous.DisposeAsync();
    }
    private static readonly JsonSerializerOptions Json = new(JsonSerializerDefaults.Web);
    protected override async Task OnBrowserEventAsync(BrowserEvent notification)
    {
        if (IsDisposed) return;
        if (notification.Name == "BindingChanged") await ApplyValueAsync(await InvokeJsonAsync<WorkbookValue>("ReadBinding"));
        else if (notification.Name == "SelectionChanged")
        {
            var value = notification.Data.Deserialize<SelectionInfo>(Json)!;
            if (value.Revision > _selectionVersion)
            {
                _selectionVersion = value.Revision; _lastSelection = value.Address;
                if (Selection != value.Address) await SelectionChanged.InvokeAsync(value.Address);
                if (Worksheet != value.Sheet) await WorksheetChanged.InvokeAsync(value.Sheet);
            }
        }
        else if (notification.Name == "CellEdited") await CellEdited.InvokeAsync(notification.Data.Deserialize<CellEditInfo>(Json)!);
        else if (notification.Name == "ZoomChanged") await ZoomChanged.InvokeAsync(notification.Data.GetProperty("zoom").GetDouble());
        else if (notification.Name == "Error") await Error.InvokeAsync(notification.Data.Deserialize<GridError>(Json)!);
        await base.OnBrowserEventAsync(notification);
    }
    private async Task ApplyValueAsync(WorkbookValue value)
    {
        if (IsDisposed || value.Revision <= _valueVersion) return;
        _valueVersion = value.Revision; _lastValue = value.Value;
        if (Value != value.Value) await ValueChanged.InvokeAsync(value.Value);
    }
    public async ValueTask FlushChangesAsync() => await ApplyValueAsync(await InvokeJsonAsync<WorkbookValue>("FlushChanges"));
    public ValueTask FocusAsync() => InvokeVoidAsync("Focus");
    public ValueTask SelectAsync(string address, string? sheet = null) => InvokeVoidAsync("Select", address, sheet);
    public ValueTask<bool> BeginEditAsync(string? input = null) => InvokeAsync<bool>("BeginEdit", input);
    public ValueTask<bool> CommitEditAsync() => InvokeAsync<bool>("CommitEdit");
    public ValueTask CancelEditAsync() => InvokeVoidAsync("CancelEdit");
    public ValueTask<string> CopySelectionAsync() => InvokeJsonAsync<string>("CopySelection");
    public ValueTask PasteTextAsync(string text) => InvokeVoidAsync("PasteText", text);
    public ValueTask<SelectionInfo> GetSelectionAsync() => InvokeJsonAsync<SelectionInfo>("GetSelection");
    public ValueTask<GridMetrics> GetMetricsAsync() => InvokeJsonAsync<GridMetrics>("GetMetrics");
    public ValueTask<IJSObjectReference> GetNativeControlAsync() => InvokeAsync<IJSObjectReference>("GetNativeControl");
    private Task? _dispose;
    private readonly object _disposeLock = new();
    public override ValueTask DisposeAsync() { lock (_disposeLock) return new(_dispose ??= DisposeCoreAsync()); }
    private async Task DisposeCoreAsync() { try { await base.DisposeAsync(); } finally { if (Client is not null) await Client.DisposeAsync(); } }
}

/// <summary>JSON workbook input participating in EditContext validation and field notifications.</summary>
public sealed class SpreadsheetInput : InputBase<string?>
{
    [Parameter] public string Theme { get; set; } = "light";
    [Parameter] public bool ReadOnly { get; set; }
    [Parameter] public string Style { get; set; } = "display:block;height:480px";
    [Parameter] public string ViewMode { get; set; } = "normal";
    [Parameter] public double Zoom { get; set; } = 1;
    [Parameter] public EventCallback<double> ZoomChanged { get; set; }
    [Parameter] public string? Selection { get; set; }
    [Parameter] public EventCallback<string?> SelectionChanged { get; set; }
    [Parameter] public long SelectionRevision { get; set; }
    [Parameter] public string? Worksheet { get; set; }
    [Parameter] public EventCallback<string?> WorksheetChanged { get; set; }
    [Parameter] public bool ShowGridLines { get; set; } = true;
    [Parameter] public string AriaLabel { get; set; } = "GridWeb spreadsheet";
    [Parameter] public IReadOnlyDictionary<string, object?>? Options { get; set; }
    [Parameter] public long ValueRevision { get; set; }
    [Parameter] public int DebounceMilliseconds { get; set; } = 150;
    [Parameter] public EventCallback<IJSObjectReference> Ready { get; set; }
    [Parameter] public EventCallback<BrowserEvent> Changed { get; set; }
    [Parameter] public EventCallback<CellEditInfo> CellEdited { get; set; }
    [Parameter] public EventCallback<GridError> Error { get; set; }
    public Spreadsheet? Editor { get; private set; }
    protected override bool TryParseValueFromString(string? value, out string? result, out string? validationErrorMessage) { result = value; validationErrorMessage = null; return true; }
    protected override void BuildRenderTree(RenderTreeBuilder b)
    {
        b.OpenComponent<Spreadsheet>(0); b.AddAttribute(1, "Value", CurrentValue);
        b.AddAttribute(2, "ValueChanged", EventCallback.Factory.Create<string?>(this, value => CurrentValue = value));
        b.AddAttribute(3, "ValueRevision", ValueRevision); b.AddAttribute(4, "Theme", Theme); b.AddAttribute(5, "ReadOnly", ReadOnly);
        b.AddAttribute(6, "Style", Style); b.AddAttribute(7, "ViewMode", ViewMode); b.AddAttribute(8, "Zoom", Zoom);
        b.AddAttribute(9, "Ready", Ready); b.AddAttribute(10, "Changed", Changed); b.AddAttribute(11, "CellEdited", CellEdited);
        b.AddAttribute(12, "Error", Error); b.AddAttribute(13, "Class", CssClass); b.AddAttribute(14, "AdditionalAttributes", AdditionalAttributes);
        b.AddAttribute(15, "DebounceMilliseconds", DebounceMilliseconds);
        b.AddAttribute(16, "ZoomChanged", ZoomChanged); b.AddAttribute(17, "Selection", Selection);
        b.AddAttribute(18, "SelectionChanged", SelectionChanged); b.AddAttribute(19, "SelectionRevision", SelectionRevision);
        b.AddAttribute(20, "Worksheet", Worksheet); b.AddAttribute(21, "WorksheetChanged", WorksheetChanged);
        b.AddAttribute(22, "ShowGridLines", ShowGridLines); b.AddAttribute(23, "AriaLabel", AriaLabel); b.AddAttribute(24, "Options", Options);
        b.AddComponentReferenceCapture(25, component => Editor = (Spreadsheet)component); b.CloseComponent();
    }
    public ValueTask FlushAsync() => Editor?.FlushChangesAsync() ?? ValueTask.CompletedTask;
}

/// <summary>Nonvisual, interactive-only module owner for access to every native browser export.</summary>
public sealed class GridProvider : BrowserProvider { }
