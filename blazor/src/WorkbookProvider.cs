using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Rendering;
using Microsoft.JSInterop;
namespace GridWeb.Blazor;

/// <summary>Owns one browser workbook and lends it to child grids. Prerendering does not invoke JavaScript.</summary>
public sealed class WorkbookProvider : ComponentBase, IAsyncDisposable
{
    [Inject] private IJSRuntime JS { get; set; } = default!;
    [Parameter] public string? InitialValue { get; set; }
    [Parameter] public RenderFragment<GridWorkbook>? ChildContent { get; set; }
    [Parameter] public RenderFragment? Loading { get; set; }
    [Parameter] public EventCallback<GridWorkbook> Ready { get; set; }
    public GridWorkbook? Workbook { get; private set; }
    private GridWebModule? _module;
    private Task? _initialization, _disposal;
    private bool _disposed, _ready;
    protected override void BuildRenderTree(RenderTreeBuilder b) { if (_ready && Workbook is not null) b.AddContent(0, ChildContent?.Invoke(Workbook)); else b.AddContent(1, Loading); }
    protected override async Task OnAfterRenderAsync(bool firstRender)
    {
        if (!firstRender || _disposed) return;
        await (_initialization ??= InitializeAsync());
        if (_disposed) return;
        _ready = true; await Ready.InvokeAsync(Workbook!);
        if (!_disposed) StateHasChanged();
    }
    private async Task InitializeAsync()
    {
        _module = new(JS);
        Workbook = await _module.CreateWorkbookAsync(InitialValue);
    }
    public ValueTask DisposeAsync() { _disposed = true; return new(_disposal ??= DisposeCoreAsync()); }
    private async Task DisposeCoreAsync()
    {
        try { if (_initialization is not null) await _initialization; }
        catch (JSDisconnectedException) { }
        finally { try { if (Workbook is not null) await Workbook.DisposeAsync(); } finally { if (_module is not null) await _module.DisposeAsync(); } }
    }
}
