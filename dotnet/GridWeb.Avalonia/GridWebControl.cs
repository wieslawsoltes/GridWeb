using System.Text.Json;
using global::Avalonia;
using global::Avalonia.Controls;
using global::Avalonia.Data;
using global::Avalonia.Threading;
using GridWeb.Client;
namespace GridWeb.Avalonia;

/// <summary>Avalonia StyledProperty wrapper around NativeWebView and the shared JS engine.</summary>
public sealed class GridWebControl : UserControl, IAsyncDisposable
{
    private readonly NativeWebView _view = new();
    private HostSession? _session;
    private LocalHostDocument? _hostDocument;
    private Task? _initialization;
    private bool _loading;
    private bool _disposed;
    private readonly SemaphoreSlim _propertyUpdates = new(1, 1);
    public SpreadsheetClient Client => _session?.Client ?? throw new InvalidOperationException("Await Ready before using Client.");
    public event EventHandler? Ready;
    public event EventHandler<JsonElement>? WorkbookChanged;
    public event EventHandler<Exception>? Error;
    public static readonly StyledProperty<string?> WorkbookJsonProperty = AvaloniaProperty.Register<GridWebControl, string?>(nameof(WorkbookJson), defaultBindingMode: BindingMode.TwoWay);
    public static readonly StyledProperty<string> SelectionProperty = AvaloniaProperty.Register<GridWebControl, string>(nameof(Selection), "A1", defaultBindingMode: BindingMode.TwoWay);
    public static readonly StyledProperty<string?> SheetProperty = AvaloniaProperty.Register<GridWebControl, string?>(nameof(Sheet));
    public static readonly StyledProperty<string> SpreadsheetThemeProperty = AvaloniaProperty.Register<GridWebControl, string>(nameof(SpreadsheetTheme), "light");
    public static readonly StyledProperty<bool> IsReadOnlyProperty = AvaloniaProperty.Register<GridWebControl, bool>(nameof(IsReadOnly));
    public static readonly StyledProperty<double> ZoomProperty = AvaloniaProperty.Register<GridWebControl, double>(nameof(Zoom), 1.0);
    public static readonly StyledProperty<string> ViewModeProperty = AvaloniaProperty.Register<GridWebControl, string>(nameof(ViewMode), "normal");
    public string? WorkbookJson { get => GetValue(WorkbookJsonProperty); set => SetValue(WorkbookJsonProperty, value); }
    public string Selection { get => GetValue(SelectionProperty); set => SetValue(SelectionProperty, value); }
    public string? Sheet { get => GetValue(SheetProperty); set => SetValue(SheetProperty, value); }
    public string SpreadsheetTheme { get => GetValue(SpreadsheetThemeProperty); set => SetValue(SpreadsheetThemeProperty, value); }
    public bool IsReadOnly { get => GetValue(IsReadOnlyProperty); set => SetValue(IsReadOnlyProperty, value); }
    public double Zoom { get => GetValue(ZoomProperty); set => SetValue(ZoomProperty, value); }
    public string ViewMode { get => GetValue(ViewModeProperty); set => SetValue(ViewModeProperty, value); }
    public GridWebControl()
    {
        Content = _view;
        Loaded += async (_, _) => { try { await InitializeAsync(); } catch (Exception e) { Error?.Invoke(this, e); } };
    }
    protected override async void OnPropertyChanged(AvaloniaPropertyChangedEventArgs change)
    {
        base.OnPropertyChanged(change);
        if (_loading || _session?.IsReady != true) return;
        if (change.Property != WorkbookJsonProperty && change.Property != SelectionProperty && change.Property != SheetProperty && change.Property != SpreadsheetThemeProperty && change.Property != IsReadOnlyProperty && change.Property != ZoomProperty && change.Property != ViewModeProperty) return;
        try { await ApplyAsync(change.Property == WorkbookJsonProperty); } catch (Exception error) { Error?.Invoke(this, error); }
    }
    public Task InitializeAsync() => _initialization ??= InitializeCoreAsync();
    private async Task InitializeCoreAsync()
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        var navigation = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        _view.NavigationStarted += (_, e) => { if (e.Request?.ToString() != _hostDocument?.FileUri.AbsoluteUri) e.Cancel = true; };
        _view.NewWindowRequested += (_, e) => e.Handled = true;
        _view.NavigationCompleted += (_, e) => { if (e.IsSuccess) navigation.TrySetResult(); else navigation.TrySetException(new InvalidOperationException("Host navigation failed")); };
        _hostDocument = new LocalHostDocument();
        _view.Navigate(_hostDocument.FileUri); await navigation.Task.WaitAsync(TimeSpan.FromSeconds(20));
        _session = new HostSession(new DelegateJavaScriptTransport(InvokeOnUiAsync)); _session.Notification += OnNotification; _session.Error += (_, e) => Error?.Invoke(this, e);
        await _session.InitializeAsync(); await ApplyAsync(); Ready?.Invoke(this, EventArgs.Empty);
    }
    private Task<string?> InvokeOnUiAsync(string script, CancellationToken token)
    {
        var result = new TaskCompletionSource<string?>(TaskCreationOptions.RunContinuationsAsynchronously);
        Dispatcher.UIThread.Post(async () => { try { token.ThrowIfCancellationRequested(); result.TrySetResult(await _view.InvokeScript(script)); } catch (Exception error) { result.TrySetException(error); } });
        return result.Task.WaitAsync(token);
    }
    private async Task ApplyAsync(bool loadDocument = true)
    {
        if (_session?.IsReady != true) return;
        await _propertyUpdates.WaitAsync();
        try
        {
            if (loadDocument && !string.IsNullOrWhiteSpace(WorkbookJson)) await Client.LoadAsync(JsonSerializer.Deserialize<JsonElement>(WorkbookJson));
            await Client.SetViewOptionsAsync(SpreadsheetTheme, IsReadOnly, Zoom, ViewMode); await Client.SelectAsync(Selection, Sheet);
        }
        finally { _propertyUpdates.Release(); }
    }
    private async void OnNotification(object? sender, JsonElement message)
    {
        try
        {
            if (message.GetProperty("type").GetString() == "selection-changed") { _loading = true; SetCurrentValue(SelectionProperty, message.GetProperty("address").GetString()!); SetCurrentValue(SheetProperty, message.GetProperty("sheet").GetString()); _loading = false; }
            else if (message.GetProperty("type").GetString() == "workbook-changed") { var json = (await Client.SaveAsync()).GetRawText(); _loading = true; SetCurrentValue(WorkbookJsonProperty, json); _loading = false; WorkbookChanged?.Invoke(this, message); }
        }
        catch (Exception error) { _loading = false; Error?.Invoke(this, error); }
    }
    public async ValueTask DisposeAsync() { if (_disposed) return; _disposed = true; if (_session is not null) await _session.DisposeAsync(); Content = null; _hostDocument?.Dispose(); }
}
