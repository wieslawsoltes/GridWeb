using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using GridWeb.Client;
using Microsoft.Web.WebView2.Wpf;
namespace GridWeb.Wpf;

/// <summary>Native WPF dependency properties with the shared browser engine in WebView2.</summary>
public sealed class GridWebControl : UserControl, IAsyncDisposable
{
    private readonly WebView2 _view = new();
    private HostSession? _session;
    private Task? _initialization;
    private bool _loading;
    private bool _disposed;
    private readonly SemaphoreSlim _propertyUpdates = new(1, 1);
    public SpreadsheetClient Client => _session?.Client ?? throw new InvalidOperationException("Await Ready before accessing Client.");
    public event EventHandler? Ready;
    public event EventHandler<JsonElement>? WorkbookChanged;
    public event EventHandler<Exception>? Error;
    public static readonly DependencyProperty WorkbookJsonProperty = DependencyProperty.Register(nameof(WorkbookJson), typeof(string), typeof(GridWebControl), new FrameworkPropertyMetadata(null, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, Changed));
    public static readonly DependencyProperty SelectionProperty = DependencyProperty.Register(nameof(Selection), typeof(string), typeof(GridWebControl), new FrameworkPropertyMetadata("A1", FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, Changed));
    public static readonly DependencyProperty SheetProperty = DependencyProperty.Register(nameof(Sheet), typeof(string), typeof(GridWebControl), new PropertyMetadata(null, Changed));
    public static readonly DependencyProperty ThemeProperty = DependencyProperty.Register(nameof(Theme), typeof(string), typeof(GridWebControl), new PropertyMetadata("light", Changed));
    public static readonly DependencyProperty IsReadOnlyProperty = DependencyProperty.Register(nameof(IsReadOnly), typeof(bool), typeof(GridWebControl), new PropertyMetadata(false, Changed));
    public static readonly DependencyProperty ZoomProperty = DependencyProperty.Register(nameof(Zoom), typeof(double), typeof(GridWebControl), new PropertyMetadata(1.0, Changed));
    public static readonly DependencyProperty ViewModeProperty = DependencyProperty.Register(nameof(ViewMode), typeof(string), typeof(GridWebControl), new PropertyMetadata("normal", Changed));
    public string? WorkbookJson { get => (string?)GetValue(WorkbookJsonProperty); set => SetValue(WorkbookJsonProperty, value); }
    public string Selection { get => (string)GetValue(SelectionProperty); set => SetValue(SelectionProperty, value); }
    public string? Sheet { get => (string?)GetValue(SheetProperty); set => SetValue(SheetProperty, value); }
    public string Theme { get => (string)GetValue(ThemeProperty); set => SetValue(ThemeProperty, value); }
    public bool IsReadOnly { get => (bool)GetValue(IsReadOnlyProperty); set => SetValue(IsReadOnlyProperty, value); }
    public double Zoom { get => (double)GetValue(ZoomProperty); set => SetValue(ZoomProperty, value); }
    public string ViewMode { get => (string)GetValue(ViewModeProperty); set => SetValue(ViewModeProperty, value); }
    public GridWebControl() { Content = _view; Loaded += async (_, _) => { try { await InitializeAsync(); } catch (Exception e) { Error?.Invoke(this, e); } }; }
    public Task InitializeAsync() => _initialization ??= InitializeCoreAsync();
    private async Task InitializeCoreAsync()
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        await _view.EnsureCoreWebView2Async();
        _view.CoreWebView2.Settings.AreDevToolsEnabled = false;
        _view.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
        _view.CoreWebView2.Settings.AreHostObjectsAllowed = false;
        _view.CoreWebView2.NavigationStarting += (_, e) => { if (e.Uri != "about:blank") e.Cancel = true; };
        _view.CoreWebView2.NewWindowRequested += (_, e) => e.Handled = true;
        _view.CoreWebView2.PermissionRequested += (_, e) => e.State = Microsoft.Web.WebView2.Core.CoreWebView2PermissionState.Deny;
        var navigation = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        _view.NavigationCompleted += (_, e) => { if (e.IsSuccess) navigation.TrySetResult(); else navigation.TrySetException(new InvalidOperationException("Host navigation failed")); };
        _view.NavigateToString(HostSession.GetHtml()); await navigation.Task.WaitAsync(TimeSpan.FromSeconds(20));
        var transport = new DelegateJavaScriptTransport(async (script, token) =>
        {
            token.ThrowIfCancellationRequested();
            if (Dispatcher.CheckAccess()) return await _view.ExecuteScriptAsync(script).WaitAsync(token);
            return await Dispatcher.InvokeAsync(() => _view.ExecuteScriptAsync(script)).Task.Unwrap().WaitAsync(token);
        });
        _session = new HostSession(transport); _session.Notification += OnNotification; _session.Error += (_, e) => Error?.Invoke(this, e);
        await _session.InitializeAsync(); await ApplyAsync(); Ready?.Invoke(this, EventArgs.Empty);
    }
    private static async void Changed(DependencyObject sender, DependencyPropertyChangedEventArgs e)
    {
        var control = (GridWebControl)sender;
        if (control._loading || control._session?.IsReady != true) return;
        try { await control.ApplyAsync(e.Property == WorkbookJsonProperty); } catch (Exception error) { control.Error?.Invoke(control, error); }
    }
    private async Task ApplyAsync(bool loadDocument = true)
    {
        if (_session?.IsReady != true) return;
        await _propertyUpdates.WaitAsync();
        try
        {
            if (loadDocument && !string.IsNullOrWhiteSpace(WorkbookJson)) await Client.LoadAsync(JsonSerializer.Deserialize<JsonElement>(WorkbookJson));
            await Client.SetViewOptionsAsync(Theme, IsReadOnly, Zoom, ViewMode); await Client.SelectAsync(Selection, Sheet);
        }
        finally { _propertyUpdates.Release(); }
    }
    private async void OnNotification(object? sender, JsonElement message)
    {
        try
        {
            if (message.GetProperty("type").GetString() == "selection-changed")
            { _loading = true; SetCurrentValue(SelectionProperty, message.GetProperty("address").GetString()); SetCurrentValue(SheetProperty, message.GetProperty("sheet").GetString()); _loading = false; }
            else if (message.GetProperty("type").GetString() == "workbook-changed")
            { var json = (await Client.SaveAsync()).GetRawText(); _loading = true; SetCurrentValue(WorkbookJsonProperty, json); _loading = false; WorkbookChanged?.Invoke(this, message); }
        }
        catch (Exception error) { _loading = false; Error?.Invoke(this, error); }
    }
    public async ValueTask DisposeAsync()
    {
        if (_disposed) return; _disposed = true;
        if (_session is not null) await _session.DisposeAsync(); _view.Dispose();
    }
}
