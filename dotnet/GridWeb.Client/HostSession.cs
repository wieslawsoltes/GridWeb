using System.Diagnostics;
using System.Reflection;
using System.Text.Json;
namespace GridWeb.Client;

/// <summary>Readiness and disposable event delivery shared by WPF, WinUI and Avalonia hosts.</summary>
public sealed class HostSession(IJavaScriptTransport transport) : IAsyncDisposable
{
    private readonly CancellationTokenSource _lifetime = new();
    private Task? _polling;
    public SpreadsheetClient Client { get; } = new(transport);
    public bool IsReady { get; private set; }
    public event EventHandler<JsonElement>? Notification;
    public event EventHandler<Exception>? Error;
    public static string GetHtml()
    {
        using var stream = typeof(HostSession).Assembly.GetManifestResourceStream("GridWeb.Host.html")
            ?? throw new InvalidOperationException("The host HTML resource was not embedded. Run npm run build before packing.");
        using var reader = new StreamReader(stream); return reader.ReadToEnd();
    }
    public async Task InitializeAsync(CancellationToken cancellationToken = default)
    {
        if (IsReady) return;
        using var linked = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, _lifetime.Token);
        var clock = Stopwatch.StartNew();
        while (clock.Elapsed < TimeSpan.FromSeconds(20))
        {
            linked.Token.ThrowIfCancellationRequested();
            try
            {
                var raw = await transport.InvokeAsync("Boolean(globalThis.gridWebHost)", linked.Token);
                if (SpreadsheetClient.Decode(raw).ValueKind == JsonValueKind.True) { IsReady = true; break; }
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception) when (clock.Elapsed < TimeSpan.FromSeconds(19)) { }
            await Task.Delay(50, linked.Token);
        }
        if (!IsReady) throw new TimeoutException("GridWeb JavaScript host did not become ready.");
        _polling = PollAsync(_lifetime.Token);
    }
    private async Task PollAsync(CancellationToken token)
    {
        using var timer = new PeriodicTimer(TimeSpan.FromMilliseconds(100));
        try
        {
            while (await timer.WaitForNextTickAsync(token))
            {
                var raw = await transport.InvokeAsync("JSON.stringify(globalThis.gridWebHost.drainEvents())", token);
                var messages = SpreadsheetClient.Decode(raw);
                if (messages.ValueKind != JsonValueKind.Array) throw new SpreadsheetException("Invalid event queue response");
                foreach (var message in messages.EnumerateArray()) Notification?.Invoke(this, message.Clone());
            }
        }
        catch (OperationCanceledException) when (token.IsCancellationRequested) { }
        catch (Exception error) { Error?.Invoke(this, error); }
    }
    public async ValueTask DisposeAsync()
    {
        _lifetime.Cancel();
        if (_polling is not null) await _polling;
        IsReady = false; _lifetime.Dispose();
    }
}
