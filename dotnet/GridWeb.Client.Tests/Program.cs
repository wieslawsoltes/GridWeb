using System.Text.Json;
using GridWeb.Client;
static void Require(bool condition, string message) { if (!condition) throw new Exception(message); }
var count = 0;
var transport = new DelegateJavaScriptTransport((script, token) =>
{
    token.ThrowIfCancellationRequested();
    const string prefix = "globalThis.gridWebHost.dispatch(";
    Require(script.StartsWith(prefix) && script.EndsWith(')'), "Fixed allowlisted dispatch expression");
    var json = JsonSerializer.Deserialize<string>(script[prefix.Length..^1])!;
    using var request = JsonDocument.Parse(json);
    var root = request.RootElement;
    Require(root.GetProperty("id").GetInt64() > 0, "Monotonic request id");
    var result = JsonSerializer.Serialize(new { id = root.GetProperty("id").GetInt64(), result = new[] { new object[] { 42 } } });
    count++;
    return Task.FromResult<string?>(JsonSerializer.Serialize(result));
});
var client = new SpreadsheetClient(transport);
var values = await client.GetRange("A1", "Sheet1").GetValuesAsync();
Require(values[0][0].GetInt32() == 42, "Double encoded WebView2 return decoded");
await client.GetRange("A1").SetValuesAsync([["'); throw new Error('not code'); //"]]);
Require(count == 2, "Two transport calls");
Require(SpreadsheetClient.Decode("{\"a\":1}").GetProperty("a").GetInt32() == 1, "Direct JSON decoded");
var mismatch = new SpreadsheetClient(new DelegateJavaScriptTransport((_, _) => Task.FromResult<string?>("{\"id\":999,\"result\":0}")));
try { await mismatch.SaveAsync(); throw new Exception("Expected mismatch rejection"); } catch (SpreadsheetException) { }
var failed = new SpreadsheetClient(new DelegateJavaScriptTransport((_, _) => Task.FromResult<string?>("{\"id\":1,\"error\":{\"message\":\"blocked\"}}")));
try { await failed.SaveAsync(); throw new Exception("Expected host error"); } catch (SpreadsheetException e) { Require(e.Message == "blocked", "Error surfaced"); }
Console.WriteLine("GridWeb.Client protocol tests passed (not a native WebView runtime test).");
