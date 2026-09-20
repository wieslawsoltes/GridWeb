using System.Text.Json;
using GridWeb.Client;
static void Require(bool condition, string message) { if (!condition) throw new Exception(message); }
var count = 0;
var requests = new List<JsonElement>();
var transport = new DelegateJavaScriptTransport((script, token) =>
{
    token.ThrowIfCancellationRequested();
    const string prefix = "globalThis.gridWebHost.dispatch(";
    Require(script.StartsWith(prefix) && script.EndsWith(')'), "Fixed allowlisted dispatch expression");
    var json = JsonSerializer.Deserialize<string>(script[prefix.Length..^1])!;
    using var request = JsonDocument.Parse(json);
    var root = request.RootElement;
    requests.Add(root.Clone());
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
await client.AddWorksheetAtAsync("Inserted", 1);
Require(requests[^1].GetProperty("method").GetString() == "worksheets.add", "Insert worksheet route");
Require(requests[^1].GetProperty("name").GetString() == "Inserted" && requests[^1].GetProperty("index").GetInt32() == 1, "Insert worksheet payload");
await client.MoveWorksheetAsync("Sheet1", 0);
Require(requests[^1].GetProperty("method").GetString() == "worksheets.move", "Move worksheet route");
Require(requests[^1].GetProperty("sheet").GetString() == "Sheet1" && requests[^1].GetProperty("index").GetInt32() == 0, "Move worksheet payload");
var callsBeforeInvalid = count;
try { await client.MoveWorksheetAsync("Sheet1", -1); throw new Exception("Expected negative index rejection"); } catch (ArgumentOutOfRangeException) { }
try { await client.AddWorksheetAtAsync("", 0); throw new Exception("Expected blank name rejection"); } catch (ArgumentException) { }
using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
try { await client.MoveWorksheetAsync("Sheet1", 0, cancellation.Token); throw new Exception("Expected cancellation"); } catch (OperationCanceledException) { }
Require(count == callsBeforeInvalid, "Invalid or cancelled calls do not invoke the transport");
Console.WriteLine("GridWeb.Client protocol and worksheet command tests passed (not a native WebView runtime test).");
