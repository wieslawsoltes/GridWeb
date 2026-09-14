using System.Text.Json;
using GridWeb.Client;
namespace GridWeb.RuntimeSmoke;
internal static class Checks
{
    internal static readonly List<string> Passed = [];
    internal static void Require(bool value, string name) { if (!value) throw new InvalidOperationException(name); Passed.Add(name); }
    internal static async Task RunAsync(SpreadsheetClient client)
    {
        var range = client.GetRange("A1:B2");
        await range.SetValuesAsync([[2,3],[4,5]]);
        await client.GetRange("C1").SetFormulasAsync([["=SUM(A1:B2)"]]);
        Require((await client.GetRange("C1").GetValuesAsync())[0][0].GetDouble() == 14, "Native WebView computes shared-engine formulas");
        await client.GetRange("A1").SetValuesAsync([[8]]);
        Require((await client.GetRange("C1").GetValuesAsync())[0][0].GetDouble() == 20, "Native edit recalculates dependencies");
        await client.UndoAsync();
        Require((await client.GetRange("C1").GetValuesAsync())[0][0].GetDouble() == 14, "Native undo restores calculation");
        await client.RedoAsync();
        Require((await client.GetRange("C1").GetValuesAsync())[0][0].GetDouble() == 20, "Native redo restores edit");
        await client.GetRange("D1").SetValuesAsync([["\"quoted\" <text> \\ \n Zażółć"]]);
        Require((await client.GetRange("D1").GetValuesAsync())[0][0].GetString() == "\"quoted\" <text> \\ \n Zażółć", "RPC retains Unicode and escaped text as data");
        await client.GetRange("C1").SetStyleAsync(new { numberFormat = "0.00", font = new { bold = true } });
        Require((await client.GetRange("C1").GetTextAsync())[0][0].GetString() == "20.00", "Native formatting uses shared formatter");
        var snapshot = await client.SaveAsync();
        await client.LoadAsync(snapshot);
        Require((await client.GetRange("C1").GetValuesAsync())[0][0].GetDouble() == 20, "Native JSON save/load retains calculation");
        await client.SelectAsync("B2:C2");
        Require((await client.InvokeAsync("view.selection.get")).GetString() == "B2:C2", "Native client selects an actual range");
        await client.SetViewOptionsAsync("dark", false, 1.2, "pageLayout");
        await Task.Delay(300);
        await client.SetViewOptionsAsync("light", false, 1, "normal");
        Require((await client.InvokeAsync("capabilities")).GetProperty("formulaFunctions").GetArrayLength() >= 330, "Native host includes the current calculation engine");
        var rejected = false;
        try { await client.InvokeAsync("not-an-allowed-method"); } catch (SpreadsheetException) { rejected = true; }
        Require(rejected, "Native RPC rejects unknown operations");
    }
    internal static void Report(string framework, Exception? error)
    {
        var target = Environment.GetEnvironmentVariable("GRIDWEB_SMOKE_REPORT") ?? Path.Combine("test-results", "native-"+framework+".json");
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(target))!);
        File.WriteAllText(target, JsonSerializer.Serialize(new { framework, passed = Passed, error = error?.ToString(), platform = System.Runtime.InteropServices.RuntimeInformation.OSDescription, runtime = Environment.Version.ToString(), qualification = "Actual native WebView startup and RPC smoke; not exhaustive native UI/device/accessibility qualification" }, new JsonSerializerOptions { WriteIndented = true }));
    }
}
