using System.IO;
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
        var functions = (await client.InvokeAsync("capabilities")).GetProperty("formulaFunctions").EnumerateArray().Select(v => v.GetString()).ToHashSet();
        Require(functions.Count >= 356 && functions.Contains("AREAS") && functions.Contains("SHEETS") && functions.Contains("AVERAGEA"), "Native host includes the current calculation engine");
        await client.GetRange("F1:G3").SetValuesAsync([["Region", "Revenue"], ["North", 10], ["North", 20]]);
        await client.AddPivotAsync("NativePivot", "F1:G3", "J1", new PivotOptions { Rows = ["Region"], Values = [new PivotValue("Revenue")] });
        Require((await client.ListPivotsAsync()).GetArrayLength() == 1, "Native client creates a managed pivot");
        Require((await client.RefreshPivotAsync("NativePivot"))[1][1].GetDouble() == 30, "Native pivot refresh uses the shared aggregation engine");
        Require((await client.GetPivotDetailsAsync("NativePivot", 1, 1)).GetArrayLength() == 3, "Native pivot drill-down returns source rows");
        await client.AddWorksheetAsync("Reference start");
        await client.AddWorksheetAsync("Reference end");
        await client.AddWorksheetAtAsync("Reference middle", 2);
        await client.GetRange("A1", "Reference start").SetValuesAsync([[1]]);
        await client.GetRange("A1", "Reference middle").SetValuesAsync([[2]]);
        await client.GetRange("A1", "Reference end").SetValuesAsync([[3]]);
        await client.GetRange("P1").SetFormulasAsync([["=SUM('Reference start:Reference end'!A1)"]]);
        Require((await client.GetRange("P1").GetValuesAsync())[0][0].GetDouble() == 6, "Native indexed insertion participates in 3-D totals");
        await client.MoveWorksheetAsync("Reference middle", 3);
        Require((await client.GetRange("P1").GetValuesAsync())[0][0].GetDouble() == 4, "Native typed sheet movement recalculates 3-D membership");
        await client.UndoAsync();
        Require((await client.GetRange("P1").GetValuesAsync())[0][0].GetDouble() == 6, "Native sheet movement undo restores reference dependencies");
        await client.GetRange("P2:P4").SetFormulasAsync([["=SUM((A1,A2))"], ["=SUM(A1:INDEX(A1:B2,2,1))"], ["=AVERAGEA({2,TRUE,\"text\"})"]]);
        var references = await client.GetRange("P2:P4").GetValuesAsync();
        Require(references[0][0].GetDouble() == 12, "Native union formula retains reference values");
        Require(references[1][0].GetDouble() == 12, "Native INDEX can form a dynamic range endpoint");
        Require(references[2][0].GetDouble() == 1, "Native inclusive statistics distinguish text and logical inputs");
        await client.DefineScopedNameAsync("NativeRate", 2);
        await client.DefineScopedNameAsync("NativeRate", 3, "Reference start", new DefinedNameOptions { Comment = "Local rate" });
        await client.GetRange("Q1").SetFormulasAsync([["='Reference start'!NativeRate+NativeRate"]]);
        Require((await client.GetRange("Q1").GetValuesAsync())[0][0].GetDouble() == 5, "Native worksheet names coexist with workbook names");
        await client.DefineScopedNameAsync("ScopedRange", "=A1", "Reference start");
        await client.GetRange("Q2").SetFormulasAsync([["=LET(data,'Reference start'!ScopedRange,ISREF(data))"]]);
        Require((await client.GetRange("Q2").GetValuesAsync())[0][0].GetBoolean(), "Native LET keeps scoped reference identity");
        await client.RenameDefinedNameAsync("NativeRate", "RenamedRate", "Reference start");
        Require((await client.GetRange("Q1").GetValuesAsync())[0][0].GetDouble() == 5 && (await client.GetDefinedNameAsync("RenamedRate", "Reference start")).GetProperty("comment").GetString() == "Local rate", "Native rename repairs formulas and preserves metadata");
        await client.UndoAsync();
        Require((await client.GetDefinedNameAsync("NativeRate", "Reference start")).GetProperty("value").GetDouble() == 3, "Native undo restores scoped name identity");
        var namesSnapshot = await client.SaveAsync(); await client.LoadAsync(namesSnapshot);
        Require((await client.GetRange("Q1").GetValuesAsync())[0][0].GetDouble() == 5 && (await client.ListDefinedNamesAsync(all: true)).GetArrayLength() >= 3, "Native JSON retains named formulas and scopes");
        await client.RemoveDefinedNameAsync("NativeRate", "Reference start");
        Require((await client.GetDefinedNameAsync("NativeRate", "Reference start")).ValueKind == JsonValueKind.Null, "Native removal targets only the selected name scope");
        await client.RemoveDefinedNameAsync("NativeRate");
        await client.GetRange("Q1:Q2").ClearAsync();
        await client.GetRange("P1:P4").ClearAsync();
        foreach (var sheet in new[] { "Reference start", "Reference middle", "Reference end" }) await client.InvokeAsync("worksheets.remove", new { sheet });
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
