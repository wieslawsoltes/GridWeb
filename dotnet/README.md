# GridWeb native hosts

GridWeb.Client contains a typed asynchronous spreadsheet API over allowlisted JSON RPC. GridWeb.Wpf, GridWeb.WinUI and GridWeb.Avalonia embed the same built JavaScript engine using WebView2 or Avalonia NativeWebView.

Run `npm run build` at the repository root before `dotnet build`; the client embeds dist/GridWeb-host.html as a resource. Projects target .NET 8; Windows UI projects require the Windows SDK and x64 configuration. Avalonia requires its platform WebView prerequisites. Each Samples.* project is a runnable desktop host.

Controls expose WorkbookJson, Selection, Sheet, Theme, IsReadOnly, Zoom and ViewMode through native dependency/styled properties, plus Ready, WorkbookChanged and Error events. Await Ready or InitializeAsync before using Client. DisposeAsync releases the session/control. WorkbookJson and Selection support view-model synchronization; no native spreadsheet engine or XAML execution is implied.

```csharp
await control.InitializeAsync();
await control.Client.GetRange("A1:B2").SetValuesAsync([[2,3],[4,5]]);
await control.Client.GetRange("C1").SetFormulasAsync([["=SUM(A1:B2)"]]);
var saved = await control.Client.SaveAsync();
```

The HTML host never navigates to user-provided URLs. Navigation/new windows are restricted, native network/filesystem capabilities are not exposed to formulas, and the host protocol uses explicit operation names. Native build, actual native WebView execution and interactive input/accessibility qualification are distinct evidence levels.
