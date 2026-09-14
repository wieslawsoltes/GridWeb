using System.Windows;
using GridWeb.Wpf;
using GridWeb.RuntimeSmoke;
internal static class Program
{
 [STAThread] public static int Main()
 {
  var app = new Application { ShutdownMode = ShutdownMode.OnExplicitShutdown };
  var grid = new GridWebControl(); var window = new Window { Title = "GridWeb WPF runtime qualification", Content = grid, Width = 1100, Height = 800 };
  var done = false; var changes = 0;
  async Task Finish(Exception? error) { if (done) return; done = true; try { await grid.DisposeAsync(); } catch (Exception e) { error ??= e; } Checks.Report("Wpf", error); app.Shutdown(error is null ? 0 : 1); }
  grid.WorkbookChanged += (_, _) => changes++;
  grid.Error += async (_, e) => await Finish(e);
  grid.Ready += async (_, _) => { try { await Checks.RunAsync(grid.Client); await Task.Delay(500); Checks.Require(changes>0, "Native workbook events reach C#"); Checks.Require(grid.Selection=="B2:C2", "Native Selection property receives JS notifications"); Checks.Require(grid.ActualWidth>0 && grid.ActualHeight>0, "Native control is laid out"); await Finish(null); } catch (Exception e) { await Finish(e); } };
  app.Startup += async (_, _) => { window.Show(); await Task.Delay(90000); if (!done) await Finish(new TimeoutException("Native WebView did not finish within 90 seconds")); };
  return app.Run();
 }
}
