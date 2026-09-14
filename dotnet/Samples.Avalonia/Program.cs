using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Themes.Fluent;
using GridWeb.Avalonia;
namespace GridWeb.Samples.Avalonia;
internal static class Program
{
 [STAThread] public static void Main(string[] args) => AppBuilder.Configure<App>().UsePlatformDetect().LogToTrace().StartWithClassicDesktopLifetime(args);
}
public sealed class App : Application
{
 public override void Initialize() => Styles.Add(new FluentTheme());
 public override void OnFrameworkInitializationCompleted()
 {
  if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
  {
   var grid = new GridWebControl(); var window = new Window { Title = "GridWeb · Avalonia shared-engine host", Width = 1100, Height = 760, Content = grid };
   grid.Ready += async (_, _) => { try { await grid.Client.GetRange("A1:C3").SetValuesAsync([["Region","Plan","Actual"],["North",100,112],["South",90,97]]); await grid.Client.GetRange("D2:D3").SetFormulasAsync([["=C2-B2"],["=C3-B3"]]); await grid.Client.GetRange("A1:D1").SetStyleAsync(new { font = new { bold = true }, fill = "#dceee3" }); } catch (Exception e) { window.Title = "GridWeb: " + e.Message; } };
   grid.Error += (_, e) => window.Title = "GridWeb: " + e.Message;
   window.Closed += async (_, _) => await grid.DisposeAsync(); desktop.MainWindow = window;
  }
  base.OnFrameworkInitializationCompleted();
 }
}
