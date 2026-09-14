using GridWeb.WinUI;
using Microsoft.UI.Xaml;
namespace GridWeb.Samples.WinUI;
public partial class App : Application
{
 private Window? _window;
 public App() => InitializeComponent();
 protected override void OnLaunched(LaunchActivatedEventArgs args)
 {
  var grid = new GridWebControl(); _window = new Window { Title = "GridWeb · WinUI shared-engine host", Content = grid };
  grid.Ready += async (_, _) => { try { await grid.Client.GetRange("A1:C3").SetValuesAsync([["Region","Units","Unit price"],["North",10,20],["South",12,30]]); await grid.Client.GetRange("D2:D3").SetFormulasAsync([["=B2*C2"],["=B3*C3"]]); } catch (Exception e) { _window.Title = "GridWeb: " + e.Message; } };
  grid.Error += (_, e) => _window.Title = "GridWeb: " + e.Message;
  _window.Closed += async (_, _) => await grid.DisposeAsync(); _window.Activate();
 }
}
