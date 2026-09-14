using Microsoft.UI.Xaml;
using GridWeb.WinUI;
namespace GridWeb.RuntimeSmoke.WinUI;
public partial class App : Application
{
 private Window? _window; private bool _done;
 public App() => InitializeComponent();
 protected override void OnLaunched(LaunchActivatedEventArgs args)
 {
  var grid=new GridWebControl(); _window=new Window { Title="GridWeb WinUI runtime qualification", Content=grid };var changes=0;
  async Task Finish(Exception? error){if(_done)return;_done=true;try{await grid.DisposeAsync();}catch(Exception e){error??=e;}Checks.Report("WinUI",error);Environment.ExitCode=error is null?0:1;Exit();}
  grid.WorkbookChanged+=(_,_)=>changes++;
  grid.Error+=async(_,e)=>await Finish(e);
  grid.Ready+=async(_,_)=>{try{await Checks.RunAsync(grid.Client);await Task.Delay(500);Checks.Require(changes>0,"Native workbook events reach C#");Checks.Require(grid.Selection=="B2:C2","Native Selection property receives JS notifications");Checks.Require(grid.ActualWidth>0&&grid.ActualHeight>0,"Native control is laid out");await Finish(null);}catch(Exception e){await Finish(e);}};
  _window.Activate();_=Deadline();
  async Task Deadline(){await Task.Delay(90000);if(!_done)await Finish(new TimeoutException("Native WebView did not finish within 90 seconds"));}
 }
}
