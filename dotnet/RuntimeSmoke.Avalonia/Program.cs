using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Themes.Fluent;
using GridWeb.Avalonia;
namespace GridWeb.RuntimeSmoke.Avalonia;
internal static class Program
{
 [STAThread] public static int Main(string[] args)
 {
  try { return AppBuilder.Configure<App>().UsePlatformDetect().LogToTrace().StartWithClassicDesktopLifetime(args); }
  catch(Exception error) { Checks.Report("Avalonia", error); return 1; }
 }
}
public sealed class App : Application
{
 public override void Initialize()=>Styles.Add(new FluentTheme());
 public override void OnFrameworkInitializationCompleted()
 {
  if(ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
  {
   var grid=new GridWebControl();var window=new Window {Title="GridWeb Avalonia runtime qualification",Width=1100,Height=800,Content=grid};var done=false;var changes=0;
   async Task Finish(Exception? error){if(done)return;done=true;try{await grid.DisposeAsync();}catch(Exception e){error??=e;}Checks.Report("Avalonia",error);desktop.Shutdown(error is null?0:1);}
   grid.WorkbookChanged+=(_,_)=>changes++;
   grid.Error+=async(_,e)=>await Finish(e);
   grid.Ready+=async(_,_)=>{try{await Checks.RunAsync(grid.Client);await Task.Delay(500);Checks.Require(changes>0,"Native workbook events reach C#");Checks.Require(grid.Selection=="B2:C2","Native Selection property receives JS notifications");Checks.Require(grid.Bounds.Width>0&&grid.Bounds.Height>0,"Native control is laid out");await Finish(null);}catch(Exception e){await Finish(e);}};
   window.Opened+=async(_,_)=>{await Task.Delay(90000);if(!done)await Finish(new TimeoutException("Native WebView did not finish within 90 seconds"));};desktop.MainWindow=window;
  }
  base.OnFrameworkInitializationCompleted();
 }
}
