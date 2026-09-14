using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Input;
namespace GridWeb.Samples.Wpf;
public partial class MainWindow : Window
{
 private readonly ViewModel _model = new();
 public MainWindow() { InitializeComponent(); _model.Undo = new Command(async () => { try { await Spreadsheet.Client.UndoAsync(); } catch (Exception e) { MessageBox.Show(e.Message); } }); DataContext = _model; Closed += async (_, _) => await Spreadsheet.DisposeAsync(); }
 private async void SpreadsheetReady(object? sender, EventArgs e) { try { var client = Spreadsheet.Client; await client.GetRange("A1:C4").SetValuesAsync([["Item","Units","Price"],["North",12,8.5],["South",9,12.0],["Total",null,null]]); await client.GetRange("D1").SetValuesAsync([["Revenue"]]); await client.GetRange("D2:D4").SetFormulasAsync([["=B2*C2"],["=B3*C3"],["=SUM(D2:D3)"]]); await client.GetRange("A1:D1").SetStyleAsync(new { font = new { bold = true, color = "#ffffff" }, fill = "#107c41" }); } catch (Exception error) { MessageBox.Show(error.Message); } }
 private void SpreadsheetError(object? sender, Exception e) => MessageBox.Show(e.Message, "GridWeb host error");
 private async void SaveClicked(object sender, RoutedEventArgs e) { try { var dialog = new Microsoft.Win32.SaveFileDialog { Filter = "GridWeb JSON|*.gridweb.json" }; if (dialog.ShowDialog() == true) await System.IO.File.WriteAllTextAsync(dialog.FileName, (await Spreadsheet.Client.SaveAsync()).GetRawText()); } catch (Exception error) { MessageBox.Show(error.Message); } }
}
public sealed class ViewModel : INotifyPropertyChanged
{
 private string _selection = "A1"; private bool _readOnly;
 public string Selection { get => _selection; set { _selection = value; Changed(); } }
 public bool ReadOnly { get => _readOnly; set { _readOnly = value; Changed(); } }
 public ICommand? Undo { get; set; }
 public event PropertyChangedEventHandler? PropertyChanged;
 private void Changed([CallerMemberName] string? name = null) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}
public sealed class Command(Action run) : ICommand { public bool CanExecute(object? parameter) => true; public void Execute(object? parameter) => run(); public event EventHandler? CanExecuteChanged { add { } remove { } } }
