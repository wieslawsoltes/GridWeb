import {Workbook,PivotReport} from '@wieslawsoltes/gridweb';
import {exportXlsx} from '@wieslawsoltes/gridweb/io';
const book=new Workbook(),source=book.ActiveWorksheet,target=book.Worksheets.Add('Pivot');
source.GetRange('A1:B2').Values=[['Key','Value'],['A',4]];
const report:PivotReport=book.PivotTables.Add('Summary',source.UsedRange,target.GetRange('A1'),{rows:['Key'],values:[{column:'Value',aggregate:'sum'}]});
report.Refresh();report.SetFilter('Key',['A']);report.DrillDown(1,1);exportXlsx(book,{pivots:'native'});
// @ts-expect-error unsupported aggregation cannot be advertised as implemented
report.Update({options:{values:[{column:1,aggregate:'magical'}]}});
