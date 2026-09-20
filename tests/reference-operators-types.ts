import {Workbook, a1ToR1C1, r1c1ToA1} from '@wieslawsoltes/gridweb';
import {GridWebElement} from '@wieslawsoltes/gridweb/controls';
import {createExcelApi} from '@wieslawsoltes/gridweb/office';
const book=new Workbook();
const jan=book.Worksheets.Add('Jan',0), feb=book.Worksheets.add('Feb',1);
const moved=book.Worksheets.Move(feb,0);
book.Worksheets.Move('Jan',0);
jan.GetCell('A1').Formula='=SUM(Jan:Feb!B1:B10)';
const reference=a1ToR1C1('=SUM(Jan:Feb!A1:A3)','D4');r1c1ToA1(reference,'D4');
const grid=new GridWebElement();grid.Workbook=book;grid.MoveWorksheet(1,jan);grid.MoveWorksheet(0,'Jan');
const dialog:HTMLDialogElement=grid.ShowMoveWorksheet();
void createExcelApi(book).run(async ctx=>{const sheet=ctx.workbook.worksheets.getItem('Jan');sheet.position=1;sheet.load('position');await ctx.sync();const index:number=sheet.position;void index;});
// @ts-expect-error Indexes are numeric, not string positions.
book.Worksheets.Move(moved,'1');
// @ts-expect-error Optional insertion position is numeric.
book.Worksheets.Add('Invalid','0');
void dialog;
