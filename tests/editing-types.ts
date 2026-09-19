// Package declarations are also exercised without relying on JavaScript inferred types.
import {Workbook as TypedWorkbook, captureRange as capture} from '@wieslawsoltes/gridweb';
import {GridWebElement as TypedGrid} from '@wieslawsoltes/gridweb/controls';
import {createExcelApi} from '@wieslawsoltes/gridweb/office';
const book = new TypedWorkbook(), sheet = book.ActiveWorksheet;
const options: import('@wieslawsoltes/gridweb').PasteOptions = {mode:'values', transpose:true, operation:'add', skipBlanks:true};
const clip = capture(sheet.GetRange('A1:B2'));
sheet.GetRange('D1').PasteSpecial(clip, options);
sheet.GetRange('F1:F10').FillSeries({type:'date',dateUnit:'month',start:new Date(2026,0,31),step:1});
const cells: import('@wieslawsoltes/gridweb').CellRange[] = sheet.UsedRange.SpecialCells('formulas',{valueTypes:['numbers']});
const grid = new TypedGrid(); grid.Workbook=book; grid.PasteSpecial(options); grid.FillSeries({type:'growth',step:2});
const dialog: HTMLDialogElement=grid.ShowGoToSpecial();grid.ShowFillSeries();grid.ShowPasteSpecial();grid.FindSpecialCells('errors');
void createExcelApi(book).run(async ctx=>{const s=ctx.workbook.worksheets.getActiveWorksheet();s.getRange('D1').copyFrom(s.getRange('A1'),'Formats',true,true);});
// @ts-expect-error Unsupported modes are rejected statically.
sheet.GetRange('A1').PasteSpecial(clip,{mode:'everything'});
// @ts-expect-error Snapshots cannot be forged from JSON-shaped objects.
const forged: import('@wieslawsoltes/gridweb').ClipboardSnapshot={kind:'GridWebClipboard',rows:1,columns:1,bounds:{r1:0,r2:0,c1:0,c2:0},hasMerges:false};
void cells;void dialog;void forged;
