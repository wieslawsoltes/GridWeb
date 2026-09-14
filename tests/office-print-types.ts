import { Workbook } from '@wieslawsoltes/gridweb';
import { createExcelApi } from '@wieslawsoltes/gridweb/office';
import { paginate, chartToSVG } from '@wieslawsoltes/gridweb/printing';
const book=new Workbook();
const Excel=createExcelApi(book);
await Excel.run(async ctx=>{const r=ctx.workbook.worksheets.getActiveWorksheet().getRange('A1:B2');r.values=[[1,2],[3,4]];r.numberFormat=[[null,'0.00'],[null,'0.00']];r.format.font.bold=true;r.load('values');await ctx.sync();const values=r.values;console.log(values);});
const pages=paginate(book.ActiveWorksheet,{paper:'A3',repeatColumns:1,fitToWidthPages:1,margins:{left:20}});
console.log(pages,chartToSVG);
// @ts-expect-error a workbook is required
createExcelApi({});
