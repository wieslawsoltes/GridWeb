import {Workbook} from './model.js';
import {exportXlsx as exportModel,importXlsx as importModel} from './io-base.js';
import {addPivotParts,importPivotParts,checkNative} from './pivots/ooxml.js';
export {exportCSV,importCSV,parseDelimited} from './io-base.js';
export function exportXlsx(book,options={}){
 if(!book.PivotTables.Count)return exportModel(book,options);
 if(options.pivots&&!['native','flatten'].includes(options.pivots))throw new TypeError('Unknown pivot export mode');
 if(options.pivots!=='flatten')for(const report of book.PivotTables)checkNative(report);
 const snapshot=Workbook.FromJSON(book.ToJSON());
 try{snapshot.PivotTables.RefreshAll();const bytes=exportModel(snapshot,options);if(options.pivots==='flatten'){options.onWarning?.('Managed pivot definitions are explicitly flattened into refreshed worksheet cells');return bytes;}return addPivotParts(snapshot,bytes);}finally{snapshot.Dispose();}
}
export async function importXlsx(bytes){const result=await importModel(bytes);await importPivotParts(result.workbook,bytes,result.warnings);return result;}
