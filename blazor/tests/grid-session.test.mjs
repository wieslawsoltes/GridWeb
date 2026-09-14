import { test } from 'node:test';
import assert from 'node:assert/strict';
import { SpreadsheetSession } from '../src/session.js';
import { Workbook } from '../../src/model.js';
import { Session } from '../src/wwwroot/interop.js';
import { createWorkerHandler } from '../../src/worker.js';

test('real ranges, calculations, undo and stream-sized Unicode DTOs', async () => {
  const s = new SpreadsheetSession(), notices = []; s.Changed.Subscribe(x => notices.push(x));
  s.Invoke('range.values.set',{address:'A1:B2',values:[[1,2],[3,4]]});
  s.Invoke('range.formulas.set',{address:'C1',formulas:[['=SUM(A1:B2)']]});
  assert.equal(s.Invoke('range.values.get',{address:'C1'})[0][0],10);
  const text = 'Zażółć🙂'.repeat(2000); s.GetRange('D1:D8').Values=Array.from({length:8},()=>[text]);
  assert.equal(s.Save().sheets[0].cells.length,13);assert.ok(s.ReadBinding().value.length>32000);
  assert.ok(notices.length);assert.equal(typeof notices[0].revision,'number');assert.ok(!JSON.stringify(notices).includes(text));
  assert.equal(s.Invoke('workbook.undo'),true);assert.equal(s.GetRange('D1').Value,null);assert.equal(s.Invoke('workbook.redo'),true);
  await s.DisposeAsync(); assert.throws(()=>s.Save(),/disposed/);
});
test('failed batch rolls back every mutation and validates its entire method list', async () => {
  const s=new SpreadsheetSession();s.GetRange('A1').Value=7;
  assert.throws(()=>s.Batch([{method:'range.values.set',arguments:{address:'A1',values:[[9]]}},{method:'worksheets.remove',arguments:{sheet:'Sheet1'}}]),/worksheet/);
  assert.equal(s.GetRange('A1').Value,7);
  assert.throws(()=>s.Batch([{method:'range.clear',arguments:{address:'A1'}},{method:'workbook.new'}]),/atomic/);assert.equal(s.GetRange('A1').Value,7);
  await s.DisposeAsync();
});
test('CSV defaults to literal formulas, respects start address, and round-trips quoting', async () => {
  const s=new SpreadsheetSession();assert.equal(s.ImportCsv('Name,Value\n"One,Two",=1+1',null,'E2'),'E2:F3');
  assert.equal(s.GetRange('F3').Value,'=1+1');assert.equal(s.GetRange('F3').Formula,null);
  assert.match(s.ExportCsv(null,'E2:F3'),/One,Two/);await s.DisposeAsync();
});
test('JSON replacement is validated before replacing the live workbook', async () => {
  const s=new SpreadsheetSession();s.GetRange('A1').Value=9;const before=s.GetWorkbook();
  assert.throws(()=>s.Load({format:'bad'}),/Unsupported/);assert.equal(s.GetWorkbook(),before);
  const doc=s.Save();s.New();assert.equal(s.GetRange('A1').Value,null);s.Load(doc);assert.equal(s.GetRange('A1').Value,9);
  assert.ok(s.Revision>1);await s.DisposeAsync();
});
test('native pivot caches, charts, print and XLSX use the same engine', async () => {
  const s=new SpreadsheetSession();s.GetRange('A1:B4').Values=[['Region','Revenue'],['North',10],['North',20],['South',5]];
  s.Invoke('worksheets.add',{name:'Report'});
  s.Invoke('pivots.add',{name:'Sales',source:{sheet:'Sheet1',address:'A1:B4'},destination:{sheet:'Report',address:'A1'},options:{rows:['Region'],values:[{column:'Revenue',aggregate:'sum',name:null}],columns:[],filters:[]}});
  assert.equal(s.GetRange('B2','Report').Value,30);
  const chart=s.CallWorksheet('Sheet1','AddChart',['A1:B4',{title:'Revenue'}]);
  assert.match(s.ChartSvg('Sheet1',chart.id),/<svg/);assert.match(s.PrintHtml('Sheet1'),/Revenue/);
  const xlsx=s.ExportXlsx(), t=new SpreadsheetSession();await t.ImportXlsx(xlsx);
  assert.equal(t.Invoke('pivots.list').length,1);assert.equal(t.GetRange('B2','Report').Value,30);
  await s.DisposeAsync();await t.DisposeAsync();
});
test('in-flight XLSX import cannot overwrite later edits', async () => {
  const s=new SpreadsheetSession();s.GetRange('A1').Value=1;const bytes=s.ExportXlsx();const loading=s.ImportXlsx(bytes);s.GetRange('A1').Value=2;
  await assert.rejects(loading,/changed during import/);assert.equal(s.GetRange('A1').Value,2);await s.DisposeAsync();
});
test('disposing borrowed sessions does not destroy external native workbooks', async () => {
  const book=new Workbook(),s=new SpreadsheetSession(book);let disposed=0;const old=book.Dispose.bind(book);book.Dispose=()=>{disposed++;old();};
  const view={Workbook:new Workbook()};s.Attach(view);await Promise.all([s.DisposeAsync(),s.DisposeAsync()]);assert.equal(disposed,0);
  book.GetRange('A1').Value=2;assert.equal(book.GetRange('A1').Value,2);book.Dispose();
});
test('public native operations preserve prototypes and reject private paths', async () => {
  const s=new SpreadsheetSession();s.SetRangeProperty('A1','Format.NumberFormat','0.00');assert.equal(s.GetRange('A1').Format.NumberFormat,'0.00');
  assert.throws(()=>s.CallRange('A1','__proto__.bad'),/Invalid/);assert.throws(()=>s.CallWorksheet(null,'_setRaw',[]),/Invalid/);
  s.GetRange('A1').Value='hello';assert.equal(s.Find('hell')[0].sheet,'Sheet1');await s.DisposeAsync();
});
test('literal application values survive the actual reusable interop resolver', async () => {
  const bridge=new Session({SpreadsheetSession});const s=await bridge.construct('SpreadsheetSession',[{$literal:null}]);
  await bridge.call(s,'Invoke',['range.values.set',{$literal:{address:'A1',values:[['$fn is plain data']]}}]);
  const blob=await bridge.transfer('call',s,'Save',[],'json',1e6);assert.ok((await blob.text()).includes('$fn is plain data'));await bridge.dispose();
});
test('worker entry uses the same allowlisted workbook protocol', () => {
  const messages=[],handler=createWorkerHandler(x=>messages.push(x));handler.handle({id:1,method:'range.values.set',address:'A1',values:[[5]]});
  handler.handle({id:2,method:'range.values.get',address:'A1'});assert.deepEqual(messages.find(x=>x.id===2).result,[[5]]);handler.Dispose();
});
