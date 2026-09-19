import test from 'node:test';import assert from 'node:assert/strict';
import {Workbook} from '../src/core.js';import {createExcelApi,RequestContext,OfficeApiError} from '../src/office.js';
test('writes are deferred; loaded values require explicit sync',async()=>{
 const b=new Workbook(),Excel=createExcelApi(b);await Excel.run(async ctx=>{const r=ctx.workbook.worksheets.getActiveWorksheet().getRange('A1:B2');r.values=[[1,2],[3,4]];r.load('values,address,rowCount,columnCount');assert.equal(b.ActiveWorksheet.GetCell('A1').Value,null);assert.throws(()=>r.values,e=>e.code==='PropertyNotLoaded');await ctx.sync();assert.deepEqual(r.values,[[1,2],[3,4]]);assert.equal(r.address,"'Sheet1'!A1:B2");assert.equal(r.rowCount,2);});assert.equal(b._history.length,1);
});
test('automatic final sync commits queued edits and discards on callback error',async()=>{
 const b=new Workbook(),Excel=createExcelApi(b);await Excel.run(ctx=>{ctx.workbook.worksheets.getActiveWorksheet().getRange('A1').values=[[42]];});assert.equal(b.ActiveWorksheet.GetCell('A1').Value,42);
 await assert.rejects(Excel.run(ctx=>{ctx.workbook.worksheets.getActiveWorksheet().getRange('A1').values=[[99]];throw Error('abort');}));assert.equal(b.ActiveWorksheet.GetCell('A1').Value,42);
});
test('null skips cells and empty string clears content in matrix assignments',async()=>{
 const b=new Workbook(),Excel=createExcelApi(b);b.ActiveWorksheet.GetRange('A1:B1').Values=[[1,2]];await Excel.run(async ctx=>{const r=ctx.workbook.worksheets.getActiveWorksheet().getRange('A1:B1');r.values=[[null,'']];r.numberFormat=[[null,'0.00']];r.load('values,numberFormat');await ctx.sync();assert.deepEqual(r.values,[[1,'']]);assert.deepEqual(r.numberFormat,[['General','0.00']]);});
});
test('a failed batch rolls back prior edits and does not publish failed loads',async()=>{
 const b=new Workbook(),ctx=new RequestContext(b),r=ctx.workbook.worksheets.getActiveWorksheet().getRange('A1');r.values=[[5]];r.load('values');r.values=[[1,2]];await assert.rejects(ctx.sync(),e=>e.code==='InvalidArgument');assert.equal(b.ActiveWorksheet.GetCell('A1').Value,null);assert.throws(()=>r.values,e=>e.code==='PropertyNotLoaded');assert.equal(b.CanUndo,false);
});
test('load commands preserve ordering and refresh formula results inside a batch',async()=>{
 const b=new Workbook(),ctx=new RequestContext(b),s=ctx.workbook.worksheets.getActiveWorksheet(),early=s.getRange('A1'),late=s.getRange('A1');early.values=[[1]];early.load('values');late.values=[[2]];late.load('values');s.getRange('B1').formulas=[['=A1*3']];const result=s.getRange('B1');result.load('values');await ctx.sync();assert.deepEqual(early.values,[[1]]);assert.deepEqual(late.values,[[2]]);assert.deepEqual(result.values,[[6]]);
});
test('nested format assignment/load and mixed formatting produce truthful snapshots',async()=>{
 const b=new Workbook(),ctx=new RequestContext(b),r=ctx.workbook.worksheets.getActiveWorksheet().getRange('A1:B1');r.set({values:[[1,2]],format:{font:{bold:true,color:'#123456'},fill:{color:'#abcdef'}}});r.load('format/font/bold,format/fill/color');await ctx.sync();assert.equal(r.format.font.bold,true);assert.equal(r.format.fill.color,'#abcdef');r.getCell(0,1).format.font.color='#654321';r.format.font.load('color');await ctx.sync();assert.equal(r.format.font.color,null);
});
test('add, rename, collection load, activate and delete resolve same model objects',async()=>{
 const b=new Workbook(),ctx=new RequestContext(b),s=ctx.workbook.worksheets.add('Added');s.getRange('A1').values=[[8]];s.name='Renamed';s.activate();ctx.workbook.worksheets.load('items/name');await ctx.sync();assert.equal(b.ActiveWorksheet.Name,'Renamed');assert.equal(ctx.workbook.worksheets.items[1].name,'Renamed');s.getRange('B1').values=[[9]];s.delete();await ctx.sync();assert.equal(b.Worksheets.Count,1);
});
test('copy, offset, resize and clear use the shared engine and translate formulas',async()=>{
 const b=new Workbook(),ctx=new RequestContext(b),s=ctx.workbook.worksheets.getActiveWorksheet();s.getRange('A1:B1').formulas=[[2,'=A1*2']];s.getRange('A2:B2').copyFrom(s.getRange('A1:B1'));const r=s.getRange('A1').getResizedRange(1,1);r.load('values');await ctx.sync();assert.deepEqual(r.values,[[2,4],[2,4]]);s.getRange('A2').getOffsetRange(0,1).clear('Contents');await ctx.sync();assert.equal(b.ActiveWorksheet.GetCell('B2').Value,null);
});
test('contexts reject unsupported operations, foreign proxies and reuse after disposal',async()=>{
 const b=new Workbook(),ctx=new RequestContext(b),other=new RequestContext(b),r=ctx.workbook.worksheets.getActiveWorksheet().getRange('A1');assert.throws(()=>r.load('notReal'),e=>e.code==='NotSupported');assert.throws(()=>ctx.load(other.workbook.worksheets.getActiveWorksheet(),'name'),e=>e.code==='InvalidObjectPath');assert.throws(()=>r.copyFrom(r,'Unsupported'),e=>e.code==='NotSupported');ctx.dispose();assert.throws(()=>r.values=[[5]],e=>e.code==='InvalidObjectPath');
});
test('read snapshots do not mutate the workbook by reference',async()=>{
 const b=new Workbook(),ctx=new RequestContext(b),r=ctx.workbook.worksheets.getActiveWorksheet().getRange('A1');r.values=[[1]];r.load('values');await ctx.sync();r.values[0][0]=100;const j=r.toJSON();j.values[0][0]=200;assert.deepEqual(r.values,[[1]]);assert.equal(b.ActiveWorksheet.GetCell('A1').Value,1);
});
test('missing sheets reject with ItemNotFound only on synchronization',async()=>{
 const ctx=new RequestContext(new Workbook());const missing=ctx.workbook.worksheets.getItem('Missing');missing.load('name');await assert.rejects(ctx.sync(),e=>e instanceof OfficeApiError&&e.code==='ItemNotFound');
});
test('values-only copying observes preceding writes rather than cached formula results',async()=>{
 const b=new Workbook();b.ActiveWorksheet.GetCell('A1').Value=1;b.ActiveWorksheet.GetCell('B1').Formula='=A1*10';assert.equal(b.ActiveWorksheet.GetCell('B1').Value,10);
 const ctx=new RequestContext(b),s=ctx.workbook.worksheets.getActiveWorksheet();s.getRange('A1').values=[[2]];s.getRange('C1').copyFrom(s.getRange('B1'),'Values');await ctx.sync();assert.equal(b.ActiveWorksheet.GetCell('C1').Value,20);
});
