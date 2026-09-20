import test from 'node:test';
import assert from 'node:assert/strict';
import {Workbook} from '../src/core.js';
import {createExcelApi} from '../src/office.js';
const fixture=()=>{const book=new Workbook();book.ActiveWorksheet.Name='North';book.ActiveWorksheet.GetRange('A1:A3').Values=[[1],[2],[3]];book.Worksheets.Add('South').GetRange('A1:A3').Values=[[10],[20],[30]];return{book,api:createExcelApi(book)};};
test('Office names are deferred, scoped, loadable and use the actual grid ranges',async()=>{
 const{book,api}=fixture();await api.run(async ctx=>{
  const north=ctx.workbook.worksheets.getItem('North'),south=ctx.workbook.worksheets.getItem('South');
  const global=ctx.workbook.names.add('Data','=North!$A$1:$A$3','Global range'),local=south.names.add('Data',south.getRange('A1:A3'),'Local range');
  assert.equal(book.Names.Has('Data'),false);assert.throws(()=>global.name,e=>e.code==='PropertyNotLoaded');
  global.load('name,scope,comment,value,type,formula,visible');local.load('name,scope,comment,value,type,formula');local.getRange().values=[[11],[22],[33]];
  const r=local.getRange().load('values');await ctx.sync();assert.equal(global.name,'Data');assert.equal(global.scope,'Workbook');assert.equal(global.type,'Range');assert.equal(global.visible,true);assert.equal(local.scope,'Worksheet');assert.equal(local.formula,"='South'!$A$1:$A$3");assert.equal(global.value,"'North'!A1:A3");assert.deepEqual(r.values,[[11],[22],[33]]);
 });assert.equal(book.Worksheets.Get('South').Names.Get('Data'),"='South'!$A$1:$A$3");
});
test('Office formula, comment and visibility writes retain scope and one batch undo',async()=>{
 const{book,api}=fixture();book.Worksheets.Get('South').Names.Add('Rate',3);book.Worksheets.Get('North').GetCell('B1').Formula='=South!Rate';book.ClearHistory();
 await api.run(async ctx=>{const item=ctx.workbook.worksheets.getItem('South').names.getItem('Rate');item.formula='=7';item.comment='Updated';item.visible=false;item.load('*');await ctx.sync();assert.equal(item.value,7);assert.equal(item.type,'Integer');assert.equal(item.visible,false);assert.equal(item.comment,'Updated');});
 assert.equal(book.ActiveWorksheet.GetCell('B1').Value,7);book.Undo();assert.equal(book.ActiveWorksheet.GetCell('B1').Value,3);assert.equal(book.Worksheets.Get('South').Names.GetDefinition('Rate').hidden,false);
});
test('Office failed name batches retain earlier workbook state and do not publish failed loads',async()=>{
 const{book,api}=fixture();const ctx=api.createRequestContext(),item=ctx.workbook.names.add('Valid','=42');item.load('name');ctx.workbook.names.add('A1','=99');const before=book.ToJSON();
 await assert.rejects(()=>ctx.sync(),e=>e.code==='InvalidArgument');assert.deepEqual(book.ToJSON(),before);assert.throws(()=>item.name,e=>e.code==='PropertyNotLoaded');ctx.dispose();
});
test('Office named scalar, array, logical and error values use spreadsheet output types',async()=>{
 const{api}=fixture();await api.run(async ctx=>{
  const cases=[['Number','=1.5','Double',1.5],['Text','="name"','String','name'],['Flag','=TRUE','Boolean',true],['Fault','=#N/A','Error','#N/A'],['ArrayName','={1,#N/A}','Array',[[1,'#N/A']]],['Callable','=LAMBDA(x,x)','Error','#CALC!']];
  for(const[name,formula,type,value]of cases){const item=ctx.workbook.names.add(name,formula).load('type,value');await ctx.sync();assert.equal(item.type,type);assert.deepEqual(item.value,value);}
 });
});
test('Office collections load all items including hidden names without mixing scopes',async()=>{
 const{api}=fixture();await api.run(async ctx=>{ctx.workbook.names.add('Global','=2');const sheet=ctx.workbook.worksheets.getItem('North');sheet.names.add('Local','=3').visible=false;const names=sheet.names.load('items/name,items/scope,items/visible');assert.throws(()=>names.items,e=>e.code==='PropertyNotLoaded');await ctx.sync();assert.equal(names.items.length,1);assert.equal(names.items[0].name,'Local');assert.equal(names.items[0].visible,false);assert.equal(names.items[0].scope,'Worksheet');const returned=names.items;returned.pop();assert.equal(names.items.length,1);});
});
test('Office named ranges observe writes queued earlier in the same batch',async()=>{
 const{api}=fixture();await api.run(async ctx=>{const s=ctx.workbook.worksheets.getItem('North');s.getRange('A1').values=[[10]];const n=s.names.add('Total','=SUM(A1:A3)').load('value');await ctx.sync();assert.equal(n.value,15);});
});
test('Office names reject foreign ranges, invalid scope, unknown properties and deleted objects',async()=>{
 const{api}=fixture(),a=api.createRequestContext(),b=api.createRequestContext();assert.throws(()=>a.workbook.names.add('Foreign',b.workbook.worksheets.getActiveWorksheet().getRange('A1')),e=>e.code==='InvalidObjectPath');
 assert.throws(()=>a.workbook.names.add('Number',42),e=>e.code==='NotSupported');assert.throws(()=>a.workbook.names.load('items/nonexistent'),e=>e.code==='NotSupported');
 a.workbook.names.getItem('Missing').load('name');await assert.rejects(()=>a.sync(),e=>e.code==='ItemNotFound');
 const global=a.workbook.names.add('Global','=1');global.worksheet.load('name');await assert.rejects(()=>a.sync(),e=>e.code==='InvalidOperation');
 const n=a.workbook.names.add('Value','=2');n.getRange().load('values');await assert.rejects(()=>a.sync(),e=>e.code==='InvalidOperation');
 const local=a.workbook.worksheets.getItem('South').names.add('Local','=7');await a.sync();a.workbook.worksheets.getItem('South').delete();local.load('name');await assert.rejects(()=>a.sync(),e=>e.code==='ItemNotFound');a.dispose();b.dispose();
});
test('Office names retain literal formula-looking text and support deferred deletion',async()=>{
 const{book,api}=fixture();await api.run(async ctx=>{const n=ctx.workbook.names.add('Literal','="=1+2"').load('value');await ctx.sync();assert.equal(n.value,'=1+2');n.delete();assert.equal(book.Names.Has('Literal'),true);await ctx.sync();assert.equal(book.Names.Has('Literal'),false);n.load('name');await assert.rejects(()=>ctx.sync(),e=>e.code==='ItemNotFound');});
});
test('Office range additions preserve cell-shaped worksheet names',async()=>{
 const{book,api}=fixture();book.ActiveWorksheet.Name='A1';await api.run(async ctx=>{ctx.workbook.names.add('CellLike',ctx.workbook.worksheets.getItem('A1').getRange('A1:A3'));});assert.equal(book.Names.Get('CellLike'),"='A1'!$A$1:$A$3");assert.equal(book.Names.GetRange('CellLike').Worksheet.Name,'A1');
});
