import test from 'node:test';
import assert from 'node:assert/strict';
import {Workbook} from '../src/core.js';
import {createHostBridge} from '../src/host.js';
import {createExcelApi} from '../src/office.js';
import {createShowcase} from '../demo/fixtures.js';

test('host and worker-compatible RPC expose only allowlisted editing operations', () => {
  const control = {Workbook:new Workbook(),Selection:'A1',Select(value){this.Selection=value;}}, bridge=createHostBridge(control);
  const call = (method,args={})=>JSON.parse(bridge.dispatch(JSON.stringify({id:1,method,...args})));
  call('range.values.set',{address:'A1:B2',values:[[1,2],[3,4]]});
  assert.equal(call('range.pasteSpecial',{address:'D1',sourceAddress:'A1:B2',options:{mode:'values',transpose:true}}).result,'D1:E2');
  assert.deepEqual(call('range.values.get',{address:'D1:E2'}).result,[[1,3],[2,4]]);
  assert.equal(call('range.fillSeries',{address:'F1:F3',options:{start:2,step:2,type:'growth'}}).result,3);
  assert.deepEqual(call('range.values.get',{address:'F1:F3'}).result,[[2],[4],[8]]);
  assert.deepEqual(call('range.specialCells',{address:'A1:F3',type:'constants'}).result,['A1','B1','D1','E1','F1','A2','B2','D2','E2','F2','F3']);
  assert.ok(call('range.pasteSpecial',{address:'A1'}).error);assert.ok(call('range.fillSeries',{address:'A1',options:{type:'eval'}}).error);
  assert.ok(call('capabilities').result.methods.includes('range.specialCells')); bridge.Dispose();
});
test('paste-special RPC supports cross-sheet source selection and one undo', () => {
  const b=new Workbook(),s=b.ActiveWorksheet,other=b.Worksheets.Add('Other');s.GetRange('A1:B1').Values=[[3,4]];b.ClearHistory();
  const bridge=createHostBridge({Workbook:b});const reply=JSON.parse(bridge.dispatch(JSON.stringify({method:'range.pasteSpecial',sheet:'Other',address:'B2',sourceSheet:s.Name,sourceAddress:'A1:B1',options:{mode:'values'}})));
  assert.equal(reply.result,'B2:C2');assert.deepEqual(other.GetRange('B2:C2').Values,[[3,4]]);b.Undo();assert.deepEqual(other.GetRange('B2:C2').Values,[[null,null]]);bridge.Dispose();
});
test('Office range copy supports transpose, blank skipping, and formats with deferred batch semantics', async () => {
  const book=new Workbook(),s=book.ActiveWorksheet;s.GetRange('A1:B2').Values=[[1,null],[3,4]];s.GetCell('A1').Style={fill:'#123456'};
  s.GetRange('D1:E2').Values=[[9,9],[9,9]];book.ClearHistory();
  await createExcelApi(book).run(async ctx=>{const sheet=ctx.workbook.worksheets.getActiveWorksheet();sheet.getRange('D1').copyFrom(sheet.getRange('A1:B2'),'Values',true,true);sheet.getRange('G1').copyFrom(sheet.getRange('A1'),'Formats');assert.equal(s.GetCell('D1').Value,9);await ctx.sync();});
  assert.deepEqual(s.GetRange('D1:E2').Values,[[1,3],[9,4]]);assert.equal(s.GetCell('G1').Style.fill,'#123456');assert.equal(book._history.length,1);
  book.Undo();assert.deepEqual(s.GetRange('D1:E2').Values,[[9,9],[9,9]]);assert.equal(s.GetCell('G1').Style.fill,undefined);
});
test('Office copy recomputes values after preceding writes in the same batch', async () => {
  const book=new Workbook();const s=book.ActiveWorksheet;s.GetCell('B1').Formula='=A1+1';assert.equal(s.GetCell('B1').Value,1);
  await createExcelApi(book).run(ctx=>{const sheet=ctx.workbook.worksheets.getActiveWorksheet();sheet.getRange('A1').values=[[5]];sheet.getRange('D1').copyFrom(sheet.getRange('B1'),'Values',false,true);});
  assert.equal(s.GetCell('D1').Value,6);
});
test('showcase retains five default worksheets and evaluates the new formulas', () => {
  const b=createShowcase(),s=b.Worksheets.Get('Formula lab');assert.equal(b.Worksheets.Count,5);
  assert.deepEqual(s.GetRange('B37:B41').Values,[[144],[14],[36],[4],[35]]);
});

test('new calculations and materialized editing results survive GridWeb JSON and XLSX round trips',async()=>{
  const {exportXlsx,importXlsx}=await import('../src/io.js');
  const book=new Workbook(),s=book.ActiveWorksheet;
  s.GetCell('A1').Formula='=MAKEARRAY(2,3,LAMBDA(row,col,row*col))';
  s.GetCell('E1').Formula='=LAMBDA(x,y,IF(ISOMITTED(y),x^2,x+y))(7,)';
  s.GetCell('E2').Formula='=AGGREGATE(9,6,{1,#N/A,3})';
  s.GetCell('E3').Formula='=DSUM({"Region","Sales";"North",2;"South",3},"Sales",{"Region";"North"})';
  s.GetRange('G1:G4').FillSeries({start:1,step:3});s.GetRange('I1').PasteSpecial(s.GetRange('G1:G4'),{mode:'values',transpose:true});
  for(const restored of [Workbook.FromJSON(book.ToJSON()),(await importXlsx(exportXlsx(book))).workbook]){
    const target=restored.ActiveWorksheet;assert.deepEqual(target.GetRange('A1:C2').Values,[[1,2,3],[2,4,6]]);
    assert.deepEqual(target.GetRange('E1:E3').Values,[[49],[4],[2]]);assert.deepEqual(target.GetRange('I1:L1').Values,[[1,4,7,10]]);
  }
});
test('database criteria cross products have an explicit evaluation budget',()=>{
  const book=new Workbook(),database=[['H'],...Array.from({length:2100},()=>[0])],criteria=[['H'],...Array.from({length:2100},()=>[1])];
  assert.throws(()=>book.Calculation.Functions.get('DSUM')(database,1,criteria),e=>e.code==='#NUM!'&&/evaluation limit/.test(e.detail));
});
test('paste flags reject non-boolean options before mutation',()=>{
  const b=new Workbook(),s=b.ActiveWorksheet;s.GetCell('A1').Value=5;
  for(const options of [{transpose:'true'},{skipBlanks:1}])assert.throws(()=>s.GetRange('C1').PasteSpecial(s.GetRange('A1'),options),/boolean/);
  assert.equal(s.GetCell('C1').Value,null);
});
test('validation clipping enforces the rule limit even when the source has no validation',()=>{
  const b=new Workbook(),s=b.ActiveWorksheet;
  s._setMeta('validations',Array.from({length:1000},()=>({type:'number',min:0,max:10,range:{r1:0,c1:0,r2:2,c2:2}})));
  b.ClearHistory();const before=b.ToJSON();
  assert.throws(()=>s.GetRange('B2').PasteSpecial(s.GetRange('F1'),{mode:'validation'}),/rule limit/);
  assert.deepEqual(b.ToJSON(),before);assert.equal(b.CanUndo,false);
});
