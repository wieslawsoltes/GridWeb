import test from 'node:test';
import assert from 'node:assert/strict';
import {Workbook,MAX_DEFINED_NAMES} from '../src/core.js';
import {createHostBridge} from '../src/host.js';
for(let bits=1;bits<16;bits++)test('Create from Selection boundary combination '+bits,()=>{
 const b=new Workbook(),s=b.ActiveWorksheet;s.GetRange('A1:F5').Values=Array.from({length:5},(_,r)=>Array.from({length:6},(_,c)=>`Label_${r}_${c}`));
 const options={topRow:!!(bits&1),bottomRow:!!(bits&2),leftColumn:!!(bits&4),rightColumn:!!(bits&8)};
 const rows=[...Array(5).keys()].filter(i=>!(options.topRow&&i===0||options.bottomRow&&i===4)),cols=[...Array(6).keys()].filter(i=>!(options.leftColumn&&i===0||options.rightColumn&&i===5));
 const expected=new Map();for(const c of cols){if(options.topRow)expected.set(`Label_0_${c}`,{r1:rows[0],r2:rows.at(-1),c1:c,c2:c});if(options.bottomRow)expected.set(`Label_4_${c}`,{r1:rows[0],r2:rows.at(-1),c1:c,c2:c});}
 for(const r of rows){if(options.leftColumn)expected.set(`Label_${r}_0`,{r1:r,r2:r,c1:cols[0],c2:cols.at(-1)});if(options.rightColumn)expected.set(`Label_${r}_5`,{r1:r,r2:r,c1:cols[0],c2:cols.at(-1)});}
 b.ClearHistory();const result=s.Names.CreateFromSelection(s.GetRange('A1:F5'),options);assert.equal(result.length,expected.size);for(const[name,bounds]of expected)assert.deepEqual(s.Names.GetRange(name).Bounds,bounds);assert.equal(s.Names.Count,expected.size);b.Undo();assert.equal(s.Names.Count,0);b.Redo();assert.equal(s.Names.Count,expected.size);
});
for(const[label,name]of[['Net sales','Net_sales'],['2026','_2026'],['A1','_A1'],['R1C1','_R1C1'],['R','_R'],['C','_C'],['TRUE','_TRUE'],['<img src=x>','_img_src_x_'],['\tTrim me\t','Trim_me'],['x'.repeat(260),'x'.repeat(255)]])test('Create from Selection label normalization '+name.slice(0,30),()=>{
 const b=new Workbook(),s=b.ActiveWorksheet;s.GetRange('A1:A2').Values=[[label],[9]];const generated=b.Names.CreateFromSelection(s.GetRange('A1:A2'));assert.equal(generated[0].name,name);assert.equal(b.Names.GetRange(name).Value,9);
});
test('selection-generated names exclude every selected header edge and retain source sheet identity',()=>{
 const b=new Workbook(),s=b.ActiveWorksheet;s.Name="Team's A1";s.GetRange('A1:C3').Values=[['Region','Net sales','Costs'],['North',10,2],['South',20,3]];
 const dest=b.Worksheets.Add('Consumer'),result=dest.Names.CreateFromSelection(s.GetRange('A1:C3'),{topRow:true,leftColumn:true});assert.equal(result.length,4);assert.equal(dest.Names.Get('Net_sales'),"='Team''s A1'!$B$2:$B$3");assert.equal(dest.Names.GetRange('North').Worksheet,s);assert.deepEqual(dest.Names.GetRange('North').Values,[[10,2]]);
 dest.GetCell('A1').Formula='=SUM(Net_sales)';assert.equal(dest.GetCell('A1').Value,30);s.GetCell('B2').Value=50;assert.equal(dest.GetCell('A1').Value,70);
});
test('blank labels are skipped and duplicate normalized labels reject the entire operation',()=>{
 const b=new Workbook(),s=b.ActiveWorksheet;s.GetRange('A1:C2').Values=[['First',null,'Last'],[1,2,3]];assert.equal(s.Names.CreateFromSelection(s.GetRange('A1:C2')).length,2);
 s.GetRange('D1:F2').Values=[['New','Dup name','Dup_name'],[4,5,6]];b.ClearHistory();const before=b.ToJSON();assert.throws(()=>s.Names.CreateFromSelection(s.GetRange('D1:F2')),/Duplicate generated/);assert.deepEqual(b.ToJSON(),before);assert.equal(b.CanUndo,false);
});
test('existing names are preserved unless overwrite is explicitly enabled and undo restores metadata',()=>{
 const b=new Workbook(),s=b.ActiveWorksheet;s.GetRange('A1:B2').Values=[['First','Existing'],[1,2]];s.Names.Create('Existing','=99',{comment:'Keep',hidden:true});b.ClearHistory();const before=b.ToJSON();assert.throws(()=>s.Names.CreateFromSelection(s.GetRange('A1:B2')),/already exists/);assert.deepEqual(b.ToJSON(),before);
 s.Names.CreateFromSelection(s.GetRange('A1:B2'),{topRow:true,overwrite:true});assert.equal(s.Names.GetRange('Existing').Value,2);assert.equal(s.Names.GetDefinition('Existing').comment,'Keep');assert.equal(s.Names.GetDefinition('Existing').hidden,true);b.Undo();assert.deepEqual(b.ToJSON(),before);
});
test('generated name options, label errors, foreign ranges and empty data reject before writes',()=>{
 const b=new Workbook(),s=b.ActiveWorksheet;s.GetRange('A1:A2').Values=[['Name'],[1]];for(const opts of [{},{topRow:1},{topRow:true,bad:true},{topRow:true,bottomRow:true}])assert.throws(()=>s.Names.CreateFromSelection(s.GetRange('A1:A2'),opts));assert.throws(()=>s.Names.CreateFromSelection(new Workbook().ActiveWorksheet.GetRange('A1:A2')));
 s.GetCell('A1').Formula='=1/0';const before=b.ToJSON();assert.throws(()=>s.Names.CreateFromSelection(s.GetRange('A1:A2')),/label/);assert.deepEqual(b.ToJSON(),before);assert.equal(s.Names.Count,0);
});
test('generated name preflight enforces the shared count and table collisions',()=>{
 const b=new Workbook(),s=b.ActiveWorksheet;s.GetRange('A1:A2').Values=[['LimitName'],[1]];for(let i=0;i<MAX_DEFINED_NAMES;i++)b._names.set('Stored_'+i,i);b.ClearHistory();assert.throws(()=>s.Names.CreateFromSelection(s.GetRange('A1:A2')),/limit/);assert.equal(b.CanUndo,false);b._names.clear();s.AddTable('A1:A2','LimitName');assert.throws(()=>s.Names.CreateFromSelection(s.GetRange('A1:A2')),/table/);
});
test('worker-compatible host creates names from source and destination scopes supplied as JSON',()=>{
 const b=new Workbook(),s=b.ActiveWorksheet,scope=b.Worksheets.Add('Scope');s.GetRange('A1:B2').Values=[['First','Second'],[1,2]];const h=createHostBridge({Workbook:b});const reply=JSON.parse(h.dispatch(JSON.stringify({method:'names.createFromSelection',sheet:s.Id,scopeSheet:scope.Id,address:'A1:B2',options:{topRow:true}})));assert.equal(reply.result.length,2);assert.equal(scope.Names.GetRange('Second').Value,2);assert.equal(b.Names.Count,0);h.Dispose();
});
