import test from 'node:test';
import assert from 'node:assert/strict';
import {Workbook,isError,shiftFormula,a1ToR1C1,MAX_DEFINED_NAMES} from '../src/core.js';
import {exportXlsx,importXlsx} from '../src/io.js';
import {readZip,writeZip} from '../src/zip.js';
import {diffDocuments,applyDocument,validateDocument} from '../src/collaboration/document.js';
import {createHostBridge} from '../src/host.js';
const norm=v=>Array.isArray(v)?v.map(norm):isError(v)?v.code:v;
function fixture(){
 const book=new Workbook(),north=book.ActiveWorksheet;north.Name='North';const south=book.Worksheets.Add('South'),summary=book.Worksheets.Add('Summary');
 north.GetRange('A1:A3').Values=[[10],[20],[30]];south.GetRange('A1:A3').Values=[[100],[200],[300]];
 book.Names.Add('Rate',2,{comment:'Global'});north.Names.Add('Rate',3);south.Names.Add('Rate',5);
 north.Names.Add('Data','=A1:A3');south.Names.Add('Data','=A1:A3');north.Names.Add('LocalFn','=LAMBDA(x,x*Rate)');
 const ev=(formula,sheet=summary,row=0,col=0)=>norm(book.Calculation.Evaluate(formula,{sheet,row,col}));return{book,north,south,summary,ev};
}
for(const [formula,scope,expected]of[
 ['=Rate','Summary',2],['=rate','North',3],['=RATE','South',5],['=North!Rate','Summary',3],['=South!Rate','North',5],
 ["='North'!Rate+'South'!Rate",'Summary',8],['=SUM(North!Data)','Summary',60],['=SUM(South!Data)','North',600],
 ['=SUM(Data)','North',60],['=SUM(Data)','Summary','#NAME?'],['=North!Missing','Summary','#NAME?'],['=Missing!Rate','North','#REF!'],
 ['=LocalFn(7)','North',21],['=North!LocalFn(7)','South',21],['=SUM(INDEX(North!Data,2))','South',20],
 ['=SUM(LET(area,North!Data,area))','South',60],['=LET(Rate,9,Rate+North!Rate)','Summary',12],['=AREAS(North!Data)','South',1],
 ['=ISREF(North!Data)','South',true],['=ISREF(North!Rate)','South',false],['=ROW(North!Data)','South',[[1],[2],[3]]],
 ['=SUM((North!Data,South!Data))','Summary',660],['=SUM(North!Data North!A2:A3)','Summary',50],['=SHEET(North!Data)','Summary',1],
])test(`scoped names ${scope} ${formula}`,()=>{const f=fixture();assert.deepEqual(f.ev(formula,f.book.Worksheets.Get(scope)),expected);});
for(const [formula,expected]of[
 ['=LET(data,North!A1:A3,ISREF(data))',true],['=LAMBDA(data,ISREF(data))(North!A1:A3)',true],['=LAMBDA(data,ISREF(data))({1;2})',false],
 ['=LET(data,(North!A1:A2,North!A3),AREAS(data))',2],['=LET(data,North!A:A,SUM(data))',60],
 ['=LET(data,North!A1:A3,SUM(INDEX(data,2):INDEX(data,3)))',50],['=SUM(LAMBDA(data,data)(North!A1:A3))',60],
 ['=ISREF(LAMBDA(data,data)(North!A1:A3))',true],['=ROW(LAMBDA(data,INDEX(data,2))(North!A1:A3))',2],
 ['=LET(data,North!A1:A3,fn,LAMBDA(SUM(data)),fn())',60],['=LET(value,1,fn,LAMBDA(value),value,2,fn())',1],
 ['=LET(data,North!A1,fn,LAMBDA(data),data,South!A1,SUM(fn()))',10],
 ['=LET(data,North!A1:A3,copy,data,SUM(copy))',60],['=LET(data,North!A1:A3,SUM(data+1))',63],
 ['=LET(data,North!A1:A3,ISREF(data+0))',false],['=LAMBDA(x,LAMBDA(y,SUM(x,y)))(North!A1:A3)(South!A1:A3)',660],
 ['=LET(data,North!A1:A3,LAMBDA(data,SUM(data))(South!A1:A3))',600],['=LET(data,North!A1:A3,LAMBDA(other,ISOMITTED(other))())','#VALUE!'],
 ['=LAMBDA(x,y,IF(ISOMITTED(y),SUM(x),0))(North!A1:A3,)',60],['=SUM(LET(data,North!A1:A3,data))',60],
 ['=ISREF(LET(data,North!A1:A3,data))',true],['=LET(data,North!A1:A3,@data)',10],
 ['=LAMBDA(x,x)(#N/A)','#N/A'],['=LET(x,42,x)',42],['=LET(x,{1,2},x)',[[1,2]]],
 ['=LET(x,"text",x)','text'],['=LET(x,North!A1:A3,fn,LAMBDA(v,SUM(v)),fn(x))',60],
])test(`reference bindings ${formula}`,()=>{assert.deepEqual(fixture().ev(formula),expected);});
for(const name of ['', 'A1','XFE1','R1C1','R','C','TRUE','FALSE','two words','with!sheet','1start','x'.repeat(256),'_xlfn.Private'])test('name validation '+name,()=>{const{book}=fixture();const before=book.ToJSON();assert.throws(()=>book.Names.Create(name,1));assert.deepEqual(book.ToJSON(),before);});
test('same name in separate scopes preserves metadata, snapshots and legacy iteration',()=>{
 const{book,north}=fixture();assert.equal(book.Names.Get('rate'),2);assert.equal(north.Names.Get('Rate'),3);assert.equal(book.Names.Count,1);
 const item=book.Names.GetDefinition('RATE');item.name='mutated';assert.equal(book.Names.GetDefinition('Rate').name,'Rate');
 assert.deepEqual([...book.Names],[['RATE',2]]);book.Names.Add('Fault',{error:'#N/A'});const fault=book.Names.Get('Fault');fault.error='#NUM!';assert.equal(book.Names.Get('Fault').error,'#N/A');
 assert.throws(()=>north.Names.Create('RATE',7));assert.throws(()=>north.Names.Update('Missing',1));
});
test('scoped upsert removal and no-op updates each maintain history and formulas',()=>{
 const{book,north,summary}=fixture();summary.GetCell('A1').Formula='=North!Rate';book.ClearHistory();const rev=book.Revision;
 north.Names.Add('Rate',3);assert.equal(book.Revision,rev);assert.equal(book.CanUndo,false);
 north.Names.Update('Rate',7);assert.equal(summary.GetCell('A1').Value,7);book.Undo();assert.equal(summary.GetCell('A1').Value,3);book.Redo();assert.equal(summary.GetCell('A1').Value,7);
 north.Names.Remove('Rate');assert.equal(summary.GetCell('A1').Value.code,'#NAME?');assert.equal(north.GetCell('B1').Worksheet.Workbook.Calculation.Evaluate('=Rate',{sheet:north}),2);book.Undo();assert.equal(summary.GetCell('A1').Value,7);
});
test('global definition context does not leak caller LET variables or worksheet names',()=>{
 const{book,north,south,ev}=fixture();book.Names.Add('GlobalFormula','=Rate+A1');assert.equal(ev('=GlobalFormula',south),13);
 assert.equal(ev('=LET(Rate,99,GlobalFormula)',south),13);book.Names.Add('GlobalFn','=LAMBDA(x,x+Rate)');assert.equal(ev('=LET(Rate,99,GlobalFn(1))',south),4);
});
test('reference-binding probing evaluates volatile/custom value calls once',()=>{
 const{book,ev}=fixture();let calls=0;book.Calculation.RegisterFunction('TICK',()=>++calls);
 assert.equal(ev('=SUM(LET(value,TICK(),value))'),1);assert.equal(calls,1);calls=0;
 assert.equal(ev('=SUM(LAMBDA(TICK())())'),1);assert.equal(calls,1);calls=0;
 assert.equal(ev('=LET(value,TICK(),value+value)'),2);assert.equal(calls,1);
});
test('LET/LAMBDA reference bindings retain hidden rows and nested totals',()=>{
 const{book,north,ev}=fixture();north.HideRows(1);
 assert.equal(ev('=LET(data,North!A1:A3,SUBTOTAL(109,data))'),40);
 assert.equal(ev('=LAMBDA(data,AGGREGATE(9,5,data))(North!A1:A3)'),40);
 assert.equal(ev('=LET(data,North!A1:A3+0,SUM(data))'),60);
 north.GetCell('A4').Formula='=SUBTOTAL(9,A1:A3)';assert.equal(ev('=LAMBDA(data,SUBTOTAL(9,data))(North!A1:A4)'),60);
});
test('bound full-column and spill references invalidate dependents on later population',()=>{
 const{book,north,summary}=fixture();summary.GetCell('C1').Formula='=LET(data,North!B:B,SUM(data))';assert.equal(summary.GetCell('C1').Value,0);north.GetCell('B200').Value=9;assert.equal(summary.GetCell('C1').Value,9);
 north.GetCell('C1').Formula='=SEQUENCE(3)';summary.GetCell('D1').Formula='=LET(data,North!C1#,SUM(data))';assert.equal(summary.GetCell('D1').Value,6);north.GetCell('C1').Formula='=SEQUENCE(4)';assert.equal(summary.GetCell('D1').Value,10);
});
test('scoped reference variables and recursive named lambdas terminate with bounded errors',()=>{
 const{north,ev}=fixture();north.Names.Add('Loop','=Loop');assert.equal(ev('=SUM(North!Loop)'),'#CIRC!');
 north.Names.Add('Factorial','=LAMBDA(n,IF(n<=1,1,n*Factorial(n-1)))');assert.equal(ev('=North!Factorial(6)'),720);
 north.Names.Add('Recurse','=LAMBDA(x,Recurse(x))');assert.equal(ev('=North!Recurse(1)'),'#NUM!');
});
test('name renaming repairs free references but preserves shadowed bindings, strings and structured selectors',()=>{
 const{book,north,south,summary,ev}=fixture();summary.GetCell('A1').Formula='=Rate+North!Rate';summary.GetCell('A2').Formula='=LET(Rate,8,Rate)+Rate';summary.GetCell('A3').Formula='=LAMBDA(Rate,Rate+1)(8)+Rate';summary.GetCell('A4').Formula='="Rate"';
 south.GetCell('B1').Formula='=Rate';book.Names.Add('Alias','=Rate',{contextSheetId:summary.Id});book.ClearHistory();
 book.Names.Rename('Rate','Tax');assert.equal(summary.GetCell('A1').Formula,'=Tax+North!Rate');assert.equal(summary.GetCell('A2').Formula,'=LET(Rate,8,Rate)+Tax');assert.equal(summary.GetCell('A3').Formula,'=LAMBDA(Rate,Rate+1)(8)+Tax');assert.equal(summary.GetCell('A4').Value,'Rate');assert.equal(south.GetCell('B1').Formula,'=Rate');assert.equal(book.Names.Get('Alias'),'=Tax');assert.equal(ev('=Alias'),2);
 book.Undo();assert.equal(summary.GetCell('A1').Formula,'=Rate+North!Rate');assert.equal(book.Names.Has('Rate'),true);assert.equal(book.Names.Has('Tax'),false);
});
test('local rename repairs qualified external uses, local aliases, validation and conditional format formulas',()=>{
 const{book,north,summary}=fixture();north.Names.Add('Alias','=Rate+1');summary.GetCell('B1').Formula='=North!Rate';north.AddValidation('B1',{type:'custom',formula:'=B1<Rate'});north.AddConditionalFormat('B1',{type:'formula',formula:'=B1>Rate'});book.ClearHistory();
 north.Names.Rename('Rate','LocalTax');assert.equal(norm(summary.GetCell('B1').Value),3);assert.equal(north.Names.Get('Alias'),'=LocalTax+1');assert.equal(north._meta.validations[0].formula,'=B1<LocalTax');assert.equal(north._meta.conditionalFormats[0].formula,'=B1>LocalTax');book.Undo();assert.equal(north.Names.Get('Alias'),'=Rate+1');
});
test('unsafe rename capture and function collisions reject atomically',()=>{
 const{book,summary}=fixture();summary.GetCell('A1').Formula='=LET(Tax,9,Rate)';book.ClearHistory();const before=book.ToJSON();assert.throws(()=>book.Names.Rename('Rate','Tax'),/capture/);assert.deepEqual(book.ToJSON(),before);assert.equal(book.CanUndo,false);
 book.Names.Add('MyFn','=LAMBDA(x,x+1)');summary.GetCell('B1').Formula='=MyFn(2)';assert.throws(()=>book.Names.Rename('MyFn','SUM'),/built-in/);
});
test('qualified names with cell-shaped and escaped worksheet names survive copy and rename',()=>{
 const{book,north,summary,ev}=fixture();north.Name='A1';assert.equal(ev('=A1!Rate'),3);assert.equal(shiftFormula('=A1!Rate+B2',1,1),'=A1!Rate+C3');assert.equal(a1ToR1C1('=A1!Rate+B2',{row:0,column:0}),'=A1!Rate+R[1]C[1]');summary.GetCell('B1').Formula='=A1!Rate';north.Name="Team's North";assert.equal(ev("='Team''s North'!Rate"),3);assert.equal(summary.GetCell('B1').Value,3);
});
test('sheet deletion tombstones qualified local references and undo restores the original names',()=>{
 const{book,north,summary}=fixture();summary.GetCell('A1').Formula='=North!Rate';summary.GetCell('B1').Formula='=North!LocalFn(2)';book.ClearHistory();book.Worksheets.Remove(north);
 assert.equal(norm(summary.GetCell('A1').Value),'#REF!');assert.equal(norm(summary.GetCell('B1').Value),'#REF!');assert.throws(()=>north.Names.Add('Other',3),/attached/);
 const replacement=book.Worksheets.Add('North');replacement.Names.Add('Rate',99);assert.equal(norm(summary.GetCell('A1').Value),'#REF!');book.Undo();book.Undo();book.Undo();assert.equal(book.Worksheets.Get('North'),north);assert.equal(summary.GetCell('A1').Value,3);
});
test('row and column edits rewrite local named ranges only in the correct definition context',()=>{
 const{book,north,south,summary,ev}=fixture();book.Names.Add('GlobalData','=A1:A3');north.InsertRows(0);assert.equal(ev('=SUM(North!Data)'),60);assert.equal(north.Names.Get('Data'),'=A2:A4');assert.equal(south.Names.Get('Data'),'=A1:A3');assert.equal(book.Names.Get('GlobalData'),'=A2:A4');book.Undo();assert.equal(north.Names.Get('Data'),'=A1:A3');
 south.InsertColumns(0);assert.equal(south.Names.Get('Data'),'=B1:B3');assert.equal(book.Names.Get('GlobalData'),'=A1:A3');
});
test('defined names resolve in range APIs but disjoint/value names cannot masquerade as rectangles',()=>{
 const{book,north,south}=fixture();assert.equal(north.GetRange('Data').FullAddress,"'North'!A1:A3");assert.equal(book.GetRange('South!Data').FullAddress,"'South'!A1:A3");assert.throws(()=>north.GetRange('Rate'));
 north.Names.Add('Parts','=(A1,A3)');assert.throws(()=>north.GetRange('Parts'),/rectangular/);
});
test('explicit relative origins preserve mixed references and caller coordinates through JSON',()=>{
 const{book,north,south,ev}=fixture();north.Names.Add('Relative','=A1+$A$1',{baseAddress:'B2'});assert.equal(ev('=North!Relative',south,2,1),30);const loaded=Workbook.FromJSON(book.ToJSON());assert.equal(loaded.Calculation.Evaluate('=North!Relative',{sheet:loaded.Worksheets.Get('South'),row:2,col:1}),30);assert.throws(()=>exportXlsx(book),/relative-origin/);
});
test('JSON restores scoped metadata and shared synchronization keeps worksheet collection ownership',()=>{
 const{book,north}=fixture();north.Names.Update('Data','=A1:A3',{comment:'Local reference',hidden:true});const saved=book.ToJSON(),loaded=Workbook.FromJSON(saved);assert.deepEqual(loaded.ToJSON(),saved);
 const oldNorth=north;loaded.Worksheets.Get('North').Names.Update('Rate',8);const delta=diffDocuments(saved,loaded.ToJSON());assert.equal(delta.kind,'replace');applyDocument(book,loaded.ToJSON());assert.equal(book.Worksheets.Get('North'),oldNorth);assert.equal(north.Names.Get('Rate'),8);north.Names.Update('Rate',9);assert.equal(north.Names.Get('Rate'),9);assert.equal(validateDocument(book.ToJSON()).sheets[0].names.find(([n])=>n==='RATE')[1],9);
});
for(const mutate of [d=>d.names.push(['RATE',99]),d=>d.sheets[0].names.push(['rate',9]),d=>d.nameMetadata.push({name:'Orphan'}),d=>d.sheets[0].nameMetadata[0].hidden='yes',d=>d.sheets[0].nameMetadata[0].comment='x'.repeat(256),d=>d.names.push(['R1C1',1]),d=>d.names='bad'])test('stored name input validation '+String(mutate),()=>{const d=fixture().book.ToJSON();mutate(d);assert.throws(()=>Workbook.FromJSON(d));});
test('old v1 documents without name metadata remain readable',()=>{const d=fixture().book.ToJSON();delete d.nameMetadata;for(const s of d.sheets){delete s.nameMetadata;delete s.names;}const b=Workbook.FromJSON(d);assert.equal(b.Names.Get('Rate'),2);});
test('defined-name cardinality and option limits are enforced before mutation',()=>{
 const{book,north}=fixture();for(const v of [{hidden:1},{comment:null},{contextSheetId:'missing'},{baseAddress:'0'},{unknown:true}])assert.throws(()=>north.Names.Create('Bad',1,v));
 // Fill the internal fixture map directly to avoid 10,000 recalculation transactions.
 for(let i=book.GetDefinedNames().length;i<MAX_DEFINED_NAMES;i++)book._names.set('Limit_'+i,i);assert.throws(()=>north.Names.Create('Overflow',1),/limit/);
});
test('XLSX retains same-spelling names on multiple worksheets, metadata and definition context',async()=>{
 const{book,north,south,summary}=fixture();north.Names.Update('Rate',3,{comment:'Local & <tax>',hidden:true});book.Names.Add('GlobalData','=A1:A3');book.Names.Add('Alias','=Rate');
 summary.GetCell('A1').Formula='=North!Rate+South!Rate+Rate';summary.GetCell('B1').Formula='=SUM(GlobalData)';summary.GetCell('C1').Formula='=Alias';
 const bytes=exportXlsx(book),parts=await readZip(bytes),xml=new TextDecoder().decode(parts.get('xl/workbook.xml'));assert.match(xml,/localSheetId="0"/);assert.match(xml,/hidden="1"/);assert.match(xml,/comment="Local &amp; &lt;tax&gt;"/);
 const {workbook:loaded}=await importXlsx(bytes);assert.equal(loaded.Worksheets.Get('Summary').GetCell('A1').Value,10);assert.equal(loaded.Worksheets.Get('Summary').GetCell('B1').Value,60);assert.equal(loaded.Worksheets.Get('Summary').GetCell('C1').Value,3);assert.equal(loaded.Worksheets.Get('North').Names.GetDefinition('Rate').hidden,true);assert.equal(loaded.Worksheets.Get('North').Names.GetDefinition('Rate').comment,'Local & <tax>');
 book.Worksheets.Move(north,2);const moved=(await importXlsx(exportXlsx(book))).workbook;assert.equal(moved.Worksheets.Get('North').Names.Has('Rate'),true);assert.equal(moved.Worksheets.Get('South').Names.Has('Rate'),true);
});
test('XLSX boolean, error, text names and missing/invalid local scope are not silently flattened',async()=>{
 const{book}=fixture();book.Names.Add('Flag',true);book.Names.Add('Fault',{error:'#N/A'});book.Names.Add('Text','literal');const bytes=exportXlsx(book),loaded=(await importXlsx(bytes)).workbook;
 for(const[name,expected]of[['Flag',true],['Fault','#N/A'],['Text','literal']])assert.equal(norm(loaded.Calculation.Evaluate('='+name,{sheet:loaded.ActiveWorksheet})),expected);
 const parts=await readZip(bytes);let xml=new TextDecoder().decode(parts.get('xl/workbook.xml'));xml=xml.replace('localSheetId="0"','localSheetId="999"');parts.set('xl/workbook.xml',xml);const result=await importXlsx(writeZip(Object.fromEntries(parts)));assert(result.warnings.some(w=>w.includes('missing sheet')));assert.equal(result.workbook.Names.Get('Rate'),'=2');
});
test('host commands expose scope-aware create list rename remove with real model updates',()=>{
 const{book,north,summary}=fixture(),bridge=createHostBridge({Workbook:book});const rpc=r=>JSON.parse(bridge.dispatch(JSON.stringify(r)));
 assert.equal(rpc({method:'names.define',sheet:north.Id,name:'LocalHost',value:7}).result,true);assert.equal(rpc({method:'names.get',sheet:north.Id,name:'localhost'}).result.value,7);
 assert.equal(rpc({method:'names.rename',sheet:north.Id,name:'LocalHost',newName:'RenamedHost'}).result,true);assert.equal(rpc({method:'names.list',all:true}).result.some(n=>n.name==='RenamedHost'&&n.sheetId===north.Id),true);
 assert.equal(rpc({method:'names.remove',sheet:north.Id,name:'RenamedHost'}).result,true);assert(rpc({method:'names.define',sheet:'missing',name:'Bad',value:7}).error);bridge.Dispose();
});

test('reference probing memoizes value-returning named LET/LAMBDA definitions',()=>{
 const{book,north,ev}=fixture();let calls=0;book.Calculation.RegisterFunction('TICK',()=>++calls);
 north.Names.Add('Computed','=LET(v,TICK(),v)');calls=0;assert.equal(ev('=SUM(North!Computed)'),1);assert.equal(calls,1);
 book.Names.Add('Callable','=LAMBDA(TICK())()');calls=0;assert.equal(ev('=SUM(Callable)'),1);assert.equal(calls,1);
});
test('renaming at the shared name limit does not transiently allocate an extra definition',()=>{
 const{book,north}=fixture();for(let i=book.GetDefinedNames().length;i<MAX_DEFINED_NAMES;i++)book._names.set('Limit_'+i,i);
 north.Names.Rename('Rate','NewRate');assert.equal(north.Names.Get('NewRate'),3);book.Undo();assert.equal(north.Names.Get('Rate'),3);
});
test('local rename refuses to capture an existing use of a workbook name',()=>{
 const{book,north}=fixture();book.Names.Add('GlobalTax',9);north.GetCell('B1').Formula='=GlobalTax';book.ClearHistory();const saved=book.ToJSON();assert.throws(()=>north.Names.Rename('Rate','GlobalTax'),/shadow/);assert.deepEqual(book.ToJSON(),saved);assert.equal(book.CanUndo,false);
});
test('global range APIs bypass active worksheet shadowing and reject value-only definitions',()=>{
 const{book,north,south}=fixture();book.Names.Add('Data','=North!A1:A3');book.ActiveWorksheet=south;
 assert.equal(book.Names.GetRange('Data').Worksheet,north);assert.equal(south.Names.GetRange('Data').Worksheet,south);
 assert.deepEqual(book.Names.Evaluate('Data'),[[10],[20],[30]]);assert.throws(()=>north.Names.GetRange('Rate'),/range/);
});
test('definition context tombstones can be edited and renamed without rebinding missing sheets',()=>{
 const{book,north,summary}=fixture();book.Names.Add('DeadRef','=A1');book.Worksheets.Remove(north);assert.equal(norm(book.Names.Evaluate('DeadRef')),'#REF!');
 book.Names.Update('DeadRef','=A1',{comment:'missing context'});book.Names.Rename('DeadRef','MissingRef');assert.equal(norm(book.Names.Evaluate('MissingRef')),'#REF!');
 const loaded=Workbook.FromJSON(book.ToJSON());assert.equal(norm(loaded.Names.Evaluate('MissingRef')),'#REF!');book.Names.Update('MissingRef','=A1',{contextSheetId:summary.Id});summary.GetCell('A1').Value=8;assert.equal(book.Names.Evaluate('MissingRef'),8);
});
test('table and defined-name collisions reject in both creation orders and on import',()=>{
 const{book,north}=fixture();north.GetRange('J1:J2').Values=[['Header'],[1]];north.Names.Add('MyTable',1);assert.throws(()=>north.AddTable('J1:J2','MyTable'));
 north.Names.Remove('MyTable');north.AddTable('J1:J2','MyTable');assert.throws(()=>book.Names.Add('MyTable',1));assert.throws(()=>north.Names.Rename('Rate','MyTable'));
 const saved=book.ToJSON();saved.names.push(['MyTable',1]);assert.throws(()=>Workbook.FromJSON(saved),/conflict/);
});
test('ambiguous relative-origin structural edits and duplicate sheet identities are rejected',()=>{
 const{book,north}=fixture();north.Names.Add('Relative','=A1',{baseAddress:'B2'});book.ClearHistory();const before=book.ToJSON();assert.throws(()=>north.InsertRows(0),/relative-origin/);assert.deepEqual(book.ToJSON(),before);assert.equal(book.CanUndo,false);
 const doc=book.ToJSON();doc.sheets[1].id=doc.sheets[0].id;assert.throws(()=>Workbook.FromJSON(doc),/identity/);
});
test('hand-authored OOXML name scopes follow original sheet indices after unsupported sheets are skipped',async()=>{
 const book=new Workbook();book.ActiveWorksheet.Name='Skipped';book.Worksheets.Add('Visible').GetCell('A1').Value=7;
 const parts=await readZip(exportXlsx(book)),decode=b=>new TextDecoder().decode(b),encode=s=>new TextEncoder().encode(s);
 parts.set('xl/worksheets/sheet1.xml',encode('<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>'));
 const xml=decode(parts.get('xl/workbook.xml')).replace('</workbook>','<definedNames><definedName name="Rate">2</definedName><definedName name="Rate" localSheetId="1" hidden="1" comment="Local">3</definedName><definedName name="Data" localSheetId="1">Visible!$A$1</definedName><definedName name="Lost" localSheetId="0">4</definedName></definedNames></workbook>');parts.set('xl/workbook.xml',encode(xml));
 const imported=await importXlsx(writeZip(Object.fromEntries(parts)));assert.equal(imported.workbook.Worksheets.Count,1);assert.equal(imported.workbook.Names.Evaluate('Rate'),2);const s=imported.workbook.ActiveWorksheet;assert.equal(s.Names.Evaluate('Rate'),3);assert.equal(s.Names.GetRange('Data').Value,7);assert.equal(s.Names.GetDefinition('Rate').hidden,true);assert.equal(imported.workbook.Names.Has('Lost'),false);assert.ok(imported.warnings.some(w=>w.includes('Lost')));
 parts.set('xl/workbook.xml',encode(xml.replace('localSheetId="1"','localSheetId="-1"')));await assert.rejects(()=>importXlsx(writeZip(Object.fromEntries(parts))),/localSheetId/);
});
test('rename never rebinds a tombstoned definition context to the active worksheet',()=>{
 const b=new Workbook(),deleted=b.ActiveWorksheet;deleted.Name='Gone';b.Names.Create('Old','=2');b.Names.Create('Formula','=Old+A1');const other=b.Worksheets.Add('Other');other.Names.Create('Old','=7');other.GetCell('A1').Value=100;b.Worksheets.Remove(deleted);b.ActiveWorksheet=other;
 b.Names.Rename('Old','Renamed');assert.match(b.Names.Get('Formula'),/Renamed/);assert.equal(norm(b.Names.Evaluate('Formula')),'#REF!');b.Undo();assert.equal(b.Names.Has('Old'),true);
});
test('XLSX export rejects deleted definition contexts instead of silently changing semantics',()=>{
 const b=new Workbook(),source=b.ActiveWorksheet;b.Names.Create('FromSource','=SUM(A1:A3)');b.Worksheets.Add('Other');b.Worksheets.Remove(source);assert.throws(()=>exportXlsx(b),/deleted name context/);const restored=Workbook.FromJSON(b.ToJSON());assert.equal(norm(restored.Names.Evaluate('FromSource')),'#REF!');
});

for(const name of ['RC','R1C','RC1'])test('R1C1-shaped name rejection '+name,()=>{const b=new Workbook();assert.throws(()=>b.Names.Create(name,1),/defined name/);assert.equal(b.Names.Count,0);});
test('ISOMITTED does not confuse a qualified name with an omitted lexical parameter',()=>{const{north,ev}=fixture();north.Names.Create('y',2);assert.equal(ev('=LAMBDA(x,y,ISOMITTED(North!y))(1,)'),false);assert.equal(ev('=LAMBDA(x,y,ISOMITTED(y))(1,)'),true);});
