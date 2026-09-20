import test from 'node:test';
import assert from 'node:assert/strict';
import {Workbook, FormulaError, parseFormula, shiftFormula, renameSheetReferences, a1ToR1C1, r1c1ToA1, quoteSheet} from '../src/core.js';
import {exportXlsx,importXlsx} from '../src/io.js';
import {createHostBridge} from '../src/host.js';
import {createExcelApi} from '../src/office.js';
function fixture(){const book=new Workbook(),sheet=book.ActiveWorksheet;sheet.Name='Jan';sheet.GetRange('A1:C4').Values=[[1,10,100],[2,20,200],[3,30,300],[4,40,400]];const feb=book.Worksheets.Add('Feb'),mar=book.Worksheets.Add('Mar'),summary=book.Worksheets.Add('Summary');feb.GetRange('A1:A3').Values=[[5],[6],[7]];mar.GetRange('A1:A3').Values=[[8],[9],[10]];book.ActiveWorksheet=summary;return{book,sheet,feb,mar,summary,ev:f=>book.Calculation.Evaluate(f,{sheet,row:1,col:1})};}
function equal(actual,expected){if(typeof expected==='string'&&expected.startsWith('#'))assert.equal(actual?.code,expected);else if(typeof expected==='number')assert.ok(Math.abs(actual-expected)<=1e-10*Math.max(1,Math.abs(expected)),`${actual} != ${expected}`);else assert.deepEqual(actual,expected);}
const cases=[
 ['=SUM((A1:A3,B1:B3))',66],['=SUM((A1:A3,A2:A4))',15],['=SUM(A1:C3 B2:C4)',550],['=SUM(A:A 2:2)',2],['=SUM(A1:A2 B1:B2)','#NULL!'],
 ['=SUM((A1:A3,B1:B3) A2:C2)',22],['=SUM(A1:INDEX(A:A,3))',6],['=SUM(INDEX(A1:C4,0,2))',100],['=SUM(INDEX(A1:C4,2,0))',222],
 ['=SUM(INDEX((A1:A3,B1:B3),0,1,2))',60],['=INDEX((A1:A3,B1:B3),2,1,2)',20],['=INDEX(A1:C1,2)',10],['=INDEX({1,2,3},2)',2],['=INDEX({1;2;3},2)',2],
 ['=INDEX(A1:B3,0,2)',[[10],[20],[30]]],['=INDEX(A1:B3,2)',[[2,20]]],['=INDEX((A1:A3,B1:B3),,,2)',[[10],[20],[30]]],
 ['=INDEX(A1:A3,-1)','#VALUE!'],['=INDEX(A1:A3,4)','#REF!'],['=INDEX((A1,A2),1,1,0)','#REF!'],['=INDEX((A1,A2),1,1,3)','#REF!'],
 ['=INDEX((Jan!A1,Feb!A1),1,1,2)','#VALUE!'],['=INDEX({1,2},1,1,1)','#VALUE!'],['=INDEX(A1,1,1,1,1)','#VALUE!'],
 ['=AREAS((B2:D4,E5,F6:I9))',3],['=AREAS(B2:D4 B2)',1],['=AREAS(1)','#VALUE!'],['=AREAS(A1,A2)','#VALUE!'],['=AREAS()','#VALUE!'],
 ['=ISREF(INDEX(A1:A3,2))',true],['=ISREF(OFFSET(A1,1,1))',true],['=ISREF(A1+1)',false],['=ISREF(#REF!)',false],['=ISREF(Unknown!A1)',false],
 ['=ROW(INDEX(A1:A3,2))',2],['=COLUMN(INDEX(A1:C3,1,3))',3],['=ROWS(INDEX(A1:C4,0,2))',4],['=COLUMNS(INDEX(A1:C4,2,0))',3],
 ['=COUNTBLANK(INDEX(D1:D4,0,1))',4],['=SUM(OFFSET(INDEX(A1:A3,2),0,1,2,1))',50],['=@INDEX(A1:A3,0,1)',2],
 ['=SUM(A1:OFFSET(A1,2,1))',66],['=SUM((A1,B1),2)',13],['=SUM( ( A1 , B1 ) )',11],['=SUM ( A1 , B1 )',11],
 ['=SUM(A1:INDEX(A:A,3))^2',36],['=-SUM(A1:A3)^2',36],['=SUM(A1:A3)%',.06],['=SUM(A1:A2)+SUM(B1:B2)*2',63],
 ['=SUM(#REF!:A1)','#REF!'],['=SUM(A1:5)','#VALUE!'],['=SUM(A1:Feb!B2)','#REF!'],['=SUM(INDEX(A1,1):Feb!B2)','#VALUE!'],['=SUM(A1,1/0)','#DIV/0!'],
 ['=SUM(Jan:Mar!A1:A3)',51],['=SUM(Mar:Jan!A1:A3)',51],['=SUM(\'Jan:Mar\'!A1:A3)',51],['=SUM(\'Jan\':\'Mar\'!A1:A3)',51],
 ['=SUM(Jan:Jan!A1:A3)',6],['=SUM(Jan:Mar!A:A)',55],['=SUM(Jan:Unknown!A1)','#REF!'],
 ['=Jan:Mar!A1','#VALUE!'],['=SUM(Jan:Mar!A1+1)','#VALUE!'],['=SUBTOTAL(9,Jan:Mar!A1)','#VALUE!'],['=AGGREGATE(9,0,Jan:Mar!A1)','#VALUE!'],
 ['=ROWS(Jan:Mar!A1)','#VALUE!'],['=INDEX(Jan:Mar!A1,1)','#VALUE!'],['=MEDIAN(Jan:Mar!A1)','#VALUE!'],['=SUM(Jan:Mar!A1 A1)','#VALUE!'],
 ['=HSTACK(Jan:Mar!A1:A2)',[[1,5,8],[2,6,9]]],['=VSTACK(Jan:Mar!A1:A2)',[[1],[2],[5],[6],[8],[9]]],
 ['=SHEET()',1],['=SHEET("Feb")',2],['=SHEET(Mar!A1)',3],['=SHEET("Missing")','#N/A'],['=SHEET(#REF!)','#REF!'],
 ['=SHEETS()',4],['=SHEETS(Jan:Mar!A1)',3],['=SHEETS(A1)',1],['=SHEETS(3)','#REF!'],['=SHEETS(#REF!)','#REF!'],['=SHEETS(A1,A2)','#VALUE!']
];
for(const [formula,expected] of cases)test('reference semantics '+formula,()=>equal(fixture().ev(formula),expected));
for(const [name,expected] of [['SUM',51],['AVERAGE',51/9],['COUNT',9],['COUNTA',9],['MIN',1],['MAX',10],['PRODUCT',907200],['VAR.S',10],['VAR.P',80/9],['STDEV.S',Math.sqrt(10)],['STDEV.P',Math.sqrt(80/9)]])test('3-D numeric aggregation '+name,()=>equal(fixture().ev(`=${name}(Jan:Mar!A1:A3)`),expected));
for(const [name,expected] of [['AVERAGEA',1],['MAXA',2],['MINA',0],['VARA',1],['VARPA',2/3],['STDEVA',1],['STDEVPA',Math.sqrt(2/3)]])test('Inclusive statistics '+name,()=>{const {book,sheet,feb,mar,ev}=fixture();sheet.GetCell('D1').Value=2;feb.GetCell('D1').Value=true;mar.GetCell('D1').Value='text';equal(ev(`=${name}(Jan:Mar!D1)`),expected);equal(ev(`=${name}({2,TRUE,"text"})`),expected);assert.ok(book.Calculation.FunctionNames.includes(name));});
for(const [formula,expected] of [['=AVERAGEA("2",TRUE)',1.5],['=AVERAGEA("bad")','#VALUE!'],['=AVERAGEA({2,TRUE,""})',1],['=AVERAGEA(D1:D4)','#DIV/0!'],['=MAXA(D1:D4)',0],['=VARA(1)','#DIV/0!'],['=VARPA(1)',0],['=MAXA(1,#N/A)','#N/A']])test('Inclusive coercion '+formula,()=>equal(fixture().ev(formula),expected));
test('Dynamic range dependencies rewire without recalculating unrelated cells',()=>{
 const {book,sheet}=fixture();sheet.GetCell('D1').Value=2;sheet.GetCell('E1').Formula='=SUM(A1:INDEX(A:A,D1))';sheet.GetCell('E2').Formula='=42';book.ClearHistory();
 assert.equal(sheet.GetCell('E1').Value,3);sheet.GetCell('D1').Value=3;assert.equal(sheet.GetCell('E1').Value,6);
 const count=book.Calculation.EvaluationCount;sheet.GetCell('C4').Value=777;assert.equal(book.Calculation.EvaluationCount,count);
 sheet.GetCell('A3').Value=30;assert.equal(sheet.GetCell('E1').Value,33);book.Undo();assert.equal(sheet.GetCell('E1').Value,6);
});
test('Initially empty 3-D full-column and named-union dependencies notice new values',()=>{
 const {book,sheet,feb,summary}=fixture();book.DefineName('Span','=Jan:Mar!D:D');book.DefineName('Parts','=(Jan!D1:D4,Feb!D1:D4)');summary.GetCell('A1').Formula='=SUM(Span)';summary.GetCell('B1').Formula='=SUM(Parts)';
 assert.equal(summary.GetCell('A1').Value,0);feb.GetCell('D3').Value=17;assert.equal(summary.GetCell('A1').Value,17);assert.equal(summary.GetCell('B1').Value,17);
 sheet.GetCell('D1').Value=11;assert.equal(summary.GetCell('A1').Value,28);assert.equal(summary.GetCell('B1').Value,28);
});
test('Spill references preserve identity through INDEX, intersection, and 3-D reads',()=>{
 const {book,sheet,feb,summary,ev}=fixture();sheet.GetCell('D1').Formula='=SEQUENCE(3)';summary.GetCell('A1').Formula='=SUM(Jan:Mar!D:D)';assert.equal(summary.GetCell('A1').Value,6);
 equal(ev('=SUM(D1#)'),6);equal(ev('=SUM(D1# 2:3)'),5);equal(ev('=ISREF(D1#)'),true);equal(ev('=SUM(INDEX(D1#,0,1))'),6);
 sheet.GetCell('D1').Formula='=SEQUENCE(4)';assert.equal(summary.GetCell('A1').Value,10);equal(ev('=SUM(INDEX(D1,1)#)'),10);
});
test('Reference INDEX retains text origin and supports formula inspection',()=>{
 const {sheet,ev}=fixture();sheet.GetCell('D1').Value='3';sheet.GetCell('D2').Formula='=A1+1';
 equal(ev('=SUM(INDEX(D1:D2,1))'),0);equal(ev('=AVERAGEA(INDEX(D1:D2,1))'),0);equal(ev('=ISFORMULA(INDEX(D1:D2,2))'),true);equal(ev('=FORMULATEXT(INDEX(D1:D2,2))'),'=A1+1');
});
test('Union aggregation respects hidden/filtered rows and nested totals',()=>{
 const {sheet,ev}=fixture();sheet.HideRows(1,1);equal(ev('=SUBTOTAL(109,(A1:A3,B1:B3))'),44);equal(ev('=AGGREGATE(9,1,(A1:A3,B1:B3))'),44);
 sheet.GetCell('D1').Formula='=SUBTOTAL(9,A1:A3)';equal(ev('=SUBTOTAL(9,(A1:A3,D1))'),6);equal(ev('=AGGREGATE(9,4,(A1:A3,D1))'),12);
});
test('Table column spans, item selection, row references and intersections share the resolver',()=>{
 const {sheet,ev}=fixture();sheet.GetRange('F1:H3').Values=[['Left','Middle','Right'],[1,2,3],[4,5,6]];sheet.AddTable('F1:H3','Data');
 for(const [f,v]of [['SUM(Data[[Left]:[Right]])',21],['SUM(Data[[#Data],[Left]:[Middle]])',12],['SUM(Data[Middle] 3:3)',5],['SUM(Data[@[Left]])',1],['ISREF(Data[Left])',true],['SUM(INDEX(Data[Left],2))',4],['SHEET(Data)',1]])equal(ev('='+f),v);
 equal(ev('=SUM(Data[[#Totals],[Left]])'),'#REF!');equal(ev('=SUM(Data[Unknown])'),'#REF!');
});
test('Worksheet insertion and moving interior sheets recalculate 3-D consumers atomically',()=>{
 const {book,feb,summary}=fixture();summary.GetCell('A1').Formula='=SUM(Jan:Mar!A1:A3)';book.ClearHistory();
 const added=book.Worksheets.Add('Extra',1);added.GetCell('A1').Value=100;assert.equal(summary.GetCell('A1').Value,151);
 book.ClearHistory();const active=book.ActiveWorksheet,rev=book.Revision;book.Worksheets.Move(feb,4);assert.equal(summary.GetCell('A1').Value,133);assert.equal(book.Revision,rev+1);assert.equal(book.ActiveWorksheet,active);
 assert.ok(book.Undo());assert.equal(summary.GetCell('A1').Value,151);assert.equal(book.Worksheets.Get(2),feb);assert.ok(book.Redo());assert.equal(summary.GetCell('A1').Value,133);
});
for(const moved of ['Jan','Mar'])test('Crossing a 3-D '+moved+' endpoint follows documented sheet semantics',()=>{
 const {book,summary}=fixture();summary.GetCell('A1').Formula='=SUM(Jan:Mar!A1:A3)';const original=summary.GetCell('A1').Formula;book.ClearHistory();book.Worksheets.Move(moved,moved==='Jan'?3:0);
 assert.equal(summary.GetCell('A1').Value,moved==='Jan'?45:24);book.Undo();assert.equal(summary.GetCell('A1').Formula,original);assert.equal(summary.GetCell('A1').Value,51);
});
test('Removing endpoints shrinks to a single sheet then permanently invalidates references',()=>{
 const {book,sheet,summary}=fixture();summary.GetCell('A1').Formula='=SUM(Jan:Mar!A1:A3)';summary.GetCell('B1').Formula='=Jan!A1';book.DefineName('Span','=Jan:Mar!A1:A3');book.ClearHistory();
 book.Worksheets.Remove(sheet);assert.equal(summary.GetCell('A1').Value,45);assert.equal(summary.GetCell('B1').Value.code,'#REF!');assert.equal(book.Names.Get('Span'),"='Feb:Mar'!A1:A3");
 book.Worksheets.Remove('Feb');assert.equal(summary.GetCell('A1').Value,27);book.Worksheets.Remove('Mar');assert.equal(summary.GetCell('A1').Value.code,'#REF!');
 book.Worksheets.Add('Jan').GetCell('A1').Value=999;assert.equal(summary.GetCell('B1').Value.code,'#REF!');
 for(let i=0;i<5;i++)book.Undo();assert.equal(book.Worksheets.Get('Jan'),sheet);assert.equal(summary.GetCell('A1').Value,51);assert.equal(summary.GetCell('B1').Value,1);
});
test('Detached worksheet removal and invalid moves never remove a different sheet',()=>{
 const {book,sheet}=fixture();book.Worksheets.Remove(sheet);const before=book.ToJSON();assert.throws(()=>book.Worksheets.Remove(sheet));assert.deepEqual(book.ToJSON(),before);
 for(const n of [-1,4,1.5,NaN,'1'])assert.throws(()=>book.Worksheets.Move('Feb',n));assert.throws(()=>book.Worksheets.Move(sheet,0));assert.throws(()=>book.Worksheets.Add('Bad',-1));
 const rev=book.Revision;book.Worksheets.Move('Feb',0);assert.equal(book.Revision,rev);
});
test('Rename repairs both 3-D endpoints, names, validations and conditional formulas',()=>{
 const {book,sheet,summary}=fixture();summary.GetCell('A1').Formula='=SUM(Jan:Mar!A1:A3)';book.DefineName('Span','=Jan:Mar!A:A');summary.AddValidation('B1',{type:'custom',formula:'=Jan!A1>0'});summary.AddConditionalFormat('B1',{type:'formula',formula:'=SUM(Jan:Mar!A1)>0'});book.ClearHistory();sheet.Name="Jan's Data";
 assert.equal(summary.GetCell('A1').Value,51);assert.equal(summary.GetCell('A1').Formula,"=SUM('Jan''s Data:Mar'!A1:A3)");assert.match(summary._meta.validations[0].formula,/Jan''s Data/);assert.match(summary._meta.conditionalFormats[0].formula,/Jan''s Data:Mar/);
 book.Undo();assert.equal(summary._meta.validations[0].formula,'=Jan!A1>0');assert.equal(summary.GetCell('A1').Value,51);
});
test('3-D structural edits fail explicitly before mutating sheets or history',()=>{
 const {book,sheet,feb,summary}=fixture();summary.GetCell('A1').Formula='=SUM(Jan:Mar!A1:A3)';book.ClearHistory();const before=book.ToJSON(),rev=book.Revision;
 assert.throws(()=>feb.InsertRows(1),/3-D/);assert.throws(()=>sheet.DeleteColumns(0),/3-D/);assert.deepEqual(book.ToJSON(),before);assert.equal(book.Revision,rev);assert.equal(book.CanUndo,false);
 feb.InsertRows(10);assert.equal(summary.GetCell('A1').Value,51);
});
test('Normal union/dynamic references survive structural edits without touching identifiers',()=>{
 const {book,sheet,summary}=fixture();summary.GetCell('A1').Formula='=SUM((Jan!A1:A3,Jan!B1:B3))';summary.GetCell('A2').Formula='=SUM(Jan!A1:INDEX(Jan!A:A,3))';sheet.InsertRows(0);
 assert.equal(summary.GetCell('A1').Value,66);assert.equal(summary.GetCell('A1').Formula,"=SUM(('Jan'!A2:A4,'Jan'!B2:B4))");assert.equal(summary.GetCell('A2').Formula,"=SUM('Jan'!A2:INDEX(Jan!A:A,3))");book.Undo();assert.equal(summary.GetCell('A1').Value,66);
});
for(const source of ['=SUM(Jan:Mar!A1:B3)','=SUM(\'Jan Data:Mar Data\'!$A1:B$3)','=SUM(\'Jan\':\'Mar\'!A:$B)','=SUM(Jan:Mar!$1:3)','=SUM((A1:A3,B1:B3))','="Jan:Mar!A1"&SUM(Table1[[A1]:[B2]])+A1'])test('Reference roundtrip '+source,()=>{
 const converted=a1ToR1C1(source,'D4'),back=r1c1ToA1(converted,'D4');assert.deepEqual(parseFormula(back),parseFormula(source));
 const shifted=shiftFormula(source,2,3);assert.ok(!shifted.includes('#REF!'));assert.deepEqual(parseFormula(shiftFormula(shifted,-2,-3)),parseFormula(source));
});
test('Cell-shaped sheet names and table selectors never shift as cell coordinates',()=>{
 assert.equal(shiftFormula('=SUM(A1:B2!$C3:D$4)+"A1:B2!C3"+Table1[[A1]:[B2]]',1,1),'=SUM(\'A1:B2\'!$C4:E$4)+"A1:B2!C3"+Table1[[A1]:[B2]]');
 assert.equal(renameSheetReferences('=SUM(Jan:Mar!A1:B2)+"Jan!A1"+Table1[Jan]', 'Jan','New'),'=SUM(\'New:Mar\'!A1:B2)+"Jan!A1"+Table1[Jan]');
});
test('Reference resolution budgets terminate cyclic names and oversized unions',()=>{
 const {book,ev}=fixture();book.DefineName('Loop','=Loop');equal(ev('=SUM(Loop)'),'#NUM!');
 const formula='=SUM(('+Array(258).fill('A1').join(',')+'))';assert.ok(ev(formula) instanceof FormulaError);
 equal(ev('=HSTACK(Jan:Mar!A1:XFD1048576)'),'#NUM!');
});
test('JSON and XLSX retain new formulas and sheet order, without claiming desktop equivalence',async()=>{
 const {book,summary}=fixture();summary.GetCell('A1').Formula='=SUM(Jan:Mar!A1:A3)';summary.GetCell('A2').Formula='=AREAS((Jan!A1,Jan!B1))';summary.GetCell('A3').Formula='=AVERAGEA(Jan:Mar!A1:A3)';book.Worksheets.Move('Feb',0);
 for(const restored of [Workbook.FromJSON(book.ToJSON()),(await importXlsx(exportXlsx(book))).workbook]){assert.deepEqual(restored.Worksheets.items.map(s=>s.Name),book.Worksheets.items.map(s=>s.Name));assert.deepEqual(restored.Worksheets.Get('Summary').GetRange('A1:A3').Values,summary.GetRange('A1:A3').Values);}
});
test('Host allowlist supports indexed sheet insertion and move with 3-D recalculation',()=>{
 const {book,summary}=fixture();summary.GetCell('A1').Formula='=SUM(Jan:Mar!A1:A3)';const bridge=createHostBridge({Workbook:book});const call=(method,args={})=>JSON.parse(bridge.dispatch(JSON.stringify({id:1,method,...args})));
 assert.equal(call('worksheets.add',{name:'Middle',index:1}).result.name,'Middle');call('range.values.set',{sheet:'Middle',address:'A1',values:[[100]]});assert.equal(summary.GetCell('A1').Value,151);
 assert.equal(call('worksheets.move',{sheet:'Middle',index:4}).result.position,4);assert.equal(summary.GetCell('A1').Value,51);assert.ok(call('worksheets.move',{sheet:'Middle',index:-1}).error);bridge.Dispose();
});
test('Office position writes defer, batch atomically, load and undo shared sheet order',async()=>{
 const {book,summary}=fixture();summary.GetCell('A1').Formula='=SUM(Jan:Mar!A1:A3)';book.ClearHistory();const api=createExcelApi(book);
 await api.run(async context=>{const sheet=context.workbook.worksheets.getItem('Feb');sheet.position=3;assert.equal(book.Worksheets.Get(1).Name,'Feb');sheet.load('position');await context.sync();assert.equal(sheet.position,3);});
 assert.equal(summary.GetCell('A1').Value,33);book.Undo();assert.equal(summary.GetCell('A1').Value,51);
 await assert.rejects(api.run(async context=>{context.workbook.worksheets.getItem('Jan').position=3;context.workbook.worksheets.getItem('Mar').position=-1;}));assert.equal(summary.GetCell('A1').Value,51);assert.equal(book.Worksheets.Get(0).Name,'Jan');
});

test('Invalid coordinate formulas remain spreadsheet errors during structural edits',()=>{
 const {book,sheet,ev}=fixture();assert.throws(()=>ev('=SUM(Jan:Mar!XFE1)'),e=>e.code==='#REF!');
 sheet.GetCell('D1').Formula='=SUM(Jan:Mar!XFE1)';assert.equal(sheet.GetCell('D1').Value.code,'#REF!');sheet.InsertRows(10);assert.equal(sheet.GetCell('D1').Value.code,'#REF!');
});
test('Reference operator geometry agrees with an independent dense rectangle oracle',()=>{
 const book=new Workbook(),sheet=book.ActiveWorksheet,values=Array.from({length:9},(_,r)=>Array.from({length:8},(_,c)=>(r+1)*10+c));sheet.GetRange('A1:H9').Values=values;
 let seed=0x12345678;const random=n=>{seed=(Math.imul(seed,1664525)+1013904223)>>>0;return seed%n;};
 const rect=()=>{const r=[random(9),random(9)].sort((a,b)=>a-b),c=[random(8),random(8)].sort((a,b)=>a-b);return{r1:r[0],r2:r[1],c1:c[0],c2:c[1]};};
 const addr=r=>String.fromCharCode(65+r.c1)+(r.r1+1)+':'+String.fromCharCode(65+r.c2)+(r.r2+1), sum=r=>{let result=0;for(let i=r.r1;i<=r.r2;i++)for(let j=r.c1;j<=r.c2;j++)result+=values[i][j];return result;};
 for(let i=0;i<250;i++){
  const a=rect(),b=rect(),intersection={r1:Math.max(a.r1,b.r1),r2:Math.min(a.r2,b.r2),c1:Math.max(a.c1,b.c1),c2:Math.min(a.c2,b.c2)};
  equal(book.Calculation.Evaluate(`=SUM((${addr(a)},${addr(b)}))`,{sheet}),sum(a)+sum(b));
  equal(book.Calculation.Evaluate(`=SUM(${addr(a)} ${addr(b)})`,{sheet}),intersection.r1>intersection.r2||intersection.c1>intersection.c2?'#NULL!':sum(intersection));
 }
});
test('Single-sheet INDEX ranges have correct dimensions for every row/column coordinate including zero',()=>{
 const {sheet,ev}=fixture(),source=sheet.GetRange('A1:C4').Values;
 for(let row=0;row<=4;row++)for(let col=0;col<=3;col++){
  const expected=source.filter((_,r)=>!row||r===row-1).map(a=>a.filter((_,c)=>!col||c===col-1));
  equal(ev(`=INDEX(A1:C4,${row},${col})`),row&&col?expected[0][0]:expected);
 }
});
test('Worksheet changes repair formulas on protected sheets without bypassing user edit protection',()=>{
 const {book,summary}=fixture();summary.GetCell('A1').Formula='=SUM(Jan:Mar!A1:A3)';summary.Protect();book.ClearHistory();book.Worksheets.Get('Jan').Name='Renamed';
 assert.equal(summary.GetCell('A1').Value,51);assert.throws(()=>summary.GetCell('A1').Formula='=1',/protected/);
 book.Worksheets.Remove('Renamed');assert.equal(summary.GetCell('A1').Value,45);book.Undo();book.Undo();assert.equal(summary.GetCell('A1').Formula,'=SUM(Jan:Mar!A1:A3)');assert.equal(summary.IsProtected,true);
});
test('Formula strings, table selectors, and literal cells remain opaque to sheet removal and rename',()=>{
 const {book,summary}=fixture();summary.GetCell('D1').Value='=Jan!A1';summary.GetCell('D2').Formula='="Jan!A1"';book.Worksheets.Remove('Jan');assert.equal(summary.GetCell('D1').Value,'=Jan!A1');assert.equal(summary.GetCell('D2').Value,'Jan!A1');
 assert.equal(renameSheetReferences('=Table1[[Jan]:[Mar]]+"Jan:Mar!A1"','Jan','New'),'=Table1[[Jan]:[Mar]]+"Jan:Mar!A1"');
});
test('SHEET without an explicit evaluation context uses the active worksheet',()=>{
 const {book}=fixture();assert.equal(book.Calculation.Evaluate('=SHEET()'),4);book.Worksheets.Move('Summary',0);assert.equal(book.Calculation.Evaluate('=SHEET()'),1);
 assert.equal(new Workbook({createSheet:false}).Calculation.Evaluate('=SHEET()').code,'#REF!');
});
for(const formula of ['=SUM(OFFSET(A1,0,0,1,1,7))','=SUM(OFFSET(A1,0))','=SUM(INDIRECT("A1",TRUE,7))','=SUM(INDIRECT())'])test('Reference-returning functions validate arguments in nested consumers '+formula,()=>equal(fixture().ev(formula),'#VALUE!'));
test('Reversed 3-D spans and inward endpoint moves preserve correct membership',()=>{
 const {book,summary}=fixture();summary.GetCell('A1').Formula='=SUM(Mar:Jan!A1:A3)';book.Worksheets.Move('Jan',1);assert.equal(summary.GetCell('A1').Value,33);book.Undo();assert.equal(summary.GetCell('A1').Value,51);
 book.Worksheets.Move('Mar',1);assert.equal(summary.GetCell('A1').Value,33);book.Undo();assert.equal(summary.GetCell('A1').Value,51);
});
