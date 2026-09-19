import test from 'node:test';
import assert from 'node:assert/strict';
import {Workbook, captureRange, pasteSpecial, fillSeries, specialCells, serialDate, isError} from '../src/core.js';
const norm = value => isError(value) ? value.code : Array.isArray(value) ? value.map(norm) : value;
const setup = () => { const book = new Workbook(); return {book, sheet:book.ActiveWorksheet}; };
const date = s => serialDate(new Date(s+'T00:00:00Z'));

test('Paste Values snapshots calculation, errors and literal formula-looking text',()=>{
 const {book,sheet:s}=setup();s.GetRange('A1:A4').Values=[[2],[null],['=NOT_A_FORMULA()'],[null]];s.GetCell('A2').Formula='=A1*3';s.GetCell('A4').Formula='=1/0';
 const copy=s.GetRange('A1:A4').Capture();s.GetCell('A1').Value=5;
 const result=s.GetRange('C1').PasteSpecial(copy,{mode:'values'});assert.equal(result.Address,'C1:C4');
 assert.deepEqual(norm(result.Values),[[2],[6],['=NOT_A_FORMULA()'],['#DIV/0!']]);assert.equal(s.GetCell('C3').Formula,null);
 assert.throws(()=>copy.cells[0][0].value=500,TypeError);book.Dispose();
});
test('Paste Formulas translates mixed references across sheets and transposes cell layout',()=>{
 const {book,sheet:s}=setup();s.GetRange('A1:B2').Values=[[1,2],[3,4]];s.GetCell('A1').Formula='=$C1+C$1+$C$1';
 const other=book.Worksheets.Add('Other');const out=other.GetRange('D4').PasteSpecial(s.GetRange('A1:B2'),{mode:'formulas',transpose:true});
 assert.equal(out.Address,'D4:E5');assert.equal(other.GetCell('D4').Formula,'=$C4+F$1+$C$1');assert.equal(other.GetCell('E4').Value,3);assert.equal(other.GetCell('D5').Value,2);book.Dispose();
});
test('Paste Formats replaces old attributes without touching data or notes',()=>{
 const {book,sheet:s}=setup();s.GetCell('A1').Style={fill:'#123456',font:{bold:true}};s.GetCell('B1').Value=9;s.GetCell('B1').Comment='retained';s.GetCell('B1').Style={numberFormat:'0.00',font:{italic:true}};
 s.GetRange('B1').PasteSpecial(s.GetRange('A1'),{mode:'formats'});assert.equal(s.GetCell('B1').Value,9);assert.equal(s.GetCell('B1').Comment,'retained');
 assert.deepEqual(s.GetCell('B1').Style,{fill:'#123456',font:{bold:true}});book.Dispose();
});
for(const mode of ['valuesAndNumberFormats','formulasAndNumberFormats'])test(mode+' leaves unrelated style and comments intact',()=>{
 const {book,sheet:s}=setup();s.GetCell('A1').Formula='=2+3';s.GetCell('A1').Style={numberFormat:'0.00',fill:'#ffffff'};s.GetCell('C1').Style={fill:'#123456'};s.GetCell('C1').Comment='keep';
 s.GetRange('C1').PasteSpecial(s.GetRange('A1'),{mode});assert.equal(s.GetCell('C1').Value,5);assert.equal(s.GetCell('C1').Style.numberFormat,'0.00');assert.equal(s.GetCell('C1').Style.fill,'#123456');assert.equal(s.GetCell('C1').Comment,'keep');assert.equal(!!s.GetCell('C1').Formula,mode.startsWith('formulas'));book.Dispose();
});
test('Paste Special repeats full source tiles and snapshots overlapping destinations',()=>{
 const {book,sheet:s}=setup();s.GetRange('A1:A3').Values=[[1],[2],[3]];s.GetRange('A2:A7').PasteSpecial(s.GetRange('A1:A3'),{mode:'values'});
 assert.deepEqual(s.GetRange('A1:A7').Values,[[1],[1],[2],[3],[1],[2],[3]]);assert.throws(()=>s.GetRange('C1:C4').PasteSpecial(s.GetRange('A1:A3')),/multiples/);book.Dispose();
});
test('Skip blanks distinguishes an absent input from a formula returning empty text',()=>{
 const {book,sheet:s}=setup();s.GetCell('A1').Value=1;s.GetCell('A3').Formula='=""';s.GetRange('C1:C3').Values=[[9],[9],[9]];
 s.GetRange('C1:C3').PasteSpecial(s.GetRange('A1:A3'),{mode:'values',skipBlanks:true});assert.deepEqual(s.GetRange('C1:C3').Values,[[1],[9],['']]);book.Dispose();
});
for(const [operation,expected]of [['add',12],['subtract',8],['multiply',20],['divide',5]])test('Arithmetic paste '+operation+' is a snapshot and one undo',()=>{
 const {book,sheet:s}=setup();s.GetCell('A1').Value=2;s.GetRange('B1:C2').Value=10;book.ClearHistory();let changes=0;book.Changed.Subscribe(()=>changes++);
 s.GetRange('B1:C2').PasteSpecial(s.GetRange('A1'),{mode:'values',operation});assert.deepEqual(s.GetRange('B1:C2').Values,[[expected,expected],[expected,expected]]);assert.equal(changes,1);
 book.Undo();assert.deepEqual(s.GetRange('B1:C2').Values,[[10,10],[10,10]]);assert.equal(book.CanUndo,false);book.Redo();assert.equal(s.GetCell('C2').Value,expected);book.Dispose();
});
test('Arithmetic paste returns bounded spreadsheet errors, not Infinity or JavaScript NaN',()=>{
 const {book,sheet:s}=setup();s.GetRange('A1:A3').Values=[[0],['no'],[{error:'#N/A'}]];s.GetRange('B1:B3').Value=10;
 s.GetRange('B1:B3').PasteSpecial(s.GetRange('A1:A3'),{mode:'values',operation:'divide'});assert.deepEqual(norm(s.GetRange('B1:B3').Values),[['#DIV/0!'],['#VALUE!'],['#N/A']]);book.Dispose();
});
test('Paste rolls back all writes if a later value violates validation',()=>{
 const {book,sheet:s}=setup();s.GetRange('A1:A2').Values=[[2],[99]];s.GetRange('C1:C2').Values=[[3],[4]];s.AddValidation('C1:C2',{type:'whole',min:0,max:10});book.ClearHistory();
 const before=book.ToJSON(),revision=book.Revision;assert.throws(()=>s.GetRange('C1').PasteSpecial(s.GetRange('A1:A2'),{mode:'values'}),/validation/);
 assert.deepEqual(book.ToJSON(),before);assert.equal(book.Revision,revision);assert.equal(book.CanUndo,false);book.Dispose();
});
test('Paste respects worksheet protection and partially selected spill ownership',()=>{
 const {book,sheet:s}=setup();s.GetCell('A1').Value=2;s.GetCell('C1').Formula='=SEQUENCE(2)';let before=book.ToJSON();
 assert.throws(()=>s.GetRange('D1:D2').PasteSpecial(s.GetRange('A1:A2'),{mode:'values',transpose:true}),/multiples/);assert.deepEqual(book.ToJSON(),before);
 assert.throws(()=>s.GetRange('C2').PasteSpecial(s.GetRange('A1'),{mode:'values'}),/spilled/);s.Protect();before=book.ToJSON();
 assert.throws(()=>s.GetRange('E1').PasteSpecial(s.GetRange('A1'),{mode:'values'}),/protected/);assert.deepEqual(book.ToJSON(),before);book.Dispose();
});
test('Full spill sources require Values mode and preserve every calculated cell',()=>{
 const {book,sheet:s}=setup();s.GetCell('A1').Formula='=SEQUENCE(2,2)';assert.throws(()=>s.GetRange('D1').PasteSpecial(s.GetRange('A1:B2')),/Values/);
 s.GetRange('D1').PasteSpecial(s.GetRange('A1:B2'),{mode:'values'});assert.deepEqual(s.GetRange('D1:E2').Values,[[1,2],[3,4]]);book.Dispose();
});
test('Comments-only and column-width paste preserve cell values',()=>{
 const {book,sheet:s}=setup();s.GetCell('A1').Comment='source';s.SetColumnWidth(0,177);s.GetCell('C1').Value=8;
 s.GetRange('C1').PasteSpecial(s.GetRange('A1'),{mode:'comments'});assert.equal(s.GetCell('C1').Comment,'source');assert.equal(s.GetCell('C1').Value,8);
 s.GetRange('C1').PasteSpecial(s.GetRange('A1'),{mode:'columnWidths'});assert.equal(s._meta.columns[2].size,177);assert.equal(s.GetCell('C1').Value,8);book.Dispose();
});
test('Validation-only paste clips old rules without deleting validation outside the target',()=>{
 const {book,sheet:s}=setup();s.AddValidation('A1:A2',{type:'whole',min:0,max:5});s.AddValidation('C1:C6',{type:'whole',min:10,max:20});book.ClearHistory();
 s.GetRange('C3').PasteSpecial(s.GetRange('A1:A2'),{mode:'validation'});assert.equal(s.Validate(2,2,3),null);assert.ok(s.Validate(2,2,12));
 assert.equal(s.Validate(0,2,12),null);assert.ok(s.Validate(0,2,3));assert.equal(s.Validate(5,2,12),null);book.Undo();assert.equal(s.Validate(2,2,12),null);book.Dispose();
});
test('Custom validation capture rebases clipped origins and pastes relative formulas',()=>{
 const {book,sheet:s}=setup();s.AddValidation('A1:A5',{type:'custom',formula:'=A1>0'});s.GetRange('C4').PasteSpecial(s.GetRange('A3:A4'),{mode:'validation'});
 const rule=s._meta.validations.find(r=>r.range.c1===2);assert.equal(rule.formula,'=C4>0');assert.equal(s.Validate(3,2,1),null);assert.ok(s.Validate(4,2,-1));book.Dispose();
});
test('All-cell paste applies copied validation instead of the overwritten destination rule',()=>{
 const {book,sheet:s}=setup();s.GetCell('A1').Value=3;s.AddValidation('A1',{type:'whole',min:1,max:5});s.GetCell('C1').Value=12;s.AddValidation('C1',{type:'whole',min:10,max:20});
 s.GetRange('C1').PasteSpecial(s.GetRange('A1'));assert.equal(s.GetCell('C1').Value,3);assert.ok(s.Validate(0,2,12));book.Dispose();
});
test('Unsupported paste combinations fail before mutation',()=>{
 const {book,sheet:s}=setup();s.GetCell('A1').Value=1;s.AddValidation('A1',{type:'whole',min:0,max:5});const before=book.ToJSON();
 for(const options of [{mode:'wrong'},{operation:'add'},{mode:'validation',skipBlanks:true},{mode:'columnWidths',transpose:true}])assert.throws(()=>s.GetRange('C1').PasteSpecial(s.GetRange('A1'),options));
 assert.deepEqual(book.ToJSON(),before);assert.throws(()=>pasteSpecial(s.GetRange('C1'),{kind:'GridWebClipboard',cells:[]}),/range/);book.Dispose();
});
test('CopyFrom rejects unknown modes, handles Formats and preserves errors in Values mode',()=>{
 const {book,sheet:s}=setup();s.GetCell('A1').Formula='=1/0';s.GetCell('A1').Style={fill:'#123456'};
 s.GetRange('B1').CopyFrom(s.GetRange('A1'),'values');assert.equal(s.GetCell('B1').Value.code,'#DIV/0!');s.GetCell('B1').Value=8;s.GetRange('B1').CopyFrom(s.GetRange('A1'),'formats');assert.equal(s.GetCell('B1').Value,8);
 assert.throws(()=>s.GetRange('B1').CopyFrom(s.GetRange('A1'),'bogus'),/mode/);book.Dispose();
});
for(const [type,step,expected]of [['linear',2,[[1],[3],[5],[7]]],['linear',-2,[[1],[-1],[-3],[-5]]],['growth',2,[[1],[2],[4],[8]]],['growth',.5,[[1],[.5],[.25],[.125]]]])test(`Fill ${type}, step ${step}`,()=>{
 const {book,sheet:s}=setup();s.GetCell('A1').Value=1;book.ClearHistory();assert.equal(s.GetRange('A1:A4').FillSeries({type,step}),4);assert.deepEqual(s.GetRange('A1:A4').Values,expected);book.Undo();assert.deepEqual(s.GetRange('A1:A4').Values,[[1],[null],[null],[null]]);book.Dispose();
});
for(const [direction,seedAddress,expected]of [['down','A1',[[1],[2],[3]]],['up','A3',[[3],[2],[1]]],['right','A1',[[1,2,3]]],['left','C1',[[3,2,1]]]])test('Series direction '+direction,()=>{
 const {book,sheet:s}=setup();s.GetCell(seedAddress).Value=1;s.GetRange(direction==='down'||direction==='up'?'A1:A3':'A1:C1').FillSeries({direction});
 assert.deepEqual(s.GetRange(direction==='down'||direction==='up'?'A1:A3':'A1:C1').Values,expected);book.Dispose();
});
test('Fill series uses each starting column, explicit start and a non-destructive stop value',()=>{
 const {book,sheet:s}=setup();s.GetRange('A1:B1').Values=[[1,10]];s.GetRange('A1:B3').FillSeries({step:2});assert.deepEqual(s.GetRange('A1:B3').Values,[[1,10],[3,12],[5,14]]);
 s.GetRange('D1:D6').Value=99;assert.equal(s.GetRange('D1:D6').FillSeries({start:1,step:2,stop:5}),3);assert.deepEqual(s.GetRange('D1:D6').Values,[[1],[3],[5],[99],[99],[99]]);book.Dispose();
});
for(const [seed,unit,step,dates]of [
 ['2025-01-31','month',1,['2025-01-31','2025-02-28','2025-03-31','2025-04-30']],
 ['2024-02-29','year',1,['2024-02-29','2025-02-28','2026-02-28','2027-02-28']],
 ['2026-09-18','weekday',1,['2026-09-18','2026-09-21','2026-09-22','2026-09-23']],
 ['2026-09-21','weekday',-1,['2026-09-21','2026-09-18','2026-09-17','2026-09-16']],
 ['2026-09-19','weekday',1,['2026-09-19','2026-09-21','2026-09-22','2026-09-23']],
 ['2026-09-20','weekday',-1,['2026-09-20','2026-09-18','2026-09-17','2026-09-16']],
 ['2026-09-18','day',1,['2026-09-18','2026-09-19','2026-09-20','2026-09-21']]
])test(`Date series ${seed} ${unit} ${step}`,()=>{
 const {book,sheet:s}=setup();s.GetCell('A1').Value=date(seed);s.GetRange('A1:A4').FillSeries({type:'date',dateUnit:unit,step});assert.deepEqual(s.GetRange('A1:A4').Values,dates.map(d=>[date(d)]));book.Dispose();
});
test('Date series preserves time fractions, fictitious leap day, and month anchor',()=>{
 const {book,sheet:s}=setup();s.GetRange('A1:A3').FillSeries({type:'date',start:60.5,dateUnit:'month'});
 assert.deepEqual(s.GetRange('A1:A3').Values,[[60.5],[date('1900-03-29')+.5],[date('1900-04-29')+.5]]);book.Dispose();
});
test('Series preflight and transaction failures leave every cell and history unchanged',()=>{
 const {book,sheet:s}=setup();s.GetRange('A1:B1').Values=[[1,'no']];const before=book.ToJSON();assert.throws(()=>s.GetRange('A1:B3').FillSeries(),/starting/);assert.deepEqual(book.ToJSON(),before);
 s.GetCell('B1').Value=1;s.AddValidation('A2:B3',{type:'whole',max:2});book.ClearHistory();const validBefore=book.ToJSON();assert.throws(()=>s.GetRange('A1:B3').FillSeries(),/validation/);assert.deepEqual(book.ToJSON(),validBefore);assert.equal(book.CanUndo,false);
 for(const options of [{step:NaN},{type:'date',step:.5},{type:'growth',step:0},{direction:'diagonal'},{type:'date',start:2958465}])assert.throws(()=>s.GetRange('C1:C3').FillSeries(options));book.Dispose();
});
test('Go To Special distinguishes formula results, constants, notes and actual blank inputs',()=>{
 const {book,sheet:s}=setup();s.GetRange('A1:B3').Values=[[1,'text'],[true,null],[null,'']];s.GetCell('A2').Formula='=1/0';s.GetCell('B2').Formula='=""';s.GetCell('A3').Comment='note';
 const r=s.GetRange('A1:B3'),addresses=(type,options)=>r.SpecialCells(type,options).map(c=>c.Address);
 assert.deepEqual(addresses('formulas'),['A2','B2']);assert.deepEqual(addresses('constants'),['A1','B1','B3']);assert.deepEqual(addresses('blanks'),['A3']);
 assert.deepEqual(addresses('errors'),['A2']);assert.deepEqual(addresses('comments'),['A3']);assert.deepEqual(addresses('formulas',{valueTypes:['errors']}),['A2']);assert.deepEqual(addresses('constants',{valueTypes:['numbers']}),['A1']);book.Dispose();
});
test('Special Cells uses sparse full-column scans for formulas, not worksheet-sized allocations',()=>{
 const {book,sheet:s}=setup();s.GetCell('A1048576').Formula='=1+1';assert.deepEqual(s.GetRange('A:A').SpecialCells('formulas').map(r=>r.Address),['A1048576']);
 assert.throws(()=>s.GetRange('A:A').SpecialCells('blanks'),/250,000/);assert.throws(()=>s.GetRange('A1').SpecialCells('missing'),/type/);book.Dispose();
});
test('Visible and validated cells observe filters, row hiding and column hiding',()=>{
 const {book,sheet:s}=setup();s.GetRange('A1:B4').Values=[['Key','Value'],['yes',1],['no',2],['yes',3]];s.AddValidation('B2:B4',{type:'number'});s.SetFilter('A1:B4',[{column:0,value:'yes'}]);s.HideRows(3);s.HideColumns(1);
 assert.deepEqual(s.GetRange('A2:B4').SpecialCells('visible').map(r=>r.Address),['A2']);assert.deepEqual(s.GetRange('A2:B4').SpecialCells('validation').map(r=>r.Address),['B2','B3','B4']);book.Dispose();
});
test('Special error cells include errors inside dynamic spill results',()=>{
 const {book,sheet:s}=setup();s.GetCell('A1').Formula='={1,#N/A}';assert.deepEqual(s.GetRange('A1:B1').SpecialCells('errors').map(r=>r.Address),['B1']);assert.deepEqual(s.GetRange('A1:B1').SpecialCells('blanks'),[]);book.Dispose();
});
