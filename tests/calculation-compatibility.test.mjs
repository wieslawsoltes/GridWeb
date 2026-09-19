import test from 'node:test';
import assert from 'node:assert/strict';
import {Workbook, createFunctionRegistry, isError} from '../src/core.js';
const normalize = value => isError(value) ? value.code : Array.isArray(value) ? value.map(normalize) : value;
function evaluate(formula, setup = () => {}) {
  const book = new Workbook();
  try { setup(book, book.ActiveWorksheet); return normalize(book.Calculation.Evaluate(formula, {sheet: book.ActiveWorksheet})); }
  finally { book.Dispose(); }
}
const lambdaCases = [
  ['=LAMBDA(x,x+1)(4)',5], ['=LAMBDA(42)()',42], ['=LAMBDA(x,y,x*y)(3,4)',12],
  ['=LET(add,LAMBDA(x,y,x+y),add(8,9))',17], ['=LET(add,LAMBDA(x,x+1),add)(3)',4],
  ['=LAMBDA(x,LAMBDA(y,x+y))(2)(3)',5], ['=(LAMBDA(x,x^2))(4)',16],
  ['=LAMBDA(x,x)(2)^2',4], ['=LAMBDA(x,y,IF(ISOMITTED(y),x*2,x+y))(3,)',6],
  ['=LAMBDA(x,y,ISOMITTED(y))(3,0)',false], ['=LAMBDA(x,y,ISOMITTED(y))(3,"")',false],
  ['=LAMBDA(x,y,LET(y,10,ISOMITTED(y)))(3,)',false],
  ['=LAMBDA(x,y,LAMBDA(z,ISOMITTED(y))(1))(3,)',true],
  ['=ISOMITTED(0)',false], ['=ISOMITTED()', '#VALUE!'], ['=ISOMITTED(1,2)','#VALUE!'],
  ['=LAMBDA(x,x)(1,2)','#VALUE!'], ['=LAMBDA(x,y,x+y)(1)','#VALUE!'],
  ['=LAMBDA(x,x,x)(1,2)','#VALUE!'], ['=LAMBDA(a.b,a.b)(1)','#VALUE!'],
  ['=LAMBDA(1,2)(3)','#VALUE!'], ['=2(3)','#VALUE!'],
  ['=MAKEARRAY(3,3,LAMBDA(row,col,row*col))',[[1,2,3],[2,4,6],[3,6,9]]],
  ['=MAKEARRAY(1,3,LAMBDA(row,col,col))',[[1,2,3]]],
  ['=MAKEARRAY(0,2,LAMBDA(row,col,row))','#VALUE!'], ['=MAKEARRAY(-1,2,LAMBDA(row,col,row))','#VALUE!'],
  ['=MAKEARRAY(1000,1000,LAMBDA(row,col,row))','#NUM!'], ['=MAKEARRAY(1,1,LAMBDA(x,x))','#VALUE!'],
  ['=MAKEARRAY(1,1,2)','#VALUE!'], ['=MAKEARRAY(1,1,LAMBDA(x,y,{1,2}))','#CALC!'],
  ['=MAP({1,2;3,4},LAMBDA(x,x*2))',[[2,4],[6,8]]],
  ['=MAP({1,2},{10,20},LAMBDA(x,y,x+y))',[[11,22]]],
  ['=MAP({1,2},{10;20},LAMBDA(x,y,x+y))','#VALUE!'],
  ['=MAP({1,2},LAMBDA(x,y,x+y))','#VALUE!'], ['=MAP({1,2},LAMBDA(x,{1,2}))','#CALC!'],
  ['=BYROW({1,2;3,4},LAMBDA(row,SUM(row)))',[[3],[7]]],
  ['=BYCOL({1,2;3,4},LAMBDA(col,SUM(col)))',[[4,6]]],
  ['=BYROW({1,2},LAMBDA(row,row))','#CALC!'], ['=BYCOL({1;2},LAMBDA(col,col))','#CALC!'],
  ['=BYROW({1,2},LAMBDA(x,y,x))','#VALUE!'], ['=BYROW({1,2},LAMBDA(x,1),3)','#VALUE!'],
  ['=SCAN(0,{1,2;3,4},LAMBDA(acc,v,acc+v))',[[1,3],[6,10]]],
  ['=REDUCE(0,{1,2;3,4},LAMBDA(acc,v,acc+v))',10],
  ['=REDUCE(0,{1,2},LAMBDA(acc,v,VSTACK(acc,v)))',[[0],[1],[2]]],
  ['=SCAN(0,{1,2},LAMBDA(acc,v,VSTACK(acc,v)))','#CALC!'],
  ['=SCAN(0,{1,2},LAMBDA(v,v))','#VALUE!'], ['=MAP({1,#N/A},LAMBDA(x,IFERROR(x,0)))',[[1,0]]]
];
for (const [formula, result] of lambdaCases) test('lambda compatibility '+formula, () => assert.deepEqual(evaluate(formula), result));
test('Named recursive LAMBDA evaluates and bounds nonterminating recursion', () => {
  assert.equal(evaluate('=FACTORIAL(6)', b => b.DefineName('FACTORIAL','=LAMBDA(n,IF(n<=1,1,n*FACTORIAL(n-1)))')),720);
  assert.equal(evaluate('=LOOP(1)', b => b.DefineName('LOOP','=LAMBDA(n,LOOP(n+1))')),'#NUM!');
});
test('Omitted arguments survive named LAMBDA calls but blank cells are supplied arguments', () => {
  assert.equal(evaluate('=OPTIONAL(3,)', b => b.DefineName('OPTIONAL','=LAMBDA(x,y,ISOMITTED(y))')),true);
  assert.equal(evaluate('=OPTIONAL(3,A1)', b => b.DefineName('OPTIONAL','=LAMBDA(x,y,ISOMITTED(y))')),false);
});
test('New lambda spills invalidate their dependents and blocked spills recover', () => {
  const b = new Workbook(), s = b.ActiveWorksheet;
  s.GetCell('A1').Value=3; s.GetCell('D1').Formula='=MAKEARRAY(2,2,LAMBDA(row,col,A1*row*col))'; s.GetCell('G1').Formula='=E2+1';
  assert.equal(s.GetCell('G1').Value,13);s.GetCell('A1').Value=5;assert.equal(s.GetCell('G1').Value,21);
  s.GetRange('D1:E2').Clear();s.GetCell('E2').Value=99;s.GetCell('D1').Formula='=MAKEARRAY(2,2,LAMBDA(row,col,row+col))';
  assert.equal(s.GetCell('D1').Value.code,'#SPILL!');s.GetCell('E2').Input=null;assert.equal(s.GetCell('G1').Value,5);b.Dispose();
});

const database = [['Tree','Height','Age','Yield','Profit'],['Apple',18,20,14,105],['Pear',12,12,10,96],['Cherry',13,14,9,105],['Apple',14,15,10,75],['Pear',9,8,8,77],['Apple',8,9,6,45]];
const allCriteria = [['Tree'],[null]], apple = [['Tree'],['=Apple']];
const dbCases = [
  ['DSUM','Profit',apple,225], ['DAVERAGE','Profit',apple,75], ['DCOUNT','Profit',apple,3], ['DCOUNTA','Tree',apple,3],
  ['DMAX','Profit',apple,105], ['DMIN','Profit',apple,45], ['DPRODUCT','Yield',apple,840],
  ['DSTDEV','Profit',apple,30], ['DSTDEVP','Profit',apple,Math.sqrt(600)], ['DVAR','Profit',apple,900], ['DVARP','Profit',apple,600],
  ['DGET','Profit',[['Tree'],['Cherry']],105], ['DGET','Profit',apple,'#NUM!'], ['DGET','Profit',[['Tree'],['No tree']],'#VALUE!'],
  ['DSUM',5,apple,225], ['DSUM','profit',apple,225], ['DSUM','Profit',[['Tree','Height','Height'],['=Apple','>10','<16'],['=Pear',null,null]],248],
  ['DSUM','Profit',[['Tree'],['App']],225], ['DSUM','Profit',[['Tree'],['=App']],0], ['DSUM','Profit',[['Tree'],['?ear']],173],
  ['DSUM','Profit',allCriteria,503], ['DCOUNT',null,apple,3], ['DCOUNTA',null,apple,3], ['DSUM','Missing',apple,'#VALUE!'],
  ['DSUM',0,apple,'#VALUE!'], ['DSUM','Profit',[['Missing'],['a']],'#VALUE!'], ['DSUM','Profit',[['Tree']],'#VALUE!'],
];
for (const [name,field,criteria,expected] of dbCases) test('database '+name+' '+JSON.stringify([field,criteria]), () => {
  let result;try{result=normalize(createFunctionRegistry().get(name)(database,field,criteria));}catch(e){result=e.code;}
  if(typeof expected==='number')assert.ok(Math.abs(result-expected)<1e-10, String(result));else assert.equal(result,expected);
});
test('Database criteria formulas are rebound for every record and track live dependencies', () => {
  const b = new Workbook(), s=b.ActiveWorksheet;
  s.GetRange('A1:E7').Values=database;s.GetCell('G1').Value='Calculated';s.GetCell('G2').Formula='=E2>$I$1';s.GetCell('I1').Value=80;
  s.GetCell('K1').Formula='=DSUM(A1:E7,"Profit",G1:G2)';assert.equal(s.GetCell('K1').Value,306);
  s.GetCell('I1').Value=100;assert.equal(s.GetCell('K1').Value,210);
  s.GetCell('E2').Value=40;assert.equal(s.GetCell('K1').Value,105);b.Dispose();
});
test('Database references, named criteria, errors and primitive distinctions', () => {
  const b = new Workbook(),s=b.ActiveWorksheet;
  s.GetRange('A1:B6').Values=[['Key','Value'],['x',1],['x',true],['x','2'],['x',null],['x',{error:'#N/A'}]];
  s.GetRange('D1:D2').Values=[['Key'],['x']];b.DefineName('CRITERIA','=Sheet1!D1:D2');
  const ev=f=>normalize(b.Calculation.Evaluate(f,{sheet:s}));
  assert.equal(ev('=DCOUNT(A1:B6,"Value",CRITERIA)'),1);assert.equal(ev('=DCOUNTA(A1:B6,"Value",CRITERIA)'),4);
  assert.equal(ev('=DSUM(A1:B6,"Value",CRITERIA)'),'#N/A');assert.equal(ev('=DSUM(#REF!,2,D1:D2)'),'#REF!');
  assert.equal(ev('=DSUM(A1:B6,2)'),'#VALUE!');b.Dispose();
});

function oracle(index, values) {
  const nums=values.filter(x=>typeof x==='number'), n=nums.length, sum=nums.reduce((a,b)=>a+b,0), mean=sum/n, sorted=[...nums].sort((a,b)=>a-b);
  if(index===2)return n;if(index===3)return values.length;
  if(values.includes('#DIV/0!'))return '#DIV/0!';
  const variance=nums.reduce((a,b)=>a+(b-mean)**2,0);
  if(index===1)return mean;if(index===4)return Math.max(...nums);if(index===5)return Math.min(...nums);if(index===6)return nums.reduce((a,b)=>a*b,1);
  if(index===7)return Math.sqrt(variance/(n-1));if(index===8)return Math.sqrt(variance/n);if(index===9)return sum;if(index===10)return variance/(n-1);if(index===11)return variance/n;
  if(index===13)return nums.filter(x=>x===10).length>1?10:'#N/A';if(index===14)return sorted.at(-2);if(index===15)return sorted[1];
  // MEDIAN and the four 50th-percentile/quartile forms have the same center.
  return n%2?sorted[(n-1)/2]:(sorted[n/2-1]+sorted[n/2])/2;
}
for(let index=1;index<=19;index++)for(let option=0;option<8;option++)test(`AGGREGATE function ${index}, option ${option}`,()=>{
  const b=new Workbook(),s=b.ActiveWorksheet;
  s.GetRange('A1:B8').Values=[['Number','Visible'],[1,'yes'],[2,'yes'],[3,'no'],[4,'yes'],[{error:'#DIV/0!'},'yes'],[null,'yes'],[null,'yes']];
  s.GetRange('D1:D2').Values=[[5],[5]];s.GetCell('A7').Formula='=SUBTOTAL(9,D1:D2)';s.GetCell('A8').Formula='=AGGREGATE(9,4,D1:D2)';
  s.HideRows(2);s.SetFilter('A1:B8',[{column:1,value:'yes'}]);
  let values=[1,...(option&1?[]:[2]),4,...(option&2?[]:['#DIV/0!']),...(option<4?[]:[10,10])];
  const k=index>=14?','+([16,18].includes(index)?'.5':2):'';
  const actual=normalize(b.Calculation.Evaluate(`=AGGREGATE(${index},${option},A2:A8${k})`,{sheet:s})),expected=oracle(index,values);
  if(typeof expected==='number')assert.ok(Math.abs(actual-expected)<1e-9,`${actual} != ${expected}`);else assert.equal(actual,expected);b.Dispose();
});
for(const [formula,expected] of [
  ['=AGGREGATE(9,6,{1,#N/A,3})',4],['=AGGREGATE(14,6,{1,#N/A,3},1)',3],
  ['=AGGREGATE(9,,{1,2,3})',6],['=AGGREGATE(0,0,{1})','#VALUE!'],['=AGGREGATE(20,0,{1})','#VALUE!'],
  ['=AGGREGATE(9,-1,{1})','#VALUE!'],['=AGGREGATE(9,8,{1})','#VALUE!'],['=AGGREGATE(15,6,{1,2})','#VALUE!'],
  ['=AGGREGATE(14,6,{1,2},1,2)','#VALUE!'],['=SUBTOTAL(200,A1)','#VALUE!'],['=SUBTOTAL(9,{1,2})','#VALUE!']
])test('aggregate validation '+formula,()=>assert.equal(evaluate(formula),expected));
test('Aggregations include spills, named/OFFSET references and invalidation of initially empty full columns',()=>{
  const b=new Workbook(),s=b.ActiveWorksheet;s.GetCell('A1').Formula='=SEQUENCE(3)';b.DefineName('SERIES','=A1:A3');s.GetCell('C1').Formula='=SUBTOTAL(9,SERIES)';
  assert.equal(s.GetCell('C1').Value,6);s.GetCell('D1').Formula='=AGGREGATE(9,4,OFFSET(A1,0,0,3,1))';assert.equal(s.GetCell('D1').Value,6);
  s.GetCell('G1').Formula='=AGGREGATE(9,4,F:F)';assert.equal(s.GetCell('G1').Value,0);s.GetCell('F1000').Value=5;assert.equal(s.GetCell('G1').Value,5);
  s.GetCell('I1').Formula='="SUBTOTAL(9,A1)"';s.GetCell('J1').Formula='=AGGREGATE(3,0,I1)';assert.equal(s.GetCell('J1').Value,1);b.Dispose();
});
test('Computed arrays retain hidden/nested values; SUBTOTAL includes hidden rows only in 1–11 modes',()=>{
  const b=new Workbook(),s=b.ActiveWorksheet;s.GetRange('A1:A3').Values=[[1],[2],[3]];s.HideRows(1);
  const ev=f=>normalize(b.Calculation.Evaluate(f,{sheet:s}));assert.equal(ev('=AGGREGATE(9,1,A1:A3)'),4);assert.equal(ev('=AGGREGATE(9,1,A1:A3*1)'),6);
  assert.equal(ev('=SUBTOTAL(9,A1:A3)'),6);assert.equal(ev('=SUBTOTAL(109,A1:A3)'),4);b.Dispose();
});
