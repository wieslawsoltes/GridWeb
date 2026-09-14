import test from 'node:test';
import assert from 'node:assert/strict';
import {Workbook,FormulaError,a1ToR1C1,r1c1ToA1,getFormulasR1C1,setFormulasR1C1} from '../src/core.js';
const origin='D5';
for(const [a,r]of [
 ['A1','R[-4]C[-3]'],['$A$1','R1C1'],['A$1','R1C[-3]'],['$A1','R[-4]C1'],['D5','RC'],
 ['A1:C3','R[-4]C[-3]:R[-2]C[-1]'],['$A:$C','C1:C3'],['3:5','R[-2]:R'],['$3:$5','R3:R5'],
 ["'My Sheet'!B2","'My Sheet'!R[-3]C[-2]"],['XFD1048576','R[1048571]C[16380]'],
 ['SUM(A1,$C5)+"A1 R1C1"','SUM(R[-4]C[-3],RC3)+"A1 R1C1"'],
 ['SUM(Table1[A1])+A1','SUM(Table1[A1])+R[-4]C[-3]'],
 ])test('R1C1 reference round trip '+a,()=>{assert.equal(a1ToR1C1('='+a,origin),'='+r);assert.equal(r1c1ToA1('='+r,origin),'='+a);});
for(const [r,a]of [['R[-1]C','#REF!'],['R1048577C1','#REF!'],['R1C16385','#REF!'],['R[+1]C[+2]','C2'],['RC','A1'],['R1C1+"RC[-1]"','$A$1+"RC[-1]"'],['Table1[R1C1]+RC','Table1[R1C1]+A1']])test('R1C1 boundary '+r,()=>assert.equal(r1c1ToA1(r),a));
test('reference origins must be valid',()=>{assert.throws(()=>a1ToR1C1('A1',{row:-1,column:0}));assert.throws(()=>r1c1ToA1('RC','XFE1'));});
test('relative formula matrices use each destination cell as origin',()=>{const b=new Workbook(),s=b.ActiveWorksheet,r=s.GetRange('C2:C3');s.GetRange('A2:B3').Values=[[2,3],[4,5]];setFormulasR1C1(r,[['=RC[-2]*RC[-1]'],['=RC[-2]*RC[-1]']]);assert.deepEqual(r.Formulas,[['=A2*B2'],['=A3*B3']]);assert.deepEqual(r.Values,[[6],[20]]);assert.deepEqual(getFormulasR1C1(r),[['=RC[-2]*RC[-1]'],['=RC[-2]*RC[-1]']]);b.Dispose();});
const cases=[
 ['=ADDRESS(2,3,2,FALSE)','R2C[3]'],['=ADDRESS(2,3,3,FALSE)','R[2]C3'],['=ADDRESS(2,3,4,FALSE)','R[2]C[3]'],['=ADDRESS(2,3,,FALSE)','R2C3'],['=ADDRESS(2,3,1,TRUE,"Sheet2")','Sheet2!$C$2'],
 ['=SUM(A1)',0],['=SUM(A2)',0],['=SUM(A1:A3)',4],['=SUM("3",TRUE)',4],['=COUNT(A1,A2,A3)',1],['=COUNT("3",TRUE,"bad",#N/A)',2],['=AVERAGE(A1:A3)',4],
 ['=SUM(OFFSET(A1,0,0))',0],['=SUM(INDIRECT("A1"))',0],['=SUM(TextRef)',0],['=SUM(NumberName)',7],
 ['=ABS({-1,-2;-3,-4})',[[1,2],[3,4]]],['=ROUND({1.234,2.456},2)',[[1.23,2.46]]],['=POWER({1;2},{2,3})',[[1,1],[4,8]]],
 ['=IFERROR({1,#N/A},{8,9})',[[1,9]]],['=IFNA({#N/A,#VALUE!},{8,9})',[[8,'#VALUE!']]],
 ['=ROW(A2:A4)',[[2],[3],[4]]],['=COLUMN(B1:D1)',[[2,3,4]]],['=@A1:A3',true],['=@A1:C3','#VALUE!'],['=@A4:A5','#VALUE!'],
 ['=INDIRECT("R3C1",FALSE)',4],['=INDIRECT("RC[-1]",FALSE)',true],['=SUM(INDIRECT("R[-1]C[-1]:R[1]C[-1]",FALSE))',4],
 ['=TEXTBEFORE("ß.A","a",1,1)','ß.'],['=TEXTAFTER("ß.A.tail","a",1,1)','.tail'],['=SUM()','#VALUE!'],['=SUM(1,2,3)',6]
];
function normalize(v){return v instanceof FormulaError?v.code:Array.isArray(v)?v.map(normalize):v;}
for(const [formula,expected]of cases)test('reference semantics '+formula,()=>{const b=new Workbook(),s=b.ActiveWorksheet;s.GetCell('A1').Input="'3";s.GetCell('A2').Value=true;s.GetCell('A3').Value=4;b.DefineName('TextRef','=A1');b.DefineName('NumberName',7);assert.deepEqual(normalize(b.Calculation.Evaluate(formula,{sheet:s,row:1,col:1})),expected);b.Dispose();});
test('INDIRECT R1C1 tracks dependency changes',()=>{const b=new Workbook(),s=b.ActiveWorksheet;s.GetCell('A2').Value=8;s.GetCell('B2').Formula='=INDIRECT("RC[-1]",FALSE)*2';assert.equal(s.GetCell('B2').Value,16);s.GetCell('A2').Value=11;assert.equal(s.GetCell('B2').Value,22);b.Dispose();});
test('scalar function broadcasting creates real worksheet spills',()=>{const b=new Workbook(),s=b.ActiveWorksheet;s.GetRange('A1:A3').Values=[[-1],[-2],[-3]];s.GetCell('C1').Formula='=ABS(A1:A3)';assert.deepEqual(s.GetRange('C1:C3').Values,[[1],[2],[3]]);b.Dispose();});
test('stable inverse tails round trip at relative precision',()=>{const b=new Workbook(),f=b.Calculation.Functions;for(const p of [1e-12,1e-8,1e-4])for(const [inverse,dist,args]of [['T.INV.2T','T.DIST.2T',[1]],['CHISQ.INV.RT','CHISQ.DIST.RT',[8]],['F.INV.RT','F.DIST.RT',[5,10]]]){const x=f.get(inverse)(p,...args),got=f.get(dist)(x,...args);assert.ok(Math.abs(got/p-1)<1e-7,`${inverse}(${p}) = ${x}; roundtrip ${got}`);}b.Dispose();});
