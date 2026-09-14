import test from 'node:test';
import assert from 'node:assert/strict';
import {Workbook,FormulaError,createFunctionRegistry} from '../src/core.js';
const book=new Workbook(), sheet=book.ActiveWorksheet;
const evaluate=formula=>{const v=book.Calculation.Evaluate(formula,{sheet});return v instanceof FormulaError?v.code:v;};
const near=(a,b,tolerance=2e-11)=>assert.ok(typeof a==='number' && Math.abs(a-b)<=tolerance*Math.max(1,Math.abs(b)),`${a} != ${b}`);
const cases=[
 ['=XMATCH(5,{1,3,5,7},0,2)',3],['=XMATCH(5,{7,5,3,1},0,-2)',2],
 ['=XMATCH(4,{1,3,5,7},-1,2)',2],['=XMATCH(4,{1,3,5,7},1,2)',3],
 ['=XMATCH(4,{7,5,3,1},-1,-2)',3],['=XMATCH(4,{7,5,3,1},1,-2)',2],
 ['=XMATCH(0,{1,3,5},-1,2)','#N/A'],['=XMATCH(9,{5,3,1},1,-2)','#N/A'],
 ['=XMATCH(3,{1,3,3,5},0,-1)',3],['=XMATCH("ab*",{"AA","Abcd","ab*"},2,-1)',3],
 ['=XMATCH("ab~*",{"abcd","ab*"},2)',2],['=XMATCH(1,{1,2;3,4})','#VALUE!'],['=XMATCH(1,{1,2},0,3)','#VALUE!'],
 ['=XLOOKUP(2,{1;2},{10,11;20,21})',[[20,21]]],['=XLOOKUP(2,{1,2},{10,20;11,21})',[[20],[21]]],
 ['=XLOOKUP(2,{1,2},{10})','#VALUE!'],['=XLOOKUP(7,{1,3,5},{10,30,50},"missing",-1,2)',50],
 ['=TEXTBEFORE("Alpha-BETA-tail","beta",1,1)','Alpha-'],['=TEXTAFTER("Alpha-BETA-tail","beta",1,1)','-tail'],
 ['=TEXTBEFORE("a/b/c","/",-1)','a/b'],['=TEXTAFTER("a/b/c","/",-2)','b/c'],
 ['=TEXTBEFORE("abc","x",1,0,1)','abc'],['=TEXTAFTER("abc","x",1,0,1)',''],
 ['=TEXTAFTER("abc","x",1,0,0,"missing")','missing'],['=TEXTBEFORE("abc","x",0)','#VALUE!'],
 ['=TEXTBEFORE("abc","",-1)','abc'],['=TEXTAFTER("abc","")','abc'],
 ['=TEXTSPLIT("a,b;c|d",{",",";","|"})',[['a','b','c','d']]],
 ['=TEXTSPLIT("a,b;c",",",";",FALSE,0,"pad")',[['a','b'],['c','pad']]],
 ['=TEXTSPLIT("aXbxc","x",,FALSE,1)',[['a','b','c']]],['=TEXTSPLIT("a,,b",",",,TRUE)',[['a','b']]],
 ['=FILTER({1,2,3;4,5,6},{TRUE,FALSE,TRUE})',[[1,3],[4,6]]],
 ['=EXPAND({1,2},2,3,0)',[[1,2,0],[0,0,0]]],['=EXPAND({1,2},1,1)','#VALUE!'],
 ['=WRAPROWS({1,2,3,4,5},3,0)',[[1,2,3],[4,5,0]]],['=WRAPCOLS({1,2,3,4,5},3,0)',[[1,4],[2,5],[3,0]]],
 ['=SORTBY({"a",2;"b",1;"c",2},{2;1;2},1,{1;2;3},-1)',[['b',1],['c',2],['a',2]]],
 ['=NETWORKDAYS.INTL(DATE(2026,9,7),DATE(2026,9,13),"0000011")',5],
 ['=NETWORKDAYS.INTL(DATE(2026,9,7),DATE(2026,9,13),11)',6],
 ['=NETWORKDAYS.INTL(DATE(2026,9,13),DATE(2026,9,7),1)',-5],
 ['=WORKDAY.INTL(DATE(2026,9,11),1,1)',46279],['=WORKDAY.INTL(DATE(2026,9,14),-1,1)',46276],
 ['=ISOWEEKNUM(DATE(2021,1,1))',53],['=WEEKNUM(DATE(2021,1,1),21)',53],
 ['=RANK.AVG(2,{1,2,2,4},1)',2.5],['=PERCENTILE.EXC({1,2,3,4},0.5)',2.5],['=QUARTILE.EXC({1,2,3,4,5,6,7},1)',2],
 ['=AVEDEV({1,2,3})',2/3],['=DEVSQ({1,2,3})',2],['=GEOMEAN({1,4,16})',4],['=HARMEAN({1,2,4})',12/7],
 ['=MODE.MULT({2,1,2,1,3})',[[2],[1]]],['=MODE.SNGL({2,1,2,1,3})',1],['=SKEW({1,2,3})',0],['=KURT({1,2,3,4,5})',-1.2],
 ['=COVARIANCE.P({2,4,6},{1,2,3})',4/3],['=COVARIANCE.S({2,4,6},{1,2,3})',2],['=CORREL({2,4,6},{1,2,3})',1],
 ['=SLOPE({3,5,7},{1,2,3})',2],['=INTERCEPT({3,5,7},{1,2,3})',1],['=FORECAST.LINEAR(4,{3,5,7},{1,2,3})',9],
 ['=STANDARDIZE(10,4,2)',3],['=GAMMA(5)',24],['=GAMMALN(5)',Math.log(24)],['=GAMMA(-0.5)',-2*Math.sqrt(Math.PI)],
 ['=BETA.DIST(0.5,2,2,TRUE)',.5],['=BETA.DIST(0,1,1,FALSE)',1],['=BETA.INV(0.5,2,2)',.5],
 ['=GAMMA.DIST(2,1,2,TRUE)',1-Math.exp(-1)],['=GAMMA.INV(0.5,1,2)',2*Math.LN2],
 ['=NORM.S.DIST(0,TRUE)',.5],['=NORM.S.DIST(0,FALSE)',1/Math.sqrt(2*Math.PI)],['=NORM.S.INV(0.975)',1.959963984540054],
 ['=NORM.DIST(10,10,2,TRUE)',.5],['=NORM.INV(0.5,10,2)',10],['=LOGNORM.DIST(1,0,1,TRUE)',.5],['=LOGNORM.INV(0.5,0,1)',1],
 ['=CHISQ.DIST(2,2,TRUE)',1-Math.exp(-1)],['=CHISQ.INV(0.5,2)',2*Math.LN2],
 ['=T.DIST(0,5,TRUE)',.5],['=T.DIST(1,1,TRUE)',.75],['=T.INV(0.75,1)',1],['=T.INV.2T(0.5,1)',1],
 ['=F.DIST(1,5,5,TRUE)',.5],['=F.INV(0.5,5,5)',1],['=BINOM.DIST(2,4,0.5,FALSE)',.375],['=BINOM.DIST(2,4,0.5,TRUE)',.6875],
 ['=BINOM.INV(4,0.5,0.5,0.5)',2],['=POISSON.DIST(2,2,FALSE)',2*Math.exp(-2)],['=POISSON.DIST(2,2,TRUE)',5*Math.exp(-2)],
 ['=EXPON.DIST(1,2,TRUE)',1-Math.exp(-2)],['=WEIBULL.DIST(1,2,1,TRUE)',1-Math.exp(-1)],['=ERF(1)',.8427007929497149],
 ['=MMULT({1,2;3,4},{5;6})',[[17],[39]]],['=MDETERM({1,2;3,4})',-2],['=MINVERSE({2,0;0,4})',[[.5,0],[0,.25]]],['=MUNIT(2)',[[1,0],[0,1]]],
 ['=BITAND(1099511627779,3)',3],['=BITOR(1099511627776,3)',1099511627779],['=BITXOR(7,3)',4],['=BITLSHIFT(1,40)',1099511627776],
 ['=BITRSHIFT(32,3)',4],['=BITLSHIFT(1,48)','#NUM!'],['=BASE(31,16,4)','001F'],['=DECIMAL("FF",16)',255],
 ['=DEC2BIN(-1)','1111111111'],['=BIN2DEC("1111111111")',-1],['=DEC2HEX(-1)','FFFFFFFFFF'],['=HEX2DEC("FFFFFFFFFF")',-1],
 ['=OCT2BIN("7777777777")','1111111111'],['=HEX2OCT("FFFFFFFFFF")','7777777777'],['=DEC2BIN(512)','#NUM!'],
 ['=COMPLEX(3,4)','3+4i'],['=IMREAL("3+4i")',3],['=IMAGINARY("3-4j")',-4],['=IMABS("3+4i")',5],
 ['=IMCONJUGATE("3+4i")','3-4i'],['=IMSUM("3+4i","2-4i")','5'],['=IMPRODUCT("1+i","1-i")','2'],['=IMDIV("3+4i","1+i")','3.5+0.5i'],
 ['=IMLN(1)','0'],['=IMEXP(0)','1'],['=IMREAL("1e+3-2e-3j")',1000],
 ['=EFFECT(0.12,12)',1.01**12-1],['=NOMINAL(EFFECT(0.12,12),12)',.12],['=IPMT(0.01,1,12,100)',-1],
 ['=XNPV(0.1,{-100,110},{DATE(2025,1,1),DATE(2026,1,1)})',0],['=XIRR({-100,110},{DATE(2025,1,1),DATE(2026,1,1)})',.1],
 ['=NORM.S.INV(0)','#NUM!'],['=MINVERSE({1,2;2,4})','#NUM!'],['=WORKDAY.INTL(40000,1,"1111111")','#VALUE!'],['=BITAND(1.5,2)','#NUM!']
];
for(const [formula,expected]of cases)test('extended '+formula,()=>{const actual=evaluate(formula);if(typeof expected==='number')near(actual,expected);else assert.deepEqual(actual,expected);});
test('binary and linear lookup agree on sorted unique ascending/descending vectors',()=>{const f=createFunctionRegistry().get('XMATCH');for(const sign of [-1,1]){const a=Array.from({length:101},(_,i)=>i*3);if(sign<0)a.reverse();for(let n=-1;n<=303;n++)for(const mode of [-1,0,1]){const call=search=>{try{return f(n,[a],mode,search);}catch(e){return e.code;}};assert.equal(call(sign*2),call(1),`${n} / ${mode} / ${sign}`);}}});
test('inverse distribution round trips across central and tail probabilities',()=>{const f=createFunctionRegistry();for(const p of [1e-9,.001,.1,.5,.9,.999,1-1e-9]){near(f.get('NORM.S.DIST')(f.get('NORM.S.INV')(p),true),p,1e-12);near(f.get('BETA.DIST')(f.get('BETA.INV')(p,2,5),2,5,true),p,1e-12);near(f.get('GAMMA.DIST')(f.get('GAMMA.INV')(p,3,2),3,2,true),p,1e-12);}});
test('matrix inversion with row pivoting multiplies to identity',()=>{const f=createFunctionRegistry(),a=[[0,2,1],[1,1,0],[2,0,1]],product=f.get('MMULT')(a,f.get('MINVERSE')(a));product.forEach((r,i)=>r.forEach((v,j)=>near(v,+(i===j))));});
test('dynamic lookup arrays spill and dependents recalculate',()=>{const b=new Workbook(),s=b.ActiveWorksheet;s.GetRange('A1:C2').Values=[[1,10,11],[2,20,21]];s.GetCell('E1').Formula='=XLOOKUP(2,A1:A2,B1:C2)';assert.equal(s.GetCell('F1').Value,21);s.GetCell('C2').Value=42;assert.equal(s.GetCell('F1').Value,42);b.Dispose();});
import fs from 'node:fs';
const numericalFixtures=JSON.parse(fs.readFileSync(new URL('./fixtures/distributions.json',import.meta.url),'utf8'));
for(const [name,args,expected] of numericalFixtures.cases)test('numerical reference '+name+' '+JSON.stringify(args),()=>near(createFunctionRegistry().get(name)(...args),expected,2e-9));
