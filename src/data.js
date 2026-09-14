import { compare } from './functions.js';
import { isError, scalar } from './errors.js';
import { cellAddress, MAX_OPERATION_CELLS } from './address.js';
export class PivotTable {
  constructor(source, { rows = [0], columns = [], values = [{ column: 1, aggregate: 'sum', name: 'Total' }], filters = [] } = {}) {
    this.Source = source; this.Options = { rows, columns, values, filters }; this.Result = []; this.Refresh();
  }
  Refresh() {
    const matrix = this.Source.Values, headers = matrix[0] ?? [], data = matrix.slice(1), { rows, columns, values, filters } = this.Options;
    const index = value => typeof value === 'string' ? headers.indexOf(value) : value;
    const rs = rows.map(index), cs = columns.map(index), vs = values.map(v => ({ ...v, column: index(v.column) }));
    if ([...rs, ...cs, ...vs.map(v => v.column)].some(v => !Number.isInteger(v) || v < 0 || v >= headers.length)) throw new RangeError('Unknown pivot field');
    const groups = new Map(), columnKeys = new Map();
    for (const row of data) {
      if (!filters.every(f => (f.values ?? []).some(v => compare(v, row[index(f.column)]) === 0))) continue;
      const rk = JSON.stringify(rs.map(i => row[i])), ck = JSON.stringify(cs.map(i => row[i]));
      columnKeys.set(ck, cs.map(i => row[i]));
      if (!groups.has(rk)) groups.set(rk, { labels: rs.map(i => row[i]), columns: new Map() });
      const group = groups.get(rk); if (!group.columns.has(ck)) group.columns.set(ck, vs.map(() => []));
      vs.forEach((v, i) => group.columns.get(ck)[i].push(row[v.column]));
    }
    const keys = [...columnKeys.keys()].sort(); if (!keys.length) keys.push('[]');
    const aggregate = (items, type) => {
      if (type === 'count') return items.filter(x => x != null && x !== '').length;
      if (type === 'distinct') return new Set(items.map(x => JSON.stringify(x))).size;
      const n = items.filter(x => typeof x === 'number'); if (!n.length) return 0;
      if (type === 'average') return n.reduce((a,b)=>a+b,0)/n.length;
      if (type === 'min') return n.reduce((a,b)=>Math.min(a,b)); if (type === 'max') return n.reduce((a,b)=>Math.max(a,b));
      if (type !== 'sum') throw new TypeError('Unknown pivot aggregate: '+type); return n.reduce((a,b)=>a+b,0);
    };
    this.Result = [[...rs.map(i=>headers[i]),...keys.flatMap(k=>vs.map(v=>[...JSON.parse(k),v.name??headers[v.column]+' '+v.aggregate].join(' / ')))], ...[...groups.values()].sort((a,b)=>String(a.labels).localeCompare(String(b.labels))).map(group=>[...group.labels,...keys.flatMap(k=>vs.map((v,i)=>aggregate(group.columns.get(k)?.[i]??[],v.aggregate??'sum')))])];
    if (this.Result.length * this.Result[0].length > MAX_OPERATION_CELLS) throw new RangeError('Pivot output exceeds the cell limit');
    return this.Result;
  }
  WriteTo(destination) { const result=this.Refresh(); destination.Resize(result.length,result[0].length).Values=result; return result; }
}
/** Secant/bisection goal seek with rollback on failure. It is not a general Solver. */
export function goalSeek(formulaCell, inputCell, target, { min=-1e6, max=1e6, guess, tolerance=1e-7, iterations=100 }={}) {
  if(formulaCell.Worksheet.Workbook!==inputCell.Worksheet.Workbook)throw new TypeError('Cells must share a workbook');
  if(inputCell.Formula)throw new Error('Changing cell must be an input, not a formula');
  const book=inputCell.Worksheet.Workbook,old=inputCell._record?structuredClone(inputCell._record):undefined,n=inputCell.Row*16384+inputCell.Column;
  const evaluate=x=>{inputCell.Worksheet._setRaw(n,{...old,input:x,literal:false});book.Calculation.Reset();book.Calculate();const v=formulaCell.Value;if(typeof v!=='number'||!Number.isFinite(v))throw new Error('Goal cell must evaluate to a finite number');return v-target;};
  let solved=null,steps=0;
  try {
    let a=min,b=max,fa=evaluate(a),fb=evaluate(b),x=guess??(typeof old?.input==='number'?old.input:(min+max)/2);
    for(;steps<iterations;steps++) {
      const fx=evaluate(x);if(Math.abs(fx)<=tolerance){solved=x;break;}
      if(fa*fb<=0){if(fa*fx<=0){b=x;fb=fx;}else{a=x;fa=fx;}x=(a+b)/2;}
      else {const delta=Math.max(1e-6,Math.abs(x)*1e-5),d=(evaluate(x+delta)-fx)/delta;if(!Number.isFinite(d)||Math.abs(d)<1e-14)break;x=Math.max(min,Math.min(max,x-fx/d));}
    }
  } finally {inputCell.Worksheet._setRaw(n,old);book.Calculation.Reset();book.Calculate();}
  if(solved==null)throw new Error('Goal seek did not converge; input was restored');
  inputCell.Value=solved;return{value:solved,result:formulaCell.Value,iterations:steps+1};
}
export function linearRegression(xs,ys) {
  if(xs.length!==ys.length||xs.length<2||[...xs,...ys].some(v=>!Number.isFinite(v)))throw new TypeError('Provide equal finite numeric vectors');
  const n=xs.length,mx=xs.reduce((a,b)=>a+b)/n,my=ys.reduce((a,b)=>a+b)/n;let xx=0,xy=0,yy=0;
  for(let i=0;i<n;i++){xx+=(xs[i]-mx)**2;xy+=(xs[i]-mx)*(ys[i]-my);yy+=(ys[i]-my)**2;}
  if(!xx)throw new Error('Predictors have zero variance');const slope=xy/xx,intercept=my-slope*mx;return{slope,intercept,rSquared:yy?xy*xy/(xx*yy):1,predict:x=>intercept+slope*x};
}
export function analyzeRange(range) {
  const all=range.Values.flat(),numbers=all.filter(v=>typeof v==='number'),sum=numbers.reduce((a,b)=>a+b,0);
  return{count:all.filter(v=>v!=null&&v!=='').length,numericCount:numbers.length,sum,average:numbers.length?sum/numbers.length:0,min:numbers.length?numbers.reduce((a,b)=>Math.min(a,b)):0,max:numbers.length?numbers.reduce((a,b)=>Math.max(a,b)):0,errors:all.filter(isError).length};
}
export function toDelimited(range,{delimiter=',',formulas=false,safe=true}={}) {
  const values=formulas?range._read('Input'):range.Values;
  return values.map(row=>row.map(v=>{let s=v==null?'':isError(v)?v.code:String(v);if(safe&&typeof v==='string'&&/^[=+\-@\t\r]/.test(s))s="'"+s;return s.includes(delimiter)||/["\r\n]/.test(s)?'"'+s.replace(/"/g,'""')+'"':s;}).join(delimiter)).join('\r\n');
}
export function parseDelimited(text,{delimiter=',',maxCells=MAX_OPERATION_CELLS}={}) {
  if(typeof text!=='string'||text.length>32*1024*1024)throw new RangeError('Delimited file too large');
  if(delimiter.length!==1||/["\r\n]/.test(delimiter))throw new TypeError('Invalid delimiter');
  text=text.replace(/^\uFEFF/,'');const rows=[[]];let value='',quoted=false,afterQuote=false,cells=0;
  const push=()=>{rows.at(-1).push(value);value='';afterQuote=false;if(++cells>maxCells)throw new RangeError('Delimited cell limit');};
  for(let i=0;i<text.length;i++) {
    const c=text[i];if(quoted){if(c==='"'){if(text[i+1]==='"'){value+='"';i++;}else{quoted=false;afterQuote=true;}}else value+=c;}
    else if(c==='"'&&!value&&!afterQuote)quoted=true;
    else if(c===delimiter)push();
    else if(c==='\n'||c==='\r'){push();if(c==='\r'&&text[i+1]==='\n')i++;rows.push([]);}
    else {if(afterQuote&&!/\s/.test(c))throw new Error('Unexpected character after a closing quote');if(!afterQuote)value+=c;}
  }
  if(quoted)throw new Error('Unterminated quoted field');if(value||rows.at(-1).length||!/[\r\n]$/.test(text))push();else rows.pop();
  const width=rows.reduce((m,r)=>Math.max(m,r.length),0);return rows.map(r=>Array.from({length:width},(_,i)=>r[i]??''));
}
export function importDelimited(sheet,text,{address='A1',delimiter=',',parseNumbers=true,allowFormulas=false}={}) {
  const rows=parseDelimited(text,{delimiter});if(!rows.length||!rows[0].length)return;
  const target=sheet.GetRange(address).Resize(rows.length,rows[0].length);
  sheet.Workbook.Transaction('Import delimited data',()=>rows.forEach((row,r)=>row.forEach((value,c)=>{const cell=target.GetCell(r,c);if(parseNumbers&&/^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?$/i.test(value.trim())&&Number.isFinite(Number(value)))cell.Value=Number(value);else if(allowFormulas&&value.startsWith('='))cell.Formula=value;else cell.Value=value;})));
  return target;
}
