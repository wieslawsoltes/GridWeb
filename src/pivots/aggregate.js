import {isError,error} from '../errors.js';
import {MAX_OPERATION_CELLS} from '../address.js';
export const aggregates=['sum','count','countNumbers','average','min','max','product','distinct','variance','variancePopulation','standardDeviation','standardDeviationPopulation'];
export const key=v=>JSON.stringify(v.map(x=>typeof x==='string'?['s',x.toLocaleUpperCase('en-US')]:isError(x)?['e',x.code]:[typeof x,x??null]));
const compare=(a,b)=>{for(let i=0;i<a.length;i++){const x=a[i],y=b[i];let n=typeof x==='number'&&typeof y==='number'?x-y:String(x??'').localeCompare(String(y??''),'en',{numeric:true,sensitivity:'base'});if(n)return n;}return 0;};
export function normalizeOptions(headers,options={}){
 const names=headers.map(x=>{if(typeof x!=='string'||!x.trim())throw new TypeError('Pivot source headers must be nonempty text');return x;});
 if(new Set(names.map(s=>s.toUpperCase())).size!==names.length)throw new TypeError('Pivot source headers must be unique');
 const field=v=>{const i=typeof v==='number'?v:names.findIndex(n=>n.toUpperCase()===String(v).toUpperCase());if(!Number.isInteger(i)||i<0||i>=names.length)throw new RangeError('Unknown pivot field: '+v);return names[i];};
 for(const flag of ['rowGrandTotals','columnGrandTotals'])if(options[flag]!=null&&typeof options[flag]!=='boolean')throw new TypeError('Pivot totals options must be boolean');
 const rows=(options.rows??[0]).map(field),columns=(options.columns??[]).map(field);
 if(rows.length>8||columns.length>4||new Set([...rows,...columns]).size!==rows.length+columns.length)throw new RangeError('Pivot axes require distinct fields (up to 8 row and 4 column fields)');
 const values=(options.values??[{column:1,aggregate:'sum'}]).map(v=>{const column=field(v.column),aggregate=v.aggregate??'sum';if(!aggregates.includes(aggregate))throw new TypeError('Unknown pivot aggregate: '+aggregate);const name=v.name??aggregate+' of '+column;if(typeof name!=='string'||!name||name.length>255)throw new TypeError('Invalid measure caption');return {column,aggregate,name};});
 if(!values.length||values.length>16||new Set(values.map(v=>v.name.toUpperCase())).size!==values.length)throw new RangeError('Use 1–16 uniquely named value fields');
 const filters=(options.filters??[]).map(f=>{if(!Array.isArray(f.values)||f.values.length>10000||f.values.some(v=>v!=null&&!['string','number','boolean'].includes(typeof v)))throw new TypeError('Pivot filter values must be scalar literals');return {column:field(f.column),values:structuredClone(f.values)};});
 if(filters.length>headers.length||new Set(filters.map(f=>f.column)).size!==filters.length)throw new RangeError('Duplicate pivot filters');
 return {rows,columns,values,filters,rowGrandTotals:options.rowGrandTotals??false,columnGrandTotals:options.columnGrandTotals??false};
}
const accumulator=()=>({count:0,n:0,sum:0,min:Infinity,max:-Infinity,product:1,mean:0,m2:0,distinct:null,error:null});
function add(a,v){if(v!=null&&v!=='')a.count++;a.distinct?.add(key([v]));if(isError(v))a.error??=v;if(typeof v!=='number')return;a.n++;a.sum+=v;a.product*=v;a.min=Math.min(a.min,v);a.max=Math.max(a.max,v);const d=v-a.mean;a.mean+=d/a.n;a.m2+=d*(v-a.mean);}
function finish(a,type){
 if(type==='count')return a.count;if(type==='countNumbers')return a.n;if(type==='distinct')return a.distinct?.size??0;if(a.error)return a.error;
 if(type==='sum')return a.sum;if(!a.n)return ['average','variance','standardDeviation'].includes(type)?error('#DIV/0!'):0;
 if(type==='average')return a.mean;if(type==='min')return a.min;if(type==='max')return a.max;if(type==='product')return a.product;
 const sample=['variance','standardDeviation'].includes(type);if(sample&&a.n<2)return error('#DIV/0!');const variance=Math.max(0,a.m2/(a.n-(sample?1:0)));return type.startsWith('standardDeviation')?Math.sqrt(variance):variance;
}
/** Streaming accumulators; no per-cell list of all input values and no unbounded output allocation. */
export function summarizePivot(matrix,options){
 if(!Array.isArray(matrix)||!matrix.length||!matrix[0].length)throw new RangeError('Pivot source requires a header row');
 const headers=matrix[0],o=normalizeOptions(headers,options),indices=list=>list.map(n=>headers.indexOf(n)),rs=indices(o.rows),cs=indices(o.columns),vs=o.values.map(v=>headers.indexOf(v.column));
 const filters=o.filters.map(f=>({index:headers.indexOf(f.column),keys:new Set(f.values.map(v=>key([v])))}));
 const rows=new Map(),cols=new Map(),buckets=new Map(),accepted=[];let budget=0;
 const get=(r,c)=>{const k=JSON.stringify([r,c]);let a=buckets.get(k);if(!a){budget+=vs.length;if(budget>MAX_OPERATION_CELLS)throw new RangeError('Pivot aggregate cell limit');a=vs.map((_,i)=>({...accumulator(),distinct:o.values[i].aggregate==='distinct'?new Set():null}));buckets.set(k,a);}return a;};
 for(let i=1;i<matrix.length;i++){
  const row=matrix[i];if(row.length!==headers.length)throw new RangeError('Pivot source must be rectangular');if(!filters.every(f=>f.keys.has(key([row[f.index]]))))continue;
  accepted.push(i);const rl=rs.map(j=>row[j]),cl=cs.map(j=>row[j]),rk=key(rl),ck=key(cl);if(!rows.has(rk))rows.set(rk,rl);if(!cols.has(ck))cols.set(ck,cl);
  for(const [r,c]of [[rk,ck],...(o.rowGrandTotals?[[rk,null]]:[]),...(o.columnGrandTotals?[[null,ck]]:[]),...(o.rowGrandTotals&&o.columnGrandTotals?[[null,null]]:[])]){const a=get(r,c);vs.forEach((j,n)=>add(a[n],row[j]));}
 }
 const rkeys=[...rows].sort((a,b)=>compare(a[1],b[1])),ckeys=[...cols].sort((a,b)=>compare(a[1],b[1]));
 if(!rs.length&&!rkeys.length)rkeys.push([key([]),[]]);
 if(!cs.length&&!ckeys.length)ckeys.push([key([]),[]]);if(o.rowGrandTotals&&cs.length)ckeys.push([null,['Grand Total']]);if(o.columnGrandTotals&&rs.length)rkeys.push([null,rs.map((_,i)=>i?'':'Grand Total')]);
 const width=rs.length+ckeys.length*vs.length;if(!width||(rkeys.length+1)*width>MAX_OPERATION_CELLS)throw new RangeError('Pivot output exceeds the cell limit');
 const result=[[...o.rows,...ckeys.flatMap(([,label])=>o.values.map(v=>[...label,v.name].join(' / ')))]];
 for(const [rk,label]of rkeys)result.push([...label,...ckeys.flatMap(([ck])=>o.values.map((v,n)=>{const a=buckets.get(JSON.stringify([rk,ck]))?.[n]??accumulator();const value=finish(a,v.aggregate);return typeof value==='number'&&!Number.isFinite(value)?error('#NUM!'):value;}))]);
 return {result,options:o,rows:rkeys,columns:ckeys,accepted,headers};
}
