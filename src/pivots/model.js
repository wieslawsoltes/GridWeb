import {parseRange,rangeAddress,cellAddress,intersects,contains,boundedCells,MAX_ROWS,MAX_COLUMNS} from '../address.js';
import {isError} from '../errors.js';
import {summarizePivot,normalizeOptions,key} from './aggregate.js';
const clone=v=>structuredClone(v),state=new WeakMap();
export function pivotId(){if(globalThis.crypto.randomUUID)return globalThis.crypto.randomUUID();const bytes=globalThis.crypto.getRandomValues(new Uint8Array(16));bytes[6]=(bytes[6]&15)|64;bytes[8]=(bytes[8]&63)|128;const hex=Array.from(bytes,b=>b.toString(16).padStart(2,'0')).join('');return hex.slice(0,8)+'-'+hex.slice(8,12)+'-'+hex.slice(12,16)+'-'+hex.slice(16,20)+'-'+hex.slice(20);}
const point=(book,range)=>{if(!range?.Worksheet||range.Worksheet.Workbook!==book)throw new TypeError('Pivot ranges must belong to the same workbook');return {sheet:range.Worksheet.Id,address:range.Address};};
const value=v=>isError(v)?{error:v.code}:v;
// Advisory stale-state fingerprint; correctness never skips a refresh based on this hash.
function stamp(matrix){const text=JSON.stringify(matrix);let a=2166136261,b=5381;for(let i=0;i<text.length;i++){a=Math.imul(a^text.charCodeAt(i),16777619);b=Math.imul(b,33)^text.charCodeAt(i);}return text.length+':'+(a>>>0)+':'+(b>>>0);}
const nameValid=name=>{if(typeof name!=='string'||!/^[A-Za-z_][\w.]{0,254}$/.test(name))throw new TypeError('Invalid pivot name');};
function sheetFor(book,id){const s=book._sheets.find(s=>s.Id===id);if(!s)throw new RangeError('Pivot worksheet no longer exists');return s;}
function ref(book,descriptor){return sheetFor(book,descriptor.sheet).GetRange(descriptor.address);}
function bounds(reference){const b=parseRange(reference);boundedCells(b);return {r1:b.r1,c1:b.c1,r2:b.r2,c2:b.c2};}
function outputBounds(destination,matrix){const b=bounds(destination.address),r2=b.r1+matrix.length-1,c2=b.c1+matrix[0].length-1;cellAddress(r2,c2);return {r1:b.r1,c1:b.c1,r2,c2};}
/** Managed definitions and cached worksheet output. Refreshes are explicit, synchronous and undoable. */
export class PivotReport {
 constructor(collection,id){this.Collection=collection;this.Id=id;}
 get _record(){const r=this.Collection._items.find(r=>r.id===this.Id);if(!r)throw new RangeError('Pivot report was removed');return r;}
 get Name(){return this._record.name;}
 get Source(){return ref(this.Collection.Workbook,this._record.source);}
 get Destination(){return ref(this.Collection.Workbook,this._record.destination);}
 get Options(){return clone(this._record.options);}
 get OutputRange(){const r=this._record;return r.lastBounds?sheetFor(this.Collection.Workbook,r.destination.sheet).GetRange(r.lastBounds):null;}
 get IsStale(){const r=this._record;if(r.invalidReason)return true;return r.sourceStamp!==stamp(this.Source.Values);}
 get InvalidReason(){return this._record.invalidReason??null;}
 _compute(){if(this.InvalidReason)throw new Error(this.InvalidReason);return summarizePivot(this.Source.Values,this.Options);}
 Refresh(){const computed=this._compute();this.Collection._write(this,computed);return computed.result.map(r=>r.map(value));}
 Update({source,destination,options,name}={}){
  const collection=this.Collection,b=collection.Workbook,old=this._record,next=clone(old);
  if(source)next.source=point(b,source);if(destination){next.destination=point(b,destination);next.destination.address=cellAddress(destination.Bounds.r1,destination.Bounds.c1);}
  if(name!=null){nameValid(name);if(collection._items.some(r=>r.id!==old.id&&r.name.toUpperCase()===name.toUpperCase()))throw new Error('Duplicate pivot name');next.name=name;}
  const input=ref(b,next.source);next.options=normalizeOptions(input.Values[0],{...old.options,...options});delete next.invalidReason;
  b.Transaction('Reconfigure pivot '+next.name,()=>{if(destination&&old.lastBounds)ref(b,old.destination).Worksheet.GetRange(old.lastBounds).Clear('contents');if(destination)next.lastBounds=null;collection._set(collection._items.map(r=>r.id===old.id?next:r));this.Refresh();});return this;
 }
 SetFilter(field,values){const f=this.Options.filters.filter(f=>f.column.toUpperCase()!==String(field).toUpperCase());if(values!==null)f.push({column:field,values});return this.Update({options:{filters:f}});}
 DrillDown(row,column,{limit=10000}={}){
  if(!Number.isInteger(limit)||limit<1||limit>100000)throw new RangeError('Drill-down limit must be 1–100,000 rows');
  const c=this._compute(),rs=c.options.rows.length,ms=c.options.values.length;
  if(!Number.isInteger(row)||row<1||row>=c.result.length||!Number.isInteger(column)||column<rs||column>=c.result[0].length)throw new RangeError('Choose a pivot value cell, using zero-based output indexes');
  const rk=c.rows[row-1][0],ck=c.columns[Math.floor((column-rs)/ms)][0],matrix=this.Source.Values,ri=c.options.rows.map(n=>c.headers.indexOf(n)),ci=c.options.columns.map(n=>c.headers.indexOf(n));const result=[c.headers];
  for(const i of c.accepted){const r=matrix[i];if(rk!==null&&key(ri.map(j=>r[j]))!==rk||ck!==null&&key(ci.map(j=>r[j]))!==ck)continue;if(result.length>limit)throw new RangeError('Drill-down result exceeds limit');result.push(r.map(value));}
  return clone(result);
 }
 ToJSON(){return clone(this._record);}
}
export class PivotTableCollection {
 constructor(book){this.Workbook=book;this._items=[];}
 get Count(){return this._items.length;}
 Get(nameOrIndex){const r=typeof nameOrIndex==='number'?this._items[nameOrIndex]:this._items.find(r=>r.name.toUpperCase()===String(nameOrIndex).toUpperCase()||r.id===nameOrIndex);return r?new PivotReport(this,r.id):null;}
 [Symbol.iterator](){return this._items.map(r=>new PivotReport(this,r.id))[Symbol.iterator]();}
 _set(items){const before=clone(this._items),after=clone(items);this.Workbook._record(()=>this._items=clone(after),()=>this._items=clone(before),{type:'pivot-definitions'});}
 Add(name,source,destination,options={}){
  nameValid(name);if(this.Get(name))throw new Error('Duplicate pivot name');if(this.Count>=100)throw new RangeError('At most 100 managed pivots per workbook');
  const input=point(this.Workbook,source),target=point(this.Workbook,destination);target.address=cellAddress(destination.Bounds.r1,destination.Bounds.c1);
  const record={id:pivotId(),name,source:input,destination:target,options:normalizeOptions(source.Values[0],options),lastBounds:null},report=new PivotReport(this,record.id);
  this.Workbook.Transaction('Create pivot '+name,()=>{this._set([...this._items,record]);report.Refresh();});return report;
 }
 Remove(name,{clear=false}={}){const report=this.Get(name);if(!report)return false;this.Workbook.Transaction('Remove pivot '+report.Name,()=>{if(clear)report.OutputRange?.Clear('contents');this._set(this._items.filter(r=>r.id!==report.Id));});return true;}
 RefreshAll(){this.Workbook.Transaction('Refresh all pivots',()=>{for(const report of this)report.Refresh();});}
 _write(report,computed){
  const record=report._record,target=ref(this.Workbook,record.destination),next=outputBounds(record.destination,computed.result),source=ref(this.Workbook,record.source);
  if(target.Worksheet===source.Worksheet&&intersects(next,source.Bounds))throw new Error('Pivot output cannot overwrite its source');
  for(const other of this._items)if(other.id!==record.id&&other.destination.sheet===record.destination.sheet&&other.lastBounds&&intersects(next,other.lastBounds))throw new Error('Pivot output overlaps another report');
  if(target.Worksheet.MergedRanges.some(r=>intersects(r,next))||target.Worksheet.Tables.some(t=>intersects(t.range,next)))throw new Error('Pivot output overlaps a table or merged range');
  const range=target.Worksheet.GetRange(next);for(let r=next.r1;r<=next.r2;r++)for(let c=next.c1;c<=next.c2;c++)if(!record.lastBounds||!contains(record.lastBounds,r,c)){const cell=target.Worksheet.GetCell(r,c);if(cell.Input!=null||cell.Value!=null)throw new Error('Pivot growth would overwrite occupied cells');}
  this.Workbook.Transaction('Refresh pivot '+record.name,()=>{
   if(record.lastBounds)target.Worksheet.GetRange(record.lastBounds).Clear('contents');range.Values=computed.result.map(r=>r.map(value));
   const updated={...record,lastBounds:next,sourceStamp:stamp(source.Values)};this._set(this._items.map(r=>r.id===record.id?updated:r));
  });
 }
 ToJSON(){return clone(this._items);}
 _load(records){
  if(!Array.isArray(records)||records.length>100)throw new TypeError('Invalid pivot definitions');const ids=new Set(),names=new Set();
  const validated=records.map(r=>{nameValid(r.name);if(typeof r.id!=='string'||!/^[-\w]{1,80}$/.test(r.id)||ids.has(r.id)||names.has(r.name.toUpperCase()))throw new TypeError('Duplicate or invalid pivot identifier');ids.add(r.id);names.add(r.name.toUpperCase());
   for(const d of [r.source,r.destination]){if(typeof d?.sheet!=='string'||!/^[-\w]{1,80}$/.test(d.sheet))throw new TypeError('Invalid pivot worksheet identifier');bounds(d.address);}
   const next={id:r.id,name:r.name,source:clone(r.source),destination:clone(r.destination),options:clone(r.options),lastBounds:r.lastBounds?bounds(rangeAddress(r.lastBounds)):null};
   if(r.invalidReason)next.invalidReason=String(r.invalidReason).slice(0,255);
   else next.options=normalizeOptions(ref(this.Workbook,next.source).Values[0],next.options);
   // Cached source stamps are advisory only; fresh output must be checked after loading.
   return next;
  });this._items=validated;
 }
}
function transform(b,axis,at,count,remove){
 const a=axis==='row'?'r1':'c1',z=axis==='row'?'r2':'c2',start=b[a],end=b[z];let first=start,last=end;
 if(!remove){if(start>=at)first+=count;if(end>=at)last+=count;}
 else if(end>=at){first=start>=at+count?start-count:start<at?start:at;last=end>=at+count?end-count:at-1;}
 if(first>last||last>=(axis==='row'?'MAX_ROWS':'MAX_COLUMNS'))return null;return {...b,[a]:first,[z]:last};
}
/** Decorate the one shared Workbook constructor, rather than creating incompatible subclass identities. */
export function installPivotModel(Workbook){
 Object.defineProperty(Workbook.prototype,'PivotTables',{get(){let collection=state.get(this);if(!collection){collection=new PivotTableCollection(this);state.set(this,collection);}return collection;}});
 const toJSON=Workbook.prototype.ToJSON,fromJSON=Workbook.FromJSON,record=Workbook.prototype._record;
 Workbook.prototype.ToJSON=function(){const json=toJSON.call(this),p=state.get(this);if(p?.Count)json.pivotTables=p.ToJSON().map(({sourceStamp,...r})=>r);return json;};
 Workbook.FromJSON=function(data){if(typeof data==='string'){if(data.length>32*1024*1024)throw new RangeError('JSON file too large');data=JSON.parse(data);}const b=fromJSON.call(this,data);try{if(data.pivotTables)b.PivotTables._load(data.pivotTables);return b;}catch(e){b.Dispose();throw e;}};
 Workbook.prototype._record=function(redo,undo,change){
  if(!this._transaction)return this.Transaction(change.type??'Edit',()=>this._record(redo,undo,change));
  const pivots=state.get(this),m=/^(insert|delete)-(row|column)$/.exec(change.type);
  if(pivots?.Count&&(m||change.type==='sheet-remove')){
   const before=pivots.ToJSON(),after=clone(before);
   for(const p of after)for(const field of ['source','destination'])if(p[field].sheet===change.sheet.Id){
    if(!m){p.invalidReason='Pivot worksheet was removed; reconfigure the report';continue;}
    const next=transform(bounds(p[field].address),m[2],change.at,change.count,m[1]==='delete');
    if(!next)p.invalidReason='Structural edit removed the pivot reference; reconfigure the report';else p[field].address=rangeAddress(next);
    if(field==='destination'&&p.lastBounds){const transformed=transform(p.lastBounds,m[2],change.at,change.count,m[1]==='delete');if(transformed)p.lastBounds=transformed;else{p.lastBounds=null;p.invalidReason='Structural edit removed pivot output; reconfigure the report';}}
   }
   return record.call(this,()=>{redo();pivots._items=clone(after);},()=>{undo();pivots._items=clone(before);},change);
  }
  return record.call(this,redo,undo,change);
 };
}
