/** Workbook and worksheet-scoped names. The legacy value/iterator API is retained. */
import {parseCell, shiftFormula} from './address.js';
import {parseFormula, tokenize} from './parser.js';
import {error, serialDate} from './errors.js';
import {referencePrefix} from './reference-syntax.js';

export const MAX_DEFINED_NAMES = 10000;
const copy = value => structuredClone(value);
const keyOf = name => String(name).toUpperCase();
export function validateDefinedName(name) {
  if (typeof name !== 'string' || name.length > 255 || !/^[A-Za-z_\\][\w.\\]*$/.test(name) || /^(?:R|C|TRUE|FALSE|R\d*C\d*)$/i.test(name)) throw new TypeError('Invalid defined name');
  // These tokens are reserved by the A1 parser, even beyond the worksheet edge.
  if (/^[A-Za-z]{1,3}[1-9]\d*$/.test(name)) throw new TypeError('A name cannot look like an A1 or R1C1 reference');
  if (/^_(?:xlfn|xlws|xlpm|xlnm)\./i.test(name)) throw new TypeError('Reserved Excel name prefix');
  return keyOf(name);
}
function primitive(value) {
  if (value instanceof Date) value = serialDate(value);
  if (value == null || typeof value === 'boolean') return value ?? null;
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  if (typeof value === 'string' && value.length <= 32767) return value;
  if (value && typeof value === 'object' && /^#(?:REF!|VALUE!|NUM!|DIV\/0!|NAME\?|N\/A|NULL!|SPILL!|CALC!|CIRC!)$/.test(value.error)) return {error:value.error};
  throw new TypeError('Invalid defined-name value');
}
export function nameOwners(book) { return [book, ...book._sheets]; }
const bookOf = owner => owner.Workbook ?? owner;
const attached = owner => { const book = bookOf(owner); if (owner !== book && !book._sheets.includes(owner)) throw new Error('Worksheet is no longer attached'); return book; };
export function nameDefinitions(book) { return nameOwners(book).flatMap(owner => owner.Names.Items); }
function info(owner, name, options = {}, previous = {}) {
  if (!options || typeof options !== 'object' || Array.isArray(options) || Object.keys(options).some(k => !['comment','hidden','contextSheetId','baseAddress'].includes(k))) throw new TypeError('Invalid name options');
  const book = bookOf(owner), result = {...previous, name};
  if ('contextSheetId' in options && options.contextSheetId!=null && (typeof options.contextSheetId!=='string' || !book._sheets.some(s=>s.Id===options.contextSheetId)&&options.contextSheetId!==previous.contextSheetId || owner!==book&&options.contextSheetId!==owner.Id)) throw new TypeError('Invalid name context worksheet');
  if ('comment' in options && (typeof options.comment !== 'string' || options.comment.length > 255)) throw new TypeError('Name comments are limited to 255 characters');
  if ('hidden' in options && typeof options.hidden !== 'boolean') throw new TypeError('Hidden must be boolean');
  result.comment = options.comment ?? previous.comment ?? '';
  result.hidden = options.hidden ?? previous.hidden ?? false;
  result.contextSheetId = owner === book ? ('contextSheetId' in options ? options.contextSheetId : 'contextSheetId' in previous ? previous.contextSheetId : book.ActiveWorksheet?.Id ?? null) : owner.Id;
  if (result.contextSheetId != null && result.contextSheetId!==previous.contextSheetId && !book._sheets.some(s => s.Id === result.contextSheetId)) throw new TypeError('Name context worksheet does not exist');
  result.baseAddress = 'baseAddress' in options ? options.baseAddress : previous.baseAddress ?? null;
  if (result.baseAddress != null) { if (typeof result.baseAddress !== 'string') throw new TypeError('Invalid name origin'); parseCell(result.baseAddress); }
  return result;
}
export function nameContext(book, definition, ctx) {
  const ownerSheet = definition.owner === book ? null : definition.owner;
  const sheet = ownerSheet ?? (definition.info?.contextSheetId!=null ? book._sheets.find(s => s.Id === definition.info.contextSheetId) ?? null : ctx.sheet);
  const value = definition.value, base = definition.info?.baseAddress;
  return {ctx:{...ctx, sheet, vars:new Map(), omitted:new Set()}, value:typeof value === 'string' && value.startsWith('=') && base ? shiftFormula(value, ctx.row-parseCell(base).row, ctx.col-parseCell(base).column) : value};
}
/** Locals shadow globals; qualified names must resolve on that worksheet. */
export function resolveDefinedName(book, name, sheet, qualifier = null) {
  name = keyOf(name);
  if (qualifier != null) {
    sheet = book._sheets.find(s => keyOf(s.Name) === keyOf(qualifier));
    if (!sheet) throw error('#REF!', 'Unknown worksheet for defined name');
    if (!sheet._names.has(name)) return null;
    return {owner:sheet, value:sheet._names.get(name), info:sheet._nameInfo.get(name), key:sheet.Id+':'+name};
  }
  const owner = sheet?._names?.has(name) ? sheet : book._names.has(name) ? book : null;
  return owner ? {owner, value:owner._names.get(name), info:owner._nameInfo.get(name), key:(owner===book?'workbook':owner.Id)+':'+name} : null;
}
export class DefinedNameCollection {
  constructor(owner) { this.owner = owner; }
  get Count() { return this.owner._names.size; }
  get Items() {
    const owner=this.owner, book=bookOf(owner);
    return [...owner._names].map(([key,value])=>({...copy(owner._nameInfo.get(key) ?? {name:key,comment:'',hidden:false,contextSheetId:null,baseAddress:null}),value:copy(value),sheetId:owner===book?null:owner.Id,scope:owner===book?'Workbook':owner.Name}));
  }
  Get(name) { return copy(this.owner._names.get(keyOf(name))); }
  GetDefinition(name) { const key=keyOf(name),owner=this.owner,book=bookOf(owner);if(!owner._names.has(key))return undefined;return {...copy(owner._nameInfo.get(key)??{name:key,comment:'',hidden:false,contextSheetId:null,baseAddress:null}),value:copy(owner._names.get(key)),sheetId:owner===book?null:owner.Id,scope:owner===book?'Workbook':owner.Name}; }
  Has(name) { return this.owner._names.has(keyOf(name)); }
  _context(name,context={}) {
    const owner=this.owner,book=attached(owner),key=keyOf(name);
    if(!owner._names.has(key))throw new RangeError('Defined name does not exist');
    const ctx={sheet:book.ActiveWorksheet,row:0,col:0,vars:new Map(),depth:0,...context};
    if(ctx.sheet&&!book._sheets.includes(ctx.sheet))throw new TypeError('Foreign name evaluation worksheet');
    if(!Number.isInteger(ctx.row)||!Number.isInteger(ctx.col)||ctx.row<0||ctx.row>=1048576||ctx.col<0||ctx.col>=16384)throw new RangeError('Invalid name evaluation origin');
    return nameContext(book,{owner,value:owner._names.get(key),info:owner._nameInfo.get(key)},ctx);
  }
  Evaluate(name,context={}) {
    const book=attached(this.owner),resolved=this._context(name,context),v=resolved.value;
    const result=typeof v==='string'&&v.startsWith('=')?book.Calculation.Evaluate(v,resolved.ctx):v?.error?error(v.error):copy(v);
    return result?.type==='lambda'?error('#CALC!','A named LAMBDA must be invoked'):result;
  }
  GetRange(name,context={}) {
    const book=attached(this.owner),resolved=this._context(name,context);
    if(typeof resolved.value!=='string'||!resolved.value.startsWith('='))throw new RangeError('Name does not refer to a range');
    const ref=book.Calculation._reference(parseFormula(resolved.value),resolved.ctx);
    if(!ref||ref.type==='multiRef')throw new RangeError('Name does not refer to one rectangular range');
    const target=book.Calculation._sheet(ref.sheet,resolved.ctx.sheet);if(!target)throw new RangeError('Name references a missing worksheet');
    return target.GetRange({...ref,sheet:null});
  }

  /** Add retains the previous DefineName upsert contract. Create rejects duplicates. */
  Add(name,value,options={}) {
    const owner=this.owner, book=attached(owner), key=validateDefinedName(name), existed=owner._names.has(key);
    if (!existed && nameOwners(book).reduce((sum,o)=>sum+o._names.size,0)>=MAX_DEFINED_NAMES) throw new RangeError('Defined-name limit');
    if (book._sheets.some(s=>s._meta.tables.some(t=>keyOf(t.name)===key))) throw new TypeError('A defined name cannot conflict with a table');
    const nextValue=primitive(value), oldValue=copy(owner._names.get(key)), oldInfo=copy(owner._nameInfo.get(key)), nextInfo=info(owner,name,options,oldInfo);
    if (existed && JSON.stringify(oldValue)===JSON.stringify(nextValue) && JSON.stringify(oldInfo)===JSON.stringify(nextInfo)) return;
    book._record(()=>{owner._names.set(key,copy(nextValue));owner._nameInfo.set(key,copy(nextInfo));},()=>{if(existed){owner._names.set(key,copy(oldValue));owner._nameInfo.set(key,copy(oldInfo));}else{owner._names.delete(key);owner._nameInfo.delete(key);}}, {type:'name',name:key,sheet:owner===book?undefined:owner});
  }
  Create(name,value,options={}) { if(this.Has(name))throw new TypeError('A name already exists in this scope');this.Add(name,value,options); }
  /** Generate absolute named ranges from boundary labels in one atomic operation. */
  CreateFromSelection(source,options={topRow:true}) {
    const owner=this.owner,book=attached(owner),sheet=source?.Worksheet;
    if(!sheet||sheet.Workbook!==book||!book._sheets.includes(sheet))throw new TypeError('Source range belongs to another or detached workbook');
    if(!options||typeof options!=='object'||Array.isArray(options)||Object.entries(options).some(([k,v])=>!['topRow','bottomRow','leftColumn','rightColumn','overwrite'].includes(k)||typeof v!=='boolean'))throw new TypeError('Invalid Create from Selection options');
    const range=sheet.GetRange(source.Bounds),b=range.Bounds,{topRow=false,bottomRow=false,leftColumn=false,rightColumn=false,overwrite=false}=options;
    if(!topRow&&!bottomRow&&!leftColumn&&!rightColumn)throw new TypeError('Choose at least one label boundary');
    const data={r1:b.r1+Number(topRow),r2:b.r2-Number(bottomRow),c1:b.c1+Number(leftColumn),c2:b.c2-Number(rightColumn)};
    if(data.r1>data.r2||data.c1>data.c2)throw new RangeError('Labels leave no data cells');
    if((Number(topRow)+Number(bottomRow))*(data.c2-data.c1+1)+(Number(leftColumn)+Number(rightColumn))*(data.r2-data.r1+1)>MAX_DEFINED_NAMES)throw new RangeError('Generated-name limit');
    const generated=[],seen=new Set(),add=(row,col,bounds)=>{
      const cell=sheet.GetCell(row,col);if(cell.Value?.code)throw new TypeError('A label contains a formula error');
      const label=cell.Text.trim();if(!label)return;
      let name=label.replace(/[^A-Za-z0-9_.\\]/g,'_');if(!/^[A-Za-z_\\]/.test(name))name='_'+name;name=name.slice(0,255);
      try{validateDefinedName(name);}catch{name=('_'+name).slice(0,255);validateDefinedName(name);}
      const key=keyOf(name);if(seen.has(key))throw new TypeError('Duplicate generated name: '+name);seen.add(key);
      if(!overwrite&&owner.Names.Has(name))throw new TypeError('A name already exists in this scope: '+name);
      if(book._sheets.some(s=>s._meta.tables.some(t=>keyOf(t.name)===key)))throw new TypeError('Generated name conflicts with a table: '+name);
      const address=sheet.GetRange(bounds).Address.replace(/([A-Z]+)(\d+)/g,'$$$1$$$2');generated.push({name,value:'='+referencePrefix(sheet.Name)+address});
    };
    if(topRow||bottomRow)for(let c=data.c1;c<=data.c2;c++){const bounds={...data,c1:c,c2:c};if(topRow)add(b.r1,c,bounds);if(bottomRow)add(b.r2,c,bounds);}
    if(leftColumn||rightColumn)for(let row=data.r1;row<=data.r2;row++){const bounds={...data,r1:row,r2:row};if(leftColumn)add(row,b.c1,bounds);if(rightColumn)add(row,b.c2,bounds);}
    const count=nameOwners(book).reduce((sum,o)=>sum+o._names.size,0)+generated.filter(d=>!owner.Names.Has(d.name)).length;if(count>MAX_DEFINED_NAMES)throw new RangeError('Defined-name limit');
    book.Transaction('Create names from selection',()=>{for(const item of generated)this.Add(item.name,item.value);});
    return generated.map(d=>this.GetDefinition(d.name));
  }
  Remove(name) {
    const owner=this.owner, book=attached(owner), key=keyOf(name);if(!owner._names.has(key))return false;
    const value=copy(owner._names.get(key)), metadata=copy(owner._nameInfo.get(key));
    book._record(()=>{owner._names.delete(key);owner._nameInfo.delete(key);},()=>{owner._names.set(key,copy(value));owner._nameInfo.set(key,copy(metadata));},{type:'name',name:key,sheet:owner===book?undefined:owner});return true;
  }
  Update(name,value,options={}) { if(!this.Has(name))throw new RangeError('Defined name does not exist');this.Add(this.GetDefinition(name).name,value,options); }
  Rename(name,newName) {
    const owner=this.owner,book=attached(owner),old=validateDefinedName(name),next=validateDefinedName(newName),definition=this.GetDefinition(name);
    if(!definition)throw new RangeError('Defined name does not exist');
    if(old!==next && this.Has(newName))throw new TypeError('A name already exists in this scope');
    if(book._sheets.some(s=>s._meta.tables.some(t=>keyOf(t.name)===next)))throw new TypeError('A defined name cannot conflict with a table');
    if(old===next){const {value,comment,hidden,contextSheetId,baseAddress}=definition;this.Add(newName,value,{comment,hidden,contextSheetId,baseAddress});return;}
    // Parse/prepare all affected formulas before any mutation. Invalid preexisting
    // formula syntax is retained; a rename cannot promise to repair invalid syntax.
    const edits=[];
    for(const sheet of book._sheets) {
      for(const n of sheet._formulaCells){const record=sheet._cells.get(n),input=renameFormulaName(record.input,book,sheet,owner,old,newName);if(input!==record.input)edits.push(()=>sheet._writeRecord(n,{...record,input}));}
      for(const field of ['validations','conditionalFormats']){
        const before=copy(sheet._meta[field]),after=before.map(rule=>{const r={...rule};for(const p of ['formula','formula1','formula2'])if(typeof r[p]==='string')r[p]=renameFormulaName(r[p],book,sheet,owner,old,newName);return r;});
        if(JSON.stringify(before)!==JSON.stringify(after))edits.push(()=>book._record(()=>{sheet._meta[field]=copy(after);},()=>{sheet._meta[field]=copy(before);},{type:field,sheet}));
      }
    }
    for(const target of nameOwners(book)) for(const item of target.Names.Items) {
      if(typeof item.value!=='string'||!item.value.startsWith('='))continue;
      const context=target===book?(item.contextSheetId==null?book.ActiveWorksheet:book._sheets.find(s=>s.Id===item.contextSheetId)??null):target;
      const value=renameFormulaName(item.value,book,context,owner,old,newName);
      if(value!==item.value)edits.push(()=>target.Names.Update(item.name,value));
    }
    book.Transaction('Rename defined name',()=>{
      for(const edit of edits)edit();
      const value=copy(owner._names.get(old)),metadata=copy(owner._nameInfo.get(old));
      book._record(()=>{owner._names.delete(old);owner._nameInfo.delete(old);owner._names.set(next,copy(value));owner._nameInfo.set(next,{...copy(metadata),name:newName});},()=>{owner._names.delete(next);owner._nameInfo.delete(next);owner._names.set(old,copy(value));owner._nameInfo.set(old,copy(metadata));},{type:'name',name:next,sheet:owner===book?undefined:owner});
    });
  }
  [Symbol.iterator]() { return this.Items.map(item=>[keyOf(item.name),item.value])[Symbol.iterator](); }
}
/** Restore stored maps without history. Old v1 documents without metadata still load. */
export function loadDefinedNames(owner, values = [], metadata = []) {
  if(!Array.isArray(values)||!Array.isArray(metadata)||values.length>MAX_DEFINED_NAMES||metadata.length>MAX_DEFINED_NAMES)throw new TypeError('Invalid stored names');
  const names=new Map(),infos=new Map(),book=bookOf(owner),seen=new Set();
  for(const entry of values){if(!Array.isArray(entry)||entry.length!==2)throw new TypeError('Invalid stored name');const k=validateDefinedName(entry[0]);if(names.has(k))throw new TypeError('Duplicate stored name');names.set(k,primitive(entry[1]));infos.set(k,{name:entry[0],comment:'',hidden:false,contextSheetId:owner===book?null:owner.Id,baseAddress:null});}
  for(const item of metadata){if(!item||typeof item!=='object')throw new TypeError('Invalid stored name metadata');const k=validateDefinedName(item.name);if(!names.has(k)||seen.has(k))throw new TypeError('Orphan or duplicate name metadata');seen.add(k);const {name,comment,hidden,contextSheetId,baseAddress,...extra}=item;if(Object.keys(extra).length)throw new TypeError('Unknown name metadata');
    // Removed definition contexts remain tombstones; they must not silently bind
    // to a new sheet that happens to reuse the old name.
    const context=contextSheetId==null?null:String(contextSheetId);if(context!=null&&!/^[\w-]{1,80}$/.test(context))throw new TypeError('Invalid name context ID');
    const checked=info(owner,name,{...(comment!==undefined?{comment}:{}),...(hidden!==undefined?{hidden}:{}),...(baseAddress!==undefined?{baseAddress}:{}),contextSheetId:owner===book?book.ActiveWorksheet?.Id:owner.Id});
    checked.contextSheetId=owner===book?context:owner.Id;infos.set(k,checked);
  }
  owner._names=names;owner._nameInfo=infos;
}
export function mapQualifiedNames(formula, transform) {
  let tokens;try{tokens=tokenize(String(formula));}catch{return formula;}
  const offset=String(formula).startsWith('=')?1:0,edits=tokens.filter(t=>t.t==='name'&&t.sheet!=null).map(t=>({start:t.start+offset,end:t.end+offset,value:transform(t)}));
  let result=String(formula);for(const e of edits.reverse())result=result.slice(0,e.start)+e.value+result.slice(e.end);return result;
}
function renameFormulaName(formula,book,sheet,owner,old,replacement){
  let ast;try{ast=parseFormula(formula);}catch{return formula;}
  const edits=[],builtins=new Set(book.Calculation.FunctionNames);
  const visit=(node,bound=new Set())=>{
    if(!node)return;
    if(node.type==='name'||node.type==='call'){
      if(owner!==book&&sheet===owner&&node.sheet==null&&node.name===keyOf(replacement)&&!bound.has(node.name)&&!(node.type==='call'&&builtins.has(node.name))&&resolveDefinedName(book,node.name,sheet)?.owner===book)throw new Error('Rename would shadow an existing workbook name use');
      if(node.name===old && (node.sheet!=null||!bound.has(old)) && !(node.type==='call'&&node.sheet==null&&builtins.has(old))){
        let def;try{def=resolveDefinedName(book,old,sheet,node.sheet);}catch{}
        if(def?.owner===owner){
          if(node.sheet==null&&bound.has(keyOf(replacement)))throw new Error('Rename would capture a LET or LAMBDA variable');
          if(node.type==='call'&&node.sheet==null&&builtins.has(keyOf(replacement)))throw new Error('Rename would call a built-in function instead');
          if(owner===book&&node.sheet==null&&sheet?._names.has(keyOf(replacement)))throw new Error('Rename would be shadowed by a worksheet name');
          edits.push({start:node.start,end:node.end,value:(node.sheet!=null?referencePrefix(node.sheet):'')+replacement});
        }
      }
      if(node.type==='call'&&node.sheet==null&&node.name==='LET'){
        const nested=new Set(bound);for(let i=0;i<node.args.length-1;i+=2){visit(node.args[i+1],nested);if(node.args[i]?.type==='name')nested.add(node.args[i].name);}visit(node.args.at(-1),nested);return;
      }
      if(node.type==='call'&&node.sheet==null&&node.name==='LAMBDA'){const nested=new Set([...bound,...node.args.slice(0,-1).filter(p=>p.type==='name').map(p=>p.name)]);visit(node.args.at(-1),nested);return;}
    }
    for(const prop of ['left','right','callee'])if(node[prop])visit(node[prop],bound);
    if(node.type==='unary')visit(node.value,bound);
    for(const arg of node.args??[])visit(arg,bound);for(const row of node.rows??[])for(const arg of row)visit(arg,bound);
  };
  visit(ast);let result=formula,offset=formula.startsWith('=')?1:0;
  for(const e of edits.sort((a,b)=>b.start-a.start))result=result.slice(0,e.start+offset)+e.value+result.slice(e.end+offset);return result;
}

export function namedRange(book,address,sheet) {
  const node=parseFormula(address);if(node.type!=='name')throw new TypeError('Expected a range address or defined name');
  const ctx={sheet,row:0,col:0,depth:0,vars:new Map()},ref=book.Calculation._reference(node,ctx);
  if(!ref||ref.type==='multiRef')throw new RangeError('Name does not refer to one rectangular range');
  const target=book.Calculation._sheet(ref.sheet,sheet);if(!target)throw new RangeError('Name references a missing worksheet');return target.GetRange({...ref,sheet:null});
}
