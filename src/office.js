import { Workbook } from './model.js';
import {referencePrefix} from './reference-syntax.js';
import { isError } from './errors.js';
/** Errors identify the implemented compatibility contract, not an Office host/runtime. */
export class OfficeApiError extends Error {
  constructor(code, message) { super(message); this.name = 'OfficeApiError'; this.code = code; }
}
const fail = (code, message) => { throw new OfficeApiError(code, message); };
const clone = value => value == null ? value : structuredClone(value);
const output = v => isError(v) ? v.code : v == null ? '' : v;
function enumValue(value, allowed) { const lower = String(value).toLowerCase(); return allowed.includes(lower) ? lower : fail('NotSupported', 'Unsupported option: ' + value); }
function mixed(range, read) { let first, set = false, same = true; range._each(c => { const v = read(c); if (!set) { first = v; set = true; } else if (!Object.is(v, first)) same = false; }); return same ? first : null; }
class ClientObject {
  constructor(context, getter, properties) {
    this.context = context; this._getter = getter; this._properties = properties; this._loaded = new Map();
    for (const [name, spec] of Object.entries(properties)) Object.defineProperty(this, name, {
      enumerable: true, get: () => this._loaded.has(name) ? clone(this._loaded.get(name)) : fail('PropertyNotLoaded', 'Call load("' + name + '") and await context.sync()'),
      ...(spec.set ? { set: value => { const copy = clone(value); this.context._enqueue(() => spec.set(this._get(), copy)); } } : {})
    });
  }
  _get() { this.context._alive(); const value = this._getter(); if (!value) fail('ItemNotFound', 'The requested object does not exist'); return value; }
  load(selection = '*') {
    this.context._alive();
    if (selection && typeof selection === 'object' && !Array.isArray(selection)) selection = selection.select ?? Object.keys(selection).filter(k => selection[k]).join(',');
    const keys = (Array.isArray(selection) ? selection : String(selection).split(',')).flatMap(s => s.trim() === '*' ? Object.keys(this._properties) : [s.trim()]);
    for (const path of keys) {
      const [key, ...tail] = path.split('/');
      if (tail.length) { const nested = this[key]; if (!(nested instanceof ClientObject)) fail('NotSupported', 'Unknown load path ' + path); nested.load(tail.join('/')); continue; }
      const spec = this._properties[key]; if (!spec) fail('NotSupported', 'Unknown load property ' + key);
      this.context._enqueue(() => { const value = clone(spec.get(this._get())); return () => this._loaded.set(key, value); }, true);
    }
    return this;
  }
  set(values) {
    if (!values || typeof values !== 'object') fail('InvalidArgument', 'A property object is required');
    for (const [key, value] of Object.entries(values)) {
      if (this._properties[key]?.set) this[key] = value;
      else if (this[key] instanceof ClientObject) this[key].set(value);
      else fail('NotSupported', 'Property is not writable: ' + key);
    }
    return this;
  }
  toJSON() { return Object.fromEntries([...this._loaded].map(([k, v]) => [k, clone(v)])); }
}
function matrixWrite(range, values, property) {
  if (!Array.isArray(values) || values.length !== range.RowCount || values.some(row => !Array.isArray(row) || row.length !== range.ColumnCount)) fail('InvalidArgument', 'Matrix dimensions must match the range');
  range._each((cell, r, c) => {
    const value = values[r][c]; if (value === null) return;
    if (property === 'numberFormat') { if (typeof value !== 'string') fail('InvalidArgument', 'Number formats must be strings or null'); cell.Style = { numberFormat: value || 'General' }; }
    else if (property === 'formulas' && typeof value === 'string' && value.startsWith('=')) cell.Formula = value;
    else if (value === '') cell.Input = null;
    else cell.Value = value;
  });
}
class RangeFormat extends ClientObject {
  constructor(context, getter) {
    const shared = (name, fallback) => ({ get: r => mixed(r, c => c.Style[name] ?? fallback), set: (r, v) => { if (v == null) fail('InvalidArgument', 'Null format value'); r.SetStyle({ [name]: v === 'Center' && name === 'verticalAlignment' ? 'middle' : typeof v === 'string' ? v.toLowerCase() : v }); } });
    super(context, getter, { wrapText: shared('wrapText', false), horizontalAlignment: shared('horizontalAlignment', 'left'), verticalAlignment: shared('verticalAlignment', 'bottom') });
    const font = {};
    for (const [name, fallback] of Object.entries({ bold:false, italic:false, underline:false, strikethrough:false, color:'#000000', name:'Calibri', size:11 })) font[name] = {
      get: r => mixed(r, c => c.Style.font?.[name] ?? fallback),
      set: (r, v) => { if (v == null) fail('InvalidArgument', 'Null font value'); if (name === 'underline' && ![true,false,'Single','single','None','none'].includes(v)) fail('NotSupported', 'Only single underline is supported'); r.SetStyle({ font: { [name]: name === 'underline' ? v === true || String(v).toLowerCase() === 'single' : v } }); }
    };
    this.font = new ClientObject(context, getter, font);
    this.fill = new ClientObject(context, getter, { color: { get: r => mixed(r, c => c.Style.fill ?? '#ffffff'), set: (r, v) => { if (typeof v !== 'string') fail('InvalidArgument', 'Fill color must be text'); r.SetStyle({ fill: v }); } } });
    this.fill.clear = () => context._enqueue(() => getter().SetStyle({ fill: '' }));
  }
  autofitColumns() { this.context._enqueue(() => this._get().Format.AutoFitColumns()); }
  autofitRows() { fail('NotSupported', 'Automatic row-height measurement is not implemented by this adapter'); }
}
export class Range extends ClientObject {
  constructor(context, getter) {
    super(context, getter, {
      address: { get: r => r.FullAddress }, rowIndex: { get: r => r.Bounds.r1 }, columnIndex: { get: r => r.Bounds.c1 }, rowCount: { get: r => r.RowCount }, columnCount: { get: r => r.ColumnCount }, cellCount: { get: r => r.Count },
      values: { get: r => r.Values.map(row => row.map(output)), set: (r, v) => matrixWrite(r, v, 'values') },
      formulas: { get: r => r._read('Formula').map((row, i) => row.map((v, j) => v ?? output(r.GetCell(i, j).Value))), set: (r, v) => matrixWrite(r, v, 'formulas') },
      text: { get: r => r.Text }, numberFormat: { get: r => r._read('Style').map(row => row.map(s => s.numberFormat ?? 'General')), set: (r, v) => matrixWrite(r, v, 'numberFormat') }
    });
    this.format = new RangeFormat(context, getter);
  }
  get worksheet() { return new Worksheet(this.context, () => this._get().Worksheet); }
  getCell(row, column) { return new Range(this.context, () => { const r = this._get(); const c = r.GetCell(row, column); return r.Worksheet.GetRange(c.Address); }); }
  getRow(index) { return new Range(this.context, () => { const r = this._get(); r.GetCell(index, 0); return r.Worksheet.GetRangeByIndexes(r.Bounds.r1 + index, r.Bounds.c1, 1, r.ColumnCount); }); }
  getColumn(index) { return new Range(this.context, () => { const r = this._get(); r.GetCell(0, index); return r.Worksheet.GetRangeByIndexes(r.Bounds.r1, r.Bounds.c1 + index, r.RowCount, 1); }); }
  getOffsetRange(rows, columns) { return new Range(this.context, () => this._get().Offset(rows, columns)); }
  getResizedRange(deltaRows, deltaColumns) { return new Range(this.context, () => { const r = this._get(); return r.Resize(r.RowCount + deltaRows, r.ColumnCount + deltaColumns); }); }
  clear(applyTo = 'All') { const mode = enumValue(applyTo, ['all','contents','formats']); this.context._enqueue(() => this._get().Clear(mode)); }
  merge(across = false) { this.context._enqueue(() => { const r = this._get(); if (across) for (let i = 0; i < r.RowCount; i++) r.Worksheet.GetRangeByIndexes(r.Bounds.r1 + i, r.Bounds.c1, 1, r.ColumnCount).Merge(); else r.Merge(); }); }
  unmerge() { this.context._enqueue(() => this._get().Unmerge()); }
  copyFrom(source, copyType = 'All', skipBlanks = false, transpose = false) {
    if (!(source instanceof Range) || source.context !== this.context) fail('InvalidObjectPath', 'Source must belong to the same request context');
    if (typeof skipBlanks !== 'boolean' || typeof transpose !== 'boolean') fail('InvalidArgument', 'Copy flags must be boolean');
    const mode = enumValue(copyType, ['all','values','formulas','formats']); this.context._enqueue(() => {
      if (mode === 'values') this.context._book.Calculation.Reset();
      if (skipBlanks || transpose || mode === 'formats') this._get().PasteSpecial(source._get(), {mode, skipBlanks, transpose});
      else this._get().CopyFrom(source._get(), mode);
    });
  }
}
export class Worksheet extends ClientObject {
  constructor(context, getter) {
    let resolved;
    const stable = () => { resolved ??= getter(); if (!resolved || !context._book._sheets.includes(resolved)) fail('ItemNotFound', 'Worksheet was not found or was deleted'); return resolved; };
    super(context, stable, { name: { get:s=>s.Name, set:(s,v)=>s.Name=v }, id: { get:s=>s.Id }, position: { get:s=>context._book._sheets.indexOf(s), set:(s,v)=>context._book.Worksheets.Move(s,v) } });
  }
  get names() { return new NamedItemCollection(this.context,()=>this._get()); }
  getRange(address) { return new Range(this.context, () => this._get().GetRange(address)); }
  getRangeByIndexes(row, column, rowCount, columnCount) { return new Range(this.context, () => this._get().GetRangeByIndexes(row, column, rowCount, columnCount)); }
  getUsedRange() { return new Range(this.context, () => this._get().UsedRange); }
  activate() { this.context._enqueue(() => { const b = this.context._book, old = b.ActiveWorksheet, next = this._get(); b._record(() => b.ActiveWorksheet = next, () => b.ActiveWorksheet = old, { type: 'sheet-activate', sheet: next }); }); }
  delete() { this.context._enqueue(() => this.context._book.Worksheets.Remove(this._get())); }
}
class WorksheetCollection {
  constructor(context) { this.context = context; this._items = null; }
  getItem(name) { return new Worksheet(this.context, () => this.context._book.Worksheets.Get(name)); }
  getItemAt(index) { return new Worksheet(this.context, () => this.context._book.Worksheets.Get(index)); }
  getActiveWorksheet() { return new Worksheet(this.context, () => this.context._book.ActiveWorksheet); }
  add(name) { let value; this.context._enqueue(() => { value = this.context._book.Worksheets.Add(name); }); return new Worksheet(this.context, () => value); }
  get items() { return this._items ? [...this._items] : fail('PropertyNotLoaded', 'Load items and synchronize the context'); }
  load(selection = 'items') {
    const fields = Array.isArray(selection) ? selection : String(selection).split(',');
    if (fields.some(x => !['items','items/name','items/id','items/position'].includes(x.trim()))) fail('NotSupported', 'Unknown worksheet collection load');
    this.context._enqueue(() => {
      const items = this.context._book._sheets.map(s => { const item = new Worksheet(this.context, () => s); for (const field of fields.map(x => x.trim().split('/')[1]).filter(Boolean)) item._loaded.set(field, item._properties[field].get(s)); return item; });
      return () => { this._items = items; };
    }, true); return this;
  }
}
const namedFormula = value => typeof value==='string' ? value.startsWith('=')?value:'"'+value.replace(/"/g,'""')+'"' : value?.error?'='+value.error:value==null?'=""':'='+String(value).toUpperCase();
/** Supported named-item subset; all writes and reads use the same request queue. */
export class NamedItem extends ClientObject {
  constructor(context,ownerGetter,name){
    const get=()=>{const owner=ownerGetter();if(owner!==context._book&&!context._book._sheets.includes(owner))fail('ItemNotFound','Worksheet was deleted');const definition=owner.Names.GetDefinition(name);if(!definition)fail('ItemNotFound','Defined name was not found');return {owner,definition};};
    const value=({owner,definition})=>{try{return owner.Names.GetRange(definition.name).FullAddress;}catch{const v=owner.Names.Evaluate(definition.name);return Array.isArray(v)?v.map(row=>row.map(output)):output(v);}};
    const type=({owner,definition})=>{try{owner.Names.GetRange(definition.name);return 'Range';}catch{}const v=owner.Names.Evaluate(definition.name);return isError(v)?'Error':Array.isArray(v)?'Array':typeof v==='boolean'?'Boolean':typeof v==='number'?Number.isInteger(v)?'Integer':'Double':'String';};
    super(context,get,{
      name:{get:d=>d.definition.name},scope:{get:d=>d.owner===context._book?'Workbook':'Worksheet'},
      formula:{get:d=>namedFormula(d.definition.value),set:(d,v)=>{if(typeof v!=='string'||!v.startsWith('='))fail('InvalidArgument','Named formulas start with =');d.owner.Names.Update(d.definition.name,v);}},
      comment:{get:d=>d.definition.comment,set:(d,v)=>d.owner.Names.Update(d.definition.name,d.definition.value,{comment:v})},
      visible:{get:d=>!d.definition.hidden,set:(d,v)=>{if(typeof v!=='boolean')fail('InvalidArgument','Visible must be boolean');d.owner.Names.Update(d.definition.name,d.definition.value,{hidden:!v});}},
      value:{get:value},type:{get:type}
    });
  }
  getRange(){return new Range(this.context,()=>{const d=this._get();try{return d.owner.Names.GetRange(d.definition.name);}catch(e){fail('InvalidOperation',e.message);}});}
  get worksheet(){return new Worksheet(this.context,()=>{const d=this._get();if(d.owner===this.context._book)fail('InvalidOperation','Workbook name has no scoped worksheet');return d.owner;});}
  delete(){this.context._enqueue(()=>{const d=this._get();d.owner.Names.Remove(d.definition.name);});}
}
export class NamedItemCollection {
  constructor(context,ownerGetter){this.context=context;this._owner=ownerGetter;this._items=null;}
  getItem(name){return new NamedItem(this.context,this._owner,name);}
  add(name,reference,comment=''){
    if(reference instanceof Range&&reference.context!==this.context)fail('InvalidObjectPath','Range belongs to another context');
    if(!(reference instanceof Range)&&typeof reference!=='string')fail('NotSupported','Use a range proxy or invariant formula string');
    this.context._enqueue(()=>{const owner=this._owner(),formula=reference instanceof Range?'='+referencePrefix(reference._get().Worksheet.Name)+reference._get().Address.replace(/([A-Z]+)(\d+)/g,'$$$1$$$2'):reference;if(typeof formula!=='string'||!formula.startsWith('='))fail('InvalidArgument','Named formulas start with =');owner.Names.Create(name,formula,{comment});});return this.getItem(name);
  }
  get items(){return this._items?[...this._items]:fail('PropertyNotLoaded','Load items and synchronize the context');}
  load(selection='items'){
    const fields=(Array.isArray(selection)?selection:String(selection).split(',')).map(s=>s.trim());
    const allowed=['name','scope','formula','comment','visible','type','value'];
    if(fields.some(f=>f!=='items'&&!allowed.some(a=>f==='items/'+a)))fail('NotSupported','Unsupported names collection load');
    this.context._enqueue(()=>{const owner=this._owner(),items=owner.Names.Items.map(d=>{const item=new NamedItem(this.context,()=>owner,d.name);for(const f of fields.map(f=>f.split('/')[1]).filter(Boolean))item._loaded.set(f,clone(item._properties[f].get(item._get())));return item;});return ()=>{this._items=items;};},true);return this;
  }
}
export class RequestContext {
  constructor(book) { if (!(book instanceof Workbook)) throw new TypeError('Expected a GridWeb Workbook'); this._book = book; this._queue = []; this._disposed = false; this.workbook = { worksheets: new WorksheetCollection(this), names: new NamedItemCollection(this,()=>book) }; }
  _alive() { if (this._disposed) fail('InvalidObjectPath', 'Request context is disposed'); }
  _enqueue(action, read = false) { this._alive(); if (this._queue.length >= 10000) fail('InvalidArgument', 'Batch exceeds 10,000 operations'); this._queue.push({ action, read }); }
  load(object, properties) { if (object.context !== this) fail('InvalidObjectPath', 'Object belongs to another context'); return object.load(properties); }
  async sync(passThroughValue) {
    this._alive(); const queue = this._queue.splice(0), publish = []; let mutated = false;
    try {
      this._book.Transaction('Office-compatible batch', () => {
        for (const op of queue) {
          if (op.read && mutated) { this._book.Calculation.Reset(); this._book.Calculation.Calculate({ full: true }); mutated = false; }
          const result = op.action(); if (op.read && typeof result === 'function') publish.push(result); else if (!op.read) mutated = true;
        }
      });
      for (const action of publish) action(); return passThroughValue;
    } catch (e) { if (e instanceof OfficeApiError) throw e; throw new OfficeApiError(e instanceof RangeError || e instanceof TypeError ? 'InvalidArgument' : 'GeneralException', e.message); }
  }
  dispose() { this._queue.length = 0; this._disposed = true; }
}
export function createExcelApi(book) {
  if (!(book instanceof Workbook)) throw new TypeError('Expected a GridWeb Workbook');
  return Object.freeze({
    async run(callback) {
      if (typeof callback !== 'function') throw new TypeError('Excel.run requires a callback');
      const context = new RequestContext(book);
      try { const result = await callback(context); if (context._queue.length) await context.sync(); return result; } finally { context.dispose(); }
    },
    createRequestContext: () => new RequestContext(book),
    ClearApplyTo: Object.freeze({ all:'All', contents:'Contents', formats:'Formats' }),
    RangeCopyType: Object.freeze({ all:'All', values:'Values', formulas:'Formulas' }),
    HorizontalAlignment: Object.freeze({ left:'Left', center:'Center', right:'Right' }),
    VerticalAlignment: Object.freeze({ top:'Top', center:'Center', bottom:'Bottom' })
  });
}
