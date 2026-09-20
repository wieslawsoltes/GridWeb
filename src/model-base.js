import {mapFormulaReferences, referencePrefix} from './reference-syntax.js';
import {rewriteWorkbookReferences, removedSheetReferences, movedSheetReferences} from './sheet-references.js';
import {captureRange,pasteSpecial,fillSeries,specialCells} from './editing.js';
import { EventSource, ObservableObject, RelayCommand } from './events.js';
import { CalculationEngine, isFormula } from './calculation.js';
import { formatValue } from './format.js';
import { compare, criteriaPredicate } from './functions.js';
import { isError, truth, scalar, serialDate } from './errors.js';
import { MAX_ROWS, MAX_COLUMNS, MAX_OPERATION_CELLS, parseCell, parseRange, columnName, cellAddress, rangeAddress, quoteSheet, contains, intersects, boundedCells, shiftFormula, rewriteReferences, rewriteAxisReferences, renameSheetReferences } from './address.js';
const clone = value => value == null ? value : structuredClone(value);
const safeName = name => typeof name === 'string' && !!name.trim() && name.length <= 31 && !/[\[\]:*?/\\]/.test(name) && !name.startsWith("'") && !name.endsWith("'");
const makeId = () => globalThis.crypto?.randomUUID?.() ?? 's' + Date.now().toString(36) + Math.random().toString(36).slice(2);
const key = (r, c) => r * MAX_COLUMNS + c;
const cellPosition = n => ({ row: Math.floor(n / MAX_COLUMNS), column: n % MAX_COLUMNS });
const validatePrimitive = value => {
  if (value instanceof Date) return serialDate(value);
  if (value == null || typeof value === 'boolean') return value ?? null;
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  if (typeof value === 'string' && value.length <= 32767) return value;
  if (value && typeof value === 'object' && /^#(?:REF!|VALUE!|NUM!|DIV\/0!|NAME\?|N\/A|NULL!|SPILL!|CALC!|CIRC!)$/.test(value.error)) return { error: value.error };
  throw new TypeError('Cell values must be finite numbers, booleans, bounded strings, dates, null, or formula errors');
};
const cleanStyle = source => {
  const result = {};
  for (const [k, v] of Object.entries(source ?? {})) {
    if (!['font', 'fill', 'numberFormat', 'horizontalAlignment', 'verticalAlignment', 'wrapText', 'border', 'locked', 'rotation'].includes(k)) continue;
    if (k === 'font') { result.font = {}; for (const [fk, fv] of Object.entries(v ?? {})) if (['name', 'size', 'bold', 'italic', 'underline', 'color', 'strikethrough'].includes(fk)) result.font[fk] = fv; }
    else result[k] = clone(v);
  }
  return result;
};
function interval(start, end, at, count, remove) {
  if (!remove) return [start >= at ? start + count : start, end >= at ? end + count : end];
  if (end < at) return [start, end]; if (start >= at + count) return [start - count, end - count];
  const first = start < at ? start : at, last = end >= at + count ? end - count : at - 1;
  return last < first ? null : [first, last];
}
function transformRange(range, axis, at, count, remove) {
  const a = axis === 'row' ? 'r1' : 'c1', b = axis === 'row' ? 'r2' : 'c2';
  const p = interval(range[a], range[b], at, count, remove); return p ? { ...range, [a]: p[0], [b]: p[1] } : null;
}
function structuralFormula(formula, target, current, axis, at, count, remove, book) {
  return mapFormulaReferences(formula, token => {
    let ref; try { ref = parseRange(token.address); } catch { return token.raw; }
    const sheet = token.sheet ?? current;
    const parts = token.address.split(':'), fullAxis = /^\$?[A-Za-z]+$/.test(parts[0]) ? 'column' : /^\$?\d+$/.test(parts[0]) ? 'row' : null;
    if (fullAxis && fullAxis !== axis) return token.raw;
    if (token.sheetEnd != null && token.sheetEnd.toUpperCase() !== sheet.toUpperCase()) {
      const names = book._sheets.map(s => s.Name.toUpperCase()), a = names.indexOf(sheet.toUpperCase()), b = names.indexOf(token.sheetEnd.toUpperCase()), t = names.indexOf(target.toUpperCase());
      // One sheet cannot rewrite coordinates shared by every sheet in the span.
      // Reject affected edits until grouped-sheet structural transforms are supported.
      if (a >= 0 && b >= 0 && t >= Math.min(a,b) && t <= Math.max(a,b) && at <= ref[axis === 'row' ? 'r2' : 'c2']) throw new Error('Structural edit affects a 3-D reference; edit its source range explicitly first');
      return token.raw;
    }
    if (sheet.toUpperCase() !== target.toUpperCase()) return token.raw;
    const next = transformRange(ref, axis, at, count, remove);
    if (!next || next.r2 >= MAX_ROWS || next.c2 >= MAX_COLUMNS) return '#REF!';
    let body;
    if (fullAxis) {
      const indexes = axis === 'row' ? [next.r1,next.r2] : [next.c1,next.c2];
      body = indexes.map((i,n) => (parts[n].startsWith('$') ? '$' : '') + (axis === 'row' ? i+1 : columnName(i))).join(':');
    } else {
      const render = (r,c,p) => (p.absoluteColumn?'$':'') + columnName(c) + (p.absoluteRow?'$':'') + (r+1);
      body = render(next.r1,next.c1,parseCell(parts[0])) + (parts.length === 2 ? ':' + render(next.r2,next.c2,parseCell(parts[1])) : '');
    }
    return referencePrefix(token.sheet,token.sheetEnd) + body;
  });
}
export class Workbook extends ObservableObject {
  constructor(options = {}) {
    super(); this.Name = options.name ?? 'Book1'; this.Locale = options.locale ?? 'en-US'; this._sheets = []; this._names = new Map(); this._history = []; this._redo = []; this._transaction = null; this._replaying = false; this.Revision = 0; this.Changed = new EventSource(); this.Calculated = new EventSource(); this.CalculationMode = 'Automatic'; this.HistoryLimit = 100;
    this.Calculation = new CalculationEngine(this); this.Worksheets = new WorksheetCollection(this);
    this.Names = { Add: (name, value) => this.DefineName(name, value), Get: name => this._names.get(name.toUpperCase()), Remove: name => this.RemoveName(name), [Symbol.iterator]: () => this._names[Symbol.iterator]() };
    this.ActiveWorksheet = null;
    if (options.createSheet !== false) { const s = new Worksheet(this, 'Sheet1'); this._sheets.push(s); this.ActiveWorksheet = s; }
  }
  get worksheets() { return this.Worksheets; } get names() { return this.Names; }
  get CanUndo() { return this._history.length > 0; } get CanRedo() { return this._redo.length > 0; }
  Transaction(label, action) {
    if (typeof label === 'function') { action = label; label = 'Edit'; }
    if (action?.constructor?.name==='AsyncFunction')throw new TypeError('Transactions must be synchronous');
    if (this._transaction) return action();
    const tx = { label, actions: [], changes: [] }; this._transaction = tx;
    try {
      const result = action(); if (result?.then) throw new TypeError('Transactions must be synchronous');
      this._transaction = null;
      if (tx.actions.length) {
        if (!this._replaying && this.HistoryLimit > 0) { this._history.push(tx); if (this._history.length > this.HistoryLimit) this._history.shift(); this._redo.length = 0; }
        this._flush(tx.changes, label);
      }
      return result;
    } catch (e) {
      for (const a of tx.actions.toReversed()) a.undo(); this._transaction = null; this.Calculation.Reset(); if (this.CalculationMode === 'Automatic') this.Calculation.Calculate(); throw e;
    }
  }
  _record(redo, undo, change) {
    if (!this._transaction) return this.Transaction(change.type ?? 'Edit', () => this._record(redo, undo, change));
    redo(); this._transaction.actions.push({ redo, undo }); this._transaction.changes.push(change);
  }
  _flush(changes, label) {
    const cellsOnly = changes.length < 500 && changes.every(x => x.type === 'cell');
    if (cellsOnly) for (const c of changes) this.Calculation.Invalidate(c.sheet, c.row, c.column); else this.Calculation.Reset();
    if (this.CalculationMode === 'Automatic') this.Calculation.Calculate();
    let filterChanged = false; for (const sheet of this._sheets) filterChanged = sheet._applyFilter() || filterChanged;
    if (filterChanged && this.CalculationMode === 'Automatic') { this.Calculation.Reset(); this.Calculation.Calculate(); }
    this.Revision++; this.Calculated.Emit({ Workbook: this, Revision: this.Revision }); this.Changed.Emit({ Workbook: this, Revision: this.Revision, Label: label, Changes: changes });
    this.PropertyChanged.Emit({ Sender: this, PropertyName: 'Revision', NewValue: this.Revision });
  }
  Calculate(full = false) { this.Calculation.Calculate({ full }); this.Calculated.Emit({ Workbook: this, Revision: this.Revision }); }
  Undo() { const tx = this._history.pop(); if (!tx) return false; for (const a of tx.actions.toReversed()) a.undo(); this._redo.push(tx); this._flush(tx.changes, 'Undo ' + tx.label); return true; }
  Redo() { const tx = this._redo.pop(); if (!tx) return false; for (const a of tx.actions) a.redo(); this._history.push(tx); this._flush(tx.changes, 'Redo ' + tx.label); return true; }
  ClearHistory() { this._history.length = this._redo.length = 0; }
  DefineName(name, value) {
    name = String(name).toUpperCase(); if (!/^[A-Z_\\][\w.\\]*$/.test(name)) throw new TypeError('Invalid defined name');
    try { parseCell(name); throw new TypeError('A name cannot be a cell reference'); } catch (e) { if (e instanceof TypeError) throw e; }
    const old = this._names.get(name), existed = this._names.has(name); value = validatePrimitive(value);
    this._record(() => this._names.set(name, value), () => existed ? this._names.set(name, old) : this._names.delete(name), { type: 'name', name });
  }
  RemoveName(name) { name = name.toUpperCase(); if (!this._names.has(name)) return false; const old = this._names.get(name); this._record(() => this._names.delete(name), () => this._names.set(name, old), { type: 'name', name }); return true; }
  GetRange(address) { const ref = parseRange(address), sheet = ref.sheet ? this.Worksheets.Get(ref.sheet) : this.ActiveWorksheet; if (!sheet) throw new RangeError('Worksheet not found'); return sheet.GetRange(ref); }
  Find(query, { matchCase = false, wholeCell = false, formulas = false, sheet = null } = {}) {
    query = String(query); if (!matchCase) query = query.toLowerCase(); const result = [];
    for (const s of sheet ? [sheet] : this._sheets) for (const [n] of s._cells) { const { row, column } = cellPosition(n); const cell = s.GetCell(row, column); let value = formulas ? String(cell.Input ?? '') : cell.Text; if (!matchCase) value = value.toLowerCase(); if (wholeCell ? value === query : value.includes(query)) result.push({ sheet: s, row, column, address: cellAddress(row, column), value: cell.Value }); }
    return result;
  }
  Replace(query, replacement, options = {}) {
    const found = this.Find(query, { ...options, formulas: true }); this.Transaction('Replace all', () => {
      for (const hit of found) { const cell = hit.sheet.GetCell(hit.row, hit.column), s = String(cell.Input ?? ''); const pattern = new RegExp(String(query).replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), options.matchCase ? 'g' : 'gi'); cell.Input = options.wholeCell ? replacement : s.replace(pattern, () => replacement); }
    }); return found.length;
  }
  ToJSON() { return { format: 'GridWeb', version: 1, name: this.Name, locale: this.Locale, activeSheet: this.ActiveWorksheet?.Id, names: [...this._names], sheets: this._sheets.map(s => s.ToJSON()) }; }
  static FromJSON(data) {
    if (typeof data === 'string') { if (data.length > 32 * 1024 * 1024) throw new RangeError('JSON file too large'); data = JSON.parse(data); }
    if (!data || data.format !== 'GridWeb' || data.version !== 1 || !Array.isArray(data.sheets) || !data.sheets.length || data.sheets.length > 256) throw new TypeError('Unsupported workbook document');
    const book = new Workbook({ createSheet: false, name: String(data.name ?? 'Book1').slice(0, 255), locale: data.locale ?? 'en-US' }); let count = 0; const ids = new Set();
    for (const source of data.sheets) {
      if (!safeName(source.name) || book._sheets.some(s => s.Name.toUpperCase() === source.name.toUpperCase())) throw new TypeError('Invalid or duplicate sheet name');
      const s = new Worksheet(book, source.name); if (typeof source.id === 'string' && /^[\w-]{1,80}$/.test(source.id) && !ids.has(source.id)) s.Id = source.id; ids.add(s.Id);
      if (!Array.isArray(source.cells) || (count += source.cells.length) > MAX_OPERATION_CELLS) throw new RangeError('Workbook cell limit');
      for (const item of source.cells) {
        if (!Array.isArray(item) || item.length !== 2 || !Number.isSafeInteger(item[0]) || item[0] < 0 || item[0] >= MAX_ROWS * MAX_COLUMNS) throw new TypeError('Invalid stored cell');
        const record = item[1]; s._setRaw(item[0], { input: validatePrimitive(record.input), literal: !!record.literal, style: cleanStyle(record.style), ...(record.comment ? { comment: String(record.comment).slice(0, 32767) } : {}) });
      }
      s._meta = sanitizeMeta(source.meta ?? {}); s._applyFilter(); book._sheets.push(s);
    }
    for (const [name, value] of data.names ?? []) { if (!/^[A-Z_\\][\w.\\]*$/i.test(name)) throw new TypeError('Invalid name'); book._names.set(name.toUpperCase(), validatePrimitive(value)); }
    book.ActiveWorksheet = book._sheets.find(s => s.Id === data.activeSheet) ?? book._sheets[0]; book.Calculation.Reset(); book.Calculate(); let changed=false;for (const s of book._sheets) changed=s._applyFilter()||changed;if(changed){book.Calculation.Reset();book.Calculate();}return book;
  }
  Dispose() { this.Changed.Clear(); this.Calculated.Clear(); this.Calculation.Reset(); super.Dispose(); }
}
function defaultMeta() { return { rows: {}, columns: {}, merges: [], tables: [], charts: [], conditionalFormats: [], validations: [], freezeRows: 0, freezeColumns: 0, protected: false, filter: null, print: { paper: 'A4', orientation: 'landscape', scale: 1, repeatRows: 0, area: null }, tabColor: null }; }
function sanitizeMeta(source) {
  const meta = defaultMeta();
  for (const axis of ['rows', 'columns']) for (const [i, v] of Object.entries(source[axis] ?? {})) if (/^\d+$/.test(i) && +i < (axis === 'rows' ? MAX_ROWS : MAX_COLUMNS)) meta[axis][i] = { size: Math.max(1, Math.min(2000, Number(v.size) || (axis === 'rows' ? 24 : 100))), hidden: !!v.hidden, outline: Math.max(0, Math.min(8, Math.trunc(v.outline) || 0)) };
  const ref = a => { cellAddress(a.r1, a.c1); cellAddress(a.r2, a.c2); if (a.r2 < a.r1 || a.c2 < a.c1) throw new TypeError('Invalid metadata range'); return { r1:a.r1, c1:a.c1, r2:a.r2, c2:a.c2 }; };
  for (const a of source.merges ?? []) { const r = ref(a); if (meta.merges.some(m => intersects(m, r))) throw new TypeError('Overlapping merges'); meta.merges.push(r); }
  for (const t of (source.tables ?? []).slice(0, 1000)) if (/^[A-Za-z_][\w.]*$/.test(t.name)) meta.tables.push({ name:t.name, range:ref(t.range), style:t.style ?? 'green', totals:!!t.totals });
  for (const t of (source.charts ?? []).slice(0, 100)) if (['bar','column','line','pie','scatter','area'].includes(t.type)) meta.charts.push({ id:String(t.id ?? makeId()), title:String(t.title ?? 'Chart').slice(0,255), type:t.type, range:ref(t.range), row:Math.max(0, Math.min(MAX_ROWS-1, Math.trunc(t.row) || 0)), column:Math.max(0,Math.min(MAX_COLUMNS-1,Math.trunc(t.column)||0)), width:Math.max(160,Math.min(2000,+t.width||480)), height:Math.max(120,Math.min(2000,+t.height||260)) });
  for (const r of (source.conditionalFormats ?? []).slice(0, 1000)) if (['cellValue','colorScale','dataBar','formula','duplicate'].includes(r.type)) meta.conditionalFormats.push({ ...clone(r), range:ref(r.range), style:cleanStyle(r.style) });
  for (const r of (source.validations ?? []).slice(0, 1000)) if (['list','number','whole','date','textLength','custom'].includes(r.type)) meta.validations.push({ ...clone(r), range:ref(r.range) });
  meta.freezeRows = Math.max(0,Math.min(100,Math.trunc(source.freezeRows)||0)); meta.freezeColumns = Math.max(0,Math.min(100,Math.trunc(source.freezeColumns)||0)); meta.protected = !!source.protected;
  if (source.filter?.range) meta.filter = { range:ref(source.filter.range), criteria:(source.filter.criteria ?? []).filter(c => Number.isInteger(c.column) && c.column >= 0).map(c => ({ column:c.column, value:validatePrimitive(c.value) })) };
  if (source.print) meta.print = { paper:source.print.paper === 'Letter' ? 'Letter' : 'A4', orientation:source.print.orientation === 'portrait' ? 'portrait' : 'landscape', scale:Math.max(.1,Math.min(4,+source.print.scale||1)), repeatRows:Math.max(0,Math.min(100,Math.trunc(source.print.repeatRows)||0)), area:source.print.area ? ref(source.print.area) : null };
  if (typeof source.tabColor === 'string') meta.tabColor = source.tabColor.slice(0,30); return meta;
}
export class WorksheetCollection {
  constructor(book) { this.book = book; this.CollectionChanged = book.Changed; }
  get Count() { return this.book._sheets.length; } get items() { return [...this.book._sheets]; }
  Get(nameOrIndex) { return typeof nameOrIndex === 'number' ? this.book._sheets[nameOrIndex] : this.book._sheets.find(s => s.Name.toUpperCase() === String(nameOrIndex).toUpperCase() || s.Id === nameOrIndex); }
  getItem(name) { const s = this.Get(name); if (!s) throw new RangeError('Worksheet not found'); return s; } getItemAt(index) { return this.getItem(index); }
  Add(name, index = this.Count) {
    if (!Number.isInteger(index) || index < 0 || index > this.Count || this.Count >= 256) throw new RangeError('Invalid worksheet insertion index or sheet limit');
    if (!name) { let i = 1; while (this.Get('Sheet' + i)) i++; name = 'Sheet' + i; }
    if (!safeName(name) || this.Get(name)) throw new TypeError('Invalid or duplicate sheet name');
    const sheet = new Worksheet(this.book, name), active = this.book.ActiveWorksheet;
    this.book._record(() => { this.book._sheets.splice(index, 0, sheet); if (!this.book.ActiveWorksheet) this.book.ActiveWorksheet = sheet; }, () => { this.book._sheets.splice(this.book._sheets.indexOf(sheet), 1); this.book.ActiveWorksheet = active; }, { type:'sheet-add', sheet }); return sheet;
  }
  add(name, index) { return this.Add(name, index); }
  Remove(sheetOrName) {
    const sheet = sheetOrName instanceof Worksheet ? sheetOrName : this.Get(sheetOrName);
    if (!sheet || !this.book._sheets.includes(sheet) || this.Count <= 1) throw new RangeError('A workbook needs at least one attached worksheet');
    const index = this.book._sheets.indexOf(sheet), active = this.book.ActiveWorksheet;
    this.book.Transaction('Remove worksheet', () => {
      rewriteWorkbookReferences(this.book, formula => removedSheetReferences(formula, sheet, this.book._sheets));
      this.book._record(() => { this.book._sheets.splice(index, 1); if (this.book.ActiveWorksheet === sheet) this.book.ActiveWorksheet = this.book._sheets[0]; }, () => { this.book._sheets.splice(index, 0, sheet); this.book.ActiveWorksheet = active; }, {type:'sheet-remove',sheet});
    });
  }
  Move(sheetOrName, index) {
    const sheet = sheetOrName instanceof Worksheet ? sheetOrName : this.Get(sheetOrName), before = this.book._sheets.slice(), from = before.indexOf(sheet);
    if (from < 0 || !Number.isInteger(index) || index < 0 || index >= this.Count) throw new RangeError('Invalid worksheet or final index');
    if (from === index) return sheet;
    const after = before.slice(); after.splice(from, 1); after.splice(index, 0, sheet);
    this.book.Transaction('Move worksheet', () => {
      rewriteWorkbookReferences(this.book, formula => movedSheetReferences(formula, sheet, before, after));
      this.book._record(() => {this.book._sheets = after.slice();}, () => {this.book._sheets = before.slice();}, {type:'sheet-move',sheet,index});
    });
    return sheet;
  }
  [Symbol.iterator]() { return this.book._sheets[Symbol.iterator](); }
}
export class Worksheet {
  constructor(workbook, name) { this.Workbook = workbook; this.Id = makeId(); this._name = name; this._cells = new Map(); this._formulaCells = new Set(); this._meta = defaultMeta(); this._filtered = new Set(); this._used = null; }
  get Name() { return this._name; }
  set Name(value) {
    const other = this.Workbook.Worksheets.Get(value); if (!safeName(value) || (other && other !== this)) throw new TypeError('Invalid or duplicate sheet name');
    const old = this._name;
    this.Workbook.Transaction('Rename worksheet', () => {
      rewriteWorkbookReferences(this.Workbook, formula => renameSheetReferences(formula,old,value));
      this.Workbook._record(() => this._name = value, () => this._name = old, { type:'sheet-rename', sheet:this });
    });
  }
  get name() { return this.Name; } set name(v) { this.Name = v; }
  get CellCount() { return this._cells.size; } get RowCount() { return MAX_ROWS; } get ColumnCount() { return MAX_COLUMNS; }
  GetCell(rowOrAddress, column) { const p = typeof rowOrAddress === 'string' ? parseCell(rowOrAddress) : { row:rowOrAddress, column }; cellAddress(p.row,p.column); return new Cell(this,p.row,p.column); }
  getCell(r,c) { return this.GetCell(r,c); }
  GetRange(address) { const ref = typeof address === 'string' ? parseRange(address) : { ...address }; if (ref.sheet && ref.sheet.toUpperCase() !== this.Name.toUpperCase()) return this.Workbook.Worksheets.getItem(ref.sheet).GetRange({ ...ref, sheet:null }); cellAddress(ref.r1,ref.c1); cellAddress(ref.r2,ref.c2); if (ref.r2 < ref.r1 || ref.c2 < ref.c1) throw new RangeError('Invalid range'); return new CellRange(this,ref); }
  getRange(address) { return this.GetRange(address); }
  GetRangeByIndexes(row,column,rowCount,columnCount) { return this.GetRange({ r1:row,c1:column,r2:row+rowCount-1,c2:column+columnCount-1 }); }
  get UsedRange() { if (!this._used) { let r1=MAX_ROWS,c1=MAX_COLUMNS,r2=0,c2=0; const include=(a,b,c,d)=>{r1=Math.min(r1,a);c1=Math.min(c1,b);r2=Math.max(r2,c);c2=Math.max(c2,d);};for (const n of this._cells.keys()) { const p=cellPosition(n); include(p.row,p.column,p.row,p.column); }for(const merge of this._meta.merges)include(merge.r1,merge.c1,merge.r2,merge.c2);for(const[key,array]of this.Workbook.Calculation.arrays){if(!key.startsWith(this.Id+':'))continue;const p=parseCell(key.slice(this.Id.length+1));include(p.row,p.column,p.row+array.length-1,p.column+array[0].length-1);}this._used={r1:r1===MAX_ROWS?0:r1,c1:c1===MAX_COLUMNS?0:c1,r2,c2}; } return this.GetRange(this._used); }
  getUsedRange() { return this.UsedRange; }
  _setRaw(n,record) { if (!record || (record.input == null && !Object.keys(record.style??{}).length && !record.comment)) this._cells.delete(n); else this._cells.set(n,record); if (isFormula(record)) this._formulaCells.add(n); else this._formulaCells.delete(n); this._used=null; }
  _writeRecord(n,next) { const before=clone(this._cells.get(n)), after=clone(next), p=cellPosition(n); if (JSON.stringify(before)===JSON.stringify(after)) return; this.Workbook._record(() => this._setRaw(n,after), () => this._setRaw(n,before), {type:'cell',sheet:this,row:p.row,column:p.column}); }
  SetCell(address,value,options={}) { const p=typeof address==='string'?parseCell(address):address; this._setInput(p.row,p.column,value,options); return this.GetCell(p.row,p.column); }
  setCell(address,value,options) { return this.SetCell(address,value,options); }
  _setInput(row,column,value,{literal=false,validate=true}={}) {
    cellAddress(row,column); value=validatePrimitive(value); const n=key(row,column), old=this._cells.get(n);
    if (this._meta.protected && old?.style?.locked!==false) throw new Error('Worksheet is protected');
    const spill=this.Workbook.Calculation.spills.get(this.Id+':'+cellAddress(row,column)); if (spill) throw new Error('Cannot edit part of a spilled array; edit its anchor');
    if (validate) { const message=this.Validate(row,column,value); if (message) throw new Error(message); }
    const style=old?.style??{}; this._writeRecord(n,{...old,input:value,literal,style});
  }
  _setMeta(field,value) { if(this.IsProtected&&!['protected','freezeRows','freezeColumns','print'].includes(field))throw new Error('Worksheet is protected'); const before=clone(this._meta[field]),after=clone(value); this.Workbook._record(() => {this._meta[field]=clone(after);},()=>{this._meta[field]=clone(before);},{type:field,sheet:this}); }
  SetRowHeight(row,size) { this._dimension('rows',row,{size}); } SetColumnWidth(column,size) { this._dimension('columns',column,{size}); }
  HideRows(start,count=1,hidden=true) { this._dimensions('rows',start,count,{hidden}); } HideColumns(start,count=1,hidden=true) { this._dimensions('columns',start,count,{hidden}); }
  GroupRows(start,count=1,level=1) { this._dimensions('rows',start,count,{outline:level}); }
  _dimension(axis,index,patch) { this._dimensions(axis,index,1,patch); }
  _dimensions(axis,start,count,patch) {
    const max=axis==='rows'?MAX_ROWS:MAX_COLUMNS; if (!Number.isInteger(start)||!Number.isInteger(count)||start<0||count<1||start+count>max||count>100000) throw new RangeError('Invalid dimension range');
    if (patch.size!=null && (!Number.isFinite(patch.size)||patch.size<1||patch.size>2000)) throw new RangeError('Dimension size must be 1–2000 pixels');
    if (this._meta.protected) throw new Error('Worksheet is protected'); const dimensions=clone(this._meta[axis]); for(let i=start;i<start+count;i++) dimensions[i]={...dimensions[i],...patch};this._setMeta(axis,dimensions);
  }
  get FrozenRows(){return this._meta.freezeRows;} set FrozenRows(n){if(!Number.isInteger(n)||n<0||n>100)throw new RangeError('Freeze 0–100 rows');this._setMeta('freezeRows',n);}
  get FrozenColumns(){return this._meta.freezeColumns;} set FrozenColumns(n){if(!Number.isInteger(n)||n<0||n>100)throw new RangeError('Freeze 0–100 columns');this._setMeta('freezeColumns',n);}
  FreezePanes(rows=1,columns=0){this.Workbook.Transaction('Freeze panes',()=>{this.FrozenRows=rows;this.FrozenColumns=columns;});}
  Protect(){this._setMeta('protected',true);} Unprotect(){this._setMeta('protected',false);} get IsProtected(){return this._meta.protected;}
  AddTable(address,name='Table1') {
    if(!/^[A-Za-z_][\w.]*$/.test(name)||this.Workbook._sheets.some(s=>s._meta.tables.some(t=>t.name.toUpperCase()===name.toUpperCase())))throw new TypeError('Invalid or duplicate table name');
    const r=this.GetRange(address).Bounds;if(r.r1===r.r2)throw new RangeError('A table needs headers and data');if(this._meta.tables.some(t=>intersects(t.range,r)))throw new Error('Tables cannot overlap');
    const table={name,range:r,style:'green',totals:false};this.Workbook.Transaction('Add table',()=>{const headers=new Set();for(let c=r.c1;c<=r.c2;c++){const cell=this.GetCell(r.r1,c),base=cell.Text||'Column'+(c-r.c1+1);let title=base,n=2;while(headers.has(title.toUpperCase()))title=base+n++;headers.add(title.toUpperCase());cell.Value=title;}this._setMeta('tables',[...this._meta.tables,table]);});return clone(table);
  }
  AddChart(address,options={}) { const chart={id:makeId(),title:options.title??'Chart',type:options.type??'column',range:this.GetRange(address).Bounds,row:options.row??2,column:options.column??7,width:options.width??480,height:options.height??260}; if(!['bar','column','line','pie','scatter','area'].includes(chart.type))throw new TypeError('Unknown chart type');this._setMeta('charts',[...this._meta.charts,chart]);return clone(chart); }
  UpdateChart(id,patch){const charts=clone(this._meta.charts),c=charts.find(c=>c.id===id);if(!c)throw new RangeError('Chart not found');for(const k of ['title','type','row','column','width','height'])if(k in patch)c[k]=patch[k];const sanitized=sanitizeMeta({charts}).charts;if(sanitized.length!==charts.length)throw new TypeError('Invalid chart');this._setMeta('charts',sanitized);}
  RemoveChart(id){this._setMeta('charts',this._meta.charts.filter(c=>c.id!==id));}
  get Charts(){return clone(this._meta.charts);} get Tables(){return clone(this._meta.tables);} get MergedRanges(){return clone(this._meta.merges);}
  AddConditionalFormat(address,rule){const r={...clone(rule),range:this.GetRange(address).Bounds};if(!['cellValue','colorScale','dataBar','formula','duplicate'].includes(r.type))throw new TypeError('Unknown rule');r.style=cleanStyle(r.style);this._setMeta('conditionalFormats',[...this._meta.conditionalFormats,r]);return clone(r);}
  AddValidation(address,rule){const r={...clone(rule),range:this.GetRange(address).Bounds};if(!['list','number','whole','date','textLength','custom'].includes(r.type))throw new TypeError('Unknown validation');this._setMeta('validations',[...this._meta.validations,r]);}
  Validate(row,col,value){
    for(const rule of this._meta.validations){if(!contains(rule.range,row,col)||((value==null||value==='')&&rule.allowBlank!==false))continue;let valid=true;
      if(rule.type==='list')valid=(rule.values??[]).some(v=>String(v).toLowerCase()===String(value).toLowerCase());
      else if(rule.type==='custom'){const engine=this.Workbook.Calculation,override=new Map([[this.Id+':'+cellAddress(row,col),value]]);const result=engine.Evaluate(shiftFormula(rule.formula,row-rule.range.r1,col-rule.range.c1),{sheet:this,row,col,vars:new Map([['VALUE',value]]),overrides:override});valid=!isError(result)&&truth(result);}
      else{const n=rule.type==='textLength'?String(value).length:Number(value);valid=Number.isFinite(n)&&(rule.type!=='whole'||Number.isInteger(n));if(rule.min!=null)valid&&=n>=Number(rule.min);if(rule.max!=null)valid&&=n<=Number(rule.max);}
      if(!valid)return rule.message??'This value does not satisfy the data validation rule.';
    }return null;
  }
  SetFilter(address,criteria=[]){this._setMeta('filter',{range:this.GetRange(address).Bounds,criteria:clone(criteria)});}
  ClearFilter(){this._setMeta('filter',null);}
  _applyFilter(){const previous=new Set(this._filtered);this._filtered.clear();const f=this._meta.filter;if(!f)return previous.size>0;const last=Math.min(f.range.r2,this.UsedRange.Bounds.r2);if(last-f.range.r1>100000)throw new RangeError('Filter is limited to 100,000 rows');const predicates=f.criteria.map(c=>[f.range.c1+c.column,criteriaPredicate(c.value)]);for(let r=f.range.r1+1;r<=last;r++)if(!predicates.every(([c,p])=>p(this.GetCell(r,c).Value)))this._filtered.add(r);return previous.size!==this._filtered.size||[...previous].some(r=>!this._filtered.has(r));}
  InsertRows(at,count=1){this._structure('row',at,count,false);} DeleteRows(at,count=1){this._structure('row',at,count,true);} InsertColumns(at,count=1){this._structure('column',at,count,false);} DeleteColumns(at,count=1){this._structure('column',at,count,true);}
  _structure(axis,at,count,remove){
    const limit=axis==='row'?MAX_ROWS:MAX_COLUMNS;if(!Number.isInteger(at)||!Number.isInteger(count)||at<0||count<1||at+count>limit)throw new RangeError('Invalid structural edit');if(this.IsProtected)throw new Error('Worksheet is protected');
    if(!remove)for(const n of this._cells.keys()){const p=cellPosition(n);if(p[axis]>=limit-count)throw new Error('Insert would discard nonempty cells at the sheet boundary');}
    const book=this.Workbook,before=book._sheets.map(s=>({sheet:s,cells:[...s._cells].map(([n,v])=>[n,clone(v)]),meta:clone(s._meta)})),namesBefore=[...book._names];
    const after=before.map(snapshot=>{
      const target=snapshot.sheet===this,cells=[];for(const [n,record] of snapshot.cells){const p=cellPosition(n);let position=n;if(target){const v=p[axis];if(remove&&v>=at&&v<at+count)continue;if(v>=at)p[axis]+=remove?-count:count;position=key(p.row,p.column);}
        const next=clone(record);if(isFormula(next))next.input=structuralFormula(next.input,this.Name,snapshot.sheet.Name,axis,at,count,remove,book);cells.push([position,next]);}
      const meta=clone(snapshot.meta);if(target){const dim=axis==='row'?'rows':'columns',out={};for(const [i,v]of Object.entries(meta[dim])){const n=+i;if(remove&&n>=at&&n<at+count)continue;out[n>=at?n+(remove?-count:count):n]=v;}meta[dim]=out;
        meta.merges=meta.merges.map(r=>transformRange(r,axis,at,count,remove)).filter(Boolean);for(const field of ['tables','conditionalFormats','validations','charts'])meta[field]=meta[field].map(item=>({...item,range:transformRange(item.range,axis,at,count,remove)})).filter(item=>item.range);if(meta.filter)meta.filter.range=transformRange(meta.filter.range,axis,at,count,remove);if(meta.filter&&!meta.filter.range)meta.filter=null;
        if(meta.print.area)meta.print.area=transformRange(meta.print.area,axis,at,count,remove);
        for(const chart of meta.charts){const position=chart[axis];chart[axis]=position<at?position:remove?Math.max(at,position-count):Math.min(limit-1,position+count);}
        const frozen=axis==='row'?'freezeRows':'freezeColumns';if(at<meta[frozen])meta[frozen]=Math.min(100,Math.max(at,meta[frozen]+(remove?-count:count)));

      }for(const list of [meta.validations,meta.conditionalFormats])for(const rule of list)for(const prop of ['formula','formula1','formula2'])if(typeof rule[prop]==='string')rule[prop]=structuralFormula(rule[prop],this.Name,snapshot.sheet.Name,axis,at,count,remove,book);return{sheet:snapshot.sheet,cells,meta};
    });
    const namesAfter=namesBefore.map(([n,v])=>[n,typeof v==='string'&&v.startsWith('=')?structuralFormula(v,this.Name,this.Name,axis,at,count,remove,book):v]);
    const apply=(snapshots,names)=>{for(const snap of snapshots){snap.sheet._cells=new Map(snap.cells.map(([n,v])=>[n,clone(v)]));snap.sheet._formulaCells=new Set(snap.cells.filter(([,v])=>isFormula(v)).map(([n])=>n));snap.sheet._meta=clone(snap.meta);snap.sheet._used=null;}book._names=new Map(names);book.Calculation.Reset();for(const snap of snapshots)snap.sheet._applyFilter();};
    book._record(()=>apply(after,namesAfter),()=>apply(before,namesBefore),{type:remove?'delete-'+axis:'insert-'+axis,sheet:this,at,count});
  }
  ToJSON(){return{id:this.Id,name:this.Name,cells:[...this._cells].sort((a,b)=>a[0]-b[0]).map(([n,v])=>[n,clone(v)]),meta:clone(this._meta)};}
}
export class Cell {
  constructor(sheet,row,column){this.Worksheet=sheet;this.Row=row;this.Column=column;}
  get Address(){return cellAddress(this.Row,this.Column);} get _record(){return this.Worksheet._cells.get(key(this.Row,this.Column));}
  get Input(){return this._record?.input??null;} set Input(v){this.Worksheet._setInput(this.Row,this.Column,v);}
  get Value(){return this.Worksheet.Workbook.Calculation.GetValue(this.Worksheet,this.Row,this.Column);} set Value(v){this.Worksheet._setInput(this.Row,this.Column,v,{literal:typeof v==='string'});}
  get Formula(){return isFormula(this._record)?this._record.input:null;} set Formula(v){this.Input=v==null?null:String(v).startsWith('=')?v:'='+v;}
  get Text(){return formatValue(this.Value,this._record?.style?.numberFormat??'General',this.Worksheet.Workbook.Locale);}
  get Style(){return clone(this._record?.style??{});} set Style(v){if(this.Worksheet.IsProtected)throw new Error('Worksheet is protected');this.Worksheet._writeRecord(key(this.Row,this.Column),{...this._record,input:this.Input,style:cleanStyle({...this.Style,...v})});}
  get Comment(){return this._record?.comment??'';} set Comment(v){if(this.Worksheet.IsProtected)throw new Error('Worksheet is protected');this.Worksheet._writeRecord(key(this.Row,this.Column),{...this._record,input:this.Input,comment:String(v).slice(0,32767)});}
  get value(){return this.Value;} set value(v){this.Value=v;} get formula(){return this.Formula;} set formula(v){this.Formula=v;} get text(){return this.Text;}
}
export class CellRange {
  constructor(sheet,bounds){this.Worksheet=sheet;this.Bounds={r1:bounds.r1,c1:bounds.c1,r2:bounds.r2,c2:bounds.c2};this.Format=new RangeFormat(this);}
  get Address(){return rangeAddress(this.Bounds);} get FullAddress(){return quoteSheet(this.Worksheet.Name)+'!'+this.Address;} get RowCount(){return this.Bounds.r2-this.Bounds.r1+1;} get ColumnCount(){return this.Bounds.c2-this.Bounds.c1+1;} get Count(){return this.RowCount*this.ColumnCount;}
  _each(fn){boundedCells(this.Bounds);for(let r=this.Bounds.r1;r<=this.Bounds.r2;r++)for(let c=this.Bounds.c1;c<=this.Bounds.c2;c++)fn(this.Worksheet.GetCell(r,c),r-this.Bounds.r1,c-this.Bounds.c1);}
  _read(property){boundedCells(this.Bounds);return Array.from({length:this.RowCount},(_,r)=>Array.from({length:this.ColumnCount},(_,c)=>this.Worksheet.GetCell(this.Bounds.r1+r,this.Bounds.c1+c)[property]));}
  _write(property,values){if(!Array.isArray(values)||values.length!==this.RowCount||values.some(row=>!Array.isArray(row)||row.length!==this.ColumnCount))throw new RangeError('Matrix dimensions must match the range');this.Worksheet.Workbook.Transaction('Set '+property,()=>this._each((cell,r,c)=>cell[property]=values[r][c]));}
  get Values(){return this._read('Value');}set Values(v){this._write('Value',v);} get Formulas(){return this._read('Formula');}set Formulas(v){this._write('Formula',v);}get Text(){return this._read('Text');}
  get Value(){return this.Worksheet.GetCell(this.Bounds.r1,this.Bounds.c1).Value;}set Value(v){this.Worksheet.Workbook.Transaction('Set value',()=>this._each(c=>c.Value=v));}
  get Formula(){return this.Worksheet.GetCell(this.Bounds.r1,this.Bounds.c1).Formula;}set Formula(v){this.Worksheet.Workbook.Transaction('Set formula',()=>this._each((c,r,col)=>c.Formula=v==null?null:shiftFormula(String(v).startsWith('=')?v:'='+v,r,col)));}
  get values(){return this.Values;}set values(v){this.Values=v;}get formulas(){return this.Formulas;}set formulas(v){this.Formulas=v;}get format(){return this.Format;}
  GetCell(r,c){if(r<0||c<0||r>=this.RowCount||c>=this.ColumnCount)throw new RangeError('Range-relative cell out of bounds');return this.Worksheet.GetCell(this.Bounds.r1+r,this.Bounds.c1+c);}
  Offset(rows,columns){return this.Worksheet.GetRange({r1:this.Bounds.r1+rows,c1:this.Bounds.c1+columns,r2:this.Bounds.r2+rows,c2:this.Bounds.c2+columns});}
  Resize(rows,columns){return this.Worksheet.GetRangeByIndexes(this.Bounds.r1,this.Bounds.c1,rows,columns);}
  Clear(mode='all'){
    if(!['all','contents','formats'].includes(mode))throw new TypeError('Unknown clear mode');if(this.Worksheet.IsProtected)throw new Error('Worksheet is protected');if(mode!=='formats')for(const [address,spill]of this.Worksheet.Workbook.Calculation.spills){if(!address.startsWith(this.Worksheet.Id+':'))continue;const cell=parseCell(address.slice(this.Worksheet.Id.length+1)),anchor=parseCell(spill.key.slice(this.Worksheet.Id.length+1));if(contains(this.Bounds,cell.row,cell.column)&&!contains(this.Bounds,anchor.row,anchor.column))throw new Error('Cannot clear part of a spilled array');}this.Worksheet.Workbook.Transaction('Clear '+mode,()=>{
      // Sparse clear supports entire rows/columns without visiting every blank cell.
      for(const [n,record]of [...this.Worksheet._cells]){const p=cellPosition(n);if(!contains(this.Bounds,p.row,p.column))continue;const next=mode==='formats'?{...record,style:{}}:mode==='contents'?{...record,input:null}:null;this.Worksheet._writeRecord(n,next);}
    });
  }
  SetStyle(patch){this.Worksheet.Workbook.Transaction('Format range',()=>this._each(c=>c.Style={...c.Style,...patch,font:patch.font?{...c.Style.font,...patch.font}:c.Style.font}));return this;}
  Merge(){if(this.Worksheet.IsProtected)throw new Error('Worksheet is protected');if(this.Worksheet._meta.merges.some(m=>intersects(m,this.Bounds)))throw new Error('Merge overlaps an existing merge');this.Worksheet.Workbook.Transaction('Merge cells',()=>{this._each((c,r,col)=>{if(r||col)c.Input=null;});this.Worksheet._setMeta('merges',[...this.Worksheet._meta.merges,this.Bounds]);});}
  Unmerge(){this.Worksheet._setMeta('merges',this.Worksheet._meta.merges.filter(m=>!intersects(m,this.Bounds)));}
  CopyFrom(source,mode='all'){
    if(!['all','values','formulas','formats'].includes(mode))throw new TypeError('Unknown copy mode');
    if(mode==='formats'){this.PasteSpecial(source,{mode:'formats'});return;}
    if(!(source instanceof CellRange))throw new TypeError('Expected source range');boundedCells(this.Bounds);boundedCells(source.Bounds);
    const copiedValues=mode==='values'?source.Values:null;const records=[];source._each((cell,r,c)=>{records[r]??=[];records[r][c]=clone(cell._record??{input:null,style:{}});});
    this.Worksheet.Workbook.Transaction('Copy cells',()=>this._each((cell,r,c)=>{
      const sr=r%source.RowCount,sc=c%source.ColumnCount,record=clone(records[sr][sc]),dr=cell.Row-(source.Bounds.r1+sr),dc=cell.Column-(source.Bounds.c1+sc);
      if(mode==='values'){const value=copiedValues[sr][sc];cell.Value=isError(value)?{error:value.code}:value;return;}
      if(isFormula(record))record.input=shiftFormula(record.input,dr,dc);if(mode==='formulas'){cell.Input=record.input;return;}
      this.Worksheet._setInput(cell.Row,cell.Column,record.input,{literal:record.literal});cell.Style=record.style??{};cell.Comment=record.comment??'';
    }));
  }
  Capture(){return captureRange(this);}
  PasteSpecial(source,options){return pasteSpecial(this,source,options);}
  FillSeries(options){return fillSeries(this,options);}
  SpecialCells(type,options){return specialCells(this,type,options);}
  FillDown(){this.CopyFrom(this.Worksheet.GetRange({...this.Bounds,r2:this.Bounds.r1}));} FillRight(){this.CopyFrom(this.Worksheet.GetRange({...this.Bounds,c2:this.Bounds.c1}));}
  AutoFill(destination,{series=true}={}){
    destination=destination instanceof CellRange?destination:this.Worksheet.GetRange(destination);const values=this.Values.flat(),numeric=series&&values.length>=2&&values.every(v=>typeof v==='number')&&(this.RowCount===1||this.ColumnCount===1);
    if(!numeric){destination.CopyFrom(this);return;}const step=values[1]-values[0],offset=this.RowCount>1?destination.Bounds.r1-this.Bounds.r1:destination.Bounds.c1-this.Bounds.c1;this.Worksheet.Workbook.Transaction('Fill series',()=>{destination._each((c,r,col)=>c.Value=values[0]+step*(offset+(this.RowCount>1?r:col)));destination.SetStyle(this.GetCell(0,0).Style);});
  }
  Sort(keys=[{column:0,ascending:true}],{hasHeaders=true}={}){
    if(this.Worksheet.IsProtected)throw new Error('Worksheet is protected');if(this.Worksheet._meta.merges.some(m=>intersects(m,this.Bounds)))throw new Error('Unmerge cells before sorting');boundedCells(this.Bounds);
    const start=this.Bounds.r1+(hasHeaders?1:0),rows=[];for(let r=start;r<=this.Bounds.r2;r++)rows.push({row:r,values:Array.from({length:this.ColumnCount},(_,c)=>this.Worksheet.GetCell(r,this.Bounds.c1+c).Value),records:Array.from({length:this.ColumnCount},(_,c)=>clone(this.Worksheet._cells.get(key(r,this.Bounds.c1+c))))});
    for(const k of keys)if(k.column<0||k.column>=this.ColumnCount)throw new RangeError('Sort key out of range');
    rows.sort((a,b)=>{for(const k of keys){let n;try{n=compare(a.values[k.column],b.values[k.column]);}catch{n=String(a.values[k.column]).localeCompare(String(b.values[k.column]));}if(n)return n*(k.ascending===false?-1:1);}return a.row-b.row;});
    this.Worksheet.Workbook.Transaction('Sort rows',()=>rows.forEach((item,i)=>item.records.forEach((record,c)=>{if(isFormula(record))record.input=shiftFormula(record.input,start+i-item.row,0);this.Worksheet._writeRecord(key(start+i,this.Bounds.c1+c),record);})));this.Worksheet._applyFilter();
  }
  RemoveDuplicates(columns=Array.from({length:this.ColumnCount},(_,i)=>i),{hasHeaders=true}={}){
    if(this.Worksheet.IsProtected)throw new Error('Worksheet is protected');if(this.Worksheet.MergedRanges.some(m=>intersects(m,this.Bounds)))throw new Error('Unmerge before removing duplicates');
    const a=this.Values,start=hasHeaders?1:0,seen=new Set(),indices=Array.from({length:start},(_,i)=>i),records=a.map((row,r)=>row.map((_,c)=>clone(this.GetCell(r,c)._record)));
    for(let i=start;i<a.length;i++){const row=a[i],k=JSON.stringify(columns.map(c=>typeof row[c]==='string'?row[c].toUpperCase():row[c]));if(!seen.has(k)){seen.add(k);indices.push(i);}}
    this.Worksheet.Workbook.Transaction('Remove duplicates',()=>{for(let r=0;r<a.length;r++)for(let c=0;c<this.ColumnCount;c++){const source=indices[r],record=source==null?undefined:records[source][c];if(isFormula(record))record.input=shiftFormula(record.input,r-source,0);this.Worksheet._writeRecord(key(this.Bounds.r1+r,this.Bounds.c1+c),record);}});return a.length-indices.length;
  }
}
export class RangeFormat {
  constructor(range){this.range=range;const object={};for(const [p,key]of Object.entries({Bold:'bold',Italic:'italic',Underline:'underline',Color:'color',Size:'size',Name:'name',Strikethrough:'strikethrough'}))for(const alias of [p,key])Object.defineProperty(object,alias,{get:()=>range.GetCell(0,0).Style.font?.[key],set:v=>range.SetStyle({font:{[key]:v}})});this.Font=object;}
  get font(){return this.Font;}get Fill(){return this.range.GetCell(0,0).Style.fill;}set Fill(v){this.range.SetStyle({fill:v});}get fill(){return this.Fill;}set fill(v){this.Fill=v;}
  get NumberFormat(){return this.range.GetCell(0,0).Style.numberFormat??'General';}set NumberFormat(v){this.range.SetStyle({numberFormat:String(v)});}get numberFormat(){return this.NumberFormat;}set numberFormat(v){this.NumberFormat=v;}
  set WrapText(v){this.range.SetStyle({wrapText:!!v});}set HorizontalAlignment(v){this.range.SetStyle({horizontalAlignment:v});}set VerticalAlignment(v){this.range.SetStyle({verticalAlignment:v});}set Borders(v){this.range.SetStyle({border:v});}
  set ColumnWidth(value){const r=this.range;this.range.Worksheet.Workbook.Transaction('Column widths',()=>{for(let c=r.Bounds.c1;c<=r.Bounds.c2;c++)r.Worksheet.SetColumnWidth(c,value);});}
  set RowHeight(value){const r=this.range;this.range.Worksheet.Workbook.Transaction('Row heights',()=>{for(let row=r.Bounds.r1;row<=r.Bounds.r2;row++)r.Worksheet.SetRowHeight(row,value);});}
  AutoFitColumns(){const r=this.range;boundedCells(r.Bounds);r.Worksheet.Workbook.Transaction('Autofit',()=>{for(let c=r.Bounds.c1;c<=r.Bounds.c2;c++){let width=64;for(let row=r.Bounds.r1;row<=r.Bounds.r2;row++)width=Math.max(width,r.Worksheet.GetCell(row,c).Text.length*7.3+16);r.Worksheet.SetColumnWidth(c,Math.min(420,width));}});}
}
export class GridViewModel extends ObservableObject {
  constructor(workbook=new Workbook()){super();this.Workbook=workbook;this.Selection='A1';this.Undo=new RelayCommand(()=>workbook.Undo(),()=>workbook.CanUndo);this.Redo=new RelayCommand(()=>workbook.Redo(),()=>workbook.CanRedo);this._sub=workbook.Changed.Subscribe(()=>{this.Undo.NotifyCanExecuteChanged();this.Redo.NotifyCanExecuteChanged();this.PropertyChanged.Emit({Sender:this,PropertyName:'Workbook',NewValue:workbook});});}
  Dispose(){this._sub.Dispose();this.Undo.Dispose();this.Redo.Dispose();super.Dispose();}
}
