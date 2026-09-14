import { Workbook } from '../model.js';
import { MAX_OPERATION_CELLS } from '../address.js';

const copy = value => structuredClone(value);
export function canonical(value) {
  if (Array.isArray(value)) return '[' + value.map(canonical).join(',') + ']';
  if (value && typeof value === 'object') return '{' + Object.keys(value).sort().filter(k => value[k] !== undefined).map(k => JSON.stringify(k) + ':' + canonical(value[k])).join(',') + '}';
  return JSON.stringify(value);
}
export const equal = (a, b) => canonical(a) === canonical(b);
export class CollaborationError extends Error {
  constructor(code, message, details = null) { super(message); this.name = 'CollaborationError'; this.code = code; this.details = details; }
}
const fail = (code, message) => { throw new CollaborationError(code, message); };
/** Validate through the same core without executing workbook-provided JavaScript. */
export function validateDocument(document) {
  if (canonical(document)?.length > 16 * 1024 * 1024) fail('LIMIT', 'Shared workbooks are limited to 16 MiB');
  if (!Array.isArray(document?.sheets) || !document.sheets.length) fail('INVALID', 'A workbook is required');
  const ids = new Set();
  for (const s of document.sheets) {
    if (typeof s.id !== 'string' || !/^[\w-]{1,80}$/.test(s.id) || ids.has(s.id)) fail('INVALID', 'Stable unique worksheet IDs are required');
    ids.add(s.id);
    if (!Array.isArray(s.cells) || s.cells.length > MAX_OPERATION_CELLS || new Set(s.cells.map(c => c[0])).size !== s.cells.length) fail('INVALID', 'Invalid or duplicate cells');
  }
  const book = Workbook.FromJSON(document);
  try { return book.ToJSON(); } finally { book.Dispose(); }
}
const content = r => ({ input: r?.input ?? null, literal: !!r?.literal });
const fieldValue = (r, field) => field === 'content' ? content(r) : field === 'style' ? r?.style ?? {} : r?.comment ?? '';
const structure = d => ({ name:d.name, locale:d.locale, names:d.names, pivotTables:d.pivotTables??[], sheets:d.sheets.map(s => ({id:s.id,name:s.name,meta:s.meta})) });
/** Content and literal flag are one atomic field; formatting and comments merge independently. */
export function diffDocuments(base, local, { exclusive = false } = {}) {
  if (exclusive || !equal(structure(base), structure(local))) return { kind:'replace', document:copy(local) };
  const changes=[];
  for(let i=0;i<base.sheets.length;i++) {
    const before=new Map(base.sheets[i].cells), after=new Map(local.sheets[i].cells), sheet=base.sheets[i].id;
    for(const cell of new Set([...before.keys(),...after.keys()])) for(const field of ['content','style','comment']) {
      const expected=fieldValue(before.get(cell),field),value=fieldValue(after.get(cell),field);
      if(!equal(expected,value)) changes.push({sheet,cell,field,expected:copy(expected),value:copy(value)});
    }
  }
  if(changes.length>10000) return {kind:'replace',document:copy(local)};
  return {kind:'cells',changes};
}
/** Apply a compare-and-swap patch to an isolated copy. Conflicts never partially mutate a room. */
export function applyChanges(document, changes, { resolve = 'reject' } = {}) {
  if(!['reject','local','remote'].includes(resolve))fail('INVALID','Unknown conflict strategy');
  if(!Array.isArray(changes)||changes.length>10000) fail('LIMIT','At most 10,000 field changes per commit');
  const result=copy(document), maps=new Map(result.sheets.map(s=>[s.id,new Map(s.cells)])), seen=new Set(), conflicts=[];
  for(const op of changes) {
    if(!op||!['content','style','comment'].includes(op.field)||!Number.isSafeInteger(op.cell)||op.cell<0||op.cell>=1048576*16384||!maps.has(op.sheet)) fail('INVALID','Invalid cell operation');
    const id=JSON.stringify([op.sheet,op.cell,op.field]);if(seen.has(id)) fail('INVALID','A field may be changed only once per commit');seen.add(id);
    const map=maps.get(op.sheet),record=map.get(op.cell),actual=fieldValue(record,op.field);
    if(!equal(actual,op.expected)&&!equal(actual,op.value)) {
      conflicts.push({...copy(op),actual:copy(actual)});
      if(resolve!=='local')continue;
    }
    const next={...record,input:record?.input??null,style:record?.style??{}};
    if(op.field==='content') {
      if(!op.value||typeof op.value!=='object'||typeof op.value.literal!=='boolean'||!Object.hasOwn(op.value,'input')) fail('INVALID','Invalid content field');
      next.input=copy(op.value.input);next.literal=op.value.literal;
    } else if(op.field==='style') {if(!op.value||typeof op.value!=='object'||Array.isArray(op.value))fail('INVALID','Invalid style');next.style=copy(op.value);}
    else {if(typeof op.value!=='string')fail('INVALID','Invalid comment');next.comment=op.value;}
    if(next.input==null&&!Object.keys(next.style).length&&!next.comment)map.delete(op.cell);else map.set(op.cell,next);
  }
  if(conflicts.length&&resolve==='reject')throw new CollaborationError('CONFLICT','Concurrent edits conflict',conflicts);
  for(const sheet of result.sheets)sheet.cells=[...maps.get(sheet.id)].sort((a,b)=>a[0]-b[0]);
  return {document:validateDocument(result),conflicts};
}
/** Rebase unsent edits onto a newer server document while retaining disjoint changes. */
export function rebaseDocuments(base, local, remote, options={}) {
  const delta=diffDocuments(base,local,options);
  if(delta.kind==='replace') {
    if(equal(structure(base),structure(remote))&&equal(base.sheets,remote.sheets))return {document:copy(local),conflicts:[]};
    if(options.resolve==='local')return {document:copy(local),conflicts:[]};
    if(options.resolve==='remote')return {document:copy(remote),conflicts:[]};
    throw new CollaborationError('CONFLICT','Workbook structure changed concurrently',[{field:'workbook'}]);
  }
  if(!delta.changes.length)return {document:copy(remote),conflicts:[]};
  return applyChanges(remote,delta.changes,options);
}
/** Keep Workbook and existing Worksheet identities stable for controls and view models. */
export function applyDocument(book, document) {
  if(book._transaction)fail('BUSY','Cannot synchronize within an active local transaction');
  const validated=Workbook.FromJSON(document),previous=new Map(book._sheets.map(s=>[s.Id,s])),active=book.ActiveWorksheet?.Id;
  try {
    book._sheets=validated._sheets.map(source=>{const target=previous.get(source.Id)??source;target.Workbook=book;target._name=source.Name;target._cells=source._cells;target._formulaCells=new Set(source._formulaCells);target._used=null;target._filtered=new Set(source._filtered);target._meta=source._meta;return target;});
    book.Name=validated.Name;book.Locale=validated.Locale;book._names=new Map(validated._names);book.PivotTables._load(validated.PivotTables.ToJSON());
    book.ActiveWorksheet=book._sheets.find(s=>s.Id===active)??book._sheets[0];
    // Inverse history closures can refer to records changed remotely. Never replay stale inverses.
    book.ClearHistory();book._flush([{type:'remote-sync'}],'Remote synchronization');
  } finally {validated.Dispose();}
}
