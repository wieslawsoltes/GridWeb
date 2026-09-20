import {resolveDefinedName} from './defined-names.js';
/** Reference values retain coordinates until an explicit value consumer reads them. */
import {error, isError, number} from './errors.js';
import {MAX_OPERATION_CELLS, parseCell, MAX_COLUMNS, contains} from './address.js';
export const REFERENCE_FUNCTIONS = ['AREAS', 'ISREF', 'SHEET', 'SHEETS'];
export const THREE_D_FUNCTIONS = new Set('SUM AVERAGE AVERAGEA COUNT COUNTA MAX MAXA MIN MINA PRODUCT STDEV STDEV.S STDEVP STDEV.P STDEVA STDEVPA VAR VAR.S VARP VAR.P VARA VARPA'.split(' '));
export const areasOf = ref => ref?.type === 'multiRef' ? ref.areas : ref ? [ref] : [];
export function singleReference(ref) {
  if (!ref || ref.type === 'multiRef') throw error('#VALUE!', 'A single worksheet range is required');
  return ref;
}
function combined(areas, threeD = false) {
  if (areas.length > 256) throw error('#NUM!', 'Reference area limit');
  if (!areas.length) throw error('#NULL!', 'References do not intersect');
  return areas.length === 1 && !threeD ? areas[0] : {type: 'multiRef', areas, threeD};
}
function resolvedSheet(engine, ref, ctx) {
  const sheet = engine._sheet(ref.sheet, ctx.sheet);
  if (!sheet) throw error('#REF!', 'Unknown worksheet');
  return sheet;
}
export function advancedReference(engine, node, ctx, depth) {
  if(node.type==='table')return tableReference(engine,node,ctx);
  if(node.type==='name'&&node.sheet==null&&!ctx.vars?.has(node.name)&&!resolveDefinedName(engine.Workbook,node.name,ctx.sheet)&&engine.Workbook._sheets.some(s=>s._meta.tables.some(t=>t.name.toUpperCase()===node.name)))return tableReference(engine,{type:'table',name:node.name,selector:'[#Data]'},ctx);
  const resolve = n => engine._reference(n, ctx, depth + 1);
  if (node.type === 'ref' && node.sheetEnd != null) {
    const first = resolvedSheet(engine, node, ctx), last = resolvedSheet(engine, {sheet: node.sheetEnd}, ctx);
    const sheets = engine.Workbook._sheets, a = sheets.indexOf(first), b = sheets.indexOf(last);
    const {sheetEnd, ...bounds} = node;
    return combined(sheets.slice(Math.min(a, b), Math.max(a, b) + 1).map(s => ({...bounds, sheet: s.Name})), true);
  }
  if (node.type === 'refop') {
    const left = resolve(node.left), right = resolve(node.right);
    if(!left||!right){const v=engine._eval(!left?node.left:node.right,ctx);if(isError(v))throw v;}
    if (!left || !right || left.threeD || right.threeD) throw error('#VALUE!', 'Reference operators require non-3-D references');
    const a = areasOf(left), b = areasOf(right);
    if (node.op === ',') return combined([...a, ...b]);
    if (node.op === ':') {
      const x = singleReference(left), y = singleReference(right);
      if (resolvedSheet(engine, x, ctx) !== resolvedSheet(engine, y, ctx)) throw error('#VALUE!', 'Range endpoints must be on one worksheet');
      return {type: 'ref', sheet: x.sheet ?? y.sheet, r1: Math.min(x.r1, y.r1), c1: Math.min(x.c1, y.c1), r2: Math.max(x.r2, y.r2), c2: Math.max(x.c2, y.c2)};
    }
    if (a.length * b.length > 4096) throw error('#NUM!', 'Reference intersection work limit');
    const out = [];
    for (const x of a) for (const y of b) {
      if (resolvedSheet(engine, x, ctx) !== resolvedSheet(engine, y, ctx)) continue;
      const r1 = Math.max(x.r1, y.r1), c1 = Math.max(x.c1, y.c1), r2 = Math.min(x.r2, y.r2), c2 = Math.min(x.c2, y.c2);
      if (r1 <= r2 && c1 <= c2) out.push({type: 'ref', sheet: x.sheet ?? y.sheet, r1, c1, r2, c2});
    }
    return combined(out);
  }
  if (node.type === 'unary' && node.op === '#') {
    const base = singleReference(resolve(node.value)), sheet = resolvedSheet(engine, base, ctx);
    if (base.r1 !== base.r2 || base.c1 !== base.c2) throw error('#REF!', 'Spill reference requires an anchor cell');
    const cell = sheet.GetCell(base.r1, base.c1);
    engine.GetValue(sheet, base.r1, base.c1, ctx);
    const values = engine.arrays.get(sheet.Id + ':' + cell.Address);
    if (!values) throw error('#REF!', 'No spill');
    return {...base, r2: base.r1 + values.length - 1, c2: base.c1 + values[0].length - 1};
  }
  if (node.type === 'call' && node.name === 'INDEX') {
    const args = node.args;
    if (args.length < 2 || args.length > 4) throw error('#VALUE!', 'INDEX argument count');
    const reference = resolve(args[0]);
    if (!reference) { if (args.length === 4) throw error('#VALUE!', 'Array INDEX has no area argument'); return undefined; }
    if (reference.threeD) throw error('#VALUE!', 'INDEX does not accept 3-D references');
    const areas = areasOf(reference), sheet = resolvedSheet(engine, areas[0], ctx);
    if (areas.some(r => resolvedSheet(engine, r, ctx) !== sheet)) throw error('#VALUE!', 'INDEX areas must be on one worksheet');
    const read = (i, fallback) => args[i] == null || args[i].omitted ? fallback : Math.trunc(number(engine._eval(args[i], ctx)));
    const area = read(3, 1);
    if (area < 1 || area > areas.length) throw error('#REF!', 'INDEX area out of bounds');
    const ref = areas[area - 1], h = ref.r2 - ref.r1 + 1, w = ref.c2 - ref.c1 + 1;
    let row = read(1, 0), col = read(2, 0);
    if (args.length < 3 && h === 1) { col = row; row = 1; }
    if (row < 0 || col < 0) throw error('#VALUE!', 'Negative INDEX coordinate');
    if (row > h || col > w) throw error('#REF!', 'INDEX coordinate out of bounds');
    return {...ref, r1: row ? ref.r1 + row - 1 : ref.r1, r2: row ? ref.r1 + row - 1 : ref.r2, c1: col ? ref.c1 + col - 1 : ref.c1, c2: col ? ref.c1 + col - 1 : ref.c2};
  }
  return undefined;
}
/** Visit populated positions rather than allocating a rectangle for whole axes. */
export function readReferenceValues(engine, reference, ctx) {
  const values = [];
  for (const ref of areasOf(reference)) {
    const sheet = resolvedSheet(engine, ref, ctx);
    engine._rangeDependency(ctx, sheet, ref);
    const positions = new Set();
    for (const [n, record] of sheet._cells) if (record.input != null && contains(ref, Math.floor(n / MAX_COLUMNS), n % MAX_COLUMNS)) positions.add(n);
    if(values.length+positions.size>MAX_OPERATION_CELLS)throw error('#NUM!','Reference value limit');
    // Evaluate formula anchors first: their newly created spills must be included.
    for (const n of positions) engine.GetValue(sheet, Math.floor(n / MAX_COLUMNS), n % MAX_COLUMNS, ctx);
    for (const key of engine.spills.keys()) if (key.startsWith(sheet.Id + ':')) {
      const p = parseCell(key.slice(sheet.Id.length + 1));
      if (contains(ref, p.row, p.column)) positions.add(p.row * MAX_COLUMNS + p.column);
    }
    if (values.length + positions.size > MAX_OPERATION_CELLS) throw error('#NUM!', 'Reference value limit');
    for (const n of [...positions].sort((a, b) => a - b)) values.push(engine.GetValue(sheet, Math.floor(n / MAX_COLUMNS), n % MAX_COLUMNS, ctx));
  }
  return [values];
}
export function referenceFunction(engine, name, args, ctx) {
  if (args.length > 1 || (!args.length && ['AREAS', 'ISREF'].includes(name))) return error('#VALUE!', name + ' argument count');
  if (!args.length) {
    if (name === 'SHEETS') return engine.Workbook._sheets.length;
    const index = engine.Workbook._sheets.indexOf(ctx.sheet ?? engine.Workbook.ActiveWorksheet);
    return index < 0 ? error('#REF!', 'No current worksheet') : index + 1;
  }
  let ref;
  if (name === 'ISREF') { try { ref = engine._reference(args[0], ctx); return !!ref && areasOf(ref).every(r => !!engine._sheet(r.sheet, ctx.sheet)); } catch { return false; } }
  ref = engine._reference(args[0], ctx);
  if (name === 'SHEET' && !ref) {
    const value = engine._eval(args[0], ctx);
    if (value?.code) return value;
    const sheet = typeof value === 'string' && engine._sheet(value, ctx.sheet);
    return sheet ? engine.Workbook._sheets.indexOf(sheet) + 1 : error('#N/A', 'Unknown worksheet');
  }
  if (!ref) {const value=engine._eval(args[0],ctx);return isError(value)?value:error(name==='SHEETS'?'#REF!':'#VALUE!',name+' requires a reference');}
  const sheets = areasOf(ref).map(r => resolvedSheet(engine, r, ctx));
  if (name === 'SHEETS') return new Set(sheets).size;
  if (name === 'SHEET') return Math.min(...sheets.map(s => engine.Workbook._sheets.indexOf(s))) + 1;
  if (ref.threeD) return error('#VALUE!', 'AREAS does not accept a 3-D reference');
  return areasOf(ref).length;
}

export function tableReference(engine, node, ctx) {
  for (const sheet of engine.Workbook._sheets) {
    const table = sheet._meta.tables.find(t => t.name.toUpperCase() === node.name.toUpperCase());
    if (!table) continue;
    engine._rangeDependency(ctx, sheet, table.range);
    const headers = sheet.GetRange({...table.range, r2:table.range.r1}).Values[0];
    const raw = node.selector.slice(1,-1), leaves = raw.includes('[') ? [...raw.matchAll(/\[([^\[\]]*)\]/g)].map(m=>m[1]) : [raw];
    const special = leaves.filter(v=>v.startsWith('#')).map(v=>v.toUpperCase());
    if (special.some(v=>!['#ALL','#HEADERS','#DATA','#THIS ROW'].includes(v)) || special.length>1) throw error('#REF!','Unsupported table item selector');
    const row = raw.startsWith('@') || special.includes('#THIS ROW');
    const ref = {...table.range,type:'ref',sheet:sheet.Name,r1:table.range.r1+1};
    if(special.includes('#ALL'))ref.r1=table.range.r1;
    else if(special.includes('#HEADERS'))ref.r1=ref.r2=table.range.r1;
    if(row){if(ctx.row<=table.range.r1||ctx.row>table.range.r2)throw error('#VALUE!','Table row outside data');ref.r1=ref.r2=ctx.row;}
    const columns = leaves.filter(v=>!v.startsWith('#')).map(v=>v.replace(/^@/,''));
    if(columns.length){
      if(columns.length>2||columns.length===2&&!raw.includes(']:['))throw error('#REF!','Unsupported table column selection');
      const indexes=columns.map(name=>headers.findIndex(h=>String(h).toUpperCase()===name.toUpperCase()));
      if(indexes.some(i=>i<0))throw error('#REF!','Unknown table column');
      ref.c1=table.range.c1+Math.min(...indexes);ref.c2=table.range.c1+Math.max(...indexes);
    }
    if(ref.r1>ref.r2)throw error('#REF!','Table has no data rows');
    return ref;
  }
  throw error('#NAME?','Unknown table');
}
