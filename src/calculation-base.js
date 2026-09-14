import { parseFormula } from './parser.js';
import { createFunctionRegistry, compare } from './functions.js';
import { error, isError, scalar, number, text, truth, matrix, flatten } from './errors.js';
import { cellAddress, parseRange, parseCell, contains, MAX_ROWS, MAX_COLUMNS, MAX_OPERATION_CELLS } from './address.js';
const keyOf = (sheet, r, c) => sheet.Id + ':' + cellAddress(r, c);
const isFormula = record => !record?.literal && typeof record?.input === 'string' && record.input.startsWith('=');
export { isFormula };
export function literalValue(record) {
  const v = record?.input ?? null;
  if (v && typeof v === 'object' && typeof v.error === 'string') return error(v.error);
  if (typeof v !== 'string' || record?.literal) return v;
  if (v.startsWith("'")) return v.slice(1);
  if (/^(true|false)$/i.test(v)) return /^true$/i.test(v);
  if (/^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?$/i.test(v.trim()) && Number.isFinite(Number(v))) return Number(v);
  if (/^[+-]?(?:\d+\.?\d*|\.\d+)%$/.test(v.trim())) return Number(v.trim().slice(0, -1)) / 100;
  return v;
}
function vectorize(a, b, fn) {
  if (!Array.isArray(a) && !Array.isArray(b)) { try { return fn(a, b); } catch (e) { return isError(e) ? e : error('#VALUE!', e.message); } }
  a = matrix(a); b = matrix(b); const h = Math.max(a.length, b.length), w = Math.max(a[0].length, b[0].length);
  if ((a.length !== h && a.length !== 1) || (b.length !== h && b.length !== 1) || (a[0].length !== w && a[0].length !== 1) || (b[0].length !== w && b[0].length !== 1) || h * w > MAX_OPERATION_CELLS) return error('#VALUE!');
  return Array.from({ length: h }, (_, r) => Array.from({ length: w }, (_, c) => vectorize(a[a.length === 1 ? 0 : r][a[0].length === 1 ? 0 : c], b[b.length === 1 ? 0 : r][b[0].length === 1 ? 0 : c], fn)));
}
export class CalculationEngine {
  constructor(workbook) {
    this.Workbook = workbook; this.Functions = createFunctionRegistry(); this.cache = new Map(); this.ast = new Map(); this.dependencies = new Map(); this.dependents = new Map(); this.ranges = new Map(); this.dirty = new Set(); this.evaluating = new Set(); this.arrays = new Map(); this.spills = new Map(); this.EvaluationCount = 0;
  }
  RegisterFunction(name, fn) { if (!/^[A-Za-z_][\w.]*$/.test(name) || typeof fn !== 'function') throw new TypeError('Invalid function'); this.Functions.set(name.toUpperCase(), fn); this.Reset(); }
  get FunctionNames() { return [...new Set([...this.Functions.keys(), 'IF', 'IFS', 'IFERROR', 'IFNA', 'SWITCH', 'LET', 'LAMBDA', 'MAP', 'REDUCE', 'SCAN', 'BYROW', 'BYCOL', 'ROW', 'COLUMN', 'ADDRESS', 'INDIRECT', 'OFFSET', 'SUBTOTAL', 'FORMULATEXT', 'ISFORMULA'])].sort(); }
  Reset() {
    this.cache.clear(); this.dependencies.clear(); this.dependents.clear(); this.ranges.clear(); this.spills.clear(); this.arrays.clear(); this.dirty.clear();
    for(const s of this.Workbook._sheets)s._used=null;
    for (const s of this.Workbook._sheets) for (const n of s._formulaCells) this.dirty.add(keyOf(s, Math.floor(n / MAX_COLUMNS), n % MAX_COLUMNS));
  }
  Invalidate(sheet, row, col) {
    sheet._used=null;
    const queue = [keyOf(sheet, row, col)], seen = new Set();
    for (const [key, refs] of this.ranges) if (refs.some(r => r.sheet === sheet.Id && contains(r, row, col))) queue.push(key);
    while (queue.length) {
      const key = queue.pop(); if (seen.has(key)) continue; seen.add(key);
      this.cache.delete(key); this.dirty.add(key);
      for (const dependent of this.dependents.get(key) ?? []) queue.push(dependent);
      const spill = this.arrays.get(key);
      if (spill) {
        for (const [cell, owner] of this.spills) if (owner.key === key) { this.spills.delete(cell); for (const dependent of this.dependents.get(cell) ?? []) queue.push(dependent); }
        this.arrays.delete(key);
      }
    }
  }
  _dependency(context, sheet, r, c) {
    if (!context?.key) return; const key = keyOf(sheet, r, c);
    if (!this.dependencies.has(context.key)) this.dependencies.set(context.key, new Set()); this.dependencies.get(context.key).add(key);
    if (!this.dependents.has(key)) this.dependents.set(key, new Set()); this.dependents.get(key).add(context.key);
  }
  _rangeDependency(context, sheet, ref) { if (!context?.key) return; if (!this.ranges.has(context.key)) this.ranges.set(context.key, []); this.ranges.get(context.key).push({ ...ref, sheet: sheet.Id }); }
  _sheet(name, fallback) { return name ? this.Workbook._sheets.find(s => s.Name.toUpperCase() === name.toUpperCase()) : fallback; }
  GetValue(sheet, row, col, context = null) {
    const tx=this.Workbook._transaction;if(tx&&!this._syncingTransaction&&(tx.processedChanges??0)<tx.changes.length){this._syncingTransaction=true;try{for(const change of tx.changes.slice(tx.processedChanges??0)){if(change.type==='cell')this.Invalidate(change.sheet,change.row,change.column);else this.Reset();}tx.processedChanges=tx.changes.length;}finally{this._syncingTransaction=false;}}
    if(context?.overrides?.has(keyOf(sheet,row,col)))return context.overrides.get(keyOf(sheet,row,col));
    this._dependency(context, sheet, row, col); const key = keyOf(sheet, row, col), record = sheet._cells.get(row * MAX_COLUMNS + col);
    if (!isFormula(record)) {
      if (record?.input != null && record.input !== '') return literalValue(record);
      const spill = this.spills.get(key); return spill ? spill.value : literalValue(record);
    }
    if (this.cache.has(key)) return this.cache.get(key);
    if (this.evaluating.has(key)) return error('#CIRC!', 'Circular reference');
    if (this.evaluating.size >= 256) return error('#NUM!', 'Calculation depth limit');
    this.evaluating.add(key); this.EvaluationCount++;
    for (const dep of this.dependencies.get(key) ?? []) this.dependents.get(dep)?.delete(key);
    this.dependencies.delete(key); this.ranges.delete(key);
    let result;
    try { result = this.Evaluate(record.input, { sheet, row, col, key, vars: new Map(), depth: 0 }); }
    catch (e) { result = isError(e) ? e : error('#VALUE!', e.message); }
    if (result?.type === 'lambda') result = error('#CALC!', 'A lambda must be invoked');
    if (Array.isArray(result)) result = this._spill(sheet, row, col, result, key);
    if (typeof result === 'number' && !Number.isFinite(result)) result = error('#NUM!');
    this.cache.set(key, result); this.dirty.delete(key); this.evaluating.delete(key); return result;
  }
  _spill(sheet, row, col, values, key) {
    values = matrix(values); const h = values.length, w = values[0]?.length ?? 0;
    if (!h || !w || h * w > MAX_OPERATION_CELLS || row + h > MAX_ROWS || col + w > MAX_COLUMNS) return error('#SPILL!', 'Spill size or worksheet limit');
    const ref = { r1: row, c1: col, r2: row + h - 1, c2: col + w - 1 };
    this._rangeDependency({ key }, sheet, ref);
    for (let r = 0; r < h; r++) for (let c = 0; c < w; c++) {
      if (!r && !c) continue;
      const record = sheet._cells.get((row + r) * MAX_COLUMNS + col + c), other = this.spills.get(keyOf(sheet, row + r, col + c));
      if ((record?.input != null && record.input !== '') || (other && other.key !== key) || sheet._meta.merges.some(m => contains(m, row + r, col + c))) return error('#SPILL!', 'Spill blocked');
    }
    this.arrays.set(key, values);sheet._used=null;
    for (let r = 0; r < h; r++) for (let c = 0; c < w; c++) if (r || c) { const target = keyOf(sheet, row + r, col + c); this.spills.set(target, { key, value: values[r][c] ?? null }); const queue = [...(this.dependents.get(target) ?? [])], seen = new Set([key]); for (const [dependent, refs] of this.ranges) if (dependent !== key && refs.some(ref => ref.sheet === sheet.Id && contains(ref, row+r, col+c))) queue.push(dependent); while (queue.length) { const dependent = queue.pop(); if (seen.has(dependent)) continue; seen.add(dependent); this.cache.delete(dependent); this.dirty.add(dependent); queue.push(...(this.dependents.get(dependent) ?? [])); } }
    return values[0][0];
  }
  Calculate({ full = false } = {}) {
    if (full) this.Reset();
    for (const s of this.Workbook._sheets) for (const n of s._formulaCells) {
      const record = s._cells.get(n);
      if (/\b(?:NOW|TODAY|RAND|RANDBETWEEN|RANDARRAY|INDIRECT|OFFSET)\s*\(/i.test(record.input)) this.Invalidate(s, Math.floor(n / MAX_COLUMNS), n % MAX_COLUMNS);
    }
    for (let pass = 0; pass < 16 && this.dirty.size; pass++) for (const key of [...this.dirty]) {
      const sep = key.indexOf(':'), s = this.Workbook._sheets.find(x => x.Id === key.slice(0, sep));
      if (!s) { this.dirty.delete(key); continue; } const p = parseCell(key.slice(sep + 1)); this.GetValue(s, p.row, p.column); this.dirty.delete(key);
    }
  }
  Evaluate(formula, context) {
    if (!this.ast.has(formula)) { if (this.ast.size >= 10000) this.ast.delete(this.ast.keys().next().value); this.ast.set(formula, parseFormula(formula)); }
    return this._eval(this.ast.get(formula), { vars: new Map(), row: 0, col: 0, depth: 0, ...context });
  }
  _read(node, ctx, sparse = false) {
    const sheet = this._sheet(node.sheet, ctx.sheet); if (!sheet) return error('#REF!', 'Unknown worksheet');
    if (node.r1 === node.r2 && node.c1 === node.c2) return this.GetValue(sheet, node.r1, node.c1, ctx);
    this._rangeDependency(ctx, sheet, node);
    const h = node.r2 - node.r1 + 1, w = node.c2 - node.c1 + 1;
    if (h * w > MAX_OPERATION_CELLS) {
      if (!sparse) return error('#NUM!', 'Range evaluation exceeds the configured cell limit');
      const result = []; for (const [n, record] of sheet._cells) { const r = Math.floor(n / MAX_COLUMNS), c = n % MAX_COLUMNS; if (contains(node, r, c) && record.input != null) result.push([this.GetValue(sheet, r, c, ctx)]); }
      for (const [key, spill] of this.spills) { if (!key.startsWith(sheet.Id + ':')) continue; const p = parseCell(key.slice(sheet.Id.length + 1)); if (contains(node, p.row, p.column)) result.push([spill.value]); }
      return result;
    }
    return Array.from({ length: h }, (_, r) => Array.from({ length: w }, (_, c) => this.GetValue(sheet, node.r1 + r, node.c1 + c, ctx)));
  }
  _eval(node, ctx) { try { return this._evaluate(node, ctx); } catch (e) { return isError(e) ? e : error('#VALUE!', e.message); } }
  _evaluate(node, ctx) {
    if (!node || ctx.depth > 128) return error('#NUM!', 'Evaluation nesting limit');
    ctx = { ...ctx, depth: ctx.depth + 1 }; const ev = n => this._eval(n, ctx);
    if (node.type === 'value') return node.value;
    if (node.type === 'array') return node.rows.map(r => r.map(ev));
    if (node.type === 'ref') return this._read(node, ctx);
    if (node.type === 'table') {
      for (const s of this.Workbook._sheets) {
        const table = s._meta.tables.find(t => t.name.toUpperCase() === node.name.toUpperCase()); if (!table) continue;
        this._rangeDependency(ctx, s, table.range);
        const headers = s.GetRange({ ...table.range, r2: table.range.r1 }).Values[0];
        const selected = node.selector.replace(/^\[|\]$/g, '').replace(/^\[|\]$/g, '');
        let ref = { ...table.range, sheet: s.Name, r1: table.range.r1 + 1 };
        if (selected === '#All') ref.r1--;
        else if (selected === '#Headers') { ref.r1--; ref.r2 = ref.r1; }
        else if (selected !== '#Data') {
          const name = selected.replace(/^@/, ''), i = headers.findIndex(h => String(h).toUpperCase() === name.toUpperCase());
          if (i < 0) return error('#REF!', 'Unknown table column'); ref.c1 += i; ref.c2 = ref.c1;
          if (selected.startsWith('@')) { if (ctx.row <= table.range.r1 || ctx.row > table.range.r2) return error('#VALUE!'); ref.r1 = ref.r2 = ctx.row; }
        }
        return this._read(ref, ctx);
      }
      return error('#NAME?', 'Unknown table');
    }
    if (node.type === 'name') {
      if (ctx.vars.has(node.name)) return ctx.vars.get(node.name);
      const value = this.Workbook._names.get(node.name); if (value == null) return error('#NAME?', 'Unknown name: ' + node.name);
      if (ctx.names?.has(node.name)) return error('#CIRC!');
      const names = new Set(ctx.names ?? []); names.add(node.name);
      return typeof value === 'string' && value.startsWith('=') ? this.Evaluate(value, { ...ctx, names }) : value;
    }
    if (node.type === 'unary') {
      if (node.op === '#') { if (node.value.type !== 'ref') return error('#REF!'); const s = this._sheet(node.value.sheet, ctx.sheet); if (!s) return error('#REF!'); this.GetValue(s, node.value.r1, node.value.c1, ctx); return this.arrays.get(keyOf(s, node.value.r1, node.value.c1)) ?? error('#REF!', 'No spill'); }
      const a = ev(node.value); if (node.op === '@') return scalar(a);
      return vectorize(a, 0, x => node.op === '-' ? -number(x) : node.op === '%' ? number(x) / 100 : number(x));
    }
    if (node.type === 'binary') return vectorize(ev(node.left), ev(node.right), (a, b) => {
      if (isError(a)) throw a; if (isError(b)) throw b;
      switch (node.op) {
        case '+': return number(a) + number(b); case '-': return number(a) - number(b); case '*': return number(a) * number(b);
        case '/': if (number(b) === 0) throw error('#DIV/0!'); return number(a) / number(b);
        case '^': { const n = number(a) ** number(b); if (!Number.isFinite(n)) throw error('#NUM!'); return n; }
        case '&': return text(a) + text(b); case '=': return compare(a, b) === 0; case '<>': return compare(a, b) !== 0;
        case '<': return compare(a, b) < 0; case '>': return compare(a, b) > 0; case '<=': return compare(a, b) <= 0; case '>=': return compare(a, b) >= 0;
      }
    });
    if (node.type !== 'call') return error('#VALUE!');
    const { name, args } = node;
    if (name === 'IF') {
      const condition = ev(args[0]); if (isError(condition)) return condition;
      if (!Array.isArray(condition)) return truth(condition) ? (args[1] ? ev(args[1]) : true) : (args[2] ? ev(args[2]) : false);
      const yes = flatten(condition).some(truth) ? ev(args[1]) : null, no = flatten(condition).some(v => !truth(v)) ? (args[2] ? ev(args[2]) : false) : null;
      const a = matrix(yes), b = matrix(no);
      return matrix(condition).map((r, ri) => r.map((v, ci) => { const m = truth(v) ? a : b; return m[m.length === 1 ? 0 : ri]?.[m[0].length === 1 ? 0 : ci] ?? error('#N/A'); }));
    }
    if (name === 'IFERROR' || name === 'IFNA') { const v = ev(args[0]); return vectorize(v, 0, x => isError(x) && (name === 'IFERROR' || x.code === '#N/A') ? ev(args[1]) : x); }
    if (name === 'IFS') { for (let i = 0; i + 1 < args.length; i += 2) if (truth(ev(args[i]))) return ev(args[i + 1]); return error('#N/A'); }
    if (name === 'SWITCH') { const v = ev(args[0]); for (let i = 1; i + 1 < args.length; i += 2) if (compare(v, ev(args[i])) === 0) return ev(args[i + 1]); return args.length % 2 === 0 ? ev(args.at(-1)) : error('#N/A'); }
    if (name === 'LET') {
      if (args.length < 3 || args.length % 2 === 0) return error('#VALUE!'); const vars = new Map(ctx.vars);
      for (let i = 0; i + 1 < args.length; i += 2) { if (args[i].type !== 'name') return error('#NAME?'); vars.set(args[i].name, this._eval(args[i + 1], { ...ctx, vars })); }
      return this._eval(args.at(-1), { ...ctx, vars });
    }
    if (name === 'LAMBDA') { if (!args.length || args.slice(0, -1).some(a => a.type !== 'name')) return error('#VALUE!'); return { type: 'lambda', parameters: args.slice(0, -1).map(a => a.name), body: args.at(-1), context: ctx }; }
    if (['MAP', 'BYROW', 'BYCOL', 'REDUCE', 'SCAN'].includes(name)) {
      const lambda = ev(args.at(-1)); if (lambda?.type !== 'lambda') return error('#VALUE!', 'Expected LAMBDA');
      if (name === 'REDUCE' || name === 'SCAN') { let value = ev(args[0]); const a = matrix(ev(args[1])); const out = a.map(r => r.map(v => (value = this._invoke(lambda, [value, v])))); return name === 'REDUCE' ? value : out; }
      let arrays = args.slice(0, -1).map(a => matrix(ev(a)));
      if (name === 'BYROW') return arrays[0].map(row => [scalar(this._invoke(lambda, [[row]]))]);
      if (name === 'BYCOL') return [arrays[0][0].map((_, c) => scalar(this._invoke(lambda, [arrays[0].map(r => [r[c]])])))];
      const a = arrays[0]; if (arrays.some(x => x.length !== a.length || x[0].length !== a[0].length)) return error('#VALUE!');
      return a.map((row, r) => row.map((_, c) => scalar(this._invoke(lambda, arrays.map(x => x[r][c])))));
    }
    if (name === 'ROW' || name === 'COLUMN') {
      if (!args.length) return name === 'ROW' ? ctx.row + 1 : ctx.col + 1;
      if (args[0].type !== 'ref') return error('#VALUE!'); return name === 'ROW' ? args[0].r1 + 1 : args[0].c1 + 1;
    }
    if ((name === 'ROWS' || name === 'COLUMNS') && args[0]?.type === 'ref') return name === 'ROWS' ? args[0].r2 - args[0].r1 + 1 : args[0].c2 - args[0].c1 + 1;
    if (name === 'ADDRESS') {
      const r = number(ev(args[0])) - 1, c = number(ev(args[1])) - 1, mode = args[2] ? number(ev(args[2])) : 1;
      if (mode < 1 || mode > 4 || (args[3] && !truth(ev(args[3])))) return error('#VALUE!', 'Only A1 address mode supported');
      const address = cellAddress(r, c).replace(/^([A-Z]+)(\d+)$/, (_, col, row) => (mode === 1 || mode === 3 ? '$' : '') + col + (mode === 1 || mode === 2 ? '$' : '') + row);
      return args[4] ? "'" + text(ev(args[4])).replace(/'/g, "''") + "'!" + address : address;
    }
    if (name === 'INDIRECT') { if (args[1] && !truth(ev(args[1]))) return error('#VALUE!', 'Only A1 mode supported'); let ref; try { ref = parseRange(text(ev(args[0]))); } catch { return error('#REF!'); } return this._read(ref, ctx); }
    if (name === 'OFFSET') {
      if (args[0]?.type !== 'ref') return error('#VALUE!'); const ref = { ...args[0] }, dr = number(ev(args[1])), dc = number(ev(args[2]));
      ref.r1 += dr; ref.r2 += dr; ref.c1 += dc; ref.c2 += dc; if (args[3]) ref.r2 = ref.r1 + number(ev(args[3])) - 1; if (args[4]) ref.c2 = ref.c1 + number(ev(args[4])) - 1;
      if (ref.r1 < 0 || ref.c1 < 0 || ref.r2 >= MAX_ROWS || ref.c2 >= MAX_COLUMNS || ref.r2 < ref.r1 || ref.c2 < ref.c1) return error('#REF!'); return this._read(ref, ctx);
    }
    if (name === 'FORMULATEXT' || name === 'ISFORMULA') {
      const ref = args[0]; if (ref?.type !== 'ref') return name === 'ISFORMULA' ? false : error('#N/A'); const s = this._sheet(ref.sheet, ctx.sheet); if (!s) return error('#REF!'); this._dependency(ctx, s, ref.r1, ref.c1); const record = s._cells.get(ref.r1 * MAX_COLUMNS + ref.c1); return name === 'ISFORMULA' ? isFormula(record) : isFormula(record) ? record.input : error('#N/A');
    }
    if (name === 'SUBTOTAL') {
      const code = number(ev(args[0])), names = ['', 'AVERAGE', 'COUNT', 'COUNTA', 'MAX', 'MIN', 'PRODUCT', 'STDEV.S', 'STDEV.P', 'SUM', 'VAR.S', 'VAR.P'], fn = this.Functions.get(names[code % 100]); if (!fn) return error('#VALUE!');
      const values = []; for (const ref of args.slice(1)) { if (ref.type !== 'ref') return error('#VALUE!'); const s = this._sheet(ref.sheet, ctx.sheet); if (!s) return error('#REF!'); this._rangeDependency(ctx, s, ref); for (const [n, record] of s._cells) { const r = Math.floor(n / MAX_COLUMNS), c = n % MAX_COLUMNS; if (!contains(ref, r, c) || s._filtered.has(r) || (code >= 100 && s._meta.rows[r]?.hidden) || /\bSUBTOTAL\s*\(/i.test(record.input)) continue; values.push(this.GetValue(s, r, c, ctx)); } }
      return fn(values);
    }
    if (name === 'COUNTBLANK' && args[0]?.type === 'ref') {
      const ref = args[0], s = this._sheet(ref.sheet, ctx.sheet); if (!s) return error('#REF!'); this._rangeDependency(ctx, s, ref);
      let count = (ref.r2 - ref.r1 + 1) * (ref.c2 - ref.c1 + 1); for (const [n] of s._cells) { const r = Math.floor(n / MAX_COLUMNS), c = n % MAX_COLUMNS; if (contains(ref, r, c)) { const v = this.GetValue(s, r, c, ctx); if (v != null && v !== '') count--; } } for(const[address,spill]of this.spills){if(!address.startsWith(s.Id+':'))continue;const p=parseCell(address.slice(s.Id.length+1));if(!s._cells.has(p.row*MAX_COLUMNS+p.column)&&contains(ref,p.row,p.column)&&spill.value!=null&&spill.value!=='')count--;} return count;
    }
    const fn = this.Functions.get(name);
    if (!fn) { const lambda = this._eval({ type: 'name', name }, ctx); return lambda?.type === 'lambda' ? this._invoke(lambda, args.map(ev)) : error('#NAME?', 'Unknown function: ' + name); }
    const sparse = ['SUM', 'AVERAGE', 'MIN', 'MAX', 'COUNT', 'COUNTA', 'PRODUCT', 'SUMSQ', 'MEDIAN', 'STDEV.S', 'STDEV.P', 'VAR.S', 'VAR.P'].includes(name);
    return fn(...args.map(a => sparse && a.type === 'ref' ? this._read(a, ctx, true) : ev(a)));
  }
  _invoke(lambda, values) {
    if (values.length !== lambda.parameters.length) return error('#VALUE!', 'Lambda argument count');
    const vars = new Map(lambda.context.vars); lambda.parameters.forEach((p, i) => vars.set(p, values[i])); return this._eval(lambda.body, { ...lambda.context, vars, depth: lambda.context.depth + 1 });
  }
  GetDependencyEdges() { return [...this.dependencies].flatMap(([target, sources]) => [...sources].map(source => ({ source, target }))); }
}
