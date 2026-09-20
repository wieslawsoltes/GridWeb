import {areasOf} from './calculation-references.js';
/** Reference-sensitive aggregations; filtering must happen before values lose their origins. */
import {error, isError, number, matrix} from './errors.js';
import {contains, parseCell, MAX_COLUMNS, MAX_OPERATION_CELLS} from './address.js';
import {parseFormula} from './parser.js';

const functions = ['', 'AVERAGE', 'COUNT', 'COUNTA', 'MAX', 'MIN', 'PRODUCT',
  'STDEV.S', 'STDEV.P', 'SUM', 'VAR.S', 'VAR.P', 'MEDIAN', 'MODE.SNGL',
  'LARGE', 'SMALL', 'PERCENTILE.INC', 'QUARTILE.INC', 'PERCENTILE.EXC', 'QUARTILE.EXC'];

function hasSubtotal(node, depth = 0) {
  if (!node || depth > 128) return false;
  if (node.type === 'call' && (node.name === 'SUBTOTAL' || node.name === 'AGGREGATE')) return true;
  return (node.args ?? []).some(n => hasSubtotal(n, depth + 1)) ||
    (node.rows ?? []).some(r => r.some(n => hasSubtotal(n, depth + 1))) ||
    ['left', 'right', 'value', 'callee'].some(k => typeof node[k] === 'object' && hasSubtotal(node[k], depth + 1));
}
function nestedFormula(engine, record) {
  if (record?.literal || typeof record?.input !== 'string' || !record.input.startsWith('=')) return false;
  let ast = engine.ast.get(record.input);
  try { ast ??= parseFormula(record.input); } catch { return false; }
  return hasSubtotal(ast);
}

export function referenceAggregate(engine, name, args, ctx, ev) {
  const subtotal = name === 'SUBTOTAL';
  if (args.length < (subtotal ? 2 : 3) || args.length > 255) throw error('#VALUE!', name + ' argument count');
  const rawCode = number(ev(args[0])), code = Math.trunc(rawCode);
  const option = subtotal ? (code >= 100 ? 1 : 0) : Math.trunc(number(ev(args[1])));
  const index = subtotal ? code % 100 : code;
  if (subtotal ? !((code >= 1 && code <= 11) || (code >= 101 && code <= 111)) : index < 1 || index > 19 || option < 0 || option > 7) throw error('#VALUE!', 'Invalid aggregation code or option');
  const arrayForm = !subtotal && index >= 14;
  if (arrayForm && args.length !== 4) throw error('#VALUE!', 'This aggregation requires one array and k');
  const ignoreHidden = !!(option & 1), ignoreErrors = !subtotal && !!(option & 2), ignoreNested = subtotal || option < 4;
  const values = [];
  const append = value => {
    if (ignoreErrors && isError(value)) return;
    if (values.length >= MAX_OPERATION_CELLS) throw error('#NUM!', 'Aggregation cell limit');
    values.push(value);
  };
  for (const arg of args.slice(subtotal ? 1 : 2, arrayForm ? 3 : undefined)) {
    const reference = engine._reference(arg, ctx);
    if (!reference) {
      if (subtotal) throw error('#VALUE!', 'SUBTOTAL requires references');
      const value = ev(arg);
      // A computed array no longer has row visibility or nested-formula metadata.
      for (const row of matrix(value)) for (const item of row) append(item);
      continue;
    }
    if(reference.threeD)throw error('#VALUE!',name+' does not accept 3-D references');
    for(const ref of areasOf(reference)){
    const sheet = engine._sheet(ref.sheet, ctx.sheet);
    if (!sheet) throw error('#REF!', 'Unknown worksheet');
    engine._rangeDependency(ctx, sheet, ref);
    const positions = new Set();
    for (const [n, record] of sheet._cells) if (record.input != null && contains(ref, Math.floor(n / MAX_COLUMNS), n % MAX_COLUMNS)) positions.add(n);
    for (const address of engine.spills.keys()) if (address.startsWith(sheet.Id + ':')) {
      const p = parseCell(address.slice(sheet.Id.length + 1));
      if (contains(ref, p.row, p.column)) positions.add(p.row * MAX_COLUMNS + p.column);
    }
    if (positions.size > MAX_OPERATION_CELLS) throw error('#NUM!', 'Aggregation cell limit');
    for (const n of [...positions].sort((a, b) => a - b)) {
      const row = Math.floor(n / MAX_COLUMNS), col = n % MAX_COLUMNS;
      if (sheet._filtered.has(row) || (ignoreHidden && sheet._meta.rows[row]?.hidden)) continue;
      if (ignoreNested && nestedFormula(engine, sheet._cells.get(n))) continue;
      append(engine.GetValue(sheet, row, col, ctx));
    }
  }
  }
  const fn = engine.Functions.get(functions[index]);
  return arrayForm ? fn(values, ev(args[3])) : fn(values);
}
