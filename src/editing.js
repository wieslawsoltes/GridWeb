/** Standalone, transaction-backed spreadsheet editing; no DOM or platform dependencies. */
import {boundedCells, contains, intersects, shiftFormula, MAX_COLUMNS, MAX_OPERATION_CELLS} from './address.js';
import {isError, number, serialDate, fromSerial} from './errors.js';

const snapshots = new WeakSet();
const clone = value => structuredClone(value);
const portable = value => isError(value) ? {error: value.code} : value;
const modes = new Set(['all', 'values', 'formulas', 'formats', 'comments', 'validation', 'columnWidths', 'valuesAndNumberFormats', 'formulasAndNumberFormats']);
const operations = new Set(['none', 'add', 'subtract', 'multiply', 'divide']);
const contentModes = new Set(['all', 'values', 'formulas', 'valuesAndNumberFormats', 'formulasAndNumberFormats']);
const valueModes = new Set(['values', 'valuesAndNumberFormats']);
function freeze(value) {
  if (value && typeof value === 'object' && !Object.isFrozen(value)) {
    for (const child of Object.values(value)) freeze(child);
    Object.freeze(value);
  }
  return value;
}
function requireRange(range) {
  if (!range?.Worksheet?.Workbook || typeof range.GetCell !== 'function') throw new TypeError('A worksheet range is required');
  boundedCells(range.Bounds);
  return range;
}
/** Capture values now, so source edits after Copy do not change a later Paste Values. */
export function captureRange(range) {
  requireRange(range);
  const sheet = range.Worksheet, bounds = {...range.Bounds}, engine = sheet.Workbook.Calculation;
  const cells = Array.from({length: range.RowCount}, (_, r) => Array.from({length: range.ColumnCount}, (_, c) => {
    const cell = range.GetCell(r, c), value = portable(cell.Value);
    return {...clone(cell._record ?? {input: null}), value, spill: engine.spills.has(sheet.Id + ':' + cell.Address)};
  }));
  const validations = sheet._meta.validations.filter(rule => intersects(rule.range, bounds)).map(rule => {
    const clipped = {r1: Math.max(bounds.r1, rule.range.r1), c1: Math.max(bounds.c1, rule.range.c1), r2: Math.min(bounds.r2, rule.range.r2), c2: Math.min(bounds.c2, rule.range.c2)};
    return {...clone(rule), range: clipped, ...(rule.type === 'custom' ? {formula: shiftFormula(rule.formula, clipped.r1 - rule.range.r1, clipped.c1 - rule.range.c1)} : {})};
  });
  const snapshot = {kind: 'GridWebClipboard', bounds, rows: range.RowCount, columns: range.ColumnCount, cells, validations,
    columnWidths: Array.from({length: range.ColumnCount}, (_, c) => sheet._meta.columns[bounds.c1 + c]?.size ?? 100),
    hasMerges: sheet._meta.merges.some(m => intersects(m, bounds))};
  snapshots.add(snapshot);
  return freeze(snapshot);
}
function subtractRect(original, cut) {
  if (!intersects(original, cut)) return [original];
  const a = {r1: Math.max(original.r1, cut.r1), r2: Math.min(original.r2, cut.r2), c1: Math.max(original.c1, cut.c1), c2: Math.min(original.c2, cut.c2)}, out = [];
  if (original.r1 < a.r1) out.push({...original, r2: a.r1 - 1});
  if (original.r2 > a.r2) out.push({...original, r1: a.r2 + 1});
  if (original.c1 < a.c1) out.push({...a, c1: original.c1, c2: a.c1 - 1});
  if (original.c2 > a.c2) out.push({...a, c1: a.c2 + 1, c2: original.c2});
  return out;
}
function replaceStyle(cell, style) {
  // The ordinary Style setter is a patch operation. Clear the old style first,
  // then use that setter to retain the model's whitelist and protection checks.
  if (cell.Worksheet.IsProtected) throw new Error('Worksheet is protected');
  cell.Worksheet._writeRecord(cell.Row * MAX_COLUMNS + cell.Column, {...cell._record, input: cell.Input, style: {}});
  cell.Style = clone(style);
}
function arithmetic(left, right, operation) {
  try {
    if (left && typeof left === 'object' && left.error) return left;
    if (right && typeof right === 'object' && right.error) return right;
    const a = number(left), b = number(right);
    if (operation === 'divide' && b === 0) return {error: '#DIV/0!'};
    const n = operation === 'add' ? a + b : operation === 'subtract' ? a - b : operation === 'multiply' ? a * b : a / b;
    return Number.isFinite(n) ? n : {error: '#NUM!'};
  } catch (e) { return {error: isError(e) ? e.code : '#VALUE!'}; }
}
/** Returns the actual destination (a single selected cell expands to the clipboard size). */
export function pasteSpecial(destination, source, {mode = 'all', operation = 'none', transpose = false, skipBlanks = false} = {}) {
  requireRange(destination);
  const snapshot = snapshots.has(source) ? source : captureRange(source);
  if (typeof transpose !== 'boolean' || typeof skipBlanks !== 'boolean') throw new TypeError('Paste flags must be boolean');
  if (!modes.has(mode) || !operations.has(operation)) throw new TypeError('Unknown paste mode or operation');
  if (operation !== 'none' && !valueModes.has(mode)) throw new TypeError('Arithmetic paste requires Values or Values and number formats');
  if (transpose && mode === 'columnWidths') throw new TypeError('Column widths cannot be transposed');
  if (snapshot.hasMerges && contentModes.has(mode) && !valueModes.has(mode)) throw new Error('Use Paste Values for merged source cells');
  if (!valueModes.has(mode) && contentModes.has(mode) && snapshot.cells.some(row => row.some(cell => cell.spill))) throw new Error('Use Paste Values to copy an entire spilled array');
  const h = transpose ? snapshot.columns : snapshot.rows, w = transpose ? snapshot.rows : snapshot.columns;
  const target = destination.Count === 1 ? destination.Resize(h, w) : destination;
  requireRange(target);
  if (target.RowCount % h || target.ColumnCount % w) throw new RangeError('Destination dimensions must be multiples of the copied range');
  const sheet = target.Worksheet, book = sheet.Workbook;
  if (sheet.MergedRanges.some(m => intersects(m, target.Bounds))) throw new Error('Unmerge destination cells before pasting');
  const edits = [];
  for (let r = 0; r < target.RowCount; r++) for (let c = 0; c < target.ColumnCount; c++) {
    const sr = transpose ? c % w : r % h, sc = transpose ? r % h : c % w, record = snapshot.cells[sr][sc], cell = target.GetCell(r, c);
    const blank = record.input == null && record.value == null;
    if (skipBlanks && blank) continue;
    if (contentModes.has(mode) && book.Calculation.spills.has(sheet.Id + ':' + cell.Address)) throw new Error('Cannot paste into part of a spilled array');
    const dr = cell.Row - (snapshot.bounds.r1 + sr), dc = cell.Column - (snapshot.bounds.c1 + sc);
    const formula = !record.literal && typeof record.input === 'string' && record.input.startsWith('=');
    edits.push({cell, record, input: formula ? shiftFormula(record.input, dr, dc) : record.input ?? null,
      value: operation === 'none' ? record.value : arithmetic(portable(cell.Value), record.value, operation)});
  }
  // Construct the complete replacement validation set before changing any cells.
  let validations;
  if (mode === 'all' || mode === 'validation') {
    if (skipBlanks && (snapshot.validations.length || sheet._meta.validations.some(rule => intersects(rule.range,target.Bounds)))) throw new TypeError('Validation paste does not combine with Skip blanks; paste values/formats separately');
    if (transpose && snapshot.validations.some(rule => rule.type === 'custom')) throw new TypeError('Transposing custom validation formulas is not supported');
    validations = sheet._meta.validations.flatMap(rule => subtractRect(rule.range, target.Bounds).map(range => ({...clone(rule), range,
      ...(rule.type === 'custom' ? {formula: shiftFormula(rule.formula, range.r1 - rule.range.r1, range.c1 - rule.range.c1)} : {})})));
    for (let tr = 0; tr < target.RowCount; tr += h) for (let tc = 0; tc < target.ColumnCount; tc += w) for (const rule of snapshot.validations) {
      const a = rule.range, row = target.Bounds.r1 + tr, col = target.Bounds.c1 + tc;
      const range = transpose ? {r1: row + a.c1 - snapshot.bounds.c1, r2: row + a.c2 - snapshot.bounds.c1, c1: col + a.r1 - snapshot.bounds.r1, c2: col + a.r2 - snapshot.bounds.r1}
        : {r1: row + a.r1 - snapshot.bounds.r1, r2: row + a.r2 - snapshot.bounds.r1, c1: col + a.c1 - snapshot.bounds.c1, c2: col + a.c2 - snapshot.bounds.c1};
      validations.push({...clone(rule), range, ...(rule.type === 'custom' ? {formula: shiftFormula(rule.formula, range.r1 - a.r1, range.c1 - a.c1)} : {})});
      if (validations.length > 1000) throw new RangeError('Pasted validation rule limit');
    }
  }
  if (validations?.length > 1000) throw new RangeError('Pasted validation rule limit');
  book.Transaction('Paste special: ' + mode, () => {
    if (validations) sheet._setMeta('validations', validations);
    for (const {cell, record, input, value} of edits) {
      if (contentModes.has(mode)) {
        if (valueModes.has(mode)) cell.Value = clone(value);
        else sheet._setInput(cell.Row, cell.Column, clone(input), {literal: !!record.literal});
      }
      if (mode === 'all' || mode === 'formats') replaceStyle(cell, record.style ?? {});
      if (mode === 'valuesAndNumberFormats' || mode === 'formulasAndNumberFormats') cell.Style = {numberFormat: record.style?.numberFormat ?? 'General'};
      if (mode === 'all' || mode === 'comments') cell.Comment = record.comment ?? '';
    }
    if (mode === 'columnWidths') for (let c = 0; c < target.ColumnCount; c++) sheet.SetColumnWidth(target.Bounds.c1 + c, snapshot.columnWidths[c % w]);
  });
  return target;
}

function dateOffset(seed, step, unit) {
  if (step === 0) return seed;
  const day = Math.floor(seed), fraction = seed - day;
  if (unit === 'day') return seed + step;
  if (unit === 'weekday') {
    let value = day, remaining = Math.abs(step), sign = Math.sign(step);
    // Weekend starting dates need normalizing before whole-week acceleration.
    while (remaining && [0, 6].includes(fromSerial(value).getUTCDay())) {
      value += sign;
      if (![0, 6].includes(fromSerial(value).getUTCDay())) remaining--;
    }
    value += Math.floor(remaining / 5) * 7 * sign; remaining %= 5;
    while (remaining) { value += sign; if (![0, 6].includes(fromSerial(value).getUTCDay())) remaining--; }
    return value + fraction;
  }
  const d = fromSerial(day), month = d.getUTCMonth() + step * (unit === 'year' ? 12 : 1);
  const end = new Date(Date.UTC(d.getUTCFullYear(), month + 1, 0));
  const phantomFebruary = end.getUTCFullYear() === 1900 && end.getUTCMonth() === 1;
  const date = Math.min(day === 60 ? 29 : d.getUTCDate(), phantomFebruary ? 29 : end.getUTCDate());
  if (phantomFebruary && date === 29) return 60 + fraction;
  return serialDate(new Date(Date.UTC(end.getUTCFullYear(), end.getUTCMonth(), date))) + fraction;
}
/** Fill independent rows or columns, preserving all formatting and using one undo transaction. */
export function fillSeries(range, {type = 'linear', direction = 'down', step = 1, start, stop, dateUnit = 'day'} = {}) {
  requireRange(range);
  if (!['linear', 'growth', 'date'].includes(type) || !['down', 'right', 'up', 'left'].includes(direction) || !['day', 'weekday', 'month', 'year'].includes(dateUnit)) throw new TypeError('Unknown series option');
  if (typeof step !== 'number' || !Number.isFinite(step) || (type === 'date' && !Number.isInteger(step)) || (type === 'growth' && step <= 0)) throw new RangeError('Invalid series step');
  const convert = value => value instanceof Date ? serialDate(value) : value;
  start = convert(start); stop = convert(stop);
  if (start != null && (typeof start !== 'number' || !Number.isFinite(start))) throw new TypeError('Series start must be a finite number or date');
  if (stop != null && (typeof stop !== 'number' || !Number.isFinite(stop))) throw new TypeError('Series stop must be a finite number or date');
  if (range.Worksheet.MergedRanges.some(m => intersects(m, range.Bounds))) throw new Error('Unmerge cells before filling a series');
  const vertical = direction === 'down' || direction === 'up', reverse = direction === 'up' || direction === 'left';
  const length = vertical ? range.RowCount : range.ColumnCount, lines = vertical ? range.ColumnCount : range.RowCount, edits = [];
  for (let line = 0; line < lines; line++) {
    const cellAt = i => { const n = reverse ? length - 1 - i : i; return range.GetCell(vertical ? n : line, vertical ? line : n); };
    const seed = start ?? cellAt(0).Value;
    if (typeof seed !== 'number' || !Number.isFinite(seed)) throw new TypeError('Each series needs a numeric or date starting cell');
    const at = i => type === 'linear' ? seed + step * i : type === 'growth' ? seed * step ** i : dateOffset(seed, step * i, dateUnit);
    const trend = Math.sign(at(1) - seed);
    for (let i = 0; i < length; i++) {
      const value = at(i);
      if (stop != null && ((trend >= 0 && value > stop) || (trend < 0 && value < stop))) break;
      if (!Number.isFinite(value) || (type === 'date' && (value < 0 || value >= 2958466))) throw new RangeError('Series exceeds the numeric or date range');
      edits.push([cellAt(i), value]);
    }
  }
  range.Worksheet.Workbook.Transaction('Fill ' + type + ' series', () => { for (const [cell, value] of edits) cell.Value = value; });
  return edits.length;
}

/** Return individual cell ranges in row-major order; this does not imply multi-area UI selection. */
export function specialCells(range, type, {valueTypes} = {}) {
  if (!range?.Worksheet) throw new TypeError('A worksheet range is required');
  if (!['formulas', 'constants', 'blanks', 'errors', 'comments', 'validation', 'visible'].includes(type)) throw new TypeError('Unknown special-cell type');
  const allowed = new Set(['numbers', 'text', 'logical', 'errors']);
  if (valueTypes && (!Array.isArray(valueTypes) || valueTypes.some(v => !allowed.has(v)))) throw new TypeError('Invalid special-cell value filter');
  const sheet = range.Worksheet, candidates = new Set();
  if (['blanks', 'visible', 'validation'].includes(type)) {
    boundedCells(range.Bounds);
    for (let r = range.Bounds.r1; r <= range.Bounds.r2; r++) for (let c = range.Bounds.c1; c <= range.Bounds.c2; c++) candidates.add(r * MAX_COLUMNS + c);
  } else {
    for (const n of sheet._cells.keys()) if (contains(range.Bounds, Math.floor(n / MAX_COLUMNS), n % MAX_COLUMNS)) candidates.add(n);
    if (type === 'errors') for (const [key] of sheet.Workbook.Calculation.spills) if (key.startsWith(sheet.Id + ':')) {
      const address = key.slice(sheet.Id.length + 1), cell = sheet.GetCell(address);
      if (contains(range.Bounds, cell.Row, cell.Column)) candidates.add(cell.Row * MAX_COLUMNS + cell.Column);
    }
  }
  if (candidates.size > MAX_OPERATION_CELLS) throw new RangeError('Special-cell result limit');
  const result = [];
  for (const n of [...candidates].sort((a, b) => a - b)) {
    const r = Math.floor(n / MAX_COLUMNS), c = n % MAX_COLUMNS, cell = sheet.GetCell(r, c);
    const match = type === 'formulas' ? !!cell.Formula : type === 'constants' ? cell.Input != null && !cell.Formula
      : type === 'blanks' ? cell.Input == null && cell.Value == null : type === 'errors' ? isError(cell.Value)
      : type === 'comments' ? !!cell.Comment : type === 'validation' ? sheet._meta.validations.some(rule => contains(rule.range, r, c))
      : !sheet._meta.rows[r]?.hidden && !sheet._meta.columns[c]?.hidden && !sheet._filtered.has(r);
    if (!match) continue;
    if (valueTypes && (type === 'formulas' || type === 'constants')) {
      const value = cell.Value, kind = isError(value) ? 'errors' : typeof value === 'number' ? 'numbers' : typeof value === 'boolean' ? 'logical' : typeof value === 'string' ? 'text' : null;
      if (!valueTypes.includes(kind)) continue;
    }
    result.push(sheet.GetRange(cell.Address));
  }
  return result;
}
