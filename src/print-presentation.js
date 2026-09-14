import { contains, shiftFormula, MAX_COLUMNS } from './address.js';
import { criteriaPredicate } from './functions.js';
import { isError, truth } from './errors.js';
const hex = s => /^#[a-f\d]{6}$/i.test(s ?? '') ? [1, 3, 5].map(i => parseInt(s.slice(i, i + 2), 16)) : [255, 255, 255];
/** Snapshot-scoped cache: never leaks obsolete rule values across revisions. */
export function createPrintPresentation(sheet) {
  const cache = new Map();
  return (row, column, cell) => {
    const value = cell.Value, table = sheet._meta.tables.find(t => contains(t.range, row, column));
    let style = cell.Style, bar = null;
    if (table) style = { fill: row === table.range.r1 ? '#107c41' : (row - table.range.r1) % 2 ? '#edf5ef' : '#fff', ...style, font: { ...(row === table.range.r1 ? { bold: true, color: '#fff' } : {}), ...style.font } };
    for (const [index, rule] of sheet._meta.conditionalFormats.entries()) {
      if (!contains(rule.range, row, column)) continue;
      let apply = false;
      if (rule.type === 'cellValue') apply = criteriaPredicate(rule.criteria ?? (({ lessThan: '<', lessThanOrEqual: '<=', greaterThan: '>', greaterThanOrEqual: '>=', equal: '=', notEqual: '<>' }[rule.operator] ?? '>') + String(rule.value ?? 0)))(value);
      else if (rule.type === 'formula') { const v = sheet.Workbook.Calculation.Evaluate(shiftFormula(rule.formula, row - rule.range.r1, column - rule.range.c1), { sheet, row, col: column }); apply = !isError(v) && truth(v); }
      else {
        if (!cache.has(index)) {
          const numbers = [], counts = new Map();
          for (const [n] of sheet._cells) if (contains(rule.range, Math.floor(n / MAX_COLUMNS), n % MAX_COLUMNS)) {
            const v = sheet.GetCell(Math.floor(n / MAX_COLUMNS), n % MAX_COLUMNS).Value;
            if (typeof v === 'number') numbers.push(v); const key = JSON.stringify([typeof v, typeof v === 'string' ? v.toUpperCase() : v]); counts.set(key, (counts.get(key) ?? 0) + 1);
          }
          cache.set(index, { min: numbers.reduce((a, b) => Math.min(a, b), numbers[0] ?? 0), max: numbers.reduce((a, b) => Math.max(a, b), numbers[0] ?? 1), counts });
        }
        const data = cache.get(index), ratio = Math.max(0, Math.min(1, (Number(value) - data.min) / (data.max - data.min || 1)));
        if (rule.type === 'dataBar' && typeof value === 'number') bar = { ratio: Math.max(0, Math.min(1, data.min >= 0 ? value / (data.max || 1) : ratio)), color: rule.color ?? '#21a366' };
        if (rule.type === 'colorScale' && typeof value === 'number') { const lo = hex(rule.minColor ?? '#fff2cc'), hi = hex(rule.maxColor ?? '#63be7b'); style = { ...style, fill: `rgb(${lo.map((v, i) => Math.round(v + (hi[i] - v) * ratio)).join(',')})` }; }
        if (rule.type === 'duplicate') apply = (data.counts.get(JSON.stringify([typeof value, typeof value === 'string' ? value.toUpperCase() : value])) ?? 0) > 1;
      }
      if (apply) style = { ...style, ...rule.style, font: { ...style.font, ...rule.style?.font } };
    }
    return { style, bar };
  };
}
