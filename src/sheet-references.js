/** Transactional reference repair shared by worksheet rename, removal and reordering. */
import {mapFormulaReferences, referencePrefix} from './reference-syntax.js';
const same = (a,b) => a?.toUpperCase() === b?.toUpperCase();
export function rewriteWorkbookReferences(book, transform) {
  for (const sheet of book._sheets) {
    for (const n of sheet._formulaCells) {
      const record = sheet._cells.get(n), input = transform(record.input);
      if (input !== record.input) sheet._writeRecord(n, {...record, input});
    }
    for (const field of ['validations', 'conditionalFormats']) {
      const before = structuredClone(sheet._meta[field]);
      const after = before.map(rule => {
        const result = {...rule};
        for (const key of ['formula', 'formula1', 'formula2']) if (typeof rule[key] === 'string') result[key] = transform(rule[key]);
        return result;
      });
      if (JSON.stringify(before) !== JSON.stringify(after)) book._record(() => {sheet._meta[field] = structuredClone(after);}, () => {sheet._meta[field] = structuredClone(before);}, {type:field,sheet});
    }
  }
  for (const [name,value] of book._names) if (typeof value === 'string' && value.startsWith('=')) {
    const next = transform(value); if (next !== value) book.DefineName(name,next);
  }
}
function span(ref, sheets) {
  const a = sheets.findIndex(s => same(s.Name,ref.sheet)), b = sheets.findIndex(s => same(s.Name,ref.sheetEnd));
  return a < 0 || b < 0 ? null : sheets.slice(Math.min(a,b),Math.max(a,b)+1);
}
const render = (ref, sheets) => sheets.length ? referencePrefix(sheets[0].Name, sheets.length > 1 ? sheets.at(-1).Name : null) + ref.address : '#REF!';
export function removedSheetReferences(formula, removed, sheets) {
  return mapFormulaReferences(formula, ref => {
    if (ref.sheetEnd == null) return same(ref.sheet,removed.Name) ? '#REF!' : ref.raw;
    const items = span(ref,sheets);
    return items?.includes(removed) ? render(ref,items.filter(s => s !== removed)) : ref.raw;
  });
}
export function movedSheetReferences(formula, moved, before, after) {
  return mapFormulaReferences(formula, ref => {
    if (ref.sheetEnd == null) return ref.raw;
    const items = span(ref,before);
    if (!items || items.length < 2) return ref.raw;
    const first = items[0], last = items.at(-1);
    if ((moved === first || moved === last) && after.indexOf(first) > after.indexOf(last)) return render(ref,items.filter(s => s !== moved));
    return ref.raw;
  });
}
