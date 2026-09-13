/** GridWeb addressing. Public numeric indexes are zero-based. */
export const MAX_ROWS = 1048576;
export const MAX_COLUMNS = 16384;
export const MAX_OPERATION_CELLS = 250000;
export function columnName(index) {
  if (!Number.isInteger(index) || index < 0 || index >= MAX_COLUMNS) throw new RangeError('Column out of bounds');
  let result = '';
  for (let n = index + 1; n; n = Math.floor((n - 1) / 26)) result = String.fromCharCode(65 + (n - 1) % 26) + result;
  return result;
}
export function columnIndex(name) {
  if (!/^[A-Za-z]{1,3}$/.test(name)) throw new RangeError('Invalid column');
  let n = 0; for (const c of name.toUpperCase()) n = n * 26 + c.charCodeAt(0) - 64;
  if (n > MAX_COLUMNS) throw new RangeError('Column out of bounds');
  return n - 1;
}
export function cellAddress(row, column) {
  if (!Number.isInteger(row) || row < 0 || row >= MAX_ROWS) throw new RangeError('Row out of bounds');
  return columnName(column) + (row + 1);
}
export function parseCell(address) {
  const m = /^(\$?)([A-Za-z]{1,3})(\$?)([1-9][0-9]*)$/.exec(address);
  if (!m || +m[4] > MAX_ROWS) throw new RangeError('Invalid address: ' + address);
  return { row: +m[4] - 1, column: columnIndex(m[2]), absoluteColumn: !!m[1], absoluteRow: !!m[3] };
}
export function splitSheet(address) {
  const m = /^(?:'((?:[^']|'')+)'|([^!]+))!(.*)$/.exec(address);
  return m ? { sheet: (m[1] ?? m[2]).replace(/''/g, "'"), address: m[3] } : { sheet: null, address };
}
export const quoteSheet = name => "'" + String(name).replace(/'/g, "''") + "'";
export function parseRange(address) {
  const s = splitSheet(String(address).trim()), p = s.address.split(':');
  if (p.length > 2) throw new RangeError('Invalid range');
  let a, b;
  if (p.length === 2 && p.every(x => /^\$?[A-Za-z]{1,3}$/.test(x))) {
    a = { row: 0, column: columnIndex(p[0].replace('$', '')) }; b = { row: MAX_ROWS - 1, column: columnIndex(p[1].replace('$', '')) };
  } else if (p.length === 2 && p.every(x => /^\$?[1-9]\d*$/.test(x))) {
    a = { row: +p[0].replace('$', '') - 1, column: 0 }; b = { row: +p[1].replace('$', '') - 1, column: MAX_COLUMNS - 1 };
    if (Math.max(a.row, b.row) >= MAX_ROWS) throw new RangeError('Invalid row');
  } else { a = parseCell(p[0]); b = parseCell(p[1] ?? p[0]); }
  return { sheet: s.sheet, r1: Math.min(a.row, b.row), c1: Math.min(a.column, b.column), r2: Math.max(a.row, b.row), c2: Math.max(a.column, b.column) };
}
export const contains = (a, r, c) => r >= a.r1 && r <= a.r2 && c >= a.c1 && c <= a.c2;
export const intersects = (a, b) => a.r1 <= b.r2 && b.r1 <= a.r2 && a.c1 <= b.c2 && b.c1 <= a.c2;
export function rangeAddress(a) { const start = cellAddress(a.r1, a.c1), end = cellAddress(a.r2, a.c2); return start === end ? start : `${start}:${end}`; }
export function boundedCells(a, limit = MAX_OPERATION_CELLS) {
  const count = (a.r2 - a.r1 + 1) * (a.c2 - a.c1 + 1);
  if (!Number.isSafeInteger(count) || count > limit) throw new RangeError(`Operation exceeds ${limit.toLocaleString('en-US')} cells; use smaller batches`);
  return count;
}
/** Rewrites cell references, ignoring string literals, identifiers and function names. */
export function rewriteReferences(formula, transform) {
  const re = /"(?:[^"]|"")*"|(?:(?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!)?\$?[A-Za-z]{1,3}\$?[1-9]\d*/g;
  return String(formula).replace(re, (token, offset, all) => {
    if (token[0] === '"' || /[\w.\[]/.test(all[offset - 1] ?? '') || /[\w.(\]]/.test(all[offset + token.length] ?? '')) return token;
    const s = splitSheet(token); let p; try { p = parseCell(s.address); } catch { return token; }
    return transform(p, s.sheet, token);
  });
}
export function shiftFormula(formula, dr, dc) {
  const ordinary=rewriteReferences(formula, (p, sheet) => {
    const r = p.row + (p.absoluteRow ? 0 : dr), c = p.column + (p.absoluteColumn ? 0 : dc);
    if (r < 0 || c < 0 || r >= MAX_ROWS || c >= MAX_COLUMNS) return '#REF!';
    return (sheet ? quoteSheet(sheet) + '!' : '') + (p.absoluteColumn ? '$' : '') + columnName(c) + (p.absoluteRow ? '$' : '') + (r + 1);
  });
  return rewriteAxisReferences(ordinary,(range,axis,parts)=>{
    const delta=axis==='row'?dr:dc,limit=axis==='row'?MAX_ROWS:MAX_COLUMNS;
    const indexes=parts.map(p=>(axis==='row'?+p.replace('$','')-1:columnIndex(p.replace('$','')))+(p.startsWith('$')?0:delta));
    if(indexes.some(i=>i<0||i>=limit))return '#REF!';
    return (range.sheet?quoteSheet(range.sheet)+'!':'')+indexes.map((i,n)=>(parts[n].startsWith('$')?'$':'')+(axis==='row'?i+1:columnName(i))).join(':');
  });
}
/** Rewrite whole-column / whole-row references without touching quoted strings or structured names. */
export function rewriteAxisReferences(formula, transform) {
  const re = /"(?:[^"]|"")*"|(?:(?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!)?(?:\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?[1-9]\d*:\$?[1-9]\d*)/g;
  return String(formula).replace(re,(token,offset,all)=>{
    if(token[0]==='"'||/[\w.\[]/.test(all[offset-1]??'')||/[\w.(\]]/.test(all[offset+token.length]??''))return token;
    const s=splitSheet(token);let range;try{range=parseRange(token);}catch{return token;}
    return transform(range,/^[\$]?[A-Za-z]/.test(s.address)?'column':'row',s.address.split(':'),token);
  });
}
export function renameSheetReferences(formula, previous, next) {
  const rename=(sheet,token)=>sheet?.toUpperCase()===previous.toUpperCase()?quoteSheet(next)+'!'+token.slice(token.lastIndexOf('!')+1):token;
  return rewriteAxisReferences(rewriteReferences(formula,(_p,sheet,token)=>rename(sheet,token)),(range,_axis,_parts,token)=>rename(range.sheet,token));
}
