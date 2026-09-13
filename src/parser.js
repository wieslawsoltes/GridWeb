import { error } from './errors.js';
import { parseCell, columnIndex, MAX_ROWS, MAX_COLUMNS } from './address.js';
const refPattern = /^(?:(?:'((?:[^']|'')+)'|([A-Za-z_][\w.]*))!)?(\$?[A-Za-z]{1,3}\$?[1-9]\d*|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?[1-9]\d*:\$?[1-9]\d*)/;
export function tokenize(source) {
  source = source.startsWith('=') ? source.slice(1) : source;
  if (source.length > 8192) throw error('#VALUE!', 'Formula exceeds 8192 characters');
  const tokens = []; let pos = 0;
  while (pos < source.length) {
    if (/\s/.test(source[pos])) { pos++; continue; }
    const rest = source.slice(pos); let m;
    if ((m = /^"((?:[^"]|"")*)"/.exec(rest))) tokens.push({ t: 'value', v: m[1].replace(/""/g, '"') });
    else if ((m = /^#(?:REF!|DIV\/0!|VALUE!|NAME\?|N\/A|NUM!|NULL!|SPILL!|CALC!|CIRC!)/i.exec(rest))) tokens.push({ t: 'value', v: error(m[0].toUpperCase()) });
    else if ((m = refPattern.exec(rest)) && !/[\w.(\[]/.test(rest[m[0].length] ?? '')) {
      tokens.push({ t: 'ref', v: m[3], sheet: m[1]?.replace(/''/g, "'") ?? m[2] ?? null });
    } else if ((m = /^(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/.exec(rest))) tokens.push({ t: 'value', v: Number(m[0]) });
    else if ((m = /^([A-Za-z_\\][\w.\\]*)(\[(?:[^\[\]]|\[[^\[\]]*\])*\])/.exec(rest))) tokens.push({ t: 'table', v: m[1], selector: m[2] });
    else if ((m = /^([A-Za-z_\\][\w.\\]*)/.exec(rest))) tokens.push({ t: 'name', v: m[0].replace(/^(?:(?:_xlfn|_xlws)\.)+/i, '').toUpperCase() });
    else if ((m = /^(?:<>|<=|>=|[+\-*/^&=<>():,;%{}#@])/.exec(rest))) tokens.push({ t: m[0], v: m[0] });
    else throw error('#NAME?', 'Unexpected formula token at ' + pos);
    pos += m[0].length;
    if (tokens.length > 4096) throw error('#VALUE!', 'Too many tokens');
  }
  tokens.push({ t: 'end' }); return tokens;
}
export function parseFormula(source) {
  const tokens = tokenize(source); let i = 0, depth = 0;
  const peek = () => tokens[i].t, take = () => tokens[i++];
  const expect = type => { if (peek() !== type) throw error('#VALUE!', 'Expected ' + type); return take(); };
  const precedence = { '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '&': 2, '+': 3, '-': 3, '*': 4, '/': 4, '^': 5 };
  function reference(t) {
    let a, b;
    if (t.v.includes(':')) {
      const [x, y] = t.v.replace(/\$/g, '').split(':');
      if (/^\d/.test(x)) { a = { row: +x - 1, column: 0 }; b = { row: +y - 1, column: MAX_COLUMNS - 1 }; }
      else { a = { row: 0, column: columnIndex(x) }; b = { row: MAX_ROWS - 1, column: columnIndex(y) }; }
    } else { a = parseCell(t.v); b = a; }
    if (peek() === ':') { take(); const end = expect('ref'); b = parseCell(end.v); if (end.sheet && end.sheet !== t.sheet) throw error('#REF!'); }
    if (Math.max(a.row, b.row) >= MAX_ROWS) throw error('#REF!');
    return { type: 'ref', sheet: t.sheet, r1: Math.min(a.row, b.row), c1: Math.min(a.column, b.column), r2: Math.max(a.row, b.row), c2: Math.max(a.column, b.column) };
  }
  function primary() {
    if (++depth > 128) throw error('#VALUE!', 'Formula nesting limit');
    const t = take(); let n;
    if (t.t === 'value') n = { type: 'value', value: t.v };
    else if (t.t === 'ref') n = reference(t);
    else if (t.t === 'table') n = { type: 'table', name: t.v, selector: t.selector };
    else if (t.t === 'name') {
      if (peek() === '(') {
        take(); const args = [];
        if (peek() !== ')') {
          while (true) {
            args.push([',', ';', ')'].includes(peek()) ? { type: 'value', value: null } : expression());
            if (peek() !== ',' && peek() !== ';') break; take();
          }
        }
        expect(')'); n = { type: 'call', name: t.v, args };
      } else n = /^(TRUE|FALSE)$/.test(t.v) ? { type: 'value', value: t.v === 'TRUE' } : { type: 'name', name: t.v };
    } else if (t.t === '(') { n = expression(); expect(')'); }
    else if (t.t === '+' || t.t === '-' || t.t === '@') n = { type: 'unary', op: t.t, value: primary() };
    else if (t.t === '{') {
      const rows = [[]];
      while (peek() !== '}') {
        rows.at(-1).push(expression());
        if (peek() === ';') { take(); rows.push([]); } else if (peek() === ',') take(); else break;
      }
      expect('}'); if (rows.some(r => r.length !== rows[0].length)) throw error('#VALUE!', 'Ragged array'); n = { type: 'array', rows };
    } else throw error('#VALUE!', 'Invalid formula expression');
    while (peek() === '%' || peek() === '#') n = { type: 'unary', op: take().t, value: n };
    depth--; return n;
  }
  function expression(min = 0) {
    let left = primary();
    while ((precedence[peek()] ?? -1) >= min) {
      const op = take().t, p = precedence[op]; const right = expression(p + 1);
      left = { type: 'binary', op, left, right };
    }
    return left;
  }
  const ast = expression(); expect('end'); return ast;
}
