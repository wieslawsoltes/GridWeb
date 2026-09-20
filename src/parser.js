import { error } from './errors.js';
import { parseRange } from './address.js';
import { readReference } from './reference-syntax.js';
export function tokenize(source) {
  source = source.startsWith('=') ? source.slice(1) : source;
  if (source.length > 8192) throw error('#VALUE!', 'Formula exceeds 8192 characters');
  const tokens = []; let pos = 0;
  while (pos < source.length) {
    const rest = source.slice(pos); let m, token, ref;
    if ((m = /^\s+/.exec(rest))) token = {t: 'space'};
    else if ((m = /^"((?:[^"]|"")*)"/.exec(rest))) token = {t: 'value', v: m[1].replace(/""/g, '"')};
    else if ((m = /^#(?:REF!|DIV\/0!|VALUE!|NAME\?|N\/A|NUM!|NULL!|SPILL!|CALC!|CIRC!)/i.exec(rest))) token = {t: 'value', v: error(m[0].toUpperCase())};
    else if ((ref = readReference(rest))) { m = [ref.raw]; token = {t: 'ref', v: ref.address, sheet: ref.sheet, sheetEnd: ref.sheetEnd}; }
    else if ((m = /^(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/.exec(rest))) token = {t: 'value', v: Number(m[0])};
    else if ((m = /^([A-Za-z_\\][\w.\\]*)(\[(?:[^\[\]]|\[[^\[\]]*\])*\])/.exec(rest))) token = {t: 'table', v: m[1], selector: m[2]};
    else if ((m = /^([A-Za-z_\\][\w.\\]*)/.exec(rest))) token = {t: 'name', v: m[0].replace(/^(?:(?:_xlfn|_xlws)\.)+/i, '').toUpperCase()};
    else if ((m = /^(?:<>|<=|>=|[+\-*/^&=<>():,;%{}#@])/.exec(rest))) token = {t: m[0], v: m[0]};
    else throw error('#NAME?', 'Unexpected formula token at ' + pos);
    pos += m[0].length; tokens.push(token);
    if (tokens.length > 4096) throw error('#VALUE!', 'Too many tokens');
  }
  // Whitespace is an operator only between potential reference expressions, not trivia.
  const result = tokens.filter((t, i) => {
    if (t.t !== 'space') return true;
    const before = tokens[i - 1]?.t, after = tokens[i + 1]?.t;
    return ['ref', 'table', 'name', ')', '#'].includes(before) && ['ref', 'table', 'name', '('].includes(after) && !(before === 'name' && after === '(');
  });
  result.push({t: 'end'}); return result;
}
export function parseFormula(source) {
  const tokens = tokenize(source); let i = 0, depth = 0;
  const peek = () => tokens[i].t, take = () => tokens[i++];
  const expect = type => { if (peek() !== type) throw error('#VALUE!', 'Expected ' + type); return take(); };
  const precedence = {'=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '&': 2, '+': 3, '-': 3, '*': 4, '/': 4, '^': 5, ',': 10, space: 11, ':': 12};
  function argumentsList() {
    expect('('); const args = [];
    if (peek() !== ')') while (true) {
      args.push([',', ';', ')'].includes(peek()) ? {type: 'value', value: null, omitted: true} : expression(0, false));
      if (peek() !== ',' && peek() !== ';') break;
      take();
    }
    expect(')'); return args;
  }
  function expression(min = 0, union = true) {
    if (++depth > 128) throw error('#VALUE!', 'Formula nesting limit');
    const t = take(); let n;
    if (t.t === 'value') n = {type: 'value', value: t.v};
    else if (t.t === 'ref') {
      try { n = {...parseRange(t.v), type: 'ref', sheet: t.sheet, ...(t.sheetEnd != null ? {sheetEnd: t.sheetEnd} : {})}; }
      catch { throw error('#REF!', 'Reference outside worksheet'); }
    } else if (t.t === 'table') n = {type: 'table', name: t.v, selector: t.selector};
    else if (t.t === 'name') {
      if (peek() === '(') n = {type: 'call', name: t.v, args: argumentsList()};
      else n = /^(TRUE|FALSE)$/.test(t.v) ? {type: 'value', value: t.v === 'TRUE'} : {type: 'name', name: t.v};
    } else if (t.t === '(') { n = expression(); expect(')'); }
    else if (['+', '-', '@'].includes(t.t)) n = {type: 'unary', op: t.t, value: expression(6, union)};
    else if (t.t === '{') {
      const rows = [[]];
      while (peek() !== '}') {
        rows.at(-1).push(expression(0, false));
        if (peek() === ';') { take(); rows.push([]); } else if (peek() === ',') take(); else break;
      }
      expect('}');
      if (rows.some(r => r.length !== rows[0].length)) throw error('#VALUE!', 'Ragged array');
      n = {type: 'array', rows};
    } else throw error('#VALUE!', 'Invalid formula expression');
    while (true) {
      if (peek() === '(' && min <= 14) { n = {type: 'invoke', callee: n, args: argumentsList()}; continue; }
      if ((peek() === '#' && min <= 13) || (peek() === '%' && min <= 7)) { n = {type: 'unary', op: take().t, value: n}; continue; }
      const op = peek(), p = precedence[op] ?? -1;
      if (p < min || (op === ',' && !union)) break;
      take(); const right = expression(p + 1, union);
      // A1:'Sheet'!B2 follows the explicitly qualified endpoint; ordinary ranges
      // remain ref nodes for existing AST consumers and dependency introspection.
      if (op === ':' && n.type === 'ref' && right.type === 'ref' && !n.sheetEnd && !right.sheetEnd && (!n.sheet || !right.sheet || n.sheet.toUpperCase() === right.sheet.toUpperCase())) {
        n = {type: 'ref', sheet: n.sheet ?? right.sheet, r1: Math.min(n.r1, right.r1), c1: Math.min(n.c1, right.c1), r2: Math.max(n.r2, right.r2), c2: Math.max(n.c2, right.c2)};
      } else n = {type: [',', 'space', ':'].includes(op) ? 'refop' : 'binary', op, left: n, right};
    }
    depth--; return n;
  }
  const ast = expression(); expect('end'); return ast;
}
