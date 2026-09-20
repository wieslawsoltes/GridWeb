import { error, isError, number, text, truth, scalar, flatten, matrix, serialDate, fromSerial } from './errors.js';
import { MAX_OPERATION_CELLS } from './address.js';
const nums = args => args.flatMap(a => Array.isArray(a) ? flatten(a).filter(v => typeof v === 'number' || isError(v)) : [number(a)]).map(number);
const sum = a => a.reduce((n, x) => n + x, 0);
const checked = value => { if (typeof value === 'number' && !Number.isFinite(value)) throw error('#NUM!'); return value; };
const round = (n, digits = 0, method = 'round') => {
  n = number(n); digits = Math.trunc(number(digits)); if (Math.abs(digits) > 308) throw error('#NUM!');
  const factor = 10 ** digits, a = Math.abs(n) * factor;
  return Math.sign(n) * (method === 'round' ? Math.floor(a + 0.5 + Number.EPSILON * a) : Math[method](a)) / factor;
};
const ensureArray = (rows, cols) => { if (!Number.isInteger(rows) || !Number.isInteger(cols) || rows < 1 || cols < 1 || rows * cols > MAX_OPERATION_CELLS) throw error('#NUM!', 'Array size limit'); };
export function compare(a, b) {
  if (isError(a)) throw a; if (isError(b)) throw b;
  if (a == null) a = typeof b === 'string' ? '' : 0;
  if (b == null) b = typeof a === 'string' ? '' : 0;
  const rank = v => typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2;
  if (rank(a) !== rank(b)) return rank(a) - rank(b);
  if (typeof a === 'string') { a = a.toLocaleUpperCase('en-US'); b = String(b).toLocaleUpperCase('en-US'); }
  return a < b ? -1 : a > b ? 1 : 0;
}
export function criteriaPredicate(criteria) {
  if (isError(criteria)) throw criteria;
  if (typeof criteria !== 'string') return value => !isError(value) && compare(value, criteria) === 0;
  const m = /^(<=|>=|<>|=|<|>)?(.*)$/.exec(criteria), op = m[1] || '=', raw = m[2];
  const target = raw !== '' && Number.isFinite(Number(raw)) ? Number(raw) : /^true$/i.test(raw) ? true : /^false$/i.test(raw) ? false : raw;
  let regex;
  if (typeof target === 'string' && /[?*~]/.test(target)) {
    let pattern = '';
    for (let i = 0; i < target.length; i++) {
      let c = target[i];
      if (c === '~' && i + 1 < target.length) { c = target[++i]; pattern += c.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }
      else pattern += c === '*' ? '.*' : c === '?' ? '.' : c.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }
    regex = new RegExp('^' + pattern + '$', 'i');
  }
  return value => {
    if (isError(value)) return false;
    const c = regex ? (regex.test(text(value)) ? 0 : 1) : compare(value, target);
    return op === '=' ? c === 0 : op === '<>' ? c !== 0 : op === '<' ? c < 0 : op === '>' ? c > 0 : op === '<=' ? c <= 0 : c >= 0;
  };
}
function conditionalAggregate(values, pairs, aggregate) {
  const arrays = pairs.map(([range, criteria]) => [flatten(range), criteriaPredicate(scalar(criteria))]);
  const v = flatten(values);
  if (arrays.some(([a]) => a.length !== v.length)) throw error('#VALUE!', 'Criteria dimensions differ');
  const selected = v.filter((_, i) => arrays.every(([a, predicate]) => predicate(a[i])));
  return aggregate(selected);
}
function variance(values, sample = true) {
  const a = nums(values), n = a.length - (sample ? 1 : 0); if (n <= 0) throw error('#DIV/0!');
  const mean = sum(a) / a.length; return sum(a.map(x => (x - mean) ** 2)) / n;
}
const datePart = (v, part) => {
  const n = number(v); if (n < 0 || n > 2958465.99999) throw error('#NUM!');
  if (Math.floor(n) === 60 && part === 'Date') return 29;
  return fromSerial(n)['getUTC' + part]();
};
const asSerial = value => typeof value === 'string' ? serialDate(new Date(value)) : number(value);
function dateValue(value) { const d = new Date(text(value)); if (!Number.isFinite(d.getTime())) throw error('#VALUE!'); return Math.floor(serialDate(d)); }
function workday(start, days, holidays = []) {
  let s = Math.floor(asSerial(start)), n = Math.trunc(number(days)); if (Math.abs(n) > 250000) throw error('#NUM!');
  const excluded = new Set(flatten(holidays).map(asSerial));
  while (n) { s += Math.sign(n); const day = fromSerial(s).getUTCDay(); if (day !== 0 && day !== 6 && !excluded.has(s)) n -= Math.sign(n); }
  return s;
}
function financeRoot(fn, guess = .1) {
  let x = number(guess);
  for (let i = 0; i < 100; i++) {
    const f = fn(x), step = Math.max(1e-7, Math.abs(x) * 1e-6), d = (fn(x + step) - fn(x - step)) / (2 * step);
    if (Math.abs(f) < 1e-9) return x; if (!Number.isFinite(d) || Math.abs(d) < 1e-15) break;
    const next = x - f / d; if (!Number.isFinite(next) || next <= -1) { x = (x - 1) / 2; continue; }
    if (Math.abs(next - x) < 1e-12) return next; x = next;
  }
  throw error('#NUM!', 'Root did not converge');
}
export function createFunctionRegistry() {
  const f = new Map();
  const add = (name, fn) => f.set(name, (...args) => checked(fn(...args)));
  add('SUM', (...a) => sum(nums(a)));
  add('PRODUCT', (...a) => { const n = nums(a); return n.length ? n.reduce((x, y) => x * y, 1) : 0; });
  add('AVERAGE', (...a) => { const n = nums(a); if (!n.length) throw error('#DIV/0!'); return sum(n) / n.length; });
  add('MIN', (...a) => { const n = nums(a); return n.length ? n.reduce((x, y) => Math.min(x, y)) : 0; });
  add('MAX', (...a) => { const n = nums(a); return n.length ? n.reduce((x, y) => Math.max(x, y)) : 0; });
  add('COUNT', (...a) => flatten(a).filter(v => typeof v === 'number').length);
  add('COUNTA', (...a) => flatten(a).filter(v => v != null).length);
  add('COUNTBLANK', a => flatten(a).filter(v => v == null || v === '').length);
  add('SUMSQ', (...a) => sum(nums(a).map(x => x * x)));
  add('MEDIAN', (...a) => { const n = nums(a).sort((x, y) => x - y); if (!n.length) throw error('#NUM!'); return (n[Math.floor((n.length - 1) / 2)] + n[Math.floor(n.length / 2)]) / 2; });
  add('LARGE', (a, k) => { const n = nums([a]).sort((x, y) => y - x); k = Math.trunc(number(k)); if (k < 1 || k > n.length) throw error('#NUM!'); return n[k - 1]; });
  add('SMALL', (a, k) => { const n = nums([a]).sort((x, y) => x - y); k = Math.trunc(number(k)); if (k < 1 || k > n.length) throw error('#NUM!'); return n[k - 1]; });
  for (const [name, sample] of [['STDEV.S', true], ['STDEV.P', false], ['STDEV', true], ['STDEVP', false]]) add(name, (...a) => Math.sqrt(variance(a, sample)));
  for (const [name, sample] of [['VAR.S', true], ['VAR.P', false], ['VAR', true], ['VARP', false]]) add(name, (...a) => variance(a, sample));
  add('RANK.EQ', (n, a, order = 0) => { n = number(n); const v = nums([a]); if (!v.includes(n)) throw error('#N/A'); return 1 + v.filter(x => number(order) ? x < n : x > n).length; });
  add('PERCENTILE.INC', (a, k) => { const n = nums([a]).sort((x, y) => x - y); k = number(k); if (k < 0 || k > 1 || !n.length) throw error('#NUM!'); const i = (n.length - 1) * k, lo = Math.floor(i); return n[lo] + (i - lo) * ((n[lo + 1] ?? n[lo]) - n[lo]); });
  add('QUARTILE.INC', (a, q) => f.get('PERCENTILE.INC')(a, number(q) / 4));
  add('SUMIF', (range, c, values = range) => conditionalAggregate(values, [[range, c]], a => sum(nums([a]))));
  add('COUNTIF', (range, c) => flatten(range).filter(criteriaPredicate(scalar(c))).length);
  add('AVERAGEIF', (range, c, values = range) => conditionalAggregate(values, [[range, c]], a => f.get('AVERAGE')(a)));
  for (const name of ['SUMIFS', 'AVERAGEIFS', 'MINIFS', 'MAXIFS']) add(name, (values, ...a) => {
    if (!a.length || a.length % 2) throw error('#VALUE!'); const pairs = []; for (let i = 0; i < a.length; i += 2) pairs.push([a[i], a[i + 1]]);
    return conditionalAggregate(values, pairs, v => f.get(name.replace('IFS', ''))(v));
  });
  add('COUNTIFS', (...a) => { if (!a.length || a.length % 2) throw error('#VALUE!'); const pairs = []; for (let i = 0; i < a.length; i += 2) pairs.push([a[i], a[i + 1]]); return conditionalAggregate(a[0], pairs, v => v.length); });
  add('SUMPRODUCT', (...a) => { const arrays = a.map(flatten); if (!arrays.length || arrays.some(x => x.length !== arrays[0].length)) throw error('#VALUE!'); return sum(arrays[0].map((_, i) => arrays.reduce((v, x) => { if (isError(x[i])) throw x[i]; return v * (typeof x[i] === 'number' ? x[i] : 0); }, 1))); });
  for (const [name, method] of Object.entries({ ABS:'abs', ACOS:'acos', ACOSH:'acosh', ASIN:'asin', ASINH:'asinh', ATAN:'atan', ATANH:'atanh', COS:'cos', COSH:'cosh', SIN:'sin', SINH:'sinh', TAN:'tan', TANH:'tanh', EXP:'exp', LN:'log', LOG10:'log10', SQRT:'sqrt', SIGN:'sign', INT:'floor', TRUNC:'trunc' })) add(name, n => Math[method](number(n)));
  add('PI', () => Math.PI); add('POWER', (a, b) => number(a) ** number(b)); add('LOG', (a, base = 10) => Math.log(number(a)) / Math.log(number(base)));
  add('ATAN2', (x, y) => { x = number(x); y = number(y); if (!x && !y) throw error('#DIV/0!'); return Math.atan2(y, x); });
  add('RADIANS', n => number(n) * Math.PI / 180); add('DEGREES', n => number(n) * 180 / Math.PI);
  add('MOD', (a, b) => { a = number(a); b = number(b); if (!b) throw error('#DIV/0!'); return a - b * Math.floor(a / b); });
  add('QUOTIENT', (a, b) => { b = number(b); if (!b) throw error('#DIV/0!'); return Math.trunc(number(a) / b); });
  add('ROUND', round); add('ROUNDUP', (a, d) => round(a, d, 'ceil')); add('ROUNDDOWN', (a, d) => round(a, d, 'floor'));
  add('MROUND', (a, multiple) => { a = number(a); multiple = number(multiple); if (a * multiple < 0) throw error('#NUM!'); return !multiple ? 0 : round(a / multiple) * multiple; });
  add('CEILING.MATH', (n, s = 1, mode = 0) => { n = number(n); s = Math.abs(number(s)); return !s ? 0 : (n < 0 && number(mode) ? Math.floor(n / s) : Math.ceil(n / s)) * s; });
  add('FLOOR.MATH', (n, s = 1, mode = 0) => { n = number(n); s = Math.abs(number(s)); return !s ? 0 : (n < 0 && number(mode) ? Math.ceil(n / s) : Math.floor(n / s)) * s; });
  add('EVEN', n => Math.sign(number(n)) * Math.ceil(Math.abs(number(n)) / 2) * 2);
  add('ODD', n => { n = number(n); return (Math.sign(n) || 1) * (Math.ceil((Math.abs(n) - 1) / 2) * 2 + 1); });
  add('FACT', n => { n = Math.trunc(number(n)); if (n < 0 || n > 170) throw error('#NUM!'); let v = 1; for (let i = 2; i <= n; i++) v *= i; return v; });
  add('COMBIN', (n, k) => { n = Math.trunc(number(n)); k = Math.trunc(number(k)); if (k < 0 || n < k) throw error('#NUM!'); k = Math.min(k, n - k); if (k > 10000) throw error('#NUM!'); let v = 1; for (let i = 1; i <= k; i++) v *= (n - i + 1) / i; return Math.round(v); });
  add('GCD', (...a) => { const n = nums(a).map(Math.trunc); if (n.some(x => x < 0)) throw error('#NUM!'); return n.reduce((x, y) => { while (y) [x, y] = [y, x % y]; return x; }, 0); });
  add('LCM', (...a) => nums(a).map(Math.trunc).reduce((x, y) => !x || !y ? 0 : Math.abs(x * y) / f.get('GCD')(x, y), 1));
  add('RAND', () => Math.random()); add('RANDBETWEEN', (a, b) => { a = Math.ceil(number(a)); b = Math.floor(number(b)); if (a > b) throw error('#NUM!'); return a + Math.floor(Math.random() * (b - a + 1)); });
  add('AND', (...a) => flatten(a).every(truth)); add('OR', (...a) => flatten(a).some(truth)); add('XOR', (...a) => flatten(a).filter(truth).length % 2 === 1); add('NOT', a => !truth(a));
  add('TRUE', () => true); add('FALSE', () => false); add('NA', () => error('#N/A'));
  add('ISERROR', a => isError(scalar(a))); add('ISERR', a => isError(scalar(a)) && scalar(a).code !== '#N/A'); add('ISNA', a => isError(scalar(a)) && scalar(a).code === '#N/A');
  add('ISNUMBER', a => typeof scalar(a) === 'number'); add('ISTEXT', a => typeof scalar(a) === 'string'); add('ISBLANK', a => scalar(a) == null); add('ISLOGICAL', a => typeof scalar(a) === 'boolean'); add('ISNONTEXT', a => typeof scalar(a) !== 'string');
  add('ISEVEN', a => Math.trunc(number(a)) % 2 === 0); add('ISODD', a => Math.abs(Math.trunc(number(a)) % 2) === 1);
  add('N', a => isError(scalar(a)) ? scalar(a) : typeof scalar(a) === 'number' ? scalar(a) : scalar(a) === true ? 1 : 0);
  add('T', a => isError(scalar(a)) ? scalar(a) : typeof scalar(a) === 'string' ? scalar(a) : '');
  add('TYPE', a => Array.isArray(a) ? 64 : isError(a) ? 16 : typeof a === 'string' ? 2 : typeof a === 'boolean' ? 4 : 1);
  add('LEN', a => text(a).length); add('LOWER', a => text(a).toLowerCase()); add('UPPER', a => text(a).toUpperCase());
  add('PROPER', a => text(a).toLowerCase().replace(/(^|[^\p{L}])\p{L}/gu, c => c.toUpperCase()));
  add('TRIM', a => text(a).replace(/^ +| +$/g, '').replace(/ +/g, ' ')); add('CLEAN', a => text(a).replace(/[\x00-\x1f]/g, ''));
  add('LEFT', (a, n = 1) => { n = Math.trunc(number(n)); if (n < 0) throw error('#VALUE!'); return text(a).slice(0, n); });
  add('RIGHT', (a, n = 1) => { n = Math.trunc(number(n)); if (n < 0) throw error('#VALUE!'); return n ? text(a).slice(-n) : ''; });
  add('MID', (a, start, n) => { start = Math.trunc(number(start)); n = Math.trunc(number(n)); if (start < 1 || n < 0) throw error('#VALUE!'); return text(a).slice(start - 1, start - 1 + n); });
  add('CONCAT', (...a) => flatten(a).map(text).join('')); add('CONCATENATE', (...a) => a.map(text).join(''));
  add('TEXTJOIN', (delimiter, ignore, ...a) => flatten(a).filter(v => !truth(ignore) || (v != null && v !== '')).map(text).join(text(delimiter)));
  add('REPT', (a, n) => { n = Math.trunc(number(n)); if (n < 0 || text(a).length * n > 32767) throw error('#VALUE!'); return text(a).repeat(n); });
  add('REPLACE', (a, start, n, b) => { a = text(a); start = Math.trunc(number(start)); n = Math.trunc(number(n)); if (start < 1 || n < 0) throw error('#VALUE!'); return a.slice(0, start - 1) + text(b) + a.slice(start - 1 + n); });
  add('SUBSTITUTE', (a, old, replacement, instance) => { a = text(a); old = text(old); replacement = text(replacement); if (!old) return a; if (instance == null) return a.split(old).join(replacement); instance = Math.trunc(number(instance)); if (instance < 1) throw error('#VALUE!'); let i = 0; return a.replace(new RegExp(old.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g'), m => ++i === instance ? replacement : m); });
  add('FIND', (what, where, start = 1) => { start = Math.trunc(number(start)); if (start < 1 || start > text(where).length + 1) throw error('#VALUE!'); const i = text(where).indexOf(text(what), start - 1); if (i < 0) throw error('#VALUE!'); return i + 1; });
  add('SEARCH', (what, where, start = 1) => f.get('FIND')(text(what).toUpperCase(), text(where).toUpperCase(), start));
  add('EXACT', (a, b) => text(a) === text(b)); add('VALUE', number);
  add('NUMBERVALUE', (a, decimal = '.', group = ',') => number(text(a).split(text(group)).join('').replace(text(decimal), '.')));
  add('CHAR', n => { n = Math.trunc(number(n)); if (n < 1 || n > 255) throw error('#VALUE!'); return String.fromCharCode(n); });
  add('UNICHAR', n => { n = Math.trunc(number(n)); if (n < 1 || n > 0x10ffff || (n >= 0xd800 && n <= 0xdfff)) throw error('#VALUE!'); return String.fromCodePoint(n); });
  add('CODE', a => { if (!text(a)) throw error('#VALUE!'); return text(a).charCodeAt(0); }); add('UNICODE', a => { if (!text(a)) throw error('#VALUE!'); return text(a).codePointAt(0); });
  add('TEXTBEFORE', (a, delimiter, instance = 1) => { const p = text(a).split(text(delimiter)); instance = Math.trunc(number(instance)); const i = instance < 0 ? p.length + instance : instance; if (!instance || i < 1 || i >= p.length) throw error('#N/A'); return p.slice(0, i).join(text(delimiter)); });
  add('TEXTAFTER', (a, delimiter, instance = 1) => { const p = text(a).split(text(delimiter)); instance = Math.trunc(number(instance)); const i = instance < 0 ? p.length + instance : instance; if (!instance || i < 1 || i >= p.length) throw error('#N/A'); return p.slice(i).join(text(delimiter)); });
  add('TEXTSPLIT', (a, col, row = null, ignoreEmpty = false) => { const rows = row == null ? [text(a)] : text(a).split(text(row)); const out = rows.map(r => r.split(text(col)).filter(x => !truth(ignoreEmpty) || x)); const width = Math.max(...out.map(r => r.length)); ensureArray(out.length, width); return out.map(r => Array.from({ length: width }, (_, i) => r[i] ?? error('#N/A'))); });
  add('DATE', (y, m, d) => { y = Math.trunc(number(y)); if (y >= 0 && y < 1900) y += 1900; if(y===1900&&number(m)===2&&number(d)===29)return 60; const v = serialDate(new Date(Date.UTC(y, Math.trunc(number(m)) - 1, Math.trunc(number(d))))); if (v < 0 || v > 2958465) throw error('#NUM!'); return v; });
  add('TODAY', () => Math.floor(serialDate(new Date()))); add('NOW', () => serialDate(new Date())); add('DATEVALUE', dateValue);
  add('DAY', a => datePart(asSerial(a), 'Date')); add('MONTH', a => datePart(asSerial(a), 'Month') + 1); add('YEAR', a => datePart(asSerial(a), 'FullYear'));
  add('HOUR', a => datePart(a, 'Hours')); add('MINUTE', a => datePart(a, 'Minutes')); add('SECOND', a => datePart(a, 'Seconds'));
  add('TIME', (h, m, s) => { const n = (Math.trunc(number(h)) * 3600 + Math.trunc(number(m)) * 60 + Math.trunc(number(s))) / 86400; if (n < 0) throw error('#NUM!'); return n % 1; });
  add('DAYS', (end, start) => Math.floor(asSerial(end)) - Math.floor(asSerial(start)));
  add('WEEKDAY', (a, type = 1) => { const d = fromSerial(asSerial(a)).getUTCDay(); type = number(type); if (type === 1) return d + 1; if (type === 2) return (d + 6) % 7 + 1; if (type === 3) return (d + 6) % 7; throw error('#NUM!'); });
  add('EDATE', (a, months) => { const d = fromSerial(asSerial(a)), m = d.getUTCMonth() + Math.trunc(number(months)), end = new Date(Date.UTC(d.getUTCFullYear(), m + 1, 0)); return serialDate(new Date(Date.UTC(end.getUTCFullYear(), end.getUTCMonth(), Math.min(d.getUTCDate(), end.getUTCDate())))); });
  add('EOMONTH', (a, months) => { const d = fromSerial(asSerial(a)); return serialDate(new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + Math.trunc(number(months)) + 1, 0))); });
  add('WORKDAY', workday);
  add('NETWORKDAYS', (a, b, holidays = []) => { a = Math.floor(asSerial(a)); b = Math.floor(asSerial(b)); const sign = b < a ? -1 : 1; if (sign < 0) [a, b] = [b, a]; if (b - a > 250000) throw error('#NUM!'); const h = new Set(flatten(holidays).map(asSerial)); let n = 0; for (let i = a; i <= b; i++) if (![0, 6].includes(fromSerial(i).getUTCDay()) && !h.has(i)) n++; return n * sign; });
  add('PMT', (rate, periods, pv, fv = 0, type = 0) => { const r = number(rate), n = number(periods), p = number(pv), v = number(fv), t = number(type); if (!n) throw error('#DIV/0!'); const q = (1 + r) ** n; return r ? -(p * q + v) * r / ((1 + r * t) * (q - 1)) : -(p + v) / n; });
  add('FV', (rate, periods, pmt, pv = 0, type = 0) => { const r = number(rate), n = number(periods), p = number(pmt), v = number(pv), t = number(type), q = (1 + r) ** n; return r ? -(v * q + p * (1 + r * t) * (q - 1) / r) : -(v + p * n); });
  add('PV', (rate, periods, pmt, fv = 0, type = 0) => { const r = number(rate), n = number(periods), p = number(pmt), v = number(fv), t = number(type), q = (1 + r) ** n; return r ? -(v + p * (1 + r * t) * (q - 1) / r) / q : -(v + p * n); });
  add('NPER', (rate, pmt, pv, fv = 0, type = 0) => { const r = number(rate), p = number(pmt), v = number(pv), end = number(fv), t = number(type); return r ? Math.log((p * (1 + r * t) - end * r) / (v * r + p * (1 + r * t))) / Math.log(1 + r) : -(v + end) / p; });
  add('NPV', (rate, ...values) => { const r = number(rate), v = nums(values); return sum(v.map((x, i) => x / (1 + r) ** (i + 1))); });
  add('IRR', (values, guess = .1) => { const v = nums([values]); if (!v.some(x => x > 0) || !v.some(x => x < 0)) throw error('#NUM!'); return financeRoot(r => sum(v.map((x, i) => x / (1 + r) ** i)), guess); });
  add('RATE', (periods, pmt, pv, fv = 0, type = 0, guess = .1) => financeRoot(r => f.get('FV')(r, periods, pmt, pv, type) - number(fv), guess));
  add('SLN', (cost, salvage, life) => (number(cost) - number(salvage)) / number(life));
  add('SYD', (cost, salvage, life, period) => 2 * (number(cost) - number(salvage)) * (number(life) - number(period) + 1) / (number(life) * (number(life) + 1)));
  add('ROWS', a => matrix(a).length); add('COLUMNS', a => matrix(a)[0]?.length ?? 0);
  add('INDEX', (a, row, col) => { a=matrix(a);if(!a.length||!a[0]?.length||a.some(r=>r.length!==a[0].length))throw error('#VALUE!');row=Math.trunc(number(row));if(col==null&&a.length===1){col=row;row=1;}col=col==null?0:Math.trunc(number(col));if(row<0||col<0)throw error('#VALUE!');if(row>a.length||col>a[0].length)throw error('#REF!');if(!row)return col?a.map(r=>[r[col-1]]):a;if(!col)return a[0].length===1?a[row-1][0]:[a[row-1]];return a[row-1][col-1]; });
  add('MATCH', (value, array, matchType = 1) => { const a = flatten(array); const t = number(matchType); let found = -1; const pred = criteriaPredicate(value); for (let i = 0; i < a.length; i++) { if (t === 0 ? pred(a[i]) : t > 0 ? compare(a[i], scalar(value)) <= 0 : compare(a[i], scalar(value)) >= 0) { found = i; if (t === 0) break; } } if (found < 0) throw error('#N/A'); return found + 1; });
  add('XMATCH', (value, array, matchMode = 0, searchMode = 1) => { const a = flatten(array), mode = number(matchMode), indices = a.map((_, i) => i); if (Math.abs(number(searchMode)) !== 1) throw error('#VALUE!', 'Binary search modes are not implemented'); if (number(searchMode) < 0) indices.reverse(); let candidate = -1; const pred = criteriaPredicate(value); for (const i of indices) { const c = compare(a[i], scalar(value)); if ((mode === 2 ? pred(a[i]) : c === 0)) return i + 1; if ((mode === -1 && c < 0) || (mode === 1 && c > 0)) if (candidate < 0 || (mode === -1 ? compare(a[i], a[candidate]) > 0 : compare(a[i], a[candidate]) < 0)) candidate = i; } if (candidate >= 0) return candidate + 1; throw error('#N/A'); });
  add('XLOOKUP', (value, lookup, returns, missing = error('#N/A'), match = 0, search = 1) => { let i; try { i = f.get('XMATCH')(value, lookup, match, search) - 1; } catch (e) { if (isError(e) && e.code === '#N/A') return missing; throw e; } const m = matrix(returns); return m.length === 1 ? m[0][i] ?? error('#VALUE!') : m[i]?.length === 1 ? m[i][0] : m[i] ? [m[i]] : error('#VALUE!'); });
  add('VLOOKUP', (value, array, col, approximate = true) => { const a = matrix(array); col = Math.trunc(number(col)); if (col < 1 || col > a[0].length) throw error('#REF!'); return a[f.get('MATCH')(value, a.map(r => r[0]), truth(approximate) ? 1 : 0) - 1][col - 1]; });
  add('HLOOKUP', (value, array, row, approximate = true) => f.get('VLOOKUP')(value, f.get('TRANSPOSE')(array), row, approximate));
  add('CHOOSE', (index, ...values) => { const n = Math.trunc(number(index)); if (n < 1 || n > values.length) throw error('#VALUE!'); return values[n - 1]; });
  add('TRANSPOSE', a => { a = matrix(a); return a[0].map((_, i) => a.map(r => r[i])); });
  add('SEQUENCE', (rows = 1, cols = 1, start = 1, step = 1) => { rows = Math.trunc(number(rows)); cols = Math.trunc(number(cols)); ensureArray(rows, cols); return Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (_, c) => number(start) + (r * cols + c) * number(step))); });
  add('RANDARRAY', (rows = 1, cols = 1, min = 0, max = 1, integer = false) => { rows = Math.trunc(number(rows)); cols = Math.trunc(number(cols)); ensureArray(rows, cols); min = number(min); max = number(max); if (max < min) throw error('#VALUE!'); return Array.from({ length: rows }, () => Array.from({ length: cols }, () => truth(integer) ? f.get('RANDBETWEEN')(min, max) : min + Math.random() * (max - min))); });
  add('FILTER', (array, include, empty = error('#CALC!')) => { const a = matrix(array), inc = flatten(include); if (inc.length !== a.length) throw error('#VALUE!'); const rows = a.filter((_, i) => truth(inc[i])); return rows.length ? rows : empty; });
  add('SORT', (array, index = 1, order = 1, byCol = false) => { let a = matrix(array).map(r => [...r]); if (truth(byCol)) a = f.get('TRANSPOSE')(a); index = Math.trunc(number(index)) - 1; order = number(order); if (index < 0 || index >= a[0].length || ![-1, 1].includes(order)) throw error('#VALUE!'); a.sort((x, y) => compare(x[index], y[index]) * order); return truth(byCol) ? f.get('TRANSPOSE')(a) : a; });
  add('UNIQUE', (array, byCol = false, exactlyOnce = false) => { let a = matrix(array); if (truth(byCol)) a = f.get('TRANSPOSE')(a); const counts = new Map(); const key = row => JSON.stringify(row.map(x => typeof x === 'string' ? x.toUpperCase() : x)); for (const r of a) counts.set(key(r), (counts.get(key(r)) ?? 0) + 1); const seen = new Set(), rows = a.filter(r => { const k = key(r); if (seen.has(k) || (truth(exactlyOnce) && counts.get(k) !== 1)) return false; seen.add(k); return true; }); if (!rows.length) throw error('#CALC!'); return truth(byCol) ? f.get('TRANSPOSE')(rows) : rows; });
  add('TAKE', (array, rows, cols) => { const a = matrix(array), r = Math.trunc(number(rows)), c = cols == null ? a[0].length : Math.trunc(number(cols)); if (!r || !c) throw error('#CALC!'); return (r > 0 ? a.slice(0, r) : a.slice(r)).map(row => c > 0 ? row.slice(0, c) : row.slice(c)); });
  add('DROP', (array, rows, cols = 0) => { const a = matrix(array), r = Math.trunc(number(rows)), c = Math.trunc(number(cols)); const out = (r >= 0 ? a.slice(r) : a.slice(0, r)).map(row => c >= 0 ? row.slice(c) : row.slice(0, c)); if (!out.length || !out[0].length) throw error('#CALC!'); return out; });
  add('CHOOSECOLS', (array, ...cols) => { const a = matrix(array), indices = flatten(cols).map(v => { const n = Math.trunc(number(v)); const i = n < 0 ? a[0].length + n : n - 1; if (i < 0 || i >= a[0].length) throw error('#VALUE!'); return i; }); return a.map(row => indices.map(i => row[i])); });
  add('CHOOSEROWS', (array, ...rows) => f.get('TRANSPOSE')(f.get('CHOOSECOLS')(f.get('TRANSPOSE')(array), ...rows)));
  add('VSTACK', (...arrays) => { const m = arrays.map(matrix), width = Math.max(...m.map(a => a[0].length)); return m.flatMap(a => a.map(r => Array.from({ length: width }, (_, i) => r[i] ?? error('#N/A')))); });
  add('HSTACK', (...arrays) => { const m = arrays.map(matrix), height = Math.max(...m.map(a => a.length)); return Array.from({ length: height }, (_, r) => m.flatMap(a => a[r] ?? a[0].map(() => error('#N/A')))); });
  add('TOCOL', (array, ignore = 0, byColumn = false) => { const a = truth(byColumn) ? f.get('TRANSPOSE')(array) : matrix(array), n = number(ignore); return flatten(a).filter(v => !((n === 1 || n === 3) && v == null) && !((n === 2 || n === 3) && isError(v))).map(v => [v ?? 0]); });
  add('TOROW', (array, ignore = 0, byColumn = false) => f.get('TRANSPOSE')(f.get('TOCOL')(array, ignore, byColumn)));
  return f;
}
