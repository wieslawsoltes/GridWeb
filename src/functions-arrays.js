import { error, isError, number, text, truth, scalar, flatten, matrix, serialDate, fromSerial } from './errors.js';
import { MAX_OPERATION_CELLS } from './address.js';

import { fail, integer, positive, probability, numeric, total, finite, checkedMatrix, vector, boundArray, realDate, logGamma, gammaP, betaI, inverseCDF, normalCDF, chooseLog, regression, weekends, textSide, complex, complexText, square } from './functions-support.js';

export function extendArrays(f, compare, criteriaPredicate) {
  const add = (name, fn) => f.set(name, (...args) => finite(fn(...args)));
  add('XMATCH', (value, lookup, matchMode = 0, searchMode = 1) => {
    const a = vector(lookup), v = scalar(value), mode = number(matchMode), search = number(searchMode);
    if (![0, -1, 1, 2].includes(mode) || ![1, -1, 2, -2].includes(search) || (mode === 2 && Math.abs(search) === 2)) fail('#VALUE!');
    if (isError(v)) throw v;
    if (Math.abs(search) === 2) {
      const direction = Math.sign(search); let low = 0, high = a.length;
      while (low < high) { const mid = (low + high) >>> 1; if (compare(a[mid], v) * direction < 0) low = mid + 1; else high = mid; }
      if (low < a.length && compare(a[low], v) === 0) return low + 1;
      const candidates = [low - 1, low].filter(i => i >= 0 && i < a.length);
      const hit = candidates.find(i => mode !== 0 && Math.sign(compare(a[i], v)) === mode); return hit == null ? fail('#N/A') : hit + 1;
    }
    let candidate = -1; const predicate = criteriaPredicate(v);
    for (let count = 0; count < a.length; count++) {
      const i = search === 1 ? count : a.length - 1 - count, c = compare(a[i], v);
      if (mode === 2 ? predicate(a[i]) : c === 0) return i + 1;
      if (Math.sign(c) === mode && (candidate < 0 || compare(a[i], a[candidate]) * mode < 0)) candidate = i;
    }
    return candidate < 0 ? fail('#N/A') : candidate + 1;
  });
  add('XLOOKUP', (value, lookup, returns, missing = error('#N/A'), match = 0, search = 1) => {
    const l = checkedMatrix(lookup), a = vector(l), r = checkedMatrix(returns), vertical = l.length > 1 || l[0].length === 1;
    if (vertical ? r.length !== a.length : r[0].length !== a.length) fail('#VALUE!', 'Return-array lookup dimension differs');
    let i; try { i = f.get('XMATCH')(value, l, match, search) - 1; } catch (e) { if (isError(e) && e.code === '#N/A') return missing; throw e; }
    const result = vertical ? [r[i]] : r.map(row => [row[i]]); return result.length * result[0].length === 1 ? result[0][0] ?? 0 : result.map(row => row.map(v => v ?? 0));
  });
  add('TEXTBEFORE', (...args) => textSide(false, ...args)); add('TEXTAFTER', (...args) => textSide(true, ...args));
  add('TEXTSPLIT', (value, columns, rows = null, ignoreEmpty = false, match = 0, pad = error('#N/A')) => {
    if (![0, 1].includes(number(match))) fail('#VALUE!');
    const delim = (v, nullable) => { if (v == null && nullable) return null; const a = flatten(v).map(text); if (!a.length || a.some(x => !x)) fail('#VALUE!'); const escaped = a.map(s => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')); return new RegExp(escaped.join('|'), number(match) ? 'i' : ''); };
    const col = delim(columns, true), row = delim(rows, true); if (!col && !row) fail('#VALUE!'); const ignore = truth(ignoreEmpty);
    let result = (row ? text(value).split(row) : [text(value)]).filter(s => !ignore || s !== '').map(s => (col ? s.split(col) : [s]).filter(v => !ignore || v !== ''));
    if (ignore) result = result.filter(r => r.length); const width = result.reduce((w, r) => Math.max(w, r.length), 0); if (!width) fail('#CALC!'); boundArray(result.length, width);
    return result.map(r => Array.from({ length: width }, (_, i) => r[i] ?? pad));
  });
  add('FILTER', (array, include, empty = error('#CALC!')) => {
    const a = checkedMatrix(array), inc = checkedMatrix(include); let result;
    if (inc[0].length === 1 && inc.length === a.length) result = a.filter((_, i) => truth(inc[i][0]));
    else if (inc.length === 1 && inc[0].length === a[0].length) result = a.map(r => r.filter((_, i) => truth(inc[0][i])));
    else fail('#VALUE!', 'Include must align with rows or columns'); return result.length && result[0].length ? result : empty;
  });
  add('EXPAND', (array, rows, columns = null, pad = error('#N/A')) => { const a = checkedMatrix(array), r = rows == null ? a.length : integer(rows), c = columns == null ? a[0].length : integer(columns); if (r < a.length || c < a[0].length) fail('#VALUE!'); boundArray(r, c); return Array.from({ length: r }, (_, i) => Array.from({ length: c }, (_, j) => a[i]?.[j] ?? pad)); });
  for (const name of ['WRAPROWS', 'WRAPCOLS']) add(name, (array, count, pad = error('#N/A')) => { const a = vector(array), n = integer(count); if (n < 1) fail(); const r = name === 'WRAPROWS' ? Math.ceil(a.length / n) : Math.min(a.length, n), c = name === 'WRAPROWS' ? Math.min(n, a.length) : Math.ceil(a.length / n); boundArray(r, c); return Array.from({ length: r }, (_, i) => Array.from({ length: c }, (_, j) => a[name === 'WRAPROWS' ? i * n + j : j * n + i] ?? pad)); });
  add('SORTBY', (array, ...args) => { const a = checkedMatrix(array), pairs = []; for (let i = 0; i < args.length;) { const by = checkedMatrix(args[i++]), order = i < args.length && !Array.isArray(args[i]) ? number(args[i++]) : 1; if (![-1, 1].includes(order)) fail('#VALUE!'); pairs.push({ by, order }); } if (!pairs.length) fail('#VALUE!'); const vertical = pairs[0].by[0].length === 1, count = vertical ? a.length : a[0].length; for (const p of pairs) if (vertical ? p.by.length !== count || p.by[0].length !== 1 : p.by.length !== 1 || p.by[0].length !== count) fail('#VALUE!'); const indices = Array.from({ length: count }, (_, i) => i).sort((i, j) => { for (const p of pairs) { const c = compare(vertical ? p.by[i][0] : p.by[0][i], vertical ? p.by[j][0] : p.by[0][j]); if (c) return c * p.order; } return i - j; }); return vertical ? indices.map(i => a[i]) : a.map(row => indices.map(i => row[i])); });
  add('NETWORKDAYS.INTL', (start, end, weekend = 1, holidays = []) => { let a = realDate(start), b = realDate(end), sign = 1; if (a > b) { [a, b] = [b, a]; sign = -1; } if (b - a > 250000) fail(); const mask = weekends(weekend), excluded = new Set(flatten(holidays).map(realDate)); let n = 0; for (let d = a; d <= b; d++) if (!mask[(fromSerial(d).getUTCDay() + 6) % 7] && !excluded.has(d)) n++; return n * sign; });
  add('WORKDAY.INTL', (start, days, weekend = 1, holidays = []) => { let date = realDate(start), remaining = integer(days); const direction = Math.sign(remaining), mask = weekends(weekend), excluded = new Set(flatten(holidays).map(realDate)); if (mask.every(Boolean) || Math.abs(remaining) > 250000) fail('#VALUE!'); let steps = 0; while (remaining) { date += direction; if (date < 0 || date >= 2958466 || ++steps > 2000000) fail(); if (!mask[(fromSerial(date).getUTCDay() + 6) % 7] && !excluded.has(date)) remaining -= direction; } return date; });
  add('ISOWEEKNUM', date => { const d = fromSerial(realDate(date)); d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7)); return Math.ceil(((d - Date.UTC(d.getUTCFullYear(), 0, 1)) / 86400000 + 1) / 7); });
  add('WEEKNUM', (date, type = 1) => { const t = integer(type), d = fromSerial(realDate(date)); if (t === 21) return f.get('ISOWEEKNUM')(date); const first = t === 1 || t === 17 ? 0 : t === 2 ? 1 : t >= 11 && t <= 16 ? t - 10 : fail(); const year = new Date(Date.UTC(d.getUTCFullYear(), 0, 1)); return Math.floor(((d - year) / 86400000 + (year.getUTCDay() - first + 7) % 7) / 7) + 1; });
  return f;
}
