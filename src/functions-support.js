import { error, isError, number, text, truth, scalar, flatten, matrix, serialDate, fromSerial } from './errors.js';
import { MAX_OPERATION_CELLS } from './address.js';

const fail = (code = '#NUM!', detail = '') => { throw error(code, detail); };
const integer = v => Math.trunc(number(v));
const positive = v => { const n = number(v); return n > 0 ? n : fail(); };
const probability = v => { const n = number(v); return n >= 0 && n <= 1 ? n : fail(); };
const numeric = v => flatten(v).filter(x => typeof x === 'number' || isError(x)).map(number);
const total = a => a.reduce((s, v) => s + v, 0);
const finite = v => { if (typeof v === 'number' && !Number.isFinite(v)) fail(); return v; };
const checkedMatrix = v => { const m = matrix(v); if (!m.length || !m[0].length || m.some(r => !Array.isArray(r) || r.length !== m[0].length) || m.length * m[0].length > MAX_OPERATION_CELLS) fail('#VALUE!', 'Rectangular bounded array required'); return m; };
const vector = v => { const m = checkedMatrix(v); if (m.length !== 1 && m[0].length !== 1) fail('#VALUE!', 'One-dimensional lookup array required'); return flatten(m); };
const boundArray = (r, c) => { if (r < 1 || c < 1 || r * c > MAX_OPERATION_CELLS) fail('#NUM!', 'Array size limit'); };
const realDate = v => { const n = typeof v === 'string' ? serialDate(new Date(v)) : number(v); if (!Number.isFinite(n) || n < 0 || n >= 2958466) fail(); return Math.floor(n); };

/** Lanczos log-gamma plus convergent series/continued fractions. No external runtime. */
function logGamma(z) {
  if (!(z > 0)) fail();
  const p = [676.5203681218851, -1259.1392167224028, 771.3234287776531, -176.6150291621406, 12.507343278686905, -0.13857109526572012, 9.984369578019572e-6, 1.5056327351493116e-7];
  if (z < .5) return Math.log(Math.PI) - Math.log(Math.sin(Math.PI * z)) - logGamma(1 - z);
  z--; let x = .99999999999980993; for (let i = 0; i < p.length; i++) x += p[i] / (z + i + 1);
  const t = z + 7.5; return .9189385332046727 + (z + .5) * Math.log(t) - t + Math.log(x);
}
function gammaP(a, x, complement = false) {
  if (x < 0 || a <= 0) fail(); if (x === 0) return complement ? 1 : 0;
  const factor = Math.exp(a * Math.log(x) - x - logGamma(a));
  if (x < a + 1) { let term = 1 / a, sum = term; for (let i = 1; i <= 10000; i++) { term *= x / (a + i); sum += term; if (Math.abs(term) < Math.abs(sum) * 2e-15) return complement ? Math.max(0, 1 - sum * factor) : Math.min(1, sum * factor); } }
  else { let b = x + 1 - a, c = 1e300, d = 1 / b, h = d; for (let i = 1; i <= 10000; i++) { const an = -i * (i - a); b += 2; d = an * d + b; if (Math.abs(d) < 1e-300) d = 1e-300; c = b + an / c; if (Math.abs(c) < 1e-300) c = 1e-300; d = 1 / d; const delta = d * c; h *= delta; if (Math.abs(delta - 1) < 2e-15) return complement ? Math.min(1, factor * h) : Math.max(0, 1 - factor * h); } }
  return fail('#NUM!', 'Incomplete gamma did not converge');
}
function betaFraction(a, b, x) {
  let c = 1, d = 1 - (a + b) * x / (a + 1); if (Math.abs(d) < 1e-300) d = 1e-300; d = 1 / d; let h = d;
  for (let m = 1; m <= 10000; m++) {
    for (const [half, aa] of [m * (b - m) * x / ((a + 2 * m - 1) * (a + 2 * m)), -(a + m) * (a + b + m) * x / ((a + 2 * m) * (a + 2 * m + 1))].entries()) {
      d = 1 + aa * d; if (Math.abs(d) < 1e-300) d = 1e-300; c = 1 + aa / c; if (Math.abs(c) < 1e-300) c = 1e-300; d = 1 / d; const delta = d * c; h *= delta;
      if (half === 1 && Math.abs(delta - 1) < 3e-15) return h;
    }
  }
  return fail('#NUM!', 'Incomplete beta did not converge');
}
function betaI(x, a, b) {
  if (x <= 0) return 0; if (x >= 1) return 1;
  const front = Math.exp(logGamma(a + b) - logGamma(a) - logGamma(b) + a * Math.log(x) + b * Math.log1p(-x));
  return x < (a + 1) / (a + b + 2) ? front * betaFraction(a, b, x) / a : 1 - front * betaFraction(b, a, 1 - x) / b;
}
function inverseCDF(p, fn, low, high) {
  for (let i = 0; i < 160; i++) { const middle = low + (high - low) / 2; if (middle === low || middle === high) return middle; if (fn(middle) < p) low = middle; else high = middle; } return (low + high) / 2;
}
const normalCDF = x => x === 0 ? .5 : x < 0 ? .5 * gammaP(.5, x * x / 2, true) : 1 - .5 * gammaP(.5, x * x / 2, true);
const chooseLog = (n, k) => logGamma(n + 1) - logGamma(k + 1) - logGamma(n - k + 1);
function paired(y, x) {
  const a = flatten(y), b = flatten(x); if (a.length !== b.length) fail('#N/A');
  const rows = []; for (let i = 0; i < a.length; i++) { if (isError(a[i])) throw a[i]; if (isError(b[i])) throw b[i]; if (typeof a[i] === 'number' && typeof b[i] === 'number') rows.push([a[i], b[i]]); }
  if (!rows.length) fail('#DIV/0!'); return rows;
}
function regression(y, x) {
  const a = paired(y, x), n = a.length, my = total(a.map(r => r[0])) / n, mx = total(a.map(r => r[1])) / n;
  let xx = 0, yy = 0, xy = 0; for (const [y, x] of a) { xx += (x - mx) ** 2; yy += (y - my) ** 2; xy += (x - mx) * (y - my); } return { n, mx, my, xx, yy, xy };
}
function weekends(value) {
  if (typeof scalar(value) === 'string') { const s = text(value); if (!/^[01]{7}$/.test(s)) fail('#VALUE!', 'Weekend mask must contain seven 0/1 digits starting Monday'); return Array.from(s, c => c === '1'); }
  const n = integer(value), out = Array(7).fill(false);
  if (n >= 1 && n <= 7) { out[(n + 4) % 7] = true; out[(n + 5) % 7] = true; }
  else if (n >= 11 && n <= 17) out[(n - 5) % 7] = true; else fail(); return out;
}
function textSide(after, source, delimiter, instance = 1, match = 0, matchEnd = 0, missing = error('#N/A')) {
  const value = text(source), delim = text(delimiter), n = integer(instance), mode = number(match), end = number(matchEnd);
  if (!n || ![0, 1].includes(mode) || ![0, 1].includes(end)) fail('#VALUE!');
  if (!delim) return n < 0 ? (after ? '' : value) : (after ? value : '');
  if (Math.abs(n) > value.length) fail('#VALUE!');
  // Keep original UTF-16 offsets: uppercasing can change length (for example ß).
  const escaped = delim.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), hits = [];
  const pattern = new RegExp(escaped, mode ? 'iyu' : 'yu');
  if (n > 0) {
    for (let i = 0; i <= value.length - delim.length;) {
      pattern.lastIndex = i; const match = pattern.exec(value);
      if (match) { hits.push(i); i += match[0].length; } else i++;
    }
    if (end) hits.push(value.length);
  } else {
    for (let i = value.length - delim.length; i >= 0;) {
      pattern.lastIndex = i; const match = pattern.exec(value);
      if (match) { hits.push(i); i -= match[0].length; } else i--;
    }
    if (end) hits.push(-delim.length);
  }
  const hit = hits[Math.abs(n) - 1]; if (hit == null) return missing;
  return after ? value.slice(Math.min(value.length, Math.max(0, hit + delim.length))) : value.slice(0, Math.max(0, hit));
}
function complex(value) {
  const s = text(value).replace(/\s/g, ''); const unit = /[ij]$/.test(s) ? s.at(-1) : 'i';
  if (!/[ij]$/.test(s)) { const r = Number(s); if (!s || !Number.isFinite(r)) fail(); return [r, 0, unit]; }
  const v = s.slice(0, -1), split = [...v].findLastIndex((c, i) => i > 0 && /[+-]/.test(c) && !/[eE]/.test(v[i - 1]));
  const a = split < 0 ? '0' : v.slice(0, split), b = split < 0 ? v : v.slice(split), imaginary = b === '' || b === '+' ? 1 : b === '-' ? -1 : Number(b), real = Number(a);
  if (!Number.isFinite(real) || !Number.isFinite(imaginary)) fail(); return [real, imaginary, unit];
}
function complexText(a, b, unit = 'i') {
  finite(a); finite(b); const clean = v => Number(v.toPrecision(15)); a = clean(a); b = clean(b);
  return !b ? String(a) : (!a ? '' : a) + (b > 0 && a ? '+' : b < 0 ? '-' : '') + (Math.abs(b) === 1 ? '' : Math.abs(b)) + unit;
}
function square(value) { const a = checkedMatrix(value).map(r => r.map(v => { if (typeof v !== 'number') fail('#VALUE!'); return v; })); if (a.length !== a[0].length || a.length > 128) fail('#VALUE!', 'Square matrix up to 128 rows required'); return a; }


export { fail, integer, positive, probability, numeric, total, finite, checkedMatrix, vector, boundArray, realDate, logGamma, gammaP, betaI, inverseCDF, normalCDF, chooseLog, regression, weekends, textSide, complex, complexText, square };
