import { fromSerial, isError } from './errors.js';
const literals = format => format.replace(/"([^"]*)"/g, '$1').replace(/\\(.)/g, '$1').replace(/_.|\*./g, '');
/** Common Excel-style formats; unsupported tokens remain a documented compatibility boundary. */
export function formatValue(value, format = 'General', locale = 'en-US') {
  if (isError(value)) return value.code;
  if (value == null) return '';
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (typeof value === 'string') { const s = format.split(';')[3]; return s ? literals(s).replace('@', value) : value; }
  if (!Number.isFinite(value)) return '#NUM!';
  const sections = String(format).split(';'); let section = sections[value < 0 && sections.length > 1 ? 1 : value === 0 && sections.length > 2 ? 2 : 0];
  if (section === '') return '';
  section = section.replace(/\[(?:Red|Blue|Green|Yellow|Black|White|Magenta|Cyan|Color\d+|[<>=][^\]]*)\]/gi, '');
  if (section === 'General' || section === '@') return Math.abs(value) >= 1e11 || (Math.abs(value) < 1e-9 && value !== 0) ? value.toExponential(5) : String(Number(value.toPrecision(12)));
  if (/[ymdhis]/i.test(section.replace(/"[^"]*"/g, '').replace(/\[[^\]]+\]/g, '')) && !/[Ee][+-]?0/.test(section)) {
    if (value < 0 || value >= 2958466) return '########'; const d = fromSerial(value), day = Math.floor(value) === 60 ? 29 : d.getUTCDate();
    let inTime = false; const hasAm = /AM\/PM/i.test(section);
    return section.replace(/"([^"]*)"|\[(h+|m+|s+)\]|AM\/PM|yyyy|yy|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s/gi, (token, quoted, elapsed, offset) => {
      if (quoted != null) return quoted;
      if (elapsed) return String(Math.floor(value * (elapsed[0].toLowerCase() === 'h' ? 24 : elapsed[0].toLowerCase() === 'm' ? 1440 : 86400))).padStart(elapsed.length, '0');
      const t = token.toLowerCase(), pad = n => String(n).padStart(2, '0');
      if (t === 'am/pm') return d.getUTCHours() < 12 ? 'AM' : 'PM';
      if (t === 'yyyy') return String(d.getUTCFullYear()); if (t === 'yy') return pad(d.getUTCFullYear() % 100);
      if (t === 'mmmm' || t === 'mmm') return new Intl.DateTimeFormat(locale, { month: t === 'mmmm' ? 'long' : 'short', timeZone: 'UTC' }).format(d);
      if (t === 'dddd' || t === 'ddd') return new Intl.DateTimeFormat(locale, { weekday: t === 'dddd' ? 'long' : 'short', timeZone: 'UTC' }).format(d);
      if (t[0] === 'h') { inTime = true; const h = hasAm ? d.getUTCHours() % 12 || 12 : d.getUTCHours(); return t.length === 2 ? pad(h) : String(h); }
      if (t[0] === 'm') { const minute = inTime || /^:s/i.test(section.slice(offset + token.length)); const n = minute ? d.getUTCMinutes() : d.getUTCMonth() + 1; return t.length === 2 ? pad(n) : String(n); }
      if (t[0] === 's') return t.length === 2 ? pad(d.getUTCSeconds()) : String(d.getUTCSeconds());
      if (t[0] === 'd') return t.length === 2 ? pad(day) : String(day); return token;
    });
  }
  const match = /[0#?][0#?,]*(?:\.[0#?]+)?(?:[eE][+-]?0+)?/.exec(section);
  if (!match) return literals(section);
  const mask = match[0], percent = section.includes('%'), decimals = (mask.split('.')[1] ?? '').replace(/[eE].*/, ''), minimumFractionDigits = (decimals.match(/0/g) ?? []).length, maximumFractionDigits = Math.min(15, decimals.length);
  let n = value * (percent ? 100 : 1); if (sections.length > 1 && value < 0) n = Math.abs(n);
  let formatted;
  if (/[eE]/.test(mask)) formatted = n.toExponential(maximumFractionDigits).replace('e', 'E');
  else formatted = new Intl.NumberFormat(locale, { useGrouping: mask.includes(','), minimumFractionDigits: Math.min(minimumFractionDigits, maximumFractionDigits), maximumFractionDigits, minimumIntegerDigits: Math.max(1, Math.min(21, (mask.split('.')[0].match(/0/g) ?? []).length)) }).format(n);
  return literals(section.slice(0, match.index)) + formatted + literals(section.slice(match.index + mask.length));
}
