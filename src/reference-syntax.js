/** Shared lexical reference scanner. String literals and table selectors are opaque. */
const atom = "(?:'(?:[^']|'')+'|[A-Za-z_][\\w.]*)";
const cell = '\\$?[A-Za-z]{1,3}\\$?[1-9]\\d*';
const axis = '(?:\\$?[A-Za-z]{1,3}:\\$?[A-Za-z]{1,3}|\\$?[1-9]\\d*:\\$?[1-9]\\d*)';
const prefix = `(?:${atom}(?::${atom})?!)?`;
const pattern = new RegExp(`^(${prefix})(${cell}(?::${cell}(?![\\w.!(\\[]))?|${axis})(?![\\w.!(\\[])`);
const unquote = name => name.startsWith("'") ? name.slice(1, -1).replace(/''/g, "'") : name;
export function readReference(source) {
  const match = pattern.exec(source);
  if (!match) return null;
  const rawPrefix = match[1].slice(0, -1);
  let sheet = null, sheetEnd = null;
  if (rawPrefix) {
    // Both 'First sheet:Last sheet' and individually quoted endpoints are accepted.
    const parts = rawPrefix.match(/'(?:[^']|'')+'|[^:]+/g).map(unquote);
    const names = parts.length === 1 ? parts[0].split(':') : parts;
    if (names.length > 2 || names.some(n => !n)) return null;
    [sheet, sheetEnd = null] = names;
  }
  return {raw: match[0], address: match[2], sheet, sheetEnd};
}
export function referencePrefix(sheet, sheetEnd = null) {
  if (sheet == null) return '';
  const name = sheetEnd == null ? sheet : `${sheet}:${sheetEnd}`;
  return "'" + name.replace(/'/g, "''") + "'!";
}
export function mapFormulaReferences(formula, transform) {
  const source = String(formula); let out = '', i = 0;
  while (i < source.length) {
    if (source[i] === '"' || source[i] === '[') {
      const start = i++, quote = source[start] === '"'; let depth = 1;
      while (i < source.length && depth) {
        const ch = source[i++];
        if (quote && ch === '"') { if (source[i] === '"') i++; else depth = 0; }
        else if (!quote) { if (ch === '[') depth++; else if (ch === ']') depth--; }
      }
      out += source.slice(start, i); continue;
    }
    const ref = !/[\w.\\]/.test(source[i - 1] ?? '') ? readReference(source.slice(i)) : null;
    if (ref) { out += transform(ref, i, source); i += ref.raw.length; }
    else if (source[i] === "'") {
      // An already translated R1C1 prefix must not be scanned as an A1 axis.
      const start = i++;
      while (i < source.length) if (source[i++] === "'") { if (source[i] === "'") i++; else break; }
      out += source.slice(start, i);
    } else out += source[i++];
  }
  return out;
}

/** A qualified defined name is not an A1 cell, a table selector, or a 3-D span. */
export function readQualifiedName(source) {
  const m = /^(('(?:[^']|'')+'|[A-Za-z_][\w.]*)!)([A-Za-z_\\][\w.\\]*)(?![\w.\\])/.exec(source);
  return m ? {raw:m[0],sheet:unquote(m[2]),name:m[3]} : null;
}
