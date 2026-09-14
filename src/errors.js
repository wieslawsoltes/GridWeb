export class FormulaError extends Error {
  constructor(code, detail = '') { super(code); this.name = 'FormulaError'; this.code = code; this.detail = detail; }
  toString() { return this.code; }
  toJSON() { return { error: this.code, detail: this.detail }; }
}
export const error = (code, detail) => new FormulaError(code, detail);
export const isError = value => value instanceof FormulaError;
export const scalar = value => Array.isArray(value) ? scalar(value[0] ?? null) : value;
export function number(value) { value = scalar(value); if (isError(value)) throw value; if (value == null || value === '') return 0; if (value === true) return 1; if (value === false) return 0; const n = Number(value); if (!Number.isFinite(n)) throw error('#VALUE!'); return n; }
export function text(value) { value = scalar(value); if (isError(value)) throw value; return value == null ? '' : value === true ? 'TRUE' : value === false ? 'FALSE' : String(value); }
export function truth(value) { value = scalar(value); if (isError(value)) throw value; if (value == null || value === '' || value === 0) return false; if (typeof value === 'string') { if (/^true$/i.test(value)) return true; if (/^false$/i.test(value)) return false; throw error('#VALUE!'); } return !!value; }
export const flatten = value => Array.isArray(value) ? value.flat(Infinity) : [value];
export const matrix = value => Array.isArray(value) ? (Array.isArray(value[0]) ? value : [value]) : [[value]];
export function serialDate(date) { const t = (date.getTime() - Date.UTC(1899, 11, 31)) / 86400000; return t >= 60 ? t + 1 : t; }
export function fromSerial(serial) { return new Date(Date.UTC(1899, 11, 31) + (serial >= 60 ? serial - 1 : serial) * 86400000); }
