import { AxisLayout } from './layout-base.js';
import { parseRange, boundedCells, MAX_ROWS, MAX_COLUMNS } from './address.js';
function integer(value, fallback, max) { const n = value ?? fallback; if (!Number.isInteger(n) || n < 0 || n > max) throw new RangeError('Invalid pagination option'); return n; }
function groups(first, last, axis, available, repeat, breaks) {
  const repeated = Array.from({ length: repeat }, (_, i) => first + i).filter(i => axis.Size(i));
  const repeatSize = repeated.reduce((n, i) => n + axis.Size(i), 0), result = []; let items = [], size = 0;
  for (let i = first; i <= last; i++) {
    const s = axis.Size(i); if (!s) continue;
    if (items.length && (breaks.has(i) || size + s > available - (result.length ? repeatSize : 0))) { result.push({ items, repeated: result.length ? repeated.filter(n => !items.includes(n)) : [] }); items = []; size = 0; }
    items.push(i); size += s;
  }
  if (items.length) result.push({ items, repeated: result.length ? repeated.filter(n => !items.includes(n)) : [] });
  return result.length ? result : [{ items: [], repeated: [] }];
}
/** Pixel geometry is shared by page preview and print output. Values are CSS pixels at 96 DPI. */
export function paginate(sheet, options = {}) {
  const config = { ...sheet._meta.print, ...options }, { rows, columns } = AxisLayout.ForSheet(sheet);
  let bounds = config.area ? typeof config.area === 'string' ? parseRange(config.area) : { ...config.area } : { ...sheet.UsedRange.Bounds };
  if (!config.area && options.includeCharts !== false) for (const c of sheet.Charts) {
    bounds.r1 = Math.min(bounds.r1, c.row); bounds.c1 = Math.min(bounds.c1, c.column);
    bounds.r2 = Math.max(bounds.r2, rows.IndexAt(rows.Offset(c.row) + c.height - 1e-5)); bounds.c2 = Math.max(bounds.c2, columns.IndexAt(columns.Offset(c.column) + c.width - 1e-5));
  }
  for (const [key, max] of [['r1', MAX_ROWS], ['r2', MAX_ROWS], ['c1', MAX_COLUMNS], ['c2', MAX_COLUMNS]]) if (!Number.isInteger(bounds[key]) || bounds[key] < 0 || bounds[key] >= max) throw new RangeError('Invalid print area');
  if (bounds.r1 > bounds.r2 || bounds.c1 > bounds.c2) throw new RangeError('Invalid print area'); boundedCells(bounds);
  const paper = { A4: [794, 1123], Letter: [816, 1056], A3: [1123, 1587], Legal: [816, 1344] }[config.paper ?? 'A4'];
  if (!paper) throw new RangeError('Unknown paper size'); const [width, height] = config.orientation === 'landscape' ? [...paper].reverse() : paper;
  const margins = { top: 40, right: 40, bottom: 40, left: 40, ...config.margins };
  for (const v of Object.values(margins)) if (!Number.isFinite(v) || v < 0) throw new RangeError('Invalid margins');
  const aw = width - margins.left - margins.right, ah = height - margins.top - margins.bottom - 20;
  if (aw < 32 || ah < 32) throw new RangeError('Margins leave no printable space');
  let scale = config.scale ?? 1; if (!Number.isFinite(scale) || scale < .1 || scale > 4) throw new RangeError('Print scale must be 0.1–4');
  const rr = Math.min(integer(config.repeatRows, 0, 100), bounds.r2 - bounds.r1 + 1), rc = Math.min(integer(config.repeatColumns, 0, 100), bounds.c2 - bounds.c1 + 1);
  const breaks = (a, max) => new Set((a ?? []).map(v => integer(v, 0, max - 1)));
  const rb = breaks(config.rowBreaks, MAX_ROWS), cb = breaks(config.columnBreaks, MAX_COLUMNS);
  const split = s => ({ rg: groups(bounds.r1, bounds.r2, rows, ah / s, rr, rb), cg: groups(bounds.c1, bounds.c2, columns, aw / s, rc, cb) });
  const fw = integer(config.fitToWidthPages, 0, 100), fh = integer(config.fitToHeightPages, 0, 100);
  if (fw || fh) {
    const fits = s => { const g = split(s); return (!fw || g.cg.length <= fw) && (!fh || g.rg.length <= fh); };
    if (!fits(.1)) throw new RangeError('Fit-to-page target cannot be met at 10% scale');
    if (!fits(scale)) { let lo = .1, hi = scale; for (let i = 0; i < 40; i++) { const mid = (lo + hi) / 2; if (fits(mid)) lo = mid; else hi = mid; } scale = lo; }
  }
  const { rg, cg } = split(scale); if (rg.length * cg.length > 100) throw new RangeError('Print preview is limited to 100 pages; narrow the print area');
  if (config.pageOrder && !['overThenDown', 'downThenOver'].includes(config.pageOrder)) throw new RangeError('Unknown page order');
  const pairs = config.pageOrder === 'downThenOver' ? cg.flatMap(c => rg.map(r => [r, c])) : rg.flatMap(r => cg.map(c => [r, c]));
  const pages = pairs.map(([r, c], i) => ({ rows: r.repeated.concat(r.items), columns: c.repeated.concat(c.items), contentRows: r.items, contentColumns: c.items, width, height, scale, number: i + 1, margins }));
  return { pages, rows, columns, config: { ...config, scale, margins, area: bounds } };
}
