import { escapeHTML, safeColor } from './layout-base.js';
import { paginate } from './pagination.js';
import { createPrintPresentation } from './print-presentation.js';
import { chartToSVG } from './chart-svg.js';
import { contains } from './address.js';
export { AxisLayout, escapeHTML, safeColor } from './layout-base.js';
export { paginate } from './pagination.js';
export { chartToSVG } from './chart-svg.js';
function runs(items, axis, all) {
  const list = [], offsets = new Map(); let offset = 0;
  for (const i of all) { offsets.set(i, offset); offset += axis.Size(i); }
  for (const i of items) { const start = axis.Offset(i), end = start + axis.Size(i), previous = list.at(-1); if (previous?.end === start) previous.end = end; else list.push({ start, end, target: offsets.get(i) }); }
  return list;
}
function chartFragments(sheet, page, layout) {
  const rs = runs(page.contentRows, layout.rows, page.rows), cs = runs(page.contentColumns, layout.columns, page.columns), output = [];
  for (const chart of sheet.Charts) {
    const x = layout.columns.Offset(chart.column), y = layout.rows.Offset(chart.row); let svg;
    for (const r of rs) for (const c of cs) {
      const left = Math.max(x, c.start), top = Math.max(y, r.start), right = Math.min(x + chart.width, c.end), bottom = Math.min(y + chart.height, r.end);
      if (right <= left || bottom <= top) continue;
      svg ??= chartToSVG(sheet, chart);
      output.push(`<div class="print-chart" style="left:${c.target + left - c.start}px;top:${r.target + top - r.start}px;width:${right - left}px;height:${bottom - top}px"><div style="transform:translate(${x - left}px,${y - top}px)">${svg}</div></div>`);
    }
  }
  return output.join('');
}
export function printPagesHTML(sheet, options = {}) {
  const layout = paginate(sheet, options), present = createPrintPresentation(sheet);
  return layout.pages.map(page => {
    const skipped = new Set(), body = [];
    for (const r of page.rows) {
      let cells = '';
      for (const c of page.columns) {
        if (skipped.has(r + ':' + c)) continue;
        const merge = sheet._meta.merges.find(m => contains(m, r, c)); let rs = 1, cs = 1, cell = sheet.GetCell(r, c);
        if (merge) {
          const mr = page.rows.filter(i => i >= merge.r1 && i <= merge.r2), mc = page.columns.filter(i => i >= merge.c1 && i <= merge.c2);
          if (r !== mr[0] || c !== mc[0]) continue; rs = mr.length; cs = mc.length; cell = sheet.GetCell(merge.r1, merge.c1);
          for (const ri of mr) for (const ci of mc) if (ri !== r || ci !== c) skipped.add(ri + ':' + ci);
        }
        const { style, bar } = present(cell.Row, cell.Column, cell), font = style.font ?? {};
        const css = `background:${safeColor(style.fill, 'white')};color:${safeColor(font.color, '#202722')};font-weight:${font.bold ? '700' : '400'};font-style:${font.italic ? 'italic' : 'normal'};font-size:${Math.max(8, Math.min(72, font.size ?? 11))}pt;text-align:${['left', 'right', 'center'].includes(style.horizontalAlignment) ? style.horizontalAlignment : typeof cell.Value === 'number' ? 'right' : 'left'};vertical-align:${['top', 'bottom', 'middle'].includes(style.verticalAlignment) ? style.verticalAlignment : 'middle'};white-space:${style.wrapText ? 'pre-wrap' : 'pre'};text-decoration:${[font.underline ? 'underline' : '', font.strikethrough ? 'line-through' : ''].filter(Boolean).join(' ') || 'none'};${style.border ? `border:1px solid ${safeColor(style.border.color, '#82998a')};` : ''}`;
        const boxHeight = page.rows.filter(i => i >= r && (!merge ? i === r : i <= merge.r2)).reduce((n, i) => n + layout.rows.Size(i), 0) - 1;
        cells += `<td data-row="${r}" data-column="${c}" rowspan="${rs}" colspan="${cs}" style="${css}"><div class="print-cell-box" style="height:${Math.max(0, boxHeight)}px;justify-content:${style.verticalAlignment === 'top' ? 'flex-start' : style.verticalAlignment === 'bottom' ? 'flex-end' : 'center'}">${bar ? `<span class="print-data-bar" aria-hidden="true" style="width:${bar.ratio * 100}%;background:${safeColor(bar.color, '#21a366')}"></span>` : ''}<span class="print-cell-text">${escapeHTML(cell.Text)}</span></div></td>`;
      }
      body.push(`<tr style="height:${layout.rows.Size(r)}px">${cells}</tr>`);
    }
    const title = template => escapeHTML(String(template).replace(/&[PNAF]/g, t => ({ '&P': page.number, '&N': layout.pages.length, '&A': sheet.Name, '&F': sheet.Workbook.Name })[t]));
    const m = page.margins, contentHeight = page.rows.reduce((n, r) => n + layout.rows.Size(r), 0);
    return `<section class="print-page" style="width:${page.width}px;height:${page.height}px;padding:${m.top}px ${m.right}px ${m.bottom}px ${m.left}px"><header>${title(layout.config.header ?? '&F · &A')}</header><div class="print-content" style="zoom:${page.scale};height:${contentHeight}px"><table style="width:${page.columns.reduce((n,c)=>n+layout.columns.Size(c),0)}px"><colgroup>${page.columns.map(c => `<col style="width:${layout.columns.Size(c)}px">`).join('')}</colgroup><tbody>${body.join('')}</tbody></table>${options.includeCharts === false ? '' : chartFragments(sheet, page, layout)}</div><footer style="left:${m.left}px;bottom:${Math.max(0, m.bottom - 16)}px">${title(layout.config.footer ?? 'Page &P of &N')}</footer></section>`;
  }).join('');
}
export const printCSS = `*{box-sizing:border-box}body{margin:0;background:#dbe0dc;font:11pt system-ui;color:#202722}.print-page{background:white;margin:24px auto;position:relative;box-shadow:0 3px 20px #0002;overflow:hidden;break-after:page;print-color-adjust:exact;-webkit-print-color-adjust:exact}.print-page:last-child{break-after:auto}.print-page header{height:20px;font-size:10pt;color:#58665d}.print-page footer{position:absolute;font-size:9pt;color:#58665d}.print-content{position:relative}.print-content table{border-collapse:collapse;table-layout:fixed}.print-content td{position:relative;border:1px solid #dce4de;padding:0;overflow:hidden}.print-content td:focus{outline:2px solid #107c41;white-space:pre-wrap}.print-cell-box{display:flex;flex-direction:column;position:relative;overflow:hidden;padding:3px 6px}.print-cell-text{display:block;width:100%;position:relative}.print-data-bar{position:absolute;left:0;top:3px;bottom:3px;opacity:.2;pointer-events:none}.print-chart{position:absolute;overflow:hidden;pointer-events:none}.print-chart svg{display:block}@media print{body{background:white}.print-page{margin:0;box-shadow:none}@page{size:auto;margin:0}}`;
export function createPrintDocument(sheet, options = {}) {
  const layout = paginate(sheet, options), page = layout.pages[0];
  return `<!doctype html><html><head><meta charset="utf-8"><title>${escapeHTML(sheet.Name)}</title><style>${printCSS}@page{size:${page.width}px ${page.height}px;margin:0}</style></head><body>${printPagesHTML(sheet, options)}</body></html>`;
}
