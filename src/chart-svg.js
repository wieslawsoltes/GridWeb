import { drawChart } from './charts.js';
import { escapeHTML, safeColor } from './layout-base.js';
let serial = 0;
const n = value => { if (!Number.isFinite(value)) throw new RangeError('Non-finite SVG coordinate'); return +value.toFixed(5); };
/** Canvas subset used by drawChart, recorded as inert, resolution-independent SVG. */
class SvgContext {
  constructor() { this.parts = []; this.path = []; this.stack = []; this.clips = []; this.id = 'gwchart' + serial++; this.fillStyle = '#000'; this.strokeStyle = '#000'; this.lineWidth = 1; this.globalAlpha = 1; this.font = '10px system-ui'; this.textAlign = 'left'; this.textBaseline = 'alphabetic'; }
  save() { this.stack.push({ fillStyle: this.fillStyle, strokeStyle: this.strokeStyle, lineWidth: this.lineWidth, globalAlpha: this.globalAlpha, font: this.font, textAlign: this.textAlign, textBaseline: this.textBaseline, clips: [...this.clips] }); }
  restore() { Object.assign(this, this.stack.pop()); }
  beginPath() { this.path = []; this.point = null; }
  moveTo(x, y) { this.path.push(`M${n(x)} ${n(y)}`); this.point = [x, y]; }
  lineTo(x, y) { this.path.push(`L${n(x)} ${n(y)}`); this.point = [x, y]; }
  closePath() { this.path.push('Z'); }
  rect(x, y, w, h) { this.moveTo(x, y); this.lineTo(x + w, y); this.lineTo(x + w, y + h); this.lineTo(x, y + h); this.closePath(); }
  arc(x, y, r, start, end) {
    const delta = Math.max(0, Math.min(Math.PI * 2, end - start)); if (!delta || !r) return;
    const sx = x + r * Math.cos(start), sy = y + r * Math.sin(start);
    if (!this.point) this.moveTo(sx, sy); else this.lineTo(sx, sy);
    const pieces = delta >= Math.PI * 2 - 1e-10 ? 2 : 1;
    for (let i = 1; i <= pieces; i++) { const angle = start + delta * i / pieces; this.path.push(`A${n(r)} ${n(r)} 0 ${delta / pieces > Math.PI ? 1 : 0} 1 ${n(x + r * Math.cos(angle))} ${n(y + r * Math.sin(angle))}`); }
    this.point = [x + r * Math.cos(end), y + r * Math.sin(end)];
  }
  emit(value) { this.parts.push(this.clips.map(id => `<g clip-path="url(#${id})">`).join('') + value + '</g>'.repeat(this.clips.length)); }
  clip() { const id = this.id + '-' + this.parts.length; this.parts.push(`<defs><clipPath id="${id}"><path d="${this.path.join(' ')}"/></clipPath></defs>`); this.clips.push(id); }
  fill() { this.emit(`<path d="${this.path.join(' ')}" fill="${escapeHTML(safeColor(this.fillStyle))}" opacity="${n(this.globalAlpha)}"/>`); }
  stroke() { this.emit(`<path d="${this.path.join(' ')}" fill="none" stroke="${escapeHTML(safeColor(this.strokeStyle))}" stroke-width="${n(this.lineWidth)}" opacity="${n(this.globalAlpha)}"/>`); }
  rectangle(x, y, w, h, stroke) {
    if (w < 0) { x += w; w = -w; } if (h < 0) { y += h; h = -h; }
    this.emit(`<rect x="${n(x)}" y="${n(y)}" width="${n(w)}" height="${n(h)}" fill="${stroke ? 'none' : escapeHTML(safeColor(this.fillStyle))}"${stroke ? ` stroke="${escapeHTML(safeColor(this.strokeStyle))}" stroke-width="${n(this.lineWidth)}"` : ''} opacity="${n(this.globalAlpha)}"/>`);
  }
  fillRect(...args) { this.rectangle(...args, false); }
  strokeRect(...args) { this.rectangle(...args, true); }
  measureText(value) { return { width: Array.from(String(value)).length * (parseFloat(/([\d.]+)px/.exec(this.font)?.[1]) || 10) * .56 }; }
  fillText(value, x, y, maxWidth) {
    const size = parseFloat(/([\d.]+)px/.exec(this.font)?.[1]) || 10;
    const measured = this.measureText(value).width;
    const baseline = this.textBaseline === 'middle' ? 'central' : this.textBaseline === 'top' ? 'hanging' : 'auto';
    this.emit(`<text x="${n(x)}" y="${n(y)}" fill="${escapeHTML(safeColor(this.fillStyle))}" font-family="system-ui,sans-serif" font-size="${n(size)}" font-weight="${/600|bold/.test(this.font) ? '600' : '400'}" text-anchor="${this.textAlign === 'center' ? 'middle' : this.textAlign === 'right' ? 'end' : 'start'}" dominant-baseline="${baseline}"${maxWidth > 0 && measured > maxWidth ? ` textLength="${n(maxWidth)}" lengthAdjust="spacingAndGlyphs"` : ''}>${escapeHTML(value)}</text>`);
  }
}
export function chartToSVG(sheet, chart, options = {}) {
  const width = chart.width, height = chart.height;
  if (!(width > 0 && height > 0 && width <= 2000 && height <= 2000)) throw new RangeError('Chart dimensions out of bounds');
  const ctx = new SvgContext(); drawChart(ctx, sheet, chart, { x: 0, y: 0, width, height }, options);
  return `<svg xmlns="http://www.w3.org/2000/svg" role="img" aria-label="${escapeHTML(chart.title)}" width="${n(width)}" height="${n(height)}" viewBox="0 0 ${n(width)} ${n(height)}"><title>${escapeHTML(chart.title)}</title>${ctx.parts.join('')}</svg>`;
}
