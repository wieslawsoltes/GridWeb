import { MAX_ROWS, MAX_COLUMNS, contains, intersects, boundedCells } from './address.js';
/** Sparse Fenwick index: overrides are O(log N), positions O(log N), hit testing O(log² N). */
export class AxisLayout {
  constructor(count,defaultSize){this.Count=count;this.DefaultSize=defaultSize;this._deltas=new Map();this._sizes=new Map();}
  SetSize(index,size){if(!Number.isInteger(index)||index<0||index>=this.Count||!Number.isFinite(size)||size<0)throw new RangeError('Invalid axis dimension');const old=this.Size(index),delta=size-old;if(size===this.DefaultSize)this._sizes.delete(index);else this._sizes.set(index,size);for(let i=index+1;i<=this.Count;i+=i&-i){const next=(this._deltas.get(i)??0)+delta;if(next)this._deltas.set(i,next);else this._deltas.delete(i);}}
  Size(index){return this._sizes.get(index)??this.DefaultSize;}
  Offset(index){index=Math.max(0,Math.min(this.Count,index));let sum=index*this.DefaultSize;for(let i=index;i>0;i-=i&-i)sum+=this._deltas.get(i)??0;return sum;}
  get TotalSize(){return this.Offset(this.Count);}
  IndexAt(offset){if(offset<=0)offset=0;let low=0,high=this.Count;while(low<high){const mid=Math.floor((low+high)/2);if(this.Offset(mid+1)<=offset)low=mid+1;else high=mid;}return Math.min(this.Count-1,low);}
  static ForSheet(sheet){const rows=new AxisLayout(MAX_ROWS,24),columns=new AxisLayout(MAX_COLUMNS,100);for(const [i,v]of Object.entries(sheet._meta.rows))rows.SetSize(+i,v.hidden?0:v.size??24);for(const r of sheet._filtered)rows.SetSize(r,0);for(const[i,v]of Object.entries(sheet._meta.columns))columns.SetSize(+i,v.hidden?0:v.size??100);return{rows,columns};}
}
export const escapeHTML = value => String(value??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
export function safeColor(value,fallback='transparent'){return typeof value==='string'&&/^(?:#[0-9a-f]{3,8}|(?:rgb|hsl)a?\([0-9.,% /-]+\)|[a-z]{1,20})$/i.test(value)?value:fallback;}
export function paginate(sheet,options={}){
  const config={...sheet._meta.print,...options},bounds=config.area??sheet.UsedRange.Bounds;boundedCells(bounds);
  const paper=config.paper==='Letter'?[816,1056]:[794,1123];if(config.orientation==='landscape')paper.reverse();
  const scale=Math.max(.1,Math.min(4,config.scale??1)),width=(paper[0]-80)/scale,height=(paper[1]-100)/scale,{rows,columns}=AxisLayout.ForSheet(sheet),repeat=Math.min(config.repeatRows??0,bounds.r2-bounds.r1+1);
  const split=(first,last,axis,available,repeatSize=0)=>{const groups=[];let items=[],size=0;for(let i=first;i<=last;i++){const s=axis.Size(i);if(!s)continue;if(items.length&&size+s>available-(groups.length?repeatSize:0)){groups.push(items);items=[];size=0;}items.push(i);size+=s;}if(items.length)groups.push(items);return groups.length?groups:[[]];};
  const colGroups=split(bounds.c1,bounds.c2,columns,width),rowGroups=split(bounds.r1,bounds.r2,rows,height,rows.Offset(bounds.r1+repeat)-rows.Offset(bounds.r1));
  if(colGroups.length*rowGroups.length>100)throw new RangeError('Print preview is limited to 100 pages; narrow the print area');
  const pages=[];for(let r=0;r<rowGroups.length;r++)for(const cols of colGroups){const list=r?Array.from({length:repeat},(_,i)=>bounds.r1+i).filter(i=>rows.Size(i)>0&&!rowGroups[r].includes(i)).concat(rowGroups[r]):rowGroups[r];pages.push({rows:list,columns:cols,width:paper[0],height:paper[1],scale,number:pages.length+1});}return{pages,rows,columns,config};
}
export function printPagesHTML(sheet,options={}){
  const layout=paginate(sheet,options);
  return layout.pages.map(page=>{
    const skipped=new Set(),body=[];
    for(const r of page.rows){let cells='';for(const c of page.columns){if(skipped.has(r+':'+c))continue;const merge=sheet._meta.merges.find(m=>contains(m,r,c));let rs=1,cs=1,cell=sheet.GetCell(r,c);if(merge){const mr=page.rows.filter(i=>i>=merge.r1&&i<=merge.r2),mc=page.columns.filter(i=>i>=merge.c1&&i<=merge.c2);if(r!==mr[0]||c!==mc[0])continue;rs=mr.length;cs=mc.length;cell=sheet.GetCell(merge.r1,merge.c1);for(const ri of mr)for(const ci of mc)if(ri!==r||ci!==c)skipped.add(ri+':'+ci);}
      const style=cell.Style,font=style.font??{};const css=`background:${safeColor(style.fill,'white')};color:${safeColor(font.color,'#202722')};font-weight:${font.bold?'700':'400'};font-style:${font.italic?'italic':'normal'};font-size:${Math.max(8,Math.min(72,font.size??11))}pt;text-align:${['left','right','center'].includes(style.horizontalAlignment)?style.horizontalAlignment:typeof cell.Value==='number'?'right':'left'};white-space:${style.wrapText?'pre-wrap':'pre'};`;
      cells+=`<td data-row="${r}" data-column="${c}" rowspan="${rs}" colspan="${cs}" style="${css}">${escapeHTML(cell.Text)}</td>`;
    }body.push(`<tr style="height:${layout.rows.Size(r)}px">${cells}</tr>`);}
    return `<section class="print-page" style="width:${page.width}px;min-height:${page.height}px"><header>${escapeHTML(sheet.Workbook.Name)} · ${escapeHTML(sheet.Name)}</header><div style="zoom:${page.scale}"><table><colgroup>${page.columns.map(c=>`<col style="width:${layout.columns.Size(c)}px">`).join('')}</colgroup><tbody>${body.join('')}</tbody></table></div><footer>Page ${page.number} of ${layout.pages.length}</footer></section>`;
  }).join('');
}
export const printCSS=`*{box-sizing:border-box}body{margin:0;background:#dbe0dc;font:11pt system-ui;color:#202722}.print-page{background:white;margin:24px auto;padding:40px;position:relative;box-shadow:0 3px 20px #0002;overflow:hidden;break-after:page}.print-page header{font-size:10pt;color:#58665d;margin-bottom:16px}.print-page footer{position:absolute;bottom:24px;left:40px;font-size:9pt;color:#58665d}table{border-collapse:collapse;table-layout:fixed}td{border:1px solid #dce4de;padding:3px 6px;max-width:480px;overflow:hidden}td:focus{outline:2px solid #107c41;white-space:pre-wrap}@media print{body{background:white}.print-page{margin:0;box-shadow:none}footer{position:fixed}@page{size:auto;margin:0}}`;
export function createPrintDocument(sheet,options={}){return`<!doctype html><html><head><meta charset="utf-8"><title>${escapeHTML(sheet.Name)}</title><style>${printCSS}</style></head><body>${printPagesHTML(sheet,options)}</body></html>`;}
