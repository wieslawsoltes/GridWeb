import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook, printPagesHTML, createPrintDocument, paginate, chartToSVG } from '../src/core.js';
function sheet() { const b = new Workbook(), s = b.ActiveWorksheet; s.GetRange('A1:C4').Values = [['Item','Value','Plan'],['North',10,20],['South',20,10],['West',15,15]]; return s; }
for (const type of ['column','bar','line','area','pie','scatter']) test('inert SVG uses the same '+type+' chart renderer', () => {
  const s = sheet(), c = s.AddChart('A1:C4', { type, title:'<unsafe> & "quoted"' }); const svg = chartToSVG(s,c);
  assert.ok(svg.startsWith('<svg')); assert.ok(svg.includes('&lt;unsafe&gt;')); assert.ok(!svg.includes('NaN')); assert.ok(!svg.includes('<script'));
});
test('page HTML prints charts and conditional formatting', () => {
  const s = sheet(); s.AddChart('A1:C4',{column:4,row:1}); s.AddConditionalFormat('B2:B4',{type:'dataBar'}); s.AddConditionalFormat('B2:B4',{type:'cellValue',criteria:'>15',style:{fill:'#ff0000'}});
  const html = printPagesHTML(s); assert.ok(html.includes('print-chart')); assert.ok(html.includes('print-data-bar')); assert.ok(html.includes('background:#ff0000'));
});
test('charts expand default area and split across page boundaries', () => {
  const s = sheet(); s.AddChart('A1:C4',{column:6,row:30,width:600,height:400});
  const layout = paginate(s,{orientation:'portrait'}); assert.ok(layout.config.area.r2>=46); assert.ok(layout.pages.length>1);
  assert.ok((printPagesHTML(s,{orientation:'portrait'}).match(/class="print-chart"/g)||[]).length>1);
});
test('fit-to-width, titles, page order and manual breaks are executable', () => {
  const s=sheet(); s.GetCell('J90').Value=1;
  const normal=paginate(s,{area:'A1:J90',orientation:'portrait'}), fit=paginate(s,{area:'A1:J90',orientation:'portrait',fitToWidthPages:1,repeatRows:1});
  assert.ok(fit.config.scale<normal.config.scale); assert.ok(fit.pages.every(p=>p.columns.length===10));
  assert.ok(fit.pages.slice(1).every(p=>p.rows[0]===0));
  const manual=paginate(s,{area:'A1:J90',rowBreaks:[20],columnBreaks:[3],repeatColumns:1,pageOrder:'downThenOver'});
  assert.ok(manual.pages[1].contentRows[0]===20); assert.ok(manual.pages.some(p=>p.columns[0]===0&&p.contentColumns[0]===3));
});
test('print headers, tables, text styles and exact page dimensions render',()=>{
  const s=sheet();s.AddTable('A1:C4','Sales');s.GetCell('A2').Style={font:{underline:true,strikethrough:true}};
  const html=createPrintDocument(s,{paper:'Letter',header:'&A &P/&N',footer:'&F'});assert.ok(html.includes('@page{size:1056px 816px'));assert.ok(html.includes('text-decoration:underline line-through'));assert.ok(html.includes('background:#107c41'));assert.ok(!html.includes('footer{position:fixed}'));
});
test('invalid pagination options fail instead of producing corrupt pages',()=>{
 const s=sheet();for(const o of [{scale:NaN},{margins:{left:10000}},{area:'A0'},{repeatColumns:-1},{fitToWidthPages:1.5},{paper:'Unknown'}])assert.throws(()=>paginate(s,o));
});
