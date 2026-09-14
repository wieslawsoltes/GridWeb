import {contains} from '../src/address.js';
import {aggregates} from '../src/pivots/aggregate.js';
const el=(tag,text)=>{const e=document.createElement(tag);if(text!=null)e.textContent=text;return e;};
const notify=message=>document.getElementById('toast').textContent=message;
function activePivot(){const {workbook,grid}=window.gridweb;return [...workbook.PivotTables].find(p=>p.Destination.Worksheet===grid.Sheet&&p.OutputRange&&contains(p.OutputRange.Bounds,grid.ActiveCell.Row,grid.ActiveCell.Column))??workbook.PivotTables.Get(0);}
function configure(existing=null){
 const {workbook,grid}=window.gridweb,dialog=el('dialog'),form=el('form'),header=el('header');dialog.id='pivot-reports-dialog';header.append(el('h2',existing?'Edit pivot report':'Create pivot report'));const close=el('button','×');close.type='button';close.setAttribute('aria-label','Close pivot reports');close.onclick=()=>dialog.close();header.append(close);form.append(header);
 const fields={};
 const field=(name,label,value,choices=null)=>{const wrapper=el('label',label),input=el(choices?'select':'input');input.name=name;input.setAttribute('aria-label',label);if(choices)for(const [key,text]of choices){const option=el('option',text);option.value=key;input.append(option);}input.value=value;input.style.cssText='display:block;width:100%;padding:7px;margin:4px 0 10px';wrapper.append(input);form.append(wrapper);fields[name]=input;return input;};
 let name='PivotReport',i=1;while(workbook.PivotTables.Get(name))name='PivotReport'+i++;
 field('name','Pivot name',existing?.Name??name);
 field('sourceSheet','Source sheet',existing?.Source.Worksheet.Id??grid.Sheet.Id,[...workbook.Worksheets].map(s=>[s.Id,s.Name]));
 field('source','Source range',existing?.Source.Address??(grid.SelectionRange.Count>1?grid.Selection:grid.Sheet.UsedRange.Address));
 field('targetSheet','Destination sheet (created if absent)',existing?.Destination.Worksheet.Name??'Pivot reports');field('target','Destination cell',existing?.Destination.Address??'A1');
 const rows=field('rows','Row fields (comma-separated)',existing?.Options.rows.join(', ')??''),cols=field('columns','Column fields (comma-separated)',existing?.Options.columns.join(', ')??'');
 const measure=field('value','Value field',existing?.Options.values[0].column??''),aggregate=field('aggregate','Summarize by',existing?.Options.values[0].aggregate??'sum',aggregates.map(a=>[a,a]));
 const total=el('label'),totalInput=el('input');totalInput.type='checkbox';totalInput.checked=existing?.Options.columnGrandTotals??true;totalInput.setAttribute('aria-label','Grand totals');total.append(totalInput,document.createTextNode(' Grand totals'));form.append(total);
 const advanced=el('details');advanced.append(el('summary','Additional measures and filters'));const config=el('textarea');config.setAttribute('aria-label','Advanced pivot options');config.rows=5;config.style.width='100%';config.value=existing?JSON.stringify({values:existing.Options.values,filters:existing.Options.filters},null,2):'{}';advanced.append(el('p','Optional JSON: values and filters arrays. These override the single value field above.'),config);form.append(advanced);
 const inspect=()=>{try{const s=workbook._sheets.find(s=>s.Id===fields.sourceSheet.value),matrix=s.GetRange(fields.source.value).Values,headers=matrix[0];if(!rows.value)rows.value=String(headers[0]??'');if(!measure.value)measure.value=String(headers.find((_,c)=>matrix.slice(1).some(r=>typeof r[c]==='number'))??headers[1]??'');rows.placeholder=headers.join(', ');measure.placeholder=headers.join(', ');}catch{}};inspect();fields.source.onchange=inspect;fields.sourceSheet.onchange=()=>{rows.value='';measure.value='';inspect();};
 const error=el('p');error.setAttribute('role','alert');const footer=el('footer'),apply=el('button',existing?'Update pivot':'Create pivot');apply.type='submit';apply.className='primary';footer.append(apply);form.append(error,footer);dialog.append(form);document.body.append(dialog);
 form.onsubmit=e=>{e.preventDefault();try{
  const source=workbook._sheets.find(s=>s.Id===fields.sourceSheet.value).GetRange(fields.source.value),split=s=>s.split(',').map(s=>s.trim()).filter(Boolean),extra=JSON.parse(config.value||'{}');
  if(!extra||typeof extra!=='object'||Array.isArray(extra))throw new Error('Advanced options must be an object');
  const options={rows:split(rows.value),columns:split(cols.value),values:[{column:measure.value,aggregate:aggregate.value}],columnGrandTotals:totalInput.checked,rowGrandTotals:totalInput.checked,...extra};let report;
  workbook.Transaction('Configure pivot report',()=>{const target=workbook.Worksheets.Get(fields.targetSheet.value)??workbook.Worksheets.Add(fields.targetSheet.value),destination=target.GetRange(fields.target.value);report=existing?existing.Update({name:fields.name.value,source,destination,options}):workbook.PivotTables.Add(fields.name.value,source,destination,options);});
  grid.Sheet=report.Destination.Worksheet;grid.Select(report.Destination.Address);notify('Pivot '+report.Name+' created. Edit source values, then Refresh pivots.');dialog.close();
 }catch(e){error.textContent=e.message;}};
 dialog.addEventListener('close',()=>dialog.remove(),{once:true});dialog.showModal();
}
function drill(){const p=activePivot();if(!p)throw new Error('Create a managed pivot first');const {workbook,grid}=window.gridweb,b=p.OutputRange.Bounds,rows=p.DrillDown(grid.ActiveCell.Row-b.r1,grid.ActiveCell.Column-b.c1);let s;
 workbook.Transaction('Pivot drill-down',()=>{let name='Pivot details',i=1;while(workbook.Worksheets.Get(name))name='Pivot details '+i++;s=workbook.Worksheets.Add(name);s.GetRange('A1').Resize(rows.length,rows[0].length).Values=rows;});grid.Sheet=s;grid.Select('A1');notify((rows.length-1)+' source records');
}
Object.assign(window.gridweb.commands,{managedPivot:()=>configure(),editManagedPivot:()=>{const p=activePivot();if(!p)throw new Error('Create a managed pivot first');configure(p);},refreshPivots:()=>{window.gridweb.workbook.PivotTables.RefreshAll();notify('Managed pivots refreshed');},drillPivot:drill});
const ribbon=document.getElementById('ribbon');function buttons(){if(ribbon.querySelector('[data-pivot-tools]')||!['Data','Insert'].includes(document.querySelector('[data-tab][aria-selected="true"]')?.dataset.tab))return;const group=el('div');group.dataset.pivotTools='true';group.style.cssText='display:flex;gap:5px;align-items:center;border-left:1px solid #ddd;padding:6px';for(const [command,label]of [['managedPivot','Pivot reports'],['editManagedPivot','Edit pivot'],['refreshPivots','Refresh pivots'],['drillPivot','Show details']]){const b=el('button',label);b.type='button';b.dataset.command=command;b.style.padding='8px';group.append(b);}ribbon.append(group);}
new MutationObserver(buttons).observe(ribbon,{childList:true});buttons();
window.gridwebPivots={configure,activePivot};
