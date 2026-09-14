import './app-base.js';
import {createExcelApi} from '../src/office.js';
import {chartToSVG,createPrintDocument,paginate} from '../src/layout.js';

const element=(tag,text)=>{const e=document.createElement(tag);if(text!=null)e.textContent=text;return e;};
let noticeTimer;
const note=message=>{const notice=document.getElementById('toast');notice.textContent=message;clearTimeout(noticeTimer);noticeTimer=setTimeout(()=>notice.textContent='',6000);};
const download=(name,content,type)=>{const url=URL.createObjectURL(new Blob([content],{type})),a=element('a');a.href=url;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);};

async function officeExample(){
  const {workbook,grid}=window.gridweb;
  let title='API example',suffix=1;while(workbook.Worksheets.Get(title))title='API example '+suffix++;
  let result;
  await createExcelApi(workbook).run(async context=>{
    const sheet=context.workbook.worksheets.add(title);
    sheet.getRange('A1:D4').values=[['Region','Units','Unit price','Revenue'],['North',12,8.5,''],['South',9,12,''],['Total','','','']];
    sheet.getRange('D2:D4').formulas=[['=B2*C2'],['=B3*C3'],['=SUM(D2:D3)']];
    sheet.getRange('A1:D1').set({format:{font:{bold:true,color:'#ffffff'},fill:{color:'#107c41'}}});
    sheet.getRange('C2:D4').numberFormat=[['0.00','0.00'],['0.00','0.00'],[null,'0.00']];
    sheet.getRange('A1:D4').format.autofitColumns();
    const total=sheet.getRange('D4');total.load('values');sheet.activate();await context.sync();result=total.values[0][0];
  });
  grid.Sheet=title;grid.Select('D4');note(`Office-style batch: ${result} total; one undo restores the workbook.`);
}
function exportVector(){
 const {grid}=window.gridweb,chart=grid.Sheet.Charts[0];
 if(!chart){note('Insert a chart on this worksheet first.');return;}
 download('GridWeb-chart.svg',chartToSVG(grid.Sheet,chart),'image/svg+xml');note('First worksheet chart exported as vector SVG.');
}
function printOptions(){
 const {grid}=window.gridweb,sheet=grid.Sheet,dialog=element('dialog'),form=element('form'),heading=element('header'),title=element('h2','Advanced print preview'),close=element('button','×');
 close.type='button';close.setAttribute('aria-label','Close advanced print');close.onclick=()=>dialog.close();heading.append(title,close);form.append(heading);
 const description=element('p','These options apply to this print preview. They do not overwrite saved worksheet settings.');description.className='hint';form.append(description);
 const fields=element('div');fields.style.cssText='display:grid;grid-template-columns:1fr 1fr;gap:12px';const inputs={};
 const add=(name,label,value,options)=>{const l=element('label',label),input=element(options?'select':'input');input.name=name;input.setAttribute('aria-label',label);
   if(options)for(const text of options){const o=element('option',text);o.value=text;input.append(o);}else if(typeof value==='number'){input.type='number';input.min='0';input.max='100';}else input.type='text';
   input.value=value;input.style.cssText='display:block;width:100%;padding:7px;margin-top:5px';l.append(input);fields.append(l);inputs[name]=input;};
 add('paper','Paper',sheet._meta.print.paper??'A4',['A4','Letter','A3','Legal']);add('orientation','Orientation',sheet._meta.print.orientation??'landscape',['portrait','landscape']);
 add('area','Print range (blank = used area)','');add('pageOrder','Page order','overThenDown',['overThenDown','downThenOver']);
 add('fitToWidthPages','Fit to pages wide (0 = scale)',1);add('fitToHeightPages','Fit to pages tall (0 = scale)',0);
 add('repeatRows','Repeat leading rows',1);add('repeatColumns','Repeat leading columns',0);
 add('header','Header tokens: &F &A &P &N','&F · &A');add('footer','Footer','Page &P of &N');
 form.append(fields);const error=element('p');error.setAttribute('role','alert');const footer=element('footer'),cancel=element('button','Cancel'),submit=element('button','Open preview');cancel.type='button';cancel.onclick=()=>dialog.close();submit.type='submit';submit.className='primary';footer.append(cancel,submit);form.append(error,footer);dialog.append(form);document.body.append(dialog);
 form.onsubmit=e=>{e.preventDefault();try{
   const options=Object.fromEntries(Object.entries(inputs).map(([k,v])=>[k,v.type==='number'?+v.value:v.value]));if(!options.area)delete options.area;
   const pages=paginate(sheet,options),html=createPrintDocument(sheet,options);
   const preview=window.open('','_blank');if(!preview)throw Error('Allow a pop-up to open the printable preview.');
   preview.opener=null;preview.document.open();preview.document.write(html);preview.document.close();note(`${pages.pages.length} printable page(s), ${Math.round(pages.config.scale*100)}% scale. Use the preview window’s Print command.`);dialog.close();
 }catch(e){error.textContent=e.message;}};
 dialog.addEventListener('close',()=>dialog.remove(),{once:true});dialog.showModal();
}
const commands=window.gridweb.commands;
Object.assign(commands,{officeExample,exportVector,advancedPrint:printOptions});
const ribbon=document.getElementById('ribbon');
function enrichRibbon(){
 if(ribbon.querySelector('[data-extra-tools]'))return;
 const tab=document.querySelector('[data-tab][aria-selected="true"]')?.dataset.tab;
 if(!['Page Layout','View','Insert'].includes(tab))return;
 const group=element('div');group.dataset.extraTools='true';group.style.cssText='display:flex;gap:6px;align-items:center;border-left:1px solid #ddd;padding:8px';
 for(const [command,label]of tab==='Page Layout'?[['advancedPrint','Advanced print'],['exportVector','Chart SVG']]:tab==='Insert'?[['exportVector','Chart SVG']]:[['officeExample','Office API example']]){
  const b=element('button',label);b.type='button';b.dataset.command=command;b.style.cssText='padding:10px 12px;white-space:normal';group.append(b);
 }
 ribbon.append(group);
}
new MutationObserver(enrichRibbon).observe(ribbon,{childList:true});enrichRibbon();
// A user-visible API entry point, bound to the current workbook rather than a stale snapshot.
window.gridwebExtensions={createExcelApi,chartToSVG,paginate,createPrintDocument,officeExample,printOptions};
