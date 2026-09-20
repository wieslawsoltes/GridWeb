import {quoteSheet} from '../src/core.js';
/** Reference demonstrations create ordinary, editable worksheets with live formulas. */
function referenceTools() {
  const {workbook,grid}=window.gridweb;
  if(grid.ReadOnly)throw new Error('The grid is read-only');
  let summary;
  workbook.Transaction('Reference tools examples',()=>{
    const unique=base=>{let name=base,i=2;while(workbook.Worksheets.Get(name))name=base+' '+i++;return name;};
    const sources=['North','Central','South'].map((name,i)=>{
      const s=workbook.Worksheets.Add(unique('Reference '+name));
      s.GetRange('A1:B4').Values=[['Category','Amount'],['Hardware',10+i*30],['Software',20+i*30],['Services',30+i*30]];
      s.GetRange('A1:B1').SetStyle({fill:'#107c41',font:{bold:true,color:'#ffffff'}});s.SetColumnWidth(0,180);return s;
    });
    summary=workbook.Worksheets.Add(unique('Reference tools'));
    const span=quoteSheet(sources[0].Name+':'+sources[2].Name);
    summary.GetCell('A1').Value='LIVE REFERENCE TOOLS';summary.GetCell('A2').Value='Three-sheet total';summary.GetCell('B2').Formula=`=SUM(${span}!B2:B4)`;
    summary.GetRange('A5:B8').Values=[[1,10],[2,20],[3,30],[4,40]];
    const examples=[['Union','=SUM((A5:A8,B5:B8))'],['Intersection','=SUM(A5:B8 B6:C7)'],['Dynamic INDEX range','=SUM(B5:INDEX(B5:B8,3))'],['INDEX area 2','=INDEX((A5:A8,B5:B8),2,1,2)'],['Worksheets in span',`=SHEETS(${span}!B2)`],['Reference identity','=ISREF(INDEX(B5:B8,2))'],['Inclusive average','=AVERAGEA({2,TRUE,"text"})']];
    examples.forEach(([label,formula],i)=>{summary.GetCell(i+4,3).Value=label;summary.GetCell(i+4,4).Formula=formula;summary.GetCell(i+4,5).Value=formula;});
    summary.GetCell('A13').Value='TRY IT';summary.GetCell('A14').Value='Edit source values, then return here to see recalculation.';
    summary.GetCell('A15').Value='Move Reference Central after Reference South: the total changes from 450 to 300.';
    summary.GetCell('A16').Value='Undo restores the sheet order and total. Rename or remove endpoints to test reference repair.';
    summary.GetCell('A18').Value='This is a tested reference profile, not complete Excel compatibility.';
    summary.GetRange('A1:F1').SetStyle({fill:'#107c41',font:{bold:true,color:'#ffffff'}});
    summary.SetColumnWidth(0,180);summary.SetColumnWidth(1,110);summary.SetColumnWidth(2,30);summary.SetColumnWidth(3,190);summary.SetColumnWidth(4,100);summary.SetColumnWidth(5,430);
    for(const row of [14,15,16,18]){summary.GetRange(`A${row}:F${row}`).Merge();summary.GetRange(`A${row}:F${row}`).SetStyle({wrapText:true});summary.SetRowHeight(row-1,38);}
    summary.GetRange('F5:F11').SetStyle({font:{name:'Consolas',color:'#526660'}});summary.FreezePanes(1,0);
  });
  grid.Sheet=summary;grid.Select('B2');grid.Focus();
}
Object.assign(window.gridweb.commands,{referenceTools,moveWorksheet:()=>window.gridweb.grid.ShowMoveWorksheet()});
const ribbon=document.getElementById('ribbon');
function buttons(){
  if(ribbon.querySelector('[data-reference-tools]')||!['Home','Formulas'].includes(document.querySelector('[data-tab][aria-selected="true"]')?.dataset.tab))return;
  const group=document.createElement('div');group.dataset.referenceTools='true';group.style.cssText='display:flex;gap:5px;align-items:center;padding:6px';
  for(const [command,label] of [['referenceTools','Reference tools'],['moveWorksheet','Move worksheet']]){const button=document.createElement('button');button.type='button';button.dataset.command=command;button.textContent=label;button.style.padding='8px';group.append(button);}
  ribbon.append(group);
}
new MutationObserver(buttons).observe(ribbon,{childList:true});buttons();
