import {quoteSheet} from '../src/core.js';
/** Examples use ordinary named formulas, editable source cells and shared undo. */
function namedFormulas() {
  const {workbook,grid}=window.gridweb;if(grid.ReadOnly)throw new Error('The grid is read-only');if(!grid.CommitEdit())throw new Error('Complete the current edit');
  let summary;
  workbook.Transaction('Scoped name examples',()=>{
    const unique=base=>{let name=base,n=2;while(workbook.Worksheets.Get(name))name=base+' '+n++;return name;};
    const source=['North','South'].map((label,i)=>{
      const s=workbook.Worksheets.Add(unique('Names '+label));
      s.GetRange('A1:C4').Values=[['Item','Revenue','Tax rate'],['Hardware',100*(i+1),i===0?.1:.2],['Software',200*(i+1),null],['Services',300*(i+1),null]];
      s.Names.Create('Revenue','=B2:B4',{comment:'Revenue on '+s.Name});s.Names.Create('TaxRate','=C2',{comment:'Local rate; the same name exists in both regions'});
      s.Names.Create('TaxDue','=LAMBDA(amount,amount*TaxRate)',{comment:'Named LAMBDA resolves this worksheet’s rate'});
      s.GetCell('A6').Value='Local total';s.GetCell('B6').Formula='=SUM(Revenue)';s.GetCell('A7').Value='Tax';s.GetCell('B7').Formula='=TaxDue(SUM(Revenue))';
      s.GetRange('A1:C1').SetStyle({fill:'#107c41',font:{bold:true,color:'#ffffff'}});s.SetColumnWidth(0,180);s.SetColumnWidth(1,140);s.SetColumnWidth(2,140);s.GetCell('C2').Style={numberFormat:'0%'};return s;
    });
    summary=workbook.Worksheets.Add(unique('Named formulas'));const n=quoteSheet(source[0].Name),s=quoteSheet(source[1].Name);
    summary.GetCell('A1').Value='SCOPED NAMES · LIVE FORMULAS';summary.GetRange('A1:D1').Merge();summary.GetRange('A1:D1').SetStyle({fill:'#107c41',font:{bold:true,color:'#ffffff'}});
    const examples=[['North revenue',`=SUM(${n}!Revenue)`],['South revenue',`=SUM(${s}!Revenue)`],['North tax',`=${n}!TaxDue(SUM(${n}!Revenue))`],['South tax',`=${s}!TaxDue(SUM(${s}!Revenue))`],['LET keeps a reference',`=LET(data,${n}!Revenue,ISREF(data))`],['Named INDEX range',`=SUM(INDEX(${n}!Revenue,1):INDEX(${n}!Revenue,2))`],['LAMBDA keeps a reference',`=LAMBDA(data,ISREF(data))(${s}!Revenue)`],['Two-area name inputs',`=SUM((${n}!Revenue,${s}!Revenue))`]];
    examples.forEach(([title,formula],i)=>{summary.GetCell(i+3,0).Value=title;summary.GetCell(i+3,1).Formula=formula;summary.GetCell(i+3,2).Value=formula;});
    summary.GetRange('A14:D14').Merge();summary.GetCell('A14').Value='Edit source revenues or rates. The two Revenue / TaxRate names remain independent.';
    summary.GetRange('A15:D15').Merge();summary.GetCell('A15').Value='Open Name manager to search, edit, rename, delete or navigate. Undo restores a rename and its formula repairs.';
    summary.SetRowHeight(13,42);summary.SetRowHeight(14,42);summary.GetRange('A14:D15').SetStyle({wrapText:true});summary.SetColumnWidth(0,260);summary.SetColumnWidth(1,120);summary.SetColumnWidth(2,570);summary.SetColumnWidth(3,40);summary.FreezePanes(1,0);
  });
  grid.Sheet=summary;grid.Select('B4');grid.Focus();
}
Object.assign(window.gridweb.commands,{namedFormulas,namesFromSelection:()=>window.gridweb.grid.ShowCreateNamesFromSelection(),names:()=>window.gridweb.grid.ShowDefineName(),nameManager:()=>window.gridweb.grid.ShowNameManager()});
const ribbon=document.getElementById('ribbon');
function buttons(){
  if(ribbon.querySelector('[data-named-formulas]')||!['Home','Formulas'].includes(document.querySelector('[data-tab][aria-selected="true"]')?.dataset.tab))return;
  const group=document.createElement('div');group.dataset.namedFormulas='true';group.style.cssText='display:flex;gap:5px;align-items:center;padding:6px';
  for(const[command,label]of[['namedFormulas','Named formulas'],['namesFromSelection','Create from selection'],['nameManager','Name manager'],['names','Define name']]){
    if(ribbon.querySelector(`[data-command="${command}"]`))continue;
    const button=document.createElement('button');button.type='button';button.dataset.command=command;button.textContent=label;button.style.padding='8px';group.append(button);
  }
  ribbon.append(group);
}
new MutationObserver(buttons).observe(ribbon,{childList:true});buttons();
