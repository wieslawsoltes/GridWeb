/** Sample commands use the same public control/model APIs as embedded applications. */
const el = (tag, text) => { const node = document.createElement(tag); if (text != null) node.textContent = text; return node; };
function examples() {
  const {workbook, grid} = window.gridweb; let sheet;
  workbook.Transaction('Calculation and editing examples', () => {
    let name = 'Calculation tools', suffix = 2; while (workbook.Worksheets.Get(name)) name = 'Calculation tools ' + suffix++;
    sheet = workbook.Worksheets.Add(name);
    sheet.GetRange('A1:D7').Values = [
      ['Region','Owner','Sales','Units'],['North','Ada',120,3],['South','Ben',90,2],
      ['North','Ada',200,5],['West','Cara',160,4],['North','Ben',120,3],['South','Ada',80,2]
    ];
    sheet.GetRange('F1:G2').Values = [['Region','Sales'],['North','>=120']];
    sheet.GetCell('F4').Value = 'Database total'; sheet.GetCell('G4').Formula = '=DSUM(A1:D7,"Sales",F1:G2)';
    sheet.GetCell('F5').Value = 'Generated multiplication table'; sheet.GetCell('F6').Formula = '=MAKEARRAY(4,4,LAMBDA(row,col,row*col))';
    sheet.GetCell('A10').Value = 'CALCULATION AND EDITING';
    sheet.GetRange('A11:B15').Values = [['Optional parameter',null],['Direct LAMBDA',null],['Ignore errors',null],['Database count',null],['Database average',null]];
    sheet.GetCell('B11').Formula = '=LAMBDA(x,y,IF(ISOMITTED(y),x*2,x+y))(7,)';
    sheet.GetCell('B12').Formula = '=LAMBDA(x,x^2)(12)';
    sheet.GetCell('B13').Formula = '=AGGREGATE(9,6,{1,#N/A,3})';
    sheet.GetCell('B14').Formula = '=DCOUNT(A1:D7,"Sales",F1:G2)';
    sheet.GetCell('B15').Formula = '=DAVERAGE(A1:D7,"Sales",F1:G2)';
    sheet.GetCell('D11').Value = 'Month-end date series';
    sheet.GetRange('D12:D17').FillSeries({type:'date',dateUnit:'month',start:new Date('2026-01-31T00:00:00Z')});
    sheet.GetRange('D12:D17').Format.NumberFormat = 'yyyy-mm-dd';
    sheet.GetCell('F12').Value = 'Growth series'; sheet.GetRange('F13:F18').FillSeries({type:'growth',start:2,step:2});
    sheet.GetCell('A18').Value = 'COPY / PASTE PRACTICE'; sheet.GetRange('A19:B22').Values = [['Item','Price'],['Coffee',4],['Tea',3],['Total',null]];
    sheet.GetCell('B22').Formula = '=SUM(B20:B21)'; sheet.GetRange('B20:B22').Format.NumberFormat = '0.00';
    sheet.GetCell('D20').Value = 'Copy A19:B22, select D23, then Paste special.';
    sheet.GetCell('D21').Value = 'Use Values, Transpose, or Formats without changing source data.';
    sheet.GetCell('D22').Value = 'Fill series and Go to special are reusable grid dialogs.';
    for (const address of ['A1:D1','F1:G1','A10:D10','A18:D18','A19:B19']) sheet.GetRange(address).SetStyle({fill:'#107c41',font:{bold:true,color:'#ffffff'}});
    sheet.SetColumnWidth(0,190); sheet.SetColumnWidth(1,120); sheet.SetColumnWidth(3,190);sheet.SetColumnWidth(5,190);sheet.FreezePanes(1,0);
  });
  grid.Sheet = sheet; grid.Select('A19:B22'); grid.Focus();
}
Object.assign(window.gridweb.commands, {
  pasteSpecial: () => window.gridweb.grid.ShowPasteSpecial(),
  fillSeries: () => window.gridweb.grid.ShowFillSeries(),
  goToSpecial: () => window.gridweb.grid.ShowGoToSpecial(),
  calculationTools: examples
});
const ribbon = document.getElementById('ribbon');
function buttons() {
  if (ribbon.querySelector('[data-editing-tools]') || !['Home','Data','Formulas'].includes(document.querySelector('[data-tab][aria-selected="true"]')?.dataset.tab)) return;
  const group = el('div'); group.dataset.editingTools = 'true'; group.style.cssText = 'display:flex;gap:5px;align-items:center;border-left:1px solid #ddd;padding:6px';
  for (const [command, label] of [['pasteSpecial','Paste special'],['fillSeries','Fill series'],['goToSpecial','Go to special'],['calculationTools','Calculation tools']]) {
    const button = el('button', label); button.type = 'button'; button.dataset.command = command; button.style.padding = '8px'; group.append(button);
  }
  ribbon.append(group);
}
new MutationObserver(buttons).observe(ribbon,{childList:true}); buttons();
