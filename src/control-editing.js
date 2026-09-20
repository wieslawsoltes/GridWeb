/** Editing dialogs shared by every GridWebElement host. Model mutations stay in the engine. */
const css = `dialog{box-sizing:border-box;width:min(440px,calc(100vw - 24px));max-height:calc(100dvh - 24px);overflow:auto;border:1px solid var(--grid-line);border-radius:12px;padding:20px;background:var(--grid-bg);color:var(--grid-fg);font:14px system-ui;box-shadow:0 12px 60px #0005}dialog::backdrop{background:#0005}form{display:grid;gap:12px}h2,p{margin:0}h2{font-size:20px}label{display:grid;gap:5px}label.check{display:flex;align-items:center;gap:8px}input,select,button{box-sizing:border-box;font:inherit;color:inherit;background:var(--grid-bg);border:1px solid var(--grid-line);border-radius:5px;padding:9px;min-width:0}input[type=checkbox]{width:18px;height:18px}button{cursor:pointer;min-height:40px}button.primary{background:var(--grid-accent);color:var(--grid-bg);font-weight:600}footer{display:flex;gap:8px;justify-content:flex-end}small,.hint{font-size:12px;line-height:1.5;opacity:.85}[role=alert]{color:#c44444;font-size:13px;white-space:pre-wrap}[role=alert]:empty{display:none}.results{display:grid;gap:5px;max-height:220px;overflow:auto}.results button{text-align:left}.results:empty{display:none}`;
let sequence = 0;
const element = (tag, text) => { const value = document.createElement(tag); if (text != null) value.textContent = text; return value; };
function dialog(grid, title, applyText, apply) {
  for (const existing of grid._editingDialogs ?? []) existing.close();
  const modal = element('dialog'), form = element('form'), heading = element('h2', title), error = element('p');
  heading.id = 'grid-editing-' + ++sequence; modal.setAttribute('aria-labelledby', heading.id);
  error.setAttribute('role', 'alert'); const style = element('style', css); form.append(heading);
  const footer = element('footer'), cancel = element('button', 'Cancel'), submit = element('button', applyText);
  const restoreFocus = () => { if (grid.isConnected && ![...(grid._editingDialogs ?? [])].some(d => d.open)) grid.Focus(); };
  const close = () => { modal.close(); restoreFocus(); };
  cancel.type = 'button'; cancel.onclick = close; submit.type = 'submit'; submit.className = 'primary';
  footer.append(cancel, submit); modal.append(style, form);
  (grid._editingDialogs ??= new Set()).add(modal);
  modal.addEventListener('close', () => { grid._editingDialogs.delete(modal); modal.remove(); restoreFocus(); }, {once: true});
  form.onsubmit = event => {
    event.preventDefault(); error.textContent = '';
    try { if (apply() !== false) close(); } catch (failure) { error.textContent = failure.message ?? String(failure); }
  };
  // The dialog is a sibling of the keyboard grid, so typing does not invoke grid shortcuts.
  const open = () => { form.append(error, footer); grid.shadowRoot.append(modal); modal.showModal(); };
  return {form, open, modal, close};
}
function select(form, title, choices, initial) {
  const label = element('label', title), input = element('select'); input.setAttribute('aria-label', title);
  for (const [value, caption] of choices) { const option = element('option', caption); option.value = value; input.append(option); }
  if (initial != null) input.value = initial; label.append(input); form.append(label); return input;
}
function checkbox(form, title) {
  const label = element('label'), input = element('input'); input.type = 'checkbox'; label.className = 'check';
  label.append(input, document.createTextNode(title)); form.append(label); return input;
}
function numeric(form, title, value = '') {
  const label = element('label', title), input = element('input'); input.type = 'number'; input.step = 'any'; input.value = String(value);
  label.append(input); form.append(label); return input;
}
function optionalNumber(input) { return input.value.trim() === '' ? undefined : input.valueAsNumber; }
export function showPasteSpecial(grid) {
  const ui = dialog(grid, 'Paste special', 'Paste', () => grid.PasteSpecial({mode: mode.value, operation: operation.value, transpose: transpose.checked, skipBlanks: skip.checked}));
  const hint = element('p', grid._clipboard?.snapshot && !grid._clipboard.cut
    ? `Copied ${grid._clipboard.snapshot.rows} × ${grid._clipboard.snapshot.columns} cells. Destination: ${grid.Selection}.`
    : 'Copy a range in this grid first. External text uses the normal Paste command.'); hint.className = 'hint'; ui.form.append(hint);
  const mode = select(ui.form, 'Paste content', [['all','All'],['values','Values'],['formulas','Formulas'],['formats','Formats'],['comments','Comments'],['validation','Validation'],['columnWidths','Column widths'],['valuesAndNumberFormats','Values and number formats'],['formulasAndNumberFormats','Formulas and number formats']]);
  const operation = select(ui.form, 'Operation', [['none','None'],['add','Add'],['subtract','Subtract'],['multiply','Multiply'],['divide','Divide']]);
  operation.onchange = () => { if (operation.value !== 'none' && !['values','valuesAndNumberFormats'].includes(mode.value)) mode.value = 'values'; };
  const skip = checkbox(ui.form, 'Skip blank source cells'), transpose = checkbox(ui.form, 'Transpose rows and columns');
  ui.form.append(element('small', 'Arithmetic uses values. Merged-cell copying and transposed custom validation are not supported. Every paste is one undoable transaction.'));
  ui.open(); return ui.modal;
}
export function showFillSeries(grid) {
  const ui = dialog(grid, 'Fill series', 'Fill', () => grid.FillSeries({type: type.value, direction: direction.value, dateUnit: dateUnit.value, step: step.valueAsNumber, start: optionalNumber(start), stop: optionalNumber(stop)}));
  ui.form.append(element('p', `Fill ${grid.Selection}. Blank start uses the first cell in each row or column, in the selected direction.`));
  const direction = select(ui.form, 'Direction', [['down','Down'],['right','Right'],['up','Up'],['left','Left']]);
  const type = select(ui.form, 'Series type', [['linear','Linear'],['growth','Growth'],['date','Date']]);
  const dateUnit = select(ui.form, 'Date unit', [['day','Day'],['weekday','Weekday'],['month','Month'],['year','Year']]);
  dateUnit.disabled = true; type.onchange = () => dateUnit.disabled = type.value !== 'date';
  const start = numeric(ui.form, 'Start value (optional)'), step = numeric(ui.form, 'Step', 1), stop = numeric(ui.form, 'Stop value (optional)');
  step.required = true; ui.form.append(element('small', 'Dates use workbook serial numbers. Month and year steps preserve the original day where possible. Stop leaves the remaining cells unchanged.'));
  ui.open(); return ui.modal;
}
export function showGoToSpecial(grid) {
  const ui = dialog(grid, 'Go to special', 'Find cells', () => {
    const matches = grid.FindSpecialCells(type.value, valueType.value ? {valueTypes: [valueType.value]} : {});
    results.replaceChildren(); status.textContent = `${matches.length} matching cells${matches.length > 500 ? ' (showing the first 500)' : ''}. Choose a cell to navigate.`;
    for (const match of matches.slice(0, 500)) {
      const button = element('button', `${match.Address} — ${match.GetCell(0, 0).Text.slice(0, 100) || '(blank)'}`); button.type = 'button';
      button.onclick = () => { grid.Select(match.Address); ui.close(); }; results.append(button);
    }
    return false;
  });
  ui.form.append(element('p', 'Search the selection, or the used range when one cell is selected. Results navigate to individual cells.'));
  const type = select(ui.form, 'Cell kind', [['formulas','Formulas'],['constants','Constants'],['blanks','Blanks'],['errors','Errors'],['comments','Comments'],['validation','Validation'],['visible','Visible cells']]);
  const valueType = select(ui.form, 'Value type', [['','All value types'],['numbers','Numbers'],['text','Text'],['logical','Logical'],['errors','Errors']]);
  type.onchange = () => { valueType.disabled = !['formulas','constants'].includes(type.value); if (valueType.disabled) valueType.value = ''; };
  const status = element('p'), results = element('div'); status.setAttribute('role', 'status'); results.className = 'results'; ui.form.append(status, results);
  ui.open(); return ui.modal;
}
export function showMoveWorksheet(grid) {
  const workbook=grid.Workbook, sheets=workbook.Worksheets.items, order=sheets.map(s=>s.Id).join(',');
  const ui=dialog(grid,'Move worksheet','Move',()=>{
    if(grid.Workbook!==workbook||workbook.Worksheets.items.map(s=>s.Id).join(',')!==order)throw new Error('Worksheet order changed. Reopen this dialog before moving.');
    grid.MoveWorksheet(Number(position.value),sheet.value);
  });
  ui.form.append(element('p','Choose the final position. Formulas using a 3-D sheet range recalculate when the sheet order changes.'));
  const sheet=select(ui.form,'Worksheet',sheets.map(s=>[s.Id,s.Name]),grid.Sheet.Id);
  const position=select(ui.form,'Final position',sheets.map((s,i)=>[String(i),`${i+1} — ${s.Name}`]),String(sheets.indexOf(grid.Sheet)));
  sheet.onchange=()=>{position.value=String(sheets.findIndex(s=>s.Id===sheet.value));};
  ui.form.append(element('small','Moving a range endpoint past its opposite endpoint shrinks that reference. Undo restores both formulas and worksheet order.'));
  ui.open();return ui.modal;
}
