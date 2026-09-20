/** Reusable defined-name UI. Untrusted names/formulas are assigned as text, never HTML. */
import {dialog,element,select} from './control-editing.js';
import {referencePrefix} from './reference-syntax.js';
const input=(form,title,value='')=>{const label=element('label',title),node=element('input');node.value=value;label.append(node);form.append(label);return node;};
const asFormula=v=>typeof v==='string'?v.startsWith('=')?v:'="'+v.replace(/"/g,'""')+'"':v?.error?'='+v.error:v===null?'=""':'='+String(v).toUpperCase();
function ownerOf(book,id){const owner=id==null?book:book.Worksheets.Get(id);if(!owner)throw new Error('Worksheet is no longer attached');return owner;}
function writeGuard(grid,book){if(grid.Workbook!==book)throw new Error('Workbook changed. Reopen the dialog.');if(grid.ReadOnly)throw new Error('The grid is read-only');if(!grid.CommitEdit())throw new Error('Complete the current edit first');}
export function showDefineName(grid,definition=null) {
  const book=grid.Workbook,original=definition?structuredClone(definition):null;
  const ui=dialog(grid,definition?'Edit defined name':'Define name','Save',()=>{
    writeGuard(grid,book);const owner=ownerOf(book,scope.value||null);
    if(!formula.value.startsWith('='))throw new Error('Refers to must start with =. Quote text constants.');
    if(original&&JSON.stringify(owner.Names.GetDefinition(original.name))!==JSON.stringify(original))throw new Error('Defined name changed. Reopen the editor.');
    book.Transaction(definition?'Edit defined name':'Define name',()=>{
      if(original){owner.Names.Update(original.name,formula.value,{comment:comment.value});owner.Names.Rename(original.name,name.value);}
      else owner.Names.Create(name.value,formula.value,{comment:comment.value});
    });
  });
  const name=input(ui.form,'Name',definition?.name??'MyRange');name.required=true;name.maxLength=255;
  const scope=select(ui.form,'Scope',[['','Workbook'],...book.Worksheets.items.map(s=>[s.Id,s.Name])],definition?.sheetId??'');scope.disabled=!!definition;
  const absolute=referencePrefix(grid.Sheet.Name)+grid.SelectionRange.Address.replace(/([A-Z]+)(\d+)/g,'$$$1$$$2');
  const formula=input(ui.form,'Refers to',definition?asFormula(definition.value):'='+absolute);formula.required=true;
  const comment=input(ui.form,'Comment',definition?.comment??'');comment.maxLength=255;
  ui.form.append(element('small','Names are case-insensitive. Worksheet names override workbook names. Scope is fixed when editing; changes are one undoable transaction.'));
  ui.open();return ui.modal;
}
export function showNameManager(grid) {
  const book=grid.Workbook;
  const ui=dialog(grid,'Name manager','New name',()=>{writeGuard(grid,book);showDefineName(grid);return false;});
  const query=input(ui.form,'Search names');query.type='search';
  const filter=select(ui.form,'Scope filter',[['all','All scopes'],['global','Workbook'],['local','Worksheets']]);
  const status=element('p'),results=element('div');status.setAttribute('role','status');results.className='results';
  const failure=element('p');failure.setAttribute('role','alert');ui.form.append(status,results,failure);
  function refresh(){
    results.replaceChildren();if(grid.Workbook!==book){status.textContent='Workbook changed. Reopen Name manager.';return;}
    const items=book.GetDefinedNames().filter(d=>!d.hidden&&(filter.value==='all'||(filter.value==='global')===(d.sheetId===null))&&(d.name+' '+d.scope+' '+d.comment).toLowerCase().includes(query.value.toLowerCase())).sort((a,b)=>a.name.localeCompare(b.name)||a.scope.localeCompare(b.scope));
    status.textContent=`${items.length} defined names${items.length>200?' (showing the first 200; refine the search)':''}. Hidden names are omitted.`;
    for(const d of items.slice(0,200)){
      const row=element('section');row.style.cssText='display:grid;gap:6px;border-bottom:1px solid var(--grid-line);padding:8px 0';
      row.append(element('strong',d.name+' — '+d.scope),element('small',asFormula(d.value)),element('small',d.comment));
      const buttons=element('div');buttons.style.cssText='display:flex;gap:6px;flex-wrap:wrap';
      for(const [caption,action]of [
        ['Edit',()=>{writeGuard(grid,book);const current=ownerOf(book,d.sheetId).Names.GetDefinition(d.name);if(!current)throw new Error('Name was deleted. Reopen Name manager.');showDefineName(grid,current);}],
        ['Delete',()=>{writeGuard(grid,book);const owner=ownerOf(book,d.sheetId);if(JSON.stringify(owner.Names.GetDefinition(d.name))!==JSON.stringify(d))throw new Error('Name changed. Reopen Name manager.');if(buttons.dataset.confirm!==d.name){buttons.dataset.confirm=d.name;remove.textContent='Confirm delete';return;}owner.Names.Remove(d.name);refresh();}],
        ['Go to',()=>{if(grid.Workbook!==book)throw new Error('Workbook changed');grid.Select(ownerOf(book,d.sheetId).Names.GetRange(d.name).FullAddress);ui.close();}]
      ]){const button=element('button',caption);button.type='button';button.setAttribute('aria-label',caption+' '+d.name+' in '+d.scope);button.disabled=caption!=='Go to'&&grid.ReadOnly;button.onclick=()=>{failure.textContent='';try{action();}catch(e){failure.textContent=e.message;}};buttons.append(button);}
      const remove=buttons.children[1];row.append(buttons);results.append(row);
    }
  }
  query.oninput=refresh;filter.onchange=refresh;refresh();ui.open();return ui.modal;
}

export function showCreateNamesFromSelection(grid) {
  const book=grid.Workbook,revision=book.Revision,source=grid.SelectionRange;
  const ui=dialog(grid,'Create names from selection','Create names',()=>{
    writeGuard(grid,book);if(book.Revision!==revision)throw new Error('Workbook changed. Reopen the dialog to review its labels.');
    ownerOf(book,scope.value||null).Names.CreateFromSelection(source,{topRow:top.checked,bottomRow:bottom.checked,leftColumn:left.checked,rightColumn:right.checked,overwrite:overwrite.checked});
  });
  ui.form.append(element('p','Use labels in '+source.FullAddress+'. Generated names use absolute data ranges; spaces and unsupported characters become underscores.'));
  const scope=select(ui.form,'Scope',[['','Workbook'],...book.Worksheets.items.map(s=>[s.Id,s.Name])]);
  const flag=(title,checked=false)=>{const label=element('label'),box=element('input');box.type='checkbox';box.checked=checked;label.className='check';label.append(box,document.createTextNode(title));ui.form.append(label);return box;};
  const top=flag('Top row',true),left=flag('Left column'),bottom=flag('Bottom row'),right=flag('Right column'),overwrite=flag('Replace existing names in this scope');
  ui.form.append(element('small','Blank labels are skipped. Duplicate generated names or invalid boundaries reject the entire operation. Changes are one undoable transaction.'));
  ui.open();return ui.modal;
}
