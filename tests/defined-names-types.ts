import {Workbook,DefinedNameCollection,DefinedNameDefinition,DefinedNameOptions,MAX_DEFINED_NAMES,validateDefinedName} from '@wieslawsoltes/gridweb';
import {GridWebElement} from '@wieslawsoltes/gridweb/controls';
import {createExcelApi,NamedItem} from '@wieslawsoltes/gridweb/office';
const b=new Workbook(),s=b.ActiveWorksheet,options:DefinedNameOptions={comment:'Local name',hidden:false,baseAddress:'A1',contextSheetId:s.Id};
s.Names.Create('Local','=A1',options);s.Names.Update('Local','=A2');s.Names.Rename('Local','Other');const item:DefinedNameDefinition|undefined=s.Names.GetDefinition('Other');const all:DefinedNameDefinition[]=b.GetDefinedNames();const names:DefinedNameCollection=b.Names;const range=s.Names.GetRange('Other');s.Names.Evaluate('Other',{sheet:s,row:0,col:1});validateDefinedName('Valid');const limit:10000=MAX_DEFINED_NAMES;
b.DefineName('Rate',2,{comment:'Constant'});const json=b.ToJSON();json.nameMetadata?.forEach(n=>console.log(n.name));json.sheets[0].names?.forEach(([name,value])=>console.log(name,value));
const grid=new GridWebElement();grid.Workbook=b;grid.DefineName('UiName','=A1',{},s);const dialog:HTMLDialogElement=grid.ShowNameManager();grid.ShowDefineName();grid.ShowCreateNamesFromSelection();s.Names.CreateFromSelection(s.GetRange('A1:B3'),{topRow:true,leftColumn:true});
void createExcelApi(b).run(async ctx=>{const sheet=ctx.workbook.worksheets.getActiveWorksheet(),name:NamedItem=sheet.names.add('Local',sheet.getRange('A1'),'Local range');name.load('formula,scope');await ctx.sync();name.formula='=A2';name.comment='Updated';name.visible=true;const local=sheet.names.load('items/name');await ctx.sync();const label:string=local.items[0].name;name.getRange().load('values');name.delete();void label;});
// @ts-expect-error hidden is boolean
b.Names.Create('Invalid',1,{hidden:'yes'});
// @ts-expect-error scope cannot be mutated on a definition collection
b.Names.Scope='South';
// @ts-expect-error unsupported primitive direct Office add argument
createExcelApi(b).createRequestContext().workbook.names.add('Bad',42);
// @ts-expect-error .name is read-only in the Office-style API
createExcelApi(b).createRequestContext().workbook.names.getItem('Rate').name='New';
void[item,all,names,range,limit,dialog];
