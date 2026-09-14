import {pivotMethods} from './pivots/host.js';
import {Workbook} from './model.js';
import {isError} from './errors.js';
const encode=value=>JSON.stringify(value,(_key,v)=>isError(v)?{error:v.code,detail:v.detail}:v);
/** JSON-only, allowlisted RPC. It never evaluates application-provided scripts or member paths. */
export function createHostBridge(control,{postMessage=()=>{}}={}){
 const eventQueue=[];const publish=postMessage;postMessage=message=>{if(eventQueue.length>=1000)eventQueue.shift();eventQueue.push(JSON.parse(message));publish(message);};
 let disposed=false;const sheet=request=>{const value=request.sheet==null?control.Workbook.ActiveWorksheet:control.Workbook.Worksheets.Get(request.sheet);if(!value)throw new RangeError('Unknown worksheet');return value;};
 const range=request=>sheet(request).GetRange(request.address??'A1');
 const methods={
  ...pivotMethods(control),
  'workbook.get':()=>control.Workbook.ToJSON(),
  'workbook.load':r=>{control.Workbook=Workbook.FromJSON(r.workbook);subscribe();return true;},
  'workbook.new':()=>{control.Workbook=new Workbook();subscribe();return true;},
  'workbook.calculate':r=>{control.Workbook.Calculate(!!r.full);return true;},
  'workbook.undo':()=>control.Workbook.Undo(),'workbook.redo':()=>control.Workbook.Redo(),
  'worksheets.list':()=>[...control.Workbook.Worksheets].map(s=>({id:s.Id,name:s.Name,cellCount:s.CellCount})),
  'worksheets.add':r=>{const s=control.Workbook.Worksheets.Add(r.name);return{name:s.Name,id:s.Id};},
  'worksheets.remove':r=>control.Workbook.Worksheets.Remove(sheet(r)),
  'worksheets.rename':r=>{sheet(r).Name=r.name;return true;},
  'range.values.get':r=>range(r).Values,'range.values.set':r=>{range(r).Values=r.values;return true;},
  'range.formulas.get':r=>range(r).Formulas,'range.formulas.set':r=>{range(r).Formulas=r.formulas;return true;},
  'range.text.get':r=>range(r).Text,'range.style.set':r=>{range(r).SetStyle(r.style);return true;},
  'range.clear':r=>{range(r).Clear(r.mode??'contents');return true;},
  'range.merge':r=>{range(r).Merge();return true;},'range.unmerge':r=>{range(r).Unmerge();return true;},
  'view.selection.get':()=>control.Selection,'view.selection.set':r=>{if(r.sheet!=null)control.Sheet=sheet(r);control.Select(r.address);return true;},
  'view.options':r=>{for(const key of ['Theme','Zoom','ReadOnly','ShowGridLines','ViewMode'])if(Object.hasOwn(r,key))control[key]=r[key];return true;},
  'sheet.freeze':r=>{sheet(r).FreezePanes(r.rows??0,r.columns??0);return true;},
  'sheet.insertRows':r=>{sheet(r).InsertRows(r.index,r.count??1);return true;},'sheet.deleteRows':r=>{sheet(r).DeleteRows(r.index,r.count??1);return true;},
  'sheet.insertColumns':r=>{sheet(r).InsertColumns(r.index,r.count??1);return true;},'sheet.deleteColumns':r=>{sheet(r).DeleteColumns(r.index,r.count??1);return true;},
  'names.define':r=>{control.Workbook.DefineName(r.name,r.value);return true;},
  'capabilities':()=>({version:1,methods:Object.keys(methods),formulaFunctions:control.Workbook.Calculation.FunctionNames,maxRows:1048576,maxColumns:16384,maxOperationCells:250000})
 };
 let subscription;function subscribe(){subscription?.Dispose();subscription=control.Workbook.Changed.Subscribe(e=>postMessage(encode({type:'workbook-changed',revision:e.Revision,label:e.Label})));}subscribe();
 const selection=e=>postMessage(encode({type:'selection-changed',address:e.detail.address,sheet:e.detail.worksheet.Name}));control.addEventListener?.('selection-change',selection);
 return{
  dispatch(json){let request;try{if(disposed)throw new Error('Bridge has been disposed');if(typeof json!=='string'||json.length>32*1024*1024)throw new Error('Request exceeds 32 MiB');request=JSON.parse(json);if(!request||typeof request!=='object'||typeof request.method!=='string'||!Object.hasOwn(methods,request.method))throw new Error('Unknown RPC method');return encode({id:request.id??null,result:methods[request.method](request)});}catch(e){return encode({id:request?.id??null,error:{message:e.message??String(e)}});}},
  drainEvents(){if(disposed)return[];return eventQueue.splice(0);},
  Dispose(){eventQueue.length=0;if(disposed)return;disposed=true;subscription?.Dispose();control.removeEventListener?.('selection-change',selection);}
 };
}
export function installHost(control,globalObject=globalThis){const postMessage=message=>{if(globalObject.chrome?.webview)globalObject.chrome.webview.postMessage(message);else if(globalObject.webkit?.messageHandlers?.invoke)globalObject.webkit.messageHandlers.invoke.postMessage(message);else if(typeof globalObject.external?.notify==='function')globalObject.external.notify(message);};const bridge=createHostBridge(control,{postMessage});globalObject.gridWebHost?.Dispose?.();globalObject.gridWebHost=bridge;postMessage(JSON.stringify({type:'ready',version:1}));return bridge;}
