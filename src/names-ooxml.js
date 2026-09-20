/** Defined-name OOXML profile: preserve localSheetId, hidden state and comments. */
import {nameDefinitions} from './defined-names.js';
import {mapFormulaReferences, referencePrefix} from './reference-syntax.js';
import {parseFormula} from './parser.js';

export function exportDefinedNames(book, escape) {
  return nameDefinitions(book).map(item=>{
    if(item.baseAddress!=null)throw new Error('XLSX export of relative-origin names is not supported; use absolute references or JSON');
    if(item.value==null)throw new Error('XLSX cannot preserve a null-valued defined name; use a blank-cell reference or JSON');
    let formula=typeof item.value==='string'?item.value.startsWith('=')?item.value.slice(1):'"'+item.value.replace(/"/g,'""')+'"':typeof item.value==='boolean'?item.value?'TRUE':'FALSE':item.value?.error??String(item.value);
    const context=book._sheets.find(s=>s.Id===item.contextSheetId);
    if(item.contextSheetId!=null&&!context)throw new Error('XLSX cannot preserve a deleted name context; redefine its context or use JSON');
    if(typeof item.value==='string'&&item.value.startsWith('=')){
      formula=mapFormulaReferences(formula,t=>t.sheet==null&&item.contextSheetId!=null?(context?referencePrefix(context.Name)+t.address:'#REF!'):t.raw);
      // A workbook name can refer to a name local to its definition worksheet.
      // Do not qualify LET/LAMBDA parameters: only free identifiers are names.
      const edits=[];let ast;try{ast=parseFormula(formula);}catch{}
      function visit(node,bound=new Set()){
        if(!node)return;
        if((node.type==='name'||node.type==='call')&&node.sheet==null&&!bound.has(node.name)&&context?._names.has(node.name)&&!(node.type==='call'&&book.Calculation.FunctionNames.includes(node.name)))edits.push({start:node.start,end:node.end,value:referencePrefix(context.Name)+node.name});
        if(node.type==='call'&&node.sheet==null&&node.name==='LET'){const nested=new Set(bound);for(let i=0;i<node.args.length-1;i+=2){visit(node.args[i+1],nested);if(node.args[i]?.type==='name')nested.add(node.args[i].name);}visit(node.args.at(-1),nested);return;}
        if(node.type==='call'&&node.sheet==null&&node.name==='LAMBDA'){visit(node.args.at(-1),new Set([...bound,...node.args.slice(0,-1).map(p=>p.name)]));return;}
        for(const key of ['left','right','callee'])visit(node[key],bound);if(node.type==='unary')visit(node.value,bound);for(const arg of node.args??[])visit(arg,bound);for(const row of node.rows??[])for(const arg of row)visit(arg,bound);
      }
      visit(ast);for(const e of edits.sort((a,b)=>b.start-a.start))formula=formula.slice(0,e.start)+e.value+formula.slice(e.end);
    }
    const local=item.sheetId==null?'':` localSheetId="${book._sheets.findIndex(s=>s.Id===item.sheetId)}"`;
    return `<definedName name="${escape(item.name)}"${local}${item.hidden?' hidden="1"':''}${item.comment?` comment="${escape(item.comment)}"`:''}>${escape(formula)}</definedName>`;
  });
}
export function importDefinedName(snapshot,node,warnings) {
  const scoped=Object.hasOwn(node.attrs,'localSheetId'),raw=node.attrs.localSheetId;
  if(scoped&&(!/^\d+$/.test(raw)||!Number.isSafeInteger(Number(raw))||Number(raw)>4294967295))throw new Error('Invalid defined-name localSheetId');
  const owner=scoped?snapshot.sheets.find(s=>s.id==='sheet-'+Number(raw)):snapshot;
  if(!owner){warnings.push('Defined name on an unsupported or missing sheet was not imported: '+node.attrs.name);return;}
  (owner.names??=[]).push([node.attrs.name,'='+node.text]);
  (owner.nameMetadata??=[]).push({name:node.attrs.name,comment:node.attrs.comment??'',hidden:['1','true'].includes(node.attrs.hidden),contextSheetId:scoped?owner.id:null,baseAddress:null});
}
