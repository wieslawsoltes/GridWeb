import {readReference, referencePrefix} from './reference-syntax.js';
import {MAX_ROWS,MAX_COLUMNS,columnName,parseRange,parseCell,splitSheet,quoteSheet,rewriteReferences,rewriteAxisReferences} from './address.js';
function originPoint(origin) {
  const p=typeof origin==='string'?parseCell(origin):origin??{row:0,column:0};
  const row=p.row,column=p.column??p.col;
  if(!Number.isInteger(row)||!Number.isInteger(column)||row<0||row>=MAX_ROWS||column<0||column>=MAX_COLUMNS)throw new RangeError('Invalid reference origin');
  return {row,column};
}
const axisR=(letter,value,absolute,origin)=>letter+(absolute?value+1:value===origin?'':`[${value-origin}]`);
const sheetPrefix=name=>name==null?'':quoteSheet(name)+'!';
/** Translate references without evaluating formulas or changing quoted strings. */
export function a1ToR1C1(formula,origin='A1') {
  const p=originPoint(origin);
  let value=rewriteReferences(formula,(r,sheet)=>sheetPrefix(sheet)+axisR('R',r.row,r.absoluteRow,p.row)+axisR('C',r.column,r.absoluteColumn,p.column));
  return rewriteAxisReferences(value,(range,axis,parts)=>sheetPrefix(range.sheet)+parts.map(part=>{
    const isRow=axis==='row',absolute=part.startsWith('$'),raw=part.replace('$','');
    const coordinate=isRow?Number(raw)-1:parseCell(raw+'1').column;
    return axisR(isRow?'R':'C',coordinate,absolute,isRow?p.row:p.column);
  }).join(':'));
}
const relative='(?:\\[[+-]?\\d+\\]|\\d*)';
const body=`(?:R${relative}C${relative}|R${relative}:R${relative}|C${relative}:C${relative})`;
const rcToken=new RegExp(`^(?:(?:'(?:[^']|'')+'|[A-Za-z_][\\w.]*)(?::(?:'(?:[^']|'')+'|[A-Za-z_][\\w.]*))?!)?${body}(?![\\w.\\[(])`,'i');
function coordinate(part,origin,max){
  const absolute=part!==''&&!part.startsWith('['),value=absolute?Number(part)-1:origin+(part?Number(part.slice(1,-1)):0);
  if(!Number.isInteger(value)||value<0||value>=max)throw new RangeError('Reference outside worksheet');
  return {value,absolute};
}
/** R1C1 formulas may use absolute, bracket-relative, mixed and whole-axis references. */
export function r1c1ToA1(formula,origin='A1') {
  const p=originPoint(origin),source=String(formula);let out='',i=0;
  while(i<source.length){
    if(source[i]==='"'){const start=i++;while(i<source.length){if(source[i++]==='"'){if(source[i]==='"')i++;else break;}}out+=source.slice(start,i);continue;}
    const match=rcToken.exec(source.slice(i));
    if(match && !/[\w.]/.test(source[i-1]??'')){
      const token=match[0],s=splitSheet(token);let replacement;
      try{
        const cell=/^R(\[[+-]?\d+\]|\d*)C(\[[+-]?\d+\]|\d*)$/i.exec(s.address);
        if(cell){const r=coordinate(cell[1],p.row,MAX_ROWS),c=coordinate(cell[2],p.column,MAX_COLUMNS);replacement=(c.absolute?'$':'')+columnName(c.value)+(r.absolute?'$':'')+(r.value+1);}
        else {const parts=s.address.split(':'),row=parts[0][0].toUpperCase()==='R';replacement=parts.map(part=>{const c=coordinate(part.slice(1),row?p.row:p.column,row?MAX_ROWS:MAX_COLUMNS);return(c.absolute?'$':'')+(row?c.value+1:columnName(c.value));}).join(':');}
        const prefix=token.includes('!')?readReference(token.slice(0,token.lastIndexOf('!')+1)+'A1'):null;
        replacement=referencePrefix(prefix?.sheet??null,prefix?.sheetEnd??null)+replacement;
      }catch{replacement='#REF!';}
      out+=replacement;i+=token.length;continue;
    }
    // Structured table selectors are opaque to reference translation.
    if(source[i]==='['){const start=i++;let depth=1;while(i<source.length&&depth){if(source[i]==='[')depth++;if(source[i]===']')depth--;i++;}out+=source.slice(start,i);continue;}
    out+=source[i++];
  }
  return out;
}
export function getFormulasR1C1(range){return range.Formulas.map((row,r)=>row.map((formula,c)=>formula?a1ToR1C1(formula,{row:range.Bounds.r1+r,column:range.Bounds.c1+c}):formula));}
export function setFormulasR1C1(range,values){
  if(!Array.isArray(values)||!values.every(Array.isArray))throw new TypeError('A formula matrix is required');
  const bounds=parseRange(range.Address);
  range.Formulas=values.map((row,r)=>row.map((formula,c)=>formula?r1c1ToA1(formula,{row:bounds.r1+r,column:bounds.c1+c}):formula));
}
