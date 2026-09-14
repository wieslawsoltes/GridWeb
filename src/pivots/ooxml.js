import {pivotId} from './model.js';
import {xmlEscape as x,parseXML,child,children} from '../xml.js';
import {readZip,writeZip,crc32} from '../zip.js';
import {parseRange,rangeAddress,cellAddress} from '../address.js';
import {isError} from '../errors.js';
import {key} from './aggregate.js';
const NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main',REL='http://schemas.openxmlformats.org/officeDocument/2006/relationships',PKG='http://schemas.openxmlformats.org/package/2006/relationships',XML='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const decode=new TextDecoder('utf-8',{fatal:true});
const subtotals={sum:'sum',count:'count',countNumbers:'countNums',average:'average',min:'min',max:'max',product:'product',variance:'var',variancePopulation:'varp',standardDeviation:'stdDev',standardDeviationPopulation:'stdDevp'};
const rel=(id,type,target)=>`<Relationship Id="${x(id)}" Type="${REL}/${type}" Target="${x(target)}"/>`;
const rels=value=>XML+`<Relationships xmlns="${PKG}">${value}</Relationships>`;
const type=(path,name)=>`<Override PartName="/${path}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.${name}+xml"/>`;
function member(value){return value==null?'<m/>':isError(value)?`<e v="${x(value.code)}"/>`:typeof value==='number'?`<n v="${value}"/>`:typeof value==='boolean'?`<b v="${+value}"/>`:`<s v="${x(value)}"/>`;}
function checkNative(report){
 const o=report.Options;
 if(o.columns.length||!o.rows.length||o.filters.some(f=>!o.rows.includes(f.column))||o.values.some(v=>!subtotals[v.aggregate]||o.rows.includes(v.column)))throw new Error('Native pivot export currently requires row axes, standard measures and row-field filters; use pivots:"flatten" explicitly for other layouts');
}
/** Decode only our freshly generated stored ZIP, never arbitrary imported packages. */
function generatedParts(bytes){
 const result=Object.create(null),view=new DataView(bytes.buffer,bytes.byteOffset,bytes.byteLength);let p=0;
 while(p+4<=bytes.length&&view.getUint32(p,true)===0x04034b50){
  if(p+30>bytes.length||view.getUint16(p+8,true)!==0||view.getUint16(p+6,true)!==0x800)throw new Error('Expected a GridWeb stored package');
  const size=view.getUint32(p+18,true),nl=view.getUint16(p+26,true),extra=view.getUint16(p+28,true),at=p+30+nl+extra,name=decode.decode(bytes.subarray(p+30,p+30+nl)),data=bytes.subarray(at,at+size);
  if(at+size>bytes.length||Object.hasOwn(result,name)||crc32(data)!==view.getUint32(p+14,true))throw new Error('Invalid generated package');result[name]=data;p=at+size;
 }return result;
}
/** Add native definitions, shared item caches and cache records to a newly generated workbook. */
export function addPivotParts(book,generated){
 const files=generatedParts(generated),text=path=>typeof files[path]==='string'?files[path]:decode.decode(files[path]),caches=[];let contentTypes=text('[Content_Types].xml'),wbRels=text('xl/_rels/workbook.xml.rels');
 const addRel=(path,value)=>files[path]=files[path]?text(path).replace('</Relationships>',value+'</Relationships>'):rels(value);
 for(const [i,report] of [...book.PivotTables].entries()){
  checkNative(report);const id=i+1,o=report.Options,source=report.Source,matrix=source.Values,headers=matrix[0],computed=report._compute(),items=headers.map(()=>[]),indexes=headers.map(()=>new Map());
  for(const row of matrix.slice(1))row.forEach((v,j)=>{const k=key([v]);if(!indexes[j].has(k)){indexes[j].set(k,items[j].length);items[j].push(v);}});
  const pivotPath=`xl/pivotTables/pivotTable${id}.xml`,cachePath=`xl/pivotCache/pivotCacheDefinition${id}.xml`,recordsPath=`xl/pivotCache/pivotCacheRecords${id}.xml`;
  const fields=headers.map((name,j)=>{
   const values=items[j],number=values.some(v=>typeof v==='number'),string=values.some(v=>typeof v==='string'),other=values.some(v=>v!=null&&typeof v!=='number'&&typeof v!=='string');
   return `<cacheField name="${x(name)}"><sharedItems count="${values.length}" containsBlank="${+values.some(v=>v==null)}" containsNumber="${+number}" containsInteger="${+(number&&values.filter(v=>typeof v==='number').every(Number.isInteger))}" containsString="${+string}" containsSemiMixedTypes="${+(string||other)}" containsNonDate="1" containsMixedTypes="${+([number,string,other,values.some(v=>v==null)].filter(Boolean).length>1)}">${values.map(member).join('')}</sharedItems></cacheField>`;
  }).join('');
  files[cachePath]=XML+`<pivotCacheDefinition xmlns="${NS}" xmlns:r="${REL}" r:id="rIdRecords" refreshOnLoad="0" enableRefresh="1" recordCount="${matrix.length-1}" createdVersion="6" refreshedVersion="6" minRefreshableVersion="3"><cacheSource type="worksheet"><worksheetSource ref="${x(source.Address)}" sheet="${x(source.Worksheet.Name)}"/></cacheSource><cacheFields count="${headers.length}">${fields}</cacheFields></pivotCacheDefinition>`;
  files[recordsPath]=XML+`<pivotCacheRecords xmlns="${NS}" count="${matrix.length-1}">${matrix.slice(1).map(row=>'<r>'+row.map((v,j)=>`<x v="${indexes[j].get(key([v]))}"/>`).join('')+'</r>').join('')}</pivotCacheRecords>`;
  files[`xl/pivotCache/_rels/pivotCacheDefinition${id}.xml.rels`]=rels(rel('rIdRecords','pivotCacheRecords',`pivotCacheRecords${id}.xml`));
  const rowFields=o.rows.map(n=>headers.indexOf(n)),dataFields=o.values.map(v=>({...v,column:headers.indexOf(v.column)}));
  const pivotFields=headers.map((name,j)=>{
   if(rowFields.includes(j)){const filter=o.filters.find(f=>f.column===name),visible=filter?new Set(filter.values.map(v=>key([v]))):null;return `<pivotField axis="axisRow" showAll="0" compact="0" outline="0" defaultSubtotal="0"><items count="${items[j].length}">${items[j].map((v,k)=>`<item x="${k}"${visible&&!visible.has(key([v]))?' h="1"':''}/>`).join('')}</items></pivotField>`;}
   return `<pivotField${dataFields.some(v=>v.column===j)?' dataField="1"':''} showAll="0" defaultSubtotal="0"/>`;
  }).join('');
  const rowItems=computed.rows.map(([k,labels])=>k===null?'<i t="grand"><x/></i>':'<i>'+labels.map((v,j)=>`<x v="${indexes[rowFields[j]].get(key([v]))}"/>`).join('')+'</i>').join('');
  const multiple=dataFields.length>1;
  files[pivotPath]=XML+`<pivotTableDefinition xmlns="${NS}" name="${x(report.Name)}" cacheId="${id}" dataCaption="Values" compact="0" compactData="0" outline="0" outlineData="0" gridDropZones="1" rowGrandTotals="0" colGrandTotals="${+o.columnGrandTotals}" createdVersion="6" updatedVersion="6" minRefreshableVersion="3"><location ref="${x(report.OutputRange.Address)}" firstHeaderRow="0" firstDataRow="1" firstDataCol="${rowFields.length}"/><pivotFields count="${headers.length}">${pivotFields}</pivotFields><rowFields count="${rowFields.length}">${rowFields.map(j=>`<field x="${j}"/>`).join('')}</rowFields><rowItems count="${computed.rows.length}">${rowItems}</rowItems>${multiple?'<colFields count="1"><field x="-2"/></colFields>':''}<colItems count="${dataFields.length}">${dataFields.map((v,j)=>`<i i="${j}"><x/></i>`).join('')}</colItems><dataFields count="${dataFields.length}">${dataFields.map(v=>`<dataField name="${x(v.name)}" fld="${v.column}" subtotal="${subtotals[v.aggregate]}" baseField="0" baseItem="0"/>`).join('')}</dataFields><pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/></pivotTableDefinition>`;
  files[`xl/pivotTables/_rels/pivotTable${id}.xml.rels`]=rels(rel('rIdCache','pivotCacheDefinition',`../pivotCache/pivotCacheDefinition${id}.xml`));
  const si=book._sheets.indexOf(report.Destination.Worksheet)+1;addRel(`xl/worksheets/_rels/sheet${si}.xml.rels`,rel('rIdPivot'+id,'pivotTable',`../pivotTables/pivotTable${id}.xml`));
  caches.push(`<pivotCache cacheId="${id}" r:id="rIdPivotCache${id}"/>`);wbRels=wbRels.replace('</Relationships>',rel('rIdPivotCache'+id,'pivotCacheDefinition',`pivotCache/pivotCacheDefinition${id}.xml`)+'</Relationships>');
  contentTypes=contentTypes.replace('</Types>',type(pivotPath,'pivotTable')+type(cachePath,'pivotCacheDefinition')+type(recordsPath,'pivotCacheRecords')+'</Types>');
 }
 files['xl/workbook.xml']=text('xl/workbook.xml').replace('</workbook>',`<pivotCaches>${caches.join('')}</pivotCaches></workbook>`);files['xl/_rels/workbook.xml.rels']=wbRels;files['[Content_Types].xml']=contentTypes;
 return writeZip(files);
}
function resolve(base,target){if(typeof target!=='string'||/[\\\0?#]/.test(target)||/^[a-z]+:/i.test(target))throw new Error('Invalid pivot relationship');const parts=target.startsWith('/')?[]:base.split('/').slice(0,-1);for(const p of target.split('/'))if(p==='..'){if(!parts.length)throw new Error('Pivot relationship escapes package');parts.pop();}else if(p&&p!=='.')parts.push(p);return parts.join('/');}
const relPath=p=>p.replace(/([^/]+)$/,'_rels/$1.rels');
export async function importPivotParts(book,bytes,warnings){
 const files=await readZip(bytes),read=p=>{if(!files.has(p))throw new Error('Missing pivot part '+p);return parseXML(decode.decode(files.get(p)));},links=p=>files.has(relPath(p))?children(read(relPath(p)),'Relationship'):[];
 const office=children(read('_rels/.rels'),'Relationship').find(r=>r.attrs.Type?.endsWith('/officeDocument'));
 const wp=office?resolve('',office.attrs.Target):'xl/workbook.xml',wb=read(wp),wr=links(wp),sheetNodes=children(child(wb,'sheets'),'sheet');
 for(const s of sheetNodes){const dest=book.Worksheets.Get(s.attrs.name),r=wr.find(r=>r.attrs.Id===s.attrs['r:id']);if(!dest||!r||r.attrs.TargetMode==='External')continue;const sp=resolve(wp,r.attrs.Target);
  for(const link of links(sp).filter(r=>r.attrs.Type?.endsWith('/pivotTable'))){try{
   if(link.attrs.TargetMode==='External')throw new Error('External pivot relationship unsupported');const pp=resolve(sp,link.attrs.Target),pivot=read(pp),cacheLink=links(pp).find(r=>r.attrs.Type?.endsWith('/pivotCacheDefinition'));
   if(!cacheLink||cacheLink.attrs.TargetMode==='External')throw new Error('Missing local pivot cache');const cache=read(resolve(pp,cacheLink.attrs.Target)),cacheSource=child(cache,'cacheSource'),ws=child(cacheSource,'worksheetSource');
   if(cacheSource?.attrs.type!=='worksheet'||!ws?.attrs.ref||!ws.attrs.sheet)throw new Error('Only local range-backed caches are supported');const sourceSheet=book.Worksheets.Get(ws.attrs.sheet);if(!sourceSheet)throw new Error('Pivot source sheet is absent');
   if(child(pivot,'pageFields')||child(pivot,'filters')||child(pivot,'pivotHierarchies')||child(cache,'calculatedItems'))throw new Error('Advanced pivot filters/calculations are not supported');
   const columns=children(child(pivot,'colFields'),'field').map(f=>+f.attrs.x);if(columns.some(i=>i!==-2))throw new Error('Column-axis pivot import is not implemented');
   const cacheFields=children(child(cache,'cacheFields'),'cacheField'),fields=children(child(pivot,'pivotFields'),'pivotField'),rowFields=children(child(pivot,'rowFields'),'field').map(f=>+f.attrs.x),values=children(child(pivot,'dataFields'),'dataField').map(f=>{const aggregate=Object.entries(subtotals).find(([,n])=>n===(f.attrs.subtotal??'sum'))?.[0];if(!aggregate||f.attrs.showDataAs&&f.attrs.showDataAs!=='normal')throw new Error('Unsupported pivot aggregation');return {column:+f.attrs.fld,aggregate,name:f.attrs.name};});
   if(!rowFields.length||rowFields.some(i=>!cacheFields[i])||fields.length!==cacheFields.length)throw new Error('Invalid pivot field indexes');
   const filters=[];
   for(const j of rowFields){const entries=children(child(fields[j],'items'),'item');if(entries.some(e=>e.attrs.h==='1'||e.attrs.h==='true')){
    const shared=child(cacheFields[j],'sharedItems')?.children??[];const vals=[];for(const item of entries.filter(e=>e.attrs.h!=='1'&&e.attrs.h!=='true'&&(!e.attrs.t||e.attrs.t==='data'))){const v=shared[+item.attrs.x];if(!v)throw new Error('Invalid shared pivot item');if(!['s','n','b','m'].includes(v.name))throw new Error('Unsupported filtered pivot item');vals.push(v.name==='m'?null:v.name==='n'?+v.attrs.v:v.name==='b'?v.attrs.v==='1':v.attrs.v);}filters.push({column:j,values:vals});
   }}
   const location=child(pivot,'location')?.attrs.ref;if(!location)throw new Error('Missing pivot location');const b=parseRange(location),definition={id:pivotId(),name:pivot.attrs.name,source:{sheet:sourceSheet.Id,address:ws.attrs.ref},destination:{sheet:dest.Id,address:cellAddress(b.r1,b.c1)},options:{rows:rowFields,columns:[],values,filters,columnGrandTotals:pivot.attrs.colGrandTotals!=='0'&&pivot.attrs.colGrandTotals!=='false'},lastBounds:{r1:b.r1,c1:b.c1,r2:b.r2,c2:b.c2}};
   book.PivotTables._load([...book.PivotTables.ToJSON(),definition]);
   const old=warnings.indexOf('Not imported: '+pp);if(old>=0)warnings.splice(old,1);warnings.push('Pivot '+pivot.attrs.name+': imported range/field/aggregation definition; layout formatting is normalized on refresh');
  }catch(e){warnings.push('Pivot import: '+e.message);}}
 }
}
export {checkNative};
