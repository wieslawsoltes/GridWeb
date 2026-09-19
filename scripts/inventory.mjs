/** Regenerate availability inventories from runtime exports, not a hand-maintained function count. */
import fs from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import * as core from '../src/core.js';
import {GridWebElement} from '../src/controls.js';
const root=path.resolve(import.meta.dirname,'..');
function members(ctor) {
  const result=new Map();
  for(let proto=ctor.prototype;proto&&proto!==Object.prototype;proto=Object.getPrototypeOf(proto))for(const [name,descriptor] of Object.entries(Object.getOwnPropertyDescriptors(proto))) {
    if(name==='constructor'||name.startsWith('_')||result.has(name))continue;
    result.set(name,{name,kind:typeof descriptor.value==='function'?'method':'property',...(descriptor.get?{get:true}:{}),...(descriptor.set?{set:true}:{})});
  }
  return [...result.values()].sort((a,b)=>a.name<b.name?-1:a.name>b.name?1:0);
}
async function modules(directory='src') {
  const result=[];
  for(const item of await fs.readdir(path.join(root,directory),{withFileTypes:true})) {
    const name=directory+'/'+item.name;
    if(item.isDirectory())result.push(...await modules(name));else if(item.name.endsWith('.js'))result.push(name);
  }
  return result.sort();
}
export async function inventories() {
  const {version}=JSON.parse(await fs.readFile(path.join(root,'package.json'),'utf8'));
  const names=new core.Workbook().Calculation.FunctionNames;
  return {
    'functions.json':{version,count:names.length,status:'Available names, not complete Excel behavioral/signature conformance',names},
    'api-inventory.json':{version,scope:'Core exports and public prototype members; instance fields, method signatures and semantic equivalence require the declarations and feature audit',coreExports:Object.keys(core).sort(),publicPrototypeMembers:Object.fromEntries(['Workbook','WorksheetCollection','Worksheet','Cell','CellRange','RangeFormat','CalculationEngine','GridViewModel','PivotTable','PivotReport','PivotTableCollection'].filter(name=>typeof core[name]==='function').map(name=>[name,members(core[name])])),controlPrototypeMembers:members(GridWebElement),sourceModules:await modules()}
  };
}
if(process.argv[1]&&path.resolve(process.argv[1])===fileURLToPath(import.meta.url)) {
  for(const [name,content] of Object.entries(await inventories()))await fs.writeFile(path.join(root,'docs',name),JSON.stringify(content,null,2)+'\n');
  console.log('Regenerated runtime function and public API inventories.');
}
