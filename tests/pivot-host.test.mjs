import test from 'node:test';import assert from 'node:assert/strict';
import {Workbook} from '../src/core.js';import {createHostBridge} from '../src/host.js';
test('native and worker RPC creates, filters, drills, updates and removes managed pivots',()=>{
 const book=new Workbook();book.ActiveWorksheet.GetRange('A1:B4').Values=[['Region','Value'],['North',10],['North',20],['South',30]];
 const control={Workbook:book},bridge=createHostBridge(control),call=(method,args={})=>{const r=JSON.parse(bridge.dispatch(JSON.stringify({id:1,method,...args})));if(r.error)throw new Error(r.error.message);return r.result;};
 assert.ok(call('capabilities').methods.includes('pivots.add'));
 call('pivots.add',{name:'Report',source:{address:'A1:B4'},destination:{address:'D1'},options:{rows:['Region'],values:[{column:'Value'}]}});
 assert.equal(call('pivots.list').length,1);assert.equal(call('pivots.refresh',{name:'Report'})[1][1],30);
 assert.equal(call('pivots.drillDown',{name:'Report',row:1,column:1}).length,3);
 call('pivots.filter',{name:'Report',field:'Region',values:['North']});assert.equal(call('pivots.refresh',{name:'Report'}).length,2);
 call('pivots.update',{name:'Report',newName:'Renamed',options:{filters:[]}});assert.equal(call('pivots.list')[0].name,'Renamed');call('pivots.refreshAll');
 assert.throws(()=>call('pivots.refresh',{name:'missing'}),/Unknown/);assert.equal(call('pivots.remove',{name:'Renamed',clear:true}),true);assert.equal(book.ActiveWorksheet.GetCell('D1').Value,null);bridge.Dispose();book.Dispose();
});
