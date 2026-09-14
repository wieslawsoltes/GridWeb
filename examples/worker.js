import {WorkbookWorkerClient} from '../src/worker.js';
const worker=new Worker(new URL('../src/worker.js',import.meta.url),{type:'module'});
const client=new WorkbookWorkerClient(worker);
await client.Call('range.values.set',{address:'A1:B1',values:[[2,3]]});
await client.Call('range.formulas.set',{address:'C1',formulas:[['=SUM(A1:B1)']]});
console.log(await client.Call('range.values.get',{address:'C1'})); // [[5]]
client.Dispose({terminate:true});
