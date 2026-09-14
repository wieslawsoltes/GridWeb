import {Worker} from 'node:worker_threads';
import {CollaborationError} from './document.js';
/** Bounded worker pool: expensive or adversarial calculations cannot block HTTP/authentication. */
export function createValidationPool({timeoutMs=5000,size=2,maxQueue=32}={}) {
  if(!Number.isInteger(timeoutMs)||timeoutMs<100||timeoutMs>60000||!Number.isInteger(size)||size<1||size>8)throw new RangeError('Invalid validation limits');
  let closed=false;const queue=[],slots=[];
  const start=slot=>{
    slot.worker=new Worker(new URL('./validation-worker.js',import.meta.url),{resourceLimits:{maxOldGenerationSizeMb:128,maxYoungGenerationSizeMb:32}});
    slot.worker.on('message',message=>finish(slot,message.error?new CollaborationError(message.error.code,message.error.message,message.error.details):null,message.result));
    slot.worker.on('error',error=>restart(slot,new CollaborationError('COMPUTE_LIMIT','Validation worker failed or exhausted its memory budget')));
    slot.worker.on('exit',code=>{if(code!==0&&slot.job)restart(slot,new CollaborationError('COMPUTE_LIMIT','Validation worker exited'));});
    slot.worker.unref();
  };
  const finish=(slot,error,value)=>{const job=slot.job;if(!job)return;clearTimeout(slot.timer);slot.job=null;slot.worker.unref();error?job.reject(error):job.resolve(value);pump();};
  const restart=(slot,error)=>{const old=slot.worker;old.removeAllListeners();old.terminate().catch(()=>{});if(!closed)start(slot);finish(slot,error);};
  const pump=()=>{if(closed)return;for(const slot of slots)if(!slot.job&&queue.length){slot.job=queue.shift();slot.worker.ref();slot.timer=setTimeout(()=>restart(slot,new CollaborationError('COMPUTE_LIMIT','Workbook calculation exceeded its server execution budget')),timeoutMs);slot.worker.postMessage({operation:slot.job.operation,args:slot.job.args});}};
  for(let i=0;i<size;i++){const slot={};slots.push(slot);start(slot);}
  const run=(operation,...args)=>new Promise((resolve,reject)=>{if(closed){reject(new Error('Validation pool is closed'));return;}if(queue.length>=maxQueue){reject(new CollaborationError('LIMIT','Validation queue is full'));return;}queue.push({operation,args,resolve,reject});pump();});
  return {validate:document=>run('validate',document),apply:(document,changes)=>run('apply',document,changes),async close(){if(closed)return;closed=true;for(const job of queue.splice(0))job.reject(new Error('Validation pool is closed'));for(const slot of slots)if(slot.job){clearTimeout(slot.timer);slot.job.reject(new Error('Validation pool is closed'));slot.job=null;}await Promise.all(slots.map(s=>s.worker.terminate()));}};
}
