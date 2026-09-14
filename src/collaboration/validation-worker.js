import {parentPort} from 'node:worker_threads';
import {validateDocument,applyChanges} from './document.js';
// Only registered document operations are executed; formula text never becomes JavaScript.
parentPort.on('message', ({operation,args}) => {
  try {
    if(!['validate','apply'].includes(operation))throw new Error('Unknown validation operation');
    const result=operation==='validate'?validateDocument(args[0]):applyChanges(args[0],args[1]);
    parentPort.postMessage({result});
  } catch(error) {parentPort.postMessage({error:{code:error.code??'INVALID',message:error.message,details:error.details}});}
});
