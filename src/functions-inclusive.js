/** A-suffixed statistics include referenced text (as zero) and logical values. */
import {error, isError, number, flatten} from './errors.js';
import {MAX_OPERATION_CELLS} from './address.js';
export function extendInclusiveStatistics(registry) {
  const values = args => {
    if (!args.length || args.length > 255) throw error('#VALUE!', 'Statistical argument count');
    const out = [];
    for (const arg of args) {
      for (const value of Array.isArray(arg) ? flatten(arg) : [arg]) {
        if (value == null) continue;
        if (isError(value)) throw value;
        if (out.length >= MAX_OPERATION_CELLS) throw error('#NUM!', 'Statistical value limit');
        out.push(Array.isArray(arg) && typeof value === 'string' ? 0 : number(value));
      }
    }
    return out;
  };
  const variance = (a,sample) => {
    if (a.length <= (sample ? 1 : 0)) throw error('#DIV/0!');
    const mean = a.reduce((s,v) => s+v,0) / a.length;
    const value = a.reduce((s,v) => s+(v-mean)**2,0) / (a.length-(sample?1:0));
    if (!Number.isFinite(value)) throw error('#NUM!');
    return value;
  };
  registry.set('AVERAGEA', (...args) => {const a=values(args);if(!a.length)throw error('#DIV/0!');const value=a.reduce((s,v)=>s+v,0)/a.length;if(!Number.isFinite(value))throw error('#NUM!');return value;});
  registry.set('MAXA', (...args) => {const a=values(args);return a.length?a.reduce((m,v)=>Math.max(m,v),-Infinity):0;});
  registry.set('MINA', (...args) => {const a=values(args);return a.length?a.reduce((m,v)=>Math.min(m,v),Infinity):0;});
  for(const [name,sample,root] of [['VARA',true,false],['VARPA',false,false],['STDEVA',true,true],['STDEVPA',false,true]]) registry.set(name,(...args)=>{const v=variance(values(args),sample);return root?Math.sqrt(v):v;});
}
