/** Reference arguments are internal values, never stored cell inputs or host objects. */
import {error} from './errors.js';
const marker = Symbol('GridWeb reference binding');
export const bindingReference = value => value?.[marker] ?? null;
export function referenceBinding(engine, reference, ctx) {
  const pin = ref => { if(ref.sheet==null&&!ctx.sheet)throw error('#REF!','Name context worksheet was deleted');return {...ref,sheet:ref.sheet ?? ctx.sheet.Name}; };
  return Object.freeze({[marker]:reference.type==='multiRef'?{...reference,areas:reference.areas.map(pin)}:pin(reference)});
}
