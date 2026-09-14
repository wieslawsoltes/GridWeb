import { extendArrays } from './functions-arrays.js';
import { extendStatistics } from './functions-statistical.js';
import { extendEngineering } from './functions-engineering.js';
export function extendFunctionRegistry(f, compare, criteriaPredicate) {
  extendArrays(f, compare, criteriaPredicate);
  extendStatistics(f, compare, criteriaPredicate);
  extendEngineering(f, compare, criteriaPredicate);
  return f;
}
