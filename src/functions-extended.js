import {extendInclusiveStatistics} from './functions-inclusive.js';
import {extendDatabase} from './functions-database.js';
import { extendArrays } from './functions-arrays.js';
import { extendStatistics } from './functions-statistical.js';
import { extendEngineering } from './functions-engineering.js';
export function extendFunctionRegistry(f, compare, criteriaPredicate) {
  extendInclusiveStatistics(f);
  extendArrays(f, compare, criteriaPredicate);
  extendStatistics(f, compare, criteriaPredicate);
  extendEngineering(f, compare, criteriaPredicate);
  extendDatabase(f,criteriaPredicate);return f;
}
