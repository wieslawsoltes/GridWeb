import { createFunctionRegistry as createBaseRegistry, compare, criteriaPredicate } from './functions-base.js';
import { extendFunctionRegistry } from './functions-extended.js';
export { compare, criteriaPredicate };
/** A separate extension module keeps legacy and newer contracts individually testable. */
export function createFunctionRegistry() {
  return extendFunctionRegistry(createBaseRegistry(), compare, criteriaPredicate);
}
