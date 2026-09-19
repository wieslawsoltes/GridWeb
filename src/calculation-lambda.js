import {error, isError, number} from './errors.js';
import {MAX_OPERATION_CELLS} from './address.js';
import {checkedMatrix} from './functions-support.js';

export const OMITTED = Symbol('omitted lambda argument');
export const LAMBDA_HELPERS = new Set(['MAKEARRAY', 'MAP', 'BYROW', 'BYCOL', 'REDUCE', 'SCAN']);
const scalarResult = value => {
  if (Array.isArray(value) || value?.type === 'lambda') throw error('#CALC!', 'Nested array results are not supported');
  return value;
};
export function evaluateLambdaHelper(engine, name, args, ev) {
  const required = name === 'MAKEARRAY' || name === 'REDUCE' || name === 'SCAN' ? 3 : 2;
  if (args.length < required || (name !== 'MAP' && args.length !== required) || args.length > 254) throw error('#VALUE!', 'Incorrect Parameters');
  const lambda = ev(args.at(-1));
  if (isError(lambda)) throw lambda;
  const arity = name === 'MAKEARRAY' || name === 'REDUCE' || name === 'SCAN' ? 2 : name === 'MAP' ? args.length - 1 : 1;
  if (lambda?.type !== 'lambda' || lambda.parameters.length !== arity) throw error('#VALUE!', 'Incorrect Parameters');
  const invoke = values => engine._invoke(lambda, values);
  if (name === 'MAKEARRAY') {
    const rows = Math.trunc(number(ev(args[0]))), cols = Math.trunc(number(ev(args[1])));
    if (rows < 1 || cols < 1) throw error('#VALUE!', 'Array dimensions must be positive');
    if (rows * cols > MAX_OPERATION_CELLS) throw error('#NUM!', 'Array size limit');
    return Array.from({length: rows}, (_, r) => Array.from({length: cols}, (_, c) => scalarResult(invoke([r + 1, c + 1]))));
  }
  if (name === 'REDUCE' || name === 'SCAN') {
    let value = ev(args[0]);
    const array = checkedMatrix(ev(args[1]));
    if (name === 'REDUCE') {
      for (const row of array) for (const item of row) value = invoke([value, item]);
      return value;
    }
    return array.map(row => row.map(item => (value = scalarResult(invoke([value, item])))));
  }
  const arrays = args.slice(0, -1).map(arg => checkedMatrix(ev(arg))), array = arrays[0];
  if (name === 'BYROW') return array.map(row => [scalarResult(invoke([[row]]))]);
  if (name === 'BYCOL') return [array[0].map((_, c) => scalarResult(invoke([array.map(row => [row[c]])])))];
  if (arrays.some(a => a.length !== array.length || a[0].length !== array[0].length)) throw error('#VALUE!', 'MAP array dimensions differ');
  return array.map((row, r) => row.map((_, c) => scalarResult(invoke(arrays.map(a => a[r][c])))));
}
