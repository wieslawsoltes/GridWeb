/** Excel-style database functions. Criteria rows are OR; columns within a row are AND. */
import {error, isError, scalar} from './errors.js';
import {checkedMatrix} from './functions-support.js';
import {MAX_OPERATION_CELLS} from './address.js';

export const DATABASE_FUNCTIONS = Object.freeze({
  DAVERAGE: 'AVERAGE', DCOUNT: 'COUNT', DCOUNTA: 'COUNTA', DGET: null,
  DMAX: 'MAX', DMIN: 'MIN', DPRODUCT: 'PRODUCT', DSTDEV: 'STDEV.S',
  DSTDEVP: 'STDEV.P', DSUM: 'SUM', DVAR: 'VAR.S', DVARP: 'VAR.P'
});
const blank = value => value == null || value === '';
const headerKey = value => String(value ?? '').toLocaleUpperCase('en-US');

/** computedCriterion is supplied only by the evaluator, never by workbook data. */
export function databaseAggregate(registry, criteriaPredicate, name, database, field, criteria, computedCriterion = null) {
  if (isError(database)) throw database;
  if (isError(criteria)) throw criteria;
  const data = checkedMatrix(database), tests = checkedMatrix(criteria);
  if (data.length < 2 || tests.length < 2) throw error('#VALUE!', 'Database and criteria require headers and data');
  const headers = data[0].map(headerKey), rawField = scalar(field);
  if (isError(rawField)) throw rawField;
  const countRecords = blank(rawField) && (name === 'DCOUNT' || name === 'DCOUNTA');
  const column = typeof rawField === 'number' ? Math.trunc(rawField) - 1 : headers.indexOf(headerKey(rawField));
  if (!countRecords && (blank(rawField) || column < 0 || column >= headers.length)) throw error('#VALUE!', 'Unknown database field');
  const columns = tests[0].map(value => { if (isError(value)) throw value; return headers.indexOf(headerKey(value)); });
  const conditions = tests.slice(1).map((row, ri) => row.map((criterion, ci) => {
    if (blank(criterion)) return null;
    if (columns[ci] < 0) {
      if (!computedCriterion) throw error('#VALUE!', 'Calculated criteria require worksheet formula references');
      return {computed: true, row: ri + 1, column: ci};
    }
    if (isError(criterion)) throw criterion;
    // Unadorned text is a prefix in database criteria, not an exact string match.
    if (typeof criterion === 'string' && !/^[<>=]/.test(criterion) && !/[?*~]/.test(criterion) && !Number.isFinite(Number(criterion)) && !/^(TRUE|FALSE)$/i.test(criterion)) criterion += '*';
    return {column: columns[ci], predicate: criteriaPredicate(criterion)};
  }));
  const selected = []; let comparisons = 0;
  const matches = (condition, row) => {
    if (++comparisons > MAX_OPERATION_CELLS * 16) throw error('#NUM!', 'Database criteria evaluation limit');
    return !condition || (condition.computed ? computedCriterion(row, condition.row, condition.column) : condition.predicate(data[row][condition.column]));
  };
  for (let i = 1; i < data.length; i++) {
    if (conditions.some(row => row.every(condition => matches(condition, i)))) selected.push(countRecords ? 1 : data[i][column]);
  }
  if (name === 'DGET') {
    if (!selected.length) throw error('#VALUE!', 'No matching record');
    if (selected.length !== 1) throw error('#NUM!', 'More than one matching record');
    return selected[0] ?? 0;
  }
  return registry.get(DATABASE_FUNCTIONS[name])(selected);
}

export function extendDatabase(registry, criteriaPredicate) {
  for (const name of Object.keys(DATABASE_FUNCTIONS)) registry.set(name, (...args) => {
    if (args.length !== 3) throw error('#VALUE!', name + ' requires database, field and criteria');
    return databaseAggregate(registry, criteriaPredicate, name, ...args);
  });
  return registry;
}
