import type {CellRange} from './index-base.js';
export type ReferenceOrigin = string | {row:number;column:number};
/** Translate formula references; strings and structured selectors are not modified. */
export function a1ToR1C1(formula:string,origin?:ReferenceOrigin):string;
export function r1c1ToA1(formula:string,origin?:ReferenceOrigin):string;
export function getFormulasR1C1(range:CellRange):(string|null)[][];
export function setFormulasR1C1(range:CellRange,values:(string|null)[][]):void;
