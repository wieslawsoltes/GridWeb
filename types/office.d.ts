import type { Workbook, Primitive } from './index.js';
export class OfficeApiError extends Error { readonly code:string; }
export type LoadSelection=string|string[]|{select:string};
interface Loadable { load(selection?:LoadSelection):this; toJSON():Record<string,unknown>; }
export interface Font extends Loadable { bold:boolean|null; italic:boolean|null; strikethrough:boolean|null; underline:boolean|'Single'|'None'|null; color:string|null; name:string|null; size:number|null; set(values:Partial<Pick<Font,'bold'|'italic'|'strikethrough'|'underline'|'color'|'name'|'size'>>):this; }
export interface Fill extends Loadable { color:string|null; clear():void; }
export interface RangeFormat extends Loadable { readonly font:Font; readonly fill:Fill; wrapText:boolean|null; horizontalAlignment:string|null; verticalAlignment:string|null; autofitColumns():void; autofitRows():never; }
export class Range implements Loadable {
 readonly context:RequestContext; readonly worksheet:Worksheet;
 readonly address:string; readonly rowIndex:number; readonly columnIndex:number; readonly rowCount:number; readonly columnCount:number; readonly cellCount:number;
 values:Primitive[][]; formulas:Primitive[][]; readonly text:string[][]; numberFormat:(string|null)[][]; readonly format:RangeFormat;
 load(selection?:LoadSelection):this; set(values:{values?:Primitive[][];formulas?:Primitive[][];numberFormat?:(string|null)[][];format?:{font?:Partial<Font>;fill?:{color:string};wrapText?:boolean;horizontalAlignment?:string;verticalAlignment?:string}}):this; toJSON():Record<string,unknown>;
 getCell(row:number,column:number):Range; getRow(index:number):Range; getColumn(index:number):Range; getOffsetRange(rows:number,columns:number):Range; getResizedRange(deltaRows:number,deltaColumns:number):Range;
 clear(applyTo?:'All'|'Contents'|'Formats'):void; merge(across?:boolean):void; unmerge():void; copyFrom(source:Range,copyType?:'All'|'Values'|'Formulas'|'Formats',skipBlanks?:boolean,transpose?:boolean):void;
}
export class Worksheet implements Loadable { readonly context:RequestContext; name:string; readonly id:string; position:number; load(selection?:LoadSelection):this; toJSON():Record<string,unknown>; getRange(address:string):Range; getRangeByIndexes(row:number,column:number,rowCount:number,columnCount:number):Range; getUsedRange():Range; activate():void; delete():void; }
export interface WorksheetCollection { readonly context:RequestContext; readonly items:Worksheet[]; getItem(name:string):Worksheet; getItemAt(index:number):Worksheet; getActiveWorksheet():Worksheet; add(name?:string):Worksheet; load(selection?:string|string[]):this; }
export class RequestContext { constructor(workbook:Workbook); readonly workbook:{worksheets:WorksheetCollection}; load<T extends Loadable & {context:RequestContext}>(object:T,properties:LoadSelection):T; sync<T=void>(passThroughValue?:T):Promise<T>; dispose():void; }
export interface ExcelApi { run<T>(callback:(context:RequestContext)=>T|Promise<T>):Promise<T>; createRequestContext():RequestContext; readonly ClearApplyTo:{readonly all:'All';readonly contents:'Contents';readonly formats:'Formats'}; readonly RangeCopyType:{readonly all:'All';readonly values:'Values';readonly formulas:'Formulas'}; readonly HorizontalAlignment:{readonly left:'Left';readonly center:'Center';readonly right:'Right'}; readonly VerticalAlignment:{readonly top:'Top';readonly center:'Center';readonly bottom:'Bottom'}; }
/** This factory does not install or claim the complete Office.js host API. */
export function createExcelApi(workbook:Workbook):ExcelApi;
