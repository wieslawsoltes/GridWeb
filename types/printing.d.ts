import type {Worksheet,ChartModel,Bounds,AxisLayout,PrintPage} from './index-base.js';
export interface ExtendedPrintOptions {
 paper?:'A4'|'Letter'|'A3'|'Legal';orientation?:'portrait'|'landscape';scale?:number;area?:Bounds|string|null;
 repeatRows?:number;repeatColumns?:number;rowBreaks?:number[];columnBreaks?:number[];fitToWidthPages?:number;fitToHeightPages?:number;
 margins?:Partial<{top:number;right:number;bottom:number;left:number}>;pageOrder?:'overThenDown'|'downThenOver';header?:string;footer?:string;includeCharts?:boolean;
}
export function paginate(sheet:Worksheet,options?:ExtendedPrintOptions):{pages:(PrintPage&{contentRows:number[];contentColumns:number[];margins:{top:number;right:number;bottom:number;left:number}})[];rows:AxisLayout;columns:AxisLayout;config:ExtendedPrintOptions};
export function printPagesHTML(sheet:Worksheet,options?:ExtendedPrintOptions):string;
export function createPrintDocument(sheet:Worksheet,options?:ExtendedPrintOptions):string;
export function chartToSVG(sheet:Worksheet,chart:ChartModel,options?:{dark?:boolean;selected?:boolean}):string;
