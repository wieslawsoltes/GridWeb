import type {ClipboardSnapshot, PasteOptions, FillSeriesOptions, SpecialCellType, SpecialCellOptions} from './editing.js';
/** Public GridWeb API. PascalCase model methods and explicit zero-based numeric indexes. */
export type Primitive = number | string | boolean | null;
export type CellInput = Primitive | Date | { error: string };
export type CellValue = Primitive | FormulaError;
export interface Disposable { Dispose(): void; dispose?(): void; unsubscribe?(): void; }
export interface Bounds { r1: number; c1: number; r2: number; c2: number; sheet?: string | null; }
export interface CellPosition { row: number; column: number; absoluteRow: boolean; absoluteColumn: boolean; }
export const MAX_ROWS: 1048576;
export const MAX_COLUMNS: 16384;
export const MAX_OPERATION_CELLS: 250000;
export function columnName(index: number): string;
export function columnIndex(name: string): number;
export function cellAddress(row: number, column: number): string;
export function parseCell(address: string): CellPosition;
export function parseRange(address: string): Bounds;
export function splitSheet(address: string): { sheet: string | null; address: string };
export function quoteSheet(name: string): string;
export function contains(bounds: Bounds, row: number, column: number): boolean;
export function intersects(a: Bounds, b: Bounds): boolean;
export function rangeAddress(bounds: Bounds): string;
export function boundedCells(bounds: Bounds, limit?: number): number;
export function shiftFormula(formula: string, rows: number, columns: number): string;
export function rewriteReferences(formula: string, transform: (cell: CellPosition, sheet: string | null, token: string) => string): string;
export function rewriteAxisReferences(formula: string, transform: (range: Bounds, axis: 'row' | 'column', parts: string[], token: string) => string): string;
export function renameSheetReferences(formula: string, previous: string, next: string): string;
export class FormulaError extends Error { code: string; detail: string; constructor(code: string, detail?: string); toJSON(): { error: string; detail: string }; }
export function error(code: string, detail?: string): FormulaError;
export function isError(value: unknown): value is FormulaError;
export function scalar(value: unknown): CellValue;
export function number(value: unknown): number;
export function text(value: unknown): string;
export function truth(value: unknown): boolean;
export function flatten(value: unknown): CellValue[];
export function matrix(value: unknown): CellValue[][];
export function serialDate(date: Date): number;
export function fromSerial(serial: number): Date;
export class EventSource<T = unknown> {
  Subscribe(listener: (event: T) => void): Disposable;
  subscribe(listener: (event: T) => void): Disposable;
  Add(listener: (event: T) => void): Disposable;
  Emit(event: T): void;
  Clear(): void;
  readonly Count: number;
}
export interface PropertyChange { Sender: object; PropertyName: string; OldValue?: unknown; NewValue?: unknown; }
export class ObservableObject { PropertyChanged: EventSource<PropertyChange>; SetProperty(name: string, value: unknown): boolean; Dispose(): void; }
export class RelayCommand<T = unknown, TResult = unknown> {
  constructor(execute: (parameter: T) => TResult, canExecute?: (parameter: T) => boolean);
  CanExecute(parameter?: T): boolean; Execute(parameter?: T): TResult | undefined;
  CanExecuteChanged: EventSource<RelayCommand<T, TResult>>; NotifyCanExecuteChanged(): void; Dispose(): void;
}
export class ObservableCollection<T> implements Iterable<T> {
  constructor(items?: Iterable<T>); readonly Count: number; Get(index: number): T | undefined;
  Add(item: T): T; Remove(item: T): boolean;
  CollectionChanged: EventSource<{ Action: string; NewItems?: T[]; OldItems?: T[]; Index: number }>;
  [Symbol.iterator](): Iterator<T>;
}
export function bind(source: ObservableObject, path: string, target: object, property: string, options?: { mode?: 'OneWay' | 'TwoWay'; event?: string; convert?: (value: unknown) => unknown }): Disposable;
export interface FontStyle { name?: string; size?: number; bold?: boolean; italic?: boolean; underline?: boolean; color?: string; strikethrough?: boolean; }
export interface CellStyle { font?: FontStyle; fill?: string; numberFormat?: string; horizontalAlignment?: 'left' | 'center' | 'right'; verticalAlignment?: 'top' | 'center' | 'bottom'; wrapText?: boolean; border?: { color?: string }; locked?: boolean; rotation?: number; }
export interface CellRecord { input: CellInput; literal?: boolean; style?: CellStyle; comment?: string; }
export interface WorksheetDocument { id: string; name: string; cells: [number, CellRecord][]; meta: Record<string, unknown>; }
export interface WorkbookDocument { format: 'GridWeb'; version: 1; name: string; locale: string; activeSheet: string; names: [string, CellInput][]; sheets: WorksheetDocument[]; }
export interface WorkbookChange { Workbook: Workbook; Revision: number; Label: string; Changes: { type: string; sheet?: Worksheet; row?: number; column?: number; [key: string]: unknown }[]; }
export interface DefinedNames extends Iterable<[string, CellInput]> { Add(name: string, value: CellInput): void; Get(name: string): CellInput | undefined; Remove(name: string): boolean; }
export class Workbook extends ObservableObject {
  constructor(options?: { name?: string; locale?: string; createSheet?: boolean });
  Name: string; Locale: string; readonly Revision: number; ActiveWorksheet: Worksheet;
  readonly Worksheets: WorksheetCollection; readonly worksheets: WorksheetCollection;
  readonly Names: DefinedNames; readonly names: DefinedNames; readonly Calculation: CalculationEngine;
  CalculationMode: 'Automatic' | 'Manual'; HistoryLimit: number;
  readonly CanUndo: boolean; readonly CanRedo: boolean;
  Changed: EventSource<WorkbookChange>; Calculated: EventSource<{ Workbook: Workbook; Revision: number }>;
  Transaction<T>(action: () => T): T; Transaction<T>(label: string, action: () => T): T;
  Calculate(full?: boolean): void; Undo(): boolean; Redo(): boolean; ClearHistory(): void;
  DefineName(name: string, value: CellInput): void; RemoveName(name: string): boolean;
  GetRange(address: string): CellRange;
  Find(query: string, options?: FindOptions): FindResult[];
  Replace(query: string, replacement: string, options?: FindOptions): number;
  ToJSON(): WorkbookDocument; static FromJSON(document: string | WorkbookDocument): Workbook;
  Dispose(): void;
}
export interface FindOptions { matchCase?: boolean; wholeCell?: boolean; formulas?: boolean; sheet?: Worksheet | null; }
export interface FindResult { sheet: Worksheet; row: number; column: number; address: string; value: CellValue; }
export class WorksheetCollection implements Iterable<Worksheet> {
  constructor(workbook: Workbook); readonly Count: number;
  Get(indexOrName: number | string): Worksheet | undefined;
  Add(name?: string, index?: number): Worksheet; Move(sheetOrName: Worksheet | string, index: number): Worksheet; Remove(sheetOrName: Worksheet | string): void;
  getItem(indexOrName: string | number): Worksheet; getItemAt(index: number): Worksheet; readonly items: Worksheet[]; CollectionChanged: EventSource<WorkbookChange>; add(name?: string, index?: number): Worksheet;
  [Symbol.iterator](): Iterator<Worksheet>;
}
export type ChartType = 'column' | 'bar' | 'line' | 'area' | 'pie' | 'scatter';
export interface ChartOptions { title?: string; type?: ChartType; row?: number; column?: number; width?: number; height?: number; }
export interface ChartModel extends Required<ChartOptions> { id: string; range: Bounds; }
export interface TableModel { name: string; range: Bounds; style?: string; totals?: boolean; }
export interface ConditionalFormat { type: 'cellValue' | 'colorScale' | 'dataBar' | 'formula' | 'duplicate'; operator?: string; value?: Primitive; criteria?: Primitive; formula?: string; minColor?: string; maxColor?: string; color?: string; style?: CellStyle; }
export interface ValidationRule { type: 'list' | 'number' | 'whole' | 'date' | 'textLength' | 'custom'; values?: Primitive[]; min?: number; max?: number; formula?: string; allowBlank?: boolean; message?: string; }
export class Worksheet {
  constructor(workbook: Workbook, name: string); readonly Workbook: Workbook; readonly Id: string; Name: string; name: string;
  readonly CellCount: number; readonly UsedRange: CellRange; readonly IsProtected: boolean;
  readonly Charts: ChartModel[]; readonly Tables: TableModel[]; readonly MergedRanges: Bounds[];
  FrozenRows: number; FrozenColumns: number;
  GetCell(address: string): Cell; GetCell(row: number, column: number): Cell;
  GetRange(address: string | Bounds): CellRange;
  GetRangeByIndexes(row: number, column: number, rowCount: number, columnCount: number): CellRange;
  getRange(address: string | Bounds): CellRange; getCell(row: number, column: number): Cell;
  SetColumnWidth(index: number, pixels: number): void; SetRowHeight(index: number, pixels: number): void;
  HideRows(start: number, count?: number, hidden?: boolean): void; HideColumns(start: number, count?: number, hidden?: boolean): void;
  GroupRows(start: number, count?: number, level?: number): void;
  FreezePanes(rows?: number, columns?: number): void; Protect(): void; Unprotect(): void;
  AddTable(address: string, name?: string): TableModel;
  AddChart(address: string, options?: ChartOptions): ChartModel;
  UpdateChart(id: string, patch: ChartOptions): void; RemoveChart(id: string): void;
  AddConditionalFormat(address: string, rule: ConditionalFormat): ConditionalFormat & { range: Bounds };
  AddValidation(address: string, rule: ValidationRule): void; Validate(row: number, column: number, value: CellInput): void;
  SetFilter(address: string, criteria?: { column: number; value: Primitive }[]): void; ClearFilter(): void;
  InsertRows(index: number, count?: number): void; DeleteRows(index: number, count?: number): void;
  InsertColumns(index: number, count?: number): void; DeleteColumns(index: number, count?: number): void;
  ToJSON(): WorksheetDocument;
}
export class Cell {
  constructor(sheet: Worksheet, row: number, column: number);
  readonly Worksheet: Worksheet; readonly Row: number; readonly Column: number; readonly Address: string;
  Input: CellInput; get Value(): CellValue; set Value(value: CellInput); Formula: string | null;
  readonly Text: string; Style: CellStyle; Comment: string;
  get value(): CellValue; set value(value: CellInput); formula: string | null; readonly text: string;
}
export class CellRange {
  constructor(sheet: Worksheet, bounds: Bounds); readonly Worksheet: Worksheet; readonly Bounds: Bounds;
  readonly Address: string; readonly FullAddress: string; readonly RowCount: number; readonly ColumnCount: number; readonly Count: number;
  get Values(): CellValue[][]; set Values(values: CellInput[][]); Formulas: (string | null)[][]; readonly Text: string[][];
  get Value(): CellValue; set Value(value: CellInput); Formula: string | null; get values(): CellValue[][]; set values(values: CellInput[][]); formulas: (string | null)[][];
  readonly Format: RangeFormat; readonly format: RangeFormat;
  GetCell(row: number, column: number): Cell; Offset(rows: number, columns: number): CellRange; Resize(rows: number, columns: number): CellRange;
  Clear(mode?: 'all' | 'contents' | 'formats'): void; SetStyle(patch: CellStyle): this; Merge(): void; Unmerge(): void;
  CopyFrom(source: CellRange, mode?: 'all' | 'values' | 'formulas' | 'formats'): void;
  Capture(): ClipboardSnapshot; PasteSpecial(source: CellRange | ClipboardSnapshot, options?: PasteOptions): CellRange;
  FillSeries(options?: FillSeriesOptions): number; SpecialCells(type: SpecialCellType, options?: SpecialCellOptions): CellRange[];
  FillDown(): void; FillRight(): void; AutoFill(destination: CellRange, options?: { series?: boolean }): void;
  Sort(keys?: { column: number; ascending?: boolean }[], options?: { hasHeaders?: boolean }): void;
  RemoveDuplicates(columns?: number[], options?: { hasHeaders?: boolean }): number;
}
export interface RangeFont extends FontStyle { Bold?: boolean; Italic?: boolean; Underline?: boolean; Color?: string; Size?: number; Name?: string; Strikethrough?: boolean; }
export class RangeFormat {
  constructor(range: CellRange); readonly Font: RangeFont; readonly font: RangeFont;
  Fill: string | undefined; fill: string | undefined; NumberFormat: string; numberFormat: string;
  set WrapText(value: boolean); set HorizontalAlignment(value: 'left' | 'center' | 'right'); set VerticalAlignment(value: 'top' | 'center' | 'bottom');
  set Borders(value: { color?: string }); set ColumnWidth(pixels: number); set RowHeight(pixels: number);
  AutoFitColumns(): void;
}
export class GridViewModel extends ObservableObject { constructor(workbook?: Workbook); Workbook: Workbook; Selection: string; Undo: RelayCommand; Redo: RelayCommand; Dispose(): void; }
export interface FormulaContext { sheet?: Worksheet; row?: number; col?: number; vars?: Map<string, unknown>; }
export class CalculationEngine {
  constructor(workbook: Workbook); readonly Workbook: Workbook; readonly FunctionNames: string[]; readonly EvaluationCount: number;
  RegisterFunction(name: string, callback: (...args: unknown[]) => unknown): void;
  Evaluate(formula: string, context?: FormulaContext): CellValue | CellValue[][];
  GetValue(sheet: Worksheet, row: number, column: number): CellValue;
  Calculate(options?: { full?: boolean }): void; Reset(): void; Invalidate(sheet: Worksheet, row: number, column: number): void;
  GetDependencyEdges(): { source: string; target: string }[];
}
export function isFormula(record: unknown): boolean;
export function tokenize(source: string): { type: string; value?: unknown; [key: string]: unknown }[];
export function parseFormula(source: string): { type: string; [key: string]: unknown };
export function createFunctionRegistry(): Map<string, (...args: unknown[]) => unknown>;
export function compare(a: unknown, b: unknown): number;
export function criteriaPredicate(criteria: unknown): (value: unknown) => boolean;
export function formatValue(value: unknown, format?: string, locale?: string): string;
export interface PivotOptions { rows?: (string | number)[]; columns?: (string | number)[]; values?: { column: string | number; aggregate?: 'sum' | 'count' | 'average' | 'min' | 'max' | 'distinct'; name?: string }[]; filters?: { column: string | number; values: Primitive[] }[]; }
export class PivotTable { constructor(source: CellRange, options?: PivotOptions); Source: CellRange; Options: PivotOptions; Result: CellInput[][]; Refresh(): CellInput[][]; WriteTo(destination: CellRange): CellInput[][]; }
export function goalSeek(formulaCell: Cell, inputCell: Cell, target: number, options?: { min?: number; max?: number; guess?: number; tolerance?: number; iterations?: number }): { value: number; result: number; iterations: number };
export function linearRegression(xs: number[], ys: number[]): { slope: number; intercept: number; rSquared: number; predict(x: number): number };
export function analyzeRange(range: CellRange): { count: number; numericCount: number; sum: number; average: number; min: number; max: number; errors: number };
export function toDelimited(range: CellRange, options?: { delimiter?: string; formulas?: boolean; safe?: boolean }): string;
export function parseDelimited(text: string, options?: { delimiter?: string; maxCells?: number }): string[][];
export function importDelimited(sheet: Worksheet, text: string, options?: { address?: string; delimiter?: string; parseNumbers?: boolean; allowFormulas?: boolean }): CellRange | undefined;
export class AxisLayout { constructor(count: number, defaultSize: number); Count: number; DefaultSize: number; SetSize(index: number, size: number): void; Size(index: number): number; Offset(index: number): number; readonly TotalSize: number; IndexAt(offset: number): number; static ForSheet(sheet: Worksheet): { rows: AxisLayout; columns: AxisLayout }; }
export interface PrintOptions { paper?: 'A4' | 'Letter'; orientation?: 'portrait' | 'landscape'; scale?: number; repeatRows?: number; area?: Bounds | null; }
export interface PrintPage { rows: number[]; columns: number[]; width: number; height: number; scale: number; number: number; }
export function paginate(sheet: Worksheet, options?: PrintOptions): { pages: PrintPage[]; rows: AxisLayout; columns: AxisLayout; config: PrintOptions };
export function printPagesHTML(sheet: Worksheet, options?: PrintOptions): string;
export function createPrintDocument(sheet: Worksheet, options?: PrintOptions): string;
export const printCSS: string;
export function safeColor(value: unknown, fallback?: string): string;
export function escapeHTML(value: unknown): string;
export function chartData(sheet: Worksheet, chart: ChartModel): { labels: CellValue[]; series: { name: string; values: (number | null)[] }[]; truncated: boolean };
export function drawChart(ctx: CanvasRenderingContext2D, sheet: Worksheet, chart: ChartModel, rect: { x: number; y: number; width: number; height: number }, options?: { dark?: boolean; selected?: boolean }): void;
