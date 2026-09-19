import type {CellRange, Bounds} from './index-base.js';
/** Opaque, immutable in-process clipboard snapshot. Obtain with Capture/captureRange, not JSON. */
declare const clipboardBrand: unique symbol;
export interface ClipboardSnapshot {
  readonly [clipboardBrand]: never;
  readonly kind: 'GridWebClipboard'; readonly bounds: Readonly<Bounds>;
  readonly rows: number; readonly columns: number; readonly hasMerges: boolean;
}
export type PasteMode = 'all' | 'values' | 'formulas' | 'formats' | 'comments' | 'validation' | 'columnWidths' | 'valuesAndNumberFormats' | 'formulasAndNumberFormats';
export interface PasteOptions {
  mode?: PasteMode; operation?: 'none' | 'add' | 'subtract' | 'multiply' | 'divide';
  transpose?: boolean; skipBlanks?: boolean;
}
export interface FillSeriesOptions {
  type?: 'linear' | 'growth' | 'date'; direction?: 'down' | 'right' | 'up' | 'left';
  start?: number | Date; step?: number; stop?: number | Date; dateUnit?: 'day' | 'weekday' | 'month' | 'year';
}
export type SpecialCellType = 'formulas' | 'constants' | 'blanks' | 'errors' | 'comments' | 'validation' | 'visible';
export interface SpecialCellOptions { valueTypes?: ('numbers' | 'text' | 'logical' | 'errors')[]; }
export function captureRange(range: CellRange): ClipboardSnapshot;
export function pasteSpecial(destination: CellRange, source: CellRange | ClipboardSnapshot, options?: PasteOptions): CellRange;
export function fillSeries(range: CellRange, options?: FillSeriesOptions): number;
/** Individual ranges, in row-major order. Does not create a multi-area selection. */
export function specialCells(range: CellRange, type: SpecialCellType, options?: SpecialCellOptions): CellRange[];
