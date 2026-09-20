import type {PasteOptions, FillSeriesOptions, SpecialCellType, SpecialCellOptions} from './editing.js';
import { Workbook, Worksheet, Cell, CellRange, Bounds, GridViewModel, EventSource, ChartModel, WorkbookChange, DefinedNameOptions, CellInput } from './index.js';
export interface SelectionChange { address: string; worksheet: Worksheet; bounds: Bounds; row: number; column: number; cell: Cell; }
export interface CellEdit { worksheet: Worksheet; row: number; column: number; value: unknown; input: unknown; }
export class GridWebElement extends HTMLElement {
  ShowCreateNamesFromSelection(): HTMLDialogElement; ShowNameManager(): HTMLDialogElement; ShowDefineName(): HTMLDialogElement; DefineName(name:string,value:CellInput,options?:DefinedNameOptions,sheet?:Worksheet|string|null):void;
  Workbook: Workbook; workbook: Workbook; Model: Workbook; Sheet: Worksheet | string;
  DataContext: GridViewModel | null; Theme: 'light' | 'dark'; Zoom: number; ReadOnly: boolean; ShowGridLines: boolean;
  ViewMode: 'normal' | 'pageLayout' | 'pageBreak'; Selection: string;
  readonly SelectionRange: CellRange; readonly ActiveCell: Cell;
  readonly Metrics: { frames: number; visibleCells: number; renderMs: number };
  SelectionChanged: EventSource<SelectionChange>; CellEdited: EventSource<CellEdit>;
  ChartHitTest?: ((x: number, y: number) => { chart: ChartModel; x: number; y: number; width: number; height: number }[]) | null;
  Select(address: string | Bounds, options?: { scroll?: boolean; extend?: boolean }): void;
  ScrollIntoView(row: number, column: number): void; Refresh(): void; Focus(): void;
  BeginEdit(initial?: string): boolean; CommitEdit(focus?: boolean): boolean; CancelEdit(): void;
  PasteText(text: string): void; CopySelection(): Promise<string>; Paste(): Promise<void>;
  PasteSpecial(options?: PasteOptions): CellRange; FillSeries(options?: FillSeriesOptions): number;
  FindSpecialCells(type: SpecialCellType, options?: SpecialCellOptions): CellRange[];
  MoveWorksheet(index: number, sheet?: Worksheet | string): Worksheet; ShowMoveWorksheet(): HTMLDialogElement;
  ShowPasteSpecial(): HTMLDialogElement; ShowFillSeries(): HTMLDialogElement; ShowGoToSpecial(): HTMLDialogElement;
  Print(): void; Notify(message: string): void; Dispose(): void;
}
export function defineGridWeb(name?: string): typeof GridWebElement;
declare global { interface HTMLElementTagNameMap { 'grid-web': GridWebElement; } }
