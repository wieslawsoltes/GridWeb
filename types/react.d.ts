import type { ForwardRefExoticComponent, RefAttributes, HTMLAttributes, CSSProperties } from 'react';
import type { Workbook, Worksheet, WorkbookChange } from './index.js';
import type { GridWebElement, SelectionChange, CellEdit } from './controls.js';
export interface GridWebProps extends Omit<HTMLAttributes<HTMLElement>, 'onChange'> {
 workbook?: Workbook; sheet?: Worksheet | string; selection?: string; theme?: 'light' | 'dark'; readOnly?: boolean; zoom?: number;
 viewMode?: 'normal' | 'pageLayout' | 'pageBreak'; showGridLines?: boolean; style?: CSSProperties;
 onSelectionChange?: (event: SelectionChange) => void; onCellEdit?: (event: CellEdit) => void; onWorkbookChange?: (event: WorkbookChange) => void;
}
export const GridWeb: ForwardRefExoticComponent<GridWebProps & RefAttributes<GridWebElement>>;
export function useWorkbook(create?: () => Workbook): Workbook;
export function useWorkbookRevision(workbook: Workbook): number;
