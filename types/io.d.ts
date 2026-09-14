import { Workbook } from './index.js';
export { toDelimited as exportCSV, importDelimited as importCSV, parseDelimited } from './index.js';
export function exportXlsx(workbook: Workbook, options?: { onWarning?: (warning: string) => void; pivots?: 'native'|'flatten' }): Uint8Array;
export function importXlsx(bytes: Uint8Array | ArrayBuffer): Promise<{ workbook: Workbook; warnings: string[] }>;
