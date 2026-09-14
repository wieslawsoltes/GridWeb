import type { Workbook, Worksheet } from './index.js';
export interface HostControl { Workbook: Workbook; Sheet: Worksheet | string; Selection: string; Select(address: string): void; addEventListener?: HTMLElement['addEventListener']; removeEventListener?: HTMLElement['removeEventListener']; }
export interface HostBridge { dispatch(requestJSON: string): string; drainEvents(): Record<string, unknown>[]; Dispose(): void; }
export function createHostBridge(control: HostControl, options?: { postMessage?: (message: string) => void }): HostBridge;
export function installHost(control: HostControl, globalObject?: typeof globalThis): HostBridge;
