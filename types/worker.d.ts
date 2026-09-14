export class WorkbookWorkerClient {
 constructor(worker: Worker, options?: { timeout?: number }); readonly Worker: Worker; Timeout: number;
 Call<T = unknown>(method: string, args?: Record<string, unknown>, signal?: AbortSignal): Promise<T>;
 Dispose(options?: { terminate?: boolean }): void;
}
export function createWorkerHandler(send: (message: unknown) => void): { handle(data: unknown): void; Dispose(): void };
