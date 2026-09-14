import * as Core from '../../src/core.js';
import * as IO from '../../src/io.js';
import * as Office from '../../src/office.js';
import * as Collaboration from '../../src/collaboration.js';
import * as Host from '../../src/host.js';
import * as Workers from '../../src/worker.js';
import * as Controls from '../../src/controls.js';
import { SpreadsheetSession } from './session.js';
import { SpreadsheetView } from './view.js';
Controls.defineGridWeb();
export const api = { ...Core, ...IO, ...Office, ...Collaboration, ...Host, ...Workers, ...Controls, SpreadsheetSession };
export async function mount(host, options) {
  const view = new SpreadsheetView(host);
  try { await view.Configure(options); return view; }
  catch (error) { await view.DisposeAsync(); throw error; }
}
export function update(view, options) { return view.Configure(options); }
