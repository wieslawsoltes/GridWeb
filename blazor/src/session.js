import * as Core from '../../src/core.js';
import * as IO from '../../src/io.js';
import { createHostBridge } from '../../src/host.js';
import { createExcelApi } from '../../src/office.js';
import { CollaborationSession } from '../../src/collaboration.js';

const dto = value => JSON.parse(JSON.stringify(value, (_key, item) => Core.isError(item) ? { error: item.code, detail: item.detail } : item));
const object = value => { if (!value || typeof value !== 'object' || Array.isArray(value)) throw new TypeError('Expected an options object.'); return value; };
const blocked = new Set(['__proto__', 'prototype', 'constructor']);
function safeMember(target, name) {
  const parts = String(name).split('.');
  if (parts.some(p => !p || blocked.has(p) || p.startsWith('_'))) throw new TypeError('Invalid public member path.');
  const key = parts.pop(); for (const p of parts) target = target?.[p];
  if (target == null) throw new TypeError('Missing public member.');
  return [target, key];
}
/** Per-owner engine session. No global bridge, DOM, eval or server-only dependency. */
export class SpreadsheetSession {
  constructor(document = null) {
    this.disposed = false; this.sequence = 0; this.views = new Set(); this._disposeTask = null;
    for (const name of ['Changed','Calculated','Replaced','Error','CollaborationChanged','CollaborationConflict','PresenceChanged']) this[name] = new Core.EventSource();
    this._book = document instanceof Core.Workbook ? document : document == null ? new Core.Workbook() : Core.Workbook.FromJSON(document);
    this._owned = !(document instanceof Core.Workbook); this._selection = 'A1'; this._collaboration = null;
    const self = this;
    this._headless = { get Workbook() { return self.Workbook; }, get Sheet() { return self.Workbook.ActiveWorksheet; },
      set Sheet(value) { self.ActivateWorksheet(value.Name); }, get Selection() { return self._selection; },
      Select(address) { self.Select(address); } };
    this._subscribe(); this._bridge = createHostBridge(this._headless);
  }
  check() { if (this.disposed) throw new Error('The workbook session has been disposed.'); }
  get Workbook() { this.check(); return this._book; }
  get Revision() { return this.sequence; }
  _subscribe() {
    this._changed?.Dispose(); this._calculated?.Dispose();
    this._changed = this._book.Changed.Subscribe(e => { this.sequence++; this.Changed.Emit(this.Info(e.Label)); });
    this._calculated = this._book.Calculated.Subscribe(() => this.Calculated.Emit(this.Info('Calculated')));
  }
  Info(label = '') { this.check(); return { revision: this.sequence, engineRevision: this._book.Revision, name: this._book.Name,
    activeWorksheet: this._book.ActiveWorksheet.Name, selection: this._selection, canUndo: this._book.CanUndo, canRedo: this._book.CanRedo, label }; }
  GetWorkbook() { return this.Workbook; }
  GetWorksheet(name = null) { const s = name == null ? this.Workbook.ActiveWorksheet : this.Workbook.Worksheets.Get(name); if (!s) throw new RangeError('Unknown worksheet'); return s; }
  GetRange(address, sheet = null) { return this.GetWorksheet(sheet).GetRange(address); }
  GetOfficeApi() { return createExcelApi(this.Workbook); }
  ReadBinding() { return { revision: this.sequence, value: JSON.stringify(this.Workbook.ToJSON()) }; }
  Save() { return this.Workbook.ToJSON(); }
  Load(document) { this.check(); if (this._collaboration) throw new Error('Disconnect collaboration before replacing the workbook.'); return this._replace(Core.Workbook.FromJSON(document)); }
  New() { this.check(); if (this._collaboration) throw new Error('Disconnect collaboration before replacing the workbook.'); return this._replace(new Core.Workbook()); }
  _replace(book) {
    if (this.disposed) { book.Dispose(); this.check(); }
    const old = this._book, owned = this._owned;
    this._bridge?.Dispose(); this._changed?.Dispose(); this._calculated?.Dispose();
    this._book = book; this._owned = true; this._selection = 'A1'; this.sequence++; this._subscribe();
    this._bridge = createHostBridge(this._headless);
    for (const view of this.views) { view.Workbook = book; view.Select('A1'); }
    if (owned) old.Dispose();
    this.Replaced.Emit(this.Info('Load workbook')); this.Changed.Emit(this.Info('Load workbook')); return true;
  }
  Attach(view) {
    this.check(); view.Workbook = this._book;
    this.views.add(view); return () => this.views.delete(view);
  }
  Select(address, sheet = null) {
    this.check(); Core.parseRange(address); if (sheet != null) this.ActivateWorksheet(sheet);
    this._selection = address; for (const v of this.views) { if (v.Sheet !== this._book.ActiveWorksheet) v.Sheet = this._book.ActiveWorksheet; v.Select(address); }
    return true;
  }
  ActivateWorksheet(name) {
    const s = this.GetWorksheet(name), changed = this._book.ActiveWorksheet !== s;
    this._book.ActiveWorksheet = s;
    for (const v of this.views) if (v.Sheet !== s) v.Sheet = s;
    if (changed) { this.sequence++; this.Changed.Emit(this.Info('Activate worksheet')); }
    return true;
  }
  /** Existing allowlisted shared-engine host protocol, with data rather than executable scripts. */
  Invoke(method, args = {}) {
    this.check(); object(args);
    if (method === 'workbook.load') return this.Load(args.workbook);
    if (method === 'workbook.new') return this.New();
    const response = JSON.parse(this._bridge.dispatch(JSON.stringify({ ...args, method, id: 1 })));
    this._bridge.drainEvents();
    if (response.error) throw new Error(response.error.message);
    return response.result;
  }
  /** Atomic synchronous native operations. Data matrices never become callback descriptors. */
  Batch(operations, label = 'Blazor batch') {
    this.check(); if (!Array.isArray(operations)) throw new TypeError('Expected operations array.');
    const allowed = new Set(['range.values.set','range.formulas.set','range.style.set','range.clear','range.merge','range.unmerge',
      'worksheets.add','worksheets.remove','worksheets.rename','sheet.freeze','sheet.insertRows','sheet.deleteRows','sheet.insertColumns','sheet.deleteColumns','names.define']);
    for (const op of operations) if (!allowed.has(op.method)) throw new TypeError(`Not an atomic mutation: ${op.method}`);
    return this._book.Transaction(label, () => operations.map(op => this.Invoke(op.method, op.arguments ?? {})));
  }
  Find(query, options = {}) { return dto(this.Workbook.Find(query, { ...options, sheet: options.sheet == null ? null : this.GetWorksheet(options.sheet) }).map(x => ({ sheet: x.sheet.Name, address: x.address, row: x.row, column: x.column, value: x.value }))); }
  Replace(query, replacement, options = {}) { return this.Workbook.Replace(query, replacement, { ...options, sheet: options.sheet == null ? null : this.GetWorksheet(options.sheet) }); }
  Copy(source, destination, sourceSheet = null, destinationSheet = null, mode = 'all') { this.GetRange(destination, destinationSheet).CopyFrom(this.GetRange(source, sourceSheet), mode); }
  Sort(address, keys, sheet = null, hasHeaders = true) { this.GetRange(address, sheet).Sort(keys, { hasHeaders }); }
  Fill(address, direction, sheet = null) { const r = this.GetRange(address, sheet); if (direction === 'down') r.FillDown(); else if (direction === 'right') r.FillRight(); else throw new TypeError('Fill direction must be down or right.'); }
  /** Complete public range/worksheet members remain available without duplicating the engine. */
  CallRange(address, method, args = [], sheet = null) { const [owner, key] = safeMember(this.GetRange(address, sheet), method); if (typeof owner[key] !== 'function') throw new TypeError('Member is not callable.'); return Reflect.apply(owner[key], owner, args); }
  CallWorksheet(sheet, method, args = []) { const [owner, key] = safeMember(this.GetWorksheet(sheet), method); if (typeof owner[key] !== 'function') throw new TypeError('Member is not callable.'); return Reflect.apply(owner[key], owner, args); }
  SetRangeProperty(address, property, value, sheet = null) { const [owner, key] = safeMember(this.GetRange(address, sheet), property); if (!(key in owner)) throw new TypeError('Unknown property.'); owner[key] = value; }
  ExportCsv(sheet = null, address = null, delimiter = ',') { const s = this.GetWorksheet(sheet); return IO.exportCSV(address ? s.GetRange(address) : s.UsedRange, { delimiter }); }
  ImportCsv(text, sheet = null, address = 'A1', options = {}) { const range = IO.importCSV(this.GetWorksheet(sheet), text, { ...options, address }); return range?.Address ?? null; }
  ExportXlsx(options = {}) { return IO.exportXlsx(this.Workbook, options); }
  async ImportXlsx(bytes) {
    this.check(); if (this._collaboration) throw new Error('Disconnect collaboration before replacing the workbook.');
    const revision = this.sequence, result = await IO.importXlsx(bytes);
    if (this.disposed || revision !== this.sequence || this._collaboration) { result.workbook.Dispose(); throw new Error('The workbook changed during import; the imported result was not applied.'); }
    this._replace(result.workbook); return result.warnings;
  }
  PrintHtml(sheet = null, options = {}) { return Core.createPrintDocument(this.GetWorksheet(sheet), options); }
  ChartSvg(sheet, chartId, options = {}) { const s = this.GetWorksheet(sheet), chart = s.Charts.find(c => c.id === chartId); if (!chart) throw new RangeError('Unknown chart'); return Core.chartToSVG(s, chart, options); }
  async Connect(options, mode = 'join') {
    this.check(); if (this._collaboration) throw new Error('Disconnect the existing collaboration session first.');
    const config = { ...options }; delete config.persistPending;
    if (options.persistPending) config.storage = globalThis.sessionStorage;
    const collaboration = new CollaborationSession(this.Workbook, config); this._collaboration = collaboration;
    this._collaborationSubscriptions = [collaboration.Changed.Subscribe(e => this.CollaborationChanged.Emit(e)),
      collaboration.Conflicted.Subscribe(e => this.CollaborationConflict.Emit(dto(e))),
      collaboration.PresenceChanged.Subscribe(e => this.PresenceChanged.Emit(dto(e))),
      collaboration.Error.Subscribe(e => this.Error.Emit({ message: e.message, code: e.code }))];
    try { await collaboration.Connect({ mode }); this.check(); return { status: collaboration.Status, revision: collaboration.Revision }; }
    catch (error) { if (this._collaboration === collaboration) await this.Disconnect(); else await collaboration.Dispose(); throw error; }
  }
  GetCollaboration() { this.check(); if (!this._collaboration) throw new Error('Collaboration is not connected.'); return this._collaboration; }
  Sync() { return this.GetCollaboration().Sync(); }
  ResolveConflict(strategy) { return this.GetCollaboration().Resolve(strategy); }
  SetPresence(sheet, selection) { return this.GetCollaboration().SetPresence(sheet, selection); }
  async Disconnect() { const c = this._collaboration; this._collaboration = null; for (const sub of this._collaborationSubscriptions ?? []) sub.Dispose(); this._collaborationSubscriptions = []; if (c) await c.Dispose(); }
  DisposeAsync() { return this._disposeTask ??= this._dispose(); }
  async _dispose() {
    if (this.disposed) return; this.disposed = true;
    try { await this.Disconnect(); }
    finally {
      this._bridge?.Dispose(); this._changed?.Dispose(); this._calculated?.Dispose();
      for (const view of this.views) view.remove?.(); this.views.clear();
      if (this._owned) this._book.Dispose();
      for (const name of ['Changed','Calculated','Replaced','Error','CollaborationChanged','CollaborationConflict','PresenceChanged']) this[name].Clear();
    }
  }
}
