import { EventSource } from '../../src/events.js';
import { isError } from '../../src/errors.js';
import { SpreadsheetSession } from './session.js';
/** Native control controller. Full workbook values are fetched by streaming, not event snapshots. */
export class SpreadsheetView {
  constructor(host) {
    this.host = host; this.grid = document.createElement('grid-web'); this.grid.style.cssText = 'display:block;width:100%;height:100%';
    this.initialWorkbook = this.grid.Workbook;
    this.disposed = false; this.initialized = false; this.selectionRevision = 0; this.bindingRevision = 0;
    for (const name of ['Changed','BindingChanged','SelectionChanged','CellEdited','ZoomChanged','Error']) this[name] = new EventSource();
    this.listeners = [];
    const listen = (event, fn) => { this.grid.addEventListener(event, fn); this.listeners.push([event, fn]); };
    listen('selection-change', e => {
      if (this.disposed) return;
      const d = e.detail; this.selectionRevision++;
      if (this.session) this.session._selection = d.address;
      this.SelectionChanged.Emit({ address: d.address, sheet: d.worksheet.Name, row: d.row, column: d.column, revision: this.selectionRevision });
    });
    listen('cell-edit', e => { const d = e.detail; this.CellEdited.Emit({ sheet: d.worksheet.Name, row: d.row, column: d.column, input: d.input, value: isError(d.value) ? { error: d.value.code } : d.value }); });
    listen('zoom-change', e => this.ZoomChanged.Emit(e.detail));
    listen('grid-error', e => this.Error.Emit(e.detail));
  }
  check() { if (this.disposed) throw new Error('The spreadsheet component has been disposed.'); }
  async Configure(o = {}) {
    this.check();
    if (!Number.isInteger(o.debounceMilliseconds ?? 150) || o.debounceMilliseconds < 0 || o.debounceMilliseconds > 60000) throw new RangeError('DebounceMilliseconds must be between 0 and 60000.');
    if (o.session != null && !(o.session instanceof SpreadsheetSession)) throw new TypeError('Workbook must be a GridWeb Blazor workbook.');
    if (o.session && o.value != null) throw new TypeError('Use either Workbook ownership or Value binding, not both.');
    const replacing = this.session == null || this.externalSession !== (o.session ?? null);
    if (replacing) {
      const next = o.session ?? new SpreadsheetSession(o.value ?? null);
      clearTimeout(this.timer); this.bindingRevision++;
      for (const sub of this.subscriptions ?? []) sub.Dispose(); this.detach?.();
      if (this.owned && this.session) await this.session.DisposeAsync();
      if (this.disposed) { if (!o.session) await next.DisposeAsync(); this.check(); }
      this.session = next; this.externalSession = o.session ?? null; this.owned = !o.session;
      this.detach = next.Attach(this.grid);
      this.initialWorkbook?.Dispose(); this.initialWorkbook = null;
      this.subscriptions = [next.Changed.Subscribe(e => {
        if (this.disposed) return;
        this.bindingRevision++;
        if (this.applying) return;
        this.Changed.Emit(e); clearTimeout(this.timer);
        if (this.enableBinding) this.timer = setTimeout(() => { if (!this.disposed) this.BindingChanged.Emit({ revision: this.bindingRevision }); }, this.debounce);
      }), next.Error.Subscribe(e => this.Error.Emit(e))];
    }
    this.debounce = o.debounceMilliseconds ?? 150; this.enableBinding = o.enableBinding ?? false;
    const force = !this.initialized || this.valueRevision !== (o.valueRevision ?? 0);
    const stale = !force && o.ackRevision >= 0 && o.ackRevision < this.bindingRevision;
    this.applying = true;
    try {
      if (!o.session && !replacing && !stale && (force || o.value !== this.lastValue)) {
        if (o.value == null) this.session.New();
        else if (this.session.ReadBinding().value !== o.value) this.session.Load(o.value);
      }
      this.lastValue = o.value; this.valueRevision = o.valueRevision ?? 0;
      const native = { Theme: o.theme ?? 'light', ReadOnly: o.readOnly ?? false, ShowGridLines: o.showGridLines ?? true, ViewMode: o.viewMode ?? 'normal', Zoom: o.zoom ?? 1, ...(o.native ?? {}) };
      for (const [key, value] of Object.entries(native)) {
        if (['Workbook','workbook','Model','Sheet','Selection','DataContext','__proto__','prototype','constructor'].includes(key) || key.startsWith('_')) throw new TypeError(`Use the explicit parameter for ${key}.`);
        if (!(key in this.grid)) throw new TypeError(`Unknown grid property ${key}.`);
        if (this.grid[key] !== value) this.grid[key] = value;
      }
      this.grid.setAttribute('aria-label', o.ariaLabel ?? 'Spreadsheet');
      if (o.sheet && (replacing || o.sheet !== this.lastSheet)) this.session.ActivateWorksheet(o.sheet);
      this.lastSheet = o.sheet;
      const forceSelection = !this.initialized || this.parameterSelectionRevision !== (o.selectionRevision ?? 0);
      const staleSelection = !forceSelection && o.ackSelectionRevision >= 0 && o.ackSelectionRevision < this.selectionRevision;
      if (o.selection && !staleSelection && (replacing || forceSelection || o.selection !== this.lastSelection) && this.grid.Selection !== o.selection) this.grid.Select(o.selection);
      this.lastSelection = o.selection; this.parameterSelectionRevision = o.selectionRevision ?? 0;
      if (!this.grid.isConnected) this.host.append(this.grid);
      this.initialized = true;
    } finally { this.applying = false; }
  }
  GetSession() { this.check(); return this.session; }
  GetNativeControl() { this.check(); return this.grid; }
  ReadBinding() { this.check(); return { ...this.session.ReadBinding(), revision: this.bindingRevision }; }
  FlushChanges() { this.check(); if (!this.grid.CommitEdit(false)) throw new Error('The active cell failed validation.'); clearTimeout(this.timer); return this.ReadBinding(); }
  Select(address, sheet = null) { this.check(); if (sheet != null) this.session.ActivateWorksheet(sheet); this.grid.Select(address); }
  Focus() { this.check(); this.grid.Focus(); }
  BeginEdit(value) { this.check(); return this.grid.BeginEdit(value ?? undefined); }
  CommitEdit() { this.check(); return this.grid.CommitEdit(); }
  CancelEdit() { this.check(); this.grid.CancelEdit(); }
  CopySelection() { this.check(); return this.grid.CopySelection(); }
  PasteText(text) { this.check(); this.grid.PasteText(text); }
  GetSelection() { this.check(); return { address: this.grid.Selection, sheet: this.grid.Sheet.Name, revision: this.selectionRevision }; }
  GetMetrics() { this.check(); return this.grid.Metrics; }
  DisposeAsync() { return this.disposeTask ??= this.dispose(); }
  async dispose() {
    if (this.disposed) return; this.disposed = true; clearTimeout(this.timer);
    for (const sub of this.subscriptions ?? []) sub.Dispose(); this.subscriptions = [];
    for (const [name, fn] of this.listeners) this.grid.removeEventListener(name, fn); this.listeners = [];
    this.detach?.(); this.grid.remove(); this.grid.Dispose(); this.initialWorkbook?.Dispose(); this.initialWorkbook = null;
    try { if (this.owned && this.session) await this.session.DisposeAsync(); }
    finally { for (const name of ['Changed','BindingChanged','SelectionChanged','CellEdited','ZoomChanged','Error']) this[name].Clear(); }
  }
}
