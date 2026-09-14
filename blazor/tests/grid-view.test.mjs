import { test } from 'node:test';
import assert from 'node:assert/strict';
import { SpreadsheetView } from '../src/view.js';
import { SpreadsheetSession } from '../src/session.js';
import { Workbook } from '../../src/core.js';

// A minimal native control contract for deterministic parameter/lifetime tests.
// Actual canvas input and layout are separately qualified in both package-restored browsers.
class NativeGrid extends EventTarget {
  constructor() { super(); this.Workbook = new Workbook(); this.style = {}; this.Theme='light'; this.ReadOnly=false; this.ShowGridLines=true; this.ViewMode='normal'; this.Zoom=1; this.Selection='A1'; this.isConnected=false; this.valid=true; this.closed=0; }
  get Sheet() { return this.Workbook.ActiveWorksheet; }
  set Sheet(value) { this.Workbook.ActiveWorksheet=value; }
  Select(address) { this.Selection=address; const e=new Event('selection-change');e.detail={address,worksheet:this.Sheet,row:0,column:0};this.dispatchEvent(e); }
  setAttribute() {}
  CommitEdit() { return this.valid; }
  remove() { this.isConnected=false; }
  Dispose() { this.closed++; }
}
function view() {
  const previous=globalThis.document;globalThis.document={createElement:()=>new NativeGrid()};
  try { return new SpreadsheetView({append(grid) { grid.isConnected=true; }}); }
  finally { if (previous===undefined) delete globalThis.document;else globalThis.document=previous; }
}
const wait=()=>new Promise(r=>setTimeout(r,20));
test('stale echoed workbook values cannot overwrite newer edits; explicit revisions reset',async()=>{
 const v=view();await v.Configure({enableBinding:true,debounceMilliseconds:0});
 v.session.GetRange('A1').Value=1;const old=v.ReadBinding();v.session.GetRange('A1').Value=2;
 await v.Configure({value:old.value,ackRevision:old.revision,enableBinding:true,debounceMilliseconds:0});
 assert.equal(v.session.GetRange('A1').Value,2);
 await v.Configure({value:old.value,ackRevision:old.revision,valueRevision:1});assert.equal(v.session.GetRange('A1').Value,1);
 await v.DisposeAsync();
});
test('switching owned/shared sessions preserves monotonic binding and borrowed workbook lifetime',async()=>{
 const v=view(), shared=new SpreadsheetSession();await v.Configure({});
 for(let i=0;i<5;i++)v.session.GetRange('A1').Value=i;
 const prior=v.ReadBinding().revision;
 await v.Configure({session:shared});assert.ok(v.ReadBinding().revision>prior);
 shared.GetRange('A1').Value=6;const second=v.ReadBinding().revision;
 await v.Configure({});assert.ok(v.ReadBinding().revision>second);assert.equal(shared.disposed,false);
 await v.DisposeAsync();assert.equal(shared.GetRange('A1').Value,6);await shared.DisposeAsync();
});
test('owned replacement cancels old binding timers and disposal suppresses queued notifications',async()=>{
 const v=view(), notices=[];v.BindingChanged.Subscribe(e=>notices.push(e));await v.Configure({enableBinding:true,debounceMilliseconds:5});
 v.session.GetRange('A1').Value=1;const shared=new SpreadsheetSession();await v.Configure({session:shared,enableBinding:true,debounceMilliseconds:5});
 await wait();assert.equal(notices.length,0);
 shared.GetRange('A1').Value=2;await v.DisposeAsync();await wait();assert.equal(notices.length,0);await shared.DisposeAsync();
});
test('repeated cleanup disposes the native grid once and rejected cell validation is visible',async()=>{
 const v=view();await v.Configure({});v.grid.valid=false;assert.throws(()=>v.FlushChanges(),/validation/);
 await Promise.all([v.DisposeAsync(),v.DisposeAsync()]);assert.equal(v.grid.closed,1);assert.throws(()=>v.ReadBinding(),/disposed/);
});
test('invalid mixing of shared and bound workbooks and private native properties fails explicitly',async()=>{
 const v=view(), shared=new SpreadsheetSession();await assert.rejects(v.Configure({session:shared,value:'{}'}),/either/);
 await v.Configure({});await assert.rejects(v.Configure({native:{_private:true}}),/explicit/);
 await assert.rejects(v.Configure({debounceMilliseconds:60001}),/Debounce/);await v.DisposeAsync();await shared.DisposeAsync();
});
test('selection acknowledgment ignores an older echo and an explicit revision can restore it',async()=>{
 const v=view();await v.Configure({selection:'A1'});v.grid.Select('B2');const revision=v.selectionRevision;v.grid.Select('C3');
 await v.Configure({selection:'B2',ackSelectionRevision:revision});assert.equal(v.grid.Selection,'C3');
 await v.Configure({selection:'B2',ackSelectionRevision:revision,selectionRevision:1});assert.equal(v.grid.Selection,'B2');await v.DisposeAsync();
});
