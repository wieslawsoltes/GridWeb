import test from 'node:test';import assert from 'node:assert/strict';
import {Workbook} from '../src/core.js';
import {CollaborationSession} from '../src/collaboration.js';
import {applyDocument} from '../src/collaboration/document.js';
test('fetch receives the browser global rather than the session as its receiver',async()=>{
 const b=new Workbook(),remote={revision:0,epoch:0,document:b.ToJSON()};
 const session=new CollaborationSession(b,{url:'https://example.test',room:'test',token:'unit-test',autoSync:false,fetch:function(){assert.equal(this,globalThis);return Promise.resolve({ok:true,json:async()=>remote});}});
 try{await session.Connect();assert.equal(session.Status,'connected');}finally{session.Dispose();}
});
test('remote snapshots replace formula indexes, used extents and stale filter state',()=>{
 const b=new Workbook(),s=b.ActiveWorksheet;s.GetCell('A1').Value=3;
 assert.equal(s.UsedRange.Address,'A1');
 const remote=Workbook.FromJSON(b.ToJSON());remote.ActiveWorksheet.GetCell('D5').Formula='=A1*2';
 applyDocument(b,remote.ToJSON());assert.equal(s.UsedRange.Address,'A1:D5');assert.equal(s._formulaCells.has(4*16384+3),true);
 s.GetCell('A1').Value=8;b.Calculate(true);assert.equal(s.GetCell('D5').Value,16);
 remote.ActiveWorksheet.GetCell('D5').Value=null;applyDocument(b,remote.ToJSON());assert.equal(s._formulaCells.size,0);
 b.Dispose();remote.Dispose();
});
test('disabling automatic synchronization cancels an already scheduled write',async()=>{
 const b=new Workbook(),remote={revision:0,epoch:0,document:b.ToJSON()};let commits=0;
 const session=new CollaborationSession(b,{url:'https://example.test',room:'test',token:'unit-test',autoSync:false,fetch:async (url,options)=>{if(options.method==='POST')commits++;return {ok:true,json:async()=>remote};}});
 await session.Connect();session._autoSync=true;b.ActiveWorksheet.GetCell('A1').Value=5;session.AutoSync=false;
 await new Promise(r=>setTimeout(r,200));assert.equal(commits,0);assert.equal(session.HasPendingChanges,true);session.Dispose();
});
