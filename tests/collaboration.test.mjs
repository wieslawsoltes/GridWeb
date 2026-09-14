import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import {Workbook} from '../src/core.js';
import {CollaborationSession, CollaborationStore, diffDocuments, rebaseDocuments, applyChanges} from '../src/collaboration.js';
import {validateDocument, applyDocument} from '../src/collaboration/document.js';
import {createCollaborationServer, bearerAuthorization, createFileStorage} from '../src/collaboration/server.js';
import {createValidationPool} from '../src/collaboration/validation.js';

const token='gridweb-test-token-with-32-characters';
const doc=()=>new Workbook().ToJSON();
const edit=(document,address,value)=>{const b=Workbook.FromJSON(document);b.ActiveWorksheet.GetCell(address).Value=value;return b.ToJSON();};
const request=(base,local,id='first')=>({...diffDocuments(base.document,local),id,clientId:'client',baseRevision:base.revision,epoch:base.epoch});
const memory=()=>{const records=new Map();return {records,load:async n=>structuredClone(records.get(n)??null),save:async(n,s)=>records.set(n,structuredClone(s))};};
const storage=()=>{const records=new Map();return {records,getItem:k=>records.get(k)??null,setItem:(k,v)=>records.set(k,v)};};
const until=async predicate=>{for(let i=0;i<150;i++){if(await predicate())return;await new Promise(r=>setTimeout(r,20));}assert.fail('Timed out waiting for synchronization');};
async function setup(t,options={}) {
 const persistence=memory();const app=await createCollaborationServer({authorize:bearerAuthorization(token),storage:persistence,...options});
 await new Promise(r=>app.server.listen(0,'127.0.0.1',r));
 const url=`http://127.0.0.1:${app.server.address().port}`,sessions=[];
 t.after(async()=>{for(const s of sessions)s.Dispose();await app.close();});
 const session=(options={})=>{const s=new CollaborationSession(new Workbook(),{url,room:'test',token,autoSync:false,...options});sessions.push(s);return s;};
 return {app,url,persistence,session};
}

test('collaboration compares content, style and comment independently',()=>{
 const base=doc(),a=Workbook.FromJSON(base),b=Workbook.FromJSON(base);
 a.ActiveWorksheet.GetCell('A1').Value=10;b.ActiveWorksheet.GetCell('A1').Style={font:{bold:true}};b.ActiveWorksheet.GetCell('A1').Comment='Reviewed';
 const merged=rebaseDocuments(base,a.ToJSON(),b.ToJSON()).document,c=Workbook.FromJSON(merged).ActiveWorksheet.GetCell('A1');
 assert.equal(c.Value,10);assert.equal(c.Style.font.bold,true);assert.equal(c.Comment,'Reviewed');
});
test('conflicting cell fields never partially apply a commit',()=>{
 const base=doc(),local=edit(edit(base,'A1',10),'B1',20),remote=edit(base,'B1',99);
 assert.throws(()=>applyChanges(remote,diffDocuments(base,local).changes),e=>e.code==='CONFLICT');
 assert.equal(Workbook.FromJSON(remote).ActiveWorksheet.GetCell('A1').Value,null);
});
for(const choice of ['local','remote'])test('explicit '+choice+' resolution preserves disjoint edits',()=>{
 const base=doc(),local=edit(edit(base,'A1',10),'B1',20),remote=edit(edit(base,'A1',99),'C1',30);
 const b=Workbook.FromJSON(rebaseDocuments(base,local,remote,{resolve:choice}).document);
 assert.deepEqual(b.ActiveWorksheet.GetRange('A1:C1').Values,[[choice==='local'?10:99,20,30]]);
});
test('whole-workbook edits require explicit conflict resolution',()=>{
 const base=doc(),local=Workbook.FromJSON(base);local.Worksheets.Add('Added');
 assert.throws(()=>rebaseDocuments(base,local.ToJSON(),edit(base,'A1',2)),e=>e.code==='CONFLICT');
});
for(const mutate of [d=>d.sheets.push(d.sheets[0]),d=>d.sheets[0].cells.push([-1,{}]),d=>d.sheets[0].cells.push([0,{}],[0,{}]),d=>d.sheets[0].id='../bad'])test('invalid shared document is rejected: '+String(mutate),()=>{const d=doc();mutate(d);assert.throws(()=>validateDocument(d));});
test('remote apply preserves Workbook and Worksheet identity and recalculates',()=>{
 const b=new Workbook(),s=b.ActiveWorksheet;b.ActiveWorksheet.GetCell('B1').Formula='=A1*2';
 applyDocument(b,edit(b.ToJSON(),'A1',7));assert.equal(b.ActiveWorksheet,s);assert.equal(s.GetCell('B1').Value,14);assert.equal(b.CanUndo,false);
});
test('store serializes disjoint concurrent commits and durably saves before acknowledgment',async()=>{
 const disk=memory(),store=new CollaborationStore(disk),base=await store.create('r',doc());
 const a=request(base,edit(base.document,'A1',1),'a'),b=request(base,edit(base.document,'B1',2),'b');
 await Promise.all([store.commit('r',a),store.commit('r',b)]);
 const result=await new CollaborationStore(disk).get('r');assert.equal(result.revision,2);assert.deepEqual(Workbook.FromJSON(result.document).ActiveWorksheet.GetRange('A1:B1').Values,[[1,2]]);
});
test('duplicate acknowledgments are idempotent and identities cannot change payload',async()=>{
 const s=new CollaborationStore(),base=await s.create('r',doc()),r=request(base,edit(base.document,'A1',1));
 await s.commit('r',r);const retry=await s.commit('r',r);assert.equal(retry.revision,1);assert.equal(retry.duplicate,true);
 r.changes[0].value.input=2;await assert.rejects(s.commit('r',r),e=>e.code==='INVALID');
});
test('structural epoch rejects stale cell locations',async()=>{
 const s=new CollaborationStore(),base=await s.create('r',doc()),b=Workbook.FromJSON(base.document);b.Worksheets.Add('New');await s.commit('r',request(base,b.ToJSON()));
 await assert.rejects(s.commit('r',request(base,edit(base.document,'A1',1),'other')),e=>e.code==='CONFLICT');
});
test('receipt expiry refuses unknown stale requests instead of applying them twice',async()=>{
 const s=new CollaborationStore(),base=await s.create('r',doc());let cur=base;
 for(let i=0;i<130;i++)cur=await s.commit('r',request(cur,edit(cur.document,'B1',i),String(i)));
 await assert.rejects(s.commit('r',request(base,edit(base.document,'A1',1),'expired')),e=>e.code==='CONFLICT');
});
test('a save failure after rename is recovered from durable state on retry',async()=>{
 const disk=memory();let fail=false;
 const s=new CollaborationStore({...disk,save:async(n,v)=>{await disk.save(n,v);if(fail){fail=false;throw Error('disk flush interrupted');}}});
 const base=await s.create('r',doc()),r=request(base,edit(base.document,'A1',3));fail=true;
 await assert.rejects(s.commit('r',r));const retry=await s.commit('r',r);assert.equal(retry.duplicate,true);assert.equal(retry.revision,1);
});
test('file storage persists atomically and confines room paths',async t=>{
 const dir=await fs.mkdtemp(path.join(os.tmpdir(),'gridweb-store-'));t.after(()=>fs.rm(dir,{recursive:true,force:true}));
 const disk=await createFileStorage(dir),s=new CollaborationStore(disk);await s.create('safe',doc());assert.equal((await new CollaborationStore(disk).get('safe')).revision,0);
 await assert.rejects(disk.load('../escape'));assert.deepEqual(await fs.readdir(dir),['safe.json']);
});
test('server requires auth, enforces editor role and explicit origins',async t=>{
 const auth=req=>req.headers.authorization==='Bearer '+token?{id:'test',role:'viewer'}:null;
 const {url}=await setup(t,{authorize:auth,allowedOrigins:['https://example.test']});
 assert.equal((await fetch(url+'/rooms/test')).status,401);
 assert.equal((await fetch(url+'/rooms/test',{headers:{Origin:'https://evil.test'}})).status,403);
 assert.equal((await fetch(url+'/rooms/test',{method:'PUT',headers:{Authorization:'Bearer '+token,'Content-Type':'application/json'},body:JSON.stringify({document:doc()})})).status,403);
 const preflight=await fetch(url+'/rooms/test',{method:'OPTIONS',headers:{Origin:'https://example.test'}});assert.equal(preflight.status,204);assert.equal(preflight.headers.get('access-control-allow-origin'),'https://example.test');
});
test('server rejects invalid request bodies and presence ranges',async t=>{
 const {url,app}=await setup(t);await app.store.create('test',doc());
 const headers={Authorization:'Bearer '+token,'Content-Type':'application/json'};
 assert.equal((await fetch(url+'/rooms/test/commits',{method:'POST',headers,body:'{bad'})).status,400);
 const p=await fetch(url+'/rooms/test/presence',{method:'POST',headers,body:JSON.stringify({clientId:'client',sheet:'missing',selection:'A1'})});assert.equal(p.status,400);
});
test('mandatory server and transport limits are validated',async()=>{
 await assert.rejects(createCollaborationServer(),/authorization/);assert.throws(()=>bearerAuthorization('weak'));
 for(const url of ['http://example.com','https://user:pass@example.com','https://example.com?token=secret'])assert.throws(()=>new CollaborationSession(new Workbook(),{url,room:'test',token}));
});
test('two real HTTP clients merge disjoint cell edits and retain formula dependencies',async t=>{
 const {session}=await setup(t),a=session(),b=session();await a.Connect({mode:'create'});await b.Connect();const ws=b.Workbook.ActiveWorksheet;
 a.Workbook.ActiveWorksheet.GetCell('A1').Value=10;b.Workbook.ActiveWorksheet.GetCell('B1').Formula='=A1*2';
 await Promise.all([a.Sync(),b.Sync()]);await a.Sync();await b.Sync();
 assert.deepEqual(a.Workbook.ActiveWorksheet.GetRange('A1:B1').Values,[[10,20]]);assert.deepEqual(b.Workbook.ActiveWorksheet.GetRange('A1:B1').Values,[[10,20]]);assert.equal(b.Workbook.ActiveWorksheet,ws);
});
for(const strategy of ['local','remote'])test('real same-cell conflict resolves '+strategy+' without losing other local edits',async t=>{
 const {session}=await setup(t),a=session(),b=session();await a.Connect({mode:'create'});await b.Connect();a.Workbook.ActiveWorksheet.GetCell('A1').Value=1;b.Workbook.ActiveWorksheet.GetRange('A1:B1').Values=[[2,3]];
 await a.Sync();assert.equal(await b.Sync(),false);assert.equal(b.Status,'conflict');assert.equal(b.Workbook.ActiveWorksheet.GetCell('B1').Value,3);
 await b.Resolve(strategy);await a.Sync();assert.deepEqual(a.Workbook.ActiveWorksheet.GetRange('A1:B1').Values,[[strategy==='local'?2:1,3]]);
});
test('offline edits survive reload and do not persist authentication tokens',async t=>{
 const {session}=await setup(t),cache=storage(),a=session({storage:cache});await a.Connect({mode:'create'});a.Workbook.ActiveWorksheet.GetCell('A1').Value=42;a.Dispose();
 assert.ok([...cache.records.values()].every(v=>!v.includes(token)));
 const restored=session({storage:cache});await restored.Connect();assert.equal(restored.Workbook.ActiveWorksheet.GetCell('A1').Value,42);assert.equal(restored.HasPendingChanges,false);
});
test('lost acknowledgment retries the same commit after reload exactly once',async t=>{
 const {session,app}=await setup(t),cache=storage();let drop=true;
 const transport=async(url,options)=>{const r=await fetch(url,options);if(drop&&url.endsWith('/commits')){drop=false;await r.text();throw Error('lost response');}return r;};
 const a=session({storage:cache,fetch:transport});await a.Connect({mode:'create'});a.Workbook.ActiveWorksheet.GetCell('A1').Value=4;await assert.rejects(a.Sync(),/lost response/);a.Dispose();
 const b=session({storage:cache});await b.Connect();assert.equal(b.Workbook.ActiveWorksheet.GetCell('A1').Value,4);assert.equal((await app.store.get('test')).revision,1);
});
test('new edits during an in-flight request are rebased rather than overwritten',async t=>{
 const {session}=await setup(t);let release,received;const entered=new Promise(r=>received=r);
 const a=session({fetch:async(url,options)=>{const response=await fetch(url,options);if(url.endsWith('/commits')){received();await new Promise(r=>release=r);}return response;}});await a.Connect({mode:'create'});a.Workbook.ActiveWorksheet.GetCell('A1').Value=1;
 const flight=a.Sync();await entered;a.Workbook.ActiveWorksheet.GetCell('B1').Value=2;release();await flight;
 assert.deepEqual(a.Workbook.ActiveWorksheet.GetRange('A1:B1').Values,[[1,2]]);assert.equal(a.HasPendingChanges,true);a.Dispose();
});
test('SSE delivers remote revisions and presence without polling timers',async t=>{
 const {session}=await setup(t),a=session({autoSync:true}),b=session({autoSync:true});await a.Connect({mode:'create'});await b.Connect();let people=[];b.PresenceChanged.Subscribe(p=>people=p);
 a.Workbook.ActiveWorksheet.GetCell('A1').Value=321;await until(()=>b.Workbook.ActiveWorksheet.GetCell('A1').Value===321);
 await a.SetPresence(a.Workbook.ActiveWorksheet.Id,'B2');await until(()=>people.some(p=>p.selection==='B2'));assert.equal(b.HasPendingChanges,false);
});
test('edits during initial join are not silently overwritten',async t=>{
 const {session}=await setup(t),a=session();await a.Connect({mode:'create'});let release,entered;const signal=new Promise(r=>entered=r);
 let first=true;const b=session({fetch:async(url,options)=>{const response=await fetch(url,options);if(first){first=false;entered();await new Promise(r=>release=r);}return response;}});
 const joining=b.Connect();await signal;b.Workbook.ActiveWorksheet.GetCell('A1').Value=9;release();await assert.rejects(joining,e=>e.code==='BUSY');assert.equal(b.Workbook.ActiveWorksheet.GetCell('A1').Value,9);
});
test('invalid offline storage is retained for user recovery',async t=>{
 const {session}=await setup(t),cache=storage();const a=session({storage:cache,storageKey:'broken'});cache.setItem('broken','{broken');await assert.rejects(a.Connect(),e=>e.code==='STORAGE');assert.equal(cache.getItem('broken'),'{broken');
});
test('validation workers return sanitized workbooks and reject unknown shapes',async t=>{
 const pool=createValidationPool();t.after(()=>pool.close());assert.equal((await pool.validate(doc())).format,'GridWeb');await assert.rejects(pool.validate({}),e=>e.code==='INVALID');
});
