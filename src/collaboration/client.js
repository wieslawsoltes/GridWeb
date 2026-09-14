import {Workbook} from '../model.js';
import { EventSource } from '../events.js';
import { diffDocuments, rebaseDocuments, applyDocument, equal, CollaborationError, validateDocument } from './document.js';
const uuid=()=>globalThis.crypto.randomUUID();
const clone=v=>structuredClone(v);
const structural=e=>e.Changes?.some(x=>/^(insert-|delete-|sheet-rename|sheet-add|sheet-remove)/.test(x.type));
/** Shared-engine browser client with durable pending edits and explicit conflict resolution. */
export class CollaborationSession {
  constructor(book,{url,room,token,fetch:fetcher=globalThis.fetch,storage=null,storageKey=null,clientId=null,autoSync=true}={}) {
    if(!(book instanceof Workbook))throw new TypeError('A GridWeb Workbook is required');
    const endpoint=new URL(url);
    if(endpoint.username||endpoint.password||endpoint.search||endpoint.hash)throw new TypeError('Endpoint must not contain credentials, query or fragment');
    if(endpoint.protocol!=='https:'&&!(endpoint.protocol==='http:'&&['localhost','127.0.0.1','[::1]'].includes(endpoint.hostname)))throw new TypeError('Use HTTPS, except for loopback development');
    if(!/^[-\w]{1,80}$/.test(room??''))throw new TypeError('Invalid room name');
    if(typeof token!=='string'||!token)throw new TypeError('Authentication token required');
    this.Workbook=book;this.Url=endpoint.href.replace(/\/$/,'');this.Room=room;this._token=token;this._fetch=fetcher;this._storage=storage;
    this._key=storageKey??'gridweb.collaboration:'+this.Url+':'+room;this.ClientId=clientId??uuid();this.AutoSync=autoSync;
    this.Changed=new EventSource();this.Conflicted=new EventSource();this.PresenceChanged=new EventSource();this.Error=new EventSource();
    this.Status='disconnected';this.Conflict=null;this._base=null;this._pending=null;this._exclusive=false;this._structuralVersion=0;this._applying=false;this._disposed=false;this._connected=false;this._flight=null;this._controller=new AbortController();this._desiredRevision=-1;this._timer=null;
  }
  _state(status){this.Status=status;this.Changed.Emit({status,revision:this.Revision,pending:this.HasPendingChanges});}
  get Revision(){return this._base?.revision??-1;}
  get HasPendingChanges(){if(!this._base)return false;const d=diffDocuments(this._base.document,this.Workbook.ToJSON(),{exclusive:this._exclusive});return !!this._pending||d.kind==='replace'||!!d.changes.length;}
  async _request(path='',method='GET',body=null){
    const response=await this._fetch(`${this.Url}/rooms/${encodeURIComponent(this.Room)}${path}`,{method,headers:{Authorization:'Bearer '+this._token,...(body?{'Content-Type':'application/json'}:{})},body:body?JSON.stringify(body):undefined,signal:this._controller.signal,cache:'no-store'});
    const value=await response.json();if(!response.ok)throw new CollaborationError(value.error?.code??'HTTP_ERROR',value.error?.message??'Request failed',value.error?.details);return value;
  }
  _persist(){
    if(!this._storage||!this._base)return;
    const record={version:1,clientId:this.ClientId,base:this._base,local:this.Workbook.ToJSON(),pending:this._pending,exclusive:this._exclusive,structuralVersion:this._structuralVersion};
    try{this._storage.setItem(this._key,JSON.stringify(record));}catch(e){this.Error.Emit(new CollaborationError('STORAGE','Offline edits could not be saved locally',e.message));}
  }
  _apply(document){const difference=diffDocuments(this.Workbook.ToJSON(),document);if(difference.kind==='cells'&&!difference.changes.length)return;this._applying=true;try{applyDocument(this.Workbook,document);}finally{this._applying=false;}}
  async Connect({mode='join'}={}){
    if(this._disposed||this._connected)throw new Error('Session cannot be connected again');
    if(!['join','create'].includes(mode))throw new TypeError('Connect mode must be join or create');
    this._state('connecting');let saved=null;
    const raw=this._storage?.getItem(this._key);if(raw){try{saved=JSON.parse(raw);if(saved.version!==1)throw new Error('Unknown offline format');validateDocument(saved.local);validateDocument(saved.base.document);}catch(e){this._state('error');throw new CollaborationError('STORAGE','Stored offline edits are invalid; back them up before resetting',e.message);}}
    const remote=mode==='create'?await this._request('','PUT',{document:this.Workbook.ToJSON()}):await this._request();
    this._base=saved?.base??remote;this.ClientId=saved?.clientId??this.ClientId;this._pending=saved?.pending??null;this._exclusive=saved?.exclusive??false;this._structuralVersion=saved?.structuralVersion??0;
    this._apply(saved?.local??remote.document);this._desiredRevision=remote.revision;this._connected=true;
    this._sub=this.Workbook.Changed.Subscribe(e=>{if(this._applying)return;if(structural(e)){this._exclusive=true;this._structuralVersion++;}this._persist();this._state(this.Conflict?'conflict':'pending');this._schedule();});
    this._persist();this._state('connected');
    // Do not overwrite a saved in-flight request: retry its immutable identity first.
    await this.Sync();if(this.AutoSync)this._events();return this;
  }
  _schedule(delay=150){if(!this.AutoSync||this._disposed||this.Conflict||this.Status==='error')return;clearTimeout(this._timer);this._timer=setTimeout(()=>{this.Sync().catch(e=>this.Error.Emit(e));},delay);this._timer.unref?.();}
  async Sync(){
    if(!this._connected||this._disposed)throw new Error('Session is not connected');if(this.Conflict)return false;if(this._flight)return this._flight;
    this._flight=this._sync().finally(()=>{this._flight=null;if(this.AutoSync&&!this._disposed&&!this.Conflict&&(this.HasPendingChanges||this._desiredRevision>this.Revision))this._schedule(this.Status==='offline'?2000:150);});return this._flight;
  }
  async _sync(){
    this._state('syncing');
    try{
      const local=this.Workbook.ToJSON();
      if(!this._pending){const delta=diffDocuments(this._base.document,local,{exclusive:this._exclusive});if(delta.kind==='replace'||delta.changes.length)this._pending={request:{...delta,id:uuid(),clientId:this.ClientId,baseRevision:this.Revision,epoch:this._base.epoch},sent:clone(local),structuralVersion:this._structuralVersion};}
      this._persist();
      if(this._pending){
        const pending=this._pending,remote=await this._request('/commits','POST',pending.request),current=this.Workbook.ToJSON();
        const exclusiveAfter=this._structuralVersion!==(pending.structuralVersion??0),expectedEpoch=pending.request.epoch+(pending.request.kind==='replace'?1:0);
        this._base={document:pending.sent,revision:remote.appliedRevision,epoch:expectedEpoch};this._pending=null;this._exclusive=exclusiveAfter;this._persist();
        const newEdits=diffDocuments(pending.sent,current,{exclusive:exclusiveAfter});
        if(remote.epoch!==expectedEpoch&&(newEdits.kind==='replace'||newEdits.changes.length))throw new CollaborationError('CONFLICT','Structure changed after the acknowledged commit',[{field:'structure'}]);
        const merged=rebaseDocuments(pending.sent,current,remote.document,{exclusive:exclusiveAfter});
        this._base=remote;this._apply(merged.document);this._persist();
      }else if(this._desiredRevision>this.Revision||!this.AutoSync){
        const remote=await this._request(),current=this.Workbook.ToJSON();
        // Edits may occur while fetching. Rebase rather than replace those edits.
        const delta=diffDocuments(this._base.document,current,{exclusive:this._exclusive});
        if(remote.epoch!==this._base.epoch&&(delta.kind==='replace'||delta.changes.length))throw new CollaborationError('CONFLICT','Structure changed while local edits were pending',[{field:'structure'}]);
        const merged=rebaseDocuments(this._base.document,current,remote.document,{exclusive:this._exclusive});this._base=remote;this._apply(merged.document);this._persist();
      }
      this._state(this.HasPendingChanges?'pending':'connected');return true;
    }catch(e){
      if(this._disposed)throw e;
      if(e.code==='CONFLICT'){
        const remote=await this._request();this.Conflict={details:e.details,base:clone(this._base),local:this.Workbook.ToJSON(),remote};this._persist();this._state('conflict');this.Conflicted.Emit(this.Conflict);return false;
      }
      this._state(e.code==='UNAUTHORIZED'||e.code==='FORBIDDEN'?'error':'offline');throw e;
    }
  }
  async Resolve(strategy){
    if(!['local','remote'].includes(strategy)||!this.Conflict)throw new TypeError('A conflict and explicit local/remote resolution are required');
    const conflict=this.Conflict,remote=await this._request();
    let result;
    if(remote.epoch!==conflict.base.epoch){result=strategy==='remote'?remote.document:this.Workbook.ToJSON();this._exclusive=strategy==='local';}
    else result=rebaseDocuments(conflict.base.document,this.Workbook.ToJSON(),remote.document,{exclusive:this._exclusive,resolve:strategy}).document;
    this._base=remote;this._pending=null;this.Conflict=null;this._apply(result);this._persist();this._state('pending');return this.Sync();
  }
  async SetPresence(sheet,selection){if(!this._connected||this._disposed)return;return this._request('/presence','POST',{clientId:this.ClientId,sheet,selection});}
  async _events(){
    let delay=500;
    while(!this._disposed){
      try{
        const response=await this._fetch(`${this.Url}/rooms/${encodeURIComponent(this.Room)}/events`,{headers:{Authorization:'Bearer '+this._token},signal:this._controller.signal,cache:'no-store'});
        if(!response.ok||!response.body)throw new Error('Event stream unavailable');
        delay=500;const reader=response.body.getReader(),decoder=new TextDecoder();let buffer='';
        try{while(!this._disposed){const {done,value}=await reader.read();if(done)break;buffer+=decoder.decode(value,{stream:true}).replace(/\r\n/g,'\n');if(buffer.length>1024*1024)throw new Error('Event frame exceeds limit');let index;
          while((index=buffer.indexOf('\n\n'))>=0){const frame=buffer.slice(0,index);buffer=buffer.slice(index+2);const data=frame.split('\n').filter(x=>x.startsWith('data:')).map(x=>x.slice(5).trimStart()).join('\n');if(!data)continue;const event=JSON.parse(data);if(event.type==='revision'){this._desiredRevision=Math.max(this._desiredRevision,event.revision);if(this._desiredRevision>this.Revision)this._schedule(0);}else if(event.type==='presence')this.PresenceChanged.Emit(event.participants);}
        }}finally{await reader.cancel().catch(()=>{});reader.releaseLock();}
      }catch(e){if(this._disposed)return;this.Error.Emit(e);}
      if(!this._disposed)await new Promise(resolve=>{const signal=this._controller.signal;const done=()=>{clearTimeout(timer);signal.removeEventListener('abort',done);resolve();};const timer=setTimeout(done,delay);timer.unref?.();signal.addEventListener('abort',done,{once:true});});delay=Math.min(30000,delay*2);
    }
  }
  Dispose(){if(this._disposed)return;this._persist();this._disposed=true;this._connected=false;clearTimeout(this._timer);this._controller.abort();this._sub?.Dispose();this._token='';this._state('disconnected');for(const event of [this.Changed,this.Conflicted,this.PresenceChanged,this.Error])event.Clear();}
}
