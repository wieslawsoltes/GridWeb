import { applyChanges, validateDocument, canonical, CollaborationError } from './document.js';
const copy=v=>structuredClone(v);
export const roomName=value=>{if(typeof value!=='string'||!/[a-z0-9]/i.test(value)||!/^[-\w]{1,80}$/.test(value))throw new CollaborationError('INVALID','Room names use 1–80 letters, digits, underscores or hyphens');return value;};
/** Transport/storage-independent, serialized room commits. Durable save precedes acknowledgment. */
export class CollaborationStore {
  constructor({load=async()=>null,save=async()=>{},notify=()=>{},maxRooms=32,validate=validateDocument,apply=applyChanges}={}) {if(!Number.isInteger(maxRooms)||maxRooms<1||maxRooms>1000)throw new RangeError('Invalid room limit');this.load=load;this.save=save;this.notify=notify;this.maxRooms=maxRooms;this.validate=validate;this.apply=apply;this.rooms=new Map();this.queues=new Map();this.faulted=new Set();this.pendingCalls=0;}
  _serial(name,action){name=roomName(name);if(this.pendingCalls>=32)return Promise.reject(new CollaborationError('LIMIT','Room operation queue is full'));const active=new Set([...this.rooms.keys(),...this.queues.keys()]);if(!active.has(name)&&active.size>=this.maxRooms)return Promise.reject(new CollaborationError('LIMIT','Loaded room limit reached'));this.pendingCalls++;const prior=this.queues.get(name)??Promise.resolve();const result=prior.catch(()=>{}).then(action);this.queues.set(name,result);result.finally(()=>{this.pendingCalls--;if(this.queues.get(name)===result)this.queues.delete(name);}).catch(()=>{});return result;}
  async _get(name){if(this.faulted.has(name))throw new CollaborationError('STORAGE','Room storage needs recovery before further access');if(!this.rooms.has(name)){if(this.rooms.size>=this.maxRooms)throw new CollaborationError('LIMIT','Loaded room limit reached');const state=copy(await this.load(name));if(state){if(state.format!=='GridWebRoom'||state.version!==1||!Number.isSafeInteger(state.revision)||!Number.isSafeInteger(state.epoch)||state.revision<0||state.epoch<0||!Array.isArray(state.receipts))throw new Error('Unsupported persisted room');state.document=await this.validate(state.document);this.rooms.set(name,state);}}return this.rooms.get(name);}
  async _save(name,state){try{await this.save(name,copy(state));}catch(error){
    // A failure after atomic rename can leave the new revision on disk. Reconcile before retrying.
    try{const persisted=await this.load(name);if(persisted){this.rooms.delete(name);await this._get(name);}else this.faulted.add(name);}catch{this.faulted.add(name);}
    throw error;
  }}
  _view(state){return {revision:state.revision,epoch:state.epoch,document:copy(state.document)};}
  get(name){return this._serial(name,async()=>{const s=await this._get(name);if(!s)throw new CollaborationError('NOT_FOUND','Room does not exist');return this._view(s);});}
  create(name,document){document=copy(document);return this._serial(name,async()=>{if(await this._get(name))throw new CollaborationError('EXISTS','Room already exists; join explicitly');const state={format:'GridWebRoom',version:1,revision:0,epoch:0,document:await this.validate(document),receipts:[]};await this._save(name,state);this.rooms.set(name,state);return this._view(state);});}
  commit(name,request,actor='editor'){request=copy(request);return this._serial(name,async()=>{
    const state=await this._get(name);if(!state)throw new CollaborationError('NOT_FOUND','Room does not exist');
    if(!request||!/^[-\w]{1,100}$/.test(request.id??'')||!/^[-\w]{1,100}$/.test(request.clientId??'')||!Number.isSafeInteger(request.baseRevision)||request.baseRevision<0||!Number.isSafeInteger(request.epoch)||request.epoch<0)throw new CollaborationError('INVALID','Invalid commit envelope');
    const identity=actor+':'+request.clientId+':'+request.id, payload=Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256',new TextEncoder().encode(canonical(request))))).map(b=>b.toString(16).padStart(2,'0')).join(''), receipt=state.receipts.find(r=>r.identity===identity);
    if(receipt){if(receipt.payload!==payload)throw new CollaborationError('INVALID','Commit identity reused for different edits');return {...this._view(state),appliedRevision:receipt.revision,duplicate:true};}
    if(request.baseRevision>state.revision)throw new CollaborationError('INVALID','Client revision is ahead of the room');
    if(request.epoch!==state.epoch)throw new CollaborationError('CONFLICT','Worksheet structure changed; refresh before applying edits',[{field:'structure',epoch:state.epoch}]);
    let document,epoch=state.epoch;
    if(request.kind==='replace'){
      if(request.baseRevision!==state.revision)throw new CollaborationError('CONFLICT','Structural edits require the current room revision',[{field:'structure'}]);
      document=await this.validate(request.document);epoch++;
    }else if(request.kind==='cells')document=(await this.apply(state.document,request.changes)).document;
    else throw new CollaborationError('INVALID','Unknown commit kind');
    const next={...state,document,epoch,revision:state.revision+1,receipts:[...state.receipts,{identity,payload,revision:state.revision+1}].slice(-128)};
    // Hashes compare immutable requests without retaining copies of large workbook payloads.
    let size=0;next.receipts=next.receipts.toReversed().filter(r=>(size+=r.payload.length)<=4*1024*1024).reverse();
    await this._save(name,next);this.rooms.set(name,next);
    try{this.notify(name,{type:'revision',revision:next.revision,epoch,actor});}catch{/* A disconnected observer cannot turn a durable commit into an error. */}
    return {...this._view(next),appliedRevision:next.revision,duplicate:false};
  });}
}
