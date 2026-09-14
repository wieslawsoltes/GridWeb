import http from 'node:http';
import fs from 'node:fs/promises';
import path from 'node:path';
import { timingSafeEqual, randomUUID } from 'node:crypto';
import { CollaborationStore, roomName } from './store.js';
import { CollaborationError } from './document.js';
import {createValidationPool} from './validation.js';
import {parseRange} from '../address.js';

/** Single-process file persistence. A shared directory must not have multiple server writers. */
export async function createFileStorage(directory) {
  const root=path.resolve(directory);await fs.mkdir(root,{recursive:true,mode:0o700});
  return {
    async load(room){try{return JSON.parse(await fs.readFile(path.join(root,roomName(room)+'.json'),'utf8'));}catch(e){if(e.code==='ENOENT')return null;throw e;}},
    async save(room,state){
      const target=path.join(root,roomName(room)+'.json'),temp=target+'.'+randomUUID()+'.tmp';
      try {
        const file=await fs.open(temp,'wx',0o600);
        try{await file.writeFile(JSON.stringify(state));await file.sync();}finally{await file.close();}
        await fs.rename(temp,target);
        if(process.platform!=='win32'){const dir=await fs.open(root,'r');try{await dir.sync();}finally{await dir.close();}}
      }finally{await fs.rm(temp,{force:true}).catch(()=>{});}
    }
  };
}
export function bearerAuthorization(token,{id='editor',role='editor'}={}) {
  if(typeof token!=='string'||token.length<24)throw new TypeError('A bearer token of at least 24 characters is required');
  const expected=Buffer.from('Bearer '+token);
  return req=>{const supplied=Buffer.from(req.headers.authorization??'');return supplied.length===expected.length&&timingSafeEqual(supplied,expected)?{id,role}:null;};
}
const status={INVALID:400,LIMIT:413,NOT_FOUND:404,EXISTS:409,CONFLICT:409,UNAUTHORIZED:401,FORBIDDEN:403,COMPUTE_LIMIT:422,STORAGE:503};
/** Authentication is mandatory and application-owned; no permissive default is provided. */
export async function createCollaborationServer({authorize,storageDirectory,storage,allowedOrigins=[],studioFile=null,maxBodyBytes=20*1024*1024,maxStreams=100,maxRooms=32,validationTimeoutMs=5000}={}) {
  if(typeof authorize!=='function')throw new TypeError('An authorization callback is required');
  if(!Number.isSafeInteger(maxBodyBytes)||maxBodyBytes<1024||maxBodyBytes>32*1024*1024)throw new RangeError('Invalid body limit');
  if(!Number.isInteger(maxStreams)||maxStreams<1||maxStreams>1000)throw new RangeError('Invalid stream limit');
  const persistence=storage??(storageDirectory?await createFileStorage(storageDirectory):null);
  if(!persistence)throw new TypeError('Durable storageDirectory or storage adapter is required');
  const subscribers=new Map(),presence=new Map();let streams=0;
  const send=(res,data)=>{if(res.destroyed)return; if(res.writableLength>1024*1024){res.destroy();return;}res.write('data: '+JSON.stringify(data)+'\n\n');};
  const broadcast=(room,data)=>{for(const res of subscribers.get(room)??[])send(res,data);};
  if(!Number.isInteger(maxRooms)||maxRooms<1||maxRooms>1000)throw new RangeError('Invalid room limit');
  const validators=createValidationPool({timeoutMs:validationTimeoutMs});
  const store=new CollaborationStore({...persistence,notify:broadcast,maxRooms,validate:validators.validate,apply:validators.apply});
  const json=(res,code,value)=>{res.writeHead(code,{'Content-Type':'application/json; charset=utf-8','Cache-Control':'no-store','X-Content-Type-Options':'nosniff'});res.end(JSON.stringify(value));};
  const body=async req=>{
    if(!String(req.headers['content-type']??'').startsWith('application/json'))throw new CollaborationError('INVALID','Use application/json');
    let bytes=0;const chunks=[];
    for await(const chunk of req){bytes+=chunk.length;if(bytes>maxBodyBytes)throw new CollaborationError('LIMIT','Request body exceeds limit');chunks.push(chunk);}
    try{return JSON.parse(Buffer.concat(chunks).toString('utf8'));}catch{throw new CollaborationError('INVALID','Invalid JSON');}
  };
  const listPresence=room=>{const map=presence.get(room);if(!map)return [];for(const [key,value]of map)if(Date.now()-value.updatedAt>=45000)map.delete(key);return [...map.values()].map(({updatedAt,...x})=>x);};
  const server=http.createServer(async(req,res)=>{
    try {
      const origin=req.headers.origin;
      if(origin&&!allowedOrigins.includes(origin))throw new CollaborationError('FORBIDDEN','Origin not allowed');
      if(origin){res.setHeader('Access-Control-Allow-Origin',origin);res.setHeader('Vary','Origin');res.setHeader('Access-Control-Allow-Headers','Authorization, Content-Type');res.setHeader('Access-Control-Allow-Methods','GET, PUT, POST, OPTIONS');}
      if(req.method==='OPTIONS'){res.writeHead(204);res.end();return;}
      const url=new URL(req.url,'http://localhost');
      if(req.method==='GET'&&url.pathname==='/studio'&&studioFile){res.writeHead(200,{'Content-Type':'text/html; charset=utf-8','Cache-Control':'no-store','X-Content-Type-Options':'nosniff'});res.end(await fs.readFile(studioFile));return;}
      const match=/^\/rooms\/([-\w]{1,80})(?:\/(commits|events|presence))?$/.exec(url.pathname);
      if(!match)throw new CollaborationError('NOT_FOUND','Unknown endpoint');
      const [,room,action]=match;roomName(room);
      const user=await authorize(req,{room,action:action??(req.method==='PUT'?'create':'read')});
      if(!user||typeof user.id!=='string'||!user.id||user.id.length>100||!['editor','viewer'].includes(user.role))throw new CollaborationError('UNAUTHORIZED','Authentication required');
      const writable=()=>{if(user.role!=='editor')throw new CollaborationError('FORBIDDEN','Editor role required');};
      if(!action&&req.method==='GET'){json(res,200,await store.get(room));return;}
      if(!action&&req.method==='PUT'){writable();json(res,201,await store.create(room,(await body(req)).document));return;}
      if(action==='commits'&&req.method==='POST'){writable();json(res,200,await store.commit(room,await body(req),user.id));return;}
      if(action==='presence'&&req.method==='POST'){
        const current=await store.get(room),value=await body(req);
        if(!/^[-\w]{1,100}$/.test(value.clientId??'')||typeof value.sheet!=='string'||value.sheet.length>80||typeof value.selection!=='string'||value.selection.length>80)throw new CollaborationError('INVALID','Invalid presence');
        if(!current.document.sheets.some(s=>s.id===value.sheet))throw new CollaborationError('INVALID','Presence worksheet does not exist');
        try{parseRange(value.selection);}catch{throw new CollaborationError('INVALID','Invalid presence range');}
        listPresence(room);const map=presence.get(room)??new Map();if(map.size>=200&&!map.has(user.id+':'+value.clientId))throw new CollaborationError('LIMIT','Too many participants');
        map.set(user.id+':'+value.clientId,{clientId:value.clientId,actor:user.id,sheet:value.sheet,selection:value.selection,updatedAt:Date.now()});presence.set(room,map);
        broadcast(room,{type:'presence',participants:listPresence(room)});json(res,200,{ok:true});return;
      }
      if(action==='events'&&req.method==='GET'){
        if(streams>=maxStreams)throw new CollaborationError('LIMIT','Too many event streams');
        const snapshot=await store.get(room);res.writeHead(200,{'Content-Type':'text/event-stream','Cache-Control':'no-store','Connection':'keep-alive','X-Accel-Buffering':'no'});res.flushHeaders();
        const set=subscribers.get(room)??new Set();set.add(res);subscribers.set(room,set);streams++;
        send(res,{type:'revision',revision:snapshot.revision,epoch:snapshot.epoch});send(res,{type:'presence',participants:listPresence(room)});
        let checking=false;
        const heartbeat=setInterval(async()=>{if(checking)return;checking=true;try{const current=await authorize(req,{room,action:'events'});if(!current||current.id!==user.id||!['editor','viewer'].includes(current.role)){res.end();return;}res.write(': keepalive\n\n');}catch{res.end();}finally{checking=false;}},15000);heartbeat.unref();
        res.on('close',()=>{clearInterval(heartbeat);if(set.delete(res))streams--;if(!set.size)subscribers.delete(room);});return;
      }
      throw new CollaborationError('NOT_FOUND','Unknown endpoint');
    }catch(e){if(res.headersSent){res.end();return;}json(res,status[e.code]??500,{error:{code:e.code??'SERVER_ERROR',message:status[e.code]?e.message:'Server operation failed',details:status[e.code]?e.details:null}});}
  });
  server.requestTimeout=30000;server.headersTimeout=10000;
  let closing;
  return {server,store,close:()=>closing??=(async()=>{for(const set of subscribers.values())for(const res of set)res.end();try{if(server.listening)await new Promise((resolve,reject)=>server.close(e=>e?reject(e):resolve()));}finally{await validators.close();}})()};
}
