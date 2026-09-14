import {CollaborationSession} from '../src/collaboration.js';
const el=(tag,text)=>{const e=document.createElement(tag);if(text!=null)e.textContent=text;return e;};
let session=null,participants=[],subscriptions=[],presenceTimer=null,dialog=null,errorMessage='';
const share=el('button','Share');share.id='collaboration-button';share.type='button';share.title='Connect to an authenticated GridWeb collaboration server';
document.querySelector('.title-actions').prepend(share);
function status(){share.textContent=session?'Share · '+session.Status:'Share';if(dialog?.open)renderState();}
function disconnect(){clearInterval(presenceTimer);for(const s of subscriptions)s.Dispose();subscriptions=[];session?.Dispose();session=null;participants=[];status();}
function presence(){if(session&&window.gridweb.workbook===session.Workbook)session.SetPresence(window.gridweb.grid.Sheet.Id,window.gridweb.grid.Selection).catch(e=>{errorMessage=e.message;status();});}
document.querySelector('#grid').addEventListener('grid-render',()=>{if(session&&session.Workbook!==window.gridweb.workbook)disconnect();});
document.querySelector('#grid').addEventListener('selection-change',()=>{clearTimeout(presence._timer);presence._timer=setTimeout(presence,250);});
document.querySelector('#book-title').addEventListener('change',()=>session?.Sync().catch(e=>{errorMessage=e.message;status();}));
function renderState(){
 const box=dialog.querySelector('[data-session-state]');box.replaceChildren();
 box.append(el('p',session?`Room ${session.Room} · revision ${session.Revision} · ${session.Status}${session.HasPendingChanges?' · edits pending':''}`:'No room connected. Your workbook stays local until you connect.'));
 if(errorMessage){const alert=el('p',errorMessage);alert.setAttribute('role','alert');box.append(alert);}
 if(!session)return;
 for(const [label,action]of [['Sync now',()=>session.Sync()],['Disconnect',async()=>disconnect()]]){const b=el('button',label);b.type='button';b.onclick=()=>Promise.resolve(action()).catch(e=>{errorMessage=e.message;status();});box.append(b);}
 if(session.Conflict){
  box.append(el('p','Concurrent edits conflict. Export a JSON backup before choosing. Cell conflicts keep disjoint edits; structural conflicts choose an entire workbook.'));
  for(const [label,strategy]of [['Keep my conflicting edits','local'],['Use server conflicting edits','remote']]){const b=el('button',label);b.type='button';b.onclick=()=>session.Resolve(strategy).catch(e=>{errorMessage=e.message;status();});box.append(b);}
 }
 const list=el('div');list.setAttribute('aria-label','Room participants');
 for(const p of participants)list.append(el('p',`${p.actor} · ${session.Workbook._sheets.find(s=>s.Id===p.sheet)?.Name??p.sheet}!${p.selection}`));box.append(list);
}
function open(){
 if(dialog?.open)return;
 dialog=el('dialog');dialog.id='collaboration-dialog';const form=el('form'),header=el('header');header.append(el('h2','Share workbook'));const close=el('button','×');close.type='button';close.setAttribute('aria-label','Close collaboration');close.onclick=()=>dialog.close();header.append(close);form.append(header);
 const intro=el('p','Requires your own GridWeb server. The token is used only in request headers, not URLs or saved workbooks. Pending edits survive reload in this tab; reconnect with the token to resume.');intro.className='hint';form.append(intro);
 const fields={};
 for(const [key,label,value,type]of [['url','Server URL',location.protocol.startsWith('http')?location.origin:'http://127.0.0.1:8099','url'],['room','Room','team-workbook','text'],['token','Access token','','password']]){
  const l=el('label',label),input=el('input');input.name=key;input.type=type;input.required=true;input.autocomplete=type==='password'?'off':'on';input.value=value;input.setAttribute('aria-label',label);input.style.cssText='display:block;width:100%;padding:8px;margin:4px 0 12px';l.append(input);form.append(l);fields[key]=input;
 }
 const label=el('label','Connection action'),mode=el('select');mode.setAttribute('aria-label','Connection action');mode.style.cssText='display:block;padding:8px;width:100%;margin:4px 0 12px';
 for(const [value,text]of [['create','Create room from this workbook'],['join','Join room and replace this local view']]){const option=el('option',text);option.value=value;mode.append(option);}label.append(mode);form.append(label);
 const confirm=el('label'),check=el('input');check.type='checkbox';confirm.append(check,document.createTextNode(' I have exported local work that I need before joining.'));check.setAttribute('aria-label','Confirm replacing local workbook');form.append(confirm);
 const state=el('section');state.dataset.sessionState='';form.append(state);
 const footer=el('footer'),backup=el('button','Export JSON backup'),connect=el('button','Connect');backup.type='button';backup.onclick=()=>window.gridweb.commands.save();connect.type='submit';connect.className='primary';footer.append(backup,connect);form.append(footer);dialog.append(form);document.body.append(dialog);
 form.onsubmit=async e=>{e.preventDefault();errorMessage='';
  if(session){errorMessage='Disconnect the current room before connecting to another.';renderState();return;}
  if(mode.value==='join'&&!check.checked){errorMessage='Confirm replacing the local view, or export it first.';renderState();return;}
  connect.disabled=true;let candidate;
  try{
   candidate=new CollaborationSession(window.gridweb.workbook,{url:fields.url.value,room:fields.room.value,token:fields.token.value,storage:sessionStorage});
   await candidate.Connect({mode:mode.value});session=candidate;fields.token.value='';
   subscriptions=[session.Changed.Subscribe(status),session.Conflicted.Subscribe(status),session.PresenceChanged.Subscribe(value=>{participants=value;status();}),session.Error.Subscribe(e=>{errorMessage=e.message;status();})];
   presence();presenceTimer=setInterval(presence,15000);status();
  }catch(e){candidate?.Dispose();errorMessage=e.message;status();}finally{connect.disabled=false;}
 };
 dialog.addEventListener('close',()=>{fields.token.value='';dialog.remove();dialog=null;},{once:true});renderState();dialog.showModal();
}
share.onclick=open;
window.gridwebCollaboration={get session(){return session;},get participants(){return participants;},open,disconnect};
