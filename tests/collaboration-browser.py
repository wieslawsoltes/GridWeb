"""Two independent Chromium contexts against the real authenticated, durable Node server."""
from pathlib import Path
import json, os, socket, subprocess, tempfile, time, urllib.request
from playwright.sync_api import sync_playwright
ROOT=Path(__file__).resolve().parents[1];OUT=ROOT/'test-results';OUT.mkdir(exist_ok=True)
TOKEN='gridweb-browser-test-token-32-characters'
with socket.socket() as s:s.bind(('127.0.0.1',0));port=s.getsockname()[1]
url=f'http://127.0.0.1:{port}';results=[];errors=[]
def require(v,m='Assertion failed'):
 if not v:raise AssertionError(m)
def check(name,action):
 try:action();results.append({'name':name,'status':'passed'});print('PASS',name,flush=True)
 except Exception as e:results.append({'name':name,'status':'failed','error':str(e)});print('FAIL',name,str(e),flush=True)
with tempfile.TemporaryDirectory(prefix='gridweb-http-') as directory:
 env={**os.environ,'PORT':str(port),'GRIDWEB_COLLAB_TOKEN':TOKEN,'GRIDWEB_DATA':directory}
 log=open(OUT/'collaboration-server.log','w')
 proc=subprocess.Popen(['node','scripts/collaboration-server.mjs'],cwd=ROOT,env=env,stdout=log,stderr=log)
 try:
  for i in range(100):
   try:urllib.request.urlopen(url+'/studio',timeout=1).close();break
   except Exception:time.sleep(.1)
  with sync_playwright() as p:
   browser=p.chromium.launch(executable_path=os.getenv('CHROMIUM_PATH','/usr/bin/chromium'),headless=True,args=['--no-sandbox'])
   ca,cb=browser.new_context(),browser.new_context();a,b=ca.new_page(),cb.new_page()
   for page in [a,b]:
    page.set_viewport_size({'width':1600,'height':1000});page.on('pageerror',lambda e:errors.append(str(e)));page.goto(url+'/studio');page.wait_for_function('!!window.gridwebCollaboration')
   def connect(page,mode):
    page.locator('#collaboration-button').click();page.get_by_label('Server URL',exact=True).fill(url);page.get_by_label('Room',exact=True).fill('browser-room');page.get_by_label('Access token',exact=True).fill(TOKEN);page.get_by_label('Connection action').select_option(mode)
    if mode=='join':page.get_by_label('Confirm replacing local workbook').check()
    page.get_by_role('button',name='Connect',exact=True).click();page.wait_for_function('gridwebCollaboration.session?.Status==="connected"');page.get_by_role('button',name='Close collaboration').click()
   check('Create and join authenticated room using Share controls',lambda:(connect(a,'create'),connect(b,'join'),require(a.evaluate('gridweb.workbook.ActiveWorksheet.Id')==b.evaluate('gridweb.workbook.ActiveWorksheet.Id'))))
   def keyboard():
    a.evaluate("gridweb.grid.Select('M1');gridweb.grid.Focus()");a.keyboard.press('F2');a.locator('grid-web .editor').fill('124');a.keyboard.press('Enter');b.wait_for_function("gridweb.grid.Sheet.GetCell('M1').Value===124")
    b.evaluate("gridweb.grid.Sheet.GetCell('N1').Formula='=M1*2'");a.wait_for_function("gridweb.grid.Sheet.GetCell('N1').Value===248")
   check('Keyboard edits and recalculated formulas synchronize across browsers',keyboard)
   def conflict():
    for page in [a,b]:page.evaluate('gridwebCollaboration.session.AutoSync=false')
    a.evaluate("gridweb.grid.Sheet.GetCell('M1').Value=200");b.evaluate("gridweb.grid.Sheet.GetCell('M1').Value=300");b.evaluate("gridweb.grid.Sheet.GetCell('O1').Value='retained'")
    a.evaluate('gridwebCollaboration.session.Sync()');require(b.evaluate('gridwebCollaboration.session.Sync()') is False)
    b.locator('#collaboration-button').click();b.get_by_role('button',name='Use server conflicting edits',exact=True).click();b.wait_for_function('!gridwebCollaboration.session.Conflict');b.get_by_role('button',name='Close collaboration').click();a.evaluate('gridwebCollaboration.session.Sync()');require(a.evaluate("gridweb.grid.Sheet.GetRange('M1:O1').Values")==[[200,400,'retained']])
   check('Conflict dialog resolves one field without dropping disjoint edits',conflict)
   def reload_pending():
    b.evaluate("gridweb.grid.Sheet.GetCell('P1').Value='reload-safe'");require(b.evaluate('gridwebCollaboration.session.HasPendingChanges'))
    b.reload();b.wait_for_function('!!gridwebCollaboration');connect(b,'join');a.evaluate('gridwebCollaboration.session.Sync()');require(a.evaluate("gridweb.grid.Sheet.GetCell('P1').Value")=='reload-safe')
   check('Unsent work survives page reload and reconnects to the same room',reload_pending)
   def presence():
    a.evaluate("gridweb.grid.Select('Q2')");b.wait_for_function("gridwebCollaboration.participants.some(p=>p.selection==='Q2')");b.locator('#collaboration-button').click();require('Q2' in b.get_by_label('Room participants').inner_text());b.screenshot(path=str(OUT/'coauthoring-studio.png'),full_page=True);b.get_by_role('button',name='Close collaboration').click()
   check('Room presence reports participant selection',presence)
   def no_token():
    require(b.evaluate('(token)=>![...Object.values(sessionStorage),...Object.values(localStorage)].some(s=>s.includes(token))',TOKEN))
   check('Access token is absent from session and workbook storage',no_token)
   def disconnect():
    b.evaluate('gridwebCollaboration.disconnect()');require(b.locator('#collaboration-button').inner_text()=='Share');b.evaluate("gridweb.grid.Sheet.GetCell('R1').Value='local-only'");a.evaluate('gridwebCollaboration.session.Sync()');require(a.evaluate("gridweb.grid.Sheet.GetCell('R1').Value") is None)
   check('Disconnect returns to independent local editing',disconnect)
   check('No uncaught collaboration page errors',lambda:require(not errors,str(errors)))
   browser.close()
 finally:
  proc.terminate()
  try:proc.wait(timeout=8)
  except subprocess.TimeoutExpired:proc.kill();proc.wait()
  log.close()
(OUT/'collaboration-browser.json').write_text(json.dumps({'results':results,'errors':errors},indent=2))
raise SystemExit(1 if len(results)!=8 or any(r['status']=='failed' for r in results) else 0)
