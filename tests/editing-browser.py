"""Real Chromium integration tests for reusable editing dialogs and calculation examples."""
from pathlib import Path
import json, os, re, time
from playwright.sync_api import sync_playwright
ROOT=Path(__file__).resolve().parents[1]
OUT=ROOT/'test-results';OUT.mkdir(exist_ok=True)
results=[]
def require(condition,message='Assertion failed'):
    if not condition: raise AssertionError(message)
def check(name,action):
    start=time.perf_counter()
    try:
        action();results.append({'name':name,'status':'passed','durationMs':round((time.perf_counter()-start)*1000,2)});print('PASS',name,flush=True)
    except Exception as e:
        results.append({'name':name,'status':'failed','error':str(e)});print('FAIL',name,str(e),flush=True)
        try: page.evaluate('for(const d of gridweb.grid.shadowRoot.querySelectorAll("dialog[open]"))d.close();gridweb.grid.ReadOnly=false')
        except Exception: pass
with sync_playwright() as p:
    browser=p.chromium.launch(executable_path=os.getenv('CHROMIUM_PATH','/usr/bin/chromium'),headless=True,args=['--no-sandbox'])
    page=browser.new_page(viewport={'width':1600,'height':1000})
    page.set_default_timeout(5000)
    errors=[];page.on('pageerror',lambda e:errors.append(str(e)))
    if os.getenv('GRIDWEB_BASE_URL'):page.goto(os.environ['GRIDWEB_BASE_URL'],wait_until='networkidle')
    else:page.set_content((ROOT/'dist/GridWeb-standalone.html').read_text(),wait_until='load')
    page.wait_for_function('!!window.gridweb?.commands.calculationTools')
    ev=page.evaluate
    def select(address):ev('(a)=>{gridweb.grid.Select(a);gridweb.grid.Focus()}',address)
    def value(address):return ev('(a)=>gridweb.grid.Sheet.GetCell(a).Value',address)
    def close_dialogs():ev('for(const d of gridweb.grid.shadowRoot.querySelectorAll("dialog[open]"))d.close()')
    def examples():
        page.get_by_role('tab',name='Home',exact=True).click()
        page.get_by_role('button',name='Calculation tools',exact=True).click()
        page.wait_for_function('gridweb.grid.Sheet.Name==="Calculation tools"')
        require(value('G4')==440);require(value('B11')==14);require(value('B12')==144);require(value('B13')==4);require(value('I9')==16)
        require(ev("gridweb.grid.Sheet.GetCell('D13').Text")=='2026-02-28')
        require(page.get_by_role('tab',name='Calculation tools',exact=True).get_attribute('aria-selected')=='true')
        page.screenshot(path=str(OUT/'calculation-editing-tools.png'),full_page=True)
    check('Calculation tools sample runs database, aggregate, lambda and series features',examples)
    def paste_values():
        ev("gridweb.grid.Sheet.GetRange('A25:B26').Values=[[2,3],[4,5]]")
        select('A25:B26');page.keyboard.press('Control+c')
        require(ev('!!gridweb.grid._clipboard?.snapshot'))
        ev("gridweb.grid.Sheet.GetCell('A25').Value=99")
        select('D25');page.get_by_role('button',name='Paste special',exact=True).click()
        page.get_by_label('Paste content',exact=True).select_option('values')
        page.get_by_label('Transpose rows and columns',exact=True).check()
        page.screenshot(path=str(OUT/'paste-special-dialog.png'),full_page=True)
        page.locator('grid-web dialog[open]').get_by_role('button',name='Paste',exact=True).click()
        require(ev("gridweb.grid.Sheet.GetRange('D25:E26').Values")==[[2,4],[3,5]])
        require(ev('gridweb.grid.Selection')=='D25:E26')
        page.keyboard.press('Control+z');require(value('D25') is None)
        page.keyboard.press('Control+y');require(value('D25')==2)
    check('Keyboard copy and dialog transpose paste retain snapshot values and one-step undo',paste_values)
    def arithmetic():
        select('H25:I26');ev("gridweb.grid.Sheet.GetRange('H25:I26').Values=[[10,10],[10,10]]")
        page.keyboard.press('Control+Alt+v')
        page.get_by_label('Operation',exact=True).select_option('multiply')
        require(page.get_by_label('Paste content',exact=True).input_value()=='values')
        page.locator('grid-web dialog[open]').get_by_role('button',name='Paste',exact=True).click()
        require(ev("gridweb.grid.Sheet.GetRange('H25:I26').Values")==[[20,30],[40,50]])
    check('Ctrl+Alt+V opens paste special and arithmetic uses copied values',arithmetic)
    def fill_series():
        select('A29:A32');page.get_by_role('button',name='Fill series',exact=True).click()
        page.get_by_label('Start value (optional)',exact=True).fill('5')
        page.get_by_label('Step',exact=True).fill('3')
        require(value('A29') is None,'Typing in dialog mutated the grid')
        page.locator('grid-web dialog[open]').get_by_role('button',name='Fill',exact=True).click()
        require(ev("gridweb.grid.Sheet.GetRange('A29:A32').Values")==[[5],[8],[11],[14]],str(ev("gridweb.grid.Sheet.GetRange('A29:A32').Values")))
        page.keyboard.press('Control+z');require(value('A29') is None,'Fill undo did not restore A29: '+str(value('A29')))
    check('Fill series dialog is keyboard-isolated and creates one undoable series',fill_series)
    def validation_rollback():
        ev("gridweb.grid.Sheet.AddValidation('J32',{type:'whole',min:0,max:5})")
        select('J30:J32');page.get_by_role('button',name='Fill series',exact=True).click()
        page.get_by_label('Start value (optional)',exact=True).fill('3');page.get_by_label('Step',exact=True).fill('3')
        page.locator('grid-web dialog[open]').get_by_role('button',name='Fill',exact=True).click()
        require(bool(page.locator('grid-web dialog[open] [role=alert]').inner_text()))
        require(value('J30') is None);require(value('J31') is None)
        page.keyboard.press('Escape');require(page.locator('grid-web dialog[open]').count()==0)
    check('Rejected series retains inline error, rolls back earlier cells and supports Escape',validation_rollback)
    def special():
        select('A11:B15');page.get_by_role('button',name='Go to special',exact=True).click()
        page.locator('grid-web dialog[open]').get_by_role('button',name='Find cells',exact=True).click()
        require('5 matching cells' in page.locator('grid-web dialog [role=status]').inner_text())
        page.get_by_role('button',name=re.compile('^B13 —')).click();require(ev('gridweb.grid.Selection')=='B13')
    check('Go to special finds formulas and navigates to actual model cells',special)
    def readonly():
        ev('gridweb.grid.ReadOnly=true');select('K30:K32')
        page.get_by_role('button',name='Fill series',exact=True).click()
        page.get_by_label('Start value (optional)',exact=True).fill('1');page.locator('grid-web dialog[open]').get_by_role('button',name='Fill',exact=True).click()
        require('read-only' in page.locator('grid-web dialog [role=alert]').inner_text());require(value('K30') is None)
        page.locator('grid-web dialog[open]').get_by_role('button',name='Cancel',exact=True).click()
        ev('gridweb.grid.ReadOnly=false')
    check('Read-only mode is enforced at apply time by the public control API',readonly)
    def fill_shortcuts():
        ev("gridweb.grid.Sheet.GetCell('L30').Formula='=ROW()'")
        select('L30:L32');page.keyboard.press('Control+d');require(value('L32')==32)
        ev("gridweb.grid.Sheet.GetCell('L34').Formula='=COLUMN()'")
        select('L34:N34');page.keyboard.press('Control+r');require(value('N34')==14)
    check('Ctrl+D and Ctrl+R fill formulas through the shared engine',fill_shortcuts)
    def context_menu():
        ev('gridweb.grid._showContext(100,100)')
        page.locator('grid-web .context').get_by_role('button',name='Go to special',exact=True).click()
        require(page.get_by_role('dialog',name='Go to special',exact=True).is_visible());page.keyboard.press('Escape')
    check('Context menu exposes the reusable editing dialogs',context_menu)
    def mobile():
        page.set_viewport_size({'width':390,'height':844});ev("gridweb.grid.Theme='dark';gridweb.grid.ShowFillSeries()")
        box=page.get_by_role('dialog',name='Fill series',exact=True).bounding_box()
        require(box['x']>=0 and box['x']+box['width']<=390 and box['height']<=844,str(box))
        page.screenshot(path=str(OUT/'editing-dialog-mobile-dark.png'),full_page=True)
        page.keyboard.press('Escape');page.set_viewport_size({'width':1600,'height':1000})
    check('Dark mobile dialog fits the viewport and uses accessible native controls',mobile)
    def lifecycle():
        ev("gridweb.grid.ShowGoToSpecial();const g=gridweb.grid,p=g.parentNode,next=g.nextSibling;g.remove();p.insertBefore(g,next)")
        page.wait_for_function('gridweb.grid.shadowRoot.querySelectorAll("dialog[open]").length===0')
        select('A1');require(ev('gridweb.grid.isConnected'))
    check('Disconnect closes editing dialogs and the control reconnects safely',lifecycle)
    check('No uncaught page errors',lambda:require(not errors,str(errors)))
    browser.close()
(OUT/'editing-browser.json').write_text(json.dumps({'results':results,'errors':errors},indent=2))
print(f'{sum(r["status"]=="passed" for r in results)}/{len(results)} editing browser groups passed')
raise SystemExit(1 if any(r['status']=='failed' for r in results) else 0)
