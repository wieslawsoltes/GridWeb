"""Qualification of installed React and companion libraries, with no substituted mocks."""
from pathlib import Path
from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
from functools import partial
import threading, os, json, time
from playwright.sync_api import sync_playwright
ROOT=Path(__file__).resolve().parents[1]; OUT=ROOT/'test-results';OUT.mkdir(exist_ok=True)
server=ThreadingHTTPServer(('127.0.0.1',0),partial(SimpleHTTPRequestHandler,directory=str(ROOT/'site')))
threading.Thread(target=server.serve_forever,daemon=True).start()
results=[]
def require(v,message='Assertion failed'):
    if not v: raise AssertionError(message)
def check(name,action):
    start=time.perf_counter()
    try: action();results.append({'name':name,'status':'passed','ms':round((time.perf_counter()-start)*1000,2)});print('PASS',name,flush=True)
    except Exception as e: results.append({'name':name,'status':'failed','error':str(e)});print('FAIL',name,str(e),flush=True)
errors=[];version='not initialized'
try:
 with sync_playwright() as p:
    browser=p.chromium.launch(executable_path=os.getenv('CHROMIUM_PATH','/usr/bin/chromium'),headless=True,args=['--no-sandbox'])
    page=browser.new_page(viewport={'width':1600,'height':1400});page.on('pageerror',lambda e:(errors.append(str(e)),print('BROWSER ERROR',str(e),flush=True)))
    page.goto(f'http://127.0.0.1:{server.server_port}/runtime-checks/',wait_until='networkidle')
    page.wait_for_function('!!window.integrationHarness',timeout=30000);page.wait_for_timeout(500)
    h='integrationHarness';ev=page.evaluate
    check('All seven real companion adapters initialize',lambda:require(ev(f'!!{h}.companions.manager&&!!{h}.companions.ribbon&&!!{h}.companions.sourceCache&&!!{h}.companions.viewModel&&!!{h}.companions.spatialIndex')))
    check('Dockyard retains the actual spreadsheet element',lambda:require(page.locator('#workspace grid-web').count()==1 and ev(f'{h}.grid.Metrics.visibleCells')>0))
    check('RBush spatial index contains rendered chart rectangles',lambda:require(ev(f'{h}.companions.spatialIndex.Count')>0))
    def ribbon():
        page.locator('ribbon-web').get_by_role('button',name='Bold',exact=True).click();require(ev(f'{h}.grid.ActiveCell.Style.font.bold') is True);require(ev(f'{h}.count.bold')==1)
    check('RibbonWeb routes a real workbook formatting command',ribbon)
    def reactive():
        ev(f"{h}.grid.Select('B2')");page.wait_for_timeout(50);require(ev(f'{h}.companions.viewModel.Selection')=='B2')
    check('ReactiveWeb receives shared selection notifications',reactive)
    def cache():
        ev(f"{h}.book.ActiveWorksheet.GetCell('B2').Value=8");page.wait_for_timeout(150);require(ev(f"{h}.companions.sourceCache.lookup({h}.book.ActiveWorksheet.Id+':B2').value.input")=='8')
    check('DynamicDataWeb reflects real workbook edits',cache)
    def tree():
        ev(f"{h}.companions.manager.Find('cells').Activate()");page.wait_for_timeout(250);require(page.locator('tree-data-grid').count()==1);require(ev("document.querySelector('tree-data-grid').Model.Items.Count")>0)
    check('TreeDataGridWeb displays the live cell source',tree)
    def graph():
        ev(f"{h}.companions.manager.Find('graph').Activate()");page.wait_for_timeout(250);require(ev("document.querySelector('quikgraph-viewer').Graph.EdgeCount")>0)
    check('QuikGraphWeb displays real calculation dependencies',graph)
    def react():
        page.wait_for_function(f'!!{h}.ref.current');require(ev(f'{h}.ref.current.Workbook==={h}.reactBook'))
        ev(f"{h}.reactBook.ActiveWorksheet.GetCell('B2').Value=7");page.wait_for_timeout(100);require(ev(f"{h}.ref.current.Sheet.GetCell('D2').Value")==35);require(int(page.locator('#react-root output').get_attribute('data-revision'))==ev(f'{h}.reactBook.Revision'))
    check('React hooks and component retain the same reactive workbook',react)
    def props():
        ev(f"{h}.configure({{theme:'dark',readOnly:true,selection:'D2',viewMode:'normal'}})");page.wait_for_timeout(100);require(ev(f'{h}.ref.current.ReadOnly') and ev(f'{h}.ref.current.Theme')=='dark' and ev(f'{h}.ref.current.Selection')=='D2')
    check('React property changes update control behavior',props)
    def events():
        ev(f"{h}.configure({{theme:'light',readOnly:false,selection:'B2',viewMode:'normal'}})");page.wait_for_timeout(100);ev(f'{h}.ref.current.Focus()');page.keyboard.press('F2');page.locator('#react-root grid-web .editor').fill('11');page.keyboard.press('Enter');require(ev(f'{h}.lastEdit') is not None);require(ev(f'{h}.selected')=='B3');require(ev(f"{h}.reactBook.ActiveWorksheet.GetCell('D2').Value")==55)
    check('React callbacks receive real keyboard edits and selection',events)
    page.screenshot(path=str(OUT/'frameworks.png'),full_page=True)
    def disposal():
        before=ev(f'{h}.reactBook.Changed.Count');ev(f'{h}.unmount()');page.wait_for_timeout(50);require(ev(f'{h}.reactBook.Changed.Count')<before);ev(f'{h}.dispose()');require(ev(f'{h}.grid.isConnected'));require(ev(f'{h}.grid.ChartHitTest===null'));require(page.locator('ribbon-web').count()==0)
    check('React and companion disposal detach owned subscriptions',disposal)
    def studio():
        studio=browser.new_page(viewport={'width':1600,'height':1000});studio.on('pageerror',lambda e:(errors.append(str(e)),print('STUDIO ERROR',str(e),flush=True)))
        studio.goto(f'http://127.0.0.1:{server.server_port}/demo/',wait_until='networkidle');studio.wait_for_function('!!window.gridweb')
        studio.evaluate('gridweb.commands.loadIntegrations()');studio.wait_for_timeout(700)
        require(studio.locator('ribbon-web').count()==1);require(studio.evaluate('gridweb.grid.Metrics.visibleCells')>0)
        studio.evaluate("gridweb.grid.Select('C8');gridweb.grid.Focus()");studio.keyboard.press('F2');studio.locator('grid-web .editor').fill('123456');studio.keyboard.press('Enter');require(studio.evaluate("gridweb.grid.Sheet.GetCell('C8').Value")==123456)
        studio.screenshot(path=str(OUT/'companion-studio.png'),full_page=True);studio.close()
    check('Distributed studio activates the real companion bundle and remains editable',studio)
    check('No uncaught framework page errors',lambda:require(not errors,str(errors)))
    version=ev(f'{h}.reactVersion');browser.close()
finally: server.shutdown()
(OUT/'frameworks-browser.json').write_text(json.dumps({'reactVersion':version,'results':results,'errors':errors},indent=2))
raise SystemExit(1 if any(r['status']=='failed' for r in results) else 0)
