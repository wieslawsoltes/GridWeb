"""Offline Chromium regression suite. Uses the same complete HTML distributed to users.
Set GRIDWEB_BASE_URL to test an HTTP deployment instead; no browser policies are changed.
"""
from pathlib import Path
import json, os, time, traceback
from playwright.sync_api import sync_playwright
ROOT=Path(__file__).resolve().parents[1]
OUT=ROOT/'test-results';OUT.mkdir(exist_ok=True)
results=[]
def check(name, action):
    started=time.perf_counter()
    try:
        action()
        results.append({'name':name,'status':'passed','durationMs':round((time.perf_counter()-started)*1000,2)})
        print('PASS',name,flush=True)
    except Exception as exc:
        results.append({'name':name,'status':'failed','error':str(exc),'durationMs':round((time.perf_counter()-started)*1000,2)})
        print('FAIL',name,str(exc),flush=True)
def require(condition,message='Assertion failed'):
    if not condition: raise AssertionError(message)
with sync_playwright() as p:
    browser=p.chromium.launch(executable_path=os.getenv('CHROMIUM_PATH','/usr/bin/chromium'),headless=True,args=['--no-sandbox'])
    page=browser.new_page(viewport={'width':1600,'height':1000},device_scale_factor=1)
    errors=[];page.on('pageerror',lambda e:errors.append(str(e)))
    if os.getenv('GRIDWEB_BASE_URL'): page.goto(os.environ['GRIDWEB_BASE_URL'],wait_until='networkidle')
    else: page.set_content((ROOT/'dist/GridWeb-standalone.html').read_text(),wait_until='load')
    page.wait_for_function('!!window.gridweb');page.wait_for_timeout(350)
    ev=page.evaluate
    def select(addr): ev('(address)=>{gridweb.grid.Select(address);gridweb.grid.Focus();}',addr)
    def value(addr): return ev('(a)=>gridweb.grid.Sheet.GetCell(a).Value',addr)
    def cmd(name): ev('(name)=>gridweb.commands[name]()',name);page.wait_for_timeout(100)
    check('Standalone starts with five actual worksheets',lambda:require(ev('gridweb.workbook.Worksheets.Count')==5))
    check('Cross-sheet operating-plan total calculates',lambda:require(value('C19')==1519700))
    check('Virtualized canvas draws bounded visible cells',lambda:require(0<ev('gridweb.grid.Metrics.visibleCells')<1500))
    check('ARIA grid exposes visible cell semantics',lambda:require(page.locator('grid-web [role="gridcell"]').count()>0))
    page.screenshot(path=str(OUT/'studio-light.png'),full_page=True)
    def keyboard_edit():
        select('C8');page.keyboard.press('F2');page.locator('grid-web .editor').fill('125000');page.keyboard.press('Enter');require(value('C8')==125000);require(ev('gridweb.grid.Selection')=='C9')
    check('F2 native editor commits value and moves selection',keyboard_edit)
    def undo_redo():
        current=value('C8');page.keyboard.press('Control+z');require(value('C8')!=current);page.keyboard.press('Control+y');require(value('C8')==current)
    check('Keyboard undo and redo update shared engine',undo_redo)
    def formula_bar():
        select('L1');page.locator('#formula').fill('=SUM(1,2,3)*7');page.locator('#formula').press('Enter');require(value('L1')==42)
    check('Formula bar evaluates safe expressions',formula_bar)
    def cancelled_edit():
        select('L1');page.keyboard.press('F2');page.locator('grid-web .editor').fill('999');page.keyboard.press('Escape');require(value('L1')==42)
    check('Escape cancels edit without mutation',cancelled_edit)
    def extend_selection():
        select('B6');page.keyboard.press('Shift+ArrowRight');page.keyboard.press('Shift+ArrowDown');require(ev('gridweb.grid.Selection')=='B6:C7')
    check('Shift keyboard rectangular selection',extend_selection)
    def format_shortcut():
        select('L1');page.keyboard.press('Control+b');require(ev('gridweb.grid.ActiveCell.Style.font.bold') is True)
    check('Keyboard formatting applies engine style',format_shortcut)
    def paste_tsv():
        select('L3');ev('gridweb.grid.PasteText("3\\t4\\n5\\t=SUM(L3:M3)")');require(value('M4')==7)
    check('TSV paste is atomic and calculates formulas',paste_tsv)
    def clipboard_snapshot():
        ev('''()=>{const g=gridweb.grid;g.Select('L3:M4');g._clipboard=g._captureClipboard(g.SelectionRange,'__clipboard_test__',false);g.Sheet.GetCell('L3').Input=20;g.Select('L6');g.PasteText('__clipboard_test__');}''');require(value('L6')==3);require(value('M7')==7)
    check('Internal clipboard retains snapshot and shifts formulas',clipboard_snapshot)
    def sheet_switch():
        page.get_by_role('tab',name='Formula lab',exact=True).click();require(value('D22')==24);require(value('H22')==324)
    check('Worksheet tabs expose dynamic arrays',sheet_switch)
    def validation():
        ev("gridweb.grid.Sheet='Assumptions';gridweb.grid.Select('B5');gridweb.grid.BeginEdit()")
        # Dedicated validation fixture avoids assumptions about sample row positions.
        ev("gridweb.grid.CancelEdit();gridweb.grid.Sheet.AddValidation('D1',{type:'list',values:['Base','Growth','Lean']});gridweb.grid.Select('D1');gridweb.grid.BeginEdit()")
        page.locator('grid-web .validation').select_option('Growth');require(value('D1')=='Growth')
    check('List validation uses native select editor',validation)
    def readonly():
        ev('gridweb.grid.ReadOnly=true');require(ev('gridweb.grid.BeginEdit()') is False);ev('gridweb.grid.ReadOnly=false')
    check('Read-only mode prevents user editing',readonly)
    def frozen_scroll():
        ev("gridweb.grid.Sheet='Operating plan';gridweb.grid.Select('A1');gridweb.grid._scroller.scrollTop=400");page.wait_for_timeout(100)
        require(ev('gridweb.grid.Sheet.FrozenRows')==5);require(ev('gridweb.grid._rect(0,0).y')==28)
    check('Frozen header geometry survives scrolling',frozen_scroll)
    def resize_column():
        ev('gridweb.grid._scroller.scrollLeft=0;gridweb.grid._scroller.scrollTop=0');page.wait_for_timeout(100)
        r=ev('gridweb.grid._rect(0,0)');box=page.locator('grid-web canvas').bounding_box();x=box['x']+r['x']+r['width']-2;y=box['y']+12
        page.mouse.move(x,y);page.mouse.down();page.mouse.move(x+36,y,steps=5);page.mouse.up();page.wait_for_timeout(100)
        require(ev('gridweb.grid._columns.Size(0)')>r['width']+25)
    check('Pointer column resize updates layout model',resize_column)
    def chart_resize():
        ev('gridweb.grid._scroller.scrollLeft=0;gridweb.grid._scroller.scrollTop=0');page.wait_for_timeout(100)
        before=ev('gridweb.grid.Sheet.Charts[0].width');rect=ev('gridweb.grid._chartRects[0]');require(bool(rect),'Chart was not rendered')
        # Exercise pointer handlers through real browser events, in canvas coordinates.
        box=page.locator('grid-web canvas').bounding_box();x=box['x']+rect['x']+rect['width']-3;y=box['y']+rect['y']+rect['height']-3
        page.mouse.move(x,y);page.mouse.down();page.mouse.move(x+20,y+12,steps=4);page.mouse.up();page.wait_for_timeout(100);require(ev('gridweb.grid.Sheet.Charts[0].width')>before)
    check('Canvas chart resize persists engine chart dimensions',chart_resize)
    def ribbon_tabs():
        for name in ['Insert','Page Layout','Formulas','Data','Review','View','Home']:
            page.get_by_role('tab',name=name,exact=True).click();require(page.locator('#ribbon button').count()>3,name+' empty')
    check('All seven ribbon tabs expose actual command controls',ribbon_tabs)
    def functions_dialog():
        cmd('functions');require(page.locator('#dialog').is_visible());require(page.locator('#dialog-title').inner_text()!='');page.locator('#close-dialog').click()
    check('Function insertion dialog opens and closes',functions_dialog)
    def named_ranges():
        ev("gridweb.workbook.DefineName('BrowserTest',\"='Operating plan'!L1\");gridweb.grid.Sheet.GetCell('L9').Formula='BrowserTest+1'");require(value('L9')==43)
    check('Named-range binding uses the shared calculation engine',named_ranges)
    def page_layout():
        ev("gridweb.grid.Sheet='Assumptions';gridweb.grid.ViewMode='pageLayout'");page.wait_for_timeout(150);require(page.locator('grid-web .pages .print-page').count()>0)
        page.screenshot(path=str(OUT/'studio-pages.png'),full_page=True)
    check('Page layout renders actual bounded print pages',page_layout)
    def page_edit():
        td=page.locator('grid-web .pages td[data-row="0"][data-column="3"]').first
        # Cell D1 is included in UsedRange by the validation editing check.
        td.dblclick();td.fill('Base');td.press('Enter');require(value('D1')=='Base')
    check('Page-layout WYSIWYG editing writes back to engine',page_edit)
    def page_break():
        ev("gridweb.grid.ViewMode='pageBreak'");page.wait_for_timeout(100);require(ev('gridweb.grid.ViewMode')=='pageBreak');ev("gridweb.grid.ViewMode='normal'")
    check('Page-break preview returns to normal editing',page_break)
    def zoom_dark():
        page.locator('#zoom-in').click();require(ev('gridweb.grid.Zoom')>1);cmd('theme');require(ev('gridweb.grid.Theme')=='dark');page.screenshot(path=str(OUT/'studio-dark.png'),full_page=True);cmd('theme');ev('gridweb.grid.Zoom=1')
    check('Zoom and dark theme preserve workbook state',zoom_dark)
    def safe_content():
        ev("gridweb.grid.Sheet.GetCell('D3').Value='<img src=x onerror=window.gridwebPwned=1>'");ev("gridweb.grid.ViewMode='pageLayout'");page.wait_for_timeout(100);require(ev('window.gridwebPwned===undefined'));require(page.locator('grid-web .pages img').count()==0);ev("gridweb.grid.ViewMode='normal'")
    check('Untrusted workbook text remains text in page views',safe_content)
    def shared_views():
        ev("""()=>{window.secondGrid=document.createElement('grid-web');secondGrid.style.cssText='position:fixed;right:20px;top:300px;width:320px;height:240px;z-index:90';secondGrid.Workbook=gridweb.workbook;document.body.append(secondGrid);gridweb.workbook.ActiveWorksheet.GetCell('F1').Value=777;}""");page.wait_for_timeout(120);require(ev('secondGrid.Sheet.GetCell("F1").Value')==777)
        ev('window.listenersBefore=gridweb.workbook.Changed.Count;secondGrid.remove()');require(ev('gridweb.workbook.Changed.Count')==ev('window.listenersBefore')-1)
    check('Independent web controls share one engine and detach cleanly',shared_views)
    def browser_xlsx():
        outcome=ev("""async()=>{const {exportXlsx,importXlsx}=await import(location.protocol.startsWith('http') ? new URL('../src/io.js',location.href).href : 'gridweb/src/io.js');const bytes=exportXlsx(gridweb.workbook);const result=await importXlsx(bytes);return {count:result.workbook.Worksheets.Count,total:result.workbook.Worksheets.Get('Operating plan').GetCell('L1').Value,warnings:result.warnings};}""")
        require(outcome['count']==5);require(outcome['total']==42)
    check('Browser XLSX ZIP/XML export-import round-trip',browser_xlsx)
    def narrow_layout():
        page.set_viewport_size({'width':760,'height':900});page.wait_for_timeout(150);require(ev('document.documentElement.scrollWidth')<=780);require(ev('gridweb.grid.Metrics.visibleCells')>0);page.screenshot(path=str(OUT/'studio-compact.png'),full_page=True);page.set_viewport_size({'width':1600,'height':1000})
    check('Compact workspace keeps document viewport usable',narrow_layout)
    def large_dataset():
        elapsed=ev("""()=>{const start=performance.now();const b=gridweb.createLargeShowcase(20000);gridweb.grid.Workbook=b;gridweb.grid.Select('XFD1048576');return performance.now()-start;}""");page.wait_for_timeout(180);require(ev('gridweb.grid.ActiveCell.Address')=='XFD1048576');require(ev('gridweb.grid.Metrics.visibleCells')<1600);require(ev('gridweb.grid.ActiveCell.Value') is not None);results.append({'name':'20,000-row creation and last-address navigation measurement','status':'measurement','durationMs':round(elapsed,2),'visibleCells':ev('gridweb.grid.Metrics.visibleCells')})
    check('Sparse full-sheet addressing with 20,000 editable rows',large_dataset)
    host_page=browser.new_page(viewport={'width':1000,'height':700})
    host_page.on('pageerror',lambda e:errors.append(str(e)))
    def embedded_host_ready():
        host_page.set_content((ROOT/'dist/GridWeb-host.html').read_text(),wait_until='load')
        host_page.wait_for_function('!!globalThis.gridWebHost')
        outcome=host_page.evaluate('JSON.parse(gridWebHost.dispatch(JSON.stringify({id:1,method:"capabilities"})))')
        require(outcome['id']==1);require('range.values.set' in outcome['result']['methods'])
    check('Embedded minimal host initializes the same engine and allowlisted RPC',embedded_host_ready)
    def embedded_host_edit():
        outcome=host_page.evaluate("""()=>{const call=(method,args={})=>JSON.parse(gridWebHost.dispatch(JSON.stringify({id:2,method,...args})));call('range.values.set',{address:'A1:B1',values:[[6,7]]});call('range.formulas.set',{address:'C1',formulas:[['=A1*B1']]});return call('range.values.get',{address:'C1'});} """)
        require(outcome['result']==[[42]])
    check('Native-host JSON range API writes and calculates real cells',embedded_host_edit)
    def embedded_host_events():
        outcome=host_page.evaluate("""()=>{gridWebHost.dispatch(JSON.stringify({id:3,method:'view.selection.set',address:'C1'}));const events=gridWebHost.drainEvents();return {events,remaining:gridWebHost.drainEvents(),blocked:JSON.parse(gridWebHost.dispatch(JSON.stringify({id:4,method:'constructor.constructor',code:'alert(1)'})))};}""")
        require(any(e['type']=='selection-changed' for e in outcome['events']));require(any(e['type']=='workbook-changed' for e in outcome['events']));require(outcome['remaining']==[]);require('error' in outcome['blocked'])
        host_page.close()
    check('Embedded host drains selection/change events and rejects arbitrary invocation',embedded_host_events)
    check('No uncaught JavaScript page errors',lambda:require(not errors,'; '.join(errors)))
    report={'browser':browser.version,'mode':'HTTP' if os.getenv('GRIDWEB_BASE_URL') else 'offline HTML via set_content','passed':sum(r['status']=='passed' for r in results),'failed':sum(r['status']=='failed' for r in results),'results':results,'pageErrors':errors}
    (OUT/'browser.json').write_text(json.dumps(report,indent=2));browser.close()
print(json.dumps({k:report[k] for k in ['browser','passed','failed']},indent=2))
raise SystemExit(1 if report['failed'] else 0)
