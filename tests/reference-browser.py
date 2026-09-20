"""Chromium checks for live reference formulas and control-owned sheet movement."""
from pathlib import Path
import json, os, time
from playwright.sync_api import sync_playwright
ROOT=Path(__file__).resolve().parents[1]
OUT=ROOT/'test-results'; OUT.mkdir(exist_ok=True)
results=[]
def require(condition, message='Assertion failed'):
    if not condition: raise AssertionError(message)
def check(name, action):
    start=time.perf_counter()
    try:
        action(); results.append({'name':name,'status':'passed','durationMs':round((time.perf_counter()-start)*1000,2)}); print('PASS',name,flush=True)
    except Exception as error:
        results.append({'name':name,'status':'failed','error':str(error)}); print('FAIL',name,str(error),flush=True)
        try: page.evaluate('for(const d of gridweb.grid.shadowRoot.querySelectorAll("dialog[open]"))d.close();gridweb.grid.ReadOnly=false')
        except Exception: pass
with sync_playwright() as p:
    browser=p.chromium.launch(executable_path=os.getenv('CHROMIUM_PATH','/usr/bin/chromium'),headless=True,args=['--no-sandbox'])
    page=browser.new_page(viewport={'width':1600,'height':1000}); page.set_default_timeout(5000)
    errors=[]; page.on('pageerror',lambda error:errors.append(str(error)))
    if os.getenv('GRIDWEB_BASE_URL'): page.goto(os.environ['GRIDWEB_BASE_URL'],wait_until='networkidle')
    else: page.set_content((ROOT/'dist/GridWeb-standalone.html').read_text(),wait_until='load')
    page.wait_for_function('!!window.gridweb?.commands.referenceTools')
    ev=page.evaluate
    def value(address): return ev('(a)=>gridweb.workbook.Worksheets.Get("Reference tools").GetCell(a).Value',address)
    def order(): return ev('gridweb.workbook.Worksheets.items.map(s=>s.Name)')
    def open_move(): ev('gridweb.grid.ShowMoveWorksheet()')
    def choose_sheet(name): page.locator('grid-web dialog[open]').get_by_label('Worksheet',exact=True).select_option(label=name)
    def move_to(index):
        page.locator('grid-web dialog[open]').get_by_label('Final position',exact=True).select_option(str(index))
        page.get_by_role('dialog',name='Move worksheet',exact=True).get_by_role('button',name='Move',exact=True).click()
    def examples():
        require(ev('gridweb.workbook.Worksheets.Count')==5)
        page.get_by_role('tab',name='Home',exact=True).click()
        page.get_by_role('button',name='Reference tools',exact=True).click()
        page.wait_for_function('gridweb.grid.Sheet.Name==="Reference tools"')
        require(ev('gridweb.workbook.Worksheets.Count')==9)
        require(value('B2')==450);require(value('E5')==110);require(value('E6')==50);require(value('E7')==60)
        require(value('E8')==20);require(value('E9')==3);require(value('E10') is True);require(value('E11')==1)
        require(ev('gridweb.grid.Sheet.GetCell("F5").Formula') is None,'Formula caption was executed')
        page.screenshot(path=str(OUT/'reference-tools-desktop.png'),full_page=True)
    check('Reference tools creates editable sources and reference-backed results without changing startup sheets',examples)
    def source_edit():
        page.get_by_role('tab',name='Reference North',exact=True).click()
        ev('gridweb.grid.Select("B2");gridweb.grid.Focus()');page.keyboard.press('F2');page.keyboard.press('Control+a');page.keyboard.type('110');page.keyboard.press('Enter')
        require(value('B2')==550);page.keyboard.press('Control+z');require(value('B2')==450)
        page.get_by_role('tab',name='Reference tools',exact=True).click()
    check('Native editor updates a source sheet and recalculates 3-D dependents',source_edit)
    def move_interior():
        before=order();page.get_by_role('button',name='Move worksheet',exact=True).click();choose_sheet('Reference Central')
        page.screenshot(path=str(OUT/'move-worksheet-dialog.png'),full_page=True)
        move_to(7);require(value('B2')==300);require(value('E9')==2)
        require(ev('gridweb.grid.Sheet.Name')=='Reference tools')
        rendered=page.get_by_role('tab').all_text_contents();require(rendered.index('Reference Central')>rendered.index('Reference South'),str(rendered))
        page.keyboard.press('Control+z');require(value('B2')==450);require(order()==before)
        page.keyboard.press('Control+y');require(value('B2')==300);page.keyboard.press('Control+z')
    check('Move worksheet dialog changes 3-D membership and keyboard undo restores order and formulas',move_interior)
    def endpoint_move():
        ev('gridweb.workbook.ClearHistory()');open_move();choose_sheet('Reference North');move_to(8)
        require(value('B2')==390);require('Reference Central:Reference South' in ev('gridweb.grid.Sheet.GetCell("B2").Formula'))
        page.keyboard.press('Control+z');require(value('B2')==450)
    check('Moving a 3-D endpoint across its opposite endpoint repairs formulas in one transaction',endpoint_move)
    def readonly():
        before=order();open_move();choose_sheet('Reference Central');ev('gridweb.grid.ReadOnly=true');move_to(7)
        require('read-only' in page.locator('grid-web dialog[open] [role=alert]').inner_text());require(order()==before)
        page.keyboard.press('Escape');ev('gridweb.grid.ReadOnly=false')
    check('Read-only changes made after opening the dialog are enforced at apply time',readonly)
    def stale():
        open_move();choose_sheet('Reference Central');ev('gridweb.workbook.Worksheets.Move("Reference Central",7)');before=order();move_to(6)
        require('order changed' in page.locator('grid-web dialog[open] [role=alert]').inner_text());require(order()==before)
        page.keyboard.press('Escape');page.keyboard.press('Control+z');require(value('B2')==450)
    check('Stale position choices are rejected without overwriting intervening sheet movement',stale)
    def reference_edit():
        ev('gridweb.grid.Select("H5");gridweb.grid.Focus()');page.keyboard.press('F2');page.keyboard.type('=SUM((A5:A8,B5:B8) A6:B7)');page.keyboard.press('Enter')
        require(value('H5')==55);page.keyboard.press('Control+z');require(value('H5') is None)
    check('Union and intersection formulas entered through the control use live engine references',reference_edit)
    def rename_delete():
        ev('gridweb.workbook.Worksheets.Get("Reference North").Name="North renamed"')
        require(value('B2')==450);require('North renamed' in ev('gridweb.grid.Sheet.GetCell("B2").Formula'))
        ev('gridweb.workbook.Worksheets.Remove("North renamed")');require(value('B2')==390)
        ev('gridweb.workbook.Undo();gridweb.workbook.Undo()');require(value('B2')==450)
    check('Rename and deletion repair displayed endpoint references, with undo retaining sheet identity',rename_delete)
    def independent_control():
        ev('window.otherGrid=document.createElement("grid-web");otherGrid.style.cssText="height:300px;display:block";document.body.append(otherGrid);otherGrid.Workbook=gridweb.workbook;otherGrid.Sheet=gridweb.grid.Sheet;otherGrid.ShowMoveWorksheet()')
        dialog=page.locator('grid-web').last.get_by_role('dialog',name='Move worksheet',exact=True)
        require(dialog.is_visible());dialog.get_by_role('button',name='Cancel',exact=True).click();ev('otherGrid.remove();delete window.otherGrid')
        require(value('B2')==450)
    check('Independent reusable controls expose sheet movement without sample-owned dialog code',independent_control)
    def context_menu():
        ev('gridweb.grid._showContext(100,100)');page.locator('grid-web .context').get_by_role('button',name='Move worksheet',exact=True).click()
        require(page.get_by_role('dialog',name='Move worksheet',exact=True).is_visible());page.keyboard.press('Escape')
        page.get_by_role('tab',name='Formulas',exact=True).click();require(page.get_by_role('button',name='Reference tools',exact=True).is_visible())
    check('Context menu and formula ribbon expose the shared reference workflows',context_menu)
    def dialog_focus():
        ev('gridweb.grid.ShowFillSeries();gridweb.grid.ShowMoveWorksheet()')
        page.wait_for_timeout(100)
        require(ev('gridweb.grid.shadowRoot.querySelector("dialog[open]").contains(gridweb.grid.shadowRoot.activeElement)'), 'Previous dialog close stole focus')
        page.keyboard.press('Escape')
    check('Closing an earlier dialog cannot steal focus from a newly opened move dialog',dialog_focus)
    def mobile():
        page.set_viewport_size({'width':390,'height':844});ev('gridweb.grid.Theme="dark";gridweb.grid.ShowMoveWorksheet()')
        box=page.get_by_role('dialog',name='Move worksheet',exact=True).bounding_box()
        require(box['x']>=0 and box['x']+box['width']<=390 and box['height']<=844,str(box))
        page.screenshot(path=str(OUT/'move-worksheet-mobile-dark.png'),full_page=True)
        page.keyboard.press('Escape');page.set_viewport_size({'width':1600,'height':1000})
    check('Move worksheet dialog fits dark mobile viewports with native labeled inputs',mobile)
    def lifecycle():
        ev('gridweb.grid.ShowMoveWorksheet();const g=gridweb.grid,p=g.parentNode,next=g.nextSibling;g.remove();p.insertBefore(g,next)')
        page.wait_for_function('gridweb.grid.shadowRoot.querySelectorAll("dialog[open]").length===0')
        require(ev('gridweb.grid.isConnected'));require(value('B2')==450)
    check('Disconnecting the control closes the dialog and reconnect retains workbook state',lifecycle)
    check('No uncaught reference workflow errors',lambda:require(not errors,str(errors)))
    browser.close()
(OUT/'reference-browser.json').write_text(json.dumps({'results':results,'errors':errors},indent=2))
print(f'{sum(r["status"]=="passed" for r in results)}/{len(results)} reference browser groups passed')
raise SystemExit(1 if any(r['status']=='failed' for r in results) else 0)
