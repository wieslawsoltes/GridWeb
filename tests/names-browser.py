"""Real Chromium interaction checks for scoped names and reusable Name Manager."""
from pathlib import Path
import json, os, time
from playwright.sync_api import sync_playwright
ROOT=Path(__file__).resolve().parents[1];OUT=ROOT/'test-results';OUT.mkdir(exist_ok=True)
results=[]
def require(ok,message='Assertion failed'):
    if not ok: raise AssertionError(message)
def check(name,action):
    start=time.perf_counter()
    try:
        action();results.append({'name':name,'status':'passed','durationMs':round((time.perf_counter()-start)*1000,2)});print('PASS',name,flush=True)
    except Exception as e:
        results.append({'name':name,'status':'failed','error':str(e)});print('FAIL',name,str(e),flush=True)
        try: page.evaluate('gridweb.grid.ReadOnly=false;for(const d of gridweb.grid.shadowRoot.querySelectorAll("dialog[open]"))d.close()')
        except Exception: pass
with sync_playwright() as p:
    browser=p.chromium.launch(executable_path=os.getenv('CHROMIUM_PATH','/usr/bin/chromium'),headless=True,args=['--no-sandbox'])
    page=browser.new_page(viewport={'width':1600,'height':1000});page.set_default_timeout(5000);errors=[];page.on('pageerror',lambda e:errors.append(str(e)))
    if os.getenv('GRIDWEB_BASE_URL'):page.goto(os.environ['GRIDWEB_BASE_URL'],wait_until='networkidle')
    else:page.set_content((ROOT/'dist/GridWeb-standalone.html').read_text(),wait_until='load')
    page.wait_for_function('!!window.gridweb?.commands.namedFormulas');ev=page.evaluate
    def dlg(title):return page.get_by_role('dialog',name=title,exact=True)
    def owner(name='Names North'):ev('(name)=>{gridweb.grid.Sheet=gridweb.workbook.Worksheets.Get(name);gridweb.grid.Focus()}',name)
    def manager():ev('gridweb.grid.ShowNameManager()');return dlg('Name manager')
    def close():page.keyboard.press('Escape')
    def summary(address):return ev('(a)=>{const v=gridweb.workbook.Worksheets.Get("Named formulas").GetCell(a).Value;return v?.code??v}',address)
    def define(name,formula,scope='Names North'):
        ev('gridweb.grid.ShowDefineName()');d=dlg('Define name');d.get_by_label('Name',exact=True).fill(name);d.get_by_label('Scope',exact=True).select_option(label=scope);d.get_by_label('Refers to',exact=True).fill(formula);return d
    def examples():
        require(ev('gridweb.workbook.Worksheets.Count')==5)
        page.get_by_role('tab',name='Home',exact=True).click();page.get_by_role('button',name='Named formulas',exact=True).click();page.wait_for_function('gridweb.grid.Sheet.Name==="Named formulas"')
        require(ev('gridweb.workbook.Worksheets.Count')==8)
        for address,value in [('B4',600),('B5',1200),('B6',60),('B7',240),('B8',True),('B9',300),('B10',True),('B11',1800)]:require(summary(address)==value,address)
        require(ev('gridweb.grid.Sheet.GetCell("C4").Formula') is None);page.screenshot(path=str(OUT/'named-formulas-desktop.png'),full_page=True)
    check('Named formulas sample creates editable independent scopes and reference bindings',examples)
    def search_scopes():
        d=manager();d.get_by_label('Search names',exact=True).fill('Revenue');d.get_by_label('Scope filter',exact=True).select_option('local')
        require(d.get_by_role('button',name='Edit Revenue in Names North',exact=True).count()==1);require(d.get_by_role('button',name='Edit Revenue in Names South',exact=True).count()==1)
        page.screenshot(path=str(OUT/'name-manager-desktop.png'),full_page=True);close()
    check('Name Manager searches duplicate names while retaining their worksheet scopes',search_scopes)
    def create_name():
        owner();d=define('UiAmount','=B2');d.get_by_label('Comment',exact=True).fill('Created through the reusable UI');d.get_by_role('button',name='Save',exact=True).click()
        require(ev('gridweb.grid.Sheet.Names.Evaluate("UiAmount")')==100);require(ev('gridweb.grid.Sheet.Names.GetDefinition("UiAmount").comment')=='Created through the reusable UI')
        page.keyboard.press('Control+z');require(not ev('gridweb.grid.Sheet.Names.Has("UiAmount")'));page.keyboard.press('Control+y');require(ev('gridweb.grid.Sheet.Names.Has("UiAmount")'))
    check('Define name UI creates metadata in the selected scope with keyboard undo and redo',create_name)
    def invalid_name():
        d=define('A1','=42');before=ev('JSON.stringify(gridweb.workbook.ToJSON())');d.get_by_role('button',name='Save',exact=True).click();require(d.locator('[role=alert]').inner_text()!='');require(ev('JSON.stringify(gridweb.workbook.ToJSON())')==before);close()
    check('Invalid name creation reports an inline error without mutating the workbook',invalid_name)
    def duplicate():
        d=define('Revenue','=99');d.get_by_role('button',name='Save',exact=True).click();require('already exists' in d.locator('[role=alert]').inner_text());require(summary('B4')==600);close()
    check('Duplicate names in one scope are rejected without overwriting existing definitions',duplicate)
    def rename():
        d=manager();d.get_by_label('Search names',exact=True).fill('Revenue');d.get_by_role('button',name='Edit Revenue in Names North',exact=True).click()
        e=dlg('Edit defined name');require(e.get_by_label('Scope',exact=True).is_disabled());e.get_by_label('Name',exact=True).fill('RegionalRevenue');e.get_by_role('button',name='Save',exact=True).click()
        require(summary('B4')==600);require('RegionalRevenue' in ev('gridweb.workbook.Worksheets.Get("Named formulas").GetCell("B4").Formula'));require(ev('gridweb.workbook.Worksheets.Get("Names South").Names.Has("Revenue")'))
        page.keyboard.press('Control+z');require(ev('gridweb.workbook.Worksheets.Get("Names North").Names.Has("Revenue")'));require(summary('B4')==600)
        owner();ev('gridweb.grid.Sheet.Names.Create("RecursiveUi","=LAMBDA(n,IF(n=0,0,n+RecursiveUi(n-1)))");gridweb.grid.Sheet.GetCell("M1").Formula="=RecursiveUi(3)"')
        d=manager();d.get_by_label('Search names',exact=True).fill('RecursiveUi');d.get_by_role('button',name='Edit RecursiveUi in Names North',exact=True).click()
        e=dlg('Edit defined name');e.get_by_label('Name',exact=True).fill('RecursiveRenamed');e.get_by_label('Refers to',exact=True).fill('=LAMBDA(n,IF(n=0,0,n+RecursiveUi(n-1)+0))');e.get_by_role('button',name='Save',exact=True).click()
        require(ev('gridweb.grid.Sheet.GetCell("M1").Value')==6);require('RecursiveRenamed(n-1)' in ev('gridweb.grid.Sheet.Names.Get("RecursiveRenamed")'));page.keyboard.press('Control+z');require(ev('gridweb.grid.Sheet.Names.Has("RecursiveUi")'));require(ev('gridweb.grid.Sheet.GetCell("M1").Value')==6)
    check('Editing a name renames free formula uses without changing other worksheet scopes',rename)
    def edit_source():
        owner();ev('gridweb.grid.Select("B2");gridweb.grid.Focus()');page.keyboard.press('F2');page.keyboard.press('Control+a');page.keyboard.type('150');page.keyboard.press('Enter')
        require(summary('B4')==650);require(summary('B6')==65);require(summary('B5')==1200);page.keyboard.press('Control+z');require(summary('B4')==600)
    check('Native cell edits recalculate scoped references and named LAMBDA results',edit_source)
    def delete():
        d=manager();d.get_by_label('Search names',exact=True).fill('Revenue');button=d.get_by_role('button',name='Delete Revenue in Names North',exact=True);button.click();require(button.inner_text()=='Confirm delete');require(summary('B4')==600);button.click();require(summary('B4')=='#NAME?');require(summary('B5')==1200);close();page.keyboard.press('Control+z');require(summary('B4')==600)
    check('Delete requires confirmation, affects one scope only, and remains undoable',delete)
    def goto():
        owner('Named formulas');d=manager();d.get_by_label('Search names',exact=True).fill('TaxRate');d.get_by_role('button',name='Go to TaxRate in Names North',exact=True).click()
        require(ev('gridweb.grid.Sheet.Name')=='Names North');require(ev('gridweb.grid.Selection')=='C2')
    check('Go to resolves a local named range across worksheets through the grid API',goto)
    def stale():
        d=manager();d.get_by_label('Search names',exact=True).fill('UiAmount');d.get_by_role('button',name='Edit UiAmount in Names North',exact=True).click();e=dlg('Edit defined name');e.get_by_label('Refers to',exact=True).fill('=999');ev('gridweb.workbook.Worksheets.Get("Names North").Names.Update("UiAmount","=123")');e.get_by_role('button',name='Save',exact=True).click();require('changed' in e.locator('[role=alert]').inner_text());require(ev('gridweb.grid.Sheet.Names.Evaluate("UiAmount")')==123);close();page.keyboard.press('Control+z')
    check('Stale name edits do not overwrite an intervening model update',stale)
    def readonly():
        d=define('ReadOnlyName','=9');ev('gridweb.grid.ReadOnly=true');d.get_by_role('button',name='Save',exact=True).click();require('read-only' in d.locator('[role=alert]').inner_text());require(not ev('gridweb.grid.Sheet.Names.Has("ReadOnlyName")'));close();ev('gridweb.grid.ReadOnly=false')
    check('Apply-time read-only enforcement protects name creation after a dialog opens',readonly)
    def unsafe():
        d=define('SafeText','="<img src=x onerror=alert(1)>"');d.get_by_label('Comment',exact=True).fill('<script>window.bad=true</script>');d.get_by_role('button',name='Save',exact=True).click();d=manager();d.get_by_label('Search names',exact=True).fill('SafeText');require('<script>' in d.inner_text());require(d.locator('img').count()==0 and d.locator('script').count()==0);require(not ev('!!window.bad'));close()
    check('Untrusted formula and comment text remains inert in Name Manager',unsafe)
    def keyboard():
        ev('gridweb.grid.Focus()');page.keyboard.press('Control+F3');require(dlg('Name manager').is_visible());close();ev('gridweb.grid._showContext(100,100)');page.locator('grid-web .context').get_by_role('button',name='Name manager',exact=True).click();require(dlg('Name manager').is_visible());close()
    check('Keyboard and context menu open the reusable Name Manager',keyboard)
    def independent():
        ev('window.otherGrid=document.createElement("grid-web");otherGrid.style.cssText="height:300px;display:block";document.body.append(otherGrid);otherGrid.Workbook=gridweb.workbook;otherGrid.ShowNameManager()');d=page.locator('grid-web').last.get_by_role('dialog',name='Name manager',exact=True);require(d.is_visible());d.get_by_role('button',name='Cancel',exact=True).click();ev('otherGrid.remove();delete window.otherGrid')
    check('Independent grid instances use Name Manager without importing sample controls',independent)
    def mobile():
        page.set_viewport_size({'width':390,'height':844});ev('gridweb.grid.Theme="dark"');d=manager();d.get_by_label('Search names',exact=True).fill('Revenue');box=d.bounding_box();require(box['x']>=0 and box['x']+box['width']<=390 and box['height']<=844,str(box));page.screenshot(path=str(OUT/'name-manager-mobile-dark.png'),full_page=True);close();page.set_viewport_size({'width':1600,'height':1000})
    check('Name Manager fits narrow dark-mode viewports with labeled native controls',mobile)
    def from_selection():
        owner();ev('gridweb.grid.Sheet.GetRange("H20:I22").Values=[["Ui category","Ui amount"],["One",2],["Two",3]];gridweb.grid.Select("H20:I22");gridweb.grid.ShowCreateNamesFromSelection()')
        d=dlg('Create names from selection');d.get_by_label('Scope',exact=True).select_option(label='Names North');d.get_by_role('button',name='Create names',exact=True).click()
        require(ev('gridweb.grid.Sheet.Names.GetRange("Ui_amount").FullAddress')=="'Names North'!I21:I22");require(ev('gridweb.grid.Sheet.Names.GetRange("Ui_category").Values')==[['One'],['Two']]);page.keyboard.press('Control+z');require(not ev('gridweb.grid.Sheet.Names.Has("Ui_amount")'))
    check('Create from Selection authors absolute named ranges from labels in one undoable operation',from_selection)
    def stale_selection():
        ev('gridweb.grid.ShowCreateNamesFromSelection();gridweb.grid.Sheet.GetCell("I21").Value=4')
        d=dlg('Create names from selection');d.get_by_role('button',name='Create names',exact=True).click();require('changed' in d.locator('[role=alert]').inner_text());require(not ev('gridweb.grid.Sheet.Names.Has("Ui_amount")'));close();page.keyboard.press('Control+z')
    check('Create from Selection refuses stale labels instead of naming intervening data',stale_selection)
    def lifecycle():
        ev('gridweb.grid.ShowDefineName();gridweb.grid.ShowNameManager()');page.wait_for_timeout(100);require(ev('gridweb.grid.shadowRoot.querySelector("dialog[open]").contains(gridweb.grid.shadowRoot.activeElement)'))
        ev('const g=gridweb.grid,p=g.parentNode,next=g.nextSibling;g.remove();p.insertBefore(g,next)');page.wait_for_function('gridweb.grid.shadowRoot.querySelectorAll("dialog[open]").length===0');require(summary('B4')==600)
    check('Dialog focus and disconnect cleanup preserve the shared workbook',lifecycle)
    check('No uncaught named-formula page errors',lambda:require(not errors,str(errors)))
    browser.close()
(OUT/'names-browser.json').write_text(json.dumps({'results':results,'errors':errors},indent=2))
print(f'{sum(r["status"]=="passed" for r in results)}/{len(results)} names browser groups passed')
raise SystemExit(1 if any(r['status']=='failed' for r in results) else 0)
