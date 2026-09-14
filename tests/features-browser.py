"""Real Chromium qualification for the Office-compatible API and print additions."""
from pathlib import Path
import json, os, time
from playwright.sync_api import sync_playwright
ROOT=Path(__file__).resolve().parents[1]
OUT=ROOT/'test-results'; OUT.mkdir(exist_ok=True)
results=[]
def require(value,message='Assertion failed'):
    if not value: raise AssertionError(message)
def check(name,action):
    start=time.perf_counter()
    try:
        action(); results.append({'name':name,'status':'passed','durationMs':round((time.perf_counter()-start)*1000,2)}); print('PASS',name,flush=True)
    except Exception as e:
        results.append({'name':name,'status':'failed','error':str(e)}); print('FAIL',name,str(e),flush=True)
with sync_playwright() as p:
    browser=p.chromium.launch(executable_path=os.getenv('CHROMIUM_PATH','/usr/bin/chromium'),headless=True,args=['--no-sandbox'])
    page=browser.new_page(viewport={'width':1600,'height':1000})
    errors=[]; page.on('pageerror',lambda e:errors.append(str(e)))
    if os.getenv('GRIDWEB_BASE_URL'): page.goto(os.environ['GRIDWEB_BASE_URL'],wait_until='networkidle')
    else: page.set_content((ROOT/'dist/GridWeb-standalone.html').read_text(),wait_until='load')
    page.wait_for_function('!!window.gridwebExtensions')
    def batch():
        value=page.evaluate('''async()=>{
          const {workbook,grid}=gridweb,Excel=gridwebExtensions.createExcelApi(workbook);let deferred=false,notLoaded=false;
          await Excel.run(async c=>{const r=c.workbook.worksheets.getActiveWorksheet().getRange('M1:N1');r.values=[[4,5]];deferred=grid.Sheet.GetCell('M1').Value===null;r.load('values');try{r.values;}catch(e){notLoaded=e.code==='PropertyNotLoaded';}await c.sync();if(r.values[0][1]!==5)throw Error('Incorrect snapshot');});
          const after=grid.Sheet.GetCell('N1').Value;workbook.Undo();return {deferred,notLoaded,after,undo:grid.Sheet.GetCell('N1').Value};
        }''')
        require(value=={'deferred':True,'notLoaded':True,'after':5,'undo':None},str(value))
    check('Batched API updates and undoes the displayed workbook',batch)
    def rollback():
        require(page.evaluate('''async()=>{const {workbook,grid}=gridweb;try{await gridwebExtensions.createExcelApi(workbook).run(async c=>{const r=c.workbook.worksheets.getActiveWorksheet().getRange('M1');r.values=[[8]];r.values=[[1,2]];await c.sync();});return false;}catch(e){return e.code==='InvalidArgument'&&grid.Sheet.GetCell('M1').Value===null;}}'''))
    check('Failed batched edit rolls back in the browser',rollback)
    def demo():
        page.get_by_role('tab',name='View',exact=True).click()
        page.get_by_role('button',name='Office API example',exact=True).click()
        page.wait_for_function("gridweb.grid.Sheet.Name==='API example'")
        require(page.evaluate("gridweb.grid.Sheet.GetCell('D4').Value")==210)
        require('210' in page.locator('#toast').inner_text())
        page.screenshot(path=str(OUT/'office-api-example.png'),full_page=True)
    check('Studio Office API example is an executable shared-engine batch',demo)
    def svg_download():
        page.evaluate("gridweb.grid.Sheet.AddChart('A1:D3',{column:4,row:1,type:'column'})")
        page.get_by_role('tab',name='Insert',exact=True).click()
        with page.expect_download() as info: page.get_by_role('button',name='Chart SVG',exact=True).click()
        download=info.value; dest=OUT/'chart-export.svg';download.save_as(dest)
        require(dest.read_text().startswith('<svg'));require('<script' not in dest.read_text())
    check('Chart SVG command exports actual vector chart artwork',svg_download)
    def preview():
        page.evaluate("gridweb.grid.Sheet.AddConditionalFormat('D2:D3',{type:'dataBar'});gridweb.grid.Sheet.AddConditionalFormat('D2:D3',{type:'cellValue',criteria:'>100',style:{fill:'#ff0000'}})")
        page.get_by_role('tab',name='Page Layout',exact=True).click()
        page.get_by_role('button',name='Advanced print',exact=True).click()
        page.get_by_label('Paper',exact=True).select_option('Letter')
        with page.expect_popup() as info: page.get_by_role('button',name='Open preview',exact=True).click()
        out=info.value;out.wait_for_load_state()
        require(out.locator('.print-chart svg').count()>0)
        require(out.locator('.print-data-bar').count()==2)
        require(out.locator('td[style*="background:#ff0000"]').count()==2)
        require(out.evaluate('window.opener===null'))
        out.screenshot(path=str(OUT/'printed-charts-and-formatting.png'),full_page=True)
        out.close()
    check('Advanced print displays vector charts and conditional formats in an isolated window',preview)
    def pages():
        value=page.evaluate('''()=>{const s=gridweb.grid.Sheet;s.GetCell('J90').Value=1;const p=gridwebExtensions.paginate(s,{area:'A1:J90',paper:'A3',orientation:'portrait',fitToWidthPages:1,repeatRows:1,repeatColumns:1});return {pages:p.pages.length,width:p.pages[0].width,columns:p.pages.every(p=>p.columns.length===10),repeated:p.pages.slice(1).every(p=>p.rows[0]===0)};}''')
        require(value['pages']>1 and value['width']==1123 and value['columns'] and value['repeated'],str(value))
    check('Fitted A3 page geometry and repeated titles work in Chromium',pages)
    def invalid_options():
        page.get_by_role('button',name='Advanced print',exact=True).click()
        page.get_by_label('Print range (blank = used area)').fill('A0')
        page.get_by_role('button',name='Open preview',exact=True).click()
        require(bool(page.locator('dialog[open] [role="alert"]').inner_text()))
        page.get_by_role('button',name='Close advanced print',exact=True).click()
    check('Invalid print area retains the dialog and explains the error',invalid_options)
    check('No uncaught page errors',lambda:require(not errors,str(errors)))
    browser.close()
(OUT/'features-browser.json').write_text(json.dumps({'results':results,'errors':errors},indent=2))
print(f'{sum(r["status"]=="passed" for r in results)}/{len(results)} feature browser groups passed')
raise SystemExit(1 if any(r['status']=='failed' for r in results) else 0)
