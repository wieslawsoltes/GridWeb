from pathlib import Path
import json,os
from playwright.sync_api import sync_playwright
ROOT=Path(__file__).resolve().parents[1];OUT=ROOT/'test-results';OUT.mkdir(exist_ok=True);results=[];errors=[]
def require(value,message='Assertion failed'):
 if not value:raise AssertionError(message)
def check(name,action):
 try:action();results.append({'name':name,'status':'passed'});print('PASS',name,flush=True)
 except Exception as e:results.append({'name':name,'status':'failed','error':str(e)});print('FAIL',name,str(e),flush=True)
with sync_playwright() as p:
 browser=p.chromium.launch(executable_path=os.getenv('CHROMIUM_PATH','/usr/bin/chromium'),headless=True,args=['--no-sandbox']);page=browser.new_page(viewport={'width':1800,'height':1200});page.set_default_timeout(5000);page.on('pageerror',lambda e:(errors.append(str(e)),print('PAGE ERROR',str(e),flush=True)))
 page.set_content((ROOT/'dist/GridWeb-standalone.html').read_text(),wait_until='load');page.wait_for_function('!!window.gridwebPivots')
 page.evaluate("{const w=gridweb.workbook,s=w.Worksheets.Add('Pivot source');s.GetRange('A1:C5').Values=[['Region','Product','Revenue'],['North','A',10],['North','B',20],['South','A',30],['South','B',40]];gridweb.grid.Sheet=s;gridweb.grid.Select('A1:C5');}")
 def create():
  page.get_by_role('tab',name='Data',exact=True).click();page.get_by_role('button',name='Pivot reports',exact=True).click();page.get_by_label('Pivot name',exact=True).fill('BrowserPivot');page.get_by_label('Row fields (comma-separated)').fill('Region');page.get_by_label('Value field',exact=True).fill('Revenue');page.get_by_role('button',name='Create pivot',exact=True).click();page.wait_for_function('gridweb.workbook.PivotTables.Count===1');require(page.evaluate("gridweb.grid.Sheet.GetRange('A1:B4').Values")==[['Region','sum of Revenue'],['North',30],['South',70],['Grand Total',100]])
 check('Create managed pivot through the studio field dialog',create)
 def refresh():
  page.evaluate("gridweb.workbook.Worksheets.Get('Pivot source').GetCell('C2').Value=15");require(page.evaluate('gridweb.workbook.PivotTables.Get(0).IsStale'));page.get_by_role('button',name='Refresh pivots',exact=True).click();require(page.evaluate("gridweb.grid.Sheet.GetCell('B2').Value")==35);require(not page.evaluate('gridweb.workbook.PivotTables.Get(0).IsStale'))
 check('Refresh command updates results from source edits',refresh)
 def edit():
  page.get_by_role('button',name='Edit pivot',exact=True).click();page.locator('#pivot-reports-dialog summary').click();page.get_by_label('Advanced pivot options').fill('{"filters":[{"column":"Region","values":["North"]}]}');page.get_by_role('button',name='Update pivot',exact=True).click();page.wait_for_function("gridweb.workbook.PivotTables.Get(0).OutputRange.RowCount===3");require(page.evaluate("gridweb.grid.Sheet.GetCell('B3').Value")==35)
 check('Edit pivot preserves definition and applies field filters',edit)
 def details():
  page.evaluate("gridweb.grid.Select('B2')");page.get_by_role('button',name='Show details',exact=True).click();require(page.evaluate("gridweb.grid.Sheet.GetRange('A1:C3').Values")==[['Region','Product','Revenue'],['North','A',15],['North','B',20]])
 check('Drill-down command creates a source-record worksheet',details)
 def save():
  with page.expect_download() as dl:page.evaluate('gridweb.commands.save()')
  path=OUT/'pivot-workbook.json';dl.value.save_as(path);doc=json.loads(path.read_text());require(len(doc['pivotTables'])==1)
  with page.expect_download() as dl:page.evaluate('gridweb.commands.xlsx()')
  path=OUT/'pivot-workbook.xlsx';dl.value.save_as(path)
  import zipfile
  with zipfile.ZipFile(path) as z:require('xl/pivotTables/pivotTable1.xml' in z.namelist())
 check('Normal studio exports retain JSON definitions and native XLSX parts',save)
 page.screenshot(path=str(OUT/'pivot-studio.png'),full_page=True)
 check('No uncaught pivot UI errors',lambda:require(not errors,str(errors)))
 browser.close()
(OUT/'pivot-browser.json').write_text(json.dumps({'results':results,'errors':errors},indent=2))
raise SystemExit(1 if any(r['status']=='failed' for r in results) else 0)
