/** Explicit pivot operations shared by web workers and native WebView hosts. */
export function pivotMethods(control){
 const book=()=>control.Workbook,report=r=>{const p=book().PivotTables.Get(r.name);if(!p)throw new RangeError('Unknown pivot report');return p;};
 const range=r=>{if(!r||typeof r.address!=='string')throw new TypeError('A sheet/address range descriptor is required');const s=r.sheet==null?book().ActiveWorksheet:book().Worksheets.Get(r.sheet);if(!s)throw new RangeError('Unknown worksheet');return s.GetRange(r.address);};
 return {
  'pivots.list':()=>book().PivotTables.ToJSON(),
  'pivots.add':r=>book().PivotTables.Add(r.name,range(r.source),range(r.destination),r.options??{}).ToJSON(),
  'pivots.refresh':r=>report(r).Refresh(),
  'pivots.refreshAll':()=>{book().PivotTables.RefreshAll();return true;},
  'pivots.update':r=>report(r).Update({options:r.options,...(r.source?{source:range(r.source)}:{}),...(r.destination?{destination:range(r.destination)}:{}),...(r.newName?{name:r.newName}:{})}).ToJSON(),
  'pivots.filter':r=>report(r).SetFilter(r.field,r.values).ToJSON(),
  'pivots.drillDown':r=>report(r).DrillDown(r.row,r.column,r.options??{}),
  'pivots.remove':r=>book().PivotTables.Remove(r.name,{clear:r.clear===true})
 };
}
