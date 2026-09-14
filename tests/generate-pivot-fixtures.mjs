import fs from 'node:fs/promises';import {Workbook} from '../src/core.js';import {exportXlsx} from '../src/io.js';
await fs.mkdir('test-results/pivots',{recursive:true});
for(const [name,options]of [['simple',{}],['totals',{columnGrandTotals:true}],['multi',{rows:['Region','Product'],values:[{column:2,aggregate:'sum',name:'Sales'},{column:2,aggregate:'average',name:'Mean'}]}],['filtered',{filters:[{column:'Region',values:['North']}]}]]){
 const book=new Workbook(),s=book.ActiveWorksheet;s.Name='Source';s.GetRange('A1:C5').Values=[['Region','Product','Revenue'],['North','A',10],['North','B',20],['South','A',30],['South','B',40]];
 const dest=book.Worksheets.Add('Report');book.PivotTables.Add('Pivot_'+name,s.UsedRange,dest.GetRange('B3'),{rows:['Region'],values:[{column:2,aggregate:'sum',name:'Revenue'}],...options});
 await fs.writeFile('test-results/pivots/'+name+'.xlsx',exportXlsx(book));book.Dispose();
}
