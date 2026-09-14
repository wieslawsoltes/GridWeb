import {createShowcase as baselineShowcase} from './fixtures-base.js';
export {createLargeShowcase} from './fixtures-base.js';
export function createShowcase(){
  const book=baselineShowcase(),lab=book.Worksheets.Get('Formula lab');
  const cases=[
    ['Binary lookup','=XMATCH(35,{10,20,30,40},-1,2)','Binary ascending lookup, next smaller'],
    ['R1C1 address','=ADDRESS(2,3,2,FALSE)','Absolute row; relative column'],
    ['Normal quantile','=NORM.S.INV(0.975)','Inverse standard normal distribution'],
    ['Student t tail','=T.DIST.2T(2.306004135,8)','Two-tailed probability'],
    ['Matrix determinant','=MDETERM({1,2;3,4})','Pivoted elimination'],
    ['48-bit arithmetic','=BITLSHIFT(1,40)','No signed 32-bit truncation'],
    ['Complex multiplication','=IMPRODUCT("1+i","1-i")','Complex arithmetic'],
    ['International workdays','=NETWORKDAYS.INTL(DATE(2026,9,7),DATE(2026,9,13),"0000011")','Weekend mask starts Monday'],
    ['Unicode text search','=TEXTBEFORE("ß.A.tail","a",1,1)','Original string offsets preserved'],
    ['Dated cash flows','=XIRR({-100,110},{DATE(2025,1,1),DATE(2026,1,1)})','Iterative dated rate of return'],
  ];
  book.Transaction('Calculation compatibility examples',()=>{
    lab.GetRange('A26:C26').Values=[['CALCULATION COMPATIBILITY','Result','Behavior']];
    lab.GetRange('A26:C26').SetStyle({fill:'#107c41',font:{bold:true,color:'#ffffff'}});
    for(const [i,[label,formula,description]]of cases.entries()){
      lab.GetCell(i+26,0).Value=label;lab.GetCell(i+26,1).Formula=formula;lab.GetCell(i+26,2).Value=description;
    }
    lab.GetCell('E27').Formula='=XLOOKUP(2,{1;2},{10,11;20,21})';
    lab.GetCell('E30').Formula='=ABS({-1,-2;-3,-4})';
    lab.GetCell('E34').Formula='=WRAPROWS({1,2,3,4,5},3,0)';
    lab.GetCell('E26').Value='Spill examples';
  });
  book.ClearHistory();return book;
}
