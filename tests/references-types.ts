import {Workbook,a1ToR1C1,r1c1ToA1,getFormulasR1C1,setFormulasR1C1} from '@wieslawsoltes/gridweb';
const book=new Workbook();
const range=book.ActiveWorksheet.GetRange('C2:C3');
setFormulasR1C1(range,[['=RC[-2]*RC[-1]'],['=RC[-2]*RC[-1]']]);
const values:(string|null)[][]=getFormulasR1C1(range);
const source:string=a1ToR1C1('=$A1+B$2','B2');
console.log(values,r1c1ToA1(source,{row:1,column:1}));
// @ts-expect-error origin indexes are numeric
r1c1ToA1('RC',{row:'1',column:0});
