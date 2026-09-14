# Calculation compatibility update

```js
import {Workbook,a1ToR1C1,setFormulasR1C1} from '@wieslawsoltes/gridweb';
const book=new Workbook();
const sheet=book.ActiveWorksheet;
sheet.GetRange('A2:B3').Values=[[2,3],[4,5]];
setFormulasR1C1(sheet.GetRange('C2:C3'),[['=RC[-2]*RC[-1]'],['=RC[-2]*RC[-1]']]);
console.log(sheet.GetRange('C2:C3').Values); // [[6],[20]]
console.log(a1ToR1C1('=$A1+B$2','B2')); // =R[-1]C1+R2C
sheet.GetCell('E1').Formula='=ABS({-1,-2;-3,-4})'; // real 2x2 spill
sheet.GetCell('H1').Formula='=NORM.S.INV(0.975)';
```

Microsoft interoperability references:
- https://support.microsoft.com/en-us/office/address-function-d0c26c0d-3991-446b-8de4-ab46431d4f89
- https://support.microsoft.com/en-us/office/xmatch-function-d966da31-7a6b-4a13-a1c6-5a33ed6a0312
- https://support.microsoft.com/en-us/excel/functions/count-function
- https://learn.microsoft.com/en-us/office/vba/api/excel.worksheetfunction.sum

The recovered compatibility matrix describes the 0.1 baseline. This update replaces its binary lookup, R1C1, selected array lifting and additional-function gaps with the tested surfaces described here. It does not close VBA, Power Query/DAX, arbitrary OOXML preservation, collaboration, Office.js/COM or printer/native-runtime gaps.
