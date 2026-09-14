import React, {useMemo} from 'react';
import {Workbook} from '@wieslawsoltes/gridweb';
import {GridWeb} from '@wieslawsoltes/gridweb/react';
export default function SpreadsheetExample(){
 const workbook=useMemo(()=>{const book=new Workbook();book.ActiveWorksheet.GetRange('A1:B2').Values=[['Quantity','Price'],[4,12]];book.ActiveWorksheet.GetCell('C2').Formula='=A2*B2';return book;},[]);
 return <GridWeb workbook={workbook} style={{height:520,display:'block'}} />;
}
