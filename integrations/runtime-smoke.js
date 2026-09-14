import React,{StrictMode,createElement,useState} from 'react';
import {createRoot} from 'react-dom/client';
import {Workbook} from '../src/core.js';
import {GridWeb,useWorkbookRevision} from '../src/react.js';
import {attachCompanions} from './index.js';

const makeBook=()=>{const b=new Workbook();b.ActiveWorksheet.GetRange('A1:C3').Values=[['Region','Units','Price'],['North',2,5],['South',3,4]];b.ActiveWorksheet.GetRange('D2:D3').Formulas=[['=B2*C2'],['=B3*C3']];return b;};
const book=makeBook(),grid=document.createElement('grid-web');grid.Workbook=book;grid.style.cssText='height:100%;display:block';document.querySelector('#document').append(grid);
book.ActiveWorksheet.AddChart('A1:D3',{row:5,column:1,width:400,height:220});
const count={},commands={};for(const name of ['undo','redo','open','save','xlsx','print','bold','italic','table','conditional','pivot','filter','sort','validation','goalSeek','functions'])commands[name]=()=>{count[name]=(count[name]??0)+1;if(name==='bold')grid.SelectionRange.SetStyle({font:{bold:true}});if(name==='undo')book.Undo();if(name==='redo')book.Redo();};
const companions=await attachCompanions({workbook:book,grid,commands,documentPane:document.querySelector('#document'),inspector:document.querySelector('#inspector'),workspace:document.querySelector('#workspace'),ribbonHost:document.querySelector('#companion-ribbon')});
const reactBook=makeBook(),root=createRoot(document.querySelector('#react-root')),ref=React.createRef();let selected='',lastEdit=null,changes=0,configure;
function ReactEditor(){const revision=useWorkbookRevision(reactBook),[props,setProps]=useState({theme:'light',readOnly:false,selection:'A1',viewMode:'normal'});configure=setProps;return createElement('section',null,createElement('output',{'data-revision':revision},'Revision '+revision),createElement(GridWeb,{...props,ref,workbook:reactBook,style:{height:300},onSelectionChange:e=>selected=e.address,onCellEdit:e=>lastEdit=e,onWorkbookChange:()=>changes++}));}
root.render(createElement(StrictMode,null,createElement(ReactEditor)));
window.integrationHarness={book,grid,count,companions,reactBook,ref,reactVersion:React.version,get selected(){return selected;},get lastEdit(){return lastEdit;},get changes(){return changes;},configure:options=>configure(options),unmount:()=>root.unmount(),dispose:()=>companions.Dispose()};
