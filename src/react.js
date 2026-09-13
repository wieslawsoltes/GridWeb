import {createElement,forwardRef,useEffect,useRef,useImperativeHandle,useState,useSyncExternalStore} from 'react';
import {Workbook} from './model.js';
import './controls.js';
/** React owns lifecycle only. The same Workbook is shared with all other hosts. */
export const GridWeb = forwardRef(function GridWeb({workbook,sheet,selection,theme='light',readOnly=false,zoom=1,viewMode='normal',showGridLines=true,onSelectionChange,onCellEdit,onWorkbookChange,style,...props},forwardedRef){
 const ref=useRef(null),callbacks=useRef({});callbacks.current={onSelectionChange,onCellEdit,onWorkbookChange};
 useImperativeHandle(forwardedRef,()=>ref.current,[]);
 useEffect(()=>{const grid=ref.current;const listen=(type,key)=>{const fn=e=>callbacks.current[key]?.(e.detail);grid.addEventListener(type,fn);return()=>grid.removeEventListener(type,fn);};const disposers=[listen('selection-change','onSelectionChange'),listen('cell-edit','onCellEdit'),listen('workbook-change','onWorkbookChange')];return()=>disposers.forEach(d=>d());},[]);
 useEffect(()=>{if(workbook&&ref.current.Workbook!==workbook)ref.current.Workbook=workbook;},[workbook]);
 useEffect(()=>{if(sheet)ref.current.Sheet=sheet;},[sheet,workbook]);
 useEffect(()=>{const grid=ref.current;grid.Theme=theme;grid.ReadOnly=readOnly;grid.Zoom=zoom;grid.ViewMode=viewMode;grid.ShowGridLines=showGridLines;},[theme,readOnly,zoom,viewMode,showGridLines]);
 useEffect(()=>{if(selection&&ref.current.Selection!==selection)ref.current.Selection=selection;},[selection]);
 return createElement('grid-web',{...props,ref,style:{display:'block',height:480,...style}});
});
export function useWorkbook(create=()=>new Workbook()){const[book]=useState(create);useWorkbookRevision(book);return book;}
export function useWorkbookRevision(workbook){return useSyncExternalStore(callback=>{const sub=workbook.Changed.Subscribe(callback);return()=>sub.Dispose();},()=>workbook.Revision,()=>0);}
