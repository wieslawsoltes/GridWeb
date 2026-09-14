import test from 'node:test';
import assert from 'node:assert/strict';
import {Workbook,paginate} from '../src/core.js';
test('fit accounts for oversized single columns rather than only page count',()=>{
 const s=new Workbook().ActiveWorksheet;s.SetColumnWidth(0,1500);assert.throws(()=>paginate(s,{area:'A1',orientation:'portrait'}),/larger/);const p=paginate(s,{area:'A1',orientation:'portrait',fitToWidthPages:1});assert.ok(p.config.scale<.5);assert.equal(p.pages.length,1);
});
test('oversized repeated title areas fail without discarding or hiding content',()=>{
 const s=new Workbook().ActiveWorksheet;assert.throws(()=>paginate(s,{area:'A1:C100',repeatRows:60}),/larger/);assert.throws(()=>paginate(s,{area:"'Other'!A1"}),/another worksheet/);assert.throws(()=>paginate(s,{orientation:'sideways'}),/orientation/);
});
