import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import {inventories} from '../scripts/inventory.mjs';
for(const [name,expected] of Object.entries(await inventories()))test('published availability inventory matches runtime: '+name,async()=>{
  const actual=JSON.parse(await fs.readFile(new URL('../docs/'+name,import.meta.url),'utf8'));
  assert.deepEqual(actual,expected,'Run npm run inventory when public availability changes');
});
