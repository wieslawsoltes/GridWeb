import fs from 'node:fs/promises';import path from 'node:path';import os from 'node:os';import assert from 'node:assert/strict';import {execFileSync}from'node:child_process';
const root=path.resolve(import.meta.dirname,'..'),npm=process.platform==='win32'?'npm.cmd':'npm';
const run=(command,args,cwd)=>execFileSync(command,args,{cwd,encoding:'utf8',stdio:['ignore','pipe','pipe'],shell:process.platform==='win32'&&command.endsWith('.cmd')});
const temporary=await fs.mkdtemp(path.join(os.tmpdir(),'gridweb-package-'));
try{
 const supplied=process.argv.indexOf('--tarball'),file=supplied>=0?path.resolve(process.argv[supplied+1]):null;
 const packed=file??path.join(temporary,JSON.parse(run(npm,['pack','--ignore-scripts','--json','--pack-destination',temporary],root))[0].filename);
 await fs.writeFile(path.join(temporary,'package.json'),'{"private":true,"type":"module"}');
 run(npm,['install','--offline','--ignore-scripts','--no-audit','--no-fund','--legacy-peer-deps',packed],temporary);
 const source=`import assert from 'node:assert/strict';import{createRequire}from'node:module';import{Workbook}from'@wieslawsoltes/gridweb';import{GridWebElement}from'@wieslawsoltes/gridweb/controls';import{exportXlsx,importXlsx}from'@wieslawsoltes/gridweb/io';const require=createRequire(import.meta.url);assert.equal(require('@wieslawsoltes/gridweb').Workbook,Workbook);assert.equal(typeof GridWebElement,'function');const b=new Workbook();b.ActiveWorksheet.GetRange('A1:B2').Values=[[1,2],[3,4]];b.ActiveWorksheet.GetCell('C1').Formula='SUM(A1:B2)';assert.equal(b.ActiveWorksheet.GetCell('C1').Value,10);const other=await importXlsx(exportXlsx(b));assert.equal(other.workbook.ActiveWorksheet.GetCell('C1').Value,10);console.log('Installed ESM, CommonJS constructor identity, headless controls import and XLSX consumers passed.');`;
 await fs.writeFile(path.join(temporary,'consumer.mjs'),source);console.log(run(process.execPath,['consumer.mjs'],temporary).trim());
 const packageRoot=path.join(temporary,'node_modules/@wieslawsoltes/gridweb');for(const name of ['types/index.d.ts','types/controls.d.ts','types/react.d.ts','types/io.d.ts','dist/GridWeb-standalone.html','dist/GridWeb-host.html','docs/compatibility.md','LICENSE','README.md'])assert((await fs.stat(path.join(packageRoot,name))).isFile(),name);
 const typeSource=await fs.readFile(path.join(root,'tests/types.ts'),'utf8');await fs.writeFile(path.join(temporary,'types.ts'),typeSource);
 try{run('tsc',['--version'],temporary);console.log(run('tsc',['--noEmit','--strict','--lib','es2022,dom','--module','nodenext','--moduleResolution','nodenext','types.ts'],temporary)||'Installed strict TypeScript consumer passed.');}catch(e){if(e.code==='ENOENT')throw new Error('TypeScript is required for package qualification; install tsc 5.8.3 or newer');throw e;}
 console.log('Verified package:',path.basename(packed));
}finally{await fs.rm(temporary,{recursive:true,force:true});}
