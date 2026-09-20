import {collectTrackedSource} from './source-archive.mjs';
import fs from'node:fs/promises';import path from'node:path';import{execFileSync}from'node:child_process';import{createHash}from'node:crypto';import{writeZip}from'../src/zip.js';
const root=path.resolve(import.meta.dirname,'..'),out=path.join(root,'artifacts'),pkg=JSON.parse(await fs.readFile(path.join(root,'package.json'))),npm=process.platform==='win32'?'npm.cmd':'npm';
await fs.mkdir(out,{recursive:true});for(const n of await fs.readdir(out))if(/\.(?:zip|tgz)$/.test(n)||n==='SHA256SUMS.txt'||n==='release-manifest.json')await fs.rm(path.join(out,n));
const packed=JSON.parse(execFileSync(npm,['pack','--ignore-scripts','--json','--pack-destination',out],{cwd:root,encoding:'utf8',shell:process.platform==='win32'}))[0];
const ignore=new Set(['.git','node_modules','bin','obj','artifacts','site','__pycache__']);
async function walk(dir,prefix='',skipDist=false){const entries={};for(const item of await fs.readdir(dir,{withFileTypes:true})){if(ignore.has(item.name)||(skipDist&&item.name==='dist')||item.name==='patch-current.py'||item.name.endsWith('.log'))continue;const name=prefix+item.name,p=path.join(dir,item.name);if(item.isDirectory())Object.assign(entries,await walk(p,name+'/',skipDist));else entries[name]=await fs.readFile(p);}return entries;}
// Archive only tracked source; never sweep private runtime files into a public release.
const source=await collectTrackedSource(root);
Object.assign(source,await walk(path.join(root,'dist'),'GridWeb/dist/'));await fs.writeFile(path.join(out,`GridWeb-${pkg.version}-source.zip`),writeZip(source));
const browser=await walk(path.join(root,'dist'));browser['LICENSE']=await fs.readFile(path.join(root,'LICENSE'));browser['NOTICE.md']=await fs.readFile(path.join(root,'NOTICE.md'));await fs.writeFile(path.join(out,`GridWeb-${pkg.version}-browser.zip`),writeZip(browser));
const files=[];for(const name of(await fs.readdir(out)).sort()){if(name==='SHA256SUMS.txt'||name==='release-manifest.json')continue;const bytes=await fs.readFile(path.join(out,name));files.push({name,bytes:bytes.length,sha256:createHash('sha256').update(bytes).digest('hex')});}
await fs.writeFile(path.join(out,'SHA256SUMS.txt'),files.map(f=>`${f.sha256}  ${f.name}`).join('\n')+'\n');
await fs.writeFile(path.join(out,'release-manifest.json'),JSON.stringify({name:pkg.name,version:pkg.version,tarball:packed.filename,files,status:'local artifacts; registry publication requires a separately verified workflow'},null,2));console.log(files.map(f=>`${f.name}: ${f.bytes.toLocaleString()} bytes`).join('\n'));
