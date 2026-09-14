import { execFileSync } from 'node:child_process';
import { readFileSync, writeFileSync } from 'node:fs';
// Immutable shared source, embedded into our own assembly and _content/GridWeb.Blazor.
await import('./runtime-source/blazor/build-consumer.mjs');
execFileSync(process.platform === 'win32' ? 'npm.cmd' : 'npm', ['exec','--yes','--package=esbuild@0.28.2','--','esbuild','src/worker.js','--bundle','--format=esm','--platform=browser','--target=es2022','--outfile=blazor/src/wwwroot/worker.js','--legal-comments=eof','--external:node:zlib'], { stdio: 'inherit' });
// Keep the shared actual-package smoke test and add real keyboard editing for this control.
const path = 'blazor/tests/smoke.py';
let smoke = readFileSync(path, 'utf8');
smoke = smoke.replaceAll("                page.locator('#action').click()", `                page.locator('#begin-edit').click()
                editor = page.locator('grid-web textarea.editor').first
                expect(editor).to_be_visible()
                editor.fill('123')
                editor.press('Enter')
                expect(page.locator('#edit-result')).to_have_text('Edited A6')
                page.locator('#action').click()`);
writeFileSync(path, smoke);
