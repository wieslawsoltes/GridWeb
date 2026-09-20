/** Root source archives record gitlinks; they never sweep submodule worktrees. */
import fs from 'node:fs/promises';
import path from 'node:path';
import {execFileSync} from 'node:child_process';

export async function collectTrackedSource(root) {
  const index = execFileSync('git', ['ls-files', '--stage', '-z'], {
    cwd: root, encoding: 'utf8', maxBuffer: 64 * 1024 * 1024
  }).split('\0').filter(Boolean);
  if (!index.length) throw new Error('Source releases require tracked files');
  const source = {}, gitlinks = [], seen = new Set();
  for (const entry of index) {
    const match = /^(\d{6}) ([0-9a-f]{40,64}) ([0-3])\t([\s\S]+)$/.exec(entry);
    if (!match) throw new Error('Invalid Git index entry');
    const [, mode, commit, stage, name] = match;
    if (stage !== '0' || seen.has(name)) throw new Error('Unmerged or duplicate source entry: ' + name);
    seen.add(name);
    const parts = name.split('/');
    if (parts.some(p => !p || p === '.' || p === '..' || /[\\:\0]/.test(p))) throw new Error('Unsafe source path');
    if (mode === '160000') { gitlinks.push({path: name, commit}); continue; }
    if (!['100644', '100755'].includes(mode)) throw new Error('Only regular source files or gitlinks may be archived: ' + name);
    // Do not follow a symlink in any parent component, even to another local directory.
    let file = root;
    for (let i = 0; i < parts.length; i++) {
      file = path.join(file, parts[i]);
      const info = await fs.lstat(file);
      if (info.isSymbolicLink() || (i === parts.length - 1 ? !info.isFile() : !info.isDirectory())) throw new Error('Unsafe source filesystem entry: ' + name);
    }
    source['GridWeb/' + name] = await fs.readFile(file);
  }
  if (gitlinks.length) {
    const manifest = 'GridWeb/SOURCE-GITLINKS.json';
    if (Object.hasOwn(source, manifest)) throw new Error('Reserved source manifest path');
    source[manifest] = Buffer.from(JSON.stringify({
      version: 1, contentsIncluded: false, submodules: gitlinks,
      note: 'This archive contains root source and built browser assets. Submodule contents are not embedded. Retrieve the exact listed commits using the URLs in .gitmodules when building those components.'
    }, null, 2) + '\n');
  }
  return source;
}
