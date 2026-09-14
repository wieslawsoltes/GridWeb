/** Public npm integrity checks. No credentials are read or persisted by this script. */
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { createHash } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { pathToFileURL } from 'node:url';

const root = path.resolve(import.meta.dirname, '..');
export const sha512 = bytes => 'sha512-' + createHash('sha512').update(bytes).digest('base64');
export const sha256 = bytes => createHash('sha256').update(bytes).digest('hex');
export function validateRelease(tag, version, expectedSha = '', actualSha = '') {
  if (!/^v\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?$/.test(tag)) throw new Error('Invalid release tag');
  if (tag !== 'v' + version) throw new Error('Tag/package version mismatch');
  if (expectedSha && (!/^[a-f0-9]{40}$/.test(expectedSha) || expectedSha !== actualSha)) throw new Error('Release commit mismatch');
}
export function validateMetadata(metadata, pkg, integrity, provenance = false) {
  if (metadata?.name !== pkg.name || metadata?.version !== pkg.version) throw new Error('Registry identity mismatch');
  if (metadata?.dist?.integrity !== integrity) throw new Error('Immutable registry version has different bytes');
  const url = new URL(metadata.dist.tarball);
  if (url.protocol !== 'https:' || url.hostname !== 'registry.npmjs.org' || url.username || url.password) throw new Error('Unexpected tarball origin');
  if (provenance && !metadata.dist.attestations?.url) throw new Error('Registry has no provenance attestation descriptor');
  return url.href;
}
export function parseChecksumManifest(text) {
  const result = new Map();
  for (const line of text.trim().split(/\r?\n/)) {
    const m = /^([a-f0-9]{64})  ([A-Za-z0-9][A-Za-z0-9._-]*)$/.exec(line);
    if (!m || result.has(m[2])) throw new Error('Invalid or duplicate checksum entry');
    result.set(m[2], m[1]);
  }
  return result;
}
async function readBounded(response, limit) {
  if (+response.headers.get('content-length') > limit) throw new Error('Registry response exceeds size limit');
  const chunks = []; let count = 0;
  for await (const chunk of response.body) {
    count += chunk.length;
    if (count > limit) { throw new Error('Registry response exceeds size limit'); }
    chunks.push(Buffer.from(chunk));
  }
  return Buffer.concat(chunks);
}
async function metadataFor(pkg, allowMissing) {
  const url = `https://registry.npmjs.org/${encodeURIComponent(pkg.name)}/${encodeURIComponent(pkg.version)}`;
  const response = await fetch(url, {redirect: 'error', signal: AbortSignal.timeout(30000)});
  if (response.status === 404 && allowMissing) return null;
  if (!response.ok) throw new Error(`Registry metadata HTTP ${response.status}`);
  return JSON.parse((await readBounded(response, 2 * 1024 * 1024)).toString('utf8'));
}
const output = async (key, value) => {
  if (process.env.GITHUB_OUTPUT) await fs.appendFile(process.env.GITHUB_OUTPUT, `${key}=${value}\n`);
  console.log(`${key}=${value}`);
};
async function main() {
  const [action, argument] = process.argv.slice(2);
  const pkg = JSON.parse(await fs.readFile(path.join(root, 'package.json'), 'utf8'));
  if (action === 'prepare') {
    const directory = path.resolve(argument ?? 'artifacts');
    const sums = parseChecksumManifest(await fs.readFile(path.join(directory, 'SHA256SUMS.txt'), 'utf8'));
    const tarballs = [...sums.keys()].filter(n => n.endsWith('.tgz'));
    if (tarballs.length !== 1) throw new Error('Expected exactly one tarball in checksum manifest');
    for (const [name, expected] of sums) {
      if (sha256(await fs.readFile(path.join(directory, name))) !== expected) throw new Error(`Release checksum mismatch: ${name}`);
    }
    await output('tarball', path.join(directory, tarballs[0])); return;
  }
  if (!['check', 'verify'].includes(action) || !argument) throw new Error('Usage: npm-registry.mjs prepare <dir> | check <tarball> | verify <tarball>');
  const file = path.resolve(argument), bytes = await fs.readFile(file), integrity = sha512(bytes);
  if (action === 'check') {
    const metadata = await metadataFor(pkg, true);
    if (metadata) validateMetadata(metadata, pkg, integrity);
    await output('exists', metadata ? 'true' : 'false'); return;
  }
  let metadata;
  for (let attempt = 0; attempt < 12; attempt++) {
    metadata = await metadataFor(pkg, true);
    if (metadata?.dist?.attestations?.url) break;
    if (attempt < 11) await new Promise(resolve => setTimeout(resolve, 5000));
  }
  const tarballUrl = validateMetadata(metadata, pkg, integrity, true);
  const response = await fetch(tarballUrl, {redirect: 'error', signal: AbortSignal.timeout(30000)});
  if (!response.ok) throw new Error(`Registry tarball HTTP ${response.status}`);
  const publicBytes = await readBounded(response, 64 * 1024 * 1024);
  if (sha512(publicBytes) !== integrity) throw new Error('Downloaded registry bytes differ from release tarball');
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'gridweb-registry-'));
  try {
    const downloaded = path.join(directory, path.basename(file));
    await fs.writeFile(downloaded, publicBytes);
    execFileSync(process.execPath, ['scripts/package-test.mjs', '--tarball', downloaded], {cwd: root, stdio: 'inherit'});
  } finally { await fs.rm(directory, {recursive: true, force: true}); }
  console.log(`Verified public ${pkg.name}@${pkg.version}: exact bytes, installed consumers, provenance descriptor.`);
  console.log('Descriptor presence is not an independent cryptographic verification of the attestation.');
}
if (process.argv[1] && import.meta.url === pathToFileURL(path.resolve(process.argv[1])).href) {
  main().catch(error => { console.error(error.message); process.exitCode = 1; });
}
