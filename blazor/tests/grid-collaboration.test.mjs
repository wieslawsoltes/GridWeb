import { test } from 'node:test';
import assert from 'node:assert/strict';
import { SpreadsheetSession } from '../src/session.js';

test('activating a worksheet is reflected in full-value binding revisions', async () => {
  const s = new SpreadsheetSession();
  s.Invoke('worksheets.add', { name: 'Second' });
  const revision = s.Revision;
  s.ActivateWorksheet('Second');
  assert.ok(s.Revision > revision);
  assert.equal(s.Info().activeWorksheet, 'Second');
  const same = s.Revision;
  s.ActivateWorksheet('Second');
  assert.equal(s.Revision, same);
  await s.DisposeAsync();
});

test('Blazor workbook sessions synchronize through the actual authenticated HTTP service', async t => {
  const { createCollaborationServer, bearerAuthorization } = await import('../../src/collaboration/server.js');
  const token = 'gridweb-blazor-integration-test-credential';
  const fs = await import('node:fs/promises'), os = await import('node:os'), path = await import('node:path');
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'gridweb-blazor-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const app = await createCollaborationServer({ authorize: bearerAuthorization(token), storageDirectory: directory });
  await new Promise(resolve => app.server.listen(0, '127.0.0.1', resolve));
  const a = new SpreadsheetSession(), b = new SpreadsheetSession();
  t.after(async () => { await a.DisposeAsync(); await b.DisposeAsync(); await app.close(); });
  const options = { url: `http://127.0.0.1:${app.server.address().port}`, room: 'blazor', token, autoSync: false };
  await a.Connect(options, 'create');
  await b.Connect(options, 'join');
  assert.throws(() => a.New(), /Disconnect/);
  a.GetRange('A1').Value = 10;
  b.GetRange('B1').Formula = '=A1*2';
  await Promise.all([a.Sync(), b.Sync()]);
  await a.Sync();
  await b.Sync();
  assert.deepEqual(a.GetRange('A1:B1').Values, [[10, 20]]);
  assert.deepEqual(b.GetRange('A1:B1').Values, [[10, 20]]);
  await a.Disconnect();
  a.New();
  assert.equal(a.GetRange('A1').Value, null);
});
