# Authenticated workbook collaboration · 0.4.0

## Use the studio

Run `npm run build` and `npm run collaboration:serve` with `GRIDWEB_COLLAB_TOKEN` supplied as an environment secret (at least 24 random characters). Open `http://127.0.0.1:8099/studio`. **Share** accepts the server endpoint, room name, token and explicit Create or Join action. Joining replaces the local view; export a JSON backup before joining an existing room. Two independent browser sessions can edit a shared room.

Share shows connection state, revision, pending changes and participant ranges. Conflicts offer Keep my conflicting edits or Use server conflicting edits. Cell conflict choices retain disjoint edits; structural conflicts choose one whole workbook. Disconnect keeps the local workbook. Replacing the workbook in the studio disconnects the previous room.

Tokens are only in memory and Authorization headers. They are not written to URLs, workbooks, offline caches or the server’s workbook files. The studio uses sessionStorage for per-tab pending edits, which survives a reload, not closing the tab. Reconnect with the token to resume. The normal local JSON save/export remains available.

## Browser library

```js
import {Workbook} from '@wieslawsoltes/gridweb';
import {CollaborationSession} from '@wieslawsoltes/gridweb/collaboration';
const book = new Workbook();
const session = new CollaborationSession(book, {
  url: 'https://your-service.example', room: 'planning',
  token: accessToken, storage: sessionStorage
});
session.Conflicted.Subscribe(conflict => showConflict(conflict));
session.Error.Subscribe(error => showError(error.message));
await session.Connect({mode:'join'});
// Existing GridWeb controls continue to use book; its identity is not replaced.
await session.Sync();
await session.SetPresence(book.ActiveWorksheet.Id, 'C8:D10');
// Call only after the user chooses a conflict strategy:
await session.Resolve('local');
session.Dispose();
```

AutoSync defaults to true. Edits are debounced; authenticated server-sent events notify clients about new revisions and presence. This is a network implementation, not several controls sharing a local object. `autoSync:false` supports explicitly synchronized applications. A saved pending request is retried under its original immutable ID after reload or lost acknowledgment. New edits made during that request are rebased rather than overwritten. Edits made while initially joining abort the join instead of discarding them.

A caller supplying localStorage must assign a distinct `storageKey` per independently edited view. Never share one pending-edit cache between simultaneous clients. Storage quota failures produce a STORAGE error; export work before closing in that case. Tokens are never included in the saved pending record. Browser-side storage is not encryption.

## Node service

The published package includes `gridweb-collaboration`. Environment variables: GRIDWEB_COLLAB_TOKEN (required), HOST (default 127.0.0.1), PORT (8099), GRIDWEB_DATA (./gridweb-data), GRIDWEB_ORIGINS (comma-separated exact origins; loopback studio origins by default). The CLI serves the bundled studio. Protect and back up GRIDWEB_DATA. Do not commit its contents or token files.

For application authentication, use `createCollaborationServer` from `/collaboration/server`. Supply an `authorize(request,{room,action})` function returning `{id,role:'editor'|'viewer'}` or null. Authorize every room independently. The single-token CLI is a development/small-team helper, not per-user enterprise identity. Server writes require editor role; read-only presentation is not an authorization boundary.

The service uses node:http. Terminate TLS at a trusted reverse proxy; browser clients reject non-HTTPS endpoints except loopback. Explicitly allow the frontend’s origin, forward Authorization, disable proxy buffering for SSE and use suitable stream timeouts. There is no wildcard CORS default. Existing SSE credentials are rechecked every 15 seconds. GitHub Pages cannot execute this Node service; deploy it separately. No public service or external account is provisioned automatically.

## Commit and storage contract

Rooms serialize writes and acknowledge only after durable save. File storage writes a temporary file, fsyncs it, atomically renames and syncs the directory where supported. If a save fails after rename, the store reloads persisted state before accepting a retry. Unrecoverable storage errors make the room unavailable rather than accepting ambiguous writes.

Cell content plus its literal flag is one atomic field; style and comment are separate fields. Disjoint edits merge. Different edits to the same field fail the entire commit without partial changes. Workbook structure/metadata changes replace a complete snapshot and increment a structural epoch; stale cell-coordinate operations cannot cross epochs. The latest 128 request receipts support idempotent retry. Unknown requests older than that window are rejected for explicit rebase; they are not blindly replayed.

Server validation runs the actual core in bounded worker threads with memory and time limits, keeping expensive formula validation off the HTTP/authentication thread. Request bodies default to 20 MiB, document JSON to 16 MiB, and core workbooks to their existing cell limits. Commits allow 10,000 field changes, otherwise snapshot replacement is used. Defaults are 32 loaded rooms, 32 queued room operations, 100 event streams and 200 presence entries per room; presence expires after 45 seconds.

## Explicit boundaries

This is not Microsoft’s coauthoring protocol, an Office add-in service or a CRDT text editor. Character-level concurrent edits and Excel collaborative undo are absent. Applying remote content clears stale local undo closures to prevent undo overwriting another user’s edits. Structural conflicts can require choosing a whole snapshot. In-place style subfields are atomic as a complete style. Full-document diffing and validation are bounded but not optimized for very large distributed workbooks.

File persistence is for one process and one writer per directory. Distributed locking, clustering, hosted identity, access-control administration, audit/compliance retention and cross-device offline queues require deployment-specific systems. Native RPC smoke tests and React integration tests do not establish these service guarantees. See test reports for the concrete qualified scenarios.
