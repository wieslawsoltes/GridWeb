# GridWeb

[![CI](https://github.com/wieslawsoltes/GridWeb/actions/workflows/ci.yml/badge.svg)](https://github.com/wieslawsoltes/GridWeb/actions/workflows/ci.yml)
[![Pages](https://github.com/wieslawsoltes/GridWeb/actions/workflows/pages.yml/badge.svg)](https://github.com/wieslawsoltes/GridWeb/actions/workflows/pages.yml)
[![npm](https://img.shields.io/npm/v/@wieslawsoltes/gridweb)](https://www.npmjs.com/package/@wieslawsoltes/gridweb)
[![Downloads](https://img.shields.io/npm/dm/@wieslawsoltes/gridweb)](https://www.npmjs.com/package/@wieslawsoltes/gridweb)
[![License](https://img.shields.io/badge/license-MIT-blue)](LICENSE)

Reusable headless spreadsheet engine, virtualized Web Component and Excel-inspired workbook studio. Plain JavaScript, TypeScript declarations, .NET-style MVVM and events, React adapter, worker protocol, native WebView hosts, Office-style batched APIs and authenticated network collaboration share the same core.

**Independent implementation, not Microsoft Excel or complete Office.js/COM/VBA compatibility.** See the [current compatibility matrix](docs/compatibility.md). No proprietary runtime, artwork or fonts are bundled.

**[Workbook Studio](https://wieslawsoltes.github.io/GridWeb/)** · [Releases](https://github.com/wieslawsoltes/GridWeb/releases) · [Core API](docs/api.md) · [Calculation](docs/functions.md) · [Office-style API](docs/office-api.md) · [Printing and SVG](docs/printing.md) · [Collaboration](docs/collaboration.md)

## Install and use

```sh
npm install @wieslawsoltes/gridweb
```

```js
import { Workbook } from '@wieslawsoltes/gridweb';
import '@wieslawsoltes/gridweb/controls';
const book = new Workbook();
const sheet = book.ActiveWorksheet;
sheet.GetRange('A1:B2').Values = [[10,20],[30,40]];
sheet.GetCell('C1').Formula = '=SUM(A1:B2)';
const grid = document.createElement('grid-web');
grid.style.cssText = 'display:block;height:520px';
grid.Workbook = book;
document.body.append(grid);
```

The root import is DOM-independent. `/controls` registers `<grid-web>`; `/browser` combines the browser control and core without duplicating model constructors. `/react`, `/worker`, `/host`, `/io`, `/office`, `/printing` and `/collaboration` provide optional entry points. The Node service is `/collaboration/server`. Node 22.16+ supports synchronous require(ESM) with shared constructor identity. React is an optional peer; the core has no mandatory runtime npm dependencies.

## Office-style migration

```js
import { createExcelApi } from '@wieslawsoltes/gridweb/office';
const Excel = createExcelApi(book);
await Excel.run(async context => {
  const range = context.workbook.worksheets.getActiveWorksheet().getRange('D1');
  range.formulas = [['=C1*2']];
  range.format.font.bold = true;
  range.load('values');
  await context.sync();
  console.log(range.values); // [[200]]
});
```

The adapter implements queued worksheet/range/format operations, loaded snapshots, null-cell skipping and atomic batch rollback. It is a documented subset, not a replacement for all Office add-in services. View → Office API example uses it directly.

## Engine and studio

Sparse worksheets and ranges; 330 available formula names including arrays, LET/LAMBDA, binary lookups, statistical distributions, matrices, complex and dated financial operations; dependency recalculation; A1/R1C1 conversion; transactional editing/history; styles, validation, conditional formatting, tables, sorting/filtering, materialized pivots, goal seek and regression.

The virtualized Canvas control provides native editing, pointer/keyboard selections, clipboard/fill, merges, frozen panes, chart manipulation, zoom, themes and multiple views. The five-sheet studio uses this engine for editing and calculations. Normal, page-layout and page-break views are available.

Print output includes inert SVG charts and fragments across pages, conditional styling and data bars, fixed cell geometry, fitted A4/Letter/A3/Legal pages, repeated titles, breaks, margins and header/footer tokens. Page Layout → Advanced print opens a printable preview. Chart SVG exports a vector chart. Browser typography is not printer-identical Excel layout.

## Frameworks and native hosts

React wrapper and hooks are in `/react`. PascalCase APIs, PropertyChanged, collection notifications, RelayCommand and disposable bindings support .NET-style JavaScript. `dotnet/` contains a typed client and WPF, WinUI and Avalonia controls and samples embedding the same JavaScript engine. CI runs native Windows WebView startup/RPC smoke tests, not just compilation. Exhaustive native input/accessibility and cross-platform qualification remain separate.

`integrations/` contains adapters for Dockyard, RibbonWeb, TreeDataGridWeb, DynamicDataWeb, ReactiveWeb, RBushWeb and QuikGraphWeb. CI exercises these actual packages with React 18/19. The default studio works offline without that bundle. See [companion instructions](integrations/README.md).

## Network collaboration

**Share** creates or joins an authenticated GridWeb room. Disjoint content/style/comment edits merge; conflicting edits require a visible choice. Pending edits survive reload in this tab’s session storage without the token. Reconnect after reload to resume.

```sh
# Set GRIDWEB_COLLAB_TOKEN to a securely generated secret of at least 24 characters.
npm run build
npm run collaboration:serve
# Installed package: gridweb-collaboration
```

The server defaults to loopback port 8099 and serves `/studio`. Configure explicit allowed origins and HTTPS for remote access. The static Pages site does not host this Node service. [API, deployment, recovery and limits](docs/collaboration.md).

## Develop and verify

```sh
npm ci --ignore-scripts --legacy-peer-deps
npm install --global typescript@5.8.3
npm run check
python -m pip install -r tests/requirements.txt
python -m playwright install chromium
npm run test:browser
npm run dev
```

`npm run build` creates the modular distribution, self-contained studio, embedded host page and Pages site. `npm run release:pack` creates tarball/source/browser archives. Tests cover the engine, types, installed consumers and actual Chromium interactions; framework and native tests run in CI. Source release archives use tracked files, excluding private runtime room data.

Successful main CI triggers checksum-verified versioned npm/GitHub publication using NPM_TOKEN and provenance. Pages deploys independently. [Release instructions](docs/publishing.md) · [Automatic releases](docs/automatic-releases.md) · [Verification](docs/verification.md) · [Security](SECURITY.md)

Full Excel parity remains unfinished: exhaustive formulas and Office semantics, lossless arbitrary XLSX, complete native pivots/drawings/data models, VBA, Power Query/DAX, enterprise coauthoring and printer fidelity. Review XLSX import warnings. MIT licensed; see [notices](NOTICE.md).
