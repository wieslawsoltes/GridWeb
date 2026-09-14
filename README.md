# GridWeb

[![CI](https://github.com/wieslawsoltes/GridWeb/actions/workflows/ci.yml/badge.svg)](https://github.com/wieslawsoltes/GridWeb/actions/workflows/ci.yml)
[![Pages](https://github.com/wieslawsoltes/GridWeb/actions/workflows/pages.yml/badge.svg)](https://github.com/wieslawsoltes/GridWeb/actions/workflows/pages.yml)
[![npm](https://img.shields.io/npm/v/@wieslawsoltes/gridweb)](https://www.npmjs.com/package/@wieslawsoltes/gridweb)
[![Downloads](https://img.shields.io/npm/dm/@wieslawsoltes/gridweb)](https://www.npmjs.com/package/@wieslawsoltes/gridweb)
[![License](https://img.shields.io/badge/license-MIT-blue)](LICENSE)

A reusable, DOM-independent spreadsheet engine and virtualized web control with an Excel-inspired workbook studio. Plain JavaScript modules, TypeScript declarations, .NET-style observable models and commands, React integration, worker execution, and WPF/WinUI/Avalonia WebView hosts share the same core.

**Independent implementation, not Microsoft Excel and not a full drop-in Office.js, COM, VBA or native framework runtime.** See [compatibility](docs/compatibility.md) and [verification](docs/verification.md). Badges reflect remote status; a prepared workflow is not evidence of a successful publication.

[Workbook Studio](https://wieslawsoltes.github.io/GridWeb/) · [Releases](https://github.com/wieslawsoltes/GridWeb/releases) · [API](docs/api.md) · [Architecture](docs/architecture.md) · [Publishing](docs/publishing.md)

## Develop

Requires Node 22.16 or newer and TypeScript 5.8.3 or newer for type checks. The runtime has no mandatory npm dependencies.

```sh
npm ci --ignore-scripts --legacy-peer-deps
npm install --global typescript@5.8.3
npm run check
npm run dev
```

`npm run build` creates modular browser files, `dist/GridWeb-standalone.html`, the embedded `dist/GridWeb-host.html`, and the GitHub Pages site. The standalone studio runs without installation or external downloads. For source ES modules use an HTTP server.

## Reuse the engine and control

```js
import { Workbook } from '@wieslawsoltes/gridweb';
import '@wieslawsoltes/gridweb/controls';
const book = new Workbook();
const sheet = book.ActiveWorksheet;
sheet.GetRange('A1:B2').Values = [[10, 20], [30, 40]];
sheet.GetCell('C1').Formula = '=SUM(A1:B2)';
console.log(sheet.GetCell('C1').Value); // 100
const grid = document.createElement('grid-web');
grid.style.cssText = 'display:block;height:520px';
grid.Workbook = book;
document.body.append(grid);
```

The package root is headless. `/controls` registers `<grid-web>`, `/react` supplies a React wrapper and hooks, `/io` handles CSV and XLSX, `/worker` supplies the worker/client protocol, and `/host` exposes allowlisted JSON RPC. ESM and Node's synchronous `require(ESM)` share constructor identity.

## Features

Sparse worksheets, range operations, cross-sheet formulas, dependency invalidation, spill arrays, named expressions, LET/LAMBDA, transactional editing and undo/redo, styles, validation, conditional formatting, tables, sorting/filtering, materialized pivots, goal seek and regression. The control includes native text editing over a virtualized Canvas viewport, resizing, frozen panes, merges, clipboard and autofill, charts, zoom, shared views, themes, and page views.

The five-sheet studio uses the actual core for its workbook, formula bar, ribbon-style commands, inspector, data tools, chart editing, JSON/CSV/XLSX exchange, and print layout. It is not a static mockup.

## React and MVVM

```jsx
import { GridWeb, useWorkbook, useWorkbookRevision } from '@wieslawsoltes/gridweb/react';
import { Workbook } from '@wieslawsoltes/gridweb';
function Editor() {
  const book = useWorkbook(() => new Workbook());
  const revision = useWorkbookRevision(book);
  return <section><p>Revision {revision}</p><GridWeb workbook={book} style={{height:600}} /></section>;
}
```

`EventSource.Subscribe`, `PropertyChanged`, `RelayCommand`, `CanExecute`, and explicit disposal support .NET-style host code. Native projects in `dotnet/` embed the same browser engine; they do not reimplement it in C#. See [native hosts](dotnet/README.md).

## Companion libraries

`integrations/` contains real adapters for Dockyard, RibbonWeb, TreeDataGridWeb, DynamicDataWeb, ReactiveWeb, RBushWeb, and QuikGraphWeb. Install and bundle them with `npm install --prefix integrations`, `npm run build:integrations`, then `npm run build`. The dependency-free studio works without this optional bundle. See [integration status](integrations/README.md).

## Verification and distribution

```sh
npm test
npm run typecheck
npm run build
npm run test:package
python -m pip install -r tests/requirements.txt
python -m playwright install chromium
npm run test:browser
npm run release:pack
```

CI checks Node 22/24, packaged consumers, Chromium interactions, and Windows native project compilation. Native compilation alone does not establish native input, rendering, accessibility or WebView runtime equivalence. Tagged releases publish an exact checksum-verified tarball using `NPM_TOKEN` and npm provenance; release retries verify immutable bytes. GitHub Pages deploys from `main`.

MIT licensed. Original source and third-party references are described in [NOTICE.md](NOTICE.md). No proprietary Microsoft runtime or artwork is bundled.
