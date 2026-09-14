# GridWeb

[![CI](https://github.com/wieslawsoltes/GridWeb/actions/workflows/ci.yml/badge.svg)](https://github.com/wieslawsoltes/GridWeb/actions/workflows/ci.yml)
[![Pages](https://github.com/wieslawsoltes/GridWeb/actions/workflows/pages.yml/badge.svg)](https://github.com/wieslawsoltes/GridWeb/actions/workflows/pages.yml)
[![npm](https://img.shields.io/npm/v/@wieslawsoltes/gridweb)](https://www.npmjs.com/package/@wieslawsoltes/gridweb)
[![Downloads](https://img.shields.io/npm/dm/@wieslawsoltes/gridweb)](https://www.npmjs.com/package/@wieslawsoltes/gridweb)
[![License](https://img.shields.io/badge/license-MIT-blue)](LICENSE)

A reusable headless spreadsheet engine, virtualized Web Component and Excel-inspired workbook studio. JavaScript, TypeScript, .NET-style MVVM, React, workers and native WebView hosts share one core. Office-style batched APIs, authenticated coauthoring, vector printing and managed pivots are optional reusable surfaces.

**Independent software, not Microsoft Excel or full Excel/Office.js/COM/VBA parity.** [Exact compatibility boundaries](docs/compatibility.md).

**[Open Workbook Studio](https://wieslawsoltes.github.io/GridWeb/)** · [Releases and full source](https://github.com/wieslawsoltes/GridWeb/releases) · [API](docs/api.md) · [Architecture](docs/architecture.md)

## Install

```sh
npm install @wieslawsoltes/gridweb
```

```js
import {Workbook} from '@wieslawsoltes/gridweb';
import '@wieslawsoltes/gridweb/controls';
const book = new Workbook();
book.ActiveWorksheet.GetRange('A1:B2').Values = [[10,20],[30,40]];
book.ActiveWorksheet.GetCell('C1').Formula = '=SUM(A1:B2)';
const grid = document.createElement('grid-web');
grid.style.cssText = 'display:block;height:520px';
grid.Workbook = book;
document.body.append(grid);
```

The root is DOM-independent. `/controls` registers `<grid-web>`; `/browser` includes core/control without duplicate model constructors. `/react`, `/worker`, `/host`, `/io`, `/office`, `/printing` and `/collaboration` expose the other surfaces. The Node service is `/collaboration/server`. Node 22.16+ supports synchronous require(ESM). React is an optional peer; the core has no mandatory runtime npm dependencies.

## Engine and editor

Sparse workbooks/ranges, 330 available formula names, dependencies, arrays and LET/LAMBDA, binary lookups, distributions, matrices, complex/dated finance, A1/R1C1 translation, transactional editing/history, formatting, validation, conditional rules, tables, sorting/filtering, goal seek and regression. The Canvas editor supports native text editing, keyboard/pointer selections, clipboard/fill, resizing, merges, frozen panes, chart manipulation, zoom, themes and multiple views. The five-sheet studio uses the same engine, not separate mock data.

[Calculation contracts](docs/functions.md) · [Core API](docs/api.md) · [React example](examples/react.jsx)

## Managed pivots

**Data / Insert → Pivot reports** creates a persistent report. Edit pivot reconfigures fields/filters; Refresh pivots updates it from source values; Show details opens underlying source records. Twelve aggregations, multiple measures, row/column axes, totals, history and JSON persistence are implemented.

```js
const source = book.Worksheets.Add('Source');
source.GetRange('A1:B3').Values = [['Region','Revenue'],['North',10],['North',20]];
const output = book.Worksheets.Add('Report');
const pivot = book.PivotTables.Add('Sales', source.UsedRange, output.GetRange('A1'), {
  rows:['Region'], values:[{column:'Revenue',aggregate:'sum'}]
});
pivot.Refresh();
```

Supported row-axis layouts export as actual native XLSX pivot definitions, cache definitions, typed shared items and cache records. Unsupported native layouts require explicit flattening; JSON retains all managed definitions. Exports refresh a snapshot without changing the open workbook. [Pivot API, native subset and C# usage](docs/pivots.md).

## Office-style batching and print

`createExcelApi(book)` from `/office` provides queued `Excel.run`, `load` and `context.sync` for implemented worksheet/range/format operations. Loaded values are snapshots; failed batches roll back atomically. View → Office API example exercises it. [API and differences](docs/office-api.md).

`/printing` exports inert SVG charts and paginated HTML including chart fragments, conditional styling/data bars, fitted A4/Letter/A3/Legal pages, repeated titles, breaks, margins and header/footer tokens. Page Layout → Advanced print and Chart SVG are working studio commands. [Printing contract](docs/printing.md).

## Collaboration

**Share** creates or joins an authenticated server room, merges disjoint content/style/comment edits, presents conflicts and reports participant selections. Pending edits survive reload in this tab’s session storage, without the token; reconnect to resume.

```sh
# Supply a securely generated GRIDWEB_COLLAB_TOKEN of at least 24 characters.
npm run build
npm run collaboration:serve
# Installed package: gridweb-collaboration
```

The Node CLI defaults to loopback port 8099 and serves `/studio`. Configure HTTPS, exact allowed origins, persistent storage and application authorization for remote access. The static Pages site does not run the service. [Deployment, API, recovery and limits](docs/collaboration.md).

## Framework and native hosts

PascalCase APIs, PropertyChanged, collection notifications, RelayCommand and disposable bindings support .NET-style JavaScript. `dotnet/` contains the typed C# client and WPF/WinUI/Avalonia controls and samples embedding the same JavaScript engine. Native Windows WebView smoke tests exercise real RPC including pivot creation, refresh and drill-down. Exhaustive native input/accessibility and cross-platform qualification remain separate.

`integrations/` contains Dockyard, RibbonWeb, TreeDataGridWeb, DynamicDataWeb, ReactiveWeb, RBushWeb and QuikGraphWeb adapters. CI runs actual package tests with React 18/19. The dependency-free studio works without the optional companion bundle. [Bundle instructions](integrations/README.md) · [Native hosts](dotnet/README.md).

## Build, test and distribute

```sh
npm ci --ignore-scripts --legacy-peer-deps
npm install --global typescript@5.8.3
npm run check
python -m pip install -r tests/requirements.txt
python -m playwright install chromium
npm run test:browser
npm run dev
```

Build produces modular browser files, a standalone studio, embedded host HTML and the Pages site. `npm run release:pack` produces tarball/source/browser archives. CI checks engine/types, installed package consumers, actual browser editing/coauthoring/pivot flows, native runtimes, companion packages and independent Open XML pivot schema validation.

Successful main CI triggers versioned npm/GitHub publication with NPM_TOKEN, provenance and exact-byte verification. Pages deploys independently. Source archives contain tracked source rather than private untracked room files. [Publishing](docs/publishing.md) · [Automatic releases](docs/automatic-releases.md) · [Security](SECURITY.md).

Full parity remains unfinished: exhaustive formulas and Office semantics, arbitrary lossless XLSX, complete native pivot/drawing/data-model behavior, VBA, Power Query/DAX, enterprise coauthoring and printer fidelity. No proprietary runtime/artwork/fonts are bundled. MIT licensed; [notices](NOTICE.md).
