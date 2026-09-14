# Architecture

`address.js`, `errors.js`, `events.js`: bounded coordinates, values/errors and disposable observable/command primitives.

`parser.js`, `functions.js`, `calculation.js`: tokenized formula AST, function registry, dependency edges, range invalidation, cached results and spill arrays. Evaluation never calls eval or Function.

`model.js`: sparse workbook/worksheet/cell/range objects, change transactions, undo/redo, metadata and model validation. A shared workbook is the source of truth for every view and host.

`data.js`: reusable data analysis, CSV and pivot summary operations. `format.js` maps values/style masks to displayed text. `layout.js` supplies sparse axis geometry and pagination. `charts.js` draws charts from live ranges.

`controls.js`: a native custom element combining a bounded Canvas viewport with native editing controls and a visible-cell accessibility mirror. Logical address space is distinct from the physical scrollbar extent. Virtualization does not allocate all worksheet cells.

`io.js`, `zip.js`, `xml.js`: bounded package/XML handling and supported OOXML exchange. `host.js` and `worker.js` expose explicit data operations; host transports do not permit arbitrary code or object-path execution. `react.js` supplies framework lifecycle/property synchronization only.

`demo/`: a five-sheet studio built on public and internal engine primitives. `integrations/`: optional real companion-package adapters, independently bundled. `dotnet/`: C# client and native WebView hosts; calculation remains in JavaScript.

Build scripts generate browser/site distributions from readable source. Release scripts qualify a tarball, checksum it, preserve immutable release identities and verify public registry bytes. Tests distinguish core behavior, packaged consumers, browser interactions, native builds and native runtime qualification.
