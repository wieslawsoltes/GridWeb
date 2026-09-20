# Compatibility matrix · 0.7.0

GridWeb is independent software. Similar names and passing tests do not establish Excel, Office.js, COM, WPF, WinUI or Avalonia equivalence.

| Area | Implemented | Remaining boundaries |
|---|---|---|
| Workbook | Sparse sheets/cells/ranges, indexed insertion/reordering with reference repair, transactions/history, styles/metadata and managed pivot definitions | Every Excel object, setting, overload and event |
| Calculation | 356 available names, dependency calculation, A1/R1C1 translation, dynamic arrays, callable/recursive LET/LAMBDA, database criteria, reference-aware AGGREGATE, binary lookup, distributions, matrix/complex/financial families; bounded 3-D references, union/intersection/dynamic-range operators, reference INDEX, sheet/area queries and A-suffixed statistics | Exhaustive signatures/coercion/locale/precision, all scalar lifting, reference-valued LET/LAMBDA bindings, grouped 3-D structural transforms, iterative calculation |
| Office-style API | run/load/sync, queued range and worksheet operations including writable position, format/transpose/skip-blank copying, loaded snapshots, format proxies and atomic batch rollback | Full Office.js, tracked objects, COM, add-in host/service APIs; atomic failure behavior is GridWeb-specific |
| Data | Tables, sorting/filtering, managed pivots with 12 aggregations, row/column grouping, measures, filters, totals, refresh and drill-down; goal seek/regression | Hierarchical subtotals, slicers, calculated pivot fields/items, date grouping, M/Power Query, DAX/Power Pivot/OLAP and connectors |
| Files | JSON definitions; supported OOXML cells/styles/names/tables/comments/validation/charts plus native range-backed row-pivot definitions/cache records | Arbitrary lossless XLSX, native column-axis pivots, advanced cache/model preservation, BIFF/XLSB/encryption and macro execution |
| Charts | Column/bar/line/area/pie/scatter, live ranges, pointer positioning and inert SVG export | Complete chart catalog, native combo/3D/drawing fidelity and exact font metrics |
| Printing | Paginated HTML with vector chart fragments, conditional formats, repeated rows/columns, fit targets, breaks, margins and headers/footers | Printer-identical Excel output, physical printer/native WebView qualification and all OOXML print-settings persistence |
| Web control | Virtualized Canvas2D, native editing, clipboard/autofill, paste-special/series/special-cell/worksheet-move dialogs, merges/freeze, themes, page modes and pivot field dialogs | WebGPU, exhaustive IME/screen-reader/RTL/native input qualification |
| Collaboration | Authenticated HTTP/SSE server, durable ordered commits, disjoint-field merge, explicit conflicts, tab-reload pending edits, presence and pivot definitions | Microsoft protocol equivalence, character-level merging, shared collaborative undo, enterprise identity/compliance, horizontal clustering |
| React/companions | React 18/19 StrictMode and seven installed companion libraries pass runtime tests | Exhaustive framework lifecycle and production-scale qualification |
| Native | Real WPF/WinUI/Avalonia WebView startup/RPC on Windows, typed C# pivot and worksheet operations; explicit Avalonia adapter readiness | Cross-platform Avalonia, exhaustive native input/accessibility, devices and printers |
| Flow documents | Embeddable engine/control primitives | No native FlowDocument/XAML runtime |

Native pivot export is a defined subset, not arbitrary original-part preservation. Unsupported managed native layouts require explicit flattening; JSON retains managed definitions. VBA/embedded executables are not run. GitHub Pages is a static site; deploy the authenticated Node collaboration service separately.

See [calculation](functions.md), [Office API](office-api.md), [printing](printing.md), [managed pivots](pivots.md), [collaboration](collaboration.md), [verification](verification.md) and SECURITY.md. Limits are explicit errors or documented bounds, not promises of allocating an entire worksheet.

The [source-level feature audit](excel-feature-audit.md) records the broader remaining work and evidence. See [editing contracts](editing.md) for unsupported paste combinations and [current verification](verification-0.7.md) for results on this increment. The separately pinned Blazor runtime was not advanced.
