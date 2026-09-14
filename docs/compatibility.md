# Compatibility matrix · 0.3.0

GridWeb is independent software. Similar names and passing tests do not establish Excel, Office.js, COM, WPF, WinUI or Avalonia equivalence.

| Area | Implemented | Remaining boundaries |
|---|---|---|
| Workbook | Sparse sheets/cells/ranges, transactions and history, common styles and metadata | Every Excel object, setting, overload and event |
| Calculation | 330 available names, dependency calculation, A1/R1C1 translation, dynamic arrays, LET/LAMBDA, binary lookup, distributions, matrix/complex/financial families | Exhaustive signatures/coercion/locale/precision, all scalar lifting, 3D references, reference unions/intersections, iterative calculation |
| Office-style API | run/load/sync, queued range and worksheet operations, loaded snapshots, format proxies and atomic batch rollback | Full Office.js, tracked objects, COM, add-in host/service APIs; atomic failure behavior is GridWeb-specific |
| Data | Tables, sorting/filtering, materialized pivots, goal seek and regression | Native pivot caches/slicers/editing, M/Power Query, DAX/Power Pivot/OLAP and connectors |
| Files | JSON, CSV/TSV and supported OOXML cells/styles/names/tables/comments/validation/conditional/chart parts | Arbitrary lossless XLSX, native pivot/model preservation, BIFF/XLSB/encryption and macro execution |
| Charts | Column/bar/line/area/pie/scatter, live ranges, pointer positioning and inert SVG export | Complete chart catalog, native combo/3D/drawing fidelity and exact font metrics |
| Printing | Paginated HTML with vector chart fragments, conditional formats, repeated rows/columns, fit targets, breaks, margins and headers/footers | Printer-identical Excel output, physical printer/native WebView qualification and all OOXML print-settings persistence |
| Web control | Virtualized Canvas2D, native editing, clipboard/autofill, merges/freeze, themes and page modes | WebGPU, exhaustive IME/screen-reader/RTL/native input qualification |
| Collaboration | Multiple controls on one in-memory workbook | Network coauthoring and offline network reconciliation |
| React/companions | Adapter source for React and the seven requested companion libraries | Installation, bundling and actual runtime qualification remain separate from baseline browser checks |
| Native | Shared JS host and typed C# client; WPF/WinUI/Avalonia controls and samples compile in CI | Actual native UI execution, OS/device/input/accessibility and printer qualification |
| Flow documents | Embeddable engine/control primitives | No native FlowDocument/XAML runtime |

The uncommitted preservation experiment discussed in the work log is NOT part of this release. Export remains the documented regeneration subset. VBA/embedded executables are not run. No production service is provisioned by GitHub Pages.

See [calculation](functions.md), [Office API](office-api.md), [printing](printing.md), [verification](verification.md) and SECURITY.md. Limits are explicit errors or documented bounds, not promises of allocating an entire worksheet.
