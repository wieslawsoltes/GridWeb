# Compatibility matrix

GridWeb is an independent implementation. Similar names are migration aids, not proof of Excel, Office.js, COM, WPF, WinUI, or Avalonia equivalence.

| Area | Working implementation | Remaining boundaries in the recovered baseline |
|---|---|---|
| Model | Sparse sheets, cells/ranges, names, transactional edits, history, common metadata | Not every Excel object, setting, overload or event |
| Formulas | Safe parser, A1 and cross-sheet references, ranges, arrays, dependency recalculation, spills, LET/LAMBDA | Exhaustive catalog/coercion/locale/precision semantics, R1C1 grammar, 3D refs, unions/intersections, iterative calculation |
| Data | Tables, filters, sort, materialized pivot summaries, goal seek, regression | Native pivot editing/cache/slicers, M/Power Query, DAX/Power Pivot/OLAP, data connections |
| Files | JSON, CSV/TSV, OOXML XLSX cells, supported styles, names, tables, comments, validation, conditional rules and chart parts | Not arbitrary lossless XLSX; no BIFF/XLSB, encryption, VBA execution, full drawing/theme/pivot preservation |
| Charts | Column/bar/line/area/pie/scatter, live ranges, basic axes/legend, pointer layout | Full chart catalog, combo/secondary axes, 3D, arbitrary drawing and exact Office fidelity |
| Rendering | Virtualized Canvas2D, native editor, high DPI, capped scrollbar, freeze panes, visible accessible mirror | Not WebGPU; full IME/screen-reader/RTL/native input certification remains |
| Printing | Normal/page-break/page-layout views, bounded A4/Letter pages, orientation, scale and repeated rows | Printed baseline omits chart/conditional visuals; no printer-identical Excel layout |
| Collaboration | Multiple views of one in-memory workbook | No network coauthoring service or offline network reconciliation in baseline |
| APIs | PascalCase model, native custom element, events/commands, React/worker/host source | Not complete Office.js/COM; React and native execution must be tested independently |
| Native | C# client and WPF/WinUI/Avalonia embedded browser controls and samples | Build and runtime results are separate qualification levels |
| Ecosystem | Source adapters for seven named companion packages | Default offline studio does not bundle them; optional build requires their packages |
| Flow documents | Shared data and rendering primitives suitable for embedding | No native FlowDocument/XAML rich document engine |

Operations are bounded: ordinary dense range work is capped, chart rendering limits its displayed series/data, and page preview is limited. These bounds are explicit errors or documented caps, not full-worksheet allocation promises. Exact current limits are in the source and API documentation.

Formula implementations never evaluate arbitrary JavaScript. Workbook input is untrusted; native RPC is allowlisted. Unsupported features must not be silently advertised as implemented. See SECURITY.md.

This matrix describes the recovered baseline. Subsequent compatibility additions must include targeted tests and update this matrix, without replacing unqualified claims with full-parity claims.
