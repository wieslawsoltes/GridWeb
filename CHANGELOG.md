# Changelog

## 0.7.0

Implement bounded 3-D reference values, unions/intersections, dynamic colon ranges and reference-form INDEX. Add AREAS, ISREF, SHEET, SHEETS, AVERAGEA, MAXA, MINA, VARA, VARPA, STDEVA and STDEVPA (356 available names). Extend reference-aware aggregation, 3-D HSTACK/VSTACK, adjacent structured table columns and A1/R1C1 conversion. Preserve reference provenance and invalidate initially empty dependencies. Share formula scanning across copy, rename and structural rewrites.

Add indexed worksheet insertion and atomic worksheet reordering; repair deleted/crossed 3-D endpoints and ordinary deleted references in formulas, names, validation and conditional-format expressions. Protect unsupported single-sheet structural changes to a multi-sheet span with preflight rejection. Add a reusable Move worksheet dialog, Office-style position setter, host/worker commands and typed C# methods. Add live Reference tools to Home/Formulas without changing the default five-sheet workbook. Fix asynchronous dialog-close focus interference.

Add 144 Node regressions, 14 Chromium groups, strict declarations and installed-package consumers. Extend actual Windows native runtime smoke to 22 assertions per framework. Harden Avalonia initialization by awaiting AdapterCreated before navigation, checking completion for the trusted host URI and cancelling initialization on disposal. This addresses a suspected readiness race exposed by a previous main-CI navigation timeout; exact native outcomes are recorded separately from local Node/browser results. Full Excel, Office.js, OOXML and native-device parity remains unfinished. See the updated feature audit and reference contract.

## 0.6.0

Add the source-level Excel feature audit and generated, drift-tested function/API inventories (345 available names, not a parity claim). Implement 12 database functions, all AGGREGATE selectors/options with reference metadata, improved SUBTOTAL, MAKEARRAY, ISOMITTED, direct/recursive LAMBDA invocation and stricter helper-array/arity handling. Bound database criteria cross-product work.

Add immutable copy snapshots; nine PasteSpecial modes, arithmetic/transpose/blank handling, exact format replacement and validation rebasing; linear/growth/anchored date series; sparse special-cell queries. Fix formula-error copying and transactional edge cases. Add reusable accessible editing dialogs, focus/read-only/lifecycle handling, fill/paste shortcuts, sample examples, host/worker commands, Office copy flags and declarations. Add 298 Node regressions, 12 Chromium groups and installed-package/type consumers. Existing native/React/Blazor qualification is not implied by local JavaScript tests.

## 0.5.0

Add managed pivot reports with persistent definitions, 12 aggregations, multiple row/column fields and measures, filters, correctly aggregated totals, refresh, stale-state checks, drill-down and transactional history. Structural edits translate pivot references. Native XLSX exports now generate supported row-layout pivot definitions, typed cache definitions/records and connected package relationships; unsupported layouts require explicit flattening. Add import of the matching native subset, field-management studio commands, TypeScript and C# APIs, host/worker routes, schema validation fixtures and installed-package/browser/native tests. This is not arbitrary lossless or complete native Excel PivotTable compatibility.

## 0.4.0

Authenticated HTTP/SSE workbook coauthoring, durable ordered room commits, independent content/style/comment merging, explicit conflicts, reload-safe pending requests and presence. Studio Share controls and deployable Node CLI. Reject expired unknown retries, protect edits made while joining, correct browser fetch binding and refresh remote formula indexes. Release archives exclude private untracked runtime files.

## 0.3.0

Office-style batched APIs, loaded snapshots and atomic rollback. Vector chart export and printed fragments, conditional formatting, fitted pages and repeated titles. Subsequent qualification covers React 18/19, seven companion packages and real Windows WPF/WinUI/Avalonia WebViews.

## 0.2.0

Expand calculation to 330 available names, modern arrays/lookups, distributions/inverses, matrix/complex/engineering/finance, A1/R1C1 translation and reference/array semantics.

## 0.1.0

Recover engine, control, studio, file exchange, framework/native source, tests and distribution as readable source; remove incomplete transfer fragments without execution.
