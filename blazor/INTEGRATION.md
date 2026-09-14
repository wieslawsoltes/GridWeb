# Integration contract

See [README.md](README.md) for installation, all public surfaces, executable snippets, forms, ownership, source builds and releases.

## Ownership

`Spreadsheet` owns a native session unless given a `GridWorkbook`. `WorkbookProvider` or a manually created `GridWebModule` owns shared/nonvisual sessions. Child controls only borrow shared sessions; dispose them before their provider. Caller-owned native `Workbook` instances passed to `SpreadsheetSession` are not destroyed by session disposal. Returned JavaScript references are independent handles: dispose each reference as well as its engine owner. Never register browser modules/workbooks as Server singletons.

## Binding and execution

`Value` carries complete GridWeb JSON with a monotonic acknowledgment revision; `ValueRevision` forces an intentional reset. Selection has an independent revision. Flush active edits before persisting/submitting. Static prerendering does not access JavaScript. Large data uses explicit JSON/binary streams with a default 64 MiB limit; native workbook limits remain applicable. Typed operations accept data, not JavaScript source, and `BrowserValue.Literal` prevents descriptor interpretation. Native synchronous callbacks execute in browser modules; an asynchronous Server callback cannot replace a synchronous calculation function.

## Compatibility

Every browser package entry is included in the native API object; the Node-only collaboration server and optional React adapter are intentionally not in the RCL. Existing native desktop clients/pipelines are untouched. The separate module worker has the same allowlisted message protocol but does not share in-memory state with UI workbooks. Collaboration requires a separately deployed service. Canvas cell rendering does not accept arbitrary Razor fragments.

## Evidence and release gates

Native CI remains separate and required before merge. The Blazor workflow packs both target assemblies, checks actual package contents, restores package consumers, runs shared managed/JavaScript tests and exercises WebAssembly/Server in Chromium. GridWeb-specific tests assert formulas, batch rollback, literal CSV, native pivot XLSX, import revision protection, native handles, worker protocol and borrowed ownership. The sample additionally verifies actual keyboard edits, EditContext changes, chart/print output and full streamed binding.

NuGet publication uses the validated artifact, checks immutable package conflicts and compares the downloaded public payload before creating a release. An upload acknowledgment is not proof of public availability. Retry the original failed publication job, not a rebuilt package of the same version.
