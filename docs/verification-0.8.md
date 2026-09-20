# Recovery and verification · 0.8.0

Date: 2026-09-20. Base main: `7195869c5950af0deecbdad1560429e87d705c49`; source tree `eaeccd1de3c4bda6adf2aae8423b5f188d688e66`. The complete snapshot and its existing `blazor/runtime-source` gitlink were restored and the tree verified before replaying interrupted source edits. The final tree contains further review fixes and is requalified, not claimed identical to the original interrupted fingerprint.

## Executed locally

| Check | Result |
|---|---|
| Baseline Node suite | 1,194 passed |
| Updated Node suite | 1,347 passed, zero failures/skips; 153 added |
| Added tests | 106 core/scoped names; 9 Office named items; 31 selection-label naming; 7 source archives |
| Strict TypeScript, build, installed ESM/CommonJS/type consumers | Passed through `npm run check` |
| Function/API inventory | Generated and drift-tested; 356 function/special-form/alias names, unchanged |
| Offline Chromium | 93 groups passed: 34 base, 8 Office/printing, 6 pivots, 12 editing, 14 references, 19 names |
| HTTP collaboration browser suite | Attempted; navigation blocked by local Chromium `ERR_BLOCKED_BY_ADMINISTRATOR` on loopback. CI retains this suite unchanged |
| Visual inspection | Desktop Named formulas and narrow dark Name Manager screenshots reviewed |

Core coverage includes local/global duplicate spelling, qualification/shadowing, stable definition contexts, safe free-reference rename, lexical capture, closures, reference identity and memoization, cycles/recursion, sparse dependencies/spills, row/column structure, undo, metadata validation, legacy JSON, shared limits, collaboration and RPC. OOXML tests include a hand-authored original-index fixture with a skipped unsupported sheet. All 15 nonempty combinations of four label edges are tested independently of normalization/conflict/overwrite tests.

Recovery review additionally fixed recursive Name Manager formula-edit+rename ordering, qualified ISOMITTED and R1C1-shaped identifiers. The actual browser rename group exercises a changed recursive body and undo. Archive regressions include regular/executable files, untracked-file exclusion, absent/initialized gitlinks, symlink path rejection, empty indexes and reserved manifests.

## Remote and publication gates

The complete seven-suite `npm run test:browser` includes 8 real authenticated HTTP collaboration groups, for **101 expected groups**. Local browser policy is not a substitute for these checks. The local environment has no .NET SDK: C# protocol/build checks, actual native WPF/WinUI/Avalonia smoke, and OpenXML schema checks must run remotely. Native reports now require **28 passes** per framework, no error and successful process exit. The standard schema job additionally reads the scoped-name fixture. React 18/19 and separately pinned Blazor checks remain separate jobs.

The actual PR/merge checks and comments record remote results. Main CI, GitHub Pages and registry/release publication are distinct outcomes. The source archive fix records submodule commits without embedding their contents; it does not advance Blazor or imply npm publication.

## Remaining limits

No independent desktop Excel oracle, native Excel open/save/reopen, physical-printer/GPU/device matrix or production-scale coauthoring qualification was performed. JSON/XLSX self-round-trips qualify only their explicit assertions. Full Unicode/external/workbook-qualified naming, native relative-name export, null/tombstoned-context export, all reference-preserving special forms, full Name Manager/Office.js APIs and arbitrary lossless OOXML remain incomplete. Full Excel parity is unfinished; see [feature audit](excel-feature-audit.md) and [defined names](defined-names.md).
