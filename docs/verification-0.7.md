# Local qualification · 0.7.0

Date: 2026-09-20. Base: main `45b59732e45c67daec3a8cbe7baccf64bd788e3e`, whose source tree is `8df8086b2ae7688bf8d487a860b93297c45c6210`. The reviewed source snapshot was obtained from the repository's existing Source snapshot workflow and verified before editing. The Blazor runtime gitlink `c833be49d472583b6f56225862e0aa7d201c1da7` is unchanged.

## Executed locally

| Check | Result |
|---|---|
| `npm run check` | Passed; includes tests, strict types, build and isolated installed-package consumers. |
| `npm test` | **1,194 passed, zero failed/skipped**, compared with 1,050 baseline tests. |
| New reference regression file | **144 tests** covering grammar, errors/coercion, 3-D/union/intersection/INDEX dependencies, structural reference repair, history, wrappers and JSON/XLSX self-round-trips. |
| Seeded rectangle comparison | 250 deterministic rectangle pairs checked against a separate dense coordinate-set calculation for union sums and intersection sums; these 500 comparisons form one Node test, not 500 additional test cases. |
| Strict TypeScript | Core, control, Office, host and existing declarations; new worksheet/index errors are tested. |
| Installed package | ESM/CommonJS constructor identity, headless imports, real formula/editing/reference/Office consumers, serialization and external strict type compilation. |
| Browser base suite | **34 groups passed**. |
| Browser Office/printing suite | **8 groups passed**. |
| Browser managed pivot suite | **6 groups passed**. |
| Browser editing suite | **12 groups passed**. |
| New reference/browser suite | **14 groups passed**, including actual keyboard editing, move/reorder/undo, stale/read-only apply guards, reference repair, independent controls, focus isolation, dark mobile layout and reconnect. |

The five local offline suites total **74 Chromium groups**, all passed with Chromium 144.0.7559.96. Screenshots were inspected for the reference example and dark narrow-screen dialog. The default five-sheet sample is retained. The live function inventory is **356 names**, not a behavioral conformance percentage.

## Environment-blocked or not executed locally

The full `npm run test:browser` command includes the real HTTP collaboration suite. That suite was attempted but Chromium rejected the loopback navigation with `ERR_BLOCKED_BY_ADMINISTRATOR`; it did not reach the collaboration assertions. This is **not** a collaboration pass. No test was disabled or weakened: the full six-suite command remains required in CI, including all eight collaboration groups. A complete run would have 82 groups. Node HTTP collaboration regressions passed as part of the 1,194 tests.

A local .NET SDK is unavailable, so C# compilation, the typed-client protocol test, and actual Windows WPF/WinUI/Avalonia execution were not run locally. CI retains native builds and adds six reference/worksheet assertions to the shared native suite. Including existing and framework-specific assertions, each native report must contain at least **22 passes**, zero error, and a successful process exit; the former threshold was 13. React/companion and OpenXML schema jobs remain part of full CI. The independently pinned Blazor runtime has its own checks and does not inherit root-engine changes automatically.

## Baseline failure and native startup hardening

The previous 0.6 PR CI passed, but the later main run [35469123373](https://github.com/wieslawsoltes/GridWeb/actions/runs/35469123373) failed in Avalonia job [105966742237](https://github.com/wieslawsoltes/GridWeb/actions/runs/35469123373/job/105966742237). Its artifact recorded `TimeoutException` in `GridWebControl.InitializeCoreAsync` while waiting for navigation, with zero completed native assertions. Compilation succeeded. Other jobs in that run passed; version publication was skipped.

The new code waits for `NativeWebView.AdapterCreated` before navigation, accepts navigation completion only for the trusted bundled URI, and cancels initialization on disposal. This addresses a suspected readiness race; the baseline evidence does not conclusively identify the sole cause. Native CI on the exact new source must qualify the change. The navigation allowlist/new-window block remains in place and no external navigation or network access is added.

## What these results do not establish

No desktop Excel oracle, independently Excel-generated differential corpus, physical GPU, printer, cross-platform native input, screen-reader or production-scale collaborative deployment was used for this increment. JSON/XLSX self-round-trips test GridWeb's own supported format profile; they are not native Excel open/save equivalence. Numerical/grammar regressions qualify their explicit cases, not every function signature, coercion, name rule or resource ceiling.

Refer to [reference semantics](reference-semantics.md), [full feature audit](excel-feature-audit.md) and the actual PR/commit workflow runs for subsequent remote qualification. This document records local evidence and does not predeclare CI or npm/Pages publication success.
