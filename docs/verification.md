# Verification

Latest increment: [0.6.0 verification](verification-0.6.md). The results below are historical recovery evidence, not current test totals.

The supplied `GridWeb-0.1.0-source.zip` SHA-256 is `eaa753d563bc24d2be920bb5bca8bb419b266e2cd471c340542d57c58c9db918`. Its complete UTF-8 source was recovered directly. The incomplete initial `.bootstrap` fragments were not executed.

On 2026-09-14 the recovered source passed 263 Node regression tests, strict TypeScript checks, browser/standalone build, and isolated npm tarball consumers. Chromium 144.0.7559.96 passed 34 browser groups with no uncaught page errors. These are local results; GitHub Actions results qualify the eventual remote commit separately.

Core coverage includes formulas, dependency invalidation, spill blocking, names, history, atomic rollback, structural edits, data tools, layout geometry, safe interchange and explicit host/worker messages. Browser coverage includes real keyboard/pointer edits, clipboard behavior, resizing, frozen panes, merges, shared views, page modes, XLSX round trips and embedded host RPC.

Native projects and companion/framework adapters have their own build/runtime requirements. A successful JavaScript browser test does not certify WPF, WinUI or Avalonia. A successful native compilation does not certify native rendering/input/accessibility. React and all seven companion packages must be installed and exercised before runtime qualification is claimed.

Run `npm run check`, followed by `npm run test:browser`. CI uploads its browser reports. Do not carry historical test totals forward as results for changed code without rerunning.
