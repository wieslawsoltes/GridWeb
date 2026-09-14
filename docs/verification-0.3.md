# Verification for the 0.3 change

The print/API additions were tested locally against the recovered 0.1 baseline: the 263 existing tests plus25 new print/API tests pass, as do strict types, installed tarball consumers and34 existing plus8 new Chromium groups. A separate13-test preservation experiment is not committed and is excluded from release test totals.

The release branch additionally includes the merged 0.2 calculation work (402 tests). Current-branch CI must pass before merge; its combined total is authoritative rather than assuming arithmetic totals establish success. Windows CI compiles the WPF, WinUI and Avalonia controls and samples and executes C# client tests. This is not native UI runtime qualification.

New browser checks execute deferred batched edits/rollback against the displayed workbook, invoke the studio API example, download SVG chart bytes, inspect printed chart/conditional-format output, exercise page geometry and invalid settings, and record uncaught errors. Reports and screenshots are uploaded by CI under test-results/.

Package qualification now imports /office, /printing and /browser in an isolated installed tarball, verifies shared model identity, executes a formula batch and vector output, and compiles new strict TypeScript consumers. The generated modular browser entry is included explicitly.
