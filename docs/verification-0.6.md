# Verification · 0.6.0

Local qualification on 2026-09-19, against source derived from main `79fe32584e646a614a6549c8a14ea12e976b2e02`. The reconstructed baseline tree, including the untouched Blazor runtime-source gitlink, matches `8530c4e7034af10ffcaa53ed0b29283bb985beed`.

## Results

| Check | Result |
|---|---|
| Baseline Node regression suite before edits | 752 passed |
| Updated Node regression suite | **1,050 passed, zero failed** (298 added assertions/test cases) |
| Strict TypeScript declarations/consumers | Passed, including new editing/control/Office signatures |
| Modular/site/standalone build | Passed |
| Isolated installed tarball | Passed ESM/CommonJS identity, headless controls, XLSX, Office, printing, collaboration exports, pivots, editing/calculation and declaration consumers |
| Offline Chromium base suite | 34 groups passed |
| Offline Chromium Office/printing suite | 8 groups passed |
| Offline Chromium managed-pivot suite | 6 groups passed |
| New offline Chromium editing suite | 12 groups passed |
| Collaboration HTTP browser suite in this container | **Environment-blocked**, not counted as passed: Chromium rejected `http://127.0.0.1:<port>/studio` with `ERR_BLOCKED_BY_ADMINISTRATOR` |
| Current desktop Excel / physical printer / native frameworks | Not run locally for this increment |

The complete `npm run test:browser` command retains the real HTTP collaboration suite. It is **not skipped or weakened** to make this environment green. GitHub CI runs that complete command, including all 12 new editing groups; inspect the CI results for the exact pushed commit independently of this local report. Node collaboration/server/store tests do pass locally.

The new numerical/semantic tests cover 19 AGGREGATE selectors × eight options, reference visibility/dependency changes, database criteria and all 12 D-functions, LAMBDA scope/omission/recursion/helper results, copy errors, paste modes/tiling/overlap, validation clipping/rollback, protected/spilled/merged targets, arithmetic, directional/date series, sparse queries and host/Office integration. Availability inventories are also checked against runtime exports. JSON/XLSX tests verify GridWeb-to-GridWeb round trips only.

The new browser tests exercise actual sample buttons, native dialogs, keyboard Copy/PasteSpecial/fill/undo, copied-value snapshots, read-only rejection, late validation rollback, special-cell navigation, Escape/focus restoration, reconnect cleanup and dark mobile geometry. Screenshots are emitted under ignored `test-results/` and uploaded by CI. Browser testing caught and fixed a focus-restoration race after closing dialogs; immediate keyboard undo now returns to the grid.

## Reproduce

```sh
npm ci --ignore-scripts --legacy-peer-deps
npm install --global typescript@5.8.3
npm run check
python -m pip install -r tests/requirements.txt
python -m playwright install chromium
npm run test:browser
```

Set `CHROMIUM_PATH` to the installed browser when needed. The offline suites can also be run individually with `python tests/browser.py`, `python tests/features-browser.py`, `python tests/pivot-browser.py` and `python tests/editing-browser.py`. Do not describe those four suites as the complete collaboration browser qualification.

Native/React/companion integration CI remains in place. The Blazor runtime-source pin is unchanged and must be advanced/qualified separately before claiming these root engine additions ship in its NuGet runtime. No workbook macro/native executable was used as a test oracle. Function availability and passing tests are not full Excel parity.
