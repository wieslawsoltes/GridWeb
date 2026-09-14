# Publishing GridWeb

## Qualification

Run `npm ci --ignore-scripts --legacy-peer-deps`, install TypeScript 5.8.3 or newer, then `npm run check`. Install the pinned Python requirements and Playwright Chromium to run `npm run test:browser`. CI checks Node 22/24 and native projects on Windows. Do not merge a known failing required check.

## GitHub Pages

The Pages workflow builds and tests `site/` on main, then uses GitHub's Pages artifact/deployment actions. Repository Pages must use GitHub Actions as its build source. Its URL is https://wieslawsoltes.github.io/GridWeb/. A successful source commit alone does not prove deployment; check the Pages workflow and deployed application.

## npm and GitHub release

Use package name `@wieslawsoltes/gridweb`. Update package.json and lockfile together, update changelog and compatibility documentation, and merge a reviewed PR. Create an immutable `v<version>` tag on the qualified main commit. The Release and npm workflow also accepts an existing tag and expected SHA for a controlled retry.

The workflow validates tag/version/SHA, reruns checks, builds npm/source/browser archives, creates draft release assets, verifies checksums and installed consumers, and publishes the exact tarball using the repository NPM_TOKEN secret plus OIDC provenance. It downloads and verifies public registry bytes before publishing the GitHub release. Existing versions/assets are not overwritten; a mismatching immutable artifact fails the workflow.

`npm run release:pack` creates local artifacts only; it does not publish. React is optional and omitted from headless offline consumer checks. The root has no mandatory runtime dependencies. The optional companion bundle is built in integrations/ and carries its dependency notices.

## Native preview packages

The manually dispatched NuGet workflow checks the preview version in dotnet/Directory.Build.props, builds/packs the client and WPF/WinUI/Avalonia adapters, and runs client protocol tests. Publishing is explicit and uses NUGET_API_KEY. Compilation and packaging do not establish native WebView runtime, input, accessibility or printer equivalence; record those results separately.
