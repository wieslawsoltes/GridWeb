# Companion integrations

This is an optional, real package build, not mocked substitutes. Install with `npm install --prefix integrations`, run `npm run build:integrations`, then `npm run build` at the repository root. The build sets demo/build-info.js and copies a bundle and dependency licenses into the distribution.

Dockyard owns docking/document/tool panes; RibbonWeb routes workbook commands; TreeDataGridWeb presents editable live cell inspection; DynamicDataWeb maintains the observable keyed cache; ReactiveWeb binds view-model state and commands; RBushWeb indexes chart hit regions; QuikGraphWeb visualizes calculation dependencies.

Use the studio's companion-workspace command after bundling. The dependency-free default studio does not require the bundle. The baseline inspector limits displayed rows to 1,000 and graph edges to 1,500. These are sample limits, not engine limits. Review bundle/license output before distribution, including QuikGraphWeb's MS-PL and optional-engine licenses. Native/React/companion runtime tests must be reported separately from baseline browser tests.
