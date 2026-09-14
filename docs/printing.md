# Vector charts and print layout

`@wieslawsoltes/gridweb/printing` exports paginate, printPagesHTML, createPrintDocument and chartToSVG. The root core export includes the same functions. Print geometry uses the worksheet's sparse AxisLayout, not a second copy of its cells.

```js
import { createPrintDocument, chartToSVG } from '@wieslawsoltes/gridweb/printing';
const html = createPrintDocument(sheet, {
  paper: 'A3', orientation: 'landscape', area: 'A1:J90',
  fitToWidthPages: 1, fitToHeightPages: 0,
  repeatRows: 1, repeatColumns: 1,
  rowBreaks: [30], columnBreaks: [5],
  margins: { top: 40, right: 40, bottom: 40, left: 40 },
  pageOrder: 'downThenOver', header: '&F · &A', footer: 'Page &P of &N'
});
const svg = chartToSVG(sheet, sheet.Charts[0]);
```

A4, Letter, A3 and Legal are supported. Scale ranges from 10% to 400%; fit targets are maximum page counts, with zero unconstrained. Index-based breaks are zero-based; a break at row30 starts a new page before row31. Repeated titles are the leading rows/columns of the chosen print area. Margins and chart dimensions are CSS pixels at 96 DPI. Header/footer tokens are workbook name (&F), sheet name (&A), page (&P), and total pages (&N).

Printed content now includes supported table styles, conditional rules, data bars, color scales, underlining/strikethrough, and the six existing chart families. Charts use the same drawing algorithm recorded into inert SVG and are clipped into fragments when crossing pages. The default area expands to include chart extents; an explicit area still clips them. `includeCharts:false` omits charts. Cell boxes and table widths use fixed axis geometry to avoid content expanding the page grid.

Studio commands: Page Layout → Advanced print opens an isolated preview with call-scoped options; Insert/Page Layout → Chart SVG exports the first chart on the active sheet. Existing page-layout editing uses the updated printable presentation. Advanced options do not overwrite saved worksheet metadata.

## Limits and qualification

This is not printer-identical Microsoft Excel output. Browser font shaping, line wrapping, printer margins, scaling, native drivers, image/drawing support and chart label measurement differ. SVG label width estimation is approximate. The inherited six chart families are not Excel's complete chart catalog. Advanced options are not persisted to the workbook's original XLSX print settings by this change.

Dense area operations remain bounded and preview is limited to 100 pages. Oversized rows/columns/titles fail explicitly unless scaling can fit them. Charts are limited by the existing chart model. Repeated-title ranges up to100 are accepted. Physical printer and native WebView printing were not exercised; Chromium HTML/SVG rendering and geometry tests are separate evidence.

References: Microsoft Excel Page Setup and scale-to-fit documentation. No proprietary fonts or drawing assets are bundled.
