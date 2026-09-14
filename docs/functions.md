# Formula inventory and reference semantics · 0.2.0

The engine exposes 330 built-in names (140 more than 0.1.0), including interpreter special forms and legacy aliases. This is an availability inventory, not an Excel behavioral-conformance percentage. `new Workbook().Calculation.FunctionNames` is the runtime source of truth; the studio's Insert Function dialog uses it directly.

New families include beta/gamma/normal/lognormal/chi-square/Student-t/F/binomial/Poisson/exponential/Weibull distributions and inverse functions; descriptive statistics and regression; matrix multiplication/determinants/inversion; 48-bit bit operations; base conversions; complex arithmetic; international workdays; dated financial functions; and modern array/text operations.

XMATCH/XLOOKUP support binary ascending/descending searches (the input must already be sorted), reverse searches, exact/approximate matches and two-dimensional return arrays. TEXTBEFORE/TEXTAFTER/TEXTSPLIT support the implemented optional instance, case, end-match, delimiter-array and padding behaviors. Unicode searches retain original text offsets.

`a1ToR1C1` / `r1c1ToA1` translate references with a specified origin, preserving quoted text and structured selectors. `getFormulasR1C1` / `setFormulasR1C1` use each destination cell as origin. ADDRESS and INDIRECT accept R1C1 mode. Implicit intersection `@` resolves unique reference intersections; ROW/COLUMN and the registered scalar-function lifting produce real arrays/spills.

Aggregation preserves the distinction between reference values and directly supplied text/logical literals. IFERROR/IFNA broadcast array fallbacks. Unsupported grammar and methods remain explicit errors rather than JavaScript execution.

## Verification

665 Node tests and 34 Chromium groups passed locally for this change. The numerical corpus includes 208 independently generated SciPy 1.17.0 cross-checks. Inverse-tail regression tests check relative precision separately. SciPy agreement is not an Excel golden-file comparison. The live Formula lab contains additional working examples starting at row 26.

## Remaining boundaries

The catalog is not exhaustive. Full argument-count/coercion/locale parity, all scalar lifting, every optional argument, 3D references, unions/intersections, iterative calculation and Excel's complete precision/error behavior remain unqualified or unimplemented. Matrix work is bounded to 128 square dimensions and other operations use explicit resource limits. Financial solvers return NUM when their bounded iteration does not converge.

References: Microsoft ADDRESS, XMATCH and COUNT function documentation, and the WorksheetFunction.Sum remarks on literal versus reference coercion.
