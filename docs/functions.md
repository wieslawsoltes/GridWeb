# Formula inventory and calculation contracts · 0.7.0

The engine exposes **356 available names**, including interpreter special forms and legacy aliases. This is not an Excel behavioral-conformance percentage. `new Workbook().Calculation.FunctionNames` is the runtime source of truth; the studio's Insert Function dialog uses it. [The JSON inventory](functions.json) is regenerated with `npm run inventory` and tested against live availability.

Existing families include arithmetic/logical/text/date functions; beta/gamma/normal/lognormal/chi-square/Student-t/F/binomial/Poisson/exponential/Weibull distributions and inverses; descriptive statistics/regression; matrix operations; 48-bit bit arithmetic and base conversions; complex arithmetic; international workdays; dated finance; arrays and lookups. Historical numerical qualification remains recorded in [calculation-0.2.md](calculation-0.2.md); it is not a desktop Excel oracle.

## Database formulas

DAVERAGE, DCOUNT, DCOUNTA, DGET, DMAX, DMIN, DPRODUCT, DSTDEV, DSTDEVP, DSUM, DVAR and DVARP accept `database, field, criteria`. Matrices require a header row and at least one data/criteria row. Fields are case-insensitive headers or one-based numeric indexes. DCOUNT/DCOUNTA accept an omitted/blank field to count matching records.

Criteria columns in a row are AND; separate rows are OR. Repeating a header permits lower/upper bounds on the same field. The profile supports comparisons, escaped wildcards, unadorned text prefixes, and formula criteria stored in a worksheet under a non-database header. Formula criteria shift relative references for each database record and retain dependencies. Literal arrays cannot supply executable criteria formulas. DGET returns VALUE for no match and NUM for multiple matches. Database/criteria matrices remain bounded; more than 4,000,000 criterion checks return NUM rather than performing an unbounded cross product.

```text
=DSUM(A1:D100,"Sales",F1:G3)
=DGET(A1:D100,"Owner",F1:F2)
```

## AGGREGATE and SUBTOTAL

AGGREGATE supports selectors 1–19 and options 0–7. Selectors 1–13 accept references/arrays; 14–19 require their fourth `k` argument. Direct and supported named/INDIRECT/OFFSET references retain row visibility and formula identity. Filtered rows are excluded; hidden rows, errors and nested SUBTOTAL/AGGREGATE formulas are excluded according to the selected option. Nested totals are detected from parsed formulas, not strings containing function names. Sparse full-column ranges and spill outputs participate without allocating full-column matrices, and range dependencies track later edits.

Computed arrays do not retain row visibility or nested-formula provenance. This is an intentional distinction, not a reason to infer hidden-row information from result coordinates. SUBTOTAL implements selectors 1–11 and 101–111 through the same reference-aware path.

```text
=AGGREGATE(9,6,A2:A100)
=AGGREGATE(14,7,A2:A100,3)
=SUBTOTAL(109,A2:A100)
```

## Callable LAMBDA and helpers

LAMBDA expressions can be invoked directly and return lexical closures. Named LAMBDAs can recurse with bounded depth/work; limit errors stay NUM rather than being rewritten as NAME. Incorrect argument counts produce VALUE. Parameters are validated for the implemented name/uniqueness rules.

Omitted call arguments are tracked separately from blank values. ISOMITTED examines this metadata, including lexical scope and LET shadowing. MAKEARRAY uses one-based row/column indexes. MAP requires equal input shapes. BYROW/BYCOL preserve row/column orientation. MAP/MAKEARRAY/BYROW/BYCOL and SCAN reject unsupported nested array results with CALC rather than silently truncating them; REDUCE permits a growing array accumulator.

```text
=LAMBDA(x,x^2)(12)
=LAMBDA(x,y,IF(ISOMITTED(y),x*2,x+y))(7,)
=MAKEARRAY(3,3,LAMBDA(row,col,row*col))
=LET(scale,4,LAMBDA(x,x*scale))(5)
```

## Reference operators and A-suffixed statistics

The reference evaluator now retains coordinates for comma unions, whitespace intersections and dynamic colon ranges, including INDEX-selected areas/rows/columns. 3-D sheet spans are supported by the documented statistical consumers and HSTACK/VSTACK. AREAS, ISREF, SHEET and SHEETS inspect reference/sheet identity without treating calculated numbers as coordinates. SUM/COUNT and related sparse consumers track dependencies even in initially empty whole-axis areas. SUBTOTAL/AGGREGATE support unions and INDEX-selected references, but reject 3-D references.

AVERAGEA, MINA, MAXA, VARA, VARPA, STDEVA and STDEVPA include logical values and referenced/array text, skip actual empty cells, propagate errors and apply their sample/population denominator. Direct numeric text is numerically coerced; text inside an array/reference contributes zero. See [reference contracts](reference-semantics.md) for all supported consumers, selectors and resource limits.

```text
=SUM(Jan:Mar!B2:B10)
=SUM((A1:A3,C1:C3))
=SUM(A1:C4 B2:D5)
=SUM(A1:INDEX(A1:A10,5))
=AREAS((A1:A3,C1:C3))
=AVERAGEA({2,TRUE,"text"})
```

## Existing reference/array semantics

XMATCH/XLOOKUP support binary ascending/descending searches (already-sorted input), reverse searches, exact/approximate matching and two-dimensional return arrays. Modern text functions implement their documented repository profile, including original Unicode offsets. A1/R1C1 conversion uses explicit origins and preserves quoted text/structured selectors. ADDRESS/INDIRECT support R1C1 mode. Implicit `@` resolves unique reference intersections. ROW/COLUMN and supported scalar lifting produce real arrays/spills.

Aggregation distinguishes reference values from directly supplied text/logical literals. IFERROR/IFNA broadcast fallbacks. Unsupported grammar does not execute JavaScript.

## Qualification and boundaries

See [0.7 verification](verification-0.7.md) and the [feature audit](excel-feature-audit.md). The new regression matrix covers all 19 AGGREGATE selectors × eight options, database criteria/error cases and LAMBDA/helper behavior. GridWeb JSON/XLSX self-round-trips are not native Excel open/save certification.

The catalog is not exhaustive. Full argument/coercion/locale/name/precision/error compatibility, every optional argument, all scalar lifting, reference-valued LET/LAMBDA bindings, grouped 3-D structural changes, iterative calculation and future-function OOXML metadata remain incomplete or unqualified. Numeric matrices and other operations use explicit bounds; financial solvers return NUM on bounded non-convergence. Large operations remain synchronous unless a worker host is used.

Target references: Microsoft [AGGREGATE](https://support.microsoft.com/en-us/excel/functions/aggregate-function), [DSUM](https://support.microsoft.com/en-us/excel/functions/dsum-function), [LAMBDA](https://support.microsoft.com/en-us/excel/functions/lambda-function), [MAKEARRAY](https://support.microsoft.com/en-us/excel/functions/makearray-function), [ISOMITTED](https://support.microsoft.com/en-us/excel/functions/isomitted-function) and [BYROW](https://support.microsoft.com/en-us/excel/functions/byrow-function).

## Scoped-name and binding changes (0.8)

Function availability remains 356. Qualified worksheet-local names/calls and direct reference-valued LET/LAMBDA arguments/returns are supported; reference probes are memoized. Named cycles report CIRC and ISOMITTED excludes explicit qualified-name expressions. This does not provide reference returns through every special form. See [defined names](defined-names.md).
