# Formula inventory

The recovered baseline has 190 available built-in names, including interpreter special forms. See the machine-readable [functions.json](functions.json) for the exact list and src/functions.js / src/calculation.js for signatures and evaluation paths.

Categories include aggregation/criteria, mathematics/statistics, text, dates, financial calculations, lookup/reference and dynamic arrays. LET, LAMBDA, lazy conditionals and reference-aware functions are interpreted rather than ordinary eager calls. Custom function registration is explicit and application-owned.

Function-name coverage is not a conformance score. Binary XMATCH/XLOOKUP search, several optional text/date arguments, advanced distribution/engineering/financial families and full error/coercion/locale equivalence were absent in the recovered baseline. Subsequent additions require tests and inventory updates.
