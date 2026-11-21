# Formula Function Directory Difference

Reference baseline: `dream-num/univer@6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68`  
Reference root: `E:\rust_test\univer`  
Comparison date: 2026-11-21

The upstream set below is taken from active entries in each `packages/engine-formula/src/functions/<category>/function-map.ts`. Commented-out registrations are excluded. Local support means the formula evaluator executes the function and the insert-function catalog lists it. Catalog exposure alone does not count as implementation.

| Category | Upstream registered | Locally executable and listed | Remaining |
| --- | ---: | ---: | ---: |
| Date | 27 | 27 | 0 |
| Text | 46 | 26 | 20 |
| Logical | 21 | 11 | 10 |
| Lookup | 37 | 11 | 26 |
| Financial | 55 | 17 | 38 |

## Date

Implemented: `DATE`, `DATEDIF`, `DATEVALUE`, `DAY`, `DAYS`, `DAYS360`, `EDATE`, `EOMONTH`, `EPOCHTODATE`, `HOUR`, `ISOWEEKNUM`, `MINUTE`, `MONTH`, `NETWORKDAYS`, `NETWORKDAYS.INTL`, `NOW`, `SECOND`, `TIME`, `TIMEVALUE`, `TODAY`, `TO_DATE`, `WEEKDAY`, `WEEKNUM`, `WORKDAY`, `WORKDAY.INTL`, `YEAR`, `YEARFRAC`.

Remaining: none in the fixed upstream date-function map. `NOW` and `TODAY` use the local clock and calendar date. A single evaluation captures one clock snapshot, and whole-workbook recalculation shares one snapshot across all formulas. A fixed `now` option supports deterministic callers and contract tests. `DATEDIF` supports `Y`, `M`, `D`, `MD`, `YM`, and `YD`, and applies the workbook's 1900/1904 date system.

`DAYS360` was cross-checked against `packages/engine-formula/src/functions/date/days360/index.ts` and `packages/engine-formula/src/basics/date.ts` at the fixed commit. NASD basis 0 adjusts a start date on the last day of February to day 30, and adjusts an end date on the last day of February to day 30 when the start date is also a February month-end. European basis 4 adjusts day 31 only. The local evaluator now shares the same 30/360 date normalization used by `YEARFRAC`, then applies the sign from the original argument order. Contract cases cover leap-year February month-end through March 31 in both directions and both methods. Scalar boundaries are verified; array broadcasting remains unsupported by the local scalar `FormulaValue` contract.

`EPOCHTODATE` uses UTC Unix timestamps with unit codes 1/2/3 for seconds, milliseconds, and microseconds, then offsets into the workbook's 1900 or 1904 serial system. `TIMEVALUE` supports ISO date-only text and 24-hour/AM-PM time text, optionally following an ISO date. `TO_DATE` preserves nonnumeric scalar values and returns numeric serials unchanged. Timezone-qualified TIMEVALUE strings and number-format metadata propagation for TO_DATE are not represented by this scalar formula runtime and remain unverified.

## Text

Implemented: `CHAR`, `CLEAN`, `CODE`, `CONCAT`, `CONCATENATE`, `EXACT`, `FIND`, `LEFT`, `LEN`, `LOWER`, `MID`, `NUMBERVALUE`, `PROPER`, `REPLACE`, `REPT`, `RIGHT`, `SEARCH`, `SUBSTITUTE`, `TEXTAFTER`, `TEXTBEFORE`, `TEXTJOIN`, `TRIM`, `UNICHAR`, `UNICODE`, `UPPER`, `VALUE`.

Remaining: `ARRAYTOTEXT`, `ASC`, `BAHTTEXT`, `DBCS`, `DOLLAR`, `EUROCONVERT`, `FINDB`, `FIXED`, `LEFTB`, `LENB`, `MIDB`, `NUMBERSTRING`, `REGEXEXTRACT`, `REGEXMATCH`, `REGEXREPLACE`, `REPLACEB`, `RIGHTB`, `SEARCHB`, `T`, `TEXT`, `TEXTSPLIT`, `VALUETOTEXT`.

`TEXTSPLIT` remains unsupported because the current formula runtime result contract (`FormulaValue`) is scalar-only and formula references, dependency recalculation, and cell caching all assume one scalar result per formula cell. Supporting its upstream 2D output requires array-valued results and dynamic-array propagation through those layers. It is intentionally not listed in the insert-function catalog until that contract is implemented. This batch has not claimed array behavior for any text function.

## Logical

Implemented: `AND`, `FALSE`, `IF`, `IFERROR`, `IFNA`, `IFS`, `NOT`, `OR`, `SWITCH`, `TRUE`, `XOR`.

Remaining: `BYCOL`, `BYROW`, `GROUPBY`, `LAMBDA`, `LET`, `MAKEARRAY`, `MAP`, `PERCENTOF`, `REDUCE`, `SCAN`.

The parser already treats bare `TRUE` and `FALSE` as Boolean literals. `TRUE()` and `FALSE()` are separately executable zero-argument function calls, and the insert-function builder accepts empty arguments for these two entries.

## Lookup

Implemented: `ADDRESS`, `AREAS`, `CHOOSE`, `COLUMNS`, `COLUMN`, `HLOOKUP`, `INDEX`, `MATCH`, `ROW`, `ROWS`, `VLOOKUP`.

Remaining: `ANCHORARRAY`, `CHOOSECOLS`, `CHOOSEROWS`, `DROP`, `EXPAND`, `FILTER`, `FORMULATEXT`, `GETPIVOTDATA`, `HSTACK`, `HYPERLINK`, `IMAGE`, `INDIRECT`, `LOOKUP`, `OFFSET`, `SORT`, `SORTBY`, `TAKE`, `TOCOL`, `TOROW`, `TRANSPOSE`, `UNIQUE`, `VSTACK`, `WRAPCOLS`, `WRAPROWS`, `XLOOKUP`, `XMATCH`.

`ADDRESS`, `AREAS`, `ROW`, and `COLUMN` support scalar references. `ROW()` and `COLUMN()` use the containing formula cell during workbook recalculation. Dynamic-array results for multi-cell `ROW(reference)` and `COLUMN(reference)` and multiple-area union references for `AREAS` remain unverified because the current formula value contract is scalar and the parser does not represent union references.

## Financial

Implemented: `DB`, `DDB`, `DOLLARDE`, `DOLLARFR`, `FV`, `FVSCHEDULE`, `IPMT`, `IRR`, `NPER`, `NPV`, `PMT`, `PPMT`, `PV`, `RATE`, `RRI`, `SLN`, `SYD`.

Remaining: `ACCRINT`, `ACCRINTM`, `AMORDEGRC`, `AMORLINC`, `COUPDAYBS`, `COUPDAYS`, `COUPDAYSNC`, `COUPNCD`, `COUPNUM`, `COUPPCD`, `CUMIPMT`, `CUMPRINC`, `DISC`, `DURATION`, `EFFECT`, `ISPMT`, `MDURATION`, `MIRR`, `NOMINAL`, `ODDFPRICE`, `ODDFYIELD`, `ODDLPRICE`, `ODDLYIELD`, `PDURATION`, `PRICE`, `PRICEDISC`, `PRICEMAT`, `RECEIVED`, `TBILLEQ`, `TBILLPRICE`, `TBILLYIELD`, `VDB`, `XIRR`, `XNPV`, `YIELD`, `YIELDDISC`, `YIELDMAT`.

The difference is intentionally not a completion claim. Functions returning arrays, depending on workbook context or external time, and requiring date-convention or financial-domain compatibility remain outside this implementation batch.

## Financial Calculation Evidence

The financial evaluator supports scalar `DB`, `DDB`, `DOLLARDE`, `DOLLARFR`, `RRI`, `SLN`, `SYD`, `NPER`, `RATE`, `IPMT`, `PPMT`, `IRR`, and `FVSCHEDULE`, and retains `NPV`'s end-of-period discounting. Depreciation functions preserve upstream defaults and period boundaries, including DB's three-decimal rate rounding and first/last-period month treatment. `DOLLARDE` and `DOLLARFR` floor the fractional denominator before conversion. Payment timing is normalized using the fixed reference's zero/end and nonzero/beginning convention. `IRR` checks for positive and negative cash flows, uses the supplied guess for Newton iteration, then performs bounded bracket search and bisection before returning `#NUM!`. Array broadcasting for `DB`, `DDB`, `SLN`, `SYD`, and `RRI` is not represented by the current scalar `FormulaValue` contract and remains unverified. `XIRR` and `XNPV` remain unimplemented because they require date-paired cashflow processing.

## Conditional Aggregation Evidence

This batch adds executable and catalog-listed support for `COUNTBLANK`, `COUNTIF`, `COUNTIFS`, `SUMIF`, `SUMIFS`, `AVERAGEIF`, and `AVERAGEIFS`. `COUNTA` was already listed and executable, and its sparse-range handling now counts implicit blank cells correctly. The evaluator covers blank cells and formula-produced empty strings separately, criteria operators, case-insensitive text, `*`/`?` wildcards with `~` escapes, numeric and ISO-date criteria, error-valued cells, and range-size validation. Multi-criteria functions require equal-sized ranges. `SUMIF` and `AVERAGEIF` align a supplied result range from its upper-left cell to the criteria-range shape, matching the fixed reference behavior.
