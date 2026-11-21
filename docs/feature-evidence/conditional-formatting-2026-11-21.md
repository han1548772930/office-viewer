# Conditional Formatting Evidence (2026-11-21)

This record covers the numeric `cellIs` slice of SHEET-14. It does not change the overall parity status in `UNIVER_FEATURE_PARITY.md`, which remains unaccepted.

## Implemented

- The sheet model stores conditional rules with ranges, operator, numeric operands, style, priority, and `stopIfTrue`.
- The rule service validates and evaluates equal/not-equal, greater/less comparisons, and inclusive/exclusive ranges. Matching styles are composed for Canvas rendering without mutating a cell's base style.
- The management panel and Sheets Ribbon entry use the Sheets feature locale catalog.
- XLSX export writes `cellIs` rules and differential styles (`dxfs`) for fill, font color, and bold. Rule priority and multiple ranges are retained.
- XLSX import reads conditional rules and differential styles into the parser payload. The workbook normalizer maps those values to the editable sheet model.

## Verification

- `cargo test --manifest-path crates/office-writer/Cargo.toml`: 48 passed.
- `cargo test --manifest-path crates/office-parser/Cargo.toml`: 19 passed, including writer-to-parser conditional-format roundtrip.
- `node --experimental-strip-types --test packages/office-studio/tests/conditional-formatting.contract.test.mjs`: 4 passed.
- `pnpm.cmd --filter @office-studio/office-studio exec node --experimental-strip-types --test tests/office-i18n.contract.test.mjs`: 4 passed.
- `pnpm.cmd --filter @office-studio/office-studio check:types`: passed.

## Remaining validation and behavior

- Formula-based rules, color scales, data bars, icon sets, and range migration after structural edits are not implemented by this slice.
- Browser interaction, Excel/LibreOffice reopening, and Office sample interoperability remain unverified. Browser validation is deferred to the consolidated run requested for this work batch.
- The upstream feature scope is documented at [Univer Sheets conditional formatting](https://docs.univer.ai/zh-CN/guides/sheets/features/conditional-formatting). The fixed source reference remains commit `eaf4cfa04fabdbcfc922c692f3450459f8a29546`.
