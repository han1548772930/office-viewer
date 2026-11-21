# DOC-02 Vertical Alignment Evidence

Status: **Not accepted**. This evidence covers superscript and subscript only. It does not establish complete DOC-02 parity.

## Reference

- Fixed Univer commit: `eaf4cfa04fabdbcfc922c692f3450459f8a29546`.
- `packages/docs-ui/src/menu/schema.ts` registers `SetInlineFormatSubscriptCommand` and `SetInlineFormatSuperscriptCommand` in the Docs START/FORMAT group.

## Local behavior

- Inline runs use `verticalAlign: 'subscript' | 'superscript'`. Normal baseline is stored without a vertical-alignment value. A `baseline` patch clears the direct field and legacy metadata field, then adjacent equivalent runs compact together.
- Ribbon commands use `doc.command.format-text`. Selected text is persisted through document history. At a collapsed caret the format remains pending until text is inserted.
- HTML snapshots use `<sub>` and `<sup>`, and browser DOM import restores their run formatting.
- Canvas layout measures superscript and subscript glyphs at 65 percent size. Canvas rendering offsets their baselines while preserving the line height.
- DOCX writes native `w:vertAlign` values. The first-class run field takes precedence over the legacy `metadata.verticalAlign` value.

## Verification

- Vertical-alignment, document-command, toolbar-state, Canvas, and Facade contract tests passed, 39/39.
- `document-toolbar.browser.test.mjs` passed, 9/9. It covers selected-range formatting, toggling back to baseline, undo/redo, collapsed-caret pending formatting, and browser DOM HTML roundtrip.
- `pnpm.cmd exec vue-tsc -b` passed.
- `pnpm.cmd build` passed.
- `cargo fmt --check --manifest-path crates/office-writer/Cargo.toml` passed.
- `cargo test --manifest-path crates/office-writer/Cargo.toml` passed, 39/39.

## Unverified

- Production Writer WASM build and writer-to-parser browser roundtrip are unverified. `pnpm.cmd build:writer` could not start because `wasm-pack` is not installed in this environment.
- Reopening the generated DOCX in a desktop Office application has not been tested.
- Full DOC-02 parity remains unaccepted. Font families, sizes, colors, highlighting, paragraph formatting, mixed ranges, and their Office import/export behavior require their own evidence.
