# DOC-05 Checklist Evidence

Status: **未验收**. This record covers the checklist slice only and does not represent full list feature parity.

## Reference contract

- Fixed Univer commit: `eaf4cfa04fabdbcfc922c692f3450459f8a29546`.
- `packages/docs-ui/src/menu/schema.ts` registers `CheckListCommand` under `RibbonStartGroup.LAYOUT`.
- `packages/docs-ui/src/commands/commands/list.command.ts` uses paragraph bullet types `CHECK_LIST` and `CHECK_LIST_CHECKED`. `CheckListCommand` applies a checklist to the selected paragraph range. `ToggleCheckListCommand` changes one item's checked state.
- `packages/docs-ui/src/controllers/render-controllers/doc-checklist.render-controller.ts` routes a click on the list symbol to the toggle command.
- `packages/core/src/docs/data-model/preset-list-type.ts` displays unchecked and checked symbols as U+2610 and U+2611.

## Local model and behavior

- `DocumentBlock.type === 'checklist'` stores list membership and checked state in serializable `checklist.listId` / `checklist.checked` fields. The checked value is list semantics, not a content-control node or inline glyph.
- `doc.command.set-block-type` creates, changes, or removes checklist membership for the selected paragraph range. Adjacent checklist paragraphs reuse the same list ID. New paragraphs created by Enter join the same list unchecked. Enter on an empty task exits the list.
- `doc.command.toggle-checklist-item` toggles one checklist block. The workbench's existing document snapshot history records the change, so undo and redo restore the checked state.
- The Docs START Ribbon now follows HISTORY, OTHERS, FORMAT, LAYOUT group ordering. Find remains in its OTHERS semantic group. Checklist is in LAYOUT beside paragraph alignment and list operations.
- Canvas layout emits a hit-testable checklist rectangle. The document scene draws unchecked/checked state. Clicking the control dispatches `doc.command.toggle-checklist-item` through the normal document command and history path.
- DOCX import maps Duke `details.listLabel` values `☐` and `☑` to the checklist model while retaining parser source metadata. DOCX export writes native Word numbering definitions using those symbols and `Segoe UI Symbol`.
- Internal HTML snapshots group consecutive items by list ID and preserve checked state through `data-office-list-type`, `data-office-list-id`, and `data-office-checked` attributes.

## Verification

- `document-checklist.contract.test.mjs`, `document-command-boundaries.contract.test.mjs`, `document-toolbar-state.contract.test.mjs`, and `rendering/document-scene.contract.test.mjs`: 26/26 passed.
- `document-toolbar.browser.test.mjs`: 6/6 passed. It clicks the Ribbon, checks a Canvas item, verifies undo/redo, checks Enter behavior and the internal snapshot round trip.
- `node --experimental-strip-types --test packages/office-studio/tests/document-toolbar.browser.test.mjs`: 6/6 passed. It verifies Ribbon creation, Canvas pointer toggling, undo/redo, list continuation after Enter, and internal HTML snapshot roundtrip.
- `rustfmt --edition 2024 --check crates/office-writer/src/docx.rs` passed. Crate-wide `cargo fmt --check` reports formatting differences in the concurrent `src/xlsx_validation.rs` file, outside this DOCS subtask.
- `cargo test --manifest-path crates/office-writer/Cargo.toml`: 38/38 passed, including the checked/unchecked checklist numbering writer test.
- A production writer-to-parser DOCX roundtrip test is added to `tests/browser/production-office-io.spec.ts`. It has not run. `wasm-pack` successfully compiled `office-writer` to Wasm, then could not install the missing `wasm-bindgen` CLI because the environment rejected both its download and Cargo install temp directory. A workspace-local temp directory retry produced the same result. The parser/writer JS glue and browser roundtrip therefore remain unverified.
- `pnpm.cmd --filter @office-studio/office-studio check:types` passed.

## Remaining scope

This does not implement complete DOC-05. Nested list level changes, numbering continuation/restart semantics, list styles, reorder behavior, richer clipboard list structure, and real Office application reopen validation remain unverified. Keep DOC-05 marked “未验收”.
