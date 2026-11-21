# DOC-03 Section Break Evidence

Status: **Not accepted**. This evidence covers editable page, column, and section break markers. It does not establish full DOC-03 parity.

## Reference

- The local Univer Docs UI registers column breaks and five section break types: `nextPage`, `continuous`, `nextColumn`, `evenPage`, and `oddPage`.
- The local command operation dispatches section break type as an argument to `doc.command.insert-section-break`. The section update command applies an undoable document mutation and constrains insertion to the document body.
- The fixed reference commit could not be independently checked out from the local Univer repository during this task. These observations come from the locally readable source tree.

## Implemented behavior

- The Docs Insert Ribbon exposes page, column, and all five section break commands through nested menus.
- Section and column commands split the current body text selection and add a first-class marker block. Table-cell insertion is rejected before mutation. The existing document snapshot history records the change for undo and redo.
- Layout starts following content on the next page for `nextPage` and `nextColumn`. It skips a page when required by `evenPage` and `oddPage`. `continuous` retains the current page. This editor has no multi-column section layout model, so `nextColumn` currently advances to the next page.
- DOCX export writes `w:br` page and column break markers. Section markers write paragraph-level `w:sectPr/w:type` for each supported break type.
- DOCX import maps section start metadata and standalone page/column break paragraphs back to marker payloads. A manual break inside a paragraph with visible text remains a line break and emits a diagnostic warning.

## Verification

- `pnpm.cmd --filter @office-studio/office-studio check:types` passed.
- Document command, Ribbon, and Canvas layout contract tests passed, 29/29.
- The Docs browser test inserted an odd-page section break from the Ribbon and verified undo and redo, 1/1 focused test passed.
- `cargo test --manifest-path crates/office-parser/Cargo.toml` passed, 11/11.
- `cargo test --manifest-path crates/office-writer/Cargo.toml` passed, 42/42.

## Unverified and remaining scope

- Continuous sections do not yet carry section-specific page geometry, columns, headers, footers, page numbering, or link-to-previous behavior.
- Multi-column layout is not implemented. Consequently `nextColumn` cannot move to a remaining column.
- Breaks embedded among visible text in one paragraph are not split into separate editable blocks on import.
- Browser coverage verifies the odd-page Ribbon path. It does not exercise all break types or DOCX browser import/export round trips.
- A generated DOCX has not been reopened in desktop Office software. DOC-03 remains unaccepted.
