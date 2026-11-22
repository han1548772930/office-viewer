# DOC-03 Section Break Evidence

Status: **Not accepted**. This evidence covers editable page, column, and section break markers. It does not establish full DOC-03 parity.

## Reference

### Current working tree, not yet tested

The following supersedes the earlier modal-only observation below. The traditional Docs workbench now connects the Insert entry and Canvas story hit areas to a page/section/variant-specific antdv-next input bridge. The body selection is retained separately. The contextual tab has four inline control groups and Close. Responsive grouping uses the toolbar-width thresholds in `packages/docs-ui/src/views/header-footer/panel/DocHeaderFooterRibbonControl.tsx`. Insert/refresh date and time uses the locale-derived masks in upstream `utils/date-time-field.ts`. Positioned fields are represented independently of literal UTF-16 text, including repeated fields and field cached results. DOCX import reads the referenced story XML, and export places simple fields between literal runs. Contracts for field deletion, Unicode, distance rounding, writer/reader reopening and a browser editing path have been added, not executed. The WASM artifact has not been rebuilt for this change.

Still unverified/incomplete: classic and grid width/overflow, localized display, permissions and IME across focus transitions, rendering rich multiline stories, unresolved nested field instructions and preservation of unknown story XML. The text bridge denotes an editable field by an object replacement character; it is not rich field rendering. Manual refresh currently handles recognized date/time masks in the selected section. This work does not establish DOC-03 acceptance.

### Earlier implementation notes

2026-11-22 (implementation, unified tests pending): v1.6.0 `docs-ui/src/menu/header-footer-ribbon.ts` registers four distinct control groups plus Close; local contextual schema now separates them rather than grouping five items under Insert. The running workbench still opens the header/footer modal, not an in-Canvas story-editing session. `views/header-footer/panel/use-header-footer-options.ts` and its tests scope header/footer distances to the current section. Local InputNumber controls write `marginHeader`/`marginFooter` through `doc.command.set-section-setup`, independently of story linking. Canvas uses the stored distances for text placement. DOCX reader/writer map section `w:pgMar` header/footer twips to/from CSS pixels, retaining the existing writer default of 720 twips (48px) when unspecified. Text/field changes no longer request body focus on each keystroke; closing the modal requests it. No new tests or WASM builds have been run for this change; real file reopening, IME, field positioning and contextual activation remain unaccepted.

- The local Univer Docs UI registers column breaks and five section break types: `nextPage`, `continuous`, `nextColumn`, `evenPage`, and `oddPage`.
- The local command operation dispatches section break type as an argument to `doc.command.insert-section-break`. The section update command applies an undoable document mutation and constrains insertion to the document body.
- The fixed reference commit could not be independently checked out from the local Univer repository during this task. These observations come from the locally readable source tree.

## Implemented behavior

- The Docs Insert Ribbon exposes page, column, and all five section break commands through nested menus.
- Section and column commands split the current body text selection and add a first-class marker block. Table-cell insertion is rejected before mutation. The existing document snapshot history records the change for undo and redo.
- Layout starts following content on the next page for `nextPage` and `nextColumn`. It skips a page when required by `evenPage` and `oddPage`. `continuous` retains the current page. This editor has no multi-column section layout model, so `nextColumn` currently advances to the next page.
- DOCX export writes `w:br` page and column break markers. Section markers write paragraph-level `w:sectPr/w:type` for each supported break type.
- DOCX import maps section start metadata and standalone page/column break paragraphs back to marker payloads. A manual break inside a paragraph with visible text remains a line break and emits a diagnostic warning.
- Header/footer models now support structured `PAGE` and `NUMPAGES` fields. The editor recalculates their displayed values from the current page layout. DOCX export writes `w:fldSimple` instructions and import preserves the field names on the matching header/footer reference.

## Verification

The November 22 position changes are excluded from the historical test counts below. The Rust parser now extracts the current section's header/footer distances, skipping historical `sectPrChange` properties. The DTO and document normalizer retain zero and missing values distinctly. Deferred contracts cover XML extraction, section-local layout with inherited content, and native writer XML with fractional pixel rounding. These are not evidence of a completed Office round trip.

- `pnpm.cmd --filter @office-studio/office-studio check:types` passed.
- Document command, Ribbon, and Canvas layout contract tests passed, 29/29.
- The Docs browser test inserted an odd-page section break from the Ribbon and verified undo and redo, 1/1 focused test passed.
- `cargo test --manifest-path crates/office-parser/Cargo.toml` passed, 11/11.
- `cargo test --manifest-path crates/office-writer/Cargo.toml` passed, 42/42.
- The page-field extension adds a layout contract proving current page and total pages update after a page break is inserted. Writer/parser DOCX contract verifies both field instructions survive export and parse.

## Unverified and remaining scope

- Continuous sections do not yet carry section-specific page geometry or columns. Header/footer content, page fields, and link-to-previous are modeled per section, but numbering formats, section-specific restart/continue settings, and page-field switches are not implemented.
- Multi-column layout is not implemented. Consequently `nextColumn` cannot move to a remaining column.
- Breaks embedded among visible text in one paragraph are not split into separate editable blocks on import.
- Browser coverage verifies the odd-page Ribbon path. It does not exercise all break types or DOCX browser import/export round trips.
- A generated DOCX has not been reopened in desktop Office software. The page-field export leaves the cached result empty and requests Word field refresh on open; the editor independently renders values from its current pagination. DOC-03 remains unaccepted.
