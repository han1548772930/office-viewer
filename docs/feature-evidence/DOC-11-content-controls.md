# DOC-11 Content Controls: implementation evidence

Status: **未验收**. This record describes the current implementation slice; it does not claim full parity.

This is the earlier implementation snapshot. Current paragraph-split/lock fixes, panel browser coverage, type-check results and remaining gaps are recorded in the [2026-11-20 continuation](continuation-2026-11-20.md#content-controls); the earlier in-flight build and UI limitations below are historical findings.

## Reference contract

- Fixed source commit: `E:\rust_test\univer` at `6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68`.
- Source types: `packages/core/src/types/interfaces/i-document-data.ts`: `SdtKind`, `SdtLock`, `SdtPlacement`, `ISdtListItem`, `ISdtRangeProperties`, and `ISdtCustomRange`.
- Public Docs 1.5.2 guide: <https://docs.univer.ai/zh-CN/guides/docs/features/content-controls>.
- Public API references: <https://docs.univer.ai/zh-CN/reference/packages/plugins/univerjs-pro/docs-sdt> and <https://docs.univer.ai/zh-CN/reference/packages/plugins/univerjs-pro/docs-sdt-ui>.
- The public contract describes insert-at-selection; text, rich text, checkbox, date, list, picture, group, repeating section, and document-part controls; content/wrapper locks; and command-backed undo/redo. Repeating-section items are created by their section.

## Current model and command path

`DocumentModel.contentControls` stores serializable SDT properties plus paragraph-local UTF-16 half-open anchors. `DocumentContentControl` preserves type, placement, alias/title/tag, appearance/color, placeholder state, lock, checkbox/date/list data, parent links, repeating-section/doc-part properties, data-binding metadata, and raw OOXML property fields.

`features/document/services/content-controls.ts` validates range bounds and type-specific date/list values; permits nested ranges only with a valid parent relationship; rejects partial overlaps; enforces content and wrapper locks; adjusts local anchors on same-paragraph text edits; and updates text, checkbox, date, dropdown/combo-box values or removes a control tree. The public facade registers these operations as `doc.command.*`; `commitDocument` makes each accepted operation one document history entry. Document range edits also update or reject anchors through the shared service.

The current callable facade is `FDocument.addContentControl(...)`, `getContentControls()`, and `getContentControl(id)`. `FDocumentContentControl` exposes `getInfo`, `getText`, `getRange`, `setValue`, `setChecked`, `selectItem`, `update`, and `remove`.

## Verification and gaps

- `node --experimental-strip-types --test packages/office-studio/tests/content-controls.contract.test.mjs packages/office-studio/tests/document-facade.contract.test.mjs packages/office-studio/tests/office-api.contract.test.mjs`: 19/19 passed. The run printed a Vite WebSocket port-in-use warning; tests completed successfully.
- The dedicated contract covers value replacement and offsets, text-edit anchor movement, content/wrapper lock behavior, date/list validation, partial-overlap rejection, group-child removal, and undo/redo.
- DOCX writer currently emits SDT wrappers and common control properties; the writer task reports 28 Rust tests passing, including final `text`, `dropDownList`, `displayText`, nested group/repeating/doc-part, checkbox, and date spellings. An independently reopened DOCX sample is still required before claiming the export loop.
- UI contextual insertion/properties and canvas rendering are being implemented separately and have not yet been browser-verified here.
- Picture controls, rich-text control behavior, repeated-item insertion/deletion and nested structure editing, custom XML binding, multiline/multi-paragraph controls, cross-paragraph anchor transforms, import preservation, and the full DOCX round trip remain unverified or incomplete.
- A full `office-studio` type check is currently blocked by unrelated in-flight Sheets errors in `api/facade.ts` / `api/office-studio.ts` (missing range/filter types and helpers); there were no remaining diagnostics in the DOC-11 service/model files at the last run.

No UI, performance, or full Office round-trip completion is claimed by this record.
