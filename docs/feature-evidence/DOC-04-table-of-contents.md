# DOC-04 Table of Contents Evidence

Reference revision: `dream-num/univer@eaf4cfa04fabdbcfc922c692f3450459f8a29546`
Review date: 2026-11-21

## Upstream source comparison

**Review scope:** Insert and update UI, menu registration, command dispatch, TOC model generation, layout metadata, selection context, localization package, and plugin registration in `docs-toc-ui` and `docs-toc`, plus the local document command, ribbon, and localization entry points.

The insert menu and selection guards are in `packages/docs-toc-ui/src/menu/menu.ts` and `menu/schema.ts`. The insertion entry is under the Insert tab (`ribbon.insert`), `ribbon.insert.cell` group, order 7. It is a submenu with automatic and custom insertion. The inspected UI operation accepts traditional documents and disables insertion for a non-collapsed selection or when the caret is already inside a TOC. The custom configuration offers 1 through 9 heading levels, four formats, page-number visibility, right alignment, and four tab leaders.

`packages/docs-toc-ui/src/commands/table-of-contents-dialog.operation.ts` opens the custom insertion and update dialogs. Update mode offers page-numbers-only and full-table updates. `TableOfContentsRibbonController` shows the contextual Ribbon tab `ribbon.table-of-contents` with group `ribbon.table-of-contents.manage` only while selection is within a TOC. The context menu exposes update and delete under the main-area Layout group, with locked TOCs disabling those actions.

The content model and mutations are implemented in `packages/docs-toc/src/commands/table-of-contents.command.ts`, with contract coverage in `packages/docs-toc/src/commands/__tests__/table-of-contents.command.spec.ts`. Entries use tab-separated page references, paragraph tab-stop leaders, heading hyperlinks, and page-reference fields. The TOC is represented as a field/custom range. This is more than a static list of heading strings.

The insertion settings view and update-mode view are `packages/docs-toc-ui/src/views/InsertTableOfContentsDialog.tsx` and `TableOfContentsDialog.tsx`. The package also includes render controllers for TOC highlighting and a contextual ribbon controller.

`packages/docs-toc-ui/src/plugin.ts` defines `UniverDocsTocUIPlugin` (`DOCS_TOC_UI_PLUGIN`) as a separate plugin dependent on Docs, render engine, Docs TOC, and Docs UI plugins. It registers dialog components, commands, menu schemas, the contextual Ribbon controller, and TOC render highlighting during plugin lifecycle hooks. This is an opt-in plugin extension, not a permanently resident Docs toolbar item. The searched presets and examples do not register this plugin at the reviewed commit, so availability depends on explicit application plugin registration.

## Local entry and command mapping

| Surface | Upstream | Local reviewed state |
| --- | --- | --- |
| Insert entry | Insert tab > `ribbon.insert.cell`, submenu `doc.menu.insert-table-of-contents`; operation `doc.operation.insert-table-of-contents` | Registered in `document-ribbon.ts` as Insert > CELL order 7 with automatic/custom actions. `DocumentWorkbench.vue` gates it to traditional body, collapsed selection outside an existing TOC, and editable permission. Automatic insertion uses the localized title. Custom insertion opens the antdv-next dialog. Both dispatch `buildDocumentTocBlocks()` via `doc.command.paste-blocks`. |
| Update entry | Contextual `ribbon.table-of-contents` > Manage; operation `doc.operation.open-table-of-contents-dialog` | `document-ribbon.ts` adds the contextual tab while a selection is contained in a TOC. Update opens the update-mode dialog; confirmation dispatches `doc.command.update-table-of-contents` with current layout-derived one-based heading page numbers. |
| Delete entry | Contextual ribbon and main-area Layout context menu; `doc.command.delete-table-of-contents` | Both entry points dispatch the existing model command with the selected TOC ID. The command validates that the selection is inside the target TOC and checks permissions for all TOC paragraphs. |
| Selection and flavor | Selection observables and traditional-document guards in `docs-toc-ui/src/menu/menu.ts` | `documentTocIdForSelection()` determines whether the selection stays inside a single TOC. `supportsDocumentTableOfContents()` defaults unspecified local flavor to traditional and hides/disables TOC for explicit modern flavor. |
| I18n | `docs-toc-ui` plugin locale namespace has 19 locale files in this revision | Feature-owned strings cover all 19 workbench locales, including insertion, update, delete, automatic title, preview and configuration options. |

## Local implementation

- `src/features/document/toc/document-toc.ts` collects eligible heading blocks, constructs paragraph-backed TOC rows with stable source heading IDs and configuration metadata, updates page numbers or rebuilds all rows, and removes every row for a TOC ID.
- `src/features/document/commands/document-commands.ts` exposes `doc.command.update-table-of-contents` and `doc.command.delete-table-of-contents`. Insertion uses the existing `doc.command.paste-blocks` command so normal document history owns the mutation.
- `src/components/document/DocumentTableOfContentsDialog.vue` provides insert settings and update mode controls using antdv-next. Its apply event is designed for workbench command dispatch.
- `src/features/document/toc/locales.ts` contains all 19 locale resources from the fixed upstream plugin locale set and is composed by `src/i18n/office-i18n.ts`.
- `tests/document-table-of-contents.contract.test.mjs` covers heading filtering, entry generation, page-number-only refresh, full rebuild, identity preservation, and deletion.

## Unverified and incomplete behavior

- Selection-driven contextual tab and Layout menu visibility are connected through current selection state. There is no TOC render highlight controller equivalent to the upstream highlight controller.
- The local paragraph layout has no tab-stop model or leader drawing. Current rows store page number after a tab, but the renderer treats tabs as ordinary spacing, so right alignment and leaders do not match upstream.
- The document model has no TOC field/custom range, PAGEREF field, or heading hyperlink target. Clicking a TOC entry does not navigate to its heading.
- The local paragraph model does not represent locked TOCs separately. Permission checks use existing paragraph/section permission rules.
- Automatic/custom insertion currently requires at least one eligible heading, matching the upstream no-headings guard.
- Current page numbers are supplied by the caller from layout and are snapshots. They do not update as calculated page fields, and no DOCX round-trip preservation has been verified.
- Permissions/locked TOC behavior and traditional-versus-modern visibility are not connected to an insertion/update UI.
- Browser interaction, undo/redo through the workbench, save/reopen, DOCX export/import, and localized UI review remain unverified.

## Verification

- `pnpm.cmd exec vitest run packages/office-studio/tests/document-table-of-contents.contract.test.mjs packages/office-studio/tests/docs-table-toolbar.contract.test.mjs`: 31 tests passed, including locale, insertion-through-command, and menu/command mapping coverage.
- `pnpm.cmd --filter @office-studio/office-studio check:types`: currently reports an unrelated shared-worktree error, `SpreadsheetCanvas.vue(204,37): Cannot find name 'columnName'`. Earlier type checks passed before concurrent spreadsheet edits.
- `git diff --check` on the DOC-04 evidence, TOC implementation, command, locale, and test files: passed.

Status remains **not accepted** until UI integration, rendering/navigation behavior, and document persistence are verified.
