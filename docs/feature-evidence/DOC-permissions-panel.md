# DOC Permissions Panel

Implementation checkpoint: 2026-11-21

## Source evidence

- The fixed Univer revision registers `OpenDocPermissionPanelOperation` in the Docs ribbon and exposes a separate paragraph-permission menu item from `E:\rust_test\univer\packages\docs-ui\src\menu\schema.ts` and `permission.menu.ts`.
- This project has a `DocumentPermissionRule` model and the `validateDocumentPermissionRules`, `resolveDocumentPermission`, and `assertDocumentPermission` services in `packages/office-studio/src/features/document/permissions/document-permissions.ts`.
- `DocumentPermissionPanel.vue` provides an antdv-next rule editor. `permission-rules.ts` provides immutable create, update, remove, and target-list helpers.

## Implemented in this checkpoint

- Document-level rules support edit, copy, comment, print, and export actions.
- Section, paragraph, and entity rules support edit, matching the current model validator and permission resolver.
- A rule may target all users or a normalized, de-duplicated list of user IDs.
- The panel receives a rule snapshot and target choices, then emits a complete updated snapshot through `save(rules)`.
- The component exposes `open`, `rules`, `targets`, and optional `readOnly` props. The host remains responsible for applying the emitted snapshot to its document transaction/history and persistence path.
- Locale strings are authored for all 19 currently supported locales in `packages/office-studio/src/i18n/docs/locales/permissions.ts`. The shared locale aggregator still needs to import this catalog.

## Not verified or still missing

- The panel is not yet mounted from the Docs Ribbon or Workbench.
- Host-side history, document dirty-state updates, DOCX persistence, remote synchronization, and server-side authorization have not been verified.
- The current rules model does not support scoped copy, comment, print, or export rules. Identity resolution is a user-ID list, not a group, role, or directory-backed model.
- No browser interaction or screen-reader verification has been run.
- DOC permissions remain unaccepted until the panel is integrated and the relevant edit/copy/print/export workflows are checked end to end.

## Validation

- Studio TypeScript check passed after adding the component and rule service.
- The panel contract test was created before test-runner migration was requested and has not been run in this checkpoint.
