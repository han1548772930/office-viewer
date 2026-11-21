# DOC-10 Thread Comments

Implementation checkpoint: 2026-11-21

## Implemented in this checkpoint

- `DocumentModel.comments` stores threads, paragraph-range anchor segments, reference text, messages, authors, timestamps, and resolved state.
- The document comment service creates anchored threads, replies, edits the current user's messages, toggles resolution, deletes threads, relocates anchors, and reports current, stale, or orphaned anchors.
- Comment mutations are stored in the document model and survive JSON snapshot serialization.
- Document-level `comment` permission rules are enforced before thread and message mutations.
- The Rust DOCX parser exposes comment messages, reply parents, resolution state, source paragraph targets, and UTF-16 offsets.
- The DOCX writer emits `comments.xml`, `commentsExtended.xml`, package relationships, and comment range/reference markers for mapped ranges.
- Single-paragraph text insertions, deletions, and replacements transform affected anchors and refresh their referenced text.
- Contract tests cover multi-paragraph anchors, replies, edit/resolve, JSON persistence, stale/orphan detection, permission rejection, and invalid empty input.
- Rust writer-to-parser round-trip verifies root/reply messages, resolution, and exported range offsets; a frontend contract verifies parser target normalization.

## Not verified or still missing

- UI creation, thread navigation, and rendering are not connected by this checkpoint.
- Cross-paragraph replacements, paragraph splits/merges, block deletion, and comment ranges overlapping inline content controls need anchor transformation or explicit unsupported-state handling. The writer emits a warning for content-control overlap.
- No service-backed multi-user synchronization or browser interaction has been tested.
- DOC-10 remains unaccepted pending full edit-anchor transformation, UI, browser, and service synchronization verification.

## Validation

- `node --experimental-strip-types --test packages/office-studio/tests/document-thread-comments.contract.test.mjs` passed.
- `node --experimental-strip-types --test packages/office-studio/tests/document-thread-comments-import.contract.test.mjs` passed.
- `cargo test --manifest-path crates/office-parser/Cargo.toml docx_comment_threads_and_range_anchors_roundtrip_through_the_writer` passed.
- `pnpm.cmd --filter @office-studio/office-studio check:types` passed after the parser import adjustments. A later workspace-level run reported unused identifiers in concurrently edited `DocumentWorkbench.vue` and `ExcelWorkbench.vue`.
