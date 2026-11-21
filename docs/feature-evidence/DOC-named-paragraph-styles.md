# Docs named paragraph styles

Reference baseline: `dream-num/univer@6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68`.

The upstream Docs paragraph menu exposes Normal, Title, Subtitle, and Heading 1 through Heading 5. Its `SetParagraphNamedStyleCommand` stores the named style separately from outline level. The fixed source defines run defaults of 26 pt bold for Title, 15 pt gray for Subtitle, and 20/18/16/14/12 pt bold for Heading 1 through 5.

This project now has a model command, `doc.command.set-paragraph-named-style`, which stores a stable style key in `DocumentBlock.metadata.namedStyleType`, applies run formatting and paragraph spacing, and maps Heading 1 through 5 to outline levels. Normal, Title, and Subtitle remain ordinary paragraphs. JSON model snapshots retain the style key.

Remaining gaps:

- The Docs Ribbon still dispatches `doc.command.set-block-type` and does not expose the named styles.
- The document normalizer does not yet map imported OOXML named styles into `metadata.namedStyleType`.
- The Rust DOCX writer/parser do not yet preserve Title and Subtitle named styles on export/import. Existing heading export only recognizes heading blocks.
- Applying a named style rewrites run font size, bold, and default color because the current run model does not distinguish inherited style from explicit character formatting.
- The contract tests cover command semantics and JSON snapshots. Browser interaction and DOCX round trips remain unverified.

Verification: `node --experimental-strip-types --test packages/office-studio/tests/document-named-paragraph-style.contract.test.mjs` passed 3/3; Studio type check passed during this task; `git diff --check` passed for the edited implementation and test files.

Status: not accepted. The toolbar entry, style-aware OOXML import/export, and browser/file sample validation remain outstanding.
