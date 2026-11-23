# office-parser

`office-parser` is the Rust/WASM boundary used by the Vue workbench. It parses
real Office bytes and returns a versioned structured DTO; the browser renderer does
not depend on internal Rust structs.

## WASM API

The crate exposes protocol 2 and three asynchronous `wasm-bindgen` functions:

```ts
parser_protocol_version(): 2
parse_office(bytes: Uint8Array, fileName: string, context: TaskControl): Promise<ParseResult>
parse_excel(bytes: Uint8Array, fileName: string, context: TaskControl): Promise<ParseResult>
parse_document(bytes: Uint8Array, fileName: string, context: TaskControl): Promise<ParseResult>
```

`TaskControl` contains `protocolVersion: 2`, `requestId`, `unitId`,
`baseRevision`, `kind: 'parse'`, and byte budgets
`{ maxInputBytes, maxOutputBytes, maxWorkingBytes }`.
`ParseResult` contains `{ context, envelope, transfer }`. `transfer` records
JS/WASM copies and DTO conversions; it does not claim end-to-end zero-copy.
Older generated artifacts are rejected by the Worker and must be rebuilt.

The structured `envelope` has this shape (shown as JSON for documentation;
Rust fields are serialized as camelCase):

```json
{
  "schemaVersion": 1,
  "parserVersion": "0.1.0",
  "format": "xlsx",
  "fileName": "book.xlsx",
  "size": 12345,
  "model": "spreadsheet",
  "payload": {
    "modelType": "spreadsheet",
    "data": { "sheets": [] }
  },
  "diagnostics": {
    "warnings": [],
    "errors": [],
    "stats": { "sheetCount": 1, "cellCount": 12 },
    "capabilities": ["typedCells", "formulas", "styles"]
  }
}
```

`payload.modelType` is `spreadsheet` or `document`, and the corresponding
value is in `payload.data`. On failure the Promise rejects with a structured
object:

```json
{
  "ok": false,
  "error": {
    "code": "INVALID_OFFICE_FILE",
    "message": "Invalid or corrupted office file",
    "details": "..."
  }
}
```

The parser yields through the shared browser/Worker timer before entering the
synchronous parsing kernel. Cancellation during the kernel terminates and
recreates the Worker; an asynchronous binding alone does not make it cooperative.

The spreadsheet DTO uses zero-based row and column coordinates. It retains the
typed value, `dataType`, formula source, cached formula type, formatted display
text, resolved cell style, rich-text runs, hyperlinks, merged ranges, row and
column geometry, hidden axes, freeze panes, view settings, defined names,
workbook properties, and style-fidelity/loss diagnostics. Formula text and its
cached value are kept separately; no formula is evaluated in the parser.

`duke-pages` reads `.doc`, `.docx`, and `.odt` into structured blocks containing
paragraphs, inline runs, tables, images, and source references. `rxls` reads
`.xls`, `.xlsx`, `.xlsb`, and `.ods` with the `full` feature set. The parser
does not promise Word/Excel pixel-identical layout or every vendor extension.
`diagnostics.capabilities` lists modeled parser features, and each worksheet
reports the style fidelity and style losses exposed by `rxls`. These fields do
not enumerate every source feature that may be omitted.

## Build

Run from the repository root:

```powershell
pnpm build:wasm
```

The generated package is written to `public/wasm` for Vite to load. For direct
Rust validation:

```powershell
cargo check --manifest-path crates/office-parser/Cargo.toml
cargo check --manifest-path crates/office-parser/Cargo.toml --target wasm32-unknown-unknown
cargo test --manifest-path crates/office-parser/Cargo.toml
```
