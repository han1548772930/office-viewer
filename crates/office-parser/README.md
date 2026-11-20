# office-parser

`office-parser` is the Rust/WASM boundary used by the Vue workbench. It parses
real Office bytes and returns a versioned JSON DTO; the browser renderer does
not depend on internal Rust structs.

## WASM API

The crate exports three synchronous `wasm-bindgen` functions:

```ts
parse_office(bytes: Uint8Array, fileName: string): string
parse_excel(bytes: Uint8Array, fileName: string): string
parse_document(bytes: Uint8Array, fileName: string): string
```

Successful calls return an envelope with this shape (Rust fields are serialized
as camelCase):

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
value is in `payload.data`. On failure the synchronous call throws a JSON
string:

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
