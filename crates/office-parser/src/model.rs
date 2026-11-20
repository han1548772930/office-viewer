use duke_pages::StructuredDocument;
use serde::Serialize;
use serde_json::Value;

pub const SCHEMA_VERSION: u32 = 1;

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ParseEnvelope {
    pub schema_version: u32,
    pub parser_version: String,
    pub format: String,
    pub file_name: String,
    pub size: usize,
    pub model: String,
    pub payload: DocumentModel,
    pub diagnostics: ParserDiagnostics,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase", tag = "modelType", content = "data")]
pub enum DocumentModel {
    Spreadsheet(WorkbookPayload),
    Document(StructuredDocument),
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ParserDiagnostics {
    pub warnings: Vec<String>,
    pub errors: Vec<ParseError>,
    pub stats: ParseStats,
    pub capabilities: Vec<String>,
}

impl ParserDiagnostics {
    pub fn spreadsheet(workbook: &WorkbookPayload, warnings: Vec<String>) -> Self {
        let cell_count = workbook.sheets.iter().map(|sheet| sheet.cells.len()).sum();
        let formula_count = workbook
            .sheets
            .iter()
            .flat_map(|sheet| sheet.cells.iter())
            .filter(|cell| cell.formula.is_some())
            .count();
        let merge_count = workbook
            .sheets
            .iter()
            .map(|sheet| sheet.merged_ranges.len())
            .sum();
        Self {
            warnings,
            errors: Vec::new(),
            stats: ParseStats {
                sheet_count: Some(workbook.sheets.len()),
                cell_count: Some(cell_count),
                block_count: None,
                formula_count: Some(formula_count),
                merge_count: Some(merge_count),
            },
            capabilities: vec![
                "typedCells".into(),
                "formulas".into(),
                "styles".into(),
                "merges".into(),
                "rowColumnGeometry".into(),
                "sheetViews".into(),
            ],
        }
    }

    pub fn document(document: &StructuredDocument) -> Self {
        let block_count = document.blocks.iter().map(count_blocks).sum();
        Self {
            warnings: document.warnings.clone(),
            errors: Vec::new(),
            stats: ParseStats {
                sheet_count: None,
                cell_count: None,
                block_count: Some(block_count),
                formula_count: None,
                merge_count: None,
            },
            capabilities: vec![
                "paragraphs".into(),
                "inlineRuns".into(),
                "tables".into(),
                "images".into(),
                "sourceReferences".into(),
            ],
        }
    }
}

fn count_blocks(block: &duke_pages::DocumentBlock) -> usize {
    1 + block.children.iter().map(count_blocks).sum::<usize>()
}

#[derive(Debug, Default, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ParseStats {
    pub sheet_count: Option<usize>,
    pub cell_count: Option<usize>,
    pub block_count: Option<usize>,
    pub formula_count: Option<usize>,
    pub merge_count: Option<usize>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ParseError {
    pub code: ParseErrorCode,
    pub message: String,
    pub details: Option<String>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "SCREAMING_SNAKE_CASE")]
pub enum ParseErrorCode {
    InvalidInput,
    UnsupportedFormat,
    InvalidOfficeFile,
    ParserFailure,
    SerializationFailure,
}

#[derive(Debug)]
pub struct ParseFailure {
    pub error: ParseError,
}

impl ParseFailure {
    pub fn new(code: ParseErrorCode, message: impl Into<String>, details: Option<String>) -> Self {
        Self {
            error: ParseError {
                code,
                message: message.into(),
                details,
            },
        }
    }

    pub fn invalid_input(message: impl Into<String>) -> Self {
        Self::new(ParseErrorCode::InvalidInput, message, None)
    }

    pub fn unsupported(message: impl Into<String>) -> Self {
        Self::new(ParseErrorCode::UnsupportedFormat, message, None)
    }

    pub fn office_file(error: impl Into<String>) -> Self {
        Self::new(
            ParseErrorCode::InvalidOfficeFile,
            "Invalid or corrupted office file",
            Some(error.into()),
        )
    }

    pub fn serialization(error: impl Into<String>) -> Self {
        Self::new(
            ParseErrorCode::SerializationFailure,
            "Could not serialize parser result",
            Some(error.into()),
        )
    }
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WorkbookPayload {
    pub sheets: Vec<SheetPayload>,
    pub date1904: bool,
    pub active_sheet: Option<usize>,
    pub defined_names: Vec<DefinedNamePayload>,
    pub local_defined_names: Vec<LocalDefinedNamePayload>,
    pub properties: WorkbookPropertiesPayload,
    pub structure_protected: bool,
    pub text_truncated: bool,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct DefinedNamePayload {
    pub name: String,
    pub refers_to: String,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct LocalDefinedNamePayload {
    pub sheet: String,
    pub name: String,
    pub refers_to: String,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WorkbookPropertiesPayload {
    pub title: Option<String>,
    pub subject: Option<String>,
    pub creator: Option<String>,
    pub keywords: Option<String>,
    pub description: Option<String>,
    pub last_modified_by: Option<String>,
    pub company: Option<String>,
    pub created: Option<String>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct SheetPayload {
    pub name: String,
    pub is_worksheet: bool,
    pub visible: String,
    pub cells: Vec<CellPayload>,
    pub dimensions: Option<DimensionPayload>,
    pub merged_ranges: Vec<RangePayload>,
    pub column_widths: Vec<AxisPayload>,
    pub row_heights: Vec<AxisPayload>,
    pub default_column_width: Option<f32>,
    pub default_row_height: Option<f32>,
    pub gridlines: Option<bool>,
    pub freeze_panes: Option<FreezePanePayload>,
    pub zoom: Option<u16>,
    pub show_headers: Option<bool>,
    pub right_to_left: bool,
    pub hidden_rows: Vec<u32>,
    pub hidden_columns: Vec<u16>,
    pub hyperlinks: Vec<HyperlinkPayload>,
    pub style_fidelity: String,
    pub style_losses: Vec<StyleLossPayload>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct CellPayload {
    pub row: u32,
    pub col: u16,
    pub value: Value,
    pub data_type: String,
    pub cached_data_type: Option<String>,
    pub display: String,
    pub formula: Option<String>,
    pub style: Option<StylePayload>,
    pub rich_text: Vec<RichTextRunPayload>,
    pub hyperlink: Option<String>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct StylePayload {
    pub font: Option<FontPayload>,
    pub fill: Option<String>,
    pub number_format: Option<String>,
    pub horizontal: Option<String>,
    pub vertical: Option<String>,
    pub wrap: bool,
    pub rotation: i16,
    pub indent: u8,
    pub shrink_to_fit: bool,
    pub fill_pattern: Option<String>,
    pub fill_background: Option<String>,
    pub fill_foreground: Option<String>,
    pub border: Option<BorderPayload>,
    pub protection: Option<ProtectionPayload>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct FontPayload {
    pub name: Option<String>,
    pub size_pt: Option<u16>,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strikethrough: bool,
    pub color: Option<String>,
    pub script: Option<String>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct RichTextRunPayload {
    pub text: String,
    pub font: FontPayload,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct BorderPayload {
    pub left: String,
    pub right: String,
    pub top: String,
    pub bottom: String,
    pub color: Option<String>,
    pub left_color: Option<String>,
    pub right_color: Option<String>,
    pub top_color: Option<String>,
    pub bottom_color: Option<String>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct FreezePanePayload {
    pub rows: u32,
    pub columns: u16,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ProtectionPayload {
    pub locked: Option<bool>,
    pub hidden: bool,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct HyperlinkPayload {
    pub row: u32,
    pub col: u16,
    pub target: String,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct StyleLossPayload {
    pub kind: String,
    pub occurrences: u32,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct DimensionPayload {
    pub first_row: u32,
    pub first_col: u16,
    pub last_row: u32,
    pub last_col: u16,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct RangePayload {
    pub first_row: u32,
    pub first_col: u16,
    pub last_row: u32,
    pub last_col: u16,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct AxisPayload {
    pub index: u32,
    pub size: f32,
}

#[cfg(test)]
mod tests {
    use super::{DocumentModel, WorkbookPayload, WorkbookPropertiesPayload};

    #[test]
    fn spreadsheet_payload_uses_tagged_model_type_and_data() {
        let payload = DocumentModel::Spreadsheet(WorkbookPayload {
            sheets: Vec::new(),
            date1904: false,
            active_sheet: None,
            defined_names: Vec::new(),
            local_defined_names: Vec::new(),
            properties: WorkbookPropertiesPayload {
                title: None,
                subject: None,
                creator: None,
                keywords: None,
                description: None,
                last_modified_by: None,
                company: None,
                created: None,
            },
            structure_protected: false,
            text_truncated: false,
        });

        let json = serde_json::to_value(payload).expect("model must serialize");
        assert_eq!(json["modelType"], "spreadsheet");
        assert!(json["data"].is_object());
        assert_eq!(json["data"]["date1904"], false);
    }
}
