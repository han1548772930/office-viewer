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
    Document(DocumentPayload),
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct DocumentPayload {
    #[serde(flatten)]
    pub document: StructuredDocument,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub numbering_definitions: Option<duke_pages::NumberingDefinitionSet>,
    pub section_layout: Vec<DocumentSectionLayoutPayload>,
    pub even_and_odd_headers: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub footnotes: Option<Vec<DocumentNotePayload>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub endnotes: Option<Vec<DocumentNotePayload>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub comments: Option<Vec<DocumentCommentPayload>>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct DocumentNotePayload {
    pub id: String,
    pub text: String,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct DocumentCommentPayload {
    pub id: String,
    pub author: Option<String>,
    pub date: Option<String>,
    pub text: String,
    pub parent_id: Option<String>,
    pub resolved: Option<bool>,
    pub targets: Vec<DocumentCommentTargetPayload>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct DocumentCommentTargetPayload {
    pub node_id: String,
    pub part_name: String,
    pub story: String,
    pub start: Option<usize>,
    pub end: Option<usize>,
    pub offset_encoding: String,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct DocumentSectionLayoutPayload {
    pub margin_header: Option<f64>,
    pub margin_footer: Option<f64>,
    pub start_node_index: i64,
    pub section_properties_xml: Option<String>,
    pub header_sections: Vec<DocumentStoryReferencePayload>,
    pub footer_sections: Vec<DocumentStoryReferencePayload>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct DocumentStoryReferencePayload {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub story_text: Option<crate::story_fields::StoryText>,
    pub part_name: String,
    pub reference_type: Option<String>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub page_fields: Vec<String>,
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
        let mut diagnostics = Self {
            warnings,
            errors: Vec::new(),
            stats: ParseStats {
                sheet_count: Some(workbook.sheets.len()),
                cell_count: Some(cell_count),
                block_count: None,
                formula_count: Some(formula_count),
                merge_count: Some(merge_count),
                chart_count: Some(workbook.sheets.iter().map(|sheet| sheet.charts.len()).sum()),
            },
            capabilities: vec![
                "typedCells".into(),
                "formulas".into(),
                "styles".into(),
                "merges".into(),
                "rowColumnGeometry".into(),
                "sheetViews".into(),
                "charts".into(),
                "dataValidations".into(),
            ],
        };
        if workbook.sheets.iter().flat_map(|sheet| &sheet.cells).any(|cell| cell.sparkline.is_some()) {
            diagnostics.capabilities.push("sparklines".into());
        }
        if !workbook.pivot_tables.is_empty() {
            diagnostics.capabilities.push("pivotTables".into());
        }
        if !workbook.tables.is_empty() {
            diagnostics.capabilities.push("structuredTables".into());
        }
        if workbook.sheets.iter().any(|sheet| !sheet.conditional_formats.is_empty()) {
            diagnostics.capabilities.push("conditionalFormatting".into());
        }
        if workbook.sheets.iter().flat_map(|sheet| &sheet.charts).any(|chart| chart.pivot_table_id.is_some()) {
            diagnostics.capabilities.push("pivotCharts".into());
        }
        diagnostics
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
                chart_count: None,
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
    pub chart_count: Option<usize>,
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
    pub pivot_tables: Vec<Value>,
    pub tables: Vec<Value>,
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
    pub data_validations: Vec<DataValidationPayload>,
    pub conditional_formats: Vec<Value>,
    pub style_fidelity: String,
    pub style_losses: Vec<StyleLossPayload>,
    /// Charts read from the worksheet drawing part. Unsupported chart
    /// constructs remain in the sidecar metadata rather than being dropped.
    pub charts: Vec<ChartPayload>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct DataValidationPayload {
    pub id: String,
    #[serde(rename = "type")]
    pub validation_type: String,
    pub ranges: Vec<DataValidationRangePayload>,
    pub operator: String,
    pub formula1: String,
    pub formula2: Option<String>,
    pub allow_blank: bool,
    pub show_drop_down: Option<bool>,
    pub show_error_message: bool,
    pub error_style: Option<String>,
    pub error_title: Option<String>,
    pub error: Option<String>,
    pub show_input_message: bool,
    pub prompt_title: Option<String>,
    pub prompt: Option<String>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct DataValidationRangePayload {
    pub start_row: u32,
    pub start_col: u32,
    pub end_row: u32,
    pub end_col: u32,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ChartPayload {
    pub index: usize,
    pub kind: String,
    pub title: Option<String>,
    pub series: Vec<ChartSeriesPayload>,
    pub legend: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub legend_position: Option<String>,
    pub data_labels: bool,
    pub x_axis_title: Option<String>,
    pub y_axis_title: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub x_axis: Option<ChartAxisPayload>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub y_axis: Option<ChartAxisPayload>,
    pub from: AnchorCellPayload,
    pub to: AnchorCellPayload,
    pub drawing: Option<ChartDrawingPayload>,
    pub pivot_table_id: Option<String>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ChartSeriesPayload {
    pub name: Option<String>,
    pub categories: Option<String>,
    pub values: String,
    pub bubble_sizes: Option<String>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ChartAxisPayload {
    pub scale_type: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub log_base: Option<f64>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct AnchorCellPayload {
    pub row: u32,
    pub col: u16,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ChartDrawingPayload {
    pub object_index: usize,
    pub from_cell: Option<AnchorCellPayload>,
    pub to_cell: Option<AnchorCellPayload>,
    pub from_offset_emu: Option<[i64; 2]>,
    pub to_offset_emu: Option<[i64; 2]>,
    pub absolute_size_emu: Option<[u64; 2]>,
    pub crop: Option<[u32; 4]>,
    pub rotation_mdeg: Option<i32>,
    pub z_order: Option<i32>,
    pub alt_text: Option<String>,
    pub name: Option<String>,
    pub behavior: String,
    pub chart_palette: Vec<String>,
    pub chart_default_latin_font_family: Option<String>,
    pub chart_category_axis_visible: Option<bool>,
    pub chart_category_axis_shifted: Option<bool>,
    pub chart_value_axis_visible: Option<bool>,
    pub chart_category_major_gridlines: Option<bool>,
    pub chart_value_major_gridlines: Option<bool>,
    pub chart_bar_direction: String,
    pub chart_frame_fill: String,
    pub chart_frame_style_losses: Vec<String>,
    pub chart_unsupported_reasons: Vec<String>,
    pub chart_series_caches: Vec<ChartSeriesCachePayload>,
    pub chart_series_styles: Vec<ChartSeriesStylePayload>,
    /// Text styles use a stable JSON shape because rxls intentionally keeps
    /// this sidecar non-exhaustive for renderer-specific consumers.
    pub chart_text_styles: Value,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ChartSeriesCachePayload {
    pub name: Vec<ChartCachedPointPayload>,
    pub categories: Vec<ChartCachedPointPayload>,
    pub values: Vec<ChartCachedPointPayload>,
    pub bubble_sizes: Vec<ChartCachedPointPayload>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ChartCachedPointPayload {
    pub index: u32,
    pub value: String,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ChartSeriesStylePayload {
    pub marker: String,
    pub marker_size: Option<u8>,
    pub line_visible: bool,
    pub line_color: Option<String>,
    pub line_width_emu: Option<u32>,
    pub losses: Vec<String>,
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
    #[serde(skip_serializing_if = "Option::is_none")]
    pub sparkline: Option<SparklinePayload>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct SparklinePayload {
    #[serde(rename = "type")]
    pub sparkline_type: String,
    pub source_range: String,
    pub color: String,
    pub negative_color: Option<String>,
    pub high_point_color: Option<String>,
    pub low_point_color: Option<String>,
    pub empty_cells: String,
    pub axis: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub min: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub max: Option<f64>,
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
    pub rotation_vertical: bool,
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
            pivot_tables: Vec::new(),
            tables: Vec::new(),
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
