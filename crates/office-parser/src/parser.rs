use crate::model::{
    AxisPayload, BorderPayload, CellPayload, DimensionPayload, DocumentModel, FontPayload,
    FreezePanePayload, HyperlinkPayload, LocalDefinedNamePayload, ParseEnvelope, ParseFailure,
    ParserDiagnostics, ProtectionPayload, RangePayload, RichTextRunPayload, SCHEMA_VERSION,
    SheetPayload, StyleLossPayload, StylePayload, WorkbookPayload, WorkbookPropertiesPayload,
};
use duke_pages::parse_document_blocks;
use rxls::{Cell, Workbook};
use serde_json::{Value, json};

pub fn parse_office_bytes(bytes: &[u8], file_name: &str) -> Result<ParseEnvelope, ParseFailure> {
    validate_input(bytes, file_name)?;
    let format = detect_format(file_name, bytes);
    match format.as_str() {
        "doc" | "docx" | "odt" => parse_document_impl(bytes, file_name, format),
        "xls" | "xlsx" | "xlsb" | "ods" => parse_workbook(bytes, file_name, format),
        _ => Err(ParseFailure::unsupported(
            "Supported formats are .doc, .docx, .odt, .xls, .xlsx, .xlsb and .ods",
        )),
    }
}

pub fn parse_excel_bytes(bytes: &[u8], file_name: &str) -> Result<ParseEnvelope, ParseFailure> {
    validate_input(bytes, file_name)?;
    let format = detect_format(file_name, bytes);
    if !matches!(format.as_str(), "xls" | "xlsx" | "xlsb" | "ods") {
        return Err(ParseFailure::unsupported(
            "Expected an Excel workbook (.xls, .xlsx, .xlsb or .ods)",
        ));
    }
    parse_workbook(bytes, file_name, format)
}

pub fn parse_document_bytes(bytes: &[u8], file_name: &str) -> Result<ParseEnvelope, ParseFailure> {
    validate_input(bytes, file_name)?;
    let format = detect_format(file_name, bytes);
    if !matches!(format.as_str(), "doc" | "docx" | "odt") {
        return Err(ParseFailure::unsupported(
            "Expected a Word document (.doc, .docx or .odt)",
        ));
    }
    parse_document_impl(bytes, file_name, format)
}

fn parse_document_impl(
    bytes: &[u8],
    file_name: &str,
    format: String,
) -> Result<ParseEnvelope, ParseFailure> {
    let document = parse_document_blocks(bytes).map_err(ParseFailure::office_file)?;
    let diagnostics = ParserDiagnostics::document(&document);
    Ok(ParseEnvelope {
        schema_version: SCHEMA_VERSION,
        parser_version: env!("CARGO_PKG_VERSION").to_owned(),
        format,
        file_name: file_name.to_owned(),
        size: bytes.len(),
        model: "document".into(),
        payload: DocumentModel::Document(document),
        diagnostics,
    })
}

fn parse_workbook(
    bytes: &[u8],
    file_name: &str,
    format: String,
) -> Result<ParseEnvelope, ParseFailure> {
    let workbook =
        Workbook::open(bytes).map_err(|error| ParseFailure::office_file(error.to_string()))?;
    let metadata = workbook.metadata();
    let sheets = workbook
        .sheets
        .iter()
        .map(sheet_payload)
        .collect::<Vec<_>>();
    let payload = WorkbookPayload {
        sheets,
        date1904: workbook.date1904,
        active_sheet: metadata.active_sheet,
        defined_names: workbook
            .defined_names
            .iter()
            .map(|(name, refers_to)| crate::model::DefinedNamePayload {
                name: name.clone(),
                refers_to: refers_to.clone(),
            })
            .collect(),
        local_defined_names: metadata
            .local_defined_names
            .iter()
            .map(|name| LocalDefinedNamePayload {
                sheet: name.sheet.clone(),
                name: name.name.clone(),
                refers_to: name.refers_to.clone(),
            })
            .collect(),
        properties: WorkbookPropertiesPayload {
            title: metadata.properties.title.clone(),
            subject: metadata.properties.subject.clone(),
            creator: metadata.properties.creator.clone(),
            keywords: metadata.properties.keywords.clone(),
            description: metadata.properties.description.clone(),
            last_modified_by: metadata.properties.last_modified_by.clone(),
            company: metadata.properties.company.clone(),
            created: metadata.properties.created.clone(),
        },
        structure_protected: metadata.structure_protected,
        text_truncated: metadata.text_truncated,
    };
    let warnings = if workbook.text_truncated {
        vec!["Some text cells were omitted after the parser allocation limit".into()]
    } else {
        Vec::new()
    };
    let diagnostics = ParserDiagnostics::spreadsheet(&payload, warnings);
    Ok(ParseEnvelope {
        schema_version: SCHEMA_VERSION,
        parser_version: env!("CARGO_PKG_VERSION").to_owned(),
        format,
        file_name: file_name.to_owned(),
        size: bytes.len(),
        model: "spreadsheet".into(),
        payload: DocumentModel::Spreadsheet(payload),
        diagnostics,
    })
}

fn sheet_payload(sheet: &rxls::Sheet) -> SheetPayload {
    let metadata = sheet.metadata();
    let cells = sheet
        .display_cells()
        .map(|cell| {
            let (value, formula) = cell_value(cell.value);
            CellPayload {
                row: cell.row,
                col: cell.col,
                value,
                data_type: cell_type(cell.value).into(),
                cached_data_type: match cell.value {
                    Cell::Formula { cached, .. } => Some(cell_type(cached).into()),
                    _ => None,
                },
                display: cell.formatted.to_owned(),
                formula,
                style: sheet
                    .resolved_cell_style(cell.row, cell.col)
                    .as_ref()
                    .map(style_payload),
                rich_text: cell
                    .rich_text
                    .map(|runs| {
                        runs.iter()
                            .map(|run| RichTextRunPayload {
                                text: run.text.clone(),
                                font: font_payload(&run.font),
                            })
                            .collect()
                    })
                    .unwrap_or_default(),
                hyperlink: cell.hyperlink.map(str::to_owned),
            }
        })
        .collect();
    let dimensions =
        metadata.dimensions.map(
            |(first_row, first_col, last_row, last_col)| DimensionPayload {
                first_row,
                first_col,
                last_row,
                last_col,
            },
        );
    let merged_ranges = metadata
        .merged_ranges
        .iter()
        .map(|&(first_row, first_col, last_row, last_col)| RangePayload {
            first_row,
            first_col,
            last_row,
            last_col,
        })
        .collect();
    let column_widths = sheet
        .column_widths()
        .iter()
        .map(|(&index, &size)| AxisPayload {
            index: index as u32,
            size,
        })
        .collect();
    let row_heights = sheet
        .row_heights()
        .iter()
        .map(|(&index, &size)| AxisPayload { index, size })
        .collect();
    let view = metadata.sheet_view;
    SheetPayload {
        name: sheet.name.clone(),
        is_worksheet: sheet.is_worksheet,
        visible: format!("{:?}", metadata.visible).to_ascii_lowercase(),
        cells,
        dimensions,
        merged_ranges,
        column_widths,
        row_heights,
        default_column_width: sheet.default_column_width(),
        default_row_height: sheet.default_row_height(),
        gridlines: Some(!view.hide_gridlines),
        freeze_panes: view
            .freeze
            .map(|(rows, columns)| FreezePanePayload { rows, columns }),
        zoom: view.zoom,
        show_headers: view.show_headers,
        right_to_left: view.right_to_left,
        hidden_rows: sheet.hidden_rows().iter().copied().collect(),
        hidden_columns: sheet.hidden_columns().iter().copied().collect(),
        hyperlinks: metadata
            .hyperlinks
            .iter()
            .map(|&(row, col, ref target)| HyperlinkPayload {
                row,
                col,
                target: target.clone(),
            })
            .collect(),
        style_fidelity: format!("{:?}", sheet.style_fidelity()).to_ascii_lowercase(),
        style_losses: sheet
            .style_losses()
            .iter()
            .map(|loss| StyleLossPayload {
                kind: format!("{:?}", loss.kind).to_ascii_lowercase(),
                occurrences: loss.occurrences,
            })
            .collect(),
    }
}

fn cell_value(cell: &Cell) -> (Value, Option<String>) {
    match cell {
        Cell::Text(value) => (Value::String(value.clone()), None),
        Cell::Number(value) | Cell::Date(value) => (json!(value), None),
        Cell::Bool(value) => (json!(value), None),
        Cell::Error(value) => (json!({ "error": value }), None),
        Cell::Formula { formula, cached } => {
            let (value, _) = cell_value(cached);
            (value, Some(formula.clone()))
        }
    }
}

fn cell_type(cell: &Cell) -> &'static str {
    match cell {
        Cell::Text(_) => "text",
        Cell::Number(_) => "number",
        Cell::Date(_) => "date",
        Cell::Bool(_) => "boolean",
        Cell::Error(_) => "error",
        Cell::Formula { .. } => "formula",
    }
}

fn style_payload(style: &rxls::CellStyle) -> StylePayload {
    let font = style.font.as_ref().map(font_payload);
    let fill = style
        .pattern_fill
        .as_ref()
        .and_then(|fill| fill.foreground.or(fill.background))
        .map(color_hex)
        .or_else(|| style.fill.map(color_hex));
    let (fill_background, fill_foreground) = style
        .pattern_fill
        .as_ref()
        .map(|fill| {
            (
                fill.background.map(color_hex),
                fill.foreground.map(color_hex),
            )
        })
        .unwrap_or((None, None));
    let fill_pattern = style
        .pattern_fill
        .as_ref()
        .map(|fill| format!("{:?}", fill.pattern).to_ascii_lowercase());
    let border = style.border.as_ref().map(border_payload);
    let protection = style
        .protection
        .as_ref()
        .map(|protection| ProtectionPayload {
            locked: protection.locked,
            hidden: protection.hidden,
        });
    let (rotation, indent, shrink_to_fit) = style
        .align
        .as_ref()
        .map(|align| (align.rotation, align.indent, align.shrink_to_fit))
        .unwrap_or((0, 0, false));
    StylePayload {
        font,
        fill,
        number_format: style.num_fmt.clone(),
        horizontal: style.align.as_ref().and_then(|align| {
            align
                .horizontal
                .map(|value| format!("{:?}", value).to_ascii_lowercase())
        }),
        vertical: style.align.as_ref().and_then(|align| {
            align
                .vertical
                .map(|value| format!("{:?}", value).to_ascii_lowercase())
        }),
        wrap: style
            .align
            .as_ref()
            .map(|align| align.wrap)
            .unwrap_or(false),
        rotation,
        indent,
        shrink_to_fit,
        fill_pattern,
        fill_background,
        fill_foreground,
        border,
        protection,
    }
}

fn font_payload(font: &rxls::Font) -> FontPayload {
    FontPayload {
        name: font.name.clone(),
        size_pt: font.size_pt,
        bold: font.bold,
        italic: font.italic,
        underline: font.underline,
        strikethrough: font.strikethrough,
        color: font.color.map(color_hex),
        script: match font.script {
            rxls::FormatScript::None => None,
            script => Some(format!("{:?}", script).to_ascii_lowercase()),
        },
    }
}

fn border_payload(border: &rxls::Border) -> BorderPayload {
    BorderPayload {
        left: format!("{:?}", border.left).to_ascii_lowercase(),
        right: format!("{:?}", border.right).to_ascii_lowercase(),
        top: format!("{:?}", border.top).to_ascii_lowercase(),
        bottom: format!("{:?}", border.bottom).to_ascii_lowercase(),
        color: border.color.map(color_hex),
        left_color: border.left_color.map(color_hex),
        right_color: border.right_color.map(color_hex),
        top_color: border.top_color.map(color_hex),
        bottom_color: border.bottom_color.map(color_hex),
    }
}

fn color_hex(color: rxls::Color) -> String {
    let [red, green, blue] = color.as_rgb();
    format!("#{red:02X}{green:02X}{blue:02X}")
}

pub fn detect_format(file_name: &str, bytes: &[u8]) -> String {
    let file_name = file_name.to_ascii_lowercase();
    if file_name.ends_with(".docx") {
        "docx".into()
    } else if file_name.ends_with(".doc") {
        "doc".into()
    } else if file_name.ends_with(".odt") {
        "odt".into()
    } else if file_name.ends_with(".xlsx") {
        "xlsx".into()
    } else if file_name.ends_with(".xlsb") {
        "xlsb".into()
    } else if file_name.ends_with(".xls") {
        "xls".into()
    } else if file_name.ends_with(".ods") {
        "ods".into()
    } else if bytes.starts_with(b"PK") {
        "xlsx".into()
    } else {
        "unknown".into()
    }
}

fn validate_input(bytes: &[u8], file_name: &str) -> Result<(), ParseFailure> {
    if bytes.is_empty() {
        return Err(ParseFailure::invalid_input("The selected file is empty"));
    }
    if file_name.trim().is_empty() {
        return Err(ParseFailure::invalid_input(
            "A file name is required for format detection",
        ));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::detect_format;

    #[test]
    fn detects_supported_extensions() {
        assert_eq!(detect_format("report.docx", &[]), "docx");
        assert_eq!(detect_format("sheet.xlsx", &[]), "xlsx");
        assert_eq!(detect_format("unknown", b"PK\x03\x04"), "xlsx");
    }
}
