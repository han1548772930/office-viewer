use crate::model::{
    AnchorCellPayload, AxisPayload, BorderPayload, CellPayload, ChartCachedPointPayload,
    ChartDrawingPayload, ChartPayload, ChartSeriesCachePayload, ChartSeriesPayload,
    ChartSeriesStylePayload, DimensionPayload, DocumentModel, FontPayload, FreezePanePayload,
    HyperlinkPayload, LocalDefinedNamePayload, ParseEnvelope, ParseFailure, ParserDiagnostics,
    ProtectionPayload, RangePayload, RichTextRunPayload, SCHEMA_VERSION, SheetPayload,
    StyleLossPayload, StylePayload, WorkbookPayload, WorkbookPropertiesPayload,
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
        charts: sheet
            .charts()
            .iter()
            .enumerate()
            .map(|(index, chart)| chart_payload(index, chart, sheet.drawing_metadata()))
            .collect(),
    }
}

fn chart_payload(
    index: usize,
    chart: &rxls::Chart,
    drawings: &[rxls::DrawingMetadata],
) -> ChartPayload {
    let drawing = drawings
        .iter()
        .find(|item| item.kind == rxls::DrawingObjectKind::Chart && item.object_index == index)
        .map(chart_drawing_payload);
    ChartPayload {
        index,
        kind: format!("{:?}", chart.kind).to_ascii_lowercase(),
        title: chart.title.clone(),
        series: chart
            .series
            .iter()
            .map(|series| ChartSeriesPayload {
                name: series.name.clone(),
                categories: series.categories.clone(),
                values: series.values.clone(),
                bubble_sizes: series.bubble_sizes.clone(),
            })
            .collect(),
        legend: chart.legend,
        data_labels: chart.data_labels,
        x_axis_title: chart.x_axis_title.clone(),
        y_axis_title: chart.y_axis_title.clone(),
        from: AnchorCellPayload {
            row: chart.from.0,
            col: chart.from.1,
        },
        to: AnchorCellPayload {
            row: chart.to.0,
            col: chart.to.1,
        },
        drawing,
    }
}

fn chart_drawing_payload(metadata: &rxls::DrawingMetadata) -> ChartDrawingPayload {
    let anchor = |value: Option<(u32, u16)>| value.map(|(row, col)| AnchorCellPayload { row, col });
    let offsets = |value: Option<(i64, i64)>| value.map(|(x, y)| [x, y]);
    let size = |value: Option<(u64, u64)>| value.map(|(width, height)| [width, height]);
    let crop = metadata.crop.map(|value| {
        [
            value.left_ppm,
            value.top_ppm,
            value.right_ppm,
            value.bottom_ppm,
        ]
    });
    let cached_points = |points: &[rxls::ChartCachedPoint]| {
        points
            .iter()
            .map(|point| ChartCachedPointPayload {
                index: point.index,
                value: point.value.clone(),
            })
            .collect()
    };
    let caches = metadata
        .chart_series_caches
        .iter()
        .map(|cache| ChartSeriesCachePayload {
            name: cached_points(&cache.name),
            categories: cached_points(&cache.categories),
            values: cached_points(&cache.values),
            bubble_sizes: cached_points(&cache.bubble_sizes),
        })
        .collect();
    let styles = metadata
        .chart_series_styles
        .iter()
        .map(|style| ChartSeriesStylePayload {
            marker: format!("{:?}", style.marker).to_ascii_lowercase(),
            marker_size: style.marker_size,
            line_visible: style.line_visible,
            line_color: style.line_color.map(color_hex),
            line_width_emu: style.line_width_emu,
            losses: style
                .losses
                .iter()
                .map(|loss| format!("{:?}", loss).to_ascii_lowercase())
                .collect(),
        })
        .collect();
    ChartDrawingPayload {
        object_index: metadata.object_index,
        from_cell: anchor(metadata.from_cell),
        to_cell: anchor(metadata.to_cell),
        from_offset_emu: offsets(metadata.from_offset_emu),
        to_offset_emu: offsets(metadata.to_offset_emu),
        absolute_size_emu: size(metadata.absolute_size_emu),
        crop,
        rotation_mdeg: metadata.rotation_mdeg,
        z_order: metadata.z_order,
        alt_text: metadata.alt_text.clone(),
        name: metadata.name.clone(),
        behavior: format!("{:?}", metadata.behavior).to_ascii_lowercase(),
        chart_palette: metadata
            .chart_palette
            .iter()
            .copied()
            .map(color_hex)
            .collect(),
        chart_default_latin_font_family: metadata.chart_default_latin_font_family.clone(),
        chart_category_axis_visible: metadata.chart_category_axis_visible,
        chart_category_axis_shifted: metadata.chart_category_axis_shifted,
        chart_value_axis_visible: metadata.chart_value_axis_visible,
        chart_category_major_gridlines: metadata.chart_category_major_gridlines,
        chart_value_major_gridlines: metadata.chart_value_major_gridlines,
        chart_bar_direction: format!("{:?}", metadata.chart_bar_direction).to_ascii_lowercase(),
        chart_frame_fill: format!("{:?}", metadata.chart_frame_fill).to_ascii_lowercase(),
        chart_frame_style_losses: metadata
            .chart_frame_style_losses
            .iter()
            .map(|loss| format!("{:?}", loss).to_ascii_lowercase())
            .collect(),
        chart_unsupported_reasons: metadata
            .chart_unsupported_reasons
            .iter()
            .map(|reason| format!("{:?}", reason).to_ascii_lowercase())
            .collect(),
        chart_series_caches: caches,
        chart_series_styles: styles,
        chart_text_styles: json!({
            "chartTitle": chart_text_style(metadata.chart_text_styles.chart_title.as_ref()),
            "categoryAxisTitle": chart_text_style(metadata.chart_text_styles.category_axis_title.as_ref()),
            "valueAxisTitle": chart_text_style(metadata.chart_text_styles.value_axis_title.as_ref()),
            "legend": chart_text_style(metadata.chart_text_styles.legend.as_ref()),
            "categoryAxisLabels": chart_text_style(metadata.chart_text_styles.category_axis_labels.as_ref()),
            "valueAxisLabels": chart_text_style(metadata.chart_text_styles.value_axis_labels.as_ref()),
            "dataLabels": chart_text_style(metadata.chart_text_styles.data_labels.as_ref()),
        }),
    }
}

fn chart_text_style(style: Option<&rxls::ChartTextStyle>) -> Value {
    let Some(style) = style else {
        return Value::Null;
    };
    json!({
        "latinFontFamily": style.latin_font_family,
        "sizeHundredthsOfPoint": style.size_hundredths_of_point,
        "color": color_hex(style.color),
        "bold": style.bold,
        "italic": style.italic,
        "underline": style.underline,
        "strikethrough": style.strikethrough,
        "kerningMinimumHundredthsOfPoint": style.kerning_minimum_hundredths_of_point,
        "rotationDegrees": style.rotation_degrees,
    })
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
    use super::{detect_format, parse_office_bytes};
    use rxls::{Chart, ChartKind, Series, Workbook};
    use serde_json::Value;

    #[test]
    fn detects_supported_extensions() {
        assert_eq!(detect_format("report.docx", &[]), "docx");
        assert_eq!(detect_format("sheet.xlsx", &[]), "xlsx");
        assert_eq!(detect_format("unknown", b"PK\x03\x04"), "xlsx");
    }

    #[test]
    fn xlsx_chart_series_titles_anchors_and_drawing_sidecar_reach_parse_payload() {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_sheet("Data");
        sheet.write(0, 0, "Month");
        sheet.write(0, 1, "Revenue");
        sheet.write(1, 0, "Jan");
        sheet.write(1, 1, 120.0);
        sheet.write(2, 0, "Feb");
        sheet.write(2, 1, 180.0);
        sheet.add_chart(
            Chart::new(ChartKind::Bar, (4, 3), (18, 12))
                .with_title("Monthly revenue")
                .with_x_axis_title("Month")
                .with_y_axis_title("Revenue")
                .with_legend(true)
                .with_data_labels(true)
                .with_series([Series::new("Data!$B$2:$B$3")
                    .with_name("Data!$B$1")
                    .with_categories("Data!$A$2:$A$3")]),
        );

        let bytes = workbook.to_xlsx();
        let envelope = parse_office_bytes(&bytes, "chart-fixture.xlsx")
            .expect("the generated xlsx fixture must parse");
        let payload = serde_json::to_value(envelope).expect("parse envelope must serialize");
        let chart = &payload["payload"]["data"]["sheets"][0]["charts"][0];

        assert_eq!(chart["kind"], "bar");
        assert_eq!(chart["title"], "Monthly revenue");
        assert_eq!(chart["legend"], true);
        assert_eq!(chart["dataLabels"], true);
        assert_eq!(chart["xAxisTitle"], "Month");
        assert_eq!(chart["yAxisTitle"], "Revenue");
        assert_eq!(chart["from"], serde_json::json!({ "row": 4, "col": 3 }));
        assert_eq!(chart["to"], serde_json::json!({ "row": 18, "col": 12 }));
        assert_eq!(chart["series"][0]["name"], "Data!$B$1");
        assert_eq!(chart["series"][0]["categories"], "Data!$A$2:$A$3");
        assert_eq!(chart["series"][0]["values"], "Data!$B$2:$B$3");
        assert!(chart["drawing"].is_object());
        assert_eq!(payload["diagnostics"]["stats"]["chartCount"], 1);
        assert!(
            payload["diagnostics"]["capabilities"]
                .as_array()
                .unwrap()
                .iter()
                .any(|capability| capability == "charts")
        );
    }

    #[test]
    fn chart_payload_retains_unsupported_constructs_as_structured_diagnostics() {
        let mut metadata = rxls::DrawingMetadata::default();
        metadata.kind = rxls::DrawingObjectKind::Chart;
        metadata.chart_unsupported_reasons = vec![rxls::ChartUnsupportedReason::Combo];
        metadata.from_cell = Some((1, 2));
        metadata.absolute_size_emu = Some((914_400, 457_200));
        let chart = Chart::new(ChartKind::Bar, (1, 2), (8, 8));
        let payload = super::chart_payload(0, &chart, &[metadata]);
        let json: Value = serde_json::to_value(payload).expect("chart payload must serialize");

        assert_eq!(json["drawing"]["chartUnsupportedReasons"][0], "combo");
        assert_eq!(
            json["drawing"]["fromCell"],
            serde_json::json!({ "row": 1, "col": 2 })
        );
        assert_eq!(
            json["drawing"]["absoluteSizeEmu"],
            serde_json::json!([914_400, 457_200])
        );
    }
}
