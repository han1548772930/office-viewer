//! XLSX OOXML writer for the canonical spreadsheet model.
//!
//! This module deliberately emits Office Open XML parts instead of serialising the
//! editor snapshot as JSON.  It owns only the spreadsheet parts; the OPC package
//! and ZIP implementation are supplied by the parent writer crate.

use crate::{ExportError, ExportWarning, opc::Package, xml};
use serde::{Deserialize, Serialize};
use serde_json::{Map, Value};
use office_io_common::chart_options::Axis as ChartAxisInput;
use std::borrow::Cow;
use std::collections::{BTreeMap, BTreeSet, HashMap};

#[path = "xlsx_filter.rs"]
mod auto_filter;
#[path = "xlsx_validation.rs"]
mod data_validation;
#[path = "xlsx_sparkline.rs"]
mod sparkline;
#[path = "xlsx_pivot.rs"]
mod pivot;
#[path = "xlsx_table.rs"]
mod table;
#[path = "xlsx_dimensions.rs"]
mod dimensions;
#[path = "xlsx_conditional_formula.rs"]
mod conditional_formula;
#[path = "xlsx_conditional_source.rs"]
mod conditional_source;
#[path = "xlsx_chart.rs"]
mod chart;
use chart::chart_xml;

const MAIN_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL_WORKBOOK: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const REL_WORKSHEET: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const REL_DRAWING: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
const REL_CHART: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
const REL_STYLES: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
const REL_CORE: &str =
    "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
const REL_APP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
const REL_HYPERLINK: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
const REL_PIVOT_TABLE: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable";
const REL_PIVOT_CACHE_DEFINITION: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition";
const REL_PIVOT_CACHE_RECORDS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords";
const MAX_ROW_INDEX: u32 = 1_048_575;
const MAX_COLUMN_INDEX: u32 = 16_383;

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct WorkbookInput {
    #[serde(default)]
    kind: Option<String>,
    #[serde(default)]
    name: String,
    #[serde(default)]
    sheets: Vec<SheetInput>,
    #[serde(default)]
    active_sheet_id: Option<String>,
    #[serde(default)]
    date1904: bool,
    #[serde(default)]
    metadata: Option<Value>,
    #[serde(default)]
    pivot_tables: Option<Vec<pivot::PivotTableInput>>,
    #[serde(default)]
    tables: Vec<table::TableInput>,
    #[serde(default)]
    charts: Option<Vec<ChartInput>>,
}

#[derive(Debug, Default, Deserialize)]
#[serde(rename_all = "camelCase")]
struct SheetInput {
    #[serde(default)]
    id: String,
    #[serde(default)]
    name: String,
    #[serde(default = "default_true")]
    visible: bool,
    #[serde(default)]
    cells: Map<String, Value>,
    #[serde(default)]
    merged_ranges: Vec<RangeInput>,
    #[serde(default)]
    column_widths: BTreeMap<String, f64>,
    #[serde(default)]
    row_heights: BTreeMap<String, f64>,
    #[serde(default)]
    axis_unit: Option<String>,
    #[serde(default)]
    row_height_spans: Vec<dimensions::AxisSpanInput<f64>>,
    #[serde(default)]
    column_width_spans: Vec<dimensions::AxisSpanInput<f64>>,
    #[serde(default)]
    row_auto_height_spans: Vec<dimensions::AxisSpanInput<bool>>,
    #[serde(default)]
    row_auto_heights: BTreeMap<String, f64>,
    #[serde(default)]
    default_row_auto_height: Option<bool>,
    #[serde(default)]
    frozen_rows: u32,
    #[serde(default)]
    frozen_cols: u32,
    #[serde(default = "default_true")]
    show_gridlines: bool,
    #[serde(default)]
    hidden_rows: Vec<u32>,
    #[serde(default)]
    row_hidden_spans: Vec<dimensions::AxisSpanInput<bool>>,
    #[serde(default)]
    column_hidden_spans: Vec<dimensions::AxisSpanInput<bool>>,
    #[serde(default)]
    outline: Option<dimensions::OutlineInput>,
    #[serde(default)]
    auto_filter: Option<auto_filter::AutoFilterInput>,
    #[serde(default)]
    data_validations: Vec<data_validation::DataValidationInput>,
    #[serde(default)]
    conditional_formats: Vec<ConditionalFormatInput>,
    #[serde(default)]
    hidden_columns: Vec<u32>,
    #[serde(default)]
    default_column_width: Option<f64>,
    #[serde(default)]
    default_row_height: Option<f64>,
    #[serde(default)]
    row_count: Option<u32>,
    #[serde(default)]
    column_count: Option<u32>,
    #[serde(default)]
    zoom: Option<f64>,
    #[serde(default)]
    metadata: Option<Value>,
}

#[derive(Debug, Deserialize, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
struct ConditionalFormatInput {
    #[serde(default)]
    id: String,
    #[serde(rename = "type", default)]
    rule_type: String,
    #[serde(default)]
    ranges: Vec<RangeInput>,
    #[serde(default)]
    formula_anchor: Option<conditional_formula::FormulaAnchor>,
    #[serde(default)]
    operator: String,
    #[serde(default)]
    formula1: String,
    #[serde(default)]
    formula2: Option<String>,
    #[serde(default)]
    style: ConditionalFormatStyleInput,
    #[serde(default)]
    priority: u32,
    #[serde(default)]
    stop_if_true: bool,
    #[serde(default)]
    sub_type: Option<String>,
    #[serde(default)]
    value: Option<Value>,
    #[serde(default)]
    is_bottom: bool,
    #[serde(default)]
    is_percent: bool,
    #[serde(default)]
    config: Option<Value>,
    #[serde(default)]
    office_source: Option<office_io_common::conditional_format_source::OfficeSource>,
    #[serde(flatten)]
    extensions: Map<String, Value>,
}

#[derive(Debug, Deserialize, Serialize, Clone, Default, PartialEq, Eq)]
#[serde(rename_all = "camelCase")]
struct ConditionalFormatStyleInput {
    fill: Option<String>,
    font_color: Option<String>,
    bold: Option<bool>,
    italic: Option<bool>,
    underline: Option<bool>,
    strike: Option<bool>,
    #[serde(flatten)]
    extensions: Map<String, Value>,
}

#[derive(Debug, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
struct ChartInput {
    #[serde(default)]
    id: String,
    #[serde(default)]
    sheet_id: String,
    #[serde(rename = "type")]
    chart_type: String,
    source: ChartSourceInput,
    position: ChartPositionInput,
    size: ChartSizeInput,
    #[serde(default)]
    title: Option<String>,
    #[serde(default)]
    subtitle: Option<String>,
    #[serde(default)]
    palette: Vec<String>,
    #[serde(default)]
    options: AuthoredChartOptions,
    #[serde(default)]
    legend: Option<ChartLegendInput>,
    #[serde(default)]
    category_field: usize,
    #[serde(default)]
    value_fields: Vec<usize>,
    #[serde(default)]
    series_to_category: BTreeMap<usize, Vec<usize>>,
    #[serde(default)]
    pivot_table_id: Option<String>,
    #[serde(default)]
    x_axis: Option<ChartAxisInput>,
    #[serde(default)]
    y_axis: Option<ChartAxisInput>,
    #[serde(skip)]
    native_bindings: Option<chart::NativeBindings>,
}

#[derive(Debug, Deserialize, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
struct ChartSourceInput {
    range: RangeInput,
    #[serde(default = "default_chart_orientation")]
    orientation: String,
    #[serde(default)]
    sheet_id: Option<String>,
}

#[derive(Debug, Deserialize, Clone)]
struct ChartPositionInput {
    x: f64,
    y: f64,
}

#[derive(Debug, Deserialize, Clone)]
struct ChartSizeInput {
    width: f64,
    height: f64,
}

#[derive(Debug, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
struct ChartLegendInput {
    #[serde(default = "default_true")]
    visible: bool,
    #[serde(default = "default_chart_legend_position")]
    position: String,
    #[serde(default)]
    font: Option<office_io_common::chart_options::Font>,
}

/// Keep the original settings next to the validated native DTO. Unknown nested
/// settings must survive the source manifest even before their native adapter
/// is implemented; this retained source never drives a second chart engine.
#[derive(Debug, Clone)]
struct AuthoredChartOptions {
    typed: office_io_common::chart_options::Options,
    source: Value,
}

impl Default for AuthoredChartOptions {
    fn default() -> Self { Self { typed: Default::default(), source: Value::Object(Map::new()) } }
}

impl<'de> Deserialize<'de> for AuthoredChartOptions {
    fn deserialize<D: serde::Deserializer<'de>>(deserializer: D) -> Result<Self, D::Error> {
        let source = Value::deserialize(deserializer)?;
        let typed: office_io_common::chart_options::Options = serde_json::from_value(source.clone()).map_err(serde::de::Error::custom)?;
        typed.validate_paints().map_err(serde::de::Error::custom)?;
        Ok(Self { typed, source })
    }
}

impl Serialize for AuthoredChartOptions {
    fn serialize<S: serde::Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> { self.source.serialize(serializer) }
}

impl std::ops::Deref for AuthoredChartOptions {
    type Target = office_io_common::chart_options::Options;
    fn deref(&self) -> &Self::Target { &self.typed }
}

fn default_chart_orientation() -> String {
    "columns".to_owned()
}

fn default_chart_legend_position() -> String {
    "bottom".to_owned()
}


#[derive(Debug, Deserialize, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
struct RangeInput {
    #[serde(default)]
    start_row: Option<u32>,
    #[serde(default)]
    start_col: Option<u32>,
    #[serde(default)]
    end_row: Option<u32>,
    #[serde(default)]
    end_col: Option<u32>,
    #[serde(default)]
    first_row: Option<u32>,
    #[serde(default)]
    first_col: Option<u32>,
    #[serde(default)]
    last_row: Option<u32>,
    #[serde(default)]
    last_col: Option<u32>,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct CellInput {
    #[serde(default)]
    id: String,
    #[serde(default)]
    row: Option<u32>,
    #[serde(default)]
    col: Option<u32>,
    #[serde(default)]
    value: Option<Value>,
    #[serde(default)]
    formula: Option<String>,
    #[serde(default)]
    display: String,
    #[serde(default)]
    style: Option<StyleInput>,
    #[serde(default)]
    data_type: Option<String>,
    #[serde(default)]
    cached_data_type: Option<String>,
    #[serde(default)]
    rich_text: Vec<Value>,
    #[serde(default)]
    hyperlink: Option<String>,
    #[serde(default)]
    metadata: Option<Value>,
}

#[derive(Debug, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
struct StyleInput {
    #[serde(default)]
    font: Option<FontInput>,
    #[serde(default)]
    fill: Option<String>,
    #[serde(default)]
    horizontal: Option<String>,
    #[serde(default)]
    vertical: Option<String>,
    #[serde(default)]
    wrap: bool,
    #[serde(default)]
    number_format: Option<String>,
    #[serde(default)]
    border: Option<Value>,
    #[serde(default)]
    rotation: Option<i32>,
    #[serde(default)]
    rotation_vertical: bool,
    #[serde(default)]
    indent: Option<i32>,
    #[serde(default)]
    shrink_to_fit: bool,
    #[serde(default)]
    metadata: Option<Value>,
}

#[derive(Debug, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
struct FontInput {
    #[serde(default)]
    family: Option<String>,
    #[serde(default)]
    name: Option<String>,
    #[serde(default)]
    size: Option<f64>,
    #[serde(default)]
    size_pt: Option<f64>,
    #[serde(default)]
    bold: bool,
    #[serde(default)]
    italic: bool,
    #[serde(default)]
    underline: bool,
    #[serde(default)]
    strike: bool,
    #[serde(default)]
    strikethrough: bool,
    #[serde(default)]
    color: Option<String>,
}

fn default_true() -> bool {
    true
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
struct StyleKey {
    font: FontKey,
    fill: String,
    horizontal: String,
    vertical: String,
    wrap: bool,
    number_format: String,
    border: String,
    rotation: i32,
    indent: i32,
    shrink_to_fit: bool,
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
struct FontKey {
    family: String,
    size: String,
    bold: bool,
    italic: bool,
    underline: bool,
    strike: bool,
    color: String,
}

#[derive(Debug, Clone)]
struct RegisteredStyle {
    key: StyleKey,
    font_id: usize,
    fill_id: usize,
    border_id: usize,
    num_fmt_id: Option<u32>,
}

#[derive(Debug, Default)]
struct StyleRegistry {
    styles: Vec<RegisteredStyle>,
    by_key: HashMap<StyleKey, usize>,
    fonts: Vec<FontKey>,
    font_ids: HashMap<FontKey, usize>,
    fills: Vec<String>,
    fill_ids: HashMap<String, usize>,
    borders: Vec<String>,
    border_ids: HashMap<String, usize>,
    num_fmts: Vec<String>,
    num_fmt_ids: HashMap<String, u32>,
    conditional_styles: Vec<String>,
    conditional_style_ids: HashMap<String, usize>,
}

impl StyleRegistry {
    fn new() -> Self {
        let mut registry = Self::default();
        registry.fills.push("none".to_owned());
        registry.fills.push("gray125".to_owned());
        registry.fill_ids.insert("none".to_owned(), 0);
        registry.fill_ids.insert("gray125".to_owned(), 1);
        registry.borders.push(String::new());
        registry.border_ids.insert(String::new(), 0);
        let default_font = FontKey {
            family: "Aptos".to_owned(),
            size: "11".to_owned(),
            bold: false,
            italic: false,
            underline: false,
            strike: false,
            color: "none".to_owned(),
        };
        registry.fonts.push(default_font.clone());
        registry.font_ids.insert(default_font, 0);
        registry
    }

    fn register(&mut self, style: Option<&StyleInput>) -> usize {
        let key = style_key(style);
        if let Some(index) = self.by_key.get(&key) {
            return *index;
        }
        let font_id = if let Some(id) = self.font_ids.get(&key.font) {
            *id
        } else {
            let id = self.fonts.len();
            self.fonts.push(key.font.clone());
            self.font_ids.insert(key.font.clone(), id);
            id
        };
        let fill_id = if let Some(id) = self.fill_ids.get(&key.fill) {
            *id
        } else {
            let id = self.fills.len();
            self.fills.push(key.fill.clone());
            self.fill_ids.insert(key.fill.clone(), id);
            id
        };
        let border_id = if let Some(id) = self.border_ids.get(&key.border) {
            *id
        } else {
            let id = self.borders.len();
            self.borders.push(key.border.clone());
            self.border_ids.insert(key.border.clone(), id);
            id
        };
        let num_fmt_id = if key.number_format.is_empty() || key.number_format == "General" {
            None
        } else if let Some(id) = self.num_fmt_ids.get(&key.number_format) {
            Some(*id)
        } else {
            let id = 164 + self.num_fmts.len() as u32;
            self.num_fmt_ids.insert(key.number_format.clone(), id);
            self.num_fmts.push(key.number_format.clone());
            Some(id)
        };
        let index = self.styles.len();
        self.styles.push(RegisteredStyle {
            key: key.clone(),
            font_id,
            fill_id,
            border_id,
            num_fmt_id,
        });
        self.by_key.insert(key, index);
        index
    }

    fn register_conditional_style_xml(&mut self, style: &office_io_common::conditional_format_xml::Element) -> usize {
        let key = style.fingerprint();
        if let Some(index) = self.conditional_style_ids.get(&key) { return *index }
        let index = self.conditional_styles.len();
        self.conditional_styles.push(style.render());
        self.conditional_style_ids.insert(key, index);
        index
    }
}

fn style_key(style: Option<&StyleInput>) -> StyleKey {
    let default_font = FontKey {
        family: "Aptos".to_owned(),
        size: "11".to_owned(),
        bold: false,
        italic: false,
        underline: false,
        strike: false,
        color: "none".to_owned(),
    };
    let Some(style) = style else {
        return StyleKey {
            font: default_font,
            fill: "none".to_owned(),
            horizontal: "general".to_owned(),
            vertical: "bottom".to_owned(),
            wrap: false,
            number_format: "General".to_owned(),
            border: String::new(),
            rotation: 0,
            indent: 0,
            shrink_to_fit: false,
        };
    };
    let font = style.font.as_ref().map(font_key).unwrap_or(default_font);
    let fill = normalize_color(style.fill.as_deref().unwrap_or("none"));
    let vertical = match style.vertical.as_deref() {
        Some("middle") | Some("center") => "center",
        Some("top") => "top",
        Some("justify") => "justify",
        Some("distributed") => "distributed",
        _ => "bottom",
    };
    let horizontal = match style.horizontal.as_deref() {
        Some("left")
        | Some("center")
        | Some("right")
        | Some("justify")
        | Some("general")
        | Some("fill")
        | Some("distributed")
        | Some("centerContinuous") => style
            .horizontal
            .clone()
            .unwrap_or_else(|| "general".to_owned()),
        _ => "general".to_owned(),
    };
    StyleKey {
        font,
        fill,
        horizontal,
        vertical: vertical.to_owned(),
        wrap: style.wrap,
        number_format: style
            .number_format
            .clone()
            .unwrap_or_else(|| "General".to_owned()),
        border: border_key(style.border.as_ref()),
        rotation: if style.rotation_vertical || style.rotation == Some(255) {
            255
        } else {
            style.rotation.unwrap_or(0).clamp(-90, 180)
        },
        indent: style.indent.unwrap_or(0).clamp(0, 255),
        shrink_to_fit: style.shrink_to_fit,
    }
}

fn font_key(font: &FontInput) -> FontKey {
    let family = font
        .family
        .clone()
        .or_else(|| font.name.clone())
        .unwrap_or_else(|| "Aptos".to_owned());
    let size = font.size.or(font.size_pt).unwrap_or(11.0);
    FontKey {
        family,
        size: format_number(size),
        bold: font.bold,
        italic: font.italic,
        underline: font.underline,
        strike: font.strike || font.strikethrough,
        color: normalize_color(font.color.as_deref().unwrap_or("none")),
    }
}

fn normalize_color(value: &str) -> String {
    let stripped = value.trim().trim_start_matches('#');
    if stripped.eq_ignore_ascii_case("none") || stripped.is_empty() {
        return "none".to_owned();
    }
    let upper = stripped.to_ascii_uppercase();
    if upper.len() == 3 && upper.bytes().all(|ch| ch.is_ascii_hexdigit()) {
        let mut rgb = String::from("FF");
        for ch in upper.chars() { rgb.push(ch); rgb.push(ch); }
        return rgb;
    }
    if upper.len() == 6 {
        format!("FF{upper}")
    } else if upper.len() == 8 {
        upper
    } else {
        "FF000000".to_owned()
    }
}

fn border_key(value: Option<&Value>) -> String {
    let Some(Value::Object(map)) = value else {
        return String::new();
    };
    let edge = |name: &str| -> String {
        let Some(item) = map.get(name) else {
            return String::new();
        };
        match item {
            Value::String(text) => text.to_owned(),
            Value::Object(obj) => {
                let style = obj.get("style").and_then(Value::as_str).unwrap_or("thin");
                let color = obj.get("color").and_then(Value::as_str).unwrap_or("");
                format!("{style}:{color}")
            }
            _ => String::new(),
        }
    };
    ["left", "right", "top", "bottom", "diagonal", "diagonalUp", "diagonalDown"]
        .iter()
        .map(|key| format!("{key}={}", edge(key)))
        .collect::<Vec<_>>()
        .join(";")
}

fn format_number(value: f64) -> String {
    if value.fract() == 0.0 {
        format!("{value:.0}")
    } else {
        format!("{value:.4}")
            .trim_end_matches('0')
            .trim_end_matches('.')
            .to_owned()
    }
}

fn esc(value: &str) -> Cow<'_, str> {
    xml::escape(value)
}

fn attr(name: &str, value: &str) -> String {
    format!(" {name}=\"{}\"", esc(value))
}

fn parse_index(key: &str) -> Option<u32> {
    key.parse().ok()
}

fn column_name(mut index: u32) -> String {
    let mut result = String::new();
    loop {
        result.insert(
            0,
            char::from_u32(u32::from(b'A') + index % 26).unwrap_or('A'),
        );
        if index < 26 {
            break;
        }
        index = index / 26 - 1;
    }
    result
}

fn cell_ref(row: u32, col: u32) -> String {
    format!("{}{}", column_name(col), row + 1)
}

fn range_ref(range: &RangeInput) -> Option<String> {
    let first_row = range.start_row.or(range.first_row)?;
    let first_col = range.start_col.or(range.first_col)?;
    let last_row = range.end_row.or(range.last_row).unwrap_or(first_row);
    let last_col = range.end_col.or(range.last_col).unwrap_or(first_col);
    Some(format!(
        "{}:{}",
        cell_ref(first_row, first_col),
        cell_ref(last_row, last_col)
    ))
}

fn cell_from_entry(key: &str, value: &Value) -> Option<CellInput> {
    let mut cell = serde_json::from_value::<CellInput>(value.clone()).ok()?;
    let (row, col) = parse_a1(key).or_else(|| Some((cell.row?, cell.col?)))?;
    cell.row = Some(row);
    cell.col = Some(col);
    Some(cell)
}

fn parse_a1(value: &str) -> Option<(u32, u32)> {
    let upper = value.to_ascii_uppercase();
    let mut letters = String::new();
    let mut digits = String::new();
    for ch in upper.chars() {
        if ch.is_ascii_alphabetic() && digits.is_empty() {
            letters.push(ch);
        } else if ch.is_ascii_digit() {
            digits.push(ch);
        } else {
            return None;
        }
    }
    if letters.is_empty() || digits.is_empty() {
        return None;
    }
    let mut col = 0u32;
    for ch in letters.chars() {
        col = col
            .checked_mul(26)?
            .checked_add(ch as u32 - 'A' as u32 + 1)?;
    }
    Some((
        digits.parse::<u32>().ok()?.checked_sub(1)?,
        col.checked_sub(1)?,
    ))
}

fn style_xml(registry: &StyleRegistry) -> String {
    let mut xml_out = String::from(xml::DECL);
    xml_out.push_str(&format!("<styleSheet xmlns=\"{MAIN_NS}\">"));
    if !registry.num_fmts.is_empty() {
        xml_out.push_str(&format!("<numFmts count=\"{}\">", registry.num_fmts.len()));
        for (index, code) in registry.num_fmts.iter().enumerate() {
            xml_out.push_str(&format!(
                "<numFmt numFmtId=\"{}\" formatCode=\"{}\"/>",
                164 + index,
                esc(code)
            ));
        }
        xml_out.push_str("</numFmts>");
    }
    xml_out.push_str(&format!("<fonts count=\"{}\">", registry.fonts.len()));
    for font in &registry.fonts {
        xml_out.push_str("<font>");
        xml_out.push_str(&format!("<sz val=\"{}\"/>", font.size));
        if font.color != "none" {
            xml_out.push_str(&format!("<color rgb=\"{}\"/>", font.color));
        }
        xml_out.push_str(&format!("<name val=\"{}\"/>", esc(&font.family)));
        if font.bold {
            xml_out.push_str("<b/>");
        }
        if font.italic {
            xml_out.push_str("<i/>");
        }
        if font.underline {
            xml_out.push_str("<u/>");
        }
        if font.strike {
            xml_out.push_str("<strike/>");
        }
        xml_out.push_str("</font>");
    }
    xml_out.push_str("</fonts>");
    xml_out.push_str(&format!("<fills count=\"{}\">", registry.fills.len()));
    for fill in &registry.fills {
        match fill.as_str() {
            "none" => xml_out.push_str("<fill><patternFill patternType=\"none\"/></fill>"),
            "gray125" => xml_out.push_str("<fill><patternFill patternType=\"gray125\"/></fill>"),
            color => xml_out.push_str(&format!("<fill><patternFill patternType=\"solid\"><fgColor rgb=\"{}\"/><bgColor indexed=\"64\"/></patternFill></fill>", color)),
        }
    }
    xml_out.push_str("</fills>");
    xml_out.push_str(&format!("<borders count=\"{}\">", registry.borders.len()));
    for border in &registry.borders {
        xml_out.push_str(&border_xml(border));
    }
    xml_out.push_str("</borders>");
    xml_out.push_str("<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>");
    xml_out.push_str(&format!("<cellXfs count=\"{}\">", registry.styles.len()));
    for style in &registry.styles {
        let num_fmt = style.num_fmt_id.unwrap_or(0);
        let mut xf = format!(
            "<xf numFmtId=\"{num_fmt}\" fontId=\"{}\" fillId=\"{}\" borderId=\"{}\" xfId=\"0\"",
            style.font_id, style.fill_id, style.border_id
        );
        let key = &style.key;
        let has_alignment = key.horizontal != "general"
            || key.vertical != "bottom"
            || key.wrap
            || key.rotation != 0
            || key.indent != 0
            || key.shrink_to_fit;
        if style.num_fmt_id.is_some() {
            xf.push_str(" applyNumberFormat=\"1\"");
        }
        if style.font_id != 0 {
            xf.push_str(" applyFont=\"1\"");
        }
        if style.fill_id > 1 {
            xf.push_str(" applyFill=\"1\"");
        }
        if style.border_id > 0 {
            xf.push_str(" applyBorder=\"1\"");
        }
        if has_alignment {
            xf.push_str(" applyAlignment=\"1\"");
        }
        if has_alignment {
            xf.push('>');
            xf.push_str("<alignment");
            if key.horizontal != "general" {
                xf.push_str(&attr("horizontal", &key.horizontal));
            }
            if key.vertical != "bottom" {
                xf.push_str(&attr("vertical", &key.vertical));
            }
            if key.wrap {
                xf.push_str(" wrapText=\"1\"");
            }
            if key.rotation != 0 {
                let rotation = if key.rotation < 0 {
                    90 - key.rotation
                } else {
                    key.rotation
                };
                xf.push_str(&format!(" textRotation=\"{rotation}\""));
            }
            if key.indent != 0 {
                xf.push_str(&format!(" indent=\"{}\"", key.indent));
            }
            if key.shrink_to_fit {
                xf.push_str(" shrinkToFit=\"1\"");
            }
            xf.push_str("/></xf>");
        } else {
            xf.push_str("/>");
        }
        xml_out.push_str(&xf);
    }
    xml_out.push_str("</cellXfs><cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>");
    xml_out.push_str(&format!("<dxfs count=\"{}\">", registry.conditional_styles.len()));
    for style in &registry.conditional_styles { xml_out.push_str(style); }
    xml_out.push_str("</dxfs><tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium2\" defaultPivotStyle=\"PivotStyleMedium9\"/></styleSheet>");
    xml_out
}

fn border_xml(value: &str) -> String {
    let mut map = BTreeMap::new();
    for part in value.split(';') {
        if let Some((key, data)) = part.split_once('=') {
            map.insert(key, data);
        }
    }
    let edge = |name: &str, key: &str| -> String {
        let Some(value) = map.get(key).copied() else {
            return format!("<{name}/>");
        };
        if value.is_empty() {
            return format!("<{name}/>");
        }
        let mut pieces = value.splitn(2, ':');
        let style = match pieces.next().unwrap_or("thin") {
            "dash-dot" => "dashDot", "dash-dot-dot" => "dashDotDot",
            "medium-dashed" => "mediumDashed", "medium-dash-dot" => "mediumDashDot",
            "medium-dash-dot-dot" => "mediumDashDotDot", "slant-dash-dot" => "slantDashDot", other => other,
        };
        let color = pieces.next().unwrap_or("");
        if color.is_empty() {
            format!("<{name} style=\"{}\"/>", esc(style))
        } else {
            format!(
                "<{name} style=\"{}\"><color rgb=\"{}\"/></{name}>",
                esc(style),
                normalize_color(color)
            )
        }
    };
    let up = map.get("diagonalUp").is_some_and(|v| !v.is_empty());
    let down = map.get("diagonalDown").is_some_and(|v| !v.is_empty());
    format!(
        "<border{}{}>{}{}</border>",
        if up { " diagonalUp=\"1\"" } else { "" },
        if down { " diagonalDown=\"1\"" } else { "" },
        ["left", "right", "top", "bottom"]
            .iter()
            .map(|name| edge(name, name))
            .collect::<String>(),
        edge("diagonal", if down { "diagonalDown" } else if up { "diagonalUp" } else { "diagonal" }),
    )
}

fn worksheet_xml(
    sheet: &SheetInput,
    sheet_index: usize,
    sheet_ids: &[String],
    sheet_names: &[String],
    registry: &mut StyleRegistry,
    warnings: &mut Vec<ExportWarning>,
    drawing_rel_id: Option<&str>,
) -> Result<(String, Vec<(String, String)>), ExportError> {
    let mut cells = Vec::<CellInput>::new();
    for (key, value) in &sheet.cells {
        if let Some(cell) = cell_from_entry(key, value) {
            cells.push(cell);
        } else {
            warnings.push(ExportWarning::new(
                "INVALID_CELL",
                format!("Ignored invalid cell key {key:?} on sheet {}", sheet.name),
            ));
        }
    }
    cells.sort_by_key(|cell| {
        (
            cell.row.unwrap_or(0),
            cell.col.unwrap_or(0),
            cell.id.clone(),
        )
    });
    let limited_borders = cells.iter().filter(|cell| {
        let Some(border) = cell.style.as_ref().and_then(|s| s.border.as_ref()).and_then(Value::as_object) else { return false; };
        ["topLeftBottomCenter", "topLeftMiddleRight", "middleLeftTopRight", "bottomCenterTopRight"].iter().any(|key| border.get(*key).is_some_and(|v| !v.is_null())) ||
            matches!((border.get("diagonalUp"), border.get("diagonalDown")), (Some(up), Some(down)) if !up.is_null() && !down.is_null() && up != down)
    }).count();
    if limited_borders > 0 {
        warnings.push(ExportWarning::new("LIMITED_DIAGONAL_BORDER", format!("{}: {limited_borders} cell(s) have compound diagonals or different diagonal pens that standard XLSX cell borders cannot fully represent", sheet.name)));
    }
    let hyperlinks = cells
        .iter()
        .filter_map(|cell| {
            let target = cell.hyperlink.as_ref()?;
            Some((cell_ref(cell.row?, cell.col?), target.clone()))
        })
        .collect::<Vec<_>>();
    let mut max_row = sheet.row_count.unwrap_or(0).saturating_sub(1);
    let mut max_col = sheet.column_count.unwrap_or(0).saturating_sub(1);
    for cell in &cells {
        max_row = max_row.max(cell.row.unwrap_or(0));
        max_col = max_col.max(cell.col.unwrap_or(0));
    }
    if let Some(filter) = &sheet.auto_filter {
        if let Some((_, _, end_row, end_col)) = range_bounds(&filter.range) {
            max_row = max_row.max(end_row);
            max_col = max_col.max(end_col);
        }
    }
    let mut out = String::from(xml::DECL);
    out.push_str(&format!(
        "<worksheet xmlns=\"{MAIN_NS}\" xmlns:r=\"{REL_NS}\">"
    ));
    if sheet
        .auto_filter
        .as_ref()
        .is_some_and(|filter| !filter.columns.is_empty())
    {
        out.push_str("<sheetPr filterMode=\"1\"/>");
    }
    out.push_str(&format!(
        "<dimension ref=\"{}:{}\"/>",
        cell_ref(0, 0),
        cell_ref(max_row, max_col)
    ));
    out.push_str("<sheetViews><sheetView workbookViewId=\"0\"");
    if let Some(zoom) = sheet.zoom {
        let scale = if zoom <= 10.0 { zoom * 100.0 } else { zoom };
        out.push_str(&format!(
            " zoomScale=\"{}\" zoomScaleNormal=\"{}\"",
            scale.round() as u32,
            scale.round() as u32
        ));
    }
    if !sheet.show_gridlines {
        out.push_str(" showGridLines=\"0\"");
    }
    out.push('>');
    if sheet.frozen_rows > 0 || sheet.frozen_cols > 0 {
        let top_row = sheet.frozen_rows;
        let left_col = sheet.frozen_cols;
        let top_left = cell_ref(top_row, left_col);
        out.push_str(&format!("<pane xSplit=\"{}\" ySplit=\"{}\" topLeftCell=\"{}\" activePane=\"bottomRight\" state=\"frozen\"/>", left_col, top_row, top_left));
        out.push_str("<selection pane=\"bottomRight\" activeCell=\"A1\" sqref=\"A1\"/>");
    } else {
        out.push_str("<selection activeCell=\"A1\" sqref=\"A1\"/>");
    }
    out.push_str("</sheetView></sheetViews>");
    let axis = dimensions::prepare(sheet)?;
    out.push_str(&axis.format_xml());
    out.push_str(&axis.columns_xml());
    out.push_str("<sheetData>");
    let mut by_row = BTreeMap::<u32, Vec<CellInput>>::new();
    for cell in cells {
        by_row.entry(cell.row.unwrap_or(0)).or_default().push(cell);
    }
    let hidden_rows = auto_filter::hidden_rows(sheet);
    for row in &hidden_rows {
        by_row.entry(*row).or_default();
    }
    let mut cell_rows = by_row.into_iter().peekable();
    let mut dimension_rows = axis.extra_row_indices().peekable();
    // Standard SpreadsheetML needs a row element for each manually hidden row;
    // stream spans into the XML without constructing a million-row lookup map.
    let mut manual_hidden_rows = axis.manual_hidden_row_indices().peekable();
    while let Some(row) = [cell_rows.peek().map(|(row, _)| *row), dimension_rows.peek().copied(), manual_hidden_rows.peek().copied()]
        .into_iter().flatten().min() {
        let mut row_cells = if cell_rows.peek().is_some_and(|(index, _)| *index == row) { cell_rows.next().unwrap().1 } else { Vec::new() };
        if dimension_rows.peek() == Some(&row) { dimension_rows.next(); }
        if manual_hidden_rows.peek() == Some(&row) { manual_hidden_rows.next(); }
        row_cells.sort_by_key(|cell| cell.col.unwrap_or(0));
        let hidden = hidden_rows.contains(&row) || axis.manual_row_hidden(row) || axis.row_size(row).height == 0.0;
        let mut attrs = format!(" r=\"{}\"", row + 1);
        attrs.push_str(&axis.row_attributes(row));
        if hidden {
            attrs.push_str(" hidden=\"1\"");
        } else if axis.default_row_height == 0.0 {
            attrs.push_str(" hidden=\"0\"");
        }
        out.push_str(&format!("<row{attrs}>"));
        for cell in row_cells {
            out.push_str(&cell_xml(&cell, registry, sheet_index, warnings));
        }
        out.push_str("</row>");
    }
    out.push_str("</sheetData>");
    out.push_str(&auto_filter::write_xml(sheet, warnings)?);
    let mut merges = sheet
        .merged_ranges
        .iter()
        .filter_map(range_ref)
        .collect::<Vec<_>>();
    merges.sort();
    merges.dedup();
    if !merges.is_empty() {
        out.push_str(&format!("<mergeCells count=\"{}\">", merges.len()));
        for range in &merges {
            out.push_str(&format!("<mergeCell ref=\"{}\"/>", range));
        }
        out.push_str("</mergeCells>");
    }
    let conditional = write_conditional_formats(sheet, registry, warnings)?;
    out.push_str(&conditional.base);
    let validation = data_validation::write_xml(sheet, warnings)?;
    out.push_str(&validation);
    if !hyperlinks.is_empty() {
        out.push_str("<hyperlinks>");
        for (link_index, (reference, target)) in hyperlinks.iter().enumerate() {
            if let Some(location) = hyperlink_location(target, sheet_ids, sheet_names) {
                out.push_str(&format!(
                    "<hyperlink ref=\"{}\" location=\"{}\"/>",
                    reference,
                    esc(&location)
                ));
            } else {
                let id = format!("rIdHyperlink{}", link_index + 1);
                out.push_str(&format!(
                    "<hyperlink ref=\"{}\" r:id=\"{}\"/>",
                    reference, id
                ));
            }
        }
        out.push_str("</hyperlinks>");
    }
    if let Some(rel_id) = drawing_rel_id {
        out.push_str(&format!("<drawing r:id=\"{}\"/>", esc(rel_id)));
    }
    let axis_extension = axis.extension_xml(sheet, &out)?;
    let validation_extension = data_validation::extension_xml(sheet, &validation)?;
    let mut extensions = sparkline::write_xml(&sheet.name, &sheet.cells, warnings)?;
    let additions = format!("{}{axis_extension}{validation_extension}", conditional.extension);
    if extensions.is_empty() { extensions = format!("<extLst>{additions}</extLst>"); }
    else if let Some(position) = extensions.rfind("</extLst>") { extensions.insert_str(position, &additions); }
    else { return Err(ExportError::invalid("worksheet extensions have no closing container")); }
    out.push_str(&extensions);
    out.push_str("</worksheet>");
    Ok((out, hyperlinks))
}

fn hyperlink_location(target: &str, sheet_ids: &[String], sheet_names: &[String]) -> Option<String> {
    if !target.starts_with('#') {
        return None;
    }
    let value = target.trim_start_matches('#');
    if !value.starts_with("gid=") {
        return Some(value.to_owned());
    }
    let mut gid = None;
    let mut range = None;
    for (key, value) in value.split('&').filter_map(|pair| pair.split_once('=')) {
        let value = decode_query_component(value)?;
        match key {
            "gid" => gid = Some(value),
            "range" => range = Some(value),
            _ => {}
        }
    }
    let index = sheet_ids.iter().position(|id| Some(id.as_str()) == gid.as_deref())?;
    let name = sheet_names.get(index)?;
    let range = range?;
    let quoted_name = if name.chars().any(|character| character.is_whitespace() || character == '\'') {
        format!("'{}'", name.replace('\'', "''"))
    } else {
        name.clone()
    };
    Some(format!("{quoted_name}!{range}"))
}

fn decode_query_component(value: &str) -> Option<String> {
    let bytes = value.as_bytes();
    let mut decoded = Vec::with_capacity(bytes.len());
    let mut index = 0;
    while index < bytes.len() {
        if bytes[index] == b'%' {
            let hex = std::str::from_utf8(bytes.get(index + 1..index + 3)?).ok()?;
            decoded.push(u8::from_str_radix(hex, 16).ok()?);
            index += 3;
        } else {
            decoded.push(if bytes[index] == b'+' { b' ' } else { bytes[index] });
            index += 1;
        }
    }
    String::from_utf8(decoded).ok()
}

fn write_conditional_formats(sheet: &SheetInput, registry: &mut StyleRegistry, warnings: &mut Vec<ExportWarning>) -> Result<conditional_source::ConditionalOutput, ExportError> {
    conditional_source::write(sheet, registry, warnings)
}

fn write_conditional_rule(out: &mut String, rule: &ConditionalFormatInput, dxf_id: Option<usize>, priority: u32, operator: &str) -> Result<(), ExportError> {
    let text_rule = matches!(rule.rule_type.as_str(), "containsText" | "notContainsText" | "beginsWith" | "endsWith") || rule.sub_type.as_deref() == Some("text");
    let text_operator = match rule.rule_type.as_str() {
        "notContainsText" => "notContainsText", "beginsWith" => "beginsWith", "endsWith" => "endsWith", _ => operator,
    };
    let average = rule.rule_type == "average" || rule.sub_type.as_deref() == Some("average");
    let xml_type = if text_rule {
        match text_operator {
            "equal" | "notEqual" => "expression",
            "notContains" | "notContainsText" => "notContainsText", "beginsWith" => "beginsWith", "endsWith" => "endsWith",
            "containsBlanks" => "containsBlanks", "notContainsBlanks" => "notContainsBlanks", "containsErrors" => "containsErrors",
            "notContainsErrors" => "notContainsErrors", _ => "containsText",
        }
    } else if average && matches!(operator, "equal" | "notEqual") { "expression" }
    else {
        match rule.rule_type.as_str() {
            "cellIs" => "cellIs",
            "highlightCell" => match rule.sub_type.as_deref().unwrap_or("number") {
                "number" => "cellIs", "formula" => "expression", "timePeriod" => "timePeriod", "duplicateValues" => "duplicateValues",
                "uniqueValues" => "uniqueValues", "rank" => "top10", "average" => "aboveAverage",
                _ => return Err(ExportError::invalid("unsupported conditional format subtype")),
            },
            "duplicateValues" => "duplicateValues", "uniqueValues" => "uniqueValues", "timePeriod" => "timePeriod",
            "top10" | "rank" => "top10", "average" => "aboveAverage", "expression" => "expression", "dataBar" => "dataBar",
            "colorScale" => "colorScale", "iconSet" => "iconSet",
            _ => return Err(ExportError::invalid(format!("conditional format {} has an unsupported rule type", rule.id))),
        }
    };
    let visual = matches!(xml_type, "dataBar" | "colorScale" | "iconSet");
    let mut attrs = format!("type=\"{}\" priority=\"{}\"", xml_type, priority);
    if let Some(id) = dxf_id { attrs.push_str(&format!(" dxfId=\"{id}\"")); }
    if xml_type == "cellIs" {
        if !matches!(operator, "equal" | "notEqual" | "greaterThan" | "greaterThanOrEqual" | "lessThan" | "lessThanOrEqual" | "between" | "notBetween") {
            return Err(ExportError::invalid("conditional format has an invalid number operator"));
        }
        attrs.push_str(&format!(" operator=\"{}\"", esc(operator)));
    }
    if xml_type == "top10" {
        if rule.is_bottom { attrs.push_str(" bottom=\"1\""); }
        if rule.is_percent { attrs.push_str(" percent=\"1\""); }
        let rank = rule.value.as_ref().and_then(json_scalar_string).and_then(|value| value.parse::<u32>().ok()).unwrap_or(10);
        if rank == 0 || rank > if rule.is_percent { 100 } else { 1000 } { return Err(ExportError::invalid("conditional format rank is invalid")); }
        attrs.push_str(&format!(" rank=\"{rank}\""));
    }
    if xml_type == "aboveAverage" {
        if matches!(operator, "lessThan" | "lessThanOrEqual") { attrs.push_str(" aboveAverage=\"0\""); }
        if matches!(operator, "greaterThanOrEqual" | "lessThanOrEqual") { attrs.push_str(" equalAverage=\"1\""); }
    }
    if xml_type == "timePeriod" { attrs.push_str(&format!(" timePeriod=\"{}\"", esc(operator))); }
    let text = rule.value.as_ref().and_then(json_scalar_string).unwrap_or_else(|| rule.formula1.clone());
    if matches!(xml_type, "containsText" | "notContainsText" | "beginsWith" | "endsWith") {
        let text_operator = if xml_type == "notContainsText" { "notContains" } else { xml_type };
        attrs.push_str(&format!(" operator=\"{}\" text=\"{}\"", text_operator, esc(&text)));
    }
    if rule.stop_if_true && !visual { attrs.push_str(" stopIfTrue=\"1\""); }
    out.push_str(&format!("<cfRule {attrs}>"));
    match xml_type {
        "cellIs" | "expression" | "containsText" | "notContainsText" | "beginsWith" | "endsWith" | "containsBlanks" | "notContainsBlanks" | "containsErrors" | "notContainsErrors" => {
            let (row, column) = rule.ranges.first().and_then(range_bounds).map(|(row, column, _, _)| (row, column)).unwrap_or((0, 0));
            let address = cell_ref(row, column);
            let literal = text.replace('"', "\"\"");
            let search = text.replace('~', "~~").replace('*', "~*").replace('?', "~?").replace('"', "\"\"");
            let formula = if text_rule {
                match text_operator {
                    "notContains" | "notContainsText" => format!("ISERROR(SEARCH(\"{search}\",{address}))"),
                    "beginsWith" => format!("LEFT({address},LEN(\"{literal}\"))=\"{literal}\""),
                    "endsWith" => format!("RIGHT({address},LEN(\"{literal}\"))=\"{literal}\""),
                    "equal" => format!("{address}&\"\"=\"{literal}\""), "notEqual" => format!("{address}&\"\"<>\"{literal}\""),
                    "containsBlanks" => format!("LEN(TRIM({address}))=0"), "notContainsBlanks" => format!("LEN(TRIM({address}))>0"),
                    "containsErrors" => format!("ISERROR({address})"), "notContainsErrors" => format!("NOT(ISERROR({address}))"),
                    _ => format!("NOT(ISERROR(SEARCH(\"{search}\",{address})))"),
                }
            } else if average { conditional_source::average_predicate(rule, &address)? }
            else { if rule.formula1.trim().is_empty() { rule.value.as_ref().and_then(json_scalar_string).unwrap_or_default() } else { rule.formula1.clone() }.trim().trim_start_matches('=').to_owned() };
            if formula.is_empty() { return Err(ExportError::invalid("conditional format is missing its formula")); }
            out.push_str(&format!("<formula>{}</formula>", esc(&formula)));
            if xml_type == "cellIs" && matches!(operator, "between" | "notBetween") {
                let second = rule.formula2.as_deref().ok_or_else(|| ExportError::invalid("conditional format is missing its second operand"))?;
                if second.trim().trim_start_matches('=').is_empty() { return Err(ExportError::invalid("conditional format has an empty second operand")); }
                out.push_str(&format!("<formula>{}</formula>", esc(second.trim().trim_start_matches('='))));
            }
        }
        "dataBar" => write_data_bar(out, rule),
        "colorScale" => write_color_scale(out, rule),
        "iconSet" => write_icon_set(out, rule),
        _ => {}
    }
    out.push_str("</cfRule>");
    Ok(())
}

fn write_data_bar(out: &mut String, rule: &ConditionalFormatInput) {
    let config = rule.config.as_ref().and_then(Value::as_object);
    let color = normalize_color(config.and_then(|item| item.get("positiveColor")).and_then(Value::as_str).unwrap_or("#638EC6"));
    let min = config.and_then(|item| item.get("min")).and_then(Value::as_object);
    let max = config.and_then(|item| item.get("max")).and_then(Value::as_object);
    let minimum = min.and_then(|item| item.get("type")).and_then(Value::as_str).unwrap_or("min");
    let maximum = max.and_then(|item| item.get("type")).and_then(Value::as_str).unwrap_or("max");
    let minimum = if minimum == "autoMin" { "min" } else { minimum };
    let maximum = if maximum == "autoMax" { "max" } else { maximum };
    let show_value = config.and_then(|item| item.get("isShowValue")).and_then(Value::as_bool).unwrap_or(true);
    let min_value = min.and_then(|item| item.get("value")).and_then(json_scalar_string);
    let max_value = max.and_then(|item| item.get("value")).and_then(json_scalar_string);
    let min_attr = min_value.map(|value| format!(" val=\"{}\"", esc(&value))).unwrap_or_default();
    let max_attr = max_value.map(|value| format!(" val=\"{}\"", esc(&value))).unwrap_or_default();
    out.push_str(&format!("<dataBar showValue=\"{}\"><cfvo type=\"{}\"{} /><cfvo type=\"{}\"{} />", if show_value { 1 } else { 0 }, esc(minimum), min_attr, esc(maximum), max_attr));
    out.push_str(&format!("<color rgb=\"{}\"/></dataBar>", color));
}

fn write_color_scale(out: &mut String, rule: &ConditionalFormatInput) {
    let thresholds = rule.config.as_ref().and_then(|value| value.get("thresholds")).and_then(Value::as_array);
    let Some(thresholds) = thresholds else { return };
    out.push_str("<colorScale>");
    for threshold in thresholds {
        let value = threshold.get("value");
        let kind = value.and_then(|item| item.get("type")).and_then(Value::as_str).unwrap_or("percentile");
        let amount = value.and_then(|item| item.get("value")).and_then(json_scalar_string).unwrap_or_default();
        out.push_str(&format!("<cfvo type=\"{}\"{} />", esc(kind), if amount.is_empty() { String::new() } else { format!(" val=\"{}\"", esc(&amount)) }));
    }
    for threshold in thresholds {
        let color = normalize_color(threshold.get("color").and_then(Value::as_str).unwrap_or("#FFFFFF"));
        out.push_str(&format!("<color rgb=\"{}\"/>", color));
    }
    out.push_str("</colorScale>");
}

fn write_icon_set(out: &mut String, rule: &ConditionalFormatInput) {
    let icons = rule.config.as_ref().and_then(|value| value.get("icons")).and_then(Value::as_array);
    let Some(icons) = icons else { return };
    let config = rule.config.as_ref().and_then(Value::as_object);
    let icon_set = config.and_then(|item| item.get("iconSet")).and_then(Value::as_str).unwrap_or("3TrafficLights1");
    let show_value = config.and_then(|item| item.get("isShowValue")).and_then(Value::as_bool).unwrap_or(true);
    let reverse = config.and_then(|item| item.get("reverse")).and_then(Value::as_bool).unwrap_or(false);
    out.push_str(&format!("<iconSet iconSet=\"{}\" showValue=\"{}\" reverse=\"{}\">", esc(icon_set), if show_value { 1 } else { 0 }, if reverse { 1 } else { 0 }));
    for (position, icon) in icons.iter().rev().enumerate() {
        if position == 0 { out.push_str("<cfvo type=\"percent\" val=\"0\"/>"); continue }
        let value = icon.get("value");
        let kind = value.and_then(|item| item.get("type")).and_then(Value::as_str).unwrap_or("percent");
        let amount = value.and_then(|item| item.get("value")).and_then(json_scalar_string).unwrap_or_default();
        out.push_str(&format!("<cfvo type=\"{}\"{}{} />", esc(kind), if amount.is_empty() { String::new() } else { format!(" val=\"{}\"", esc(&amount)) },
            if icon.get("operator").and_then(Value::as_str) == Some("greaterThan") { " gte=\"0\"" } else { "" }));
    }
    out.push_str("</iconSet>");
}

fn json_scalar_string(value: &Value) -> Option<String> {
    match value { Value::String(text) => Some(text.clone()), Value::Number(number) => Some(number.to_string()), _ => None }
}

fn range_bounds(range: &RangeInput) -> Option<(u32, u32, u32, u32)> {
    let start_row = range.start_row.or(range.first_row)?;
    let start_col = range.start_col.or(range.first_col)?;
    let end_row = range.end_row.or(range.last_row).unwrap_or(start_row);
    let end_col = range.end_col.or(range.last_col).unwrap_or(start_col);
    Some((start_row, start_col, end_row, end_col))
}

fn chart_range_formula(
    sheet_name: &str,
    first_row: u32,
    first_col: u32,
    last_row: u32,
    last_col: u32,
) -> String {
    let escaped = sheet_name.replace('\'', "''");
    format!(
        "'{}'!${}${}:${}${}",
        escaped,
        column_name(first_col),
        first_row + 1,
        column_name(last_col),
        last_row + 1
    )
}

fn chart_cell_value(sheet: &SheetInput, row: u32, col: u32) -> Option<Value> {
    let key = format!("{row}:{col}");
    if let Some(value) = sheet.cells.get(&key) {
        let cell = cell_from_entry(&key, value)?;
        return cell
            .value
            .or_else(|| (!cell.display.is_empty()).then(|| Value::String(cell.display)));
    }
    let key = cell_ref(row, col);
    let value = sheet.cells.get(&key)?;
    let cell = cell_from_entry(&key, value)?;
    cell.value
        .or_else(|| (!cell.display.is_empty()).then(|| Value::String(cell.display)))
}

fn chart_value_text(value: Option<Value>, display: bool) -> String {
    match value {
        Some(Value::Number(number)) => number.to_string(),
        Some(Value::Bool(value)) => {
            if value {
                "1".to_owned()
            } else {
                "0".to_owned()
            }
        }
        Some(Value::String(value)) => value,
        Some(Value::Object(map)) if display => map
            .get("error")
            .and_then(Value::as_str)
            .unwrap_or("")
            .to_owned(),
        _ => String::new(),
    }
}

fn chart_number_text(value: Option<Value>) -> String {
    let Some(value) = value else {
        return String::new();
    };
    match value {
        Value::Number(number) => number
            .as_f64()
            .filter(|number| number.is_finite())
            .map(|number| number.to_string())
            .unwrap_or_default(),
        Value::String(text) => text
            .trim()
            .parse::<f64>()
            .ok()
            .filter(|number| number.is_finite())
            .map(|number| number.to_string())
            .unwrap_or_default(),
        _ => String::new(),
    }
}

fn emu(value: f64) -> Result<u32, ExportError> {
    if !value.is_finite() || value < 0.0 || value > u32::MAX as f64 / 9525.0 {
        return Err(ExportError::invalid(
            "chart position and size must fit DrawingML coordinate limits",
        ));
    }
    Ok((value * 9525.0).round() as u32)
}

fn drawing_xml(charts: &[(&ChartInput, usize)]) -> Result<String, ExportError> {
    let mut out = format!(
        "{}<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">",
        xml::DECL
    );
    for (local_index, (chart, _global_index)) in charts.iter().enumerate() {
        if chart.size.width <= 0.0 || chart.size.height <= 0.0 {
            return Err(ExportError::invalid(format!(
                "chart {} size must be greater than zero",
                chart.id
            )));
        }
        let x = emu(chart.position.x)?;
        let y = emu(chart.position.y)?;
        let cx = emu(chart.size.width)?;
        let cy = emu(chart.size.height)?;
        out.push_str(&format!("<xdr:absoluteAnchor><xdr:pos x=\"{}\" y=\"{}\"/><xdr:ext cx=\"{}\" cy=\"{}\"/><xdr:graphicFrame macro=\"\"><xdr:nvGraphicFramePr><xdr:cNvPr id=\"{}\" name=\"Chart {}\"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{}\" cy=\"{}\"/></xdr:xfrm><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rIdChart{}\"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:absoluteAnchor>", x, y, cx, cy, local_index + 1, local_index + 1, cx, cy, local_index + 1));
    }
    out.push_str("</xdr:wsDr>");
    Ok(out)
}

fn cell_xml(
    cell: &CellInput,
    registry: &mut StyleRegistry,
    _sheet_index: usize,
    warnings: &mut Vec<ExportWarning>,
) -> String {
    let row = cell.row.unwrap_or(0);
    let col = cell.col.unwrap_or(0);
    let reference = cell_ref(row, col);
    let style_id = registry.register(cell.style.as_ref());
    let mut out = format!("<c r=\"{}\" s=\"{}\"", reference, style_id);
    let formula = cell.formula.as_deref().filter(|value| !value.is_empty());
    let rich_text = if formula.is_none() {
        rich_runs(&cell.rich_text)
    } else {
        None
    };
    if formula.is_some() {
        let cached_kind = value_kind(
            cell.value.as_ref(),
            cell.cached_data_type
                .as_deref()
                .or(cell.data_type.as_deref()),
            !cell.display.is_empty(),
        );
        match cached_kind.as_str() {
            "str" => out.push_str(" t=\"str\""),
            "b" => out.push_str(" t=\"b\""),
            "e" => out.push_str(" t=\"e\""),
            _ => {}
        }
        out.push('>');
        let expression = formula.unwrap_or_default().trim_start_matches('=');
        out.push_str(&format!("<f>{}</f>", esc(expression)));
        let cached = if cached_kind == "e" {
            error_text(cell.value.as_ref(), &cell.display)
        } else {
            cell.value
                .as_ref()
                .and_then(value_text)
                .or_else(|| (!cell.display.is_empty()).then(|| cell.display.clone()))
        };
        if let Some(cached) = cached {
            out.push_str(&format!("<v>{}</v>", esc(&cached)));
        } else {
            out.push_str("<v/>");
        }
        out.push_str("</c>");
        return out;
    }
    if let Some(rich) = rich_text {
        out.push_str(" t=\"inlineStr\"><is>");
        out.push_str(&rich);
        out.push_str("</is></c>");
        return out;
    }
    let kind = value_kind(
        cell.value.as_ref(),
        cell.data_type.as_deref(),
        !cell.display.is_empty(),
    );
    match (kind.as_str(), cell.value.as_ref()) {
        ("str", Some(Value::String(value))) => {
            out.push_str(" t=\"inlineStr\"><is><t xml:space=\"preserve\">");
            out.push_str(&esc(value));
            out.push_str("</t></is></c>");
        }
        ("str", _) if !cell.display.is_empty() => {
            out.push_str(" t=\"inlineStr\"><is><t xml:space=\"preserve\">");
            out.push_str(&esc(&cell.display));
            out.push_str("</t></is></c>");
        }
        ("b", Some(Value::Bool(value))) => out.push_str(&format!(
            " t=\"b\"><v>{}</v></c>",
            if *value { 1 } else { 0 }
        )),
        ("e", value) => {
            out.push_str(" t=\"e\"><v>");
            if let Some(error) = error_text(value, &cell.display) {
                out.push_str(&esc(&error));
            }
            out.push_str("</v></c>");
        }
        ("n", Some(value)) => out.push_str(&format!(
            " t=\"n\"><v>{}</v></c>",
            esc(&value_text(value).unwrap_or_default())
        )),
        _ if cell.value.is_none() && cell.display.is_empty() => out.push_str("/>"),
        _ => {
            if !cell.display.is_empty() {
                out.push_str(" t=\"inlineStr\"><is><t xml:space=\"preserve\">");
                out.push_str(&esc(&cell.display));
                out.push_str("</t></is></c>");
            } else {
                warnings.push(ExportWarning::new(
                    "UNSUPPORTED_CELL_VALUE",
                    format!("Cell {reference} has no serialisable value"),
                ));
                out.push_str("/>");
            }
        }
    }
    out
}

fn value_kind(value: Option<&Value>, data_type: Option<&str>, has_display: bool) -> String {
    if let Some(data_type) = data_type {
        let lower = data_type.to_ascii_lowercase();
        if lower == "b" || lower.contains("bool") {
            return "b".to_owned();
        }
        if lower == "e" || lower.contains("error") {
            return "e".to_owned();
        }
        if lower == "n" || lower.contains("number") || lower.contains("date") {
            return "n".to_owned();
        }
        if lower == "s" || lower.contains("string") || lower.contains("text") {
            return "str".to_owned();
        }
    }
    match value {
        Some(Value::Bool(_)) => "b",
        Some(Value::Number(_)) => "n",
        Some(Value::String(_)) => "str",
        Some(Value::Object(map)) if map.contains_key("error") => "e",
        _ if has_display => "str",
        _ => "str",
    }
    .to_owned()
}

fn value_text(value: &Value) -> Option<String> {
    match value {
        Value::String(value) => Some(value.clone()),
        Value::Number(value) => Some(value.to_string()),
        Value::Bool(value) => Some(if *value { "1" } else { "0" }.to_owned()),
        _ => None,
    }
}

fn error_text(value: Option<&Value>, display: &str) -> Option<String> {
    match value {
        Some(Value::String(value)) => Some(value.clone()),
        Some(Value::Object(map)) => map.get("error").and_then(value_text),
        _ if !display.is_empty() => Some(display.to_owned()),
        _ => None,
    }
}

fn rich_runs(runs: &[Value]) -> Option<String> {
    if runs.is_empty() {
        return None;
    }
    let mut out = String::new();
    for run in runs {
        let Some(map) = run.as_object() else {
            continue;
        };
        let text = map.get("text").and_then(Value::as_str).unwrap_or("");
        let font = map.get("font").and_then(Value::as_object);
        out.push_str("<r>");
        if let Some(font) = font {
            let bold = font.get("bold").and_then(Value::as_bool).unwrap_or(false);
            let italic = font.get("italic").and_then(Value::as_bool).unwrap_or(false);
            let underline = font
                .get("underline")
                .and_then(Value::as_bool)
                .unwrap_or(false);
            let strike = font
                .get("strike")
                .or_else(|| font.get("strikethrough"))
                .and_then(Value::as_bool)
                .unwrap_or(false);
            let family = font
                .get("family")
                .or_else(|| font.get("name"))
                .and_then(Value::as_str);
            let size = font
                .get("size")
                .or_else(|| font.get("sizePt"))
                .and_then(Value::as_f64);
            let color = font.get("color").and_then(Value::as_str);
            if bold
                || italic
                || underline
                || strike
                || family.is_some()
                || size.is_some()
                || color.is_some()
            {
                out.push_str("<rPr>");
                if bold {
                    out.push_str("<b/>");
                }
                if italic {
                    out.push_str("<i/>");
                }
                if underline {
                    out.push_str("<u/>");
                }
                if strike {
                    out.push_str("<strike/>");
                }
                if let Some(family) = family {
                    out.push_str(&format!("<rFont val=\"{}\"/>", esc(family)));
                }
                if let Some(size) = size {
                    out.push_str(&format!("<sz val=\"{}\"/>", format_number(size)));
                }
                if let Some(color) = color {
                    out.push_str(&format!("<color rgb=\"{}\"/>", normalize_color(color)));
                }
                out.push_str("</rPr>");
            }
        }
        out.push_str(&format!("<t xml:space=\"preserve\">{}</t></r>", esc(text)));
    }
    Some(out)
}

fn workbook_xml(workbook: &WorkbookInput) -> String {
    let sheet_names = safe_sheet_names(workbook.sheets.iter().map(|sheet| sheet.name.as_str()));
    let active_index = workbook
        .active_sheet_id
        .as_ref()
        .and_then(|id| workbook.sheets.iter().position(|sheet| &sheet.id == id))
        .unwrap_or(0);
    let mut out = String::from(xml::DECL);
    out.push_str(&format!(
        "<workbook xmlns=\"{MAIN_NS}\" xmlns:r=\"{REL_NS}\">"
    ));
    out.push_str(&format!("<fileVersion appName=\"xl\" lastEdited=\"7\" lowestEdited=\"7\" rupBuild=\"1\"/><workbookPr date1904=\"{}\"/>", if workbook.date1904 { 1 } else { 0 }));
    out.push_str(&format!("<bookViews><workbookView xWindow=\"0\" yWindow=\"0\" windowWidth=\"24000\" windowHeight=\"12000\" activeTab=\"{}\"/></bookViews><sheets>", active_index));
    for (index, (sheet, name)) in workbook.sheets.iter().zip(&sheet_names).enumerate() {
        let visibility = if sheet.visible { "visible" } else { "hidden" };
        out.push_str(&format!(
            "<sheet name=\"{}\" sheetId=\"{}\" state=\"{}\" r:id=\"rIdSheet{}\"/>",
            esc(name),
            index + 1,
            visibility,
            index + 1
        ));
    }
    out.push_str("</sheets>");
    let pivot_count = workbook.pivot_tables.as_ref().map_or(0, Vec::len);
    if pivot_count > 0 {
        out.push_str(&format!("<pivotCaches count=\"{pivot_count}\">"));
        for index in 0..pivot_count {
            out.push_str(&format!("<pivotCache cacheId=\"{}\" r:id=\"rIdPivotCache{}\"/>", index + 1, index + 1));
        }
        out.push_str("</pivotCaches>");
    }
    if let Some(Value::Object(metadata)) = workbook.metadata.as_ref() {
        let mut names = Vec::<(String, String, Option<String>)>::new();
        if let Some(items) = metadata.get("definedNames").and_then(Value::as_array) {
            for item in items {
                if let Some(map) = item.as_object() {
                    if let (Some(name), Some(reference)) = (
                        map.get("name").and_then(Value::as_str),
                        map.get("refersTo")
                            .or_else(|| map.get("refers_to"))
                            .and_then(Value::as_str),
                    ) {
                        names.push((name.to_owned(), reference.to_owned(), None));
                    }
                }
            }
        }
        if let Some(items) = metadata.get("localDefinedNames").and_then(Value::as_array) {
            for item in items {
                if let Some(map) = item.as_object() {
                    if let (Some(name), Some(reference), Some(sheet)) = (
                        map.get("name").and_then(Value::as_str),
                        map.get("refersTo")
                            .or_else(|| map.get("refers_to"))
                            .and_then(Value::as_str),
                        map.get("sheet").and_then(Value::as_str),
                    ) {
                        names.push((
                            name.to_owned(),
                            reference.to_owned(),
                            Some(sheet.to_owned()),
                        ));
                    }
                }
            }
        }
        if !names.is_empty() {
            out.push_str("<definedNames>");
            for (name, reference, sheet) in names {
                if let Some(sheet) = sheet {
                    let index = workbook
                        .sheets
                        .iter()
                        .position(|item| item.name == sheet)
                        .or_else(|| {
                            sheet_names
                                .iter()
                                .position(|safe_name| safe_name.as_str() == sheet)
                        });
                    if let Some(index) = index {
                        out.push_str(&format!(
                            "<definedName name=\"{}\" localSheetId=\"{}\">{}</definedName>",
                            esc(&name),
                            index,
                            esc(&reference)
                        ));
                    } else {
                        out.push_str(&format!(
                            "<definedName name=\"{}\">{}</definedName>",
                            esc(&name),
                            esc(&reference)
                        ));
                    }
                } else {
                    out.push_str(&format!(
                        "<definedName name=\"{}\">{}</definedName>",
                        esc(&name),
                        esc(&reference)
                    ));
                }
            }
            out.push_str("</definedNames>");
        }
    }
    out.push_str("<calcPr calcId=\"191029\" fullCalcOnLoad=\"1\" forceFullCalc=\"1\"/></workbook>");
    out
}

fn safe_sheet_name(name: &str, index: usize) -> String {
    let mut value = name
        .trim()
        .replace([':', '\\', '/', '?', '*', '[', ']'], "_");
    if value.is_empty() {
        value = format!("Sheet{}", index + 1);
    }
    value.chars().take(31).collect()
}

fn safe_sheet_names<'a>(names: impl IntoIterator<Item = &'a str>) -> Vec<String> {
    let mut used = BTreeSet::new();
    names
        .into_iter()
        .enumerate()
        .map(|(index, name)| {
            let base = safe_sheet_name(name, index);
            if used.insert(base.to_lowercase()) {
                return base;
            }
            for suffix_index in 2usize.. {
                let suffix = format!(" ({suffix_index})");
                let prefix = base
                    .chars()
                    .take(31 - suffix.chars().count())
                    .collect::<String>();
                let candidate = format!("{prefix}{suffix}");
                if used.insert(candidate.to_lowercase()) {
                    return candidate;
                }
            }
            unreachable!("a unique worksheet name is available")
        })
        .collect()
}

fn core_properties_xml(workbook: &WorkbookInput) -> String {
    let props = workbook
        .metadata
        .as_ref()
        .and_then(Value::as_object)
        .and_then(|map| map.get("properties"))
        .and_then(Value::as_object);
    let value = |name: &str| {
        props
            .and_then(|map| map.get(name))
            .and_then(Value::as_str)
            .unwrap_or("")
    };
    format!(
        "{}<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"><dc:title>{}</dc:title><dc:creator>{}</dc:creator><cp:lastModifiedBy>{}</cp:lastModifiedBy><dc:subject>{}</dc:subject><cp:keywords>{}</cp:keywords><dc:description>{}</dc:description></cp:coreProperties>",
        xml::DECL,
        esc(value("title")),
        esc(value("creator")),
        esc(value("lastModifiedBy")),
        esc(value("subject")),
        esc(value("keywords")),
        esc(value("description"))
    )
}

fn app_properties_xml(workbook: &WorkbookInput) -> String {
    let sheet_names = safe_sheet_names(workbook.sheets.iter().map(|sheet| sheet.name.as_str()));
    let names = workbook
        .sheets
        .iter()
        .zip(&sheet_names)
        .map(|(_sheet, name)| format!("<vt:lpstr>{}</vt:lpstr>", esc(name)))
        .collect::<String>();
    format!(
        "{}<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\"><Application>Office Viewer Rust Writer</Application><AppVersion>1.0</AppVersion><HeadingPairs><vt:vector size=\"2\" baseType=\"variant\"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>{}</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size=\"{}\" baseType=\"lpstr\">{}</vt:vector></TitlesOfParts></Properties>",
        xml::DECL,
        workbook.sheets.len(),
        workbook.sheets.len(),
        names
    )
}

fn add_unsupported_warnings(workbook: &WorkbookInput, warnings: &mut Vec<ExportWarning>) {
    let unsupported = [
        "slicers",
        "conditionalFormatting",
        "tables",
        "drawings",
        "comments",
        "images",
    ];
    let mut seen = BTreeSet::new();
    let mut inspect = |value: Option<&Value>, seen: &mut BTreeSet<String>| {
        if let Some(Value::Object(map)) = value {
            for key in unsupported {
                if map.contains_key(key) && seen.insert(key.to_owned()) {
                    warnings.push(ExportWarning::new("UNSUPPORTED_FEATURE", format!("Workbook feature {key} is not represented by the current canonical model and was not exported")));
                }
            }
        }
    };
    inspect(workbook.metadata.as_ref(), &mut seen);
    for sheet in &workbook.sheets {
        inspect(sheet.metadata.as_ref(), &mut seen);
        for cell in sheet.cells.values() {
            if let Some(map) = cell.as_object() {
                inspect(map.get("metadata"), &mut seen);
            }
        }
    }
}

/// Serialize the canonical workbook JSON into XLSX parts.
pub(crate) fn write(data: &str, package: &mut Package) -> Result<Vec<ExportWarning>, ExportError> {
    let mut workbook: WorkbookInput = serde_json::from_str(data)
        .map_err(|error| ExportError::invalid(format!("workbook JSON is invalid: {error}")))?;
    if workbook
        .kind
        .as_deref()
        .is_some_and(|kind| kind != "workbook")
    {
        return Err(ExportError::invalid(
            "XLSX writer received a non-workbook model",
        ));
    }
    if workbook.sheets.is_empty() {
        return Err(ExportError::invalid(
            "workbook must contain at least one worksheet",
        ));
    }
    if workbook.metadata.as_ref().and_then(|metadata| metadata.get("unresolvedOfficeCharts")).and_then(Value::as_array).is_some_and(|charts| !charts.is_empty()) {
        return Err(ExportError::unsupported("the imported workbook contains unresolved native charts; their original source is retained and cannot be silently omitted from XLSX export"));
    }
    if workbook.sheets.iter().any(|sheet| {
        sheet
            .row_count
            .is_some_and(|count| count > MAX_ROW_INDEX + 1)
            || sheet
                .column_count
                .is_some_and(|count| count > MAX_COLUMN_INDEX + 1)
            || sheet.frozen_rows > MAX_ROW_INDEX
            || sheet.frozen_cols > MAX_COLUMN_INDEX
            || sheet.hidden_rows.iter().any(|row| *row > MAX_ROW_INDEX)
            || sheet
                .hidden_columns
                .iter()
                .any(|col| *col > MAX_COLUMN_INDEX)
            || sheet
                .row_heights
                .keys()
                .filter_map(|key| parse_index(key))
                .any(|row| row > MAX_ROW_INDEX)
            || sheet
                .column_widths
                .keys()
                .filter_map(|key| parse_index(key))
                .any(|col| col > MAX_COLUMN_INDEX)
            || sheet.cells.iter().any(|(key, value)| {
                cell_from_entry(key, value).is_some_and(|cell| {
                    cell.row.is_some_and(|row| row > MAX_ROW_INDEX)
                        || cell.col.is_some_and(|col| col > MAX_COLUMN_INDEX)
                })
            })
            || sheet.merged_ranges.iter().any(|range| {
                range
                    .start_row
                    .or(range.first_row)
                    .is_some_and(|row| row > MAX_ROW_INDEX)
                    || range
                        .end_row
                        .or(range.last_row)
                        .is_some_and(|row| row > MAX_ROW_INDEX)
                    || range
                        .start_col
                        .or(range.first_col)
                        .is_some_and(|col| col > MAX_COLUMN_INDEX)
                    || range
                        .end_col
                        .or(range.last_col)
                        .is_some_and(|col| col > MAX_COLUMN_INDEX)
            })
    }) {
        return Err(ExportError::invalid(
            "worksheet coordinates exceed the XLSX row or column limit",
        ));
    }
    let mut validation_budget = office_io_common::sheet_validation_source::ValidationSourceBudget::default();
    for sheet in &workbook.sheets {
        auto_filter::validate(sheet)?;
        data_validation::validate_with_budget(sheet, &mut validation_budget)?;
    }
    let conditional_theme = conditional_source::theme_for_workbook(&workbook)?;
    let mut warnings = Vec::new();
    chart::prepare_native_bindings(&mut workbook, &mut warnings)?;
    add_unsupported_warnings(&workbook, &mut warnings);
    let mut registry = StyleRegistry::new();
    registry.register(None);
    let sheet_ids = workbook.sheets.iter().map(|sheet| sheet.id.clone()).collect::<Vec<_>>();
    let charts = workbook.charts.as_deref().unwrap_or(&[]);
    let mut chart_index_by_sheet = HashMap::<String, Vec<(usize, usize)>>::new();
    for (chart_index, chart) in charts.iter().enumerate() {
        let sheet_index = workbook
            .sheets
            .iter()
            .position(|sheet| sheet.id == chart.sheet_id)
            .ok_or_else(|| {
                ExportError::invalid(format!(
                    "chart {} references an unknown worksheet",
                    chart.id
                ))
            })?;
        chart_index_by_sheet
            .entry(chart.sheet_id.clone())
            .or_default()
            .push((sheet_index, chart_index));
    }
    let pivot_tables = workbook.pivot_tables.as_deref().unwrap_or(&[]);
    let sheet_names = safe_sheet_names(workbook.sheets.iter().map(|sheet| sheet.name.as_str()));
    let workbook_filename = format!("{}.xlsx", safe_sheet_name(&workbook.name, 0));
    let mut pivot_exports = Vec::<(String, String, usize, usize, pivot::PivotParts)>::new();
    for (pivot_index, pivot_table) in pivot_tables.iter().enumerate() {
        let source_index = workbook.sheets.iter().position(|sheet| {
            sheet.id == pivot_table.source.sub_unit_id || sheet.name == pivot_table.source.sheet_name
        }).ok_or_else(|| ExportError::invalid(format!("pivot table {} references an unknown source worksheet", pivot_table.id)))?;
        let target_index = workbook.sheets.iter().position(|sheet| sheet.id == pivot_table.anchor.sub_unit_id).ok_or_else(|| ExportError::invalid(format!("pivot table {} references an unknown target worksheet", pivot_table.id)))?;
        let source_bounds = pivot_table.source.range.bounds().ok_or_else(|| ExportError::invalid(format!("pivot table {} has no source range", pivot_table.id)))?;
        let output_bounds = pivot_table.output_range.as_ref().and_then(pivot::RangeInput::bounds).ok_or_else(|| ExportError::invalid(format!("pivot table {} has no persisted output range", pivot_table.id)))?;
        if (output_bounds.0, output_bounds.1) != (pivot_table.anchor.row, pivot_table.anchor.col) {
            return Err(ExportError::invalid(format!("pivot table {} output range does not match its anchor", pivot_table.id)));
        }
        let table_name = format!("PivotTable{}", pivot_index + 1);
        let parts = pivot::build_parts(pivot_table, &sheet_names[source_index], output_bounds, &workbook.sheets[source_index].cells, source_bounds, &table_name, pivot_index + 1)?;
        let cache_path = format!("xl/pivotCache/pivotCacheDefinition{}.xml", pivot_index + 1);
        let records_path = format!("xl/pivotCache/pivotCacheRecords{}.xml", pivot_index + 1);
        let table_path = format!("xl/pivotTables/pivotTable{}.xml", pivot_index + 1);
        package.add_xml(&cache_path, "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml", parts.cache_definition.clone())?;
        package.add_xml(&records_path, "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml", parts.cache_records.clone())?;
        package.add_xml(&table_path, "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml", parts.table_definition.clone())?;
        package.add_relationship(Some(&cache_path), "rIdRecords", REL_PIVOT_CACHE_RECORDS, &format!("pivotCacheRecords{}.xml", pivot_index + 1), false)?;
        package.add_relationship(Some(&table_path), "rIdCache", REL_PIVOT_CACHE_DEFINITION, &format!("../pivotCache/pivotCacheDefinition{}.xml", pivot_index + 1), false)?;
        package.add_relationship(Some("xl/workbook.xml"), &format!("rIdPivotCache{}", pivot_index + 1), REL_PIVOT_CACHE_DEFINITION, &format!("pivotCache/pivotCacheDefinition{}.xml", pivot_index + 1), false)?;
        package.add_relationship(Some(&format!("xl/worksheets/sheet{}.xml", target_index + 1)), &format!("rIdPivotTable{}", pivot_index + 1), REL_PIVOT_TABLE, &format!("../pivotTables/pivotTable{}.xml", pivot_index + 1), false)?;
        pivot_exports.push((pivot_table.id.clone(), table_name, source_index, target_index, parts));
    }
    let mut table_exports = Vec::<(usize, usize, table::TableInput, String)>::new();
    for (table_index, table) in workbook.tables.iter().enumerate() {
        let sheet_index = workbook.sheets.iter().position(|sheet| sheet.id == table.sheet_id)
            .ok_or_else(|| ExportError::invalid(format!("table {} references an unknown worksheet", table.id)))?;
        table::validate(&workbook.tables, table)?;
        let path = format!("xl/tables/table{}.xml", table_index + 1);
        package.add_xml(&path, "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml", table::write(table, table_index + 1)?)?;
        table_exports.push((sheet_index, table_index, table.clone(), path));
    }
    let mut conditional_source_bytes = 0usize;
    let mut conditional_source_elements = 0usize;
    for (index, sheet) in workbook.sheets.iter().enumerate() {
        let sheet_charts = chart_index_by_sheet
            .get(&sheet.id)
            .map(|items| {
                items
                    .iter()
                    .map(|(_, chart_index)| (&charts[*chart_index], *chart_index))
                    .collect::<Vec<_>>()
            })
            .unwrap_or_default();
        let drawing_rel_id = (!sheet_charts.is_empty()).then_some("rIdDrawing");
        let (mut content, hyperlinks) =
            worksheet_xml(sheet, index, &sheet_ids, &sheet_names, &mut registry, &mut warnings, drawing_rel_id)?;
        let sheet_pivots = pivot_exports.iter().enumerate().filter(|(_, item)| item.3 == index).collect::<Vec<_>>();
        if !sheet_pivots.is_empty() {
            let mut pivot_parts = format!("<pivotTableParts count=\"{}\">", sheet_pivots.len());
            for (pivot_index, _) in &sheet_pivots { pivot_parts.push_str(&format!("<pivotTablePart r:id=\"rIdPivotTable{}\"/>", pivot_index + 1)); }
            pivot_parts.push_str("</pivotTableParts>");
            let insertion = content.find("<extLst").unwrap_or_else(|| content.find("</worksheet>").unwrap_or(content.len()));
            content.insert_str(insertion, &pivot_parts);
        }
        let sheet_tables = table_exports.iter().filter(|item| item.0 == index).collect::<Vec<_>>();
        if !sheet_tables.is_empty() {
            let mut table_parts = format!("<tableParts count=\"{}\">", sheet_tables.len());
            for (_, table_index, _, _) in &sheet_tables {
                table_parts.push_str(&format!("<tablePart r:id=\"rIdTable{}\"/>", table_index + 1));
            }
            table_parts.push_str("</tableParts>");
            let insertion = content.find("<extLst").unwrap_or_else(|| content.find("</worksheet>").unwrap_or(content.len()));
            content.insert_str(insertion, &table_parts);
        }
        let (source_bytes, source_elements) = conditional_source::output_source_usage(&content)?;
        conditional_source_bytes = conditional_source_bytes.checked_add(source_bytes)
            .ok_or_else(|| ExportError::invalid("CF_SOURCE_TOTAL_LIMIT: conditional-format workbook output byte counter overflow"))?;
        conditional_source_elements = conditional_source_elements.checked_add(source_elements)
            .ok_or_else(|| ExportError::invalid("CF_SOURCE_TOTAL_LIMIT: conditional-format workbook output element counter overflow"))?;
        if conditional_source_bytes > office_io_common::conditional_format_source::SOURCE_RESOURCE_BYTE_LIMIT
            || conditional_source_elements > office_io_common::conditional_format_source::SOURCE_RESOURCE_ELEMENT_LIMIT {
            return Err(ExportError::invalid("CF_SOURCE_TOTAL_LIMIT: conditional-format worksheet output resources exceed the workbook budget"));
        }
        let path = format!("xl/worksheets/sheet{}.xml", index + 1);
        package.add_xml(
            &path,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
            content,
        )?;
        package.add_relationship(
            Some("xl/workbook.xml"),
            &format!("rIdSheet{}", index + 1),
            REL_WORKSHEET,
            &format!("worksheets/sheet{}.xml", index + 1),
            false,
        )?;
        for (_, table_index, _, table_path) in &sheet_tables {
            package.add_relationship(Some(&path), &format!("rIdTable{}", table_index + 1), "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table", &format!("../tables/table{}.xml", table_index + 1), false)?;
            let _ = table_path;
        }
        if !sheet_charts.is_empty() {
            let drawing_path = format!("xl/drawings/drawing{}.xml", index + 1);
            package.add_xml(
                &drawing_path,
                "application/vnd.openxmlformats-officedocument.drawing+xml",
                drawing_xml(&sheet_charts)?,
            )?;
            package.add_relationship(
                Some(&path),
                "rIdDrawing",
                REL_DRAWING,
                &format!("../drawings/drawing{}.xml", index + 1),
                false,
            )?;
            for (local_index, (_sheet_index, chart_index)) in chart_index_by_sheet
                .get(&sheet.id)
                .into_iter()
                .flat_map(|items| items.iter())
                .enumerate()
            {
                let chart = &charts[*chart_index];
                package.add_xml(
                    &format!("xl/charts/chart{}.xml", chart_index + 1),
                    "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
                    chart_xml(&workbook, chart, *chart_index, chart.pivot_table_id.as_deref().and_then(|id| pivot_exports.iter().find(|item| item.0 == id).map(|item| format!("[{}]{}!{}", workbook_filename, sheet_names[item.3], item.1))).as_deref())?,
                )?;
                package.add_relationship(
                    Some(&drawing_path),
                    &format!("rIdChart{}", local_index + 1),
                    REL_CHART,
                    &format!("../charts/chart{}.xml", chart_index + 1),
                    false,
                )?;
            }
        }
        for (link_index, (_reference, target)) in hyperlinks.iter().enumerate() {
            if !target.starts_with('#') {
                package.add_relationship(
                    Some(&path),
                    &format!("rIdHyperlink{}", link_index + 1),
                    REL_HYPERLINK,
                    target,
                    true,
                )?;
            }
        }
    }
    conditional_source::validate_output_resource_budget(
        conditional_source_bytes, conditional_source_elements, &registry, conditional_theme.as_deref(),
    )?;
    package.add_xml(
        "xl/styles.xml",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
        style_xml(&registry),
    )?;
    package.add_xml(
        "xl/workbook.xml",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
        workbook_xml(&workbook),
    )?;
    package.add_xml(
        "docProps/core.xml",
        "application/vnd.openxmlformats-package.core-properties+xml",
        core_properties_xml(&workbook),
    )?;
    package.add_xml(
        "docProps/app.xml",
        "application/vnd.openxmlformats-officedocument.extended-properties+xml",
        app_properties_xml(&workbook),
    )?;
    package.add_relationship(None, "rIdWorkbook", REL_WORKBOOK, "xl/workbook.xml", false)?;
    package.add_relationship(None, "rIdCore", REL_CORE, "docProps/core.xml", false)?;
    package.add_relationship(None, "rIdApp", REL_APP, "docProps/app.xml", false)?;
    package.add_relationship(
        Some("xl/workbook.xml"),
        "rIdStyles",
        REL_STYLES,
        "styles.xml",
        false,
    )?;
    if let Some(theme) = conditional_theme {
        package.add_xml("xl/theme/theme1.xml", "application/vnd.openxmlformats-officedocument.theme+xml", theme)?;
        package.add_relationship(Some("xl/workbook.xml"), "rIdTheme", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme", "theme/theme1.xml", false)?;
    }
    Ok(warnings)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn converts_a1_and_ranges() {
        assert_eq!(cell_ref(0, 0), "A1");
        assert_eq!(cell_ref(9, 26), "AA10");
        let range = RangeInput {
            start_row: Some(1),
            start_col: Some(2),
            end_row: Some(4),
            end_col: Some(5),
            first_row: None,
            first_col: None,
            last_row: None,
            last_col: None,
        };
        assert_eq!(range_ref(&range).as_deref(), Some("C2:F5"));
    }

    #[test]
    fn writes_diagonal_flags_and_ooxml_line_styles() {
        let value = serde_json::json!({ "diagonalDown": { "style": "medium-dash-dot", "color": "#123456" }, "diagonalUp": { "style": "medium-dash-dot", "color": "#123456" }, "bottom": { "style": "dash-dot-dot", "color": "#abcdef" } });
        let xml = border_xml(&border_key(Some(&value)));
        assert!(xml.contains("diagonalUp=\"1\" diagonalDown=\"1\""));
        assert!(xml.contains("<diagonal style=\"mediumDashDot\"><color rgb=\"FF123456\"/></diagonal>"));
        assert!(xml.contains("<bottom style=\"dashDotDot\"><color rgb=\"FFABCDEF\"/></bottom>"));
    }

    #[test]
    fn writes_slanted_dash_dot_and_cleared_side_segments() {
        let value = serde_json::json!({
            "diagonalDown": { "style": "slant-dash-dot", "color": "#123456" },
            "left": null,
            "right": { "style": "thin", "color": "#abcdef" }
        });
        let xml = border_xml(&border_key(Some(&value)));
        assert!(xml.contains("<diagonal style=\"slantDashDot\"><color rgb=\"FF123456\"/></diagonal>"));
        assert!(xml.contains("<left/>"));
        assert!(xml.contains("<right style=\"thin\"><color rgb=\"FFABCDEF\"/></right>"));
        assert!(!xml.contains("style=\"none\""));
    }

    #[test]
    fn styles_are_deduplicated() {
        let mut registry = StyleRegistry::new();
        let style = StyleInput {
            font: None,
            fill: Some("#ffffff".to_owned()),
            horizontal: None,
            vertical: None,
            wrap: false,
            number_format: None,
            border: None,
            rotation: None,
            rotation_vertical: false,
            indent: None,
            shrink_to_fit: false,
            metadata: None,
        };
        assert_eq!(
            registry.register(Some(&style)),
            registry.register(Some(&style))
        );
    }

    #[test]
    fn reset_colors_are_distinct_from_an_explicit_white_fill_and_font() {
        let reset: StyleInput = serde_json::from_str(r##"{"fill":null,"font":{"color":null}}"##).unwrap();
        let white: StyleInput = serde_json::from_str(r##"{"fill":"#fff","font":{"color":"#ffffff"}}"##).unwrap();
        let mut registry = StyleRegistry::new();
        let reset_id = registry.register(Some(&reset));
        let white_id = registry.register(Some(&white));
        assert_ne!(reset_id, white_id);
        assert_eq!(style_key(Some(&reset)).fill, "none");
        let xml = style_xml(&registry);
        assert!(xml.contains("<fgColor rgb=\"FFFFFFFF\"/>"));
        assert!(xml.contains("<color rgb=\"FFFFFFFF\"/>"));
        assert_eq!(xml.matches("<color rgb=").count(), 1);
    }

    #[test]
    fn writes_conditional_format_differential_styles_and_rules() {
        let data = serde_json::json!({
            "kind": "workbook",
            "sheets": [{
                "id": "sheet-1", "name": "Scores", "cells": {},
                "conditionalFormats": [
                    { "id": "high", "type": "cellIs", "ranges": [
                        { "startRow": 0, "startCol": 0, "endRow": 4, "endCol": 0 },
                        { "startRow": 0, "startCol": 2, "endRow": 4, "endCol": 2 }
                    ], "operator": "greaterThanOrEqual", "formula1": "90", "priority": 2,
                      "stopIfTrue": true, "style": { "fill": "#00aa55", "fontColor": "#ffffff", "bold": true } },
                    { "id": "low", "type": "cellIs", "ranges": [
                        { "startRow": 0, "startCol": 0, "endRow": 4, "endCol": 0 }
                    ], "operator": "lessThan", "formula1": "60", "priority": 1,
                      "style": { "fill": "#ee3344" } },
                    { "id": "text", "type": "containsText", "ranges": [
                        { "startRow": 2, "startCol": 1, "endRow": 4, "endCol": 1 }
                    ], "operator": "containsText", "formula1": "A\"B", "value": "A\"B", "priority": 3,
                      "style": { "italic": true } }
                ]
            }]
        }).to_string();
        let mut package = Package::new();
        let warnings = write(&data, &mut package).expect("conditional formats export");
        assert!(warnings.is_empty(), "{warnings:?}");
        let bytes = package.finish().unwrap();
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
        let mut worksheet = String::new();
        let mut styles = String::new();
        std::io::Read::read_to_string(&mut archive.by_name("xl/worksheets/sheet1.xml").unwrap(), &mut worksheet).unwrap();
        std::io::Read::read_to_string(&mut archive.by_name("xl/styles.xml").unwrap(), &mut styles).unwrap();

        let root = office_io_common::conditional_format_xml::parse(&worksheet).unwrap();
        let groups = root.children.iter().filter(|node| node.is_main("conditionalFormatting")).collect::<Vec<_>>();
        assert_eq!(groups.len(), 3);
        for (index, (reference, operator, operand)) in [("A1:A5", "lessThan", "60"), ("A1:A5 C1:C5", "greaterThanOrEqual", "90")].into_iter().enumerate() {
            let rule = groups[index].child("cfRule").unwrap();
            assert_eq!(groups[index].attrs.get("sqref").map(String::as_str), Some(reference));
            assert_eq!(rule.attrs.get("priority"), Some(&(index + 1).to_string())); assert_eq!(rule.attrs.get("dxfId"), Some(&index.to_string()));
            assert_eq!(rule.attrs.get("operator").map(String::as_str), Some(operator)); assert_eq!(rule.child("formula").unwrap().text, operand);
        }
        assert_eq!(groups[1].child("cfRule").unwrap().attrs.get("stopIfTrue").map(String::as_str), Some("1"));
        let text = groups[2].child("cfRule").unwrap();
        assert_eq!(groups[2].attrs.get("sqref").map(String::as_str), Some("B3:B5"));
        assert_eq!(text.attrs.get("type").map(String::as_str), Some("containsText")); assert_eq!(text.attrs.get("text").map(String::as_str), Some("A\"B"));
        assert_eq!(text.child("formula").unwrap().text, "NOT(ISERROR(SEARCH(\"A\"\"B\",B3)))");
        let root = office_io_common::conditional_format_xml::parse(&styles).unwrap(); let dxfs = root.child("dxfs").unwrap();
        assert_eq!(dxfs.children.len(), 3);
        assert_eq!(dxfs.children[0].child("fill").unwrap().child("patternFill").unwrap().child("fgColor").unwrap().attrs.get("rgb").map(String::as_str), Some("FFEE3344"));
        let font = dxfs.children[1].child("font").unwrap(); assert!(font.child("b").is_some());
        assert_eq!(font.child("color").unwrap().attrs.get("rgb").map(String::as_str), Some("FFFFFFFF"));
    }

    #[test]
    fn vertical_text_uses_the_ooxml_255_rotation_sentinel() {
        let data = r#"{"kind":"workbook","sheets":[{"id":"sheet-1","name":"Vertical","cells":{"A1":{"value":"Vertical","style":{"rotationVertical":true}}}}]}"#;
        let mut package = Package::new();
        write(data, &mut package).unwrap();
        let bytes = package.finish().unwrap();
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
        let mut styles = String::new();
        std::io::Read::read_to_string(&mut archive.by_name("xl/styles.xml").unwrap(), &mut styles)
            .unwrap();

        assert!(styles.contains("textRotation=\"255\""), "{styles}");
    }

    #[test]
    fn worksheet_names_are_safe_and_unique() {
        let names = safe_sheet_names(["Sales", "Sales", "sales", "Q1/2026", ""]);
        assert_eq!(
            names,
            ["Sales", "Sales (2)", "sales (3)", "Q1_2026", "Sheet5"]
        );
    }

    #[test]
    fn local_defined_name_tracks_sanitized_sheet_name() {
        let workbook: WorkbookInput = serde_json::from_str(
            r#"{"sheets":[{"id":"sheet-1","name":"Quarter/1"}],"metadata":{"localDefinedNames":[{"sheet":"Quarter/1","name":"LocalTotal","refersTo":"A1"}]}}"#,
        )
        .unwrap();
        let xml = workbook_xml(&workbook);

        assert!(xml.contains("<sheet name=\"Quarter_1\""));
        assert!(
            xml.contains("<definedName name=\"LocalTotal\" localSheetId=\"0\">A1</definedName>"),
            "{xml}"
        );
    }

    #[test]
    fn writes_a_minimal_xlsx_package_with_consistent_parts() {
        let data = r#"{"kind":"workbook","sheets":[{"id":"sheet-1","name":"Sales","cells":{}}]}"#;
        let mut package = Package::new();
        write(data, &mut package).unwrap();
        let bytes = package.finish().unwrap();
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
        let mut workbook = String::new();
        std::io::Read::read_to_string(
            &mut archive.by_name("xl/workbook.xml").unwrap(),
            &mut workbook,
        )
        .unwrap();
        let mut app_properties = String::new();
        std::io::Read::read_to_string(
            &mut archive.by_name("docProps/app.xml").unwrap(),
            &mut app_properties,
        )
        .unwrap();
        let mut content_types = String::new();
        std::io::Read::read_to_string(
            &mut archive.by_name("[Content_Types].xml").unwrap(),
            &mut content_types,
        )
        .unwrap();
        let mut styles = String::new();
        std::io::Read::read_to_string(&mut archive.by_name("xl/styles.xml").unwrap(), &mut styles)
            .unwrap();

        assert!(workbook.contains("<sheet name=\"Sales\""));
        assert!(app_properties.contains("<vt:lpstr>Sales</vt:lpstr>"));
        assert!(content_types.contains("Extension=\"rels\""));
        assert!(content_types.contains("PartName=\"/xl/workbook.xml\""));
        assert!(content_types.contains("PartName=\"/xl/styles.xml\""));
        assert!(styles.contains("<cellXfs count=\"1\">"));
    }

    #[test]
    fn writes_a1_keyed_external_and_internal_hyperlinks() {
        let data = r##"{"kind":"workbook","sheets":[{"id":"sheet-1","name":"Links","cells":{"A1":{"value":"external","hyperlink":"https://example.com/?a=1&b=2"},"A2":{"value":"internal","hyperlink":"#Links!A1"},"B2":{"value":"external after internal","hyperlink":"https://example.com/second"}}}]}"##;
        let mut package = Package::new();
        write(data, &mut package).unwrap();
        let bytes = package.finish().unwrap();
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
        let mut worksheet = String::new();
        std::io::Read::read_to_string(
            &mut archive.by_name("xl/worksheets/sheet1.xml").unwrap(),
            &mut worksheet,
        )
        .unwrap();
        let mut relationships = String::new();
        std::io::Read::read_to_string(
            &mut archive
                .by_name("xl/worksheets/_rels/sheet1.xml.rels")
                .unwrap(),
            &mut relationships,
        )
        .unwrap();

        assert!(worksheet.contains("<hyperlink ref=\"A1\" r:id=\"rIdHyperlink1\"/>"));
        assert!(worksheet.contains("<hyperlink ref=\"A2\" location=\"Links!A1\"/>"));
        assert!(worksheet.contains("<hyperlink ref=\"B2\" r:id=\"rIdHyperlink3\"/>"));
        assert!(relationships.contains("Id=\"rIdHyperlink1\""));
        assert!(relationships.contains("Target=\"https://example.com/?a=1&amp;b=2\""));
        assert!(relationships.contains("Id=\"rIdHyperlink3\""));
        assert!(relationships.contains("Target=\"https://example.com/second\""));
        assert!(relationships.contains("TargetMode=\"External\""));
        assert!(!relationships.contains("Links!A1"));
        assert_eq!(relationships.matches("<Relationship ").count(), 2);
    }

    #[test]
    fn writes_internal_link_payload_as_an_xlsx_location() {
        let data = r##"{"kind":"workbook","sheets":[{"id":"source","name":"Source","cells":{"A1":{"value":"internal","hyperlink":"#gid=target&range=B2%3AC4"}}},{"id":"target","name":"Quarter's Report","cells":{}}]}"##;
        let mut package = Package::new();
        write(data, &mut package).unwrap();
        let bytes = package.finish().unwrap();
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
        let mut worksheet = String::new();
        std::io::Read::read_to_string(
            &mut archive.by_name("xl/worksheets/sheet1.xml").unwrap(),
            &mut worksheet,
        )
        .unwrap();
        assert!(worksheet.contains("<hyperlink ref=\"A1\" location=\"&apos;Quarter&apos;&apos;s Report&apos;!B2:C4\"/>"));
    }

    #[test]
    fn writes_hidden_rows_columns_and_custom_sizes() {
        let data = r#"{"kind":"workbook","sheets":[{"id":"sheet-1","name":"Axes","rowCount":5,"columnCount":7,"hiddenRows":[2],"hiddenColumns":[4],"rowHeights":{"3":32},"columnWidths":{"5":120}}]}"#;
        let mut package = Package::new();
        write(data, &mut package).unwrap();
        let bytes = package.finish().unwrap();
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
        let mut worksheet = String::new();
        std::io::Read::read_to_string(
            &mut archive.by_name("xl/worksheets/sheet1.xml").unwrap(),
            &mut worksheet,
        )
        .unwrap();

        assert!(
            worksheet.contains(
                "<col min=\"5\" max=\"5\" width=\"16\" customWidth=\"1\" hidden=\"1\" />"
            )
        );
        assert!(
            worksheet.contains("<col min=\"6\" max=\"6\" width=\"17.1445\" customWidth=\"1\" />")
        );
        assert!(worksheet.contains("<row r=\"3\" hidden=\"1\"></row>"));
        assert!(worksheet.contains("<row r=\"4\" ht=\"24\" customHeight=\"1\"></row>"));
    }

    #[test]
    fn writes_x14_sparkline_groups_with_sources_and_supported_display_options() {
        let (worksheet, warnings) = exported_filter_sheet(serde_json::json!({
            "kind": "workbook",
            "sheets": [{
                "id": "sparkline-sheet", "name": "Data", "cells": {
                    "F1": { "row": 0, "col": 5, "value": null, "sparkline": {
                        "type": "column", "sourceRange": "B1:E1", "color": "#167a55",
                        "negativeColor": "#c43b3b", "highPointColor": "#f0a000", "lowPointColor": "#2b579a",
                        "emptyCells": "zero", "axis": "custom", "min": -5, "max": 20
                    } }
                }
            }]
        }));
        assert!(warnings.is_empty(), "{warnings:?}");
        assert!(worksheet.contains("<extLst><ext uri=\"{05C60535-1F16-4fd2-B633-F4F36F0B64E0}\"><x14:sparklineGroups"));
        assert!(worksheet.contains("type=\"column\" displayEmptyCellsAs=\"zero\" minAxisType=\"custom\" maxAxisType=\"custom\" manualMin=\"-5\" manualMax=\"20\""));
        assert!(worksheet.contains("<x14:colorSeries rgb=\"FF167A55\"/><x14:colorNegative rgb=\"FFC43B3B\"/><x14:colorHigh rgb=\"FFF0A000\"/><x14:colorLow rgb=\"FF2B579A\"/>"));
        assert!(worksheet.contains("<xm:f>&apos;Data&apos;!B1:E1</xm:f><xm:sqref>F1</xm:sqref>"), "{worksheet}");
        assert!(worksheet.contains("</x14:sparklineGroups></ext><ext uri=\"{F41EED10-CA42-42C0-A66C-9152ABCB9AE1}\""));
        assert!(worksheet.ends_with("</ov:axisDimensions></ext></extLst></worksheet>"));
    }

    #[test]
    fn writes_structured_table_parts_and_sheet_relationships() {
        let read_part = |archive: &mut zip::ZipArchive<std::io::Cursor<Vec<u8>>>, path: &str| {
            let mut content = String::new();
            std::io::Read::read_to_string(&mut archive.by_name(path).unwrap(), &mut content).unwrap();
            content
        };
        let mut package = Package::new();
        let model = serde_json::json!({
            "kind":"workbook","name":"Tables","sheets":[{"id":"sheet-1","name":"Data","cells":{}}],
            "tables":[{"id":"table-stable-id","name":"Orders","sheetId":"sheet-1","range":{"startRow":0,"startCol":0,"endRow":4,"endCol":1},"columns":[{"id":"c1","name":"Product"},{"id":"c2","name":"Amount"}],"showHeaderRow":true,"showTotalsRow":true,"styleName":"TableStyleMedium4"}]
        });
        let warnings = write(&model.to_string(), &mut package).expect("table writes");
        assert!(warnings.is_empty(), "{warnings:?}");
        let bytes = package.finish().unwrap();
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
        let worksheet = read_part(&mut archive, "xl/worksheets/sheet1.xml");
        let table = read_part(&mut archive, "xl/tables/table1.xml");
        let rels = read_part(&mut archive, "xl/worksheets/_rels/sheet1.xml.rels");
        assert!(worksheet.contains("<tableParts count=\"1\"><tablePart r:id=\"rIdTable1\"/></tableParts>"));
        assert!(table.contains("displayName=\"Orders\" ref=\"A1:B5\" headerRowCount=\"1\" totalsRowShown=\"1\""));
        assert!(table.contains("<tableColumn id=\"1\" name=\"Product\"/>"));
        assert!(table.contains("<tableStyleInfo name=\"TableStyleMedium4\""));
        assert!(rels.contains("relationships/table"));
        assert!(rels.contains("../tables/table1.xml"));
    }

    fn exported_filter_sheet(model: serde_json::Value) -> (String, Vec<ExportWarning>) {
        let mut package = Package::new();
        let warnings = write(&model.to_string(), &mut package).unwrap();
        let mut archive =
            zip::ZipArchive::new(std::io::Cursor::new(package.finish().unwrap())).unwrap();
        let mut worksheet = String::new();
        std::io::Read::read_to_string(
            &mut archive.by_name("xl/worksheets/sheet1.xml").unwrap(),
            &mut worksheet,
        )
        .unwrap();
        (worksheet, warnings)
    }

    #[test]
    fn exports_filter_columns_relative_to_range_and_unions_hidden_row_reasons() {
        let (worksheet, warnings) = exported_filter_sheet(serde_json::json!({
            "kind": "workbook", "sheets": [{"id": "data", "name": "Filters", "hiddenRows": [1],
                "autoFilter": { "range": {"startRow": 2, "startCol": 1, "endRow": 6, "endCol": 3},
                    "columns": {"1": {"kind": "values", "selectedItems": ["East", null, ""]},
                        "3": {"kind": "condition", "operator": "between", "value": 100, "secondValue": 200}},
                    "filteredRows": [3, 6]}
            }]
        }));
        assert!(warnings.is_empty(), "{warnings:?}");
        assert!(worksheet.contains("<sheetPr filterMode=\"1\"/>"));
        assert!(worksheet.contains("<dimension ref=\"A1:D7\"/>"));
        assert!(worksheet.contains("<autoFilter ref=\"B3:D7\"><filterColumn colId=\"0\"><filters blank=\"1\"><filter val=\"East\"/></filters></filterColumn>"));
        assert!(worksheet.contains("<filterColumn colId=\"2\"><customFilters and=\"1\"><customFilter operator=\"greaterThanOrEqual\" val=\"100\"/><customFilter operator=\"lessThanOrEqual\" val=\"200\"/></customFilters></filterColumn>"));
        for row in [2, 4, 7] {
            assert!(
                worksheet.contains(&format!("<row r=\"{row}\" hidden=\"1\">")),
                "{worksheet}"
            );
        }
        assert!(!worksheet.contains("<row r=\"3\" hidden=\"1\">"));
        assert!(worksheet.find("</sheetData>").unwrap() < worksheet.find("<autoFilter ").unwrap());
    }

    #[test]
    fn serializes_comparisons_and_escaped_text_filter_patterns() {
        let operators = [
            (
                "equals",
                serde_json::json!("A*?~&"),
                None,
                "<customFilter operator=\"equal\" val=\"A~*~?~~&amp;\"/>",
            ),
            (
                "doesNotEqual",
                serde_json::json!(4),
                None,
                "<customFilter operator=\"notEqual\" val=\"4\"/>",
            ),
            (
                "greaterThan",
                serde_json::json!(4),
                None,
                "<customFilter operator=\"greaterThan\" val=\"4\"/>",
            ),
            (
                "greaterThanOrEqual",
                serde_json::json!(4),
                None,
                "<customFilter operator=\"greaterThanOrEqual\" val=\"4\"/>",
            ),
            (
                "lessThan",
                serde_json::json!(4),
                None,
                "<customFilter operator=\"lessThan\" val=\"4\"/>",
            ),
            (
                "lessThanOrEqual",
                serde_json::json!(4),
                None,
                "<customFilter operator=\"lessThanOrEqual\" val=\"4\"/>",
            ),
            (
                "contains",
                serde_json::json!("*?~"),
                None,
                "<customFilter operator=\"equal\" val=\"*~*~?~~*\"/>",
            ),
            (
                "doesNotContain",
                serde_json::json!("*?~"),
                None,
                "<customFilter operator=\"notEqual\" val=\"*~*~?~~*\"/>",
            ),
            (
                "beginsWith",
                serde_json::json!("*?~"),
                None,
                "<customFilter operator=\"equal\" val=\"~*~?~~*\"/>",
            ),
            (
                "endsWith",
                serde_json::json!("*?~"),
                None,
                "<customFilter operator=\"equal\" val=\"*~*~?~~\"/>",
            ),
            (
                "between",
                serde_json::json!(4),
                Some(serde_json::json!(8)),
                "<customFilters and=\"1\"><customFilter operator=\"greaterThanOrEqual\" val=\"4\"/><customFilter operator=\"lessThanOrEqual\" val=\"8\"/></customFilters>",
            ),
            (
                "notBetween",
                serde_json::json!(4),
                Some(serde_json::json!(8)),
                "<customFilters and=\"0\"><customFilter operator=\"lessThan\" val=\"4\"/><customFilter operator=\"greaterThan\" val=\"8\"/></customFilters>",
            ),
            (
                "isBlank",
                serde_json::Value::Null,
                None,
                "<filters blank=\"1\"/>",
            ),
            (
                "isNotBlank",
                serde_json::Value::Null,
                None,
                "<customFilter operator=\"notEqual\" val=\" \"/>",
            ),
        ];
        for (operator, value, second, expected) in operators {
            let mut criterion =
                serde_json::json!({"kind": "condition", "operator": operator, "value": value});
            if let Some(second) = second {
                criterion["secondValue"] = second;
            }
            let (worksheet, warnings) = exported_filter_sheet(
                serde_json::json!({"sheets": [{"name": "Filters", "autoFilter": {
                    "range": {"startRow": 0, "startCol": 0, "endRow": 3, "endCol": 0}, "columns": {"0": criterion}, "filteredRows": []
                }}]}),
            );
            assert!(
                worksheet.contains(expected),
                "operator={operator}: {worksheet}"
            );
            assert!(warnings.is_empty(), "operator={operator}: {warnings:?}");
        }
    }

    #[test]
    fn empty_value_selection_is_an_impossible_predicate_and_error_filter_is_a_diagnosed_snapshot() {
        let base = serde_json::json!({"sheets": [{"name": "Errors", "cells": {
            "A2": {"value": "#DIV/0!", "dataType": "error"}, "A3": {"value": "#N/A", "dataType": "e"}, "A4": {"value": "normal"}
        }, "autoFilter": {"range": {"startRow": 0, "startCol": 0, "endRow": 3, "endCol": 0}, "columns": {"0": {"kind": "values", "selectedItems": []}}, "filteredRows": [1, 2, 3]}}]});
        let (empty, warnings) = exported_filter_sheet(base.clone());
        assert!(warnings.is_empty());
        assert!(empty.contains("<customFilters and=\"1\"><customFilter operator=\"equal\" val=\" \"/><customFilter operator=\"notEqual\" val=\" \"/></customFilters>"));
        let mut errors = base;
        errors["sheets"][0]["autoFilter"]["columns"]["0"] =
            serde_json::json!({"kind": "condition", "operator": "isError"});
        let (worksheet, warnings) = exported_filter_sheet(errors);
        assert!(worksheet.contains("<filter val=\"#DIV/0!\"/><filter val=\"#N/A\"/>"));
        assert!(
            warnings
                .iter()
                .any(|warning| warning.code == "AUTOFILTER_ERROR_SNAPSHOT")
        );
    }

    #[test]
    fn diagnoses_filter_semantics_that_office_cannot_preserve_exactly() {
        let (worksheet, warnings) = exported_filter_sheet(
            serde_json::json!({"sheets": [{"name": "Typed", "cells": {"A2": {"value": 1}}, "autoFilter": {
                "range": {"startRow": 0, "startCol": 0, "endRow": 3, "endCol": 2}, "columns": {
                    "0": {"kind": "condition", "operator": "contains", "value": "East", "caseSensitive": true},
                    "1": {"kind": "values", "selectedItems": [null]},
                    "2": {"kind": "values", "selectedItems": [1, "1", true, "1"]}
                }, "filteredRows": []
            }}]}),
        );
        assert!(worksheet.contains("<autoFilter "));
        for expected in [
            "AUTOFILTER_CASE_SENSITIVE",
            "AUTOFILTER_BLANK_TYPES",
            "AUTOFILTER_VALUE_TYPES",
            "AUTOFILTER_MIXED_TYPES",
        ] {
            assert!(
                warnings.iter().any(|warning| warning.code == expected),
                "missing {expected}: {warnings:?}"
            );
        }
    }

    #[test]
    fn rejects_invalid_filter_ranges_columns_rows_and_operands() {
        let cases = [
            serde_json::json!({"range": {"startRow": 2, "startCol": 0, "endRow": 2, "endCol": 0}}),
            serde_json::json!({"range": {"startRow": 0, "startCol": 2, "endRow": 3, "endCol": 1}}),
            serde_json::json!({"range": {"startRow": 0, "startCol": 0, "endRow": 1048576, "endCol": 0}}),
            serde_json::json!({"range": {"startRow": 0, "startCol": 0, "endRow": 3, "endCol": 0}, "columns": {"1": {"kind": "values", "selectedItems": []}}}),
            serde_json::json!({"range": {"startRow": 0, "startCol": 0, "endRow": 3, "endCol": 0}, "filteredRows": [0]}),
            serde_json::json!({"range": {"startRow": 0, "startCol": 0, "endRow": 3, "endCol": 0}, "filteredRows": [4]}),
            serde_json::json!({"range": {"startRow": 0, "startCol": 0, "endRow": 3, "endCol": 0}, "columns": {"0": {"kind": "condition", "operator": "between", "value": 2}}}),
            serde_json::json!({"range": {"startRow": 0, "startCol": 0, "endRow": 3, "endCol": 0}, "columns": {"0": {"kind": "condition", "operator": "unknown"}}}),
            serde_json::json!({"range": {"startRow": 0, "startCol": 0, "endRow": 3, "endCol": 0}, "columns": {"0": {"kind": "values", "selectedItems": [{}]}}}),
        ];
        for filter in cases {
            let data = serde_json::json!({"sheets": [{"name": "Invalid", "autoFilter": filter}]})
                .to_string();
            let error = write(&data, &mut Package::new()).unwrap_err();
            assert_eq!(error.code, "INVALID_MODEL", "{data}");
        }
    }

    #[test]
    fn rejects_axis_and_freeze_coordinates_beyond_xlsx_limits() {
        let invalid_sheets = [
            r#"{"hiddenRows":[1048576]}"#,
            r#"{"hiddenColumns":[16384]}"#,
            r#"{"rowHeights":{"1048576":24}}"#,
            r#"{"columnWidths":{"16384":120}}"#,
            r#"{"frozenRows":1048576}"#,
            r#"{"frozenCols":16384}"#,
        ];

        for sheet in invalid_sheets {
            let data = format!(
                r#"{{"kind":"workbook","sheets":[{{"name":"Invalid",{}}}]}}"#,
                sheet.trim_start_matches('{').trim_end_matches('}')
            );
            let mut package = Package::new();
            assert!(write(&data, &mut package).is_err(), "accepted {sheet}");
        }
    }

    #[test]
    fn exports_pivot_cache_table_and_linked_pivot_chart_parts() {
        let data = r##"{"kind":"workbook","name":"Pivot","sheets":[{"id":"data","name":"Data","cells":{"A1":{"value":"Region"},"B1":{"value":"Sales"},"A2":{"value":"East"},"B2":{"value":10},"A3":{"value":"West"},"B3":{"value":20}}},{"id":"pivot","name":"Pivot View","cells":{"A1":{"value":"Region"},"B1":{"value":"Sales"},"A2":{"value":"East"},"B2":{"value":10},"A3":{"value":"West"},"B3":{"value":20}}}],"pivotTables":[{"id":"pivot-1","source":{"subUnitId":"data","sheetName":"Data","range":{"startRow":0,"startCol":0,"endRow":2,"endCol":1}},"anchor":{"subUnitId":"pivot","row":0,"col":0},"outputRange":{"startRow":0,"startCol":0,"endRow":2,"endCol":1},"fields":[{"sourceColumn":0,"name":"Region","area":"row","position":0},{"sourceColumn":1,"name":"Sales","area":"value","position":0,"aggregation":"sum"}],"options":{}}],"charts":[{"id":"pivot-chart","sheetId":"pivot","type":"bar","source":{"range":{"startRow":0,"startCol":0,"endRow":2,"endCol":1},"orientation":"columns"},"position":{"x":20,"y":20},"size":{"width":400,"height":240},"categoryField":0,"valueFields":[1],"pivotTableId":"pivot-1"}]}"##;
        let mut package = Package::new();
        let warnings = write(data, &mut package).unwrap();
        assert!(warnings.is_empty(), "{warnings:?}");
        let bytes = package.finish().unwrap();
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
        let read_part = |archive: &mut zip::ZipArchive<std::io::Cursor<Vec<u8>>>, path: &str| {
            let mut text = String::new();
            std::io::Read::read_to_string(&mut archive.by_name(path).unwrap(), &mut text).unwrap();
            text
        };
        let workbook = read_part(&mut archive, "xl/workbook.xml");
        let workbook_rels = read_part(&mut archive, "xl/_rels/workbook.xml.rels");
        let cache = read_part(&mut archive, "xl/pivotCache/pivotCacheDefinition1.xml");
        let cache_rels = read_part(&mut archive, "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels");
        let records = read_part(&mut archive, "xl/pivotCache/pivotCacheRecords1.xml");
        let pivot_table = read_part(&mut archive, "xl/pivotTables/pivotTable1.xml");
        let table_rels = read_part(&mut archive, "xl/pivotTables/_rels/pivotTable1.xml.rels");
        let pivot_sheet = read_part(&mut archive, "xl/worksheets/sheet2.xml");
        let pivot_chart = read_part(&mut archive, "xl/charts/chart1.xml");
        assert!(workbook.contains("<pivotCaches count=\"1\">"));
        assert!(workbook_rels.contains("pivotCache/pivotCacheDefinition1.xml"));
        assert!(cache.contains("<worksheetSource ref=\"A1:B3\" sheet=\"Data\"/>"));
        assert!(cache_rels.contains("pivotCacheRecords1.xml"));
        assert!(records.contains("<r><x v=\"0\"/><x v=\"0\"/></r>"));
        assert!(pivot_table.contains("<rowFields count=\"1\"><field x=\"0\"/>"));
        assert!(table_rels.contains("pivotCacheDefinition1.xml"));
        assert!(pivot_sheet.contains("<pivotTableParts count=\"1\"><pivotTablePart r:id=\"rIdPivotTable1\"/>"));
        assert!(pivot_chart.contains("<c:pivotSource><c:name>[Pivot.xlsx]Pivot View!PivotTable1</c:name>"));
    }

    #[test]
    fn writes_chart_drawing_and_series_parts_with_source_ranges() {
        let data = r##"{"kind":"workbook","name":"Chart book","sheets":[{"id":"sheet-1","name":"Q1 Sales","cells":{"A1":{"value":"Region"},"B1":{"value":"Revenue"},"A2":{"value":"North"},"B2":{"value":10},"A3":{"value":"South"},"B3":{"value":20}}}],"charts":[{"id":"chart-1","sheetId":"sheet-1","type":"bar","source":{"range":{"startRow":0,"startCol":0,"endRow":2,"endCol":1},"orientation":"columns"},"position":{"x":16,"y":32},"size":{"width":640,"height":360},"title":"Revenue by region","legend":{"visible":true,"position":"bottom"},"categoryField":0,"valueFields":[1],"xAxis":{"title":"Region","scaleType":"linear"},"yAxis":{"title":"Revenue","scaleType":"linear"}}]}"##;
        let mut package = Package::new();
        let warnings = write(data, &mut package).unwrap();
        assert!(warnings.is_empty(), "{warnings:?}");
        let bytes = package.finish().unwrap();
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
        let read_part = |archive: &mut zip::ZipArchive<std::io::Cursor<Vec<u8>>>, path: &str| {
            let mut text = String::new();
            std::io::Read::read_to_string(&mut archive.by_name(path).unwrap(), &mut text).unwrap();
            text
        };
        let worksheet = read_part(&mut archive, "xl/worksheets/sheet1.xml");
        let sheet_relationships = read_part(&mut archive, "xl/worksheets/_rels/sheet1.xml.rels");
        let drawing = read_part(&mut archive, "xl/drawings/drawing1.xml");
        let drawing_relationships = read_part(&mut archive, "xl/drawings/_rels/drawing1.xml.rels");
        let chart = read_part(&mut archive, "xl/charts/chart1.xml");
        let content_types = read_part(&mut archive, "[Content_Types].xml");

        assert!(worksheet.contains("<drawing r:id=\"rIdDrawing\"/>"));
        assert!(
            sheet_relationships
                .contains("relationships/drawing\" Target=\"../drawings/drawing1.xml\"")
        );
        assert!(drawing.contains("<xdr:absoluteAnchor><xdr:pos x=\"152400\" y=\"304800\"/>"));
        assert!(drawing.contains("r:id=\"rIdChart1\""));
        assert!(
            drawing_relationships.contains("relationships/chart\" Target=\"../charts/chart1.xml\"")
        );
        assert!(chart.contains("<c:barChart>"));
        assert!(chart.contains("<a:t>Revenue by region</a:t>"));
        assert!(chart.contains("&apos;Q1 Sales&apos;!$B$1:$B$1"));
        assert!(chart.contains("&apos;Q1 Sales&apos;!$A$2:$A$3"));
        assert!(chart.contains("&apos;Q1 Sales&apos;!$B$2:$B$3"));
        assert!(chart.contains("<c:v>North</c:v>"));
        assert!(chart.contains("<c:v>20</c:v>"));
        assert!(content_types.contains("/xl/drawings/drawing1.xml"));
        assert!(content_types.contains("/xl/charts/chart1.xml"));
    }

    #[test]
    fn rejects_chart_series_outside_source_and_unsupported_type() {
        let base = r#"{"kind":"workbook","sheets":[{"id":"sheet-1","name":"Data","cells":{}}],"charts":[{"id":"chart-1","sheetId":"sheet-1","type":"bar","source":{"range":{"startRow":0,"startCol":0,"endRow":1,"endCol":1}},"position":{"x":0,"y":0},"size":{"width":300,"height":200},"categoryField":0,"valueFields":[4]}]}"#;
        let mut package = Package::new();
        assert!(write(base, &mut package).is_err());

        let unsupported = base.replace("\"bar\"", "\"surface\"").replace("[4]", "[1]");
        let mut package = Package::new();
        let error = write(&unsupported, &mut package).unwrap_err();
        assert_eq!(error.code, "UNSUPPORTED_EXPORT");
    }

    #[test]
    fn writes_line_chart_with_row_oriented_source_and_log_axis() {
        let workbook: WorkbookInput = serde_json::from_str(
            r##"{"sheets":[{"id":"sheet-1","name":"Data","cells":{"A1":{"value":"Metric"},"B1":{"value":"Jan"},"C1":{"value":"Feb"},"A2":{"value":"Net"},"B2":{"value":5},"C2":{"value":8}}}],"charts":[{"id":"line-1","sheetId":"sheet-1","type":"line","source":{"range":{"startRow":0,"startCol":0,"endRow":1,"endCol":2},"orientation":"rows"},"position":{"x":0,"y":0},"size":{"width":320,"height":200},"legend":{"visible":false,"position":"right"},"categoryField":0,"valueFields":[1],"yAxis":{"scaleType":"log","logBase":2}}]}"##,
        )
        .unwrap();
        let chart = chart_xml(&workbook, &workbook.charts.as_ref().unwrap()[0], 0, None).unwrap();

        assert!(chart.contains("<c:lineChart>"));
        assert!(chart.contains("<c:grouping val=\"standard\"/>"));
        assert!(chart.contains("<c:f>&apos;Data&apos;!$B$1:$C$1</c:f>"));
        assert!(chart.contains("<c:f>&apos;Data&apos;!$B$2:$C$2</c:f>"));
        assert!(chart.contains("<c:logBase val=\"2\"/>"));
        assert!(!chart.contains("<c:legend>"));
    }

    #[test]
    fn writes_scatter_chart_with_numeric_x_and_y_series() {
        let workbook: WorkbookInput = serde_json::from_str(
            r##"{"sheets":[{"id":"sheet-1","name":"Data","cells":{"A1":{"value":"X"},"B1":{"value":"Y"},"A2":{"value":1},"B2":{"value":10},"A3":{"value":2},"B3":{"value":20}}}],"charts":[{"id":"scatter-1","sheetId":"sheet-1","type":"scatter","source":{"range":{"startRow":0,"startCol":0,"endRow":2,"endCol":1},"orientation":"columns"},"position":{"x":0,"y":0},"size":{"width":320,"height":200},"categoryField":0,"valueFields":[1]}]}"##,
        )
        .unwrap();
        let chart = chart_xml(&workbook, &workbook.charts.as_ref().unwrap()[0], 0, None).unwrap();

        assert!(chart.contains("<c:scatterChart><c:scatterStyle val=\"marker\"/>"));
        assert!(chart.contains("<c:xVal><c:numRef><c:f>&apos;Data&apos;!$A$2:$A$3</c:f>"));
        assert!(chart.contains("<c:yVal><c:numRef><c:f>&apos;Data&apos;!$B$2:$B$3</c:f>"));
        assert!(chart.contains("<c:valAx><c:axId val=\"1000003\"/>"));
        assert!(!chart.contains("<c:catAx>"));
    }

    #[test]
    fn writes_radar_chart_with_category_and_value_axes() {
        let mut workbook: WorkbookInput = serde_json::from_value(serde_json::json!({
            "sheets": [{"id": "data", "name": "Survey", "cells": {
                "B1": {"value": "Quality"}, "C1": {"value": "Speed"}, "D1": {"value": "Cost"},
                "A2": {"value": "Team A"}, "B2": {"value": 8}, "C2": {"value": 7}, "D2": {"value": 9}
            }}],
            "charts": [{
                "id": "radar-1", "sheetId": "data", "type": "radar",
                "source": {"range": {"startRow": 0, "startCol": 0, "endRow": 1, "endCol": 3}, "orientation": "rows"},
                "position": {"x": 12, "y": 16}, "size": {"width": 480, "height": 320},
                "title": "Survey scores", "legend": {"visible": true, "position": "right"},
                "categoryField": 0, "valueFields": [1], "yAxis": {"scaleType": "linear"}
            }]
        }))
        .unwrap();
        let chart = &workbook.charts.as_ref().unwrap()[0];
        let xml = chart_xml(&workbook, chart, 0, None).unwrap();
        assert!(xml.contains("<c:radarChart><c:radarStyle val=\"standard\"/>"));
        assert!(xml.contains(
            "<c:cat><c:strRef><c:f>&apos;Survey&apos;!$B$1:$D$1</c:f>"
        ));
        assert!(xml.contains("<c:val><c:numRef><c:f>&apos;Survey&apos;!$B$2:$D$2</c:f>"));
        assert!(xml.contains(
            "<c:axId val=\"1000002\"/><c:axId val=\"1000003\"/></c:radarChart>"
        ));
        assert!(xml.contains("<c:catAx><c:axId val=\"1000002\"/>"));
        assert!(xml.contains("<c:valAx><c:axId val=\"1000003\"/>"));
        assert!(xml.contains("<c:legendPos val=\"r\"/>"));

        workbook.charts.as_mut().unwrap()[0]
            .y_axis
            .as_mut()
            .unwrap()
            .scale_type = "log".to_owned();
        let chart = &workbook.charts.as_ref().unwrap()[0];
        let error = chart_xml(&workbook, chart, 0, None).unwrap_err();
        assert!(error.message.contains("radar charts require a linear value axis"));
    }

    #[test]
    fn writes_area_pie_doughnut_and_stacked_column_chart_types() {
        let base = r##"{"sheets":[{"id":"sheet-1","name":"Data","cells":{"A1":{"value":"Category"},"B1":{"value":"Value"},"A2":{"value":"One"},"B2":{"value":4},"A3":{"value":"Two"},"B3":{"value":7}}}],"charts":[{"id":"chart-1","sheetId":"sheet-1","type":"bar","source":{"range":{"startRow":0,"startCol":0,"endRow":2,"endCol":1}},"position":{"x":0,"y":0},"size":{"width":320,"height":200},"categoryField":0,"valueFields":[1]}]}"##;
        let cases = [
            ("area", "<c:areaChart>", "<c:catAx>"),
            ("pie", "<c:pieChart>", "<c:firstSliceAng val=\"0\"/>"),
            ("doughnut", "<c:doughnutChart>", "<c:holeSize val=\"55\"/>"),
            (
                "stacked-column",
                "<c:barChart>",
                "<c:overlap val=\"100\"/>",
            ),
        ];
        for (chart_type, expected_chart, expected_setting) in cases {
            let data = base.replace("\"bar\"", &format!("\"{chart_type}\""));
            let workbook: WorkbookInput = serde_json::from_str(&data).unwrap();
            let chart = chart_xml(&workbook, &workbook.charts.as_ref().unwrap()[0], 0, None).unwrap();
            assert!(
                chart.contains(expected_chart),
                "missing {expected_chart} for {chart_type}"
            );
            assert!(
                chart.contains(expected_setting),
                "missing {expected_setting} for {chart_type}"
            );
            if matches!(chart_type, "pie" | "doughnut") {
                assert!(!chart.contains("<c:valAx>"));
            }
        }

        let pivot_data = base.replace(
            "\"type\":\"bar\"",
            "\"type\":\"bar\",\"pivotTableId\":\"pivot-1\"",
        );
        let workbook: WorkbookInput = serde_json::from_str(&pivot_data).unwrap();
        let error = chart_xml(&workbook, &workbook.charts.as_ref().unwrap()[0], 0, None).unwrap_err();
        assert_eq!(error.code, "UNSUPPORTED_EXPORT");
        assert!(error.message.contains("pivot chart parts"));

        let negative_data = base
            .replace("\"type\":\"bar\"", "\"type\":\"pie\"")
            .replace("\"value\":4", "\"value\":-4");
        let workbook: WorkbookInput = serde_json::from_str(&negative_data).unwrap();
        assert!(chart_xml(&workbook, &workbook.charts.as_ref().unwrap()[0], 0, None).is_err());
    }

    #[test]
    fn emits_formula_and_cached_value() {
        let mut registry = StyleRegistry::new();
        let cell = CellInput {
            id: "0:0".to_owned(),
            row: Some(0),
            col: Some(0),
            value: Some(Value::Number(serde_json::Number::from(7))),
            formula: Some("=SUM(A2)".to_owned()),
            display: "7".to_owned(),
            style: None,
            data_type: Some("number".to_owned()),
            cached_data_type: None,
            rich_text: Vec::new(),
            hyperlink: None,
            metadata: None,
        };
        let mut warnings = Vec::new();
        let value = cell_xml(&cell, &mut registry, 0, &mut warnings);
        assert!(value.contains("<f>SUM(A2)</f>"));
        assert!(value.contains("<v>7</v>"));
    }

    #[test]
    fn emits_error_values_and_formula_error_cache() {
        let mut registry = StyleRegistry::new();
        let error = CellInput {
            id: "0:0".to_owned(),
            row: Some(0),
            col: Some(0),
            value: Some(serde_json::json!({"error": "#DIV/0!"})),
            formula: None,
            display: "#DIV/0!".to_owned(),
            style: None,
            data_type: Some("error".to_owned()),
            cached_data_type: None,
            rich_text: Vec::new(),
            hyperlink: None,
            metadata: None,
        };
        let formula_error = CellInput {
            id: "1:0".to_owned(),
            row: Some(1),
            col: Some(0),
            value: Some(serde_json::json!({"error": "#N/A"})),
            formula: Some("=MATCH(A1,B:B,0)".to_owned()),
            display: "#N/A".to_owned(),
            style: None,
            data_type: Some("formula".to_owned()),
            cached_data_type: None,
            rich_text: Vec::new(),
            hyperlink: None,
            metadata: None,
        };
        let mut warnings = Vec::new();
        let error_xml = cell_xml(&error, &mut registry, 0, &mut warnings);
        let formula_error_xml = cell_xml(&formula_error, &mut registry, 0, &mut warnings);

        assert!(error_xml.contains("<c r=\"A1\" s=\"0\" t=\"e\"><v>#DIV/0!</v></c>"));
        assert!(
            formula_error_xml
                .contains("<c r=\"A2\" s=\"0\" t=\"e\"><f>MATCH(A1,B:B,0)</f><v>#N/A</v></c>")
        );
        assert!(warnings.is_empty());
    }

    #[test]
    fn emits_formula_boolean_cache_type() {
        let mut registry = StyleRegistry::new();
        let cell = CellInput {
            id: "0:0".to_owned(),
            row: Some(0),
            col: Some(0),
            value: Some(Value::Bool(true)),
            formula: Some("=A1>0".to_owned()),
            display: "TRUE".to_owned(),
            style: None,
            data_type: Some("formula".to_owned()),
            cached_data_type: Some("boolean".to_owned()),
            rich_text: Vec::new(),
            hyperlink: None,
            metadata: None,
        };
        let mut warnings = Vec::new();
        let value = cell_xml(&cell, &mut registry, 0, &mut warnings);

        assert!(value.contains("<c r=\"A1\" s=\"0\" t=\"b\"><f>A1&gt;0</f><v>1</v></c>"));
        assert!(warnings.is_empty());
    }
}
