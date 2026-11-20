//! XLSX OOXML writer for the canonical spreadsheet model.
//!
//! This module deliberately emits Office Open XML parts instead of serialising the
//! editor snapshot as JSON.  It owns only the spreadsheet parts; the OPC package
//! and ZIP implementation are supplied by the parent writer crate.

use crate::{ExportError, ExportWarning, opc::Package, xml};
use serde::Deserialize;
use serde_json::{Map, Value};
use std::borrow::Cow;
use std::collections::{BTreeMap, BTreeSet, HashMap};

const MAIN_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL_WORKBOOK: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const REL_WORKSHEET: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const REL_STYLES: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
const REL_CORE: &str =
    "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
const REL_APP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
const REL_HYPERLINK: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
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
    pivot_tables: Option<Vec<Value>>,
}

#[derive(Debug, Deserialize)]
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
    frozen_rows: u32,
    #[serde(default)]
    frozen_cols: u32,
    #[serde(default = "default_true")]
    show_gridlines: bool,
    #[serde(default)]
    hidden_rows: Vec<u32>,
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

#[derive(Debug, Deserialize, Clone)]
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
            color: "FF1F2937".to_owned(),
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
}

fn style_key(style: Option<&StyleInput>) -> StyleKey {
    let default_font = FontKey {
        family: "Aptos".to_owned(),
        size: "11".to_owned(),
        bold: false,
        italic: false,
        underline: false,
        strike: false,
        color: "FF1F2937".to_owned(),
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
    let fill = if fill == "FFFFFFFF" || fill == "FF" {
        "none".to_owned()
    } else {
        fill
    };
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
        rotation: style.rotation.unwrap_or(0).clamp(-90, 180),
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
        color: normalize_color(font.color.as_deref().unwrap_or("FF1F2937")),
    }
}

fn normalize_color(value: &str) -> String {
    let stripped = value.trim().trim_start_matches('#');
    if stripped.eq_ignore_ascii_case("none") || stripped.is_empty() {
        return "none".to_owned();
    }
    let upper = stripped.to_ascii_uppercase();
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
    ["left", "right", "top", "bottom", "diagonal"]
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
        xml_out.push_str(&format!("<color rgb=\"{}\"/>", font.color));
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
    xml_out.push_str("</cellXfs><cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles><dxfs count=\"0\"/><tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium2\" defaultPivotStyle=\"PivotStyleMedium9\"/></styleSheet>");
    xml_out
}

fn border_xml(value: &str) -> String {
    let mut map = BTreeMap::new();
    for part in value.split(';') {
        if let Some((key, data)) = part.split_once('=') {
            map.insert(key, data);
        }
    }
    let edge = |name: &str| -> String {
        let Some(value) = map.get(name).copied() else {
            return format!("<{name}/>");
        };
        if value.is_empty() {
            return format!("<{name}/>");
        }
        let mut pieces = value.splitn(2, ':');
        let style = pieces.next().unwrap_or("thin");
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
    format!(
        "<border>{}<diagonal/></border>",
        ["left", "right", "top", "bottom"]
            .iter()
            .map(|name| edge(name))
            .collect::<String>()
    )
}

fn worksheet_xml(
    sheet: &SheetInput,
    sheet_index: usize,
    registry: &mut StyleRegistry,
    warnings: &mut Vec<ExportWarning>,
) -> (String, Vec<(String, String)>) {
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
    let mut out = String::from(xml::DECL);
    out.push_str(&format!(
        "<worksheet xmlns=\"{MAIN_NS}\" xmlns:r=\"{REL_NS}\">"
    ));
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
    out.push_str("<sheetFormatPr");
    if let Some(width) = sheet.default_column_width {
        out.push_str(&format!(
            " defaultColWidth=\"{}\"",
            format_number(width_to_excel(width))
        ));
    }
    if let Some(height) = sheet.default_row_height {
        out.push_str(&format!(" defaultRowHeight=\"{}\"", format_number(height)));
    }
    out.push_str("/>");
    let mut cols = BTreeMap::<u32, (f64, bool)>::new();
    for (key, width) in &sheet.column_widths {
        if let Some(index) = parse_index(key) {
            cols.insert(index, (*width, false));
        }
    }
    for index in &sheet.hidden_columns {
        let entry = cols
            .entry(*index)
            .or_insert((sheet.default_column_width.unwrap_or(8.43), false));
        entry.1 = true;
    }
    if !cols.is_empty() {
        out.push_str("<cols>");
        for (index, (width, hidden)) in cols {
            out.push_str(&format!(
                "<col min=\"{}\" max=\"{}\" width=\"{}\" customWidth=\"1\"{} />",
                index + 1,
                index + 1,
                format_number(width_to_excel(width)),
                if hidden { " hidden=\"1\"" } else { "" }
            ));
        }
        out.push_str("</cols>");
    }
    out.push_str("<sheetData>");
    let mut by_row = BTreeMap::<u32, Vec<CellInput>>::new();
    for cell in cells {
        by_row.entry(cell.row.unwrap_or(0)).or_default().push(cell);
    }
    for row in &sheet.hidden_rows {
        by_row.entry(*row).or_default();
    }
    for row in sheet.row_heights.keys().filter_map(|key| parse_index(key)) {
        by_row.entry(row).or_default();
    }
    for (row, mut row_cells) in by_row {
        row_cells.sort_by_key(|cell| cell.col.unwrap_or(0));
        let hidden = sheet.hidden_rows.contains(&row);
        let mut attrs = format!(" r=\"{}\"", row + 1);
        if let Some(height) = sheet.row_heights.get(&row.to_string()) {
            attrs.push_str(&format!(
                " ht=\"{}\" customHeight=\"1\"",
                format_number(*height)
            ));
        }
        if hidden {
            attrs.push_str(" hidden=\"1\"");
        }
        out.push_str(&format!("<row{attrs}>"));
        for cell in row_cells {
            out.push_str(&cell_xml(&cell, registry, sheet_index, warnings));
        }
        out.push_str("</row>");
    }
    out.push_str("</sheetData>");
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
    if !hyperlinks.is_empty() {
        out.push_str("<hyperlinks>");
        for (link_index, (reference, target)) in hyperlinks.iter().enumerate() {
            if target.starts_with('#') {
                out.push_str(&format!(
                    "<hyperlink ref=\"{}\" location=\"{}\"/>",
                    reference,
                    esc(target.trim_start_matches('#'))
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
    out.push_str("</worksheet>");
    (out, hyperlinks)
}

fn width_to_excel(width: f64) -> f64 {
    if width > 30.0 { width / 7.0 } else { width }
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
        "charts",
        "pivotTables",
        "slicers",
        "sparklines",
        "conditionalFormatting",
        "dataValidations",
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
    if workbook
        .pivot_tables
        .as_ref()
        .is_some_and(|pivot_tables| !pivot_tables.is_empty())
        && seen.insert("pivotTables".to_owned())
    {
        warnings.push(ExportWarning::new(
            "UNSUPPORTED_FEATURE",
            "Workbook feature pivotTables is not represented by the current XLSX writer and was not exported",
        ));
    }
}

/// Serialize the canonical workbook JSON into XLSX parts.
pub(crate) fn write(data: &str, package: &mut Package) -> Result<Vec<ExportWarning>, ExportError> {
    let workbook: WorkbookInput = serde_json::from_str(data)
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
    let mut warnings = Vec::new();
    add_unsupported_warnings(&workbook, &mut warnings);
    let mut registry = StyleRegistry::new();
    registry.register(None);
    for (index, sheet) in workbook.sheets.iter().enumerate() {
        let (content, hyperlinks) = worksheet_xml(sheet, index, &mut registry, &mut warnings);
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
                "<col min=\"5\" max=\"5\" width=\"8.43\" customWidth=\"1\" hidden=\"1\" />"
            )
        );
        assert!(
            worksheet.contains("<col min=\"6\" max=\"6\" width=\"17.1429\" customWidth=\"1\" />")
        );
        assert!(worksheet.contains("<row r=\"3\" hidden=\"1\"></row>"));
        assert!(worksheet.contains("<row r=\"4\" ht=\"32\" customHeight=\"1\"></row>"));
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
    fn warns_when_canonical_pivot_tables_are_not_exported() {
        let data = r#"{"kind":"workbook","name":"Pivot","sheets":[{"id":"sheet-1","name":"Data","cells":{}}],"pivotTables":[{"id":"pivot-1","source":{},"anchor":{},"fields":[],"options":{}}]}"#;
        let mut package = Package::new();
        let warnings = write(data, &mut package).unwrap();
        assert_eq!(warnings.len(), 1);
        assert_eq!(warnings[0].code, "UNSUPPORTED_FEATURE");
        assert!(warnings[0].message.contains("pivotTables"));
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
