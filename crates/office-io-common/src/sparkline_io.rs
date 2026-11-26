//! One editable sparkline DTO and bounded Office source contract shared by import/export.

use std::collections::BTreeMap;
use serde::{Deserialize, Serialize};
use crate::conditional_format_xml as source_xml;

#[path = "sparkline_native.rs"]
mod native;
pub use native::{NativeSparklineGroup, decode_native_group, native_group_for_spec, native_group_fingerprint, native_group_signature, native_projection_signature, normalize_native_group};

pub const NATIVE_EXTENSION_URI: &str = "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}";
pub const MODEL_EXTENSION_URI: &str = "{0D86756A-1E63-4A9B-9FDB-0F5A91DD1E77}";
pub const MODEL_NAMESPACE: &str = "urn:office-studio:sparklines:1";
pub const MAX_GROUP_SOURCE_BYTES: usize = 128 * 1024;
pub const MAX_THEME_BYTES: usize = 16 * 1024 * 1024;
pub const MAX_MODEL_BYTES: usize = 64 * 1024 * 1024;
pub const MAX_MEMBERS: usize = 262_144;
pub const COLOR_NAMES: [&str; 8] = ["colorSeries", "colorNegative", "colorAxis", "colorMarkers", "colorFirst", "colorLast", "colorHigh", "colorLow"];

fn line() -> String { "line".into() }
fn gap() -> String { "gap".into() }
fn individual() -> String { "individual".into() }
fn default_color() -> String { "#4472C4".into() }

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct SparklineSpec {
    #[serde(rename = "type", default = "line")]
    pub sparkline_type: String,
    pub source_range: String,
    #[serde(default = "default_color")]
    pub color: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub negative_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub high_point_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub low_point_color: Option<String>,
    #[serde(default = "gap")]
    pub empty_cells: String,
    #[serde(default = "individual")]
    pub axis: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub min: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub max: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub group_id: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub theme_type: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub points: Option<BTreeMap<String, SparklinePoint>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub axis_options: Option<SparklineAxis>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extremum_min: Option<SparklineExtremum>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extremum_max: Option<SparklineExtremum>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub non_numeric_cells: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub contain_hidden_cells: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_width: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_gradient: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_radius: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub office_source: Option<OfficeSparklineSource>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct SparklinePoint {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub visible: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct SparklineAxis {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub visible: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub reverse: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct SparklineExtremum {
    #[serde(rename = "type")]
    pub mode: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<f64>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct OfficeSparklineSource {
    pub version: u8,
    pub group_xml: String,
    pub resolved_colors: BTreeMap<String, String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub theme_xml: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub indexed_colors: Option<Vec<String>>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct SparklineModelRecord {
    pub target: String,
    /// None is reserved for Pie, which deliberately has no native projection.
    pub native_signature: Option<String>,
    pub spec: SparklineSpec,
}

impl SparklineSpec {
    pub fn validate(&self) -> Result<(), String> {
        if !matches!(self.sparkline_type.as_str(), "line" | "column" | "winloss" | "pie") { return Err("SPARKLINE_TYPE_INVALID: unknown sparkline type".into()); }
        normalize_source_reference(&self.source_range)?;
        if !matches!(self.empty_cells.as_str(), "gap" | "zero" | "connect")
            || self.non_numeric_cells.as_deref().is_some_and(|value| !matches!(value, "gap" | "zero" | "connect")) {
            return Err("SPARKLINE_EMPTY_MODE_INVALID: unknown empty/non-numeric mode".into());
        }
        if !matches!(self.axis.as_str(), "individual" | "custom") { return Err("SPARKLINE_AXIS_INVALID: unknown legacy axis mode".into()); }
        if self.theme_type.as_deref().is_some_and(|value| !matches!(value, "standard" | "specialty" | "distinctive" | "soft" | "fresh" | "custom")) {
            return Err("SPARKLINE_THEME_INVALID: unknown theme name".into());
        }
        if self.group_id.as_ref().is_some_and(|value| value.is_empty() || value.trim() != value || value.chars().count() > 200 || value.starts_with("ungrouped:")) {
            return Err("SPARKLINE_GROUP_INVALID: invalid group ID".into());
        }
        if self.line_width.is_some_and(|value| !value.is_finite() || !(0.0..=2112.0).contains(&value)) {
            return Err("SPARKLINE_WIDTH_INVALID: CSS line width is outside the Office admission budget".into());
        }
        for color in [Some(&self.color), self.negative_color.as_ref(), self.high_point_color.as_ref(), self.low_point_color.as_ref(), self.axis_options.as_ref().and_then(|axis| axis.color.as_ref())].into_iter().flatten() {
            validate_color(color)?;
        }
        if let Some(points) = &self.points {
            for (name, point) in points {
                if !matches!(name.as_str(), "highPoint" | "firstPoint" | "lowPoint" | "lastPoint" | "negativePoint" | "markersPoint") {
                    return Err("SPARKLINE_POINT_INVALID: unknown marker".into());
                }
                if let Some(color) = &point.color { validate_color(color)?; }
            }
        }
        let (min_mode, min) = self.extremum(false);
        let (max_mode, max) = self.extremum(true);
        for (mode, value) in [(min_mode, min), (max_mode, max)] {
            if !matches!(mode, "individual" | "group" | "custom") || value.is_some_and(|value| !value.is_finite()) || mode == "custom" && value.is_none() {
                return Err("SPARKLINE_EXTREMUM_INVALID: invalid independent axis mode/value".into());
            }
        }
        if min_mode == "custom" && max_mode == "custom" && min.zip(max).is_some_and(|(min, max)| min >= max) {
            return Err("SPARKLINE_EXTREMUM_INVALID: custom minimum must be below maximum".into());
        }
        if self.min.is_some_and(|value| !value.is_finite()) || self.max.is_some_and(|value| !value.is_finite()) {
            return Err("SPARKLINE_EXTREMUM_INVALID: non-finite retained extrema".into());
        }
        if let Some(source) = &self.office_source { source.validate()?; }
        Ok(())
    }

    pub fn extremum(&self, maximum: bool) -> (&str, Option<f64>) {
        let current = if maximum { self.extremum_max.as_ref() } else { self.extremum_min.as_ref() };
        match current {
            Some(current) => (&current.mode, current.value),
            None => (&self.axis, if maximum { self.max } else { self.min }),
        }
    }

    pub fn point(&self, name: &str) -> (bool, String) {
        let legacy = match name { "negativePoint" => self.negative_color.as_ref(), "highPoint" => self.high_point_color.as_ref(), "lowPoint" => self.low_point_color.as_ref(), _ => None };
        let point = self.points.as_ref().and_then(|points| points.get(name));
        let secondary = match self.theme_type.as_deref().unwrap_or("standard") {
            "specialty" => "#1eb1bd", "distinctive" => "#fc753c", "soft" => "#f468f9", "fresh" => "#f8be15", _ => "#ec64a9",
        };
        (point.and_then(|point| point.visible).unwrap_or(legacy.is_some()),
            point.and_then(|point| point.color.as_ref()).or(legacy).map(String::as_str).unwrap_or(secondary).to_owned())
    }
}

fn validate_color(value: &str) -> Result<(), String> {
    if value.len() > 256 { return Err("SPARKLINE_COLOR_INVALID: color exceeds the admission budget".into()); }
    crate::chart_color::rgba(value).map(|_| ())
}

impl OfficeSparklineSource {
    pub fn validate(&self) -> Result<(), String> {
        if self.version != 1 || self.group_xml.len() > MAX_GROUP_SOURCE_BYTES { return Err("SPARKLINE_SOURCE_INVALID: version or group source budget".into()); }
        let group = source_xml::parse(&self.group_xml)?;
        if !group.is(source_xml::X14_NAMESPACE, "sparklineGroup") || group.children.iter().any(|node| node.is(source_xml::X14_NAMESPACE, "sparklines")) {
            return Err("SPARKLINE_SOURCE_INVALID: source must be a detached x14 group without members".into());
        }
        group.validate_fragment()?;
        for (name, value) in &self.resolved_colors {
            if !COLOR_NAMES.contains(&name.as_str()) { return Err("SPARKLINE_SOURCE_INVALID: unknown color baseline".into()); }
            validate_color(value)?;
        }
        if let Some(theme) = &self.theme_xml {
            if theme.len() > MAX_THEME_BYTES { return Err("SPARKLINE_SOURCE_LIMIT: theme exceeds 16 MiB".into()); }
            let theme = source_xml::parse(theme)?;
            if theme.name != "theme" || !matches!(theme.namespace.as_str(), "http://schemas.openxmlformats.org/drawingml/2006/main" | "http://purl.oclc.org/ooxml/drawingml/main") {
                return Err("SPARKLINE_SOURCE_INVALID: invalid theme source".into());
            }
        }
        if self.indexed_colors.as_ref().is_some_and(|colors| colors.len() > 256 || colors.iter().any(|value| argb(value).is_none())) {
            return Err("SPARKLINE_SOURCE_INVALID: invalid indexed palette".into());
        }
        Ok(())
    }
}

/// Charge serialized DTO bytes per member, including repeated retained themes.
/// A source template is not a second editable model and is never exempt from memory admission.
#[derive(Debug, Default)]
pub struct SparklineModelBudget { bytes: usize, members: usize }
impl SparklineModelBudget {
    pub fn include(&mut self, spec: &SparklineSpec) -> Result<(), String> {
        let size = serde_json::to_vec(spec).map_err(|error| error.to_string())?.len();
        self.bytes = self.bytes.checked_add(size).ok_or("SPARKLINE_MODEL_LIMIT: byte budget overflow")?;
        self.members = self.members.checked_add(1).ok_or("SPARKLINE_MODEL_LIMIT: member budget overflow")?;
        if self.bytes > MAX_MODEL_BYTES || self.members > MAX_MEMBERS { return Err("SPARKLINE_MODEL_LIMIT: retained sparkline model exceeds workbook admission budget".into()); }
        Ok(())
    }
    pub fn replace(&mut self, previous: &SparklineSpec, next: &SparklineSpec) -> Result<(), String> {
        let previous_size = serde_json::to_vec(previous).map_err(|error| error.to_string())?.len();
        let next_size = serde_json::to_vec(next).map_err(|error| error.to_string())?.len();
        self.bytes = self.bytes.checked_sub(previous_size).and_then(|bytes| bytes.checked_add(next_size))
            .ok_or("SPARKLINE_MODEL_LIMIT: replacement byte budget overflow")?;
        if self.bytes > MAX_MODEL_BYTES { return Err("SPARKLINE_MODEL_LIMIT: retained sparkline model exceeds workbook admission budget".into()); }
        Ok(())
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SourceReference { pub sheet: Option<String>, pub first: String, pub last: Option<String> }

/// Preserve relative/absolute anchors and quoted worksheet names. Never strip a foreign qualifier.
pub fn parse_source_reference(input: &str) -> Result<SourceReference, String> {
    let text = input.trim().strip_prefix('=').unwrap_or(input.trim());
    if text.len() > 4096 { return Err("SPARKLINE_REFERENCE_INVALID: source exceeds the reference budget".into()); }
    let mut quoted = false;
    let mut qualifier = None;
    let mut chars = text.char_indices().peekable();
    while let Some((index, character)) = chars.next() {
        if character == '\'' {
            if quoted && chars.peek().is_some_and(|(_, character)| *character == '\'') { chars.next(); } else { quoted = !quoted; }
        } else if character == '!' && !quoted {
            if qualifier.replace(index).is_some() { return Err("SPARKLINE_REFERENCE_INVALID: multiple worksheet qualifiers".into()); }
        }
    }
    if quoted { return Err("SPARKLINE_REFERENCE_INVALID: unclosed worksheet quote".into()); }
    let (sheet, reference) = match qualifier {
        Some(index) => {
            let raw = &text[..index];
            let sheet = if let Some(name) = raw.strip_prefix('\'').and_then(|name| name.strip_suffix('\'')) {
                if name.replace("''", "").contains('\'') { return Err("SPARKLINE_REFERENCE_INVALID: invalid worksheet quote".into()); }
                name.replace("''", "'")
            } else {
                if raw.contains('\'') || raw.trim() != raw { return Err("SPARKLINE_REFERENCE_INVALID: invalid worksheet qualifier".into()); }
                raw.to_owned()
            };
            if sheet.is_empty() || sheet.contains(['[', ']', ':', '\\', '/', '?', '*']) || sheet.chars().any(char::is_control) {
                return Err("SPARKLINE_REFERENCE_UNRESOLVED: external/3D/invalid worksheet source".into());
            }
            (Some(sheet), &text[index + 1..])
        }
        None => (None, text),
    };
    let mut addresses = reference.split(':');
    let first = normalize_address(addresses.next().unwrap_or(""), true).ok_or("SPARKLINE_REFERENCE_UNRESOLVED: source must resolve to an A1 address")?;
    let last = addresses.next().map(|value| normalize_address(value, true).ok_or("SPARKLINE_REFERENCE_INVALID: invalid range endpoint".to_owned())).transpose()?;
    if addresses.next().is_some() { return Err("SPARKLINE_REFERENCE_INVALID: multiple range delimiters".into()); }
    Ok(SourceReference { sheet, first, last })
}

fn quote_sheet(name: &str) -> String { format!("'{}'!", name.replace('\'', "''")) }
pub fn normalize_source_reference(input: &str) -> Result<String, String> {
    let source = parse_source_reference(input)?;
    Ok(format!("{}{}{}", source.sheet.as_deref().map(quote_sheet).unwrap_or_default(), source.first,
        source.last.map(|last| format!(":{last}")).unwrap_or_default()))
}
pub fn source_formula(sheet_name: &str, input: &str) -> Result<String, String> {
    let mut source = parse_source_reference(input)?;
    if source.sheet.is_none() { source.sheet = Some(sheet_name.into()); }
    normalize_source_reference(&format!("{}{}{}", quote_sheet(source.sheet.as_deref().unwrap()), source.first,
        source.last.map(|last| format!(":{last}")).unwrap_or_default()))
}

pub fn normalize_target(input: &str) -> Option<String> { normalize_address(input, false) }
fn normalize_address(input: &str, keep_anchors: bool) -> Option<String> {
    let input = input.trim();
    let mut chars = input.chars().peekable();
    let abs_column = chars.next_if_eq(&'$').is_some();
    let mut column = 0u32;
    let mut letters = String::new();
    while let Some(character) = chars.peek().copied().filter(char::is_ascii_alphabetic) {
        chars.next();
        column = column.checked_mul(26)?.checked_add(character.to_ascii_uppercase() as u32 - 'A' as u32 + 1)?;
        if column > 16_384 { return None; }
        letters.push(character.to_ascii_uppercase());
    }
    let abs_row = chars.next_if_eq(&'$').is_some();
    let digits = chars.collect::<String>();
    if letters.is_empty() || digits.starts_with('0') || !digits.bytes().all(|byte| byte.is_ascii_digit()) { return None; }
    let row = digits.parse::<u32>().ok()?;
    if !(1..=1_048_576).contains(&row) { return None; }
    Some(format!("{}{}{}{}", if keep_anchors && abs_column { "$" } else { "" }, letters, if keep_anchors && abs_row { "$" } else { "" }, row))
}

pub fn argb(input: &str) -> Option<([u8; 3], f64)> {
    let hex = input.strip_prefix('#').unwrap_or(input);
    if !matches!(hex.len(), 6 | 8) || !hex.bytes().all(|byte| byte.is_ascii_hexdigit()) { return None; }
    let (alpha, rgb) = if hex.len() == 8 { (u8::from_str_radix(&hex[..2], 16).ok()?, &hex[2..]) } else { (255, hex) };
    Some(([u8::from_str_radix(&rgb[..2], 16).ok()?, u8::from_str_radix(&rgb[2..4], 16).ok()?, u8::from_str_radix(&rgb[4..], 16).ok()?], f64::from(alpha) / 255.0))
}
pub fn css_argb(input: &str) -> Result<String, String> {
    let (rgb, alpha) = crate::chart_color::rgba(input)?;
    Ok(format!("{:02X}{:02X}{:02X}{:02X}", (alpha * 255.0).round().clamp(0.0, 255.0) as u8, rgb[0], rgb[1], rgb[2]))
}

#[cfg(test)]
#[path = "sparkline_io_tests.rs"]
mod tests;
