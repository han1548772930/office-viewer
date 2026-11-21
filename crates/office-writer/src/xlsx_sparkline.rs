//! XLSX x14 sparkline groups mapped from the canonical per-cell model.
use serde_json::Value;

use crate::{ExportError, ExportWarning, xml};
use std::borrow::Cow;

#[derive(Debug)]
struct Sparkline {
    target: String,
    source: String,
    kind: String,
    color: String,
    negative: Option<String>,
    high: Option<String>,
    low: Option<String>,
    empty: String,
    axis: String,
    min: Option<f64>,
    max: Option<f64>,
}

pub(crate) fn write_xml(
    sheet_name: &str,
    cells: &serde_json::Map<String, Value>,
    warnings: &mut Vec<ExportWarning>,
) -> Result<String, ExportError> {
    let mut groups = Vec::new();
    for (target, cell) in cells {
        let Some(spec) = cell.get("sparkline").filter(|value| !value.is_null()) else { continue };
        let Some(object) = spec.as_object() else {
            warnings.push(ExportWarning::new("INVALID_SPARKLINE", format!("Ignored malformed sparkline at {target}")));
            continue;
        };
        let Some((row, column)) = parse_a1(target) else {
            warnings.push(ExportWarning::new("INVALID_SPARKLINE", format!("Ignored sparkline with invalid target {target:?}")));
            continue;
        };
        if row > 1_048_575 || column > 16_383 {
            return Err(ExportError::invalid("sparkline target exceeds XLSX worksheet limits"));
        }
        let source = object.get("sourceRange").and_then(Value::as_str).unwrap_or_default();
        let Some(source) = source_formula(sheet_name, source) else {
            warnings.push(ExportWarning::new("INVALID_SPARKLINE", format!("Ignored sparkline with invalid source range at {target}")));
            continue;
        };
        let kind = match object.get("type").and_then(Value::as_str).unwrap_or("line") {
            "line" => "line",
            "column" => "column",
            "winloss" => "stacked",
            other => {
                warnings.push(ExportWarning::new("UNSUPPORTED_SPARKLINE_TYPE", format!("Ignored unsupported sparkline type {other:?} at {target}")));
                continue;
            }
        };
        let empty = match object.get("emptyCells").and_then(Value::as_str).unwrap_or("gap") {
            "gap" => "gap",
            "zero" => "zero",
            "connect" => "span",
            other => {
                warnings.push(ExportWarning::new("UNSUPPORTED_SPARKLINE_EMPTY_MODE", format!("Ignored unsupported sparkline empty-cell mode {other:?} at {target}")));
                continue;
            }
        };
        let Some(series_color) = color(object, "color") else {
            warnings.push(ExportWarning::new("INVALID_SPARKLINE_COLOR", format!("Ignored sparkline with invalid color at {target}")));
            continue;
        };
        let axis = if object.get("axis").and_then(Value::as_str) == Some("custom") { "custom" } else { "individual" };
        let min = object.get("min").and_then(Value::as_f64);
        let max = object.get("max").and_then(Value::as_f64);
        if axis == "custom" && (!min.zip(max).is_some_and(|(minimum, maximum)| minimum < maximum)) {
            warnings.push(ExportWarning::new("INVALID_SPARKLINE_AXIS", format!("Ignored sparkline with invalid custom axis at {target}")));
            continue;
        }
        groups.push(Sparkline {
            target: cell_ref(row, column), source, kind: kind.into(), color: series_color,
            negative: color(object, "negativeColor"), high: color(object, "highPointColor"),
            low: color(object, "lowPointColor"), empty: empty.into(), axis: axis.into(), min, max,
        });
    }
    if groups.is_empty() { return Ok(String::new()) }
    let mut out = format!("<extLst><ext uri=\"{{05C60535-1F16-4fd2-B633-F4F36F0B64E0}}\"><x14:sparklineGroups xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\">");
    for group in groups {
        let axis_attrs = if group.axis == "custom" { format!("minAxisType=\"custom\" maxAxisType=\"custom\" manualMin=\"{}\" manualMax=\"{}\"", format_number(group.min.unwrap()), format_number(group.max.unwrap())) } else { "minAxisType=\"individual\" maxAxisType=\"individual\"".into() };
        out.push_str(&format!("<x14:sparklineGroup type=\"{}\" displayEmptyCellsAs=\"{}\" {} negative=\"{}\" high=\"{}\" low=\"{}\"><x14:colorSeries rgb=\"{}\"/>", group.kind, group.empty, axis_attrs, if group.negative.is_some() { 1 } else { 0 }, if group.high.is_some() { 1 } else { 0 }, if group.low.is_some() { 1 } else { 0 }, group.color));
        if let Some(color) = group.negative { out.push_str(&format!("<x14:colorNegative rgb=\"{color}\"/>")); }
        if let Some(color) = group.high { out.push_str(&format!("<x14:colorHigh rgb=\"{color}\"/>")); }
        if let Some(color) = group.low { out.push_str(&format!("<x14:colorLow rgb=\"{color}\"/>")); }
        out.push_str(&format!("<x14:sparklines><x14:sparkline><xm:f>{}</xm:f><xm:sqref>{}</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup>", esc(&group.source), group.target));
    }
    out.push_str("</x14:sparklineGroups></ext></extLst>");
    Ok(out)
}

fn color(spec: &serde_json::Map<String, Value>, key: &str) -> Option<String> {
    let value = spec.get(key)?.as_str()?;
    let hex = value.strip_prefix('#')?;
    if hex.len() != 6 || !hex.bytes().all(|byte| byte.is_ascii_hexdigit()) { return None }
    Some(format!("FF{}", hex.to_ascii_uppercase()))
}

fn source_formula(sheet_name: &str, source: &str) -> Option<String> {
    let source = source.trim();
    let (sheet, reference) = source.rsplit_once('!').map_or((sheet_name, source), |(sheet, reference)| (sheet.trim_matches('\''), reference));
    let refs = reference.replace('$', "");
    let mut parts = refs.split(':');
    let first = parts.next()?;
    let last = parts.next().unwrap_or(first);
    if parts.next().is_some() || parse_a1(first).is_none() || parse_a1(last).is_none() { return None }
    let sheet = sheet.replace('\'', "''");
    Some(format!("'{}'!{}:{}", sheet, first.to_ascii_uppercase(), last.to_ascii_uppercase()))
}

fn parse_a1(value: &str) -> Option<(u32, u32)> {
    let value = value.trim().replace('$', "");
    let mut column = 0u32;
    let mut chars = value.chars().peekable();
    let mut letters = 0;
    while let Some(ch) = chars.peek().copied().filter(char::is_ascii_alphabetic) {
        chars.next(); letters += 1;
        column = column.checked_mul(26)?.checked_add(ch.to_ascii_uppercase() as u32 - 'A' as u32 + 1)?;
    }
    if letters == 0 { return None }
    let row = chars.collect::<String>().parse::<u32>().ok()?.checked_sub(1)?;
    Some((row, column.checked_sub(1)?))
}

fn cell_ref(row: u32, column: u32) -> String {
    let mut label = String::new();
    let mut current = column;
    loop {
        label.insert(0, char::from_u32((current % 26 + 65) as u32).unwrap());
        if current < 26 { break }
        current = current / 26 - 1;
    }
    format!("{label}{}", row + 1)
}

fn format_number(value: f64) -> String {
    let mut result = format!("{value:.10}");
    while result.contains('.') && result.ends_with('0') { result.pop(); }
    if result.ends_with('.') { result.pop(); }
    result
}

fn esc(value: &str) -> Cow<'_, str> { xml::escape(value) }
