//! Native DrawingML appearance. All dimensions enter as editor CSS pixels.

use crate::{ExportError, xml};
use office_io_common::chart_options::{Axis, ChartColor, FillPattern, Font, Label, Line, Marker, Series};
use serde::{Serialize, de::DeserializeOwned};
use serde_json::Value;

pub(super) fn overlay<T: Default + Serialize + DeserializeOwned>(base: Option<&T>, specific: Option<&T>) -> Result<T, ExportError> {
    fn merge(target: &mut Value, source: Value) {
        if let (Some(left), Some(right)) = (target.as_object_mut(), source.as_object()) {
            for (key, value) in right {
                if let Some(previous) = left.get_mut(key) { merge(previous, value.clone()); }
                else { left.insert(key.clone(), value.clone()); }
            }
        } else { *target = source; }
    }
    let mut value = serde_json::to_value(base).map_err(|error| ExportError::invalid(error.to_string()))?;
    if value.is_null() { value = serde_json::json!({}); }
    if let Some(specific) = specific {
        merge(&mut value, serde_json::to_value(specific).map_err(|error| ExportError::invalid(error.to_string()))?);
    }
    serde_json::from_value(value).map_err(|error| ExportError::invalid(error.to_string()))
}

pub(super) fn number(value: f64, min: f64, max: f64, name: &str) -> Result<f64, ExportError> {
    if !value.is_finite() || value < min || value > max {
        return Err(ExportError::invalid(format!("chart {name} must be between {min} and {max}")));
    }
    Ok(value)
}

pub(super) fn integer(value: f64, min: f64, max: f64, name: &str) -> Result<i64, ExportError> {
    Ok(number(value, min, max, name)?.round() as i64)
}

pub(super) fn color(value: &str, opacity: Option<f64>) -> Result<String, ExportError> {
    let (rgb, alpha) = office_io_common::chart_color::rgba(value).map_err(ExportError::invalid)?;
    let alpha = alpha * number(opacity.unwrap_or(1.0), 0.0, 1.0, "opacity")?;
    Ok(format!("<a:srgbClr val=\"{:02X}{:02X}{:02X}\"><a:alpha val=\"{}\"/></a:srgbClr>",
        rgb[0], rgb[1], rgb[2], (alpha * 100_000.0).round() as u32))
}

pub(super) fn line(value: Option<&Line>, fallback_color: Option<&str>) -> Result<String, ExportError> {
    let Some(value) = value else { return Ok(String::new()); };
    let width = value.width.map(|width| number(width, 0.0, 100.0, "line width")).transpose()?;
    let mut out = match width {
        Some(width) => format!("<a:ln w=\"{}\">", (width * 9525.0).round() as u32),
        None => "<a:ln>".into(),
    };
    if value.visible == Some(false) || width == Some(0.0) { out.push_str("<a:noFill/>"); }
    else if let Some(value_color) = value.color.as_deref().or(fallback_color) {
        out.push_str(&format!("<a:solidFill>{}</a:solidFill>", color(value_color, value.opacity)?));
    }
    if let Some(dash) = value.dash_type.as_deref() {
        let dash = match dash { "solid" => "solid", "dashed" => "dash", "dotted" => "dot", _ => return Err(ExportError::invalid("chart line dash type is invalid")) };
        out.push_str(&format!("<a:prstDash val=\"{dash}\"/>"));
    }
    out.push_str("</a:ln>");
    Ok(out)
}

pub(super) fn shape(fill: Option<&str>, opacity: Option<f64>, border: Option<&Line>, stroke_color: Option<&str>) -> Result<String, ExportError> {
    if fill.is_none() && border.is_none() { return Ok(String::new()); }
    let mut out = "<c:spPr>".to_owned();
    if let Some(fill) = fill { out.push_str(&format!("<a:solidFill>{}</a:solidFill>", color(fill, opacity)?)); }
    out.push_str(&line(border, stroke_color)?);
    out.push_str("</c:spPr>");
    Ok(out)
}

pub(super) fn painted_shape(fill: &ChartColor, opacity: Option<f64>, pattern: Option<&FillPattern>, border: Option<&Line>, stroke_color: Option<&str>) -> Result<String, ExportError> {
    Ok(format!("<c:spPr>{}{}</c:spPr>", super::paint::fill(fill, opacity, pattern)?, line(border, stroke_color)?))
}

fn painted_line(value: &Line, fallback: &ChartColor) -> Result<String, ExportError> {
    let width = value.width.map(|width| number(width, 0.0, 100.0, "line width")).transpose()?;
    let mut out = width.map(|width| format!("<a:ln w=\"{}\">", (width * 9525.0).round() as u32)).unwrap_or_else(|| "<a:ln>".into());
    if value.visible == Some(false) || width == Some(0.0) { out.push_str("<a:noFill/>"); }
    else {
        let authored = value.color.as_deref().map(ChartColor::from);
        out.push_str(&super::paint::fill(authored.as_ref().unwrap_or(fallback), value.opacity, None)?);
    }
    if let Some(dash) = value.dash_type.as_deref() {
        let dash = match dash { "solid" => "solid", "dashed" => "dash", "dotted" => "dot", _ => return Err(ExportError::invalid("chart line dash type is invalid")) };
        out.push_str(&format!("<a:prstDash val=\"{dash}\"/>"));
    }
    out.push_str("</a:ln>");
    Ok(out)
}

fn font_run(font: &Font, tag: &str) -> Result<String, ExportError> {
    let mut out = format!("<{tag} lang=\"zh-CN\"");
    if let Some(size) = font.font_size { out.push_str(&format!(" sz=\"{}\"", integer(size * 75.0, 75.0, 19200.0, "font size")?)); }
    if let Some(bold) = font.bold { out.push_str(&format!(" b=\"{}\"", u8::from(bold))); }
    if let Some(italic) = font.italic { out.push_str(&format!(" i=\"{}\"", u8::from(italic))); }
    out.push('>');
    if let Some(value) = font.color.as_deref() { out.push_str(&format!("<a:solidFill>{}</a:solidFill>", color(value, None)?)); }
    if let Some(value) = font.family.as_deref() { out.push_str(&format!("<a:latin typeface=\"{}\"/><a:ea typeface=\"{}\"/>", xml::escape(value), xml::escape(value))); }
    out.push_str(&format!("</{tag}>"));
    Ok(out)
}

pub(super) fn text_properties(font: Option<&Font>, rotate: Option<f64>) -> Result<String, ExportError> {
    if font.is_none() && rotate.is_none() { return Ok(String::new()); }
    let font = font.cloned().unwrap_or_default();
    let body = rotate.map(|rotate| integer(rotate * 60_000.0, -21_600_000.0, 21_600_000.0, "label rotation")).transpose()?
        .map(|rotation| format!("<a:bodyPr rot=\"{rotation}\"/>"))
        .unwrap_or_else(|| "<a:bodyPr/>".into());
    Ok(format!("<c:txPr>{body}<a:lstStyle/><a:p><a:pPr>{}</a:pPr>{}</a:p></c:txPr>",
        font_run(&font, "a:defRPr")?, font_run(&font, "a:endParaRPr")?))
}

pub(super) fn title(title: Option<&str>, subtitle: Option<&str>, title_font: Option<&Font>, subtitle_font: Option<&Font>) -> Result<String, ExportError> {
    let rows = [(title, title_font), (subtitle, subtitle_font)].into_iter().filter_map(|(text, font)| text.filter(|text| !text.is_empty()).map(|text| (text, font)));
    let mut rich = String::new();
    for (text, font) in rows {
        rich.push_str(&format!("<a:p><a:pPr/><a:r>{}<a:t>{}</a:t></a:r><a:endParaRPr lang=\"zh-CN\"/></a:p>",
            font_run(&font.cloned().unwrap_or_default(), "a:rPr")?, xml::escape(text)));
    }
    if rich.is_empty() { return Ok(String::new()); }
    Ok(format!("<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/>{rich}</c:rich></c:tx><c:layout/><c:overlay val=\"0\"/></c:title>"))
}

pub(super) fn marker(value: Option<&Marker>, fallback_color: Option<&str>) -> Result<String, ExportError> {
    let fallback = fallback_color.map(ChartColor::from);
    painted_marker(value, fallback.as_ref(), None, None)
}

pub(super) fn painted_marker(value: Option<&Marker>, fallback_color: Option<&ChartColor>, opacity: Option<f64>, pattern: Option<&FillPattern>) -> Result<String, ExportError> {
    let Some(value) = value else { return Ok(String::new()); };
    let symbol = match value.shape.as_deref().unwrap_or("circle") {
        "circle" => "circle", "rect" => "square", "triangle" => "triangle", "diamond" => "diamond", "none" => "none",
        _ => return Err(ExportError::invalid("chart marker shape is invalid")),
    };
    let mut out = format!("<c:marker><c:symbol val=\"{symbol}\"/>");
    if let Some(size) = value.size {
        if size == 0.0 { return Ok("<c:marker><c:symbol val=\"none\"/></c:marker>".into()); }
        let size = integer(size * 0.75, 0.0, 192.0, "marker size")?;
        // Native markers support 2..72 pt. Keep the exact authored pixel size in
        // the source manifest; the export reports the native representation limit.
        out.push_str(&format!("<c:size val=\"{}\"/>", size.clamp(2, 72)));
    }
    let authored = value.color.as_deref().map(ChartColor::from);
    let color = match fallback_color {
        Some(gradient @ ChartColor::LinearGradient(_)) => Some(gradient),
        fallback => authored.as_ref().or(fallback),
    };
    if let Some(color) = color { out.push_str(&painted_shape(color, opacity, pattern, None, None)?); }
    out.push_str("</c:marker>");
    Ok(out)
}

pub(super) fn labels(value: Option<&Label>, leader_lines: Option<bool>, pie: bool) -> Result<String, ExportError> {
    let default = Label::default();
    let value = value.unwrap_or(&default);
    let visible = value.visible.unwrap_or(false);
    let content = value.content_type.as_deref().unwrap_or("value");
    if !matches!(content, "value" | "category" | "series" | "percentage") { return Err(ExportError::invalid("chart label content type is invalid")); }
    let mut out = format!("<c:dLbls><c:delete val=\"{}\"/>", u8::from(!visible));
    if let Some(format) = value.format.as_deref() { out.push_str(&format!("<c:numFmt formatCode=\"{}\" sourceLinked=\"0\"/>", xml::escape(format))); }
    out.push_str(&text_properties(Some(&value.font), None)?);
    if let Some(position) = value.position.as_deref() {
        let position = match position {
            "inside" => "inEnd", "outside" => "outEnd", "center" => "ctr", "top" => "t", "bottom" => "b", "left" => "l", "right" => "r",
            _ => return Err(ExportError::invalid("chart label position is invalid")),
        };
        out.push_str(&format!("<c:dLblPos val=\"{position}\"/>"));
    }
    out.push_str(&format!("<c:showLegendKey val=\"0\"/><c:showVal val=\"{}\"/><c:showCatName val=\"{}\"/><c:showSerName val=\"{}\"/><c:showPercent val=\"{}\"/>",
        u8::from(visible && content == "value"), u8::from(visible && content == "category"), u8::from(visible && content == "series"), u8::from(visible && content == "percentage" && pie)));
    if let Some(visible) = leader_lines { out.push_str(&format!("<c:showLeaderLines val=\"{}\"/>", u8::from(visible))); }
    out.push_str("</c:dLbls>");
    Ok(out)
}

pub(super) fn axis(axis: Option<&Axis>, id: u32, cross: u32, position: &str, kind: &str, percentage: bool, secondary: bool) -> Result<String, ExportError> {
    let default = Axis::default();
    let axis = axis.unwrap_or(&default);
    axis.validate_intervals(kind == "dateAx").map_err(ExportError::invalid)?;
    if axis.title_alignment.as_deref().is_some_and(|alignment| !matches!(alignment, "start" | "center" | "end")) {
        return Err(ExportError::invalid("chart axis title alignment is invalid"));
    }
    if !matches!(axis.scale_type.as_str(), "linear" | "log") { return Err(ExportError::invalid("chart axis scale type is invalid")); }
    if axis.scale_type == "log" && kind != "valAx" { return Err(ExportError::invalid("chart category and depth axes must use a linear scale")); }
    if axis.min.is_some_and(|min| !min.is_finite()) || axis.max.is_some_and(|max| !max.is_finite()) || axis.min.zip(axis.max).is_some_and(|(min, max)| min >= max) {
        return Err(ExportError::invalid("chart axis limits must be finite and increasing"));
    }
    if axis.scale_type == "log" && (axis.min.is_some_and(|min| min <= 0.0) || axis.max.is_some_and(|max| max <= 0.0)) { return Err(ExportError::invalid("chart logarithmic axis limits must be positive")); }
    let mut out = format!("<c:{kind}><c:axId val=\"{id}\"/><c:scaling>");
    if axis.log_base.is_some() || axis.scale_type == "log" {
        let base = number(axis.log_base.unwrap_or(10.0), 2.0, 1000.0, "logarithmic axis base")?;
        if axis.scale_type == "log" { out.push_str(&format!("<c:logBase val=\"{base}\"/>")); }
    }
    out.push_str(&format!("<c:orientation val=\"{}\"/>", if axis.reverse == Some(true) { "maxMin" } else { "minMax" }));
    let native_value = |value: f64| if percentage { value / 100.0 } else { value };
    if let Some(max) = axis.max { out.push_str(&format!("<c:max val=\"{}\"/>", native_value(max))); }
    if let Some(min) = axis.min { out.push_str(&format!("<c:min val=\"{}\"/>", native_value(min))); }
    let hidden = secondary && (matches!(kind, "catAx" | "dateAx") || position == "b");
    out.push_str(&format!("</c:scaling><c:delete val=\"{}\"/><c:axPos val=\"{position}\"/>", u8::from(hidden)));
    if axis.grid_line.as_ref().is_some_and(|line| line.visible != Some(false)) || axis.grid_line.is_none() && kind == "valAx" && !secondary && position != "b" && position != "t" {
        out.push_str(&format!("<c:majorGridlines>{}</c:majorGridlines>", shape(None, None, axis.grid_line.as_ref(), Some("#d9d9d9"))?));
    }
    out.push_str(&title(axis.title.as_deref(), None, axis.title_style.as_ref(), None)?);
    if kind != "serAx" {
        let format = axis.format.as_deref().unwrap_or(if percentage { "0%" } else if kind == "dateAx" { "yyyy-mm-dd" } else { "General" });
        out.push_str(&format!("<c:numFmt formatCode=\"{}\" sourceLinked=\"{}\"/>", xml::escape(format), u8::from(axis.format.is_none() && !percentage && kind != "dateAx")));
    }
    let tick = axis.tick.as_ref();
    let tick_position = if tick.is_some_and(|tick| tick.line.visible != Some(false)) {
        match tick.and_then(|tick| tick.position.as_deref()).unwrap_or("outside") { "inside" => "in", "outside" => "out", _ => return Err(ExportError::invalid("chart tick position is invalid")) }
    } else { "none" };
    out.push_str(&format!("<c:majorTickMark val=\"{tick_position}\"/><c:minorTickMark val=\"none\"/><c:tickLblPos val=\"{}\"/>", if axis.label.as_ref().is_some_and(|label| label.visible == Some(false)) { "none" } else { "nextTo" }));
    out.push_str(&shape(None, None, axis.line.as_ref(), Some("#767676"))?);
    out.push_str(&text_properties(axis.label.as_ref().map(|label| &label.font), axis.label.as_ref().and_then(|label| label.rotate))?);
    out.push_str(&format!("<c:crossAx val=\"{cross}\"/><c:crosses val=\"{}\"/>", if secondary { "max" } else { "autoZero" }));
    if kind == "catAx" { out.push_str("<c:auto val=\"1\"/><c:lblAlgn val=\"ctr\"/><c:lblOffset val=\"100\"/>"); }
    if kind == "dateAx" {
        let base = axis.base_time_unit.as_deref().unwrap_or("days");
        out.push_str(&format!("<c:auto val=\"0\"/><c:lblOffset val=\"100\"/><c:baseTimeUnit val=\"{base}\"/>"));
        if let Some(interval) = axis.major_unit { out.push_str(&format!("<c:majorUnit val=\"{interval}\"/>")); }
        // Office's omitted majorTimeUnit need not inherit baseTimeUnit. Emit
        // the effective unit whenever the author supplies an explicit interval.
        if axis.major_time_unit.is_some() || axis.major_unit.is_some() {
            let unit = axis.major_time_unit.as_deref().unwrap_or(base);
            out.push_str(&format!("<c:majorTimeUnit val=\"{unit}\"/>"));
        }
    }
    if kind == "valAx" {
        out.push_str("<c:crossBetween val=\"between\"/>");
        if let Some(interval) = axis.major_unit { out.push_str(&format!("<c:majorUnit val=\"{}\"/>", native_value(interval))); }
    }
    out.push_str(&format!("</c:{kind}>"));
    Ok(out)
}

pub(super) fn series_shape(series: &Series, color: &ChartColor, fill_opacity: Option<f64>, line_chart: bool) -> Result<String, ExportError> {
    let representative = color.representative().map_err(ExportError::invalid)?;
    if line_chart {
        let mut line_style = series.border.clone().unwrap_or_default();
        // The editor's texture override disables a line gradient; it does not
        // texture the line stroke. Filled series/markers use the pattern fill.
        if line_style.color.is_none() && series.fill_pattern.as_ref().and_then(Option::as_ref).is_some() { line_style.color = Some(representative); }
        Ok(format!("<c:spPr>{}</c:spPr>", painted_line(&line_style, color)?))
    } else { painted_shape(color, fill_opacity, series.fill_pattern.as_ref().and_then(Option::as_ref), series.border.as_ref(), Some(&representative)) }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn native_percentage_limits_and_axis_direction_use_office_units() {
        let value: Axis = serde_json::from_value(serde_json::json!({"scaleType":"linear","min":0,"max":100,"reverse":true,"label":{"rotate":30}})).unwrap();
        let xml = axis(Some(&value), 10, 11, "b", "valAx", true, false).unwrap();
        assert!(xml.contains("<c:orientation val=\"maxMin\"/><c:max val=\"1\"/><c:min val=\"0\"/>"));
        assert!(xml.contains("formatCode=\"0%\" sourceLinked=\"0\""));
        assert!(xml.contains("rot=\"1800000\""));
        assert!(xml.find("<c:txPr>").unwrap() < xml.find("<c:crossAx").unwrap());
    }

    #[test]
    fn nested_series_defaults_preserve_specific_marker_and_label_settings() {
        let base: Series = serde_json::from_value(serde_json::json!({"color":"red","point":{"size":8,"color":"blue"},"label":{"visible":true,"fontSize":16}})).unwrap();
        let specific: Series = serde_json::from_value(serde_json::json!({"point":{"shape":"diamond"},"label":{"color":"green"}})).unwrap();
        let merged: Series = overlay(Some(&base), Some(&specific)).unwrap();
        assert_eq!(merged.point.as_ref().unwrap().size, Some(8.0));
        assert_eq!(merged.point.as_ref().unwrap().shape.as_deref(), Some("diamond"));
        assert_eq!(merged.label.as_ref().unwrap().font.font_size, Some(16.0));
        assert_eq!(merged.label.as_ref().unwrap().visible, Some(true));
    }
}
