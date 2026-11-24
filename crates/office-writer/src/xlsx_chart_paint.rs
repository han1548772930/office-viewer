//! Author paints -> DrawingML fills. The exact author vector/stops stay in the
//! signature-bound manifest; native angle quantization is reported separately.

use crate::ExportError;
use office_io_common::chart_color::{ChartColor, ChartLinearGradient, GradientKind, GradientPoint, GradientStop, rgba};
use office_io_common::chart_options::{FillPattern, PlotArea, Series};
use super::{ChartInput, SeriesSpec, style};

pub(super) fn fill(color: &ChartColor, opacity: Option<f64>, pattern: Option<&FillPattern>) -> Result<String, ExportError> {
    let opacity = style::number(opacity.unwrap_or(1.0), 0.0, 1.0, "fill opacity")?;
    if let Some(pattern) = pattern {
        pattern.validate().map_err(ExportError::invalid)?;
        return Ok(format!("<a:pattFill prst=\"{}\"><a:fgClr>{}</a:fgClr><a:bgClr>{}</a:bgClr></a:pattFill>",
            pattern.preset, style::color(&pattern.foreground_color, Some(opacity))?, style::color(&pattern.background_color, Some(opacity))?));
    }
    match color {
        ChartColor::Solid(color) => Ok(format!("<a:solidFill>{}</a:solidFill>", style::color(color, Some(opacity))?)),
        ChartColor::LinearGradient(gradient) => linear_fill(gradient, opacity),
    }
}

struct NativeStop { position: f64, rgb: [u8; 3], alpha: f64 }

fn sample(stops: &[NativeStop], position: f64) -> ([u8; 3], f64) {
    if position < stops[0].position { return (stops[0].rgb, stops[0].alpha); }
    let last = &stops[stops.len() - 1];
    if position >= last.position { return (last.rgb, last.alpha); }
    for pair in stops.windows(2) {
        if position < pair[1].position {
            let ratio = (position - pair[0].position) / (pair[1].position - pair[0].position);
            let mut rgb = [0; 3];
            for (index, component) in rgb.iter_mut().enumerate() {
                *component = (f64::from(pair[0].rgb[index]) + (f64::from(pair[1].rgb[index]) - f64::from(pair[0].rgb[index])) * ratio).round() as u8;
            }
            return (rgb, pair[0].alpha + (pair[1].alpha - pair[0].alpha) * ratio);
        }
    }
    (last.rgb, last.alpha)
}

fn linear_fill(gradient: &ChartLinearGradient, opacity: f64) -> Result<String, ExportError> {
    gradient.validate().map_err(ExportError::invalid)?;
    let dx = gradient.end.x - gradient.start.x;
    let dy = gradient.end.y - gradient.start.y;
    let length = dx * dx + dy * dy;
    if length == 0.0 { return Ok("<a:noFill/>".into()); }
    let extent = dx.abs() + dy.abs();
    let minimum = dx.min(0.0) + dy.min(0.0);
    let origin = gradient.start.x * dx + gradient.start.y * dy;
    if ![length, extent, origin].iter().all(|value| value.is_finite()) {
        return Err(ExportError::invalid("chart gradient exceeds native finite geometry"));
    }
    let mut stops = gradient.stops.iter().map(|stop| {
        let (rgb, alpha) = rgba(&stop.color).map_err(ExportError::invalid)?;
        let position = (origin + length * stop.offset - minimum) / extent;
        if !position.is_finite() { return Err(ExportError::invalid("chart gradient stop exceeds native finite geometry")); }
        Ok(NativeStop { position, rgb, alpha: alpha * stop.opacity.unwrap_or(1.0) * opacity })
    }).collect::<Result<Vec<_>, ExportError>>()?;
    stops.sort_by(|left, right| left.position.total_cmp(&right.position));
    let mut visible = stops.iter().filter(|stop| (0.0..=1.0).contains(&stop.position))
        .map(|stop| NativeStop { position: stop.position, rgb: stop.rgb, alpha: stop.alpha }).collect::<Vec<_>>();
    // Replace out-of-mark stops with the interpolated edge colors. Do not clamp
    // their positions: that changes padding and hard-stop transitions.
    if stops[0].position < 0.0 || visible.is_empty() {
        let (rgb, alpha) = sample(&stops, 0.0);
        visible.insert(0, NativeStop { position: 0.0, rgb, alpha });
    }
    if stops[stops.len() - 1].position > 1.0 || visible.len() == 1 {
        let (rgb, alpha) = sample(&stops, 1.0);
        visible.push(NativeStop { position: 1.0, rgb, alpha });
    }
    let angle = (dy.atan2(dx).to_degrees().rem_euclid(360.0) * 60_000.0).round() as u32 % 21_600_000;
    let mut out = "<a:gradFill rotWithShape=\"1\"><a:gsLst>".to_owned();
    for stop in visible {
        out.push_str(&format!("<a:gs pos=\"{}\"><a:srgbClr val=\"{:02X}{:02X}{:02X}\"><a:alpha val=\"{}\"/></a:srgbClr></a:gs>",
            (stop.position * 100_000.0).round() as u32, stop.rgb[0], stop.rgb[1], stop.rgb[2], (stop.alpha * 100_000.0).round() as u32));
    }
    out.push_str(&format!("</a:gsLst><a:lin ang=\"{angle}\" scaled=\"1\"/><a:tileRect/></a:gradFill>"));
    Ok(out)
}

pub(super) fn automatic_enabled(chart: &ChartInput) -> bool {
    chart.options.auto_gradient_fill.unwrap_or(matches!(chart.chart_type.as_str(), "area" | "areaStacked" | "areaPercentStacked" | "area3D" | "areaStacked3D" | "areaPercentStacked3D"))
}

/// Mirrors the shared editor preset; this is a paint adapter, not a data engine.
pub(super) fn series_color(chart: &ChartInput, spec: &SeriesSpec, color: &ChartColor, style: &Series, negative: bool) -> Result<ChartColor, ExportError> {
    let ChartColor::Solid(solid) = color else { return Ok(color.clone()); };
    if !automatic_enabled(chart) || style.fill_pattern.as_ref().and_then(Option::as_ref).is_some()
        || !matches!(spec.kind.plot, "barChart" | "bar3DChart" | "areaChart" | "area3DChart" | "lineChart" | "line3DChart" | "pieChart" | "pie3DChart" | "doughnutChart" | "ofPieChart") {
        return Ok(color.clone());
    }
    let (rgb, alpha) = rgba(solid).map_err(ExportError::invalid)?;
    let tint = |amount: f64| {
        let mut tinted = [0u8; 3];
        for (out, channel) in tinted.iter_mut().zip(rgb) {
            let channel = f64::from(channel) / 255.0;
            let value = if amount >= 0.0 { channel + (1.0 - channel) * amount } else { channel * (1.0 + amount) };
            *out = (value * 255.0).round() as u8;
        }
        office_io_common::chart_color::css_rgba(tinted, alpha)
    };
    let pie = spec.kind.pie;
    let area = matches!(spec.kind.plot, "areaChart" | "area3DChart");
    let horizontal = spec.kind.horizontal;
    let axis = if spec.secondary { chart.options.right_y_axis.as_ref() } else { chart.y_axis.as_ref() };
    let reverse = axis.is_some_and(|axis| axis.reverse == Some(true));
    let (mut start, mut end) = if pie { (GradientPoint { x: 0.0, y: 0.0 }, GradientPoint { x: 1.0, y: 1.0 }) }
        else if horizontal { (GradientPoint { x: 0.0, y: 0.0 }, GradientPoint { x: 1.0, y: 0.0 }) }
        else { (GradientPoint { x: 0.0, y: 1.0 }, GradientPoint { x: 0.0, y: 0.0 }) };
    if !pie && negative != reverse { std::mem::swap(&mut start, &mut end); }
    Ok(ChartColor::LinearGradient(ChartLinearGradient {
        kind: GradientKind::Linear, start, end,
        stops: vec![GradientStop { offset: 0.0, color: tint(if area { 0.0 } else { 0.28 }), opacity: Some(if area { 0.06 } else { 1.0 }) },
            GradientStop { offset: 1.0, color: tint(if area { 0.0 } else if pie { -0.12 } else { -0.08 }), opacity: None }],
    }))
}

pub(super) fn plot_layout(area: Option<&PlotArea>) -> Result<String, ExportError> {
    let Some(area) = area else { return Ok("<c:layout/>".into()); };
    area.validate().map_err(ExportError::invalid)?;
    Ok(format!("<c:layout><c:manualLayout><c:layoutTarget val=\"inner\"/><c:xMode val=\"edge\"/><c:yMode val=\"edge\"/><c:wMode val=\"factor\"/><c:hMode val=\"factor\"/><c:x val=\"{}\"/><c:y val=\"{}\"/><c:w val=\"{}\"/><c:h val=\"{}\"/></c:manualLayout></c:layout>", area.left, area.top, area.width, area.height))
}

#[cfg(test)]
#[path = "xlsx_chart_paint_tests.rs"]
mod tests;
