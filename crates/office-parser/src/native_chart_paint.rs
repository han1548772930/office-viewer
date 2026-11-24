//! DrawingML paints enter the same typed author model as editor colors.

use std::collections::HashMap;
use office_io_common::chart_color::{ChartColor, ChartLinearGradient, GradientKind, GradientPoint, GradientStop};
use office_io_common::chart_options::{FillPattern, PlotArea};
use office_io_common::chart_xml::{CHART_NS, DRAWING_NS, Node};
use super::{boolean, color, css_color, number};

#[derive(Default)]
pub(super) struct NativePaint {
    pub color: Option<ChartColor>,
    pub opacity: Option<f64>,
    pub pattern: Option<FillPattern>,
}

fn drawing_child<'a>(node: &'a Node, name: &str) -> Option<&'a Node> {
    node.children.iter().find(|child| child.name == name && child.namespace.as_deref() == Some(DRAWING_NS))
}
fn chart_child<'a>(node: &'a Node, name: &str) -> Option<&'a Node> {
    node.children.iter().find(|child| child.name == name && child.namespace.as_deref() == Some(CHART_NS))
}
fn chart_value<'a>(node: &'a Node, name: &str) -> Option<&'a str> {
    chart_child(node, name)?.attrs.get("val").map(String::as_str)
}

pub(super) fn read(shape: Option<&Node>, theme: &HashMap<String, String>) -> Result<NativePaint, String> {
    let Some(shape) = shape else { return Ok(NativePaint::default()); };
    if drawing_child(shape, "noFill").is_some() {
        return Ok(NativePaint { color: Some(ChartColor::from("transparent")), ..Default::default() });
    }
    if let Some(pattern) = drawing_child(shape, "pattFill") {
        let pattern = FillPattern {
            preset: pattern.attrs.get("prst").cloned().unwrap_or_else(|| "pct5".into()),
            foreground_color: css_color(color(drawing_child(pattern, "fgClr"), theme)).unwrap_or_else(|| "#000000".into()),
            background_color: css_color(color(drawing_child(pattern, "bgClr"), theme)).unwrap_or_else(|| "#ffffff".into()),
        };
        pattern.validate()?;
        return Ok(NativePaint { color: Some(ChartColor::from(pattern.foreground_color.clone())), pattern: Some(pattern), opacity: None });
    }
    if let Some(gradient) = drawing_child(shape, "gradFill") {
        let linear = drawing_child(gradient, "lin").ok_or("CHART_NATIVE_GRADIENT_PATH_UNMAPPED")?;
        let angle = match linear.attrs.get("ang") {
            Some(value) => number(Some(value)).filter(|value| (0.0..21_600_000.0).contains(value)).ok_or("chart gradient angle is invalid")?,
            None => 0.0,
        } / 60_000.0;
        let angle = angle.to_radians();
        let snap = |value: f64| if value.abs() < 1e-12 { 0.0 } else if (value.abs() - 1.0).abs() < 1e-12 { value.signum() } else { value };
        let dx = snap(angle.cos());
        let dy = snap(angle.sin());
        let extent = dx.abs() + dy.abs();
        let list = drawing_child(gradient, "gsLst").ok_or("chart gradient is missing its stops")?;
        let stops = list.children.iter().filter(|stop| stop.name == "gs" && stop.namespace.as_deref() == Some(DRAWING_NS)).map(|stop| {
            Ok(GradientStop {
                offset: number(stop.attrs.get("pos").map(String::as_str)).ok_or("chart gradient is missing its stop offset")? / 100_000.0,
                color: css_color(color(Some(stop), theme)).ok_or("chart gradient stop color could not be resolved")?,
                opacity: None,
            })
        }).collect::<Result<Vec<_>, String>>()?;
        let gradient = ChartLinearGradient { kind: GradientKind::Linear,
            start: GradientPoint { x: 0.5 - dx * extent / 2.0, y: 0.5 - dy * extent / 2.0 },
            end: GradientPoint { x: 0.5 + dx * extent / 2.0, y: 0.5 + dy * extent / 2.0 }, stops };
        gradient.validate()?;
        return Ok(NativePaint { color: Some(ChartColor::LinearGradient(gradient)), ..Default::default() });
    }
    let fill = color(drawing_child(shape, "solidFill"), theme);
    Ok(NativePaint { color: fill.as_ref().map(|(color, _)| ChartColor::from(color.clone())), opacity: fill.map(|(_, opacity)| opacity), pattern: None })
}

/// Edge-origin + factor-sized inner layouts share the editor's frame contract.
/// Default-relative or outer layouts require Office's automatic axis geometry.
pub(super) fn plot_area(plot: &Node) -> Result<Option<PlotArea>, String> {
    let Some(layout) = chart_child(plot, "layout").and_then(|layout| chart_child(layout, "manualLayout")) else { return Ok(None); };
    if chart_value(layout, "layoutTarget") != Some("inner") || chart_value(layout, "xMode") != Some("edge") || chart_value(layout, "yMode") != Some("edge")
        || chart_value(layout, "wMode").unwrap_or("factor") != "factor" || chart_value(layout, "hMode").unwrap_or("factor") != "factor" {
        return Ok(None);
    }
    let (Some(left), Some(top), Some(width), Some(height)) = (number(chart_value(layout, "x")), number(chart_value(layout, "y")), number(chart_value(layout, "w")), number(chart_value(layout, "h"))) else { return Ok(None); };
    let area = PlotArea { left, top, width, height };
    area.validate()?;
    Ok(Some(area))
}

pub(super) fn source_differences(root: &Node, mapped_area: bool) -> Vec<String> {
    let mut differences = Vec::new();
    for gradient in root.descendants("gradFill").into_iter().filter(|node| node.namespace.as_deref() == Some(DRAWING_NS)) {
        if let Some(linear) = drawing_child(gradient, "lin") {
            let angle = number(linear.attrs.get("ang").map(String::as_str)).unwrap_or(0.0) / 60_000.0;
            if boolean(linear.attrs.get("scaled").map(String::as_str)) != Some(true) && angle.rem_euclid(90.0).abs() > 1e-9 {
                differences.push("CHART_NATIVE_UNSCALED_GRADIENT_GEOMETRY".into());
            }
        }
        if drawing_child(gradient, "tileRect").is_some_and(|rect| ["l", "t", "r", "b"].iter().any(|key| number(rect.attrs.get(*key).map(String::as_str)).is_some_and(|value| value != 0.0))) {
            differences.push("CHART_NATIVE_GRADIENT_TILE_GEOMETRY".into());
        }
    }
    if root.descendants("ln").into_iter().any(|line| line.namespace.as_deref() == Some(DRAWING_NS) && drawing_child(line, "pattFill").is_some()) {
        differences.push("CHART_NATIVE_PATTERNED_LINE_STROKE".into());
    }
    if !mapped_area && chart_child(root, "chart").and_then(|chart| chart_child(chart, "plotArea")).and_then(|plot| chart_child(plot, "layout")).and_then(|layout| chart_child(layout, "manualLayout")).is_some() {
        differences.push("CHART_NATIVE_DEFAULT_RELATIVE_OR_OUTER_LAYOUT".into());
    }
    differences.sort(); differences.dedup(); differences
}

#[cfg(test)]
#[path = "native_chart_paint_tests.rs"]
mod tests;
