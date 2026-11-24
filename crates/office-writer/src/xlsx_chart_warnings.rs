//! Observable native representation differences are returned with the file.
//! Retaining authored settings is not a claim that Office renders them.

use crate::{ExportError, ExportWarning};
use office_io_common::chart_options::{Axis, ChartColor, Label, Series};
use serde_json::{Map, Value};
use super::{ChartInput, native_chart_kind};

pub(super) fn append(chart: &ChartInput, warnings: &mut Vec<ExportWarning>) -> Result<(), ExportError> {
    let Some(kind) = native_chart_kind(if chart.chart_type == "combination" { "column" } else { &chart.chart_type }) else { return Ok(()); };
    let options = &chart.options;
    let report = |warnings: &mut Vec<ExportWarning>, code, message: String| warnings.push(ExportWarning::new(code, format!("Chart {}: {message}; exact authored settings remain in the source manifest", chart.id)));
    if options.axis_pointer.as_ref().is_some_and(|pointer| pointer.indicator_label_color.is_some() || pointer.indicator_line_type.is_some()
        || pointer.indicator_line_color.is_some() || pointer.indicator_label_text_color.is_some()) {
        report(warnings, "CHART_NATIVE_AXIS_POINTER", "native Office charts have no equivalent authored hover-indicator configuration".into());
    }
    if options.mapping.as_ref().and_then(|mapping| mapping.use_date_axis) == Some(true) && (kind.three_d || chart.pivot_table_id.is_some()) {
        report(warnings, "CHART_NATIVE_DATE_CATEGORY_SPACING", "native 3D and pivot charts retain date labels on a category axis; elapsed-time spacing needs its native host adapter".into());
    }
    if options.three_d.as_ref().and_then(|three_d| three_d.view.as_ref()).is_some_and(|view| view.height_percent.is_some_and(|height| height < 5.0) || view.depth_percent.is_some_and(|depth| depth < 20.0)) && kind.three_d {
        report(warnings, "CHART_NATIVE_3D_DIMENSION_CLAMPED", "native 3D height/depth are at least 5%/20%".into());
    }
    if chart.chart_type == "doughnut" && options.pie.as_ref().and_then(|pie| pie.doughnut_hole).is_some_and(|hole| hole < 0.1) {
        report(warnings, "CHART_NATIVE_DOUGHNUT_HOLE_CLAMPED", "native doughnut holes are at least 10%".into());
    }
    if kind.plot == "bubbleChart" && options.bubble.as_ref().and_then(|bubble| bubble.scale).is_some_and(|scale| !(0.01..=3.0).contains(&scale)) {
        report(warnings, "CHART_NATIVE_BUBBLE_SCALE_CLAMPED", "native bubble scales are between 1% and 300%".into());
    }
    if kind.plot == "stockChart" && options.candlestick.as_ref().and_then(|candle| candle.candle_width).is_some_and(|width| width < 100.0 / 6.0) {
        report(warnings, "CHART_NATIVE_CANDLE_WIDTH_CLAMPED", "native stock gap widths are at most 500%".into());
    }
    if kind.plot == "radarChart" && options.radar.as_ref().is_some_and(|radar| radar.shape.as_deref() == Some("circle")) {
        report(warnings, "CHART_NATIVE_RADAR_GRID", "the traditional native radar chart uses a polygon grid".into());
    }
    if matches!(kind.plot, "areaChart" | "area3DChart") && options.area.as_ref().is_some_and(|area| area.smooth == Some(true)) {
        report(warnings, "CHART_NATIVE_AREA_SMOOTH", "traditional native areas do not carry a smooth-boundary setting".into());
    }
    if kind.pie {
        if let Some(pie) = &options.pie {
            if pie.rose_pie == Some(true) || pie.is_half_pie == Some(true) {
                report(warnings, "CHART_NATIVE_PIE_GEOMETRY", "traditional native pies do not carry rose radii or half-pie geometry".into());
            }
            label_warnings(pie.label.as_ref(), true, chart, warnings);
        }
    }
    let aggregation = options.extensions.get("aggregation");
    if aggregation.is_some_and(|value| value.get("aggregate").and_then(serde_json::Value::as_bool) == Some(true) || value.get("topN").is_some_and(|value| !value.is_null())) {
        report(warnings, "CHART_NATIVE_SOURCE_TRANSFORM", "native category aggregation/top-N formula bindings still require their source-transform adapter".into());
    }
    series_warnings(options.all_series_style.as_ref(), kind.pie, chart, warnings);
    for series in options.series.values() { series_warnings(Some(series), kind.pie, chart, warnings); }
    for entry in &options.named_series { series_warnings(Some(&entry.style), kind.pie, chart, warnings); }
    paint_warnings(chart, kind, warnings);
    let unmapped = unmapped_paths(chart)?;
    if !unmapped.is_empty() {
        report(warnings, "CHART_NATIVE_UNMAPPED_SETTING", format!("native settings adapters have not mapped {}", unmapped.join(", ")));
    }
    for axis in [chart.x_axis.as_ref(), chart.y_axis.as_ref(), options.right_y_axis.as_ref()].into_iter().flatten() { axis_warnings(axis, chart, warnings); }
    if let Some(depth) = &options.depth_axis { axis_warnings(&depth.to_axis(), chart, warnings); }
    // Diagnostics may be reached through both an inherited and a field-specific
    // style. Return one instance of each distinct message per chart.
    let mut seen = std::collections::BTreeSet::new();
    warnings.retain(|warning| seen.insert((warning.code, warning.message.clone())));
    Ok(())
}

fn paint_warnings(chart: &ChartInput, kind: office_io_common::chart_kind::NativeChartKind, warnings: &mut Vec<ExportWarning>) {
    let styles = chart.options.all_series_style.iter().chain(chart.options.series.values()).chain(chart.options.named_series.iter().map(|entry| &entry.style));
    let mut explicit = false;
    let mut geometry = false;
    let mut patterned_line = false;
    for style in styles {
        let is_line = style.kind.as_deref().unwrap_or(&chart.chart_type) == "line" || matches!(kind.plot, "lineChart" | "line3DChart" | "scatterChart" | "stockChart");
        patterned_line |= is_line && style.fill_pattern.as_ref().and_then(Option::as_ref).is_some();
        for color in style.color.iter().chain(style.data_points.values().filter_map(|point| point.color.as_ref())) {
            if let ChartColor::LinearGradient(gradient) = color {
                explicit = true;
                let horizontal = gradient.start.y == gradient.end.y && gradient.start.x.min(gradient.end.x) == 0.0 && gradient.start.x.max(gradient.end.x) == 1.0;
                let vertical = gradient.start.x == gradient.end.x && gradient.start.y.min(gradient.end.y) == 0.0 && gradient.start.y.max(gradient.end.y) == 1.0;
                geometry |= !horizontal && !vertical;
            }
        }
        patterned_line |= is_line && style.data_points.values().any(|point| point.fill_pattern.as_ref().and_then(Option::as_ref).is_some());
    }
    let report = |warnings: &mut Vec<ExportWarning>, code, message| warnings.push(ExportWarning::new(code,
        format!("Chart {}: {message}; exact authored paints remain in the signature-bound source manifest", chart.id)));
    if geometry { report(warnings, "CHART_NATIVE_GRADIENT_GEOMETRY", "DrawingML uses a quantized linear angle and a native fill rectangle; diagonal or nonstandard editor vectors need independent visual verification"); }
    if (explicit || super::paint::automatic_enabled(chart)) && (kind.pie || kind.three_d || matches!(kind.plot, "areaChart" | "area3DChart")) {
        report(warnings, "CHART_NATIVE_GRADIENT_FILL_BOUNDS", "native pie, 3D and area fill bounds can differ from the editor's mark or plot bounds");
    }
    if chart.options.auto_gradient_fill == Some(true) && !matches!(kind.plot, "barChart" | "bar3DChart" | "areaChart" | "area3DChart" | "lineChart" | "line3DChart" | "pieChart" | "pie3DChart" | "doughnutChart" | "ofPieChart") {
        report(warnings, "CHART_NATIVE_AUTOMATIC_GRADIENT_UNMAPPED", "this native plot has no automatic-gradient paint adapter yet");
    }
    if patterned_line { report(warnings, "CHART_NATIVE_PATTERNED_LINE_STROKE", "the editor texture setting paints fills and markers; a patterned native line stroke still requires its own adapter"); }
}

/// Compare retained source controls with the typed native surface. The raw
/// source remains in the manifest; recognizing an extension is not rendering it.
fn unmapped_paths(chart: &ChartInput) -> Result<Vec<String>, ExportError> {
    fn remove_extensions(value: Option<&mut Value>, extensions: &Map<String, Value>) {
        if let Some(object) = value.and_then(Value::as_object_mut) { for key in extensions.keys() { object.remove(key); } }
    }
    fn strip_series(value: Option<&mut Value>, series: &Series) {
        let Some(value) = value else { return; };
        remove_extensions(Some(&mut *value), &series.extensions);
        for (index, point) in &series.data_points {
            remove_extensions(value.get_mut("dataPoints").and_then(|points| points.get_mut(index.to_string())), &point.extensions);
        }
    }
    fn visit(source: &Value, native: Option<&Value>, path: &str, output: &mut Vec<String>) {
        if source.is_null() { return; }
        match source {
            Value::Object(object) => {
                for (key, value) in object {
                    let child = format!("{path}[{}]", Value::String(key.clone()));
                    let mapped = native.and_then(|native| native.get(key));
                    if mapped.is_none() && !value.is_null() { output.push(child); }
                    else { visit(value, mapped, &child, output); }
                }
            }
            Value::Array(array) => for (index, value) in array.iter().enumerate() {
                visit(value, native.and_then(|native| native.get(index)), &format!("{path}[{index}]"), output);
            },
            _ => {}
        }
    }
    let options = &chart.options;
    let mut native = serde_json::to_value(&options.typed).map_err(|error| ExportError::invalid(error.to_string()))?;
    remove_extensions(Some(&mut native), &options.extensions);
    if let Some(mapping) = &options.mapping { remove_extensions(native.get_mut("mapping"), &mapping.extensions); }
    if let Some(series) = &options.all_series_style { strip_series(native.get_mut("allSeriesStyle"), series); }
    for (field, series) in &options.series { strip_series(native.get_mut("series").and_then(|series| series.get_mut(field.to_string())), series); }
    for (index, entry) in options.named_series.iter().enumerate() {
        strip_series(native.get_mut("namedSeries").and_then(|series| series.get_mut(index)).and_then(|entry| entry.get_mut("style")), &entry.style);
    }
    let mut paths = Vec::new();
    visit(&options.source, Some(&native), "options", &mut paths);
    Ok(paths)
}

fn series_warnings(style: Option<&Series>, pie: bool, chart: &ChartInput, warnings: &mut Vec<ExportWarning>) {
    let Some(style) = style else { return; };
    if style.point.as_ref().and_then(|marker| marker.size).is_some_and(clamped_marker)
        || style.data_points.values().any(|point| point.size.is_some_and(clamped_marker)) {
        warnings.push(ExportWarning::new("CHART_NATIVE_MARKER_SIZE_CLAMPED", format!("Chart {}: native markers are between 2 and 72 pt; the original CSS pixel sizes are retained", chart.id)));
    }
    label_warnings(style.label.as_ref(), pie, chart, warnings);
}

fn clamped_marker(size: f64) -> bool { size != 0.0 && !(2.0..=72.0).contains(&(size * 0.75)) }

fn label_warnings(label: Option<&Label>, pie: bool, chart: &ChartInput, warnings: &mut Vec<ExportWarning>) {
    if !pie && label.is_some_and(|label| label.visible == Some(true) && label.content_type.as_deref() == Some("percentage")) {
        warnings.push(ExportWarning::new("CHART_NATIVE_PERCENTAGE_LABEL", format!("Chart {}: percentage labels outside pies require native label formula bindings; authored labels are retained", chart.id)));
    }
}

fn axis_warnings(axis: &Axis, chart: &ChartInput, warnings: &mut Vec<ExportWarning>) {
    if axis.title_alignment.is_some() {
        warnings.push(ExportWarning::new("CHART_NATIVE_AXIS_TITLE_POSITION", format!("Chart {}: native axis titles use Office's automatic position; the authored along-axis title position is retained independently of paragraph text alignment", chart.id)));
    }
    if axis.tick.as_ref().is_some_and(|tick| tick.length.is_some() || tick.line.color.is_some() || tick.line.width.is_some() || tick.line.dash_type.is_some() || tick.line.opacity.is_some()) {
        warnings.push(ExportWarning::new("CHART_NATIVE_TICK_APPEARANCE", format!("Chart {}: native ticks share their axis stroke and Office tick length; independent authored tick settings are retained", chart.id)));
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;
    use super::super::{WorkbookInput, chart_xml, context};

    #[test]
    fn unknown_nested_controls_remain_exact_and_emit_a_native_mapping_diagnostic() {
        let authored = json!({
            "allSeriesStyle":{"customCurve":{"tension":0.3},"point":{"customSymbol":"star"}},
            "series":{"1":{"border":{"customJoin":"round"},"dataPoints":{"0":{"customPoint":{"flags":[1,2]}}}}},
            "namedSeries":[{"name":"Revenue","style":{"label":{"customPosition":"auto"}}}],
            "mapping":{"futureRole":{"field":2}},"customTop":{"enabled":true}
        });
        let book: WorkbookInput = serde_json::from_value(json!({"kind":"workbook","sheets":[{"id":"data","name":"Data","visible":true,
            "rowCount":2,"columnCount":2,"cells":{"0:0":{"value":"Category"},"0:1":{"value":"Revenue"},"1:0":{"value":"A"},"1:1":{"value":10}}}],
            "charts":[{"id":"chart","sheetId":"data","type":"line","categoryField":0,"valueFields":[1],
                "position":{"x":0,"y":0},"size":{"width":480,"height":300},
                "source":{"orientation":"columns","range":{"startRow":0,"startCol":0,"endRow":1,"endCol":1}},"options":authored}]})).unwrap();
        let chart = &book.charts.as_ref().unwrap()[0];
        let paths = unmapped_paths(chart).unwrap();
        assert!(paths.contains(&"options[\"allSeriesStyle\"][\"customCurve\"]".to_owned()));
        assert!(paths.contains(&"options[\"series\"][\"1\"][\"border\"][\"customJoin\"]".to_owned()));
        assert!(paths.contains(&"options[\"namedSeries\"][0][\"style\"][\"label\"][\"customPosition\"]".to_owned()));
        let mut warnings = Vec::new();
        append(chart, &mut warnings).unwrap();
        assert!(warnings.iter().any(|warning| warning.code == "CHART_NATIVE_UNMAPPED_SETTING"));
        let specs = super::super::series_plan::specifications(&context(&book, chart).unwrap(), chart).unwrap();
        assert_eq!(specs[0].style.extensions["customCurve"], json!({"tension":0.3}));
        assert_eq!(specs[0].style.data_points[&0].extensions["customPoint"], json!({"flags":[1,2]}));
        let native = office_io_common::chart_xml::parse(&chart_xml(&book, chart, 0, None).unwrap()).unwrap();
        let manifest: Value = serde_json::from_str(&native.descendants("source")[0].text).unwrap();
        assert_eq!(manifest["authored"]["options"], authored);
    }
}
