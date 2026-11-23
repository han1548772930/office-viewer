//! Native, editable DrawingML chart plots. Placement stays in the XLSX host.

use std::collections::{BTreeMap, BTreeSet};
use crate::{ExportError, xml};
use office_io_common::chart_kind::{NativeChartKind, native_chart_kind};
use office_io_common::chart_options::Series;
use office_io_common::chart_xml;
use super::{WorkbookInput, SheetInput, ChartInput, range_bounds, safe_sheet_names, chart_range_formula, cell_ref, MAX_ROW_INDEX, MAX_COLUMN_INDEX};

#[path = "xlsx_chart_style.rs"]
mod style;
#[path = "xlsx_chart_bindings.rs"]
mod bindings;
#[path = "xlsx_chart_source.rs"]
mod source_view;
#[path = "xlsx_chart_warnings.rs"]
mod warnings;
#[path = "xlsx_chart_dates.rs"]
mod dates;
#[path = "xlsx_chart_series.rs"]
mod series_plan;
use series_plan::specifications;

#[derive(Debug, Clone)]
pub(super) struct NativeVector { formula: String, values: Vec<Option<f64>> }
#[derive(Debug, Clone)]
pub(super) struct NativeCategory { formula: String, values: Vec<String>, dates: Option<Vec<Option<f64>>> }
#[derive(Debug, Clone)]
pub(super) struct NativeGroup { name: String, vectors: BTreeMap<(usize, bool), NativeVector>, date_vectors: BTreeMap<usize, NativeVector>, default_size: Option<NativeVector> }
#[derive(Debug, Clone)]
pub(super) struct NativeHelper { sheet_name: String, key: String }
#[derive(Debug, Default, Clone)]
pub(super) struct NativeBindings {
    category: Option<NativeCategory>,
    vectors: BTreeMap<(usize, bool), NativeVector>,
    date_vectors: BTreeMap<usize, NativeVector>,
    default_size: Option<NativeVector>,
    groups: Vec<NativeGroup>,
    helper: Option<NativeHelper>,
    visible_fields: Option<BTreeSet<usize>>,
    point_count: usize,
    empty: bool,
}

pub(super) fn prepare_native_bindings(workbook: &mut WorkbookInput, diagnostics: &mut Vec<crate::ExportWarning>) -> Result<(), ExportError> {
    bindings::prepare(workbook, diagnostics)
}

const MAX_SERIES: usize = 1024;
const MAX_STATISTICS: usize = 4096;
const DEFAULT_PALETTE: &[&str] = &["#5470c6", "#91cc75", "#fac858", "#ee6666", "#73c0de", "#3ba272", "#fc8452", "#9a60b4", "#ea7ccc"];

pub(super) struct Context<'a> {
    sheet: &'a SheetInput,
    sheet_name: String,
    bounds: (u32, u32, u32, u32),
    field_count: usize,
    point_count: usize,
}

pub(super) fn context<'a>(workbook: &'a WorkbookInput, chart: &ChartInput) -> Result<Context<'a>, ExportError> {
    let pivot = chart.pivot_table_id.as_ref().and_then(|id| workbook.pivot_tables.as_deref().unwrap_or(&[]).iter().find(|pivot| pivot.id == *id));
    let source_id = pivot.map(|pivot| pivot.anchor.sub_unit_id.as_str()).or(chart.source.sheet_id.as_deref()).unwrap_or(&chart.sheet_id);
    let sheet_index = workbook.sheets.iter().position(|sheet| sheet.id == source_id)
        .ok_or_else(|| ExportError::invalid(format!("chart {} references an unknown data source worksheet", chart.id)))?;
    let sheet = &workbook.sheets[sheet_index];
    let sheet_name = safe_sheet_names(workbook.sheets.iter().map(|sheet| sheet.name.as_str())).into_iter().nth(sheet_index)
        .ok_or_else(|| ExportError::invalid("missing chart source worksheet name"))?;
    let bounds = range_bounds(pivot.and_then(|pivot| pivot.output_range.as_ref()).unwrap_or(&chart.source.range))
        .ok_or_else(|| ExportError::invalid(format!("chart {} source range is missing", chart.id)))?;
    let (first_row, first_col, last_row, last_col) = bounds;
    if first_row >= last_row || first_col >= last_col || last_row > MAX_ROW_INDEX || last_col > MAX_COLUMN_INDEX
        || sheet.row_count.is_some_and(|count| last_row >= count) || sheet.column_count.is_some_and(|count| last_col >= count) {
        return Err(ExportError::invalid(format!("chart {} source must contain a header and data within XLSX worksheet limits", chart.id)));
    }
    let (field_count, point_count) = match chart.source.orientation.as_str() {
        "columns" => ((last_col - first_col + 1) as usize, (last_row - first_row) as usize),
        "rows" => ((last_row - first_row + 1) as usize, (last_col - first_col) as usize),
        _ => return Err(ExportError::invalid(format!("chart {} has an invalid source orientation", chart.id))),
    };
    if chart.category_field >= field_count { return Err(ExportError::invalid(format!("chart {} category field is outside the source range", chart.id))); }
    Ok(Context { sheet, sheet_name, bounds, field_count, point_count })
}

fn scalar(sheet: &SheetInput, row: u32, col: u32) -> Option<&serde_json::Value> {
    let coordinate = format!("{row}:{col}");
    let cell = sheet.cells.get(&coordinate).or_else(|| sheet.cells.get(&cell_ref(row, col)))?;
    cell.get("value")
}

fn text(value: Option<&serde_json::Value>) -> String {
    match value { Some(serde_json::Value::String(value)) => value.clone(), Some(serde_json::Value::Number(value)) => value.as_f64().map(|value| ryu_js::Buffer::new().format_finite(value).to_owned()).unwrap_or_default(), Some(serde_json::Value::Bool(value)) => value.to_string(), _ => String::new() }
}

fn numeric(value: Option<&serde_json::Value>) -> Option<f64> {
    match value {
        Some(serde_json::Value::Number(value)) => value.as_f64().filter(|value| value.is_finite()),
        Some(serde_json::Value::String(value)) if !value.trim().is_empty() => value.trim().parse::<f64>().ok().filter(|value| value.is_finite()),
        _ => None,
    }
}

impl Context<'_> {
    fn coordinates(&self, chart: &ChartInput, field: usize, point: Option<usize>) -> Result<(u32, u32), ExportError> {
        if field >= self.field_count { return Err(ExportError::invalid(format!("chart {} field {field} is outside the source range", chart.id))); }
        let (row, col, _, _) = self.bounds;
        let offset = point.map_or(0, |point| point as u32 + 1);
        Ok(if chart.source.orientation == "columns" { (row + offset, col + field as u32) } else { (row + field as u32, col + offset) })
    }
    fn header(&self, chart: &ChartInput, field: usize) -> Result<(String, String), ExportError> {
        let (row, col) = self.coordinates(chart, field, None)?;
        Ok((chart_range_formula(&self.sheet_name, row, col, row, col), text(scalar(self.sheet, row, col))))
    }
    fn vector(&self, chart: &ChartInput, field: usize) -> Result<(String, Vec<Option<f64>>, Vec<String>), ExportError> {
        let (first_row, first_col) = self.coordinates(chart, field, Some(0))?;
        let (last_row, last_col) = self.coordinates(chart, field, Some(self.point_count - 1))?;
        let mut numbers = Vec::with_capacity(self.point_count);
        let mut strings = Vec::with_capacity(self.point_count);
        for point in 0..self.point_count {
            let (row, col) = self.coordinates(chart, field, Some(point))?;
            let value = scalar(self.sheet, row, col);
            numbers.push(numeric(value)); strings.push(text(value));
        }
        Ok((chart_range_formula(&self.sheet_name, first_row, first_col, last_row, last_col), numbers, strings))
    }
}

fn string_cache(values: &[String]) -> String {
    let mut out = format!("<c:strCache><c:ptCount val=\"{}\"/>", values.len());
    for (index, value) in values.iter().enumerate() { out.push_str(&format!("<c:pt idx=\"{index}\"><c:v>{}</c:v></c:pt>", xml::escape(value))); }
    out.push_str("</c:strCache>"); out
}

fn number_cache(values: &[Option<f64>]) -> String {
    number_cache_format(values, "General")
}

fn number_cache_format(values: &[Option<f64>], format: &str) -> String {
    let mut out = format!("<c:numCache><c:formatCode>{}</c:formatCode><c:ptCount val=\"{}\"/>", xml::escape(format), values.len());
    for (index, value) in values.iter().enumerate() { if let Some(value) = value { out.push_str(&format!("<c:pt idx=\"{index}\"><c:v>{value}</c:v></c:pt>")); } }
    out.push_str("</c:numCache>"); out
}

fn number_reference(tag: &str, formula: &str, values: &[Option<f64>]) -> String {
    format!("<c:{tag}><c:numRef><c:f>{}</c:f>{}</c:numRef></c:{tag}>", xml::escape(formula), number_cache(values))
}

#[derive(Debug, Clone)]
struct SeriesSpec {
    field: usize,
    kind: NativeChartKind,
    secondary: bool,
    x_field: Option<usize>,
    size_field: Option<usize>,
    style: Series,
    group_index: Option<usize>,
    source_name: Option<String>,
}

fn statistics(context: &Context<'_>, chart: &ChartInput, spec: &SeriesSpec) -> Result<String, ExportError> {
    let trends = chart.options.trendlines.iter().filter(|trend| trend.series_index == spec.field);
    let errors = chart.options.error_bars.iter().filter(|error| error.series_index == spec.field);
    let mut out = String::new();
    for trend in trends {
        if spec.kind.pie || spec.kind.three_d || spec.kind.plot == "radarChart" || matches!(spec.kind.grouping, Some("stacked" | "percentStacked")) { return Err(ExportError::unsupported("this native plot cannot represent the authored trendline")); }
        if !matches!(trend.kind.as_str(), "exp" | "linear" | "log" | "movingAvg" | "poly" | "pow") { return Err(ExportError::invalid("chart trendline type is invalid")); }
        out.push_str("<c:trendline>");
        if let Some(name) = &trend.name { out.push_str(&format!("<c:name>{}</c:name>", xml::escape(name))); }
        out.push_str(&style::shape(None, None, Some(&trend.line), Some("#666666"))?);
        out.push_str(&format!("<c:trendlineType val=\"{}\"/>", trend.kind));
        if trend.kind == "poly" {
            let order = trend.order.unwrap_or(2);
            if !(2..=6).contains(&order) { return Err(ExportError::invalid("polynomial trendline order must be 2..6")); }
            out.push_str(&format!("<c:order val=\"{order}\"/>"));
        }
        if trend.kind == "movingAvg" {
            let period = trend.period.unwrap_or(2);
            if !(2..=255).contains(&period) { return Err(ExportError::invalid("moving-average trendline period must be 2..255")); }
            out.push_str(&format!("<c:period val=\"{period}\"/>"));
        }
        out.push_str(&format!("<c:dispRSqr val=\"{}\"/><c:dispEq val=\"{}\"/></c:trendline>", u8::from(trend.show_r2.unwrap_or(false)), u8::from(trend.show_formula.unwrap_or(false))));
    }
    let mut directions = BTreeSet::new();
    for error in errors {
        if spec.kind.pie || spec.kind.three_d || spec.kind.plot == "radarChart" || spec.kind.plot == "stockChart" { return Err(ExportError::unsupported("this native plot cannot represent the authored error bars")); }
        if !matches!(error.direction.as_str(), "x" | "y") || error.direction == "x" && !matches!(spec.kind.plot, "scatterChart" | "bubbleChart") { return Err(ExportError::invalid("error bar direction is invalid for this native plot")); }
        if !directions.insert(error.direction.as_str()) { return Err(ExportError::invalid("each native series can contain one error bar per direction")); }
        if !matches!(error.kind.as_str(), "both" | "plus" | "minus") || !matches!(error.value_type.as_str(), "fixedVal" | "percentage" | "stdDev" | "stdErr" | "cust") { return Err(ExportError::invalid("chart error bar type is invalid")); }
        out.push_str(&format!("<c:errBars><c:errDir val=\"{}\"/><c:errBarType val=\"{}\"/><c:errValType val=\"{}\"/><c:noEndCap val=\"{}\"/>", error.direction, error.kind, error.value_type, u8::from(error.no_end_cap.unwrap_or(false))));
        if error.value_type == "cust" {
            for (tag, field, required) in [("plus", error.plus_index, error.kind != "minus"), ("minus", error.minus_index, error.kind != "plus")] {
                if !required { continue; }
                let field = field.ok_or_else(|| ExportError::invalid("custom error bars require their source fields"))?;
                let (formula, values) = bound_vector(context, chart, spec, field, false)?;
                if values.iter().flatten().any(|value| *value < 0.0) { return Err(ExportError::invalid("custom error magnitudes must be non-negative")); }
                out.push_str(&number_reference(tag, &formula, &values));
            }
        } else if error.value_type != "stdErr" {
            let value = style::number(error.value.unwrap_or(1.0), 0.0, f64::MAX, "error value")?;
            out.push_str(&format!("<c:val val=\"{value}\"/>"));
        }
        out.push_str(&style::shape(None, None, error.line.as_ref(), Some("#666666"))?);
        out.push_str("</c:errBars>");
    }
    Ok(out)
}

fn series(context: &Context<'_>, chart: &ChartInput, spec: &SeriesSpec, index: usize) -> Result<String, ExportError> {
    let (name_formula, name) = context.header(chart, spec.field)?;
    let (value_formula, values) = bound_vector(context, chart, spec, spec.field, chart.options.invalid_value_strategy.as_deref() == Some("zero"))?;
    if spec.kind.pie && (!values.iter().flatten().any(|value| *value > 0.0) || values.iter().flatten().any(|value| *value < 0.0)) { return Err(ExportError::invalid("pie values must be non-negative with a positive total")); }
    let color = spec.style.color.as_deref().or_else(|| chart.palette.get(index % chart.palette.len().max(1)).map(String::as_str)).unwrap_or(DEFAULT_PALETTE[index % DEFAULT_PALETTE.len()]);
    let mut out = format!("<c:ser><c:idx val=\"{index}\"/><c:order val=\"{index}\"/>");
    if let Some(name) = spec.style.name.as_deref() { out.push_str(&format!("<c:tx><c:v>{}</c:v></c:tx>", xml::escape(name))); }
    else if let Some(group) = spec.group_index.and_then(|index| chart.native_bindings.as_ref()?.groups.get(index)) { out.push_str(&format!("<c:tx><c:v>{}</c:v></c:tx>", xml::escape(&group.name))); }
    else { out.push_str(&format!("<c:tx><c:strRef><c:f>{}</c:f>{}</c:strRef></c:tx>", xml::escape(&name_formula), string_cache(&[name]))); }
    out.push_str(&style::series_shape(&spec.style, color, matches!(spec.kind.plot, "lineChart" | "line3DChart" | "scatterChart" | "stockChart"))?);
    if matches!(spec.kind.plot, "lineChart" | "line3DChart" | "scatterChart" | "radarChart" | "stockChart") { out.push_str(&style::marker(spec.style.point.as_ref(), Some(color))?); }
    if spec.kind.pie {
        if let Some(explosion) = chart.options.pie.as_ref().and_then(|pie| pie.explosion) { out.push_str(&format!("<c:explosion val=\"{}\"/>", style::integer(explosion, 0.0, 100.0, "pie explosion")?)); }
    }
    for (point_index, point) in &spec.style.data_points {
        if *point_index >= chart.native_bindings.as_ref().map_or(context.point_count, |bindings| bindings.point_count) { continue; }
        out.push_str(&format!("<c:dPt><c:idx val=\"{point_index}\"/>"));
        if matches!(spec.kind.plot, "lineChart" | "line3DChart" | "scatterChart" | "radarChart" | "stockChart") {
            let point_marker = office_io_common::chart_options::Marker { shape: point.shape.clone().or_else(|| spec.style.point.as_ref().and_then(|marker| marker.shape.clone())), size: point.size.or_else(|| spec.style.point.as_ref().and_then(|marker| marker.size)), color: point.color.clone().or_else(|| spec.style.point.as_ref().and_then(|marker| marker.color.clone())) };
            out.push_str(&style::marker(Some(&point_marker), Some(color))?);
        }
        out.push_str(&style::shape(point.color.as_deref().or(Some(color)), point.fill_opacity.or(spec.style.fill_opacity), point.border.as_ref().or(spec.style.border.as_ref()), Some(color))?);
        out.push_str("</c:dPt>");
    }
    let pie = chart.options.pie.as_ref();
    let label = if spec.kind.pie { style::overlay(spec.style.label.as_ref(), pie.and_then(|pie| pie.label.as_ref()))? } else { spec.style.label.clone().unwrap_or_default() };
    out.push_str(&style::labels(Some(&label), pie.and_then(|pie| pie.show_label_line), spec.kind.pie)?);
    out.push_str(&statistics(context, chart, spec)?);
    if let Some(x_field) = spec.x_field {
        let (formula, numbers) = if dates::enabled(chart, spec.kind) { bound_date_vector(chart, spec, x_field)? }
            else { bound_vector(context, chart, spec, x_field, false)? };
        out.push_str(&number_reference("xVal", &formula, &numbers));
        out.push_str(&number_reference("yVal", &value_formula, &values));
        if spec.kind.plot == "bubbleChart" {
            if let Some(size_field) = spec.size_field {
                let (formula, numbers) = bound_vector(context, chart, spec, size_field, false)?;
                out.push_str(&number_reference("bubbleSize", &formula, &numbers));
            } else if let Some(vector) = chart.native_bindings.as_ref().and_then(|bindings| {
                if let Some(index) = spec.group_index { bindings.groups.get(index)?.default_size.as_ref() } else { bindings.default_size.as_ref() }
            }) {
                out.push_str(&number_reference("bubbleSize", &vector.formula, &vector.values));
            } else {
                let values = vec![Some(1.0); values.len()];
                out.push_str(&format!("<c:bubbleSize><c:numLit>{}</c:numLit></c:bubbleSize>", number_cache(&values).replace("<c:numCache>", "").replace("</c:numCache>", "")));
            }
        }
    } else {
        if dates::enabled(chart, spec.kind) {
            let category = chart.native_bindings.as_ref().and_then(|bindings| bindings.category.as_ref())
                .ok_or_else(|| ExportError::invalid("date category formula bindings were not prepared"))?;
            let values = category.dates.as_ref().ok_or_else(|| ExportError::invalid("date categories are missing their numeric serials"))?;
            let format = chart.x_axis.as_ref().and_then(|axis| axis.format.as_deref()).unwrap_or("yyyy-mm-dd");
            out.push_str(&format!("<c:cat><c:numRef><c:f>{}</c:f>{}</c:numRef></c:cat>", xml::escape(&category.formula), number_cache_format(values, format)));
        } else {
            let (formula, categories) = if let Some(categories) = chart.native_bindings.as_ref().and_then(|bindings| bindings.category.as_ref()) { (categories.formula.clone(), categories.values.clone()) }
                else { let (formula, _, categories) = context.vector(chart, chart.category_field)?; (formula, categories) };
            out.push_str(&format!("<c:cat><c:strRef><c:f>{}</c:f>{}</c:strRef></c:cat>", xml::escape(&formula), string_cache(&categories)));
        }
        out.push_str(&number_reference("val", &value_formula, &values));
    }
    if matches!(spec.kind.plot, "lineChart" | "line3DChart" | "scatterChart" | "stockChart") { out.push_str("<c:smooth val=\"0\"/>"); }
    out.push_str("</c:ser>"); Ok(out)
}

fn view(chart: &ChartInput, kind: NativeChartKind) -> Result<String, ExportError> {
    if !kind.three_d { return Ok(String::new()); }
    let view = chart.options.three_d.as_ref().and_then(|settings| settings.view.as_ref());
    let mut out = "<c:view3D>".to_owned();
    let x = view.and_then(|view| view.rot_x).unwrap_or(28.0);
    let mut y = view.and_then(|view| view.rot_y).unwrap_or(28.0);
    if kind.pie { y += chart.options.pie.as_ref().and_then(|pie| pie.start_angle).unwrap_or(-90.0) + 90.0; }
    out.push_str(&format!("<c:rotX val=\"{}\"/>", style::integer(x, -90.0, 90.0, "3D pitch")?));
    if let Some(height) = view.and_then(|view| view.height_percent) { out.push_str(&format!("<c:hPercent val=\"{}\"/>", style::integer(height, 1.0, 500.0, "3D height")?.max(5))); }
    out.push_str(&format!("<c:rotY val=\"{}\"/>", y.rem_euclid(360.0).round() as i64));
    if let Some(depth) = view.and_then(|view| view.depth_percent) { out.push_str(&format!("<c:depthPercent val=\"{}\"/>", style::integer(depth, 1.0, 500.0, "3D depth")?.max(20))); }
    out.push_str(&format!("<c:rAngAx val=\"{}\"/><c:perspective val=\"{}\"/></c:view3D>", u8::from(view.and_then(|view| view.right_angle_axes).unwrap_or(false)), style::integer(view.and_then(|view| view.perspective).unwrap_or(30.0), 0.0, 120.0, "3D perspective")?));
    Ok(out)
}

fn plot_start(kind: NativeChartKind, chart: &ChartInput) -> Result<String, ExportError> {
    let mut out = format!("<c:{}>", kind.plot);
    if kind.plot.starts_with("bar") { out.push_str(&format!("<c:barDir val=\"{}\"/>", if kind.horizontal { "bar" } else { "col" })); }
    if let Some(grouping) = kind.grouping {
        let grouping = if kind.plot == "bar3DChart" && grouping == "clustered" { chart.options.three_d.as_ref().and_then(|three_d| three_d.grouping.as_deref()).unwrap_or(grouping) } else { grouping };
        if !matches!(grouping, "standard" | "clustered" | "stacked" | "percentStacked") { return Err(ExportError::invalid("native chart grouping is invalid")); }
        out.push_str(&format!("<c:grouping val=\"{grouping}\"/>"));
    }
    if let Some(kind) = kind.of_pie_type { out.push_str(&format!("<c:ofPieType val=\"{kind}\"/>")); }
    if kind.plot == "scatterChart" { out.push_str("<c:scatterStyle val=\"marker\"/>"); }
    if kind.plot == "radarChart" { out.push_str(&format!("<c:radarStyle val=\"{}\"/>", if chart.options.radar.as_ref().is_some_and(|radar| radar.fill == Some(true)) { "filled" } else { "standard" })); }
    if kind.plot != "stockChart" { out.push_str(&format!("<c:varyColors val=\"{}\"/>", u8::from(kind.pie))); }
    Ok(out)
}

fn plot_end(kind: NativeChartKind, chart: &ChartInput, series_count: usize, x: u32, y: u32, z: u32) -> Result<String, ExportError> {
    let mut out = String::new();
    if kind.plot == "barChart" || kind.plot == "bar3DChart" {
        let bar = chart.options.bar.as_ref();
        out.push_str(&format!("<c:gapWidth val=\"{}\"/>", style::integer(bar.and_then(|bar| bar.gap_width).unwrap_or(150.0), 0.0, 500.0, "bar gap width")?));
        if kind.three_d { out.push_str(&format!("<c:gapDepth val=\"{}\"/>", style::integer(chart.options.three_d.as_ref().and_then(|settings| settings.gap_depth).unwrap_or(150.0), 0.0, 500.0, "3D gap depth")?)); }
        else { out.push_str(&format!("<c:overlap val=\"{}\"/>", style::integer(bar.and_then(|bar| bar.overlap).unwrap_or(if matches!(kind.grouping, Some("stacked" | "percentStacked")) { 100.0 } else { 0.0 }), -100.0, 100.0, "bar overlap")?)); }
    }
    if matches!(kind.plot, "line3DChart" | "area3DChart") { out.push_str(&format!("<c:gapDepth val=\"{}\"/>", style::integer(chart.options.three_d.as_ref().and_then(|settings| settings.gap_depth).unwrap_or(150.0), 0.0, 500.0, "3D gap depth")?)); }
    if matches!(kind.plot, "pieChart" | "doughnutChart") {
        let angle = chart.options.pie.as_ref().and_then(|pie| pie.start_angle).unwrap_or(-90.0);
        style::number(angle, -360.0, 360.0, "pie start angle")?;
        out.push_str(&format!("<c:firstSliceAng val=\"{}\"/>", (angle + 90.0).rem_euclid(360.0).round() as i64));
    }
    if kind.plot == "doughnutChart" { out.push_str(&format!("<c:holeSize val=\"{}\"/>", style::integer(chart.options.pie.as_ref().and_then(|pie| pie.doughnut_hole).unwrap_or(0.55) * 100.0, 0.0, 95.0, "doughnut hole")?.max(10))); }
    if kind.plot == "ofPieChart" {
        let secondary = chart.options.pie.as_ref().and_then(|pie| pie.secondary_plot.as_ref());
        out.push_str(&format!("<c:gapWidth val=\"{}\"/>", style::integer(secondary.and_then(|secondary| secondary.gap_width).unwrap_or(50.0), 0.0, 500.0, "secondary pie gap")?));
        if let Some(split) = secondary.and_then(|secondary| secondary.split.as_ref()) {
            let split_type = match split.kind.as_str() { "auto" => "auto", "position" => "pos", "value" => "val", "percentage" => "percent", "custom" => "cust", _ => return Err(ExportError::invalid("secondary pie split type is invalid")) };
            out.push_str(&format!("<c:splitType val=\"{split_type}\"/>"));
            if split.kind == "position" { out.push_str(&format!("<c:splitPos val=\"{}\"/>", split.count.unwrap_or(3))); }
            if matches!(split.kind.as_str(), "value" | "percentage") { out.push_str(&format!("<c:splitPos val=\"{}\"/>", style::number(split.less_than.unwrap_or(0.0), 0.0, if split.kind == "percentage" { 100.0 } else { f64::MAX }, "secondary pie threshold")?)); }
            if split.kind == "custom" {
                out.push_str("<c:custSplit>");
                for point in split.point_indexes.as_deref().unwrap_or(&[]) { out.push_str(&format!("<c:secondPiePt val=\"{point}\"/>")); }
                out.push_str("</c:custSplit>");
            }
        }
        out.push_str(&format!("<c:secondPieSize val=\"{}\"/><c:serLines/>", style::integer(secondary.and_then(|secondary| secondary.size).unwrap_or(65.0), 15.0, 100.0, "secondary pie size")?));
    }
    if kind.plot == "bubbleChart" {
        let bubble = chart.options.bubble.as_ref();
        let scale = style::integer(bubble.and_then(|bubble| bubble.scale).unwrap_or(1.0) * 100.0, 0.0, 1_000_000.0, "bubble scale")?.clamp(1, 300);
        let represents = bubble.and_then(|bubble| bubble.size_represents.as_deref()).unwrap_or("area");
        if !matches!(represents, "area" | "width") { return Err(ExportError::invalid("bubble size meaning is invalid")); }
        out.push_str(&format!("<c:bubbleScale val=\"{scale}\"/><c:showNegBubbles val=\"0\"/><c:sizeRepresents val=\"{}\"/>", if represents == "width" { "w" } else { "area" }));
    }
    if kind.plot == "stockChart" {
        out.push_str("<c:hiLowLines/>");
        if series_count == 4 {
            let candle = chart.options.candlestick.as_ref();
            let up = candle.and_then(|candle| candle.rising_color.as_deref()).unwrap_or("#ef4444");
            let down = candle.and_then(|candle| candle.falling_color.as_deref()).unwrap_or("#16a34a");
            let gap = candle.and_then(|candle| candle.candle_width).map(|width| (100.0 / width - 1.0) * 100.0).unwrap_or(150.0).clamp(0.0, 500.0);
            out.push_str(&format!("<c:upDownBars><c:gapWidth val=\"{}\"/><c:upBars>{}</c:upBars><c:downBars>{}</c:downBars></c:upDownBars>", gap.round() as i64,
                style::shape(if candle.is_some_and(|candle| candle.hollow == Some(true)) { Some("transparent") } else { Some(up) }, None, Some(&office_io_common::chart_options::Line { color: Some(up.into()), ..Default::default() }), Some(up))?,
                style::shape(Some(down), None, None, None)?));
        }
    }
    if !kind.pie {
        out.push_str(&format!("<c:axId val=\"{x}\"/><c:axId val=\"{y}\"/>"));
        if kind.three_d { out.push_str(&format!("<c:axId val=\"{z}\"/>")); }
    }
    out.push_str(&format!("</c:{}>", kind.plot)); Ok(out)
}

pub(super) fn chart_xml(workbook: &WorkbookInput, chart: &ChartInput, chart_index: usize, pivot_table_name: Option<&str>) -> Result<String, ExportError> {
    if chart.pivot_table_id.is_some() && pivot_table_name.is_none() { return Err(ExportError::unsupported(format!("XLSX pivot chart {} requires pivot chart parts", chart.id))); }
    if chart.options.mapping.as_ref().and_then(|mapping| mapping.bubble.as_ref()).is_some_and(|mapping| mapping.series_index.is_some()) && chart.native_bindings.is_none() {
        return Err(ExportError::invalid("grouped bubble chart formula bindings were not prepared"));
    }
    let context = context(workbook, chart)?;
    let specs = specifications(&context, chart)?;
    let first_kind = specs.first().map(|spec| spec.kind).or_else(|| native_chart_kind(if chart.chart_type == "combination" { "column" } else { &chart.chart_type }))
        .ok_or_else(|| ExportError::unsupported("native chart type is not supported"))?;
    let date_mode = dates::enabled(chart, first_kind);
    if date_mode { dates::validate_axis(chart, workbook.date1904)?; }
    if chart.options.trendlines.len() + chart.options.error_bars.len() > MAX_STATISTICS { return Err(ExportError::invalid("chart statistics configuration exceeds its export budget")); }
    for field in chart.options.trendlines.iter().map(|trend| trend.series_index).chain(chart.options.error_bars.iter().map(|error| error.series_index)) { context.coordinates(chart, field, None)?; }
    if chart.options.mapping.as_ref().and_then(|mapping| mapping.category_indexes.as_ref()).is_some_and(|fields| fields.as_slice() != [chart.category_field])
        && !matches!(first_kind.plot, "scatterChart" | "bubbleChart" | "stockChart") && !specs.is_empty()
        && chart.native_bindings.as_ref().and_then(|bindings| bindings.category.as_ref()).is_none() { return Err(ExportError::invalid("multi-field chart category formula bindings were not prepared")); }
    let axis_base = u32::try_from(chart_index).ok().and_then(|index| index.checked_mul(5)).and_then(|index| index.checked_add(1_000_002)).filter(|index| *index <= u32::MAX - 4).ok_or_else(|| ExportError::invalid("too many charts"))?;
    let mut out = format!("{}<c:chartSpace xmlns:c=\"{}\" xmlns:a=\"{}\"><c:date1904 val=\"{}\"/>", xml::DECL, chart_xml::CHART_NS, chart_xml::DRAWING_NS, u8::from(workbook.date1904));
    if let Some(name) = pivot_table_name { out.push_str(&format!("<c:pivotSource><c:name>{}</c:name><c:fmtId val=\"0\"/></c:pivotSource>", xml::escape(name))); }
    out.push_str("<c:chart>");
    out.push_str(&style::title(chart.title.as_deref(), chart.subtitle.as_deref(), chart.options.title_style.as_ref(), chart.options.subtitle_style.as_ref())?);
    out.push_str("<c:autoTitleDeleted val=\"0\"/>");
    out.push_str(&view(chart, first_kind)?);
    out.push_str("<c:plotArea><c:layout/>");
    let mut groups: Vec<(NativeChartKind, bool, Vec<usize>)> = Vec::new();
    for (index, spec) in specs.iter().enumerate() {
        if let Some(group) = groups.iter_mut().find(|group| group.0 == spec.kind && group.1 == spec.secondary) { group.2.push(index); }
        else { groups.push((spec.kind, spec.secondary, vec![index])); }
    }
    // Native plots permit zero series. Keep a real empty chart when its source
    // roles or all samples are hidden, while the manifest retains the binding.
    if groups.is_empty() { groups.push((first_kind, false, Vec::new())); }
    for (kind, secondary, indexes) in &groups {
        out.push_str(&plot_start(*kind, chart)?);
        for index in indexes { out.push_str(&series(&context, chart, &specs[*index], *index)?); }
        let base = axis_base + if *secondary { 2 } else { 0 };
        out.push_str(&plot_end(*kind, chart, indexes.len(), base, base + 1, axis_base + 4)?);
    }
    if !first_kind.pie {
        let xy = matches!(first_kind.plot, "scatterChart" | "bubbleChart");
        let horizontal = first_kind.horizontal;
        let mut x_axis = chart.x_axis.clone().unwrap_or_default();
        if date_mode && x_axis.format.is_none() { x_axis.format = Some("yyyy-mm-dd".into()); }
        let x_kind = if xy { "valAx" } else if date_mode && !first_kind.three_d && chart.pivot_table_id.is_none() { "dateAx" } else { "catAx" };
        for secondary in [false, true] {
            if !(specs.is_empty() && !secondary) && !specs.iter().any(|spec| spec.secondary == secondary) { continue; }
            let base = axis_base + if secondary { 2 } else { 0 };
            let percentage = specs.iter().filter(|spec| spec.secondary == secondary).all(|spec| spec.kind.grouping == Some("percentStacked"));
            out.push_str(&style::axis(Some(&x_axis), base, base + 1, if horizontal { "l" } else { "b" }, x_kind, false, secondary)?);
            out.push_str(&style::axis(if secondary { chart.options.right_y_axis.as_ref() } else { chart.y_axis.as_ref() }, base + 1, base, if horizontal { if secondary { "t" } else { "b" } } else if secondary { "r" } else { "l" }, "valAx", percentage, secondary)?);
        }
        if first_kind.three_d {
            let depth = chart.options.depth_axis.as_ref().map(|depth| depth.to_axis());
            out.push_str(&style::axis(depth.as_ref(), axis_base + 4, axis_base, "b", "serAx", false, false)?);
        }
    }
    out.push_str("</c:plotArea>");
    if !specs.is_empty() && chart.legend.as_ref().is_none_or(|legend| legend.visible) {
        let position = match chart.legend.as_ref().map(|legend| legend.position.as_str()).unwrap_or("bottom") { "top" => "t", "right" => "r", "bottom" => "b", "left" => "l", _ => return Err(ExportError::invalid("chart legend position is invalid")) };
        out.push_str(&format!("<c:legend><c:legendPos val=\"{position}\"/><c:layout/><c:overlay val=\"0\"/>{}</c:legend>", style::text_properties(chart.legend.as_ref().and_then(|legend| legend.font.as_ref()), None)?));
    }
    let blanks = match chart.options.invalid_value_strategy.as_deref().unwrap_or("break") { "zero" => "zero", "break" => "gap", "link" => "span", _ => return Err(ExportError::invalid("chart invalid-value strategy is invalid")) };
    // Helpers already exclude the authored source's current hidden data.
    // They reside on a hidden worksheet, so Office must read their data.
    let native_visible_only = chart.native_bindings.as_ref().is_none_or(|bindings| bindings.helper.is_none()) && chart.options.plot_visible_only.unwrap_or(true);
    out.push_str(&format!("<c:plotVisOnly val=\"{}\"/><c:dispBlanksAs val=\"{blanks}\"/></c:chart>", u8::from(native_visible_only)));
    let appearance = chart.options.appearance.as_ref();
    out.push_str(&style::shape(appearance.and_then(|appearance| appearance.background.as_deref()), None, appearance.and_then(|appearance| appearance.border.as_ref()), None)?);
    out.push_str(&style::text_properties(appearance.and_then(|appearance| appearance.font.as_ref()), None)?);
    let native = format!("{out}</c:chartSpace>");
    let signature = chart_xml::signature(&chart_xml::parse(&native).map_err(ExportError::invalid)?);
    let (first_row, first_col, last_row, last_col) = context.bounds;
    let source = serde_json::json!({"version":1,"nativeSignature":signature,"authored":{
        "type":chart.chart_type,"source":{"sheetName":context.sheet_name,"orientation":chart.source.orientation,"range":{"startRow":first_row,"startCol":first_col,"endRow":last_row,"endCol":last_col}},
        "categoryField":chart.category_field,"valueFields":chart.value_fields,"seriesToCategory":chart.series_to_category,"options":chart.options,
        "title":chart.title,"subtitle":chart.subtitle,"palette":chart.palette,"xAxis":chart.x_axis,"yAxis":chart.y_axis,
        "legend":chart.legend.as_ref().map(|legend| serde_json::json!({"visible":legend.visible,"position":legend.position,"font":legend.font})),
        "helperSheets":chart.native_bindings.as_ref().and_then(|bindings| bindings.helper.as_ref()).map(|helper| vec![serde_json::json!({"name":helper.sheet_name,"key":helper.key})]).unwrap_or_default()
    }});
    out.push_str(&format!("<c:extLst><c:ext uri=\"{}\"><ov:source xmlns:ov=\"{}\">{}</ov:source></c:ext></c:extLst></c:chartSpace>", office_io_common::chart_options::SOURCE_EXTENSION_URI, office_io_common::chart_options::SOURCE_EXTENSION_NS, xml::escape(&source.to_string())));
    Ok(out)
}

fn bound_vector(context: &Context<'_>, chart: &ChartInput, spec: &SeriesSpec, field: usize, as_zero: bool) -> Result<(String, Vec<Option<f64>>), ExportError> {
    if let Some(group_index) = spec.group_index {
        let vector = chart.native_bindings.as_ref().and_then(|bindings| bindings.groups.get(group_index)).and_then(|group| group.vectors.get(&(field, as_zero)))
            .ok_or_else(|| ExportError::invalid("native bubble group is missing its formula vector"))?;
        return Ok((vector.formula.clone(), vector.values.clone()));
    }
    if let Some(vector) = chart.native_bindings.as_ref().and_then(|bindings| bindings.vectors.get(&(field, as_zero))) { return Ok((vector.formula.clone(), vector.values.clone())); }
    let (formula, mut values, _) = context.vector(chart, field)?;
    if as_zero { for value in &mut values { if value.is_none() { *value = Some(0.0); } } }
    Ok((formula, values))
}

fn bound_date_vector(chart: &ChartInput, spec: &SeriesSpec, field: usize) -> Result<(String, Vec<Option<f64>>), ExportError> {
    let binding = chart.native_bindings.as_ref().ok_or_else(|| ExportError::invalid("date X formula bindings were not prepared"))?;
    let vector = if let Some(group) = spec.group_index {
        binding.groups.get(group).and_then(|group| group.date_vectors.get(&field))
    } else { binding.date_vectors.get(&field) }.ok_or_else(|| ExportError::invalid("date X formula binding is missing"))?;
    Ok((vector.formula.clone(), vector.values.clone()))
}
