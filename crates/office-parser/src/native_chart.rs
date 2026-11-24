//! Native chart XML complements rxls' workbook/cell reader without narrowing its
//! parsed chart list to the types recognized by the dependency.

use std::collections::HashMap;
use office_io_common::chart_kind::{model_chart_kind, NATIVE_CHART_PLOTS};
use office_io_common::chart_options::{self as config, Axis, AxisLabel, AxisTick, Font, Label, Line, Marker, Options, Point, Series};
use office_io_common::chart_xml::{self, Node};
use crate::model::{AnchorCellPayload, ChartErrorBarPayload, ChartPayload, ChartSeriesPayload};

#[path = "native_chart_paint.rs"]
mod paint;

pub(crate) struct NativeChart {
    pub(crate) payload: ChartPayload,
    pub(crate) pivot_source_name: Option<String>,
}

fn boolean(value: Option<&str>) -> Option<bool> { match value? { "1" | "true" | "on" => Some(true), "0" | "false" | "off" => Some(false), _ => None } }
fn number(value: Option<&str>) -> Option<f64> { value?.parse::<f64>().ok().filter(|value| value.is_finite()) }
fn property(node: &Node, name: &str) -> Option<f64> { number(node.value(name)) }

fn color(node: Option<&Node>, theme: &HashMap<String, String>) -> Option<(String, f64)> {
    let fill = node?;
    let drawing = |name: &str| fill.children.iter().find(|child| child.name == name && child.namespace.as_deref() == Some(chart_xml::DRAWING_NS));
    let value = drawing("srgbClr").or_else(|| drawing("schemeClr")).or_else(|| drawing("sysClr"))?;
    let hex = if value.name == "schemeClr" { theme.get(value.attrs.get("val")?).map(String::as_str) }
        else if value.name == "sysClr" { value.attrs.get("lastClr").map(String::as_str) }
        else { value.attrs.get("val").map(String::as_str) }?;
    let hex = hex.trim_start_matches('#');
    if hex.len() != 6 || !hex.bytes().all(|byte| byte.is_ascii_hexdigit()) { return None; }
    let mut channels = [0f64; 3];
    for (index, channel) in channels.iter_mut().enumerate() { *channel = f64::from(u8::from_str_radix(&hex[index * 2..index * 2 + 2], 16).ok()?) / 255.0; }
    let mut alpha = 1.0;
    for transform in &value.children {
        if transform.namespace.as_deref() != Some(chart_xml::DRAWING_NS) { continue; }
        let amount = number(transform.attrs.get("val").map(String::as_str)).unwrap_or(0.0) / 100_000.0;
        match transform.name.as_str() {
            "alpha" => alpha = amount, "alphaMod" => alpha *= amount, "alphaOff" => alpha += amount,
            "tint" => channels.iter_mut().for_each(|channel| *channel = *channel + (1.0 - *channel) * amount),
            "shade" => channels.iter_mut().for_each(|channel| *channel *= amount),
            "lum" => transform_luminance(&mut channels, |_: f64| amount),
            "lumMod" => transform_luminance(&mut channels, |luminance| luminance * amount),
            "lumOff" => transform_luminance(&mut channels, |luminance| luminance + amount),
            _ => {}
        }
        // Each DrawingML transform clamps before the next transform consumes
        // the channels. Delaying this until the final color changes HSL hue.
        for channel in &mut channels { *channel = channel.clamp(0.0, 1.0); }
        alpha = alpha.clamp(0.0, 1.0);
    }
    Some((format!("#{:02x}{:02x}{:02x}", (channels[0].clamp(0.0, 1.0) * 255.0).round() as u8, (channels[1].clamp(0.0, 1.0) * 255.0).round() as u8, (channels[2].clamp(0.0, 1.0) * 255.0).round() as u8), alpha.clamp(0.0, 1.0)))
}

/// DrawingML lum transforms operate on HSL lightness, preserving hue and
/// saturation. Applying them as RGB multiplication loses saturated colors.
fn transform_luminance(channels: &mut [f64; 3], transform: impl FnOnce(f64) -> f64) {
    let max = channels.iter().copied().fold(f64::NEG_INFINITY, f64::max).clamp(0.0, 1.0);
    let min = channels.iter().copied().fold(f64::INFINITY, f64::min).clamp(0.0, 1.0);
    let lightness = (max + min) / 2.0;
    let chroma = max - min;
    let saturation = if chroma == 0.0 { 0.0 } else { chroma / (1.0 - (2.0 * lightness - 1.0).abs()) };
    let hue = if chroma == 0.0 { 0.0 }
        else if max == channels[0] { ((channels[1] - channels[2]) / chroma).rem_euclid(6.0) }
        else if max == channels[1] { (channels[2] - channels[0]) / chroma + 2.0 }
        else { (channels[0] - channels[1]) / chroma + 4.0 };
    let lightness = transform(lightness).clamp(0.0, 1.0);
    let chroma = (1.0 - (2.0 * lightness - 1.0).abs()) * saturation;
    let x = chroma * (1.0 - (hue.rem_euclid(2.0) - 1.0).abs());
    let rgb = if hue < 1.0 { [chroma, x, 0.0] } else if hue < 2.0 { [x, chroma, 0.0] }
        else if hue < 3.0 { [0.0, chroma, x] } else if hue < 4.0 { [0.0, x, chroma] }
        else if hue < 5.0 { [x, 0.0, chroma] } else { [chroma, 0.0, x] };
    let offset = lightness - chroma / 2.0;
    for (channel, component) in channels.iter_mut().zip(rgb) { *channel = (component + offset).clamp(0.0, 1.0); }
}

fn css_color(value: Option<(String, f64)>) -> Option<String> {
    value.map(|(hex, alpha)| if alpha == 1.0 { hex } else {
        let red = u8::from_str_radix(&hex[1..3], 16).unwrap_or(0);
        let green = u8::from_str_radix(&hex[3..5], 16).unwrap_or(0);
        let blue = u8::from_str_radix(&hex[5..7], 16).unwrap_or(0);
        format!("rgba({red}, {green}, {blue}, {alpha})")
    })
}

fn read_line(node: Option<&Node>, theme: &HashMap<String, String>) -> Option<Line> {
    let node = node?;
    let stroke = node.child("ln").unwrap_or(node);
    let fill = color(stroke.child("solidFill"), theme);
    Some(Line {
        visible: Some(stroke.child("noFill").is_none()),
        color: fill.as_ref().map(|fill| fill.0.clone()),
        opacity: fill.map(|fill| fill.1),
        width: number(stroke.attrs.get("w").map(String::as_str)).map(|width| width / 9525.0),
        dash_type: stroke.value("prstDash").and_then(|value| match value { "solid" => Some("solid"), "dash" | "sysDash" | "lgDash" => Some("dashed"), "dot" | "sysDot" => Some("dotted"), _ => None }).map(str::to_owned),
    })
}

fn read_font(node: Option<&Node>, theme: &HashMap<String, String>) -> Option<Font> {
    let node = node?;
    let properties = node.descendants("rPr").into_iter().next().or_else(|| node.descendants("defRPr").into_iter().next()).or_else(|| node.descendants("endParaRPr").into_iter().next()).unwrap_or(node);
    Some(Font {
        color: css_color(color(properties.child("solidFill"), theme)),
        font_size: number(properties.attrs.get("sz").map(String::as_str)).map(|size| size / 75.0),
        bold: boolean(properties.attrs.get("b").map(String::as_str)),
        italic: boolean(properties.attrs.get("i").map(String::as_str)),
        family: properties.child("latin").or_else(|| properties.child("ea")).and_then(|font| font.attrs.get("typeface")).cloned(),
    })
}

fn read_marker(node: Option<&Node>, theme: &HashMap<String, String>) -> Option<Marker> {
    let node = node?;
    Some(Marker {
        shape: node.value("symbol").and_then(|value| match value { "circle" => Some("circle"), "square" => Some("rect"), "triangle" => Some("triangle"), "diamond" => Some("diamond"), "none" => Some("none"), _ => None }).map(str::to_owned),
        size: property(node, "size").map(|size| size / 0.75),
        color: css_color(color(node.child("spPr").and_then(|shape| shape.child("solidFill")), theme)),
    })
}

fn read_label(node: Option<&Node>, theme: &HashMap<String, String>) -> Option<Label> {
    let node = node?;
    let value = boolean(node.value("showVal")).unwrap_or(false);
    let category = boolean(node.value("showCatName")).unwrap_or(false);
    let series = boolean(node.value("showSerName")).unwrap_or(false);
    let percent = boolean(node.value("showPercent")).unwrap_or(false);
    let deleted = boolean(node.value("delete")).unwrap_or(false);
    Some(Label {
        font: read_font(node.child("txPr"), theme).unwrap_or_default(),
        visible: Some(!deleted && (value || category || series || percent)),
        content_type: Some(if percent { "percentage" } else if category { "category" } else if series { "series" } else { "value" }.into()),
        position: node.value("dLblPos").and_then(|value| match value { "inEnd" | "inBase" => Some("inside"), "outEnd" | "bestFit" => Some("outside"), "ctr" => Some("center"), "t" => Some("top"), "b" => Some("bottom"), "l" => Some("left"), "r" => Some("right"), _ => None }).map(str::to_owned),
        format: node.child("numFmt").and_then(|format| format.attrs.get("formatCode")).cloned(),
    })
}

fn read_axis(node: Option<&Node>, theme: &HashMap<String, String>, percentage: bool) -> Result<Axis, String> {
    let Some(node) = node else { return Ok(Axis::default()); };
    let scaling = node.child("scaling");
    let log_base = scaling.and_then(|scaling| property(scaling, "logBase"));
    let limits = |name: &str| scaling.and_then(|scaling| property(scaling, name)).map(|value| if percentage { value * 100.0 } else { value });
    let label_visible = node.value("tickLblPos").map(|value| value != "none").unwrap_or(true) && boolean(node.value("delete")) != Some(true);
    let tick = node.value("majorTickMark");
    let major_unit = node.value("majorUnit").map(|value| value.parse::<f64>().map_err(|_| "chart major tick interval is not a number".to_owned())).transpose()?;
    let axis = Axis {
        title: node.child("title").map(plain_text).filter(|text| !text.is_empty()),
        title_style: read_font(node.child("title"), theme),
        // Native paragraph alignment is not an along-axis title position.
        // Exact authored positions are restored only by the matching manifest.
        title_alignment: None,
        scale_type: if log_base.is_some() { "log" } else { "linear" }.into(), log_base,
        min: limits("min"), max: limits("max"), reverse: scaling.and_then(|scaling| scaling.value("orientation")).map(|value| value == "maxMin"),
        base_time_unit: (node.name == "dateAx").then(|| node.value("baseTimeUnit").map(str::to_owned)).flatten(),
        major_time_unit: (node.name == "dateAx").then(|| node.value("majorTimeUnit").map(str::to_owned)).flatten(),
        major_unit: major_unit.map(|value| if percentage { value * 100.0 } else { value }),
        line: node.child("spPr").and_then(|shape| read_line(shape.child("ln"), theme)),
        label: Some(AxisLabel { font: read_font(node.child("txPr"), theme).unwrap_or_default(), visible: Some(label_visible),
            rotate: node.child("txPr").and_then(|text| text.child("bodyPr")).and_then(|body| number(body.attrs.get("rot").map(String::as_str))).map(|rotation| rotation / 60_000.0) }),
        grid_line: Some(node.child("majorGridlines").map(|grid| read_line(grid.child("spPr").and_then(|shape| shape.child("ln")), theme).unwrap_or(Line { visible: Some(true), ..Default::default() })).unwrap_or(Line { visible: Some(false), ..Default::default() })),
        tick: tick.map(|position| AxisTick { line: Line { visible: Some(position != "none"), ..Default::default() }, length: None, position: Some(if position == "in" { "inside" } else { "outside" }.into()) }),
        format: node.child("numFmt").and_then(|format| format.attrs.get("formatCode")).cloned(),
    };
    axis.validate_intervals(node.name == "dateAx")?;
    Ok(axis)
}

fn plain_text(node: &Node) -> String {
    let runs = node.descendants("t");
    if runs.is_empty() { node.descendants("v").iter().map(|value| value.text.as_str()).collect::<Vec<_>>().join("") }
    else { runs.iter().map(|value| value.text.as_str()).collect::<Vec<_>>().join("") }
}

fn formula(node: Option<&Node>) -> Option<String> {
    node?.descendants("f").into_iter().next().map(|node| node.text.trim().to_owned()).filter(|formula| !formula.is_empty())
}

fn plot_kind(plot: &Node) -> Option<String> { model_chart_kind(&plot.name, plot.value("grouping"), plot.value("barDir"), plot.value("ofPieType")).map(str::to_owned) }

fn series_options(series: &Node, plot: &Node, theme: &HashMap<String, String>, secondary: bool) -> Result<Series, String> {
    let shape = series.child("spPr");
    let fill = paint::read(shape, theme)?;
    let border = shape.and_then(|shape| read_line(shape.child("ln"), theme));
    let line = matches!(plot.name.as_str(), "lineChart" | "line3DChart" | "scatterChart" | "stockChart");
    let stroke = paint::read(shape.and_then(|shape| shape.child("ln")), theme)?;
    let effective = if line && stroke.color.is_some() { stroke } else { fill };
    let kind = plot_kind(plot).and_then(|kind| match kind.as_str() { "bar" => Some("column"), "stacked-column" => Some("columnStacked"), "line" => Some("line"), "area" => Some("area"), "areaStacked" => Some("areaStacked"), _ => None }).map(str::to_owned);
    let mut options = Series { name: series.child("tx").map(plain_text).filter(|name| !name.is_empty()), kind,
        color: effective.color, fill_opacity: effective.opacity, fill_pattern: effective.pattern.map(Some), border, point: read_marker(series.child("marker"), theme),
        right_axis: Some(secondary), label: read_label(series.child("dLbls").or_else(|| plot.child("dLbls")), theme), ..Default::default() };
    for point in series.children_named("dPt") {
        let Some(index) = point.value("idx").and_then(|index| index.parse::<usize>().ok()) else { continue; };
        let mut fill = paint::read(point.child("spPr"), theme)?;
        if fill.color.is_none() { fill = paint::read(point.child("marker").and_then(|marker| marker.child("spPr")), theme)?; }
        let marker = read_marker(point.child("marker"), theme);
        options.data_points.insert(index, Point { color: fill.color.or_else(|| marker.as_ref().and_then(|marker| marker.color.clone()).map(config::ChartColor::from)), fill_opacity: fill.opacity, fill_pattern: fill.pattern.map(Some),
            size: marker.as_ref().and_then(|marker| marker.size), shape: marker.and_then(|marker| marker.shape), border: point.child("spPr").and_then(|shape| read_line(shape.child("ln"), theme)), ..Default::default() });
    }
    Ok(options)
}

fn series_payload(index: usize, series: &Node, plot: &Node, theme: &HashMap<String, String>, secondary: bool) -> Result<ChartSeriesPayload, String> {
    let mut trendlines = Vec::new();
    for trend in series.children_named("trendline") {
        trendlines.push(config::Trendline {
            series_index: index, kind: trend.value("trendlineType").unwrap_or("linear").into(),
            name: trend.child("name").map(|name| name.text.clone()), order: property(trend, "order").map(|order| order as usize), period: property(trend, "period").map(|period| period as usize),
            show_formula: boolean(trend.value("dispEq")), show_r2: boolean(trend.value("dispRSqr")), line: read_line(trend.child("spPr").and_then(|shape| shape.child("ln")), theme).unwrap_or_default(),
        });
    }
    let mut error_bars = Vec::new();
    for error in series.children_named("errBars") {
        error_bars.push(ChartErrorBarPayload { options: config::ErrorBar { series_index: index,
            direction: error.value("errDir").unwrap_or("y").into(), kind: error.value("errBarType").unwrap_or("both").into(), value_type: error.value("errValType").unwrap_or("fixedVal").into(),
            value: property(error, "val"), no_end_cap: boolean(error.value("noEndCap")), line: read_line(error.child("spPr").and_then(|shape| shape.child("ln")), theme), plus_index: None, minus_index: None },
            plus_formula: formula(error.child("plus")), minus_formula: formula(error.child("minus")) });
    }
    Ok(ChartSeriesPayload { name: series.child("tx").map(plain_text).filter(|name| !name.is_empty()),
        categories: formula(series.child("xVal").or_else(|| series.child("cat"))), values: formula(series.child("yVal").or_else(|| series.child("val"))).unwrap_or_default(),
        bubble_sizes: formula(series.child("bubbleSize")), native_kind: plot_kind(plot), options: Some(series_options(series, plot, theme, secondary)?), trendlines, error_bars })
}

fn global_options(root: &Node, chart: &Node, plot: &Node, theme: &HashMap<String, String>) -> Options {
    let appearance = root.child("spPr");
    let title = chart.child("title");
    let mut options = Options {
        invalid_value_strategy: chart.value("dispBlanksAs").and_then(|value| match value { "zero" => Some("zero"), "span" => Some("link"), "gap" => Some("break"), _ => None }).map(str::to_owned),
        plot_visible_only: Some(boolean(chart.value("plotVisOnly")).unwrap_or(true)),
        // External native areas with a solid fill must not acquire the editor's
        // default automatic gradient. Valid authored manifests restore defaults.
        auto_gradient_fill: Some(false),
        appearance: Some(config::Appearance { background: css_color(color(appearance.and_then(|shape| shape.child("solidFill")), theme)), border: appearance.and_then(|shape| read_line(shape.child("ln"), theme)), font: read_font(root.child("txPr"), theme) }),
        title_style: read_font(title, theme), ..Default::default()
    };
    if let Some(view) = chart.child("view3D") {
        options.three_d = Some(config::ThreeD { view: Some(config::ThreeDView {
            rot_x: property(view, "rotX"), rot_y: property(view, "rotY"), perspective: property(view, "perspective"), right_angle_axes: boolean(view.value("rAngAx")),
            height_percent: property(view, "hPercent"), depth_percent: property(view, "depthPercent") }),
            gap_depth: property(plot, "gapDepth"), grouping: plot.value("grouping").filter(|group| matches!(*group, "standard" | "clustered")).map(str::to_owned) });
    }
    if plot.name.starts_with("bar") { options.bar = Some(config::Bar { gap_width: property(plot, "gapWidth"), overlap: property(plot, "overlap") }); }
    if matches!(plot.name.as_str(), "pieChart" | "pie3DChart" | "doughnutChart" | "ofPieChart") {
        let split = plot.value("splitType").and_then(|kind| match kind { "auto" => Some("auto"), "pos" => Some("position"), "val" => Some("value"), "percent" => Some("percentage"), "cust" => Some("custom"), _ => None });
        options.pie = Some(config::Pie {
            start_angle: property(plot, "firstSliceAng").map(|angle| angle - 90.0), doughnut_hole: property(plot, "holeSize").map(|size| size / 100.0),
            explosion: plot.child("ser").and_then(|series| property(series, "explosion")), label: read_label(plot.child("dLbls"), theme),
            show_label_line: plot.child("dLbls").and_then(|labels| boolean(labels.value("showLeaderLines"))),
            secondary_plot: (plot.name == "ofPieChart").then(|| config::SecondaryPie { kind: plot.value("ofPieType").map(str::to_owned), size: property(plot, "secondPieSize"), gap_width: property(plot, "gapWidth"),
                split: split.map(|kind| config::PieSplit { kind: kind.into(), count: (kind == "position").then(|| property(plot, "splitPos").unwrap_or(3.0) as usize), less_than: matches!(kind, "value" | "percentage").then(|| property(plot, "splitPos").unwrap_or(0.0)),
                    point_indexes: plot.child("custSplit").map(|split| split.children_named("secondPiePt").filter_map(|point| point.attrs.get("val").and_then(|value| value.parse().ok())).collect()) }) }), ..Default::default()
        });
    }
    if plot.name == "bubbleChart" { options.bubble = Some(config::Bubble { scale: property(plot, "bubbleScale").map(|scale| scale / 100.0), size_represents: plot.value("sizeRepresents").map(|value| if value == "w" { "width" } else { "area" }.into()) }); }
    if plot.name == "radarChart" { options.radar = Some(config::Radar { shape: Some("polygon".into()), fill: boolean(plot.value("radarStyle").map(|value| if value == "filled" { "1" } else { "0" })) }); }
    if plot.name == "stockChart" {
        let bars = plot.child("upDownBars");
        options.candlestick = Some(config::Candlestick {
            rising_color: css_color(color(bars.and_then(|bars| bars.child("upBars")).and_then(|bars| bars.child("spPr")).and_then(|shape| shape.child("solidFill")), theme)),
            falling_color: css_color(color(bars.and_then(|bars| bars.child("downBars")).and_then(|bars| bars.child("spPr")).and_then(|shape| shape.child("solidFill")), theme)),
            candle_width: bars.and_then(|bars| property(bars, "gapWidth")).map(|gap| 100.0 / (1.0 + gap / 100.0)), hollow: None });
    }
    options
}

pub(crate) fn read(index: usize, source: &str, theme: &HashMap<String, String>) -> Result<NativeChart, String> {
    let root = chart_xml::parse(source)?;
    if root.name != "chartSpace" || root.namespace.as_deref() != Some(chart_xml::CHART_NS) { return Err("unsupported native chart namespace".into()); }
    let chart = root.child("chart").ok_or_else(|| "native chart is missing chart content".to_owned())?;
    let plot_area = chart.child("plotArea").ok_or_else(|| "native chart is missing its plot area".to_owned())?;
    let plots = plot_area.children.iter().filter(|node| node.namespace.as_deref() == Some(chart_xml::CHART_NS) && NATIVE_CHART_PLOTS.contains(&node.name.as_str())).collect::<Vec<_>>();
    let first = *plots.first().ok_or_else(|| "native chart plot requires its chartEx or custom adapter".to_owned())?;
    let axes = plot_area.children.iter().filter(|node| node.namespace.as_deref() == Some(chart_xml::CHART_NS) && matches!(node.name.as_str(), "catAx" | "dateAx" | "valAx" | "serAx")).collect::<Vec<_>>();
    let axis = |id: &str| axes.iter().copied().find(|axis| axis.value("axId") == Some(id));
    let plot_axes = |plot: &Node| {
        let ids = plot.children_named("axId").filter_map(|id| id.attrs.get("val")).collect::<Vec<_>>();
        let x = ids.first().and_then(|id| axis(id));
        let y = ids.get(1).and_then(|id| axis(id));
        (x, y)
    };
    let horizontal = first.value("barDir") == Some("bar");
    let primary = plots.iter().copied().find(|plot| plot_axes(plot).1.is_some_and(|axis| axis.value("axPos") == Some(if horizontal { "b" } else { "l" }))).unwrap_or(first);
    let (x, y) = plot_axes(primary);
    let percentage = primary.value("grouping") == Some("percentStacked");
    let mut options = global_options(&root, chart, primary, theme);
    options.plot_area = paint::plot_area(plot_area)?;
    if x.is_some_and(|axis| axis.name == "dateAx") {
        options.mapping = Some(config::Mapping { use_date_axis: Some(true), ..Default::default() });
    }
    if let Some(depth) = axes.iter().find(|axis| axis.name == "serAx") { options.depth_axis = Some(config::DepthAxis::from_axis(read_axis(Some(depth), theme, false)?)); }
    let mut raw_series = Vec::new();
    for plot in &plots {
        let (_, plot_y) = plot_axes(plot);
        let secondary = plot_y.and_then(|axis| axis.value("axId")) != y.and_then(|axis| axis.value("axId"));
        if secondary { options.right_y_axis = Some(read_axis(plot_y, theme, plot.value("grouping") == Some("percentStacked"))?); }
        for series in plot.children_named("ser") { raw_series.push((property(series, "order").unwrap_or(raw_series.len() as f64), series, *plot, secondary)); }
    }
    raw_series.sort_by(|left, right| left.0.total_cmp(&right.0));
    let series = raw_series.iter().enumerate().map(|(position, (_, series, plot, secondary))| series_payload(position, series, plot, theme, *secondary)).collect::<Result<Vec<_>, String>>()?;
    let legend = chart.child("legend");
    let title = chart.child("title");
    let title_rows = title.map(|title| title.descendants("p").into_iter().map(plain_text).filter(|text| !text.is_empty()).collect::<Vec<_>>()).unwrap_or_default();
    let mut unsupported_reasons = paint::source_differences(&root, options.plot_area.is_some());
    if series.iter().any(|series| series.values.is_empty() || series.categories.is_none()) { unsupported_reasons.push("CHART_LITERAL_OR_UNBOUND_SERIES".into()); }
    let authored = root.descendants("source").into_iter().find(|node| node.namespace.as_deref() == Some(config::SOURCE_EXTENSION_NS))
        .and_then(|node| serde_json::from_str::<serde_json::Value>(&node.text).ok()).and_then(|manifest| {
            if manifest.get("version").and_then(serde_json::Value::as_u64) != Some(1) || manifest.get("nativeSignature").and_then(serde_json::Value::as_str) != Some(chart_xml::signature(&root).as_str()) {
                unsupported_reasons.push("CHART_AUTHORED_SOURCE_STALE".into()); None
            } else { manifest.get("authored").cloned() }
        });
    let kinds = plots.iter().filter_map(|plot| plot_kind(plot)).collect::<std::collections::BTreeSet<_>>();
    let kind = if kinds.len() == 1 { kinds.into_iter().next().unwrap_or_else(|| "unknown".into()) }
        else if plots.len() > 1 { "combination".into() } else { "unknown".into() };
    let payload = ChartPayload { index, kind, title: title_rows.first().cloned().or_else(|| title.map(plain_text).filter(|text| !text.is_empty())), subtitle: (title_rows.len() > 1).then(|| title_rows[1..].join("\n")),
        legend: legend.is_some(), legend_position: legend.and_then(|legend| legend.value("legendPos")).and_then(|position| match position { "t" => Some("top"), "b" => Some("bottom"), "l" => Some("left"), "r" => Some("right"), _ => None }).map(str::to_owned),
        legend_font: read_font(legend.and_then(|legend| legend.child("txPr")), theme), data_labels: series.iter().any(|series| series.options.as_ref().and_then(|series| series.label.as_ref()).is_some_and(|label| label.visible == Some(true))),
        x_axis_title: x.and_then(|axis| axis.child("title")).map(plain_text).filter(|text| !text.is_empty()), y_axis_title: y.and_then(|axis| axis.child("title")).map(plain_text).filter(|text| !text.is_empty()),
        x_axis: Some(read_axis(x, theme, false)?), y_axis: Some(read_axis(y, theme, percentage)?), series, options: Some(options), authored, unsupported_reasons, office_source: Some(source.into()),
        from: AnchorCellPayload::default(), to: AnchorCellPayload::default(), drawing: None, pivot_table_id: None };
    let pivot_source_name = root.child("pivotSource").and_then(|pivot| pivot.child("name")).map(|name| name.text.clone());
    Ok(NativeChart { payload, pivot_source_name })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn an_external_date_axis_restores_date_mapping_and_font_without_inventing_title_position() {
        let source = format!(r#"<c:chartSpace xmlns:c="{}" xmlns:a="{}"><c:chart><c:plotArea><c:lineChart><c:grouping val="standard"/><c:ser><c:cat><c:numRef><c:f>'Data'!$A$2:$A$3</c:f></c:numRef></c:cat><c:val><c:numRef><c:f>'Data'!$B$2:$B$3</c:f></c:numRef></c:val></c:ser><c:axId val="10"/><c:axId val="11"/></c:lineChart><c:dateAx><c:axId val="10"/><c:scaling/><c:axPos val="b"/><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr algn="r"/><a:r><a:rPr sz="1200"><a:latin typeface="Arial"/></a:rPr><a:t>Dates</a:t></a:r></a:p></c:rich></c:tx></c:title><c:numFmt formatCode="yyyy-mm-dd" sourceLinked="0"/><c:crossAx val="11"/></c:dateAx><c:valAx><c:axId val="11"/><c:scaling/><c:axPos val="l"/><c:crossAx val="10"/></c:valAx></c:plotArea></c:chart></c:chartSpace>"#, chart_xml::CHART_NS, chart_xml::DRAWING_NS);
        let payload = read(0, &source, &HashMap::new()).unwrap().payload;
        assert_eq!(payload.options.as_ref().unwrap().mapping.as_ref().unwrap().use_date_axis, Some(true));
        let axis = payload.x_axis.unwrap();
        assert_eq!(axis.title.as_deref(), Some("Dates"));
        assert_eq!(axis.title_style.as_ref().unwrap().font_size, Some(16.0));
        assert_eq!(axis.title_style.as_ref().unwrap().family.as_deref(), Some("Arial"));
        assert_eq!(axis.title_alignment, None);
    }

    #[test]
    fn luminance_transforms_keep_saturated_hues_and_apply_in_xml_order() {
        let theme = HashMap::from([("accent1".into(), "#ff0000".into())]);
        let source = format!(r#"<a:solidFill xmlns:a="{}"><a:schemeClr val="accent1"><a:lumMod val="50000"/><a:lumOff val="20000"/></a:schemeClr></a:solidFill>"#, chart_xml::DRAWING_NS);
        let root = chart_xml::parse(&source).unwrap();
        assert_eq!(color(Some(&root), &theme), Some(("#e60000".into(), 1.0)));
        let source = format!(r#"<a:solidFill xmlns:a="{}"><a:srgbClr val="808080"><a:lum val="25000"/></a:srgbClr></a:solidFill>"#, chart_xml::DRAWING_NS);
        assert_eq!(color(Some(&chart_xml::parse(&source).unwrap()), &theme), Some(("#404040".into(),1.0)));
    }

    #[test]
    fn every_transform_clamps_before_later_color_and_alpha_operations() {
        let source = format!(r#"<a:solidFill xmlns:a="{}"><a:srgbClr val="ff0000"><a:shade val="200000"/><a:lumMod val="50000"/><a:alphaOff val="200000"/><a:alphaMod val="50000"/></a:srgbClr></a:solidFill>"#, chart_xml::DRAWING_NS);
        assert_eq!(color(Some(&chart_xml::parse(&source).unwrap()), &HashMap::new()), Some(("#800000".into(), 0.5)));
    }

    #[test]
    fn horizontal_percent_stacked_axes_follow_logical_role_and_native_ids() {
        let source = format!("<c:chartSpace xmlns:c=\"{}\"><c:chart><c:plotArea><c:barChart><c:barDir val=\"bar\"/><c:grouping val=\"percentStacked\"/><c:ser><c:idx val=\"0\"/><c:order val=\"0\"/><c:cat><c:strRef><c:f>'Data'!$A$2:$A$3</c:f></c:strRef></c:cat><c:val><c:numRef><c:f>'Data'!$B$2:$B$3</c:f></c:numRef></c:val></c:ser><c:axId val=\"10\"/><c:axId val=\"11\"/></c:barChart><c:catAx><c:axId val=\"10\"/><c:scaling><c:orientation val=\"maxMin\"/></c:scaling><c:axPos val=\"l\"/></c:catAx><c:valAx><c:axId val=\"11\"/><c:scaling><c:max val=\"1\"/><c:min val=\"0\"/></c:scaling><c:axPos val=\"b\"/></c:valAx></c:plotArea></c:chart></c:chartSpace>", chart_xml::CHART_NS);
        let chart = read(0, &source, &HashMap::new()).unwrap().payload;
        assert_eq!(chart.kind, "barPercentStacked");
        assert_eq!(chart.x_axis.as_ref().unwrap().reverse, Some(true));
        assert_eq!(chart.y_axis.as_ref().unwrap().min, Some(0.0));
        assert_eq!(chart.y_axis.as_ref().unwrap().max, Some(100.0));
        assert_eq!(chart.series[0].values, "'Data'!$B$2:$B$3");
    }
}
