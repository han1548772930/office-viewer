//! Authored chart settings shared by the native Office reader and writer.
//! These DTOs contain source-field identities, never derived pixels or samples.

use std::collections::BTreeMap;
use serde::{Deserialize, Serialize};
use serde_json::{Map, Value};

#[derive(Debug, Default, Clone, Deserialize, Serialize)]
#[serde(default, rename_all = "camelCase")]
pub struct Font {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_size: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bold: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub italic: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub family: Option<String>,
}

#[derive(Debug, Default, Clone, Deserialize, Serialize)]
#[serde(default, rename_all = "camelCase")]
pub struct Line {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub visible: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub width: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub opacity: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub dash_type: Option<String>,
}

#[derive(Debug, Default, Clone, Deserialize, Serialize)]
#[serde(default, rename_all = "camelCase")]
pub struct Label {
    #[serde(flatten)]
    pub font: Font,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub visible: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub position: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub content_type: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub format: Option<String>,
}

#[derive(Debug, Default, Clone, Deserialize, Serialize)]
#[serde(default, rename_all = "camelCase")]
pub struct Marker {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub shape: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub size: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
}

#[derive(Debug, Default, Clone, Deserialize, Serialize)]
#[serde(default, rename_all = "camelCase")]
pub struct Point {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fill_opacity: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub size: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub shape: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub border: Option<Line>,
    #[serde(flatten)]
    pub extensions: Map<String, Value>,
}

#[derive(Debug, Default, Clone, Deserialize, Serialize)]
#[serde(default, rename_all = "camelCase")]
pub struct Series {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    #[serde(rename = "type", skip_serializing_if = "Option::is_none")]
    pub kind: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fill_opacity: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub border: Option<Line>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub point: Option<Marker>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub right_axis: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub label: Option<Label>,
    #[serde(skip_serializing_if = "BTreeMap::is_empty")]
    pub data_points: BTreeMap<usize, Point>,
    #[serde(flatten)]
    pub extensions: Map<String, Value>,
}

/// Names identify the original source header or bubble group. A display alias
/// in `style.name` does not change this selector, including empty/space names.
#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct NamedSeries {
    pub name: String,
    pub style: Series,
}

#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(default, rename_all = "camelCase")]
pub struct Axis {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub title: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub title_style: Option<Font>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub title_alignment: Option<String>,
    pub scale_type: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub log_base: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub min: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub max: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub reverse: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub line: Option<Line>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub label: Option<AxisLabel>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub grid_line: Option<Line>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tick: Option<AxisTick>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub format: Option<String>,
}

impl Default for Axis {
    fn default() -> Self {
        Self { title: None, title_style: None, title_alignment: None, scale_type: "linear".into(), log_base: None, min: None, max: None,
            reverse: None, line: None, label: None, grid_line: None, tick: None, format: None }
    }
}

#[derive(Debug, Default, Clone, Deserialize, Serialize)]
#[serde(default, rename_all = "camelCase")]
pub struct DepthAxis {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub title: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub title_style: Option<Font>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub title_alignment: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub reverse: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub line: Option<Line>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub label: Option<AxisLabel>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub grid_line: Option<Line>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tick: Option<AxisTick>,
}

impl DepthAxis {
    pub fn to_axis(&self) -> Axis {
        Axis { title: self.title.clone(), title_style: self.title_style.clone(), title_alignment: self.title_alignment.clone(), reverse: self.reverse, line: self.line.clone(), label: self.label.clone(),
            grid_line: self.grid_line.clone(), tick: self.tick.clone(), ..Axis::default() }
    }
    pub fn from_axis(axis: Axis) -> Self {
        Self { title: axis.title, title_style: axis.title_style, title_alignment: axis.title_alignment, reverse: axis.reverse, line: axis.line, label: axis.label, grid_line: axis.grid_line, tick: axis.tick }
    }
}

#[derive(Debug, Default, Clone, Deserialize, Serialize)]
#[serde(default, rename_all = "camelCase")]
pub struct AxisLabel {
    #[serde(flatten)]
    pub font: Font,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub visible: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rotate: Option<f64>,
}

#[derive(Debug, Default, Clone, Deserialize, Serialize)]
#[serde(default, rename_all = "camelCase")]
pub struct AxisTick {
    #[serde(flatten)]
    pub line: Line,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub length: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub position: Option<String>,
}

#[derive(Debug, Default, Clone, Deserialize, Serialize)]
#[serde(default, rename_all = "camelCase")]
pub struct Trendline {
    pub series_index: usize,
    #[serde(rename = "type")]
    pub kind: String,
    #[serde(flatten)]
    pub line: Line,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub order: Option<usize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub period: Option<usize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub show_formula: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub show_r2: Option<bool>,
}

#[derive(Debug, Default, Clone, Deserialize, Serialize)]
#[serde(default, rename_all = "camelCase")]
pub struct ErrorBar {
    pub series_index: usize,
    pub direction: String,
    #[serde(rename = "type")]
    pub kind: String,
    pub value_type: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub value: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub plus_index: Option<usize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub minus_index: Option<usize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub no_end_cap: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub line: Option<Line>,
}

#[derive(Debug, Default, Clone, Deserialize, Serialize)]
#[serde(default, rename_all = "camelCase")]
pub struct Appearance {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub background: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub border: Option<Line>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font: Option<Font>,
}

#[derive(Debug, Default, Clone, Deserialize, Serialize)]
#[serde(default, rename_all = "camelCase")]
pub struct Mapping {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub use_date_axis: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub category_indexes: Option<Vec<usize>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub hierarchy_indexes: Option<Vec<usize>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bubble: Option<BubbleMapping>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub candlestick: Option<CandlestickMapping>,
    #[serde(flatten)]
    pub extensions: Map<String, Value>,
}

#[derive(Debug, Default, Clone, Deserialize, Serialize)]
#[serde(default, rename_all = "camelCase")]
pub struct BubbleMapping {
    pub x_index: usize,
    pub y_index: usize,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub size_index: Option<usize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub label_index: Option<usize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub series_index: Option<usize>,
}

#[derive(Debug, Default, Clone, Deserialize, Serialize)]
#[serde(default, rename_all = "camelCase")]
pub struct CandlestickMapping {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub open_index: Option<usize>,
    pub high_index: usize,
    pub low_index: usize,
    pub close_index: usize,
}

#[derive(Debug, Default, Clone, Deserialize, Serialize)]
#[serde(default, rename_all = "camelCase")]
pub struct ThreeD {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub view: Option<ThreeDView>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub grouping: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub gap_depth: Option<f64>,
}

#[derive(Debug, Default, Clone, Deserialize, Serialize)]
#[serde(default, rename_all = "camelCase")]
pub struct ThreeDView {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rot_x: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rot_y: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub perspective: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub right_angle_axes: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub height_percent: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub depth_percent: Option<f64>,
}

#[derive(Debug, Default, Clone, Deserialize, Serialize)]
#[serde(default, rename_all = "camelCase")]
pub struct Bar {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub gap_width: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub overlap: Option<f64>,
}

#[derive(Debug, Default, Clone, Deserialize, Serialize)]
#[serde(default, rename_all = "camelCase")]
pub struct Pie {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub start_angle: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub doughnut_hole: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub explosion: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rose_pie: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub is_half_pie: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub show_label_line: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub label: Option<Label>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub secondary_plot: Option<SecondaryPie>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub aggregate_label: Option<String>,
}

#[derive(Debug, Default, Clone, Deserialize, Serialize)]
#[serde(default, rename_all = "camelCase")]
pub struct SecondaryPie {
    #[serde(rename = "type", skip_serializing_if = "Option::is_none")]
    pub kind: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub size: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub gap_width: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub split: Option<PieSplit>,
}

#[derive(Debug, Default, Clone, Deserialize, Serialize)]
#[serde(default, rename_all = "camelCase")]
pub struct PieSplit {
    #[serde(rename = "type")]
    pub kind: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub count: Option<usize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub less_than: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub point_indexes: Option<Vec<usize>>,
}

#[derive(Debug, Default, Clone, Deserialize, Serialize)]
#[serde(default, rename_all = "camelCase")]
pub struct Bubble {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub scale: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub size_represents: Option<String>,
}

#[derive(Debug, Default, Clone, Deserialize, Serialize)]
#[serde(default, rename_all = "camelCase")]
pub struct Radar {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub shape: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fill: Option<bool>,
}

#[derive(Debug, Default, Clone, Deserialize, Serialize)]
#[serde(default, rename_all = "camelCase")]
pub struct Area {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub smooth: Option<bool>,
}

#[derive(Debug, Default, Clone, Deserialize, Serialize)]
#[serde(default, rename_all = "camelCase")]
pub struct Candlestick {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub candle_width: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rising_color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub falling_color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub hollow: Option<bool>,
}

#[derive(Debug, Default, Clone, Deserialize, Serialize)]
#[serde(default, rename_all = "camelCase")]
pub struct AxisPointer {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub indicator_label_color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub indicator_line_type: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub indicator_line_color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub indicator_label_text_color: Option<String>,
}

#[derive(Debug, Default, Clone, Deserialize, Serialize)]
#[serde(default, rename_all = "camelCase")]
pub struct Options {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub invalid_value_strategy: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub plot_visible_only: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub appearance: Option<Appearance>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub title_style: Option<Font>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub subtitle_style: Option<Font>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub all_series_style: Option<Series>,
    #[serde(skip_serializing_if = "BTreeMap::is_empty")]
    pub series: BTreeMap<usize, Series>,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub named_series: Vec<NamedSeries>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub axis_pointer: Option<AxisPointer>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub right_y_axis: Option<Axis>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub depth_axis: Option<DepthAxis>,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub trendlines: Vec<Trendline>,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub error_bars: Vec<ErrorBar>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub mapping: Option<Mapping>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub three_d: Option<ThreeD>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bar: Option<Bar>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pie: Option<Pie>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bubble: Option<Bubble>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub radar: Option<Radar>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub area: Option<Area>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub candlestick: Option<Candlestick>,
    #[serde(flatten)]
    pub extensions: Map<String, Value>,
}

/// Source metadata supplements native formulas when helpers or aliases are used.
/// The native chart remains the rendered and editable Office object.
pub const SOURCE_EXTENSION_URI: &str = "{8E13DA70-6490-4C80-AE82-63A85BFD26D3}";
pub const SOURCE_EXTENSION_NS: &str = "urn:office-studio:chart-source:1";
