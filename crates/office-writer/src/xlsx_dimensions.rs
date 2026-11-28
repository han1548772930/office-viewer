use super::{SheetInput, MAX_COLUMN_INDEX, MAX_ROW_INDEX, esc, format_number, parse_index};
use crate::ExportError;
use office_io_common::{pixels_to_points, pixels_to_stored_character_width, worksheet_axis_fingerprint, worksheet_filter_fingerprint_with_colors, SHEET_AXIS_EXTENSION_URI, SHEET_AXIS_NAMESPACE};
use serde::Deserialize;
use serde_json::json;
use std::collections::{BTreeMap, BTreeSet};

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub(super) struct AxisSpanInput<T> { pub start: u32, pub end: u32, pub value: T }

#[derive(Debug, Clone, Copy, PartialEq)]
pub(super) struct RowSize { pub height: f64, pub automatic: bool }
#[derive(Debug, Clone)]
struct RowRun { start: u32, end: u32, size: RowSize }
#[derive(Debug, Clone, Copy, PartialEq)]
struct ColumnSize { width: f64, hidden: bool, explicit: bool }
#[derive(Debug, Clone)]
struct ColumnRun { start: u32, end: u32, size: ColumnSize }

pub(super) struct WorksheetAxisPlan {
    pub default_row_height: f64,
    pub default_row_automatic: bool,
    pub default_column_width: f64,
    rows: Vec<RowRun>,
    extra_rows: Vec<RowRun>,
    columns: Vec<ColumnRun>,
    manual_hidden_rows: Vec<AxisSpanInput<bool>>,
}

fn span_at<T: Copy>(spans: &[AxisSpanInput<T>], index: u32) -> Option<T> {
    let position = spans.partition_point(|span| span.start <= index);
    spans.get(position.wrapping_sub(1)).filter(|span| index <= span.end).map(|span| span.value)
}
fn check_spans<T>(spans: &[AxisSpanInput<T>], maximum: u32, valid: impl Fn(&T) -> bool) -> Result<(), ExportError> {
    let mut previous = None;
    for span in spans {
        if span.start > span.end || span.end > maximum || previous.is_some_and(|end| span.start <= end) || !valid(&span.value) {
            return Err(ExportError::invalid("worksheet axis spans must be finite, ordered, disjoint and within the XLSX limit"));
        }
        previous = Some(span.end);
    }
    Ok(())
}
fn points(values: &BTreeMap<String, f64>, maximum: u32) -> Result<BTreeMap<u32, f64>, ExportError> {
    let mut result = BTreeMap::new();
    for (key, value) in values {
        let index = parse_index(key).filter(|index| *index <= maximum).ok_or_else(|| ExportError::invalid("worksheet axis index is invalid"))?;
        if !value.is_finite() || *value < 0.0 { return Err(ExportError::invalid("worksheet axis sizes must be finite and non-negative")); }
        result.insert(index, *value);
    }
    Ok(result)
}
fn boundaries<T>(spans: &[AxisSpanInput<T>], positions: impl IntoIterator<Item = u32>, maximum: u32) -> BTreeSet<u32> {
    let mut result = BTreeSet::from([0, maximum + 1]);
    for span in spans { result.insert(span.start); result.insert(span.end + 1); }
    for position in positions { result.insert(position); result.insert(position + 1); }
    result
}
fn numeric_spans(spans: &[AxisSpanInput<f64>], values: &BTreeMap<u32, f64>, maximum: u32) -> Vec<AxisSpanInput<f64>> {
    let positions = boundaries(spans, values.keys().copied(), maximum).into_iter().collect::<Vec<_>>();
    let mut result: Vec<AxisSpanInput<f64>> = Vec::new();
    for pair in positions.windows(2) {
        let Some(value) = values.get(&pair[0]).copied().or_else(|| span_at(spans, pair[0])) else { continue; };
        if let Some(previous) = result.last_mut() {
            if previous.end + 1 == pair[0] && previous.value == value { previous.end = pair[1] - 1; continue; }
        }
        result.push(AxisSpanInput { start: pair[0], end: pair[1] - 1, value });
    }
    result
}

pub(super) fn prepare(sheet: &SheetInput) -> Result<WorksheetAxisPlan, ExportError> {
    if sheet.axis_unit.as_deref().is_some_and(|unit| unit != "cssPixel") { return Err(ExportError::invalid("unsupported worksheet axis unit")); }
    let valid_size = |value: &f64| value.is_finite() && *value >= 0.0;
    if sheet.row_count.is_some_and(|count| count == 0 || count > MAX_ROW_INDEX + 1)
        || sheet.column_count.is_some_and(|count| count == 0 || count > MAX_COLUMN_INDEX + 1)
        || sheet.hidden_rows.iter().any(|row| *row > MAX_ROW_INDEX || sheet.row_count.is_some_and(|count| *row >= count))
        || sheet.hidden_columns.iter().any(|column| *column > MAX_COLUMN_INDEX || sheet.column_count.is_some_and(|count| *column >= count)) {
        return Err(ExportError::invalid("worksheet axis counts or hidden indices exceed the worksheet or XLSX limit"));
    }
    check_spans(&sheet.row_height_spans, MAX_ROW_INDEX, valid_size)?;
    check_spans(&sheet.column_width_spans, MAX_COLUMN_INDEX, valid_size)?;
    check_spans(&sheet.row_auto_height_spans, MAX_ROW_INDEX, |_| true)?;
    check_spans(&sheet.row_hidden_spans, sheet.row_count.map(|count| count - 1).unwrap_or(MAX_ROW_INDEX), |_| true)?;
    check_spans(&sheet.column_hidden_spans, sheet.column_count.map(|count| count - 1).unwrap_or(MAX_COLUMN_INDEX), |_| true)?;
    for group in sheet.outline.iter().flat_map(|outline| &outline.groups) {
        let maximum = match group.axis.as_str() {
            "row" => sheet.row_count.map(|count| count - 1).unwrap_or(MAX_ROW_INDEX),
            "column" => sheet.column_count.map(|count| count - 1).unwrap_or(MAX_COLUMN_INDEX),
            _ => return Err(ExportError::invalid("worksheet outline axis is invalid")),
        };
        if group.start > group.end || group.end > maximum { return Err(ExportError::invalid("worksheet outline group is outside the worksheet")); }
    }
    let row_points = points(&sheet.row_heights, MAX_ROW_INDEX)?;
    let column_points = points(&sheet.column_widths, MAX_COLUMN_INDEX)?;
    let auto_points = points(&sheet.row_auto_heights, MAX_ROW_INDEX)?;
    let source_default_row = sheet.default_row_height.unwrap_or(26.0);
    let default_column_width = sheet.default_column_width.unwrap_or(112.0);
    if !valid_size(&source_default_row) || !valid_size(&default_column_width) { return Err(ExportError::invalid("worksheet default dimensions must be finite and non-negative")); }
    let manual_rows = numeric_spans(&sheet.row_height_spans, &row_points, MAX_ROW_INDEX);
    let manual_columns = numeric_spans(&sheet.column_width_spans, &column_points, MAX_COLUMN_INDEX);
    let mut row_boundaries = boundaries(&manual_rows, auto_points.keys().copied(), MAX_ROW_INDEX);
    for mode in &sheet.row_auto_height_spans { row_boundaries.insert(mode.start); row_boundaries.insert(mode.end + 1); }
    let positions = row_boundaries.into_iter().collect::<Vec<_>>();
    let mut rows: Vec<RowRun> = Vec::new();
    for pair in positions.windows(2) {
        let manual = span_at(&manual_rows, pair[0]);
        let automatic = span_at(&sheet.row_auto_height_spans, pair[0]).unwrap_or(sheet.default_row_auto_height != Some(false) && manual.is_none());
        let height = if automatic { auto_points.get(&pair[0]).copied() } else { manual }.unwrap_or(source_default_row);
        let size = RowSize { height, automatic };
        if let Some(previous) = rows.last_mut() {
            if previous.end + 1 == pair[0] && previous.size == size { previous.end = pair[1] - 1; continue; }
        }
        rows.push(RowRun { start: pair[0], end: pair[1] - 1, size });
    }
    // Choose the most prevalent geometry/mode pair for the standard default.
    // A full manual axis with a few automatic exceptions still needs only those
    // exceptions, while the extension retains the original declared defaults.
    let source_default_automatic = sheet.default_row_auto_height != Some(false);
    let mut coverage = BTreeMap::<(u64, bool), (f64, u64)>::new();
    for run in &rows {
        let height = if run.size.height == 0.0 { 0.0 } else { run.size.height };
        let entry = coverage.entry((height.to_bits(), run.size.automatic)).or_insert((height, 0));
        entry.1 += u64::from(run.end - run.start) + 1;
    }
    let ((_, default_row_automatic), (default_row_height, _)) = coverage.iter().max_by_key(|((_, automatic), (height, count))|
        (*count, *height == source_default_row && *automatic == source_default_automatic)).unwrap();
    let (default_row_height, default_row_automatic) = (*default_row_height, *default_row_automatic);
    let extra_rows = rows.iter().filter(|run| run.size.height != default_row_height || run.size.automatic != default_row_automatic).cloned().collect();
    let hidden_columns = sheet.hidden_columns.iter().copied().collect::<BTreeSet<_>>();
    let mut column_boundaries = boundaries(&manual_columns, hidden_columns.iter().copied(), MAX_COLUMN_INDEX);
    for span in &sheet.column_hidden_spans { column_boundaries.insert(span.start); column_boundaries.insert(span.end + 1); }
    if sheet.default_column_width.is_none() { column_boundaries.insert(1); }
    let positions = column_boundaries.into_iter().collect::<Vec<_>>();
    let mut columns: Vec<ColumnRun> = Vec::new();
    for pair in positions.windows(2) {
        let explicit = span_at(&manual_columns, pair[0]);
        let width = explicit.unwrap_or_else(|| if sheet.default_column_width.is_none() && pair[0] == 0 { 152.0 } else { default_column_width });
        let size = ColumnSize { width, hidden: hidden_columns.contains(&pair[0]) || span_at(&sheet.column_hidden_spans, pair[0]).unwrap_or(false) || width == 0.0, explicit: explicit.is_some() };
        if width == default_column_width && !size.hidden && !size.explicit { continue; }
        if let Some(previous) = columns.last_mut() {
            if previous.end + 1 == pair[0] && previous.size == size { previous.end = pair[1] - 1; continue; }
        }
        columns.push(ColumnRun { start: pair[0], end: pair[1] - 1, size });
    }
    Ok(WorksheetAxisPlan { default_row_height, default_row_automatic, default_column_width, rows, extra_rows, columns,
        manual_hidden_rows: sheet.row_hidden_spans.clone() })
}

impl WorksheetAxisPlan {
    pub fn row_size(&self, row: u32) -> RowSize {
        let position = self.rows.partition_point(|run| run.start <= row);
        self.rows[position - 1].size
    }
    pub fn extra_row_indices(&self) -> impl Iterator<Item = u32> + '_ { self.extra_rows.iter().flat_map(|run| run.start..=run.end) }
    pub fn manual_hidden_row_indices(&self) -> impl Iterator<Item = u32> + '_ { self.manual_hidden_rows.iter().filter(|span| span.value).flat_map(|span| span.start..=span.end) }
    pub fn manual_row_hidden(&self, row: u32) -> bool { span_at(&self.manual_hidden_rows, row).unwrap_or(false) }
    pub fn column_hidden(&self, column: u32) -> bool {
        let position = self.columns.partition_point(|run| run.start <= column);
        self.columns.get(position.wrapping_sub(1)).is_some_and(|run| column <= run.end && run.size.hidden)
    }
    pub fn row_attributes(&self, row: u32) -> String {
        let size = self.row_size(row);
        let position = self.extra_rows.partition_point(|run| run.start <= row);
        let explicit = self.extra_rows.get(position.wrapping_sub(1)).is_some_and(|run| row <= run.end);
        if !explicit { return String::new(); }
        format!(" ht=\"{}\" customHeight=\"{}\"", format_number(pixels_to_points(size.height)), if size.automatic { 0 } else { 1 })
    }
    pub fn format_xml(&self) -> String {
        format!("<sheetFormatPr defaultColWidth=\"{}\" defaultRowHeight=\"{}\" customHeight=\"{}\"{}/>",
            format_number(pixels_to_stored_character_width(self.default_column_width)), format_number(pixels_to_points(self.default_row_height)),
            if self.default_row_automatic { 0 } else { 1 }, if self.default_row_height == 0.0 { " zeroHeight=\"1\"" } else { "" })
    }
    pub fn columns_xml(&self) -> String {
        if self.columns.is_empty() { return String::new(); }
        let mut output = "<cols>".to_owned();
        for run in &self.columns {
            output.push_str(&format!("<col min=\"{}\" max=\"{}\" width=\"{}\" customWidth=\"1\"{} />", run.start + 1, run.end + 1,
                format_number(pixels_to_stored_character_width(run.size.width)), if run.size.hidden { " hidden=\"1\"" } else { "" }));
        }
        output.push_str("</cols>"); output
    }
    pub fn extension_xml(&self, sheet: &SheetInput, worksheet: &str, registry: &super::StyleRegistry) -> Result<String, ExportError> {
        let fingerprint = worksheet_axis_fingerprint(worksheet).map_err(ExportError::invalid)?;
        let filter_fingerprint = worksheet_filter_fingerprint_with_colors(worksheet, |node| super::auto_filter::color_signature(node, registry)).map_err(ExportError::invalid)?;
        let numeric = |spans: &[AxisSpanInput<f64>]| spans.iter().map(|span| json!({ "start": span.start, "end": span.end, "value": span.value })).collect::<Vec<_>>();
        let points = |values: &BTreeMap<String, f64>| values.iter().map(|(key, size)| json!({ "index": parse_index(key).unwrap(), "size": size })).collect::<Vec<_>>();
        let state = json!({
            "version": 1, "axisUnit": "cssPixel", "rowHeights": points(&sheet.row_heights), "columnWidths": points(&sheet.column_widths),
            "rowHeightSpans": numeric(&sheet.row_height_spans), "columnWidthSpans": numeric(&sheet.column_width_spans),
            "rowAutoHeightSpans": sheet.row_auto_height_spans.iter().map(|span| json!({ "start": span.start, "end": span.end, "value": span.value })).collect::<Vec<_>>(),
            "rowAutoHeights": points(&sheet.row_auto_heights),
            "defaultRowHeight": sheet.default_row_height, "defaultColumnWidth": sheet.default_column_width,
            "defaultRowAutoHeight": sheet.default_row_auto_height, "rowCount": sheet.row_count, "columnCount": sheet.column_count,
            "hiddenRows": sheet.hidden_rows, "hiddenColumns": sheet.hidden_columns,
            "rowHiddenSpans": sheet.row_hidden_spans.iter().map(|span| json!({ "start": span.start, "end": span.end, "value": span.value })).collect::<Vec<_>>(),
            "columnHiddenSpans": sheet.column_hidden_spans.iter().map(|span| json!({ "start": span.start, "end": span.end, "value": span.value })).collect::<Vec<_>>(),
            "autoFilter": super::auto_filter::snapshot_value(sheet)?, "filterFingerprint": filter_fingerprint,
        });
        let state = state.to_string();
        if state.len() > 32 * 1024 * 1024 { return Err(ExportError::invalid("worksheet pixel/filter snapshot exceeds the parser's 32 MiB budget")); }
        Ok(format!("<ext uri=\"{}\"><ov:axisDimensions xmlns:ov=\"{}\" version=\"1\" fingerprint=\"{}\">{}</ov:axisDimensions></ext>",
            SHEET_AXIS_EXTENSION_URI, SHEET_AXIS_NAMESPACE, fingerprint, esc(&state)))
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    fn sheet(value: serde_json::Value) -> SheetInput { serde_json::from_value(value).unwrap() }

    #[test]
    fn whole_axis_manual_height_is_compact_and_exceptions_keep_their_office_modes() {
        let value = sheet(json!({ "rowHeightSpans": [{ "start": 0, "end": MAX_ROW_INDEX, "value": 32 }], "rowAutoHeightSpans": [{ "start": 0, "end": MAX_ROW_INDEX, "value": false }], "rowHeights": { "9": 40 } }));
        let plan = prepare(&value).unwrap();
        assert_eq!(plan.default_row_height, 32.0);
        assert!(!plan.default_row_automatic);
        assert_eq!(plan.extra_row_indices().collect::<Vec<_>>(), [9]);
        assert_eq!(plan.row_attributes(9), " ht=\"30\" customHeight=\"1\"");
        assert!(plan.format_xml().contains("defaultRowHeight=\"24\" customHeight=\"1\""));
    }

    #[test]
    fn zero_width_and_auto_height_are_distinct_from_manual_overrides() {
        let value = sheet(json!({ "defaultColumnWidth": 112, "columnWidthSpans": [{ "start": 1, "end": 9, "value": 0 }],
            "rowAutoHeightSpans": [{ "start": 4, "end": 4, "value": true }], "rowAutoHeights": { "4": 48 }, "rowHeights": { "5": 48 } }));
        let plan = prepare(&value).unwrap();
        assert!(plan.columns_xml().contains("min=\"2\" max=\"10\" width=\"0\" customWidth=\"1\" hidden=\"1\""));
        assert_eq!(plan.row_attributes(4), " ht=\"36\" customHeight=\"0\"");
        assert_eq!(plan.row_attributes(5), " ht=\"36\" customHeight=\"1\"");
    }

    #[test]
    fn mostly_manual_axis_with_automatic_exceptions_uses_a_compact_standard_projection() {
        let value = sheet(json!({ "rowHeightSpans": [{ "start": 0, "end": MAX_ROW_INDEX, "value": 32 }],
            "rowAutoHeightSpans": [{ "start": 9, "end": 9, "value": true }], "rowAutoHeights": { "9": 48 } }));
        let plan = prepare(&value).unwrap();
        assert_eq!(plan.default_row_height, 32.0);
        assert!(!plan.default_row_automatic);
        assert_eq!(plan.extra_row_indices().collect::<Vec<_>>(), [9]);
        assert_eq!(plan.row_attributes(9), " ht=\"36\" customHeight=\"0\"");
    }

    #[test]
    fn rejects_invalid_indices_negative_sizes_and_overlapping_spans() {
        for value in [json!({ "rowHeights": { "bad": 30 } }), json!({ "columnWidths": { "0": -1 } }),
            json!({ "rowHeightSpans": [{ "start": 0, "end": 2, "value": 20 }, { "start": 2, "end": 3, "value": 30 }] })] {
            assert!(prepare(&sheet(value)).is_err());
        }
    }
}
