//! Restore native outlines after dimension/filter provenance has been restored.
//! Project model and manual visibility have independent staleness guards.

use std::collections::BTreeMap;
use std::io::{BufReader, Cursor, Read};

use office_io_common::sheet_outline_source::{HiddenSpan, MAX_COLUMN_INDEX, MAX_ROW_INDEX,
    MAX_WORKSHEET_BYTES, WorksheetOutlineRead, infer_native_state, read_worksheet_outline, validate_state};
use serde_json::json;
use zip::ZipArchive;

use crate::model::{AxisSpanPayload, SheetPayload};
use crate::xlsx_parts::worksheet_parts;

#[derive(Debug, Default)]
pub(crate) struct SheetOutlineSidecar {
    by_sheet: BTreeMap<String, WorksheetOutlineRead>,
    pub(crate) warnings: Vec<String>,
    pub(crate) failure: Option<String>,
}

impl SheetOutlineSidecar {
    pub(crate) fn read(bytes: &[u8]) -> Self {
        match read_sidecar(bytes) {
            Ok(value) => value,
            Err(warning) => Self { failure: Some(format!("Could not read XLSX outlines: {warning}")), ..Self::default() },
        }
    }

    pub(crate) fn restore(&mut self, sheet: &mut SheetPayload) {
        let Some(mut supplement) = self.by_sheet.remove(&sheet.name) else { return; };
        let (outline_fingerprint, visibility_fingerprint) = match supplement.source.fingerprints() {
            Ok(value) => value,
            Err(warning) => { self.failure = Some(format!("{}: {warning}", sheet.name)); return; }
        };
        let model_matched = supplement.extension.as_ref().is_some_and(|state| state.outline_fingerprint.eq_ignore_ascii_case(&outline_fingerprint));
        let visibility_matched = model_matched && supplement.extension.as_ref().is_some_and(|state| state.visibility_fingerprint.eq_ignore_ascii_case(&visibility_fingerprint));
        let axis_exact = sheet.source_axis_dimensions.pointer("/projectAxisExtension/restored").and_then(serde_json::Value::as_bool) == Some(true)
            && !(sheet.source_auto_filter.is_some() && sheet.auto_filter.is_none());
        let state = if model_matched { supplement.extension.as_ref().unwrap().state.clone() } else { infer_native_state(&supplement.source) };
        if let Err(warning) = validate_state(&state, MAX_ROW_INDEX + 1, MAX_COLUMN_INDEX + 1) {
            self.failure = Some(format!("{}: native outline model could not be normalized: {warning}", sheet.name));
        } else {
            // Empty detail rows/columns are still part of the authored outline.
            // Summary markers may live beyond the original used-cell rectangle.
            let row_extent = state.groups.iter().filter(|group| group.axis == "row").map(|group| group.end + 1)
                .chain(supplement.source.row_runs.iter().filter(|run| run.level > 0 || run.collapsed).map(|run| run.end + 1)).max();
            let column_extent = state.groups.iter().filter(|group| group.axis == "column").map(|group| group.end + 1)
                .chain(supplement.source.column_runs.iter().filter(|run| run.level > 0 || run.collapsed).map(|run| run.end + 1)).max();
            if let Some(count) = row_extent { sheet.row_count = Some(sheet.row_count.unwrap_or(0).max(count)); }
            if let Some(count) = column_extent { sheet.column_count = Some(sheet.column_count.unwrap_or(0).max(count)); }
            if visibility_matched {
                let manual = &supplement.extension.as_ref().unwrap().manual_visibility;
                sheet.hidden_rows = manual.hidden_rows.clone();
                sheet.hidden_columns = manual.hidden_columns.iter().map(|column| *column as u16).collect();
                sheet.row_hidden_spans = manual.row_hidden_spans.iter().map(payload_span).collect();
                sheet.column_hidden_spans = manual.column_hidden_spans.iter().map(payload_span).collect();
            } else if !axis_exact {
                // Standard `hidden` is a composite bit. Expand must reveal the
                // inferred collapsed range; potential manual overlaps remain
                // explicitly unverified in sourceOutline, never called exact.
                for (axis, native, spans) in [("row", &supplement.source.row_runs, &mut sheet.row_hidden_spans),
                    ("column", &supplement.source.column_runs, &mut sheet.column_hidden_spans)] {
                    let mut cuts = collapsed_ranges(&state, axis);
                    if axis == "row" {
                        cuts.extend(sheet.auto_filter.iter().flat_map(|filter| filter.filtered_rows.iter()).map(|row| (*row, *row)));
                        cuts = merge_ranges(cuts);
                    }
                    let hidden = native.iter().filter(|run| run.hidden == Some(true)).map(|run| HiddenSpan { start: run.start, end: run.end, value: true }).collect();
                    *spans = subtract_ranges(hidden, &cuts).iter().map(payload_span).collect();
                }
                let mut row_cuts = collapsed_ranges(&state, "row");
                row_cuts.extend(sheet.auto_filter.iter().flat_map(|filter| filter.filtered_rows.iter()).map(|row| (*row, *row)));
                let row_cuts = merge_ranges(row_cuts);
                let column_cuts = collapsed_ranges(&state, "column");
                sheet.hidden_rows.retain(|row| !in_ranges(&row_cuts, *row));
                sheet.hidden_columns.retain(|column| !in_ranges(&column_cuts, u32::from(*column)));
            }
            if !state.groups.is_empty() || supplement.source.outline_pr_xml.is_some() || supplement.extension.is_some() { sheet.outline = Some(state); }
        }
        if supplement.extension.is_some() && !model_matched {
            supplement.warnings.push("OUTLINE_SNAPSHOT_STALE: external Office outline edits take precedence over the authored project model".into());
        }
        if supplement.extension.is_some() && model_matched && !visibility_matched && !axis_exact {
            supplement.warnings.push("OUTLINE_VISIBILITY_STALE: current Office hidden bits take precedence over the manual visibility snapshot".into());
        }
        if !visibility_matched && !axis_exact && (!supplement.source.row_runs.is_empty() || !supplement.source.column_runs.is_empty()) {
            supplement.source.diagnostics.push("NATIVE_HIDDEN_REASONS_AMBIGUOUS: standard hidden bits cannot identify manual/outline/filter/zero-size overlaps; complete native runs retained".into());
        }
        self.warnings.extend(supplement.warnings.iter().map(|warning| format!("{}: {warning}", sheet.name)));
        sheet.source_outline = json!({
            "version": 1, "native": supplement.source, "manualVisibilityNormalized": true,
            "authored": supplement.extension.as_ref().and_then(|state| state.source.as_ref()),
            "projectSnapshot": supplement.extension,
            "projectOutlineExtension": { "version": 1, "restored": model_matched, "projectionMatched": model_matched,
                "visibilityProjectionMatched": visibility_matched, "hiddenReasonsExact": visibility_matched || axis_exact },
        });
    }
}

fn read_sidecar(bytes: &[u8]) -> Result<SheetOutlineSidecar, String> {
    let mut archive = ZipArchive::new(Cursor::new(bytes)).map_err(|error| error.to_string())?;
    let parts = worksheet_parts(&mut archive)?;
    let mut result = SheetOutlineSidecar::default();
    for part in parts {
        let read = (|| {
            let entry = archive.by_name(&part.path).map_err(|error| error.to_string())?;
            if entry.size() > MAX_WORKSHEET_BYTES { return Err("worksheet exceeds the outline scanner's 512 MiB budget".into()); }
            read_worksheet_outline(BufReader::new(entry.take(MAX_WORKSHEET_BYTES + 1)), false)
        })();
        match read {
            Ok(source) => { result.by_sheet.insert(part.name, source); }
            Err(warning) => { result.failure.get_or_insert_with(|| format!("{}: XLSX outline metadata cannot be restored: {warning}", part.name)); },
        }
    }
    Ok(result)
}

fn payload_span(span: &HiddenSpan) -> AxisSpanPayload<bool> { AxisSpanPayload { start: span.start, end: span.end, value: span.value } }

fn collapsed_ranges(state: &office_io_common::sheet_outline_source::SheetOutlineState, axis: &str) -> Vec<(u32, u32)> {
    let mut ranges: Vec<(u32, u32)> = Vec::new();
    let mut groups = state.groups.iter().filter(|group| group.axis == axis && group.collapsed).collect::<Vec<_>>();
    groups.sort_by_key(|group| (group.start, group.end));
    for group in groups {
        if let Some(previous) = ranges.last_mut() { if group.start <= previous.1 + 1 { previous.1 = previous.1.max(group.end); continue; } }
        ranges.push((group.start, group.end));
    }
    ranges
}

fn in_ranges(ranges: &[(u32, u32)], index: u32) -> bool {
    let position = ranges.partition_point(|range| range.0 <= index);
    ranges.get(position.wrapping_sub(1)).is_some_and(|range| index <= range.1)
}

fn merge_ranges(mut ranges: Vec<(u32, u32)>) -> Vec<(u32, u32)> {
    ranges.sort_unstable();
    let mut result: Vec<(u32, u32)> = Vec::new();
    for (start, end) in ranges {
        if let Some(previous) = result.last_mut() { if start <= previous.1 + 1 { previous.1 = previous.1.max(end); continue; } }
        result.push((start, end));
    }
    result
}

fn subtract_ranges(spans: Vec<HiddenSpan>, cuts: &[(u32, u32)]) -> Vec<HiddenSpan> {
    let mut result: Vec<HiddenSpan> = Vec::new();
    for span in spans {
        let mut cursor = span.start;
        for &(start, end) in cuts.iter().skip_while(|(_, end)| *end < span.start).take_while(|(start, _)| *start <= span.end) {
            if start > cursor { push_span(&mut result, cursor, start - 1); }
            cursor = cursor.max(end + 1);
            if cursor > span.end { break; }
        }
        if cursor <= span.end { push_span(&mut result, cursor, span.end); }
    }
    result
}

fn push_span(spans: &mut Vec<HiddenSpan>, start: u32, end: u32) {
    if let Some(previous) = spans.last_mut() { if start <= previous.end + 1 { previous.end = previous.end.max(end); return; } }
    spans.push(HiddenSpan { start, end, value: true });
}
