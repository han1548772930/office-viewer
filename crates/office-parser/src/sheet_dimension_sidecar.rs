//! Restore the project's exact pixel geometry only while its Office projection
//! still matches the worksheet. External Office dimension edits take precedence.

use std::collections::{BTreeMap, BTreeSet};
use std::io::{BufRead, BufReader, Cursor, Read};

use office_io_common::{SHEET_AXIS_EXTENSION_URI, SHEET_AXIS_NAMESPACE, OfficeFilterNode, WorksheetAxisFingerprint, WorksheetFilterProjection};
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::{NsReader, XmlVersion};
use serde::Deserialize;
use serde_json::json;
use zip::ZipArchive;

use crate::model::{AxisPayload, AxisSpanPayload, SheetPayload};
use crate::sheet_filter_import::{self, AutoFilterState};
use crate::xlsx_parts::{spreadsheet_namespace, worksheet_parts};

const MAX_ROW_INDEX: u32 = 1_048_575;
const MAX_COLUMN_INDEX: u32 = 16_383;
const MAX_AXIS_JSON_BYTES: usize = 32 * 1024 * 1024;
const MAX_WORKSHEET_XML_BYTES: u64 = 512 * 1024 * 1024;

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct AxisState {
    version: u32,
    axis_unit: String,
    row_heights: Vec<AxisPayload>,
    column_widths: Vec<AxisPayload>,
    row_height_spans: Vec<AxisSpanPayload<f64>>,
    column_width_spans: Vec<AxisSpanPayload<f64>>,
    row_auto_height_spans: Vec<AxisSpanPayload<bool>>,
    row_auto_heights: Vec<AxisPayload>,
    default_row_height: Option<f64>,
    default_column_width: Option<f64>,
    default_row_auto_height: Option<bool>,
    row_count: Option<u32>,
    column_count: Option<u32>,
    hidden_rows: Vec<u32>,
    hidden_columns: Vec<u32>,
    #[serde(default)]
    row_hidden_spans: Vec<AxisSpanPayload<bool>>,
    #[serde(default)]
    column_hidden_spans: Vec<AxisSpanPayload<bool>>,
    auto_filter: Option<AutoFilterState>,
    filter_fingerprint: Option<String>,
}

#[derive(Debug)]
struct WorksheetSidecar {
    state: Option<AxisState>,
    axis_matched: bool,
    filter_fingerprint: String,
    filter_projection: Option<OfficeFilterNode>,
    warnings: Vec<String>,
}

#[derive(Debug, Default)]
pub(crate) struct SheetDimensionSidecar {
    by_sheet: BTreeMap<String, WorksheetSidecar>,
    colors: crate::sheet_filter_color::FilterColors,
    pub(crate) warnings: Vec<String>,
}

impl SheetDimensionSidecar {
    pub(crate) fn read(bytes: &[u8]) -> Self {
        match read_sidecar(bytes) {
            Ok(value) => value,
            Err(warning) => Self { warnings: vec![format!("Could not restore XLSX pixel dimensions: {warning}")], ..Self::default() },
        }
    }

    pub(crate) fn restore(&mut self, sheet: &mut SheetPayload) {
        let Some(mut supplement) = self.by_sheet.remove(&sheet.name) else { return; };
        let standard_hidden_rows = sheet.hidden_rows.clone();
        let filter_matched = supplement.state.as_ref().and_then(|state| state.filter_fingerprint.as_deref()) == Some(supplement.filter_fingerprint.as_str());
        let filter = if filter_matched {
            supplement.state.as_mut().and_then(|state| state.auto_filter.take())
        } else if let Some(source) = &supplement.filter_projection {
            match sheet_filter_import::standard_filter_with_colors(source, &sheet.cells, &standard_hidden_rows, &self.colors) {
                Ok(filter) => Some(filter),
                Err(warning) => { self.warnings.push(format!("{}: {warning}; complete source XML and current visibility are retained", sheet.name)); None }
            }
        } else { None };
        sheet.source_auto_filter = supplement.filter_projection;
        sheet.auto_filter = filter;
        if let Some(source) = sheet.source_axis_dimensions.as_object_mut() {
            source.insert("projectAutoFilterExtension".into(), json!({ "version": 1, "restored": filter_matched, "projectionMatched": filter_matched, "hiddenReasonsExact": supplement.axis_matched && filter_matched }));
        }
        let Some(state) = supplement.state else { return; };
        if !supplement.axis_matched {
            // A resize makes exact pixels stale without changing filter criteria.
            // Keep newly authored Office hides, and remove known filter hides
            // from the manual set while preserving an original manual overlap.
            if filter_matched {
                if let Some(filter) = &sheet.auto_filter {
                    let manual = state.hidden_rows.iter().copied().collect::<BTreeSet<_>>();
                    let filtered = filter.filtered_rows.iter().copied().collect::<BTreeSet<_>>();
                    sheet.hidden_rows.retain(|row| manual.contains(row) || state.row_hidden_spans.iter().any(|span| span.value && *row >= span.start && *row <= span.end) || !filtered.contains(row));
                }
            }
            return;
        }
        sheet.axis_unit = state.axis_unit;
        sheet.row_heights = state.row_heights;
        sheet.column_widths = state.column_widths;
        sheet.row_height_spans = state.row_height_spans;
        sheet.column_width_spans = state.column_width_spans;
        sheet.row_auto_height_spans = state.row_auto_height_spans;
        sheet.row_auto_heights = state.row_auto_heights;
        sheet.default_row_height = state.default_row_height;
        sheet.default_column_width = state.default_column_width;
        sheet.default_row_auto_height = state.default_row_auto_height;
        sheet.row_count = state.row_count;
        sheet.column_count = state.column_count;
        // The standard hidden attribute combines zero dimensions and manual
        // hidden state. The extension keeps the latter independently.
        sheet.hidden_rows = state.hidden_rows;
        sheet.hidden_columns = state.hidden_columns.into_iter().map(|column| column as u16).collect();
        sheet.row_hidden_spans = state.row_hidden_spans;
        sheet.column_hidden_spans = state.column_hidden_spans;
        // Unprojected external filter criteria still keep their observable
        // hidden rows. The raw tree remains in sourceAutoFilter for its adapter.
        if sheet.source_auto_filter.is_some() && sheet.auto_filter.is_none() {
            sheet.hidden_rows.extend(standard_hidden_rows);
            sheet.hidden_rows.sort_unstable();
            sheet.hidden_rows.dedup();
        }
        if let Some(source) = sheet.source_axis_dimensions.as_object_mut() {
            source.insert("projectAxisExtension".into(), json!({ "version": 1, "restored": true, "projectionMatched": true }));
        }
    }
}

fn read_sidecar(bytes: &[u8]) -> Result<SheetDimensionSidecar, String> {
    let mut archive = ZipArchive::new(Cursor::new(bytes)).map_err(|error| error.to_string())?;
    let parts = worksheet_parts(&mut archive)?;
    let mut result = SheetDimensionSidecar::default();
    for part in parts {
        let read = (|| {
            let entry = archive.by_name(&part.path).map_err(|error| error.to_string())?;
            if entry.size() > MAX_WORKSHEET_XML_BYTES { return Err("worksheet XML exceeds the pixel-dimension reader budget".into()); }
            // Read once as events: large cell XML is not retained beside the
            // parser's workbook model just to inspect its compact extension.
            read_worksheet_state(BufReader::new(entry.take(MAX_WORKSHEET_XML_BYTES + 1)))
        })();
        match read {
            Ok(mut state) => {
                for warning in state.warnings.drain(..) { result.warnings.push(format!("{}: {warning}", part.name)); }
                if state.state.is_some() || state.filter_projection.is_some() { result.by_sheet.insert(part.name, state); }
            }
            Err(warning) => result.warnings.push(format!("{}: skipped XLSX pixel-dimension extension: {warning}", part.name)),
        }
    }
    match crate::sheet_filter_color::referenced_ids(result.by_sheet.values().filter_map(|state| state.filter_projection.as_ref()))
        .and_then(|ids| crate::sheet_filter_color::FilterColors::read(&mut archive, &ids)) {
        Ok(colors) => result.colors = colors,
        Err(error) => { result.warnings.push(error.clone()); result.colors.error = Some(error); }
    }
    for (name, state) in &mut result.by_sheet {
        match office_io_common::office_filter_fingerprint(state.filter_projection.as_ref(), |node| result.colors.signature(node)) {
            Ok(fingerprint) => state.filter_fingerprint = fingerprint,
            Err(error) => { state.filter_fingerprint.clear(); result.warnings.push(format!("{name}: {error}; current filter source and visibility retained")); }
        }
        if state.state.as_ref().and_then(|state| state.filter_fingerprint.as_deref()).is_some_and(|expected| expected != state.filter_fingerprint) {
            result.warnings.push(format!("{name}: Office AutoFilter criteria changed outside this snapshot; using current filter colors and retaining source XML"));
        }
    }
    Ok(result)
}

struct Capture { depth: usize, fingerprint: String, text: String, invalid: bool }

fn read_worksheet_state<R: BufRead>(input: R) -> Result<WorksheetSidecar, String> {
    let mut reader = NsReader::from_reader(input);
    let mut fingerprint = WorksheetAxisFingerprint::default();
    let mut buffer = Vec::new();
    let mut root_seen = false;
    let mut extension_depth = None;
    let mut capture: Option<Capture> = None;
    let mut candidate: Option<(String, AxisState)> = None;
    let mut axis_extension_count = 0usize;
    let mut warnings = Vec::new();
    let mut filter_projection = WorksheetFilterProjection::default();
    loop {
        let event = reader.read_event_into(&mut buffer).map_err(|error| error.to_string())?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                let name = element.local_name();
                let namespace = reader.resolver().resolve_element(element.name()).0;
                let is_sheet = spreadsheet_namespace(namespace.clone());
                if fingerprint.depth() == 0 {
                    if root_seen || name.as_ref() != "worksheet" || !is_sheet { return Err("invalid worksheet root".into()); }
                    root_seen = true;
                }
                if let Some(capture) = &mut capture {
                    if !capture.invalid { warnings.push("nested markup inside the pixel-dimension JSON; using standard worksheet data".into()); }
                    capture.invalid = true;
                    capture.text.clear();
                }
                let filter_root = fingerprint.path_is(&[b"worksheet"]) && name.as_ref() == "autoFilter" && is_sheet;
                filter_projection.start(&reader, element, filter_root, matches!(&event, Event::Empty(_)))?;
                if fingerprint.path_is(&[b"worksheet", b"extLst"]) && name.as_ref() == "ext" && is_sheet {
                    match attributes(&reader, element) {
                        Ok(attributes) if attributes.get("uri").is_some_and(|uri| uri == SHEET_AXIS_EXTENSION_URI) && matches!(&event, Event::Start(_)) => {
                            extension_depth = Some(fingerprint.depth() + 1);
                        }
                        Err(warning) => warnings.push(format!("invalid worksheet extension attributes: {warning}")),
                        _ => {}
                    }
                }
                if extension_depth == Some(fingerprint.depth()) && fingerprint.path_is(&[b"worksheet", b"extLst", b"ext"])
                    && name.as_ref() == "axisDimensions" && matches!(namespace, ResolveResult::Bound(value) if value.as_ref() == SHEET_AXIS_NAMESPACE) {
                    axis_extension_count += 1;
                    if axis_extension_count > 1 { warnings.push("duplicate pixel-dimension extensions; using standard worksheet data".into()); }
                    if matches!(&event, Event::Empty(_)) {
                        warnings.push("empty pixel-dimension extension; using standard worksheet data".into());
                    } else {
                        let expected = (|| -> Result<String, String> {
                            let attributes = attributes(&reader, element)?;
                            if attributes.get("version").map(String::as_str) != Some("1") { return Err("unsupported pixel-dimension extension version".into()); }
                            attributes.get("fingerprint").filter(|value| value.len() == 16 && value.bytes().all(|byte| byte.is_ascii_hexdigit()))
                                .map(|value| value.to_ascii_lowercase()).ok_or_else(|| "invalid pixel-dimension projection fingerprint".to_owned())
                        })();
                        let invalid = expected.is_err();
                        let expected = match expected { Ok(value) => value, Err(warning) => { warnings.push(warning); String::new() } };
                        capture = Some(Capture { depth: fingerprint.depth() + 1, fingerprint: expected, text: String::new(), invalid });
                    }
                }
                fingerprint.start(element, matches!(&event, Event::Empty(_)))?;
            }
            Event::Text(text) => {
                let text = text.xml_content(XmlVersion::Implicit1_0);
                if let Some(capture) = &mut capture { append_capture(capture, &text, &mut warnings); }
                else { filter_projection.text(&text)?; }
            }
            Event::CData(text) => {
                let text = text.xml_content(XmlVersion::Implicit1_0);
                if let Some(capture) = &mut capture { append_capture(capture, &text, &mut warnings); }
                else { filter_projection.text(&text)?; }
            }
            Event::GeneralRef(reference) => {
                let decoded = if let Some(character) = reference.resolve_char_ref().map_err(|error| error.to_string())? { character.to_string() }
                    else { let name = reference.as_ref(); quick_xml::escape::resolve_xml_entity(&name).ok_or_else(|| "unknown worksheet XML entity".to_owned())?.to_owned() };
                if let Some(capture) = &mut capture { append_capture(capture, &decoded, &mut warnings); }
                else { filter_projection.text(&decoded)?; }
            }
            Event::End(_) => {
                if capture.as_ref().is_some_and(|capture| capture.depth == fingerprint.depth()) {
                    let capture = capture.take().unwrap();
                    if !capture.invalid {
                        let parsed: Result<AxisState, String> = (|| {
                            let state: AxisState = serde_json::from_str(&capture.text).map_err(|error| format!("invalid pixel-dimension JSON: {error}"))?;
                            validate_state(&state)?;
                            Ok(state)
                        })();
                        match parsed {
                            Ok(state) => candidate = Some((capture.fingerprint, state)),
                            Err(warning) => warnings.push(format!("{warning}; using standard worksheet data")),
                        }
                    }
                }
                if extension_depth == Some(fingerprint.depth()) { extension_depth = None; }
                filter_projection.end();
                fingerprint.end()?;
            }
            Event::DocType(_) => return Err("worksheet DTDs are unsupported".into()),
            Event::Eof => break,
            _ => {}
        }
        if reader.buffer_position() > MAX_WORKSHEET_XML_BYTES { return Err("worksheet XML exceeds the pixel-dimension reader budget".into()); }
        buffer.clear();
    }
    if !root_seen || fingerprint.depth() != 0 || capture.is_some() { return Err("incomplete worksheet XML".into()); }
    if axis_extension_count != 1 { candidate = None; }
    let axis_matched = candidate.as_ref().is_some_and(|(expected, _)| *expected == fingerprint.finish());
    let filter_fingerprint = filter_projection.fingerprint()?;
    if candidate.is_some() && !axis_matched { warnings.push("Office row/column dimensions changed outside this snapshot; using the current Office dimensions".into()); }
    Ok(WorksheetSidecar { state: candidate.map(|(_, state)| state), axis_matched, filter_fingerprint, filter_projection: filter_projection.take_root(), warnings })
}

fn attributes<R: BufRead>(reader: &NsReader<R>, element: &BytesStart<'_>) -> Result<BTreeMap<String, String>, String> {
    let mut result = BTreeMap::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| error.to_string())?;
        if attribute.key.as_ref().as_bytes().contains(&b':') || attribute.key.as_ref().as_bytes() == b"xmlns" { continue; }
        let name = std::str::from_utf8(attribute.key.as_ref().as_bytes()).map_err(|error| error.to_string())?.to_owned();
        let value = attribute.normalized_value(XmlVersion::Implicit1_0).map_err(|error| error.to_string())?.into_owned();
        result.insert(name, value);
    }
    Ok(result)
}

fn append(target: &mut String, value: &str) -> Result<(), String> {
    if target.len().saturating_add(value.len()) > MAX_AXIS_JSON_BYTES { return Err("pixel-dimension JSON exceeds the parser budget".into()); }
    target.push_str(value);
    Ok(())
}

fn append_capture(capture: &mut Capture, value: &str, warnings: &mut Vec<String>) {
    if capture.invalid { return; }
    if let Err(warning) = append(&mut capture.text, value) {
        capture.invalid = true;
        capture.text.clear();
        warnings.push(format!("{warning}; using standard worksheet data"));
    }
}

fn validate_state(state: &AxisState) -> Result<(), String> {
    if state.version != 1 || state.axis_unit != "cssPixel" { return Err("unsupported pixel-dimension payload schema".into()); }
    let valid_size = |size: &f64| size.is_finite() && *size >= 0.0;
    if state.default_row_height.as_ref().is_some_and(|size| !valid_size(size))
        || state.default_column_width.as_ref().is_some_and(|size| !valid_size(size)) {
        return Err("invalid pixel-dimension defaults".into());
    }
    validate_points(&state.row_heights, MAX_ROW_INDEX)?;
    validate_points(&state.row_auto_heights, MAX_ROW_INDEX)?;
    validate_points(&state.column_widths, MAX_COLUMN_INDEX)?;
    validate_spans(&state.row_height_spans, MAX_ROW_INDEX, valid_size)?;
    validate_spans(&state.column_width_spans, MAX_COLUMN_INDEX, valid_size)?;
    validate_spans(&state.row_auto_height_spans, MAX_ROW_INDEX, |_| true)?;
    validate_spans(&state.row_hidden_spans, state.row_count.filter(|count| *count > 0).map(|count| count - 1).unwrap_or(MAX_ROW_INDEX), |_| true)?;
    validate_spans(&state.column_hidden_spans, state.column_count.filter(|count| *count > 0).map(|count| count - 1).unwrap_or(MAX_COLUMN_INDEX), |_| true)?;
    if let Some(filter) = &state.auto_filter { sheet_filter_import::validate(filter)?; }
    if state.auto_filter.is_some() && state.filter_fingerprint.is_none()
        || state.filter_fingerprint.as_ref().is_some_and(|value| value.len() != 16 || !value.bytes().all(|byte| byte.is_ascii_hexdigit())) {
        return Err("invalid typed AutoFilter projection fingerprint".into());
    }
    if state.row_count.is_some_and(|count| count == 0 || count > MAX_ROW_INDEX + 1)
        || state.column_count.is_some_and(|count| count == 0 || count > MAX_COLUMN_INDEX + 1)
        || state.hidden_rows.iter().any(|index| *index > MAX_ROW_INDEX || state.row_count.is_some_and(|count| *index >= count))
        || state.hidden_columns.iter().any(|index| *index > MAX_COLUMN_INDEX || state.column_count.is_some_and(|count| *index >= count)) {
        return Err("pixel-dimension counts or hidden indices exceed the worksheet or Office limit".into());
    }
    Ok(())
}

fn validate_points(points: &[AxisPayload], maximum: u32) -> Result<(), String> {
    let mut indices = BTreeSet::new();
    for point in points {
        if point.index > maximum || !point.size.is_finite() || point.size < 0.0 || !indices.insert(point.index) {
            return Err("invalid or duplicate pixel-dimension point".into());
        }
    }
    Ok(())
}

fn validate_spans<T>(spans: &[AxisSpanPayload<T>], maximum: u32, valid: impl Fn(&T) -> bool) -> Result<(), String> {
    let mut previous = None;
    for span in spans {
        if span.start > span.end || span.end > maximum || previous.is_some_and(|end| span.start <= end) || !valid(&span.value) {
            return Err("invalid, unordered or overlapping pixel-dimension span".into());
        }
        previous = Some(span.end);
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use office_io_common::worksheet_axis_fingerprint;

    fn read_axis_state<R: BufRead>(input: R) -> Result<Option<AxisState>, String> {
        let result = read_worksheet_state(input)?;
        if result.state.is_some() && !result.axis_matched { return Err("stale Office axis projection".into()); }
        Ok(result.state)
    }

    fn worksheet(state: serde_json::Value) -> String {
        let prefix = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetFormatPr defaultRowHeight="19.5" defaultColWidth="16"/><sheetData/>"#;
        let fingerprint = worksheet_axis_fingerprint(prefix).unwrap();
        format!(r#"{prefix}<extLst><ext uri="{SHEET_AXIS_EXTENSION_URI}"><axis:axisDimensions xmlns:axis="{SHEET_AXIS_NAMESPACE}" version="1" fingerprint="{fingerprint}">{}</axis:axisDimensions></ext></extLst></worksheet>"#, quick_xml::escape::escape(&state.to_string()))
    }
    fn state() -> serde_json::Value {
        json!({ "version": 1, "axisUnit": "cssPixel", "rowHeights": [{ "index": 9, "size": 31.123456789 }], "columnWidths": [],
            "rowHeightSpans": [], "columnWidthSpans": [{ "start": 2, "end": 4, "value": 0 }], "rowAutoHeightSpans": [], "rowAutoHeights": [],
            "defaultRowHeight": null, "defaultColumnWidth": null, "defaultRowAutoHeight": null, "rowCount": 400, "columnCount": 30, "hiddenRows": [8], "hiddenColumns": [5] })
    }

    #[test]
    fn reads_entity_split_json_exact_pixels_modes_counts_and_hidden_reasons() {
        let parsed = read_axis_state(Cursor::new(worksheet(state()).into_bytes())).unwrap().unwrap();
        assert_eq!(parsed.row_heights[0].size, 31.123456789);
        assert_eq!(parsed.column_width_spans[0].value, 0.0);
        assert_eq!((parsed.row_count, parsed.column_count), (Some(400), Some(30)));
        assert_eq!(parsed.hidden_rows, [8]);
        assert_eq!(parsed.default_row_height, None);
        assert_eq!(parsed.default_row_auto_height, None);
    }

    #[test]
    fn external_dimension_edits_skip_the_extension_while_cell_content_edits_do_not() {
        let original = worksheet(state());
        assert!(read_axis_state(Cursor::new(original.replace("defaultRowHeight=\"19.5\"", "defaultRowHeight=\"30\"").into_bytes())).is_err());
        assert!(read_axis_state(Cursor::new(original.replace("<sheetData/>", "<sheetData><row r=\"1\"><c r=\"A1\"><v>9</v></c></row></sheetData>").into_bytes())).unwrap().is_some());
    }

    #[test]
    fn rejects_invalid_and_duplicate_spans_instead_of_restoring_partial_geometry() {
        for spans in [json!([{ "start": 2, "end": 1, "value": 20 }]), json!([{ "start": 0, "end": 2, "value": 20 }, { "start": 2, "end": 3, "value": 30 }]), json!([{ "start": 0, "end": 0, "value": -1 }])] {
            let mut invalid = state(); invalid["rowHeightSpans"] = spans;
            assert!(read_axis_state(Cursor::new(worksheet(invalid).into_bytes())).unwrap().is_none());
        }
        let xml = worksheet(state()).replace(SHEET_AXIS_NAMESPACE, "urn:other");
        assert!(read_axis_state(Cursor::new(xml.into_bytes())).unwrap().is_none());
    }

    #[test]
    fn malformed_project_extensions_do_not_discard_standard_filter_source() {
        let filter = r#"<autoFilter ref="A1:C9"><filterColumn colId="1"><top10 val="4"/></filterColumn></autoFilter>"#;
        let original = worksheet(state()).replace("<extLst>", &format!("{filter}<extLst>"));
        let invalid_json = original.replace(&quick_xml::escape::escape(&state().to_string()).to_string(), "{broken");
        let invalid_nested = original.replace("</axis:axisDimensions>", "<unexpected/></axis:axisDimensions>");
        let invalid_version = original.replace("version=\"1\"", "version=\"99\"");
        let invalid_fingerprint = original.replace(" fingerprint=\"", " fingerprint=\"invalid");
        for xml in [invalid_json, invalid_nested, invalid_version, invalid_fingerprint] {
            let result = read_worksheet_state(Cursor::new(xml.into_bytes())).unwrap();
            assert!(result.state.is_none());
            assert!(!result.warnings.is_empty());
            let source = result.filter_projection.unwrap();
            assert!(source.is_standard("autoFilter"));
            assert!(source.children[0].children[0].is_standard("top10"));
        }
    }
}
