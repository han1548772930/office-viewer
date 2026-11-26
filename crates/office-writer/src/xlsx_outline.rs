//! Native outline projection with an exact, independently guarded model snapshot.
//! Axis geometry and manual hidden state remain owned by xlsx_dimensions.

use std::collections::{BTreeMap, BTreeSet};

use office_io_common::conditional_format_xml::{Element, MAIN_NAMESPACE, STRICT_MAIN_NAMESPACE, parse};
use office_io_common::sheet_outline_source::{HiddenSpan, ManualVisibility, NativeOutlineRun, NativeOutlineSource,
    OutlineExtension, SheetOutlineState, MAX_NATIVE_OUTLINE_LEVEL, MAX_OUTLINE_JSON_BYTES,
    run_at, validate_extension, validate_source, validate_state, worksheet_outline_fingerprints,
    SHEET_OUTLINE_EXTENSION_URI, SHEET_OUTLINE_NAMESPACE};
use serde_json::{Value, json};

use super::{SheetInput, MAX_COLUMN_INDEX, MAX_ROW_INDEX, esc};
use crate::ExportError;

pub(super) type OutlineInput = SheetOutlineState;

#[derive(Debug, Clone)]
struct ProjectionRun { start: u32, end: u32, level: u8, hidden: bool, collapsed: bool, extra_attributes: String }

pub(super) struct WorksheetOutlinePlan {
    state: Option<SheetOutlineState>,
    source: Option<NativeOutlineSource>,
    snapshot_extra: BTreeMap<String, Value>,
    rows: Vec<ProjectionRun>,
    columns: Vec<ProjectionRun>,
    pub warnings: Vec<String>,
}

pub(super) fn prepare(sheet: &SheetInput) -> Result<WorksheetOutlinePlan, ExportError> {
    if sheet.row_count.is_some_and(|count| count == 0 || count > MAX_ROW_INDEX + 1)
        || sheet.column_count.is_some_and(|count| count == 0 || count > MAX_COLUMN_INDEX + 1) {
        return Err(ExportError::invalid("outline worksheet counts exceed the XLSX limits"));
    }
    let mut state = sheet.outline.clone();
    if let Some(state) = &mut state {
        for (index, group) in state.groups.iter_mut().enumerate() {
            // Legacy internal export inputs had no ID. Assign a deterministic
            // ID once; canonical authored IDs and all unknown JSON stay intact.
            if group.id.is_empty() { group.id = format!("xlsx-outline-legacy-{index}"); }
        }
        validate_state(state, sheet.row_count.unwrap_or(MAX_ROW_INDEX + 1), sheet.column_count.unwrap_or(MAX_COLUMN_INDEX + 1)).map_err(ExportError::invalid)?;
    }
    let (source, snapshot_extra) = source_from_metadata(sheet.metadata.as_ref())?;
    let mut warnings = Vec::new();
    let empty = SheetOutlineState::default();
    let current = state.as_ref().unwrap_or(&empty);
    let rows = project_axis(current, "row", sheet.row_count.map(|count| count - 1).unwrap_or(MAX_ROW_INDEX),
        source.as_ref().map(|source| source.row_runs.as_slice()).unwrap_or(&[]), &mut warnings)?;
    let columns = project_axis(current, "column", sheet.column_count.map(|count| count - 1).unwrap_or(MAX_COLUMN_INDEX),
        source.as_ref().map(|source| source.column_runs.as_slice()).unwrap_or(&[]), &mut warnings)?;
    Ok(WorksheetOutlinePlan { state, source, snapshot_extra, rows, columns, warnings })
}

fn source_from_metadata(metadata: Option<&Value>) -> Result<(Option<NativeOutlineSource>, BTreeMap<String, Value>), ExportError> {
    let Some(raw) = metadata.and_then(|value| value.get("outlineOfficeSource").or_else(|| value.pointer("/source/sourceOutline")).or_else(|| value.get("sourceOutline")))
        .filter(|value| !value.is_null()) else { return Ok((None, BTreeMap::new())); };
    let candidate = raw.get("native").or_else(|| raw.get("authored")).unwrap_or(raw);
    let mut source = if candidate.is_null() { None } else {
        let source: NativeOutlineSource = serde_json::from_value(candidate.clone()).map_err(|error| ExportError::invalid(format!("invalid outline Office source: {error}")))?;
        validate_source(&source).map_err(ExportError::invalid)?;
        Some(source)
    };
    let extra = if let Some(snapshot) = raw.get("projectSnapshot").filter(|value| !value.is_null()) {
        let snapshot: OutlineExtension = serde_json::from_value(snapshot.clone()).map_err(|error| ExportError::invalid(format!("invalid retained outline snapshot: {error}")))?;
        snapshot.extra
    } else { BTreeMap::new() };
    if let Some(current) = &mut source {
        if let Some(authored) = raw.get("authored").or_else(|| raw.pointer("/projectSnapshot/source")).filter(|value| !value.is_null()) {
            let authored: NativeOutlineSource = serde_json::from_value(authored.clone()).map_err(|error| ExportError::invalid(format!("invalid authored outline source: {error}")))?;
            validate_source(&authored).map_err(ExportError::invalid)?;
            for (key, value) in authored.extra { current.extra.entry(key).or_insert(value); }
        }
        if raw.get("native").is_some() {
            let unknown = raw.as_object().unwrap().iter().filter(|(key, _)| !matches!(key.as_str(),
                "version" | "native" | "authored" | "projectSnapshot" | "projectOutlineExtension" | "manualVisibilityNormalized"))
                .map(|(key, value)| (key.clone(), value.clone())).collect::<serde_json::Map<String, Value>>();
            if !unknown.is_empty() { current.extra.insert("retainedImportMetadata".into(), Value::Object(unknown)); }
        }
        validate_source(current).map_err(ExportError::invalid)?;
    }
    Ok((source, extra))
}

fn project_axis(state: &SheetOutlineState, axis: &str, maximum: u32, source: &[NativeOutlineRun], warnings: &mut Vec<String>) -> Result<Vec<ProjectionRun>, ExportError> {
    let mut events = BTreeMap::<u32, (i32, i32)>::new();
    let mut markers = BTreeSet::new();
    let mut marker_owners = BTreeMap::<u32, Vec<bool>>::new();
    let at_start = if axis == "row" { state.summary_above } else { state.summary_left };
    for group in state.groups.iter().filter(|group| group.axis == axis) {
        let start = events.entry(group.start).or_default(); start.0 += 1; start.1 += i32::from(group.collapsed);
        let end = events.entry(group.end + 1).or_default(); end.0 -= 1; end.1 -= i32::from(group.collapsed);
        let marker = if at_start { group.start.checked_sub(1) } else { group.end.checked_add(1).filter(|index| *index <= maximum) };
        if let Some(marker) = marker {
            marker_owners.entry(marker).or_default().push(group.collapsed);
            if group.collapsed { markers.insert(marker); }
        } else if group.collapsed {
            warnings.push(format!("OUTLINE_BOUNDARY_SNAPSHOT: {axis} group {} has no adjacent summary dimension; exact collapse remains in the project extension", group.id));
        }
    }
    for owners in marker_owners.values() {
        if owners.iter().any(|value| *value) && owners.iter().any(|value| !*value) {
            warnings.push(format!("OUTLINE_SHARED_MARKER_SNAPSHOT: {axis} groups share one native collapsed marker; exact per-group state remains in the project extension")); break;
        }
    }
    let mut boundaries = BTreeSet::from([0, maximum + 1]);
    boundaries.extend(events.keys().copied());
    for marker in &markers { boundaries.insert(*marker); boundaries.insert(*marker + 1); }
    for run in source {
        if run.start <= maximum && has_unknown_attributes(run) { boundaries.insert(run.start); boundaries.insert((run.end + 1).min(maximum + 1)); }
    }
    let positions = boundaries.into_iter().collect::<Vec<_>>();
    let mut depth = 0i32;
    let mut collapsed_depth = 0i32;
    let mut result = Vec::<ProjectionRun>::new();
    let mut clipped = false;
    for pair in positions.windows(2) {
        if let Some(change) = events.get(&pair[0]) { depth += change.0; collapsed_depth += change.1; }
        if depth > i32::from(MAX_NATIVE_OUTLINE_LEVEL) { clipped = true; }
        let level = depth.min(i32::from(MAX_NATIVE_OUTLINE_LEVEL)) as u8;
        let hidden = collapsed_depth > 0;
        let collapsed = markers.contains(&pair[0]);
        let extra_attributes = run_at(source, pair[0]).filter(|run| has_unknown_attributes(run)).map(preserved_attributes).transpose()?.unwrap_or_default();
        if level == 0 && !collapsed && extra_attributes.is_empty() { continue; }
        if let Some(previous) = result.last_mut() {
            if previous.end + 1 == pair[0] && previous.level == level && previous.hidden == hidden && previous.collapsed == collapsed && previous.extra_attributes == extra_attributes {
                previous.end = pair[1] - 1; continue;
            }
        }
        result.push(ProjectionRun { start: pair[0], end: pair[1] - 1, level, hidden, collapsed, extra_attributes });
    }
    if clipped { warnings.push(format!("OUTLINE_DEPTH_SNAPSHOT: {axis} level eight is projected as native level seven; all eight authored model layers remain in the project extension")); }
    Ok(result)
}

fn has_unknown_attributes(run: &NativeOutlineRun) -> bool {
    run.attributes.keys().any(|name| !matches!(name.as_str(), "outlineLevel" | "collapsed" | "hidden"))
}

fn preserved_attributes(run: &NativeOutlineRun) -> Result<String, ExportError> {
    let mut node = Element::new(MAIN_NAMESPACE, "row");
    for (prefix, namespace) in &run.namespaces {
        if !prefix.is_empty() { node.set_attr(&format!("xmlns:{prefix}"), namespace); }
    }
    let mut generated = BTreeMap::<String, String>::new();
    for (key, value) in &run.attributes {
        if matches!(key.as_str(), "outlineLevel" | "collapsed" | "hidden") { continue; }
        if let Some(tail) = key.strip_prefix('{') {
            let (namespace, name) = tail.split_once('}').ok_or_else(|| ExportError::invalid("invalid expanded outline attribute"))?;
            let existing = run.namespaces.iter().find(|(prefix, uri)| !prefix.is_empty() && uri.as_str() == namespace).map(|(prefix, _)| prefix.clone());
            let prefix = if let Some(prefix) = existing { prefix } else {
                let mut index = generated.len();
                let prefix = loop { let prefix = format!("oso{index}"); if !run.namespaces.contains_key(&prefix) && !generated.values().any(|value| value == &prefix) { break prefix; } index += 1; };
                generated.entry(namespace.into()).or_insert(prefix).clone()
            };
            node.set_attr_ns(namespace, &format!("{prefix}:{name}"), value);
        } else { node.set_attr(key, value); }
    }
    node.validate_fragment().map_err(ExportError::invalid)?;
    let rendered = node.render();
    let parsed = parse(&rendered).map_err(ExportError::invalid)?;
    let rendered = parsed.render();
    let start = rendered.find(' ').unwrap_or(rendered.len() - 2);
    let end = rendered.rfind("/>").ok_or_else(|| ExportError::invalid("invalid outline attribute fragment"))?;
    Ok(rendered[start..end].to_owned())
}

fn projection_at(runs: &[ProjectionRun], index: u32) -> Option<&ProjectionRun> {
    let position = runs.partition_point(|run| run.start <= index);
    runs.get(position.wrapping_sub(1)).filter(|run| index <= run.end)
}

fn normalize_main_namespace(node: &mut Element) {
    if node.namespace == STRICT_MAIN_NAMESPACE { node.namespace = MAIN_NAMESPACE.into();
        if node.attrs.get("xmlns").is_some_and(|value| value == STRICT_MAIN_NAMESPACE) { node.set_attr("xmlns", MAIN_NAMESPACE); }
        for (name, value) in &mut node.attrs { if (name == "xmlns" || name.starts_with("xmlns:")) && value == STRICT_MAIN_NAMESPACE { *value = MAIN_NAMESPACE.into(); } }
    }
    for child in &mut node.children { normalize_main_namespace(child); }
}

impl WorksheetOutlinePlan {
    pub fn sheet_properties_xml(&self, filter_mode: bool) -> Result<String, ExportError> {
        if self.state.is_none() && self.source.as_ref().and_then(|source| source.outline_pr_xml.as_ref()).is_none() {
            return Ok(if filter_mode { "<sheetPr filterMode=\"1\"/>".into() } else { String::new() });
        }
        let mut properties = Element::new(MAIN_NAMESPACE, "sheetPr");
        if filter_mode { properties.set_attr("filterMode", "1"); }
        let mut outline = self.source.as_ref().and_then(|source| source.outline_pr_xml.as_ref()).map(|xml| parse(xml).map_err(ExportError::invalid)).transpose()?
            .unwrap_or_else(|| Element::new(MAIN_NAMESPACE, "outlinePr"));
        normalize_main_namespace(&mut outline);
        let state = self.state.as_ref();
        outline.set_attr("summaryBelow", if state.is_some_and(|state| state.summary_above) { "0" } else { "1" });
        outline.set_attr("summaryRight", if state.is_some_and(|state| state.summary_left) { "0" } else { "1" });
        properties.push(outline);
        properties.validate_fragment().map_err(ExportError::invalid)?;
        Ok(properties.render())
    }

    pub fn format_xml(&self, axis_format: &str) -> Result<String, ExportError> {
        let mut format = parse(axis_format).map_err(ExportError::invalid)?;
        for (name, runs) in [("outlineLevelRow", &self.rows), ("outlineLevelCol", &self.columns)] {
            let level = runs.iter().map(|run| run.level).max().unwrap_or(0);
            if level > 0 { format.set_attr(name, level.to_string()); }
        }
        Ok(format.render())
    }

    pub fn columns_xml(&self, axis_columns: &str) -> Result<String, ExportError> {
        if self.columns.is_empty() { return Ok(axis_columns.to_owned()); }
        let base = if axis_columns.is_empty() { Element::new(MAIN_NAMESPACE, "cols") } else { parse(axis_columns).map_err(ExportError::invalid)? };
        let mut boundaries = BTreeSet::from([0, MAX_COLUMN_INDEX + 1]);
        let mut authored = Vec::<(u32, u32, Element)>::new();
        for col in &base.children {
            let index = |key: &str| col.attrs.get(key).ok_or_else(|| ExportError::invalid("axis column lacks an index"))?.parse::<u32>()
                .map_err(|_| ExportError::invalid("invalid axis column index"))?.checked_sub(1).ok_or_else(|| ExportError::invalid("axis column index is zero"));
            let (start, end) = (index("min")?, index("max")?);
            boundaries.insert(start); boundaries.insert(end + 1); authored.push((start, end, col.clone()));
        }
        for run in &self.columns { boundaries.insert(run.start); boundaries.insert(run.end + 1); }
        let positions = boundaries.into_iter().collect::<Vec<_>>();
        let mut output = Element::new(MAIN_NAMESPACE, "cols");
        for pair in positions.windows(2) {
            let position = authored.partition_point(|(start, _, _)| *start <= pair[0]);
            let base = authored.get(position.wrapping_sub(1)).filter(|(_, end, _)| pair[0] <= *end);
            let projection = projection_at(&self.columns, pair[0]);
            if base.is_none() && projection.is_none() { continue; }
            let mut col = base.map(|(_, _, node)| node.clone()).unwrap_or_else(|| Element::new(MAIN_NAMESPACE, "col"));
            // xlsx_dimensions supplies an unqualified standalone fragment.
            // Its columns belong to the containing worksheet main namespace.
            col.namespace = MAIN_NAMESPACE.into();
            if col.attrs.get("xmlns").is_some_and(String::is_empty) { col.remove_attr("xmlns"); }
            col.set_attr("min", (pair[0] + 1).to_string()); col.set_attr("max", pair[1].to_string());
            if let Some(run) = projection {
                if run.level > 0 { col.set_attr("outlineLevel", run.level.to_string()); }
                if run.collapsed { col.set_attr("collapsed", "1"); }
                if run.hidden { col.set_attr("hidden", "1"); }
                if !run.extra_attributes.is_empty() {
                    let extras = parse(&format!("<col{}/>", run.extra_attributes)).map_err(ExportError::invalid)?;
                    for (name, value) in extras.attrs { col.set_attr_ns(extras.attribute_namespaces.get(&name).map(String::as_str).unwrap_or(""), &name, value); }
                }
            }
            output.push(col);
        }
        output.validate_fragment().map_err(ExportError::invalid)?;
        Ok(output.render())
    }

    pub fn row_indices(&self) -> impl Iterator<Item = u32> + '_ { self.rows.iter().flat_map(|run| run.start..=run.end) }
    pub fn row_hidden(&self, row: u32) -> bool { projection_at(&self.rows, row).is_some_and(|run| run.hidden) }
    pub fn row_attributes(&self, row: u32) -> String {
        let Some(run) = projection_at(&self.rows, row) else { return String::new(); };
        let mut output = run.extra_attributes.clone();
        if run.level > 0 { output.push_str(&format!(" outlineLevel=\"{}\"", run.level)); }
        if run.collapsed { output.push_str(" collapsed=\"1\""); }
        output
    }

    pub fn extension_xml(&self, sheet: &SheetInput, worksheet_prefix: &str) -> Result<String, ExportError> {
        if self.state.is_none() && self.source.is_none() { return Ok(String::new()); }
        let (outline_fingerprint, visibility_fingerprint) = worksheet_outline_fingerprints(worksheet_prefix).map_err(ExportError::invalid)?;
        let manual_visibility = ManualVisibility {
            hidden_rows: sheet.hidden_rows.clone(), hidden_columns: sheet.hidden_columns.clone(),
            row_hidden_spans: sheet.row_hidden_spans.iter().map(|span| HiddenSpan { start: span.start, end: span.end, value: span.value }).collect(),
            column_hidden_spans: sheet.column_hidden_spans.iter().map(|span| HiddenSpan { start: span.start, end: span.end, value: span.value }).collect(),
        };
        let mut extra = self.snapshot_extra.clone();
        extra.insert("projectionDiagnostics".into(), json!(self.warnings));
        let extension = OutlineExtension { version: 1, outline_fingerprint, visibility_fingerprint,
            state: self.state.clone().unwrap_or_default(), manual_visibility, source: self.source.clone(), extra };
        validate_extension(&extension).map_err(ExportError::invalid)?;
        let state = serde_json::to_string(&extension).map_err(|error| ExportError::invalid(error.to_string()))?;
        if state.len() > MAX_OUTLINE_JSON_BYTES { return Err(ExportError::invalid("outline snapshot exceeds the parser's 32 MiB budget")); }
        Ok(format!("<ext uri=\"{SHEET_OUTLINE_EXTENSION_URI}\"><oo:outline xmlns:oo=\"{SHEET_OUTLINE_NAMESPACE}\" version=\"1\">{}</oo:outline></ext>", esc(&state)))
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use office_io_common::sheet_outline_source::{read_worksheet_outline, infer_native_state};
    use std::io::Cursor;

    // Meaningful draft checks for the deferred shared Rust verification pass.
    fn input(value: Value) -> SheetInput { serde_json::from_value(value).unwrap() }
    fn fixture() -> SheetInput { input(json!({ "rowCount": 20, "columnCount": 10, "outline": {
        "summaryAbove": false, "summaryLeft": false, "futureState": { "tag": "state" }, "groups": [
            { "id": "outer", "axis": "row", "start": 1, "end": 8, "collapsed": false },
            { "id": "inner", "axis": "row", "start": 3, "end": 5, "collapsed": true, "futureGroup": { "tag": [1, 2] } },
            { "id": "col", "axis": "column", "start": 2, "end": 4, "collapsed": true } ] } })) }

    #[test]
    fn nested_detail_and_summary_rows_are_sorted_unique_and_keep_independent_hidden_state() {
        let plan = prepare(&fixture()).unwrap();
        assert_eq!(plan.row_indices().collect::<Vec<_>>(), (1..=8).collect::<Vec<_>>());
        assert!(!plan.row_hidden(2)); assert!(plan.row_hidden(3)); assert!(!plan.row_hidden(6));
        assert!(plan.row_attributes(3).contains("outlineLevel=\"2\""));
        assert!(plan.row_attributes(6).contains("collapsed=\"1\""));
        assert!(!plan.row_attributes(8).contains("collapsed="));
    }

    #[test]
    fn column_projection_splits_geometry_runs_and_keeps_main_namespace_widths_and_manual_hidden() {
        let plan = prepare(&fixture()).unwrap();
        let xml = plan.columns_xml("<cols><col min=\"1\" max=\"6\" width=\"20\" customWidth=\"1\" hidden=\"1\"/></cols>").unwrap();
        let columns = parse(&xml).unwrap();
        assert!(columns.children.iter().all(|column| column.namespace == MAIN_NAMESPACE));
        assert!(columns.children.iter().all(|column| column.attrs.get("width").map(String::as_str) == Some("20")));
        assert!(columns.children.iter().all(|column| column.attrs.get("hidden").map(String::as_str) == Some("1")));
        assert!(columns.children.iter().any(|column| column.attrs.get("min").map(String::as_str) == Some("3")
            && column.attrs.get("max").map(String::as_str) == Some("5") && column.attrs.get("outlineLevel").map(String::as_str) == Some("1")));
        let worksheet = format!("<worksheet xmlns=\"{MAIN_NAMESPACE}\">{xml}<sheetData/></worksheet>");
        let native = read_worksheet_outline(Cursor::new(worksheet.as_bytes()), false).unwrap();
        assert_eq!(infer_native_state(&native.source).groups.iter().filter(|group| group.axis == "column").count(), 1);
    }

    #[test]
    fn outline_properties_combine_with_filter_mode_in_one_native_container() {
        let mut sheet = fixture(); sheet.outline.as_mut().unwrap().summary_above = true;
        let properties = parse(&prepare(&sheet).unwrap().sheet_properties_xml(true).unwrap()).unwrap();
        assert_eq!(properties.name, "sheetPr"); assert_eq!(properties.children.len(), 1);
        assert_eq!(properties.attrs.get("filterMode").map(String::as_str), Some("1"));
        assert_eq!(properties.children[0].attrs.get("summaryBelow").map(String::as_str), Some("0"));
    }

    #[test]
    fn eight_equal_layers_are_native_clipped_but_the_snapshot_keeps_every_id_and_unknown_field() {
        let mut sheet = fixture();
        let state = sheet.outline.as_mut().unwrap(); state.groups = (0..8).map(|index| office_io_common::sheet_outline_source::SheetOutlineGroup {
            id: format!("layer-{index}"), axis: "row".into(), start: 1, end: 8, collapsed: true, extra: BTreeMap::new() }).collect();
        let plan = prepare(&sheet).unwrap();
        assert!(plan.warnings.iter().any(|warning| warning.starts_with("OUTLINE_DEPTH_SNAPSHOT")));
        assert!(plan.row_attributes(2).contains("outlineLevel=\"7\""));
        let format = plan.format_xml("<sheetFormatPr defaultRowHeight=\"20\"/>").unwrap();
        let mut prefix = format!("<worksheet xmlns=\"{MAIN_NAMESPACE}\">{}{}<sheetData>", plan.sheet_properties_xml(false).unwrap(), format);
        for row in plan.row_indices() { prefix.push_str(&format!("<row r=\"{}\"{} hidden=\"1\"/>", row + 1, plan.row_attributes(row))); }
        prefix.push_str("</sheetData>");
        let extension = plan.extension_xml(&sheet, &prefix).unwrap();
        let xml = format!("{prefix}<extLst>{extension}</extLst></worksheet>");
        let read = read_worksheet_outline(Cursor::new(xml.as_bytes()), false).unwrap();
        assert_eq!(read.extension.unwrap().state, sheet.outline.unwrap());
    }

    #[test]
    fn a_full_declared_axis_boundary_keeps_extent_without_creating_an_extra_summary_dimension() {
        let mut sheet = fixture(); let state = sheet.outline.as_mut().unwrap(); state.groups.truncate(1);
        state.groups[0].start = 0; state.groups[0].end = 19; state.groups[0].collapsed = true;
        let plan = prepare(&sheet).unwrap();
        assert_eq!(plan.row_indices().collect::<Vec<_>>(), (0..20).collect::<Vec<_>>());
        assert!(plan.warnings.iter().any(|warning| warning.starts_with("OUTLINE_BOUNDARY_SNAPSHOT")));
        assert!(plan.row_hidden(19)); assert_eq!(plan.row_attributes(20), "");
    }

    #[test]
    fn future_source_state_and_namespace_attributes_survive_prepare_and_export() {
        let source = json!({ "summaryAbove": false, "summaryLeft": false, "zeroHeight": false,
            "rowRuns": [{ "start": 1, "end": 1, "level": 1, "collapsed": false, "hidden": null,
                "attributes": { "{urn:future}flag": "f:value" }, "namespaces": { "f": "urn:future" } }], "columnRuns": [], "futureOfficeSource": { "tag": [1, 2] } });
        let mut sheet = fixture(); sheet.metadata = Some(json!({ "outlineOfficeSource": { "native": source, "futureImport": { "kept": true } } }));
        let plan = prepare(&sheet).unwrap();
        let row = parse(&format!("<row{}/>", plan.row_attributes(1))).unwrap();
        assert_eq!(row.attribute_ns("urn:future", "flag"), Some("f:value"));
        assert_eq!(plan.source.as_ref().unwrap().extra["futureOfficeSource"], json!({ "tag": [1, 2] }));
        assert_eq!(plan.source.as_ref().unwrap().extra["retainedImportMetadata"]["futureImport"], json!({ "kept": true }));
    }

    #[test]
    fn crossing_invalid_counts_and_oversized_depth_reject_before_any_projection() {
        let mut sheet = fixture(); sheet.outline.as_mut().unwrap().groups[1].end = 11;
        assert!(prepare(&sheet).is_err());
        let sheet = input(json!({ "rowCount": 0 })); assert!(prepare(&sheet).is_err());
        let mut sheet = fixture(); sheet.outline.as_mut().unwrap().groups = (0..9).map(|index| office_io_common::sheet_outline_source::SheetOutlineGroup {
            id: format!("layer-{index}"), axis: "row".into(), start: 1, end: 8, ..Default::default() }).collect();
        assert!(prepare(&sheet).is_err());
    }
}
