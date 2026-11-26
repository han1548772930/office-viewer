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
    let source = if candidate.is_null() { None } else {
        let source: NativeOutlineSource = serde_json::from_value(candidate.clone()).map_err(|error| ExportError::invalid(format!("invalid outline Office source: {error}")))?;
        validate_source(&source).map_err(ExportError::invalid)?;
        Some(source)
    };
    let extra = if let Some(snapshot) = raw.get("projectSnapshot").filter(|value| !value.is_null()) {
        let snapshot: OutlineExtension = serde_json::from_value(snapshot.clone()).map_err(|error| ExportError::invalid(format!("invalid retained outline snapshot: {error}")))?;
        snapshot.extra
    } else { BTreeMap::new() };
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
