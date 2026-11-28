//! Native SpreadsheetML outlines and an independently guarded project snapshot.
//! The scanner retains axis metadata, never worksheet cell contents. Fingerprints
//! detect accidental external edits; they are not authentication signatures.

use std::collections::{BTreeMap, BTreeSet};
use std::io::BufRead;

use quick_xml::{Reader, XmlVersion, events::{BytesStart, Event}};
use serde::{Deserialize, Serialize};
use serde_json::Value;

use crate::conditional_format_xml::{Content, Element, MAIN_NAMESPACE, STRICT_MAIN_NAMESPACE};
use crate::xml_namespaces::{XmlNamespaces, validate_declaration, xml_character};

pub const SHEET_OUTLINE_NAMESPACE: &str = "urn:office-studio:sheet-outline:1";
pub const SHEET_OUTLINE_EXTENSION_URI: &str = "{94802E78-595D-40AD-BA94-29DDF46ED673}";
pub const MAX_OUTLINE_JSON_BYTES: usize = 32 * 1024 * 1024;
pub const MAX_WORKSHEET_BYTES: u64 = 512 * 1024 * 1024;
pub const MAX_ROW_INDEX: u32 = 1_048_575;
pub const MAX_COLUMN_INDEX: u32 = 16_383;
pub const MAX_NATIVE_OUTLINE_LEVEL: u8 = 7;
pub const MAX_MODEL_OUTLINE_DEPTH: u8 = 8;
const MAX_SOURCE_RUNS: usize = 262_144;
const MAX_SOURCE_BYTES: usize = 32 * 1024 * 1024;

#[derive(Debug, Clone, Default, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct SheetOutlineGroup {
    #[serde(default)] pub id: String,
    pub axis: String,
    pub start: u32,
    pub end: u32,
    #[serde(default)] pub collapsed: bool,
    #[serde(default, flatten)] pub extra: BTreeMap<String, Value>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct SheetOutlineState {
    #[serde(default)] pub summary_above: bool,
    #[serde(default)] pub summary_left: bool,
    #[serde(default)] pub groups: Vec<SheetOutlineGroup>,
    #[serde(default, flatten)] pub extra: BTreeMap<String, Value>,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct HiddenSpan { pub start: u32, pub end: u32, pub value: bool }

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct ManualVisibility {
    #[serde(default)] pub hidden_rows: Vec<u32>,
    #[serde(default)] pub hidden_columns: Vec<u32>,
    #[serde(default)] pub row_hidden_spans: Vec<HiddenSpan>,
    #[serde(default)] pub column_hidden_spans: Vec<HiddenSpan>,
}

/// Expanded attribute names use `{namespace}local`; indices are stored once in
/// the run. Unknown authored attributes survive in the source and project JSON.
#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeOutlineRun {
    pub start: u32,
    pub end: u32,
    pub level: u8,
    pub collapsed: bool,
    pub hidden: Option<bool>,
    #[serde(default)] pub attributes: BTreeMap<String, String>,
    /// Authored bindings also retain prefixes appearing inside future QName
    /// attribute values, even when no known attribute uses that prefix.
    #[serde(default)] pub namespaces: BTreeMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOutlineSource {
    #[serde(default)] pub summary_above: bool,
    #[serde(default)] pub summary_left: bool,
    #[serde(default)] pub zero_height: bool,
    #[serde(default)] pub outline_pr_xml: Option<String>,
    #[serde(default)] pub row_runs: Vec<NativeOutlineRun>,
    #[serde(default)] pub column_runs: Vec<NativeOutlineRun>,
    #[serde(default)] pub diagnostics: Vec<String>,
    #[serde(default, flatten)] pub extra: BTreeMap<String, Value>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct OutlineExtension {
    pub version: u32,
    pub outline_fingerprint: String,
    pub visibility_fingerprint: String,
    pub state: SheetOutlineState,
    pub manual_visibility: ManualVisibility,
    #[serde(default)] pub source: Option<NativeOutlineSource>,
    #[serde(default, flatten)] pub extra: BTreeMap<String, Value>,
}

#[derive(Debug)]
pub struct WorksheetOutlineRead {
    pub source: NativeOutlineSource,
    pub extension: Option<OutlineExtension>,
    pub warnings: Vec<String>,
}

pub fn validate_state(state: &SheetOutlineState, row_count: u32, column_count: u32) -> Result<(), String> {
    if row_count == 0 || row_count > MAX_ROW_INDEX + 1 || column_count == 0 || column_count > MAX_COLUMN_INDEX + 1 {
        return Err("outline worksheet extent is outside the XLSX limits".into());
    }
    if state.groups.len() > MAX_SOURCE_RUNS { return Err("outline groups exceed the source budget".into()); }
    let mut ids = BTreeSet::new();
    for group in &state.groups {
        let maximum = match group.axis.as_str() { "row" => row_count - 1, "column" => column_count - 1,
            _ => return Err("outline group axis must be row or column".into()) };
        if group.id.trim().is_empty() || !ids.insert(&group.id) || group.start > group.end || group.end > maximum {
            return Err("outline group id or inclusive range is invalid".into());
        }
    }
    // A sweep validates containment and equal-range layers without quadratic
    // pair comparisons. Equal ranges are independent authored model layers.
    for axis in ["row", "column"] {
        let mut groups = state.groups.iter().filter(|group| group.axis == axis).collect::<Vec<_>>();
        groups.sort_by_key(|group| (group.start, std::cmp::Reverse(group.end), &group.id));
        let mut stack: Vec<&SheetOutlineGroup> = Vec::new();
        for group in groups {
            while stack.last().is_some_and(|parent| parent.end < group.start) { stack.pop(); }
            if stack.last().is_some_and(|parent| group.end > parent.end) { return Err("outline ranges cross".into()); }
            stack.push(group);
            if stack.len() > usize::from(MAX_MODEL_OUTLINE_DEPTH) { return Err("outline depth exceeds eight model groups".into()); }
        }
    }
    Ok(())
}

fn validate_spans(spans: &[HiddenSpan], maximum: u32) -> Result<(), String> {
    let mut previous = None;
    for span in spans {
        if span.start > span.end || span.end > maximum || previous.is_some_and(|end| span.start <= end) {
            return Err("manual hidden spans must be ordered, disjoint and inside the XLSX axis".into());
        }
        previous = Some(span.end);
    }
    Ok(())
}

pub fn validate_extension(extension: &OutlineExtension) -> Result<(), String> {
    let fingerprint = |value: &str| value.len() == 16 && value.bytes().all(|byte| byte.is_ascii_hexdigit());
    if extension.version != 1 || !fingerprint(&extension.outline_fingerprint) || !fingerprint(&extension.visibility_fingerprint) {
        return Err("unsupported outline snapshot version or fingerprint".into());
    }
    validate_state(&extension.state, MAX_ROW_INDEX + 1, MAX_COLUMN_INDEX + 1)?;
    let manual = &extension.manual_visibility;
    if manual.hidden_rows.iter().any(|row| *row > MAX_ROW_INDEX) || manual.hidden_columns.iter().any(|column| *column > MAX_COLUMN_INDEX) {
        return Err("manual hidden indices exceed the XLSX limits".into());
    }
    validate_spans(&manual.row_hidden_spans, MAX_ROW_INDEX)?;
    validate_spans(&manual.column_hidden_spans, MAX_COLUMN_INDEX)?;
    if let Some(source) = &extension.source { validate_source(source)?; }
    Ok(())
}

pub fn validate_source(source: &NativeOutlineSource) -> Result<(), String> {
    for (runs, maximum) in [(&source.row_runs, MAX_ROW_INDEX), (&source.column_runs, MAX_COLUMN_INDEX)] {
        let mut previous = None;
        if runs.len() > MAX_SOURCE_RUNS { return Err("native outline source exceeds its run budget".into()); }
        for run in runs {
            if run.start > run.end || run.end > maximum || run.level > MAX_MODEL_OUTLINE_DEPTH || previous.is_some_and(|end| run.start <= end) {
                return Err("native outline runs must be ordered, disjoint and inside the XLSX axis".into());
            }
            previous = Some(run.end);
        }
    }
    if let Some(xml) = &source.outline_pr_xml {
        let node = crate::conditional_format_xml::parse(xml)?;
        if !node.is_main("outlinePr") || !matches!(node.namespace.as_str(), MAIN_NAMESPACE | STRICT_MAIN_NAMESPACE) {
            return Err("outline source fragment must be a standard outlinePr".into());
        }
    }
    if serde_json::to_vec(source).map_err(|error| error.to_string())?.len() > MAX_SOURCE_BYTES {
        return Err("native outline source exceeds its 32 MiB budget".into());
    }
    Ok(())
}

pub fn run_at(runs: &[NativeOutlineRun], index: u32) -> Option<&NativeOutlineRun> {
    let position = runs.partition_point(|run| run.start <= index);
    runs.get(position.wrapping_sub(1)).filter(|run| index <= run.end)
}

/// Native levels express nested ranges, but a shared summary marker cannot
/// identify every collapsed layer. Infer only the level immediately above the
/// marker and retain the complete source for all ambiguous hidden overlaps.
pub fn infer_native_state(source: &NativeOutlineSource) -> SheetOutlineState {
    let mut state = SheetOutlineState { summary_above: source.summary_above, summary_left: source.summary_left, ..SheetOutlineState::default() };
    for (axis, runs, at_start, maximum) in [("row", &source.row_runs, source.summary_above, MAX_ROW_INDEX),
        ("column", &source.column_runs, source.summary_left, MAX_COLUMN_INDEX)] {
        for level in 1..=MAX_MODEL_OUTLINE_DEPTH {
            let mut ranges = Vec::<(u32, u32)>::new();
            for run in runs.iter().filter(|run| run.level >= level) {
                if let Some(previous) = ranges.last_mut() {
                    if previous.1 + 1 == run.start { previous.1 = run.end; continue; }
                }
                ranges.push((run.start, run.end));
            }
            for (start, end) in ranges {
                let marker = if at_start { start.checked_sub(1) } else { end.checked_add(1).filter(|index| *index <= maximum) };
                let collapsed = marker.and_then(|index| run_at(runs, index)).is_some_and(|run| run.collapsed && run.level + 1 == level)
                    && range_is_hidden(runs, start, end, if axis == "row" { source.zero_height } else { false });
                state.groups.push(SheetOutlineGroup { id: format!("xlsx-outline-{axis}-{level}-{start}-{end}"), axis: axis.into(), start, end, collapsed, extra: BTreeMap::new() });
            }
        }
    }
    state
}

fn range_is_hidden(runs: &[NativeOutlineRun], start: u32, end: u32, default_hidden: bool) -> bool {
    let mut cursor = start;
    for run in runs.iter().filter(|run| run.end >= start && run.start <= end) {
        if run.start > cursor && !default_hidden || !run.hidden.unwrap_or(default_hidden) { return false; }
        cursor = cursor.max(run.end + 1);
    }
    cursor > end || default_hidden
}

fn hash(bytes: &[u8]) -> String {
    let mut value = 0xcbf29ce484222325u64;
    for byte in bytes { value ^= u64::from(*byte); value = value.wrapping_mul(0x100000001b3); }
    format!("{value:016x}")
}

fn canonical_runs(runs: &[NativeOutlineRun], visibility: bool, default_hidden: bool) -> Vec<(u32, u32, u8, bool)> {
    let mut result: Vec<(u32, u32, u8, bool)> = Vec::new();
    for run in runs {
        let (level, flag) = if visibility { (0, run.hidden.unwrap_or(default_hidden)) } else { (run.level, run.collapsed) };
        if visibility && flag == default_hidden || !visibility && level == 0 && !flag { continue; }
        if let Some(previous) = result.last_mut() {
            if previous.1 + 1 == run.start && previous.2 == level && previous.3 == flag { previous.1 = run.end; continue; }
        }
        result.push((run.start, run.end, level, flag));
    }
    result
}

impl NativeOutlineSource {
    pub fn fingerprints(&self) -> Result<(String, String), String> {
        let outline = serde_json::to_vec(&(self.summary_above, self.summary_left,
            canonical_runs(&self.row_runs, false, false), canonical_runs(&self.column_runs, false, false))).map_err(|error| error.to_string())?;
        let visibility = serde_json::to_vec(&(self.zero_height, canonical_runs(&self.row_runs, true, self.zero_height),
            canonical_runs(&self.column_runs, true, false))).map_err(|error| error.to_string())?;
        Ok((hash(&outline), hash(&visibility)))
    }
}

fn attributes(element: &BytesStart<'_>, namespaces: &XmlNamespaces) -> Result<BTreeMap<String, String>, String> {
    let mut result = BTreeMap::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| error.to_string())?;
        if attribute.key.as_namespace_binding().is_some() { continue; }
        let name = std::str::from_utf8(attribute.key.local_name().as_ref().as_bytes()).map_err(|error| error.to_string())?.to_owned();
        let namespace = namespaces.namespace(attribute.key, true)?.unwrap_or("");
        let key = if namespace.is_empty() { name } else { format!("{{{namespace}}}{name}") };
        let value = attribute.normalized_value_with(XmlVersion::Implicit1_0, 1, quick_xml::escape::resolve_xml_entity)
            .map_err(|error| error.to_string())?.into_owned();
        if !value.chars().all(xml_character) || result.insert(key, value).is_some() { return Err("invalid or duplicate expanded outline attribute".into()); }
    }
    Ok(result)
}

/// Render expanded names through the shared XML tree, which escapes values and
/// declares namespaces locally. No source string is interpolated into XML names.
pub fn node_with_attributes(namespace: &str, name: &str, attributes: &BTreeMap<String, String>) -> Result<Element, String> {
    let mut node = Element::new(namespace, name);
    let mut prefixes = BTreeMap::<String, String>::new();
    for (key, value) in attributes {
        if let Some(tail) = key.strip_prefix('{') {
            let (namespace, local) = tail.split_once('}').ok_or_else(|| "invalid expanded outline attribute name".to_owned())?;
            let next = format!("osn{}", prefixes.len());
            let prefix = prefixes.entry(namespace.into()).or_insert(next).clone();
            node.set_attr_ns(namespace, &format!("{prefix}:{local}"), value);
        } else { node.set_attr(key, value); }
    }
    node.validate_fragment()?;
    // parse validates XML names and expanded attributes before any re-emission.
    crate::conditional_format_xml::parse(&node.render())
}

fn flag(value: Option<&String>, default: bool) -> Result<bool, String> {
    match value.map(String::as_str) { None => Ok(default), Some("1" | "true") => Ok(true), Some("0" | "false") => Ok(false),
        _ => Err("invalid native outline boolean".into()) }
}

fn append_run(runs: &mut Vec<NativeOutlineRun>, run: NativeOutlineRun, source_bytes: &mut usize) -> Result<(), String> {
    if let Some(previous) = runs.last_mut() {
        if previous.end >= run.start { return Err("native outline axis elements overlap or are out of order".into()); }
        if previous.end + 1 == run.start && previous.level == run.level && previous.collapsed == run.collapsed
            && previous.hidden == run.hidden && previous.attributes == run.attributes && previous.namespaces == run.namespaces { previous.end = run.end; return Ok(()); }
    }
    if runs.len() >= MAX_SOURCE_RUNS { return Err("native outline run count exceeds the parser budget".into()); }
    *source_bytes = source_bytes.saturating_add(serde_json::to_vec(&run).map_err(|error| error.to_string())?.len());
    if *source_bytes > MAX_SOURCE_BYTES { return Err("native outline source exceeds its byte budget".into()); }
    runs.push(run); Ok(())
}

fn native_run(mut attrs: BTreeMap<String, String>, bindings: &BTreeMap<String, String>, is_row: bool, start: u32, end: u32, warnings: &mut Vec<String>) -> Result<Option<NativeOutlineRun>, String> {
    let raw_level = attrs.get("outlineLevel").map(|value| value.parse::<u32>()).transpose().map_err(|_| "invalid native outline level")?.unwrap_or(0);
    if raw_level > u32::from(MAX_MODEL_OUTLINE_DEPTH) { return Err("native outline level exceeds the complete eight-layer model; cannot restore without dropping groups".into()); }
    let level = raw_level as u8;
    if raw_level > u32::from(MAX_NATIVE_OUTLINE_LEVEL) && !warnings.iter().any(|warning| warning.starts_with("NATIVE_LEVEL_OUTSIDE_OFFICE")) {
        warnings.push("NATIVE_LEVEL_OUTSIDE_OFFICE: source exceeds Office level seven; its complete eighth model layer and authored attributes are retained".into());
    }
    let collapsed = flag(attrs.get("collapsed"), false)?;
    let hidden = attrs.get("hidden").map(|value| flag(Some(value), false)).transpose()?;
    let managed: &[&str] = if is_row { &["r", "ht", "customHeight", "s", "customFormat", "spans", "thickTop", "thickBot", "ph"] }
        else { &["min", "max", "width", "customWidth", "style", "bestFit", "phonetic"] };
    for name in managed { attrs.remove(*name); }
    if level == 0 && !collapsed && hidden.is_none() && attrs.is_empty() { return Ok(None); }
    let namespaces = if attrs.keys().any(|name| !matches!(name.as_str(), "outlineLevel" | "collapsed" | "hidden")) { bindings.clone() } else { BTreeMap::new() };
    Ok(Some(NativeOutlineRun { start, end, level, collapsed, hidden, attributes: attrs, namespaces }))
}

fn source_node(tag: &BytesStart<'_>, namespaces: &XmlNamespaces, bindings: &BTreeMap<String, String>) -> Result<Element, String> {
    let namespace = namespaces.namespace(tag.name(), false)?.unwrap_or("");
    let name = std::str::from_utf8(tag.local_name().as_ref().as_bytes()).map_err(|error| error.to_string())?.to_owned();
    let mut node = Element::new(namespace, &name);
    node.qualified_name = std::str::from_utf8(tag.name().as_ref().as_bytes()).map_err(|error| error.to_string())?.to_owned();
    for (prefix, namespace) in bindings {
        node.set_attr(&if prefix.is_empty() { "xmlns".into() } else { format!("xmlns:{prefix}") }, namespace);
    }
    for attribute in tag.attributes() {
        let attribute = attribute.map_err(|error| error.to_string())?;
        if attribute.key.as_namespace_binding().is_some() { continue; }
        let key = std::str::from_utf8(attribute.key.as_ref().as_bytes()).map_err(|error| error.to_string())?;
        let value = attribute.normalized_value_with(XmlVersion::Implicit1_0, 1, quick_xml::escape::resolve_xml_entity).map_err(|error| error.to_string())?;
        node.set_attr_ns(namespaces.namespace(attribute.key, true)?.unwrap_or(""), key, value.into_owned());
    }
    Ok(node)
}

struct JsonCapture { depth: usize, text: String, invalid: bool }

fn parse_snapshot(text: &str) -> Result<OutlineExtension, String> {
    let raw: Value = serde_json::from_str(text).map_err(|error| format!("invalid outline JSON: {error}"))?;
    let state = raw.get("state").and_then(Value::as_object).ok_or_else(|| "outline snapshot is missing its complete state".to_owned())?;
    let groups = state.get("groups").and_then(Value::as_array).ok_or_else(|| "outline snapshot is missing its groups array".to_owned())?;
    if !state.get("summaryAbove").is_some_and(Value::is_boolean) || !state.get("summaryLeft").is_some_and(Value::is_boolean)
        || groups.iter().any(|group| !group.get("collapsed").is_some_and(Value::is_boolean)) {
        return Err("outline snapshot is missing authored summary or collapse flags".into());
    }
    let manual = raw.get("manualVisibility").and_then(Value::as_object).ok_or_else(|| "outline snapshot is missing manual visibility".to_owned())?;
    if ["hiddenRows", "hiddenColumns", "rowHiddenSpans", "columnHiddenSpans"].iter().any(|key| !manual.get(*key).is_some_and(Value::is_array)) {
        return Err("outline snapshot is missing authored manual hidden arrays or spans".into());
    }
    let extension: OutlineExtension = serde_json::from_value(raw).map_err(|error| format!("invalid outline JSON: {error}"))?;
    validate_extension(&extension)?;
    Ok(extension)
}

/// `allow_prefix` is exclusively for the writer after sheetData and before
/// extLst. The importer requires one closed standard worksheet document.
pub fn read_worksheet_outline<R: BufRead>(input: R, allow_prefix: bool) -> Result<WorksheetOutlineRead, String> {
    let mut reader = Reader::from_reader(input);
    let mut namespaces = XmlNamespaces::default();
    let mut path = Vec::<(String, String)>::new();
    let mut buffer = Vec::new();
    let mut source = NativeOutlineSource::default();
    let mut source_bytes = 0usize;
    let mut bindings = BTreeMap::<String, String>::new();
    let mut binding_scopes = Vec::<Vec<(String, Option<String>)>>::new();
    let mut root_seen = false;
    let mut next_row = 0u32;
    let mut previous_row = None;
    let mut tree = Vec::<Element>::new();
    let mut tree_bytes = 0usize;
    let mut tree_nodes = 0usize;
    let mut outline_pr_count = 0usize;
    let mut extension_depth = None;
    let mut extension_containers = 0usize;
    let mut extension_count = 0usize;
    let mut capture: Option<JsonCapture> = None;
    let mut candidate = None;
    let warnings = Vec::new();
    loop {
        let event = reader.read_event_into(&mut buffer).map_err(|error| error.to_string())?;
        match event {
            Event::Start(ref tag) | Event::Empty(ref tag) => {
                let empty = matches!(&event, Event::Empty(_));
                if path.len() >= 128 { return Err("outline worksheet nesting exceeds 128".into()); }
                namespaces.push(tag)?;
                let mut scope = Vec::new();
                for attribute in tag.attributes() {
                    let attribute = attribute.map_err(|error| error.to_string())?;
                    let key = std::str::from_utf8(attribute.key.as_ref().as_bytes()).map_err(|error| error.to_string())?;
                    let prefix = if key == "xmlns" { Some("") } else { key.strip_prefix("xmlns:") };
                    if let Some(prefix) = prefix {
                        let value = attribute.normalized_value_with(XmlVersion::Implicit1_0, 1, quick_xml::escape::resolve_xml_entity).map_err(|error| error.to_string())?.into_owned();
                        scope.push((prefix.to_owned(), bindings.insert(prefix.to_owned(), value)));
                    }
                }
                let namespace = namespaces.namespace(tag.name(), false)?.unwrap_or("").to_owned();
                let name = std::str::from_utf8(tag.local_name().as_ref().as_bytes()).map_err(|error| error.to_string())?.to_owned();
                let main = matches!(namespace.as_str(), MAIN_NAMESPACE | STRICT_MAIN_NAMESPACE);
                let attrs = attributes(tag, &namespaces)?;
                if path.is_empty() {
                    if root_seen || !main || name != "worksheet" { return Err("invalid or duplicate outline worksheet root".into()); }
                    root_seen = true;
                }
                let standard_parent = |names: &[&str]| path.len() == names.len() && path.iter().zip(names)
                    .all(|((namespace, name), expected)| matches!(namespace.as_str(), MAIN_NAMESPACE | STRICT_MAIN_NAMESPACE) && name == expected);
                if capture.is_some() { capture.as_mut().unwrap().invalid = true; }
                if main && name == "sheetFormatPr" && standard_parent(&["worksheet"]) { source.zero_height = flag(attrs.get("zeroHeight"), false)?; }
                if main && name == "outlinePr" && standard_parent(&["worksheet", "sheetPr"]) {
                    outline_pr_count += 1;
                    if outline_pr_count > 1 { return Err("duplicate standard outlinePr".into()); }
                    source.summary_above = !flag(attrs.get("summaryBelow"), true)?;
                    source.summary_left = !flag(attrs.get("summaryRight"), true)?;
                    let node = source_node(tag, &namespaces, &bindings)?;
                    tree_bytes += tag.len();
                    tree_nodes += 1;
                    if empty { source.outline_pr_xml = Some(node.render()); } else { tree.push(node); }
                } else if !tree.is_empty() {
                    tree_bytes = tree_bytes.saturating_add(tag.len());
                    tree_nodes += 1;
                    let node = source_node(tag, &namespaces, &bindings)?;
                    if empty { tree.last_mut().unwrap().push(node); } else { tree.push(node); }
                }
                if main && name == "row" && standard_parent(&["worksheet", "sheetData"]) {
                    let row = attrs.get("r").map(|value| value.parse::<u32>()).transpose().map_err(|_| "invalid outline row number")?
                        .map(|value| value.checked_sub(1).ok_or_else(|| "outline row number is zero".to_owned())).transpose()?.unwrap_or(next_row);
                    if row > MAX_ROW_INDEX || previous_row.is_some_and(|previous| row <= previous) { return Err("outline row numbers are invalid or out of order".into()); }
                    previous_row = Some(row); next_row = row + 1;
                    if let Some(run) = native_run(attrs.clone(), &bindings, true, row, row, &mut source.diagnostics)? { append_run(&mut source.row_runs, run, &mut source_bytes)?; }
                }
                if main && name == "col" && standard_parent(&["worksheet", "cols"]) {
                    let index = |key: &str| attrs.get(key).ok_or_else(|| format!("outline column lacks {key}"))?.parse::<u32>()
                        .map_err(|_| "invalid outline column number".to_owned())?.checked_sub(1).ok_or_else(|| "outline column number is zero".to_owned());
                    let (start, end) = (index("min")?, index("max")?);
                    if start > end || end > MAX_COLUMN_INDEX { return Err("outline columns exceed the XLSX limits".into()); }
                    if let Some(run) = native_run(attrs.clone(), &bindings, false, start, end, &mut source.diagnostics)? { append_run(&mut source.column_runs, run, &mut source_bytes)?; }
                }
                if main && name == "ext" && standard_parent(&["worksheet", "extLst"]) && attrs.get("uri").map(String::as_str) == Some(SHEET_OUTLINE_EXTENSION_URI) {
                    extension_containers += 1;
                    extension_depth = (!empty).then_some(path.len() + 1);
                }
                if namespace == SHEET_OUTLINE_NAMESPACE && name == "outline" && extension_depth == Some(path.len())
                    && standard_parent(&["worksheet", "extLst", "ext"]) {
                    extension_count += 1;
                    if empty { return Err("empty authored outline snapshot; native fallback would lose complete model state".into()); }
                    capture = Some(JsonCapture { depth: path.len() + 1, text: String::new(), invalid: attrs.get("version").map(String::as_str) != Some("1") });
                }
                if empty { restore_bindings(&mut bindings, scope); namespaces.pop(); } else { binding_scopes.push(scope); path.push((namespace, name)); }
            }
            Event::End(_) => {
                if capture.as_ref().is_some_and(|capture| capture.depth == path.len()) {
                    let completed = capture.take().unwrap();
                    let parsed = if completed.invalid { Err("invalid nested outline snapshot markup or version".into()) } else {
                        parse_snapshot(&completed.text)
                    };
                    match parsed { Ok(state) => candidate = Some(state), Err(warning) => return Err(warning) }
                }
                if !tree.is_empty() {
                    let node = tree.pop().unwrap();
                    if let Some(parent) = tree.last_mut() { parent.push(node); } else { node.validate_fragment()?; source.outline_pr_xml = Some(node.render()); }
                }
                if extension_depth == Some(path.len()) { extension_depth = None; }
                path.pop().ok_or_else(|| "unmatched outline worksheet closing tag".to_owned())?;
                restore_bindings(&mut bindings, binding_scopes.pop().ok_or_else(|| "unmatched outline namespace scope".to_owned())?);
                namespaces.pop();
            }
            Event::Text(text) => { let value = text.xml_content(XmlVersion::Implicit1_0);
                if path.is_empty() && !value.trim().is_empty() { return Err("text outside the outline worksheet root".into()); }
                append_text(&mut capture, &mut tree, &value, &mut tree_bytes)?; }
            Event::CData(text) => { let value = text.xml_content(XmlVersion::Implicit1_0);
                if path.is_empty() { return Err("CDATA outside the outline worksheet root".into()); }
                append_text(&mut capture, &mut tree, &value, &mut tree_bytes)?; }
            Event::GeneralRef(reference) => {
                let value = if let Some(character) = reference.resolve_char_ref().map_err(|error| error.to_string())? { character.to_string() }
                    else { let name = reference.as_ref();
                        quick_xml::escape::resolve_xml_entity(&name).ok_or_else(|| "unknown outline XML entity".to_owned())?.to_owned() };
                append_text(&mut capture, &mut tree, &value, &mut tree_bytes)?;
            }
            Event::Comment(comment) => {
                if let Some(node) = tree.last_mut() { let value = comment.as_ref().to_owned();
                    tree_bytes += value.len(); node.content.push(Content::Comment(value)); }
            }
            Event::PI(instruction) => { if let Some(node) = tree.last_mut() {
                let value = instruction.as_ref().to_owned();
                tree_bytes += value.len(); node.content.push(Content::Instruction(value));
            } }
            Event::Decl(declaration) => { if root_seen { return Err("misplaced outline XML declaration".into()); } validate_declaration(declaration.as_ref())?; }
            Event::DocType(_) => return Err("outline worksheet DTDs are unsupported".into()),
            Event::Eof => break,
        }
        if reader.buffer_position() > MAX_WORKSHEET_BYTES || tree_bytes > 16 * 1024 * 1024 || tree_nodes > 262_144 { return Err("outline worksheet or fragment exceeds the parser budget".into()); }
        buffer.clear();
    }
    if !root_seen || !tree.is_empty() || capture.is_some() || !(path.is_empty() || allow_prefix && path.len() == 1 && path[0].1 == "worksheet") {
        return Err("incomplete outline worksheet XML".into());
    }
    if extension_containers > 0 && (extension_containers != 1 || extension_count != 1 || candidate.is_none()) {
        return Err("duplicate, unsupported or incomplete authored outline snapshots; native fallback would lose complete model state".into());
    }
    validate_source(&source)?;
    Ok(WorksheetOutlineRead { source, extension: candidate, warnings })
}

fn restore_bindings(bindings: &mut BTreeMap<String, String>, scope: Vec<(String, Option<String>)>) {
    for (prefix, previous) in scope.into_iter().rev() { match previous { Some(value) => { bindings.insert(prefix, value); }, None => { bindings.remove(&prefix); } } }
}

fn append_text(capture: &mut Option<JsonCapture>, tree: &mut [Element], value: &str, tree_bytes: &mut usize) -> Result<(), String> {
    if !value.chars().all(xml_character) { return Err("invalid XML character in outline metadata".into()); }
    if let Some(capture) = capture {
        if capture.text.len().saturating_add(value.len()) > MAX_OUTLINE_JSON_BYTES { capture.invalid = true; capture.text.clear(); }
        if !capture.invalid { capture.text.push_str(value); }
    }
    if let Some(node) = tree.last_mut() { *tree_bytes = tree_bytes.saturating_add(value.len()); node.text.push_str(value); node.content.push(Content::Text(value.into())); }
    Ok(())
}

pub fn worksheet_outline_fingerprints(xml: &str) -> Result<(String, String), String> {
    read_worksheet_outline(std::io::Cursor::new(xml.as_bytes()), true)?.source.fingerprints()
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    // Contract drafts for the deferred unified test pass; not run in this batch.
    fn worksheet(body: &str) -> String { format!("<worksheet xmlns=\"{MAIN_NAMESPACE}\">{body}</worksheet>") }
    fn read(xml: &str) -> WorksheetOutlineRead { read_worksheet_outline(std::io::Cursor::new(xml.as_bytes()), false).unwrap() }
    fn group(id: &str, axis: &str, start: u32, end: u32) -> SheetOutlineGroup {
        SheetOutlineGroup { id: id.into(), axis: axis.into(), start, end, ..SheetOutlineGroup::default() }
    }
    fn native_fixture() -> String { worksheet(concat!(
        "<sheetPr><outlinePr summaryBelow=\"1\" summaryRight=\"0\" applyStyles=\"1\"/></sheetPr>",
        "<cols><col min=\"1\" max=\"1\" collapsed=\"1\"/><col min=\"2\" max=\"4\" outlineLevel=\"1\" hidden=\"1\"/></cols>",
        "<sheetData><row r=\"2\" outlineLevel=\"1\" hidden=\"1\"/><row r=\"3\" outlineLevel=\"2\" hidden=\"1\"/>",
        "<row r=\"4\" outlineLevel=\"2\" hidden=\"1\"/><row r=\"5\" outlineLevel=\"1\" collapsed=\"1\" hidden=\"1\"/>",
        "<row r=\"6\" outlineLevel=\"1\" hidden=\"1\"/><row r=\"7\" collapsed=\"1\"/></sheetData>")) }
    fn snapshot(xml: &str) -> OutlineExtension {
        let source = read(xml).source;
        let (outline_fingerprint, visibility_fingerprint) = source.fingerprints().unwrap();
        OutlineExtension { version: 1, outline_fingerprint, visibility_fingerprint, state: infer_native_state(&source),
            manual_visibility: ManualVisibility::default(), source: None, extra: BTreeMap::new() }
    }
    fn with_snapshot(xml: &str, extension: &OutlineExtension, copies: usize) -> String {
        let json = quick_xml::escape::escape(&serde_json::to_string(extension).unwrap()).into_owned();
        let fragment = format!("<ext uri=\"{SHEET_OUTLINE_EXTENSION_URI}\"><oo:outline xmlns:oo=\"{SHEET_OUTLINE_NAMESPACE}\" version=\"1\">{json}</oo:outline></ext>");
        xml.replace("</worksheet>", &format!("<extLst>{}</extLst></worksheet>", fragment.repeat(copies)))
    }

    #[test]
    fn native_nested_rows_columns_markers_and_summary_directions_form_real_groups() {
        let source = read(&native_fixture()).source;
        let state = infer_native_state(&source);
        assert!(!state.summary_above); assert!(state.summary_left);
        assert_eq!(state.groups.iter().map(|group| (group.axis.as_str(), group.start, group.end, group.collapsed)).collect::<Vec<_>>(),
            [("row", 1, 5, true), ("row", 2, 3, true), ("column", 1, 3, true)]);
        assert!(source.outline_pr_xml.as_ref().unwrap().contains("applyStyles=\"1\""));
    }

    #[test]
    fn shared_marker_identifies_only_the_adjacent_outline_level() {
        let source = read(&worksheet("<sheetData><row r=\"2\" outlineLevel=\"2\" hidden=\"1\"/><row r=\"3\" collapsed=\"1\"/></sheetData>")).source;
        let groups = infer_native_state(&source).groups;
        assert_eq!(groups.len(), 2);
        assert!(groups[0].collapsed); assert!(!groups[1].collapsed);
        assert_eq!((groups[0].start, groups[0].end), (groups[1].start, groups[1].end));
    }

    #[test]
    fn collapsed_marker_with_visible_detail_does_not_hide_authored_visible_content() {
        let source = read(&worksheet("<sheetData><row r=\"2\" outlineLevel=\"1\"/><row r=\"3\" collapsed=\"1\"/></sheetData>")).source;
        assert!(!infer_native_state(&source).groups[0].collapsed);
        assert!(run_at(&source.row_runs, 2).unwrap().collapsed);
    }

    #[test]
    fn fingerprints_ignore_cells_styles_prefixes_attribute_order_and_default_false_flags() {
        let first = worksheet("<sheetPr><outlinePr summaryBelow=\"1\"/></sheetPr><sheetData><row r=\"2\" outlineLevel=\"1\" hidden=\"1\"><c r=\"A2\"><v>4</v></c></row></sheetData>");
        let second = format!("<s:worksheet xmlns:s=\"{STRICT_MAIN_NAMESPACE}\"><s:sheetPr><s:outlinePr summaryRight=\"true\" summaryBelow=\"true\"/></s:sheetPr><s:sheetData><s:row hidden=\"true\" s=\"2\" outlineLevel=\"01\" collapsed=\"false\" r=\"2\"><s:c r=\"A2\"><s:v>99</s:v></s:c></s:row></s:sheetData></s:worksheet>");
        assert_eq!(read(&first).source.fingerprints().unwrap(), read(&second).source.fingerprints().unwrap());
    }

    #[test]
    fn hidden_edits_invalidate_visibility_without_invalidating_the_outline_model() {
        let xml = native_fixture();
        let changed = xml.replacen("hidden=\"1\"", "hidden=\"0\"", 1);
        let before = read(&xml).source.fingerprints().unwrap();
        let after = read(&changed).source.fingerprints().unwrap();
        assert_eq!(before.0, after.0); assert_ne!(before.1, after.1);
    }

    #[test]
    fn native_zero_height_visible_exceptions_change_only_the_visibility_projection() {
        let before = worksheet("<sheetFormatPr zeroHeight=\"1\"/><sheetData><row r=\"3\" hidden=\"0\"/></sheetData>");
        let after = before.replace("hidden=\"0\"", "hidden=\"1\"");
        let first = read(&before).source.fingerprints().unwrap(); let second = read(&after).source.fingerprints().unwrap();
        assert_eq!(first.0, second.0); assert_ne!(first.1, second.1);
    }

    #[test]
    fn exact_snapshot_preserves_future_state_group_and_snapshot_fields() {
        let xml = native_fixture();
        let mut extension = snapshot(&xml);
        extension.extra.insert("futureSnapshot".into(), json!({ "tag": [1, 2] }));
        extension.state.extra.insert("futureState".into(), json!({ "version": 2 }));
        extension.state.groups[0].extra.insert("futureGroup".into(), json!({ "label": "A & B" }));
        let restored = read(&with_snapshot(&xml, &extension, 1)).extension.unwrap();
        assert_eq!(restored.state, extension.state); assert_eq!(restored.extra, extension.extra);
        assert_eq!(restored.outline_fingerprint, read(&xml).source.fingerprints().unwrap().0);
    }

    #[test]
    fn duplicate_or_nested_snapshot_markup_fails_without_returning_partial_state() {
        let xml = native_fixture();
        let extension = snapshot(&xml);
        let duplicate = with_snapshot(&xml, &extension, 2);
        assert!(read_worksheet_outline(std::io::Cursor::new(duplicate.as_bytes()), false).is_err());
        let invalid = with_snapshot(&xml, &extension, 1).replace("</oo:outline>", "<oo:bad/></oo:outline>");
        assert!(read_worksheet_outline(std::io::Cursor::new(invalid.as_bytes()), false).is_err());
    }

    #[test]
    fn foreign_local_names_cannot_spoof_native_outline_properties_or_rows() {
        let xml = worksheet("<sheetPr><f:outlinePr xmlns:f=\"urn:foreign\" summaryBelow=\"0\"/></sheetPr><sheetData><f:row xmlns:f=\"urn:foreign\" r=\"2\" outlineLevel=\"3\" hidden=\"1\"/></sheetData>");
        let source = read(&xml).source;
        assert!(source.outline_pr_xml.is_none()); assert!(source.row_runs.is_empty()); assert!(!source.summary_above);
    }

    #[test]
    fn unknown_outline_xml_preserves_expanded_names_qname_bindings_and_mixed_content() {
        let xml = format!("<worksheet xmlns=\"{MAIN_NAMESPACE}\" xmlns:f=\"urn:future\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"><sheetPr><outlinePr mc:Ignorable=\"f\" f:owner=\"f:team\">left<f:node f:value=\"a&amp;b\"/>right</outlinePr></sheetPr><sheetData><row r=\"2\" outlineLevel=\"1\" f:flag=\"f:row\"/></sheetData></worksheet>");
        let source = read(&xml).source;
        let node = crate::conditional_format_xml::parse(source.outline_pr_xml.as_ref().unwrap()).unwrap();
        assert_eq!(node.attribute_ns("urn:future", "owner"), Some("f:team"));
        assert_eq!(node.children[0].namespace, "urn:future"); assert_eq!(node.text, "leftright");
        assert_eq!(source.row_runs[0].attributes.get("{urn:future}flag").map(String::as_str), Some("f:row"));
        assert_eq!(source.row_runs[0].namespaces.get("f").map(String::as_str), Some("urn:future"));
    }

    #[test]
    fn namespace_escapes_are_resolved_and_invalid_documents_reject_before_restore() {
        let escaped = native_fixture().replace("spreadsheetml/2006/main", "spreadsheetml/2006/ma&#105;n");
        assert_eq!(infer_native_state(&read(&escaped).source).groups.len(), 3);
        for xml in ["<!DOCTYPE worksheet [<!ENTITY x 'bad'>]><worksheet/>",
            "<worksheet xmlns='urn:foreign'/>", "<worksheet xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'><x:row/></worksheet>",
            "<worksheet xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'><sheetData></worksheet>"] {
            assert!(read_worksheet_outline(std::io::Cursor::new(xml.as_bytes()), false).is_err(), "{xml}");
        }
    }

    #[test]
    fn only_the_writer_may_fingerprint_a_closed_content_prefix() {
        let xml = native_fixture(); let prefix = xml.strip_suffix("</worksheet>").unwrap();
        assert!(read_worksheet_outline(std::io::Cursor::new(prefix.as_bytes()), false).is_err());
        assert_eq!(worksheet_outline_fingerprints(prefix).unwrap(), read(&xml).source.fingerprints().unwrap());
        let incomplete = worksheet("<sheetData>").replace("</worksheet>", "");
        assert!(worksheet_outline_fingerprints(&incomplete).is_err());
    }

    #[test]
    fn ordered_axes_compress_identical_adjacent_rows_without_expanding_columns() {
        let mut data = String::from("<cols><col min=\"1\" max=\"16384\" outlineLevel=\"1\"/></cols><sheetData>");
        for row in 1..=100 { data.push_str(&format!("<row r=\"{row}\" outlineLevel=\"1\" hidden=\"1\"/>")); }
        data.push_str("</sheetData>");
        let source = read(&worksheet(&data)).source;
        assert_eq!(source.row_runs.len(), 1); assert_eq!(source.row_runs[0].end, 99);
        assert_eq!(source.column_runs.len(), 1); assert_eq!(source.column_runs[0].end, MAX_COLUMN_INDEX);
    }

    #[test]
    fn eight_equal_range_model_layers_are_preserved_and_the_ninth_is_rejected() {
        let mut state = SheetOutlineState::default();
        for index in 0..8 { state.groups.push(group(&format!("g{index}"), "row", 1, 5)); }
        validate_state(&state, 20, 10).unwrap();
        state.groups.push(group("ninth", "row", 1, 5)); assert!(validate_state(&state, 20, 10).is_err());
        let crossing = SheetOutlineState { groups: vec![group("a", "row", 1, 5), group("b", "row", 3, 8)], ..SheetOutlineState::default() };
        assert!(validate_state(&crossing, 20, 10).is_err());
    }

    #[test]
    fn invalid_manual_visibility_snapshot_fails_instead_of_returning_a_lossy_native_fallback() {
        let xml = native_fixture(); let mut extension = snapshot(&xml);
        extension.manual_visibility.row_hidden_spans = vec![HiddenSpan { start: 2, end: 5, value: true }, HiddenSpan { start: 4, end: 7, value: false }];
        let invalid = with_snapshot(&xml, &extension, 1);
        assert!(read_worksheet_outline(std::io::Cursor::new(invalid.as_bytes()), false).is_err());
        let mut missing = serde_json::to_value(snapshot(&xml)).unwrap();
        missing["state"].as_object_mut().unwrap().remove("groups");
        assert!(parse_snapshot(&missing.to_string()).is_err());
        let mut missing = serde_json::to_value(snapshot(&xml)).unwrap();
        missing["manualVisibility"].as_object_mut().unwrap().remove("hiddenRows");
        assert!(parse_snapshot(&missing.to_string()).is_err());
    }

    #[test]
    fn native_level_eight_keeps_eight_model_groups_and_level_nine_cannot_succeed() {
        let xml = worksheet("<sheetData><row r=\"2\" outlineLevel=\"8\"/></sheetData>");
        let source = read(&xml).source;
        assert_eq!(infer_native_state(&source).groups.len(), 8);
        assert!(source.diagnostics[0].starts_with("NATIVE_LEVEL_OUTSIDE_OFFICE"));
        let invalid = xml.replace("outlineLevel=\"8\"", "outlineLevel=\"9\"");
        assert!(read_worksheet_outline(std::io::Cursor::new(invalid.as_bytes()), false).is_err());
    }
}
