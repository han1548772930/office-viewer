//! Versioned conditional-format provenance and accidental-staleness guards.
//! Office rules and styles remain standard XML; the private extension retains
//! local identity and fields which have no corresponding Office representation.

use serde::{Deserialize, Serialize};
use serde_json::{Map, Value, json};

use std::io::{self, Write};

use crate::conditional_format_xml::{Element, MAIN_NAMESPACE, MODEL_EXTENSION_URI, MODEL_NAMESPACE, X14_NAMESPACE, XML_FRAGMENT_BYTE_LIMIT};

/// Version 1 retains five independently bounded XML fields.
pub const SOURCE_XML_TOTAL_BYTE_LIMIT: usize = 5 * XML_FRAGMENT_BYTE_LIMIT;
/// Sum of serialized conditional-format models, including duplicated provenance.
pub const CONDITIONAL_MODEL_BYTE_LIMIT: usize = 256 * 1024 * 1024;
pub const SOURCE_RESOURCE_BYTE_LIMIT: usize = 128 * 1024 * 1024;
pub const SOURCE_RESOURCE_ELEMENT_LIMIT: usize = 2_097_152;

struct JsonByteCounter { bytes: usize, limit: usize, exceeded: bool }
impl Write for JsonByteCounter {
    fn write(&mut self, buffer: &[u8]) -> io::Result<usize> {
        let Some(next) = self.bytes.checked_add(buffer.len()).filter(|next| *next <= self.limit) else {
            self.exceeded = true; return Err(io::Error::other("conditional-format JSON exceeds its byte budget"))
        };
        self.bytes = next; Ok(buffer.len())
    }
    fn flush(&mut self) -> io::Result<()> { Ok(()) }
}

/// Use serde's actual UTF-8/escape representation without allocating a JSON string.
pub fn serialized_model_size<T: Serialize + ?Sized>(value: &T, limit: usize) -> Result<usize, String> {
    let mut output = JsonByteCounter { bytes: 0, limit, exceeded: false };
    if let Err(error) = serde_json::to_writer(&mut output, value) {
        return Err(if output.exceeded { "CF_SOURCE_TOTAL_LIMIT: serialized conditional-format models exceed the byte budget".into() }
            else { format!("CF_SOURCE_XML_INVALID: conditional-format model cannot be serialized: {error}") })
    }
    Ok(output.bytes)
}

pub struct ConditionalFormatModelBudget { used: usize, entries: usize, limit: usize }
impl Default for ConditionalFormatModelBudget {
    fn default() -> Self { Self::new(CONDITIONAL_MODEL_BYTE_LIMIT) }
}
impl ConditionalFormatModelBudget {
    pub fn new(limit: usize) -> Self { Self { used: 2, entries: 0, limit } }
    pub fn used(&self) -> usize { self.used }
    pub fn reserve_bytes(&mut self, bytes: usize) -> Result<(), String> {
        let next = self.used.checked_add(bytes).filter(|next| *next <= self.limit)
            .ok_or("CF_SOURCE_TOTAL_LIMIT: serialized conditional-format models exceed the byte budget")?;
        self.used = next; Ok(())
    }
    pub fn consume<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), String> {
        let separator = usize::from(self.entries > 0);
        let remaining = self.limit.checked_sub(self.used).and_then(|remaining| remaining.checked_sub(separator))
            .ok_or("CF_SOURCE_TOTAL_LIMIT: serialized conditional-format models exceed the byte budget")?;
        let bytes = serialized_model_size(value, remaining)?;
        self.used += separator + bytes; self.entries += 1; Ok(())
    }
}

pub const RULE_FIELDS: &[&str] = &["type", "subType", "ranges", "formulaAnchor", "operator", "formula1", "formula2",
    "value", "isBottom", "isPercent", "config", "style", "priority", "stopIfTrue"];
pub const STYLE_FIELDS: &[&str] = &["fill", "fontColor", "bold", "italic", "underline", "strike"];
pub const CONFIG_FIELDS: &[&str] = &["min", "max", "isGradient", "positiveColor", "nativeColor", "thresholds", "icons",
    "isShowValue", "iconSet", "reverse", "minLength", "maxLength", "border", "borderColor", "negativeBorderColor", "axisColor",
    "negativeBarColorSameAsPositive", "negativeBarBorderColorSameAsPositive", "direction", "axisPosition"];

pub fn office_icon_type(name: &str) -> Option<&str> {
    match name {
        "3TrafficLights" => Some("3TrafficLights1"), "EMPTY_ICON_TYPE" => Some("NoIcons"),
        "3Arrows" | "3ArrowsGray" | "3Flags" | "3TrafficLights1" | "3TrafficLights2" | "3Signs" | "3Symbols" | "3Symbols2"
        | "4Arrows" | "4ArrowsGray" | "4RedToBlack" | "4Rating" | "4TrafficLights" | "5Arrows" | "5ArrowsGray"
        | "5Rating" | "5Quarters" | "3Stars" | "3Triangles" | "5Boxes" | "NoIcons" => Some(name), _ => None,
    }
}
pub fn local_icon_type(name: &str) -> &str { match name { "NoIcons" => "EMPTY_ICON_TYPE", "3TrafficLights" => "3TrafficLights1", _ => name } }
pub fn icon_count(name: &str) -> Option<usize> { office_icon_type(name)?.chars().next()?.to_digit(10).map(|value| value as usize) }
/// The editor's arrows are high-to-low; its rating glyph arrays use Office's
/// low-to-high glyph order. Map actual glyph identities, not only rule positions.
pub fn icon_index(name: &str, index: usize) -> Option<usize> {
    if matches!(name, "NoIcons" | "EMPTY_ICON_TYPE") { return Some(0) }
    let count = icon_count(name)?; if index >= count { return None }
    Some(if matches!(name, "4Rating" | "5Rating") { index } else { count - index - 1 })
}

#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct OfficeSource {
    pub version: u32,
    pub format: String,
    pub original_id: String,
    pub storage: String,
    pub rule_xml: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub group_xml: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub differential_format_xml: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub x14_xml: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub theme_xml: Option<String>,
    pub baseline: Value,
    #[serde(flatten)]
    pub extensions: Map<String, Value>,
}

impl OfficeSource {
    pub fn validate(&self) -> Result<(), String> {
        if self.version != 1 || self.format != "xlsx" || !matches!(self.storage.as_str(), "base" | "x14")
            || self.original_id.trim().is_empty() || self.original_id.len() > 4096 || !valid_known_fields(&self.baseline) || self.rule_xml.trim().is_empty() {
            return Err("CF_SOURCE_XML_INVALID: invalid or unsupported conditional-format Office provenance".into());
        }
        let mut total_bytes = 0usize;
        for xml in std::iter::once(&self.rule_xml).chain(self.group_xml.iter()).chain(self.differential_format_xml.iter())
            .chain(self.x14_xml.iter()).chain(self.theme_xml.iter()) {
            if xml.len() > XML_FRAGMENT_BYTE_LIMIT { return Err("CF_SOURCE_FRAGMENT_LIMIT: conditional-format source XML exceeds the fragment byte budget".into()); }
            total_bytes = total_bytes.checked_add(xml.len()).ok_or("CF_SOURCE_TOTAL_LIMIT: conditional-format source XML byte counter overflow")?;
        }
        if total_bytes > SOURCE_XML_TOTAL_BYTE_LIMIT { return Err("CF_SOURCE_TOTAL_LIMIT: conditional-format source XML exceeds its total byte budget".into()); }
        serialized_model_size(self, CONDITIONAL_MODEL_BYTE_LIMIT)?;
        Ok(())
    }
}

fn selected(value: &Value, fields: &[&str]) -> Value {
    let mut output = Map::new();
    for key in fields { if let Some(value) = value.get(*key).filter(|value| !value.is_null()) { output.insert((*key).into(), value.clone()); } }
    Value::Object(output)
}

fn threshold(value: &Value) -> Value { selected(value, &["type", "value"]) }

/// A baseline contains only fields the canonical model can inspect. Source
/// metadata and unknown JSON do not cause unrelated Office fields to be rebuilt.
pub fn known_snapshot(rule: &Value) -> Value {
    let mut output = selected(rule, RULE_FIELDS);
    for (key, fallback) in [("operator", json!("")), ("formula1", json!("")), ("priority", json!(1)), ("stopIfTrue", json!(false))] {
        if output.get(key).is_none() { output[key] = fallback; }
    }
    for key in ["isBottom", "isPercent"] { if output.get(key) == Some(&Value::Bool(false)) { output.as_object_mut().unwrap().remove(key); } }
    output["style"] = selected(rule.get("style").unwrap_or(&Value::Null), STYLE_FIELDS);
    if let Some(config) = rule.get("config").filter(|value| value.is_object()) {
        let mut known = selected(config, CONFIG_FIELDS);
        for key in ["min", "max"] { if let Some(value) = config.get(key).filter(|value| value.is_object()) { known[key] = threshold(value); } }
        if let Some(entries) = config.get("thresholds").and_then(Value::as_array) {
            known["thresholds"] = Value::Array(entries.iter().map(|entry| {
                let mut item = selected(entry, &["color"]); item["value"] = threshold(&entry["value"]); item
            }).collect());
        }
        if let Some(entries) = config.get("icons").and_then(Value::as_array) {
            known["icons"] = Value::Array(entries.iter().map(|entry| {
                let mut item = selected(entry, &["operator", "iconType", "iconId"]);
                if let Some(id) = item.get("iconId").and_then(Value::as_u64) { item["iconId"] = json!(id.to_string()); }
                item["value"] = threshold(&entry["value"]); item
            }).collect());
        }
        if known.as_object().is_some_and(|value| !value.is_empty()) { output["config"] = known; }
        else { output.as_object_mut().unwrap().remove("config"); }
    } else { output.as_object_mut().unwrap().remove("config"); }
    output
}

pub fn without_source(rule: &Value) -> Value {
    let mut output = rule.clone();
    if let Some(object) = output.as_object_mut() { object.remove("officeSource"); }
    output
}

fn extension_list(node: &Element) -> bool { node.name == "extLst" && (node.is_main("extLst") || node.namespace == X14_NAMESPACE) }
fn recognized_extension(node: &Element) -> bool {
    node.is_main("ext") && node.attrs.get("uri").is_some_and(|uri| uri == MODEL_EXTENSION_URI)
        && node.children.iter().any(|child| child.is(MODEL_NAMESPACE, "model") && child.attrs.get("version").is_some_and(|version| version == "1"))
}

/// Strip only the extension version this implementation understands. Unknown
/// Office/vendor extensions keep their attributes, namespaces and mixed content.
pub fn strip_projection(node: &mut Element) {
    for child in &mut node.children {
        if extension_list(child) { child.replace_children(recognized_extension, Vec::new()); }
        strip_projection(child);
    }
    node.replace_children(|child| extension_list(child) && child.children.is_empty() && child.text.trim().is_empty(), Vec::new());
}

fn group_fingerprint(group: &Element) -> String {
    let mut skeleton = group.clone();
    skeleton.replace_children(|child| child.name == "cfRule" && (child.is_main("cfRule") || child.namespace == X14_NAMESPACE), Vec::new());
    strip_projection(&mut skeleton);
    if let Some(sqref) = skeleton.attrs.get("sqref").cloned() { skeleton.set_attr("sqref", sqref.split_whitespace().collect::<Vec<_>>().join(" ")); }
    for child in &mut skeleton.children { if child.name == "sqref" { child.set_text(child.text.split_whitespace().collect::<Vec<_>>().join(" ")); } }
    skeleton.fingerprint()
}

#[derive(Debug, Clone, Deserialize, Serialize, PartialEq, Eq)]
#[serde(rename_all = "camelCase")]
pub struct ProjectionHashes {
    pub rule: String,
    pub group: String,
    pub differential_format: Option<String>,
    pub x14: Option<String>,
    pub theme: Option<String>,
}

impl ProjectionHashes {
    pub fn of(rule: &Element, group: &Element, dxf: Option<&Element>, x14: Option<&Element>, theme: Option<&Element>) -> Self {
        let mut rule = rule.clone(); strip_projection(&mut rule);
        let x14 = x14.map(|group| { let mut group = group.clone(); strip_projection(&mut group); group.fingerprint() });
        Self { rule: rule.fingerprint(), group: group_fingerprint(group), differential_format: dxf.map(Element::fingerprint),
            x14, theme: theme.map(Element::fingerprint) }
    }
}

#[derive(Debug, Clone, Deserialize, Serialize)]
pub struct Projection { pub hashes: ProjectionHashes, pub model: Value }

pub fn read_projection(rule: &Element) -> Option<Projection> {
    for list in rule.children.iter().filter(|child| extension_list(child)) {
        for ext in list.children.iter().filter(|child| recognized_extension(child)) {
            let payload = ext.children.iter().find(|child| child.is(MODEL_NAMESPACE, "model") && child.attrs.get("version").is_some_and(|version| version == "1"))?;
            let projection = serde_json::from_str::<Projection>(&payload.text).ok()?;
            if valid_saved_model(&projection.model) { return Some(projection); }
        }
    }
    None
}

pub fn attach_projection(rule: &mut Element, model: &Value, hashes: ProjectionHashes) -> Result<(), String> {
    if !valid_saved_model(model) { return Err("Invalid conditional-format model snapshot".into()) }
    strip_projection(rule);
    let projection = Projection { hashes, model: without_source(model) };
    let mut payload = Element::prefixed(MODEL_NAMESPACE, "oscf", "model"); payload.set_attr("version", "1");
    payload.set_text(serde_json::to_string(&projection).map_err(|error| error.to_string())?);
    let mut ext = Element::new(MAIN_NAMESPACE, "ext"); ext.set_attr("uri", MODEL_EXTENSION_URI); ext.push(payload);
    if let Some(list) = rule.children.iter_mut().find(|child| extension_list(child)) { list.push(ext); }
    else {
        let mut list = if rule.namespace == X14_NAMESPACE { Element::prefixed(X14_NAMESPACE, "x14", "extLst") } else { Element::new(MAIN_NAMESPACE, "extLst") };
        list.push(ext); rule.push(list);
    }
    Ok(())
}

pub fn valid_saved_model(model: &Value) -> bool {
    let Some(id) = model.get("id").and_then(Value::as_str) else { return false };
    if id.trim().is_empty() || id.len() > 4096 || !valid_known_fields(model) { return false }
    let kind = model["type"].as_str().unwrap();
    let subtype = model.get("subType").and_then(Value::as_str).unwrap_or("");
    let operator = model.get("operator").and_then(Value::as_str).unwrap_or("");
    if kind == "office" { return true }
    let nonempty = |key: &str| model.get(key).and_then(Value::as_str).is_some_and(|value| !value.trim().trim_start_matches('=').trim().is_empty());
    if kind == "expression" || subtype == "formula" {
        if !nonempty("formula1") && !nonempty("value") { return false }
    }
    if kind == "cellIs" || kind == "highlightCell" && matches!(subtype, "" | "number") {
        if !matches!(operator, "equal" | "notEqual" | "greaterThan" | "greaterThanOrEqual" | "lessThan" | "lessThanOrEqual" | "between" | "notBetween")
            || !nonempty("formula1") || matches!(operator, "between" | "notBetween") && !nonempty("formula2") { return false }
    }
    if kind == "average" || subtype == "average" {
        if !matches!(operator, "equal" | "notEqual" | "greaterThan" | "greaterThanOrEqual" | "lessThan" | "lessThanOrEqual") { return false }
    }
    if kind == "timePeriod" || subtype == "timePeriod" {
        if !matches!(operator, "today" | "yesterday" | "tomorrow" | "last7Days" | "thisWeek" | "lastWeek" | "nextWeek" | "thisMonth" | "lastMonth" | "nextMonth") { return false }
    }
    if kind == "top10" || subtype == "rank" {
        if let Some(value) = model.get("value") {
            let rank = value.as_u64().or_else(|| value.as_str().and_then(|value| value.trim().parse().ok()));
            let maximum = if model.get("isPercent").and_then(Value::as_bool).unwrap_or(false) { 100 } else { 1000 };
            if rank.is_none_or(|rank| rank == 0 || rank > maximum) { return false }
        }
    }
    let config = model.get("config").unwrap_or(&Value::Null);
    if kind == "dataBar" {
        for (key, automatic) in [("min", "autoMin"), ("max", "autoMax")] {
            if let Some(value) = config.get(key) { if !valid_threshold(value, Some(automatic)) { return false } }
        }
    }
    if kind == "colorScale" {
        let Some(stops) = config.get("thresholds").and_then(Value::as_array) else { return false };
        if !(2..=3).contains(&stops.len()) || stops.iter().any(|entry| !valid_threshold(&entry["value"], None)) { return false }
    }
    if kind == "iconSet" {
        let Some(icons) = config.get("icons").and_then(Value::as_array) else { return false };
        if !(3..=5).contains(&icons.len()) { return false }
        let set = config.get("iconSet").and_then(Value::as_str).unwrap_or("3TrafficLights1");
        if office_icon_type(set).is_none() && set != "_5Felling" { return false }
        for icon in icons {
            if !valid_threshold(&icon["value"], None) || icon.get("operator").is_some_and(|value| !matches!(value.as_str(), Some("greaterThan" | "greaterThanOrEqual"))) { return false }
            let kind = icon.get("iconType").and_then(Value::as_str).unwrap_or(set);
            if office_icon_type(kind).is_none() && kind != "_5Felling" { return false }
            if let Some(value) = icon.get("iconId") {
                if matches!(kind, "EMPTY_ICON_TYPE" | "NoIcons") && value.as_str() == Some("") { continue }
                let id = value.as_u64().or_else(|| value.as_str().and_then(|value| value.parse().ok()));
                if id.is_none_or(|id| if kind == "_5Felling" { id >= 5 } else { icon_index(kind, id as usize).is_none() }) { return false }
            }
        }
    }
    true
}

fn valid_color(value: &Value) -> bool {
    value.as_str().and_then(|value| value.strip_prefix('#')).is_some_and(|value|
        matches!(value.len(), 3 | 6 | 8) && value.bytes().all(|value| value.is_ascii_hexdigit()))
}
fn valid_scalar(value: &Value) -> bool { value.is_string() || value.as_f64().is_some_and(f64::is_finite) }
fn threshold_shape(value: &Value) -> bool {
    value.is_object() && value.get("type").is_some_and(Value::is_string) && value.get("value").is_none_or(valid_scalar)
}
fn valid_threshold(value: &Value, automatic: Option<&str>) -> bool {
    if !threshold_shape(value) { return false }
    let kind = value["type"].as_str().unwrap();
    if matches!(kind, "min" | "max") { return true }
    if matches!(kind, "autoMin" | "autoMax") { return automatic == Some(kind) }
    if kind == "formula" { return value.get("value").and_then(Value::as_str).is_some_and(|value| !value.trim().trim_start_matches('=').trim().is_empty()) }
    if !matches!(kind, "num" | "percent" | "percentile") { return false }
    let numeric = value.get("value").and_then(|value| value.as_f64().or_else(|| value.as_str().and_then(|value| value.trim().parse().ok())));
    numeric.is_some_and(|value| value.is_finite() && (!matches!(kind, "percent" | "percentile") || (0.0..=100.0).contains(&value)))
}

/// Primitive/shape validation is also used for provenance baselines. Semantic
/// validation of a private model is stricter because it can replace parsed data.
pub fn valid_known_fields(model: &Value) -> bool {
    if !model.is_object() || !matches!(model.get("type").and_then(Value::as_str), Some("cellIs" | "highlightCell" | "top10" | "average"
        | "duplicateValues" | "uniqueValues" | "timePeriod" | "containsText" | "notContainsText" | "beginsWith" | "endsWith"
        | "expression" | "dataBar" | "colorScale" | "iconSet" | "office")) { return false }
    let Some(ranges) = model.get("ranges").and_then(Value::as_array) else { return false };
    if ranges.is_empty() || ranges.len() > 100_000 || !model.get("style").is_some_and(Value::is_object) { return false }
    for range in ranges {
        let (Some(sr), Some(sc), Some(er), Some(ec)) = (range["startRow"].as_u64(), range["startCol"].as_u64(), range["endRow"].as_u64(), range["endCol"].as_u64()) else { return false };
        if er < sr || ec < sc || er > 1_048_575 || ec > 16_383 { return false }
    }
    if model.get("priority").is_some_and(|value| !value.as_u64().is_some_and(|value| value > 0 && value <= i32::MAX as u64)) { return false }
    for key in ["operator", "formula1", "formula2"] { if model.get(key).is_some_and(|value| !value.is_string()) { return false } }
    if model.get("subType").is_some_and(|value| !matches!(value.as_str(), Some("number" | "text" | "formula" | "timePeriod" | "rank" | "average" | "duplicateValues" | "uniqueValues"))) { return false }
    if let Some(anchor) = model.get("formulaAnchor") {
        if !anchor.is_object() || !anchor["row"].as_u64().is_some_and(|value| value <= 1_048_575)
            || !anchor["col"].as_u64().is_some_and(|value| value <= 16_383) { return false }
    }
    for key in ["stopIfTrue", "isBottom", "isPercent"] { if model.get(key).is_some_and(|value| !value.is_boolean()) { return false } }
    if model.get("value").is_some_and(|value| !valid_scalar(value) && !value.as_array().is_some_and(|values|
        values.len() == 2 && values.iter().all(|value| value.as_f64().is_some_and(f64::is_finite)))) { return false }
    let style = &model["style"];
    for key in ["fill", "fontColor"] { if style.get(key).is_some_and(|value| !valid_color(value)) { return false } }
    for key in ["bold", "italic", "underline", "strike"] { if style.get(key).is_some_and(|value| !value.is_boolean()) { return false } }
    let Some(config) = model.get("config") else { return true };
    if !config.is_object() { return false }
    for key in ["isGradient", "isShowValue", "reverse", "border", "negativeBarColorSameAsPositive", "negativeBarBorderColorSameAsPositive"] {
        if config.get(key).is_some_and(|value| !value.is_boolean()) { return false }
    }
    for key in ["positiveColor", "nativeColor", "borderColor", "negativeBorderColor", "axisColor"] {
        if config.get(key).is_some_and(|value| !valid_color(value)) { return false }
    }
    for key in ["minLength", "maxLength"] { if config.get(key).is_some_and(|value| !value.as_u64().is_some_and(|value| value <= 100)) { return false } }
    if config["minLength"].as_u64().unwrap_or(0) > config["maxLength"].as_u64().unwrap_or(100) { return false }
    if config.get("direction").is_some_and(|value| !matches!(value.as_str(), Some("context" | "leftToRight" | "rightToLeft")))
        || config.get("axisPosition").is_some_and(|value| !matches!(value.as_str(), Some("automatic" | "middle" | "none")))
        || config.get("iconSet").is_some_and(|value| !value.is_string()) { return false }
    for key in ["min", "max"] { if config.get(key).is_some_and(|value| !threshold_shape(value)) { return false } }
    if let Some(stops) = config.get("thresholds") {
        let Some(stops) = stops.as_array() else { return false };
        if stops.len() > 100_000 || stops.iter().any(|entry| !entry.is_object() || !threshold_shape(&entry["value"]) || !valid_color(&entry["color"])) { return false }
    }
    if let Some(icons) = config.get("icons") {
        let Some(icons) = icons.as_array() else { return false };
        if icons.len() > 100_000 || icons.iter().any(|entry| !entry.is_object() || !threshold_shape(&entry["value"])
            || ["operator", "iconType"].iter().any(|key| entry.get(*key).is_some_and(|value| !value.is_string()))
            || entry.get("iconId").is_some_and(|value| !value.is_string() && value.as_u64().is_none())) { return false }
    }
    true
}

fn copy_unknown(output: &mut Value, source: &Value, known: &[&str]) {
    let (Some(output), Some(source)) = (output.as_object_mut(), source.as_object()) else { return };
    for (key, value) in source { if !known.contains(&key.as_str()) && key != "officeSource" { output.insert(key.clone(), value.clone()); } }
}

fn nested_unknown(output: &mut Value, source: &Value) {
    copy_unknown(&mut output["style"], &source["style"], STYLE_FIELDS);
    if !output.get("config").is_some_and(Value::is_object) { return }
    let config = &mut output["config"]; let before = &source["config"];
    copy_unknown(config, before, CONFIG_FIELDS);
    for key in ["min", "max"] { if config.get(key).is_some_and(Value::is_object) { copy_unknown(&mut config[key], &before[key], &["type", "value"]); } }
    for (key, fields) in [("thresholds", &["value", "color"][..]), ("icons", &["value", "operator", "iconType", "iconId"][..])] {
        if let Some(entries) = config.get_mut(key).and_then(Value::as_array_mut) {
            for (index, entry) in entries.iter_mut().enumerate() {
                let prior = &before[key][index]; copy_unknown(entry, prior, fields);
                if entry.get("value").is_some_and(Value::is_object) { copy_unknown(&mut entry["value"], &prior["value"], &["type", "value"]); }
            }
        }
    }
}

/// External Office edits win over stale local fields. Local identity and
/// unmapped JSON remain separate from that standard projection.
pub fn restore_projection(output: &mut Value, projection: &Projection, actual: &ProjectionHashes) -> bool {
    if !valid_saved_model(&projection.model) { return false }
    if projection.hashes == *actual { *output = without_source(&projection.model); return true }
    let mut known = RULE_FIELDS.to_vec(); known.push("id");
    copy_unknown(output, &projection.model, &known); nested_unknown(output, &projection.model);
    output["id"] = projection.model["id"].clone();
    false
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::conditional_format_xml::parse;

    fn model() -> Value { json!({ "id": "stable-rule", "type": "expression", "ranges": [{ "startRow": 0, "startCol": 0, "endRow": 0, "endCol": 0 }],
        "operator": "", "formula1": "=A1>0", "style": { "bold": true, "vendor": { "width": 2 } }, "priority": 8, "stopIfTrue": false,
        "metadata": { "note": "retained" }, "officeSource": { "recursive": "must not be saved" } }) }

    #[test]
    fn extension_retains_local_identity_without_recursively_embedding_provenance() {
        let mut rule = parse(r#"<cfRule type="expression" priority="1"><formula>A1&gt;0</formula></cfRule>"#).unwrap();
        let group = parse(r#"<conditionalFormatting sqref="A1"/>"#).unwrap();
        let hashes = ProjectionHashes::of(&rule, &group, None, None, None);
        attach_projection(&mut rule, &model(), hashes.clone()).unwrap();
        let first = rule.render(); let projection = read_projection(&parse(&first).unwrap()).unwrap();
        assert!(projection.model.get("officeSource").is_none()); assert_eq!(projection.model["id"], "stable-rule");
        assert_eq!(hashes, ProjectionHashes::of(&rule, &group, None, None, None));
        attach_projection(&mut rule, &model(), hashes).unwrap(); assert_eq!(rule.render(), first);
    }

    #[test]
    fn stale_projection_preserves_standard_edits_identity_and_unknown_fields_independently() {
        let rule = parse(r#"<cfRule type="expression" priority="1"><formula>A1&gt;0</formula></cfRule>"#).unwrap();
        let group = parse(r#"<conditionalFormatting sqref="A1"/>"#).unwrap();
        let hashes = ProjectionHashes::of(&rule, &group, None, None, None);
        let projection = Projection { hashes: hashes.clone(), model: model() };
        let mut output = model(); output["formula1"] = json!("A1>9"); output["priority"] = json!(1); output["style"] = json!({ "bold": false });
        let mut edited = hashes; edited.rule = "edited".into();
        assert!(!restore_projection(&mut output, &projection, &edited));
        assert_eq!(output["formula1"], "A1>9"); assert_eq!(output["style"]["bold"], false);
        assert_eq!(output["style"]["vendor"], json!({ "width": 2 })); assert_eq!(output["id"], "stable-rule");
        assert_eq!(output["metadata"]["note"], "retained");
    }

    #[test]
    fn provenance_and_unknown_json_are_excluded_from_the_editable_baseline() {
        let mut first = model(); first["config"] = json!({ "icons": [{ "value": { "type": "num", "value": 40, "vendor": 1 }, "iconId": 2, "vendor": 2 }], "vendor": 3 });
        let mut next = first.clone(); next["officeSource"] = json!({ "baseline": { "vendor": "changed" } }); next["style"]["vendor"] = json!(99);
        next["config"]["icons"][0]["iconId"] = json!("2"); next["config"]["icons"][0]["value"]["vendor"] = json!(9);
        assert_eq!(known_snapshot(&first), known_snapshot(&next));
        next["config"]["icons"][0]["value"]["value"] = json!(41); assert_ne!(known_snapshot(&first), known_snapshot(&next));
    }

    #[test]
    fn invalid_private_known_fields_cannot_replace_valid_standard_xml_even_with_matching_hashes() {
        let rule = parse(r#"<cfRule type="expression" priority="1"><formula>A1&gt;0</formula></cfRule>"#).unwrap();
        let group = parse(r#"<conditionalFormatting sqref="A1"/>"#).unwrap();
        let hashes = ProjectionHashes::of(&rule, &group, None, None, None);
        for (key, value) in [("type", json!("futureCondition")), ("formula1", json!(2)), ("priority", json!(0)),
            ("style", json!({ "bold": "false" })), ("formulaAnchor", json!({ "row": -1, "col": 0 })),
            ("config", json!({ "minLength": 101 }))] {
            let mut corrupt = model(); corrupt[key] = value;
            let projection = Projection { hashes: hashes.clone(), model: corrupt.clone() };
            let mut standard = model(); standard["formula1"] = json!("A1>9"); let before = standard.clone();
            assert!(!restore_projection(&mut standard, &projection, &hashes)); assert_eq!(standard, before);
            assert!(attach_projection(&mut rule.clone(), &corrupt, hashes.clone()).is_err());
        }
    }

    #[test]
    fn source_baselines_validate_known_fields_without_needing_local_identity_or_unknown_json() {
        let mut source = OfficeSource { version: 1, format: "xlsx".into(), original_id: "local-id".into(), storage: "base".into(),
            rule_xml: "<cfRule/>".into(), group_xml: None, differential_format_xml: None, x14_xml: None, theme_xml: None,
            baseline: known_snapshot(&model()), extensions: Map::new() };
        assert!(source.validate().is_ok()); assert!(source.baseline.get("id").is_none());
        source.baseline["style"]["underline"] = json!("double"); assert!(source.validate().is_err());
        source.baseline = known_snapshot(&model()); source.baseline["priority"] = json!(i32::MAX as u64 + 1); assert!(source.validate().is_err());
    }

    #[test]
    fn automatic_thresholds_are_limited_to_corresponding_data_bar_boundaries() {
        let mut bar = model(); bar["type"] = json!("dataBar");
        bar["config"] = json!({ "min": { "type": "autoMin" }, "max": { "type": "autoMax" } });
        assert!(valid_saved_model(&bar));
        for (key, kind) in [("min", "autoMax"), ("max", "autoMin")] {
            let mut invalid = bar.clone(); invalid["config"][key]["type"] = json!(kind);
            assert!(!valid_saved_model(&invalid));
        }
        for kind in ["autoMin", "autoMax"] {
            let mut scale = bar.clone(); scale["type"] = json!("colorScale");
            scale["config"] = json!({ "thresholds": [{ "value": { "type": kind }, "color": "#112233" },
                { "value": { "type": "max" }, "color": "#aabbcc" }] });
            assert!(!valid_saved_model(&scale));
            let mut icons = bar.clone(); icons["type"] = json!("iconSet");
            icons["config"] = json!({ "iconSet": "3Arrows", "icons": [
                { "value": { "type": kind }, "operator": "greaterThanOrEqual", "iconType": "3Arrows", "iconId": "0" },
                { "value": { "type": "num", "value": 10 }, "operator": "greaterThan", "iconType": "3Arrows", "iconId": "1" },
                { "value": { "type": "num", "value": 0 }, "iconType": "3Arrows", "iconId": "2" }] });
            assert!(!valid_saved_model(&icons));
        }
    }

    #[test]
    fn icon_identity_mapping_is_involutive_for_stock_glyphs_and_no_icon_is_explicit() {
        for set in ["3Arrows", "4Arrows", "5Arrows", "4Rating", "5Rating", "3Stars", "5Boxes"] {
            for index in 0..icon_count(set).unwrap() { assert_eq!(icon_index(set, icon_index(set, index).unwrap()), Some(index)); }
        }
        assert_eq!(office_icon_type("EMPTY_ICON_TYPE"), Some("NoIcons")); assert_eq!(icon_index("NoIcons", 0), Some(0));
        assert!(office_icon_type("_5Felling").is_none()); assert!(icon_index("3Arrows", 3).is_none());
    }

    #[test]
    fn json_byte_budget_counts_actual_utf8_escapes_and_unknown_extensions() {
        let value = json!({ "rule": model(), "vendor": { "text": "汉😀\n\\\"\u{1}", "copies": ["source", "source"] } });
        let actual = serde_json::to_vec(&value).unwrap().len();
        assert_eq!(serialized_model_size(&value, actual).unwrap(), actual);
        assert!(serialized_model_size(&value, actual - 1).unwrap_err().starts_with("CF_SOURCE_TOTAL_LIMIT"));
    }

    #[test]
    fn duplicated_rule_provenance_spends_the_cumulative_model_budget_atomically() {
        let value = json!({ "rule": model(), "officeSource": { "ruleXml": "<cfRule>汉😀</cfRule>", "vendor": { "retained": true } } });
        let bytes = serde_json::to_vec(&value).unwrap().len();
        let mut exact = ConditionalFormatModelBudget::new(2 * bytes + 3);
        exact.consume(&value).unwrap(); exact.consume(&value).unwrap(); assert_eq!(exact.used(), 2 * bytes + 3);
        assert!(exact.reserve_bytes(1).unwrap_err().starts_with("CF_SOURCE_TOTAL_LIMIT")); assert_eq!(exact.used(), 2 * bytes + 3);
        let mut short = ConditionalFormatModelBudget::new(2 * bytes + 2);
        short.consume(&value).unwrap(); let before = short.used();
        assert!(short.consume(&value).unwrap_err().starts_with("CF_SOURCE_TOTAL_LIMIT")); assert_eq!(short.used(), before);
    }

    #[test]
    fn version_one_xml_fields_are_limited_by_utf8_bytes_and_never_truncated() {
        let mut source = OfficeSource { version: 1, format: "xlsx".into(), original_id: "local-id".into(), storage: "base".into(),
            rule_xml: "<cfRule/>".into(), group_xml: None, differential_format_xml: None, x14_xml: None, theme_xml: None,
            baseline: known_snapshot(&model()), extensions: Map::new() };
        source.rule_xml = "汉".repeat(XML_FRAGMENT_BYTE_LIMIT / 3 + 1);
        let original_bytes = source.rule_xml.len();
        assert!(source.validate().unwrap_err().starts_with("CF_SOURCE_FRAGMENT_LIMIT"));
        assert_eq!(source.rule_xml.len(), original_bytes);
        source.rule_xml = "<cfRule/>".into(); source.original_id = "汉".repeat(4096 / 3 + 1);
        assert!(source.validate().unwrap_err().starts_with("CF_SOURCE_XML_INVALID"));
    }
}
