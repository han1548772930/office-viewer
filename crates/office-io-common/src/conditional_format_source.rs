//! Versioned conditional-format provenance and accidental-staleness guards.
//! Office rules and styles remain standard XML; the private extension retains
//! local identity and fields which have no corresponding Office representation.

use serde::{Deserialize, Serialize};
use serde_json::{Map, Value, json};

use crate::conditional_format_xml::{Element, MAIN_NAMESPACE, MODEL_EXTENSION_URI, MODEL_NAMESPACE, X14_NAMESPACE};

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
            || self.original_id.trim().is_empty() || !self.baseline.is_object() || self.rule_xml.is_empty() {
            return Err("Invalid or unsupported conditional-format Office provenance".into());
        }
        for xml in std::iter::once(&self.rule_xml).chain(self.group_xml.iter()).chain(self.differential_format_xml.iter())
            .chain(self.x14_xml.iter()).chain(self.theme_xml.iter()) {
            if xml.len() > 16 * 1024 * 1024 { return Err("Conditional-format source XML exceeds the fragment limit".into()); }
        }
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
    let Some(ranges) = model.get("ranges").and_then(Value::as_array) else { return false };
    if id.trim().is_empty() || id.len() > 4096 || ranges.is_empty() || !model.get("type").is_some_and(Value::is_string)
        || !model.get("style").is_some_and(Value::is_object) { return false }
    for range in ranges {
        let (Some(sr), Some(sc), Some(er), Some(ec)) = (range["startRow"].as_u64(), range["startCol"].as_u64(), range["endRow"].as_u64(), range["endCol"].as_u64()) else { return false };
        if er < sr || ec < sc || er > 1_048_575 || ec > 16_383 { return false }
    }
    model.get("priority").is_none_or(|value| value.as_u64().is_some_and(|value| value > 0 && value <= i32::MAX as u64))
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
}
