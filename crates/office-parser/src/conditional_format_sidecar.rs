use std::collections::{HashMap, HashSet};
use std::io::Read;

use office_io_common::conditional_format_color::{color_value, uses_theme};
use office_io_common::conditional_format_source::{OfficeSource, ProjectionHashes, icon_index, known_snapshot, local_icon_type, read_projection, restore_projection, strip_projection};
use office_io_common::conditional_format_xml::{self as source_xml, Element as Node, RULE_LINK_EXTENSION_URI, X14_NAMESPACE, XM_NAMESPACE};
use serde_json::{Value, json};
use zip::ZipArchive;

const REL_NAMESPACE: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_REL_NAMESPACE: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const REVISION_NAMESPACE: &str = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision";

#[derive(Debug, Default)]
pub(crate) struct ConditionalFormatSidecar {
    by_sheet: HashMap<usize, Vec<Value>>,
    pub(crate) warnings: Vec<String>,
}

#[derive(Debug)]
struct Relationship { id: String, target: String, kind: String, external: bool }

impl ConditionalFormatSidecar {
    pub(crate) fn read(bytes: &[u8]) -> Self {
        match read_sidecar(bytes) {
            Ok(value) => value,
            Err(error) => Self { warnings: vec![error], ..Self::default() },
        }
    }
    pub(crate) fn take(&mut self, sheet_index: usize) -> Vec<Value> { self.by_sheet.remove(&sheet_index).unwrap_or_default() }
}

fn read_sidecar(bytes: &[u8]) -> Result<ConditionalFormatSidecar, String> {
    let mut archive = ZipArchive::new(std::io::Cursor::new(bytes))
        .map_err(|_| "Invalid XLSX ZIP while reading conditional formatting".to_owned())?;
    let workbook = parse_part(&mut archive, "xl/workbook.xml")?
        .ok_or_else(|| "Missing XLSX workbook part while reading conditional formatting".to_owned())?;
    let workbook = source_xml::parse(&workbook)?;
    let workbook_rels = parse_part(&mut archive, "xl/_rels/workbook.xml.rels")?.unwrap_or_default();
    let relation_map = relationships(&workbook_rels)?.into_iter().filter(|relationship| !relationship.external).filter_map(|relationship| {
        resolve("xl/workbook.xml", &relationship.target).map(|path| (relationship.id, (relationship.kind, path)))
    }).collect::<HashMap<_, _>>();
    let style_path = relation_map.values().find_map(|(kind, path)| (kind == "styles").then_some(path.as_str()));
    let dxfs = match style_path {
        Some(path) => match parse_part(&mut archive, path)? {
            Some(xml) => source_xml::parse(&xml)?.descendants("dxf").into_iter().filter(|node| node.is_main("dxf")).collect(),
            None => Vec::new(),
        },
        None => Vec::new(),
    };
    let theme_path = relation_map.values().find_map(|(kind, path)| (kind == "theme").then_some(path.as_str()));
    let theme_xml = match theme_path { Some(path) => parse_part(&mut archive, path)?, None => None };
    let mut sidecar = ConditionalFormatSidecar::default();
    let sheets = workbook.child("sheets").map(|sheets| sheets.children.iter().filter(|sheet| sheet.is_main("sheet")).collect::<Vec<_>>()).unwrap_or_default();
    for (index, sheet) in sheets.into_iter().enumerate() {
        let Some(relationship_id) = sheet.attribute_ns(REL_NAMESPACE, "id").or_else(|| sheet.attribute_ns(STRICT_REL_NAMESPACE, "id")) else { continue };
        let Some((kind, path)) = relation_map.get(relationship_id) else { continue };
        if kind != "worksheet" { continue }
        let Some(xml) = parse_part(&mut archive, path)? else { continue };
        let rules = parse_rules_with_theme(&xml, &dxfs, theme_xml.as_deref(), path, &mut sidecar.warnings);
        if !rules.is_empty() { sidecar.by_sheet.insert(index, rules); }
    }
    Ok(sidecar)
}

fn detached_rule(group: &Node, rule: &Node) -> Node {
    let mut rule = rule.clone();
    for (key, value) in &group.attrs {
        if (key == "xmlns" || key.starts_with("xmlns:")) && !rule.attrs.contains_key(key) { rule.set_attr(key, value); }
        if group.attribute_namespaces.get(key).is_some_and(|namespace| namespace == "http://schemas.openxmlformats.org/markup-compatibility/2006") && !rule.attrs.contains_key(key) {
            rule.set_attr_ns("http://schemas.openxmlformats.org/markup-compatibility/2006", key, value);
        }
    }
    rule
}
fn is_rule(node: &Node) -> bool { node.is_main("cfRule") || node.is(X14_NAMESPACE, "cfRule") }
fn skeleton(group: &Node) -> Node {
    let mut group = group.clone(); group.replace_children(is_rule, Vec::new()); strip_projection(&mut group); group
}
fn single_rule_group(group: &Node, rule: &Node) -> Node {
    let mut group = group.clone(); group.replace_children(is_rule, vec![rule.clone()]); group
}
fn sqref(group: &Node) -> Option<&str> {
    if group.is(X14_NAMESPACE, "conditionalFormatting") {
        group.children.iter().find(|child| child.is(XM_NAMESPACE, "sqref")).map(|child| child.text.as_str())
    } else { group.attrs.get("sqref").map(String::as_str) }
}
fn ranges(group: &Node, warnings: &mut Vec<String>) -> Option<Vec<Value>> {
    let ranges = sqref(group).unwrap_or("").split_whitespace().map(parse_range).collect::<Option<Vec<_>>>();
    if ranges.as_ref().is_none_or(|ranges| ranges.is_empty()) {
        warnings.push("Skipped XLSX conditional formatting with invalid sqref".into()); return None
    }
    ranges
}
fn flag(node: &Node, key: &str, default: bool) -> bool {
    node.attrs.get(key).map(|value| value == "1" || value.eq_ignore_ascii_case("true")).unwrap_or(default)
}
fn linked_id(rule: &Node) -> Option<String> {
    let list = rule.children.iter().find(|child| child.is_main("extLst"))?;
    let ext = list.children.iter().find(|child| child.is_main("ext") && child.attrs.get("uri").is_some_and(|uri| uri.eq_ignore_ascii_case(RULE_LINK_EXTENSION_URI)))?;
    ext.children.iter().find(|child| child.is(X14_NAMESPACE, "id")).map(|child| child.text.trim().to_ascii_uppercase())
}
fn identity(rule: &Node, group: &Node, sheet_path: &str, ordinal: usize) -> String {
    if let Some(projection) = read_projection(rule) { return projection.model["id"].as_str().unwrap().to_owned() }
    if let Some(id) = rule.attribute_ns(REVISION_NAMESPACE, "uid") { return id.to_owned() }
    if rule.namespace == X14_NAMESPACE && let Some(id) = rule.attrs.get("id") { return id.clone() }
    if let Some(id) = linked_id(rule) { return id }
    let mut authored = rule.clone(); strip_projection(&mut authored);
    authored.remove_attr("priority"); authored.remove_attr("dxfId");
    let mut key = Node::new("urn:office-studio:conditional-formatting:identity", "identity");
    key.set_attr("part", sheet_path); key.set_attr("sqref", sqref(group).unwrap_or("")); key.push(authored);
    format!("conditional-{}-{ordinal}", key.fingerprint())
}

#[derive(Clone)]
struct X14Rule { rule: Node, group: Node, used: bool }

#[cfg(test)]
fn parse_rules(xml: &str, dxfs: &[Node], warnings: &mut Vec<String>) -> Vec<Value> {
    parse_rules_with_theme(xml, dxfs, None, "xl/worksheets/sheet1.xml", warnings)
}

fn parse_rules_with_theme(xml: &str, dxfs: &[Node], theme_xml: Option<&str>, sheet_path: &str, warnings: &mut Vec<String>) -> Vec<Value> {
    let root = match source_xml::parse(xml) { Ok(root) => root, Err(error) => { warnings.push(format!("Could not preserve XLSX conditional formatting: {error}")); return Vec::new() } };
    let theme = theme_xml.and_then(|xml| source_xml::parse(xml).ok());
    let mut x14 = root.descendants("conditionalFormatting").into_iter().filter(|group| group.is(X14_NAMESPACE, "conditionalFormatting")).flat_map(|group| {
        group.children.iter().filter(|rule| rule.is(X14_NAMESPACE, "cfRule")).map(|rule| {
            let rule = detached_rule(&group, rule); X14Rule { group: single_rule_group(&group, &rule), rule, used: false }
        }).collect::<Vec<_>>()
    }).collect::<Vec<_>>();
    let mut output = Vec::<(u32, usize, Value)>::new();
    let mut ordinal = 0usize;
    for group in root.children.iter().filter(|group| group.is_main("conditionalFormatting")) {
        let group = detached_rule(&root, group);
        let Some(rule_ranges) = ranges(&group, warnings) else { continue };
        for rule in group.children.iter().filter(|rule| rule.is_main("cfRule")) {
            ordinal += 1;
            let rule = detached_rule(&group, rule);
            let linked = linked_id(&rule).and_then(|id| x14.iter_mut().find(|candidate| !candidate.used && !candidate.rule.attrs.contains_key("priority")
                && candidate.rule.attrs.get("id").is_some_and(|candidate_id| candidate_id.trim().eq_ignore_ascii_case(&id))));
            let linked_group = linked.map(|candidate| { candidate.used = true; candidate.group.clone() });
            let linked_rule = linked_group.as_ref().and_then(|group| group.children.iter().find(|rule| rule.is(X14_NAMESPACE, "cfRule")));
            let dxf = rule.attrs.get("dxfId").and_then(|value| value.parse::<usize>().ok()).and_then(|index| dxfs.get(index));
            if rule.attrs.contains_key("dxfId") && dxf.is_none() { warnings.push("XLSX conditional formatting references an unavailable differential style; its source reference was retained".into()); }
            let priority = rule.attrs.get("priority").and_then(|value| value.parse::<u32>().ok()).filter(|value| *value > 0).unwrap_or(ordinal as u32);
            let mut model = parse_model(&rule, linked_rule, &rule_ranges, dxf, theme.as_ref(), priority, warnings);
            if let Some(linked_group) = &linked_group {
                if sqref(linked_group).map(|value| value.split_whitespace().collect::<Vec<_>>().join(" "))
                    != sqref(&group).map(|value| value.split_whitespace().collect::<Vec<_>>().join(" ")) {
                    model["type"] = json!("office"); model.as_object_mut().unwrap().remove("subType");
                    warnings.push("Retained a linked x14 conditional format with a different scope without local evaluation".into());
                }
            }
            model["id"] = json!(identity(&rule, &group, sheet_path, ordinal));
            finish_model(&mut model, &rule, &group, dxf, linked_group.as_ref(), theme_xml, theme.as_ref(), "base");
            output.push((priority, ordinal, model));
        }
    }
    for source in x14.into_iter().filter(|source| !source.used) {
        ordinal += 1;
        let Some(rule_ranges) = ranges(&source.group, warnings) else { continue };
        let dxf = source.rule.children.iter().find(|child| child.is(X14_NAMESPACE, "dxf") || child.is_main("dxf"));
        let priority = source.rule.attrs.get("priority").and_then(|value| value.parse::<u32>().ok()).filter(|value| *value > 0).unwrap_or(ordinal as u32);
        let mut model = parse_model(&source.rule, None, &rule_ranges, dxf, theme.as_ref(), priority, warnings);
        if !source.rule.attrs.contains_key("priority") {
            model["type"] = json!("office"); model.as_object_mut().unwrap().remove("subType");
            warnings.push("Retained an orphan x14 conditional-format enhancement without local evaluation".into());
        }
        model["id"] = json!(identity(&source.rule, &source.group, sheet_path, ordinal));
        finish_model(&mut model, &source.rule, &source.group, dxf, None, theme_xml, theme.as_ref(), "x14");
        output.push((priority, ordinal, model));
    }
    // Standard priorities retain the emitted stack order even when a private
    // snapshot restores sparse or tied local priorities across base/x14 storage.
    output.sort_by_key(|(priority, ordinal, _)| (*priority, *ordinal));
    let mut ids = HashSet::new();
    output.into_iter().map(|(_, ordinal, mut rule)| {
        let id = rule["id"].as_str().unwrap_or("conditional").to_owned();
        if !ids.insert(id.clone()) {
            let mut unique = format!("{id}-import-{ordinal}");
            while !ids.insert(unique.clone()) { unique.push('-'); }
            rule["id"] = json!(unique);
            // Keep originalId at the source identity so the writer remaps copied GUIDs.
            warnings.push("Disambiguated duplicate XLSX conditional-format identities".into());
        }
        rule
    }).collect()
}

fn finish_model(output: &mut Value, rule: &Node, group: &Node, dxf: Option<&Node>, x14: Option<&Node>, theme_xml: Option<&str>, theme: Option<&Node>, storage: &str) {
    let references_theme = uses_theme(rule) || dxf.is_some_and(uses_theme) || x14.is_some_and(uses_theme);
    let theme_xml = theme_xml.filter(|_| references_theme);
    let theme = theme.filter(|_| references_theme);
    let hashes = ProjectionHashes::of(rule, group, dxf, x14, theme);
    if let Some(projection) = read_projection(rule) { restore_projection(output, &projection, &hashes); }
    let original_id = output["id"].as_str().unwrap().to_owned();
    let mut rule = rule.clone(); strip_projection(&mut rule);
    let group = skeleton(group);
    let dxf = dxf.map(|dxf| { let mut dxf = dxf.clone(); strip_projection(&mut dxf); dxf.render() });
    let x14 = x14.map(|group| { let mut group = group.clone(); strip_projection(&mut group); group.render() });
    let source = OfficeSource { version: 1, format: "xlsx".into(), original_id, storage: storage.into(), rule_xml: rule.render(),
        group_xml: Some(group.render()), differential_format_xml: dxf, x14_xml: x14, theme_xml: theme_xml.map(str::to_owned),
        baseline: known_snapshot(output), extensions: serde_json::Map::new() };
    output["officeSource"] = serde_json::to_value(source).unwrap();
}

fn parse_model(rule: &Node, linked: Option<&Node>, rule_ranges: &[Value], dxf: Option<&Node>, theme: Option<&Node>, priority: u32, warnings: &mut Vec<String>) -> Value {
    let xml_type = rule.attrs.get("type").map(String::as_str).unwrap_or_default();
    let (mut rule_type, mut sub_type) = match xml_type {
        "cellIs" => ("cellIs", Some("number")), "expression" => ("expression", Some("formula")),
        "containsText" => match rule.attrs.get("operator").map(String::as_str).unwrap_or("containsText") {
            "notContains" => ("notContainsText", Some("text")), "beginsWith" => ("beginsWith", Some("text")),
            "endsWith" => ("endsWith", Some("text")), _ => ("containsText", Some("text")),
        },
        "notContainsText" => ("notContainsText", Some("text")), "beginsWith" => ("beginsWith", Some("text")), "endsWith" => ("endsWith", Some("text")),
        "containsBlanks" | "notContainsBlanks" | "containsErrors" | "notContainsErrors" => ("containsText", Some("text")),
        "duplicateValues" => ("duplicateValues", None), "uniqueValues" => ("uniqueValues", None),
        "timePeriod" => ("timePeriod", Some("timePeriod")), "top10" => ("top10", Some("rank")),
        "aboveAverage" => ("average", Some("average")), "dataBar" => ("dataBar", None), "colorScale" => ("colorScale", None), "iconSet" => ("iconSet", None),
        _ => ("office", None),
    };
    let operator = match xml_type {
        "aboveAverage" => match (flag(rule, "aboveAverage", true), flag(rule, "equalAverage", false)) {
            (true, true) => "greaterThanOrEqual", (true, false) => "greaterThan", (false, true) => "lessThanOrEqual", (false, false) => "lessThan",
        },
        "containsBlanks" | "notContainsBlanks" | "containsErrors" | "notContainsErrors" => xml_type,
        "notContainsText" => "notContainsText", "beginsWith" => "beginsWith", "endsWith" => "endsWith",
        "containsText" if rule_type == "notContainsText" => "notContainsText",
        _ => rule.attrs.get("operator").or_else(|| rule.attrs.get("timePeriod")).map(String::as_str).unwrap_or(""),
    };
    let formulas = rule.children.iter().filter(|node| node.is_main("formula") || node.is(XM_NAMESPACE, "f")).map(|node| node.text.trim().to_owned()).collect::<Vec<_>>();
    let unsupported = rule_type == "office"
        || xml_type == "cellIs" && (!matches!(operator, "equal" | "notEqual" | "greaterThan" | "greaterThanOrEqual" | "lessThan" | "lessThanOrEqual" | "between" | "notBetween")
            || formulas.is_empty() || matches!(operator, "between" | "notBetween") && formulas.len() < 2)
        || xml_type == "expression" && formulas.is_empty()
        || xml_type == "aboveAverage" && rule.attrs.get("stdDev").is_some_and(|value| value != "0")
        || matches!(xml_type, "containsText" | "notContainsText" | "beginsWith" | "endsWith") && !rule.attrs.contains_key("text");
    if unsupported {
        rule_type = "office"; sub_type = None;
        warnings.push(format!("Retained XLSX conditional-format rule {xml_type} without local evaluation"));
    }
    let mut output = json!({ "id": "", "type": rule_type, "ranges": rule_ranges, "operator": operator,
        "formula1": formulas.first().cloned().unwrap_or_default(), "style": dxf.map(|dxf| differential_style(dxf, theme)).unwrap_or_else(|| json!({})),
        "priority": priority, "stopIfTrue": flag(rule, "stopIfTrue", false) });
    let first = &rule_ranges[0];
    let earliest = rule_ranges.iter().min_by_key(|range| (range["startRow"].as_u64(), range["startCol"].as_u64())).unwrap();
    if first["startRow"] != earliest["startRow"] || first["startCol"] != earliest["startCol"] {
        output["formulaAnchor"] = json!({ "row": first["startRow"], "col": first["startCol"] });
    }
    if let Some(sub_type) = sub_type { output["subType"] = json!(sub_type); }
    if let Some(formula) = formulas.get(1) { output["formula2"] = json!(formula); }
    if sub_type == Some("text") && let Some(text) = rule.attrs.get("text") { output["value"] = json!(text); }
    if let Some(rank) = rule.attrs.get("rank").and_then(|value| value.parse::<u32>().ok()) { output["value"] = json!(rank); }
    if flag(rule, "bottom", false) { output["isBottom"] = json!(true); }
    if flag(rule, "percent", false) { output["isPercent"] = json!(true); }
    if let Some(config) = parse_visual_config(linked.unwrap_or(rule), theme).or_else(|| parse_visual_config(rule, theme)) { output["config"] = config; }
    output
}

fn threshold(node: &Node) -> Value {
    let mut output = json!({ "type": node.attrs.get("type").map(String::as_str).unwrap_or("num") });
    let value = node.children.iter().find(|child| child.is(XM_NAMESPACE, "f")).map(|node| node.text.trim().to_owned()).or_else(|| node.attrs.get("val").cloned());
    if let Some(value) = value { output["value"] = json!(value); }
    output
}
fn visual_child(rule: &Node, name: &str) -> Option<Node> {
    rule.children.iter().find(|child| child.is_main(name) || child.is(X14_NAMESPACE, name)).cloned()
}
fn parse_visual_config(rule: &Node, theme: Option<&Node>) -> Option<Value> {
    if let Some(node) = visual_child(rule, "dataBar") {
        let values = node.children.iter().filter(|child| child.is_main("cfvo") || child.is(X14_NAMESPACE, "cfvo")).collect::<Vec<_>>();
        let mut config = json!({ "isShowValue": flag(&node, "showValue", true), "isGradient": flag(&node, "gradient", true) });
        if let Some(value) = values.first() { config["min"] = threshold(value); }
        if let Some(value) = values.get(1) { config["max"] = threshold(value); }
        for (xml_name, field) in [("color", "positiveColor"), ("fillColor", "positiveColor"), ("negativeFillColor", "nativeColor"),
            ("borderColor", "borderColor"), ("negativeBorderColor", "negativeBorderColor"), ("axisColor", "axisColor")] {
            if let Some(color) = node.children.iter().find(|child| (child.is_main(xml_name) || child.is(X14_NAMESPACE, xml_name))).and_then(|child| color_value(child, theme)) { config[field] = json!(color); }
        }
        for key in ["minLength", "maxLength"] { if let Some(value) = node.attrs.get(key).and_then(|value| value.parse::<u32>().ok()) { config[key] = json!(value); } }
        for key in ["direction", "axisPosition"] { if let Some(value) = node.attrs.get(key) { config[key] = json!(value); } }
        for key in ["border", "negativeBarColorSameAsPositive", "negativeBarBorderColorSameAsPositive"] {
            if node.attrs.contains_key(key) { config[key] = json!(flag(&node, key, false)); }
        }
        return Some(config)
    }
    if let Some(node) = visual_child(rule, "colorScale") {
        let values = node.children.iter().filter(|child| child.is_main("cfvo") || child.is(X14_NAMESPACE, "cfvo")).collect::<Vec<_>>();
        let colors = node.children.iter().filter(|child| child.is_main("color") || child.is(X14_NAMESPACE, "color")).collect::<Vec<_>>();
        return Some(json!({ "thresholds": values.into_iter().zip(colors).filter_map(|(value, color)| Some(json!({
            "value": threshold(value), "color": color_value(color, theme)?
        }))).collect::<Vec<_>>() }));
    }
    if let Some(node) = visual_child(rule, "iconSet") {
        let values = node.children.iter().filter(|child| child.is_main("cfvo") || child.is(X14_NAMESPACE, "cfvo")).collect::<Vec<_>>();
        let custom = node.children.iter().filter(|child| child.is(X14_NAMESPACE, "cfIcon")).collect::<Vec<_>>();
        let set = local_icon_type(node.attrs.get("iconSet").map(String::as_str).unwrap_or("3TrafficLights1"));
        let icons = values.iter().enumerate().rev().map(|(index, value)| {
            let icon = custom.get(index);
            let name = icon.and_then(|icon| icon.attrs.get("iconSet").map(|name| local_icon_type(name))).unwrap_or(set);
            let authored_index = icon.and_then(|icon| icon.attrs.get("iconId")).and_then(|id| id.parse::<usize>().ok()).unwrap_or(index);
            json!({ "value": threshold(value), "iconType": name, "iconId": if name == "EMPTY_ICON_TYPE" { String::new() } else { icon_index(name, authored_index).unwrap_or(authored_index).to_string() },
                "operator": if flag(value, "gte", true) { "greaterThanOrEqual" } else { "greaterThan" } })
        }).collect::<Vec<_>>();
        return Some(json!({ "icons": icons, "iconSet": set, "reverse": flag(&node, "reverse", false), "isShowValue": flag(&node, "showValue", true) }))
    }
    None
}

fn differential_style(node: &Node, theme: Option<&Node>) -> Value {
    let font = node.children.iter().find(|child| child.is_main("font"));
    let fill = node.children.iter().find(|child| child.is_main("fill"));
    let mut style = serde_json::Map::new();
    if let Some(color) = font.and_then(|font| font.children.iter().find(|child| child.is_main("color"))).and_then(|color| color_value(color, theme)) {
        style.insert("fontColor".into(), json!(color));
    }
    for (xml_name, field) in [("b", "bold"), ("i", "italic"), ("u", "underline"), ("strike", "strike")] {
        if let Some(node) = font.and_then(|font| font.children.iter().find(|child| child.is_main(xml_name))) {
            let enabled = if xml_name == "u" { node.attrs.get("val").is_none_or(|value| value != "none" && value != "0" && !value.eq_ignore_ascii_case("false")) }
                else { flag(node, "val", true) };
            style.insert(field.into(), json!(enabled));
        }
    }
    if let Some(color) = fill.and_then(|fill| fill.descendants("fgColor").into_iter().find(|child| child.is_main("fgColor"))).and_then(|color| color_value(&color, theme)) {
        style.insert("fill".into(), json!(color));
    }
    Value::Object(style)
}

fn parse_range(value: &str) -> Option<Value> {
    let mut parts = value.split(':');
    let (start_row, start_col) = parse_cell(parts.next()?)?;
    let (end_row, end_col) = parse_cell(parts.next().unwrap_or(value))?;
    if parts.next().is_some() || end_row < start_row || end_col < start_col { return None }
    Some(json!({ "startRow": start_row, "startCol": start_col, "endRow": end_row, "endCol": end_col }))
}

fn parse_cell(value: &str) -> Option<(u32, u32)> {
    let value = value.trim();
    let value = value.strip_prefix('$').unwrap_or(value);
    let split = value.find(|character: char| !character.is_ascii_alphabetic())?;
    let (letters, digits) = value.split_at(split);
    let digits = digits.strip_prefix('$').unwrap_or(digits);
    if letters.is_empty() || !letters.bytes().all(|byte| byte.is_ascii_alphabetic()) { return None }
    if digits.is_empty() || !digits.bytes().all(|byte| byte.is_ascii_digit()) { return None }
    let mut column = 0u32;
    for byte in letters.bytes() {
        column = column.checked_mul(26)?.checked_add(u32::from(byte.to_ascii_uppercase() - b'A' + 1))?;
    }
    let row = digits.parse::<u32>().ok()?;
    if row == 0 || row > 1_048_576 || column == 0 || column > 16_384 { return None }
    Some((row - 1, column - 1))
}

#[cfg(test)]
fn parse_xml(xml: &str) -> Node { source_xml::parse(xml).unwrap() }

fn relationships(xml: &str) -> Result<Vec<Relationship>, String> {
    if xml.is_empty() { return Ok(Vec::new()) }
    let document = source_xml::parse(xml)?;
    Ok(document.descendants("Relationship").into_iter().filter(|node| matches!(node.namespace.as_str(),
        "http://schemas.openxmlformats.org/package/2006/relationships" | "http://purl.oclc.org/ooxml/package/relationships" | "")).filter_map(|node| Some(Relationship {
        id: node.attrs.get("Id")?.clone(), target: node.attrs.get("Target")?.clone(),
        kind: node.attrs.get("Type")?.rsplit('/').next()?.to_owned(),
        external: node.attrs.get("TargetMode").is_some_and(|value| value == "External"),
    })).collect())
}

fn parse_part(archive: &mut ZipArchive<std::io::Cursor<&[u8]>>, path: &str) -> Result<Option<String>, String> {
    let mut entry = match archive.by_name(path) { Ok(entry) => entry, Err(zip::result::ZipError::FileNotFound) => return Ok(None), Err(error) => return Err(format!("Could not read XLSX package part {path}: {error}")) };
    if entry.size() > 16 * 1024 * 1024 { return Err(format!("XLSX metadata XML part {path} exceeds the parser limit")) }
    let mut xml = String::with_capacity(entry.size() as usize);
    entry.read_to_string(&mut xml).map_err(|error| format!("Could not decode XLSX package part {path}: {error}"))?;
    Ok(Some(xml))
}

fn resolve(source: &str, target: &str) -> Option<String> {
    if target.contains(':') || target.starts_with("//") { return None }
    let mut parts = if target.starts_with('/') { Vec::new() } else { source.rsplit_once('/').map(|(directory, _)| directory.split('/').collect()).unwrap_or_default() };
    for part in target.trim_start_matches('/').split('/') { match part { "" | "." => {}, ".." => { parts.pop()?; }, value => parts.push(value) } }
    (!parts.is_empty()).then(|| parts.join("/"))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_conditional_format_ranges_and_differential_style() {
        let style = parse_xml("<styleSheet><dxfs count=\"1\"><dxf><font><b/><color rgb=\"FFFFFFFF\"/></font><fill><patternFill patternType=\"solid\"><fgColor rgb=\"FF00AA55\"/></patternFill></fill></dxf></dxfs></styleSheet>");
        let rule_xml = "<worksheet><conditionalFormatting sqref=\"$A$1:$A$5 $C$1:$C$5\"><cfRule type=\"cellIs\" dxfId=\"0\" priority=\"3\" operator=\"between\" stopIfTrue=\"1\"><formula>10</formula><formula>20</formula></cfRule></conditionalFormatting></worksheet>";
        let mut warnings = Vec::new();
        let styles = style.descendants("dxf");
        let rules = parse_rules(rule_xml, &styles, &mut warnings);
        assert!(warnings.is_empty());
        assert_eq!(rules[0]["ranges"], json!([
            { "startRow": 0, "startCol": 0, "endRow": 4, "endCol": 0 },
            { "startRow": 0, "startCol": 2, "endRow": 4, "endCol": 2 }
        ]));
        assert_eq!(rules[0]["formula1"], "10");
        assert_eq!(rules[0]["formula2"], "20");
        assert_eq!(rules[0]["priority"], 3);
        assert_eq!(rules[0]["stopIfTrue"], true);
        assert_eq!(rules[0]["style"], json!({ "fill": "#00AA55", "fontColor": "#FFFFFF", "bold": true }));
    }

    #[test]
    fn reordered_sqref_retains_the_original_formula_basis_for_every_rule_in_the_group() {
        let xml = r#"<worksheet><conditionalFormatting sqref="A2:A4 A1"><cfRule type="expression" priority="1"><formula>A1&gt;5</formula></cfRule><cfRule type="cellIs" priority="2" operator="between"><formula>A1</formula><formula>$B$1+1</formula></cfRule></conditionalFormatting></worksheet>"#;
        let mut warnings = Vec::new();
        let rules = parse_rules(xml, &[], &mut warnings);
        assert!(warnings.is_empty(), "{warnings:?}");
        assert_eq!(rules.len(), 2);
        for rule in &rules {
            assert_eq!(rule["formulaAnchor"], json!({ "row": 1, "col": 0 }));
            assert_eq!(rule["ranges"], json!([
                { "startRow": 1, "startCol": 0, "endRow": 3, "endCol": 0 },
                { "startRow": 0, "startCol": 0, "endRow": 0, "endCol": 0 }
            ]));
        }
        assert_eq!(rules[0]["formula1"], "A1>5");
        assert_eq!(rules[1]["formula2"], "$B$1+1");
    }

    #[test]
    fn normal_sqref_keeps_the_implicit_basis_and_exact_worksheet_upper_limits() {
        let xml = r#"<worksheet><conditionalFormatting sqref="$A$1:$A$4 $XFD$1048576"><cfRule type="expression" priority="1"><formula>ROW()+COLUMN()</formula></cfRule></conditionalFormatting></worksheet>"#;
        let mut warnings = Vec::new();
        let rules = parse_rules(xml, &[], &mut warnings);
        assert!(warnings.is_empty(), "{warnings:?}");
        assert!(rules[0].get("formulaAnchor").is_none());
        assert_eq!(rules[0]["ranges"][1], json!({ "startRow": 1_048_575, "startCol": 16_383, "endRow": 1_048_575, "endCol": 16_383 }));
    }

    #[test]
    fn invalid_sqref_rejects_the_whole_group_instead_of_changing_its_formula_origin() {
        for invalid in ["A0 A2:A4", "A2:A4 XFE1", "A2:A4 A1048577", "A$1$ A2:A4", "A$$1 A2:A4"] {
            let xml = format!(r#"<worksheet><conditionalFormatting sqref="{invalid}"><cfRule type="expression" priority="1"><formula>A1</formula></cfRule></conditionalFormatting></worksheet>"#);
            let mut warnings = Vec::new();
            assert!(parse_rules(&xml, &[], &mut warnings).is_empty(), "{invalid}");
            assert_eq!(warnings, ["Skipped XLSX conditional formatting with invalid sqref"]);
        }
    }

    #[test]
    fn visual_flags_are_boolean_and_formula_thresholds_keep_the_nondefault_origin() {
        let xml = r#"<worksheet><conditionalFormatting sqref="B2:B4 A1"><cfRule type="dataBar" priority="1"><dataBar showValue="false" gradient="true"><cfvo type="formula" val="A1"/><cfvo type="num" val="100"/><color rgb="FF123456"/><negativeFillColor rgb="FF334455"/></dataBar></cfRule><cfRule type="iconSet" priority="2"><iconSet iconSet="3Arrows" showValue="0" reverse="true"><cfvo type="num" val="0"/><cfvo type="num" val="40"/><cfvo type="formula" val="A1"/></iconSet></cfRule></conditionalFormatting></worksheet>"#;
        let mut warnings = Vec::new();
        let rules = parse_rules(xml, &[], &mut warnings);
        assert!(warnings.is_empty(), "{warnings:?}");
        assert_eq!(rules[0]["formulaAnchor"], json!({ "row": 1, "col": 1 }));
        assert_eq!(rules[0]["config"]["min"], json!({ "type": "formula", "value": "A1" }));
        assert_eq!(rules[0]["config"]["isShowValue"], false);
        assert_eq!(rules[0]["config"]["isGradient"], true);
        assert_eq!(rules[0]["config"]["nativeColor"], "#334455");
        assert_eq!(rules[1]["config"]["reverse"], true);
        assert_eq!(rules[1]["config"]["isShowValue"], false);
        assert_eq!(rules[1]["config"]["icons"][0]["value"]["value"], "A1");
    }

    fn exported_rules(rules: Value) -> Vec<u8> {
        let model = json!({ "kind": "workbook", "name": "Conditional formula boundaries", "sheets": [{
            "id": "data", "name": "Data", "cells": {
                "0:0": { "row": 0, "col": 0, "value": 7 },
                "0:1": { "row": 0, "col": 1, "value": 3 },
                "1:0": { "row": 1, "col": 0, "value": 2 },
                "2:0": { "row": 2, "col": 0, "value": 8 },
                "3:0": { "row": 3, "col": 0, "value": 1 }
            }, "conditionalFormats": rules
        }] });
        let (bytes, warnings) = office_writer::export_xlsx_bytes(&model.to_string()).unwrap();
        assert!(warnings.is_empty(), "{warnings:?}");
        bytes
    }

    fn worksheet_xml(bytes: &[u8]) -> String {
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
        let mut xml = String::new();
        archive.by_name("xl/worksheets/sheet1.xml").unwrap().read_to_string(&mut xml).unwrap();
        xml
    }

    fn reopened_rules(bytes: &[u8]) -> Vec<Value> {
        let payload = serde_json::to_value(crate::parser::parse_office_bytes(bytes, "conditional-origin.xlsx").unwrap()).unwrap();
        assert!(payload["diagnostics"]["errors"].as_array().unwrap().is_empty(), "{payload}");
        payload["payload"]["data"]["sheets"][0]["conditionalFormats"].as_array().unwrap().clone()
    }

    fn covered_cells(rule: &Value) -> std::collections::BTreeSet<(u64, u64)> {
        let mut result = std::collections::BTreeSet::new();
        for range in rule["ranges"].as_array().unwrap() {
            for row in range["startRow"].as_u64().unwrap()..=range["endRow"].as_u64().unwrap() {
                for col in range["startCol"].as_u64().unwrap()..=range["endCol"].as_u64().unwrap() {
                    result.insert((row, col));
                }
            }
        }
        result
    }

    #[test]
    fn native_binary_roundtrip_preserves_partially_valid_rows_without_splitting_the_rule() {
        let source = json!({
            "id": "partial-rows", "type": "expression",
            "ranges": [{ "startRow": 0, "startCol": 0, "endRow": 3, "endCol": 0 }],
            "formulaAnchor": { "row": 1, "col": 0 }, "formula1": "=A1>5",
            "style": { "fill": "#123456" }, "priority": 1, "stopIfTrue": true
        });
        let bytes = exported_rules(json!([source]));
        let xml = worksheet_xml(&bytes);
        assert_eq!(xml.matches("<conditionalFormatting ").count(), 1);
        assert_eq!(xml.matches("<cfRule ").count(), 1);
        assert!(xml.contains(r#"sqref="A2:A4 A1""#), "{xml}");
        assert!(!xml.contains("formulaAnchor"), "{xml}");
        let rules = reopened_rules(&bytes);
        assert_eq!(rules.len(), 1);
        assert_eq!(rules[0]["formulaAnchor"], json!({ "row": 1, "col": 0 }));
        assert_eq!(rules[0]["formula1"], "A1>5");
        assert_eq!(covered_cells(&rules[0]), covered_cells(&source));
        let again = reopened_rules(&exported_rules(json!(rules)));
        assert_eq!(again[0]["formulaAnchor"], json!({ "row": 1, "col": 0 }));
        assert_eq!(again[0]["formula1"], "A1>5");
        assert_eq!(covered_cells(&again[0]), covered_cells(&source));
    }

    #[test]
    fn native_binary_roundtrip_preserves_independent_iferror_operands_with_no_safe_sqref_origin() {
        let source = json!({
            "id": "opposing-bounds", "type": "expression",
            "ranges": [{ "startRow": 0, "startCol": 0, "endRow": 0, "endCol": 0 }],
            "formulaAnchor": { "row": 1, "col": 0 }, "formula1": "=IFERROR(A1,0)+IFERROR(A1048576,0)+$B$1",
            "style": { "bold": true }, "priority": 1
        });
        let bytes = exported_rules(json!([source]));
        let xml = worksheet_xml(&bytes);
        assert_eq!(xml.matches("<cfRule ").count(), 1);
        assert!(xml.contains("INDEX($A$1:$XFD$1048576"), "{xml}");
        assert!(xml.contains("IF(ROW()-1&lt;1,1048577,ROW()-1)"), "{xml}");
        let rules = reopened_rules(&bytes);
        assert_eq!(rules[0]["formula1"].as_str().unwrap().matches("IFERROR(").count(), 2);
        assert!(rules[0]["formula1"].as_str().unwrap().ends_with("+$B$1"));
        assert_eq!(covered_cells(&rules[0]), covered_cells(&source));
        let again = reopened_rules(&exported_rules(json!(rules)));
        assert_eq!(again[0]["formula1"], rules[0]["formula1"]);
        assert_eq!(covered_cells(&again[0]), covered_cells(&source));
    }

    #[test]
    fn native_visual_roundtrip_preserves_formula_origins_statistics_scopes_and_boolean_flags() {
        let ranges = json!([{ "startRow": 0, "startCol": 0, "endRow": 3, "endCol": 0 }]);
        let anchor = json!({ "row": 1, "col": 0 });
        let source = json!([
            { "id": "bar", "type": "dataBar", "ranges": ranges, "formulaAnchor": anchor, "priority": 1,
                "config": { "min": { "type": "formula", "value": "=A1" }, "max": { "type": "num", "value": 100 },
                    "isGradient": false, "positiveColor": "#123456", "nativeColor": "#334455", "isShowValue": false } },
            { "id": "scale", "type": "colorScale", "ranges": ranges, "formulaAnchor": anchor, "priority": 2,
                "config": { "thresholds": [{ "value": { "type": "formula", "value": "=A1" }, "color": "#123456" },
                    { "value": { "type": "percentile", "value": 50 }, "color": "#888888" },
                    { "value": { "type": "num", "value": 100 }, "color": "#ffffff" }] } },
            { "id": "icons", "type": "iconSet", "ranges": ranges, "formulaAnchor": anchor, "priority": 3,
                "config": { "iconSet": "3Arrows", "reverse": true, "isShowValue": false, "icons": [
                    { "value": { "type": "formula", "value": "=A1" }, "operator": "greaterThanOrEqual", "iconType": "3Arrows", "iconId": "0" },
                    { "value": { "type": "num", "value": 40 }, "operator": "greaterThanOrEqual", "iconType": "3Arrows", "iconId": "1" },
                    { "value": { "type": "num", "value": 0 }, "iconType": "3Arrows", "iconId": "2" }
                ] } }
        ]);
        let bytes = exported_rules(source.clone());
        let xml = worksheet_xml(&bytes);
        assert_eq!(xml.matches("<cfRule ").count(), 3);
        assert_eq!(xml.matches(r#"sqref="A2:A4 A1""#).count(), 3);
        let rules = reopened_rules(&bytes);
        for (index, rule) in rules.iter().enumerate() {
            assert_eq!(rule["formulaAnchor"], anchor);
            assert_eq!(covered_cells(rule), covered_cells(&source[index]));
        }
        assert_eq!(rules[0]["config"]["min"]["value"], "A1");
        assert_eq!(rules[0]["config"]["isGradient"], false);
        assert_eq!(rules[0]["config"]["nativeColor"], "#334455");
        assert_eq!(rules[1]["config"]["thresholds"][0]["value"]["value"], "A1");
        assert_eq!(rules[2]["config"]["icons"][0]["value"]["value"], "A1");
        assert_eq!(rules[2]["config"]["reverse"], true);
        assert_eq!(rules[2]["config"]["isShowValue"], false);
    }

    #[test]
    fn equal_priorities_keep_the_model_stack_order_in_native_binary_output() {
        let source = json!([
            { "id": "z-first", "type": "expression", "priority": 1, "ranges": [{ "startRow": 0, "startCol": 0, "endRow": 0, "endCol": 0 }], "formula1": "10" },
            { "id": "a-second", "type": "expression", "priority": 1, "ranges": [{ "startRow": 0, "startCol": 0, "endRow": 0, "endCol": 0 }], "formula1": "20" }
        ]);
        let rules = reopened_rules(&exported_rules(source));
        assert_eq!(rules.iter().map(|rule| rule["formula1"].as_str().unwrap()).collect::<Vec<_>>(), ["10", "20"]);
        assert_eq!(rules.iter().map(|rule| rule["priority"].as_u64().unwrap()).collect::<Vec<_>>(), [1, 2]);
    }
}
