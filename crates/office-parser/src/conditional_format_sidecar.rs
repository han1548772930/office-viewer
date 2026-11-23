use std::collections::{HashMap, HashSet};
use std::io::{BufReader, Read};

use office_io_common::conditional_format_color::{color_value, uses_theme, validate_theme_references};
use office_io_common::conditional_format_source::{CONDITIONAL_MODEL_BYTE_LIMIT, SOURCE_RESOURCE_BYTE_LIMIT, SOURCE_RESOURCE_ELEMENT_LIMIT, ConditionalFormatModelBudget, OfficeSource, ProjectionHashes, icon_index, known_snapshot, local_icon_type, read_projection, restore_projection, strip_projection, valid_known_fields, valid_saved_model};
use office_io_common::conditional_format_xml::{self as source_xml, Element as Node, RULE_LINK_EXTENSION_URI, X14_NAMESPACE, XM_NAMESPACE, XmlName, XmlSelection, XmlSelectionLimits};
use serde_json::{Value, json};
use zip::ZipArchive;

const REL_NAMESPACE: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_REL_NAMESPACE: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const REVISION_NAMESPACE: &str = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision";
const MC_NAMESPACE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const PACKAGE_REL_NAMESPACE: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const STRICT_PACKAGE_REL_NAMESPACE: &str = "http://purl.oclc.org/ooxml/package/relationships";
const DRAWING_NAMESPACE: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const STRICT_DRAWING_NAMESPACE: &str = "http://purl.oclc.org/ooxml/drawingml/main";

#[derive(Debug, Default)]
pub(crate) struct ConditionalFormatSidecar {
    by_sheet: HashMap<usize, Vec<Value>>,
    pub(crate) warnings: Vec<String>,
    /// The caller must reject the import; an empty success would lose declared sources.
    pub(crate) failure: Option<String>,
}

#[derive(Debug)]
struct Relationship { id: String, target: String, kind: String }

#[derive(Clone, Copy)]
struct SidecarLimits { part: XmlSelectionLimits, resource_bytes: usize, resource_elements: usize, model_bytes: usize }
impl Default for SidecarLimits {
    fn default() -> Self { Self { part: XmlSelectionLimits::default(), resource_bytes: SOURCE_RESOURCE_BYTE_LIMIT,
        resource_elements: SOURCE_RESOURCE_ELEMENT_LIMIT, model_bytes: CONDITIONAL_MODEL_BYTE_LIMIT } }
}
struct ReadBudget { limits: SidecarLimits, bytes: usize, elements: usize }
impl ReadBudget {
    fn include(&mut self, selected: &XmlSelection) -> Result<(), String> {
        self.bytes = self.bytes.checked_add(selected.selected_bytes).ok_or("CF_SOURCE_TOTAL_LIMIT: selected resource byte counter overflow")?;
        self.elements = self.elements.checked_add(selected.selected_elements).ok_or("CF_SOURCE_TOTAL_LIMIT: selected resource element counter overflow")?;
        if self.bytes > self.limits.resource_bytes || self.elements > self.limits.resource_elements {
            return Err("CF_SOURCE_TOTAL_LIMIT: selected XLSX conditional-format resources exceed the workbook budget".into())
        }
        Ok(())
    }
}

fn invalid_xml(message: impl std::fmt::Display) -> String { format!("CF_SOURCE_XML_INVALID: {message}") }
fn required_resource(message: impl std::fmt::Display) -> String { format!("CF_SOURCE_REQUIRED_RESOURCE: {message}") }
fn package_error(message: impl std::fmt::Display) -> String { format!("CF_SOURCE_PACKAGE_INVALID: {message}") }

impl ConditionalFormatSidecar {
    pub(crate) fn read(bytes: &[u8]) -> Self {
        match read_sidecar(bytes) {
            Ok(value) => value,
            Err(error) => Self { warnings: vec![error.clone()], failure: Some(error), ..Self::default() },
        }
    }
    pub(crate) fn take(&mut self, sheet_index: usize) -> Vec<Value> { self.by_sheet.remove(&sheet_index).unwrap_or_default() }
}

fn read_sidecar(bytes: &[u8]) -> Result<ConditionalFormatSidecar, String> {
    read_sidecar_with_limits(bytes, SidecarLimits::default())
}

fn read_sidecar_with_limits(bytes: &[u8], limits: SidecarLimits) -> Result<ConditionalFormatSidecar, String> {
    let mut archive = ZipArchive::new(std::io::Cursor::new(bytes))
        .map_err(|error| package_error(format!("invalid XLSX ZIP while reading conditional formatting: {error}")))?;
    let mut budget = ReadBudget { limits, bytes: 0, elements: 0 };
    let workbook = selected_part(&mut archive, "xl/workbook.xml", &mut budget, |path, _| Ok(path.len() == 3
        && path[0].is_main("workbook") && path[1].is_main("sheets") && path[2].is_main("sheet")))?
        .ok_or_else(|| required_resource("missing XLSX workbook part"))?;
    if !workbook.root.is_main("workbook") { return Err(invalid_xml("invalid XLSX workbook root")) }
    let sheets = workbook.fragments;
    let sheet_ids = sheets.iter().map(|sheet| sheet.attribute_ns(REL_NAMESPACE, "id").or_else(|| sheet.attribute_ns(STRICT_REL_NAMESPACE, "id"))
        .map(str::to_owned).ok_or_else(|| required_resource("XLSX sheet has no workbook relationship"))).collect::<Result<HashSet<_>, _>>()?;
    if sheet_ids.len() != sheets.len() { return Err(invalid_xml("duplicate XLSX sheet relationship IDs")) }
    let workbook_rels = selected_part(&mut archive, "xl/_rels/workbook.xml.rels", &mut budget, |path, node| {
        if path.len() != 2 || !package_relationship(&path[0], "Relationships") || !package_relationship(&path[1], "Relationship") { return Ok(false) }
        if node.attrs.get("TargetMode").is_some_and(|mode| mode.eq_ignore_ascii_case("External")) { return Ok(false) }
        Ok(node.attrs.get("Id").is_some_and(|id| sheet_ids.contains(id))
            || node.attrs.get("Type").and_then(|kind| standard_relationship_kind(kind)).is_some_and(|kind| matches!(kind, "styles" | "theme")))
    })?;
    let relation_map = match workbook_rels {
        Some(selected) => {
            if !package_relationship(&XmlName { namespace: selected.root.namespace.clone(), name: selected.root.name.clone() }, "Relationships") {
                return Err(invalid_xml("invalid XLSX workbook relationships root"))
            }
            let mut relationships_by_id = HashMap::new();
            for node in selected.fragments {
                let relationship = relationship(&node)?;
                let path = resolve("xl/workbook.xml", &relationship.target).ok_or_else(|| package_error("invalid internal XLSX relationship target"))?;
                if relationships_by_id.insert(relationship.id, (relationship.kind, path)).is_some() { return Err(invalid_xml("duplicate XLSX workbook relationship IDs")) }
            }
            relationships_by_id
        }
        None if sheets.is_empty() => HashMap::new(),
        None => return Err(required_resource("missing XLSX workbook relationships part")),
    };
    // Workbook sheets/relationships have their own metadata budget. The shared
    // CF resource budget below counts worksheet fragments, referenced dxfs/theme.
    let mut budget = ReadBudget { limits, bytes: 0, elements: 0 };
    let mut worksheets = Vec::new();
    for (index, sheet) in sheets.into_iter().enumerate() {
        let relationship_id = sheet.attribute_ns(REL_NAMESPACE, "id").or_else(|| sheet.attribute_ns(STRICT_REL_NAMESPACE, "id"))
            .ok_or_else(|| required_resource("XLSX sheet has no workbook relationship"))?;
        let (kind, path) = relation_map.get(relationship_id).ok_or_else(|| required_resource("XLSX sheet references an unavailable internal relationship"))?;
        if kind != "worksheet" { continue }
        let selected = selected_part(&mut archive, path, &mut budget, worksheet_group)?
            .ok_or_else(|| required_resource(format!("missing XLSX worksheet {path}")))?;
        if !selected.root.is_main("worksheet") { return Err(invalid_xml(format!("invalid worksheet root in {path}"))) }
        let mut root = selected.root;
        for group in selected.fragments { root.push(group); }
        worksheets.push((index, path.clone(), root));
    }
    let dxf_ids = referenced_dxfs(worksheets.iter().map(|(_, _, root)| root))?;
    let dxfs = if dxf_ids.is_empty() { HashMap::new() } else {
        let style_path = relation_map.values().find_map(|(kind, path)| (kind == "styles").then_some(path.as_str()))
            .ok_or_else(|| required_resource("conditional-format dxfId references have no styles relationship"))?;
        let mut index = 0usize;
        let mut selected_indices = Vec::new();
        let styles = selected_part(&mut archive, style_path, &mut budget, |path, _| {
            if path.len() != 3 || !path[0].is_main("styleSheet") || !path[1].is_main("dxfs") || !path[2].is_main("dxf") { return Ok(false) }
            let original_index = index;
            index = index.checked_add(1).ok_or("CF_SOURCE_TOTAL_LIMIT: differential-style index counter overflow")?;
            let selected = dxf_ids.contains(&original_index);
            if selected { selected_indices.push(original_index); }
            Ok(selected)
        })?.ok_or_else(|| required_resource(format!("missing referenced XLSX styles part {style_path}")))?;
        if !styles.root.is_main("styleSheet") { return Err(invalid_xml("invalid XLSX styles root")) }
        let dxfs = selected_indices.into_iter().zip(styles.fragments).collect::<HashMap<_, _>>();
        if let Some(index) = dxf_ids.iter().filter(|index| !dxfs.contains_key(*index)).min() {
            return Err(required_resource(format!("conditional-format dxfId {index} has no differential style")))
        }
        dxfs
    };
    let needs_theme = worksheets.iter().any(|(_, _, root)| root.children.iter().any(uses_theme)) || dxfs.values().any(uses_theme);
    let theme = if needs_theme {
        let theme_path = relation_map.values().find_map(|(kind, path)| (kind == "theme").then_some(path.as_str()))
            .ok_or_else(|| required_resource("conditional-format theme references have no theme relationship"))?;
        let selected = selected_part(&mut archive, theme_path, &mut budget, |path, _| Ok(path.len() == 1))?
            .ok_or_else(|| required_resource(format!("missing referenced XLSX theme part {theme_path}")))?;
        let theme = selected.fragments.into_iter().next().ok_or_else(|| required_resource("empty XLSX theme source"))?;
        if theme.name != "theme" || !matches!(theme.namespace.as_str(), DRAWING_NAMESPACE | STRICT_DRAWING_NAMESPACE) {
            return Err(invalid_xml("invalid XLSX theme root"))
        }
        let xml = checked_render(&theme)?; Some((xml, theme))
    } else { None };
    let mut sidecar = ConditionalFormatSidecar::default();
    let mut model_budget = ConditionalFormatModelBudget::new(limits.model_bytes);
    for (index, path, root) in worksheets {
        let rules = parse_rules_tree(&root, &dxfs, theme.as_ref().map(|(xml, _)| xml.as_str()), theme.as_ref().map(|(_, theme)| theme),
            &path, &mut sidecar.warnings, &mut model_budget)?;
        if !rules.is_empty() { sidecar.by_sheet.insert(index, rules); }
    }
    // Publish only after every selected source and required resource succeeded.
    Ok(sidecar)
}

fn detached_rule(group: &Node, rule: &Node) -> Node {
    group.detached_child(rule)
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
fn ranges(group: &Node) -> Result<Vec<Value>, String> {
    let ranges = sqref(group).unwrap_or("").split_whitespace().map(parse_range).collect::<Option<Vec<_>>>()
        .filter(|ranges| !ranges.is_empty()).ok_or_else(|| invalid_xml("conditional formatting has invalid sqref"))?;
    if ranges.len() > 100_000 { return Err("CF_SOURCE_FRAGMENT_LIMIT: conditional-format sqref exceeds the area budget".into()) }
    Ok(ranges)
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

struct X14Rule<'a> { rule: &'a Node, group: &'a Node, used: bool }

#[cfg(test)]
fn parse_rules(xml: &str, dxfs: &[Node], warnings: &mut Vec<String>) -> Vec<Value> {
    parse_rules_with_theme(xml, dxfs, None, "xl/worksheets/sheet1.xml", warnings)
}

#[cfg(test)]
fn parse_rules_with_theme(xml: &str, dxfs: &[Node], theme_xml: Option<&str>, sheet_path: &str, warnings: &mut Vec<String>) -> Vec<Value> {
    match parse_rules_result(xml, dxfs, theme_xml, sheet_path, warnings) {
        Ok(rules) => rules,
        Err(error) => { warnings.push(error); Vec::new() },
    }
}

#[cfg(test)]
fn parse_rules_result(xml: &str, dxfs: &[Node], theme_xml: Option<&str>, sheet_path: &str, warnings: &mut Vec<String>) -> Result<Vec<Value>, String> {
    let selected = source_xml::extract(std::io::Cursor::new(xml.as_bytes()), XmlSelectionLimits::default(), worksheet_group)?;
    let mut root = selected.root;
    if !root.is_main("worksheet") { return Err(invalid_xml("invalid worksheet root")) }
    for group in selected.fragments { root.push(group); }
    let theme = theme_xml.map(source_xml::parse).transpose().map_err(invalid_xml)?;
    let dxfs = dxfs.iter().cloned().enumerate().collect::<HashMap<_, _>>();
    parse_rules_tree(&root, &dxfs, theme_xml, theme.as_ref(), sheet_path, warnings, &mut ConditionalFormatModelBudget::default())
}

fn parse_rules_tree(root: &Node, dxfs: &HashMap<usize, Node>, theme_xml: Option<&str>, theme: Option<&Node>, sheet_path: &str,
    warnings: &mut Vec<String>, model_budget: &mut ConditionalFormatModelBudget) -> Result<Vec<Value>, String> {
    for group in &root.children {
        ranges(group)?;
        if !group.children.iter().any(is_rule) { return Err(invalid_xml("conditional-format group has no rule")) }
    }
    let mut x14 = root.children.iter().filter(|group| group.is(X14_NAMESPACE, "conditionalFormatting")).flat_map(|group|
        group.children.iter().filter(|rule| rule.is(X14_NAMESPACE, "cfRule")).map(move |rule| X14Rule { group, rule, used: false })
    ).collect::<Vec<_>>();
    let mut output = Vec::<(u32, usize, Value)>::new();
    let mut ordinal = 0usize;
    for group in root.children.iter().filter(|group| group.is_main("conditionalFormatting")) {
        let group = detached_rule(&root, group);
        let rule_ranges = ranges(&group)?;
        for rule in group.children.iter().filter(|rule| rule.is_main("cfRule")) {
            ordinal += 1;
            let rule = detached_rule(&group, rule);
            let linked = linked_id(&rule).and_then(|id| x14.iter_mut().find(|candidate| !candidate.used && !candidate.rule.attrs.contains_key("priority")
                && candidate.rule.attrs.get("id").is_some_and(|candidate_id| candidate_id.trim().eq_ignore_ascii_case(&id))));
            let linked_group = linked.map(|candidate| {
                candidate.used = true; let rule = detached_rule(candidate.group, candidate.rule); single_rule_group(candidate.group, &rule)
            });
            let linked_rule = linked_group.as_ref().and_then(|group| group.children.iter().find(|rule| rule.is(X14_NAMESPACE, "cfRule")));
            let dxf = referenced_dxf(&rule, dxfs)?;
            let priority = rule_priority(&rule, ordinal)?;
            let mut model = parse_model(&rule, linked_rule, &rule_ranges, dxf, theme, priority, warnings);
            if let Some(linked_group) = &linked_group {
                if sqref(linked_group).map(|value| value.split_whitespace().collect::<Vec<_>>().join(" "))
                    != sqref(&group).map(|value| value.split_whitespace().collect::<Vec<_>>().join(" ")) {
                    model["type"] = json!("office"); model.as_object_mut().unwrap().remove("subType");
                    warnings.push("Retained a linked x14 conditional format with a different scope without local evaluation".into());
                }
            }
            model["id"] = json!(identity(&rule, &group, sheet_path, ordinal));
            finish_model(&mut model, &rule, &group, dxf, linked_group.as_ref(), theme_xml, theme, "base")?;
            model_budget.consume(&model)?;
            output.push((priority, ordinal, model));
        }
    }
    for source in x14.into_iter().filter(|source| !source.used) {
        ordinal += 1;
        let rule = detached_rule(source.group, source.rule);
        let group = single_rule_group(source.group, &rule);
        let rule_ranges = ranges(&group)?;
        let dxf = rule.children.iter().find(|child| child.is(X14_NAMESPACE, "dxf") || child.is_main("dxf"))
            .or(referenced_dxf(&rule, dxfs)?);
        let priority = rule_priority(&rule, ordinal)?;
        let mut model = parse_model(&rule, None, &rule_ranges, dxf, theme, priority, warnings);
        if !rule.attrs.contains_key("priority") {
            model["type"] = json!("office"); model.as_object_mut().unwrap().remove("subType");
            warnings.push("Retained an orphan x14 conditional-format enhancement without local evaluation".into());
        }
        model["id"] = json!(identity(&rule, &group, sheet_path, ordinal));
        finish_model(&mut model, &rule, &group, dxf, None, theme_xml, theme, "x14")?;
        model_budget.consume(&model)?;
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
            model_budget.reserve_bytes(unique.len().saturating_sub(id.len()))?;
            rule["id"] = json!(unique);
            // Keep originalId at the source identity so the writer remaps copied GUIDs.
            warnings.push("Disambiguated duplicate XLSX conditional-format identities".into());
        }
        Ok(rule)
    }).collect()
}

fn finish_model(output: &mut Value, rule: &Node, group: &Node, dxf: Option<&Node>, x14: Option<&Node>, theme_xml: Option<&str>, theme: Option<&Node>, storage: &str) -> Result<(), String> {
    let references_theme = uses_theme(rule) || dxf.is_some_and(uses_theme) || x14.is_some_and(uses_theme);
    let theme_xml = theme_xml.filter(|_| references_theme);
    let theme = theme.filter(|_| references_theme);
    if references_theme && (theme_xml.is_none() || theme.is_none()) { return Err(required_resource("conditional-format theme references have no preserved theme")) }
    for node in std::iter::once(rule).chain(dxf).chain(x14) { validate_theme_references(node, theme)?; }
    let hashes = ProjectionHashes::of(rule, group, dxf, x14, theme);
    if let Some(projection) = read_projection(rule) { restore_projection(output, &projection, &hashes); }
    let original_id = output["id"].as_str().unwrap().to_owned();
    let mut rule = rule.clone(); strip_projection(&mut rule);
    let group = skeleton(group);
    let dxf = dxf.map(|dxf| { let mut dxf = dxf.clone(); strip_projection(&mut dxf); checked_render(&dxf) }).transpose()?;
    let x14 = x14.map(|group| { let mut group = group.clone(); strip_projection(&mut group); checked_render(&group) }).transpose()?;
    let source = OfficeSource { version: 1, format: "xlsx".into(), original_id, storage: storage.into(), rule_xml: checked_render(&rule)?,
        group_xml: Some(checked_render(&group)?), differential_format_xml: dxf, x14_xml: x14, theme_xml: theme_xml.map(str::to_owned),
        baseline: known_snapshot(output), extensions: serde_json::Map::new() };
    source.validate()?;
    output["officeSource"] = serde_json::to_value(source).map_err(invalid_xml)?;
    Ok(())
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
    let unsupported = rule_type == "office" || formulas.len() > 2
        || xml_type == "cellIs" && (!matches!(operator, "equal" | "notEqual" | "greaterThan" | "greaterThanOrEqual" | "lessThan" | "lessThanOrEqual" | "between" | "notBetween")
            || formulas.is_empty() || matches!(operator, "between" | "notBetween") && formulas.len() < 2)
        || xml_type == "expression" && formulas.is_empty()
        || xml_type == "aboveAverage" && rule.attrs.get("stdDev").is_some_and(|value| value.parse::<f64>().ok() != Some(0.0))
        || xml_type == "top10" && !rule.attrs.get("rank").and_then(|value| value.parse::<u32>().ok())
            .is_some_and(|value| value > 0 && value <= if flag(rule, "percent", false) { 100 } else { 1000 })
        || xml_type == "timePeriod" && !matches!(operator, "today" | "yesterday" | "tomorrow" | "last7Days" | "thisWeek" | "lastWeek" | "nextWeek" | "thisMonth" | "lastMonth" | "nextMonth")
        || linked.is_some_and(|linked| xml_type != "dataBar" || linked.attrs.get("type").map(String::as_str) != Some("dataBar")
            || visual_child(linked, "dataBar").is_none_or(|bar| bar.children.iter().filter(|node| node.is(X14_NAMESPACE, "cfvo")).count() != 2))
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
    let base_config = parse_visual_config(rule, theme);
    let linked_config = linked.and_then(|linked| parse_visual_config(linked, theme));
    if let Some(mut config) = base_config.or_else(|| linked_config.clone()) {
        if let Some(linked) = linked_config.and_then(|config| config.as_object().cloned()) { config.as_object_mut().unwrap().extend(linked); }
        output["config"] = config;
    }
    let visual = match xml_type { "dataBar" | "colorScale" | "iconSet" => Some(xml_type), _ => None };
    if let Some(name) = visual {
        let node = visual_child(rule, name);
        let invalid = node.as_ref().is_none_or(|node| {
            let values = node.children.iter().filter(|node| node.is_main("cfvo") || node.is(X14_NAMESPACE, "cfvo")).count();
            match name {
                "dataBar" => values != 2 || node.namespace != X14_NAMESPACE && !output["config"]["positiveColor"].is_string(),
                "colorScale" => !(2..=3).contains(&values) || output["config"]["thresholds"].as_array().is_none_or(|stops| stops.len() != values),
                "iconSet" => {
                    let custom = node.children.iter().filter(|node| node.is(X14_NAMESPACE, "cfIcon")).collect::<Vec<_>>();
                    !(3..=5).contains(&values) || flag(node, "custom", false) && (custom.len() != values || custom.iter().any(|node|
                        node.attrs.get("iconSet").is_none() || node.attrs.get("iconId").and_then(|id| id.parse::<u32>().ok()).is_none()))
                }
                _ => false,
            }
        });
        if invalid {
            output["type"] = json!("office"); output.as_object_mut().unwrap().remove("subType");
            warnings.push(format!("Retained XLSX conditional-format rule {xml_type} with an uninspectable visual body without local evaluation"));
        }
    }
    let mut candidate = output.clone(); candidate["id"] = json!("inspectable-source-rule");
    if !valid_saved_model(&candidate) {
        output["type"] = json!("office"); output.as_object_mut().unwrap().remove("subType");
        if !valid_known_fields(&output) { output.as_object_mut().unwrap().remove("config"); }
        warnings.push(format!("Retained XLSX conditional-format rule {xml_type} with unsupported or invalid known fields without local evaluation"));
    }
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

fn package_relationship(name: &XmlName, local: &str) -> bool {
    name.name == local && matches!(name.namespace.as_str(), PACKAGE_REL_NAMESPACE | STRICT_PACKAGE_REL_NAMESPACE | "")
}

fn standard_relationship_kind(kind: &str) -> Option<&str> {
    let (namespace, local) = kind.rsplit_once('/')?;
    matches!(namespace, REL_NAMESPACE | STRICT_REL_NAMESPACE).then_some(local)
}

fn relationship(node: &Node) -> Result<Relationship, String> {
    let id = node.attrs.get("Id").filter(|id| !id.trim().is_empty()).ok_or_else(|| package_error("relevant XLSX relationship has no ID"))?;
    let target = node.attrs.get("Target").filter(|target| !target.trim().is_empty()).ok_or_else(|| package_error("relevant XLSX relationship has no target"))?;
    let kind = node.attrs.get("Type").filter(|kind| !kind.trim().is_empty()).ok_or_else(|| package_error("relevant XLSX relationship has no type"))?;
    if node.attrs.get("TargetMode").is_some_and(|mode| !mode.eq_ignore_ascii_case("Internal")) { return Err(package_error("relevant XLSX relationship is not internal")) }
    Ok(Relationship { id: id.clone(), target: target.clone(), kind: standard_relationship_kind(kind).unwrap_or(kind).to_owned() })
}

fn worksheet_group(path: &[XmlName], _: &Node) -> Result<bool, String> {
    if path.len() < 2 || !path[0].is_main("worksheet") { return Ok(false) }
    Ok(path.len() == 2 && path[1].is_main("conditionalFormatting")
        || path.len() >= 4 && path[1].is_main("extLst") && path.last().is_some_and(|name| name.is(X14_NAMESPACE, "conditionalFormatting")))
}

fn selected_part(archive: &mut ZipArchive<std::io::Cursor<&[u8]>>, path: &str, budget: &mut ReadBudget,
    select: impl FnMut(&[XmlName], &Node) -> Result<bool, String>) -> Result<Option<XmlSelection>, String> {
    let entry = match archive.by_name(path) {
        Ok(entry) => entry,
        Err(zip::result::ZipError::FileNotFound) => return Ok(None),
        Err(error) => return Err(package_error(format!("cannot read XLSX part {path}: {error}"))),
    };
    let mut limits = budget.limits.part;
    limits.total_bytes = limits.total_bytes.min(budget.limits.resource_bytes.saturating_sub(budget.bytes));
    limits.total_elements = limits.total_elements.min(budget.limits.resource_elements.saturating_sub(budget.elements));
    let selected = source_xml::extract(BufReader::new(entry), limits, select).map_err(|error| format!("{error} [XLSX part {path}]"))?;
    budget.include(&selected)?; Ok(Some(selected))
}

fn referenced_dxfs<'a>(roots: impl Iterator<Item = &'a Node>) -> Result<HashSet<usize>, String> {
    let mut indices = HashSet::new();
    for root in roots {
        for rule in root.children.iter().flat_map(|group| group.children.iter().filter(|node| is_rule(node))) {
            if let Some(value) = rule.attrs.get("dxfId") {
                indices.insert(dxf_index(value)?);
            }
        }
    }
    Ok(indices)
}

fn referenced_dxf<'a>(rule: &Node, dxfs: &'a HashMap<usize, Node>) -> Result<Option<&'a Node>, String> {
    let Some(value) = rule.attrs.get("dxfId") else { return Ok(None) };
    let index = dxf_index(value)?;
    dxfs.get(&index).map(Some).ok_or_else(|| required_resource(format!("conditional-format dxfId {index} has no preserved differential style")))
}

fn dxf_index(value: &str) -> Result<usize, String> {
    value.parse::<u32>().map(|index| index as usize).map_err(|_| invalid_xml("conditional-format dxfId is not a valid Office unsigned index"))
}

fn rule_priority(rule: &Node, ordinal: usize) -> Result<u32, String> {
    let priority = match rule.attrs.get("priority") {
        Some(value) => value.parse::<u32>().map_err(|_| invalid_xml("conditional-format priority is not a valid unsigned integer"))?,
        None => u32::try_from(ordinal).map_err(|_| "CF_SOURCE_TOTAL_LIMIT: conditional-format rule count exceeds the priority budget".to_owned())?,
    };
    if priority == 0 || priority > i32::MAX as u32 { return Err(invalid_xml("conditional-format priority exceeds the Office range")) }
    Ok(priority)
}

fn checked_render(node: &Node) -> Result<String, String> { node.validate_fragment()?; Ok(node.render()) }

fn resolve(source: &str, target: &str) -> Option<String> {
    if target.contains(':') || target.starts_with("//") { return None }
    let mut parts = if target.starts_with('/') { Vec::new() } else { source.rsplit_once('/').map(|(directory, _)| directory.split('/').collect()).unwrap_or_default() };
    for part in target.trim_start_matches('/').split('/') { match part { "" | "." => {}, ".." => { parts.pop()?; }, value => parts.push(value) } }
    (!parts.is_empty()).then(|| parts.join("/"))
}

#[cfg(test)]
#[path = "conditional_format_sidecar_stream_tests.rs"]
mod stream_tests;

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
            assert_eq!(warnings, ["CF_SOURCE_XML_INVALID: conditional formatting has invalid sqref"]);
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

    pub(super) fn exported_rules(rules: Value) -> Vec<u8> {
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

    pub(super) fn worksheet_xml(bytes: &[u8]) -> String {
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
        let mut xml = String::new();
        archive.by_name("xl/worksheets/sheet1.xml").unwrap().read_to_string(&mut xml).unwrap();
        xml
    }

    pub(super) fn reopened_rules(bytes: &[u8]) -> Vec<Value> {
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
        let standard = source_xml::parse(&xml).unwrap();
        assert!(standard.children.iter().filter(|node| node.is_main("conditionalFormatting"))
            .all(|group| group.attrs.get("formulaAnchor").is_none() && group.children.iter().filter(|node| node.is_main("cfRule")).all(|rule| rule.attrs.get("formulaAnchor").is_none())));
        let rules = reopened_rules(&bytes);
        assert_eq!(rules.len(), 1);
        assert_eq!(rules[0]["formulaAnchor"], json!({ "row": 1, "col": 0 }));
        assert_eq!(rules[0]["id"], "partial-rows"); assert_eq!(rules[0]["formula1"], "=A1>5");
        assert_eq!(covered_cells(&rules[0]), covered_cells(&source));
        let again = reopened_rules(&exported_rules(json!(rules)));
        assert_eq!(again[0]["formulaAnchor"], json!({ "row": 1, "col": 0 }));
        assert_eq!(again[0]["formula1"], "=A1>5");
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
        assert_eq!(rules[0]["formula1"], source["formula1"]);
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
        assert_eq!(rules[0]["config"]["min"]["value"], "=A1");
        assert_eq!(rules[0]["config"]["isGradient"], false);
        assert_eq!(rules[0]["config"]["nativeColor"], "#334455");
        assert_eq!(rules[1]["config"]["thresholds"][0]["value"]["value"], "=A1");
        assert_eq!(rules[2]["config"]["icons"][0]["value"]["value"], "=A1");
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
        assert_eq!(rules.iter().map(|rule| rule["id"].as_str().unwrap()).collect::<Vec<_>>(), ["z-first", "a-second"]);
        assert_eq!(rules.iter().map(|rule| rule["priority"].as_u64().unwrap()).collect::<Vec<_>>(), [1, 1]);
    }

    #[test]
    fn native_repeated_roundtrips_retain_local_identity_sparse_priorities_and_unknown_json_without_source_growth() {
        let original = json!({ "id": "stable-local", "type": "expression", "ranges": [{ "startRow": 0, "startCol": 0, "endRow": 3, "endCol": 0 }],
            "formula1": "=A1>0", "priority": 17, "style": { "bold": false, "vendor": { "font": "retained" } },
            "config": { "vendor": { "ref": "A1" } }, "metadata": { "note": "retained" }, "vendor": { "root": [1, 2] } });
        let first_bytes = exported_rules(json!([original])); let first = reopened_rules(&first_bytes);
        assert_eq!(first[0]["id"], "stable-local"); assert_eq!(first[0]["priority"], 17); assert_eq!(first[0]["formula1"], "=A1>0");
        assert_eq!(first[0]["style"]["vendor"], original["style"]["vendor"]); assert_eq!(first[0]["vendor"], original["vendor"]);
        let second_bytes = exported_rules(json!(first)); let second = reopened_rules(&second_bytes);
        assert_eq!(second[0]["id"], "stable-local"); assert_eq!(second[0]["config"]["vendor"], original["config"]["vendor"]);
        assert_eq!(second[0]["officeSource"]["ruleXml"], first[0]["officeSource"]["ruleXml"]);
        assert!(second[0]["officeSource"]["ruleXml"].as_str().unwrap().find("urn:office-studio:conditional-formatting:1").is_none());
        assert!(second[0]["officeSource"]["baseline"].get("officeSource").is_none());
        assert_eq!(worksheet_xml(&second_bytes).matches("<oscf:model ").count(), 1);
    }

    #[test]
    fn standard_office_formula_edits_override_stale_private_conditions_but_keep_identity_and_vendor_json() {
        let original = json!({ "id": "stable-local", "type": "expression", "ranges": [{ "startRow": 0, "startCol": 0, "endRow": 3, "endCol": 0 }],
            "formula1": "=A1>0", "priority": 17, "style": { "vendor": { "font": "retained" } }, "vendor": { "note": "retained" } });
        let xml = worksheet_xml(&exported_rules(json!([original]))).replace("<formula>A1&gt;0</formula>", "<formula>A1&gt;9</formula>");
        let mut warnings = Vec::new(); let edited = parse_rules(&xml, &[], &mut warnings); assert!(warnings.is_empty(), "{warnings:?}");
        assert_eq!(edited[0]["id"], "stable-local"); assert_eq!(edited[0]["formula1"], "A1>9"); assert_eq!(edited[0]["priority"], 1);
        assert_eq!(edited[0]["style"]["vendor"], original["style"]["vendor"]); assert_eq!(edited[0]["vendor"], original["vendor"]);
        let reopened = reopened_rules(&exported_rules(json!(edited)));
        assert_eq!(reopened[0]["formula1"], "A1>9"); assert_eq!(reopened[0]["officeSource"]["baseline"]["formula1"], "A1>9");
    }

    pub(super) fn theme_fixture() -> String {
        let colors = [("dk1", "000000"), ("lt1", "FFFFFF"), ("dk2", "1F497D"), ("lt2", "EEECE1"), ("accent1", "FF0000"), ("accent2", "336699"),
            ("accent3", "9BBB59"), ("accent4", "8064A2"), ("accent5", "4BACC6"), ("accent6", "F79646"), ("hlink", "0000FF"), ("folHlink", "800080")];
        let colors = colors.into_iter().map(|(name, color)| format!(r#"<a:{name}><a:srgbClr val="{color}"/></a:{name}>"#)).collect::<String>();
        let fills = r#"<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>"#.repeat(3);
        let lines = r#"<a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>"#.repeat(3);
        let effects = "<a:effectStyle><a:effectLst/></a:effectStyle>".repeat(3);
        format!(r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Contract"><a:themeElements><a:clrScheme name="Contract">{colors}</a:clrScheme><a:fontScheme name="Contract"><a:majorFont><a:latin typeface="Aptos Display"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Aptos"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Contract"><a:fillStyleLst>{fills}</a:fillStyleLst><a:lnStyleLst>{lines}</a:lnStyleLst><a:effectStyleLst>{effects}</a:effectStyleLst><a:bgFillStyleLst>{fills}</a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>"#)
    }
    pub(super) fn part_xml(bytes: &[u8], part: &str) -> String {
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap(); let mut text = String::new();
        archive.by_name(part).unwrap().read_to_string(&mut text).unwrap(); text
    }
    pub(super) fn themed_rule() -> Value {
        let styles = parse_xml(&format!(r#"<styleSheet xmlns="{}" xmlns:v="urn:vendor"><dxfs><dxf><font><name val="Office Font"/><sz val="14"/><u val="double"/><color theme="4" tint="0.5"/><scheme val="minor"/><b val="0"/></font><numFmt numFmtId="166" formatCode="0.000"/><alignment wrapText="1"/><border><bottom style="double"><color indexed="10"/></bottom></border><protection locked="0"/><extLst><ext uri="vendor"><v:payload ref="A1"/></ext></extLst></dxf></dxfs></styleSheet>"#, source_xml::MAIN_NAMESPACE));
        let xml = format!(r#"<worksheet xmlns="{}"><conditionalFormatting sqref="A1:A4"><cfRule type="expression" dxfId="0" priority="3"><formula>A1&gt;0</formula></cfRule></conditionalFormatting></worksheet>"#, source_xml::MAIN_NAMESPACE);
        let mut warnings = Vec::new();
        let rules = parse_rules_with_theme(&xml, &styles.descendants("dxf"), Some(&theme_fixture()), "xl/worksheets/sheet1.xml", &mut warnings);
        assert!(warnings.is_empty(), "{warnings:?}"); rules[0].clone()
    }

    #[test]
    fn native_theme_and_rich_dxf_roundtrip_keeps_unedited_office_fields_after_a_bold_edit() {
        let source = themed_rule(); assert_eq!(source["style"]["fontColor"], "#FF8080"); assert_eq!(source["style"]["underline"], true);
        let bytes = exported_rules(json!([source])); let mut rules = reopened_rules(&bytes);
        assert_eq!(rules[0]["style"]["fontColor"], "#FF8080"); assert_eq!(rules[0]["priority"], 3);
        assert!(part_xml(&bytes, "xl/_rels/workbook.xml.rels").contains("relationships/theme"));
        rules[0]["style"]["bold"] = json!(true);
        let edited_bytes = exported_rules(json!(rules)); let styles = source_xml::parse(&part_xml(&edited_bytes, "xl/styles.xml")).unwrap();
        let dxf = &styles.child("dxfs").unwrap().children[0]; let font = dxf.child("font").unwrap();
        assert_eq!(font.child("u").unwrap().attrs.get("val").map(String::as_str), Some("double"));
        assert_eq!(font.child("color").unwrap().attrs.get("theme").map(String::as_str), Some("4"));
        assert_eq!(font.child("color").unwrap().attrs.get("tint").map(String::as_str), Some("0.5"));
        assert_eq!(font.child("name").unwrap().attrs.get("val").map(String::as_str), Some("Office Font"));
        for name in ["numFmt", "alignment", "border", "protection", "extLst"] { assert!(dxf.child(name).is_some(), "{name}"); }
        assert_eq!(dxf.descendants("payload")[0].attrs.get("ref").map(String::as_str), Some("A1"));
        assert_eq!(source_xml::parse(&part_xml(&edited_bytes, "xl/theme/theme1.xml")).unwrap().fingerprint(), source_xml::parse(&theme_fixture()).unwrap().fingerprint());
        let reopened = reopened_rules(&edited_bytes); assert_eq!(reopened[0]["style"]["bold"], true); assert_eq!(reopened[0]["style"]["underline"], true);
    }

    #[test]
    fn linked_data_bars_roundtrip_their_guid_and_remap_both_link_and_revision_id_for_a_copy() {
        const LINK: &str = "{11111111-2222-4333-8444-555555555555}";
        const REVISION: &str = "{AAAAAAAA-BBBB-4CCC-8DDD-EEEEEEEEEEEE}";
        let xml = format!(r#"<worksheet xmlns="{}" xmlns:x14="{X14_NAMESPACE}" xmlns:xm="{XM_NAMESPACE}" xmlns:xr="{REVISION_NAMESPACE}" xmlns:mc="{MC_NAMESPACE}" mc:Ignorable="xr"><conditionalFormatting sqref="A1:A4"><cfRule type="dataBar" priority="5" xr:uid="{REVISION}"><dataBar showValue="0"><cfvo type="num" val="0"/><cfvo type="num" val="100"/><color rgb="FF336699"/></dataBar><extLst><ext uri="{RULE_LINK_EXTENSION_URI}"><x14:id>{LINK}</x14:id></ext></extLst></cfRule></conditionalFormatting><extLst><ext uri="{}"><x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type="dataBar" id="{LINK}"><x14:dataBar gradient="0" showValue="0" border="1" axisPosition="middle" direction="rightToLeft" negativeBarColorSameAsPositive="0"><x14:cfvo type="num"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="num"><xm:f>100</xm:f></x14:cfvo><x14:fillColor rgb="FF336699"/><x14:borderColor rgb="FF112233"/><x14:negativeFillColor rgb="FFAA3344"/><x14:axisColor rgb="FF000000"/></x14:dataBar></x14:cfRule><xm:sqref>A1:A4</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst></worksheet>"#, source_xml::MAIN_NAMESPACE, source_xml::WORKSHEET_EXTENSION_URI);
        let mut warnings = Vec::new(); let mut rules = parse_rules(&xml, &[], &mut warnings); assert!(warnings.is_empty(), "{warnings:?}");
        assert_eq!(rules[0]["id"], REVISION); assert_eq!(rules[0]["config"]["direction"], "rightToLeft");
        rules[0]["config"]["nativeColor"] = json!("#443388");
        let mut copied = rules[0].clone(); copied["id"] = json!("copied-bar"); copied["ranges"] = json!([{ "startRow": 0, "startCol": 2, "endRow": 3, "endCol": 2 }]); rules.push(copied);
        let bytes = exported_rules(json!(rules)); let standard = source_xml::parse(&worksheet_xml(&bytes)).unwrap();
        let base = standard.children.iter().filter(|node| node.is_main("conditionalFormatting")).map(|group| group.child("cfRule").unwrap()).collect::<Vec<_>>();
        assert_eq!(base.len(), 2); assert_eq!(linked_id(base[0]).as_deref(), Some(LINK)); assert_ne!(linked_id(base[1]).as_deref(), Some(LINK));
        assert_eq!(base[0].attribute_ns(REVISION_NAMESPACE, "uid"), Some(REVISION)); assert_ne!(base[1].attribute_ns(REVISION_NAMESPACE, "uid"), Some(REVISION));
        let x14 = standard.descendants("cfRule").into_iter().filter(|rule| rule.is(X14_NAMESPACE, "cfRule")).collect::<Vec<_>>();
        assert_eq!(x14.len(), 2); assert!(x14.iter().all(|rule| rule.attrs.get("priority").is_none()));
        for (base, x14) in base.iter().zip(&x14) { assert_eq!(linked_id(base).as_ref(), x14.attrs.get("id")); }
        let reopened = reopened_rules(&bytes); assert_eq!(reopened[0]["id"], REVISION); assert_eq!(reopened[1]["id"], "copied-bar");
        assert_eq!(reopened[0]["config"]["nativeColor"], "#443388"); assert_eq!(reopened[1]["config"]["axisPosition"], "middle");
    }

    #[test]
    fn standalone_x14_custom_icons_and_inline_dxf_reopen_with_noicons_and_strict_thresholds() {
        const ID: &str = "{11111111-2222-4333-8444-555555555555}";
        let xml = format!(r#"<worksheet xmlns="{}" xmlns:x14="{X14_NAMESPACE}" xmlns:xm="{XM_NAMESPACE}"><extLst><ext uri="{}"><x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type="iconSet" priority="3" id="{ID}"><x14:iconSet iconSet="3Arrows" custom="1" showValue="0"><x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="num"><xm:f>40</xm:f></x14:cfvo><x14:cfvo type="num" gte="0"><xm:f>70</xm:f></x14:cfvo><x14:cfIcon iconSet="3Arrows" iconId="0"/><x14:cfIcon iconSet="NoIcons" iconId="0"/><x14:cfIcon iconSet="3Stars" iconId="2"/></x14:iconSet><x14:dxf><font><u val="none"/></font></x14:dxf></x14:cfRule><xm:sqref>A1:A4</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst></worksheet>"#, source_xml::MAIN_NAMESPACE, source_xml::WORKSHEET_EXTENSION_URI);
        let mut warnings = Vec::new(); let rules = parse_rules(&xml, &[], &mut warnings); assert!(warnings.is_empty(), "{warnings:?}");
        assert_eq!(rules[0]["type"], "iconSet"); assert_eq!(rules[0]["officeSource"]["storage"], "x14");
        assert_eq!(rules[0]["config"]["icons"][1]["iconType"], "EMPTY_ICON_TYPE"); assert_eq!(rules[0]["style"]["underline"], false);
        let bytes = exported_rules(json!(rules)); let standard = source_xml::parse(&worksheet_xml(&bytes)).unwrap();
        assert_eq!(standard.children.iter().filter(|node| node.is_main("conditionalFormatting")).count(), 0);
        let extended = standard.descendants("cfRule").into_iter().find(|rule| rule.is(X14_NAMESPACE, "cfRule")).unwrap();
        assert_eq!(extended.attrs.get("priority").map(String::as_str), Some("1")); assert_eq!(extended.attrs.get("id").map(String::as_str), Some(ID));
        assert_eq!(extended.descendants("cfIcon").iter().filter(|node| node.attrs.get("iconSet").map(String::as_str) == Some("NoIcons")).count(), 1);
        assert_eq!(extended.descendants("cfvo")[2].attrs.get("gte").map(String::as_str), Some("0"));
        let reopened = reopened_rules(&bytes); assert_eq!(reopened[0]["config"]["icons"], rules[0]["config"]["icons"]); assert_eq!(reopened[0]["priority"], 3);
    }

    #[test]
    fn unknown_icon_sets_stddev_and_mismatched_linked_types_are_retained_as_office_instead_of_evaluated() {
        for body in [r#"<cfRule type="iconSet" priority="1"><iconSet iconSet="vendorIcons"><cfvo type="percent" val="0"/><cfvo type="percent" val="40"/><cfvo type="percent" val="70"/></iconSet></cfRule>"#,
            r#"<cfRule type="aboveAverage" priority="1" stdDev="2"/>"#] {
            let xml = format!(r#"<worksheet xmlns="{}"><conditionalFormatting sqref="A1:A4">{body}</conditionalFormatting></worksheet>"#, source_xml::MAIN_NAMESPACE);
            let mut warnings = Vec::new(); let rules = parse_rules(&xml, &[], &mut warnings);
            assert_eq!(rules.len(), 1); assert_eq!(rules[0]["type"], "office"); assert!(!warnings.is_empty()); assert!(rules[0]["officeSource"]["ruleXml"].is_string());
            let reopened = reopened_rules(&exported_rules(json!(rules))); assert_eq!(reopened[0]["type"], "office");
        }
        let mismatched = format!(r#"<worksheet xmlns="{}" xmlns:x14="{X14_NAMESPACE}" xmlns:xm="{XM_NAMESPACE}"><conditionalFormatting sqref="A1:A4"><cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF336699"/></dataBar><extLst><ext uri="{RULE_LINK_EXTENSION_URI}"><x14:id>{{11111111-2222-4333-8444-555555555555}}</x14:id></ext></extLst></cfRule></conditionalFormatting><extLst><ext uri="{}"><x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type="iconSet" id="{{11111111-2222-4333-8444-555555555555}}"><x14:iconSet/></x14:cfRule><xm:sqref>A1:A4</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst></worksheet>"#, source_xml::MAIN_NAMESPACE, source_xml::WORKSHEET_EXTENSION_URI);
        let mut warnings = Vec::new(); let rules = parse_rules(&mismatched, &[], &mut warnings); assert_eq!(rules[0]["type"], "office"); assert!(!warnings.is_empty());
    }

    #[test]
    fn conflicting_preserved_themes_reject_export_instead_of_rebinding_colors_to_another_workbook_theme() {
        let first = themed_rule(); let mut second = first.clone(); second["id"] = json!("other-source");
        second["officeSource"]["themeXml"] = json!(theme_fixture().replace("FF0000", "0000FF"));
        let workbook = json!({ "kind": "workbook", "sheets": [{ "id": "data", "name": "Data", "cells": {}, "conditionalFormats": [first, second] }] });
        let error = office_writer::export_xlsx_bytes(&workbook.to_string()).unwrap_err(); assert_eq!(error.code, "INVALID_MODEL");
    }

    #[test]
    fn additional_standard_formulas_remain_opaque_and_survive_native_range_edits() {
        let xml = format!(r#"<worksheet xmlns="{}"><conditionalFormatting sqref="A1:A4"><cfRule type="expression" priority="1"><formula>A1&gt;0</formula><formula>B1&gt;0</formula><formula>C1&gt;0</formula></cfRule></conditionalFormatting></worksheet>"#, source_xml::MAIN_NAMESPACE);
        let mut warnings = Vec::new(); let mut rules = parse_rules(&xml, &[], &mut warnings);
        assert_eq!(rules[0]["type"], "office"); assert!(!warnings.is_empty());
        let mut known = rules[0].clone(); known["type"] = json!("expression"); known["subType"] = json!("formula");
        let baseline = known_snapshot(&known); known["officeSource"]["baseline"] = baseline;
        let workbook = json!({ "kind": "workbook", "sheets": [{ "id": "data", "name": "Data", "cells": {}, "conditionalFormats": [known] }] });
        assert_eq!(office_writer::export_xlsx_bytes(&workbook.to_string()).unwrap_err().code, "INVALID_MODEL");
        rules[0]["ranges"] = json!([{ "startRow": 2, "startCol": 0, "endRow": 5, "endCol": 0 }]);
        let bytes = exported_rules(json!(rules)); let standard = source_xml::parse(&worksheet_xml(&bytes)).unwrap();
        let formulas = standard.descendants("formula").into_iter().map(|node| node.text).collect::<Vec<_>>();
        assert_eq!(formulas, vec!["A3>0", "B3>0", "C3>0"]);
        assert_eq!(reopened_rules(&bytes)[0]["type"], "office");
    }
}
