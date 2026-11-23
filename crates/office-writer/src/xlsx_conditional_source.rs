//! Standard SpreadsheetML rules with lossless, field-by-field Office provenance.
use super::{ConditionalFormatInput, ConditionalFormatStyleInput, RangeInput, SheetInput, StyleRegistry, WorkbookInput,
    conditional_formula, normalize_color, range_bounds, range_ref, write_conditional_rule};
use crate::{ExportError, ExportWarning};
use office_io_common::conditional_format_color::{uses_theme, validate_theme_references};
use office_io_common::conditional_format_source::{ConditionalFormatModelBudget, OfficeSource, ProjectionHashes, SOURCE_RESOURCE_BYTE_LIMIT, SOURCE_RESOURCE_ELEMENT_LIMIT, STYLE_FIELDS, attach_projection,
    icon_index, known_snapshot, office_icon_type, strip_projection, without_source};
use office_io_common::conditional_format_xml::{self as source_xml, Element, MAIN_NAMESPACE, RULE_LINK_EXTENSION_URI,
    STRICT_MAIN_NAMESPACE, WORKSHEET_EXTENSION_URI, X14_NAMESPACE, XM_NAMESPACE, XmlSelectionLimits};
use serde_json::{Value, json};
use std::collections::HashSet;

const REVISION_NAMESPACE: &str = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision";
const MC_NAMESPACE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

#[derive(Default)]
pub(super) struct ConditionalOutput { pub(super) base: String, pub(super) extension: String }

#[derive(Default)]
struct SourceParts { rule: Option<Element>, group: Option<Element>, dxf: Option<Element>, x14: Option<Element>, theme: Option<Element> }

fn invalid(message: impl Into<String>) -> ExportError { ExportError::invalid(message) }
fn parsed(xml: &str) -> Result<Element, ExportError> { source_xml::parse(xml).map_err(invalid) }
fn spreadsheet(node: &mut Element) {
    if node.namespace == STRICT_MAIN_NAMESPACE || node.namespace.is_empty() { node.namespace = MAIN_NAMESPACE.into(); }
    for value in node.attrs.values_mut() { if value == STRICT_MAIN_NAMESPACE { *value = MAIN_NAMESPACE.into(); } }
    // Extension payloads can use a deliberate empty/default namespace. Their
    // resolved names remain authored even when the surrounding Office dialect changes.
    if node.is_main("ext") { return }
    for child in &mut node.children { spreadsheet(child); }
}
fn rename(node: &mut Element, namespace: &str, prefix: &str, name: &str) {
    node.namespace = namespace.into(); node.name = name.into(); node.qualified_name = if prefix.is_empty() { name.into() } else { format!("{prefix}:{name}") };
}
fn dxf_root(node: &mut Element, x14: bool) {
    spreadsheet(node); rename(node, if x14 { X14_NAMESPACE } else { MAIN_NAMESPACE }, if x14 { "x14" } else { "" }, "dxf");
}
fn is_rule(node: &Element) -> bool { node.is_main("cfRule") || node.is(X14_NAMESPACE, "cfRule") }
fn is_visual(node: &Element, name: &str) -> bool { node.is_main(name) || node.is(X14_NAMESPACE, name) }
fn ext_list(node: &Element) -> bool { node.name == "extLst" && (node.is_main("extLst") || node.namespace == X14_NAMESPACE) }
fn group_reference(group: &Element) -> Option<&str> {
    if group.namespace == X14_NAMESPACE { group.children.iter().find(|node| node.is(XM_NAMESPACE, "sqref")).map(|node| node.text.as_str()) }
    else { group.attrs.get("sqref").map(String::as_str) }
}
fn read_source(source: Option<&OfficeSource>) -> Result<SourceParts, ExportError> {
    let Some(source) = source else { return Ok(SourceParts::default()) };
    source.validate().map_err(invalid)?;
    let mut rule = parsed(&source.rule_xml)?; strip_projection(&mut rule);
    if !is_rule(&rule) || (source.storage == "x14") != (rule.namespace == X14_NAMESPACE) { return Err(invalid("conditional-format source rule namespace does not match its storage")); }
    if rule.attrs.get("dxfId").is_some_and(|value| value.parse::<u32>().is_err()) { return Err(invalid("CF_SOURCE_XML_INVALID: conditional-format source dxfId is not a valid Office unsigned index")) }
    if source.baseline["type"] != "office" && rule.children.iter().filter(|node| formula_child(node)).count() > 2 {
        return Err(invalid("conditional-format source has additional formulas outside the known editable model"));
    }
    let mut group = source.group_xml.as_deref().map(parsed).transpose()?;
    if let Some(group) = &mut group {
        if !is_visual(group, "conditionalFormatting") || (source.storage == "x14") != (group.namespace == X14_NAMESPACE) { return Err(invalid("invalid conditional-format source group")); }
        group.replace_children(is_rule, Vec::new()); strip_projection(group);
    }
    let dxf = source.differential_format_xml.as_deref().map(parsed).transpose()?
        .or_else(|| rule.children.iter().find(|node| is_visual(node, "dxf")).cloned());
    if dxf.as_ref().is_some_and(|node| !is_visual(node, "dxf")) { return Err(invalid("invalid conditional-format differential-style source")); }
    // An edited replacement cannot recover the fields lost with an unresolved
    // original dxfId. Reject the source before differential-style merging.
    if rule.attrs.contains_key("dxfId") && dxf.is_none() { return Err(invalid("conditional-format source differential style is unavailable")) }
    let mut x14 = source.x14_xml.as_deref().map(parsed).transpose()?;
    if let Some(group) = &mut x14 {
        if !group.is(X14_NAMESPACE, "conditionalFormatting") || group.children.iter().filter(|node| node.is(X14_NAMESPACE, "cfRule")).count() != 1 { return Err(invalid("invalid linked x14 conditional-format source")); }
        strip_projection(group);
    }
    let theme = source.theme_xml.as_deref().map(parsed).transpose()?;
    if theme.as_ref().is_some_and(|theme| theme.name != "theme" || !matches!(theme.namespace.as_str(),
        "http://schemas.openxmlformats.org/drawingml/2006/main" | "http://purl.oclc.org/ooxml/drawingml/main")) { return Err(invalid("invalid conditional-format source theme")); }
    for node in std::iter::once(&rule).chain(group.as_ref()).chain(dxf.as_ref()).chain(x14.as_ref()).chain(theme.as_ref()) {
        node.validate_fragment().map_err(invalid)?;
    }
    for node in std::iter::once(&rule).chain(dxf.as_ref()).chain(x14.as_ref()) {
        validate_theme_references(node, theme.as_ref()).map_err(invalid)?;
    }
    Ok(SourceParts { rule: Some(rule), group, dxf, x14, theme })
}

struct Diff { current: Value, baseline: Option<Value> }
impl Diff {
    fn new(model: &Value, source: Option<&OfficeSource>) -> Self { Self { current: known_snapshot(model), baseline: source.map(|source| known_snapshot(&source.baseline)) } }
    fn field(&self, key: &str) -> bool { self.baseline.as_ref().is_none_or(|baseline| baseline[key] != self.current[key]) }
    fn config(&self, key: &str) -> bool { self.baseline.as_ref().is_none_or(|baseline| baseline["config"][key] != self.current["config"][key]) }
    fn rebase(&self) -> bool { self.field("ranges") || self.field("formulaAnchor") }
    fn semantic_change(&self) -> bool { ["type", "subType", "operator", "formula1", "formula2", "value", "isBottom", "isPercent", "config"].iter().any(|key| self.field(key)) }
}

fn snapshot(rule: &ConditionalFormatInput, sheet: &SheetInput, ordinal: usize) -> Result<Value, ExportError> {
    let mut model = without_source(&serde_json::to_value(rule).map_err(|error| invalid(error.to_string()))?);
    for key in ["formulaAnchor", "formula2", "subType", "value", "config"] { if model[key].is_null() { model.as_object_mut().unwrap().remove(key); } }
    for key in ["isBottom", "isPercent"] { if model[key] == Value::Bool(false) { model.as_object_mut().unwrap().remove(key); } }
    for key in STYLE_FIELDS { if model["style"][*key].is_null() { model["style"].as_object_mut().unwrap().remove(*key); } }
    model["ranges"] = Value::Array(rule.ranges.iter().map(|range| {
        let (sr, sc, er, ec) = range_bounds(range).ok_or_else(|| invalid("conditional-format range has missing coordinates"))?;
        if sr > er || sc > ec || er > 1_048_575 || ec > 16_383 { return Err(invalid("conditional-format range exceeds the worksheet")) }
        Ok(json!({ "startRow": sr, "startCol": sc, "endRow": er, "endCol": ec }))
    }).collect::<Result<Vec<_>, ExportError>>()?);
    if rule.ranges.is_empty() { return Err(invalid("conditional format has no ranges")) }
    model["priority"] = json!(rule.priority.max(1));
    if rule.id.trim().is_empty() { model["id"] = json!(format!("conditional-{}-{ordinal}", sheet.id)); }
    Ok(model)
}

#[derive(Default)]
struct Guids { reserved: HashSet<String>, used: HashSet<String> }
fn valid_guid(value: &str) -> bool {
    let bytes = value.as_bytes(); bytes.len() == 38 && bytes[0] == b'{' && bytes[37] == b'}'
        && bytes[1..37].iter().enumerate().all(|(index, byte)| if [8, 13, 18, 23].contains(&index) { *byte == b'-' } else { byte.is_ascii_hexdigit() })
}
fn link_id(rule: &Element) -> Option<String> {
    rule.children.iter().find(|node| node.is_main("extLst"))?.children.iter().find(|node| node.is_main("ext")
        && node.attrs.get("uri").is_some_and(|uri| uri.eq_ignore_ascii_case(RULE_LINK_EXTENSION_URI)))?.children.iter()
        .find(|node| node.is(X14_NAMESPACE, "id")).map(|node| node.text.trim().to_owned())
}
impl Guids {
    fn reserve(&mut self, parts: &SourceParts) {
        if let Some(rule) = &parts.rule {
            for id in [rule.attribute_ns(REVISION_NAMESPACE, "uid").map(str::to_owned), rule.attrs.get("id").cloned(), link_id(rule)].into_iter().flatten() {
                if valid_guid(&id) { self.reserved.insert(id.to_ascii_uppercase()); }
            }
        }
    }
    fn assign(&mut self, original: Option<&str>, copied: bool, key: &str) -> Result<String, ExportError> {
        if let Some(original) = original {
            if !valid_guid(original) { return Err(invalid("conditional-format Office GUID is malformed")) }
            let normalized = original.to_ascii_uppercase();
            if !copied && self.used.insert(normalized.clone()) { return Ok(normalized) }
        }
        for ordinal in 0u64.. {
            let mut identity = Element::new("urn:office-studio:conditional-formatting:identity", "guid"); identity.set_attr("key", key); identity.set_attr("ordinal", ordinal.to_string());
            let first = identity.fingerprint(); identity.set_attr("salt", "second-half");
            let mut digits = format!("{first}{}", identity.fingerprint()).to_ascii_uppercase().into_bytes(); digits[12] = b'4'; digits[16] = b'8';
            let digits = String::from_utf8(digits).unwrap(); let id = format!("{{{}-{}-{}-{}-{}}}", &digits[0..8], &digits[8..12], &digits[12..16], &digits[16..20], &digits[20..32]);
            if !self.reserved.contains(&id) && self.used.insert(id.clone()) { return Ok(id) }
        }
        unreachable!()
    }
}

fn upsert(parent: &mut Element, child: Element) {
    let namespace = child.namespace.clone(); let name = child.name.clone();
    let select = |node: &Element| node.namespace == namespace && node.name == name;
    if parent.children.iter().any(&select) { parent.replace_children(select, vec![child]); }
    else { parent.insert_before(ext_list, child); }
}
fn copy_attr(target: &mut Element, generated: &Element, key: &str) {
    if let Some(value) = generated.attrs.get(key) { target.set_attr(key, value); } else { target.remove_attr(key); }
}
fn set_link(rule: &mut Element, id: &str) {
    let mut list = rule.children.iter().find(|node| node.is_main("extLst")).cloned().unwrap_or_else(|| Element::new(MAIN_NAMESPACE, "extLst"));
    let mut ext = list.children.iter().find(|node| node.is_main("ext") && node.attrs.get("uri").is_some_and(|uri| uri.eq_ignore_ascii_case(RULE_LINK_EXTENSION_URI)))
        .cloned().unwrap_or_else(|| { let mut ext = Element::new(MAIN_NAMESPACE, "ext"); ext.set_attr("uri", RULE_LINK_EXTENSION_URI); ext });
    let mut identity = ext.children.iter().find(|node| node.is(X14_NAMESPACE, "id")).cloned().unwrap_or_else(|| Element::prefixed(X14_NAMESPACE, "x14", "id"));
    identity.set_text(id); upsert(&mut ext, identity);
    if list.children.iter().any(|node| node.is_main("ext") && node.attrs.get("uri").is_some_and(|uri| uri.eq_ignore_ascii_case(RULE_LINK_EXTENSION_URI))) {
        list.replace_children(|node| node.is_main("ext") && node.attrs.get("uri").is_some_and(|uri| uri.eq_ignore_ascii_case(RULE_LINK_EXTENSION_URI)), vec![ext]);
    } else { list.push(ext); }
    upsert(rule, list);
}
fn remap_revision_id(rule: &mut Element, guids: &mut Guids, copied: bool, key: &str) -> Result<(), ExportError> {
    let attr = rule.attribute_namespaces.iter().find(|(name, namespace)| namespace.as_str() == REVISION_NAMESPACE && name.rsplit(':').next() == Some("uid"))
        .map(|(name, _)| name.clone());
    if let Some(attr) = attr {
        let value = guids.assign(rule.attrs.get(&attr).map(String::as_str), copied, &format!("{key}:revision"))?;
        rule.set_attr_ns(REVISION_NAMESPACE, &attr, value);
        let prefix = attr.split_once(':').map(|(prefix, _)| prefix).unwrap_or("xr");
        let ignorable = rule.attribute_namespaces.iter().find(|(name, namespace)| namespace.as_str() == MC_NAMESPACE && name.rsplit(':').next() == Some("Ignorable"))
            .map(|(name, _)| name.clone()).unwrap_or_else(|| "mc:Ignorable".into());
        let mut prefixes = rule.attrs.get(&ignorable).map(|value| value.split_whitespace().map(str::to_owned).collect::<Vec<_>>()).unwrap_or_default();
        if !prefixes.iter().any(|value| value == prefix) { prefixes.push(prefix.into()); }
        rule.set_attr_ns(MC_NAMESPACE, &ignorable, prefixes.join(" "));
    }
    Ok(())
}

pub(super) fn write(sheet: &SheetInput, registry: &mut StyleRegistry, warnings: &mut Vec<ExportWarning>) -> Result<ConditionalOutput, ExportError> {
    let mut ordered = sheet.conditional_formats.iter().collect::<Vec<_>>(); ordered.sort_by_key(|rule| rule.priority);
    let mut prepared = Vec::new(); let mut guids = Guids::default();
    for rule in ordered { let parts = read_source(rule.office_source.as_ref())?; guids.reserve(&parts); prepared.push((rule, parts)); }
    let mut output = ConditionalOutput::default(); let mut extended = Element::prefixed(X14_NAMESPACE, "x14", "conditionalFormattings"); let mut ids = HashSet::new();
    for (index, (rule, mut parts)) in prepared.into_iter().enumerate() {
        let model = snapshot(rule, sheet, index + 1)?; let id = model["id"].as_str().unwrap();
        if !ids.insert(id.to_owned()) { return Err(invalid("conditional-format local IDs must be unique within a worksheet")) }
        let source = rule.office_source.as_ref(); let copied = source.is_some_and(|source| source.original_id != id);
        let diff = Diff::new(&model, source); let priority = u32::try_from(index + 1).map_err(|_| invalid("too many conditional-format rules"))?;
        if priority > i32::MAX as u32 { return Err(invalid("conditional-format priorities exceed the Office limit")) }
        let key = format!("{}:{id}", sheet.id);
        let mut dxf = differential_format(&rule.style, parts.dxf.take(), &diff)?;
        let standalone = source.is_some_and(|source| source.storage == "x14") || rule.rule_type == "iconSet" && icon_requires_x14(rule);
        let dxf_id = if standalone { None } else { dxf.as_ref().map(|dxf| registry.register_conditional_style_xml(dxf)) };
        let mut projected = conditional_formula::project_rule(rule)?;
        let mut authored = if rule.rule_type == "office" {
            if source.is_none() || diff.semantic_change() { return Err(invalid("unknown Office conditional-format semantics must be retained")) }
            let original = parts.rule.take().ok_or_else(|| invalid("unknown Office conditional format has no XML source"))?;
            let (projected_rule, original, linked) = project_opaque(rule, original, parts.x14.take(), parts.group.as_ref(), &diff)?;
            projected = projected_rule; parts.x14 = linked; original
        } else {
            let mut generated = generated_rule(&projected, dxf_id, priority, &diff)?;
            if standalone { to_x14(&mut generated); }
            if matches!(rule.rule_type.as_str(), "dataBar" | "colorScale" | "iconSet") {
                let visual = generated_visual(&projected, standalone, &diff, warnings)?; upsert(&mut generated, visual);
            }
            merge_rule(parts.rule.take(), generated, &diff, &projected, warnings)?
        };
        if standalone {
            if !authored.is(X14_NAMESPACE, "cfRule") { to_x14(&mut authored); }
            let original = authored.attrs.get("id").cloned(); let guid = guids.assign(original.as_deref(), copied, &format!("{key}:x14"))?;
            authored.set_attr("id", guid); authored.remove_attr("dxfId");
            if rule.rule_type == "office" && source.is_some_and(|source| source.storage == "x14") && !authored.attrs.contains_key("priority") { authored.remove_attr("priority"); }
            else { authored.set_attr("priority", priority.to_string()); }
            if diff.field("stopIfTrue") { if rule.stop_if_true { authored.set_attr("stopIfTrue", "1"); } else { authored.remove_attr("stopIfTrue"); } }
            remap_revision_id(&mut authored, &mut guids, copied, &key)?;
            if let Some(dxf) = &mut dxf { dxf_root(dxf, true); upsert(&mut authored, dxf.clone()); }
            else { authored.replace_children(|node| is_visual(node, "dxf"), Vec::new()); }
            let mut group = make_group(parts.group, true, &projected.ranges)?;
            let theme = checked_theme(&authored, dxf.as_ref(), None, parts.theme.as_ref())?;
            let hashes = ProjectionHashes::of(&authored, &group, dxf.as_ref(), None, theme);
            attach_projection(&mut authored, &model, hashes).map_err(invalid)?; group.insert_before(|node| node.is(XM_NAMESPACE, "sqref") || ext_list(node), authored);
            extended.push(group);
        } else {
            spreadsheet(&mut authored); authored.set_attr("priority", priority.to_string());
            if let Some(index) = dxf_id { authored.set_attr("dxfId", index.to_string()); } else { authored.remove_attr("dxfId"); }
            if diff.field("stopIfTrue") { if rule.stop_if_true { authored.set_attr("stopIfTrue", "1"); } else { authored.remove_attr("stopIfTrue"); } }
            remap_revision_id(&mut authored, &mut guids, copied, &key)?;
            let linked = if rule.rule_type == "dataBar" {
                let original = link_id(&authored); let guid = guids.assign(original.as_deref(), copied, &format!("{key}:x14"))?;
                set_link(&mut authored, &guid);
                // These legacy fields belong to x14, never the base dataBar schema.
                if let Some(bar) = authored.child_mut("dataBar") { bar.remove_attr("gradient"); bar.replace_children(|node| node.is_main("negativeFillColor"), Vec::new()); }
                Some(linked_data_bar(parts.x14.take(), &projected, &diff, &guid, warnings)?)
            } else if rule.rule_type == "office" && let Some(mut linked) = parts.x14.take() {
                if copied {
                    let original = link_id(&authored); let guid = guids.assign(original.as_deref(), true, &format!("{key}:x14"))?; set_link(&mut authored, &guid);
                    if let Some(rule) = linked.children.iter_mut().find(|node| node.is(X14_NAMESPACE, "cfRule")) { rule.set_attr("id", guid); }
                }
                Some(linked)
            } else { None };
            let mut group = make_group(parts.group, false, &projected.ranges)?;
            let theme = checked_theme(&authored, dxf.as_ref(), linked.as_ref(), parts.theme.as_ref())?;
            let hashes = ProjectionHashes::of(&authored, &group, dxf.as_ref(), linked.as_ref(), theme);
            attach_projection(&mut authored, &model, hashes).map_err(invalid)?; group.insert_before(ext_list, authored); output.base.push_str(&group.render());
            if let Some(linked) = linked { extended.push(linked); }
        }
    }
    if !extended.children.is_empty() { let mut ext = Element::new(MAIN_NAMESPACE, "ext"); ext.set_attr("uri", WORKSHEET_EXTENSION_URI); ext.push(extended); output.extension = ext.render(); }
    Ok(output)
}

fn generated_rule(rule: &ConditionalFormatInput, dxf_id: Option<usize>, priority: u32, diff: &Diff) -> Result<Element, ExportError> {
    let mut xml = String::new(); write_conditional_rule(&mut xml, rule, dxf_id, priority, &rule.operator)?;
    let root = parsed(&format!("<worksheet xmlns=\"{MAIN_NAMESPACE}\">{xml}</worksheet>"))?;
    let mut authored = root.children.into_iter().find(|node| node.is_main("cfRule")).ok_or_else(|| invalid("conditional-format writer emitted no rule"))?;
    let text = matches!(rule.rule_type.as_str(), "containsText" | "notContainsText" | "beginsWith" | "endsWith") || rule.sub_type.as_deref() == Some("text");
    let literal = rule.value.as_ref().and_then(super::json_scalar_string).unwrap_or_else(|| rule.formula1.clone());
    let blanks_or_errors = matches!(rule.operator.as_str(), "containsBlanks" | "notContainsBlanks" | "containsErrors" | "notContainsErrors");
    if text && !rule.formula1.trim().is_empty() && (rule.formula1 != literal || blanks_or_errors)
        && (diff.baseline.is_none() || !diff.field("value") && !diff.field("operator")) {
        let mut formula = Element::new(MAIN_NAMESPACE, "formula"); formula.set_text(rule.formula1.trim().trim_start_matches('='));
        authored.replace_children(formula_child, vec![formula]);
    }
    Ok(authored)
}
fn make_group(source: Option<Element>, x14: bool, ranges: &[RangeInput]) -> Result<Element, ExportError> {
    let mut group = source.unwrap_or_else(|| if x14 { Element::prefixed(X14_NAMESPACE, "x14", "conditionalFormatting") } else { Element::new(MAIN_NAMESPACE, "conditionalFormatting") });
    if x14 { rename(&mut group, X14_NAMESPACE, "x14", "conditionalFormatting"); group.set_attr("xmlns:xm", XM_NAMESPACE); }
    else { spreadsheet(&mut group); }
    group.replace_children(is_rule, Vec::new());
    set_group_reference(&mut group, x14, ranges)?;
    Ok(group)
}
fn set_group_reference(group: &mut Element, x14: bool, ranges: &[RangeInput]) -> Result<(), ExportError> {
    let references = ranges.iter().map(|range| range_ref(range).ok_or_else(|| invalid("invalid conditional-format range"))).collect::<Result<Vec<_>, _>>()?.join(" ");
    if x14 { group.remove_attr("sqref"); let mut sqref = group.children.iter().find(|node| node.is(XM_NAMESPACE, "sqref")).cloned().unwrap_or_else(|| Element::prefixed(XM_NAMESPACE, "xm", "sqref")); sqref.set_text(references); upsert(group, sqref); }
    else { group.set_attr("sqref", references); }
    Ok(())
}
fn checked_theme<'a>(rule: &Element, dxf: Option<&Element>, x14: Option<&Element>, theme: Option<&'a Element>) -> Result<Option<&'a Element>, ExportError> {
    if uses_theme(rule) || dxf.is_some_and(uses_theme) || x14.is_some_and(uses_theme) { theme.map(Some).ok_or_else(|| invalid("conditional-format theme references have no preserved source theme")) }
    else { Ok(None) }
}

/// Inspect the final worksheet package XML with the importer's fragment limits.
/// The caller accumulates this tuple across worksheets before committing ZIP bytes.
pub(super) fn output_source_usage(xml: &str) -> Result<(usize, usize), ExportError> {
    let selected = source_xml::extract(std::io::Cursor::new(xml.as_bytes()), XmlSelectionLimits::default(), |path, _| {
        if path.len() < 2 || !path[0].is_main("worksheet") { return Ok(false) }
        Ok(path.len() == 2 && path[1].is_main("conditionalFormatting")
            || path.len() >= 4 && path[1].is_main("extLst") && path.last().is_some_and(|name| name.is(X14_NAMESPACE, "conditionalFormatting")))
    }).map_err(invalid)?;
    if !selected.root.is_main("worksheet") { return Err(invalid("CF_SOURCE_XML_INVALID: conditional-format output has an invalid worksheet root")) }
    Ok((selected.selected_bytes, selected.selected_elements))
}

/// Worksheet fragments plus each emitted differential style and selected theme.
pub(super) fn validate_output_resource_budget(bytes: usize, elements: usize, registry: &StyleRegistry, theme: Option<&str>) -> Result<(), ExportError> {
    validate_output_resource_budget_with_style_limits(bytes, elements, registry, theme, XmlSelectionLimits::default())
}

fn validate_output_resource_budget_with_style_limits(bytes: usize, elements: usize, registry: &StyleRegistry, theme: Option<&str>,
    style_limits: XmlSelectionLimits) -> Result<(), ExportError> {
    let mut bytes = bytes; let mut elements = elements;
    if bytes > SOURCE_RESOURCE_BYTE_LIMIT || elements > SOURCE_RESOURCE_ELEMENT_LIMIT {
        return Err(invalid("CF_SOURCE_TOTAL_LIMIT: conditional-format worksheet output resources exceed the workbook budget"))
    }
    let mut include = |node: &Element| -> Result<(usize, usize), ExportError> {
        node.validate_fragment().map_err(invalid)?;
        fn count(node: &Element) -> usize { 1usize.saturating_add(node.children.iter().map(count).fold(0usize, usize::saturating_add)) }
        let node_bytes = node.rendered_size(); let node_elements = count(node);
        bytes = bytes.checked_add(node_bytes).ok_or_else(|| invalid("CF_SOURCE_TOTAL_LIMIT: conditional-format output resource byte counter overflow"))?;
        elements = elements.checked_add(node_elements).ok_or_else(|| invalid("CF_SOURCE_TOTAL_LIMIT: conditional-format output resource element counter overflow"))?;
        if bytes > SOURCE_RESOURCE_BYTE_LIMIT || elements > SOURCE_RESOURCE_ELEMENT_LIMIT {
            return Err(invalid("CF_SOURCE_TOTAL_LIMIT: conditional-format output resources exceed the workbook budget"))
        }
        Ok((node_bytes, node_elements))
    };
    let mut style_bytes = 0usize; let mut style_elements = 0usize;
    for dxf in &registry.conditional_styles {
        let (node_bytes, node_elements) = include(&parsed(dxf)?)?;
        style_bytes = style_bytes.checked_add(node_bytes).ok_or_else(|| invalid("CF_SOURCE_TOTAL_LIMIT: conditional-format styles part byte counter overflow"))?;
        style_elements = style_elements.checked_add(node_elements).ok_or_else(|| invalid("CF_SOURCE_TOTAL_LIMIT: conditional-format styles part element counter overflow"))?;
        if style_bytes > style_limits.total_bytes || style_elements > style_limits.total_elements {
            return Err(invalid("CF_SOURCE_TOTAL_LIMIT: selected conditional-format differential styles exceed the styles part budget"))
        }
    }
    if let Some(theme) = theme { include(&parsed(theme)?)?; }
    Ok(())
}

pub(super) fn theme_for_workbook(workbook: &WorkbookInput) -> Result<Option<String>, ExportError> {
    let mut model_budget = ConditionalFormatModelBudget::default();
    for rule in workbook.sheets.iter().flat_map(|sheet| &sheet.conditional_formats) { model_budget.consume(rule).map_err(invalid)?; }
    let mut selected = None::<(String, String)>;
    for source in workbook.sheets.iter().flat_map(|sheet| &sheet.conditional_formats).filter_map(|rule| rule.office_source.as_ref()) {
        let parts = read_source(Some(source))?;
        if let Some(theme) = parts.theme {
            let fingerprint = theme.fingerprint();
            if selected.as_ref().is_some_and(|(current, _)| current != &fingerprint) { return Err(invalid("conditional formats reference conflicting source themes; a single workbook theme cannot preserve both")) }
            if selected.is_none() { selected = Some((fingerprint, source.theme_xml.as_ref().unwrap().clone())); }
        }
    }
    Ok(selected.map(|(_, xml)| xml))
}

fn color(name: &str, value: &str, x14: bool) -> Result<Element, ExportError> {
    let rgb = normalize_color(value);
    if rgb.len() != 8 || !rgb.bytes().all(|byte| byte.is_ascii_hexdigit()) { return Err(invalid("conditional-format color must be #RRGGBB")) }
    let mut node = if x14 { Element::prefixed(X14_NAMESPACE, "x14", name) } else { Element::new(MAIN_NAMESPACE, name) }; node.set_attr("rgb", rgb); Ok(node)
}
fn merged_color(mut authored: Element, generated: &Element) -> Element {
    for key in ["rgb", "theme", "indexed", "auto", "tint"] { copy_attr(&mut authored, generated, key); }
    authored
}
fn changed_style(diff: &Diff, key: &str) -> bool { diff.baseline.as_ref().is_none_or(|before| before["style"][key] != diff.current["style"][key]) }
fn differential_format(style: &ConditionalFormatStyleInput, source: Option<Element>, diff: &Diff) -> Result<Option<Element>, ExportError> {
    let mut dxf = source.unwrap_or_else(|| Element::new(MAIN_NAMESPACE, "dxf")); dxf_root(&mut dxf, false);
    let mut font = dxf.child("font").cloned().unwrap_or_else(|| Element::new(MAIN_NAMESPACE, "font"));
    if changed_style(diff, "fontColor") {
        if let Some(value) = &style.font_color {
            let generated = color("color", value, false)?; let existing = font.child("color").cloned().unwrap_or_else(|| generated.clone()); upsert(&mut font, merged_color(existing, &generated));
        } else { font.replace_children(|node| node.is_main("color"), Vec::new()); }
    }
    for (field, name, flag) in [("bold", "b", style.bold), ("italic", "i", style.italic), ("underline", "u", style.underline), ("strike", "strike", style.strike)] {
        if !changed_style(diff, field) { continue }
        if let Some(flag) = flag {
            let mut node = font.child(name).cloned().unwrap_or_else(|| Element::new(MAIN_NAMESPACE, name));
            if flag { if name == "u" { node.set_attr("val", "single"); } else { node.remove_attr("val"); } }
            else { node.set_attr("val", if name == "u" { "none" } else { "0" }); }
            upsert(&mut font, node);
        } else { font.replace_children(|node| node.is_main(name), Vec::new()); }
    }
    if !font.content.is_empty() || !font.attrs.is_empty() {
        if dxf.child("font").is_some() { upsert(&mut dxf, font); }
        else { dxf.insert_before(|node| node.is_main("numFmt") || node.is_main("fill") || node.is_main("alignment") || node.is_main("border") || node.is_main("protection") || ext_list(node), font); }
    } else { dxf.replace_children(|node| node.is_main("font"), Vec::new()); }
    if changed_style(diff, "fill") {
        if let Some(value) = &style.fill {
            let mut fill = dxf.child("fill").cloned().unwrap_or_else(|| Element::new(MAIN_NAMESPACE, "fill"));
            let mut pattern = fill.child("patternFill").cloned().unwrap_or_else(|| Element::new(MAIN_NAMESPACE, "patternFill")); pattern.set_attr("patternType", "solid");
            let generated = color("fgColor", value, false)?; let existing = pattern.child("fgColor").cloned().unwrap_or_else(|| generated.clone()); upsert(&mut pattern, merged_color(existing, &generated));
            if fill.children.iter().any(|node| node.is_main("patternFill") || node.is_main("gradientFill")) {
                fill.replace_children(|node| node.is_main("patternFill") || node.is_main("gradientFill"), vec![pattern]);
            } else { fill.push(pattern); }
            if dxf.child("fill").is_some() { upsert(&mut dxf, fill); }
            else { dxf.insert_before(|node| node.is_main("alignment") || node.is_main("border") || node.is_main("protection") || ext_list(node), fill); }
        } else { dxf.replace_children(|node| node.is_main("fill"), Vec::new()); }
    }
    Ok((!dxf.content.is_empty() || dxf.attrs.iter().any(|(key, _)| key != "xmlns" && !key.starts_with("xmlns:"))).then_some(dxf))
}

fn cfvo(value: Option<&Value>, default: &str, x14: bool, operator: Option<&str>) -> Result<Element, ExportError> {
    let value = value.unwrap_or(&Value::Null); let authored_kind = value["type"].as_str().unwrap_or(default);
    let kind = if !x14 && authored_kind == "autoMin" { "min" } else if !x14 && authored_kind == "autoMax" { "max" } else { authored_kind };
    if !matches!(kind, "min" | "max" | "num" | "percent" | "percentile" | "formula" | "autoMin" | "autoMax") { return Err(invalid("unsupported Office conditional-format threshold type")) }
    let mut node = if x14 { Element::prefixed(X14_NAMESPACE, "x14", "cfvo") } else { Element::new(MAIN_NAMESPACE, "cfvo") }; node.set_attr("type", kind);
    if operator == Some("greaterThan") { node.set_attr("gte", "0"); }
    if !matches!(kind, "min" | "max" | "autoMin" | "autoMax") {
        let amount = value.get("value").and_then(super::json_scalar_string).ok_or_else(|| invalid("conditional-format threshold has no operand"))?;
        let amount = if kind == "formula" { amount.trim().trim_start_matches('=').to_owned() } else { amount };
        if amount.trim().is_empty() { return Err(invalid("conditional-format threshold has an empty operand")) }
        if kind != "formula" {
            let numeric = amount.parse::<f64>().map_err(|_| invalid("conditional-format threshold is not numeric"))?;
            if !numeric.is_finite() || matches!(kind, "percent" | "percentile") && !(0.0..=100.0).contains(&numeric) { return Err(invalid("conditional-format numeric threshold is outside its allowed range")) }
        }
        if x14 { let mut formula = Element::prefixed(XM_NAMESPACE, "xm", "f"); formula.set_text(amount); node.push(formula); }
        else { node.set_attr("val", amount); }
    }
    Ok(node)
}
fn merged_cfvo(mut authored: Element, generated: &Element) -> Element {
    for key in ["type", "val", "gte"] { copy_attr(&mut authored, generated, key); }
    let formulas = generated.children.iter().filter(|node| node.is(XM_NAMESPACE, "f")).cloned().collect::<Vec<_>>();
    if authored.children.iter().any(|node| node.is(XM_NAMESPACE, "f")) { authored.replace_children(|node| node.is(XM_NAMESPACE, "f"), formulas); }
    else { for formula in formulas { authored.insert_before(ext_list, formula); } }
    authored
}
fn merged_threshold(authored: Element, generated: &Element) -> Element {
    // gte is editable in an icon condition; bars/scales do not expose that
    // field, so their original attribute stays independent of an operand edit.
    let gte = authored.attrs.get("gte").cloned(); let mut node = merged_cfvo(authored, generated);
    if let Some(gte) = gte { node.set_attr("gte", gte); } else { node.remove_attr("gte"); }
    node
}

fn visual_order(name: &str) -> usize {
    ["cfvo", "color", "fillColor", "borderColor", "negativeFillColor", "negativeBorderColor", "axisColor", "cfIcon", "extLst"]
        .iter().position(|entry| *entry == name).unwrap_or(usize::MAX)
}
fn upsert_visual(parent: &mut Element, child: Element) {
    let name = child.name.clone();
    if parent.children.iter().any(|node| is_visual(node, &name)) {
        parent.replace_children(|node| is_visual(node, &name), vec![child]);
    } else {
        let position = visual_order(&name);
        parent.insert_before(|node| ext_list(node) || is_visual(node, &node.name) && visual_order(&node.name) > position, child);
    }
}
fn series(parent: &Element, name: &str) -> Vec<Element> { parent.children.iter().filter(|node| is_visual(node, name)).cloned().collect() }
fn set_series(parent: &mut Element, name: &str, children: Vec<Element>) {
    if parent.children.iter().any(|node| is_visual(node, name)) { parent.replace_children(|node| is_visual(node, name), children); }
    else {
        let position = visual_order(name);
        for child in children { parent.insert_before(|node| ext_list(node) || is_visual(node, &node.name) && visual_order(&node.name) > position, child); }
    }
}
fn merge_series_node(authored: Element, generated: &Element, name: &str) -> Element {
    if name == "cfvo" { merged_threshold(authored, generated) }
    else if name == "cfIcon" {
        let mut node = authored; for key in ["iconSet", "iconId"] { copy_attr(&mut node, generated, key); } node
    } else { merged_color(authored, generated) }
}
fn merge_bar_color(authored: &mut Element, generated: &Element, name: &str, changed: bool) {
    if !changed { return }
    if let Some(color) = generated.children.iter().find(|node| is_visual(node, name)) {
        let color = authored.children.iter().find(|node| is_visual(node, name)).cloned()
            .map(|before| merged_color(before, color)).unwrap_or_else(|| color.clone());
        upsert_visual(authored, color);
    } else { authored.replace_children(|node| is_visual(node, name), Vec::new()); }
}
fn merge_visual(mut authored: Element, generated: &Element, diff: &Diff) -> Element {
    let baseline = diff.baseline.as_ref().unwrap_or(&Value::Null);
    match generated.name.as_str() {
        "dataBar" => {
            for (field, attr) in [("isShowValue", "showValue"), ("minLength", "minLength"), ("maxLength", "maxLength"),
                ("isGradient", "gradient"), ("border", "border"), ("direction", "direction"), ("axisPosition", "axisPosition")] {
                if diff.config(field) { copy_attr(&mut authored, generated, attr); }
            }
            let fill_flags = diff.config("negativeBarColorSameAsPositive") || diff.config("nativeColor");
            let border_flags = diff.config("negativeBarBorderColorSameAsPositive") || diff.config("negativeBorderColor");
            if fill_flags { copy_attr(&mut authored, generated, "negativeBarColorSameAsPositive"); }
            if border_flags { copy_attr(&mut authored, generated, "negativeBarBorderColorSameAsPositive"); }
            let old = series(&authored, "cfvo"); let new = series(generated, "cfvo");
            let values = new.iter().enumerate().map(|(index, node)| {
                let key = if index == 0 { "min" } else { "max" };
                let changed = diff.config(key) || diff.current["config"][key]["type"] == "formula";
                match old.get(index) { Some(before) if !changed => before.clone(), Some(before) => merged_threshold(before.clone(), node), None => node.clone() }
            }).collect();
            set_series(&mut authored, "cfvo", values);
            let x14 = generated.namespace == X14_NAMESPACE;
            merge_bar_color(&mut authored, generated, if x14 { "fillColor" } else { "color" }, diff.config("positiveColor"));
            if x14 {
                merge_bar_color(&mut authored, generated, "borderColor", diff.config("borderColor") || diff.config("border")
                    || diff.config("positiveColor") && diff.current["config"]["borderColor"].is_null());
                merge_bar_color(&mut authored, generated, "negativeFillColor", fill_flags);
                merge_bar_color(&mut authored, generated, "negativeBorderColor", border_flags || diff.config("border")
                    || diff.config("nativeColor") && diff.current["config"]["negativeBorderColor"].is_null());
                merge_bar_color(&mut authored, generated, "axisColor", diff.config("axisColor") || diff.config("axisPosition"));
            }
        }
        "colorScale" => {
            let before = baseline["config"]["thresholds"].as_array().map(Vec::as_slice).unwrap_or(&[]);
            let current = diff.current["config"]["thresholds"].as_array().map(Vec::as_slice).unwrap_or(&[]);
            for (name, field) in [("cfvo", "value"), ("color", "color")] {
                let old = series(&authored, name); let new = series(generated, name);
                let values = new.iter().enumerate().map(|(index, node)| {
                    // A two/three-stop conversion retains the original low/high XML.
                    let origin = if old.len() == new.len() || index == 0 { Some(index) }
                        else if index + 1 == new.len() { old.len().checked_sub(1) } else { None };
                    let changed = origin.is_none_or(|origin| before.get(origin).map(|entry| &entry[field]) != current.get(index).map(|entry| &entry[field]))
                        || name == "cfvo" && current.get(index).is_some_and(|entry| entry["value"]["type"] == "formula");
                    match origin.and_then(|index| old.get(index)) {
                        Some(before) if !changed => before.clone(), Some(before) => merge_series_node(before.clone(), node, name), None => node.clone(),
                    }
                }).collect();
                set_series(&mut authored, name, values);
            }
        }
        "iconSet" => {
            for (field, attr) in [("iconSet", "iconSet"), ("isShowValue", "showValue"), ("reverse", "reverse")] {
                if diff.config(field) { copy_attr(&mut authored, generated, attr); }
            }
            let before = baseline["config"]["icons"].as_array().map(Vec::as_slice).unwrap_or(&[]);
            let current = diff.current["config"]["icons"].as_array().map(Vec::as_slice).unwrap_or(&[]);
            let old = series(&authored, "cfvo"); let new = series(generated, "cfvo");
            let count_changed = before.len() != current.len() || old.len() != new.len();
            let values = new.iter().enumerate().map(|(index, node)| {
                let local = current.len().saturating_sub(index + 1);
                let changed = count_changed || before.get(local).map(|entry| (&entry["value"], &entry["operator"]))
                    != current.get(local).map(|entry| (&entry["value"], &entry["operator"]))
                    || current.get(local).is_some_and(|entry| entry["value"]["type"] == "formula");
                match old.get(index) { Some(before) if !changed => before.clone(), Some(before) => merged_cfvo(before.clone(), node), None => node.clone() }
            }).collect();
            set_series(&mut authored, "cfvo", values);
            let glyphs_changed = count_changed || diff.config("iconSet") || before.iter().zip(current).any(|(before, after)|
                before["iconType"] != after["iconType"] || before["iconId"] != after["iconId"]);
            if glyphs_changed {
                copy_attr(&mut authored, generated, "custom");
                let old = series(&authored, "cfIcon"); let new = series(generated, "cfIcon");
                let values = new.iter().enumerate().map(|(index, node)| old.get(index).cloned()
                    .map(|before| merge_series_node(before, node, "cfIcon")).unwrap_or_else(|| node.clone())).collect();
                set_series(&mut authored, "cfIcon", values);
            }
        }
        _ => {},
    }
    authored
}

fn formula_child(node: &Element) -> bool { node.is_main("formula") || node.is(XM_NAMESPACE, "f") }
fn rule_body(node: &Element) -> bool { formula_child(node) || ["dataBar", "colorScale", "iconSet"].iter().any(|name| is_visual(node, name)) }
fn remove_link(rule: &mut Element) {
    for list in rule.children.iter_mut().filter(|node| ext_list(node)) {
        list.replace_children(|node| node.is_main("ext") && node.attrs.get("uri").is_some_and(|uri| uri.eq_ignore_ascii_case(RULE_LINK_EXTENSION_URI)), Vec::new());
    }
    rule.replace_children(|node| ext_list(node) && node.content.is_empty(), Vec::new());
}
fn merge_rule(source: Option<Element>, generated: Element, diff: &Diff, projected: &ConditionalFormatInput,
    _warnings: &mut Vec<ExportWarning>) -> Result<Element, ExportError> {
    let Some(mut source) = source else { return Ok(generated) };
    if generated.namespace == X14_NAMESPACE { to_x14(&mut source); } else { spreadsheet(&mut source); }
    let schema_changed = source.attrs.get("type") != generated.attrs.get("type");
    if schema_changed {
        for key in ["type", "operator", "text", "timePeriod", "rank", "bottom", "percent", "aboveAverage", "equalAverage", "stdDev"] {
            copy_attr(&mut source, &generated, key);
        }
        let bodies = generated.children.iter().filter(|node| rule_body(node)).cloned().collect::<Vec<_>>();
        if source.children.iter().any(rule_body) { source.replace_children(rule_body, bodies); }
        else { for body in bodies { source.insert_before(|node| is_visual(node, "dxf") || ext_list(node), body); } }
        if projected.rule_type != "dataBar" { remove_link(&mut source); }
    } else {
        for (field, attr) in [("operator", "operator"), ("operator", "timePeriod"), ("value", "text"), ("value", "rank"),
            ("isBottom", "bottom"), ("isPercent", "percent"), ("operator", "aboveAverage"), ("operator", "equalAverage")] {
            if diff.field(field) { copy_attr(&mut source, &generated, attr); }
        }
        let changed = diff.field("type") || diff.field("subType") || diff.field("operator") || diff.field("formula1")
            || diff.field("formula2") || diff.field("value") || diff.rebase() || !projected.formula1.trim().is_empty();
        if changed {
            let text = matches!(projected.rule_type.as_str(), "containsText" | "notContainsText" | "beginsWith" | "endsWith")
                || projected.sub_type.as_deref() == Some("text");
            let literal = projected.value.as_ref().and_then(super::json_scalar_string).unwrap_or_else(|| projected.formula1.clone());
            let blanks_or_errors = matches!(projected.operator.as_str(), "containsBlanks" | "notContainsBlanks" | "containsErrors" | "notContainsErrors");
            let preserve_predicate = !projected.formula1.trim().is_empty() && (!text || projected.formula1 != literal || blanks_or_errors)
                && !diff.field("type") && !diff.field("subType") && !diff.field("operator") && (!text || !diff.field("value"));
            let originals = source.children.iter().filter(|node| formula_child(node)).cloned().collect::<Vec<_>>();
            let formulas = if preserve_predicate {
                [Some(projected.formula1.as_str()), projected.formula2.as_deref()].into_iter().flatten().enumerate().map(|(index, value)| {
                    let mut formula = originals.get(index).cloned().unwrap_or_else(||
                        if generated.namespace == X14_NAMESPACE { Element::prefixed(XM_NAMESPACE, "xm", "f") } else { Element::new(MAIN_NAMESPACE, "formula") });
                    formula.set_text(value.trim().trim_start_matches('=')); formula
                }).collect::<Vec<_>>()
            } else { generated.children.iter().filter(|node| formula_child(node)).enumerate().map(|(index, generated)| {
                let mut formula = originals.get(index).cloned().unwrap_or_else(|| generated.clone());
                formula.set_text(&generated.text); formula
            }).collect() };
            if source.children.iter().any(formula_child) { source.replace_children(formula_child, formulas); }
            else { for formula in formulas { source.insert_before(|node| is_visual(node, "dxf") || ext_list(node), formula); } }
        }
        for generated in generated.children.iter().filter(|node| ["dataBar", "colorScale", "iconSet"].iter().any(|name| is_visual(node, name))) {
            let authored = source.children.iter().find(|node| is_visual(node, &generated.name)).cloned();
            let visual = authored.map(|authored| merge_visual(authored, generated, diff)).unwrap_or_else(|| generated.clone());
            upsert(&mut source, visual);
        }
    }
    for key in ["priority", "dxfId"] { copy_attr(&mut source, &generated, key); }
    if diff.field("stopIfTrue") { copy_attr(&mut source, &generated, "stopIfTrue"); }
    Ok(source)
}

fn linked_data_bar(source: Option<Element>, rule: &ConditionalFormatInput, diff: &Diff, id: &str,
    warnings: &mut Vec<ExportWarning>) -> Result<Element, ExportError> {
    let generated = generated_visual(rule, true, diff, warnings)?;
    let (mut group, original) = match source {
        Some(mut group) => {
            let original = group.children.iter().find(|node| node.is(X14_NAMESPACE, "cfRule")).cloned();
            group.replace_children(is_rule, Vec::new()); (group, original)
        }
        None => (Element::prefixed(X14_NAMESPACE, "x14", "conditionalFormatting"), None),
    };
    let mut authored = original.unwrap_or_else(|| Element::prefixed(X14_NAMESPACE, "x14", "cfRule"));
    if authored.attrs.get("type").is_some_and(|kind| kind != "dataBar") { return Err(invalid("linked x14 conditional-format source does not describe a data bar")) }
    let before = authored.children.iter().find(|node| node.is(X14_NAMESPACE, "dataBar")).cloned();
    let visual = before.map(|before| merge_visual(before, &generated, diff)).unwrap_or(generated);
    authored.set_attr("type", "dataBar"); authored.set_attr("id", id); authored.remove_attr("priority");
    upsert(&mut authored, visual);
    set_group_reference(&mut group, true, &rule.ranges)?;
    group.insert_before(|node| node.is(XM_NAMESPACE, "sqref") || ext_list(node), authored);
    Ok(group)
}

fn icon_requires_x14(rule: &ConditionalFormatInput) -> bool {
    let config = rule.config.as_ref().unwrap_or(&Value::Null); let name = config["iconSet"].as_str().unwrap_or("3TrafficLights1");
    if matches!(name, "3Stars" | "3Triangles" | "5Boxes" | "_5Felling") || office_icon_type(name).is_none() { return true }
    let Some(icons) = config["icons"].as_array() else { return false };
    icons.iter().enumerate().any(|(index, icon)| {
        let kind = icon["iconType"].as_str().unwrap_or(name); let id = icon["iconId"].as_str().and_then(|value| value.parse().ok()).or_else(|| icon["iconId"].as_u64().map(|value| value as usize));
        office_icon_type(kind) != office_icon_type(name) || id.and_then(|id| icon_index(kind, id)) != Some(icons.len() - index - 1)
    })
}
fn generated_visual(rule: &ConditionalFormatInput, x14: bool, diff: &Diff, warnings: &mut Vec<ExportWarning>) -> Result<Element, ExportError> {
    let config = rule.config.as_ref().unwrap_or(&Value::Null);
    let mut node = if x14 { Element::prefixed(X14_NAMESPACE, "x14", &rule.rule_type) } else { Element::new(MAIN_NAMESPACE, &rule.rule_type) };
    match rule.rule_type.as_str() {
        "dataBar" => {
            let min_length = config["minLength"].as_u64().unwrap_or(0); let max_length = config["maxLength"].as_u64().unwrap_or(100);
            if min_length > max_length || max_length > 100 { return Err(invalid("conditional-format bar length is outside 0..100")) }
            node.set_attr("minLength", min_length.to_string()); node.set_attr("maxLength", max_length.to_string());
            node.set_attr("showValue", if config["isShowValue"].as_bool().unwrap_or(true) { "1" } else { "0" });
            node.push(cfvo(config.get("min"), "min", x14, None)?); node.push(cfvo(config.get("max"), "max", x14, None)?);
            let positive = config["positiveColor"].as_str().unwrap_or("#638EC6");
            node.push(color(if x14 { "fillColor" } else { "color" }, positive, x14)?);
            if x14 {
                node.set_attr("gradient", if config["isGradient"].as_bool().unwrap_or(false) { "1" } else { "0" });
                let border = config["border"].as_bool().unwrap_or(false); node.set_attr("border", if border { "1" } else { "0" });
                let direction = config["direction"].as_str().unwrap_or("context"); let axis = config["axisPosition"].as_str().unwrap_or("automatic");
                if !matches!(direction, "context" | "leftToRight" | "rightToLeft") || !matches!(axis, "automatic" | "middle" | "none") { return Err(invalid("invalid conditional-format data-bar axis or direction")) }
                node.set_attr("direction", direction); node.set_attr("axisPosition", axis);
                if border { node.push(color("borderColor", config["borderColor"].as_str().unwrap_or(positive), true)?); }
                let same_fill = config["negativeBarColorSameAsPositive"].as_bool().unwrap_or(false)
                    && !(diff.config("nativeColor") && config["nativeColor"].is_string());
                node.set_attr("negativeBarColorSameAsPositive", if same_fill { "1" } else { "0" });
                let negative = config["nativeColor"].as_str().unwrap_or("#F8696B");
                if !same_fill { node.push(color("negativeFillColor", negative, true)?); }
                let same_border = config["negativeBarBorderColorSameAsPositive"].as_bool().unwrap_or(true)
                    && !(diff.config("negativeBorderColor") && config["negativeBorderColor"].is_string());
                node.set_attr("negativeBarBorderColorSameAsPositive", if same_border { "1" } else { "0" });
                if border && !same_border { node.push(color("negativeBorderColor", config["negativeBorderColor"].as_str().unwrap_or(negative), true)?); }
                if axis != "none" { node.push(color("axisColor", config["axisColor"].as_str().unwrap_or("#000000"), true)?); }
            }
        }
        "colorScale" => {
            let thresholds = config["thresholds"].as_array().ok_or_else(|| invalid("conditional-format color scale has no stops"))?;
            if !(2..=3).contains(&thresholds.len()) { return Err(invalid("conditional-format color scale needs two or three stops")) }
            for threshold in thresholds { node.push(cfvo(threshold.get("value"), "num", x14, None)?); }
            for threshold in thresholds { node.push(color("color", threshold["color"].as_str().ok_or_else(|| invalid("color-scale stop has no color"))?, x14)?); }
        }
        "iconSet" => {
            let icons = config["icons"].as_array().ok_or_else(|| invalid("conditional-format icon set has no thresholds"))?;
            if !(3..=5).contains(&icons.len()) { return Err(invalid("conditional-format icon set needs three to five icons")) }
            let selected = config["iconSet"].as_str().unwrap_or("3TrafficLights1");
            let fallback = if icons.len() == 3 { "3Arrows" } else if icons.len() == 4 { "4Rating" } else { "5Rating" };
            let name = office_icon_type(selected).filter(|name| *name != "NoIcons").unwrap_or(fallback);
            if office_icon_type(selected).is_none() { warnings.push(ExportWarning::new("CONDITIONAL_ICON_OFFICE_MAPPING", format!("Conditional-format {} uses {selected}; Office receives {name}, while the private extension retains its local glyphs", rule.id))); }
            node.set_attr("iconSet", name); node.set_attr("showValue", if config["isShowValue"].as_bool().unwrap_or(true) { "1" } else { "0" });
            node.set_attr("reverse", if config["reverse"].as_bool().unwrap_or(false) { "1" } else { "0" });
            for (position, icon) in icons.iter().rev().enumerate() {
                if position == 0 { node.push(cfvo(Some(&json!({ "type": "percent", "value": 0 })), "percent", x14, None)?); }
                else { node.push(cfvo(icon.get("value"), "num", x14, icon["operator"].as_str())?); }
            }
            if x14 && icon_requires_x14(rule) {
                node.set_attr("custom", "1");
                for (position, icon) in icons.iter().rev().enumerate() {
                    let authored = icon["iconType"].as_str().unwrap_or(selected); let kind = office_icon_type(authored).unwrap_or(name);
                    let local_id = icon["iconId"].as_str().and_then(|value| value.parse::<usize>().ok()).or_else(|| icon["iconId"].as_u64().map(|value| value as usize)).unwrap_or(icons.len() - position - 1);
                    let office_id = if office_icon_type(authored).is_none() { icons.len() - local_id.min(icons.len() - 1) - 1 }
                        else { icon_index(authored, local_id).ok_or_else(|| invalid("conditional-format icon identity is outside its selected set"))? };
                    let mut icon = Element::prefixed(X14_NAMESPACE, "x14", "cfIcon"); icon.set_attr("iconSet", kind); icon.set_attr("iconId", office_id.to_string()); node.push(icon);
                }
            }
        }
        _ => return Err(invalid("unsupported conditional-format visual type")),
    }
    Ok(node)
}

fn to_x14(node: &mut Element) {
    if node.is_main("cfRule") { rename(node, X14_NAMESPACE, "x14", "cfRule"); node.remove_attr("dxfId"); }
    else if node.is_main("formula") { rename(node, XM_NAMESPACE, "xm", "f"); }
    else if ["dataBar", "colorScale", "iconSet", "cfvo", "color", "extLst", "dxf"].iter().any(|name| node.is_main(name)) {
        let name = node.name.clone(); rename(node, X14_NAMESPACE, "x14", &name);
        if name == "cfvo" && let Some(value) = node.attrs.get("val").cloned() {
            node.remove_attr("val"); let mut formula = Element::prefixed(XM_NAMESPACE, "xm", "f"); formula.set_text(value.trim().trim_start_matches('=')); node.insert_before(ext_list, formula);
        }
    }
    // Differential-format children use the base SpreadsheetML types.
    if node.is(X14_NAMESPACE, "dxf") { return }
    for child in &mut node.children { if child.is_main("ext") { continue } to_x14(child); }
}

#[derive(Clone)]
struct FormulaBinding { path: Vec<usize>, attribute: bool, value: String }
fn bindings(root: &Element) -> Vec<FormulaBinding> {
    fn collect(node: &Element, path: &mut Vec<usize>, result: &mut Vec<FormulaBinding>, office: bool) {
        let office = office || is_rule(node) || is_visual(node, "conditionalFormatting");
        if office && (node.is_main("formula") || node.is(XM_NAMESPACE, "f")) { result.push(FormulaBinding { path: path.clone(), attribute: false, value: node.text.clone() }); return }
        if office && is_visual(node, "cfvo") && node.attrs.get("type").is_some_and(|kind| kind == "formula") && let Some(value) = node.attrs.get("val") {
            result.push(FormulaBinding { path: path.clone(), attribute: true, value: value.clone() });
        }
        for (index, child) in node.children.iter().enumerate() {
            // Vendor extensions can contain references with unverified semantics.
            if ext_list(child) || !(is_rule(child) || is_visual(child, "dataBar") || is_visual(child, "colorScale") || is_visual(child, "iconSet")
                || is_visual(child, "cfvo") || child.is_main("formula") || child.is(XM_NAMESPACE, "f")) { continue }
            path.push(index); collect(child, path, result, office); path.pop();
        }
    }
    let mut result = Vec::new(); collect(root, &mut Vec::new(), &mut result, false); result
}
fn apply_bindings(root: &mut Element, bindings: &[FormulaBinding], formulas: &[String]) {
    for (binding, formula) in bindings.iter().zip(formulas) {
        if binding.value == *formula { continue }
        let mut node = &mut *root; for index in &binding.path { node = &mut node.children[*index]; }
        if binding.attribute { node.set_attr("val", formula); } else { node.set_text(formula); }
    }
}
fn project_opaque(rule: &ConditionalFormatInput, mut source: Element, mut x14: Option<Element>, group: Option<&Element>, diff: &Diff)
    -> Result<(ConditionalFormatInput, Element, Option<Element>), ExportError> {
    if diff.rebase() && x14.as_ref().is_some_and(|linked| group_reference(linked) != group.and_then(group_reference)) {
        return Err(invalid("range transforms of an unknown linked conditional format with differing source scopes require a verified mapping"))
    }
    let authored = bindings(&source); let linked = x14.as_ref().map(bindings).unwrap_or_default();
    let origin = group.and_then(group_reference).and_then(|reference| reference.split_whitespace().next()).and_then(|reference| reference.split(':').next())
        .and_then(|cell| super::parse_a1(&cell.replace('$', ""))).ok_or_else(|| invalid("unknown conditional format has no original formula origin"))?;
    let mut synthetic = rule.clone(); synthetic.rule_type = "expression".into(); synthetic.sub_type = None; synthetic.formula1.clear(); synthetic.formula2 = None; synthetic.value = None;
    synthetic.formula_anchor = Some(conditional_formula::FormulaAnchor { row: origin.0, col: origin.1 });
    synthetic.config = Some(json!({ "thresholds": authored.iter().chain(&linked).map(|binding| json!({ "value": { "type": "formula", "value": binding.value } })).collect::<Vec<_>>() }));
    let projected = conditional_formula::project_rule(&synthetic)?;
    let formulas = projected.config.as_ref().unwrap()["thresholds"].as_array().unwrap().iter().map(|entry| entry["value"]["value"].as_str().unwrap().to_owned()).collect::<Vec<_>>();
    apply_bindings(&mut source, &authored, &formulas[..authored.len()]);
    if let Some(x14) = &mut x14 {
        apply_bindings(x14, &linked, &formulas[authored.len()..]);
        if diff.rebase() { set_group_reference(x14, true, &projected.ranges)?; }
    }
    let mut result = rule.clone(); result.ranges = projected.ranges; Ok((result, source, x14))
}

pub(super) fn average_predicate(rule: &ConditionalFormatInput, address: &str) -> Result<String, ExportError> {
    // AVERAGE must count every cell once, including overlapping noncontiguous areas.
    let mut areas = Vec::<(u32, u32, u32, u32)>::new();
    for range in &rule.ranges {
        let mut remaining = vec![range_bounds(range).ok_or_else(|| invalid("invalid average conditional-format range"))?];
        for &(sr, sc, er, ec) in &areas {
            remaining = remaining.into_iter().flat_map(|(ar, ac, br, bc)| {
                let (top, left, bottom, right) = (ar.max(sr), ac.max(sc), br.min(er), bc.min(ec));
                if top > bottom || left > right { return vec![(ar, ac, br, bc)] }
                let mut pieces = Vec::new();
                if ar < top { pieces.push((ar, ac, top - 1, bc)); } if bottom < br { pieces.push((bottom + 1, ac, br, bc)); }
                if ac < left { pieces.push((top, ac, bottom, left - 1)); } if right < bc { pieces.push((top, right + 1, bottom, bc)); } pieces
            }).collect();
        }
        areas.extend(remaining);
        if areas.len() > 100_000 { return Err(invalid("average conditional-format scope exceeds the range budget")) }
    }
    let references = areas.into_iter().map(|(sr, sc, er, ec)| format!("${}${}:${}${}", super::column_name(sc), sr + 1, super::column_name(ec), er + 1)).collect::<Vec<_>>().join(",");
    let formula = format!("AND(ISNUMBER({address}),{address}{}AVERAGE({references}))", if rule.operator == "notEqual" { "<>" } else { "=" });
    if formula.len() > 8192 { return Err(invalid("average conditional-format formula exceeds the Office formula limit")) } Ok(formula)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn model(kind: &str, config: Value, style: Value) -> Value {
        json!({ "id": "local-rule", "type": kind, "ranges": [{ "startRow": 0, "startCol": 0, "endRow": 3, "endCol": 0 }],
            "operator": "", "formula1": "", "priority": 17, "stopIfTrue": false, "config": config, "style": style })
    }
    fn diff(before: &Value, after: &Value) -> Diff { Diff { current: known_snapshot(after), baseline: Some(known_snapshot(before)) } }
    fn input(model: &Value) -> ConditionalFormatInput { serde_json::from_value(model.clone()).unwrap() }

    #[test]
    fn one_font_edit_preserves_other_dxf_fields_theme_colors_double_underline_and_extensions() {
        let before = model("expression", json!({}), json!({ "bold": false, "underline": true, "fontColor": "#FF8080", "fill": "#336699" }));
        let mut after = before.clone(); after["style"]["bold"] = json!(true);
        let authored = parsed(&format!(r#"<dxf xmlns="{MAIN_NAMESPACE}" xmlns:v="urn:vendor"><font><name val="Office Font"/><sz val="14"/><u val="double"/><color theme="4" tint="0.5"/><scheme val="minor"/><b val="0"/></font><numFmt numFmtId="166" formatCode="0.000"/><fill><patternFill patternType="solid"><fgColor theme="5"/></patternFill></fill><alignment wrapText="1"/><border><bottom style="double"><color indexed="10"/></bottom></border><protection locked="0"/><extLst><ext uri="vendor"><v:payload>before<v:b/>after</v:payload></ext></extLst></dxf>"#)).unwrap();
        let retained = differential_format(&input(&after).style, Some(authored.clone()), &diff(&before, &after)).unwrap().unwrap();
        let font = retained.child("font").unwrap();
        assert_eq!(font.child("u").unwrap().attrs.get("val").map(String::as_str), Some("double"));
        assert!(font.child("b").unwrap().attrs.get("val").is_none());
        for name in ["name", "sz", "color", "scheme"] { assert_eq!(font.child(name).unwrap().fingerprint(), authored.child("font").unwrap().child(name).unwrap().fingerprint()); }
        for name in ["numFmt", "fill", "alignment", "border", "protection", "extLst"] {
            assert_eq!(retained.child(name).unwrap().fingerprint(), authored.child(name).unwrap().fingerprint());
        }
    }

    #[test]
    fn editing_one_color_scale_stop_keeps_the_other_threshold_and_color_xml() {
        let before = model("colorScale", json!({ "thresholds": [
            { "value": { "type": "min" }, "color": "#111111" }, { "value": { "type": "num", "value": 50 }, "color": "#222222" },
            { "value": { "type": "max" }, "color": "#333333" }] }), json!({}));
        let mut after = before.clone(); after["config"]["thresholds"][1]["value"]["value"] = json!(70);
        let authored = parsed(&format!(r#"<colorScale xmlns="{MAIN_NAMESPACE}" xmlns:v="urn:vendor"><cfvo type="min" v:tag="low"/><cfvo type="num" val="50" v:tag="mid"/><cfvo type="max"/><color theme="0" tint="-0.2"/><color rgb="FF222222"/><color indexed="10"/></colorScale>"#)).unwrap();
        let generated = generated_visual(&input(&after), false, &diff(&before, &after), &mut Vec::new()).unwrap();
        let retained = merge_visual(authored.clone(), &generated, &diff(&before, &after));
        let values = series(&retained, "cfvo"); assert_eq!(values.len(), 3);
        assert_eq!(values[1].attrs.get("val").map(String::as_str), Some("70")); assert_eq!(values[1].attrs.get("v:tag").map(String::as_str), Some("mid"));
        assert_eq!(values[0].fingerprint(), series(&authored, "cfvo")[0].fingerprint());
        assert_eq!(values[2].fingerprint(), series(&authored, "cfvo")[2].fingerprint());
        assert_eq!(series(&retained, "color").iter().map(Element::fingerprint).collect::<Vec<_>>(), series(&authored, "color").iter().map(Element::fingerprint).collect::<Vec<_>>());
    }

    #[test]
    fn linked_bar_changes_only_the_negative_fill_and_retains_axis_border_thresholds_and_guid() {
        const ID: &str = "{11111111-2222-4333-8444-555555555555}";
        let before = model("dataBar", json!({ "min": { "type": "formula", "value": "A1" }, "max": { "type": "num", "value": 100 },
            "isGradient": true, "positiveColor": "#336699", "isShowValue": false, "border": true, "borderColor": "#112233", "axisColor": "#445566",
            "direction": "rightToLeft", "axisPosition": "middle", "negativeBarColorSameAsPositive": true }), json!({}));
        let mut after = before.clone(); after["config"]["nativeColor"] = json!("#AA3344");
        let authored = parsed(&format!(r#"<x14:conditionalFormatting xmlns:x14="{X14_NAMESPACE}" xmlns:xm="{XM_NAMESPACE}"><x14:cfRule type="dataBar" id="{ID}"><x14:dataBar showValue="0" border="1" gradient="1" direction="rightToLeft" axisPosition="middle" negativeBarColorSameAsPositive="1"><x14:cfvo type="formula"><xm:f>A1</xm:f></x14:cfvo><x14:cfvo type="num"><xm:f>100</xm:f></x14:cfvo><x14:fillColor rgb="FF336699"/><x14:borderColor theme="4" tint="0.5"/><x14:axisColor indexed="10"/></x14:dataBar></x14:cfRule><xm:sqref>A1:A4</xm:sqref></x14:conditionalFormatting>"#)).unwrap();
        let original = authored.children.iter().find(|node| node.is(X14_NAMESPACE, "cfRule")).unwrap().children[0].clone();
        let retained = linked_data_bar(Some(authored), &input(&after), &diff(&before, &after), ID, &mut Vec::new()).unwrap();
        let rule = retained.children.iter().find(|node| node.is(X14_NAMESPACE, "cfRule")).unwrap();
        assert_eq!(rule.attrs.get("id").map(String::as_str), Some(ID)); assert!(rule.attrs.get("priority").is_none());
        let bar = rule.children.iter().find(|node| node.is(X14_NAMESPACE, "dataBar")).unwrap();
        assert_eq!(bar.attrs.get("negativeBarColorSameAsPositive").map(String::as_str), Some("0"));
        assert_eq!(series(bar, "negativeFillColor")[0].attrs.get("rgb").map(String::as_str), Some("FFAA3344"));
        for name in ["cfvo", "fillColor", "borderColor", "axisColor"] {
            assert_eq!(series(bar, name).iter().map(Element::fingerprint).collect::<Vec<_>>(), series(&original, name).iter().map(Element::fingerprint).collect::<Vec<_>>());
        }
    }

    #[test]
    fn custom_icon_threshold_edit_preserves_cficon_glyphs_and_noicons_instead_of_rebuilding_them() {
        let before = model("iconSet", json!({ "iconSet": "3Arrows", "isShowValue": false, "icons": [
            { "value": { "type": "num", "value": 70 }, "operator": "greaterThanOrEqual", "iconType": "3Stars", "iconId": "0" },
            { "value": { "type": "num", "value": 40 }, "operator": "greaterThanOrEqual", "iconType": "EMPTY_ICON_TYPE", "iconId": "" },
            { "value": { "type": "percent", "value": 0 }, "iconType": "3Arrows", "iconId": "2" }] }), json!({}));
        let mut after = before.clone(); after["config"]["icons"][0]["value"]["value"] = json!(80); after["config"]["icons"][0]["operator"] = json!("greaterThan");
        let authored = parsed(&format!(r#"<x14:iconSet xmlns:x14="{X14_NAMESPACE}" xmlns:xm="{XM_NAMESPACE}" xmlns:v="urn:vendor" iconSet="3Arrows" custom="1" showValue="0"><x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="num"><xm:f>40</xm:f></x14:cfvo><x14:cfvo type="num"><xm:f>70</xm:f></x14:cfvo><x14:cfIcon iconSet="3Arrows" iconId="0" v:tag="low"/><x14:cfIcon iconSet="NoIcons" iconId="0"/><x14:cfIcon iconSet="3Stars" iconId="2"/></x14:iconSet>"#)).unwrap();
        let generated = generated_visual(&input(&after), true, &diff(&before, &after), &mut Vec::new()).unwrap();
        let retained = merge_visual(authored.clone(), &generated, &diff(&before, &after));
        assert_eq!(series(&retained, "cfIcon").iter().map(Element::fingerprint).collect::<Vec<_>>(), series(&authored, "cfIcon").iter().map(Element::fingerprint).collect::<Vec<_>>());
        let values = series(&retained, "cfvo"); assert_eq!(values.len(), 3); assert_eq!(values[2].attrs.get("gte").map(String::as_str), Some("0"));
        assert_eq!(values[2].children[0].text, "80"); assert_eq!(values[0].fingerprint(), series(&authored, "cfvo")[0].fingerprint());
    }

    #[test]
    fn office_guids_are_reserved_stable_for_originals_and_remapped_for_copies() {
        const ID: &str = "{11111111-2222-4333-8444-555555555555}";
        let mut guids = Guids::default(); guids.reserved.insert(ID.into());
        assert_eq!(guids.assign(Some(ID), false, "original").unwrap(), ID);
        let copied = guids.assign(Some(ID), true, "copy").unwrap(); assert_ne!(copied, ID); assert!(valid_guid(&copied));
        let duplicate = guids.assign(Some(ID), false, "duplicate").unwrap(); assert_ne!(duplicate, ID); assert_ne!(duplicate, copied);
        assert!(guids.assign(Some("bad-id"), false, "invalid").is_err());
    }

    #[test]
    fn updating_opaque_standard_formula_paths_retains_linked_rules_and_vendor_references() {
        let before = json!({ "id": "opaque", "type": "office", "ranges": [{ "startRow": 1, "startCol": 1, "endRow": 4, "endCol": 1 }],
            "operator": "future", "formula1": "A1>0", "style": {}, "priority": 1, "stopIfTrue": false });
        let mut after = before.clone(); after["ranges"] = json!([{ "startRow": 3, "startCol": 3, "endRow": 6, "endCol": 3 }]);
        let source = parsed(&format!(r#"<cfRule xmlns="{MAIN_NAMESPACE}" type="future"><formula>A1&gt;0</formula><extLst><ext uri="vendor"><v:payload xmlns:v="urn:vendor" ref="A1"/></ext></extLst></cfRule>"#)).unwrap();
        let group = parsed(&format!(r#"<conditionalFormatting xmlns="{MAIN_NAMESPACE}" sqref="B2:B5"/>"#)).unwrap();
        let linked = parsed(&format!(r#"<x14:conditionalFormatting xmlns:x14="{X14_NAMESPACE}" xmlns:xm="{XM_NAMESPACE}"><x14:cfRule type="future"><xm:f>B2</xm:f></x14:cfRule><xm:sqref>B2:B5</xm:sqref></x14:conditionalFormatting>"#)).unwrap();
        let (_, retained, linked) = project_opaque(&input(&after), source, Some(linked), Some(&group), &diff(&before, &after)).unwrap();
        assert_eq!(retained.child("formula").unwrap().text, "C3>0"); assert_eq!(retained.descendants("payload")[0].attrs.get("ref").map(String::as_str), Some("A1"));
        let linked = linked.unwrap(); assert_eq!(linked.children.iter().filter(|node| node.is(X14_NAMESPACE, "cfRule")).count(), 1);
        assert_eq!(group_reference(&linked), Some("D4:D7")); assert_eq!(linked.children.iter().find(|node| is_rule(node)).unwrap().children[0].text, "D4");
    }

    #[test]
    fn average_equality_uses_absolute_disjoint_scope_and_excludes_overlapping_cells_from_double_counting() {
        let mut rule = input(&model("average", json!({}), json!({}))); rule.operator = "notEqual".into();
        rule.ranges = serde_json::from_value(json!([{ "startRow": 0, "startCol": 0, "endRow": 2, "endCol": 0 },
            { "startRow": 1, "startCol": 0, "endRow": 3, "endCol": 0 }])).unwrap();
        assert_eq!(average_predicate(&rule, "A1").unwrap(), "AND(ISNUMBER(A1),A1<>AVERAGE($A$1:$A$3,$A$4:$A$4))");
    }

    #[test]
    fn unresolved_source_dxf_rejects_export_even_when_an_edit_supplies_a_new_style() {
        let mut before = model("expression", json!({}), json!({})); before["formula1"] = json!("A1>0");
        let source = OfficeSource { version: 1, format: "xlsx".into(), original_id: "local-rule".into(), storage: "base".into(),
            rule_xml: format!(r#"<cfRule xmlns="{MAIN_NAMESPACE}" type="expression" priority="1" dxfId="9"><formula>A1&gt;0</formula></cfRule>"#),
            group_xml: Some(format!(r#"<conditionalFormatting xmlns="{MAIN_NAMESPACE}" sqref="A1:A4"/>"#)),
            differential_format_xml: None, x14_xml: None, theme_xml: None,
            baseline: known_snapshot(&before), extensions: serde_json::Map::new() };
        for style in [json!({}), json!({ "bold": true }), json!({ "fill": "#336699" })] {
            let mut after = before.clone(); after["style"] = style; after["officeSource"] = serde_json::to_value(&source).unwrap();
            let sheet: SheetInput = serde_json::from_value(json!({ "id": "sheet", "name": "Data", "conditionalFormats": [after] })).unwrap();
            let mut registry = StyleRegistry::new(); let mut warnings = Vec::new();
            let error = write(&sheet, &mut registry, &mut warnings).err().expect("an unavailable source dxf must reject before an edited replacement is registered");
            assert_eq!(error.code, "INVALID_MODEL"); assert!(registry.conditional_styles.is_empty());
        }
    }

    #[test]
    fn final_worksheet_usage_keeps_late_cf_after_unrelated_large_text_and_validates_the_suffix() {
        let data = "x".repeat(source_xml::XML_FRAGMENT_BYTE_LIMIT + 1);
        let group = r#"<conditionalFormatting sqref="A1"><cfRule type="expression" priority="1"><formula>A1&gt;0</formula></cfRule></conditionalFormatting>"#;
        let xml = format!(r#"<worksheet xmlns="{MAIN_NAMESPACE}"><sheetData><row><c><is><t>{data}</t></is></c></row></sheetData>{group}</worksheet>"#);
        let (bytes, elements) = output_source_usage(&xml).unwrap();
        assert!(bytes > 0 && bytes < 1024); assert_eq!(elements, 3);
        let error = output_source_usage(&xml.replace("</worksheet>", "<broken></worksheet>")).unwrap_err();
        assert_eq!(error.code, "INVALID_MODEL"); assert!(error.message.starts_with("CF_SOURCE_XML_INVALID"));
    }

    #[test]
    fn output_resource_budget_counts_registered_dxfs_theme_and_individual_style_part_boundaries() {
        let mut registry = StyleRegistry::new();
        let dxf = parsed(&format!(r#"<dxf xmlns="{MAIN_NAMESPACE}"><font><b/></font></dxf>"#)).unwrap();
        registry.register_conditional_style_xml(&dxf);
        let theme = r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements/></a:theme>"#;
        let theme_size = parsed(theme).unwrap().rendered_size();
        let boundary = SOURCE_RESOURCE_BYTE_LIMIT - dxf.rendered_size() - theme_size;
        assert!(validate_output_resource_budget(boundary, SOURCE_RESOURCE_ELEMENT_LIMIT - 5, &registry, Some(theme)).is_ok());
        for (bytes, elements) in [(boundary + 1, 0), (boundary, SOURCE_RESOURCE_ELEMENT_LIMIT - 4)] {
            let error = validate_output_resource_budget(bytes, elements, &registry, Some(theme)).unwrap_err();
            assert_eq!(error.code, "INVALID_MODEL"); assert!(error.message.starts_with("CF_SOURCE_TOTAL_LIMIT"));
        }
        let second_dxf = parsed(&format!(r#"<dxf xmlns="{MAIN_NAMESPACE}"><font><i/></font></dxf>"#)).unwrap();
        registry.register_conditional_style_xml(&second_dxf);
        assert_eq!(registry.conditional_styles.len(), 2);
        let style_bytes = dxf.rendered_size() + second_dxf.rendered_size();
        let exact = XmlSelectionLimits { total_bytes: style_bytes, total_elements: 6, ..XmlSelectionLimits::default() };
        assert!(validate_output_resource_budget_with_style_limits(0, 0, &registry, None, exact).is_ok());
        for limits in [XmlSelectionLimits { total_bytes: style_bytes - 1, ..exact },
            XmlSelectionLimits { total_elements: 5, ..exact }] {
            let error = validate_output_resource_budget_with_style_limits(0, 0, &registry, None, limits).unwrap_err();
            assert_eq!(error.code, "INVALID_MODEL"); assert!(error.message.starts_with("CF_SOURCE_TOTAL_LIMIT"));
            assert!(error.message.contains("styles part"));
        }
    }

    #[test]
    fn composed_cf_output_depth_is_checked_with_the_same_fragment_guard_as_import() {
        let payload = format!("{}x{}", "<v:n>".repeat(128), "</v:n>".repeat(128));
        let xml = format!(r#"<worksheet xmlns="{MAIN_NAMESPACE}" xmlns:v="urn:vendor"><conditionalFormatting sqref="A1"><cfRule type="expression" priority="1"><formula>A1</formula><extLst><ext uri="vendor">{payload}</ext></extLst></cfRule></conditionalFormatting></worksheet>"#);
        let error = output_source_usage(&xml).unwrap_err(); assert_eq!(error.code, "INVALID_MODEL");
        assert!(error.message.starts_with("CF_SOURCE_FRAGMENT_LIMIT"));
    }

    #[test]
    fn malformed_dxf_indices_and_unavailable_source_theme_resources_reject_before_any_replacement_style() {
        let mut before = model("expression", json!({}), json!({})); before["formula1"] = json!("A1>0");
        for index in ["-1", "4294967296", "not-index"] {
            let source = OfficeSource { version: 1, format: "xlsx".into(), original_id: "local-rule".into(), storage: "base".into(),
                rule_xml: format!(r#"<cfRule xmlns="{MAIN_NAMESPACE}" type="expression" priority="1" dxfId="{index}"><formula>A1&gt;0</formula></cfRule>"#),
                group_xml: Some(format!(r#"<conditionalFormatting xmlns="{MAIN_NAMESPACE}" sqref="A1:A4"/>"#)),
                differential_format_xml: Some(format!(r#"<dxf xmlns="{MAIN_NAMESPACE}"><font><b/></font></dxf>"#)), x14_xml: None, theme_xml: None,
                baseline: known_snapshot(&before), extensions: serde_json::Map::new() };
            let mut after = before.clone(); after["style"] = json!({ "bold": true }); after["officeSource"] = serde_json::to_value(source).unwrap();
            let sheet: SheetInput = serde_json::from_value(json!({ "id": "sheet", "name": "Data", "conditionalFormats": [after] })).unwrap();
            let mut registry = StyleRegistry::new();
            let error = write(&sheet, &mut registry, &mut Vec::new()).err().expect("an invalid original dxfId must reject before a replacement style");
            assert_eq!(error.code, "INVALID_MODEL"); assert!(error.message.starts_with("CF_SOURCE_XML_INVALID"));
            assert!(registry.conditional_styles.is_empty());
        }
        let empty_theme = r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements/></a:theme>"#.to_owned();
        for (dxf_body, theme_xml) in [
            (r#"<font><color theme="4"/></font>"#, None),
            (r#"<font><color theme="4"/></font>"#, Some(empty_theme.clone())),
            (r#"<font><scheme val="major"/></font>"#, Some(empty_theme.clone())),
        ] {
            let source = OfficeSource { version: 1, format: "xlsx".into(), original_id: "local-rule".into(), storage: "base".into(),
                rule_xml: format!(r#"<cfRule xmlns="{MAIN_NAMESPACE}" type="expression" priority="1" dxfId="0"><formula>A1&gt;0</formula></cfRule>"#),
                group_xml: Some(format!(r#"<conditionalFormatting xmlns="{MAIN_NAMESPACE}" sqref="A1:A4"/>"#)),
                differential_format_xml: Some(format!(r#"<dxf xmlns="{MAIN_NAMESPACE}">{dxf_body}</dxf>"#)), x14_xml: None, theme_xml,
                baseline: known_snapshot(&before), extensions: serde_json::Map::new() };
            let mut after = before.clone(); after["style"] = json!({ "bold": true }); after["officeSource"] = serde_json::to_value(source).unwrap();
            let sheet: SheetInput = serde_json::from_value(json!({ "id": "sheet", "name": "Data", "conditionalFormats": [after] })).unwrap();
            let mut registry = StyleRegistry::new();
            let error = write(&sheet, &mut registry, &mut Vec::new()).err().expect("declared theme resources must exist before replacement styles");
            assert_eq!(error.code, "INVALID_MODEL"); assert!(error.message.starts_with("CF_SOURCE_REQUIRED_RESOURCE"));
            assert!(registry.conditional_styles.is_empty());
        }
    }
}
