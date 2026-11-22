//! Standard SpreadsheetML rules with lossless, field-by-field Office provenance.
use super::{ConditionalFormatInput, ConditionalFormatStyleInput, RangeInput, SheetInput, StyleRegistry, WorkbookInput,
    cell_ref, conditional_formula, normalize_color, range_bounds, range_ref, write_conditional_rule};
use crate::{ExportError, ExportWarning};
use office_io_common::conditional_format_color::uses_theme;
use office_io_common::conditional_format_source::{CONFIG_FIELDS, OfficeSource, ProjectionHashes, STYLE_FIELDS, attach_projection,
    icon_count, icon_index, known_snapshot, office_icon_type, strip_projection, without_source};
use office_io_common::conditional_format_xml::{self as source_xml, Element, MAIN_NAMESPACE, RULE_LINK_EXTENSION_URI,
    STRICT_MAIN_NAMESPACE, WORKSHEET_EXTENSION_URI, X14_NAMESPACE, XM_NAMESPACE};
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
    let mut group = source.group_xml.as_deref().map(parsed).transpose()?;
    if let Some(group) = &mut group {
        if !is_visual(group, "conditionalFormatting") || (source.storage == "x14") != (group.namespace == X14_NAMESPACE) { return Err(invalid("invalid conditional-format source group")); }
        group.replace_children(is_rule, Vec::new()); strip_projection(group);
    }
    let dxf = source.differential_format_xml.as_deref().map(parsed).transpose()?;
    if dxf.as_ref().is_some_and(|node| !is_visual(node, "dxf")) { return Err(invalid("invalid conditional-format differential-style source")); }
    let mut x14 = source.x14_xml.as_deref().map(parsed).transpose()?;
    if let Some(group) = &mut x14 {
        if !group.is(X14_NAMESPACE, "conditionalFormatting") || group.children.iter().filter(|node| node.is(X14_NAMESPACE, "cfRule")).count() != 1 { return Err(invalid("invalid linked x14 conditional-format source")); }
        strip_projection(group);
    }
    let theme = source.theme_xml.as_deref().map(parsed).transpose()?;
    if theme.as_ref().is_some_and(|theme| theme.name != "theme" || !matches!(theme.namespace.as_str(),
        "http://schemas.openxmlformats.org/drawingml/2006/main" | "http://purl.oclc.org/ooxml/drawingml/main")) { return Err(invalid("invalid conditional-format source theme")); }
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
            let original = original.to_ascii_uppercase();
            if !copied && self.used.insert(original.clone()) { return Ok(original) }
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
        if parts.rule.as_ref().is_some_and(|rule| rule.attrs.contains_key("dxfId")) && dxf.is_none() && parts.rule.as_ref().unwrap().attrs.get("dxfId").is_some() {
            // An unresolved source index cannot safely be rebound to an unrelated new stylesheet.
            if source.and_then(|source| source.differential_format_xml.as_ref()).is_none() { return Err(invalid("conditional-format source differential style is unavailable")) }
        }
        let standalone = source.is_some_and(|source| source.storage == "x14") || rule.rule_type == "iconSet" && icon_requires_x14(rule);
        let dxf_id = if standalone { None } else { dxf.as_ref().map(|dxf| registry.register_conditional_style_xml(dxf)) };
        let mut projected = conditional_formula::project_rule(rule)?;
        let mut authored = if rule.rule_type == "office" {
            if source.is_none() || diff.semantic_change() { return Err(invalid("unknown Office conditional-format semantics must be retained")) }
            let original = parts.rule.take().ok_or_else(|| invalid("unknown Office conditional format has no XML source"))?;
            let (projected_rule, original, linked) = project_opaque(rule, original, parts.x14.take(), parts.group.as_ref(), &diff)?;
            projected = projected_rule; parts.x14 = linked; original
        } else {
            let mut generated = generated_rule(&projected, dxf_id, priority)?;
            if matches!(rule.rule_type.as_str(), "dataBar" | "colorScale" | "iconSet") {
                let visual = generated_visual(&projected, false, &diff, warnings)?; upsert(&mut generated, visual);
            }
            if standalone { to_x14(&mut generated); }
            merge_rule(parts.rule.take(), generated, &diff, &projected, warnings)?
        };
        if standalone {
            if !authored.is(X14_NAMESPACE, "cfRule") { to_x14(&mut authored); }
            let original = authored.attrs.get("id").cloned(); let guid = guids.assign(original.as_deref(), copied, &format!("{key}:x14"))?;
            authored.set_attr("id", guid); authored.remove_attr("dxfId");
            if rule.rule_type == "office" && source.is_some_and(|source| source.storage == "x14") && !authored.attrs.contains_key("priority") { authored.remove_attr("priority"); }
            else { authored.set_attr("priority", priority.to_string()); }
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
            } else if let Some(mut linked) = parts.x14.take() {
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

fn generated_rule(rule: &ConditionalFormatInput, dxf_id: Option<usize>, priority: u32) -> Result<Element, ExportError> {
    let mut xml = String::new(); write_conditional_rule(&mut xml, rule, dxf_id, priority, &rule.operator)?;
    let root = parsed(&format!("<worksheet xmlns=\"{MAIN_NAMESPACE}\">{xml}</worksheet>"))?;
    root.children.into_iter().find(|node| node.is_main("cfRule")).ok_or_else(|| invalid("conditional-format writer emitted no rule"))
}
fn make_group(source: Option<Element>, x14: bool, ranges: &[RangeInput]) -> Result<Element, ExportError> {
    let mut group = source.unwrap_or_else(|| if x14 { Element::prefixed(X14_NAMESPACE, "x14", "conditionalFormatting") } else { Element::new(MAIN_NAMESPACE, "conditionalFormatting") });
    if x14 { rename(&mut group, X14_NAMESPACE, "x14", "conditionalFormatting"); group.set_attr("xmlns:xm", XM_NAMESPACE); }
    else { spreadsheet(&mut group); }
    group.replace_children(is_rule, Vec::new());
    let references = ranges.iter().map(|range| range_ref(range).ok_or_else(|| invalid("invalid conditional-format range"))).collect::<Result<Vec<_>, _>>()?.join(" ");
    if x14 { let mut sqref = group.children.iter().find(|node| node.is(XM_NAMESPACE, "sqref")).cloned().unwrap_or_else(|| Element::prefixed(XM_NAMESPACE, "xm", "sqref")); sqref.set_text(references); upsert(&mut group, sqref); }
    else { group.set_attr("sqref", references); }
    Ok(group)
}
fn checked_theme<'a>(rule: &Element, dxf: Option<&Element>, x14: Option<&Element>, theme: Option<&'a Element>) -> Result<Option<&'a Element>, ExportError> {
    if uses_theme(rule) || dxf.is_some_and(uses_theme) || x14.is_some_and(uses_theme) { theme.map(Some).ok_or_else(|| invalid("conditional-format theme references have no preserved source theme")) }
    else { Ok(None) }
}
pub(super) fn theme_for_workbook(workbook: &WorkbookInput) -> Result<Option<String>, ExportError> {
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
