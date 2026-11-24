//! Native SDT property merging, row wrappers and authored-property provenance.
use std::collections::{BTreeMap, HashSet};

use office_io_common::{chart_xml, document_content_controls::{self as contract, ControlSourceEntry}};
use serde_json::Value;

use super::{CONTENT_NAMESPACES, ContentControl, WriteContext, content_control_xml, get, source_xml};
use crate::{ExportError, xml};

#[derive(Default)]
pub(super) struct ControlWriteState { entries: Vec<ControlSourceEntry>, bytes: usize }

fn managed(node: &source_xml::Element) -> bool {
    if node.namespace == source_xml::WORD_NAMESPACE {
        return ["id", "alias", "tag", "lock", "dataBinding", "text", "richText", "date", "dropDownList", "comboBox", "picture",
            "group", "docPartObj", "appearance", "color", "temporary", "showingPlcHdr", "placeholder"].contains(&node.local.as_str());
    }
    node.namespace == contract::WORD_2010_NAMESPACE && node.local == "checkbox"
        || node.namespace == contract::WORD_2012_NAMESPACE && ["repeatingSection", "repeatingSectionItem", "appearance", "color"].contains(&node.local.as_str())
}

fn preserved_root(raw: &str, tag: &str) -> Result<source_xml::Element, ExportError> {
    // Parser supplies namespace-complete properties. Locally authored historical
    // fragments can use the standard vocabulary provided by this wrapper.
    let wrapper = source_xml::parse_xml(&format!("<office {CONTENT_NAMESPACES}>{raw}</office>"), true).map_err(ExportError::invalid)?;
    if wrapper.children().count() != 1 { return Err(ExportError::invalid("DOCX SDT raw properties require one root")); }
    let mut root = wrapper.children().next().cloned().ok_or_else(|| ExportError::invalid("DOCX SDT raw properties have no root"))?;
    source_xml::canonicalize_office_namespaces(&mut root).map_err(ExportError::invalid)?;
    if !root.is_word(tag) { return Err(ExportError::invalid(format!("DOCX SDT raw property root must be {tag}"))); }
    // Root standard bindings must survive removal from the temporary wrapper.
    let standard = source_xml::parse_xml(&format!("<w:{tag} {CONTENT_NAMESPACES}/>"), true).map_err(ExportError::invalid)?;
    for attribute in standard.attributes.into_iter().filter(|attribute| attribute.namespace == source_xml::XMLNS_NAMESPACE) {
        if !root.attributes.iter().any(|current| current.name == attribute.name) { root.attributes.push(attribute); }
    }
    source_xml::validate_namespace_bindings(&root, &BTreeMap::new()).map_err(ExportError::invalid)?;
    Ok(root)
}

fn bind_generated_extensions(node: &mut source_xml::Element) {
    fn uses(node: &source_xml::Element, namespace: &str) -> bool {
        node.namespace == namespace || node.attributes.iter().any(|attribute| attribute.namespace == namespace)
            || node.children().any(|child| uses(child, namespace))
    }
    // The source may use w14/w15 for another extension. Bind each generated
    // property in its own scope, preserving those unknown source declarations.
    for (prefix, namespace) in [("w14", contract::WORD_2010_NAMESPACE), ("w15", contract::WORD_2012_NAMESPACE)] {
        if uses(node, namespace) {
            node.attributes.push(source_xml::Attribute { name: format!("xmlns:{prefix}"), namespace: source_xml::XMLNS_NAMESPACE.into(),
                local: prefix.into(), value: namespace.into() });
        }
    }
}

pub(super) fn merge_properties(control: &ContentControl, generated: &str) -> Result<(String, String), ExportError> {
    let mut generated = source_xml::parse_xml(&format!("<w:sdtPr {CONTENT_NAMESPACES}>{generated}</w:sdtPr>"), true).map_err(ExportError::invalid)?;
    for content in &mut generated.content {
        if let source_xml::Content::Element(node) = content { bind_generated_extensions(node); }
    }
    let mut properties = if let Some(raw) = get(&control.value, "rawSdtPr").as_str().filter(|raw| !raw.is_empty()) {
        preserved_root(raw, "sdtPr")?
    } else { source_xml::Element::word("sdtPr") };
    properties.content.retain(|node| !matches!(node, source_xml::Content::Element(node) if managed(node)));
    properties.content.extend(generated.content);
    properties.empty = properties.content.is_empty();
    // Newly generated native IDs always replace the source ID, including a
    // repeated item copied from a raw SDT. Unknown properties retain their names.
    let namespaces = BTreeMap::from([
        ("w".into(), source_xml::WORD_NAMESPACE.into()), ("w14".into(), contract::WORD_2010_NAMESPACE.into()), ("w15".into(), contract::WORD_2012_NAMESPACE.into()),
    ]);
    source_xml::validate_namespace_bindings(&properties, &namespaces).map_err(ExportError::invalid)?;
    let properties = properties.render(Some(&namespaces));
    chart_xml::parse(&properties).map_err(ExportError::invalid)?;
    let end_properties = get(&control.value, "rawSdtEndPr").as_str().filter(|raw| !raw.is_empty())
        .map(|raw| preserved_root(raw, "sdtEndPr").map(|root| root.render(None))).transpose()?.unwrap_or_default();
    Ok((properties, end_properties))
}

pub(super) fn wrap(context: &mut WriteContext<'_>, control: &ContentControl, content: &str) -> Result<String, ExportError> {
    let result = content_control_xml(control, content)?.ok_or_else(|| ExportError::invalid(format!("Unsupported native content-control type {:?}", control.kind)))?;
    record(context, control, result)
}

pub(super) fn record(context: &mut WriteContext<'_>, control: &ContentControl, result: String) -> Result<String, ExportError> {
    if get(&control.value, "tag").as_str().is_some_and(contract::internal_feature_tag) {
        if !context.used_controls.insert(control.id.clone()) { return Err(ExportError::invalid("A content control was written more than once")); }
        return Ok(result);
    }
    let native = chart_xml::parse(&result).map_err(ExportError::invalid)?;
    let mut properties = control.value.as_object().cloned().ok_or_else(|| ExportError::invalid("DOCX SDT properties must be an object"))?;
    for key in ["id", "blockId", "startOffset", "endOffset", "parentId", "tableRowRange", "blockRange", "value"] { properties.remove(key); }
    let entry = ControlSourceEntry { native_id: control.numeric_id.to_string(), id: control.id.clone(),
        properties_signature: contract::properties_signature(native.child("sdtPr"), native.child("sdtEndPr")).map_err(ExportError::invalid)?, properties: Value::Object(properties) };
    let bytes = serde_json::to_vec(&entry).map_err(|error| ExportError::invalid(error.to_string()))?.len();
    context.control_sources.bytes = context.control_sources.bytes.checked_add(bytes).filter(|bytes| *bytes <= contract::MAX_CONTROL_SOURCE_BYTES)
        .ok_or_else(|| ExportError::invalid("DOCX_SDT_RESOURCE_LIMIT: authored properties exceed 16 MiB"))?;
    if context.control_sources.entries.len() >= contract::MAX_CONTROL_COUNT { return Err(ExportError::invalid("DOCX_SDT_RESOURCE_LIMIT: too many content controls")); }
    if !context.used_controls.insert(control.id.clone()) { return Err(ExportError::invalid("A content control was written more than once")); }
    context.control_sources.entries.push(entry);
    Ok(result)
}

pub(super) fn row_range(control: &ContentControl, table_id: &str, rows: &[Value]) -> Result<(usize, usize), ExportError> {
    let range = get(&control.value, "tableRowRange");
    if get(range, "tableId").as_str() != Some(table_id) { return Err(ExportError::invalid("DOCX row SDT belongs to a different table")); }
    let ids = get(range, "rowIds").as_array().filter(|ids| !ids.is_empty()).ok_or_else(|| ExportError::invalid("DOCX row SDT requires stable physical row IDs"))?;
    let ids = ids.iter().map(|value| value.as_str().filter(|id| !id.trim().is_empty()).ok_or_else(|| ExportError::invalid("DOCX row SDT contains an invalid row identity"))).collect::<Result<Vec<_>, _>>()?;
    let start = rows.iter().position(|row| get(row, "id").as_str() == ids.first().copied()).ok_or_else(|| ExportError::invalid("DOCX row SDT first row does not exist"))?;
    if start + ids.len() > rows.len() || ids.iter().enumerate().any(|(offset, id)| get(&rows[start + offset], "id").as_str() != Some(*id)) {
        return Err(ExportError::invalid("DOCX row SDT identities must be ordered and contiguous"));
    }
    Ok((start, start + ids.len()))
}

/// Row contents are built by the ordinary table writer first. Wrapping those
/// exact tr strings preserves geometry, continuation paragraphs and inner tables.
pub(super) fn wrap_rows(context: &mut WriteContext<'_>, table_id: &str, rows: &[Value], contents: &[String]) -> Result<String, ExportError> {
    let controls = context.controls_by_block.get(table_id).cloned().unwrap_or_default();
    let controls = controls.into_iter().filter(|control| get(&control.value, "placement").as_str() == Some("row")).collect::<Vec<_>>();
    if controls.is_empty() { return Ok(contents.concat()); }
    let mut ranges = HashMapRanges::default();
    for control in &controls { ranges.values.insert(control.id.clone(), row_range(control, table_id, rows)?); }
    fn render(context: &mut WriteContext<'_>, controls: &[ContentControl], ranges: &HashMapRanges, contents: &[String], start: usize, end: usize,
        parent: Option<&str>, depth: usize, written: &mut HashSet<String>) -> Result<String, ExportError> {
        if depth > 64 { return Err(ExportError::invalid("DOCX row SDT nesting exceeds 64")); }
        let mut children = controls.iter().filter(|control| {
            let local_parent = get(&control.value, "parentId").as_str().filter(|id| controls.iter().any(|candidate| candidate.id == *id));
            local_parent == parent
        }).collect::<Vec<_>>();
        children.sort_by_key(|control| { let range = ranges.values[&control.id]; (range.0, std::cmp::Reverse(range.1), control.numeric_id) });
        let mut cursor = start; let mut result = String::new();
        for control in children {
            let (from, to) = ranges.values[&control.id];
            if from < cursor || from < start || to > end { return Err(ExportError::invalid("DOCX row SDT has crossing siblings or escapes its parent")); }
            result.push_str(&contents[cursor..from].concat());
            let inner = render(context, controls, ranges, contents, from, to, Some(&control.id), depth + 1, written)?;
            result.push_str(&wrap(context, control, &inner)?); written.insert(control.id.clone()); cursor = to;
        }
        result.push_str(&contents[cursor..end].concat()); Ok(result)
    }
    let mut written = HashSet::new();
    let result = render(context, &controls, &ranges, contents, 0, rows.len(), None, 0, &mut written)?;
    if written.len() != controls.len() { return Err(ExportError::invalid("DOCX row SDT has a missing or cyclic structural parent")); }
    Ok(result)
}

pub(super) fn block_range(control: &ContentControl, blocks: &[Value]) -> Result<(usize, usize), ExportError> {
    let declared = get(&control.value, "blockRange");
    let ids = if declared.is_null() {
        let id = get(&control.value, "blockId").as_str().ok_or_else(|| ExportError::invalid("DOCX block SDT requires an input block"))?;
        let block = blocks.iter().find(|block| get(block, "id").as_str() == Some(id)).ok_or_else(|| ExportError::invalid("DOCX block SDT input block is not in this container"))?;
        let length = get(block, "runs").as_array().into_iter().flatten().filter_map(|run| get(run, "text").as_str()).map(|text| text.encode_utf16().count()).sum::<usize>();
        if control.start != 0 || control.end != length { return Err(ExportError::invalid("DOCX block SDT must own the complete paragraph; use an inline SDT for a partial text range")); }
        vec![id]
    } else {
        let ids = get(declared, "blockIds").as_array().filter(|ids| !ids.is_empty()).ok_or_else(|| ExportError::invalid("DOCX block SDT requires a complete block ID sequence"))?;
        if ids.len() > 100_000 || declared.as_object().is_none_or(|range| range.keys().any(|key| key != "blockIds")) { return Err(ExportError::invalid("DOCX block SDT has an invalid range contract")); }
        ids.iter().map(|id| id.as_str().filter(|id| !id.trim().is_empty()).ok_or_else(|| ExportError::invalid("DOCX block SDT contains an invalid block ID"))).collect::<Result<Vec<_>, _>>()?
    };
    if ids.iter().copied().collect::<HashSet<_>>().len() != ids.len() { return Err(ExportError::invalid("DOCX block SDT block IDs are duplicated")); }
    let start = blocks.iter().position(|block| get(block, "id").as_str() == ids.first().copied()).ok_or_else(|| ExportError::invalid("DOCX block SDT first block is outside its container"))?;
    if start + ids.len() > blocks.len() || ids.iter().enumerate().any(|(offset, id)| get(&blocks[start + offset], "id").as_str() != Some(*id)) {
        return Err(ExportError::invalid("DOCX block SDT must contain ordered contiguous complete blocks"));
    }
    Ok((start, start + ids.len()))
}

/// The ordinary writers produce p/tbl nodes first; native block SDTs wrap those
/// complete nodes, preserving paragraph properties, inner rows and inline SDTs.
pub(super) fn wrap_blocks(context: &mut WriteContext<'_>, blocks: &[Value], contents: &[String]) -> Result<String, ExportError> {
    if blocks.len() != contents.len() { return Err(ExportError::invalid("DOCX block SDT contents do not match their physical blocks")); }
    let controls = blocks.iter().filter_map(|block| get(block, "id").as_str())
        .flat_map(|id| context.controls_by_block.get(id).into_iter().flatten())
        .filter(|control| get(&control.value, "placement").as_str() == Some("block")).cloned().collect::<Vec<_>>();
    if controls.is_empty() { return Ok(contents.concat()); }
    let mut ranges = HashMapRanges::default();
    for control in &controls { ranges.values.insert(control.id.clone(), block_range(control, blocks)?); }
    fn render(context: &mut WriteContext<'_>, controls: &[ContentControl], ranges: &HashMapRanges, contents: &[String], start: usize, end: usize,
        parent: Option<&str>, depth: usize, written: &mut HashSet<String>) -> Result<String, ExportError> {
        if depth > 64 { return Err(ExportError::invalid("DOCX block SDT nesting exceeds 64")); }
        let mut children = controls.iter().filter(|control| {
            let local_parent = get(&control.value, "parentId").as_str().filter(|id| controls.iter().any(|candidate| candidate.id == *id));
            local_parent == parent
        }).collect::<Vec<_>>();
        children.sort_by_key(|control| { let (from, to) = ranges.values[&control.id]; (from, std::cmp::Reverse(to), control.numeric_id) });
        let mut cursor = start; let mut result = String::new();
        for control in children {
            let (from, to) = ranges.values[&control.id];
            if from < cursor || from < start || to > end { return Err(ExportError::invalid("DOCX block SDT has crossing siblings or escapes its parent")); }
            result.push_str(&contents[cursor..from].concat());
            let inner = render(context, controls, ranges, contents, from, to, Some(&control.id), depth + 1, written)?;
            result.push_str(&wrap(context, control, &inner)?); written.insert(control.id.clone()); cursor = to;
        }
        result.push_str(&contents[cursor..end].concat()); Ok(result)
    }
    let mut written = HashSet::new();
    let result = render(context, &controls, &ranges, contents, 0, blocks.len(), None, 0, &mut written)?;
    if written.len() != controls.len() { return Err(ExportError::invalid("DOCX block SDT has a missing or cyclic structural parent")); }
    Ok(result)
}

#[derive(Default)]
struct HashMapRanges { values: std::collections::HashMap<String, (usize, usize)> }

pub(super) fn write_sources(context: &mut WriteContext<'_>) -> Result<(), ExportError> {
    if context.control_sources.entries.is_empty() { return Ok(()); }
    contract::validate_source_entries(&context.control_sources.entries).map_err(ExportError::invalid)?;
    let mut body = String::new();
    for entry in &context.control_sources.entries {
        body.push_str(&format!("<control>{}</control>", xml::escape(&serde_json::to_string(entry).map_err(|error| ExportError::invalid(error.to_string()))?)));
    }
    let xml = format!("<controls xmlns=\"{}\" version=\"1\">{body}</controls>", contract::CONTROL_SOURCE_NAMESPACE);
    if xml.len() > contract::MAX_CONTROL_SOURCE_BYTES { return Err(ExportError::invalid("DOCX_SDT_RESOURCE_LIMIT: source XML exceeds 16 MiB")); }
    context.package.add_xml(contract::CONTROL_SOURCE_PART, "application/xml", xml)?;
    context.relationship(office_io_common::document_custom_xml::CUSTOM_XML_RELATIONSHIP, contract::CONTROL_SOURCE_TARGET, false)?;
    Ok(())
}
