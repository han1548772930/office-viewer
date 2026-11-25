//! Note authoring, native references and special stories share the document's
//! resource/field/control writer instead of flattening them into label text.
use std::collections::BTreeSet;

use office_io_common::document_note_source::{self, DocumentNoteKind, DocumentNoteProperties, NOTE_SAFE_INTEGER,
    NOTE_EXTENSION_NAMESPACE, WORD_2013_NAMESPACE};
use serde_json::{Value, json};

use super::{CONTENT_NAMESPACES, NoteBundle, NoteEntry, NoteKind, WriteContext, block_hosts, control_writer, get,
    note_text, source_xml::{self, Attribute, Content, Element}, xml, xml_text};
use crate::ExportError;

const DECORATIONS: [&str; 3] = ["separator", "continuationSeparator", "continuationNotice"];
const MC_NAMESPACE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

#[path = "docx_note_decorations.rs"]
mod decoration_writer;
pub(super) use decoration_writer::{decoration_contents, reserved_ids};
#[path = "docx_note_markers.rs"]
mod marker_writer;
pub(super) use marker_writer::{ReferenceProperties, reference_run};
use marker_writer::body_marker;
#[path = "docx_note_positions.rs"]
mod positions;
pub(super) use positions::{EventScope, NoteBoundary, NotePositions, optional_reference_count};

fn native_kind(kind: NoteKind) -> DocumentNoteKind {
    if kind == NoteKind::Footnote { DocumentNoteKind::Footnote } else { DocumentNoteKind::Endnote }
}

pub(super) fn allocate_id(next: &mut i64, used: &mut BTreeSet<i64>) -> Result<i64, ExportError> {
    while used.contains(next) { *next = (*next).checked_add(1).ok_or_else(|| ExportError::invalid("注释标识超出容量。"))?; }
    if *next <= 0 || (*next).unsigned_abs() > NOTE_SAFE_INTEGER { return Err(ExportError::invalid("注释标识超出安全整数范围。")); }
    let id = *next;
    used.insert(id);
    *next = id.checked_add(1).ok_or_else(|| ExportError::invalid("注释标识超出容量。"))?;
    Ok(id)
}

fn append_marker(output: &mut String, marker: &str) -> Result<bool, ExportError> {
    if marker.is_empty() { return Ok(true); }
    if !output.starts_with("<w:p>") { return Ok(false); }
    let position = if output.starts_with("<w:p><w:pPr") {
        let opening = output["<w:p>".len()..].find('>').map(|index| index + "<w:p>".len())
            .ok_or_else(|| ExportError::invalid("注释首段的段落属性开头无效。"))?;
        if output.as_bytes().get(opening.saturating_sub(1)) == Some(&b'/') { opening + 1 }
        else { output.find("</w:pPr>").ok_or_else(|| ExportError::invalid("注释首段缺少段落属性结束标记。"))? + "</w:pPr>".len() }
    } else { "<w:p>".len() };
    output.insert_str(position, marker); Ok(true)
}

fn block_body(context: &mut WriteContext<'_>, blocks: &[Value], marker: &str) -> Result<String, ExportError> {
    let mut prefix = String::new();
    let mut contents = Vec::with_capacity(blocks.len());
    for (index, block) in blocks.iter().enumerate() {
        let mut output = context.block(block)?;
        if index == 0 && !append_marker(&mut output, marker)? { prefix.push_str(&format!("<w:p>{marker}</w:p>")); }
        contents.push(output);
    }
    prefix.push_str(&control_writer::wrap_blocks(context, blocks, &contents)?);
    if blocks.is_empty() { prefix.push_str(&format!("<w:p>{marker}</w:p>")); }
    Ok(prefix)
}

fn legacy_body(context: &mut WriteContext<'_>, value: &Value, kind: NoteKind, marker: &str) -> Result<String, ExportError> {
    let style = kind.style();
    let runs = [get(value, "content"), value, get(value, "body")].into_iter().find_map(|content| get(content, "runs").as_array());
    if let Some(runs) = runs.filter(|runs| !runs.is_empty()) {
        let text = context.runs(runs)?;
        return Ok(format!("<w:p><w:pPr><w:pStyle w:val=\"{style}\"/></w:pPr>{marker}{text}</w:p>"));
    }
    let text = note_text(value).unwrap_or_default();
    let mut result = String::new();
    for (index, line) in text.replace("\r\n", "\n").replace('\r', "\n").split('\n').enumerate() {
        let run = context.text_run_xml(&json!({}), line)?;
        result.push_str(&format!("<w:p><w:pPr><w:pStyle w:val=\"{style}\"/></w:pPr>{}{run}</w:p>", if index == 0 { marker } else { "" }));
    }
    Ok(result)
}

fn note_element(source: &Value, kind: NoteKind, id: i64, note_type: Option<&str>, public_id: Option<&str>, body: &str) -> Result<String, ExportError> {
    let opening = source.as_str().map(|xml| source_xml::parse_xml(xml, false)).transpose().map_err(ExportError::invalid)?;
    let mut element = opening.unwrap_or_else(|| Element::word(native_kind(kind).name()));
    source_xml::canonicalize_office_namespaces(&mut element).map_err(ExportError::invalid)?;
    if !element.is_word(native_kind(kind).name()) { return Err(ExportError::invalid("注释来源的元素类型无效。")); }
    element.set_word_attribute("id", id.to_string());
    if let Some(note_type) = note_type { element.set_word_attribute("type", note_type); }
    else { element.attributes.retain(|attribute| attribute.namespace != source_xml::WORD_NAMESPACE || attribute.local != "type"); }
    if let Some(public_id) = public_id {
        extension_attribute(&mut element, "sourceIdHex", &document_note_source::encode_public_note_source_id(public_id));
    }
    Ok(format!("{}{body}</{}>", element.opening(None, false), element.name))
}

pub(super) fn notes_xml(context: &mut WriteContext<'_>, kind: NoteKind, notes: &[NoteEntry], settings: &Value) -> Result<String, ExportError> {
    let path = format!("word/{}.xml", kind.collection());
    context.with_relationship_owner(&path, |context| {
        let mut body = String::new();
        let empty_content = Value::Null;
        for decoration in decoration_writer::decorations(settings, kind)? {
            let content = decoration.content.unwrap_or(&empty_content);
            let source = get(get(content, "metadata"), "decorationSource");
            let marker = if decoration.content.is_none() || get(source, "marker").as_str().is_some_and(|marker| marker == "separator" || marker == "continuationSeparator") {
                if decoration.kind == "continuationNotice" { String::new() } else { format!("<w:r><w:{}/></w:r>", decoration.kind) }
            } else { String::new() };
            let paragraphs = if let Some(blocks) = block_hosts::content_blocks(content)? { block_body(context, blocks, &marker)? }
                else { format!("<w:p>{marker}</w:p>") };
            body.push_str(&note_element(get(source, "sourceXml"), kind, decoration.id, Some(decoration.kind), decoration.public_id, &paragraphs)?);
        }
        for note in notes {
            let marker = body_marker(context, &note.value, kind)?;
            let paragraphs = if let Some(blocks) = block_hosts::content_blocks(get(&note.value, "content"))? { block_body(context, blocks, &marker)? }
                else { legacy_body(context, &note.value, kind, &marker)? };
            body.push_str(&note_element(get(get(get(&note.value, "metadata"), "officeNoteSource"), "openingXml"), kind, note.id, None, None, &paragraphs)?);
        }
        Ok(format!("{}<w:{} {CONTENT_NAMESPACES}>{body}</w:{}>", xml::DECL, kind.collection(), kind.collection()))
    })
}

fn namespace_prefix(element: &mut Element, preferred: &str, namespace: &str) -> String {
    if let Some(existing) = element.attributes.iter().find(|attribute| attribute.namespace == source_xml::XMLNS_NAMESPACE
        && attribute.name.starts_with("xmlns:") && attribute.value == namespace) { return existing.local.clone(); }
    let mut index = 0;
    let prefix = loop {
        let candidate = if index == 0 { preferred.to_owned() } else { format!("{preferred}{index}") };
        if !element.attributes.iter().any(|attribute| attribute.name == format!("xmlns:{candidate}")) { break candidate; }
        index += 1;
    };
    element.attributes.push(Attribute { name: format!("xmlns:{prefix}"), namespace: source_xml::XMLNS_NAMESPACE.into(), local: prefix.clone(), value: namespace.into() });
    prefix
}

fn extension_attribute(element: &mut Element, local: &str, value: &str) {
    let prefix = namespace_prefix(element, "officeNote", NOTE_EXTENSION_NAMESPACE);
    element.attributes.retain(|attribute| attribute.namespace != NOTE_EXTENSION_NAMESPACE || attribute.local != local);
    element.attributes.push(Attribute { name: format!("{prefix}:{local}"), namespace: NOTE_EXTENSION_NAMESPACE.into(), local: local.into(), value: value.into() });
    let mc_prefix = namespace_prefix(element, "officeNoteMc", MC_NAMESPACE);
    if let Some(ignorable) = element.attributes.iter_mut().find(|attribute| attribute.namespace == MC_NAMESPACE && attribute.local == "Ignorable") {
        if !ignorable.value.split_whitespace().any(|candidate| candidate == prefix) { ignorable.value.push(' '); ignorable.value.push_str(&prefix); }
    } else { element.attributes.push(Attribute { name: format!("{mc_prefix}:Ignorable"), namespace: MC_NAMESPACE.into(), local: "Ignorable".into(), value: prefix }); }
}

fn update_property(root: &mut Element, tag: &str, value: Option<String>, selector: &str) {
    let mut element = root.child(tag).cloned().unwrap_or_else(|| Element::word(tag));
    root.content.retain(|content| !matches!(content, Content::Element(element) if element.is_word(tag)));
    let Some(value) = value else { return; };
    element.set_word_attribute("val", value);
    let rank = |name: &str| ["pos", "numFmt", "numStart", "numRestart", selector].iter().position(|candidate| *candidate == name);
    let target = rank(tag).unwrap_or(0);
    let at = root.content.iter().position(|content| matches!(content, Content::Element(element) if element.namespace == source_xml::WORD_NAMESPACE
        && rank(&element.local).is_some_and(|rank| rank > target))).unwrap_or(root.content.len());
    root.content.insert(at, Content::Element(element));
}

fn property_xml(settings: &Value, kind: NoteKind, special_ids: Option<&[i64]>) -> Result<String, ExportError> {
    if settings.is_null() && special_ids.is_none() { return Ok(String::new()); }
    if !settings.is_null() && !settings.is_object() { return Err(ExportError::invalid("注释设置必须为对象。")); }
    if settings.as_object().is_some_and(|settings| settings.keys().any(|key| !["position", "restart", "startNumber", "numberFormat", "columnCount", "metadata", "separator", "continuationSeparator", "continuationNotice"].contains(&key.as_str()))) {
        return Err(ExportError::invalid("注释设置包含未定义的字段。"));
    }
    let properties: DocumentNoteProperties = serde_json::from_value(if settings.is_null() { json!({}) } else { settings.clone() }).map_err(|error| ExportError::invalid(format!("注释属性无效：{error}")))?;
    document_note_source::validate_properties(&properties, native_kind(kind)).map_err(ExportError::invalid)?;
    let source = get(get(settings, "metadata"), "officeNoteProperties");
    let mut root = if source.is_null() { Element::word(native_kind(kind).properties_tag()) } else {
        if get(source, "version").as_u64() != Some(1) || get(source, "kind").as_str() != Some(native_kind(kind).name()) { return Err(ExportError::invalid("注释属性来源版本或类型无效。")); }
        source_xml::parse_xml(get(source, "xml").as_str().ok_or_else(|| ExportError::invalid("注释属性来源 XML 缺失。"))?, false).map_err(ExportError::invalid)?
    };
    source_xml::canonicalize_office_namespaces(&mut root).map_err(ExportError::invalid)?;
    if !root.is_word(native_kind(kind).properties_tag()) { return Err(ExportError::invalid("注释属性来源根节点无效。")); }
    for (key, tag) in [("position", "pos"), ("numberFormat", "numFmt"), ("startNumber", "numStart"), ("restart", "numRestart")] {
        let author = get(settings, key);
        if !author.is_null() || source.is_null() || author != get(get(source, "baseline"), key) {
            update_property(&mut root, tag, (!author.is_null()).then(|| author.as_str().map(str::to_owned).unwrap_or_else(|| author.to_string())), native_kind(kind).name());
        }
    }
    if let Some(ids) = special_ids {
        root.attributes.retain(|attribute| attribute.namespace != NOTE_EXTENSION_NAMESPACE || attribute.local != "columnCount");
        if let Some(count) = properties.column_count { extension_attribute(&mut root, "columnCount", &count.to_string()); }
        root.content.retain(|content| !matches!(content, Content::Element(element) if element.is_word(native_kind(kind).name())));
        for id in ids { let mut element = Element::word(native_kind(kind).name()); element.set_word_attribute("id", id.to_string()); root.content.push(Content::Element(element)); }
    } else if DECORATIONS.iter().any(|key| settings.get(key).is_some()) { return Err(ExportError::invalid("分隔内容只能存放在文档默认注释设置中。")); }
    Ok(root.render(None))
}

pub(super) fn document_properties(settings: &Value, notes: &NoteBundle) -> Result<String, ExportError> {
    let mut result = String::new();
    for (kind, entries) in [(NoteKind::Footnote, &notes.footnotes), (NoteKind::Endnote, &notes.endnotes)] {
        let value = get(settings, native_kind(kind).name());
        if value.is_null() && entries.is_empty() { continue; }
        let active_ids = decoration_writer::decorations(value, kind)?.into_iter().filter(|value| value.active).map(|value| value.id).collect::<Vec<_>>();
        result.push_str(&property_xml(value, kind, Some(&active_ids))?);
    }
    Ok(result)
}

pub(super) fn section_properties(setup: &Value, document_settings: &Value) -> Result<String, ExportError> {
    let mut result = String::new();
    let notes = get(setup, "noteProperties");
    if !notes.is_null() && (notes.as_object().is_none() || notes.as_object().is_some_and(|notes| notes.keys().any(|key| key != "footnote" && key != "endnote"))) {
        return Err(ExportError::invalid("节注释设置类型无效。"));
    }
    for kind in [NoteKind::Footnote, NoteKind::Endnote] { result.push_str(&property_xml(get(notes, native_kind(kind).name()), kind, None)?); }
    let footnotes = get(notes, "footnote");
    let own_count = get(footnotes, "columnCount").as_u64();
    let inherited_count = get(get(document_settings, "footnote"), "columnCount").as_u64();
    if let Some(count) = own_count.or(inherited_count) {
        if count > 4 { return Err(ExportError::invalid("脚注列数需要在 0 到 4 之间。")); }
        let provenance = get(get(footnotes, "metadata"), "officeNoteColumnsSource");
        if !provenance.is_null() && get(provenance, "version").as_u64() != Some(1) { return Err(ExportError::invalid("脚注列数来源版本无效。")); }
        let mut element = if let Some(xml) = get(provenance, "xml").as_str() { source_xml::parse_xml(xml, false).map_err(ExportError::invalid)? }
            else { Element { name: "w15:footnoteColumns".into(), namespace: WORD_2013_NAMESPACE.into(), local: "footnoteColumns".into(), attributes: Vec::new(), content: Vec::new(), empty: true } };
        if element.namespace != WORD_2013_NAMESPACE || element.local != "footnoteColumns" { return Err(ExportError::invalid("脚注列数来源根节点无效。")); }
        element.attributes.retain(|attribute| !(attribute.namespace == WORD_2013_NAMESPACE && attribute.local == "val") && !(attribute.namespace == NOTE_EXTENSION_NAMESPACE && attribute.local == "inheritedColumnsFromDocument"));
        let prefix = namespace_prefix(&mut element, "w15", WORD_2013_NAMESPACE);
        element.name = format!("{prefix}:footnoteColumns");
        element.attributes.push(Attribute { name: format!("{prefix}:val"), namespace: WORD_2013_NAMESPACE.into(), local: "val".into(), value: count.to_string() });
        if own_count.is_none() { extension_attribute(&mut element, "inheritedColumnsFromDocument", "1"); }
        result.push_str(&element.render(None));
    }
    if let Some(value) = setup.get("suppressEndnotes") {
        let enabled = value.as_bool().ok_or_else(|| ExportError::invalid("节尾注延后标志必须是布尔值。"))?;
        result.push_str(if enabled { "<w:suppressEndnotes/>" } else { "<w:suppressEndnotes w:val=\"0\"/>" });
    }
    Ok(result)
}
