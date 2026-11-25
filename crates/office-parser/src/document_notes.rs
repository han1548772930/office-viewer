//! Restore note properties and special note stories through the same Duke
//! extraction path as ordinary notes, retaining their actual OPC owner.
use std::{cell::RefCell, collections::{BTreeMap, HashMap}};

use duke_pages::{DocModel, DocumentNoteDefinition, OoxmlPackage, TextStyle,
    parse::{ParseContext, ParsedStyleSheet, parse_content_types, parse_document_xml, relationship_part_name_for_part, resolve_part_path}};
use office_io_common::{chart_xml, document_content_controls::{word, word_attribute},
    document_note_source::{self, DocumentNoteKind, DocumentNoteProperties, DocumentSectionNoteProperties, NOTE_EXTENSION_NAMESPACE, WORD_2013_NAMESPACE}};
use quick_xml::{Reader, events::Event};
use serde_json::{Value, json};

use crate::{document_format_sidecar::parse_xml, document_native_xml::{complete_fragment, scope},
    model::{DocumentNoteDecorationPayload, DocumentNotePayload, DocumentNoteSettingsKindPayload, DocumentNoteSettingsPayload}};

#[path = "document_note_references.rs"]
mod references;

pub(crate) struct NativeNoteDefinition {
    pub note_type: String,
    pub part_name: String,
    pub opening_xml: String,
    pub show_body_reference: bool,
    pub custom_mark: Option<String>,
    pub reference_style: Option<TextStyle>,
    pub body_marker_source: Option<Value>,
    pub body_marker_layout: Option<Value>,
}

#[derive(Default)]
pub(crate) struct NativeNotes {
    pub definitions: BTreeMap<(DocumentNoteKind, i64), NativeNoteDefinition>,
    pub settings: DocumentNoteSettingsPayload,
    pub warnings: Vec<String>,
    pub namespaces: HashMap<String, BTreeMap<String, String>>,
}

impl NativeNotes {
    pub fn is_decoration(&self, kind: DocumentNoteKind, id: i64) -> bool {
        self.definitions.get(&(kind, id)).is_some_and(|note| note.note_type != "normal")
    }

    pub fn payload(&self, kind: DocumentNoteKind, note: &DocumentNoteDefinition) -> DocumentNotePayload {
        let native = self.definitions.get(&(kind, note.id));
        let native_id = note.id;
        DocumentNotePayload { id: note.id.to_string(), text: note.text.clone(),
            custom_mark: native.and_then(|note| note.custom_mark.clone()),
            reference_text_style: native.and_then(|note| note.reference_style.clone()),
            show_body_reference: native.map(|note| note.show_body_reference),
            metadata: native.map(|note| {
                let mut source = json!({ "version": 1, "partName": note.part_name, "nativeId": native_id,
                    "noteType": note.note_type, "openingXml": note.opening_xml });
                if let Some(style) = &note.reference_style { source["referenceTextStyleBaseline"] = json!(style); }
                if let Some(marker) = &note.body_marker_source { source["bodyMarkerSource"] = marker.clone(); }
                if let Some(layout) = &note.body_marker_layout { source["bodyMarkerLayout"] = layout.clone(); }
                json!({ "officeNoteSource": source })
            }),
        }
    }

    pub fn restore_references(&mut self, document: &mut duke_pages::StructuredDocument, source: &DocModel, styles: &ParsedStyleSheet) -> Result<(), String> {
        references::restore(self, document, source, styles)
    }

    pub fn into_settings(self) -> Option<DocumentNoteSettingsPayload> {
        (self.settings.footnote.is_some() || self.settings.endnote.is_some()).then_some(self.settings)
    }
}

fn settings_slot(settings: &mut DocumentNoteSettingsPayload, kind: DocumentNoteKind) -> &mut Option<DocumentNoteSettingsKindPayload> {
    if kind == DocumentNoteKind::Footnote { &mut settings.footnote } else { &mut settings.endnote }
}

fn opening_fragment(xml: &str) -> Result<String, String> {
    let mut reader = Reader::from_str(xml);
    loop {
        match reader.read_event().map_err(|error| error.to_string())? {
            Event::Start(_) => {
                let end = reader.buffer_position() as usize;
                return Ok(format!("{}/>", xml[..end].trim_end_matches('>')));
            }
            Event::Empty(_) => return Ok(xml[..reader.buffer_position() as usize].to_owned()),
            Event::Eof => return Err("Native note source has no opening element".into()),
            _ => {},
        }
    }
}

/// Read the native and converted-DOC packages already owned by the parser.
/// This does not reparse or clone the whole Office file.
pub(crate) fn read(package: Option<&OoxmlPackage>, source: &mut DocModel, styles: &ParsedStyleSheet) -> Result<NativeNotes, String> {
    let mut result = NativeNotes::default();
    let Some(package) = package else { return Ok(result); };
    for story in crate::document_story_source::stories(source) {
        if let Some(part) = package.parts.get(story.part) {
            result.namespaces.insert(story.part.into(), crate::document_native_xml::opening_namespaces(&part.content)?);
        }
    }
    let settings_part = duke_pages::related_document_part_name(package, "settings");
    let mut selected_decorations: BTreeMap<DocumentNoteKind, Vec<i64>> = BTreeMap::new();
    if let Some(part) = package.parts.get(&settings_part) {
        let raw = parse_xml(&part.content)?;
        let root = chart_xml::parse(&part.content)?;
        if !word(&root, "settings") { return Err("DOCX_NOTE_SETTINGS_INVALID: settings part has the wrong root".into()); }
        let inherited = scope(&raw, &BTreeMap::new());
        result.namespaces.insert(settings_part.clone(), inherited.clone());
        for kind in [DocumentNoteKind::Footnote, DocumentNoteKind::Endnote] {
            let mut properties = raw.children.iter().zip(&root.children).filter(|(_, node)| word(node, kind.properties_tag()));
            if let Some((raw, node)) = properties.next() {
                if properties.next().is_some() { return Err("DOCX_NOTE_SETTINGS_INVALID: duplicate document note properties".into()); }
                let xml = complete_fragment(&part.content, raw, &scope(raw, &inherited))?;
                let mut parsed = document_note_source::read_properties(&xml, kind, &settings_part)?;
                result.warnings.append(&mut parsed.warnings);
                settings_slot(&mut result.settings, kind).get_or_insert_with(Default::default).properties = parsed.properties;
                let mut ids = Vec::new();
                for reference in node.children.iter().filter(|node| word(node, kind.name())) {
                    let id = word_attribute(reference, "id").and_then(|value| value.parse::<i64>().ok())
                        .filter(|value| value.unsigned_abs() <= document_note_source::NOTE_SAFE_INTEGER)
                        .ok_or("DOCX_NOTE_SETTINGS_INVALID: special note selector has an invalid ID")?;
                    if ids.contains(&id) { return Err("DOCX_NOTE_SETTINGS_INVALID: special note selector is duplicated".into()); }
                    ids.push(id);
                }
                selected_decorations.insert(kind, ids);
            }
        }
    }
    let content_types = parse_content_types(package);
    for kind in [DocumentNoteKind::Footnote, DocumentNoteKind::Endnote] {
        let part_name = duke_pages::related_document_part_name(package, kind.collection());
        let Some(part) = package.parts.get(&part_name) else { continue; };
        let raw = parse_xml(&part.content)?;
        let root = chart_xml::parse(&part.content)?;
        if !word(&root, kind.collection()) { return Err("DOCX_NOTE_INVALID: note part has the wrong namespace or root".into()); }
        let inherited = scope(&raw, &BTreeMap::new());
        result.namespaces.insert(part_name.clone(), inherited.clone());
        let relationships = package.parts.get(&relationship_part_name_for_part(&part_name))
            .map(|part| office_io_common::document_custom_xml::relationships(&part.content)).transpose()?
            .into_iter().flatten().map(|relationship| {
                let target = if relationship.external { relationship.target } else { resolve_part_path(&part_name, &relationship.target) };
                (relationship.id, target)
            }).collect::<HashMap<_, _>>();
        let context = ParseContext { relationships, content_types: content_types.clone(), parts: &package.parts,
            binary_assets: &package.binary_assets, style_sheet: styles, warnings: RefCell::new(Vec::new()) };
        let notes = if kind == DocumentNoteKind::Footnote { &mut source.metadata.footnotes } else { &mut source.metadata.endnotes };
        let mut decorations = Vec::new();
        let mut public_ids = BTreeMap::new();
        for (raw, node) in raw.children.iter().zip(&root.children).filter(|(_, node)| word(node, kind.name())) {
            let id = word_attribute(node, "id").and_then(|value| value.parse::<i64>().ok())
                .filter(|value| value.unsigned_abs() <= document_note_source::NOTE_SAFE_INTEGER)
                .ok_or("DOCX_NOTE_INVALID: note ID must be an integer within the JavaScript safe range")?;
            let note_type = word_attribute(node, "type").unwrap_or("normal");
            if !["normal", "separator", "continuationSeparator", "continuationNotice"].contains(&note_type) {
                return Err(format!("DOCX_NOTE_INVALID: unrecognized note type {note_type:?}"));
            }
            let xml = complete_fragment(&part.content, raw, &scope(raw, &inherited))?;
            let native = NativeNoteDefinition { note_type: note_type.into(), part_name: part_name.clone(), opening_xml: opening_fragment(&xml)?,
                show_body_reference: node.descendants(kind.body_reference_tag()).into_iter().any(|node| word(node, kind.body_reference_tag())),
                custom_mark: None, reference_style: None, body_marker_source: None, body_marker_layout: None };
            if result.definitions.insert((kind, id), native).is_some() { return Err("DOCX_NOTE_INVALID: duplicate note ID in its owning part".into()); }
            if note_type == "normal" { continue; }
            let public_source = node.attribute_ns(NOTE_EXTENSION_NAMESPACE, "sourceIdHex")
                .map(document_note_source::decode_public_note_source_id).transpose()?;
            if let Some(public_id) = &public_source {
                if public_ids.insert(public_id.clone(), id).is_some() {
                    return Err("DOCX_NOTE_SOURCE_INVALID: special definitions share a public source identity".into());
                }
            }
            let canonical = crate::document_format_sidecar::canonical_resource(&xml, &format!("w:{}", kind.name()))?;
            let nodes = parse_document_xml(&canonical, &context);
            let text = native_node_text(&nodes);
            if notes.as_ref().is_some_and(|notes| notes.iter().any(|note| note.id == id)) {
                return Err("DOCX_NOTE_INVALID: special note ID conflicts with a normal note".into());
            }
            notes.get_or_insert_with(Vec::new).push(DocumentNoteDefinition { id, text: text.clone(), nodes: Some(nodes) });
            let marker = ["separator", "continuationSeparator"].into_iter().find(|name| node.descendants(name).into_iter().any(|node| word(node, name)));
            decorations.push((id, note_type.to_owned(), DocumentNoteDecorationPayload { text, decoration_source: json!({ "version": 1, "kind": kind,
                "type": note_type, "noteId": id, "partName": part_name, "sourceXml": xml, "marker": marker }),
                univer_note_decoration_source: public_source.map(|public_id| json!({ "version": 1, "kind": kind, "type": note_type, "noteId": public_id })) }));
        }
        let selectors = selected_decorations.get(&kind);
        let mut active = BTreeMap::new();
        for note_type in ["separator", "continuationSeparator", "continuationNotice"] {
            let candidates = decorations.iter().filter(|(_, candidate_type, _)| candidate_type == note_type);
            let selected = if let Some(ids) = selectors {
                ids.iter().find_map(|id| candidates.clone().find(|(candidate, _, _)| candidate == id).map(|(id, _, _)| *id))
            } else {
                let default_id = if note_type == "separator" { -1 } else if note_type == "continuationSeparator" { 0 } else { 1 };
                candidates.clone().find(|(id, _, _)| *id == default_id).or_else(|| candidates.clone().next()).map(|(id, _, _)| *id)
            };
            if let Some(id) = selected { active.insert(note_type, id); }
        }
        if let Some(ids) = selectors { for id in ids {
            if !decorations.iter().any(|(candidate, _, _)| candidate == id) {
                result.warnings.push(format!("Native {} special-note selector {id} has no special definition; its XML remains preserved.", kind.name()));
            }
        } }
        if !decorations.is_empty() {
            let slot = settings_slot(&mut result.settings, kind).get_or_insert_with(Default::default);
            for (id, note_type, decoration) in decorations {
                if active.get(note_type.as_str()) == Some(&id) {
                    match note_type.as_str() {
                        "separator" => slot.separator = Some(decoration),
                        "continuationSeparator" => slot.continuation_separator = Some(decoration),
                        _ => slot.continuation_notice = Some(decoration),
                    }
                } else { slot.dormant_decorations.push(decoration); }
            }
        }
        result.warnings.extend(context.warnings.into_inner());
    }
    // Newly admitted special stories need the exact same IDs as all existing
    // body/table/story nodes. The document identity and content stay unchanged.
    let document_id = source.metadata.document_id.clone().unwrap_or_else(|| "unsaved".into());
    duke_pages::structured::assign_source_ids(source, &document_id);
    Ok(result)
}

fn native_node_text(nodes: &[duke_pages::DocNode]) -> String {
    fn paragraph(node: &duke_pages::ParagraphNode) -> String {
        node.children.iter().filter_map(|child| match child {
            duke_pages::ParagraphChildNode::Text(run) => Some(run.text.as_str()),
            _ => None,
        }).collect()
    }
    fn table(node: &duke_pages::TableNode) -> String {
        node.rows.iter().flat_map(|row| &row.cells).flat_map(|cell| &cell.nodes).map(|node| match node {
            duke_pages::TableCellContentNode::Paragraph(node) => paragraph(node),
            duke_pages::TableCellContentNode::Table(node) => table(node),
        }).collect::<Vec<_>>().join("\n")
    }
    nodes.iter().map(|node| match node {
        duke_pages::DocNode::Paragraph(node) => paragraph(node),
        duke_pages::DocNode::Table(node) => table(node),
    }).collect::<Vec<_>>().join("\n")
}

pub(crate) fn read_section(xml: Option<&str>, inherited: &BTreeMap<String, String>, part_name: &str, document_columns: Option<u8>) -> Result<(Option<DocumentSectionNoteProperties>, Option<bool>, Vec<String>), String> {
    let Some(xml) = xml else { return Ok((None, None, Vec::new())); };
    let raw = parse_xml(xml)?;
    let xml = complete_fragment(xml, &raw, &scope(&raw, inherited))?;
    let raw = parse_xml(&xml)?;
    let root = chart_xml::parse(&xml)?;
    if !word(&root, "sectPr") { return Err("DOCX_NOTE_SETTINGS_INVALID: section has the wrong root".into()); }
    let inherited = scope(&raw, &BTreeMap::new());
    let mut properties = DocumentSectionNoteProperties::default();
    let mut warnings = Vec::new();
    for kind in [DocumentNoteKind::Footnote, DocumentNoteKind::Endnote] {
        let mut nodes = raw.children.iter().zip(&root.children).filter(|(_, node)| word(node, kind.properties_tag()));
        if let Some((raw, _)) = nodes.next() {
            if nodes.next().is_some() { return Err("DOCX_NOTE_SETTINGS_INVALID: duplicate section note properties".into()); }
            let xml = complete_fragment(&xml, raw, &scope(raw, &inherited))?;
            let mut parsed = document_note_source::read_properties(&xml, kind, part_name)?;
            warnings.append(&mut parsed.warnings);
            if kind == DocumentNoteKind::Footnote { properties.footnote = Some(parsed.properties); }
            else { properties.endnote = Some(parsed.properties); }
        }
    }
    let mut columns = raw.children.iter().zip(&root.children).filter(|(_, node)| node.namespace.as_deref() == Some(WORD_2013_NAMESPACE) && node.name == "footnoteColumns");
    if let Some((raw, node)) = columns.next() {
        if columns.next().is_some() { return Err("DOCX_NOTE_SETTINGS_INVALID: duplicate footnote columns".into()); }
        let count = node.attribute_ns(WORD_2013_NAMESPACE, "val").and_then(|value| value.parse::<u8>().ok()).filter(|value| *value <= 4)
            .ok_or("DOCX_NOTE_SETTINGS_INVALID: footnote columns must be between zero and four")?;
        let marked = node.attribute_ns(NOTE_EXTENSION_NAMESPACE, "inheritedColumnsFromDocument")
            .map(|value| document_note_source::on_off(Some(value))).transpose()?.unwrap_or(false);
        let inherited_columns = marked && document_columns == Some(count);
        if marked && !inherited_columns { warnings.push("A section's inherited footnote columns differ from the document default; its native value is retained as an explicit override.".into()); }
        let slot = properties.footnote.get_or_insert_with(DocumentNoteProperties::default);
        if !inherited_columns { slot.column_count = Some(count); }
        let metadata = slot.metadata.get_or_insert_with(|| json!({})).as_object_mut()
            .ok_or("DOCX_NOTE_SETTINGS_INVALID: note metadata must be an object")?;
        metadata.insert("officeNoteColumnsSource".into(), json!({ "version": 1, "partName": part_name,
            "xml": complete_fragment(&xml, raw, &scope(raw, &inherited))?,
            "baseline": { "columnCount": count, "inheritedFromDocument": inherited_columns } }));
    }
    let mut suppress_nodes = root.children.iter().filter(|node| word(node, "suppressEndnotes"));
    let suppress = suppress_nodes.next().map(|node| document_note_source::on_off(word_attribute(node, "val"))).transpose()?;
    if suppress_nodes.next().is_some() { return Err("DOCX_NOTE_SETTINGS_INVALID: duplicate endnote suppression".into()); }
    Ok(((properties.footnote.is_some() || properties.endnote.is_some()).then_some(properties), suppress, warnings))
}
