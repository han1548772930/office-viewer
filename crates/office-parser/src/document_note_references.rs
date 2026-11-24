//! Reference attributes and styling use the paragraph's real native owner.
//! They are annotations of the Duke DTO, never a second document parser.
use std::collections::{BTreeMap, HashMap};

use duke_pages::{BlockDetails, DocModel, DocumentBlock, StructuredDocument, TextStyle, parse::ParsedStyleSheet};
use office_io_common::{chart_xml::{self, Node}, document_content_controls::{word, word_attribute},
    document_note_source::{self, DocumentNoteKind}};
use serde_json::{Value, json};

use crate::{document_format_sidecar::{Element, parse_xml}, document_native_xml::{complete_fragment, scope, standard_word_namespaces},
    document_story_source::{SourceAddress, collect, stories}};
use super::NativeNotes;

struct Reference {
    kind: DocumentNoteKind,
    id: i64,
    custom_mark_follows: bool,
    custom_mark: Option<String>,
    source_format: Value,
    source_xml: String,
}

fn references(xml: &str, inherited: &BTreeMap<String, String>) -> Result<Vec<Reference>, String> {
    let raw = parse_xml(xml)?;
    let xml = complete_fragment(xml, &raw, &scope(&raw, inherited))?;
    let raw = parse_xml(&xml)?;
    let root = chart_xml::parse(&xml)?;
    struct SourceRun<'a> { raw: &'a Element, node: &'a Node, namespaces: BTreeMap<String, String> }
    fn collect_runs<'a>(raw: &'a Element, node: &'a Node, inherited: &BTreeMap<String, String>, result: &mut Vec<SourceRun<'a>>) {
        let namespaces = scope(raw, inherited);
        if word(node, "r") {
            result.push(SourceRun { raw, node, namespaces });
            return;
        }
        if ["pPr", "rPr", "del", "moveFrom", "drawing", "pict", "txbxContent"].iter().any(|name| word(node, name)) { return; }
        for (raw, child) in raw.children.iter().zip(&node.children) { collect_runs(raw, child, &namespaces, result); }
    }
    fn marker_text(run: &SourceRun<'_>, from: usize, xml: &str) -> Result<Option<String>, String> {
        let mut text = String::new();
        let mut present = false;
        for (raw, child) in run.raw.children.iter().zip(&run.node.children).skip(from) {
            if word(child, "rPr") { continue; }
            if word(child, "t") { present = true; text.push_str(&child.text); }
            else if word(child, "sym") {
                present = true;
                let fragment = complete_fragment(xml, raw, &scope(raw, &run.namespaces))?;
                let canonical = crate::document_format_sidecar::canonical_resource(&fragment, "w:sym")?;
                text.extend(duke_pages::parse::parse_run_text_tokens(&format!("<w:r>{canonical}</w:r>")).into_iter().map(|token| token.text));
            } else { break; }
        }
        Ok(present.then_some(text))
    }
    let mut runs = Vec::new();
    collect_runs(&raw, &root, &BTreeMap::new(), &mut runs);
    let mut result = Vec::new();
    for (run_index, run) in runs.iter().enumerate() {
            let properties = run.raw.children.iter().zip(&run.node.children).find(|(_, node)| word(node, "rPr")).map(|(raw, _)| raw);
            let format = json!({ "version": 1, "propertiesXml": properties.map(|properties| xml[properties.start..properties.end].to_owned()).unwrap_or_else(|| "<w:rPr/>".into()),
                "namespaces": properties.map(|properties| scope(properties, &run.namespaces)).unwrap_or_else(|| run.namespaces.clone()) });
            for (index, (source, child)) in run.raw.children.iter().zip(&run.node.children).enumerate() {
                let kind = if word(child, "footnoteReference") { DocumentNoteKind::Footnote }
                    else if word(child, "endnoteReference") { DocumentNoteKind::Endnote } else { continue; };
                let id = word_attribute(child, "id").and_then(|value| value.parse::<i64>().ok())
                    .filter(|value| value.unsigned_abs() <= document_note_source::NOTE_SAFE_INTEGER)
                    .ok_or("DOCX_NOTE_INVALID: reference ID is not a safe integer")?;
                let follows = word_attribute(child, "customMarkFollows").map(|value| document_note_source::on_off(Some(value))).transpose()?.unwrap_or(false);
                let custom_mark = if follows {
                    match marker_text(run, index + 1, &xml)? {
                        Some(mark) => Some(mark),
                        None => runs.get(run_index + 1).map(|next| marker_text(next, 0, &xml)).transpose()?.flatten(),
                    }
                } else { None };
                result.push(Reference { kind, id, custom_mark_follows: follows,
                    custom_mark, source_format: format.clone(), source_xml: complete_fragment(&xml, source, &scope(source, &run.namespaces))? });
            }
    }
    Ok(result)
}

pub(super) fn restore(native: &mut NativeNotes, document: &mut StructuredDocument, source: &DocModel, _styles: &ParsedStyleSheet) -> Result<(), String> {
    let canonical = crate::document_paragraph_source::canonical_scope(source)?;
    let mut indexed = HashMap::<SourceAddress, Vec<Reference>>::new();
    for story in stories(source) {
        let mut inherited = native.namespaces.get(story.part).cloned().unwrap_or_default();
        inherited.extend(standard_word_namespaces());
        if story.story == "body" { inherited.extend(canonical.clone()); }
        for (address, paragraph) in collect(&story).paragraphs {
            let Some(xml) = &paragraph.source_xml else { continue; };
            if !xml.contains("footnoteReference") && !xml.contains("endnoteReference") { continue; }
            let refs = references(xml, &inherited)?;
            if !refs.is_empty() { indexed.insert(address, refs); }
        }
    }
    fn visit(blocks: &mut [DocumentBlock], note_id: Option<i64>, indexed: &HashMap<SourceAddress, Vec<Reference>>, native: &mut NativeNotes) -> Result<(), String> {
        for block in blocks {
            let note_id = block.note_id.or(note_id);
            let refs = block.source_refs.iter().find_map(|source| indexed.get(&SourceAddress::from_ref(source, note_id)));
            if let Some(refs) = refs.filter(|_| matches!(block.details, BlockDetails::Paragraph(_) | BlockDetails::Raw(_))) {
                let mut details = block.details.to_json_value();
                let inlines = details.get_mut("inlines").and_then(Value::as_array_mut);
                if let Some(inlines) = inlines {
                    let matching = inlines.iter().filter(|inline| inline.get("noteReference").is_some_and(|reference| !reference.is_null())).count();
                    if matching != refs.len() { return Err(format!("DOCX_NOTE_REFERENCE_MAPPING: paragraph {} has {} native references and {} extracted references", block.id, refs.len(), matching)); }
                    let mut index = 0;
                    for inline in inlines.iter_mut().filter(|inline| inline.get("noteReference").is_some_and(|reference| !reference.is_null())) {
                        let reference = &refs[index]; index += 1;
                        let kind = inline["noteReference"]["kind"].as_str();
                        let id = inline["noteReference"]["id"].as_i64();
                        if kind != Some(reference.kind.name()) || id != Some(reference.id) { return Err("DOCX_NOTE_REFERENCE_MAPPING: reference sequence does not match its native paragraph".into()); }
                        inline["sourceFormat"] = reference.source_format.clone();
                        inline["noteReference"]["customMarkFollows"] = json!(reference.custom_mark_follows);
                        inline["noteReference"]["sourceXml"] = json!(reference.source_xml);
                        if let Some(mark) = &reference.custom_mark { inline["noteReference"]["customMark"] = json!(mark); }
                        if let Some(definition) = native.definitions.get_mut(&(reference.kind, reference.id)) {
                            if let Some(mark) = &reference.custom_mark {
                                if definition.custom_mark.as_ref().is_some_and(|existing| existing != mark) {
                                    return Err("DOCX_NOTE_REFERENCE_MAPPING: one note has conflicting custom reference marks".into());
                                }
                                definition.custom_mark = Some(mark.clone());
                            }
                            if !inline["style"].is_null() {
                                let style = serde_json::from_value::<TextStyle>(inline["style"].clone()).map_err(|error| error.to_string())?;
                                definition.reference_style.get_or_insert(style);
                            }
                        }
                    }
                    block.details = BlockDetails::Raw(details);
                }
            }
            visit(&mut block.children, note_id, indexed, native)?;
        }
        Ok(())
    }
    visit(&mut document.blocks, None, &indexed, native)?;
    let mut body_markers = HashMap::<SourceAddress, (DocumentNoteKind, String)>::new();
    for story in stories(source) {
        let Some(id) = story.note_id else { continue; };
        let kind = if story.story == "footnote" { DocumentNoteKind::Footnote } else if story.story == "endnote" { DocumentNoteKind::Endnote } else { continue; };
        let Some(mark) = native.definitions.get(&(kind, id)).and_then(|definition| definition.custom_mark.as_ref()).filter(|mark| !mark.is_empty()).cloned() else { continue; };
        let hosts = collect(&story);
        let Some((address, paragraph)) = hosts.paragraphs.first() else { continue; };
        let Some(xml) = paragraph.source_xml.as_deref() else { continue; };
        let mut inherited = native.namespaces.get(story.part).cloned().unwrap_or_default();
        inherited.extend(standard_word_namespaces());
        let raw = parse_xml(xml)?;
        let xml = complete_fragment(xml, &raw, &scope(&raw, &inherited))?;
        let root = chart_xml::parse(&xml)?;
        let Some(run) = root.descendants("r").into_iter().find(|node| word(node, "r")) else { continue; };
        let reference_style = if kind == DocumentNoteKind::Footnote { "FootnoteReference" } else { "EndnoteReference" };
        let styled = run.children.iter().find(|node| word(node, "rPr")).is_some_and(|properties| properties.children.iter()
            .any(|node| word(node, "rStyle") && word_attribute(node, "val") == Some(reference_style)));
        let mut text = String::new();
        for child in run.children.iter().filter(|node| !word(node, "rPr")) {
            if word(child, "t") { text.push_str(&child.text); } else { text.clear(); break; }
        }
        if styled && text == mark {
            body_markers.insert(address.clone(), (kind, mark));
            if let Some(definition) = native.definitions.get_mut(&(kind, id)) { definition.show_body_reference = true; }
        }
    }
    fn annotate_body(blocks: &mut [DocumentBlock], inherited_note: Option<i64>, markers: &HashMap<SourceAddress, (DocumentNoteKind, String)>) -> Result<(), String> {
        for block in blocks {
            let note = block.note_id.or(inherited_note);
            if let Some((kind, marker)) = block.source_refs.iter().find_map(|source| markers.get(&SourceAddress::from_ref(source, note))) {
                let mut details = block.details.to_json_value();
                if let Some(inline) = details.get_mut("inlines").and_then(Value::as_array_mut).and_then(|inlines| inlines.iter_mut()
                    .find(|inline| matches!(inline["type"].as_str(), Some("text" | "form_field")) && !inline["text"].as_str().unwrap_or_default().is_empty())) {
                    if inline["text"].as_str().is_some_and(|text| text.starts_with(marker)) {
                        inline["noteBodyMarker"] = json!({ "version": 1, "kind": kind, "text": marker });
                        block.details = BlockDetails::Raw(details);
                    }
                }
            }
            annotate_body(&mut block.children, note, markers)?;
        }
        Ok(())
    }
    annotate_body(&mut document.blocks, None, &body_markers)
}
