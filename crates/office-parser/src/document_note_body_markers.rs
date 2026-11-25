//! Body markers retain their own paragraph, UTF-16 boundary and source runs.
//! Generated identities survive destination rStyle remapping and repeated saves.
use std::collections::{BTreeMap, HashMap, HashSet};

use duke_pages::{BlockDetails, DocModel, DocumentBlock, StructuredDocument};
use office_io_common::{chart_xml, document_content_controls::{word, word_attribute},
    document_note_source::{self, DocumentNoteKind, NOTE_EXTENSION_NAMESPACE}};
use serde_json::{Value, json};

use crate::{document_format_sidecar::parse_xml, document_native_xml::{complete_fragment, scope, standard_word_namespaces},
    document_story_source::{SourceAddress, collect, stories}};
use super::{SourceRun, collect_runs, marker_text, paragraph_style, source_format};
use super::super::NativeNotes;

struct BodyMarker { id: String, offset: usize, note_count: usize, text: String, source: Value }

fn marker_identity(run: &SourceRun<'_>, kind: DocumentNoteKind, note_id: i64, ordinal: &mut usize,
    identities: &mut HashSet<String>) -> Result<String, String> {
    if *ordinal >= 4096 { return Err("DOCX_NOTE_BODY_MARKER_INVALID: body marker occurrence budget exceeded".into()); }
    *ordinal += 1;
    let identity = generated_id(run)?.unwrap_or_else(|| format!("{}:{note_id}:body:{ordinal}", kind.name()));
    if !identities.insert(identity.clone()) { return Err("DOCX_NOTE_BODY_MARKER_INVALID: duplicated body marker identity".into()); }
    Ok(identity)
}

fn advance_offset(offset: &mut usize, text: &str) -> Result<(), String> {
    *offset = offset.checked_add(text.encode_utf16().count())
        .filter(|offset| *offset as u64 <= document_note_source::NOTE_SAFE_INTEGER)
        .ok_or("DOCX_NOTE_BODY_MARKER_INVALID: body marker geometry exceeds the safe integer range")?;
    Ok(())
}

fn generated_kind(run: &SourceRun<'_>) -> Result<Option<&'static str>, String> {
    match run.node.attribute_ns(NOTE_EXTENSION_NAMESPACE, "bodyMarker") {
        Some("custom") => Ok(Some("custom")), Some("automatic") => Ok(Some("automatic")), None => Ok(None),
        Some(_) => Err("DOCX_NOTE_BODY_MARKER_INVALID: unrecognized generated marker kind".into()),
    }
}

fn generated_id(run: &SourceRun<'_>) -> Result<Option<String>, String> {
    run.node.attribute_ns(NOTE_EXTENSION_NAMESPACE, "bodyMarkerIdHex")
        .map(document_note_source::decode_public_note_source_id).transpose()
}

fn canonical_custom(run: &SourceRun<'_>, kind: DocumentNoteKind) -> bool {
    let canonical = if kind == DocumentNoteKind::Footnote { "FootnoteReference" } else { "EndnoteReference" };
    run.node.children.iter().find(|node| word(node, "rPr")).is_some_and(|properties| properties.children.iter()
        .any(|node| word(node, "rStyle") && word_attribute(node, "val") == Some(canonical)))
}

/// Canonical Word custom markers may span several real runs. Consume only the
/// exact expected mark; the following styled body text remains ordinary text.
fn custom_group(runs: &[SourceRun<'_>], xml: &str, paragraph_style: Option<&str>, kind: DocumentNoteKind,
    expected: Option<&str>) -> Result<Option<(String, Value, usize)>, String> {
    let Some(first) = runs.first() else { return Ok(None); };
    let generated = generated_kind(first)? == Some("custom");
    if !generated && (!canonical_custom(first, kind) || expected.is_none()) { return Ok(None); }
    let identity = generated_id(first)?;
    let mut text = String::new(); let mut sources = Vec::new();
    for run in runs {
        if generated {
            if generated_kind(run)? != Some("custom") || generated_id(run)? != identity { break; }
        } else if !canonical_custom(run, kind) { break; }
        let only_mark = run.node.children.iter().all(|node| word(node, "rPr") || word(node, "t") || word(node, "sym"));
        let marker = if only_mark { marker_text(run, 0, xml, paragraph_style)? } else { None };
        let Some(marker) = marker else {
            if generated { return Err("DOCX_NOTE_BODY_MARKER_INVALID: custom marker contains other content".into()); }
            return Ok(None);
        };
        text.push_str(&marker.text); sources.push(marker.source);
        if text.encode_utf16().count() > 4096 || sources.len() > 4096 {
            return Err("DOCX_NOTE_BODY_MARKER_INVALID: custom marker exceeds its budget".into());
        }
        if !generated {
            if expected == Some(text.as_str()) { break; }
            if !expected.is_some_and(|expected| expected.starts_with(&text)) { return Ok(None); }
        }
    }
    if expected.is_some_and(|expected| expected != text) {
        if generated { return Err("DOCX_NOTE_BODY_MARKER_INVALID: body and reference custom markers differ".into()); }
        return Ok(None);
    }
    let consumed = sources.len();
    Ok(Some((text.clone(), json!({ "version": 1, "text": text, "runs": sources }), consumed)))
}

fn native_text(raw: &crate::document_format_sidecar::Element, node: &chart_xml::Node, run: &SourceRun<'_>, xml: &str) -> Result<String, String> {
    if !["t", "tab", "br", "cr", "noBreakHyphen", "softHyphen", "sym", "footnoteReference", "endnoteReference"]
        .iter().any(|name| word(node, name)) { return Ok(String::new()); }
    let fragment = complete_fragment(xml, raw, &scope(raw, &run.namespaces))?;
    let canonical = crate::document_format_sidecar::canonical_resource(&fragment, &format!("w:{}", node.name))?;
    Ok(duke_pages::parse::parse_run_text_tokens(&format!("<w:r>{canonical}</w:r>")).into_iter().map(|token| token.text).collect())
}

pub(super) fn restore(native: &mut NativeNotes, document: &mut StructuredDocument, source: &DocModel) -> Result<(), String> {
    let mut markers = HashMap::<SourceAddress, (DocumentNoteKind, Vec<BodyMarker>)>::new();
    for story in stories(source) {
        let Some(id) = story.note_id else { continue; };
        let kind = match story.story { "footnote" => DocumentNoteKind::Footnote, "endnote" => DocumentNoteKind::Endnote, _ => continue };
        if native.is_decoration(kind, id) { continue; }
        let mut identities = HashSet::new(); let mut ordinal = 0usize; let mut complete = true; let mut occurrence_ids = Vec::new();
        for (address, paragraph) in collect(&story).paragraphs {
            let Some(xml) = paragraph.source_xml.as_deref() else { complete = false; continue; };
            let mut inherited = native.namespaces.get(story.part).cloned().unwrap_or_default(); inherited.extend(standard_word_namespaces());
            let raw = parse_xml(xml)?; let xml = complete_fragment(xml, &raw, &scope(&raw, &inherited))?;
            let raw = parse_xml(&xml)?; let root = chart_xml::parse(&xml)?;
            let mut runs = Vec::new(); collect_runs(&raw, &root, &BTreeMap::new(), &mut runs);
            let mut offset = 0usize; let mut note_count = 0usize; let mut index = 0usize; let mut found = Vec::new();
            while let Some(run) = runs.get(index) {
                let generated = generated_kind(run)?;
                if generated.is_none() && generated_id(run)?.is_some() {
                    return Err("DOCX_NOTE_BODY_MARKER_INVALID: generated identity has no marker kind".into());
                }
                let expected = native.definitions.get(&(kind, id)).and_then(|definition| definition.custom_mark.as_deref());
                if let Some((mark, marker_source, consumed)) = custom_group(&runs[index..], &xml, paragraph_style(&root), kind, expected)? {
                    let identity = marker_identity(run, kind, id, &mut ordinal, &mut identities)?;
                    occurrence_ids.push(identity.clone());
                    found.push(BodyMarker { id: identity, offset, note_count, text: mark.clone(), source: marker_source.clone() });
                    if let Some(definition) = native.definitions.get_mut(&(kind, id)) {
                        if definition.custom_mark.as_ref().is_some_and(|existing| existing != &mark) {
                            return Err("DOCX_NOTE_BODY_MARKER_INVALID: conflicting custom body marker text".into());
                        }
                        definition.show_body_reference = true; definition.custom_mark = Some(mark.clone());
                        definition.body_marker_source.get_or_insert(marker_source);
                    }
                    advance_offset(&mut offset, &mark)?; index += consumed; continue;
                }
                if generated == Some("custom") { return Err("DOCX_NOTE_BODY_MARKER_INVALID: custom marker group could not be restored".into()); }
                if generated == Some("automatic") && (run.node.children.iter().filter(|node| word(node, kind.body_reference_tag())).count() != 1
                    || run.node.children.iter().any(|node| !word(node, "rPr") && !word(node, kind.body_reference_tag()))) {
                    return Err("DOCX_NOTE_BODY_MARKER_INVALID: automatic marker contains other content".into());
                }
                for (raw, node) in run.raw.children.iter().zip(&run.node.children) {
                    if word(node, kind.body_reference_tag()) {
                        let identity = marker_identity(run, kind, id, &mut ordinal, &mut identities)?;
                        occurrence_ids.push(identity.clone());
                        let marker_source = json!({ "version": 1, "sourceFormat": source_format(run, &xml, paragraph_style(&root)),
                            "sourceXml": complete_fragment(&xml, raw, &scope(raw, &run.namespaces))? });
                        found.push(BodyMarker { id: identity, offset, note_count, text: String::new(), source: marker_source.clone() });
                        if let Some(definition) = native.definitions.get_mut(&(kind, id)) {
                            definition.show_body_reference = true; definition.body_marker_source.get_or_insert(marker_source);
                        }
                    }
                    if word(node, "footnoteReference") || word(node, "endnoteReference") { note_count += 1; }
                    advance_offset(&mut offset, &native_text(raw, node, run, &xml)?)?;
                }
                index += 1;
            }
            if !found.is_empty() && markers.insert(address, (kind, found)).is_some() {
                return Err("DOCX_NOTE_BODY_MARKER_MAPPING: duplicate native marker paragraph identity".into());
            }
        }
        if let Some(definition) = native.definitions.get_mut(&(kind, id)) {
            let mut layout = json!({ "version": 1, "complete": complete, "showBodyReference": definition.show_body_reference,
                "occurrenceIds": occurrence_ids });
            if let Some(mark) = &definition.custom_mark { layout["customMark"] = json!(mark); }
            definition.body_marker_layout = Some(layout);
            if !complete { native.warnings.push(format!("Native {} {id} has paragraphs without marker-position source XML; its marker layout remains unresolved.", kind.name())); }
        }
    }
    let mut mapped = HashSet::new();
    annotate_body(&mut document.blocks, None, &markers, &mut mapped)?;
    if mapped.len() != markers.len() { return Err("DOCX_NOTE_BODY_MARKER_MAPPING: a native marker paragraph has no extracted block".into()); }
    Ok(())
}

fn utf16_byte(text: &str, offset: usize) -> Option<usize> {
    let mut count = 0usize;
    for (index, character) in text.char_indices() {
        if count == offset { return Some(index); }
        count += character.len_utf16();
        if count > offset { return None; }
    }
    (count == offset).then_some(text.len())
}

fn annotate_inlines(inlines: &[Value], kind: DocumentNoteKind, markers: &[BodyMarker]) -> Result<Vec<Value>, String> {
    let mut result = Vec::new(); let mut offset = 0usize;
    for inline in inlines {
        if !matches!(inline["type"].as_str(), Some("text" | "form_field")) { result.push(inline.clone()); continue; }
        let text = inline["text"].as_str().unwrap_or_default(); let length = text.encode_utf16().count();
        let mut cursor = 0usize;
        for marker in markers.iter().filter(|marker| !marker.text.is_empty()) {
            let end = marker.offset + marker.text.encode_utf16().count();
            let from = marker.offset.max(offset); let to = end.min(offset + length);
            if from >= to { continue; }
            if inline.get("noteReference").is_some_and(|reference| !reference.is_null()) { return Err("DOCX_NOTE_BODY_MARKER_MAPPING: body marker overlaps a note reference".into()); }
            let start_byte = utf16_byte(text, from - offset).ok_or("DOCX_NOTE_BODY_MARKER_MAPPING: marker splits a surrogate pair")?;
            let end_byte = utf16_byte(text, to - offset).ok_or("DOCX_NOTE_BODY_MARKER_MAPPING: marker splits a surrogate pair")?;
            if cursor < start_byte { let mut prefix = inline.clone(); prefix["text"] = json!(&text[cursor..start_byte]); result.push(prefix); }
            let expected_start = utf16_byte(&marker.text, from - marker.offset).ok_or("DOCX_NOTE_BODY_MARKER_MAPPING: invalid marker boundary")?;
            let expected_end = utf16_byte(&marker.text, to - marker.offset).ok_or("DOCX_NOTE_BODY_MARKER_MAPPING: invalid marker boundary")?;
            if text[start_byte..end_byte] != marker.text[expected_start..expected_end] { return Err("DOCX_NOTE_BODY_MARKER_MAPPING: native marker differs from extracted text".into()); }
            let mut current = inline.clone(); current["text"] = json!(&text[start_byte..end_byte]);
            current["noteBodyMarker"] = json!({ "version": 1, "kind": kind, "text": &text[start_byte..end_byte], "id": marker.id });
            result.push(current); cursor = end_byte;
        }
        if cursor < text.len() || text.is_empty() { let mut suffix = inline.clone(); suffix["text"] = json!(&text[cursor..]); result.push(suffix); }
        offset += length;
    }
    if markers.iter().any(|marker| marker.offset + marker.text.encode_utf16().count() > offset) {
        return Err("DOCX_NOTE_BODY_MARKER_MAPPING: extracted paragraph omitted marker text".into());
    }
    Ok(result)
}

fn annotate_body(blocks: &mut [DocumentBlock], inherited_note: Option<i64>, markers: &HashMap<SourceAddress, (DocumentNoteKind, Vec<BodyMarker>)>,
    mapped: &mut HashSet<SourceAddress>) -> Result<(), String> {
    for block in blocks {
        let note = block.note_id.or(inherited_note);
        let paragraph = matches!(&block.details, BlockDetails::Paragraph(_))
            || matches!(&block.details, BlockDetails::Raw(details) if details["kind"].as_str() == Some("paragraph"));
        if let Some((address, (kind, entries))) = paragraph.then(|| block.source_refs.iter().find_map(|source| {
            let address = SourceAddress::from_ref(source, note); markers.get(&address).map(|markers| (address, markers))
        })).flatten() {
            if !mapped.insert(address.clone()) { return Err("DOCX_NOTE_BODY_MARKER_MAPPING: one native marker paragraph has several extracted text owners".into()); }
            let mut details = block.details.to_json_value();
            let empty = Vec::new();
            let inlines = match details.get("inlines") {
                Some(value) => value.as_array().ok_or("DOCX_NOTE_BODY_MARKER_MAPPING: marker paragraph inlines are invalid")?,
                None if block.content.is_empty() => &empty,
                None => return Err("DOCX_NOTE_BODY_MARKER_MAPPING: marker has no extracted paragraph".into()),
            };
            details["inlines"] = json!(annotate_inlines(inlines, *kind, entries)?);
            details["noteBodyMarkers"] = json!({ "version": 1, "kind": kind, "markers": entries.iter().map(|marker|
                json!({ "id": marker.id, "sourceOffset": marker.offset, "sourceNoteReferenceCount": marker.note_count,
                    "text": marker.text, "source": marker.source })).collect::<Vec<_>>() });
            block.details = BlockDetails::Raw(details);
        }
        annotate_body(&mut block.children, note, markers, mapped)?;
    }
    Ok(())
}
