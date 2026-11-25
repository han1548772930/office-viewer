//! Native comment endpoints carry a rank as well as a literal UTF-16 offset.
//! Two references with empty source text must remain separately addressable.
use std::collections::{BTreeMap, BTreeSet, HashMap};

use duke_pages::{BlockDetails, DocModel, DocumentBlock, StructuredDocument};
use office_io_common::{chart_xml::{self, Node}, document_content_controls::{word, word_attribute}};

use crate::{document_format_sidecar::{Element, parse_xml}, document_native_xml::{complete_fragment, scope},
    document_story_source::{collect, stories}, model::{DocumentCommentPayload, DocumentCommentTargetPayload}};

#[derive(Clone, Copy, Default, PartialEq, Eq)]
struct Position { offset: usize, rank: usize }
#[derive(Clone, Copy)]
enum Boundary { Start, End, Reference }
#[derive(Default)]
struct Projection { text: String, position: Position, has_notes: bool, boundaries: Vec<(Position, Boundary, String)> }

impl Projection {
    fn append(&mut self, text: &str) -> Result<(), String> {
        self.position.offset = self.position.offset.checked_add(text.encode_utf16().count()).ok_or("DOCX_COMMENT_NOTE_RANGE: text offset overflow")?;
        self.text.push_str(text);
        Ok(())
    }

    fn visit(&mut self, source: &Element, node: &Node, xml: &str, inherited: &BTreeMap<String, String>) -> Result<(), String> {
        let namespaces = scope(source, inherited);
        if ["pPr", "rPr", "sdtPr", "sdtEndPr", "instrText", "del", "moveFrom", "drawing", "pict", "object", "txbxContent"].iter().any(|name| word(node, name)) { return Ok(()); }
        if word(node, "footnoteReference") || word(node, "endnoteReference") {
            self.has_notes = true;
            self.position.rank = self.position.rank.checked_add(1).ok_or("DOCX_COMMENT_NOTE_RANGE: reference rank overflow")?;
            return Ok(());
        }
        if ["commentRangeStart", "commentRangeEnd", "commentReference"].iter().any(|name| word(node, name)) {
            let id = word_attribute(node, "id").and_then(|id| id.parse::<i64>().ok()).filter(|id| *id >= 0)
                .ok_or("DOCX_COMMENT_NOTE_RANGE: invalid native comment ID")?.to_string();
            let kind = if word(node, "commentRangeStart") { Boundary::Start } else if word(node, "commentRangeEnd") { Boundary::End } else { Boundary::Reference };
            self.boundaries.push((self.position, kind, id));
            return Ok(());
        }
        if word(node, "t") { return self.append(&node.text.replace("\r\n", " ").replace(['\r', '\n'], " ")); }
        if word(node, "tab") { return self.append("\t"); }
        if word(node, "br") || word(node, "cr") { return self.append("\n"); }
        if word(node, "noBreakHyphen") { return self.append("\u{2011}"); }
        if word(node, "softHyphen") { return self.append("\u{ad}"); }
        if word(node, "sym") {
            let fragment = complete_fragment(xml, source, &namespaces)?;
            let canonical = crate::document_format_sidecar::canonical_resource(&fragment, "w:sym")?;
            for token in duke_pages::parse::parse_run_text_tokens(&format!("<w:r>{canonical}</w:r>")) { self.append(&token.text)?; }
            return Ok(());
        }
        for (source, node) in source.children.iter().zip(&node.children) { self.visit(source, node, xml, &namespaces)?; }
        Ok(())
    }
}

fn target(part: &str, block: &str, start: Position, end: Position) -> DocumentCommentTargetPayload {
    DocumentCommentTargetPayload { node_id: block.into(), part_name: part.into(), story: "body".into(),
        start: Some(start.offset), end: Some(end.offset), start_note_reference_count: Some(start.rank),
        end_note_reference_count: Some(end.rank), offset_encoding: "utf-16".into() }
}

fn append_range(part: &str, block: &str, active: &BTreeSet<String>, start: Position, end: Position,
    output: &mut BTreeMap<String, Vec<DocumentCommentTargetPayload>>) {
    if start == end { return; }
    for id in active {
        let targets = output.entry(id.clone()).or_default();
        // Another comment's endpoint cannot split this owner's continuous
        // range. Merge only when both literal position and note rank meet.
        if let Some(previous) = targets.last_mut().filter(|previous| previous.node_id == block && previous.part_name == part
            && previous.end == Some(start.offset) && previous.end_note_reference_count == Some(start.rank)) {
            previous.end = Some(end.offset);
            previous.end_note_reference_count = Some(end.rank);
        } else { targets.push(target(part, block, start, end)); }
    }
}

/// Source text is compared with the actual extracted paragraph before any
/// endpoint is replaced. Unmapped object anchors remain intact and diagnosed.
pub(crate) fn restore(source: &DocModel, document: &StructuredDocument, comments: &mut [DocumentCommentPayload],
    inherited: &BTreeMap<String, String>) -> Result<Vec<String>, String> {
    if comments.is_empty() { return Ok(Vec::new()); }
    let source_stories = stories(source);
    let Some(body) = source_stories.iter().find(|story| story.story == "body") else { return Ok(Vec::new()); };
    let nodes = collect(body);
    if !nodes.paragraphs.iter().any(|(_, paragraph)| paragraph.source_xml.as_ref().is_some_and(|xml|
        xml.contains("footnoteReference") || xml.contains("endnoteReference"))) { return Ok(Vec::new()); }
    let mut extracted = HashMap::<String, String>::new();
    fn collect_text(blocks: &[DocumentBlock], part: &str, output: &mut HashMap<String, String>) {
        for block in blocks {
            if matches!(&block.details, BlockDetails::Paragraph(_)) || matches!(&block.details, BlockDetails::Raw(details) if details["kind"].as_str() == Some("paragraph")) {
                for reference in block.source_refs.iter().filter(|reference| reference.story == "body" && reference.part_name == part) {
                    output.insert(reference.node_id.clone(), block.content.clone());
                }
            }
            collect_text(&block.children, part, output);
        }
    }
    collect_text(&document.blocks, body.part, &mut extracted);
    let mut paragraphs = Vec::new();
    for (address, paragraph) in nodes.paragraphs {
        let Some(xml) = paragraph.source_xml.as_deref() else { return Ok(vec![format!("Native comment/note projection for {} has no source XML; original anchors remain retained.", address.node_id)]); };
        let raw = parse_xml(xml)?;
        let xml = complete_fragment(xml, &raw, &scope(&raw, inherited))?;
        let raw = parse_xml(&xml)?;
        let root = chart_xml::parse(&xml)?;
        let mut projection = Projection::default();
        projection.visit(&raw, &root, &xml, &BTreeMap::new())?;
        if extracted.get(&address.node_id) != Some(&projection.text) {
            return Ok(vec![format!("Native comment/note projection for {} differs from the extracted text; original anchors remain retained.", address.node_id)]);
        }
        paragraphs.push((address.node_id, projection));
    }
    let mut active = BTreeSet::new();
    let mut affected = BTreeSet::new();
    let mut targets = BTreeMap::<String, Vec<DocumentCommentTargetPayload>>::new();
    let mut references = BTreeMap::new();
    for (block, projection) in paragraphs {
        if projection.has_notes { affected.extend(active.iter().cloned()); }
        let mut from = Position::default();
        for (position, kind, id) in projection.boundaries {
            append_range(body.part, &block, &active, from, position, &mut targets);
            if projection.has_notes { affected.insert(id.clone()); }
            match kind {
                Boundary::Start => { active.insert(id); },
                Boundary::End => { active.remove(&id); },
                Boundary::Reference => { references.insert(id, target(body.part, &block, position, position)); },
            }
            from = position;
        }
        append_range(body.part, &block, &active, from, projection.position, &mut targets);
    }
    let mut warnings = Vec::new();
    for comment in comments.iter_mut().filter(|comment| affected.contains(&comment.id)) {
        if comment.targets.iter().any(|target| target.start.is_none() || target.end.is_none()) {
            warnings.push(format!("Native comment {} also owns an object anchor; its targets remain retained without invented note ranks.", comment.id));
            continue;
        }
        let mut replacement = targets.remove(&comment.id).unwrap_or_default();
        if replacement.is_empty() { if let Some(reference) = references.remove(&comment.id) { replacement.push(reference); } }
        comment.targets = replacement;
    }
    Ok(warnings)
}
