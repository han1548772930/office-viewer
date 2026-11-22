use std::collections::{BTreeMap, BTreeSet};

use duke_pages::{CommentTarget, DocModel, DocNode, DocumentComment, FormFieldType, ParagraphChildNode, ParagraphNode, SourceRef, TableCellContentNode, TableNode};

use super::{Element, parse_xml};

#[derive(Clone, Copy)]
enum Boundary { Start, End, Reference }

#[derive(Default)]
struct Projection {
    text: String,
    offset: usize,
    boundaries: Vec<(usize, Boundary, i64)>,
    has_sdt: bool,
}

impl Projection {
    fn append(&mut self, text: &str) {
        self.offset += text.encode_utf16().count();
        self.text.push_str(text);
    }

    fn visit(&mut self, element: &Element) -> Result<(), String> {
        match element.name.as_str() {
            "pPr" | "rPr" | "sdtPr" | "sdtEndPr" | "instrText" | "del" | "moveFrom" | "drawing" | "pict" | "object" | "txbxContent" => return Ok(()),
            "commentRangeStart" | "commentRangeEnd" | "commentReference" => {
                let id = element.integer("id").filter(|id| *id >= 0).ok_or("A source comment boundary has an invalid ID.")?;
                let boundary = match element.name.as_str() { "commentRangeStart" => Boundary::Start, "commentRangeEnd" => Boundary::End, _ => Boundary::Reference };
                self.boundaries.push((self.offset, boundary, id));
                return Ok(());
            }
            "t" => { self.append(&element.text.replace("\r\n", " ").replace(['\r', '\n'], " ")); return Ok(()); },
            "tab" => { self.append("\t"); return Ok(()); },
            "br" | "cr" => { self.append("\n"); return Ok(()); },
            "noBreakHyphen" => { self.append("\u{2011}"); return Ok(()); },
            "softHyphen" => { self.append("\u{ad}"); return Ok(()); },
            "sdt" => self.has_sdt = true,
            _ => {},
        }
        for child in &element.children { self.visit(child)?; }
        Ok(())
    }
}

fn visible_text(paragraph: &ParagraphNode) -> String {
    paragraph.children.iter().map(|child| match child {
        ParagraphChildNode::Text(run) => run.text.as_str(),
        ParagraphChildNode::FormField(field) => field.value.as_deref().unwrap_or_else(|| {
            if field.field_type == FormFieldType::Checkbox {
                if field.checked == Some(true) { field.checked_symbol.as_deref().unwrap_or("☒") } else { field.unchecked_symbol.as_deref().unwrap_or("☐") }
            } else { "" }
        }),
        ParagraphChildNode::Image(_) => "",
    }).collect()
}

fn paragraphs<'a>(nodes: &'a [DocNode], root: &str, output: &mut Vec<(&'a ParagraphNode, String)>) {
    fn table<'a>(value: &'a TableNode, root: &str, output: &mut Vec<(&'a ParagraphNode, String)>) {
        for (row_index, row) in value.rows.iter().enumerate() { for (cell_index, cell) in row.cells.iter().enumerate() { for (node_index, node) in cell.nodes.iter().enumerate() {
            let fallback = format!("{root}/row/{row_index}/cell/{cell_index}/{node_index}");
            match node {
                TableCellContentNode::Paragraph(paragraph) => output.push((paragraph, paragraph.block_id.clone().unwrap_or(fallback))),
                TableCellContentNode::Table(nested) => table(nested, nested.block_id.as_deref().unwrap_or(&fallback), output),
            }
        } } }
    }
    for (index, node) in nodes.iter().enumerate() {
        let fallback = format!("{root}/{index}");
        match node {
            DocNode::Paragraph(paragraph) => output.push((paragraph, paragraph.block_id.clone().unwrap_or(fallback))),
            DocNode::Table(value) => table(value, value.block_id.as_deref().unwrap_or(&fallback), output),
        }
    }
}

fn source_reference(model: &DocModel, part: &str, node_id: &str, start: usize, end: usize) -> SourceRef {
    SourceRef { document_id: model.metadata.document_id.clone().unwrap_or_else(|| "unsaved".into()),
        node_id: node_id.into(), part_name: part.into(), story: "body".into(), start: Some(start), end: Some(end), offset_encoding: "utf-16".into() }
}

fn append_range(model: &DocModel, part: &str, paragraph: &str, active: &BTreeSet<i64>, start: usize, end: usize, targets: &mut BTreeMap<i64, Vec<CommentTarget>>) {
    if start == end { return; }
    for id in active {
        let ranges = targets.entry(*id).or_default();
        if let Some(last) = ranges.last_mut().filter(|last| last.source.node_id == paragraph && last.source.end == Some(start)) { last.source.end = Some(end); }
        else { ranges.push(CommentTarget { source: source_reference(model, part, paragraph, start, end), child_start: 0, child_end: 0 }); }
    }
}

/// Duke aggregates typed SDTs into one FormField child, while its review walker
/// counts their inner runs. Restore affected body anchors from visible XML only
/// when that projection exactly matches the production model's paragraph text.
pub(super) fn restore_sdt_comment_targets(model: &DocModel, comments: &mut [DocumentComment]) -> Vec<String> {
    if comments.is_empty() { return Vec::new(); }
    let part = model.metadata.source_part_names.as_ref().and_then(|parts| parts.get("document")).map_or("word/document.xml", String::as_str);
    let root = format!("docx:{}/{part}/body", model.metadata.document_id.as_deref().unwrap_or("unsaved"));
    let mut source = Vec::new();
    paragraphs(&model.nodes, &root, &mut source);
    if !source.iter().any(|(paragraph, _)| paragraph.source_xml.as_ref().is_some_and(|xml| xml.contains(":sdt") || xml.contains("<sdt"))) { return Vec::new(); }
    let mut projections = Vec::new();
    for (paragraph, id) in &source {
        let Some(xml) = &paragraph.source_xml else { return vec![format!("DOCX SDT comment projection for {id} has no source XML; existing comment targets were retained.")]; };
        let mut projection = Projection::default();
        let result = parse_xml(xml).and_then(|root| projection.visit(&root));
        if let Err(error) = result { return vec![format!("DOCX SDT comment projection for {id} could not be read: {error}; existing comment targets were retained.")]; }
        if projection.text != visible_text(paragraph) { return vec![format!("DOCX SDT comment projection for {id} differs from visible paragraph text; existing comment targets were retained.")]; }
        projections.push((id, projection));
    }
    let mut active = BTreeSet::new();
    let mut affected = BTreeSet::new();
    let mut targets = BTreeMap::<i64, Vec<CommentTarget>>::new();
    let mut references = BTreeMap::new();
    for (id, projection) in projections {
        if projection.has_sdt { affected.extend(active.iter().copied()); }
        let mut offset = 0;
        for (position, boundary, comment_id) in projection.boundaries {
            append_range(model, part, id, &active, offset, position, &mut targets);
            if projection.has_sdt { affected.insert(comment_id); }
            match boundary {
                Boundary::Start => { active.insert(comment_id); },
                Boundary::End => { active.remove(&comment_id); },
                Boundary::Reference => { references.insert(comment_id, source_reference(model, part, id, position, position)); },
            }
            offset = position;
        }
        append_range(model, part, id, &active, offset, projection.offset, &mut targets);
    }
    let mut warnings = Vec::new();
    for comment in comments.iter_mut().filter(|comment| affected.contains(&comment.id)) {
        if comment.targets.iter().any(|target| target.source.story == "body" && target.source.start.is_none()) {
            warnings.push(format!("DOCX SDT comment {} also targets a non-text object; existing targets were retained.", comment.id));
            continue;
        }
        comment.targets.retain(|target| target.source.story != "body");
        comment.targets.extend(targets.remove(&comment.id).unwrap_or_default());
        if let Some(reference) = references.remove(&comment.id) { comment.reference = Some(reference); }
    }
    warnings
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    fn reviewed_control() -> (DocModel, Vec<DocumentComment>) {
        let input = json!({ "kind": "document", "name": "projection-guard", "blocks": [
            { "id": "lead", "type": "paragraph", "runs": [{ "text": "Lead" }] },
            { "id": "body", "type": "paragraph", "runs": [{ "text": "Alpha" }] },
        ], "contentControls": [{ "id": "typed", "type": "plainText", "blockId": "body", "startOffset": 1, "endOffset": 4 }],
            "comments": { "review": { "anchor": [{ "blockId": "body", "startOffset": 1, "endOffset": 5 }], "messages": [{ "authorName": "Author", "text": "Retain current targets" }] } } });
        let (bytes, warnings) = office_writer::export_docx_bytes(&input.to_string()).unwrap();
        assert!(warnings.is_empty(), "{warnings:?}");
        let model = duke_pages::parse_document_model(&bytes).unwrap();
        let comments = duke_pages::extract_document_review(&model).comments;
        assert_eq!(comments.len(), 1);
        (model, comments)
    }

    #[test]
    fn unmatched_invalid_or_missing_source_projection_retains_all_existing_targets() {
        for xml in [
            Some("<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:r><w:t>Changed</w:t></w:r></w:p>"),
            Some("<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:sdt>"),
            None,
        ] {
            let (mut model, mut comments) = reviewed_control();
            let before = comments.clone();
            let DocNode::Paragraph(lead) = &mut model.nodes[0] else { panic!("fixture starts with a paragraph"); };
            lead.source_xml = xml.map(str::to_owned);
            let warnings = restore_sdt_comment_targets(&model, &mut comments);
            assert_eq!(warnings.len(), 1, "{warnings:?}");
            assert!(warnings[0].contains("existing comment targets were retained"));
            assert_eq!(comments, before);
        }
    }

    #[test]
    fn a_comment_that_also_targets_a_non_text_object_is_not_partially_replaced() {
        let (model, mut comments) = reviewed_control();
        let mut object = comments[0].targets[0].clone();
        object.source.start = None;
        object.source.end = None;
        comments[0].targets.push(object);
        let before = comments.clone();
        let warnings = restore_sdt_comment_targets(&model, &mut comments);
        assert_eq!(warnings.len(), 1, "{warnings:?}");
        assert!(warnings[0].contains("non-text object"));
        assert_eq!(comments, before);
    }
}
