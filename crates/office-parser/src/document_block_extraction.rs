//! Preserve the paragraphs in every story that an Office editor must address.
//! duke-pages 0.6.0 filters empty/invisible paragraphs and flattens object-only
//! paragraphs. A temporary extraction bookmark keeps their existing structure;
//! it is removed from both the source model and result before either is exposed.
use std::collections::HashSet;

use duke_pages::{BlockDetails, DocModel, DocNode, DocumentBlock, ParagraphBookmark, ParagraphChildNode,
    ParagraphNode, StructuredDocument, TableCellContentNode, TableNode, extract_document_blocks};

fn visit_table(table: &mut TableNode, visit: &mut impl FnMut(&mut ParagraphNode)) {
    for row in &mut table.rows {
        for cell in &mut row.cells {
            for node in &mut cell.nodes {
                match node {
                    TableCellContentNode::Paragraph(paragraph) => visit(paragraph),
                    TableCellContentNode::Table(table) => visit_table(table, visit),
                }
            }
        }
    }
}

fn visit_body(nodes: &mut [DocNode], visit: &mut impl FnMut(&mut ParagraphNode)) {
    for node in nodes {
        match node {
            DocNode::Paragraph(paragraph) => visit(paragraph),
            DocNode::Table(table) => visit_table(table, visit),
        }
    }
}

fn visible(text: &str) -> bool {
    text.chars().any(|ch| !ch.is_whitespace() && !matches!(ch, '\u{200b}' | '\u{feff}' | '\u{00ad}'))
}

fn needs_anchor(paragraph: &ParagraphNode) -> bool {
    paragraph.bookmarks.is_empty() && !paragraph.children.iter().any(|child| match child {
        ParagraphChildNode::Text(run) => visible(&run.text) || run.note_reference.is_some(),
        ParagraphChildNode::FormField(field) => field.checked.is_some() || field.value.as_deref().is_some_and(visible),
        ParagraphChildNode::Image(_) => false,
    })
}

fn collect_names(nodes: &[DocNode], used: &mut HashSet<String>) {
    fn paragraph(paragraph: &ParagraphNode, used: &mut HashSet<String>) {
        used.extend(paragraph.bookmarks.iter().map(|bookmark| bookmark.name.clone()));
        for child in &paragraph.children {
            if let ParagraphChildNode::Image(image) = child {
                if let Some(nodes) = &image.content_nodes { collect_names(nodes, used); }
            }
        }
    }
    fn table(node: &TableNode, used: &mut HashSet<String>) {
        for row in &node.rows {
            for cell in &row.cells {
                for node in &cell.nodes {
                    match node {
                        TableCellContentNode::Paragraph(node) => paragraph(node, used),
                        TableCellContentNode::Table(node) => table(node, used),
                    }
                }
            }
        }
    }
    for node in nodes {
        match node {
            DocNode::Paragraph(node) => paragraph(node, used),
            DocNode::Table(node) => table(node, used),
        }
    }
}

fn remove_anchors(blocks: &mut [DocumentBlock], names: &HashSet<String>) {
    for block in blocks {
        match &mut block.details {
            BlockDetails::Paragraph(details) => details.bookmarks.retain(|bookmark| !names.contains(&bookmark.name)),
            BlockDetails::Raw(details) => {
                if let Some(bookmarks) = details.get_mut("bookmarks").and_then(serde_json::Value::as_array_mut) {
                    bookmarks.retain(|bookmark| bookmark["name"].as_str().is_none_or(|name| !names.contains(name)));
                }
            }
            _ => {},
        }
        remove_anchors(&mut block.children, names);
    }
}

pub(crate) fn extract(source: &mut DocModel) -> StructuredDocument {
    extract_with_resources(source, None, &duke_pages::parse::ParsedStyleSheet::empty())
}

pub(crate) fn extract_with_resources(source: &mut DocModel, package: Option<&duke_pages::OoxmlPackage>,
    styles: &duke_pages::parse::ParsedStyleSheet) -> StructuredDocument {
    let (sdt_nodes, warnings) = crate::document_sdt_projection::project(source, package, styles);
    let mut used = HashSet::new();
    collect_names(&source.nodes, &mut used);
    // Extraction returns other stories too. Their real bookmarks must never
    // collide with a temporary body anchor, including section-local stories.
    for header in source.metadata.header_sections.iter().chain(source.metadata.sections.iter().flatten().flat_map(|section| &section.header_sections)) {
        collect_names(&header.nodes, &mut used);
    }
    for footer in source.metadata.footer_sections.iter().chain(source.metadata.sections.iter().flatten().flat_map(|section| &section.footer_sections)) {
        collect_names(&footer.nodes, &mut used);
    }
    for notes in [&source.metadata.footnotes, &source.metadata.endnotes] {
        for note in notes.iter().flatten() {
            if let Some(nodes) = &note.nodes { collect_names(nodes, &mut used); }
        }
    }
    for comment in source.metadata.comments.iter().flatten() {
        if let Some(nodes) = &comment.nodes { collect_names(nodes, &mut used); }
    }
    let mut names = HashSet::new(); let mut sequence = 0usize;
    crate::document_story_source::visit_nodes_mut(source, &mut |nodes| visit_body(nodes, &mut |paragraph| {
        if !needs_anchor(paragraph) { return; }
        let name = loop {
            let name = format!("office_studio_internal_keep_{sequence}"); sequence += 1;
            if used.insert(name.clone()) { break name; }
        };
        names.insert(name.clone());
        paragraph.bookmarks.push(ParagraphBookmark { child_index: 0, name });
    }));
    let mut document = extract_document_blocks(source);
    crate::document_story_source::visit_nodes_mut(source, &mut |nodes| visit_body(nodes, &mut |paragraph| {
        paragraph.bookmarks.retain(|bookmark| !names.contains(&bookmark.name));
    }));
    remove_anchors(&mut document.blocks, &names);
    crate::document_sdt_projection::restore(source, sdt_nodes);
    document.warnings.extend(warnings);
    document
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::{Value, json};

    fn source(nodes: Value) -> DocModel {
        serde_json::from_value(json!({ "nodes": nodes, "metadata": { "sourceParts": 1, "warnings": [],
            "headerSections": [], "footerSections": [], "paragraphStyles": [] } })).unwrap()
    }

    #[test]
    fn empty_and_invisible_paragraphs_keep_source_positions_without_exposing_extraction_anchors() {
        let mut source = source(json!([
            { "type": "paragraph", "blockId": "first", "children": [{ "type": "text", "text": "A" }] },
            { "type": "paragraph", "blockId": "empty", "children": [] },
            { "type": "paragraph", "blockId": "invisible", "children": [{ "type": "text", "text": " \u{200b}\u{feff}\u{00ad}" }] },
            { "type": "paragraph", "blockId": "last", "children": [{ "type": "text", "text": "😀B" }] },
        ]));
        let original = source.clone(); let document = extract(&mut source);
        assert_eq!(source, original);
        assert_eq!(document.blocks.iter().map(|block| block.id.as_str()).collect::<Vec<_>>(), vec!["first", "empty", "invisible", "last"]);
        assert_eq!(document.blocks[1].content, "");
        assert_eq!(document.blocks[2].content, " \u{200b}\u{feff}\u{00ad}");
        assert_eq!(document.blocks[1].source_refs[0].node_id, "empty");
        assert!(!serde_json::to_string(&document).unwrap().contains("office_studio_internal_keep_"));
    }

    #[test]
    fn a_real_bookmark_with_the_extraction_prefix_is_never_removed() {
        let mut source = source(json!([
            { "type": "paragraph", "blockId": "authored", "children": [], "bookmarks": [{ "childIndex": 0, "name": "office_studio_internal_keep_0" }] },
            { "type": "paragraph", "blockId": "empty", "children": [] },
        ]));
        let original = source.clone(); let document = extract(&mut source);
        assert_eq!(source, original);
        assert_eq!(document.blocks.len(), 2);
        assert_eq!(document.blocks[0].details.to_json_value()["bookmarks"][0]["name"], "office_studio_internal_keep_0");
        assert!(document.blocks[1].details.to_json_value()["bookmarks"].as_array().is_none_or(Vec::is_empty));
    }

    #[test]
    fn body_anchors_do_not_remove_authored_bookmarks_from_a_different_story() {
        let mut source = source(json!([{ "type": "paragraph", "blockId": "empty-body", "children": [] }]));
        source.metadata.header_sections = serde_json::from_value(json!([{ "partName": "word/header1.xml", "nodes": [
            { "type": "paragraph", "blockId": "authored-header", "children": [],
                "bookmarks": [{ "childIndex": 0, "name": "office_studio_internal_keep_0" }] }
        ] }])).unwrap();
        let original = source.clone(); let document = extract(&mut source);
        assert_eq!(source, original);
        fn find<'a>(blocks: &'a [DocumentBlock], id: &str) -> Option<&'a DocumentBlock> {
            blocks.iter().find_map(|block| if block.id == id { Some(block) } else { find(&block.children, id) })
        }
        let header = find(&document.blocks, "authored-header").unwrap();
        assert_eq!(header.details.to_json_value()["bookmarks"][0]["name"], "office_studio_internal_keep_0");
        assert_eq!(find(&document.blocks, "empty-body").unwrap().content, "");
    }

    #[test]
    fn empty_table_cell_paragraphs_keep_their_parent_and_do_not_disappear_from_native_field_mapping() {
        let mut source = source(json!([{ "type": "table", "blockId": "table", "rows": [
            { "cells": [{ "nodes": [{ "type": "paragraph", "blockId": "cell-empty", "children": [] }] }] }
        ] }]));
        let original = source.clone(); let document = extract(&mut source);
        fn find<'a>(blocks: &'a [DocumentBlock], id: &str) -> Option<&'a DocumentBlock> {
            blocks.iter().find_map(|block| if block.id == id { Some(block) } else { find(&block.children, id) })
        }
        assert_eq!(source, original);
        let empty = find(&document.blocks, "cell-empty").unwrap();
        assert!(empty.parent_block_id.is_some());
        assert_eq!(empty.content, "");
        assert_eq!(empty.details.to_json_value()["kind"], "paragraph");
    }
}
