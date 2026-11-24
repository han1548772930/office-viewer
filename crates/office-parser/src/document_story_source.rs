//! Borrow source nodes with their real package/story owner. Local paragraph IDs
//! and note indices may repeat in other parts; they must never share a lookup.
use std::collections::HashSet;

use duke_pages::{DocModel, DocNode, ParagraphNode, SourceRef, TableCellContentNode, TableNode};

#[derive(Clone, Debug, Hash, PartialEq, Eq)]
pub(crate) struct SourceAddress {
    pub part: String,
    pub story: String,
    pub note_id: Option<i64>,
    pub node_id: String,
}

impl SourceAddress {
    pub fn from_ref(source: &SourceRef, note_id: Option<i64>) -> Self {
        Self { part: source.part_name.clone(), story: source.story.clone(), note_id, node_id: source.node_id.clone() }
    }
}

pub(crate) struct SourceStory<'a> {
    pub part: &'a str,
    pub story: &'static str,
    pub note_id: Option<i64>,
    pub parent: Option<String>,
    pub nodes: &'a [DocNode],
}

impl SourceStory<'_> {
    fn address(&self, id: String) -> SourceAddress {
        SourceAddress { part: self.part.into(), story: self.story.into(), note_id: self.note_id, node_id: id }
    }
}

fn part<'a>(model: &'a DocModel, key: &str, fallback: &'a str) -> &'a str {
    model.metadata.source_part_names.as_ref().and_then(|names| names.get(key)).map(String::as_str).unwrap_or(fallback)
}

pub(crate) fn stories(source: &DocModel) -> Vec<SourceStory<'_>> {
    let mut result = vec![SourceStory { part: part(source, "document", "word/document.xml"), story: "body", note_id: None, parent: None, nodes: &source.nodes }];
    let document_id = source.metadata.document_id.as_deref().unwrap_or("unsaved");
    let mut seen = HashSet::new();
    for header in source.metadata.header_sections.iter().chain(source.metadata.sections.iter().flatten().flat_map(|section| &section.header_sections)) {
        if seen.insert(header.part_name.as_str()) {
            result.push(SourceStory { part: part(source, &header.part_name, &header.part_name), story: "header", note_id: None,
                parent: Some(format!("docx:{document_id}/{}/story", header.part_name)), nodes: &header.nodes });
        }
    }
    for footer in source.metadata.footer_sections.iter().chain(source.metadata.sections.iter().flatten().flat_map(|section| &section.footer_sections)) {
        if seen.insert(footer.part_name.as_str()) {
            result.push(SourceStory { part: part(source, &footer.part_name, &footer.part_name), story: "footer", note_id: None,
                parent: Some(format!("docx:{document_id}/{}/story", footer.part_name)), nodes: &footer.nodes });
        }
    }
    for (story, collection, fallback, notes) in [("footnote", "footnotes", "word/footnotes.xml", &source.metadata.footnotes),
        ("endnote", "endnotes", "word/endnotes.xml", &source.metadata.endnotes)] {
        for note in notes.iter().flatten() {
            if let Some(nodes) = &note.nodes {
                result.push(SourceStory { part: part(source, collection, fallback), story, note_id: Some(note.id), parent: None, nodes });
            }
        }
    }
    for comment in source.metadata.comments.iter().flatten() {
        if let Some(nodes) = &comment.nodes {
            result.push(SourceStory { part: part(source, "comments", "word/comments.xml"), story: "comment", note_id: None, parent: None, nodes });
        }
    }
    result
}

pub(crate) struct SourceNodes<'a> {
    pub paragraphs: Vec<(SourceAddress, &'a ParagraphNode)>,
    pub tables: Vec<(SourceAddress, &'a TableNode)>,
}

pub(crate) fn collect<'a>(story: &SourceStory<'a>) -> SourceNodes<'a> {
    fn table<'a>(source: &'a TableNode, fallback: &str, story: &SourceStory<'a>, output: &mut SourceNodes<'a>) {
        let id = source.block_id.as_deref().unwrap_or(fallback);
        output.tables.push((story.address(id.into()), source));
        for (row_index, row) in source.rows.iter().enumerate() {
            for (cell_index, cell) in row.cells.iter().enumerate() {
                for (index, node) in cell.nodes.iter().enumerate() {
                    let fallback = format!("{id}/row/{row_index}/cell/{cell_index}/{index}");
                    match node {
                        TableCellContentNode::Paragraph(source) => output.paragraphs.push((story.address(source.block_id.clone().unwrap_or(fallback)), source)),
                        TableCellContentNode::Table(source) => table(source, &fallback, story, output),
                    }
                }
            }
        }
    }
    let mut result = SourceNodes { paragraphs: Vec::new(), tables: Vec::new() };
    for (index, node) in story.nodes.iter().enumerate() {
        let fallback = format!("{}/{index}", story.parent.as_deref().unwrap_or("document"));
        match node {
            DocNode::Paragraph(source) => result.paragraphs.push((story.address(source.block_id.clone().unwrap_or(fallback)), source)),
            DocNode::Table(source) => table(source, &fallback, story, &mut result),
        }
    }
    result
}

/// Visit every physical copy. Duke deduplicates references at extraction, but
/// temporary changes and their cleanup must also cover section-local copies.
pub(crate) fn visit_nodes_mut(source: &mut DocModel, visit: &mut impl FnMut(&mut [DocNode])) {
    visit(&mut source.nodes);
    for header in &mut source.metadata.header_sections { visit(&mut header.nodes); }
    for footer in &mut source.metadata.footer_sections { visit(&mut footer.nodes); }
    for section in source.metadata.sections.iter_mut().flatten() {
        for header in &mut section.header_sections { visit(&mut header.nodes); }
        for footer in &mut section.footer_sections { visit(&mut footer.nodes); }
    }
    for notes in [&mut source.metadata.footnotes, &mut source.metadata.endnotes] {
        for note in notes.iter_mut().flatten() { if let Some(nodes) = &mut note.nodes { visit(nodes); } }
    }
    for comment in source.metadata.comments.iter_mut().flatten() { if let Some(nodes) = &mut comment.nodes { visit(nodes); } }
}
