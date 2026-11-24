//! Native SDTs are structural wrappers. Expand their literal rich runs only
//! during extraction; preserve the source model's fields, XML and identities.
use std::{cell::RefCell, collections::{BTreeMap, HashMap, HashSet}};

use duke_pages::{CommentRange, DocModel, DocNode, OoxmlPackage, ParagraphBookmark, ParagraphChildNode, ParagraphNode,
    TableCellContentNode, TableNode, parse::{ParseContext, ParsedStyleSheet, parse_content_types, parse_paragraph, relationship_part_name_for_part, resolve_part_path}};

use crate::document_format_sidecar::{Element, parse_xml};
use crate::document_native_xml::{complete_fragment, opening_namespaces, scope, standard_word_namespaces};
use office_io_common::{chart_xml::{self, Node}, document_content_controls::{self as controls, word, word_attribute}};

pub(crate) struct ParagraphProjectionSource {
    children: Vec<ParagraphChildNode>,
    bookmarks: Vec<ParagraphBookmark>,
    comment_ranges: Option<Vec<CommentRange>>,
}
pub(crate) type ProjectionSources = HashMap<usize, ParagraphProjectionSource>;

fn ordinary_sdt(node: &Node) -> bool {
    word(node, "sdt") && !node.children.iter().find(|node| word(node, "sdtPr"))
        .and_then(|properties| properties.children.iter().find(|node| word(node, "tag")))
        .and_then(|tag| word_attribute(tag, "val")).is_some_and(controls::internal_feature_tag)
}

/// Duke supplies canonical Word prefixes in source nodes. Strip only ordinary
/// Word wrappers; formulas/code and unknown namespaces retain their source XML.
fn expand(xml: &str, source: &Element, node: &Node, inherited: &BTreeMap<String, String>, detached: bool, output: &mut String) -> Result<(), String> {
    let namespaces = scope(source, inherited);
    if ordinary_sdt(node) {
        let (content, native) = source.children.iter().zip(&node.children).find(|(_, node)| word(node, "sdtContent"))
            .ok_or("Native SDT has no Word content element")?;
        let namespaces = scope(content, &namespaces);
        for (child, native) in content.children.iter().zip(&native.children) { expand(xml, child, native, &namespaces, true, output)?; }
        return Ok(());
    }
    let start = output.len();
    // Internal feature SDTs keep their own typed FormField representation.
    if word(node, "sdt") {
        output.push_str(xml.get(source.start..source.end).ok_or("SDT source offsets exceed their paragraph")?);
    } else {
        let mut cursor = source.start;
        for (child, native) in source.children.iter().zip(&node.children) {
            output.push_str(xml.get(cursor..child.start).ok_or("SDT source child offsets overlap")?);
            expand(xml, child, native, &namespaces, false, output)?;
            cursor = child.end;
        }
        output.push_str(xml.get(cursor..source.end).ok_or("SDT source offsets exceed their paragraph")?);
    }
    if detached {
        // Declarations on an SDT or sdtContent disappear with their wrapper.
        // Restore their actual scope on each retained content root, including
        // unknown extensions and internal formula/code/lock SDTs.
        let fragment = complete_fragment(&output[start..], &Element { start: 0, end: output.len() - start,
            namespaces: source.namespaces.clone(), ..Default::default() }, &namespaces)?;
        output.truncate(start); output.push_str(&fragment);
    }
    if output.len() > chart_xml::MAX_XML_BYTES { return Err("DOCX_SDT_RESOURCE_LIMIT: projected rich content exceeds 16 MiB".into()); }
    Ok(())
}

pub(crate) fn expand_native_sdt_xml(xml: &str, inherited: &BTreeMap<String, String>) -> Result<String, String> {
    // Source fragments supplied by Duke can inherit standard declarations from
    // their part. Pair expanded names with byte offsets inside the same wrapper.
    let mut namespaces = standard_word_namespaces(); namespaces.extend(inherited.clone());
    let declarations = namespaces.iter().map(|(prefix, namespace)| {
        let name = if prefix.is_empty() { "xmlns".into() } else { format!("xmlns:{prefix}") };
        format!(" {name}=\"{}\"", quick_xml::escape::escape(namespace))
    }).collect::<String>();
    let wrapped = format!("<office{declarations}>{xml}</office>");
    let root = parse_xml(&wrapped)?;
    let native = chart_xml::parse(&wrapped)?;
    if root.children.len() != 1 || native.children.len() != 1 { return Err("Native SDT projection requires one source root".into()); }
    let mut result = String::with_capacity(xml.len());
    expand(&wrapped, &root.children[0], &native.children[0], &scope(&root, &BTreeMap::new()), false, &mut result)?;
    Ok(result)
}

fn visit_table(table: &mut TableNode, visit: &mut impl FnMut(&mut ParagraphNode)) {
    for row in &mut table.rows { for cell in &mut row.cells { for node in &mut cell.nodes {
        match node {
            TableCellContentNode::Paragraph(paragraph) => visit(paragraph),
            TableCellContentNode::Table(table) => visit_table(table, visit),
        }
    } } }
}
fn visit_body(nodes: &mut [DocNode], visit: &mut impl FnMut(&mut ParagraphNode)) {
    for node in nodes { match node {
        DocNode::Paragraph(paragraph) => visit(paragraph),
        DocNode::Table(table) => visit_table(table, visit),
    } }
}

fn paragraph_properties(paragraph: &ParagraphNode) -> Result<String, String> {
    if let Some(xml) = &paragraph.source_xml {
        let root = parse_xml(xml)?;
        if let Some(properties) = root.child("pPr") { return Ok(xml[properties.start..properties.end].into()); }
    }
    Ok(paragraph.style.as_ref().and_then(|style| style.style_id.as_deref())
        .map(|id| format!("<w:pPr><w:pStyle w:val=\"{}\"/></w:pPr>", quick_xml::escape::escape(id))).unwrap_or_default())
}

fn project_paragraph(paragraph: &mut ParagraphNode, context: &ParseContext<'_, &ParsedStyleSheet>, inherited: &BTreeMap<String, String>) -> Result<Option<ParagraphProjectionSource>, String> {
    if !paragraph.children.iter().any(|child| matches!(child, ParagraphChildNode::FormField(field)
        if field.source_kind == Some(duke_pages::FormFieldSourceKind::Sdt)
        && !field.tag.as_deref().is_some_and(office_io_common::document_content_controls::internal_feature_tag))) { return Ok(None); }
    let properties = paragraph_properties(paragraph)?;
    let namespaces = paragraph.source_xml.as_deref().map(|xml| parse_xml(xml).map(|root| scope(&root, inherited)))
        .transpose()?.unwrap_or_else(|| inherited.clone());
    let mut children = Vec::with_capacity(paragraph.children.len());
    let mut boundaries = Vec::with_capacity(paragraph.children.len() + 1);
    let mut inner_bookmarks = Vec::new();
    for child in &paragraph.children {
        boundaries.push(children.len());
        let field = match child {
            ParagraphChildNode::FormField(field) if field.source_kind == Some(duke_pages::FormFieldSourceKind::Sdt)
                && !field.tag.as_deref().is_some_and(office_io_common::document_content_controls::internal_feature_tag) => field,
            _ => { children.push(child.clone()); continue; },
        };
        let xml = field.source_xml.as_deref().ok_or("Native SDT has no source XML for rich-run projection")?;
        let content = expand_native_sdt_xml(xml, &namespaces)?;
        let mut parsed = parse_paragraph(&format!("<w:p>{properties}{content}</w:p>"), context);
        for bookmark in &mut parsed.bookmarks { bookmark.child_index += children.len(); }
        inner_bookmarks.extend(parsed.bookmarks);
        // The enclosing hyperlink lies outside the SDT fragment. Inner explicit
        // links still win and each run keeps its independently resolved style.
        for node in &mut parsed.children { match node {
            ParagraphChildNode::Text(run) if run.link.is_none() => run.link.clone_from(&field.link),
            ParagraphChildNode::FormField(run) if run.link.is_none() => run.link.clone_from(&field.link),
            _ => {},
        } }
        children.extend(parsed.children);
    }
    boundaries.push(children.len());
    let interior_names = inner_bookmarks.iter().map(|bookmark| bookmark.name.as_str()).collect::<HashSet<_>>();
    let mut bookmarks = paragraph.bookmarks.iter().filter(|bookmark| !interior_names.contains(bookmark.name.as_str()))
        .map(|bookmark| boundaries.get(bookmark.child_index).copied().map(|index| ParagraphBookmark { child_index: index, name: bookmark.name.clone() })
            .ok_or_else(|| "SDT bookmark child boundary exceeds source paragraph".to_owned())).collect::<Result<Vec<_>, _>>()?;
    bookmarks.extend(inner_bookmarks);
    bookmarks.sort_by_key(|bookmark| bookmark.child_index);
    let comments = paragraph.comment_ranges.as_ref().map(|comments| comments.iter().map(|comment| {
        let mut mapped = comment.clone();
        mapped.child_start = *boundaries.get(comment.child_start).ok_or("SDT comment start boundary exceeds source paragraph")?;
        mapped.child_end = *boundaries.get(comment.child_end).ok_or("SDT comment end boundary exceeds source paragraph")?;
        Ok(mapped)
    }).collect::<Result<Vec<_>, String>>()).transpose()?;
    Ok(Some(ParagraphProjectionSource { children: std::mem::replace(&mut paragraph.children, children),
        bookmarks: std::mem::replace(&mut paragraph.bookmarks, bookmarks), comment_ranges: std::mem::replace(&mut paragraph.comment_ranges, comments) }))
}

pub(crate) fn project(source: &mut DocModel, package: Option<&OoxmlPackage>, styles: &ParsedStyleSheet) -> (ProjectionSources, Vec<String>) {
    let empty = OoxmlPackage { parts: HashMap::new(), binary_assets: HashMap::new(), warnings: Vec::new() };
    let package = package.unwrap_or(&empty);
    let content_types = parse_content_types(package);
    let mut originals = HashMap::new(); let mut warnings = Vec::new();
    let canonical = match crate::document_paragraph_source::canonical_scope(source) {
        Ok(namespaces) => namespaces,
        Err(error) => return (originals, vec![format!("Native SDT paragraph namespaces could not be read: {error}; source fields remain retained")]),
    };
    let body_part = source.metadata.source_part_names.as_ref().and_then(|names| names.get("document")).cloned()
        .unwrap_or_else(|| "word/document.xml".into());
    let mut contexts = HashMap::new();
    let mut namespaces = HashMap::new();
    crate::document_story_source::visit_parts_mut(source, &mut |part, nodes| {
        if !contexts.contains_key(part) {
            let inherited = package.parts.get(part).map(|source| opening_namespaces(&source.content)).transpose();
            let mut inherited = match inherited {
                Ok(namespaces) => namespaces.unwrap_or_default(),
                Err(error) => { warnings.push(format!("Native SDT namespaces in {part} could not be read: {error}; source fields remain retained")); return; },
            };
            // Only the body's canonical opening can supply its unknown aliases.
            // Other stories keep their own part bindings, plus the standard
            // vocabulary; local paragraph/SDT declarations override both.
            inherited.extend(if part == body_part { canonical.clone() } else { standard_word_namespaces() });
            namespaces.insert(part.to_owned(), inherited);
            // Use the project's expanded-name OPC reader. Duke's public helper
            // scans unprefixed Relationship tags; source parts may legally use
            // a prefix even though their paragraph vocabulary was normalized.
            let relationships = package.parts.get(&relationship_part_name_for_part(part))
                .map(|source| office_io_common::document_custom_xml::relationships(&source.content)).transpose();
            let relationships = match relationships {
                Ok(entries) => entries.into_iter().flatten().map(|entry| {
                    let target = if entry.external { entry.target } else { resolve_part_path(part, &entry.target) };
                    (entry.id, target)
                }).collect(),
                Err(error) => { warnings.push(format!("Native SDT relationships in {part} could not be read: {error}; source fields remain retained")); return; },
            };
            contexts.insert(part.to_owned(), ParseContext { relationships, content_types: content_types.clone(),
                parts: &package.parts, binary_assets: &package.binary_assets, style_sheet: styles, warnings: RefCell::new(Vec::new()) });
        }
        let context = contexts.get(part).expect("SDT story context was inserted");
        let inherited = namespaces.get(part).expect("SDT story namespace scope was inserted");
        visit_body(nodes, &mut |paragraph| match project_paragraph(paragraph, context, inherited) {
            Ok(Some(original)) => { originals.insert(paragraph as *const ParagraphNode as usize, original); },
            Ok(None) => {},
            Err(error) => warnings.push(format!("Native SDT rich content in {part} {:?} could not be projected: {error}; source fields remain retained", paragraph.block_id)),
        });
    });
    for (part, context) in contexts { warnings.extend(context.warnings.into_inner().into_iter().map(|warning| format!("Native SDT content in {part}: {warning}"))); }
    (originals, warnings)
}

pub(crate) fn restore(source: &mut DocModel, mut originals: ProjectionSources) {
    crate::document_story_source::visit_nodes_mut(source, &mut |nodes| visit_body(nodes, &mut |paragraph| {
        if let Some(original) = originals.remove(&(paragraph as *const ParagraphNode as usize)) {
            paragraph.children = original.children; paragraph.bookmarks = original.bookmarks; paragraph.comment_ranges = original.comment_ranges;
        }
    }));
    debug_assert!(originals.is_empty(), "Every physical source paragraph must restore its SDT rich-run projection");
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    // Drafts for unified native/wasm verification, not executed in this batch.
    #[test]
    fn detached_sdt_content_keeps_relationship_drawing_and_extension_namespaces() {
        let mut inherited = standard_word_namespaces(); inherited.insert("x".into(), "urn:owner-extension".into());
        let xml = "<w:sdt xmlns:y='urn:wrapper-extension'><w:sdtPr><w:text/></w:sdtPr><w:sdtContent>
            <w:hyperlink r:id='link'><w:r><w:t>Link</w:t></w:r></w:hyperlink>
            <w:r><w:drawing><wp:inline><a:graphic><a:graphicData><pic:pic x:flag='owner'/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>
            <y:marker x:flag='retained'/></w:sdtContent></w:sdt>";
        let projected = expand_native_sdt_xml(xml, &inherited).unwrap();
        let root = chart_xml::parse(&format!("<office>{projected}</office>")).unwrap();
        let hyperlink = root.descendants("hyperlink");
        assert_eq!(hyperlink[0].attribute_ns("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id"), Some("link"));
        let picture = root.descendants("pic");
        assert_eq!(picture[0].namespace.as_deref(), Some("http://schemas.openxmlformats.org/drawingml/2006/picture"));
        assert_eq!(picture[0].attribute_ns("urn:owner-extension", "flag"), Some("owner"));
        assert_eq!(root.descendants("marker")[0].namespace.as_deref(), Some("urn:wrapper-extension"));
        assert!(root.descendants("sdt").is_empty());
    }

    #[test]
    fn namespace_names_and_internal_feature_tags_distinguish_transparent_user_wrappers() {
        let xml = "<w:sdt xmlns:x='urn:foreign'><w:sdtPr><w:group/></w:sdtPr><w:sdtContent>
            <x:sdt><x:sdtContent><x:t>Foreign</x:t></x:sdtContent></x:sdt>
            <w:sdt><w:sdtPr><w:tag w:val='office-latex:v1:sum'/></w:sdtPr><w:sdtContent><w:r><w:t>Formula</w:t></w:r></w:sdtContent></w:sdt>
            </w:sdtContent></w:sdt>";
        let projected = expand_native_sdt_xml(xml, &BTreeMap::new()).unwrap();
        let root = chart_xml::parse(&format!("<office>{projected}</office>")).unwrap();
        let wrappers = root.descendants("sdt");
        assert_eq!(wrappers.iter().filter(|node| word(node, "sdt")).count(), 1);
        assert_eq!(wrappers.iter().filter(|node| node.namespace.as_deref() == Some("urn:foreign")).count(), 1);
        assert!(projected.contains("office-latex:v1:sum")); assert!(projected.contains("Foreign"));
    }

    #[test]
    fn rich_projection_preserves_styles_links_bookmarks_comment_boundaries_and_restores_the_source() {
        let xml = format!("<w:document xmlns:w='{}' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'>
            <w:body><w:p><w:r><w:t>pre </w:t></w:r><w:sdt><w:sdtPr><w:comboBox/></w:sdtPr><w:sdtContent>
            <w:r><w:rPr><w:b/><w:color w:val='112233'/></w:rPr><w:t xml:space='preserve'> 中😀 </w:t></w:r>
            <w:bookmarkStart w:id='7' w:name='inside'/><w:hyperlink r:id='link'><w:r><w:rPr><w:i/></w:rPr><w:t>Link</w:t></w:r></w:hyperlink>
            <w:bookmarkEnd w:id='7'/></w:sdtContent></w:sdt><w:r><w:t> end</w:t></w:r></w:p></w:body></w:document>", controls::WORD_NAMESPACE);
        let mut package = duke_pages::create_minimal_docx_package(Some(&xml));
        package.parts.get_mut("word/_rels/document.xml.rels").unwrap().content =
            "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='link' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink' Target='https://example.test/inside' TargetMode='External'/></Relationships>".into();
        let mut source = duke_pages::build_doc_model(&package);
        let DocNode::Paragraph(paragraph) = &mut source.nodes[0] else { panic!("fixture must start with a paragraph") };
        assert_eq!(paragraph.children.len(), 3);
        paragraph.bookmarks.push(ParagraphBookmark { name: "after".into(), child_index: 2 });
        paragraph.comment_ranges = Some(vec![CommentRange { comment_id: 5, child_start: 1, child_end: 2, ends_here: true }]);
        let original = serde_json::to_value(&source).unwrap();
        let (retained, warnings) = project(&mut source, Some(&package), &ParsedStyleSheet::empty());
        assert!(warnings.is_empty(), "{warnings:?}"); assert_eq!(retained.len(), 1);
        let DocNode::Paragraph(paragraph) = &source.nodes[0] else { unreachable!() };
        let text = paragraph.children.iter().filter_map(|node| match node { ParagraphChildNode::Text(run) => Some(run.text.as_str()), _ => None }).collect::<String>();
        assert_eq!(text, "pre  中😀 Link end");
        let styled = paragraph.children.iter().filter_map(|node| match node { ParagraphChildNode::Text(run) => Some(run), _ => None }).collect::<Vec<_>>();
        assert_eq!(styled[1].style.as_ref().unwrap().bold, Some(true));
        assert_eq!(styled[1].style.as_ref().unwrap().color.as_deref(), Some("#112233"));
        assert_eq!(styled[2].style.as_ref().unwrap().italic, Some(true));
        assert_eq!(styled[2].link.as_deref(), Some("https://example.test/inside"));
        assert!(paragraph.bookmarks.iter().any(|bookmark| bookmark.name == "inside" && bookmark.child_index == 2));
        assert!(paragraph.bookmarks.iter().any(|bookmark| bookmark.name == "after" && bookmark.child_index == 3));
        assert_eq!(serde_json::to_value(&paragraph.comment_ranges).unwrap(), json!([{ "commentId": 5, "childStart": 1, "childEnd": 3, "endsHere": true }]));
        restore(&mut source, retained);
        assert_eq!(serde_json::to_value(&source).unwrap(), original);
    }
}
