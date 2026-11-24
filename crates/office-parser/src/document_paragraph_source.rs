//! Preserve paragraph and ordinary-run properties in every editable block host.
//! The current normalized style remains authoritative; unmodeled pPr/rPr and
//! named-style dependencies stay available to the shared DOCX source adapter.
use std::{collections::{BTreeMap, HashMap}, io::Cursor};

use duke_pages::{BlockDetails, DocModel, DocumentBlock, StructuredDocument};
use serde_json::{Value, json};
use zip::ZipArchive;

use crate::document_format_sidecar::{Element, parse_xml, read_xml};
use crate::document_story_source::{SourceAddress, collect, stories};

fn scope(node: &Element, inherited: &BTreeMap<String, String>) -> BTreeMap<String, String> {
    let mut result = inherited.clone();
    result.extend(node.namespaces.iter().map(|(prefix, namespace)| (prefix.clone(), namespace.clone())));
    result
}

pub(crate) fn canonical_scope(source: &DocModel) -> Result<BTreeMap<String, String>, String> {
    let mut result = source.metadata.document_open_tag.as_deref().map(crate::document_native_xml::opening_namespaces).transpose()?.unwrap_or_default();
    for (prefix, namespace) in crate::document_native_xml::standard_word_namespaces() { result.entry(prefix).or_insert(namespace); }
    Ok(result)
}

fn property(node: Option<&Element>, tag: &str, xml: &str, style_id: Option<&str>, inherited: &BTreeMap<String, String>) -> Value {
    let mut value = json!({ "version": 1, "propertiesXml": node.map(|node| xml[node.start..node.end].to_owned()).unwrap_or_else(|| format!("<w:{tag}/>")),
        "namespaces": node.map(|node| scope(node, inherited)).unwrap_or_else(|| inherited.clone()) });
    if let Some(style_id) = style_id { value["paragraphStyleId"] = json!(style_id); }
    value
}

pub(crate) fn restore(document: &mut StructuredDocument, source: &DocModel, mut archive: Option<&mut ZipArchive<Cursor<&[u8]>>>) -> Result<Vec<String>, String> {
    let canonical = canonical_scope(source)?;
    let sources = stories(source).iter().flat_map(|story| collect(story).paragraphs).collect::<HashMap<_, _>>();
    let mut scopes = HashMap::<String, BTreeMap<String, String>>::new();
    for address in sources.keys() {
        if scopes.contains_key(&address.part) { continue; }
        let mut namespaces = if let Some(archive) = archive.as_mut() {
            read_xml(archive, &address.part)?.map(|xml| parse_xml(&xml).map(|root| root.namespaces.into_iter().collect::<BTreeMap<_, _>>())).transpose()?.unwrap_or_default()
        } else { BTreeMap::new() };
        // Extracted XML uses Duke's canonical vocabulary. Extra bindings belong
        // to its real owning part, while the canonical vocabulary wins clashes.
        namespaces.extend(canonical.iter().map(|(prefix, namespace)| (prefix.clone(), namespace.clone())));
        scopes.insert(address.part.clone(), namespaces);
    }
    fn visit(blocks: &mut [DocumentBlock], inherited_note: Option<i64>, sources: &HashMap<SourceAddress, &duke_pages::ParagraphNode>,
        scopes: &HashMap<String, BTreeMap<String, String>>, warnings: &mut Vec<String>) {
        for block in blocks {
            let note_id = block.note_id.or(inherited_note);
            let paragraph = matches!(&block.details, BlockDetails::Paragraph(_))
                || matches!(&block.details, BlockDetails::Raw(details) if details["kind"].as_str() == Some("paragraph"));
            if paragraph {
                if let Some((address, source)) = block.source_refs.iter().find_map(|reference| {
                    let address = SourceAddress::from_ref(reference, note_id);
                    sources.get(&address).map(|source| (address, *source))
                }) {
                    if let Some(xml) = source.source_xml.as_deref() {
                        match parse_xml(xml) {
                            Ok(root) => {
                                let inherited = scopes.get(&address.part).cloned().unwrap_or_default();
                                let namespaces = scope(&root, &inherited);
                                let properties = root.child("pPr");
                                let style_id = properties.and_then(|properties| properties.child("pStyle")).and_then(|style| style.attribute("val"))
                                    .or_else(|| source.style.as_ref().and_then(|style| style.style_id.as_deref()));
                                let mut details = block.details.to_json_value();
                                details["sourceFormat"] = property(properties, "pPr", xml, style_id, &namespaces);
                                let mark_scope = properties.map(|properties| scope(properties, &namespaces)).unwrap_or_else(|| namespaces.clone());
                                details["paragraphMarkSourceFormat"] = property(properties.and_then(|properties| properties.child("rPr")), "rPr", xml, style_id, &mark_scope);
                                if !xml.contains("fldChar") && !xml.contains("fldSimple") {
                                    if let Err(error) = crate::document_text_fields::restore_run_formats(&mut details, xml, &block.content, &namespaces) {
                                        warnings.push(format!("DOCX run provenance for {} could not be mapped: {error}; paragraph source was retained.", block.id));
                                    }
                                }
                                block.details = BlockDetails::Raw(details);
                            }
                            Err(error) => warnings.push(format!("DOCX paragraph source for {} could not be retained: {error}", block.id)),
                        }
                    }
                }
            }
            visit(&mut block.children, note_id, sources, scopes, warnings);
        }
    }
    let mut warnings = Vec::new();
    visit(&mut document.blocks, None, &sources, &scopes, &mut warnings);
    Ok(warnings)
}
