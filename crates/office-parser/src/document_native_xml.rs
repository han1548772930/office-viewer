//! Pair native expanded names with bounded byte offsets and inherited scopes.
use std::collections::BTreeMap;

use office_io_common::{chart_xml::Node, document_content_controls::word};
use quick_xml::{Reader, events::Event};

use crate::document_format_sidecar::Element;

pub(crate) fn opening_namespaces(xml: &str) -> Result<BTreeMap<String, String>, String> {
    let mut reader = Reader::from_str(xml);
    loop {
        match reader.read_event().map_err(|error| error.to_string())? {
            Event::Start(tag) => {
                let name = String::from_utf8_lossy(tag.name().as_ref()).into_owned();
                let end = reader.buffer_position() as usize;
                return crate::document_format_sidecar::parse_xml(&format!("{}</{name}>", &xml[..end]))
                    .map(|root| root.namespaces.into_iter().collect());
            }
            Event::Empty(_) => return crate::document_format_sidecar::parse_xml(xml).map(|root| root.namespaces.into_iter().collect()),
            Event::Eof => return Err("Document namespace context has no root.".into()),
            Event::DocType(_) => return Err("Document namespace context declares a document type.".into()),
            _ => {},
        }
    }
}

/// Duke 0.6 canonicalizes this standard vocabulary in its source fragments.
pub(crate) fn standard_word_namespaces() -> BTreeMap<String, String> {
    [
        ("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"),
        ("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"),
        ("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"),
        ("a", "http://schemas.openxmlformats.org/drawingml/2006/main"),
        ("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture"),
        ("c", "http://schemas.openxmlformats.org/drawingml/2006/chart"),
        ("m", "http://schemas.openxmlformats.org/officeDocument/2006/math"),
        ("w14", "http://schemas.microsoft.com/office/word/2010/wordml"),
        ("w15", "http://schemas.microsoft.com/office/word/2012/wordml"),
        ("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006"),
        ("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"),
        ("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"),
        ("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"),
        ("dgm", "http://schemas.openxmlformats.org/drawingml/2006/diagram"),
        ("dsp", "http://schemas.microsoft.com/office/drawing/2008/diagram"),
        ("a14", "http://schemas.microsoft.com/office/drawing/2010/main"),
        ("asvg", "http://schemas.microsoft.com/office/drawing/2016/SVG/main"),
        ("am3d", "http://schemas.microsoft.com/office/drawing/2017/model3d"),
        ("v", "urn:schemas-microsoft-com:vml"),
        ("o", "urn:schemas-microsoft-com:office:office"),
        ("w10", "urn:schemas-microsoft-com:office:word"),
    ].into_iter().map(|(prefix, namespace)| (prefix.into(), namespace.into())).collect()
}

pub(crate) fn scope(node: &Element, inherited: &BTreeMap<String, String>) -> BTreeMap<String, String> {
    let mut scope = inherited.clone();
    scope.extend(node.namespaces.iter().map(|(prefix, namespace)| (prefix.clone(), namespace.clone())));
    scope
}

pub(crate) fn content_children<'a>(source: &'a Element, native: &'a Node, name: &str,
    inherited: &BTreeMap<String, String>) -> Vec<(&'a Element, &'a Node, BTreeMap<String, String>)> {
    fn visit<'a>(source: &'a Element, native: &'a Node, name: &str, inherited: &BTreeMap<String, String>,
        result: &mut Vec<(&'a Element, &'a Node, BTreeMap<String, String>)>) {
        for (source, native) in source.children.iter().zip(&native.children) {
            let namespaces = scope(source, inherited);
            if word(native, name) { result.push((source, native, namespaces)); }
            else if word(native, "sdt") {
                if let Some((content, node)) = source.children.iter().zip(&native.children).find(|(_, node)| word(node, "sdtContent")) {
                    visit(content, node, name, &scope(content, &namespaces), result);
                }
            } else if ["sdtContent", "customXml", "ins", "moveTo"].iter().any(|tag| word(native, tag)) {
                visit(source, native, name, &namespaces, result);
            }
        }
    }
    let mut result = Vec::new();
    visit(source, native, name, inherited, &mut result);
    result
}

pub(crate) fn complete_fragment(xml: &str, node: &Element, namespaces: &BTreeMap<String, String>) -> Result<String, String> {
    let fragment = xml.get(node.start..node.end).ok_or("Native XML offsets exceed their owning part")?;
    let mut reader = Reader::from_str(fragment);
    let insertion = loop {
        match reader.read_event().map_err(|error| error.to_string())? {
            Event::Start(_) => break (reader.buffer_position() as usize).checked_sub(1).ok_or("Native XML opening is invalid")?,
            Event::Empty(_) => break (reader.buffer_position() as usize).checked_sub(2).ok_or("Native XML empty opening is invalid")?,
            Event::Eof => return Err("Native XML fragment has no element".into()),
            _ => {}
        }
    };
    let declarations = namespaces.iter().filter(|(prefix, _)| !node.namespaces.contains_key(*prefix)).map(|(prefix, namespace)| {
        let name = if prefix.is_empty() { "xmlns".into() } else { format!("xmlns:{prefix}") };
        format!(" {name}=\"{}\"", quick_xml::escape::escape(namespace))
    }).collect::<String>();
    let result = format!("{}{declarations}{}", &fragment[..insertion], &fragment[insertion..]);
    if result.len() > office_io_common::chart_xml::MAX_XML_BYTES { return Err("DOCX_SDT_RESOURCE_LIMIT: namespace-complete source exceeds 16 MiB".into()); }
    Ok(result)
}
