//! A bounded, namespace-resolved representation of standard AutoFilter XML.
//! Cell contents and worksheet axes are outside this projection. The fingerprint
//! detects stale typed criteria; it is not a security signature.

use std::collections::BTreeMap;
use std::io::BufRead;

use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::{NsReader, XmlVersion};
use serde::{Deserialize, Serialize};

const MAIN_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_NS: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const MAX_FILTER_BYTES: usize = 16 * 1024 * 1024;
const MAX_FILTER_NODES: usize = 200_000;

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct OfficeFilterNode {
    pub name: String,
    pub namespace: String,
    pub attributes: BTreeMap<String, String>,
    pub children: Vec<OfficeFilterNode>,
    pub text: String,
}

impl OfficeFilterNode {
    pub fn is_standard(&self, name: &str) -> bool {
        self.name == name && matches!(self.namespace.as_str(), MAIN_NS | STRICT_NS)
    }
}

#[derive(Debug, Default)]
pub struct WorksheetFilterProjection {
    stack: Vec<OfficeFilterNode>,
    root: Option<OfficeFilterNode>,
    byte_count: usize,
    node_count: usize,
}

fn namespace_value(namespace: ResolveResult<'_>) -> Result<String, String> {
    match namespace {
        ResolveResult::Bound(value) => std::str::from_utf8(value.as_ref()).map(str::to_owned).map_err(|error| error.to_string()),
        ResolveResult::Unbound => Ok(String::new()),
        ResolveResult::Unknown(_) => Err("unbound namespace prefix in AutoFilter XML".into()),
    }
}

impl WorksheetFilterProjection {
    /// The caller identifies worksheet/autoFilter using the resolved namespace
    /// and parent path. Foreign extension elements with that name are ignored.
    pub fn start<R: BufRead>(&mut self, reader: &NsReader<R>, element: &BytesStart<'_>, is_root: bool, empty: bool) -> Result<(), String> {
        if self.stack.is_empty() && !is_root { return Ok(()); }
        if is_root && (self.root.is_some() || !self.stack.is_empty()) { return Err("duplicate worksheet AutoFilter".into()); }
        if self.stack.len() >= 64 { return Err("AutoFilter nesting exceeds the parser budget".into()); }
        self.node_count += 1;
        if self.node_count > MAX_FILTER_NODES { return Err("AutoFilter node count exceeds the parser budget".into()); }
        let name = std::str::from_utf8(element.local_name().as_ref()).map_err(|error| error.to_string())?.to_owned();
        let namespace = namespace_value(reader.resolver().resolve_element(element.name()).0)?;
        let mut attributes = BTreeMap::new();
        self.reserve(name.len().saturating_add(namespace.len()))?;
        for attribute in element.attributes() {
            let attribute = attribute.map_err(|error| error.to_string())?;
            if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") { continue; }
            let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
            let namespace = namespace_value(namespace)?;
            let local = std::str::from_utf8(local.as_ref()).map_err(|error| error.to_string())?;
            let key = if namespace.is_empty() { local.to_owned() } else { format!("{{{namespace}}}{local}") };
            let value = attribute.decode_and_unescape_value(reader.decoder(), XmlVersion::Implicit1_0).map_err(|error| error.to_string())?.into_owned();
            self.reserve(key.len().saturating_add(value.len()))?;
            if attributes.insert(key, value).is_some() { return Err("duplicate resolved AutoFilter attribute".into()); }
        }
        let node = OfficeFilterNode { name, namespace, attributes, children: Vec::new(), text: String::new() };
        if empty { self.attach(node); } else { self.stack.push(node); }
        Ok(())
    }

    pub fn text(&mut self, text: &str) -> Result<(), String> {
        if self.stack.is_empty() || text.trim().is_empty() { return Ok(()); }
        self.reserve(text.len())?;
        self.stack.last_mut().unwrap().text.push_str(text);
        Ok(())
    }

    pub fn end(&mut self) {
        if let Some(node) = self.stack.pop() { self.attach(node); }
    }

    fn attach(&mut self, node: OfficeFilterNode) {
        if let Some(parent) = self.stack.last_mut() { parent.children.push(node); } else { self.root = Some(node); }
    }

    fn reserve(&mut self, bytes: usize) -> Result<(), String> {
        self.byte_count = self.byte_count.saturating_add(bytes);
        if self.byte_count > MAX_FILTER_BYTES { return Err("AutoFilter XML exceeds the parser budget".into()); }
        Ok(())
    }

    pub fn root(&self) -> Option<&OfficeFilterNode> { self.root.as_ref() }
    pub fn take_root(&mut self) -> Option<OfficeFilterNode> { self.root.take() }
    pub fn fingerprint(&self) -> Result<String, String> {
        if !self.stack.is_empty() { return Err("incomplete AutoFilter XML".into()); }
        let mut canonical = self.root.clone();
        if let Some(node) = &mut canonical { canonicalize_node(node); }
        let bytes = serde_json::to_vec(&canonical).map_err(|error| error.to_string())?;
        let mut hash = 0xcbf29ce484222325u64;
        for byte in bytes { hash ^= u64::from(byte); hash = hash.wrapping_mul(0x100000001b3); }
        Ok(format!("{hash:016x}"))
    }
}

fn canonicalize_node(node: &mut OfficeFilterNode) {
    if matches!(node.namespace.as_str(), MAIN_NS | STRICT_NS) {
        node.namespace = MAIN_NS.into();
        for (key, default) in [("blank", false), ("and", false), ("hiddenButton", false), ("showButton", true)] {
            let Some(value) = node.attributes.get(key) else { continue; };
            // An unsupported flag is retained verbatim. The typed importer can
            // reject it without discarding the complete standard source tree.
            let enabled = match value.as_str() { "1" | "true" => true, "0" | "false" => false, _ => continue };
            if enabled == default { node.attributes.remove(key); } else { node.attributes.insert(key.into(), if enabled { "1" } else { "0" }.into()); }
        }
        if node.name == "filterColumn" {
            if let Some(value) = node.attributes.get("colId") {
                if let Ok(value) = value.parse::<u32>() { node.attributes.insert("colId".into(), value.to_string()); }
            }
        }
    }
    for child in &mut node.children { canonicalize_node(child); }
}

/// Accepts the worksheet prefix produced before its extLst and closing tag.
pub fn worksheet_filter_fingerprint(xml: &str) -> Result<String, String> {
    let mut reader = NsReader::from_str(xml);
    let mut projection = WorksheetFilterProjection::default();
    let mut path: Vec<Vec<u8>> = Vec::new();
    loop {
        let event = reader.read_event().map_err(|error| error.to_string())?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                if path.len() >= 256 { return Err("worksheet nesting exceeds the parser budget".into()); }
                let name = element.local_name();
                let namespace = namespace_value(reader.resolver().resolve_element(element.name()).0)?;
                let root = path.len() == 1 && path[0] == b"worksheet" && name.as_ref() == b"autoFilter" && matches!(namespace.as_str(), MAIN_NS | STRICT_NS);
                projection.start(&reader, element, root, matches!(&event, Event::Empty(_)))?;
                if matches!(&event, Event::Start(_)) { path.push(name.as_ref().to_vec()); }
            }
            Event::End(_) => { projection.end(); path.pop().ok_or_else(|| "unmatched worksheet closing tag".to_owned())?; }
            Event::Text(value) => projection.text(&value.xml_content(XmlVersion::Implicit1_0).map_err(|error| error.to_string())?)?,
            Event::CData(value) => projection.text(&value.xml_content(XmlVersion::Implicit1_0).map_err(|error| error.to_string())?)?,
            Event::GeneralRef(value) => {
                if let Some(character) = value.resolve_char_ref().map_err(|error| error.to_string())? { projection.text(&character.to_string())?; }
                else {
                    let name = value.decode().map_err(|error| error.to_string())?;
                    projection.text(quick_xml::escape::resolve_xml_entity(&name).ok_or_else(|| "unknown AutoFilter XML entity".to_owned())?)?;
                }
            }
            Event::DocType(_) => return Err("worksheet DTDs are unsupported".into()),
            Event::Eof => break,
            _ => {}
        }
    }
    projection.fingerprint()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn namespace_prefixes_attribute_order_false_defaults_and_cell_edits_do_not_stale_typed_criteria() {
        let first = format!(r#"<worksheet xmlns="{MAIN_NS}"><sheetData><row r="2" hidden="1"/></sheetData><autoFilter ref="A1:C9"><filterColumn colId="1"><filters blank="0"><filter val="A &amp; B"/></filters></filterColumn></autoFilter>"#);
        let rewritten = format!(r#"<s:worksheet xmlns:s="{MAIN_NS}"><s:sheetData><s:row r="2"><s:c><s:v>99</s:v></s:c></s:row></s:sheetData><s:autoFilter ref="A1:C9"><s:filterColumn showButton="true" hiddenButton="false" colId="01"><s:filters><s:filter val="A &amp; B"/></s:filters></s:filterColumn></s:autoFilter></s:worksheet>"#);
        assert_eq!(worksheet_filter_fingerprint(&first).unwrap(), worksheet_filter_fingerprint(&rewritten).unwrap());
        assert_ne!(worksheet_filter_fingerprint(&first).unwrap(), worksheet_filter_fingerprint(&first.replace("A &amp; B", "C")).unwrap());
        assert_ne!(worksheet_filter_fingerprint(&first).unwrap(), worksheet_filter_fingerprint(&first.replace("A1:C9", "A1:D9")).unwrap());
    }
}
