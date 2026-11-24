//! Bounded, namespace-aware native-chart XML and authored-source freshness.

use std::collections::BTreeMap;
use quick_xml::{NsReader, XmlVersion, events::{BytesStart, Event}, name::ResolveResult};
use crate::chart_options::{SOURCE_EXTENSION_NS, SOURCE_EXTENSION_URI};

pub const CHART_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
pub const DRAWING_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
pub const MAX_XML_BYTES: usize = 16 * 1024 * 1024;
const MAX_NODES: usize = 300_000;
const MAX_DEPTH: usize = 128;

#[derive(Debug, Default, Clone)]
pub struct Node {
    pub name: String,
    pub namespace: Option<String>,
    pub attrs: BTreeMap<String, String>,
    pub attr_namespaces: BTreeMap<String, Option<String>>,
    /// Expanded names retain distinct extension attributes sharing a local name.
    pub expanded_attrs: BTreeMap<String, BTreeMap<String, String>>,
    pub text: String,
    pub children: Vec<Node>,
}

impl Node {
    pub fn attribute_ns(&self, namespace: &str, name: &str) -> Option<&str> {
        self.expanded_attrs.get(namespace).and_then(|attributes| attributes.get(name)).map(String::as_str)
    }
    pub fn child(&self, name: &str) -> Option<&Self> { self.children.iter().find(|child| child.name == name) }
    pub fn children_named<'a>(&'a self, name: &'a str) -> impl Iterator<Item = &'a Self> { self.children.iter().filter(move |child| child.name == name) }
    pub fn value(&self, name: &str) -> Option<&str> { self.child(name)?.attrs.get("val").map(String::as_str) }
    pub fn descendants<'a>(&'a self, name: &str) -> Vec<&'a Self> {
        let mut out = Vec::new();
        let mut stack = vec![self];
        while let Some(node) = stack.pop() {
            if node.name == name { out.push(node); }
            stack.extend(node.children.iter().rev());
        }
        out
    }
}

fn element(reader: &NsReader<&[u8]>, value: &BytesStart<'_>, namespace: Option<String>) -> Result<Node, String> {
    let name = String::from_utf8_lossy(value.local_name().as_ref()).into_owned();
    let mut attrs = BTreeMap::new();
    let mut attr_namespaces = BTreeMap::new();
    let mut expanded_attrs = BTreeMap::<String, BTreeMap<String, String>>::new();
    for attr in value.attributes() {
        let attr = attr.map_err(|error| error.to_string())?;
        if attr.key.as_ref() == b"xmlns" || attr.key.as_ref().starts_with(b"xmlns:") { continue; }
        let key = String::from_utf8_lossy(attr.key.local_name().as_ref()).into_owned();
        let (resolved, _) = reader.resolver().resolve_attribute(attr.key);
        let attr_namespace = match resolved {
            ResolveResult::Bound(value) => Some(String::from_utf8_lossy(value.as_ref()).into_owned()),
            ResolveResult::Unbound => None,
            ResolveResult::Unknown(value) => return Err(format!("unbound native XML attribute prefix {}", String::from_utf8_lossy(&value))),
        };
        let decoded = attr.decoded_and_normalized_value_with(XmlVersion::Implicit1_0, reader.decoder(), 1, quick_xml::escape::resolve_xml_entity)
            .map_err(|error| error.to_string())?.into_owned();
        if expanded_attrs.entry(attr_namespace.clone().unwrap_or_default()).or_default().insert(key.clone(), decoded.clone()).is_some() {
            return Err("duplicate native XML attribute expanded name".into());
        }
        // Chart attributes normally have no namespace; Word properties use the
        // Word namespace. Keep their convenience lookup while preserving all
        // other expanded names for ownership and native property signatures.
        let rank = |namespace: Option<&str>| match namespace {
            None => 3, Some("http://schemas.openxmlformats.org/wordprocessingml/2006/main" | "http://purl.oclc.org/ooxml/wordprocessingml/main") => 2, Some(_) => 1,
        };
        if !attrs.contains_key(&key) || rank(attr_namespace.as_deref()) > rank(attr_namespaces.get(&key).and_then(Option::as_deref)) {
            attr_namespaces.insert(key.clone(), attr_namespace);
            attrs.insert(key, decoded);
        }
    }
    Ok(Node { name, namespace, attrs, attr_namespaces, expanded_attrs, ..Node::default() })
}

pub fn parse(source: &str) -> Result<Node, String> {
    parse_with_reserved_nodes(source, 0)
}

/// A writer can reserve the three manifest element nodes before computing a
/// signature, so its final XML stays within the reader's identical tree budget.
pub fn parse_with_reserved_nodes(source: &str, reserved_nodes: usize) -> Result<Node, String> {
    let node_limit = MAX_NODES.checked_sub(reserved_nodes).ok_or("CHART_XML_RESOURCE_LIMIT: reserved chart nodes exceed the tree budget")?;
    if source.len() > MAX_XML_BYTES { return Err("CHART_XML_RESOURCE_LIMIT: native chart part exceeds 16 MiB".into()); }
    let mut reader = NsReader::from_str(source);
    let mut stack = vec![Node::default()];
    let mut count = 0usize;
    loop {
        let (namespace, event) = reader.read_resolved_event().map_err(|error| error.to_string())?;
        let namespace = match namespace {
            ResolveResult::Bound(value) => Some(String::from_utf8_lossy(value.as_ref()).into_owned()),
            ResolveResult::Unbound => None,
            ResolveResult::Unknown(value) => return Err(format!("unbound chart XML prefix {}", String::from_utf8_lossy(&value))),
        };
        match event {
            Event::Start(value) => {
                count += 1;
                if count > node_limit || stack.len() > MAX_DEPTH { return Err("CHART_XML_RESOURCE_LIMIT: chart tree is too large or deep".into()); }
                let node = element(&reader, &value, namespace)?;
                stack.push(node);
            }
            Event::Empty(value) => {
                count += 1;
                if count > node_limit || stack.len() > MAX_DEPTH { return Err("CHART_XML_RESOURCE_LIMIT: chart tree is too large or deep".into()); }
                let node = element(&reader, &value, namespace)?;
                stack.last_mut().ok_or_else(|| "empty chart XML root".to_owned())?.children.push(node);
            }
            Event::Text(value) => {
                if let Some(node) = stack.last_mut() { node.text.push_str(&value.decode().map_err(|error| error.to_string())?); }
            }
            Event::CData(value) => {
                if let Some(node) = stack.last_mut() { node.text.push_str(&value.decode().map_err(|error| error.to_string())?); }
            }
            Event::GeneralRef(value) => {
                let decoded = value.decode().map_err(|error| error.to_string())?;
                let resolved = value.resolve_char_ref().map_err(|error| error.to_string())?.map(|value| value.to_string())
                    .or_else(|| quick_xml::escape::resolve_xml_entity(&decoded).map(str::to_owned))
                    .ok_or_else(|| format!("unknown chart XML entity &{decoded};"))?;
                if let Some(node) = stack.last_mut() { node.text.push_str(&resolved); }
            }
            Event::End(_) => {
                if stack.len() <= 1 { return Err("unbalanced chart XML tree".into()); }
                let node = stack.pop().ok_or_else(|| "empty chart XML stack".to_owned())?;
                stack.last_mut().ok_or_else(|| "empty chart XML root".to_owned())?.children.push(node);
            }
            Event::DocType(_) => return Err("chart XML must not declare a DTD".into()),
            Event::Eof => break,
            _ => {}
        }
    }
    if stack.len() != 1 { return Err("unclosed chart XML tree".into()); }
    let root = stack.pop().ok_or_else(|| "missing chart XML tree".to_owned())?;
    if root.children.len() != 1 { return Err("chart XML must contain one root element".into()); }
    root.children.into_iter().next().ok_or_else(|| "missing chart XML root".into())
}

fn is_source_extension(node: &Node) -> bool {
    node.name == "ext" && node.namespace.as_deref() == Some(CHART_NS) && node.attrs.get("uri").is_some_and(|uri| uri == SOURCE_EXTENSION_URI)
}

/// Caches may be recalculated by Office. Only native settings and references
/// determine whether the exact authored DTO is still current.
pub fn signature(root: &Node) -> String {
    fn bytes(hash: &mut u64, value: &[u8]) {
        for byte in (value.len() as u64).to_le_bytes().iter().chain(value) { *hash ^= u64::from(*byte); *hash = hash.wrapping_mul(0x100000001b3); }
    }
    fn ignored(node: &Node) -> bool {
        node.namespace.as_deref() == Some(CHART_NS) && matches!(node.name.as_str(), "strCache" | "numCache" | "multiLvlStrCache")
            || is_source_extension(node)
            || node.namespace.as_deref() == Some(SOURCE_EXTENSION_NS)
            || node.name == "extLst" && !node.children.is_empty() && node.children.iter().all(is_source_extension)
    }
    let mut hash = 0xcbf29ce484222325u64;
    let mut stack = vec![(root, false)];
    while let Some((node, close)) = stack.pop() {
        if ignored(node) { continue; }
        bytes(&mut hash, if close { b"end" } else { b"start" });
        bytes(&mut hash, node.namespace.as_deref().unwrap_or("").as_bytes());
        bytes(&mut hash, node.name.as_bytes());
        if close { continue; }
        for (namespace, attributes) in &node.expanded_attrs {
            for (key, value) in attributes {
                bytes(&mut hash, namespace.as_bytes()); bytes(&mut hash, key.as_bytes()); bytes(&mut hash, value.as_bytes());
            }
        }
        if !node.text.trim().is_empty() { bytes(&mut hash, node.text.as_bytes()); }
        stack.push((node, true));
        stack.extend(node.children.iter().rev().map(|child| (child, false)));
    }
    format!("{hash:016x}")
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn attributes_entities_namespaces_and_cache_recalculation_preserve_settings_identity() {
        let a = parse(&format!("<c:chartSpace xmlns:c=\"{CHART_NS}\"><c:chart><c:title>A &amp; B &#x4E2D;</c:title><c:lineChart><c:numCache><c:v>1</c:v></c:numCache><c:f>'A &amp; B'!$B$2:$B$3</c:f></c:lineChart></c:chart></c:chartSpace>")).unwrap();
        let b = parse(&format!("<z:chartSpace xmlns:z=\"{CHART_NS}\"><z:chart><z:title>A &amp; B 中</z:title><z:lineChart><z:numCache><z:v>9</z:v></z:numCache><z:f>'A &amp; B'!$B$2:$B$3</z:f></z:lineChart></z:chart></z:chartSpace>")).unwrap();
        assert_eq!(a.descendants("title")[0].text, "A & B 中");
        assert_eq!(signature(&a), signature(&b));
        let c = parse(&format!("<c:chartSpace xmlns:c=\"{CHART_NS}\"><c:chart><c:lineChart><c:f>'A &amp; B'!$C$2:$C$3</c:f></c:lineChart></c:chart></c:chartSpace>")).unwrap();
        assert_ne!(signature(&a), signature(&c));
    }

    #[test]
    fn expanded_attributes_keep_distinct_namespaces_and_default_namespaces_do_not_bind_attributes() {
        let root = parse("<root xmlns=\"urn:root\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"urn:future\" val=\"plain\" w:val=\"word\" x:val=\"extension\"/>").unwrap();
        assert_eq!(root.namespace.as_deref(), Some("urn:root"));
        assert_eq!(root.attribute_ns("", "val"), Some("plain"));
        assert_eq!(root.attribute_ns("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "val"), Some("word"));
        assert_eq!(root.attribute_ns("urn:future", "val"), Some("extension"));
        assert_eq!(root.attrs.get("val").map(String::as_str), Some("plain"));
    }

    #[test]
    fn duplicate_expanded_names_and_unbound_attribute_prefixes_are_rejected() {
        assert!(parse("<root xmlns:a=\"urn:same\" xmlns:b=\"urn:same\" a:val=\"1\" b:val=\"2\"/>").is_err());
        assert!(parse("<root missing:val=\"1\"/>").is_err());
    }

    #[test]
    fn chart_freshness_includes_all_expanded_attributes_and_ignores_prefix_spelling() {
        let first = parse(&format!("<c:chartSpace xmlns:c=\"{CHART_NS}\" xmlns:a=\"urn:first\" xmlns:b=\"urn:second\" a:val=\"1\" b:val=\"2\"/>")).unwrap();
        let renamed = parse(&format!("<z:chartSpace xmlns:z=\"{CHART_NS}\" xmlns:x=\"urn:second\" xmlns:y=\"urn:first\" x:val=\"2\" y:val=\"1\"/>")).unwrap();
        let changed = parse(&format!("<c:chartSpace xmlns:c=\"{CHART_NS}\" xmlns:a=\"urn:first\" xmlns:b=\"urn:second\" a:val=\"1\" b:val=\"3\"/>")).unwrap();
        assert_eq!(signature(&first), signature(&renamed));
        assert_ne!(signature(&first), signature(&changed));
    }
}
