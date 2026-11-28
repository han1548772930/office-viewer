//! Namespace-aware, bounded XML fragments shared by conditional-format import and export.

use std::collections::{BTreeMap, BTreeSet};

use quick_xml::{XmlVersion, events::{BytesStart, Event}, name::ResolveResult, reader::NsReader};

#[path = "conditional_format_xml_selection.rs"]
mod selection;
pub use selection::{XmlName, XmlSelection, XmlSelectionLimits, extract};

pub const MAIN_NAMESPACE: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
pub const STRICT_MAIN_NAMESPACE: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
pub const X14_NAMESPACE: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
pub const XM_NAMESPACE: &str = "http://schemas.microsoft.com/office/excel/2006/main";
pub const MODEL_NAMESPACE: &str = "urn:office-studio:conditional-formatting:1";
pub const MODEL_EXTENSION_URI: &str = "{A7174764-91A7-4C36-96D4-5C1391AF06EB}";
pub const RULE_LINK_EXTENSION_URI: &str = "{B025F937-C7B1-47D3-B67F-A62EFF666E3E}";
pub const WORKSHEET_EXTENSION_URI: &str = "{78C0D931-6437-407D-A8EE-F0AAD7539E65}";
const XML_NAMESPACE: &str = "http://www.w3.org/XML/1998/namespace";
const XMLNS_NAMESPACE: &str = "http://www.w3.org/2000/xmlns/";
const XML_LIMIT: usize = 16 * 1024 * 1024;
const DEPTH_LIMIT: usize = 128;
const ELEMENT_LIMIT: usize = 262_144;
pub const XML_FRAGMENT_BYTE_LIMIT: usize = XML_LIMIT;
pub const XML_FRAGMENT_DEPTH_LIMIT: usize = DEPTH_LIMIT;
pub const XML_FRAGMENT_ELEMENT_LIMIT: usize = ELEMENT_LIMIT;

#[derive(Debug, Clone)]
pub enum Content { Child(usize), Text(String), Comment(String), Instruction(String) }

#[derive(Debug, Default, Clone)]
pub struct Element {
    /// Local element name; `qualified_name` retains its authored prefix.
    pub name: String,
    pub qualified_name: String,
    pub namespace: String,
    pub attrs: BTreeMap<String, String>,
    pub attribute_namespaces: BTreeMap<String, String>,
    pub text: String,
    pub children: Vec<Element>,
    pub content: Vec<Content>,
}

impl Element {
    pub fn new(namespace: &str, name: &str) -> Self {
        Self { name: name.into(), qualified_name: name.into(), namespace: namespace.into(), ..Self::default() }
    }

    pub fn prefixed(namespace: &str, prefix: &str, name: &str) -> Self {
        Self { qualified_name: format!("{prefix}:{name}"), ..Self::new(namespace, name) }
    }

    pub fn is(&self, namespace: &str, name: &str) -> bool { self.name == name && self.namespace == namespace }
    pub fn is_main(&self, name: &str) -> bool { self.name == name && matches!(self.namespace.as_str(), MAIN_NAMESPACE | STRICT_MAIN_NAMESPACE | "") }
    pub fn attribute_ns(&self, namespace: &str, name: &str) -> Option<&str> {
        self.attrs.iter().find(|(key, _)| key.rsplit(':').next() == Some(name)
            && self.attribute_namespaces.get(*key).map(String::as_str).unwrap_or("") == namespace).map(|(_, value)| value.as_str())
    }
    pub fn set_attr(&mut self, name: &str, value: impl Into<String>) {
        self.attrs.insert(name.into(), value.into());
        self.attribute_namespaces.insert(name.into(), if name == "xmlns" || name.starts_with("xmlns:") { XMLNS_NAMESPACE.into() } else { String::new() });
    }
    pub fn set_attr_ns(&mut self, namespace: &str, name: &str, value: impl Into<String>) {
        self.attrs.insert(name.into(), value.into());
        self.attribute_namespaces.insert(name.into(), namespace.into());
    }
    pub fn remove_attr(&mut self, name: &str) { self.attrs.remove(name); self.attribute_namespaces.remove(name); }
    pub fn push(&mut self, child: Element) { self.content.push(Content::Child(self.children.len())); self.children.push(child); }
    pub fn insert_before(&mut self, select: impl Fn(&Element) -> bool, child: Element) {
        let position = self.content.iter().position(|content| matches!(content, Content::Child(index) if self.children.get(*index).is_some_and(&select))).unwrap_or(self.content.len());
        self.content.insert(position, Content::Child(self.children.len())); self.children.push(child);
    }
    pub fn set_text(&mut self, text: impl Into<String>) {
        self.text = text.into(); self.children.clear(); self.content = vec![Content::Text(self.text.clone())];
    }
    pub fn child(&self, name: &str) -> Option<&Element> { self.children.iter().find(|child| child.is_main(name)) }
    pub fn child_mut(&mut self, name: &str) -> Option<&mut Element> { self.children.iter_mut().find(|child| child.is_main(name)) }
    pub fn detached_child(&self, child: &Element) -> Element { selection::detach_child(self, child) }

    /// Replace selected children without losing mixed text or unrelated extension order.
    pub fn replace_children(&mut self, select: impl Fn(&Element) -> bool, replacement: Vec<Element>) {
        let old_children = std::mem::take(&mut self.children);
        let old_content = std::mem::take(&mut self.content);
        let mut replacement = Some(replacement);
        for item in old_content {
            match item {
                Content::Child(index) if old_children.get(index).is_some_and(&select) => {
                    if let Some(children) = replacement.take() { for child in children { self.push(child); } }
                }
                Content::Child(index) => { if let Some(child) = old_children.get(index) { self.push(child.clone()); } }
                item => self.content.push(item),
            }
        }
        if let Some(children) = replacement { for child in children { self.push(child); } }
    }

    /// Metadata children in these Office containers have a schema-defined order.
    pub fn order_children(&mut self, names: &[&str]) {
        let mut children = std::mem::take(&mut self.children);
        children.sort_by_key(|child| names.iter().position(|name| child.name == *name).unwrap_or(names.len()));
        self.content.clear(); self.text.clear();
        for child in children { self.push(child); }
    }

    pub fn descendants(&self, name: &str) -> Vec<Element> {
        fn collect(node: &Element, name: &str, inherited: &BTreeMap<String, String>, result: &mut Vec<Element>) {
            let mut scope = inherited.clone();
            for (key, value) in &node.attrs {
                if key == "xmlns" { scope.insert(String::new(), value.clone()); }
                else if let Some(prefix) = key.strip_prefix("xmlns:") { scope.insert(prefix.into(), value.clone()); }
            }
            for child in &node.children {
                if child.name == name {
                    let mut owned = child.clone();
                    for (prefix, namespace) in &scope {
                        let key = if prefix.is_empty() { "xmlns".into() } else { format!("xmlns:{prefix}") };
                        if !owned.attrs.contains_key(&key) { owned.set_attr(&key, namespace); }
                    }
                    result.push(owned);
                }
                collect(child, name, &scope, result);
            }
        }
        let mut result = Vec::new(); collect(self, name, &BTreeMap::new(), &mut result); result
    }

    pub fn render(&self) -> String {
        let mut out = String::new(); render_element(self, &BTreeMap::new(), &mut out); out
    }

    /// Count the same canonical bytes as render without allocating a second XML string.
    pub fn rendered_size(&self) -> usize {
        let mut out = XmlByteCount::default(); render_element(self, &BTreeMap::new(), &mut out); out.0
    }

    /// Check detached output before render allocates its full canonical XML string.
    pub fn validate_fragment(&self) -> Result<(), String> {
        fn visit(node: &Element, depth: usize, elements: &mut usize) -> Result<(), String> {
            *elements = elements.saturating_add(1);
            if depth > DEPTH_LIMIT || *elements > ELEMENT_LIMIT {
                return Err("CF_SOURCE_FRAGMENT_LIMIT: conditional-format XML exceeds its element or depth budget".into())
            }
            for child in &node.children { visit(child, depth + 1, elements)?; }
            Ok(())
        }
        visit(self, 1, &mut 0)?;
        if self.rendered_size() > XML_LIMIT { return Err("CF_SOURCE_FRAGMENT_LIMIT: canonical conditional-format XML exceeds its byte budget".into()) }
        Ok(())
    }

    /// Resolved names, attribute order and indentation do not invalidate an untouched projection.
    /// This detects accidental stale metadata; it is not a security signature.
    pub fn fingerprint(&self) -> String {
        fn feed(hash: &mut u64, value: &str) {
            for byte in (value.len() as u64).to_le_bytes().into_iter().chain(value.bytes()) { *hash ^= u64::from(byte); *hash = hash.wrapping_mul(0x100000001b3); }
        }
        fn visit(node: &Element, hash: &mut u64) {
            feed(hash, "element"); feed(hash, &node.namespace); feed(hash, &node.name);
            let mut attrs = node.attrs.iter().filter(|(key, _)| *key != "xmlns" && !key.starts_with("xmlns:")).map(|(name, value)|
                (node.attribute_namespaces.get(name).map(String::as_str).unwrap_or(""), name.rsplit(':').next().unwrap_or(name), value.as_str())).collect::<Vec<_>>();
            attrs.sort();
            for (namespace, name, value) in attrs { feed(hash, "attribute"); feed(hash, namespace); feed(hash, name); feed(hash, value); }
            let mut text = String::new();
            let flush = |text: &mut String, hash: &mut u64| {
                if !text.is_empty() && (node.children.is_empty() || !text.trim().is_empty()) { feed(hash, "text"); feed(hash, text); }
                text.clear();
            };
            for content in &node.content {
                match content {
                    Content::Child(index) => { flush(&mut text, hash); if let Some(child) = node.children.get(*index) { visit(child, hash); } }
                    Content::Text(value) => text.push_str(value),
                    _ => {},
                }
            }
            flush(&mut text, hash);
            feed(hash, "/");
        }
        let mut hash = 0xcbf29ce484222325u64; visit(self, &mut hash); format!("{hash:016x}")
    }
}

trait XmlOutput {
    fn raw(&mut self, value: &str);
    fn escaped(&mut self, value: &str);
}
impl XmlOutput for String {
    fn raw(&mut self, value: &str) { self.push_str(value); }
    fn escaped(&mut self, value: &str) { self.push_str(&escape(value)); }
}
#[derive(Default)]
struct XmlByteCount(usize);
impl XmlOutput for XmlByteCount {
    fn raw(&mut self, value: &str) { self.0 = self.0.saturating_add(value.len()); }
    fn escaped(&mut self, value: &str) {
        for byte in value.bytes() { self.0 = self.0.saturating_add(match byte { b'&' => 5, b'<' | b'>' => 4, b'"' | b'\'' => 6, _ => 1 }); }
    }
}

fn render_element(node: &Element, inherited: &BTreeMap<String, String>, out: &mut impl XmlOutput) {
            let mut scope = inherited.clone();
            let mut declarations = BTreeMap::new();
            for (name, value) in &node.attrs {
                if name == "xmlns" { scope.insert(String::new(), value.clone()); }
                else if let Some(prefix) = name.strip_prefix("xmlns:") { scope.insert(prefix.into(), value.clone()); }
            }
            let mut bind = |name: &str, namespace: &str, attribute: bool| {
                let prefix = name.split_once(':').map(|(prefix, _)| prefix).unwrap_or("");
                if attribute && prefix.is_empty() || prefix == "xml" { return }
                if scope.get(prefix).map(String::as_str).unwrap_or("") != namespace {
                    declarations.insert(if prefix.is_empty() { "xmlns".into() } else { format!("xmlns:{prefix}") }, namespace.to_owned());
                    scope.insert(prefix.to_owned(), namespace.to_owned());
                }
            };
            bind(&node.qualified_name, &node.namespace, false);
            for (name, namespace) in &node.attribute_namespaces { if namespace != XMLNS_NAMESPACE { bind(name, namespace, true); } }
            out.raw("<"); out.raw(&node.qualified_name);
            let mut attributes = node.attrs.clone(); attributes.extend(declarations);
            for (name, value) in &attributes {
                out.raw(" "); out.raw(name); out.raw("=\""); out.escaped(value); out.raw("\"");
            }
            if node.content.is_empty() { out.raw("/>"); return }
            out.raw(">");
            for content in &node.content {
                match content {
                    Content::Child(index) => { if let Some(child) = node.children.get(*index) { render_element(child, &scope, out); } }
                    Content::Text(value) => out.escaped(value),
                    Content::Comment(value) => { out.raw("<!--"); out.raw(value); out.raw("-->"); }
                    Content::Instruction(value) => { out.raw("<?"); out.raw(value); out.raw("?>"); }
                }
            }
            out.raw("</"); out.raw(&node.qualified_name); out.raw(">");
}

fn escape(value: &str) -> String { value.replace('&', "&amp;").replace('<', "&lt;").replace('>', "&gt;").replace('"', "&quot;").replace('\'', "&apos;") }
fn valid_text(value: &str) -> bool { !value.chars().any(|value| value < '\u{20}' && !matches!(value, '\t' | '\n' | '\r') || matches!(value, '\u{fffe}' | '\u{ffff}')) }
fn namespace(result: ResolveResult<'_>) -> Result<String, String> {
    match result {
        ResolveResult::Bound(value) => {
            let value = value.as_ref().to_owned();
            let normalized = value.replace("\r\n", " ").replace(['\r', '\n', '\t'].as_slice(), " ");
            quick_xml::escape::unescape(&normalized).map(|value| value.into_owned()).map_err(|error| error.to_string())
        }
        ResolveResult::Unbound => Ok(String::new()),
        ResolveResult::Unknown(prefix) => Err(format!("Conditional-format XML uses undeclared prefix {}", prefix)),
    }
}

fn push_text(stack: &mut [Element], value: String) -> Result<(), String> {
    if !valid_text(&value) { return Err("Conditional-format XML has invalid characters".into()); }
    if let Some(node) = stack.last_mut() { node.text.push_str(&value); node.content.push(Content::Text(value)); }
    else if !value.trim().is_empty() { return Err("Conditional-format XML has text outside its root".into()); }
    Ok(())
}

fn valid_namespace_declaration(name: &str, value: &str) -> bool {
    !(name == "xmlns:xmlns" || name == "xmlns:xml" && value != XML_NAMESPACE || value == XMLNS_NAMESPACE
        || value == XML_NAMESPACE && name != "xmlns:xml" || name != "xmlns" && value.is_empty())
}

fn node_from_tag<R>(reader: &NsReader<R>, tag: &BytesStart<'_>, resolved_namespace: String) -> Result<Element, String> {
    let mut node = Element { name: String::from_utf8(tag.local_name().as_ref().as_bytes().to_vec()).map_err(|error| error.to_string())?,
        qualified_name: String::from_utf8(tag.name().as_ref().as_bytes().to_vec()).map_err(|error| error.to_string())?, namespace: resolved_namespace, ..Element::default() };
    let mut expanded = BTreeSet::new();
    for attribute in tag.attributes() {
        let attribute = attribute.map_err(|error| error.to_string())?;
        let name = String::from_utf8(attribute.key.as_ref().as_bytes().to_vec()).map_err(|error| error.to_string())?;
        let (resolved, local) = reader.resolver().resolve_attribute(attribute.key);
        let declaration = name == "xmlns" || name.starts_with("xmlns:");
        let namespace = if declaration { XMLNS_NAMESPACE.into() } else { namespace(resolved)? };
        if !expanded.insert((namespace.clone(), local.as_ref().to_owned())) { return Err("Conditional-format XML has duplicate expanded attributes".into()); }
        let value = attribute.normalized_value_with(XmlVersion::Implicit1_0, 1, quick_xml::escape::resolve_xml_entity).map_err(|error| error.to_string())?.into_owned();
        if !valid_text(&value) { return Err("Conditional-format XML has invalid characters".into()); }
        if declaration && !valid_namespace_declaration(&name, &value) { return Err("Conditional-format XML has an invalid namespace declaration".into()); }
        node.attribute_namespaces.insert(name.clone(), namespace); node.attrs.insert(name, value);
    }
    Ok(node)
}

pub fn parse(source: &str) -> Result<Element, String> {
    if source.len() > XML_LIMIT { return Err("Conditional-format XML exceeds the 16 MiB limit".into()); }
    let mut reader = NsReader::from_str(source); reader.config_mut().enable_all_checks(true);
    let mut stack = Vec::<Element>::new(); let mut root = None; let mut elements = 0;
    loop {
        let (resolved, event) = reader.read_resolved_event().map_err(|error| error.to_string())?;
        let empty = matches!(&event, Event::Empty(_));
        match event {
            Event::Start(tag) | Event::Empty(tag) => {
                elements += 1;
                if elements > ELEMENT_LIMIT || stack.len() >= DEPTH_LIMIT { return Err("Conditional-format XML exceeds its element or depth limit".into()); }
                let resolved_namespace = namespace(resolved)?;
                let node = node_from_tag(&reader, &tag, resolved_namespace)?;
                if empty { if let Some(parent) = stack.last_mut() { parent.push(node); } else if root.replace(node).is_some() { return Err("Conditional-format XML has multiple roots".into()); } }
                else { stack.push(node); }
            }
            Event::End(_) => {
                let node = stack.pop().ok_or("Conditional-format XML has an unmatched closing element")?;
                if let Some(parent) = stack.last_mut() { parent.push(node); } else if root.replace(node).is_some() { return Err("Conditional-format XML has multiple roots".into()); }
            }
            Event::Text(value) => push_text(&mut stack, value.xml_content(XmlVersion::Implicit1_0).into_owned())?,
            Event::CData(value) => push_text(&mut stack, value.xml_content(XmlVersion::Implicit1_0).into_owned())?,
            Event::GeneralRef(reference) => {
                let value = if let Some(value) = reference.resolve_char_ref().map_err(|error| error.to_string())? { value.to_string() }
                    else { quick_xml::escape::resolve_xml_entity(&reference.as_ref()).ok_or("Conditional-format XML has an unsupported entity")?.to_owned() };
                if !valid_text(&value) { return Err("Conditional-format XML has invalid characters".into()); }
                let node = stack.last_mut().ok_or("Conditional-format XML has an entity outside its root")?; node.text.push_str(&value); node.content.push(Content::Text(value));
            }
            Event::Comment(value) => { if let Some(node) = stack.last_mut() { node.content.push(Content::Comment(value.as_ref().to_owned())); } }
            Event::PI(value) => { if let Some(node) = stack.last_mut() { node.content.push(Content::Instruction(value.as_ref().to_owned())); } }
            Event::Decl(value) if root.is_none() && stack.is_empty() => { if value.version().map_err(|error| error.to_string())?.as_ref() != "1.0" { return Err("Conditional-format XML must use XML 1.0".into()); } }
            Event::Eof => { if !stack.is_empty() { return Err("Conditional-format XML has unclosed elements".into()); } break; }
            Event::DocType(_) | Event::Decl(_) => return Err("Conditional-format XML has an unsupported declaration".into()),
        }
    }
    root.ok_or_else(|| "Conditional-format XML has no root".into())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn subtrees_keep_inherited_prefixes_expanded_attributes_and_mixed_extension_text() {
        let document = parse(r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:v="urn:vendor"><conditionalFormatting sqref="A1"><cfRule type="expression" priority="1" v:type="custom"><formula>A1&gt;0</formula><extLst><ext uri="vendor"><v:item>before<v:n value="a&amp;b"/>after</v:item></ext></extLst></cfRule></conditionalFormatting></worksheet>"#).unwrap();
        let rule = document.descendants("cfRule").remove(0); let rendered = rule.render(); let reopened = parse(&rendered).unwrap();
        assert_eq!(reopened.attrs["type"], "expression"); assert_eq!(reopened.attribute_ns("urn:vendor", "type"), Some("custom"));
        assert_eq!(reopened.child("formula").unwrap().text, "A1>0"); assert_eq!(rule.fingerprint(), reopened.fingerprint());
        assert!(rendered.contains("before<v:n")); assert!(rendered.contains("/>after")); assert!(rendered.contains("xmlns:v=\"urn:vendor\""));
    }

    #[test]
    fn fingerprint_ignores_prefix_and_attribute_order_but_tracks_rule_and_style_edits() {
        let first = parse(r#"<dxf xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><font><b val="1"/><color rgb="FF112233"/></font></dxf>"#).unwrap();
        let renamed = parse(r#"<s:dxf xmlns:s="http://schemas.openxmlformats.org/spreadsheetml/2006/main"> <s:font> <s:b val="1"/> <s:color rgb="FF112233"/> </s:font> </s:dxf>"#).unwrap();
        assert_eq!(first.fingerprint(), renamed.fingerprint());
        assert_ne!(first.fingerprint(), parse(&first.render().replace("FF112233", "FF334455")).unwrap().fingerprint());
    }

    #[test]
    fn fingerprint_is_independent_of_entity_cdata_and_comment_text_segmentation() {
        let ordinary = parse("<formula>A1&gt;0&amp;B1</formula>").unwrap();
        let split = parse("<formula>A1<![CDATA[>]]>0<!-- split -->&#38;B1</formula>").unwrap();
        assert_eq!(ordinary.text, split.text);
        assert_eq!(ordinary.fingerprint(), split.fingerprint());
    }

    #[test]
    fn source_fragments_reject_entities_unbound_names_duplicate_attributes_and_unclosed_elements() {
        for malformed in ["<!DOCTYPE cfRule [<!ENTITY value 'x'>]><cfRule/>", "<v:cfRule/>", "<cfRule type='a' type='b'/>", "<cfRule><formula>&vendor;</formula></cfRule>", "<cfRule>", "<cfRule/><cfRule/>"] { assert!(parse(malformed).is_err(), "{malformed}"); }
    }

    #[test]
    fn canonical_byte_count_includes_utf8_escaping_inherited_namespaces_comments_and_instructions() {
        let parent = parse(r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:v="urn:vendor"><conditionalFormatting sqref="A1"><cfRule v:tag="汉&amp;&quot;&apos;"><formula><![CDATA[A1<5&😀]]></formula><!--keep--><?vendor keep?></cfRule></conditionalFormatting></worksheet>"#).unwrap();
        let group = parent.detached_child(&parent.children[0]); let rule = group.detached_child(&group.children[0]);
        assert_eq!(rule.rendered_size(), rule.render().len());
        assert!(rule.validate_fragment().is_ok());
        assert_eq!(rule.attrs.get("xmlns:v").map(String::as_str), Some("urn:vendor"));
    }

    #[test]
    fn composed_output_fragments_are_bounded_even_when_each_input_tree_was_shallow() {
        let mut node = Element::new(MAIN_NAMESPACE, "payload");
        for _ in 0..128 { let mut parent = Element::new(MAIN_NAMESPACE, "payload"); parent.push(node); node = parent; }
        assert!(node.validate_fragment().unwrap_err().starts_with("CF_SOURCE_FRAGMENT_LIMIT"));
    }
}
