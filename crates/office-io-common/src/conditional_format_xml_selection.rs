//! Scan an Office XML part without constructing trees for unselected content.
//! Limits apply to detached source fragments, independently of worksheet data.

use std::collections::BTreeMap;
use std::io::BufRead;

use quick_xml::{XmlVersion, events::Event, reader::NsReader};

use super::{Content, Element, XML_NAMESPACE, XML_FRAGMENT_BYTE_LIMIT,
    XML_FRAGMENT_DEPTH_LIMIT, XML_FRAGMENT_ELEMENT_LIMIT, namespace, node_from_tag, push_text, valid_text};

const MC_NAMESPACE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
pub const SELECTED_PART_BYTE_LIMIT: usize = 64 * 1024 * 1024;
pub const SELECTED_PART_ELEMENT_LIMIT: usize = 1_048_576;
pub const SCAN_DEPTH_LIMIT: usize = 4096;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XmlName { pub namespace: String, pub name: String }

impl XmlName {
    pub fn is(&self, namespace: &str, name: &str) -> bool { self.namespace == namespace && self.name == name }
    pub fn is_main(&self, name: &str) -> bool {
        self.name == name && matches!(self.namespace.as_str(), super::MAIN_NAMESPACE | super::STRICT_MAIN_NAMESPACE | "")
    }
}

#[derive(Debug, Clone, Copy)]
pub struct XmlSelectionLimits {
    pub fragment_bytes: usize,
    pub fragment_depth: usize,
    pub fragment_elements: usize,
    pub total_bytes: usize,
    pub total_elements: usize,
    /// A separate guard for the scanner; unselected nodes do not spend fragment depth.
    pub scan_depth: usize,
}

impl Default for XmlSelectionLimits {
    fn default() -> Self {
        Self { fragment_bytes: XML_FRAGMENT_BYTE_LIMIT, fragment_depth: XML_FRAGMENT_DEPTH_LIMIT,
            fragment_elements: XML_FRAGMENT_ELEMENT_LIMIT, total_bytes: SELECTED_PART_BYTE_LIMIT,
            total_elements: SELECTED_PART_ELEMENT_LIMIT, scan_depth: SCAN_DEPTH_LIMIT }
    }
}

#[derive(Debug)]
pub struct XmlSelection {
    /// Only the root name and attributes; unselected children and text are absent.
    pub root: Element,
    /// Each selected outermost subtree is complete and independently renderable.
    pub fragments: Vec<Element>,
    pub selected_bytes: usize,
    pub selected_elements: usize,
}

#[derive(Clone)]
struct ContextToken { lexical: String, namespace: Option<String>, local: Option<String> }

#[derive(Clone)]
struct ContextAttribute {
    qualified_name: String,
    namespace: String,
    value: String,
    tokens: Option<Vec<ContextToken>>,
}

type ContextKey = (String, String);
type Context = BTreeMap<ContextKey, ContextAttribute>;

#[derive(Default)]
struct ScopeUndo {
    namespaces: Vec<(String, Option<String>)>,
    context: Vec<(ContextKey, Option<ContextAttribute>)>,
}

fn context_tokens(namespace: &str, local: &str, value: &str, bindings: &BTreeMap<String, String>) -> Option<Vec<ContextToken>> {
    if namespace != MC_NAMESPACE { return None }
    let prefix_list = matches!(local, "Ignorable" | "MustUnderstand");
    if !prefix_list && !matches!(local, "ProcessContent" | "PreserveElements" | "PreserveAttributes") { return None }
    Some(value.split_whitespace().map(|token| {
        let (prefix, local) = if prefix_list { (token, None) }
            else { token.split_once(':').map(|(prefix, local)| (prefix, Some(local.to_owned()))).unwrap_or(("", Some(token.to_owned()))) };
        ContextToken { lexical: token.into(), namespace: bindings.get(prefix).cloned(), local }
    }).collect())
}

fn enter_scope(node: &Element, bindings: &mut BTreeMap<String, String>, context: &mut Context) -> ScopeUndo {
    let mut undo = ScopeUndo::default();
    for (key, value) in &node.attrs {
        let prefix = if key == "xmlns" { Some("") } else { key.strip_prefix("xmlns:") };
        if let Some(prefix) = prefix { undo.namespaces.push((prefix.into(), bindings.insert(prefix.into(), value.clone()))); }
    }
    for (key, value) in &node.attrs {
        let namespace = node.attribute_namespaces.get(key).map(String::as_str).unwrap_or("");
        let local = key.rsplit(':').next().unwrap_or(key);
        if namespace == MC_NAMESPACE || namespace == XML_NAMESPACE && matches!(local, "lang" | "space" | "base") {
            let context_key = (namespace.to_owned(), local.to_owned());
            let attribute = ContextAttribute { qualified_name: key.clone(), namespace: namespace.into(), value: value.clone(),
                tokens: context_tokens(namespace, local, value, bindings) };
            undo.context.push((context_key.clone(), context.insert(context_key, attribute)));
        }
    }
    undo
}

fn leave_scope(undo: ScopeUndo, bindings: &mut BTreeMap<String, String>, context: &mut Context) {
    for (key, before) in undo.context.into_iter().rev() {
        if let Some(before) = before { context.insert(key, before); } else { context.remove(&key); }
    }
    for (prefix, before) in undo.namespaces.into_iter().rev() {
        if let Some(before) = before { bindings.insert(prefix, before); } else { bindings.remove(&prefix); }
    }
}

fn matching_prefix(node: &mut Element, bindings: &mut BTreeMap<String, String>, preferred: &str, namespace: &str) -> String {
    if preferred == "xml" && namespace == XML_NAMESPACE { return "xml".into() }
    if !preferred.is_empty() && bindings.get(preferred).is_some_and(|value| value == namespace) { return preferred.into() }
    if let Some((prefix, _)) = bindings.iter().find(|(prefix, value)| !prefix.is_empty() && value.as_str() == namespace) { return prefix.clone() }
    let mut index = 0usize;
    let mut prefix = if preferred.is_empty() || matches!(preferred, "xml" | "xmlns") { "cfctx".into() } else { preferred.to_owned() };
    while bindings.contains_key(&prefix) { index += 1; prefix = format!("cfctx{index}"); }
    bindings.insert(prefix.clone(), namespace.into()); node.set_attr(&format!("xmlns:{prefix}"), namespace); prefix
}

fn detached(mut node: Element, bindings: &BTreeMap<String, String>, context: &Context) -> Element {
    // Preserve every active binding, including prefixes used only in QName values.
    let mut local_bindings = bindings.clone();
    for (prefix, value) in bindings {
        if prefix == "xml" { continue }
        let key = if prefix.is_empty() { "xmlns".into() } else { format!("xmlns:{prefix}") };
        if !node.attrs.contains_key(&key) { node.set_attr(&key, value); }
    }
    for ((namespace, local), attribute) in context {
        if node.attribute_ns(namespace, local).is_some() { continue }
        let authored_prefix = attribute.qualified_name.split_once(':').map(|(prefix, _)| prefix).unwrap_or("");
        let prefix = matching_prefix(&mut node, &mut local_bindings, authored_prefix, &attribute.namespace);
        let mut value = attribute.value.clone();
        if let Some(tokens) = &attribute.tokens {
            value = tokens.iter().map(|token| {
                let Some(namespace) = &token.namespace else { return token.lexical.clone() };
                if namespace.is_empty() { return token.local.clone().unwrap_or_else(|| token.lexical.clone()) }
                let authored_prefix = token.lexical.split_once(':').map(|(prefix, _)| prefix)
                    .unwrap_or(if token.local.is_none() { token.lexical.as_str() } else { "" });
                let prefix = matching_prefix(&mut node, &mut local_bindings, authored_prefix, namespace);
                token.local.as_ref().map(|local| format!("{prefix}:{local}")).unwrap_or(prefix)
            }).collect::<Vec<_>>().join(" ");
        }
        node.set_attr_ns(&attribute.namespace, &format!("{prefix}:{local}"), value);
    }
    node
}

pub(super) fn detach_child(parent: &Element, child: &Element) -> Element {
    let mut bindings = BTreeMap::from([("xml".into(), XML_NAMESPACE.into())]);
    let mut context = Context::new();
    enter_scope(parent, &mut bindings, &mut context);
    enter_scope(child, &mut bindings, &mut context);
    detached(child.clone(), &bindings, &context)
}

struct Capture { stack: Vec<Element>, elements: usize, input_start: u64 }

fn xml_error(error: impl std::fmt::Display) -> String { format!("CF_SOURCE_XML_INVALID: {error}") }
fn fragment_error(reason: &str) -> String { format!("CF_SOURCE_FRAGMENT_LIMIT: {reason}") }
fn total_error(reason: &str) -> String { format!("CF_SOURCE_TOTAL_LIMIT: {reason}") }

fn check_raw(capture: &Capture, position: u64, limits: XmlSelectionLimits) -> Result<(), String> {
    if position.saturating_sub(capture.input_start) > limits.fragment_bytes as u64 {
        return Err(fragment_error("selected XML subtree exceeds the input byte budget"))
    }
    Ok(())
}

fn publish(node: Element, elements: usize, limits: XmlSelectionLimits, fragments: &mut Vec<Element>, bytes: &mut usize, total_elements: &mut usize) -> Result<(), String> {
    let size = node.rendered_size();
    if size > limits.fragment_bytes { return Err(fragment_error("detached canonical XML exceeds the byte budget")) }
    let next_bytes = bytes.checked_add(size).ok_or_else(|| total_error("selected XML byte counter overflow"))?;
    let next_elements = total_elements.checked_add(elements).ok_or_else(|| total_error("selected XML element counter overflow"))?;
    if next_bytes > limits.total_bytes || next_elements > limits.total_elements {
        return Err(total_error("selected XML fragments exceed the part budget"))
    }
    *bytes = next_bytes; *total_elements = next_elements; fragments.push(node); Ok(())
}

/// The selector sees the resolved path and a transient name/attribute node.
/// It is not called inside an active capture, preventing duplicate nested copies.
/// No result is returned until the complete part, including its suffix, is valid.
pub fn extract<R: BufRead>(input: R, limits: XmlSelectionLimits,
    mut select: impl FnMut(&[XmlName], &Element) -> Result<bool, String>) -> Result<XmlSelection, String> {
    let mut reader = NsReader::from_reader(input); reader.config_mut().enable_all_checks(true);
    let mut buffer = Vec::new();
    let mut path = Vec::<XmlName>::new();
    let mut scopes = Vec::<ScopeUndo>::new();
    let mut bindings = BTreeMap::from([("xml".into(), XML_NAMESPACE.into())]);
    let mut context = Context::new();
    let mut root = None;
    let mut capture = None::<Capture>;
    let mut fragments = Vec::new();
    let mut selected_bytes = 0usize;
    let mut selected_elements = 0usize;
    loop {
        let input_start = reader.buffer_position();
        let (resolved, event) = reader.read_resolved_event_into(&mut buffer).map_err(xml_error)?;
        let resolved_namespace = namespace(resolved).map_err(xml_error)?;
        let position = reader.buffer_position();
        let empty = matches!(&event, Event::Empty(_));
        if let Some(active) = &capture { check_raw(active, position, limits)?; }
        match event {
            Event::Start(tag) | Event::Empty(tag) => {
                if path.len() >= limits.scan_depth { return Err(total_error("XML scanner exceeds its separate depth guard")) }
                let node = node_from_tag(&reader, &tag, resolved_namespace).map_err(xml_error)?;
                if path.is_empty() {
                    if root.is_some() { return Err(xml_error("XML part has multiple roots")) }
                    root = Some(node.clone());
                }
                path.push(XmlName { namespace: node.namespace.clone(), name: node.name.clone() });
                let undo = enter_scope(&node, &mut bindings, &mut context);
                if capture.is_none() && select(&path, &node)? {
                    let node = detached(node, &bindings, &context);
                    capture = Some(Capture { stack: vec![node], elements: 1, input_start });
                } else if let Some(active) = &mut capture {
                    active.elements = active.elements.checked_add(1).ok_or_else(|| fragment_error("XML element counter overflow"))?;
                    if active.stack.len() >= limits.fragment_depth { return Err(fragment_error("selected XML subtree exceeds the depth budget")) }
                    active.stack.push(node);
                }
                if let Some(active) = &capture {
                    check_raw(active, position, limits)?;
                    if active.elements > limits.fragment_elements || active.stack.len() > limits.fragment_depth {
                        return Err(fragment_error("selected XML subtree exceeds the element or depth budget"))
                    }
                    if selected_elements.saturating_add(active.elements) > limits.total_elements {
                        return Err(total_error("selected XML fragments exceed the part element budget"))
                    }
                }
                if empty {
                    if let Some(active) = &mut capture {
                        let node = active.stack.pop().ok_or_else(|| xml_error("empty capture stack"))?;
                        if let Some(parent) = active.stack.last_mut() { parent.push(node); }
                        else {
                            let elements = active.elements;
                            publish(node, elements, limits, &mut fragments, &mut selected_bytes, &mut selected_elements)?;
                            capture = None;
                        }
                    }
                    path.pop(); leave_scope(undo, &mut bindings, &mut context);
                } else { scopes.push(undo); }
            }
            Event::End(_) => {
                if path.pop().is_none() { return Err(xml_error("XML part has an unmatched closing element")) }
                if let Some(active) = &mut capture {
                    let node = active.stack.pop().ok_or_else(|| xml_error("capture has an unmatched closing element"))?;
                    if let Some(parent) = active.stack.last_mut() { parent.push(node); }
                    else {
                        let elements = active.elements;
                        publish(node, elements, limits, &mut fragments, &mut selected_bytes, &mut selected_elements)?;
                        capture = None;
                    }
                }
                let undo = scopes.pop().ok_or_else(|| xml_error("XML part has an unmatched namespace scope"))?;
                leave_scope(undo, &mut bindings, &mut context);
            }
            Event::Text(value) => {
                let value = value.xml_content(XmlVersion::Implicit1_0).map_err(xml_error)?.into_owned();
                if !valid_text(&value) || path.is_empty() && !value.trim().is_empty() { return Err(xml_error("invalid text or text outside the root")) }
                if let Some(active) = &mut capture { push_text(&mut active.stack, value).map_err(xml_error)?; }
            }
            Event::CData(value) => {
                let value = value.xml_content(XmlVersion::Implicit1_0).map_err(xml_error)?.into_owned();
                if !valid_text(&value) || path.is_empty() { return Err(xml_error("invalid CDATA or CDATA outside the root")) }
                if let Some(active) = &mut capture { push_text(&mut active.stack, value).map_err(xml_error)?; }
            }
            Event::GeneralRef(reference) => {
                let value = if let Some(value) = reference.resolve_char_ref().map_err(xml_error)? { value.to_string() }
                    else { quick_xml::escape::resolve_xml_entity(&reference.decode().map_err(xml_error)?).ok_or_else(|| xml_error("unsupported XML entity"))?.to_owned() };
                if !valid_text(&value) || path.is_empty() { return Err(xml_error("invalid entity or entity outside the root")) }
                if let Some(active) = &mut capture { push_text(&mut active.stack, value).map_err(xml_error)?; }
            }
            Event::Comment(value) => {
                let value = String::from_utf8(value.as_ref().to_vec()).map_err(xml_error)?;
                if !valid_text(&value) { return Err(xml_error("invalid XML comment characters")) }
                if let Some(active) = &mut capture { active.stack.last_mut().ok_or_else(|| xml_error("comment outside capture"))?.content.push(Content::Comment(value)); }
            }
            Event::PI(value) => {
                let value = String::from_utf8(value.as_ref().to_vec()).map_err(xml_error)?;
                if !valid_text(&value) { return Err(xml_error("invalid XML instruction characters")) }
                if let Some(active) = &mut capture { active.stack.last_mut().ok_or_else(|| xml_error("instruction outside capture"))?.content.push(Content::Instruction(value)); }
            }
            Event::Decl(value) if root.is_none() && path.is_empty() => {
                if value.version().map_err(xml_error)?.as_ref() != b"1.0" { return Err(xml_error("XML part must use XML 1.0")) }
            }
            Event::Eof => {
                if !path.is_empty() || !scopes.is_empty() || capture.is_some() { return Err(xml_error("XML part has unclosed elements")) }
                break
            }
            Event::DocType(_) | Event::Decl(_) => return Err(xml_error("unsupported XML declaration")),
        }
        buffer.clear();
    }
    let root = root.ok_or_else(|| xml_error("XML part has no root"))?;
    Ok(XmlSelection { root, fragments, selected_bytes, selected_elements })
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::{BufReader, Cursor};

    fn groups(xml: &str, limits: XmlSelectionLimits) -> Result<XmlSelection, String> {
        extract(BufReader::with_capacity(17, Cursor::new(xml.as_bytes())), limits, |path, _| Ok(path.last().is_some_and(|name|
            name.is_main("conditionalFormatting") || name.is(super::super::X14_NAMESPACE, "conditionalFormatting"))))
    }

    #[test]
    fn unselected_large_cell_data_and_deep_containers_do_not_spend_fragment_budgets() {
        let data = "<row><c><v>7</v></c></row>".repeat(90_000);
        let text = "x".repeat(XML_FRAGMENT_BYTE_LIMIT + 1);
        let deep = format!("{}ignored{}", "<v:n>".repeat(200), "</v:n>".repeat(200));
        let xml = format!(r#"<worksheet xmlns:v="urn:vendor"><sheetData>{data}<row><c><is><t>{text}</t></is></c></row></sheetData>{deep}<conditionalFormatting sqref="A1"><cfRule type="expression"><formula>A1&gt;0</formula></cfRule></conditionalFormatting></worksheet>"#);
        let selected = groups(&xml, XmlSelectionLimits { fragment_bytes: 1024, fragment_depth: 3, fragment_elements: 3,
            total_bytes: 1024, total_elements: 3, ..XmlSelectionLimits::default() }).unwrap();
        assert!(xml.len() > XML_FRAGMENT_BYTE_LIMIT);
        assert!(selected.root.children.is_empty()); assert!(selected.root.text.is_empty());
        assert_eq!(selected.fragments.len(), 1); assert_eq!(selected.selected_elements, 3);
        assert_eq!(selected.fragments[0].child("cfRule").unwrap().child("formula").unwrap().text, "A1>0");
    }

    #[test]
    fn strict_alias_qname_only_namespaces_and_mixed_vendor_content_survive_detaching() {
        let xml = format!(r#"<s:worksheet xmlns:s="{}" xmlns:compat="{MC_NAMESPACE}" xmlns:q="urn:qname-only" xmlns:v="urn:outer" compat:Ignorable="q v" xml:lang="zh-CN" xml:space="preserve"><s:conditionalFormatting sqref="A1"><s:cfRule type="future"><v:payload>汉&amp;<!--retained--><?vendor keep?><v:n xmlns:v="urn:inner"/>字😀</v:payload></s:cfRule></s:conditionalFormatting></s:worksheet>"#, super::super::STRICT_MAIN_NAMESPACE);
        let selected = groups(&xml, XmlSelectionLimits::default()).unwrap();
        let group = &selected.fragments[0];
        assert_eq!(group.namespace, super::super::STRICT_MAIN_NAMESPACE);
        assert_eq!(group.attrs.get("xmlns:q").map(String::as_str), Some("urn:qname-only"));
        assert_eq!(group.attribute_ns(MC_NAMESPACE, "Ignorable"), Some("q v"));
        assert_eq!(group.attribute_ns(XML_NAMESPACE, "lang"), Some("zh-CN"));
        assert_eq!(group.attribute_ns(XML_NAMESPACE, "space"), Some("preserve"));
        let rule = group.detached_child(&group.children[0]);
        let rendered = rule.render();
        let reopened = super::super::parse(&rendered).unwrap();
        assert_eq!(rule.fingerprint(), reopened.fingerprint());
        assert_eq!(reopened.children[0].children[0].namespace, "urn:inner");
        assert!(rendered.contains("<!--retained--><?vendor keep?>")); assert!(rendered.contains("字😀"));
    }

    #[test]
    fn inherited_mc_qnames_keep_their_authored_namespace_after_prefix_rebinding() {
        let xml = format!(r#"<worksheet xmlns:mc="{MC_NAMESPACE}" xmlns:v="urn:authored" xmlns:a="urn:authored" xmlns:x14="{}" mc:Ignorable="v" mc:PreserveAttributes="v:tag"><extLst xmlns:v="urn:local"><ext><x14:conditionalFormatting><x14:cfRule type="future"/></x14:conditionalFormatting></ext></extLst></worksheet>"#, super::super::X14_NAMESPACE);
        let selected = groups(&xml, XmlSelectionLimits::default()).unwrap();
        let group = &selected.fragments[0];
        assert_eq!(group.attrs.get("xmlns:v").map(String::as_str), Some("urn:local"));
        assert_eq!(group.attribute_ns(MC_NAMESPACE, "Ignorable"), Some("a"));
        assert_eq!(group.attribute_ns(MC_NAMESPACE, "PreserveAttributes"), Some("a:tag"));
        assert_eq!(group.attrs.get("xmlns:a").map(String::as_str), Some("urn:authored"));
        assert!(super::super::parse(&group.render()).is_ok());
    }

    #[test]
    fn corrupt_suffix_or_multiple_roots_discards_earlier_complete_captures() {
        let prefix = "<worksheet><conditionalFormatting sqref=\"A1\"><cfRule type=\"future\"/></conditionalFormatting>";
        for suffix in ["<sheetData><row>", "<sheetData></wrong></worksheet>", "</worksheet><second/>", "<v:unknown/></worksheet>", "<!--bad--comment--></worksheet>"] {
            let error = groups(&format!("{prefix}{suffix}"), XmlSelectionLimits::default()).unwrap_err();
            assert!(error.starts_with("CF_SOURCE_XML_INVALID"), "{error}");
        }
    }

    #[test]
    fn selected_fragment_depth_elements_and_raw_bytes_fail_with_distinct_fragment_class() {
        let xml = "<worksheet><conditionalFormatting sqref=\"A1\"><cfRule><formula>A1&gt;0</formula></cfRule></conditionalFormatting></worksheet>";
        for limits in [XmlSelectionLimits { fragment_depth: 2, ..XmlSelectionLimits::default() },
            XmlSelectionLimits { fragment_elements: 2, ..XmlSelectionLimits::default() },
            XmlSelectionLimits { fragment_bytes: 30, ..XmlSelectionLimits::default() }] {
            assert!(groups(xml, limits).unwrap_err().starts_with("CF_SOURCE_FRAGMENT_LIMIT"));
        }
    }

    #[test]
    fn canonical_cdata_escaping_and_unicode_bytes_are_counted_before_publish() {
        let text = format!("{}汉😀", "&".repeat(200));
        let xml = format!("<worksheet><conditionalFormatting sqref=\"A1\"><cfRule><![CDATA[{text}]]></cfRule></conditionalFormatting></worksheet>");
        let first = groups(&xml, XmlSelectionLimits::default()).unwrap();
        let canonical_bytes = first.fragments[0].render().len();
        assert_eq!(first.selected_bytes, canonical_bytes);
        assert!(canonical_bytes > xml.len());
        let error = groups(&xml, XmlSelectionLimits { fragment_bytes: canonical_bytes - 1, ..XmlSelectionLimits::default() }).unwrap_err();
        assert!(error.starts_with("CF_SOURCE_FRAGMENT_LIMIT")); assert!(error.contains("canonical"));
    }

    #[test]
    fn many_small_fragments_exceed_the_part_total_without_returning_partial_rules() {
        let group = "<conditionalFormatting sqref=\"A1\"><cfRule><formula>A1</formula></cfRule></conditionalFormatting>";
        let xml = format!("<worksheet>{group}{group}</worksheet>");
        let all = groups(&xml, XmlSelectionLimits::default()).unwrap();
        let one_bytes = all.fragments[0].rendered_size();
        for limits in [XmlSelectionLimits { total_bytes: one_bytes, ..XmlSelectionLimits::default() },
            XmlSelectionLimits { total_elements: 3, ..XmlSelectionLimits::default() }] {
            assert!(groups(&xml, limits).unwrap_err().starts_with("CF_SOURCE_TOTAL_LIMIT"));
        }
    }

    #[test]
    fn the_scanner_has_a_separate_depth_guard_for_unselected_content() {
        let xml = "<worksheet><sheetData><row><c/></row></sheetData></worksheet>";
        let error = groups(xml, XmlSelectionLimits { scan_depth: 3, ..XmlSelectionLimits::default() }).unwrap_err();
        assert!(error.starts_with("CF_SOURCE_TOTAL_LIMIT")); assert!(error.contains("scanner"));
        assert!(groups(xml, XmlSelectionLimits { fragment_depth: 1, fragment_elements: 1, ..XmlSelectionLimits::default() }).unwrap().fragments.is_empty());
    }

    #[test]
    fn unselected_malformed_attributes_namespaces_and_declarations_still_reject_the_part() {
        for xml in [r#"<worksheet><sheetData><row r='1' r='2'/></sheetData></worksheet>"#,
            r#"<worksheet xmlns:p="urn:a" xmlns:q="urn:a"><sheetData><row p:x="1" q:x="2"/></sheetData></worksheet>"#,
            r#"<worksheet xmlns:xml="urn:invalid"><sheetData/></worksheet>"#,
            r#"<?xml version="1.1"?><worksheet><sheetData/></worksheet>"#,
            "<worksheet><sheetData>&unknown;</sheetData></worksheet>"] {
            assert!(groups(xml, XmlSelectionLimits::default()).unwrap_err().starts_with("CF_SOURCE_XML_INVALID"), "{xml}");
        }
    }

    #[test]
    fn capturing_an_outer_subtree_keeps_nested_unknown_content_without_duplicate_selections() {
        let xml = "<worksheet><conditionalFormatting sqref=\"A1\"><cfRule><extLst><conditionalFormatting><unknown/></conditionalFormatting></extLst></cfRule></conditionalFormatting></worksheet>";
        let selected = groups(xml, XmlSelectionLimits::default()).unwrap();
        assert_eq!(selected.fragments.len(), 1); assert_eq!(selected.selected_elements, 5);
        assert_eq!(selected.fragments[0].descendants("conditionalFormatting").len(), 1);
    }
}
