use std::collections::{BTreeMap, HashSet};

use quick_xml::{XmlVersion, events::Event, name::ResolveResult, reader::NsReader};

use crate::xml;

pub(super) const WORD_NAMESPACE: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
pub(super) const DRAWING_NAMESPACE: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
pub(super) const XML_NAMESPACE: &str = "http://www.w3.org/XML/1998/namespace";
pub(super) const XMLNS_NAMESPACE: &str = "http://www.w3.org/2000/xmlns/";
const STRICT_WORD_NAMESPACE: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const STRICT_DRAWING_NAMESPACE: &str = "http://purl.oclc.org/ooxml/drawingml/main";
const RELATIONSHIP_NAMESPACE: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_RELATIONSHIP_NAMESPACE: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const XML_LIMIT: usize = 16 * 1024 * 1024;
const XML_DEPTH_LIMIT: usize = 128;
const XML_ELEMENT_LIMIT: usize = 262_144;

#[derive(Clone)]
pub(super) struct Attribute {
    pub name: String,
    pub namespace: String,
    pub local: String,
    pub value: String,
}

#[derive(Clone)]
pub(super) enum Content { Element(Element), Text(String) }

#[derive(Clone)]
pub(super) struct Element {
    pub name: String,
    pub namespace: String,
    pub local: String,
    pub attributes: Vec<Attribute>,
    pub content: Vec<Content>,
    pub empty: bool,
}

impl Element {
    pub fn word(local: &str) -> Self {
        Self { name: format!("w:{local}"), namespace: WORD_NAMESPACE.into(), local: local.into(), attributes: Vec::new(), content: Vec::new(), empty: true }
    }

    pub fn is_word(&self, local: &str) -> bool { self.namespace == WORD_NAMESPACE && self.local == local }

    pub fn attribute(&self, local: &str) -> Option<&str> {
        self.attributes.iter().find(|attribute| attribute.namespace == WORD_NAMESPACE && attribute.local == local).map(|attribute| attribute.value.as_str())
    }

    pub fn set_word_attribute(&mut self, local: &str, value: impl Into<String>) {
        let value = value.into();
        if let Some(attribute) = self.attributes.iter_mut().find(|attribute| attribute.namespace == WORD_NAMESPACE && attribute.local == local) { attribute.value = value; }
        else { self.attributes.push(Attribute { name: format!("w:{local}"), namespace: WORD_NAMESPACE.into(), local: local.into(), value }); }
    }

    pub fn children(&self) -> impl Iterator<Item = &Element> {
        self.content.iter().filter_map(|content| match content { Content::Element(element) => Some(element), _ => None })
    }

    pub fn child(&self, local: &str) -> Option<&Element> { self.children().find(|element| element.is_word(local)) }

    pub fn text(&self) -> String {
        self.content.iter().filter_map(|content| match content { Content::Text(text) => Some(text.as_str()), _ => None }).collect()
    }

    pub fn opening(&self, namespaces: Option<&BTreeMap<String, String>>, empty: bool) -> String {
        let mut result = format!("<{}", self.name);
        if let Some(namespaces) = namespaces {
            for (prefix, namespace) in namespaces {
                let name = if prefix.is_empty() { "xmlns".to_owned() } else { format!("xmlns:{prefix}") };
                if !self.attributes.iter().any(|attribute| attribute.name == name) { result.push_str(&xml::attr(&name, namespace)); }
            }
        }
        for attribute in &self.attributes { result.push_str(&xml::attr(&attribute.name, &attribute.value)); }
        result.push_str(if empty { "/>" } else { ">" });
        result
    }

    pub fn render(&self, namespaces: Option<&BTreeMap<String, String>>) -> String {
        if self.empty && self.content.is_empty() { return self.opening(namespaces, true); }
        let mut result = self.opening(namespaces, false);
        for content in &self.content {
            match content {
                Content::Element(element) => result.push_str(&element.render(None)),
                Content::Text(text) => result.push_str(&xml::escape(text)),
            }
        }
        result.push_str(&format!("</{}>", self.name));
        result
    }
}

/// The writer emits transitional OOXML. Source fragments may use either
/// vocabulary; normalize resolved names and declarations together.
pub(super) fn canonicalize_office_namespaces(element: &mut Element) -> Result<(), String> {
    fn mapped(namespace: &str) -> Option<(&'static str, &'static str)> {
        match namespace {
            WORD_NAMESPACE | STRICT_WORD_NAMESPACE => Some(("w", WORD_NAMESPACE)),
            DRAWING_NAMESPACE | STRICT_DRAWING_NAMESPACE => Some(("a", DRAWING_NAMESPACE)),
            RELATIONSHIP_NAMESPACE | STRICT_RELATIONSHIP_NAMESPACE => Some(("r", RELATIONSHIP_NAMESPACE)),
            _ => None,
        }
    }
    fn visit(element: &mut Element) -> Result<(), String> {
        if let Some((prefix, namespace)) = mapped(&element.namespace) {
            element.name = format!("{prefix}:{}", element.local);
            element.namespace = namespace.into();
        }
        for attribute in &mut element.attributes {
            if attribute.namespace == XMLNS_NAMESPACE {
                if let Some((_, namespace)) = mapped(&attribute.value) { attribute.value = namespace.into(); }
                let expected = match attribute.name.as_str() {
                    "xmlns:w" => Some(WORD_NAMESPACE), "xmlns:a" => Some(DRAWING_NAMESPACE), "xmlns:r" => Some(RELATIONSHIP_NAMESPACE), _ => None,
                };
                if expected.is_some_and(|namespace| attribute.value != namespace) { return Err("Source XML rebinds a canonical Office namespace prefix.".into()); }
            } else if let Some((prefix, namespace)) = mapped(&attribute.namespace) {
                attribute.name = format!("{prefix}:{}", attribute.local);
                attribute.namespace = namespace.into();
            }
        }
        for content in &mut element.content {
            if let Content::Element(child) = content { visit(child)?; }
        }
        Ok(())
    }
    visit(element)?;
    // Canonical names must be bound even when the source used another prefix.
    for (prefix, namespace) in [("w", WORD_NAMESPACE), ("a", DRAWING_NAMESPACE), ("r", RELATIONSHIP_NAMESPACE)] {
        let name = format!("xmlns:{prefix}");
        if let Some(declaration) = element.attributes.iter_mut().find(|attribute| attribute.name == name) {
            declaration.value = namespace.into();
        } else {
            element.attributes.push(Attribute { name, namespace: XMLNS_NAMESPACE.into(), local: prefix.into(), value: namespace.into() });
        }
    }
    Ok(())
}

/// Check the bindings that will be serialized, after fragments from independent
/// source scopes have been merged. A failed binding must never be emitted as
/// another extension's property merely because the lexical prefix is equal.
pub(super) fn validate_namespace_bindings(element: &Element, inherited: &BTreeMap<String, String>) -> Result<(), String> {
    let mut scope = inherited.clone();
    scope.insert("xml".into(), XML_NAMESPACE.into());
    for attribute in &element.attributes {
        if attribute.namespace == XMLNS_NAMESPACE {
            let prefix = attribute.name.strip_prefix("xmlns:").unwrap_or("");
            scope.insert(prefix.into(), attribute.value.clone());
        }
    }
    let bound = |name: &str, expected: &str, attribute: bool| -> Result<(), String> {
        let prefix = name.split_once(':').map(|(prefix, _)| prefix).unwrap_or("");
        let actual = if attribute && prefix.is_empty() { "" } else { scope.get(prefix).map(String::as_str).unwrap_or("") };
        if actual != expected { return Err(format!("Source namespace prefix {prefix:?} conflicts after property inheritance for {name:?}.")); }
        Ok(())
    };
    bound(&element.name, &element.namespace, false)?;
    for attribute in &element.attributes {
        if attribute.namespace != XMLNS_NAMESPACE { bound(&attribute.name, &attribute.namespace, true)?; }
    }
    for child in element.children() { validate_namespace_bindings(child, &scope)?; }
    Ok(())
}

pub(super) fn valid_name(name: &str) -> bool {
    let mut characters = name.chars();
    characters.next().is_some_and(|character| character == '_' || character.is_alphabetic()) &&
        characters.all(|character| character == '_' || character == '-' || character == '.' || character.is_alphanumeric())
}

fn qualified_name(name: &str) -> bool {
    name.split_once(':').map_or_else(|| valid_name(name), |(prefix, local)| valid_name(prefix) && valid_name(local))
}

fn namespace(result: ResolveResult<'_>) -> Result<String, String> {
    match result {
        ResolveResult::Bound(value) => Ok(value.as_ref().to_owned()),
        ResolveResult::Unbound => Ok(String::new()),
        ResolveResult::Unknown(prefix) => Err(format!("Office XML contains undeclared namespace prefix {prefix:?}.")),
    }
}

pub(super) fn valid_xml_characters(text: &str) -> bool {
    !text.chars().any(|character| (character < '\u{20}' && !matches!(character, '\t' | '\n' | '\r')) || matches!(character, '\u{fffe}' | '\u{ffff}'))
}

fn push_text(stack: &mut [Element], text: String) -> Result<(), String> {
    if !valid_xml_characters(&text) { return Err("Office XML contains invalid XML characters.".into()); }
    if let Some(node) = stack.last_mut() { node.content.push(Content::Text(text)); }
    else if !text.trim().is_empty() { return Err("Office XML contains text outside its root element.".into()); }
    Ok(())
}

/// Read source resources and marker fragments through the same namespace-aware,
/// bounded parser. Re-emission never inserts unchecked source XML.
pub(super) fn parse_xml(source: &str, document_nodes: bool) -> Result<Element, String> {
    if source.len() > XML_LIMIT { return Err("Office XML exceeds the 16 MiB limit.".into()); }
    let mut reader = NsReader::from_str(source);
    reader.config_mut().enable_all_checks(true);
    let mut stack = Vec::<Element>::new();
    let mut root = None;
    let mut elements = 0;
    let mut declaration = false;
    loop {
        let (resolved, event) = reader.read_resolved_event().map_err(|error| error.to_string())?;
        let empty = matches!(&event, Event::Empty(_));
        match event {
            Event::Start(tag) | Event::Empty(tag) => {
                elements += 1;
                if elements > XML_ELEMENT_LIMIT || stack.len() >= XML_DEPTH_LIMIT { return Err("Office XML exceeds the element or nesting limit.".into()); }
                let name = tag.name().as_ref().to_owned();
                if !qualified_name(&name) { return Err("Office XML contains an invalid element name.".into()); }
                let mut node = Element { name, namespace: namespace(resolved)?, local: tag.local_name().as_ref().to_owned(), attributes: Vec::new(), content: Vec::new(), empty };
                let mut attributes = HashSet::new();
                for attribute in tag.attributes() {
                    let attribute = attribute.map_err(|error| error.to_string())?;
                    let name = attribute.key.as_ref().to_owned();
                    if !qualified_name(&name) { return Err("Office XML contains an invalid attribute name.".into()); }
                    let declaration = name == "xmlns" || name.starts_with("xmlns:");
                    let (resolved, local) = reader.resolver().resolve_attribute(attribute.key);
                    let attribute_namespace = if declaration { XMLNS_NAMESPACE.to_owned() } else { namespace(resolved)? };
                    let local = local.as_ref().to_owned();
                    if !attributes.insert((attribute_namespace.clone(), local.clone())) { return Err("Office XML contains duplicate attributes.".into()); }
                    let value = attribute.normalized_value_with(XmlVersion::Implicit1_0, 1, quick_xml::escape::resolve_xml_entity).map_err(|error| error.to_string())?.into_owned();
                    if !valid_xml_characters(&value) { return Err("Office XML contains invalid XML characters.".into()); }
                    if declaration && (name == "xmlns:xmlns" || name == "xmlns:xml" && value != XML_NAMESPACE || value == XMLNS_NAMESPACE || value == XML_NAMESPACE && name != "xmlns:xml" || name != "xmlns" && value.is_empty()) { return Err("Office XML contains an invalid namespace declaration.".into()); }
                    node.attributes.push(Attribute { name, namespace: attribute_namespace, local, value });
                }
                if empty {
                    if let Some(parent) = stack.last_mut() { parent.content.push(Content::Element(node)); }
                    else if root.replace(node).is_some() { return Err("Office XML has multiple root elements.".into()); }
                } else { stack.push(node); }
            }
            Event::End(_) => {
                let node = stack.pop().ok_or("Office XML has an unmatched closing element.")?;
                if let Some(parent) = stack.last_mut() { parent.content.push(Content::Element(node)); }
                else if root.replace(node).is_some() { return Err("Office XML has multiple root elements.".into()); }
            }
            Event::Text(text) => push_text(&mut stack, text.xml_content(XmlVersion::Implicit1_0).into_owned())?,
            Event::CData(text) => push_text(&mut stack, text.xml_content(XmlVersion::Implicit1_0).into_owned())?,
            Event::GeneralRef(reference) => {
                let text = if let Some(character) = reference.resolve_char_ref().map_err(|error| error.to_string())? { character.to_string() }
                else { quick_xml::escape::resolve_xml_entity(&reference.as_ref()).ok_or("Office XML contains an unsupported entity.")?.to_owned() };
                push_text(&mut stack, text)?;
            }
            Event::Decl(value) if document_nodes && !declaration && root.is_none() && stack.is_empty() => {
                if value.version().map_err(|error| error.to_string())?.as_ref() != "1.0" { return Err("Office XML must use XML 1.0.".into()); }
                if let Some(encoding) = value.encoding() {
                    let encoding = encoding.map_err(|error| error.to_string())?;
                    if !encoding.eq_ignore_ascii_case("utf-8") && !encoding.eq_ignore_ascii_case("utf8") { return Err("Office XML source resources must be UTF-8.".into()); }
                }
                declaration = true;
            }
            Event::Comment(_) if document_nodes => {},
            Event::Eof => { if !stack.is_empty() { return Err("Office XML has unclosed elements.".into()); } break; },
            Event::DocType(_) | Event::Decl(_) | Event::PI(_) | Event::Comment(_) => return Err("Office XML contains an unsupported declaration or processing node.".into()),
        }
    }
    root.ok_or_else(|| "Office XML has no root element.".into())
}
