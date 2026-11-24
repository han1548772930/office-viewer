//! Normalized XML 1.0 namespace bindings shared by Office tree and streaming readers.
//! quick-xml's NamespaceResolver stores raw xmlns values; normalize before adding them.
use quick_xml::{XmlVersion, encoding::Decoder, events::BytesStart,
    name::{Namespace, NamespaceResolver, QName, ResolveResult}};

const XML_NAMESPACE: &str = "http://www.w3.org/XML/1998/namespace";
const XMLNS_NAMESPACE: &str = "http://www.w3.org/2000/xmlns/";
const MAX_DECLARATIONS: usize = 256;

fn name_start(character: char) -> bool {
    matches!(character as u32, 0x41..=0x5a | 0x5f | 0x61..=0x7a | 0xc0..=0xd6 | 0xd8..=0xf6
        | 0xf8..=0x2ff | 0x370..=0x37d | 0x37f..=0x1fff | 0x200c..=0x200d
        | 0x2070..=0x218f | 0x2c00..=0x2fef | 0x3001..=0xd7ff | 0xf900..=0xfdcf
        | 0xfdf0..=0xfffd | 0x10000..=0xeffff)
}

fn name_character(character: char) -> bool {
    name_start(character) || matches!(character as u32, 0x2d | 0x2e | 0x30..=0x39 | 0xb7 | 0x300..=0x36f | 0x203f..=0x2040)
}

pub(crate) fn xml_character(character: char) -> bool {
    matches!(character as u32, 0x9 | 0xa | 0xd | 0x20..=0xd7ff | 0xe000..=0xfffd | 0x10000..=0x10ffff)
}

fn ncname(value: &str) -> bool {
    let mut characters = value.chars();
    characters.next().is_some_and(name_start) && characters.all(name_character)
}

pub(crate) fn validate_qname(value: QName<'_>) -> Result<(), String> {
    let name = std::str::from_utf8(value.as_ref()).map_err(|error| error.to_string())?;
    let valid = match name.split_once(':') { Some((prefix, local)) => ncname(prefix) && ncname(local), None => ncname(name) };
    if !valid { return Err(format!("invalid XML qualified name: {name}")); }
    Ok(())
}

pub(crate) fn validate_pi_target(value: &[u8]) -> Result<(), String> {
    let name = std::str::from_utf8(value).map_err(|error| error.to_string())?;
    let mut characters = name.chars();
    if name.eq_ignore_ascii_case("xml") || !characters.next().is_some_and(|value| value == ':' || name_start(value))
        || !characters.all(|value| value == ':' || name_character(value)) { return Err("invalid XML processing-instruction target".into()); }
    Ok(())
}

#[derive(Default)]
pub(crate) struct XmlNamespaces { resolver: NamespaceResolver }

impl XmlNamespaces {
    pub(crate) fn push(&mut self, element: &BytesStart<'_>, decoder: Decoder) -> Result<(), String> {
        validate_qname(element.name())?;
        // Begin a scope without letting the resolver read unnormalized xmlns attributes.
        self.resolver.push(&BytesStart::new("scope")).map_err(|error| error.to_string())?;
        let mut declarations = 0usize;
        for attribute in element.attributes() {
            let attribute = attribute.map_err(|error| error.to_string())?;
            validate_qname(attribute.key)?;
            if attribute.value.contains(&b'<') { return Err("unescaped angle bracket in XML attribute".into()); }
            let Some(prefix) = attribute.key.as_namespace_binding() else { continue; };
            declarations += 1;
            if declarations > MAX_DECLARATIONS { return Err("XML namespace declarations exceed 256 per element".into()); }
            let namespace = attribute.decoded_and_normalized_value_with(XmlVersion::Implicit1_0, decoder, 1, quick_xml::escape::resolve_xml_entity)
                .map_err(|error| error.to_string())?;
            if !namespace.chars().all(xml_character) { return Err("invalid character in XML namespace".into()); }
            if namespace == XMLNS_NAMESPACE || matches!(prefix, quick_xml::name::PrefixDeclaration::Default) && namespace == XML_NAMESPACE
                || !matches!(prefix, quick_xml::name::PrefixDeclaration::Default) && namespace.is_empty() {
                return Err("invalid reserved or empty XML 1.0 namespace binding".into());
            }
            self.resolver.add(prefix, Namespace(namespace.as_bytes())).map_err(|error| error.to_string())?;
        }
        if element.name().prefix().is_some_and(|prefix| prefix.as_ref() == b"xmlns") { return Err("the xmlns prefix cannot name an XML element".into()); }
        Ok(())
    }

    pub(crate) fn pop(&mut self) { self.resolver.pop(); }

    pub(crate) fn namespace(&self, name: QName<'_>, attribute: bool) -> Result<Option<&str>, String> {
        let (namespace, _) = self.resolver.resolve(name, !attribute);
        match namespace {
            ResolveResult::Bound(value) => Ok(Some(std::str::from_utf8(value.into_inner()).map_err(|error| error.to_string())?)),
            ResolveResult::Unbound => Ok(None),
            ResolveResult::Unknown(prefix) => Err(format!("undeclared XML namespace prefix: {}", String::from_utf8_lossy(&prefix))),
        }
    }
}

pub(crate) fn validate_declaration(value: &[u8]) -> Result<(), String> {
    let invalid = || "invalid XML 1.0 declaration".to_owned();
    let raw = std::str::from_utf8(value).map_err(|_| invalid())?;
    let mut tail = raw.strip_prefix("xml").ok_or_else(invalid)?;
    let whitespace = |character: char| matches!(character, ' ' | '\t' | '\r' | '\n');
    let mut stage = 0usize;
    loop {
        if tail.is_empty() { break; }
        let trimmed = tail.trim_start_matches(whitespace);
        if trimmed.len() == tail.len() { return Err(invalid()); }
        tail = trimmed;
        if tail.is_empty() { break; }
        let name_end = tail.find(|character: char| !character.is_ascii_alphabetic()).ok_or_else(invalid)?;
        let name = &tail[..name_end];
        tail = tail[name_end..].trim_start_matches(whitespace).strip_prefix('=').ok_or_else(invalid)?.trim_start_matches(whitespace);
        let quote @ ('\'' | '"') = tail.chars().next().ok_or_else(invalid)? else { return Err(invalid()); };
        let end = tail[1..].find(quote).ok_or_else(invalid)? + 1;
        let content = &tail[1..end];
        stage = match name {
            "version" if stage == 0 && content == "1.0" => 1,
            "encoding" if stage == 1 && content.as_bytes().first().is_some_and(u8::is_ascii_alphabetic)
                && content.bytes().all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'.' | b'_' | b'-')) => 2,
            "standalone" if matches!(stage, 1 | 2) && matches!(content, "yes" | "no") => 3,
            _ => return Err(invalid()),
        };
        tail = &tail[end + 1..];
    }
    if stage == 0 { return Err(invalid()); }
    Ok(())
}
