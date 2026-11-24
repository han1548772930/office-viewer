//! Actual OPC custom-XML stores and retained related resources, separate from editor sidecars.
use std::collections::{BTreeMap, BTreeSet};
use std::borrow::Cow;
use base64::{Engine, engine::general_purpose::STANDARD};
use quick_xml::{Reader, XmlVersion, events::Event};
use crate::xml_namespaces::{XmlNamespaces, validate_declaration, validate_pi_target, xml_character};
use serde::{Deserialize, Serialize};

pub const CUSTOM_XML_NAMESPACE: &str = "http://schemas.openxmlformats.org/officeDocument/2006/customXml";
pub const STRICT_CUSTOM_XML_NAMESPACE: &str = "http://purl.oclc.org/ooxml/officeDocument/customXml";
pub const RELATIONSHIP_NAMESPACE: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
pub const STRICT_RELATIONSHIP_NAMESPACE: &str = "http://purl.oclc.org/ooxml/package/relationships";
pub const CUSTOM_XML_RELATIONSHIP: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";
pub const CUSTOM_XML_PROPERTIES_RELATIONSHIP: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps";
pub const CUSTOM_XML_PROPERTIES_CONTENT_TYPE: &str = "application/vnd.openxmlformats-officedocument.customXmlProperties+xml";
pub const RELATIONSHIP_CONTENT_TYPE: &str = "application/vnd.openxmlformats-package.relationships+xml";
pub const MAX_PART_BYTES: usize = 8 * 1024 * 1024;
pub const MAX_TOTAL_BYTES: usize = 64 * 1024 * 1024;
pub const MAX_PARTS: usize = 256;
pub const MAX_RELATED_PARTS: usize = 2048;
pub const MAX_ELEMENTS: usize = 100_000;
pub const MAX_DEPTH: usize = 128;

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "lowercase")]
pub enum CustomXmlEncoding { Utf8, Utf16le, Utf16be }

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct DocumentCustomXmlResource {
    pub part_name: String,
    pub content_type: String,
    pub data_base64: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct DocumentCustomXmlPart {
    pub part_name: String,
    pub xml: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub content_type: Option<String>,
    #[serde(rename = "storeItemID", skip_serializing_if = "Option::is_none")]
    pub store_item_id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub properties_part_name: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub properties_xml: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub relationships_xml: Option<String>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub related_parts: Vec<DocumentCustomXmlResource>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub encoding: Option<CustomXmlEncoding>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub byte_order_mark: Option<bool>,
}

#[derive(Debug, Clone)]
pub struct CustomXmlRelationship { pub id: String, pub kind: String, pub target: String, pub external: bool }

pub fn relationship_is(kind: &str, suffix: &str) -> bool {
    kind.strip_prefix("http://schemas.openxmlformats.org/officeDocument/2006/relationships/")
        .or_else(|| kind.strip_prefix("http://purl.oclc.org/ooxml/officeDocument/relationships/")).is_some_and(|kind| kind == suffix)
}

pub fn valid_part_name(value: &str) -> bool {
    canonical_part_name(value).is_ok()
}

pub fn valid_content_type(value: &str) -> bool {
    let token = |value: &str| !value.is_empty() && value.bytes().all(|byte| byte.is_ascii_alphanumeric()
        || matches!(byte, b'!' | b'#' | b'$' | b'%' | b'&' | b'\'' | b'*' | b'+' | b'-' | b'.' | b'^' | b'_' | b'`' | b'|' | b'~'));
    value.split_once('/').is_some_and(|(kind, subtype)| token(kind) && token(subtype))
}

fn encode_part_segment(segment: &str, escapes: bool) -> Result<String, String> {
    let mut decoded = Vec::with_capacity(segment.len());
    let bytes = segment.as_bytes();
    let mut index = 0usize;
    let hex = |byte: u8| match byte { b'0'..=b'9' => Some(byte - b'0'), b'a'..=b'f' => Some(byte - b'a' + 10), b'A'..=b'F' => Some(byte - b'A' + 10), _ => None };
    while index < bytes.len() {
        if escapes && bytes[index] == b'%' {
            let high = bytes.get(index + 1).and_then(|byte| hex(*byte)).ok_or("DOCX_CUSTOM_XML_INVALID_PART: invalid percent escape")?;
            let low = bytes.get(index + 2).and_then(|byte| hex(*byte)).ok_or("DOCX_CUSTOM_XML_INVALID_PART: invalid percent escape")?;
            decoded.push(high * 16 + low); index += 3;
        } else { decoded.push(bytes[index]); index += 1; }
    }
    let text = std::str::from_utf8(&decoded).map_err(|_| "DOCX_CUSTOM_XML_INVALID_PART: percent escapes are not UTF-8")?;
    if text.is_empty() || matches!(text, "." | "..") || text.ends_with('.') || text.chars().any(|character| character.is_control() || matches!(character, '/' | '\\')) {
        return Err("DOCX_CUSTOM_XML_INVALID_PART: invalid package path segment".into());
    }
    const HEX: &[u8; 16] = b"0123456789ABCDEF";
    let mut result = String::with_capacity(decoded.len());
    for byte in decoded {
        if byte.is_ascii_alphanumeric() || matches!(byte, b'-' | b'_' | b'.' | b'~' | b'!' | b'\'' | b'(' | b')' | b'*') { result.push(byte as char); }
        else { result.push('%'); result.push(HEX[(byte >> 4) as usize] as char); result.push(HEX[(byte & 15) as usize] as char); }
    }
    Ok(result)
}

fn canonical_name(value: &str, escapes: bool) -> Result<String, String> {
    if value.is_empty() || value.encode_utf16().count() > 1024 || value.starts_with('/') || value.chars().any(|character| character.is_control() || character == '\\'
        || escapes && matches!(character, ':' | '#' | '?')) { return Err("DOCX_CUSTOM_XML_INVALID_PART: invalid package part name".into()); }
    let result = value.split('/').map(|segment| encode_part_segment(segment, escapes)).collect::<Result<Vec<_>, _>>()?.join("/");
    if result.len() > 1024 { return Err("DOCX_CUSTOM_XML_INVALID_PART: encoded package part name exceeds its limit".into()); }
    Ok(result)
}

/// Models and OPC attributes carry a canonical relative URI; ZIP names are separate.
pub fn canonical_part_name(value: &str) -> Result<String, String> { canonical_name(value, true) }
pub fn part_name_key(value: &str) -> Result<String, String> { Ok(canonical_part_name(value)?.to_ascii_lowercase()) }
pub fn part_name_from_zip(value: &str) -> Result<String, String> { canonical_name(value, false) }

pub fn zip_part_name(value: &str) -> Result<String, String> {
    let name = canonical_part_name(value)?;
    let mut bytes = Vec::with_capacity(name.len());
    let mut index = 0usize;
    while index < name.len() {
        if name.as_bytes()[index] == b'%' {
            bytes.push(u8::from_str_radix(&name[index + 1..index + 3], 16).map_err(|_| "DOCX_CUSTOM_XML_INVALID_PART: invalid percent escape")?); index += 3;
        } else { bytes.push(name.as_bytes()[index]); index += 1; }
    }
    String::from_utf8(bytes).map_err(|_| "DOCX_CUSTOM_XML_INVALID_PART: decoded package name is not UTF-8".into())
}

pub fn relationship_part(owner: &str) -> String {
    match owner.rsplit_once('/') { Some((parent, file)) => format!("{parent}/_rels/{file}.rels"), None => format!("_rels/{owner}.rels") }
}

pub fn relationship_owner(part: &str) -> Option<String> {
    let name = canonical_part_name(part).ok()?;
    let lower = name.to_ascii_lowercase();
    if !lower.ends_with(".rels") { return None; }
    let owner = if let Some(index) = lower.rfind("/_rels/") { format!("{}/{}", &name[..index], &name[index + 7..name.len() - 5]) }
        else if lower.starts_with("_rels/") { name[6..name.len() - 5].to_owned() } else { return None; };
    canonical_part_name(&owner).ok()
}

pub fn resolve_target(owner: &str, target: &str) -> Result<String, String> {
    if target.is_empty() || target.encode_utf16().count() > 2048 || target.starts_with("//") || target.chars().any(|character| character.is_control() || matches!(character, '\\' | ':' | '#' | '?')) { return Err("DOCX_CUSTOM_XML_INVALID_TARGET: invalid package relationship target".into()); }
    let mut segments = if target.starts_with('/') || owner.is_empty() { Vec::new() }
        else { let name = canonical_part_name(owner)?; name.rsplit_once('/').map_or(Vec::new(), |(parent, _)| parent.split('/').map(str::to_owned).collect()) };
    for segment in target.strip_prefix('/').unwrap_or(target).split('/') {
        match segment { "." => {}, ".." => { segments.pop().ok_or("DOCX_CUSTOM_XML_INVALID_TARGET: target leaves the package")?; }, _ => segments.push(canonical_part_name(segment)?) }
    }
    canonical_part_name(&segments.join("/"))
}

/// Streaming validation allocates no text/tree copy. XML content remains exactly in its source String.
pub fn validate_xml(xml: &str) -> Result<(), String> {
    if xml.len() > MAX_PART_BYTES { return Err("DOCX_CUSTOM_XML_RESOURCE_LIMIT: XML part exceeds 8 MiB".into()); }
    if !xml.chars().all(xml_character) { return Err("DOCX_CUSTOM_XML_INVALID_XML: invalid XML character".into()); }
    let mut reader = Reader::from_str(xml);
    reader.config_mut().enable_all_checks(true);
    let mut namespaces = XmlNamespaces::default();
    let mut depth = 0usize;
    let mut elements = 0usize;
    let mut roots = 0usize;
    let mut declaration_seen = false;
    loop {
        let position = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(|error| format!("DOCX_CUSTOM_XML_INVALID_XML: {error}"))?;
        let empty = matches!(&event, Event::Empty(_));
        match event {
            Event::Start(tag) | Event::Empty(tag) => {
                elements += 1;
                if depth == 0 { roots += 1; }
                if elements > MAX_ELEMENTS || depth >= MAX_DEPTH || roots > 1 { return Err("DOCX_CUSTOM_XML_RESOURCE_LIMIT: XML element count/depth or root count exceeds its contract".into()); }
                namespaces.push(&tag, reader.decoder()).map_err(|error| format!("DOCX_CUSTOM_XML_INVALID_XML: {error}"))?;
                namespaces.namespace(tag.name(), false).map_err(|error| format!("DOCX_CUSTOM_XML_INVALID_XML: {error}"))?;
                let mut names = BTreeSet::new();
                for attribute in tag.attributes() {
                    let attribute = attribute.map_err(|error| format!("DOCX_CUSTOM_XML_INVALID_XML: {error}"))?;
                    let value = attribute.decoded_and_normalized_value_with(XmlVersion::Implicit1_0, reader.decoder(), 1, quick_xml::escape::resolve_xml_entity).map_err(|error| format!("DOCX_CUSTOM_XML_INVALID_XML: {error}"))?;
                    if !value.chars().all(xml_character) { return Err("DOCX_CUSTOM_XML_INVALID_XML: invalid attribute character".into()); }
                    if attribute.key.as_namespace_binding().is_some() { continue; }
                    let namespace = namespaces.namespace(attribute.key, true).map_err(|error| format!("DOCX_CUSTOM_XML_INVALID_XML: {error}"))?.unwrap_or_default();
                    if !names.insert((namespace, attribute.key.local_name().into_inner())) { return Err("DOCX_CUSTOM_XML_INVALID_XML: duplicate expanded attribute name".into()); }
                }
                if empty { namespaces.pop(); } else { depth += 1; }
            },
            Event::End(_) => { depth = depth.checked_sub(1).ok_or("DOCX_CUSTOM_XML_INVALID_XML: unbalanced end tag")?; namespaces.pop(); },
            Event::Text(value) => {
                let text = value.decode().map_err(|error| error.to_string())?;
                if text.contains("]]>") || depth == 0 && !text.chars().all(|character| matches!(character, ' ' | '\t' | '\r' | '\n')) { return Err("DOCX_CUSTOM_XML_INVALID_XML: invalid text outside/inside XML root".into()); }
            },
            Event::CData(_) if depth == 0 => return Err("DOCX_CUSTOM_XML_INVALID_XML: CDATA outside XML root".into()),
            Event::GeneralRef(value) => {
                if depth == 0 { return Err("DOCX_CUSTOM_XML_INVALID_XML: entity outside XML root".into()); }
                let named = value.decode().map_err(|error| error.to_string())?;
                let resolved = value.resolve_char_ref().map_err(|error| error.to_string())?;
                if resolved.is_some_and(|character| !xml_character(character)) || resolved.is_none() && quick_xml::escape::resolve_xml_entity(&named).is_none() { return Err("DOCX_CUSTOM_XML_INVALID_XML: unknown or invalid entity".into()); }
            },
            Event::DocType(_) => return Err("DOCX_CUSTOM_XML_INVALID_XML: document types and entity declarations are not permitted".into()),
            Event::Decl(declaration) => {
                if declaration_seen || position > 3 || !xml.trim_start_matches('\u{feff}').starts_with("<?xml") { return Err("DOCX_CUSTOM_XML_INVALID_XML: misplaced or repeated XML declaration".into()); }
                validate_declaration(declaration.as_ref()).map_err(|error| format!("DOCX_CUSTOM_XML_INVALID_XML: {error}"))?;
                declaration_seen = true;
            },
            Event::PI(instruction) => {
                validate_pi_target(instruction.target()).map_err(|error| format!("DOCX_CUSTOM_XML_INVALID_XML: {error}"))?;
                if !instruction.content().is_empty() && !instruction.content()[0].is_ascii_whitespace() { return Err("DOCX_CUSTOM_XML_INVALID_XML: invalid processing-instruction separator".into()); }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if roots != 1 || depth != 0 { return Err("DOCX_CUSTOM_XML_INVALID_XML: missing or unfinished XML root".into()); }
    Ok(())
}

pub fn properties_store_id(xml: &str) -> Result<String, String> {
    validate_xml(xml)?;
    let root = crate::chart_xml::parse(xml)?;
    if root.name != "datastoreItem" || !matches!(root.namespace.as_deref(), Some(CUSTOM_XML_NAMESPACE | STRICT_CUSTOM_XML_NAMESPACE)) { return Err("DOCX_CUSTOM_XML_INVALID_PROPERTIES: expected datastoreItem properties root".into()); }
    let id = root.attribute_ns(CUSTOM_XML_NAMESPACE, "itemID").or_else(|| root.attribute_ns(STRICT_CUSTOM_XML_NAMESPACE, "itemID")).ok_or("DOCX_CUSTOM_XML_INVALID_PROPERTIES: store ID is missing")?;
    if id.is_empty() || id.encode_utf16().count() > 128 || id.chars().any(char::is_control) { return Err("DOCX_CUSTOM_XML_INVALID_PROPERTIES: invalid store ID".into()); }
    Ok(id.to_owned())
}

pub fn relationships(xml: &str) -> Result<Vec<CustomXmlRelationship>, String> {
    validate_xml(xml)?;
    let root = crate::chart_xml::parse(xml)?;
    if root.name != "Relationships" || !matches!(root.namespace.as_deref(), Some(RELATIONSHIP_NAMESPACE | STRICT_RELATIONSHIP_NAMESPACE)) { return Err("DOCX_CUSTOM_XML_INVALID_RELATIONSHIPS: expected OPC Relationships root".into()); }
    if !root.text.trim().is_empty() { return Err("DOCX_CUSTOM_XML_INVALID_RELATIONSHIPS: relationships root contains text".into()); }
    let mut ids = BTreeSet::new();
    let mut result = Vec::new();
    for relationship in &root.children {
        if relationship.name != "Relationship" || relationship.namespace != root.namespace || !relationship.children.is_empty() || !relationship.text.trim().is_empty() { return Err("DOCX_CUSTOM_XML_INVALID_RELATIONSHIPS: unknown relationship element or content".into()); }
        let get = |name: &str| relationship.expanded_attrs.get("").and_then(|attributes| attributes.get(name)).filter(|value| !value.is_empty()).cloned().ok_or_else(|| format!("DOCX_CUSTOM_XML_INVALID_RELATIONSHIPS: missing {name}"));
        let id = get("Id")?;
        if !ids.insert(id.clone()) { return Err("DOCX_CUSTOM_XML_INVALID_RELATIONSHIPS: duplicate relationship ID".into()); }
        let mode = relationship.expanded_attrs.get("").and_then(|attributes| attributes.get("TargetMode")).map(String::as_str);
        if mode.is_some_and(|mode| !matches!(mode, "Internal" | "External")) { return Err("DOCX_CUSTOM_XML_INVALID_RELATIONSHIPS: invalid target mode".into()); }
        result.push(CustomXmlRelationship { id, kind: get("Type")?, target: get("Target")?, external: mode == Some("External") });
    }
    Ok(result)
}

pub fn validate_prefix_mappings(value: &str) -> Result<(), String> {
    if value.encode_utf16().count() > 8192 { return Err("DOCX_CUSTOM_XML_INVALID_BINDING: namespace mappings exceed their limit".into()); }
    let xml = format!("<binding {value}/>");
    validate_xml(&xml)?;
    let root = crate::chart_xml::parse(&xml)?;
    if root.name != "binding" || !root.attrs.is_empty() || !root.children.is_empty() || !root.text.is_empty() {
        return Err("DOCX_CUSTOM_XML_INVALID_BINDING: namespace mappings must contain xmlns declarations only".into());
    }
    Ok(())
}

/// Decode before classifying an editor sidecar; callers supply their distinct
/// bounded source budget. Actual data stores additionally call validate_xml.
pub fn decode_xml_text(bytes: &[u8], maximum: usize) -> Result<(String, CustomXmlEncoding, bool), String> {
    if bytes.len() > maximum { return Err("DOCX_CUSTOM_XML_RESOURCE_LIMIT: source part exceeds its byte limit".into()); }
    let (encoding, bom, content) = if bytes.starts_with(&[0xff, 0xfe]) { (CustomXmlEncoding::Utf16le, true, &bytes[2..]) }
        else if bytes.starts_with(&[0xfe, 0xff]) { (CustomXmlEncoding::Utf16be, true, &bytes[2..]) }
        else if bytes.starts_with(&[0xef, 0xbb, 0xbf]) { (CustomXmlEncoding::Utf8, true, &bytes[3..]) }
        else if bytes.starts_with(&[b'<', 0]) { (CustomXmlEncoding::Utf16le, false, bytes) }
        else if bytes.starts_with(&[0, b'<']) { (CustomXmlEncoding::Utf16be, false, bytes) }
        else { (CustomXmlEncoding::Utf8, false, bytes) };
    let xml = match encoding {
        CustomXmlEncoding::Utf8 => String::from_utf8(content.to_vec()).map_err(|error| format!("DOCX_CUSTOM_XML_INVALID_ENCODING: {error}"))?,
        CustomXmlEncoding::Utf16le | CustomXmlEncoding::Utf16be => {
            if content.len() % 2 != 0 { return Err("DOCX_CUSTOM_XML_INVALID_ENCODING: incomplete UTF-16 code unit".into()); }
            let units = content.chunks_exact(2).map(|bytes| if encoding == CustomXmlEncoding::Utf16le { u16::from_le_bytes([bytes[0], bytes[1]]) } else { u16::from_be_bytes([bytes[0], bytes[1]]) });
            char::decode_utf16(units).collect::<Result<String, _>>().map_err(|error| format!("DOCX_CUSTOM_XML_INVALID_ENCODING: {error}"))?
        },
    };
    Ok((xml, encoding, bom))
}

pub fn decode_xml(bytes: &[u8]) -> Result<(String, CustomXmlEncoding, bool), String> {
    let result = decode_xml_text(bytes, MAX_PART_BYTES)?;
    validate_xml(&result.0)?;
    Ok(result)
}

pub fn xml_root_identity(xml: &str) -> Result<(String, Option<String>), String> {
    let mut reader = Reader::from_str(xml);
    let mut namespaces = XmlNamespaces::default();
    loop {
        let event = reader.read_event().map_err(|error| error.to_string())?;
        match event {
            Event::Start(tag) | Event::Empty(tag) => {
                namespaces.push(&tag, reader.decoder()).map_err(|error| format!("DOCX_CUSTOM_XML_INVALID_XML: {error}"))?;
                let namespace = namespaces.namespace(tag.name(), false).map_err(|error| format!("DOCX_CUSTOM_XML_INVALID_XML: {error}"))?.map(str::to_owned);
                return Ok((std::str::from_utf8(tag.local_name().as_ref()).map_err(|error| error.to_string())?.to_owned(), namespace));
            },
            Event::Eof | Event::DocType(_) => return Err("DOCX_CUSTOM_XML_INVALID_XML: missing XML root or forbidden DTD".into()),
            _ => {},
        }
    }
}

fn xml_encoding_range(xml: &str) -> Option<std::ops::Range<usize>> {
    if !xml.starts_with("<?xml") || !xml.as_bytes().get(5).is_some_and(u8::is_ascii_whitespace) { return None; }
    let end = xml.find("?>")?;
    let declaration = &xml[..end];
    let start = declaration.find("encoding")?;
    if start != 0 && !declaration.as_bytes()[start - 1].is_ascii_whitespace() { return None; }
    let tail = declaration[start + "encoding".len()..].trim_start().strip_prefix('=')?.trim_start();
    let quote @ ('\'' | '"') = tail.chars().next()? else { return None; };
    let length = tail[1..].find(quote)?;
    let value_start = end - tail.len() + 1;
    Some(value_start..value_start + length)
}

pub fn normalize_xml_encoding(mut xml: String) -> String {
    if let Some(range) = xml_encoding_range(&xml) { xml.replace_range(range, "UTF-8"); }
    xml
}

fn xml_for_encoding(value: &str, encoding: CustomXmlEncoding) -> Cow<'_, str> {
    let declaration = match encoding { CustomXmlEncoding::Utf8 => "UTF-8", CustomXmlEncoding::Utf16le => "UTF-16LE", CustomXmlEncoding::Utf16be => "UTF-16BE" };
    let mut xml = Cow::Borrowed(value);
    if let Some(range) = xml_encoding_range(&xml) {
        let existing = &xml[range.clone()];
        let matching = existing.eq_ignore_ascii_case(declaration)
            || encoding != CustomXmlEncoding::Utf8 && existing.eq_ignore_ascii_case("UTF-16");
        if !matching { xml.to_mut().replace_range(range, declaration); }
    }
    xml
}

fn encoded_xml_length(xml: &str, encoding: CustomXmlEncoding, bom: bool) -> Result<usize, String> {
    let length = match encoding { CustomXmlEncoding::Utf8 => xml.len().checked_add(usize::from(bom) * 3),
        _ => xml.encode_utf16().count().checked_mul(2).and_then(|length| length.checked_add(usize::from(bom) * 2)) };
    length.ok_or_else(|| "DOCX_CUSTOM_XML_RESOURCE_LIMIT: encoded size overflow".into())
}

pub fn encode_xml(part: &DocumentCustomXmlPart) -> Result<Vec<u8>, String> {
    validate_xml(&part.xml)?;
    let encoding = part.encoding.unwrap_or(CustomXmlEncoding::Utf8);
    let bom = part.byte_order_mark.unwrap_or(false);
    let xml = xml_for_encoding(&part.xml, encoding);
    validate_xml(&xml)?;
    let length = encoded_xml_length(&xml, encoding, bom)?;
    if length > MAX_PART_BYTES { return Err("DOCX_CUSTOM_XML_RESOURCE_LIMIT: encoded XML exceeds 8 MiB".into()); }
    let mut bytes = Vec::new();
    bytes.try_reserve_exact(length).map_err(|_| "DOCX_CUSTOM_XML_RESOURCE_LIMIT: XML byte allocation failed")?;
    match encoding {
        CustomXmlEncoding::Utf8 => { if bom { bytes.extend_from_slice(&[0xef, 0xbb, 0xbf]); } bytes.extend_from_slice(xml.as_bytes()); },
        CustomXmlEncoding::Utf16le | CustomXmlEncoding::Utf16be => {
            if bom { bytes.extend_from_slice(if encoding == CustomXmlEncoding::Utf16le { &[0xff, 0xfe] } else { &[0xfe, 0xff] }); }
            for unit in xml.encode_utf16() { bytes.extend_from_slice(&if encoding == CustomXmlEncoding::Utf16le { unit.to_le_bytes() } else { unit.to_be_bytes() }); }
        },
    }
    Ok(bytes)
}

pub fn encode_resource(bytes: &[u8]) -> String { STANDARD.encode(bytes) }

/// Canonical base64 validation without allocating the decoded resource.
pub fn resource_byte_length(value: &str) -> Result<usize, String> {
    if value.len() > MAX_PART_BYTES.div_ceil(3) * 4 { return Err("DOCX_CUSTOM_XML_RESOURCE_LIMIT: retained resource exceeds 8 MiB".into()); }
    if value.len() % 4 != 0 { return Err("DOCX_CUSTOM_XML_INVALID_RESOURCE: incomplete base64".into()); }
    let padding = if value.ends_with("==") { 2 } else { usize::from(value.ends_with('=')) };
    let end = value.len() - padding;
    let digit = |byte: u8| -> Option<u8> { match byte { b'A'..=b'Z' => Some(byte - b'A'), b'a'..=b'z' => Some(byte - b'a' + 26),
        b'0'..=b'9' => Some(byte - b'0' + 52), b'+' => Some(62), b'/' => Some(63), _ => None } };
    if value.as_bytes()[..end].iter().any(|byte| digit(*byte).is_none()) { return Err("DOCX_CUSTOM_XML_INVALID_RESOURCE: invalid base64 character".into()); }
    if padding != 0 && (end < 2 || digit(value.as_bytes()[end - 1]).is_none_or(|digit| digit % if padding == 2 { 16 } else { 4 } != 0)) {
        return Err("DOCX_CUSTOM_XML_INVALID_RESOURCE: noncanonical base64 trailing bits".into());
    }
    let bytes = value.len() / 4 * 3 - padding;
    if bytes > MAX_PART_BYTES { return Err("DOCX_CUSTOM_XML_RESOURCE_LIMIT: retained resource exceeds 8 MiB".into()); }
    Ok(bytes)
}

pub fn decode_resource(value: &str) -> Result<Vec<u8>, String> {
    resource_byte_length(value)?;
    let bytes = STANDARD.decode(value).map_err(|error| format!("DOCX_CUSTOM_XML_INVALID_RESOURCE: {error}"))?;
    if bytes.len() > MAX_PART_BYTES { return Err("DOCX_CUSTOM_XML_RESOURCE_LIMIT: retained resource exceeds 8 MiB".into()); }
    Ok(bytes)
}

pub fn validate_parts(parts: &[DocumentCustomXmlPart]) -> Result<(), String> {
    if parts.len() > MAX_PARTS { return Err("DOCX_CUSTOM_XML_RESOURCE_LIMIT: too many data sources".into()); }
    let mut names = BTreeSet::new();
    let mut stores = BTreeSet::new();
    let mut resources = BTreeMap::<String, (&str, &str)>::new();
    let mut total = 0usize;
    let mut resource_count = 0usize;
    let mut count = |bytes: usize| -> Result<(), String> { total = total.checked_add(bytes).filter(|total| *total <= MAX_TOTAL_BYTES).ok_or("DOCX_CUSTOM_XML_RESOURCE_LIMIT: data sources exceed 64 MiB")?; Ok(()) };
    for part in parts {
        let name_key = part_name_key(&part.part_name)?;
        if name_key.ends_with(".rels") || name_key == "%5bcontent_types%5d.xml" || !names.insert(name_key) { return Err("DOCX_CUSTOM_XML_INVALID_SOURCE: invalid or duplicate data part name".into()); }
        validate_xml(&part.xml)?;
        if part.content_type.as_ref().is_some_and(|kind| !valid_content_type(kind)
            || !kind.eq_ignore_ascii_case("application/xml") && !kind.eq_ignore_ascii_case("text/xml") && !kind.to_ascii_lowercase().ends_with("+xml")) {
            return Err("DOCX_CUSTOM_XML_INVALID_SOURCE: data-store content type must be XML".into());
        }
        let encoding = part.encoding.unwrap_or(CustomXmlEncoding::Utf8);
        let serialized = xml_for_encoding(&part.xml, encoding);
        let encoded_bytes = encoded_xml_length(&serialized, encoding, part.byte_order_mark.unwrap_or(false))?;
        if serialized.len() > MAX_PART_BYTES || encoded_bytes > MAX_PART_BYTES { return Err("DOCX_CUSTOM_XML_RESOURCE_LIMIT: encoded XML exceeds 8 MiB".into()); }
        count(part.xml.len().max(serialized.len()).max(encoded_bytes))?;
        if let Some(id) = &part.store_item_id {
            if id.is_empty() || id.encode_utf16().count() > 128 || id.chars().any(char::is_control) || !stores.insert(id.to_lowercase()) { return Err("DOCX_CUSTOM_XML_INVALID_SOURCE: invalid or duplicate store ID".into()); }
        }
        match (&part.properties_part_name, &part.properties_xml) {
            (Some(name), Some(xml)) => {
                let key = part_name_key(name)?;
                if key.ends_with(".rels") || key == "%5bcontent_types%5d.xml" || !names.insert(key) || Some(properties_store_id(xml)?.to_lowercase()) != part.store_item_id.as_ref().map(|id| id.to_lowercase()) { return Err("DOCX_CUSTOM_XML_INVALID_PROPERTIES: property owner/ID mismatch".into()); }
                let encoded = xml_for_encoding(xml, CustomXmlEncoding::Utf8);
                if encoded.len() > MAX_PART_BYTES { return Err("DOCX_CUSTOM_XML_RESOURCE_LIMIT: encoded properties exceed 8 MiB".into()); }
                count(xml.len().max(encoded.len()))?;
            },
            (None, None) => {}, _ => return Err("DOCX_CUSTOM_XML_INVALID_PROPERTIES: property name and XML must both be present".into()),
        }
        if let Some(xml) = &part.relationships_xml {
            relationships(xml)?;
            let encoded = xml_for_encoding(xml, CustomXmlEncoding::Utf8);
            if encoded.len() > MAX_PART_BYTES { return Err("DOCX_CUSTOM_XML_RESOURCE_LIMIT: encoded relationships exceed 8 MiB".into()); }
            count(xml.len().max(encoded.len()))?;
        }
        resource_count = resource_count.checked_add(part.related_parts.len()).filter(|count| *count <= MAX_RELATED_PARTS).ok_or("DOCX_CUSTOM_XML_RESOURCE_LIMIT: too many related resources")?;
        for resource in &part.related_parts {
            let key = part_name_key(&resource.part_name)?;
            if key == "%5bcontent_types%5d.xml" || !valid_content_type(&resource.content_type) { return Err("DOCX_CUSTOM_XML_INVALID_RESOURCE: invalid resource name or content type".into()); }
            resource_byte_length(&resource.data_base64)?;
            let value = (resource.content_type.as_str(), resource.data_base64.as_str());
            if let Some(previous) = resources.insert(key, value) { if !previous.0.eq_ignore_ascii_case(value.0) || previous.1 != value.1 { return Err("DOCX_CUSTOM_XML_INVALID_RESOURCE: duplicate resource content conflict".into()); } }
            else { count(resource.data_base64.len())?; }
        }
    }
    if resources.keys().any(|name| names.contains(name)) { return Err("DOCX_CUSTOM_XML_INVALID_RESOURCE: data/property part conflicts with retained resource".into()); }
    Ok(())
}

#[cfg(test)]
#[path = "document_custom_xml_tests.rs"]
mod tests;
