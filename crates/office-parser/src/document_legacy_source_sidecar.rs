//! Read only the dedicated historical-resource part. Its Word owner and native
//! styles/theme remain independent from the legacy resources stored as data.

use std::{collections::{BTreeMap, BTreeSet}, io::{Cursor, Read}};

use office_io_common::document_legacy_source::{
    DocumentOfficeSourceOrigin, JSON_BYTE_LIMIT, LEGACY_SOURCE_CONTENT_TYPE,
    LEGACY_SOURCE_PART, LEGACY_SOURCE_RELATIONSHIP, decode_origins,
};
use quick_xml::{NsReader, XmlVersion, events::Event, name::ResolveResult};
use zip::ZipArchive;

use super::{XML_ELEMENT_LIMIT, part_for_type, read_xml};

struct Declaration { name: String, attributes: BTreeMap<String, String> }
impl Declaration {
    fn attribute(&self, name: &str) -> Option<&str> { self.attributes.get(name).map(String::as_str) }
}

fn declarations(xml: &str, root_name: &[u8], namespaces: &[&[u8]]) -> Result<Vec<Declaration>, String> {
    let mut reader = NsReader::from_str(xml);
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut result = Vec::new();
    loop {
        let event = reader.read_event().map_err(|error| error.to_string())?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                let namespace = reader.resolver().resolve_element(element.name()).0;
                let valid_namespace = matches!(namespace, ResolveResult::Bound(value) if namespaces.contains(&value.as_ref()));
                if !valid_namespace { return Err("Office legacy source metadata has an invalid namespace.".into()); }
                if depth == 0 {
                    if root_seen || element.local_name().as_ref() != root_name { return Err("Office legacy source metadata has an invalid root.".into()); }
                    root_seen = true;
                } else {
                    let name = element.local_name();
                    let valid_name = if root_name == b"Types" { matches!(name.as_ref(), b"Default" | b"Override") } else { name.as_ref() == b"Relationship" };
                    if depth != 1 || !valid_name { return Err("Office legacy source metadata has an invalid child element.".into()); }
                    if result.len() >= XML_ELEMENT_LIMIT { return Err("Office legacy source metadata exceeds the element limit.".into()); }
                    let mut attributes = BTreeMap::new();
                    for attribute in element.attributes() {
                        let attribute = attribute.map_err(|error| error.to_string())?;
                        // OPC contract attributes are unqualified. An extension
                        // attribute named x:Type must not masquerade as Type.
                        if attribute.key.as_ref().contains(&b':') || attribute.key.as_ref() == b"xmlns" { continue; }
                        let key = std::str::from_utf8(attribute.key.as_ref()).map_err(|error| error.to_string())?.to_owned();
                        let value = attribute.decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder()).map_err(|error| error.to_string())?.into_owned();
                        if attributes.insert(key, value).is_some() { return Err("Office legacy source metadata repeats an attribute.".into()); }
                    }
                    result.push(Declaration { name: String::from_utf8(name.as_ref().to_vec()).map_err(|error| error.to_string())?, attributes });
                }
                if matches!(event, Event::Start(_)) { depth += 1; }
            },
            Event::End(_) => { depth = depth.checked_sub(1).ok_or("Office legacy source metadata has an unmatched closing tag.")?; },
            Event::Text(text) => if !text.decode().map_err(|error| error.to_string())?.trim().is_empty() { return Err("Office legacy source metadata contains text content.".into()); },
            Event::DocType(_) => return Err("Office legacy source metadata must not declare a DTD.".into()),
            Event::CData(_) | Event::GeneralRef(_) => return Err("Office legacy source metadata contains unsupported text content.".into()),
            Event::Eof => break,
            _ => {},
        }
    }
    if !root_seen || depth != 0 { return Err("Office legacy source metadata is incomplete.".into()); }
    Ok(result)
}

fn relationships(xml: &str) -> Result<Vec<Declaration>, String> {
    let result = declarations(xml, b"Relationships", &[
        b"http://schemas.openxmlformats.org/package/2006/relationships",
        b"http://purl.oclc.org/ooxml/package/relationships",
    ])?;
    let mut ids = BTreeSet::new();
    for entry in &result {
        let id = entry.attribute("Id").filter(|id| !id.is_empty()).ok_or("Office legacy source relationship has no ID.")?;
        if !ids.insert(id) { return Err("Office legacy source relationship ID is duplicated.".into()); }
        if entry.attribute("Type").is_none() || entry.attribute("Target").is_none() { return Err("Office legacy source relationship lacks its type or target.".into()); }
    }
    Ok(result)
}

fn relationship_path(owner: &str) -> String {
    let (directory, file) = owner.rsplit_once('/').unwrap_or(("", owner));
    if directory.is_empty() { format!("_rels/{file}.rels") } else { format!("{directory}/_rels/{file}.rels") }
}

fn resolve_target(owner: &str, target: &str) -> Result<String, String> {
    let mut decoded = Vec::with_capacity(target.len());
    let mut input = target.bytes();
    while let Some(byte) = input.next() {
        decoded.push(if byte == b'%' {
            let high = char::from(input.next().ok_or("Office legacy source URI escape is incomplete.")?).to_digit(16).ok_or("Office legacy source URI escape is invalid.")?;
            let low = char::from(input.next().ok_or("Office legacy source URI escape is incomplete.")?).to_digit(16).ok_or("Office legacy source URI escape is invalid.")?;
            (high * 16 + low) as u8
        } else { byte });
    }
    let target = std::str::from_utf8(&decoded).map_err(|_| "Office legacy source URI escape is not UTF-8.")?;
    if target.is_empty() || target.starts_with("//") || target.contains(['\\', ':', '?', '#']) || target.chars().any(char::is_control) {
        return Err("Office legacy source relationship target has an invalid package path.".into());
    }
    let directory = owner.rsplit_once('/').map_or("", |(directory, _)| directory);
    let path = if target.starts_with('/') || directory.is_empty() { target.trim_start_matches('/').to_owned() } else { format!("{directory}/{target}") };
    let mut components = Vec::new();
    for component in path.split('/') {
        match component {
            "" | "." => {},
            ".." => { if components.pop().is_none() { return Err("Office legacy source relationship target escapes the package root.".into()); } },
            component => components.push(component),
        }
    }
    if components.is_empty() { return Err("Office legacy source relationship target is empty.".into()); }
    Ok(components.join("/"))
}

fn internal_target(owner: &str, relationship: &Declaration) -> Result<String, String> {
    if !matches!(relationship.attribute("TargetMode"), None | Some("Internal")) { return Err("Office legacy source relationship must be internal.".into()); }
    resolve_target(owner, relationship.attribute("Target").ok_or("Office legacy source relationship has no target.")?)
}

pub(crate) fn read_legacy_origins(archive: &mut ZipArchive<Cursor<&[u8]>>) -> Result<Vec<DocumentOfficeSourceOrigin>, String> {
    let types = read_xml(archive, "[Content_Types].xml")?.map(|xml| declarations(&xml, b"Types", &[
        b"http://schemas.openxmlformats.org/package/2006/content-types",
        b"http://purl.oclc.org/ooxml/package/content-types",
    ])).transpose()?.unwrap_or_default();
    let main = resolve_target("", &part_for_type(archive, ".document.main+xml", "word/document.xml")?)?;
    let document_relationships = read_xml(archive, &relationship_path(&main))?.map(|xml| relationships(&xml)).transpose()?.unwrap_or_default();
    let legacy_relationships = document_relationships.iter().filter(|entry| entry.attribute("Type") == Some(LEGACY_SOURCE_RELATIONSHIP)).collect::<Vec<_>>();
    let declares_type = types.iter().any(|entry| entry.attribute("ContentType") == Some(LEGACY_SOURCE_CONTENT_TYPE));
    let mut names = Vec::new();
    for index in 0..archive.len() { names.push(archive.by_index(index).map_err(|error| error.to_string())?.name().to_owned()); }
    let part_count = names.iter().filter(|name| name.as_str() == LEGACY_SOURCE_PART).count();
    if !declares_type && legacy_relationships.is_empty() && part_count == 0 { return Ok(Vec::new()); }
    if legacy_relationships.len() != 1 { return Err("Declared Office legacy source must have one document relationship.".into()); }
    if internal_target(&main, legacy_relationships[0])? != LEGACY_SOURCE_PART { return Err("Office legacy source relationship does not target the registered part.".into()); }
    if part_count != 1 { return Err("Declared Office legacy source part is missing or duplicated.".into()); }

    // With provenance declared, bind its owner to the one current Office
    // document rather than reading an unrelated document's relationship list.
    let root = read_xml(archive, "_rels/.rels")?.ok_or("Office legacy source owner has no root relationships.")?;
    let root = relationships(&root)?;
    let office = root.iter().filter(|entry| matches!(entry.attribute("Type"), Some("http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" | "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument"))).collect::<Vec<_>>();
    if office.len() != 1 || internal_target("", office[0])? != main { return Err("Office legacy source owner does not match the current document part.".into()); }
    if types.iter().filter(|entry| entry.name == "Override" && entry.attribute("ContentType").is_some_and(|kind| kind.ends_with(".document.main+xml"))).count() > 1 {
        return Err("Office legacy source owner has ambiguous main document content types.".into());
    }

    let defaults = types.iter().filter(|entry| entry.name == "Default" && entry.attribute("Extension").is_some_and(|extension| extension.eq_ignore_ascii_case("json"))).collect::<Vec<_>>();
    if defaults.len() > 1 { return Err("Office legacy source has duplicate JSON content-type defaults.".into()); }
    let mut overrides = BTreeMap::new();
    for entry in types.iter().filter(|entry| entry.name == "Override") {
        let Some(part) = entry.attribute("PartName") else { continue; };
        let part = resolve_target("", part)?;
        if overrides.insert(part, entry.attribute("ContentType")).is_some() { return Err("Office legacy source has duplicate part content-type overrides.".into()); }
    }
    for (part, kind) in &overrides {
        if *kind == Some(LEGACY_SOURCE_CONTENT_TYPE) && part != LEGACY_SOURCE_PART { return Err("Office legacy source content type names an unregistered part.".into()); }
    }
    let content_type = overrides.get(LEGACY_SOURCE_PART).copied().unwrap_or_else(|| defaults.first().and_then(|entry| entry.attribute("ContentType")));
    if content_type != Some(LEGACY_SOURCE_CONTENT_TYPE) { return Err("Office legacy source part has a missing or inconsistent content type.".into()); }
    for name in &names {
        if name != LEGACY_SOURCE_PART && name.rsplit('.').next().is_some_and(|extension| extension.eq_ignore_ascii_case("json")) {
            let kind = overrides.get(name).copied().unwrap_or_else(|| defaults.first().and_then(|entry| entry.attribute("ContentType")));
            if kind == Some(LEGACY_SOURCE_CONTENT_TYPE) { return Err("Office legacy source content type resolves to multiple parts.".into()); }
        }
    }

    let entry = archive.by_name(LEGACY_SOURCE_PART).map_err(|error| error.to_string())?;
    if entry.size() > JSON_BYTE_LIMIT as u64 { return Err("Office legacy source JSON exceeds the 64 MiB part limit.".into()); }
    let mut bytes = Vec::new();
    entry.take(JSON_BYTE_LIMIT as u64 + 1).read_to_end(&mut bytes).map_err(|error| error.to_string())?;
    if bytes.len() > JSON_BYTE_LIMIT { return Err("Office legacy source JSON exceeds the 64 MiB part limit.".into()); }
    decode_origins(&bytes)
}

#[cfg(test)]
#[path = "document_legacy_source_sidecar_tests.rs"]
mod tests;
