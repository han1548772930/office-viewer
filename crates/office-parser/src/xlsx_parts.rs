//! Resolve worksheet parts through OPC relationships, including relocated parts.

use std::collections::{BTreeMap, BTreeSet};
use std::io::{Read, Seek};

use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::{NsReader, XmlVersion};
use zip::ZipArchive;

const MAX_METADATA_XML_BYTES: u64 = 16 * 1024 * 1024;
const PACKAGE_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const STRICT_PACKAGE_NS: &str = "http://purl.oclc.org/ooxml/package/relationships";
const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_REL_NS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";

pub(crate) struct WorksheetPart {
    pub(crate) name: String,
    pub(crate) path: String,
    pub(crate) office_sheet_id: u32,
}
pub(crate) struct WorkbookParts {
    pub(crate) workbook_path: String,
    pub(crate) sheets: Vec<WorksheetPart>,
    pub(crate) custom_xml: Vec<String>,
}
pub(crate) struct Relationship {
    pub(crate) id: String,
    pub(crate) target: String,
    pub(crate) kind: String,
    pub(crate) external: bool,
}

pub(crate) fn spreadsheet_namespace(namespace: ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(value) if matches!(value.as_ref(), "http://schemas.openxmlformats.org/spreadsheetml/2006/main" | "http://purl.oclc.org/ooxml/spreadsheetml/main"))
}

pub(crate) fn worksheet_parts<R: Read + Seek>(archive: &mut ZipArchive<R>) -> Result<Vec<WorksheetPart>, String> {
    Ok(workbook_parts(archive)?.sheets)
}

pub(crate) fn workbook_parts<R: Read + Seek>(archive: &mut ZipArchive<R>) -> Result<WorkbookParts, String> {
    let root = required_xml_part(archive, "_rels/.rels")?;
    let root_relationships = relationships(&root)?;
    let workbooks = root_relationships.iter().filter(|rel| rel.kind == "officeDocument" && !rel.external).collect::<Vec<_>>();
    if workbooks.len() != 1 { return Err("XLSX package must identify one internal workbook part".into()); }
    let workbook_path = resolve_target("", &workbooks[0].target).ok_or_else(|| "Invalid XLSX workbook relationship target".to_owned())?;
    let workbook = required_xml_part(archive, &workbook_path)?;
    let rels = required_xml_part(archive, &relationship_part_path(&workbook_path))?;
    let workbook_relationships = relationships(&rels)?;
    let mut custom_xml = BTreeSet::new();
    for (source, rels) in [("", &root_relationships), (workbook_path.as_str(), &workbook_relationships)] {
        for rel in rels.iter().filter(|rel| rel.kind == "customXml" && !rel.external) {
            custom_xml.insert(resolve_target(source, &rel.target).ok_or_else(|| "Invalid XLSX custom XML relationship target".to_owned())?);
            if custom_xml.len() > 4096 { return Err("XLSX custom XML relationships exceed the parser limit".into()); }
        }
    }
    let rels = workbook_relationships.into_iter().map(|rel| (rel.id.clone(), rel)).collect::<BTreeMap<_, _>>();
    let sheets = workbook_sheets(&workbook)?.into_iter().filter_map(|(name, id, office_sheet_id)| {
        let Some(rel) = rels.get(&id) else { return Some(Err(format!("Missing XLSX worksheet relationship for {name}"))); };
        if rel.kind != "worksheet" { return None; }
        if rel.external { return Some(Err(format!("XLSX worksheet {name} has an external relationship"))); }
        Some(resolve_target(&workbook_path, &rel.target).map(|path| WorksheetPart { name, path, office_sheet_id })
            .ok_or_else(|| "Invalid XLSX worksheet relationship target".to_owned()))
    }).collect::<Result<Vec<_>, _>>()?;
    Ok(WorkbookParts { workbook_path, sheets, custom_xml: custom_xml.into_iter().collect() })
}

pub(crate) fn required_xml_part<R: Read + Seek>(archive: &mut ZipArchive<R>, path: &str) -> Result<String, String> {
    let entry = archive.by_name(path).map_err(|error| format!("Could not read XLSX package part {path}: {error}"))?;
    if entry.size() > MAX_METADATA_XML_BYTES { return Err(format!("XLSX metadata part {path} exceeds the parser limit")); }
    let mut text = String::new();
    entry.take(MAX_METADATA_XML_BYTES + 1).read_to_string(&mut text).map_err(|error| format!("Could not decode XLSX package part {path}: {error}"))?;
    if text.len() as u64 > MAX_METADATA_XML_BYTES { return Err(format!("XLSX metadata part {path} exceeds the parser limit")); }
    Ok(text)
}

pub(crate) fn unqualified_attributes(_reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<BTreeMap<String, String>, String> {
    let mut attributes = BTreeMap::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| error.to_string())?;
        if attribute.key.as_ref().as_bytes().contains(&b':') || attribute.key.as_ref().as_bytes() == b"xmlns" { continue; }
        let name = std::str::from_utf8(attribute.key.as_ref().as_bytes()).map_err(|error| error.to_string())?.to_owned();
        let value = attribute.normalized_value(XmlVersion::Implicit1_0).map_err(|error| error.to_string())?.into_owned();
        attributes.insert(name, value);
    }
    Ok(attributes)
}

pub(crate) fn relationships(xml: &str) -> Result<Vec<Relationship>, String> {
    let mut reader = NsReader::from_str(xml);
    let mut depth: usize = 0;
    let mut root = false;
    let mut identifiers = BTreeSet::new();
    let mut result = Vec::new();
    loop {
        let event = reader.read_event().map_err(|error| error.to_string())?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                let namespace = reader.resolver().resolve_element(element.name()).0;
                let package_namespace = matches!(namespace, ResolveResult::Bound(value) if value.as_ref() == PACKAGE_NS || value.as_ref() == STRICT_PACKAGE_NS);
                if depth == 0 {
                    if root || element.local_name().as_ref().as_bytes() != b"Relationships" || !package_namespace { return Err("Invalid XLSX relationships root".into()); }
                    root = true;
                } else if depth == 1 && element.local_name().as_ref().as_bytes() == b"Relationship" && package_namespace {
                    let attributes = unqualified_attributes(&reader, element)?;
                    let id = attributes.get("Id").filter(|value| !value.is_empty()).ok_or_else(|| "Missing XLSX relationship ID".to_owned())?.clone();
                    if !identifiers.insert(id.clone()) { return Err("Duplicate XLSX relationship ID".into()); }
                    let target = attributes.get("Target").ok_or_else(|| "Missing XLSX relationship target".to_owned())?.clone();
                    let relation_type = attributes.get("Type").ok_or_else(|| "Missing XLSX relationship type".to_owned())?;
                    let kind = relation_type.strip_prefix("http://schemas.openxmlformats.org/officeDocument/2006/relationships/")
                        .or_else(|| relation_type.strip_prefix("http://purl.oclc.org/ooxml/officeDocument/relationships/"))
                        .unwrap_or("").to_owned();
                    let external = match attributes.get("TargetMode").map(String::as_str) {
                        None | Some("Internal") => false, Some("External") => true, _ => return Err("Invalid XLSX relationship target mode".into()),
                    };
                    result.push(Relationship { id, target, kind, external });
                }
                if matches!(event, Event::Start(_)) { depth += 1; }
                if depth > 64 { return Err("XLSX relationships nesting exceeds the parser limit".into()); }
            }
            Event::End(_) => depth = depth.checked_sub(1).ok_or_else(|| "Unmatched XLSX relationships end tag".to_owned())?,
            Event::DocType(_) => return Err("XLSX package DTDs are unsupported".into()),
            Event::Eof => break,
            _ => {}
        }
    }
    if !root || depth != 0 { return Err("Incomplete XLSX relationships document".into()); }
    Ok(result)
}

fn workbook_sheets(xml: &str) -> Result<Vec<(String, String, u32)>, String> {
    let mut reader = NsReader::from_str(xml);
    let mut path: Vec<Vec<u8>> = Vec::new();
    let mut root = false;
    let mut names = BTreeSet::new();
    let mut sheet_ids = BTreeSet::new();
    let mut result = Vec::new();
    loop {
        let event = reader.read_event().map_err(|error| error.to_string())?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                let name = element.local_name();
                let namespace = reader.resolver().resolve_element(element.name()).0;
                let namespace_is_sheet = spreadsheet_namespace(namespace);
                if path.is_empty() {
                    if root || name.as_ref() != "workbook" || !namespace_is_sheet { return Err("Invalid XLSX workbook root".into()); }
                    root = true;
                }
                if path.len() == 2 && path[0] == b"workbook" && path[1] == b"sheets" && name.as_ref() == "sheet" && namespace_is_sheet {
                    let attributes = unqualified_attributes(&reader, element)?;
                    let name = attributes.get("name").filter(|value| !value.is_empty()).ok_or_else(|| "Missing XLSX worksheet name".to_owned())?.clone();
                    if !names.insert(name.to_lowercase()) { return Err("Duplicate XLSX worksheet name".into()); }
                    let office_sheet_id = attributes.get("sheetId").and_then(|value| value.parse::<u32>().ok())
                        .filter(|id| *id > 0).ok_or_else(|| "Missing or invalid XLSX worksheet native ID".to_owned())?;
                    if !sheet_ids.insert(office_sheet_id) { return Err("Duplicate XLSX worksheet native ID".into()); }
                    let mut relationship_id = None;
                    for attribute in element.attributes() {
                        let attribute = attribute.map_err(|error| error.to_string())?;
                        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
                        if local.as_ref() == "id" && matches!(namespace, ResolveResult::Bound(value) if value.as_ref() == REL_NS || value.as_ref() == STRICT_REL_NS) {
                            if relationship_id.is_some() { return Err("Duplicate XLSX sheet relationship attribute".into()); }
                            relationship_id = Some(attribute.normalized_value(XmlVersion::Implicit1_0).map_err(|error| error.to_string())?.into_owned());
                        }
                    }
                    result.push((name, relationship_id.ok_or_else(|| "Missing XLSX worksheet relationship ID".to_owned())?, office_sheet_id));
                }
                if matches!(event, Event::Start(_)) { path.push(name.as_ref().as_bytes().to_vec()); }
                if path.len() > 64 { return Err("XLSX workbook nesting exceeds the parser limit".into()); }
            }
            Event::End(_) => { path.pop().ok_or_else(|| "Unmatched XLSX workbook end tag".to_owned())?; }
            Event::DocType(_) => return Err("XLSX workbook DTDs are unsupported".into()),
            Event::Eof => break,
            _ => {}
        }
    }
    if !root || !path.is_empty() { return Err("Incomplete XLSX workbook document".into()); }
    Ok(result)
}

pub(crate) fn relationship_part_path(source: &str) -> String {
    match source.rsplit_once('/') { Some((directory, file)) => format!("{directory}/_rels/{file}.rels"), None => format!("_rels/{source}.rels") }
}

pub(crate) fn resolve_target(source: &str, target: &str) -> Option<String> {
    let mut decoded = Vec::with_capacity(target.len());
    let mut input = target.bytes();
    while let Some(byte) = input.next() {
        decoded.push(if byte == b'%' {
            let high = char::from(input.next()?).to_digit(16)?;
            let low = char::from(input.next()?).to_digit(16)?;
            (high * 16 + low) as u8
        } else { byte });
    }
    let target = std::str::from_utf8(&decoded).ok()?;
    if target.contains([':', '\\', '?', '#']) || target.starts_with("//") || target.chars().any(char::is_control) { return None; }
    let mut parts = if target.starts_with('/') { Vec::new() } else { source.rsplit_once('/').map(|(directory, _)| directory.split('/').collect()).unwrap_or_default() };
    for part in target.trim_start_matches('/').split('/') {
        match part { "" | "." => {}, ".." => { parts.pop()?; }, value => parts.push(value) }
    }
    (!parts.is_empty()).then(|| parts.join("/"))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn resolves_escaped_relocated_targets_and_rejects_escape_from_the_package() {
        assert_eq!(resolve_target("custom/book.xml", "../sheets/sheet%20one.xml"), Some("sheets/sheet one.xml".into()));
        assert_eq!(resolve_target("custom/book.xml", "/sheets/one.xml"), Some("sheets/one.xml".into()));
        for invalid in ["../../outside.xml", "https://example.org/sheet.xml", "..\\outside.xml", "%00.xml", "//example.org/sheet.xml", "bad%GG.xml"] {
            assert!(resolve_target("custom/book.xml", invalid).is_none());
        }
    }

    #[test]
    fn workbook_relationship_ids_are_namespace_resolved_and_names_decode_entities() {
        let workbook = r#"<s:workbook xmlns:s="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><s:sheets><s:sheet name="A &amp; B" sheetId="37" rel:id="moved"/></s:sheets></s:workbook>"#;
        assert_eq!(workbook_sheets(workbook).unwrap(), [("A & B".into(), "moved".into(), 37)]);
        assert!(workbook_sheets(&workbook.replace("xmlns:rel=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"", "xmlns:rel=\"urn:other\"")).is_err());
    }
}
