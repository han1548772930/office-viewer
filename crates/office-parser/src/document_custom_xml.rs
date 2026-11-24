//! Retain actual DOCX custom-XML data stores through their OPC dependency graph.
use std::collections::{BTreeMap, BTreeSet, VecDeque};
use std::io::{Cursor, Read};
use office_io_common::{chart_xml, document_custom_xml::{self as contract, CustomXmlRelationship, DocumentCustomXmlPart, DocumentCustomXmlResource}};
use zip::ZipArchive;

const CONTENT_TYPES_NAMESPACE: &str = "http://schemas.openxmlformats.org/package/2006/content-types";
const SOURCE_CLASSIFICATION_BYTES: usize = 16 * 1024 * 1024;

struct PackageNames { entries: BTreeMap<String, Option<usize>> }
impl PackageNames {
    fn read(archive: &mut ZipArchive<Cursor<&[u8]>>) -> Result<Self, String> {
        let mut entries = BTreeMap::new();
        for index in 0..archive.len() {
            let entry = archive.by_index(index).map_err(|error| error.to_string())?;
            if entry.is_dir() { continue; }
            // ZIP producers use both literal names and URI-escaped names. A
            // reference is ambiguous if either spelling aliases two entries.
            let keys = [contract::part_name_from_zip(entry.name()), contract::canonical_part_name(entry.name())]
                .into_iter().filter_map(Result::ok).map(|name| name.to_ascii_lowercase()).collect::<BTreeSet<_>>();
            for key in keys {
                let previous = entries.entry(key).or_insert(Some(index));
                if *previous != Some(index) { *previous = None; }
            }
        }
        Ok(Self { entries })
    }
}

fn bytes(archive: &mut ZipArchive<Cursor<&[u8]>>, names: &PackageNames, name: &str, maximum: usize) -> Result<Option<Vec<u8>>, String> {
    let index = match names.entries.get(&contract::part_name_key(name)?) {
        Some(Some(index)) => *index, Some(None) => return Err(format!("DOCX_CUSTOM_XML_DUPLICATE_PART: ambiguous package URI {name}")), None => return Ok(None),
    };
    let entry = archive.by_index(index).map_err(|error| error.to_string())?;
    if entry.size() > maximum as u64 { return Err(format!("DOCX_CUSTOM_XML_RESOURCE_LIMIT: {name} exceeds its source budget")); }
    let mut result = Vec::new();
    result.try_reserve_exact(entry.size() as usize).map_err(|_| "DOCX_CUSTOM_XML_RESOURCE_LIMIT: package part allocation failed")?;
    entry.take(maximum as u64 + 1).read_to_end(&mut result).map_err(|error| error.to_string())?;
    if result.len() > maximum { return Err(format!("DOCX_CUSTOM_XML_RESOURCE_LIMIT: {name} exceeds its source budget")); }
    Ok(Some(result))
}

fn required_bytes(archive: &mut ZipArchive<Cursor<&[u8]>>, names: &PackageNames, name: &str) -> Result<Vec<u8>, String> {
    bytes(archive, names, name, contract::MAX_PART_BYTES)?.ok_or_else(|| format!("DOCX_CUSTOM_XML_MISSING_PART: declared package part {name} is missing"))
}

fn related(archive: &mut ZipArchive<Cursor<&[u8]>>, names: &PackageNames, owner: &str) -> Result<Option<(String, Vec<CustomXmlRelationship>)>, String> {
    let Some(bytes) = bytes(archive, names, &contract::relationship_part(owner), contract::MAX_PART_BYTES)? else { return Ok(None); };
    let (xml, _, _) = contract::decode_xml(&bytes)?;
    let xml = contract::normalize_xml_encoding(xml);
    let entries = contract::relationships(&xml)?;
    Ok(Some((xml, entries)))
}

fn main_part(archive: &mut ZipArchive<Cursor<&[u8]>>, names: &PackageNames) -> Result<String, String> {
    let Some(bytes) = bytes(archive, names, "_rels/.rels", contract::MAX_PART_BYTES)? else {
        return crate::document_format_sidecar::part_for_type(archive, ".document.main+xml", "word/document.xml");
    };
    let (xml, _, _) = contract::decode_xml(&bytes)?;
    let entries = contract::relationships(&xml)?;
    let mains = entries.iter().filter(|entry| contract::relationship_is(&entry.kind, "officeDocument")).collect::<Vec<_>>();
    if mains.len() != 1 || mains[0].external { return Err("DOCX_CUSTOM_XML_INVALID_PACKAGE: package must identify one internal document part".into()); }
    contract::resolve_target("", &mains[0].target)
}

#[derive(Default)]
struct ContentTypes { defaults: BTreeMap<String, String>, overrides: BTreeMap<String, String> }
impl ContentTypes {
    fn read(archive: &mut ZipArchive<Cursor<&[u8]>>, names: &PackageNames) -> Result<Self, String> {
        let (xml, _, _) = contract::decode_xml(&required_bytes(archive, names, "[Content_Types].xml")?)?;
        let root = chart_xml::parse(&xml)?;
        if root.name != "Types" || root.namespace.as_deref() != Some(CONTENT_TYPES_NAMESPACE) {
            return Err("DOCX_CUSTOM_XML_INVALID_PACKAGE: invalid content-types root".into());
        }
        if !root.text.trim().is_empty() { return Err("DOCX_CUSTOM_XML_INVALID_PACKAGE: content-types root contains text".into()); }
        let mut result = Self::default();
        for node in &root.children {
            if node.namespace != root.namespace || !node.children.is_empty() || !node.text.trim().is_empty() { return Err("DOCX_CUSTOM_XML_INVALID_PACKAGE: invalid content-types child".into()); }
            let attr = |name: &str| node.expanded_attrs.get("").and_then(|values| values.get(name)).filter(|value| !value.is_empty()).cloned()
                .ok_or_else(|| format!("DOCX_CUSTOM_XML_INVALID_PACKAGE: missing content-types {name}"));
            let kind = attr("ContentType")?;
            if !contract::valid_content_type(&kind) { return Err("DOCX_CUSTOM_XML_INVALID_PACKAGE: invalid MIME content type".into()); }
            let previous = match node.name.as_str() {
                "Default" => result.defaults.insert(attr("Extension")?.to_ascii_lowercase(), kind),
                "Override" => {
                    let name = attr("PartName")?;
                    if !name.starts_with('/') { return Err("DOCX_CUSTOM_XML_INVALID_PACKAGE: content-type override must use an absolute part URI".into()); }
                    result.overrides.insert(contract::part_name_key(&contract::resolve_target("", &name)?)?, kind)
                },
                _ => return Err("DOCX_CUSTOM_XML_INVALID_PACKAGE: unknown content-types entry".into()),
            };
            if previous.is_some() { return Err("DOCX_CUSTOM_XML_INVALID_PACKAGE: duplicate content-types declaration".into()); }
        }
        Ok(result)
    }
    fn for_part(&self, name: &str) -> Result<String, String> {
        let kind = self.overrides.get(&contract::part_name_key(name)?).or_else(|| name.rsplit_once('.').and_then(|(_, extension)| self.defaults.get(&extension.to_ascii_lowercase())))
            .cloned().ok_or_else(|| format!("DOCX_CUSTOM_XML_MISSING_CONTENT_TYPE: {name} has no declared content type"))?;
        if !contract::valid_content_type(&kind) || name.to_ascii_lowercase().ends_with(".rels") && !kind.eq_ignore_ascii_case(contract::RELATIONSHIP_CONTENT_TYPE) {
            return Err("DOCX_CUSTOM_XML_INVALID_PACKAGE: invalid part content type".into());
        }
        Ok(kind)
    }
}

fn editor_sidecar(xml: &str) -> Result<bool, String> {
    let (name, namespace) = contract::xml_root_identity(xml)?;
    Ok(matches!((name.as_str(), namespace.as_deref()),
        ("officeNumbering", Some("urn:office-studio:numbering:1"))
        | ("fields", Some("urn:office-studio:document-fields:1"))
        | ("tables", Some("urn:office-studio:document-tables:1"))
        | ("controls", Some("urn:office-studio:content-controls:1"))))
}

fn charge(total: &mut usize, amount: usize) -> Result<(), String> {
    *total = total.checked_add(amount).filter(|value| *value <= contract::MAX_TOTAL_BYTES)
        .ok_or("DOCX_CUSTOM_XML_RESOURCE_LIMIT: data stores exceed 64 MiB")?;
    Ok(())
}

/// Related resources share a document-level name space. Store/property parts
/// are their own owners; a shared resource is retained once at its first root.
pub(crate) fn read(archive: &mut ZipArchive<Cursor<&[u8]>>) -> Result<Option<Vec<DocumentCustomXmlPart>>, String> {
    let names = PackageNames::read(archive)?;
    let main = main_part(archive, &names)?;
    let Some((_, entries)) = related(archive, &names, &main)? else { return Ok(None); };
    let mut roots = BTreeMap::new();
    for entry in entries.iter().filter(|entry| contract::relationship_is(&entry.kind, "customXml")) {
        if entry.external { return Err("DOCX_CUSTOM_XML_EXTERNAL_STORE: external data stores cannot be fetched during local import".into()); }
        let name = contract::resolve_target(&main, &entry.target)?;
        roots.entry(contract::part_name_key(&name)?).or_insert(name);
    }
    if roots.is_empty() { return Ok(None); }
    let types = ContentTypes::read(archive, &names)?;
    types.for_part(&contract::relationship_part(&main))?;
    if names.entries.contains_key(&contract::part_name_key("_rels/.rels")?) { types.for_part("_rels/.rels")?; }
    let mut parts = Vec::new();
    let mut total = 0usize;
    let mut primary = BTreeSet::new();
    let mut root_edges = Vec::new();
    for (name_key, name) in roots {
        let raw = bytes(archive, &names, &name, SOURCE_CLASSIFICATION_BYTES)?.ok_or_else(|| format!("DOCX_CUSTOM_XML_MISSING_PART: {name}"))?;
        let raw_length = raw.len();
        let (xml, encoding, bom) = contract::decode_xml_text(&raw, SOURCE_CLASSIFICATION_BYTES)?;
        if editor_sidecar(&xml)? { continue; }
        if parts.len() >= contract::MAX_PARTS { return Err("DOCX_CUSTOM_XML_RESOURCE_LIMIT: too many custom-XML data sources".into()); }
        if raw_length > contract::MAX_PART_BYTES { return Err("DOCX_CUSTOM_XML_RESOURCE_LIMIT: data store exceeds 8 MiB".into()); }
        contract::validate_xml(&xml)?;
        charge(&mut total, xml.len().max(raw_length))?;
        let content_type = types.for_part(&name)?;
        if !content_type.eq_ignore_ascii_case("application/xml") && !content_type.eq_ignore_ascii_case("text/xml") {
            // OOXML allows XML subtypes; retain them as graph resources through
            // the normal XML data-store part rather than silently treating binary as XML.
            if !content_type.to_ascii_lowercase().ends_with("+xml") { return Err("DOCX_CUSTOM_XML_INVALID_SOURCE: data-store part is not XML".into()); }
        }
        let source_relationships = related(archive, &names, &name)?;
        if source_relationships.is_some() { types.for_part(&contract::relationship_part(&name))?; }
        let links = source_relationships.as_ref().map(|(_, entries)| entries.as_slice()).unwrap_or_default();
        let props = links.iter().filter(|entry| contract::relationship_is(&entry.kind, "customXmlProps")).collect::<Vec<_>>();
        if props.len() > 1 || props.first().is_some_and(|entry| entry.external) { return Err("DOCX_CUSTOM_XML_INVALID_PROPERTIES: ambiguous or external properties part".into()); }
        let mut part = DocumentCustomXmlPart { part_name: name.clone(), xml, content_type: Some(content_type), store_item_id: None, properties_part_name: None,
            properties_xml: None, relationships_xml: source_relationships.as_ref().map(|(xml, _)| xml.clone()), related_parts: Vec::new(),
            encoding: Some(encoding), byte_order_mark: Some(bom) };
        primary.insert(name_key);
        if let Some(properties) = props.first() {
            let property_name = contract::resolve_target(&name, &properties.target)?;
            if !types.for_part(&property_name)?.eq_ignore_ascii_case(contract::CUSTOM_XML_PROPERTIES_CONTENT_TYPE) { return Err("DOCX_CUSTOM_XML_INVALID_PROPERTIES: wrong properties content type".into()); }
            let (properties_xml, _, _) = contract::decode_xml(&required_bytes(archive, &names, &property_name)?)?;
            let properties_xml = contract::normalize_xml_encoding(properties_xml);
            let id = contract::properties_store_id(&properties_xml)?;
            charge(&mut total, properties_xml.len())?;
            part.store_item_id = Some(id);
            part.properties_part_name = Some(property_name.clone());
            part.properties_xml = Some(properties_xml);
            primary.insert(contract::part_name_key(&property_name)?);
        }
        if let Some(xml) = &part.relationships_xml { charge(&mut total, xml.len())?; }
        root_edges.push(source_relationships.map(|(_, entries)| entries).unwrap_or_default());
        parts.push(part);
    }
    if parts.is_empty() { return Ok(None); }
    let source_names = parts.iter().map(|part| contract::part_name_key(&part.part_name)).collect::<Result<BTreeSet<_>, _>>()?;
    let mut retained = BTreeSet::new();
    for (part, edges) in parts.iter_mut().zip(root_edges) {
        let mut queue = VecDeque::new();
        if let Some(name) = &part.properties_part_name { queue.push_back(name.clone()); }
        for edge in &edges { if !edge.external { queue.push_back(contract::resolve_target(&part.part_name, &edge.target)?); } }
        let mut visited = BTreeSet::new();
        while let Some(name) = queue.pop_front() {
            let name_key = contract::part_name_key(&name)?;
            if name_key.ends_with(".rels") || name_key == "%5bcontent_types%5d.xml" { return Err("DOCX_CUSTOM_XML_INVALID_TARGET: relationships cannot target OPC system parts".into()); }
            if !visited.insert(name_key.clone()) { continue; }
            if visited.len() > contract::MAX_RELATED_PARTS + contract::MAX_PARTS * 2 { return Err("DOCX_CUSTOM_XML_RESOURCE_LIMIT: custom-XML graph exceeds its node limit".into()); }
            if !primary.contains(&name_key) && retained.insert(name_key.clone()) {
                let bytes = required_bytes(archive, &names, &name)?;
                charge(&mut total, bytes.len().div_ceil(3) * 4)?;
                part.related_parts.push(DocumentCustomXmlResource { content_type: types.for_part(&name)?, part_name: name.clone(), data_base64: contract::encode_resource(&bytes) });
            }
            let Some((xml, edges)) = related(archive, &names, &name)? else { continue; };
            types.for_part(&contract::relationship_part(&name))?;
            if !source_names.contains(&name_key) {
                let relationships_name = contract::relationship_part(&name);
                if retained.insert(contract::part_name_key(&relationships_name)?) {
                    charge(&mut total, xml.len().div_ceil(3) * 4)?;
                    part.related_parts.push(DocumentCustomXmlResource { part_name: relationships_name, content_type: contract::RELATIONSHIP_CONTENT_TYPE.into(), data_base64: contract::encode_resource(xml.as_bytes()) });
                }
            }
            for edge in edges { if !edge.external { queue.push_back(contract::resolve_target(&name, &edge.target)?); } }
            if retained.len() > contract::MAX_RELATED_PARTS { return Err("DOCX_CUSTOM_XML_RESOURCE_LIMIT: too many retained graph resources".into()); }
        }
    }
    contract::validate_parts(&parts)?;
    Ok(Some(parts))
}
