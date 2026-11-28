//! Actual data stores and OPC graph preservation, independent from editor sidecars.
use std::collections::{BTreeMap, BTreeSet};
use office_io_common::document_custom_xml::{self as contract, CustomXmlRelationship, DocumentCustomXmlPart};
use quick_xml::{Reader, XmlVersion, events::Event};
use serde::Deserialize;
use serde_json::Value;
use super::{WriteContext, get, xml, ExportError};

struct Part { name: String, content_type: String, bytes: Vec<u8> }
struct Owner { name: String, xml: String, additional: Vec<CustomXmlRelationship> }
#[derive(Default)]
pub(super) struct Prepared {
    pub(super) stores: BTreeSet<String>,
    parts: BTreeMap<String, Part>,
    owners: BTreeMap<String, Owner>,
    sources: Vec<String>,
}

fn insert_part(parts: &mut BTreeMap<String, Part>, name: &str, kind: &str, bytes: Vec<u8>) -> Result<(), ExportError> {
    let name = contract::canonical_part_name(name).map_err(ExportError::invalid)?;
    let key = name.to_ascii_lowercase();
    if key.ends_with(".rels") || key == "%5bcontent_types%5d.xml" {
        return Err(ExportError::invalid("DOCX_CUSTOM_XML_INVALID_PART: invalid retained ordinary part"));
    }
    if let Some(previous) = parts.get(&key) {
        if !previous.content_type.eq_ignore_ascii_case(kind) || previous.bytes != bytes { return Err(ExportError::invalid(format!("DOCX_CUSTOM_XML_RESOURCE_CONFLICT: {name}"))); }
        return Ok(());
    }
    parts.insert(key, Part { name, content_type: kind.to_owned(), bytes });
    Ok(())
}

fn insert_owner(owners: &mut BTreeMap<String, Owner>, name: &str, xml: String) -> Result<(), ExportError> {
    let name = contract::canonical_part_name(name).map_err(ExportError::invalid)?;
    let key = name.to_ascii_lowercase();
    contract::relationships(&xml).map_err(ExportError::invalid)?;
    if let Some(previous) = owners.get(&key) {
        if previous.xml != xml { return Err(ExportError::invalid(format!("DOCX_CUSTOM_XML_RELATIONSHIP_CONFLICT: {name}"))); }
        return Ok(());
    }
    owners.insert(key, Owner { name, xml, additional: Vec::new() });
    Ok(())
}

fn empty_relationships() -> String {
    format!("{}<Relationships xmlns=\"{}\"/>", xml::DECL, contract::RELATIONSHIP_NAMESPACE)
}

pub(super) fn prepare(model: &Value) -> Result<Prepared, ExportError> {
    for control in get(model, "contentControls").as_array().into_iter().flatten() {
        let Some(binding) = control.get("dataBinding") else { continue; };
        if !binding.is_object()
            || binding.get("xpath").is_some_and(|value| value.as_str().is_none_or(|value| value.encode_utf16().count() > 8192))
            || binding.get("storeItemID").is_some_and(|value| value.as_str().is_none_or(|value| value.encode_utf16().count() > 128 || value.chars().any(char::is_control)))
            || binding.get("prefixMappings").is_some_and(|value| value.as_str().is_none_or(|value| value.encode_utf16().count() > 8192))
            || binding.get("cachedResult").is_some_and(|value| !value.is_string())
            || binding.get("dirty").is_some_and(|value| !value.is_boolean()) {
            return Err(ExportError::invalid("DOCX_CUSTOM_XML_INVALID_BINDING: invalid content-control XML binding"));
        }
        if active_binding_store(binding).is_some() {
            if let Some(mappings) = get(binding, "prefixMappings").as_str() { contract::validate_prefix_mappings(mappings).map_err(ExportError::invalid)?; }
        }
    }
    let Some(value) = model.get("customXmlParts") else { return Ok(Prepared::default()); };
    if value.is_null() { return Err(ExportError::invalid("DOCX_CUSTOM_XML_INVALID_SOURCE: use an empty array to clear data sources")); }
    // Deserializing &Value creates typed owners without cloning a second JSON tree.
    let mut sources = Vec::<DocumentCustomXmlPart>::deserialize(value).map_err(|error| ExportError::invalid(format!("DOCX_CUSTOM_XML_INVALID_SOURCE: {error}")))?;
    for part in &mut sources {
        part.part_name = contract::canonical_part_name(&part.part_name).map_err(ExportError::invalid)?;
        if let Some(name) = &mut part.properties_part_name { *name = contract::canonical_part_name(name).map_err(ExportError::invalid)?; }
        for resource in &mut part.related_parts { resource.part_name = contract::canonical_part_name(&resource.part_name).map_err(ExportError::invalid)?; }
    }
    contract::validate_parts(&sources).map_err(ExportError::invalid)?;
    let mut prepared = Prepared::default();
    for part in &sources {
        insert_part(&mut prepared.parts, &part.part_name, part.content_type.as_deref().unwrap_or("application/xml"), contract::encode_xml(part).map_err(ExportError::invalid)?)?;
        prepared.sources.push(part.part_name.clone());
        if let (Some(name), Some(xml)) = (&part.properties_part_name, &part.properties_xml) {
            let xml = contract::normalize_xml_encoding(xml.clone());
            insert_part(&mut prepared.parts, name, contract::CUSTOM_XML_PROPERTIES_CONTENT_TYPE, xml.into_bytes())?;
        }
        if let Some(xml) = &part.relationships_xml { insert_owner(&mut prepared.owners, &part.part_name, contract::normalize_xml_encoding(xml.clone()))?; }
        for resource in &part.related_parts {
            let bytes = contract::decode_resource(&resource.data_base64).map_err(ExportError::invalid)?;
            if resource.part_name.to_ascii_lowercase().ends_with(".rels") {
                if !resource.content_type.eq_ignore_ascii_case(contract::RELATIONSHIP_CONTENT_TYPE) { return Err(ExportError::invalid("DOCX_CUSTOM_XML_INVALID_RELATIONSHIPS: incorrect relationship content type")); }
                let owner = contract::relationship_owner(&resource.part_name).ok_or_else(|| ExportError::invalid("DOCX_CUSTOM_XML_INVALID_RELATIONSHIPS: invalid related owner"))?;
                let (xml, _, _) = contract::decode_xml(&bytes).map_err(ExportError::invalid)?;
                insert_owner(&mut prepared.owners, &owner, contract::normalize_xml_encoding(xml))?;
            } else { insert_part(&mut prepared.parts, &resource.part_name, &resource.content_type, bytes)?; }
        }
    }
    for (index, part) in sources.iter().enumerate() {
        let source_key = part.part_name.to_ascii_lowercase();
        let edges = prepared.owners.get(&source_key).map(|owner| contract::relationships(&owner.xml)).transpose().map_err(ExportError::invalid)?.unwrap_or_default();
        let property_edges = edges.iter().filter(|edge| contract::relationship_is(&edge.kind, "customXmlProps")).collect::<Vec<_>>();
        if property_edges.len() > 1 || property_edges.first().is_some_and(|edge| edge.external) { return Err(ExportError::invalid("DOCX_CUSTOM_XML_INVALID_PROPERTIES: ambiguous or external properties target")); }
        let existing = property_edges.first().map(|edge| contract::resolve_target(&part.part_name, &edge.target)).transpose().map_err(ExportError::invalid)?;
        if part.properties_part_name.as_ref().zip(existing.as_ref()).is_some_and(|(declared, related)| !declared.eq_ignore_ascii_case(related)) {
            return Err(ExportError::invalid("DOCX_CUSTOM_XML_INVALID_PROPERTIES: property relationship differs from its declared part"));
        }
        let mut property = part.properties_part_name.clone().or(existing.clone());
        if property.is_none() {
            if let Some(id) = &part.store_item_id {
                let mut serial = index + 1;
                let name = loop {
                    let name = format!("customXml/office-data-properties/item{serial}.xml");
                    if !prepared.parts.contains_key(&name.to_ascii_lowercase()) { break name; }
                    serial += 1;
                };
                let xml = format!("{}<ds:datastoreItem xmlns:ds=\"{}\" ds:itemID=\"{}\"><ds:schemaRefs/></ds:datastoreItem>", xml::DECL, contract::CUSTOM_XML_NAMESPACE, xml::escape(id));
                insert_part(&mut prepared.parts, &name, contract::CUSTOM_XML_PROPERTIES_CONTENT_TYPE, xml.into_bytes())?;
                property = Some(name);
            }
        }
        if let Some(name) = &property {
            let property_part = prepared.parts.get(&name.to_ascii_lowercase()).ok_or_else(|| ExportError::invalid(format!("DOCX_CUSTOM_XML_MISSING_PART: properties {name}")))?;
            if !property_part.content_type.eq_ignore_ascii_case(contract::CUSTOM_XML_PROPERTIES_CONTENT_TYPE) { return Err(ExportError::invalid("DOCX_CUSTOM_XML_INVALID_PROPERTIES: invalid property content type")); }
            let (xml, _, _) = contract::decode_xml(&property_part.bytes).map_err(ExportError::invalid)?;
            let id = contract::properties_store_id(&xml).map_err(ExportError::invalid)?;
            if part.store_item_id.as_ref().is_some_and(|expected| !expected.eq_ignore_ascii_case(&id)) { return Err(ExportError::invalid("DOCX_CUSTOM_XML_INVALID_PROPERTIES: properties store ID mismatch")); }
            if !prepared.stores.insert(id.to_lowercase()) { return Err(ExportError::invalid("DOCX_CUSTOM_XML_INVALID_PROPERTIES: duplicate store ID")); }
            if existing.is_none() {
                let owner = prepared.owners.entry(source_key).or_insert_with(|| Owner { name: part.part_name.clone(), xml: empty_relationships(), additional: Vec::new() });
                let ids = edges.iter().map(|edge| edge.id.as_str()).collect::<BTreeSet<_>>();
                let mut serial = 1;
                let id = loop { let id = format!("rId{serial}"); if !ids.contains(id.as_str()) { break id; } serial += 1; };
                owner.additional.push(CustomXmlRelationship { id, kind: contract::CUSTOM_XML_PROPERTIES_RELATIONSHIP.into(), target: relative_target(&part.part_name, name), external: false });
            }
        }
    }
    for (key, owner) in &prepared.owners {
        let name = &owner.name;
        if !prepared.parts.contains_key(key) { return Err(ExportError::invalid(format!("DOCX_CUSTOM_XML_MISSING_PART: relationship owner {name}"))); }
        for edge in contract::relationships(&owner.xml).map_err(ExportError::invalid)?.into_iter().chain(owner.additional.iter().cloned()) {
            if !edge.external {
                let target = contract::resolve_target(name, &edge.target).map_err(ExportError::invalid)?;
                if !prepared.parts.contains_key(&target.to_ascii_lowercase()) { return Err(ExportError::invalid(format!("DOCX_CUSTOM_XML_MISSING_PART: related {target}"))); }
            }
        }
    }
    for control in get(model, "contentControls").as_array().into_iter().flatten() {
        if let Some(id) = active_binding_store(get(control, "dataBinding")) {
            if !prepared.stores.contains(&id.to_lowercase()) { return Err(ExportError::invalid("DOCX_CUSTOM_XML_MISSING_STORE: content control references a store absent from the explicit source list")); }
        }
    }
    Ok(prepared)
}

// Univer retains partially configured bindings and activates only a nonempty store + XPath pair.
pub(super) fn active_binding_store(binding: &Value) -> Option<&str> {
    let store = get(binding, "storeItemID").as_str().filter(|value| !value.is_empty())?;
    get(binding, "xpath").as_str().filter(|value| !value.is_empty())?;
    Some(store)
}

fn relative_target(owner: &str, target: &str) -> String {
    let parent = owner.rsplit_once('/').map_or("", |(parent, _)| parent);
    let from = parent.split('/').filter(|segment| !segment.is_empty()).collect::<Vec<_>>();
    let to = target.split('/').collect::<Vec<_>>();
    let common = from.iter().zip(&to).take_while(|(left, right)| left == right).count();
    std::iter::repeat_n("..", from.len() - common).chain(to[common..].iter().copied()).collect::<Vec<_>>().join("/")
}

fn reserved_generated_path(name: &str) -> bool {
    let name = name.to_ascii_lowercase();
    name.starts_with("word/") || name.starts_with("docprops/") || name.starts_with("xl/") || name.starts_with("ppt/")
        || name.starts_with("_rels/") || matches!(name.as_str(), "customxml/office-numbering.xml"
            | "customxml/office-numbering-properties.xml" | "customxml/office-fields.xml" | "customxml/office-tables.xml" | "customxml/office-content-controls.xml")
}

fn relationship_xml(edge: &CustomXmlRelationship, qualified_name: &str) -> String {
    format!("<{qualified_name}{}{}{}{} />", xml::attr("Id", &edge.id), xml::attr("Type", &edge.kind), xml::attr("Target", &edge.target),
        if edge.external { xml::attr("TargetMode", "External") } else { String::new() })
}

/// Replace only OPC targets and append generated property relationships. Other
/// XML tokens, namespace declarations, comments and processing instructions stay.
fn remap_relationships(owner: &str, source: &Owner, mapping: &BTreeMap<String, String>) -> Result<String, ExportError> {
    let mut reader = Reader::from_str(&source.xml);
    let mut output = String::new();
    output.try_reserve(source.xml.len()).map_err(|_| ExportError::invalid("DOCX_CUSTOM_XML_RESOURCE_LIMIT: relationship allocation failed"))?;
    let mut cursor = 0usize;
    let mut depth = 0usize;
    let mut root_name = String::new();
    let mut added = false;
    loop {
        let start = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(|error| ExportError::invalid(error.to_string()))?;
        let end = reader.buffer_position() as usize;
        let empty = matches!(&event, Event::Empty(_));
        match event {
            Event::Start(tag) | Event::Empty(tag) => {
                let name = std::str::from_utf8(tag.name().as_ref().as_bytes()).map_err(|error| ExportError::invalid(error.to_string()))?.to_owned();
                if depth == 0 { root_name = name.clone(); }
                let root_empty = depth == 0 && empty && !source.additional.is_empty();
                if depth == 1 || root_empty {
                    let mut attributes = Vec::new();
                    for attribute in tag.attributes() {
                        let attribute = attribute.map_err(|error| ExportError::invalid(error.to_string()))?;
                        let name = std::str::from_utf8(attribute.key.as_ref().as_bytes()).map_err(|error| ExportError::invalid(error.to_string()))?.to_owned();
                        let value = attribute.normalized_value_with(XmlVersion::Implicit1_0, 1, quick_xml::escape::resolve_xml_entity)
                            .map_err(|error| ExportError::invalid(error.to_string()))?.into_owned();
                        attributes.push((name, value));
                    }
                    let external = attributes.iter().any(|(name, value)| name == "TargetMode" && value == "External");
                    let target = attributes.iter_mut().find(|(name, _)| name == "Target");
                    let mut changed = root_empty;
                    if !external { if let Some((_, value)) = target {
                        let old = contract::resolve_target(owner, value).map_err(ExportError::invalid)?;
                        let new = relative_target(&mapping[&owner.to_ascii_lowercase()], &mapping[&old.to_ascii_lowercase()]);
                        changed |= new != *value; *value = new;
                    } }
                    if changed {
                        output.push_str(&source.xml[cursor..start]); output.push('<'); output.push_str(&name);
                        for (name, value) in attributes { output.push_str(&xml::attr(&name, &value)); }
                        output.push_str(if empty && !root_empty { "/>" } else { ">" });
                        if root_empty {
                            let qualified = root_name.rsplit_once(':').map_or("Relationship".to_owned(), |(prefix, _)| format!("{prefix}:Relationship"));
                            for edge in &source.additional {
                                let mut edge = edge.clone(); let old = contract::resolve_target(owner, &edge.target).map_err(ExportError::invalid)?;
                                edge.target = relative_target(&mapping[&owner.to_ascii_lowercase()], &mapping[&old.to_ascii_lowercase()]); output.push_str(&relationship_xml(&edge, &qualified));
                            }
                            output.push_str(&format!("</{root_name}>")); added = true;
                        }
                        cursor = end;
                    }
                }
                if !empty { depth += 1; }
            },
            Event::End(_) => {
                if depth == 1 && !added && !source.additional.is_empty() {
                    output.push_str(&source.xml[cursor..start]);
                    let qualified = root_name.rsplit_once(':').map_or("Relationship".to_owned(), |(prefix, _)| format!("{prefix}:Relationship"));
                    for edge in &source.additional {
                        let mut edge = edge.clone(); let old = contract::resolve_target(owner, &edge.target).map_err(ExportError::invalid)?;
                        edge.target = relative_target(&mapping[&owner.to_ascii_lowercase()], &mapping[&old.to_ascii_lowercase()]); output.push_str(&relationship_xml(&edge, &qualified));
                    }
                    cursor = start; added = true;
                }
                depth -= 1;
            },
            Event::Eof => break,
            _ => {},
        }
        if output.len() > contract::MAX_PART_BYTES { return Err(ExportError::invalid("DOCX_CUSTOM_XML_RESOURCE_LIMIT: relationship XML exceeds 8 MiB")); }
    }
    output.push_str(&source.xml[cursor..]);
    contract::relationships(&output).map_err(ExportError::invalid)?;
    Ok(output)
}

impl Prepared {
    pub(super) fn write(self, context: &mut WriteContext<'_>) -> Result<(), ExportError> {
        let mut mapping = BTreeMap::new();
        let mut used = self.parts.values().filter(|part| !reserved_generated_path(&part.name)).map(|part| part.name.to_ascii_lowercase()).collect::<BTreeSet<_>>();
        let mut relocated = 0usize;
        let mut serial = 1usize;
        for (key, part) in &self.parts {
            let name = &part.name;
            let path = if !reserved_generated_path(name) && !context.package.contains_part(name) { name.clone() } else {
                relocated += 1;
                let extension = name.rsplit_once('.').map_or("bin", |(_, extension)| extension);
                loop {
                    let candidate = format!("customXml/office-studio-data/item{serial}.{extension}"); serial += 1;
                    if !used.contains(&candidate.to_ascii_lowercase()) && !context.package.contains_part(&candidate) { break candidate; }
                }
            };
            used.insert(path.to_ascii_lowercase()); mapping.insert(key.clone(), path);
        }
        let owners = self.owners.iter().map(|(key, owner)| remap_relationships(&owner.name, owner, &mapping).map(|xml| (mapping[key].clone(), xml.into_bytes()))).collect::<Result<Vec<_>, _>>()?;
        let mut output_bytes = 0usize;
        for length in self.parts.values().map(|part| part.bytes.len()).chain(owners.iter().map(|(_, bytes)| bytes.len())) {
            output_bytes = output_bytes.checked_add(length).filter(|total| *total <= contract::MAX_TOTAL_BYTES)
                .ok_or_else(|| ExportError::invalid("DOCX_CUSTOM_XML_RESOURCE_LIMIT: encoded dependency graph exceeds 64 MiB"))?;
        }
        for (key, part) in self.parts { context.package.add_binary(&mapping[&key], &part.content_type, part.bytes)?; }
        for (name, bytes) in owners { context.package.add_preserved_relationships(&name, bytes)?; }
        for name in self.sources { context.relationship(contract::CUSTOM_XML_RELATIONSHIP, &relative_target("word/document.xml", &mapping[&name.to_ascii_lowercase()]), false)?; }
        if relocated != 0 { context.warn("DOCX_CUSTOM_XML_PART_RELOCATED", format!("{relocated} custom-XML dependency part(s) were relocated with their OPC targets to avoid generated document-part collisions.")); }
        Ok(())
    }
}

#[cfg(test)]
#[path = "docx_custom_xml_tests.rs"]
mod tests;
