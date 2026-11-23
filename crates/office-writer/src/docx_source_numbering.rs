//! Import numbering resources referenced by native source styles/properties.
//! Domain lists and sequence evaluation stay in docx_numbering.rs. This module
//! only allocates package-local IDs and preserves the native resource graph.

use std::collections::{BTreeMap, BTreeSet};
use serde::Deserialize;
use serde_json::Value;
use office_io_common::document_numbering_source::{DocumentNumberingSource, DocumentNumberingMedia,
    package_part_name, resolve_relationship};
use super::{Source, Element, Content, WORD_NAMESPACE, DRAWING_NAMESPACE, XMLNS_NAMESPACE, XML_NAMESPACE,
    STYLE_DEPTH_LIMIT, STYLE_LIMIT, attach_namespaces, namespaces, parse_xml, canonicalize_office_namespaces,
    source_style, style_chain, tables};
use crate::{ExportError, opc::Package, xml};

#[path = "docx_source_numbering_dependencies.rs"]
mod dependencies;
#[path = "docx_source_numbering_media.rs"]
mod media;
pub(super) use dependencies::{expand, has_live_numbering};

const NUMBERING_LIMIT: usize = 4096;
const MAX_ID: u32 = i32::MAX as u32;
const VML_NAMESPACE: &str = "urn:schemas-microsoft-com:vml";
const OFFICE_NAMESPACE: &str = "urn:schemas-microsoft-com:office:office";
const REL_NAMESPACE: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_REL_NAMESPACE: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const MC_NAMESPACE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

pub(super) struct SourceNumbering {
    part_name: String,
    numbers: BTreeMap<u32, Element>,
    abstracts: BTreeMap<u32, Element>,
    pictures: BTreeMap<u32, Element>,
    shape_types: BTreeMap<String, Element>,
    relationships: BTreeMap<String, Relationship>,
    media: BTreeMap<String, DocumentNumberingMedia>,
    vml_ids: BTreeSet<String>,
}

struct Relationship { kind: String, target: String, external: bool }

#[derive(Default)]
pub(super) struct Dependencies {
    pub(super) styles: BTreeMap<String, Element>,
    numbers: BTreeMap<u32, Element>,
    abstracts: BTreeMap<u32, Element>,
    pictures: BTreeMap<u32, Element>,
    shape_types: BTreeMap<String, Element>,
}

impl Dependencies {
    pub(super) fn theme_requirements(&self, initial: (bool, bool)) -> (bool, bool) {
        self.styles.values().chain(self.numbers.values()).chain(self.abstracts.values()).chain(self.pictures.values()).chain(self.shape_types.values())
            .fold(initial, |required, node| { let next = super::theme_requirements(node); (required.0 || next.0, required.1 || next.1) })
    }
}

#[derive(Default)]
pub(super) struct NumberingRegistry {
    generated: Option<Element>,
    used_numbers: BTreeSet<u32>,
    used_abstracts: BTreeSet<u32>,
    used_pictures: BTreeSet<u32>,
    number_ids: BTreeMap<u32, u32>,
    abstract_ids: BTreeMap<u32, u32>,
    picture_ids: BTreeMap<u32, u32>,
    numbers: BTreeMap<u32, Element>,
    abstracts: BTreeMap<u32, Element>,
    pictures: BTreeMap<u32, Element>,
    shape_types: BTreeMap<String, Element>,
    used_vml_ids: BTreeSet<String>,
    vml_ids: BTreeMap<String, String>,
    next_vml_id: u32,
}

fn unsigned(value: Option<&str>, label: &str) -> Result<u32, String> {
    value.and_then(|value| value.parse::<u32>().ok()).filter(|value| *value <= MAX_ID)
        .ok_or_else(|| format!("DOCX_NUMBERING_ID: {label} requires a nonnegative 32-bit signed Office ID"))
}

fn plain_attribute<'a>(node: &'a Element, name: &str) -> Option<&'a str> {
    node.attributes.iter().find(|attribute| attribute.namespace.is_empty() && attribute.local == name).map(|attribute| attribute.value.as_str())
}

fn supported_namespace(namespace: &str) -> bool {
    matches!(namespace, WORD_NAMESPACE | DRAWING_NAMESPACE | VML_NAMESPACE | OFFICE_NAMESPACE | MC_NAMESPACE
        | "urn:schemas-microsoft-com:office:word")
        || namespace.starts_with("http://schemas.openxmlformats.org/drawingml/2006/")
        || namespace.starts_with("http://schemas.microsoft.com/office/word/")
        || namespace.starts_with("http://schemas.microsoft.com/office/drawing/")
}

fn validate_resource(node: &Element) -> Result<(), String> {
    if !supported_namespace(&node.namespace) {
        return Err(format!("DOCX_NUMBERING_RESOURCE: unsupported numbering element namespace {:?}", node.namespace));
    }
    for attribute in &node.attributes {
        if matches!(attribute.namespace.as_str(), "" | XML_NAMESPACE | XMLNS_NAMESPACE) || supported_namespace(&attribute.namespace) { continue; }
        if attribute.namespace == REL_NAMESPACE && matches!(attribute.local.as_str(), "id" | "embed" | "link") { continue; }
        return Err("DOCX_NUMBERING_RESOURCE: a numbering attribute depends on an unmapped namespace or relationship".into());
    }
    for child in node.children() { validate_resource(child)?; }
    Ok(())
}

fn definitions(root: &Element, local: &str, attribute: &str) -> Result<BTreeMap<u32, Element>, String> {
    let scope = namespaces(root);
    let mut output = BTreeMap::new();
    for original in root.children().filter(|node| node.is_word(local)) {
        let id = unsigned(original.attribute(attribute), attribute)?;
        let mut node = original.clone();
        attach_namespaces(&mut node, &scope);
        if output.insert(id, node).is_some() { return Err(format!("DOCX_NUMBERING_DUPLICATE: duplicate {local} ID {id}")); }
        if output.len() > NUMBERING_LIMIT { return Err("DOCX_NUMBERING_RESOURCE_LIMIT: too many source numbering definitions".into()); }
    }
    Ok(output)
}

fn shape_signature(node: &Element) -> String {
    fn visit(node: &Element, output: &mut String) {
        output.push_str(&format!("{:?}:{:?}[", node.namespace, node.local));
        let mut attributes = node.attributes.iter().filter(|attribute| attribute.namespace != XMLNS_NAMESPACE)
            .map(|attribute| (&attribute.namespace, &attribute.local, &attribute.value)).collect::<Vec<_>>();
        attributes.sort();
        for attribute in attributes { output.push_str(&format!("{attribute:?}")); }
        output.push(']');
        for child in &node.content { match child {
            Content::Element(child) => visit(child, output), Content::Text(text) if !text.trim().is_empty() => output.push_str(&format!("{text:?}")), _ => {},
        } }
        output.push('/');
    }
    let mut output = String::new(); visit(node, &mut output); output
}

fn collect_shapes(node: &Element, inherited: &BTreeMap<String, String>, shape_types: &mut BTreeMap<String, Element>, ids: &mut BTreeSet<String>) -> Result<(), String> {
    let mut scope = inherited.clone(); scope.extend(namespaces(node));
    if node.namespace == VML_NAMESPACE && matches!(node.local.as_str(), "shape" | "group" | "shapetype") {
        let id = plain_attribute(node, "id").filter(|id| !id.is_empty());
        if node.local == "shapetype" {
            let id = id.ok_or("DOCX_NUMBERING_VML: a numbering shape type has no ID")?;
            let mut definition = node.clone(); attach_namespaces(&mut definition, &scope);
            if let Some(prior) = shape_types.get(id) {
                if shape_signature(prior) != shape_signature(&definition) { return Err("DOCX_NUMBERING_VML: duplicate shape type IDs have different definitions".into()); }
            } else {
                if ids.contains(id) { return Err("DOCX_NUMBERING_VML: a shape type ID conflicts with a shape ID".into()); }
                shape_types.insert(id.into(), definition);
                ids.insert(id.into());
            }
        } else {
            let mut aliases = BTreeSet::new();
            if let Some(id) = id { aliases.insert(id); }
            if let Some(id) = node.attributes.iter().find(|attribute| attribute.namespace == OFFICE_NAMESPACE && attribute.local == "spid").map(|attribute| attribute.value.as_str()) { aliases.insert(id); }
            for id in aliases {
                if id.is_empty() || !ids.insert(id.into()) { return Err("DOCX_NUMBERING_VML: numbering shapes contain duplicate or empty IDs".into()); }
            }
        }
    }
    for child in node.children() { collect_shapes(child, &scope, shape_types, ids)?; }
    Ok(())
}

pub(super) fn read(value: &Value) -> Result<Option<SourceNumbering>, String> {
    if value.is_null() { return Ok(None); }
    let source = DocumentNumberingSource::deserialize(value).map_err(|error| format!("DOCX_NUMBERING_SOURCE: {error}"))?;
    source.validate()?;
    let mut root = parse_xml(&source.xml, true)?;
    canonicalize_office_namespaces(&mut root)?;
    if !root.is_word("numbering") { return Err("DOCX_NUMBERING_SOURCE: numbering XML has an invalid root".into()); }
    let mut relationships = BTreeMap::new();
    if let Some(xml) = &source.relationships_xml {
        let root = parse_xml(xml, true)?;
        if root.namespace != PACKAGE_REL_NAMESPACE || root.local != "Relationships" { return Err("DOCX_NUMBERING_SOURCE: invalid numbering relationships root".into()); }
        for node in root.children() {
            if node.namespace != PACKAGE_REL_NAMESPACE || node.local != "Relationship" { return Err("DOCX_NUMBERING_SOURCE: invalid numbering relationship element".into()); }
            let id = plain_attribute(node, "Id").filter(|id| !id.is_empty()).ok_or("DOCX_NUMBERING_RELATIONSHIP: a source relationship has no ID")?;
            let kind = plain_attribute(node, "Type").filter(|kind| !kind.is_empty()).ok_or("DOCX_NUMBERING_RELATIONSHIP: a source relationship has no type")?;
            let external = match plain_attribute(node, "TargetMode") { None | Some("Internal") => false, Some("External") => true, _ => return Err("DOCX_NUMBERING_RELATIONSHIP: invalid relationship target mode".into()) };
            let target = plain_attribute(node, "Target").filter(|target| !target.is_empty()).ok_or("DOCX_NUMBERING_RELATIONSHIP: a source relationship has no target")?;
            if !external { resolve_relationship(&source.part_name, target)?; }
            if relationships.insert(id.into(), Relationship { kind: kind.into(), target: target.into(), external }).is_some() {
                return Err("DOCX_NUMBERING_RELATIONSHIP: duplicate source relationship ID".into());
            }
            if relationships.len() > office_io_common::document_numbering_source::RESOURCE_COUNT_LIMIT { return Err("DOCX_NUMBERING_RESOURCE_LIMIT: too many source numbering relationships".into()); }
        }
    }
    let numbers = definitions(&root, "num", "numId")?;
    let abstracts = definitions(&root, "abstractNum", "abstractNumId")?;
    let pictures = definitions(&root, "numPicBullet", "numPicBulletId")?;
    let mut shape_types = BTreeMap::new(); let mut vml_ids = BTreeSet::new();
    collect_shapes(&root, &BTreeMap::new(), &mut shape_types, &mut vml_ids)?;
    let media = source.resources.into_iter().map(|media| Ok((package_part_name(&media.part_name)?, media)))
        .collect::<Result<BTreeMap<_, _>, String>>()?;
    Ok(Some(SourceNumbering { part_name: source.part_name, numbers, abstracts, pictures, shape_types, relationships, media, vml_ids }))
}

fn allocate(used: &mut BTreeSet<u32>, first: u32) -> Result<u32, String> {
    let mut next = first;
    while used.contains(&next) { next = next.checked_add(1).filter(|next| *next <= MAX_ID).ok_or("DOCX_NUMBERING_ID_LIMIT: imported numbering IDs are exhausted")?; }
    used.insert(next); Ok(next)
}

impl NumberingRegistry {
    pub(super) fn reserve(&mut self, generated: &str) -> Result<(), String> {
        if self.generated.is_some() || !self.numbers.is_empty() || !self.abstracts.is_empty() || !self.pictures.is_empty() {
            return Err("DOCX_NUMBERING_ORDER: generated numbering must be reserved exactly once before source import".into());
        }
        let mut root = parse_xml(generated, true)?; canonicalize_office_namespaces(&mut root)?;
        if !root.is_word("numbering") { return Err("DOCX_NUMBERING_ORDER: generated numbering has an invalid root".into()); }
        self.used_numbers = definitions(&root, "num", "numId")?.into_keys().collect();
        self.used_abstracts = definitions(&root, "abstractNum", "abstractNumId")?.into_keys().collect();
        self.used_pictures = definitions(&root, "numPicBullet", "numPicBulletId")?.into_keys().collect();
        collect_shapes(&root, &BTreeMap::new(), &mut BTreeMap::new(), &mut self.used_vml_ids)?;
        self.next_vml_id = 1_073_741_824;
        self.generated = Some(root); Ok(())
    }

    pub(super) fn import(&mut self, source: &Source, needed: &Dependencies) -> Result<(), String> {
        if needed.numbers.is_empty() && needed.abstracts.is_empty() && needed.pictures.is_empty() { return Ok(()); }
        if self.generated.is_none() { return Err("DOCX_NUMBERING_ORDER: source numbering was imported before generated IDs were reserved".into()); }
        for (id, node) in &needed.numbers {
            if self.number_ids.contains_key(id) { continue; }
            self.number_ids.insert(*id, allocate(&mut self.used_numbers, 1)?); self.numbers.insert(*id, node.clone());
        }
        for (id, node) in &needed.abstracts {
            if self.abstract_ids.contains_key(id) { continue; }
            self.abstract_ids.insert(*id, allocate(&mut self.used_abstracts, 0)?); self.abstracts.insert(*id, node.clone());
        }
        for (id, node) in &needed.pictures {
            if self.picture_ids.contains_key(id) { continue; }
            self.picture_ids.insert(*id, allocate(&mut self.used_pictures, 1)?); self.pictures.insert(*id, node.clone());
        }
        for (id, node) in &needed.shape_types { self.shape_types.entry(id.clone()).or_insert_with(|| node.clone()); }
        let source = source.numbering.as_ref().map_err(Clone::clone)?.as_ref().ok_or("DOCX_NUMBERING_SOURCE_MISSING: referenced source numbering was not retained")?;
        self.reserve_shapes(source)?;
        Ok(())
    }

    /// numId=0 suppresses numbering. It deliberately needs no resource mapping.
    pub(super) fn remap(&self, node: &mut Element) -> Result<(), String> {
        if node.namespace == WORD_NAMESPACE {
            let mapping = match node.local.as_str() { "numId" => Some((&self.number_ids, true)), "abstractNumId" => Some((&self.abstract_ids, false)),
                "lvlPicBulletId" => Some((&self.picture_ids, false)), _ => None };
            if let Some((mapping, suppression)) = mapping {
                let id = unsigned(node.attribute("val"), &node.local)?;
                if !(suppression && id == 0) {
                    let mapped = mapping.get(&id).ok_or_else(|| format!("DOCX_NUMBERING_UNMAPPED: {} {id} has no imported source resource", node.local))?;
                    node.set_word_attribute("val", mapped.to_string());
                }
            }
        }
        for child in &mut node.content { if let Content::Element(child) = child { self.remap(child)?; } }
        Ok(())
    }

    pub(super) fn finish(&self, source: Option<&SourceNumbering>, styles: &BTreeMap<String, String>, package: &mut Package) -> Result<String, ExportError> {
        let mut root = self.generated.clone().ok_or_else(|| ExportError::invalid("DOCX_NUMBERING_ORDER: generated numbering was not reserved"))?;
        let mut imported_pictures = Vec::new();
        for (id, original) in &self.pictures {
            let mut node = original.clone(); node.set_word_attribute("numPicBulletId", self.picture_ids[id].to_string());
            imported_pictures.push(node);
        }
        self.write_picture_resources(source, &mut imported_pictures, package)?;
        let mut pictures = Vec::new(); let mut abstracts = Vec::new(); let mut numbers = Vec::new(); let mut tail = Vec::new();
        for content in std::mem::take(&mut root.content) { match &content {
            Content::Element(node) if node.is_word("numPicBullet") => pictures.push(content),
            Content::Element(node) if node.is_word("abstractNum") => abstracts.push(content),
            Content::Element(node) if node.is_word("num") => numbers.push(content), _ => tail.push(content),
        } }
        for mut node in imported_pictures {
            tables::remap_references(&mut node, styles); self.remap(&mut node).map_err(ExportError::invalid)?;
            pictures.push(Content::Element(node));
        }
        for (id, original) in &self.abstracts {
            let mut node = original.clone(); node.set_word_attribute("abstractNumId", self.abstract_ids[id].to_string());
            tables::remap_references(&mut node, styles); self.remap(&mut node).map_err(ExportError::invalid)?;
            abstracts.push(Content::Element(node));
        }
        for (id, original) in &self.numbers {
            let mut node = original.clone(); node.set_word_attribute("numId", self.number_ids[id].to_string());
            tables::remap_references(&mut node, styles); self.remap(&mut node).map_err(ExportError::invalid)?;
            numbers.push(Content::Element(node));
        }
        root.content.extend(pictures); root.content.extend(abstracts); root.content.extend(numbers); root.content.extend(tail);
        super::super::source_xml::validate_namespace_bindings(&root, &BTreeMap::new()).map_err(ExportError::invalid)?;
        Ok(format!("{}{}", xml::DECL, root.render(None)))
    }
}
