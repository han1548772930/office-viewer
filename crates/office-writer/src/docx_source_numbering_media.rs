//! Numbering picture parts, VML IDs, and package relationships share one owner.
use office_io_common::document_numbering_source::decode_media;
use super::*;

const IMAGE_RELATIONSHIP: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

fn external_source(value: &str) -> bool { value.contains(':') || value.starts_with("//") }

fn relationship_attribute(attribute: &super::super::Attribute) -> bool {
    attribute.namespace == REL_NAMESPACE && matches!(attribute.local.as_str(), "id" | "embed" | "link")
        || attribute.namespace == OFFICE_NAMESPACE && attribute.local == "relid"
}

fn picture_references(node: &Element, relationships: &mut BTreeSet<String>, sources: &mut BTreeSet<String>) {
    for attribute in &node.attributes {
        if relationship_attribute(attribute) { relationships.insert(attribute.value.clone()); }
        if node.namespace == VML_NAMESPACE && node.local == "imagedata" && attribute.namespace.is_empty() && attribute.local == "src"
            && !attribute.value.is_empty() && !external_source(&attribute.value) { sources.insert(attribute.value.clone()); }
    }
    for child in node.children() { picture_references(child, relationships, sources); }
}

fn remove_shape_types(node: &mut Element) {
    node.content.retain(|content| !matches!(content, Content::Element(child) if child.namespace == VML_NAMESPACE && child.local == "shapetype"));
    for child in &mut node.content { if let Content::Element(child) = child { remove_shape_types(child); } }
}

fn shape_ids(node: &Element, output: &mut BTreeSet<String>) {
    if node.namespace == VML_NAMESPACE && matches!(node.local.as_str(), "shape" | "group" | "shapetype") {
        if let Some(id) = plain_attribute(node, "id") { output.insert(id.into()); }
        if let Some(id) = node.attributes.iter().find(|attribute| attribute.namespace == OFFICE_NAMESPACE && attribute.local == "spid").map(|attribute| attribute.value.as_str()) { output.insert(id.into()); }
    }
    for child in node.children() { shape_ids(child, output); }
}

fn write_media(source: &SourceNumbering, path: &str, package: &mut Package, mapping: &mut BTreeMap<String, String>, next: &mut u32) -> Result<String, ExportError> {
    if let Some(mapped) = mapping.get(path) { return Ok(mapped.clone()); }
    let media = source.media.get(path).ok_or_else(|| ExportError::invalid(format!("DOCX_NUMBERING_MEDIA_MISSING: referenced numbering media {path:?} was not retained")))?;
    if !media.content_type.starts_with("image/") { return Err(ExportError::invalid("DOCX_NUMBERING_MEDIA_TYPE: a numbering image refers to a non-image content type")); }
    let extension = path.rsplit_once('.').map(|(_, extension)| extension).filter(|extension| !extension.is_empty() && extension.len() <= 16 && extension.bytes().all(|byte| byte.is_ascii_alphanumeric())).unwrap_or("bin");
    let target = loop {
        let target = format!("media/office-source-numbering-{}.{extension}", *next);
        *next = next.checked_add(1).ok_or_else(|| ExportError::invalid("DOCX_NUMBERING_MEDIA_LIMIT: image part IDs are exhausted"))?;
        if !package.contains_part(&format!("word/{target}")) { break target; }
    };
    package.add_binary(&format!("word/{target}"), &media.content_type, decode_media(&media.data_base64).map_err(ExportError::invalid)?)?;
    mapping.insert(path.into(), target.clone()); Ok(target)
}

impl NumberingRegistry {
    pub(super) fn reserve_shapes(&mut self, source: &SourceNumbering) -> Result<(), String> {
        let mut needed = BTreeSet::new();
        for node in self.pictures.values().chain(self.shape_types.values()) { shape_ids(node, &mut needed); }
        for id in needed {
            if self.vml_ids.contains_key(&id) { continue; }
            let prefix = if self.shape_types.contains_key(&id) { 't' } else { 's' };
            let mapped = loop {
                let candidate = format!("_x0000_{prefix}{}", self.next_vml_id);
                self.next_vml_id = self.next_vml_id.checked_add(1).filter(|next| *next <= MAX_ID).ok_or("DOCX_NUMBERING_VML: imported VML IDs are exhausted")?;
                if !source.vml_ids.contains(&candidate) && !self.used_vml_ids.contains(&candidate) { break candidate; }
            };
            self.used_vml_ids.insert(mapped.clone()); self.vml_ids.insert(id, mapped);
        }
        Ok(())
    }

    fn remap_shapes(&self, node: &mut Element) -> Result<(), String> {
        for attribute in &mut node.attributes {
            if node.namespace == VML_NAMESPACE && matches!(node.local.as_str(), "shape" | "group" | "shapetype") && attribute.namespace.is_empty() {
                if attribute.local == "id" {
                    attribute.value = self.vml_ids.get(&attribute.value).ok_or("DOCX_NUMBERING_VML: a shape ID has no reserved mapping")?.clone();
                } else if attribute.local == "type" {
                    if let Some(id) = attribute.value.strip_prefix('#') {
                        attribute.value = format!("#{}", self.vml_ids.get(id).ok_or("DOCX_NUMBERING_VML: a shape type ID has no reserved mapping")?);
                    }
                }
            } else if attribute.namespace == OFFICE_NAMESPACE && attribute.local == "spid" {
                attribute.value = self.vml_ids.get(&attribute.value).ok_or("DOCX_NUMBERING_VML: a shape's Office ID has no reserved mapping")?.clone();
            }
        }
        for child in &mut node.content { if let Content::Element(child) = child { self.remap_shapes(child)?; } }
        Ok(())
    }

    pub(super) fn write_picture_resources(&self, source: Option<&SourceNumbering>, pictures: &mut [Element], package: &mut Package) -> Result<(), ExportError> {
        if pictures.is_empty() { return Ok(()); }
        let source = source.ok_or_else(|| ExportError::invalid("DOCX_NUMBERING_SOURCE_MISSING: numbering picture resources were not retained"))?;
        for picture in pictures.iter_mut() { remove_shape_types(picture); }
        if !self.shape_types.is_empty() {
            let pict = pictures[0].content.iter_mut().find_map(|content| match content { Content::Element(node) if node.is_word("pict") => Some(node), _ => None })
                .ok_or_else(|| ExportError::invalid("DOCX_NUMBERING_PICTURE: a numbering picture has no pict container for shape types"))?;
            pict.content.splice(0..0, self.shape_types.values().cloned().map(Content::Element));
        }
        let mut relationship_ids = BTreeSet::new(); let mut direct_sources = BTreeSet::new();
        for picture in pictures.iter() { picture_references(picture, &mut relationship_ids, &mut direct_sources); }
        let mut media_mapping = BTreeMap::new(); let mut relationship_mapping = BTreeMap::new(); let mut next_part = 1u32; let mut next_relationship = 1u32;
        for id in relationship_ids {
            let relationship = source.relationships.get(&id).ok_or_else(|| ExportError::invalid(format!("DOCX_NUMBERING_RELATIONSHIP_MISSING: numbering picture relationship {id:?} was not retained")))?;
            if !matches!(relationship.kind.as_str(), IMAGE_RELATIONSHIP | "http://purl.oclc.org/ooxml/officeDocument/relationships/image") {
                return Err(ExportError::invalid("DOCX_NUMBERING_RELATIONSHIP: a numbering picture depends on a non-image relationship"));
            }
            let target = if relationship.external { relationship.target.clone() } else {
                let path = resolve_relationship(&source.part_name, &relationship.target).map_err(ExportError::invalid)?;
                write_media(source, &path, package, &mut media_mapping, &mut next_part)?
            };
            let mapped = loop {
                let candidate = format!("rIdOfficeSourceNumbering{next_relationship}");
                next_relationship = next_relationship.checked_add(1).ok_or_else(|| ExportError::invalid("DOCX_NUMBERING_RELATIONSHIP_LIMIT: imported relationship IDs are exhausted"))?;
                if !package.contains_relationship(Some("word/numbering.xml"), &candidate) { break candidate; }
            };
            package.add_relationship(Some("word/numbering.xml"), &mapped, IMAGE_RELATIONSHIP, &target, relationship.external)?;
            relationship_mapping.insert(id, mapped);
        }
        let mut direct_mapping = BTreeMap::new();
        for original in direct_sources {
            let path = resolve_relationship(&source.part_name, &original).map_err(ExportError::invalid)?;
            let target = write_media(source, &path, package, &mut media_mapping, &mut next_part)?;
            direct_mapping.insert(original, target);
        }
        fn rewrite(node: &mut Element, relationships: &BTreeMap<String, String>, direct: &BTreeMap<String, String>) -> Result<(), String> {
            for attribute in &mut node.attributes {
                if relationship_attribute(attribute) {
                    attribute.value = relationships.get(&attribute.value).ok_or("DOCX_NUMBERING_RELATIONSHIP: a picture relationship has no output mapping")?.clone();
                } else if node.namespace == VML_NAMESPACE && node.local == "imagedata" && attribute.namespace.is_empty() && attribute.local == "src" {
                    if let Some(mapped) = direct.get(&attribute.value) { attribute.value = mapped.clone(); }
                }
            }
            for child in &mut node.content { if let Content::Element(child) = child { rewrite(child, relationships, direct)?; } }
            Ok(())
        }
        for picture in pictures {
            rewrite(picture, &relationship_mapping, &direct_mapping).map_err(ExportError::invalid)?;
            self.remap_shapes(picture).map_err(ExportError::invalid)?;
        }
        Ok(())
    }
}
