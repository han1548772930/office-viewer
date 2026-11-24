use crate::error::ExportError;
use crate::xml;
use std::collections::{BTreeMap, BTreeSet, HashMap};
use std::io::Cursor;
use zip::write::FileOptions;

#[derive(Debug, Clone)]
struct Relationship {
    id: String,
    kind: String,
    target: String,
    external: bool,
}

#[derive(Debug, Default)]
pub(crate) struct Package {
    parts: BTreeMap<String, (String, Vec<u8>)>,
    part_keys: BTreeSet<String>,
    part_names: BTreeMap<String, String>,
    relationships: HashMap<Option<String>, Vec<Relationship>>,
    preserved_relationships: HashMap<Option<String>, BTreeSet<String>>,
}

impl Package {
    pub(crate) fn new() -> Self {
        Self::default()
    }

    pub(crate) fn contains_part(&self, path: &str) -> bool {
        office_io_common::document_custom_xml::part_name_key(path.trim_start_matches('/')).is_ok_and(|key| self.part_keys.contains(&key))
    }

    pub(crate) fn contains_relationship(&self, source: Option<&str>, id: &str) -> bool {
        let source = source.map(|source| office_io_common::document_custom_xml::part_name_key(source.trim_start_matches('/'))).transpose();
        let Ok(source) = source else { return false; };
        self.relationships.get(&source).is_some_and(|entries| entries.iter().any(|entry| entry.id == id))
            || self.preserved_relationships.get(&source).is_some_and(|ids| ids.contains(id))
    }

    pub(crate) fn add_xml(
        &mut self,
        path: &str,
        content_type: &str,
        content: String,
    ) -> Result<(), ExportError> {
        self.add_part(path, content_type, content.into_bytes())
    }

    pub(crate) fn add_binary(
        &mut self,
        path: &str,
        content_type: &str,
        bytes: Vec<u8>,
    ) -> Result<(), ExportError> {
        self.add_part(path, content_type, bytes)
    }

    fn add_part(
        &mut self,
        path: &str,
        content_type: &str,
        content: Vec<u8>,
    ) -> Result<(), ExportError> {
        let normalized = office_io_common::document_custom_xml::canonical_part_name(path.trim_start_matches('/')).map_err(ExportError::invalid)?;
        let key = normalized.to_ascii_lowercase();
        if !office_io_common::document_custom_xml::valid_content_type(content_type) { return Err(ExportError::invalid("invalid OPC MIME content type")); }
        if key.starts_with("_rels/")
            || key.ends_with(".rels")
            || key == "%5bcontent_types%5d.xml"
        {
            return Err(ExportError::invalid(format!(
                "invalid OPC part path: {path}"
            )));
        }
        if self.part_keys.contains(&key) {
            return Err(ExportError::invalid(format!(
                "duplicate OPC part: {normalized}"
            )));
        }
        self.part_names.insert(key.clone(), normalized.clone());
        self.part_keys.insert(key);
        self.parts.insert(normalized, (content_type.to_string(), content));
        Ok(())
    }

    /// Preserve a relationship owner's actual XML only when no generated owner
    /// exists. Later additions are rejected instead of silently replacing it.
    pub(crate) fn add_preserved_relationships(&mut self, source: &str, bytes: Vec<u8>) -> Result<(), ExportError> {
        use office_io_common::document_custom_xml::{canonical_part_name, decode_xml, relationship_part, relationships, RELATIONSHIP_CONTENT_TYPE};
        let source = canonical_part_name(source).map_err(ExportError::invalid)?;
        let source_key = source.to_ascii_lowercase();
        if !self.part_keys.contains(&source_key) { return Err(ExportError::invalid("DOCX_CUSTOM_XML_MISSING_PART: preserved relationship owner does not exist")); }
        let key = Some(source_key);
        let path = relationship_part(&source);
        let path_key = path.to_ascii_lowercase();
        if self.relationships.contains_key(&key) || self.preserved_relationships.contains_key(&key) || self.part_keys.contains(&path_key) {
            return Err(ExportError::invalid("DOCX_CUSTOM_XML_RELATIONSHIP_CONFLICT: relationship owner already exists"));
        }
        let (xml, _, _) = decode_xml(&bytes).map_err(ExportError::invalid)?;
        let ids = relationships(&xml).map_err(ExportError::invalid)?.into_iter().map(|relationship| relationship.id).collect();
        self.part_names.insert(path_key.clone(), path.clone());
        self.part_keys.insert(path_key);
        self.parts.insert(path, (RELATIONSHIP_CONTENT_TYPE.into(), bytes));
        self.preserved_relationships.insert(key, ids);
        Ok(())
    }

    pub(crate) fn add_relationship(
        &mut self,
        source: Option<&str>,
        id: &str,
        kind: &str,
        target: &str,
        external: bool,
    ) -> Result<(), ExportError> {
        if id.is_empty() || kind.is_empty() || target.is_empty() {
            return Err(ExportError::invalid(
                "relationship id, type and target are required",
            ));
        }
        let source = source.map(|value| office_io_common::document_custom_xml::part_name_key(value.trim_start_matches('/'))).transpose().map_err(ExportError::invalid)?;
        if self.preserved_relationships.contains_key(&source) {
            return Err(ExportError::invalid("A preserved OPC relationship owner cannot receive generated relationships."));
        }
        let entries = self.relationships.entry(source).or_default();
        if entries.iter().any(|entry| entry.id == id) {
            return Err(ExportError::invalid(format!(
                "duplicate relationship id: {id}"
            )));
        }
        entries.push(Relationship {
            id: id.to_string(),
            kind: kind.to_string(),
            target: target.to_string(),
            external,
        });
        Ok(())
    }

    pub(crate) fn finish(self) -> Result<Vec<u8>, ExportError> {
        let mut package = self;
        let content_types = package.content_types();
        package.parts.insert(
            "[Content_Types].xml".to_string(),
            ("application/xml".to_string(), content_types.into_bytes()),
        );
        let relationships = std::mem::take(&mut package.relationships);
        for (source, entries) in relationships {
            let source_name = source.as_ref().map(|key| package.part_names.get(key)
                .ok_or_else(|| ExportError::invalid(format!("missing OPC relationship owner: {key}")))).transpose()?;
            let path = relationship_path(source_name.map(String::as_str));
            for entry in entries.iter().filter(|entry| !entry.external) {
                let target = office_io_common::document_custom_xml::resolve_target(source_name.map_or("", String::as_str), &entry.target).map_err(ExportError::invalid)?;
                if !package.contains_part(&target) { return Err(ExportError::invalid(format!("missing OPC relationship target: {target}"))); }
            }
            let body = relationship_xml(&entries);
            if package.part_keys.contains(&path.to_ascii_lowercase()) { return Err(ExportError::invalid(format!("duplicate OPC relationship part: {path}"))); }
            package.part_keys.insert(path.to_ascii_lowercase());
            package.parts.insert(
                path,
                (
                    "application/vnd.openxmlformats-package.relationships+xml".to_string(),
                    body.into_bytes(),
                ),
            );
        }
        let mut output = Cursor::new(Vec::new());
        let options =
            FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);
        let mut archive = zip::ZipWriter::new(&mut output);
        for (path, (_, bytes)) in package.parts {
            let zip_name = office_io_common::document_custom_xml::zip_part_name(&path).map_err(ExportError::invalid)?;
            archive
                .start_file(zip_name, options)
                .map_err(|error| ExportError::package(error.to_string()))?;
            std::io::Write::write_all(&mut archive, &bytes)
                .map_err(|error| ExportError::package(error.to_string()))?;
        }
        archive
            .finish()
            .map_err(|error| ExportError::package(error.to_string()))?;
        Ok(output.into_inner())
    }

    fn content_types(&self) -> String {
        let mut by_extension = BTreeMap::<String, Vec<(String, String)>>::new();
        for (path, (content_type, _)) in &self.parts {
            if path == "[Content_Types].xml" {
                continue;
            }
            if let Some(extension) = path.rsplit('.').next().filter(|value| !value.is_empty()) {
                by_extension
                    .entry(extension.to_ascii_lowercase())
                    .or_default()
                    .push((path.clone(), content_type.clone()));
            } else {
                by_extension
                    .entry(String::new())
                    .or_default()
                    .push((path.clone(), content_type.clone()));
            }
        }

        // The package content-types part itself is XML, so retain the standard
        // XML default and add overrides for the OOXML parts with specialised
        // content types. Relationship parts are generated after this method and
        // therefore need their required default declared here as well.
        let mut defaults = BTreeMap::<String, String>::from([
            (
                "rels".to_owned(),
                "application/vnd.openxmlformats-package.relationships+xml".to_owned(),
            ),
            ("xml".to_owned(), "application/xml".to_owned()),
        ]);
        let mut overrides = BTreeMap::<String, String>::new();
        for (extension, entries) in by_extension {
            if extension.is_empty() {
                for (path, content_type) in entries {
                    overrides.insert(format!("/{path}"), content_type);
                }
                continue;
            }
            let first_type = entries
                .first()
                .map(|(_, content_type)| content_type.clone());
            let all_same = first_type
                .as_ref()
                .is_some_and(|content_type| entries.iter().all(|(_, value)| value == content_type));
            let reserved_default = defaults.get(&extension).cloned();
            if all_same && reserved_default.is_none() {
                if let Some(content_type) = first_type {
                    defaults.insert(extension, content_type);
                }
            } else {
                for (path, content_type) in entries {
                    if reserved_default
                        .as_ref()
                        .is_none_or(|default| default != &content_type)
                    {
                        overrides.insert(format!("/{path}"), content_type);
                    }
                }
            }
        }
        let mut body = format!(
            "{}<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">",
            xml::DECL
        );
        for (extension, content_type) in defaults {
            body.push_str(&format!(
                "<Default Extension=\"{}\" ContentType=\"{}\"/>",
                xml::escape(&extension),
                xml::escape(&content_type)
            ));
        }
        for (part, content_type) in overrides {
            body.push_str(&format!(
                "<Override PartName=\"{}\" ContentType=\"{}\"/>",
                xml::escape(&part),
                xml::escape(&content_type)
            ));
        }
        body.push_str("</Types>");
        body
    }
}

fn relationship_path(source: Option<&str>) -> String {
    match source {
        None => "_rels/.rels".to_string(),
        Some(path) => {
            let (directory, file) = path.rsplit_once('/').unwrap_or(("", path));
            if directory.is_empty() {
                format!("_rels/{file}.rels")
            } else {
                format!("{directory}/_rels/{file}.rels")
            }
        }
    }
}

fn relationship_xml(entries: &[Relationship]) -> String {
    let mut body = format!(
        "{}<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">",
        xml::DECL
    );
    for entry in entries {
        body.push_str(&format!(
            "<Relationship Id=\"{}\" Type=\"{}\" Target=\"{}\"{} />",
            xml::escape(&entry.id),
            xml::escape(&entry.kind),
            xml::escape(&entry.target),
            if entry.external {
                " TargetMode=\"External\""
            } else {
                ""
            }
        ));
    }
    body.push_str("</Relationships>");
    body
}

#[cfg(test)]
mod tests {
    use super::Package;

    #[test]
    fn writes_deterministic_zip_entries_without_a_host_clock() {
        let create = || {
            let mut package = Package::new();
            package
                .add_xml("word/document.xml", "application/xml", "<document/>".into())
                .unwrap();
            package
                .add_relationship(None, "rId1", "officeDocument", "word/document.xml", false)
                .unwrap();
            package.finish().unwrap()
        };
        let first = create();
        assert_eq!(first, create());
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(first)).unwrap();
        for index in 0..archive.len() {
            let entry = archive.by_index(index).unwrap();
            assert_eq!(entry.last_modified(), Some(zip::DateTime::default()));
        }
    }

    #[test]
    fn creates_root_relationships_and_content_types() {
        let mut package = Package::new();
        package
            .add_xml(
                "xl/workbook.xml",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                "<workbook/>".into(),
            )
            .unwrap();
        package
            .add_relationship(None, "rId1", "type", "xl/workbook.xml", false)
            .unwrap();
        let bytes = package.finish().unwrap();
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
        assert!(archive.by_name("[Content_Types].xml").is_ok());
        assert!(archive.by_name("_rels/.rels").is_ok());
        let mut content_types = String::new();
        std::io::Read::read_to_string(
            &mut archive.by_name("[Content_Types].xml").unwrap(),
            &mut content_types,
        )
        .unwrap();
        assert!(content_types.contains("Extension=\"rels\""));
        assert!(content_types.contains("PartName=\"/xl/workbook.xml\""));
    }

    #[test]
    fn emits_overrides_for_mixed_xml_content_types() {
        let mut package = Package::new();
        package
            .add_xml(
                "xl/workbook.xml",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                "<workbook/>".into(),
            )
            .unwrap();
        package
            .add_xml(
                "xl/styles.xml",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
                "<styleSheet/>".into(),
            )
            .unwrap();
        package
            .add_relationship(None, "rId1", "type", "xl/workbook.xml", false)
            .unwrap();
        let bytes = package.finish().unwrap();
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
        let mut content_types = String::new();
        std::io::Read::read_to_string(
            &mut archive.by_name("[Content_Types].xml").unwrap(),
            &mut content_types,
        )
        .unwrap();
        assert!(content_types.contains("PartName=\"/xl/workbook.xml\""));
        assert!(content_types.contains("PartName=\"/xl/styles.xml\""));
        assert!(!content_types.contains("Extension=\"xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\""));
    }
}
