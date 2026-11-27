//! Retain native numbering, its relationships, and their actual package media.
//! No remote relationship is fetched and no source list is evaluated here.

use std::{collections::BTreeSet, io::{Cursor, Read}};
use duke_pages::OoxmlPackage;
use zip::ZipArchive;
use office_io_common::document_numbering_source::{DocumentNumberingSource, DocumentNumberingMedia,
    MEDIA_BYTE_LIMIT, TOTAL_BYTE_LIMIT, RESOURCE_COUNT_LIMIT, decoded_media_bytes, encode_media, resolve_relationship};
use super::{DocumentOfficeXmlPartPayload, Element, parse_xml, read_xml};

fn relationships_path(owner: &str) -> String {
    let (directory, file) = owner.rsplit_once('/').unwrap_or(("", owner));
    if directory.is_empty() { format!("_rels/{file}.rels") } else { format!("{directory}/_rels/{file}.rels") }
}

fn content_type<'a>(types: &'a Element, path: &str) -> Option<&'a str> {
    types.children.iter().find(|node| node.name == "Override" && node.attribute("PartName").is_some_and(|name| name.trim_start_matches('/') == path))
        .and_then(|node| node.attribute("ContentType"))
        .or_else(|| path.rsplit_once('.').and_then(|(_, extension)| types.children.iter().find(|node|
            node.name == "Default" && node.attribute("Extension").is_some_and(|name| name.eq_ignore_ascii_case(extension)))
                .and_then(|node| node.attribute("ContentType"))))
}

fn retain_media(source: &mut DocumentNumberingSource, content_types: Option<&str>,
    mut encode_part: impl FnMut(&str) -> Result<String, String>, errors: &mut Vec<String>) {
    let result = (|| -> Result<(), String> {
        let types = content_types.map(parse_xml).transpose()?;
        let mut targets = BTreeSet::new();
        if let Some(relationships) = &source.relationships_xml {
            let root = parse_xml(relationships)?;
            if root.name != "Relationships" { return Err("Numbering relationships have an invalid root".into()); }
            for relationship in root.children.iter().filter(|node| node.name == "Relationship" && node.attribute("TargetMode") != Some("External")) {
                let Some(target) = relationship.attribute("Target") else { errors.push("Numbering relationship has no target".into()); continue; };
                match resolve_relationship(&source.part_name, target) {
                    Ok(path) => { targets.insert(path); },
                    Err(error) => errors.push(error),
                }
            }
        }
        fn direct_sources(node: &Element, owner: &str, targets: &mut BTreeSet<String>, errors: &mut Vec<String>) {
            if node.name == "imagedata" {
                if let Some(target) = node.attribute("src").filter(|target| !target.is_empty() && !target.contains(':') && !target.starts_with("//")) {
                    match resolve_relationship(owner, target) { Ok(path) => { targets.insert(path); }, Err(error) => errors.push(error) }
                }
            }
            for child in &node.children { direct_sources(child, owner, targets, errors); }
        }
        direct_sources(&parse_xml(&source.xml)?, &source.part_name, &mut targets, errors);
        if targets.len() > RESOURCE_COUNT_LIMIT { return Err("Numbering media count exceeds 1024 resources".into()); }
        let mut total = source.xml.len() + source.relationships_xml.as_ref().map_or(0, String::len);
        for path in targets {
            let retain = (|| -> Result<DocumentNumberingMedia, String> {
                let kind = types.as_ref().and_then(|types| content_type(types, &path))
                    .ok_or_else(|| format!("Numbering media {path:?} has no retained content type"))?;
                let data_base64 = encode_part(&path)?;
                total = total.checked_add(decoded_media_bytes(&data_base64)?).filter(|size| *size <= TOTAL_BYTE_LIMIT)
                    .ok_or("Numbering XML and media exceed their 32 MiB source budget")?;
                Ok(DocumentNumberingMedia { part_name: path, content_type: kind.into(), data_base64 })
            })();
            match retain { Ok(media) => source.resources.push(media), Err(error) => errors.push(error) }
        }
        Ok(())
    })();
    if let Err(error) = result { errors.push(error); }
    if let Err(error) = source.validate() { errors.push(error); }
}

pub(super) fn read(archive: &mut ZipArchive<Cursor<&[u8]>>, part: DocumentOfficeXmlPartPayload) -> (DocumentNumberingSource, Vec<String>) {
    let mut errors = Vec::new();
    let relationships_xml = match read_xml(archive, &relationships_path(&part.part_name)) {
        Ok(xml) => xml, Err(error) => { errors.push(error); None },
    };
    let types = match read_xml(archive, "[Content_Types].xml") {
        Ok(xml) => xml, Err(error) => { errors.push(error); None },
    };
    let mut source = DocumentNumberingSource { part_name: part.part_name, xml: part.xml, relationships_xml, resources: Vec::new() };
    retain_media(&mut source, types.as_deref(), |path| {
        let entry = archive.by_name(path).map_err(|error| format!("Numbering media {path:?} could not be read: {error}"))?;
        if entry.size() > MEDIA_BYTE_LIMIT as u64 { return Err(format!("Numbering media {path:?} exceeds 16 MiB")); }
        let mut bytes = Vec::new();
        bytes.try_reserve_exact(entry.size() as usize).map_err(|_| "Unable to allocate numbering media")?;
        entry.take(MEDIA_BYTE_LIMIT as u64 + 1).read_to_end(&mut bytes).map_err(|error| error.to_string())?;
        encode_media(&bytes)
    }, &mut errors);
    (source, errors)
}

/// DOC resources are explicitly synthesized from the same OoxmlPackage used by
/// Duke's parser. Binary assets are borrowed during encoding, never cloned.
pub(super) fn synthesized(package: &OoxmlPackage) -> (Option<DocumentNumberingSource>, Vec<String>) {
    let Some(part) = package.parts.get("word/numbering.xml") else { return (None, Vec::new()); };
    let mut source = DocumentNumberingSource { part_name: "word/numbering.xml".into(), xml: part.content.clone(),
        relationships_xml: package.parts.get("word/_rels/numbering.xml.rels").map(|part| part.content.clone()), resources: Vec::new() };
    let mut errors = Vec::new();
    retain_media(&mut source, package.parts.get("[Content_Types].xml").map(|part| part.content.as_str()), |path| {
        let bytes = package.binary_assets.get(path).ok_or_else(|| format!("Synthesized DOC numbering media {path:?} is missing"))?;
        encode_media(bytes)
    }, &mut errors);
    (Some(source), errors)
}

#[cfg(test)]
mod tests {
    use std::io::Write;
    use zip::{ZipWriter, write::SimpleFileOptions};
    use super::*;

    const WORD: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
    const IMAGE: &[u8] = b"real retained picture bytes";

    fn archive_parts(parts: &[(&str, &[u8])]) -> Vec<u8> {
        let mut archive = ZipWriter::new(Cursor::new(Vec::new()));
        for (path, data) in parts {
            archive.start_file(*path, SimpleFileOptions::default()).unwrap();
            archive.write_all(data).unwrap();
        }
        archive.finish().unwrap().into_inner()
    }

    #[test]
    fn package_relationships_locate_the_real_numbering_owner_and_media_without_fetching_external_images() {
        let types = r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="png" ContentType="image/png"/><Override PartName="/pkg/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/pkg/list/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>"#;
        let main_rels = format!(r#"<Relationships xmlns="{REL}"><Relationship Id="list" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="list/numbering.xml"/></Relationships>"#);
        let numbering = format!(r#"<w:numbering xmlns:w="{WORD}" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:numPicBullet w:numPicBulletId="4"><w:pict><v:shape><v:imagedata r:id="picture"/></v:shape></w:pict></w:numPicBullet></w:numbering>"#);
        let rels = format!(r#"<Relationships xmlns="{REL}"><Relationship Id="picture" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/bullet.png"/><Relationship Id="external" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="https://example.test/bullet.png" TargetMode="External"/></Relationships>"#);
        let bytes = archive_parts(&[("[Content_Types].xml", types.as_bytes()), ("pkg/_rels/document.xml.rels", main_rels.as_bytes()),
            ("pkg/list/numbering.xml", numbering.as_bytes()), ("pkg/list/_rels/numbering.xml.rels", rels.as_bytes()), ("pkg/media/bullet.png", IMAGE)]);
        let mut archive = ZipArchive::new(Cursor::new(bytes.as_slice())).unwrap();
        let (source, warnings) = super::super::read_office_source(&mut archive);
        assert!(warnings.is_empty(), "{warnings:?}");
        let source = source.unwrap();
        assert!(source.resource_errors.is_empty());
        let numbering_source = source.numbering.unwrap();
        assert_eq!(numbering_source.part_name, "pkg/list/numbering.xml");
        assert_eq!(numbering_source.xml, numbering);
        assert_eq!(numbering_source.relationships_xml.as_deref(), Some(rels.as_str()));
        assert_eq!(numbering_source.resources.len(), 1);
        assert_eq!(numbering_source.resources[0].part_name, "pkg/media/bullet.png");
        assert_eq!(numbering_source.resources[0].content_type, "image/png");
        assert_eq!(office_io_common::document_numbering_source::decode_media(&numbering_source.resources[0].data_base64).unwrap(), IMAGE);
    }

    #[test]
    fn direct_vml_sources_keep_available_bytes_and_missing_parts_remain_explicit_diagnostics() {
        let types = r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="png" ContentType="image/png"/></Types>"#;
        let xml = format!(r#"<w:numbering xmlns:w="{WORD}" xmlns:v="urn:schemas-microsoft-com:vml"><w:numPicBullet w:numPicBulletId="1"><w:pict><v:shape><v:imagedata src="media/found.png"/></v:shape></w:pict></w:numPicBullet><w:numPicBullet w:numPicBulletId="2"><w:pict><v:shape><v:imagedata src="media/missing.png"/></v:shape></w:pict></w:numPicBullet></w:numbering>"#);
        let bytes = archive_parts(&[("[Content_Types].xml", types.as_bytes()), ("word/media/found.png", IMAGE)]);
        let mut archive = ZipArchive::new(Cursor::new(bytes.as_slice())).unwrap();
        let (source, errors) = read(&mut archive, DocumentOfficeXmlPartPayload { part_name: "word/numbering.xml".into(), xml: xml.clone() });
        assert_eq!(source.xml, xml);
        assert_eq!(source.resources.len(), 1);
        assert_eq!(source.resources[0].part_name, "word/media/found.png");
        assert!(errors.iter().any(|error| error.contains("missing.png")), "{errors:?}");
        assert!(source.relationships_xml.is_none());
    }
}
