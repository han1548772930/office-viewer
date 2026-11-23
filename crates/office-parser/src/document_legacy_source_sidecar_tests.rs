// Drafts only: not compiled or executed in the 2026-11-22 continuation.
use std::io::Write;

use serde_json::{Value, json};
use zip::{ZipWriter, write::SimpleFileOptions};

use super::*;

fn model(source: Option<Value>) -> Value {
    let mut model = json!({ "kind": "document", "name": "legacy source", "blocks": [{ "id": "p", "type": "paragraph", "runs": [{ "text": "Current Word content" }] }],
        "defaultStyle": { "version": 1, "text": { "fontFamily": "Georgia", "eastAsiaFontFamily": "宋体", "fontSize": 17, "color": "auto", "bold": false, "italic": false, "underline": false, "strike": false },
            "paragraph": { "align": "left", "indent": 0, "lineHeight": 1, "spacingRule": "auto", "spaceBefore": 0, "spaceAfter": 0 } } });
    if let Some(source) = source { model["metadata"] = json!({ "source": { "officeSource": source } }); }
    model
}

fn with_history() -> Vec<u8> {
    let source = json!({ "version": 1, "origin": { "format": "doc", "synthesized": true,
        "resources": [{ "kind": "binary", "streamName": "1Table", "offset": 32, "dataHex": "00ab" }], "resourceErrors": ["An original DOC property is unresolved."] } });
    office_writer::export_docx_bytes(&model(Some(source)).to_string()).unwrap().0
}

fn parts(bytes: &[u8]) -> Vec<(String, Vec<u8>)> {
    let mut archive = ZipArchive::new(Cursor::new(bytes)).unwrap();
    (0..archive.len()).map(|index| {
        let mut entry = archive.by_index(index).unwrap();
        let name = entry.name().to_owned(); let mut bytes = Vec::new(); entry.read_to_end(&mut bytes).unwrap(); (name, bytes)
    }).collect()
}

fn package(parts: Vec<(String, Vec<u8>)>) -> Vec<u8> {
    let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
    let options = SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    for (name, bytes) in parts { writer.start_file(name, options).unwrap(); writer.write_all(&bytes).unwrap(); }
    writer.finish().unwrap().into_inner()
}

fn rewrite(bytes: &[u8], mut edit: impl FnMut(&str, Vec<u8>) -> Option<Vec<u8>>) -> Vec<u8> {
    package(parts(bytes).into_iter().filter_map(|(name, bytes)| edit(&name, bytes).map(|bytes| (name, bytes))).collect())
}

fn read(bytes: &[u8]) -> Result<Vec<DocumentOfficeSourceOrigin>, String> {
    let mut archive = ZipArchive::new(Cursor::new(bytes)).unwrap(); read_legacy_origins(&mut archive)
}

fn data(bytes: &[u8]) -> Value {
    let envelope = crate::parser::parse_document_bytes(bytes, "reopened.docx").unwrap();
    serde_json::to_value(envelope).unwrap()["payload"]["data"].clone()
}

#[test]
fn registered_part_is_history_and_does_not_pollute_current_native_word_defaults() {
    let bytes = with_history();
    let history = read(&bytes).unwrap();
    assert_eq!(history.len(), 1);
    assert_eq!(history[0].resource_errors, ["An original DOC property is unresolved."]);
    let parsed = data(&bytes);
    let source = &parsed["officeSource"];
    assert!(source.get("origin").is_none());
    assert_eq!(source["legacyOrigins"][0]["format"], "doc");
    assert!(source.get("resourceErrors").is_none());
    assert!(source["styles"]["xml"].as_str().unwrap().contains("w:styles"));
    assert_eq!(parsed["defaultStyle"]["text"]["fontSizePt"], 17.0);
    assert_eq!(parsed["defaultStyle"]["text"]["fontFamilyAscii"], "Georgia");
}

#[test]
fn older_docx_without_any_sidecar_declaration_remains_compatible() {
    let bytes = office_writer::export_docx_bytes(&model(None).to_string()).unwrap().0;
    assert!(read(&bytes).unwrap().is_empty());
    let parsed = data(&bytes);
    assert!(parsed["officeSource"].get("origin").is_none());
    assert!(parsed["officeSource"].get("legacyOrigins").is_none());
    assert_eq!(parsed["defaultStyle"]["text"]["fontSizePt"], 17.0);
}

#[test]
fn package_paths_decode_once_and_reject_encoded_controls_authorities_and_root_escape() {
    assert_eq!(resolve_target("word/document.xml", "office-studio/legacy%2Dsources.json").unwrap(), LEGACY_SOURCE_PART);
    assert_eq!(resolve_target("custom/document.xml", "../word/office-studio/legacy-sources.json").unwrap(), LEGACY_SOURCE_PART);
    for target in ["%00.json", "//example.org/history.json", "%2e%2e/%2e%2e/outside.json", "https%3A//example.org/history.json", "bad%GG.json"] {
        assert!(resolve_target("word/document.xml", target).is_err());
    }
}

#[test]
fn declared_missing_wrong_type_external_and_escaping_sidecars_are_atomic_parse_errors() {
    let original = with_history();
    let missing = rewrite(&original, |name, bytes| (name != LEGACY_SOURCE_PART).then_some(bytes));
    let wrong_type = rewrite(&original, |name, bytes| Some(if name == "[Content_Types].xml" { String::from_utf8(bytes).unwrap().replace(LEGACY_SOURCE_CONTENT_TYPE, "application/json").into_bytes() } else { bytes }));
    let external = rewrite(&original, |name, bytes| Some(if name == "word/_rels/document.xml.rels" { String::from_utf8(bytes).unwrap().replace(&format!("Type=\"{LEGACY_SOURCE_RELATIONSHIP}\""), &format!("Type=\"{LEGACY_SOURCE_RELATIONSHIP}\" TargetMode=\"External\"")).into_bytes() } else { bytes }));
    let escaping = rewrite(&original, |name, bytes| Some(if name == "word/_rels/document.xml.rels" { String::from_utf8(bytes).unwrap().replace("office-studio/legacy-sources.json", "../../outside.json").into_bytes() } else { bytes }));
    for bytes in [missing, wrong_type, external, escaping] {
        assert!(read(&bytes).is_err());
        assert!(crate::parser::parse_document_bytes(&bytes, "invalid.docx").is_err());
    }
    assert_eq!(read(&original).unwrap().len(), 1);
}

#[test]
fn sidecar_schema_extension_truncation_and_global_resource_budget_are_rejected() {
    let original = with_history();
    let unknown = rewrite(&original, |name, bytes| Some(if name == LEGACY_SOURCE_PART { br#"{"version":1,"legacyOrigins":[],"unretained":true}"#.to_vec() } else { bytes }));
    let truncated = rewrite(&original, |name, bytes| Some(if name == LEGACY_SOURCE_PART { br#"{"version":1,"legacyOrigins": ["#.to_vec() } else { bytes }));
    let resources = |start: u32| (start..start + 600).map(|offset| json!({ "kind": "binary", "streamName": "1Table", "offset": offset, "dataHex": "00" })).collect::<Vec<_>>();
    let oversized = json!({ "version": 1, "legacyOrigins": [
        { "format": "doc", "synthesized": true, "resources": resources(0) },
        { "format": "doc", "synthesized": true, "resources": resources(600) },
    ] }).to_string().into_bytes();
    let oversized = rewrite(&original, |name, bytes| Some(if name == LEGACY_SOURCE_PART { oversized.clone() } else { bytes }));
    for bytes in [unknown, truncated, oversized] {
        assert!(read(&bytes).is_err());
        assert!(crate::parser::parse_document_bytes(&bytes, "invalid.docx").is_err());
    }
}

#[test]
fn duplicate_relationships_and_unqualified_contract_attribute_spoofing_are_rejected() {
    let original = with_history();
    let duplicate = rewrite(&original, |name, bytes| Some(if name == "word/_rels/document.xml.rels" {
        let text = String::from_utf8(bytes).unwrap();
        let position = text.find(LEGACY_SOURCE_RELATIONSHIP).unwrap();
        let start = text[..position].rfind("<Relationship ").unwrap();
        let end = position + text[position..].find("/>").unwrap() + 2;
        let declarations = relationships(&text).unwrap();
        let id = declarations.iter().find(|entry| entry.attribute("Type") == Some(LEGACY_SOURCE_RELATIONSHIP)).unwrap().attribute("Id").unwrap();
        let extra = text[start..end].replace(&format!("Id=\"{id}\""), "Id=\"history-copy\"");
        text.replace("</Relationships>", &format!("{extra}</Relationships>")).into_bytes()
    } else { bytes }));
    let spoof = rewrite(&original, |name, bytes| Some(if name == "word/_rels/document.xml.rels" {
        String::from_utf8(bytes).unwrap().replace("<Relationships ", "<Relationships xmlns:x=\"urn:spoof\" ")
            .replace(&format!("Type=\"{LEGACY_SOURCE_RELATIONSHIP}\""), &format!("x:Type=\"{LEGACY_SOURCE_RELATIONSHIP}\"")).into_bytes()
    } else { bytes }));
    assert!(read(&duplicate).is_err()); assert!(read(&spoof).is_err());
    let mut extra_parts = parts(&original); extra_parts.push(("word/office-studio/extra.json".into(), b"{}".to_vec()));
    assert!(read(&package(extra_parts)).is_err());
}
