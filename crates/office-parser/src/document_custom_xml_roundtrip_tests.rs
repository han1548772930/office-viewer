//! Deferred real-package contracts. These drafts have not been executed.
use std::io::{Cursor, Read, Write};
use serde_json::{Value, json};
use zip::{ZipArchive, ZipWriter, write::SimpleFileOptions};
use office_io_common::document_custom_xml as contract;
use crate::parser::parse_office_bytes;

fn model(parts: Value) -> Value {
    json!({ "kind": "document", "name": "custom-xml-roundtrip", "blocks": [
        { "id": "p", "type": "paragraph", "runs": [{ "text": "Stored XML" }] }
    ], "customXmlParts": parts })
}

fn export(value: &Value) -> Vec<u8> {
    office_writer::export_docx_bytes(&value.to_string()).expect("DOCX export").0
}

fn reopen(bytes: &[u8]) -> Value {
    serde_json::to_value(parse_office_bytes(bytes, "custom-xml.docx").expect("DOCX reopen")).unwrap()["payload"]["data"].clone()
}

fn stores(bytes: &[u8]) -> Result<Option<Vec<contract::DocumentCustomXmlPart>>, String> {
    crate::document_custom_xml::read(&mut ZipArchive::new(Cursor::new(bytes)).unwrap())
}

fn package_bytes(bytes: &[u8], name: &str) -> Vec<u8> {
    let mut archive = ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut value = Vec::new(); archive.by_name(name).unwrap().read_to_end(&mut value).unwrap(); value
}

fn rewrite(bytes: &[u8], name: &str, replacement: Option<&[u8]>, append: Option<(&str, &[u8])>) -> Vec<u8> {
    let mut archive = ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
    for index in 0..archive.len() {
        let mut source = archive.by_index(index).unwrap();
        let source_name = source.name().unwrap().into_owned();
        if source_name == name && replacement.is_none() { continue; }
        let mut data = Vec::new();
        if source_name == name { data.extend_from_slice(replacement.unwrap()); }
        else { source.read_to_end(&mut data).unwrap(); }
        writer.start_file(source_name, SimpleFileOptions::default()).unwrap();
        writer.write_all(&data).unwrap();
    }
    if let Some((name, data)) = append {
        writer.start_file(name, SimpleFileOptions::default()).unwrap();
        writer.write_all(data).unwrap();
    }
    writer.finish().unwrap().into_inner()
}

fn relationships(entries: &str) -> String {
    format!("<r:Relationships xmlns:r='{}'>{entries}</r:Relationships>", contract::RELATIONSHIP_NAMESPACE)
}

#[test]
fn real_sources_properties_shared_binary_and_cyclic_relationships_reopen_without_sidecar_substitution() {
    let first_rels = relationships("<r:Relationship Id='data' Type='urn:payload' Target='shared/payload.bin'/><r:Relationship Id='external' Type='urn:external' Target='https://example.test/schema' TargetMode='External'/>");
    let second_rels = relationships("<r:Relationship Id='shared' Type='urn:payload' Target='shared/payload.bin'/>");
    let payload_rels = relationships("<r:Relationship Id='cycle' Type='urn:back' Target='../item1.xml'/>");
    let document = model(json!([
        { "partName": "customXml/item1.xml", "xml": "<p:root xmlns:p='urn:person'><p:value>A &amp; 😀</p:value></p:root>", "storeItemID": "{FIRST}", "relationshipsXml": first_rels,
          "relatedParts": [
              { "partName": "customXml/shared/payload.bin", "contentType": "application/octet-stream", "dataBase64": "AAE=" },
              { "partName": "customXml/shared/_rels/payload.bin.rels", "contentType": contract::RELATIONSHIP_CONTENT_TYPE,
                "dataBase64": contract::encode_resource(payload_rels.as_bytes()) }
          ] },
        { "partName": "customXml/item2.xml", "xml": "<second><value>B</value></second>", "storeItemID": "{SECOND}", "relationshipsXml": second_rels }
    ]));
    let bytes = export(&document);
    let reopened = reopen(&bytes);
    let parts = reopened["customXmlParts"].as_array().unwrap();
    assert_eq!(parts.len(), 2);
    assert_eq!(parts[0]["storeItemID"], "{FIRST}"); assert_eq!(parts[1]["storeItemID"], "{SECOND}");
    assert_eq!(parts[0]["xml"], document["customXmlParts"][0]["xml"]);
    let resources = parts.iter().flat_map(|part| part["relatedParts"].as_array().into_iter().flatten()).collect::<Vec<_>>();
    assert_eq!(resources.iter().filter(|part| part["partName"] == "customXml/shared/payload.bin").count(), 1);
    assert_eq!(resources.iter().filter(|part| part["partName"] == "customXml/shared/_rels/payload.bin.rels").count(), 1);
    assert_eq!(package_bytes(&bytes, "customXml/shared/payload.bin"), vec![0, 1]);
    let source_edges = contract::relationships(parts[0]["relationshipsXml"].as_str().unwrap()).unwrap();
    assert!(source_edges.iter().any(|edge| edge.external && edge.target == "https://example.test/schema"));
    assert!(source_edges.iter().any(|edge| contract::relationship_is(&edge.kind, "customXmlProps")));
    let mut again = document.clone(); again["customXmlParts"] = reopened["customXmlParts"].clone();
    let second = reopen(&export(&again));
    assert_eq!(second["customXmlParts"], reopened["customXmlParts"]);
}

#[test]
fn utf16_bom_and_literal_zip_names_keep_the_declared_uri_encoding_and_unicode() {
    let document = model(json!([{ "partName": "customXml/with%20space.xml", "xml": "<?xml version='1.0' encoding='UTF-16'?><root>中😀</root>",
        "storeItemID": "{UNICODE}", "encoding": "utf16le", "byteOrderMark": true }]));
    let bytes = export(&document);
    assert!(package_bytes(&bytes, "customXml/with space.xml").starts_with(&[0xff, 0xfe]));
    let parts = stores(&bytes).unwrap().unwrap();
    assert_eq!(parts[0].part_name, "customXml/with%20space.xml");
    assert_eq!(parts[0].encoding, Some(contract::CustomXmlEncoding::Utf16le));
    assert_eq!(parts[0].byte_order_mark, Some(true)); assert!(parts[0].xml.contains("中😀"));
}

#[test]
fn missing_graph_resources_and_uri_aliased_zip_entries_fail_instead_of_dropping_sources() {
    let document = model(json!([{ "partName": "customXml/with%20space.xml", "xml": "<root/>", "storeItemID": "{STORE}",
        "relationshipsXml": relationships("<r:Relationship Id='data' Type='urn:payload' Target='payload.bin'/>"),
        "relatedParts": [{ "partName": "customXml/payload.bin", "contentType": "application/octet-stream", "dataBase64": "AAE=" }] }]));
    let bytes = export(&document);
    let missing = rewrite(&bytes, "customXml/payload.bin", None, None);
    assert!(stores(&missing).unwrap_err().contains("MISSING_PART"));
    let aliased = rewrite(&bytes, "", None, Some(("customXml/with%20space.xml", b"<root/>")));
    assert!(stores(&aliased).unwrap_err().contains("DUPLICATE_PART"));
}

#[test]
fn actual_relationships_mime_and_properties_id_are_checked_on_reopen() {
    let bytes = export(&model(json!([{ "partName": "customXml/item.xml", "xml": "<root/>", "storeItemID": "{STORE}" }])));
    let types = String::from_utf8(package_bytes(&bytes, "[Content_Types].xml")).unwrap();
    let invalid_types = types.replace(contract::RELATIONSHIP_CONTENT_TYPE, "application/octet-stream");
    assert!(stores(&rewrite(&bytes, "[Content_Types].xml", Some(invalid_types.as_bytes()), None)).is_err());
    let parts = stores(&bytes).unwrap().unwrap();
    let props = contract::zip_part_name(parts[0].properties_part_name.as_ref().unwrap()).unwrap();
    let invalid = format!("<ds:datastoreItem xmlns:ds='{}'/>", contract::CUSTOM_XML_NAMESPACE);
    assert!(stores(&rewrite(&bytes, &props, Some(invalid.as_bytes()), None)).unwrap_err().contains("INVALID_PROPERTIES"));
}

#[test]
fn a_business_xml_root_is_retained_even_when_its_requested_name_collides_with_an_editor_sidecar() {
    let document = model(json!([{ "partName": "customXml/office-numbering.xml", "xml": "<officeNumbering xmlns='urn:business'><value>keep</value></officeNumbering>",
        "storeItemID": "{BUSINESS}" }]));
    let bytes = export(&document);
    let parts = stores(&bytes).unwrap().unwrap();
    assert_eq!(parts.len(), 1); assert_eq!(parts[0].store_item_id.as_deref(), Some("{BUSINESS}"));
    assert!(parts[0].xml.contains("urn:business"));
    assert!(parts[0].part_name.starts_with("customXml/office-studio-data/"));
}
