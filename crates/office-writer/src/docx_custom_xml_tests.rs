//! Deferred writer contracts, including real OPC output inspected without opening a browser.
use super::*;
use std::io::{Cursor, Read};
use serde_json::json;
use zip::ZipArchive;

fn source(name: &str, id: &str) -> Value {
    json!({ "partName": name, "xml": "<root><value>old</value></root>", "storeItemID": id })
}

fn model(parts: Value) -> Value {
    json!({ "kind": "document", "name": "xml-draft", "blocks": [{ "id": "p", "type": "paragraph", "runs": [{ "text": "old" }] }], "customXmlParts": parts })
}

fn part(bytes: &[u8], name: &str) -> String {
    let mut archive = ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut value = String::new(); archive.by_name(name).unwrap().read_to_string(&mut value).unwrap(); value
}

#[test]
fn missing_sources_remain_legacy_but_null_and_explicit_dangling_stores_are_rejected() {
    let mut value = model(json!([])); value.as_object_mut().unwrap().remove("customXmlParts");
    value["contentControls"] = json!([{ "id": "bound", "type": "text", "placement": "inline", "blockId": "p", "startOffset": 0, "endOffset": 3,
        "dataBinding": { "xpath": "/root/value", "storeItemID": "{STORE}" } }]);
    assert!(prepare(&value).is_ok());
    value["customXmlParts"] = json!([]);
    assert!(prepare(&value).err().unwrap().message.contains("MISSING_STORE"));
    value["customXmlParts"] = Value::Null;
    assert!(prepare(&value).is_err());
    value.as_object_mut().unwrap().remove("customXmlParts");
    value["contentControls"][0]["dataBinding"]["prefixMappings"] = json!("class='ordinary'");
    assert!(prepare(&value).is_err());
}

#[test]
fn incomplete_bindings_are_retained_as_dormant_configurations() {
    let mut document = model(json!([]));
    for binding in [json!({}), json!({ "xpath": "/root/value" }), json!({ "storeItemID": "{STORE}" }),
        json!({ "xpath": "", "storeItemID": "{STORE}" }), json!({ "xpath": "/root/value", "storeItemID": "" }),
        json!({ "prefixMappings": "pending namespace mapping", "dirty": false, "cachedResult": "old" })] {
        document["contentControls"] = json!([{ "id": "dormant", "type": "text", "placement": "inline", "blockId": "p", "startOffset": 0, "endOffset": 3,
            "dataBinding": binding }]);
        assert!(prepare(&document).is_ok(), "{document}");
    }
    for binding in [json!({ "xpath": null }), json!({ "storeItemID": 3 }), json!({ "dirty": "false" }), json!({ "cachedResult": false })] {
        document["contentControls"][0]["dataBinding"] = binding;
        assert!(prepare(&document).is_err(), "{document}");
    }
}

#[test]
fn source_properties_can_be_inferred_from_the_real_relationship_resource() {
    let properties = format!("<ds:datastoreItem xmlns:ds='{}' ds:itemID='{{STORE}}'/>", contract::CUSTOM_XML_NAMESPACE);
    let relationships = format!("<Relationships xmlns='{}'><Relationship Id='props' Type='{}' Target='properties.xml'/></Relationships>",
        contract::RELATIONSHIP_NAMESPACE, contract::CUSTOM_XML_PROPERTIES_RELATIONSHIP);
    let value = model(json!([{ "partName": "customXml/item.xml", "xml": "<root/>", "relationshipsXml": relationships,
        "relatedParts": [{ "partName": "customXml/properties.xml", "contentType": contract::CUSTOM_XML_PROPERTIES_CONTENT_TYPE,
            "dataBase64": contract::encode_resource(properties.as_bytes()) }] }]));
    let prepared = prepare(&value).unwrap();
    assert!(prepared.stores.contains("{store}"));
    assert_eq!(prepared.sources.len(), 1);
}

#[test]
fn graph_validation_rejects_missing_targets_owners_and_actual_store_id_conflicts() {
    let missing = format!("<Relationships xmlns='{}'><Relationship Id='missing' Type='urn:resource' Target='missing.bin'/></Relationships>", contract::RELATIONSHIP_NAMESPACE);
    let mut value = model(json!([source("customXml/a.xml", "{STORE}")]));
    value["customXmlParts"][0]["relationshipsXml"] = json!(missing);
    assert!(prepare(&value).is_err());
    value["customXmlParts"][0].as_object_mut().unwrap().remove("relationshipsXml");
    value["customXmlParts"][0]["relatedParts"] = json!([{ "partName": "customXml/_rels/missing.bin.rels", "contentType": contract::RELATIONSHIP_CONTENT_TYPE,
        "dataBase64": contract::encode_resource(empty_relationships().as_bytes()) }]);
    assert!(prepare(&value).is_err());
    value["customXmlParts"][0].as_object_mut().unwrap().remove("relatedParts");
    value["customXmlParts"][0]["propertiesPartName"] = json!("customXml/properties.xml");
    value["customXmlParts"][0]["propertiesXml"] = json!(format!("<ds:datastoreItem xmlns:ds='{}' ds:itemID='{{WRONG}}'/>", contract::CUSTOM_XML_NAMESPACE));
    assert!(prepare(&value).is_err());
}

#[test]
fn remapping_preserves_prefixed_relationships_comments_pi_and_external_urls() {
    let relationships = format!("<?xml version='1.0'?><!-- literal <!DOCTYPE remains> --><p:Relationships xmlns:p='{}'><?keep x?><p:Relationship Id='r' Type='urn:resource' Target='payload.bin'/><p:Relationship Id='external' Type='urn:external' Target='https://example.test/schema' TargetMode='External'/></p:Relationships>", contract::RELATIONSHIP_NAMESPACE);
    let value = model(json!([{ "partName": "word/source.xml", "xml": "<root/>", "storeItemID": "{STORE}", "relationshipsXml": relationships,
        "relatedParts": [{ "partName": "word/payload.bin", "contentType": "application/octet-stream", "dataBase64": "AA==" }] }]));
    let (bytes, warnings) = crate::export_docx_bytes(&value.to_string()).unwrap();
    assert!(warnings.iter().any(|warning| warning.code == "DOCX_CUSTOM_XML_PART_RELOCATED"));
    let mut archive = ZipArchive::new(Cursor::new(&bytes)).unwrap();
    let names = archive.file_names().filter(|name| name.starts_with("customXml/office-studio-data/_rels/")).map(str::to_owned).collect::<Vec<_>>();
    assert_eq!(names.len(), 1);
    let xml = part(&bytes, &names[0]);
    assert!(xml.contains("<!-- literal <!DOCTYPE remains> -->")); assert!(xml.contains("<?keep x?>"));
    assert!(xml.contains("<p:Relationship")); assert!(xml.contains("https://example.test/schema"));
    assert_eq!(contract::relationships(&xml).unwrap().len(), 3);
}

#[test]
fn date_full_date_keeps_an_existing_time_and_adds_midnight_only_to_a_date() {
    for (value, expected) in [("2026-11-25", "2026-11-25T00:00:00Z"), ("2026-11-25T10:20:30Z", "2026-11-25T10:20:30Z")] {
        let mut document = model(json!([]));
        document["contentControls"] = json!([{ "id": "date", "type": "date", "placement": "inline", "blockId": "p", "startOffset": 0, "endOffset": 3,
            "dateValue": value, "dateFormat": "yyyy-MM-dd" }]);
        let (bytes, _) = crate::export_docx_bytes(&document.to_string()).unwrap();
        let xml = part(&bytes, "word/document.xml");
        assert!(xml.contains(&format!("w:fullDate=\"{expected}\"")), "{xml}");
        assert!(!xml.contains("ZT00:00:00Z"));
    }
}
