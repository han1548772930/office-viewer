// Drafts for the final unified verification stage; not executed in this phase.
use std::io::{Cursor, Read};

use serde_json::{Value, json};

use super::validate;

const WORD: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

fn content(body: &str) -> Value {
    json!({ "blocks": [], "runs": [], "paragraphs": [], "metadata": {
        "officeStoryBlockSource": { "version": 1, "partName": "word/header1.xml" },
        "officeSource": { "version": 1, "xml": format!("<w:hdr xmlns:w=\"{WORD}\">{body}</w:hdr>"), "unsupportedFeatures": [] }
    } })
}

#[test]
fn native_tables_and_general_fields_use_block_adapters_despite_legacy_feature_flags() {
    let mut value = content("<w:tbl><w:tblPr/><w:tr><w:tc><w:p><w:fldSimple w:instr=\"DOCVARIABLE stamp\"><w:r><w:t>8</w:t></w:r></w:fldSimple></w:p></w:tc></w:tr></w:tbl>");
    value["metadata"]["officeSource"]["unsupportedFeatures"] = json!(["tables", "unknownFields", "nestedFieldCodes"]);
    assert!(validate(&value, "hdr").is_ok());
    value["metadata"].as_object_mut().unwrap().remove("officeStoryBlockSource");
    assert!(validate(&value, "hdr").unwrap_err().message.contains("nativeBlockMapping"));
}

#[test]
fn unknown_sdt_properties_and_unmapped_drawings_cannot_disappear_in_tree_export() {
    for (body, feature) in [
        ("<w:p><w:sdt><w:sdtPr><w:id w:val=\"1\"/><w:date/></w:sdtPr><w:sdtContent><w:r><w:t>Date</w:t></w:r></w:sdtContent></w:sdt></w:p>", "contentControls"),
        ("<w:p><w:r><w:drawing/></w:r></w:p>", "drawing"),
        ("<w:customXml><w:p><w:r><w:t>Bound</w:t></w:r></w:p></w:customXml>", "customXml"),
    ] {
        let value = content(body);
        let before = value.clone();
        assert!(validate(&value, "hdr").unwrap_err().message.contains(feature));
        assert_eq!(value, before, "rejection must keep the source capsule");
    }
}

#[test]
fn insertions_and_property_revisions_are_detected_even_when_flat_reader_did_not_flag_them() {
    for body in [
        "<w:p><w:ins w:id=\"2\"><w:r><w:t>Inserted</w:t></w:r></w:ins></w:p>",
        "<w:p><w:pPr><w:pPrChange w:id=\"2\"><w:pPr><w:keepNext/></w:pPr></w:pPrChange></w:pPr><w:r><w:t>Text</w:t></w:r></w:p>",
    ] { assert!(validate(&content(body), "hdr").unwrap_err().message.contains("trackedChanges")); }
}

#[test]
fn a_mapped_story_lock_is_allowed_but_a_missing_lock_is_rejected() {
    let mut value = content("<w:p><w:sdt><w:sdtPr><w:id w:val=\"7\"/><w:tag w:val=\"office_story_lock_v1_7461696c\"/><w:lock w:val=\"contentLocked\"/><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>Tail</w:t></w:r></w:sdtContent></w:sdt></w:p>");
    value["locks"] = json!([{ "id": "tail", "startOffset": 0, "endOffset": 4, "lock": "contentLocked" }]);
    assert!(validate(&value, "hdr").is_ok());
    value["locks"] = json!([]);
    assert!(validate(&value, "hdr").unwrap_err().message.contains("contentControls"));
}

#[test]
fn a_source_hyperlink_needs_a_resolved_supported_target() {
    let mut value = content("<w:p xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><w:hyperlink r:id=\"rId1\"><w:r><w:t>Link</w:t></w:r></w:hyperlink></w:p>");
    assert!(validate(&value, "hdr").unwrap_err().message.contains("unresolvedHyperlink"));
    value["metadata"]["officeSource"]["relationshipsXml"] = json!("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"https://example.test/\" TargetMode=\"External\"/></Relationships>");
    assert!(validate(&value, "hdr").is_ok());
}

fn model_with_header(content: Value) -> Value {
    json!({ "kind": "document", "name": "story-guard", "blocks": [{ "id": "body", "type": "paragraph", "runs": [{ "text": "Body" }] }],
        "sectionHeaderFooters": { "first": { "header": "", "storyContent": { "header": content } } } })
}

#[test]
fn real_docx_block_tree_export_rejects_opaque_source_instead_of_saving_a_successful_empty_header() {
    let model = model_with_header(content("<w:p><w:r><w:drawing/></w:r></w:p>"));
    let before = model.clone();
    assert!(crate::export_docx_bytes(&model.to_string()).unwrap_err().message.contains("drawing"));
    assert_eq!(model, before);
}

#[test]
fn real_docx_tree_export_writes_the_current_table_field_and_empty_tree_does_not_restore_old_text() {
    let mut header = content("<w:tbl><w:tr><w:tc><w:p><w:fldSimple w:instr=\"DOCVARIABLE stamp\"><w:r><w:t>Old</w:t></w:r></w:fldSimple></w:p></w:tc></w:tr></w:tbl>");
    header["metadata"]["officeSource"]["unsupportedFeatures"] = json!(["tables", "unknownFields"]);
    header["blocks"] = json!([{ "id": "table", "type": "table", "runs": [], "metadata": { "table": { "rows": [{ "cells": [{ "blocks": [
        { "id": "header-p", "type": "paragraph", "runs": [{ "text": "New" }], "metadata": { "documentTextFields": { "version": 1, "fields": [
            { "id": "stamp", "field": "DOCVARIABLE", "instruction": "DOCVARIABLE stamp", "startOffset": 0, "endOffset": 3 }
        ] } } }
    ] }] }] } } }]);
    let mut model = model_with_header(header);
    for empty in [false, true] {
        if empty { model["sectionHeaderFooters"]["first"]["storyContent"]["header"]["blocks"] = json!([]); }
        let (bytes, _) = crate::export_docx_bytes(&model.to_string()).unwrap();
        let mut zip = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
        let mut xml = String::new();
        zip.by_name("word/header1.xml").unwrap().read_to_string(&mut xml).unwrap();
        assert!(!xml.contains("Old"));
        assert_eq!(xml.contains("<w:tbl>"), !empty);
        assert_eq!(xml.contains("DOCVARIABLE stamp"), !empty);
        assert_eq!(xml.contains("New"), !empty);
    }
}
