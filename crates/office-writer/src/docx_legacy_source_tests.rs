// Drafts only: not compiled or executed in the 2026-11-22 continuation.
use std::collections::BTreeMap;

use office_io_common::document_legacy_source::decode_origins;
use serde_json::json;

use super::prepare;
use super::super::{source_styles::SourceStyles, source_xml::{WORD_NAMESPACE, parse_xml}};

const GENERATED: &str = r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Normal"/><w:style w:type="character" w:styleId="Strong"/></w:styles>"#;

#[test]
fn writer_merges_current_origin_with_history_and_migrates_only_its_diagnostics() {
    let origin = json!({ "format": "doc", "synthesized": true, "resources": [{ "kind": "binary", "streamName": "1Table", "offset": 12, "dataHex": "00Ab" }] });
    let model = json!({ "metadata": { "source": { "officeSource": { "version": 1,
        "origin": origin, "resourceErrors": ["Original STSH property is unresolved."], "legacyOrigins": [origin] } } } });
    let bytes = prepare(&model).unwrap().unwrap();
    let origins = decode_origins(&bytes).unwrap();
    assert_eq!(origins.len(), 1);
    assert_eq!(origins[0].resource_errors, ["Original STSH property is unresolved."]);
    assert_eq!(serde_json::to_value(&origins[0]).unwrap()["resources"][0]["dataHex"], "00Ab");
    let native = json!({ "metadata": { "source": { "officeSource": { "version": 1, "resourceErrors": ["Native Word resource failed."] } } } });
    assert!(prepare(&native).unwrap().is_none());
}

#[test]
fn current_odt_does_not_materialize_synthetic_word_absence_defaults() {
    let model = json!({ "metadata": { "source": { "officeSource": { "version": 1,
        "origin": { "format": "odt", "synthesized": false, "resources": [{ "kind": "xml", "partName": "content.xml", "xml": "<odf/>" }] } } } } });
    let mut registry = SourceStyles::new(&model, GENERATED.into()).unwrap();
    let mut marker = vec![parse_xml(&format!(r#"<w:r xmlns:w="{WORD_NAMESPACE}"><w:t>1.</w:t></w:r>"#), false).unwrap()];
    registry.restore_marker(&mut marker, &BTreeMap::from([("w".into(), WORD_NAMESPACE.into())]), &json!({})).unwrap();
    assert!(marker[0].child("rPr").is_none());
    assert_eq!(registry.styles_xml().unwrap(), GENERATED);
}

#[test]
fn historical_odt_diagnostics_do_not_disable_native_word_style_resolution_and_remap() {
    let source = format!(r#"<w:styles xmlns:w="{WORD_NAMESPACE}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Georgia"/><w:sz w:val="34"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="character" w:styleId="Strong"><w:rPr><w:b/><w:caps/></w:rPr></w:style></w:styles>"#);
    let model = json!({ "metadata": { "source": { "officeSource": { "version": 1,
        "styles": { "partName": "word/styles.xml", "xml": source },
        "legacyOrigins": [{ "format": "odt", "synthesized": false, "resources": [{ "kind": "xml", "partName": "styles.xml", "xml": "<odf/>" }], "resourceErrors": ["An old ODF property is unresolved."] }] } } } });
    assert!(prepare(&model).unwrap().is_some());
    let mut registry = SourceStyles::new(&model, GENERATED.into()).unwrap();
    let mut marker = vec![parse_xml(&format!(r#"<w:r xmlns:w="{WORD_NAMESPACE}"><w:rPr><w:rStyle w:val="Strong"/></w:rPr><w:t>1.</w:t></w:r>"#), false).unwrap()];
    registry.restore_marker(&mut marker, &BTreeMap::from([("w".into(), WORD_NAMESPACE.into())]), &json!({})).unwrap();
    let properties = marker[0].child("rPr").unwrap();
    assert_eq!(properties.child("rFonts").unwrap().attribute("ascii"), Some("Georgia"));
    assert_eq!(properties.child("sz").unwrap().attribute("val"), Some("34"));
    assert!(properties.child("b").is_some() && properties.child("caps").is_some());
    let mapped = properties.child("rStyle").unwrap().attribute("val").unwrap();
    assert_ne!(mapped, "Strong");
    assert!(registry.styles_xml().unwrap().contains(&format!("w:styleId=\"{mapped}\"")));
}

#[test]
fn writer_refuses_unsupported_or_invalid_source_before_emitting_docx_bytes() {
    let base = json!({ "kind": "document", "blocks": [{ "id": "p", "type": "paragraph", "runs": [{ "text": "unchanged" }] }],
        "metadata": { "source": { "officeSource": { "version": 1, "legacyOrigins": [] } } } });
    for source in [
        json!({ "version": 1, "legacyOrigins": null }),
        json!({ "version": 1, "origin": { "format": "odt", "synthesized": false, "resources": [] }, "styles": { "partName": "styles.xml", "xml": "<odf/>" } }),
        json!({ "version": 1, "origin": { "format": "doc", "synthesized": true, "resources": [], "unretained": true } }),
        json!({ "version": 1, "unknownSourceExtension": {} }),
    ] {
        let mut model = base.clone(); model["metadata"]["source"]["officeSource"] = source;
        assert!(crate::export_docx_bytes(&model.to_string()).is_err());
        assert_eq!(model["blocks"], base["blocks"]);
    }
}
