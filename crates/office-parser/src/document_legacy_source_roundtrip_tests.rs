// Drafts only. These exercise native parser/writer Office bytes; the helper
// assembles the published domain contract and does not invoke the TS normalizer.
use std::io::Read;

use serde_json::{Value, json};
use office_io_common::document_legacy_source::{LEGACY_SOURCE_PART, decode_origins};

use super::*;

fn text_style(raw: &Value) -> Value {
    let font = raw["fontFamilyAscii"].as_str().or_else(|| raw["resolvedFontFamilyAscii"].as_str()).or_else(|| raw["fontFamily"].as_str()).unwrap_or("Aptos");
    json!({ "fontFamily": font, "eastAsiaFontFamily": raw["fontFamilyEastAsia"].as_str().or_else(|| raw["resolvedFontFamilyEastAsia"].as_str()).unwrap_or(font),
        "fontSize": raw["fontSizePt"].as_f64().unwrap_or(12.0), "color": raw["color"].as_str().unwrap_or("auto"),
        "bold": raw["bold"] == true, "italic": raw["italic"] == true, "underline": raw["underline"] == true, "strike": raw["strike"] == true })
}

fn paragraph_style(raw: &Value) -> Value {
    let indent = &raw["indent"];
    let spacing = &raw["spacing"];
    json!({ "align": raw["align"].as_str().unwrap_or("left"), "indent": 0,
        "indentStart": indent["leftTwips"].as_f64().unwrap_or(0.0) / 15.0, "indentEnd": indent["rightTwips"].as_f64().unwrap_or(0.0) / 15.0,
        "indentFirstLine": indent["firstLineTwips"].as_f64().unwrap_or(0.0) / 15.0, "hanging": indent["hangingTwips"].as_f64().unwrap_or(0.0) / 15.0,
        "spacingRule": "auto", "lineHeight": spacing["lineTwips"].as_f64().map(|value| value / 240.0).or_else(|| spacing["lineMultiplier"].as_f64()).unwrap_or(1.0),
        "spaceBefore": spacing["beforeTwips"].as_f64().map(|value| value / 15.0).or_else(|| spacing["beforePt"].as_f64().map(|value| value * 4.0 / 3.0)).unwrap_or(0.0),
        "spaceAfter": spacing["afterTwips"].as_f64().map(|value| value / 15.0).or_else(|| spacing["afterPt"].as_f64().map(|value| value * 4.0 / 3.0)).unwrap_or(0.0) })
}

fn editable_model(payload: &DocumentPayload) -> Value {
    let data = serde_json::to_value(payload).unwrap();
    let defaults = json!({ "version": 1, "text": text_style(&data["defaultStyle"]["text"]), "paragraph": paragraph_style(&data["defaultStyle"]["paragraph"]) });
    let blocks = data["blocks"].as_array().unwrap().iter().filter(|block| block["type"] == "paragraph").enumerate().map(|(index, raw)| {
        let inlines = raw["details"]["inlines"].as_array().unwrap();
        let runs = inlines.iter().enumerate().map(|(run_index, inline)| {
            let mut run = text_style(&inline["style"]);
            run["id"] = json!(format!("body-{index}-run-{run_index}"));
            run["text"] = inline["text"].clone();
            run["metadata"] = json!({ "source": inline });
            run
        }).collect::<Vec<_>>();
        json!({ "id": format!("body-{index}"), "type": "paragraph", "level": 0,
            "runs": runs, "style": paragraph_style(&raw["details"]["style"]),
            "metadata": { "source": raw } })
    }).collect::<Vec<_>>();
    assert!(!blocks.is_empty());
    json!({ "kind": "document", "name": "legacy-resource-roundtrip", "defaultStyle": defaults, "blocks": blocks, "metadata": { "source": data } })
}

fn exported_origins(bytes: &[u8]) -> Vec<DocumentOfficeSourceOriginPayload> {
    let mut archive = ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut bytes = Vec::new(); archive.by_name(LEGACY_SOURCE_PART).unwrap().read_to_end(&mut bytes).unwrap();
    decode_origins(&bytes).unwrap()
}

fn two_docx_rounds(payload: DocumentPayload) {
    let original = payload.office_source.as_ref().unwrap().origin.as_ref().unwrap().clone();
    let expected_texts = payload.document.blocks.iter().filter(|block| !block.content.is_empty()).map(|block| block.content.clone()).collect::<Vec<_>>();
    let expected_size = payload.default_style.as_ref().unwrap().text.as_ref().unwrap().font_size_pt;
    let mut model = editable_model(&payload);
    for _ in 0..2 {
        let bytes = office_writer::export_docx_bytes(&model.to_string()).unwrap().0;
        assert_eq!(exported_origins(&bytes), [original.clone()]);
        let (reopened, capabilities) = document_payload(&bytes, "converted.docx");
        assert!(capabilities.iter().any(|name| name == "documentDefaultStyle"));
        let source = reopened.office_source.as_ref().unwrap();
        assert!(source.origin.is_none());
        assert!(source.resource_errors.is_empty());
        assert_eq!(source.legacy_origins, [original.clone()]);
        assert!(source.styles.as_ref().unwrap().xml.contains("w:styles"));
        assert_eq!(reopened.default_style.as_ref().unwrap().text.as_ref().unwrap().font_size_pt, expected_size);
        assert_eq!(reopened.document.blocks.iter().filter(|block| !block.content.is_empty()).map(|block| block.content.clone()).collect::<Vec<_>>(), expected_texts);
        model = editable_model(&reopened);
    }
}

#[test]
fn doc_style_font_binary_resources_survive_two_real_docx_exports_and_reopens_without_new_current_origin() {
    let (word, table) = doc_streams();
    let bytes = compound(&word, &table, "1Table");
    let (payload, _) = document_payload(&bytes, "origin.doc");
    let source = payload.office_source.as_ref().unwrap().origin.as_ref().unwrap();
    assert_eq!(serde_json::to_value(source).unwrap()["resources"][0]["dataHex"], encode_hex(&stsh()));
    assert_eq!(serde_json::to_value(source).unwrap()["resources"][1]["dataHex"], encode_hex(&fonts()));
    two_docx_rounds(payload);
}

#[test]
fn odt_original_xml_with_unmodelled_extensions_survives_two_real_docx_exports_and_reopens() {
    let styles = odt_styles(true);
    let content = odt_content();
    let bytes = odt_bytes(Some(&styles), Some(&content));
    let (payload, _) = document_payload(&bytes, "origin.odt");
    let source = serde_json::to_value(payload.office_source.as_ref().unwrap().origin.as_ref().unwrap()).unwrap();
    assert_eq!(source["resources"][0]["xml"], styles);
    assert_eq!(source["resources"][1]["xml"], content);
    assert!(source["resources"][0]["xml"].as_str().unwrap().contains("extra:unmodelled"));
    two_docx_rounds(payload);
}
