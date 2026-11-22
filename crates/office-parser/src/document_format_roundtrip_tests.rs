use std::io::{Cursor, Read};

use serde_json::{Value, json};
use zip::ZipArchive;

use crate::parser::parse_office_bytes;

fn definition() -> Value {
    let levels = json!([
        {
            "ilvl": 0, "start": 0, "format": "decimal", "text": "%1.", "suffix": "space", "lvlRestart": 0,
            "indent": { "leftTwips": 600, "rightTwips": -150, "firstLineTwips": 150, "hangingTwips": 450 },
            "runStyle": { "fontFamilyAscii": "Georgia", "fontSizePt": 18, "bold": true, "italic": true, "color": "auto" }
        },
        { "ilvl": 1, "start": 3, "format": "upperRoman", "text": "%1.%2)", "suffix": "tab", "lvlRestart": 1, "isLgl": true }
    ]);
    let mut effective = levels.clone();
    effective[0]["text"] = json!("[%1]");
    json!({
        "abstract": { "abstractNumId": 7, "levels": levels },
        "instance": { "numId": 9, "abstractNumId": 7, "levelStartOverrides": { "0": 0 }, "levelOverrides": [{ "ilvl": 0, "text": "[%1]" }] },
        "levels": effective
    })
}

fn paragraph(id: &str, text: &str, numbering: &Value) -> Value {
    json!({
        "id": id, "type": "numbered", "level": 0,
        "runs": [{ "id": format!("{id}-run"), "text": text }], "style": {},
        "metadata": { "numberingListId": "shared-list", "numberingDefinition": numbering }
    })
}

fn export(model: &Value) -> Vec<u8> {
    let (bytes, warnings) = office_writer::export_docx_bytes(&model.to_string()).expect("DOCX export succeeds");
    assert!(warnings.is_empty(), "{warnings:?}");
    bytes
}

fn reopen(bytes: &[u8]) -> Value {
    let envelope = parse_office_bytes(bytes, "paragraph-format.docx").expect("exported DOCX reopens");
    serde_json::to_value(envelope).expect("parse result serializes")["payload"]["data"].clone()
}

fn read_part(bytes: &[u8], part: &str) -> String {
    let mut archive = ZipArchive::new(Cursor::new(bytes)).expect("DOCX package is readable");
    let mut text = String::new();
    archive.by_name(part).expect("DOCX part exists").read_to_string(&mut text).expect("DOCX part is UTF-8 XML");
    text
}

fn paragraphs(data: &Value) -> Vec<&Value> {
    fn visit<'a>(blocks: &'a [Value], result: &mut Vec<&'a Value>) {
        for block in blocks {
            if block["details"]["kind"].as_str() == Some("paragraph") { result.push(block); }
            if let Some(children) = block["children"].as_array() { visit(children, result); }
        }
    }
    let mut result = Vec::new();
    visit(data["blocks"].as_array().expect("document exposes blocks"), &mut result);
    result
}

fn abstract_for_instance(data: &Value, num_id: i64) -> &Value {
    let definitions = &data["numberingDefinitions"];
    let instance = definitions["instances"].as_array().unwrap().iter().find(|value| value["numId"].as_i64() == Some(num_id)).expect("numbering instance reopens");
    definitions["abstracts"].as_array().unwrap().iter().find(|value| value["abstractNumId"] == instance["abstractNumId"]).expect("instance keeps its abstract definition")
}

#[test]
fn custom_numbering_reopens_zero_start_full_overrides_and_one_instance_across_body_and_table() {
    let numbering = definition();
    let mut heading = paragraph("heading", "Numbered heading", &numbering);
    heading["type"] = json!("heading");
    heading["level"] = json!(3);
    heading["metadata"]["numberingLevel"] = json!(1);
    let model = json!({ "kind": "document", "name": "custom-numbering", "blocks": [
        paragraph("first", "First item", &numbering),
        { "id": "table", "type": "table", "metadata": { "table": { "rows": [{ "id": "row", "cells": [{
            "id": "cell", "rowSpan": 1, "colSpan": 1, "blocks": [paragraph("in-cell", "Cell item", &numbering)]
        }] }] } } },
        heading,
        paragraph("last", "Last item", &numbering)
    ] });
    let bytes = export(&model);
    let data = reopen(&bytes);
    let blocks = paragraphs(&data);
    let first = blocks.iter().find(|block| block["content"] == "First item").unwrap();
    let num_id = first["details"]["numbering"]["numId"].as_i64().unwrap();
    for content in ["First item", "Cell item", "Numbered heading", "Last item"] {
        let block = blocks.iter().find(|block| block["content"] == content).expect("numbered paragraph survives in its container");
        assert_eq!(block["details"]["numbering"]["numId"].as_i64(), Some(num_id));
    }
    let heading = blocks.iter().find(|block| block["content"] == "Numbered heading").unwrap();
    assert_eq!(heading["type"], "heading");
    assert_eq!(heading["details"]["headingLevel"], 3);
    assert_eq!(heading["details"]["numbering"]["ilvl"], 1);
    let base = abstract_for_instance(&data, num_id);
    assert_eq!(base["levels"][0]["start"], 0);
    assert_eq!(base["levels"][0]["lvlRestart"], 0);
    assert_eq!(base["levels"][1]["lvlRestart"], 1);
    assert_eq!(base["levels"][1]["isLgl"], true);
    let instance = data["numberingDefinitions"]["instances"].as_array().unwrap().iter().find(|value| value["numId"].as_i64() == Some(num_id)).unwrap();
    assert_eq!(instance["levelStartOverrides"]["0"], 0);
    let level = &instance["levelOverrides"][0];
    assert_eq!(level["text"], "[%1]");
    assert_eq!(level["format"], "decimal");
    assert_eq!(level["suffix"], "space");
    assert_eq!(level["indent"]["leftTwips"], 600);
    assert_eq!(level["indent"]["rightTwips"], -150);
    assert_eq!(level["indent"]["hangingTwips"], 300);
    assert_eq!(level["runStyle"]["fontFamilyAscii"], "Georgia");
    assert_eq!(level["runStyle"]["fontSizePt"].as_f64(), Some(18.0));
    assert_eq!(level["runStyle"]["bold"], true);
    assert_eq!(level["runStyle"]["italic"], true);
    assert_eq!(level["runStyle"]["color"], "auto");
    assert_eq!(level["bulletColor"], "auto");
}

#[test]
fn decimal_padding_uses_native_custom_formats_and_reopens_base_and_override_widths() {
    let mut numbering = definition();
    numbering["abstract"]["levels"][0]["minimumDigits"] = json!(3);
    numbering["instance"]["levelOverrides"][0]["minimumDigits"] = json!(4);
    numbering["levels"][0]["minimumDigits"] = json!(4);
    let bytes = export(&json!({ "kind": "document", "name": "padded-numbering", "blocks": [paragraph("padded", "Padded item", &numbering)] }));
    let xml = read_part(&bytes, "word/numbering.xml");
    assert!(xml.contains("<w:numFmt w:val=\"decimal\" w:format=\"001\"/>"));
    assert!(xml.contains("<w:numFmt w:val=\"decimal\" w:format=\"0001\"/>"));
    let data = reopen(&bytes);
    let block = paragraphs(&data).into_iter().find(|block| block["content"] == "Padded item").unwrap();
    let num_id = block["details"]["numbering"]["numId"].as_i64().unwrap();
    assert_eq!(abstract_for_instance(&data, num_id)["levels"][0]["minimumDigits"], 3);
    let instance = data["numberingDefinitions"]["instances"].as_array().unwrap().iter().find(|value| value["numId"].as_i64() == Some(num_id)).unwrap();
    assert_eq!(instance["levelOverrides"][0]["minimumDigits"], 4);
}

#[test]
fn paragraph_spacing_indent_wrap_grid_and_automatic_text_color_survive_docx_reopen() {
    let cases = [("Auto", "auto", 0.75, 180, false, true), ("Exact", "exact", 24.0, 360, true, false), ("Minimum", "atLeast", 20.0, 300, false, false)];
    let blocks = cases.iter().map(|(name, rule, height, _, wrap, grid)| json!({
        "id": name, "type": "paragraph",
        "style": { "spacingRule": rule, "lineHeight": height, "spaceBefore": 0, "spaceAfter": 0,
            "indentStart": -12, "indentEnd": 9, "indentFirstLine": 4, "hanging": 18, "wordWrap": wrap, "snapToGrid": grid },
        "runs": [{ "text": format!("{name} & 😀"), "color": "auto", "bold": true }, { "text": " red", "color": "#FF0000" }]
    })).collect::<Vec<_>>();
    let bytes = export(&json!({ "kind": "document", "name": "paragraph-properties", "blocks": blocks }));
    let xml = read_part(&bytes, "word/document.xml");
    assert!(xml.contains("w:line=\"180\" w:lineRule=\"auto\""));
    assert!(xml.contains("w:line=\"360\" w:lineRule=\"exact\""));
    assert!(xml.contains("w:line=\"300\" w:lineRule=\"atLeast\""));
    assert!(xml.contains("<w:ind w:left=\"-180\" w:right=\"135\" w:hanging=\"210\"/>"));
    let data = reopen(&bytes);
    let blocks = paragraphs(&data);
    for (name, rule, _, line, wrap, grid) in cases {
        let content = format!("{name} & 😀 red");
        let block = blocks.iter().find(|block| block["content"].as_str() == Some(content.as_str())).unwrap();
        let style = &block["details"]["style"];
        assert_eq!(style["spacing"]["lineRule"], rule);
        assert_eq!(style["spacing"]["lineTwips"], line);
        assert_eq!(style["spacing"]["beforeTwips"], 0);
        assert_eq!(style["spacing"]["afterTwips"], 0);
        assert_eq!(style["indent"]["leftTwips"], -180);
        assert_eq!(style["indent"]["rightTwips"], 135);
        assert_eq!(style["indent"]["hangingTwips"], 210);
        assert_eq!(style["wordWrap"], wrap);
        assert_eq!(style["snapToGrid"], grid);
        let inlines = block["details"]["inlines"].as_array().unwrap();
        assert_eq!(inlines[0]["style"]["color"], "auto");
        assert_eq!(inlines[0]["style"]["bold"], true);
        assert_eq!(inlines.last().unwrap()["style"]["color"], "#FF0000");
    }
}

#[test]
fn picture_numbering_writes_its_shape_media_and_relationship_and_reopens_the_image() {
    const IMAGE: &str = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=";
    let levels = json!([{ "ilvl": 0, "start": 1, "format": "bullet", "text": "", "suffix": "tab", "pictureBulletId": 5,
        "pictureBullet": { "numPicBulletId": 5, "src": IMAGE, "widthPx": 12, "heightPx": 9 } }]);
    let numbering = json!({ "abstract": { "abstractNumId": 7, "levels": levels }, "instance": { "numId": 9, "abstractNumId": 7 }, "levels": levels });
    let bytes = export(&json!({ "kind": "document", "name": "picture-numbering", "blocks": [paragraph("picture", "Picture item", &numbering)] }));
    let xml = read_part(&bytes, "word/numbering.xml");
    assert!(xml.contains("<v:shapetype id=\"_x0000_t75\""));
    assert!(xml.contains("type=\"#_x0000_t75\" style=\"width:12px;height:9px\""));
    assert!(xml.contains("<w:lvlPicBulletId w:val=\"1\"/>"));
    let relationships = read_part(&bytes, "word/_rels/numbering.xml.rels");
    assert!(relationships.contains("Target=\"media/numbering-picture-1.png\""));
    let mut archive = ZipArchive::new(Cursor::new(bytes.as_slice())).unwrap();
    let mut media = Vec::new();
    archive.by_name("word/media/numbering-picture-1.png").unwrap().read_to_end(&mut media).unwrap();
    assert!(media.starts_with(b"\x89PNG\r\n\x1a\n"));
    let data = reopen(&bytes);
    let block = paragraphs(&data).into_iter().find(|block| block["content"] == "Picture item").unwrap();
    let num_id = block["details"]["numbering"]["numId"].as_i64().unwrap();
    let picture = &abstract_for_instance(&data, num_id)["levels"][0]["pictureBullet"];
    assert_eq!(picture["src"], IMAGE);
    assert_eq!(picture["widthPx"].as_f64(), Some(12.0));
    assert_eq!(picture["heightPx"].as_f64(), Some(9.0));
}

#[test]
fn unsupported_consecutive_numbering_and_invalid_padding_fail_before_creating_a_file() {
    let mut numbering = definition();
    numbering["abstract"]["consecutiveNumbering"] = json!(true);
    numbering["instance"]["consecutiveListId"] = json!(9);
    let model = json!({ "kind": "document", "name": "consecutive-numbering", "blocks": [paragraph("item", "Item", &numbering)] });
    let error = office_writer::export_docx_bytes(&model.to_string()).expect_err("consecutive numbering cannot be silently downgraded");
    assert_eq!(error.code, "UNSUPPORTED_EXPORT");
    let mut custom = definition();
    custom["abstract"]["levels"][0]["format"] = json!("custom");
    let model = json!({ "kind": "document", "name": "unmapped-custom-format", "blocks": [paragraph("item", "Item", &custom)] });
    let error = office_writer::export_docx_bytes(&model.to_string()).expect_err("unmapped custom formats cannot be silently downgraded");
    assert_eq!(error.code, "UNSUPPORTED_EXPORT");
    for digits in [json!(-1), json!(1.5), json!(256)] {
        let mut numbering = definition();
        numbering["abstract"]["levels"][0]["minimumDigits"] = digits;
        let model = json!({ "kind": "document", "name": "invalid-padding", "blocks": [paragraph("item", "Item", &numbering)] });
        let error = office_writer::export_docx_bytes(&model.to_string()).expect_err("invalid padding cannot be rounded or dropped");
        assert_eq!(error.code, "INVALID_MODEL");
    }
}
