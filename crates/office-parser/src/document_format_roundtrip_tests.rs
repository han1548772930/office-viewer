use std::io::{Cursor, Read, Write};

use serde_json::{Value, json};
use zip::{ZipArchive, ZipWriter, write::SimpleFileOptions};

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

fn rewrite_part(bytes: &[u8], part: &str, replacement: Option<&str>) -> Vec<u8> {
    let mut archive = ZipArchive::new(Cursor::new(bytes)).expect("DOCX package is readable");
    let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
    let mut found = false;
    for index in 0..archive.len() {
        let mut source = archive.by_index(index).expect("DOCX part is readable");
        let name = source.name().to_owned();
        if name == part { found = true; }
        if name == part && replacement.is_none() { continue; }
        let mut contents = Vec::new();
        if name == part { contents.extend_from_slice(replacement.unwrap().as_bytes()); }
        else { source.read_to_end(&mut contents).expect("DOCX part can be copied"); }
        writer.start_file(name, SimpleFileOptions::default()).expect("replacement part can be created");
        writer.write_all(&contents).expect("replacement part can be written");
    }
    if !found {
        if let Some(contents) = replacement {
            writer.start_file(part, SimpleFileOptions::default()).expect("additional part can be created");
            writer.write_all(contents.as_bytes()).expect("additional part can be written");
        }
    }
    writer.finish().expect("replacement package closes").into_inner()
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
fn consecutive_numbering_uses_updateable_fields_and_restores_one_counter_across_body_tables_and_headings() {
    let mut numbering = definition();
    numbering["abstract"]["consecutiveNumbering"] = json!(true);
    numbering["instance"]["consecutiveListId"] = json!(9);
    let mut cell = paragraph("cell-number", "Cell number", &numbering);
    cell["level"] = json!(1);
    let mut heading = paragraph("numbered-heading", "Numbered heading", &numbering);
    heading["type"] = json!("heading");
    heading["level"] = json!(3);
    heading["metadata"]["numberingLevel"] = json!(1);
    let model = json!({ "kind": "document", "name": "consecutive-numbering", "blocks": [
        paragraph("first", "First item", &numbering),
        { "id": "table", "type": "table", "metadata": { "table": { "rows": [{ "id": "row", "cells": [{ "id": "cell", "rowSpan": 1, "colSpan": 1, "blocks": [cell] }] }] } } },
        heading,
        paragraph("last", "Last item", &numbering)
    ] });
    let bytes = export(&model);
    let xml = read_part(&bytes, "word/document.xml");
    assert!(xml.contains("SEQ office_seq_1 \\r 0 \\h"));
    assert!(xml.contains("SEQ office_seq_1_l1 \\r"));
    assert!(xml.contains("w:fldCharType=\"separate\""));
    assert!(!xml.contains("<w:numPr>"));
    assert!(read_part(&bytes, "word/settings.xml").contains("<w:updateFields w:val=\"true\"/>"));
    let mapping = read_part(&bytes, "customXml/office-numbering.xml");
    assert!(mapping.contains("urn:office-studio:numbering:1"));
    assert!(read_part(&bytes, "word/_rels/document.xml.rels").contains("Target=\"../customXml/office-numbering.xml\""));

    let data = reopen(&bytes);
    let blocks = paragraphs(&data);
    let first = blocks.iter().find(|block| block["content"] == "First item").expect("sequence fields do not become body text");
    let num_id = first["details"]["numbering"]["numId"].as_i64().expect("list instance is restored");
    for (content, marker, level) in [("First item", "[0]", 0), ("Cell number", "0.1)", 1), ("Numbered heading", "0.2)", 1), ("Last item", "[3]", 0)] {
        let block = blocks.iter().find(|block| block["content"] == content).expect("paragraph remains in its container");
        assert_eq!(block["details"]["numbering"]["numId"].as_i64(), Some(num_id));
        assert_eq!(block["details"]["numbering"]["ilvl"], level);
        assert_eq!(block["details"]["numberingListId"], "shared-list");
        assert_eq!(block["details"]["listLabel"], marker);
    }
    let base = abstract_for_instance(&data, num_id);
    assert_eq!(base["consecutiveNumbering"], true);
    let instance = data["numberingDefinitions"]["instances"].as_array().unwrap().iter().find(|instance| instance["numId"].as_i64() == Some(num_id)).unwrap();
    assert_eq!(instance["consecutiveListId"].as_i64(), Some(num_id));
    let heading = blocks.iter().find(|block| block["content"] == "Numbered heading").unwrap();
    assert_eq!(heading["details"]["headingLevel"], 3);
}

#[test]
fn a_first_nested_consecutive_level_writes_and_restores_its_own_start_override() {
    let mut numbering = definition();
    numbering["abstract"]["consecutiveNumbering"] = json!(true);
    numbering["instance"]["consecutiveListId"] = json!("restarted-list");
    numbering["instance"]["levelStartOverrides"]["1"] = json!(7);
    let mut first = paragraph("nested-first", "Nested first", &numbering);
    first["level"] = json!(1);
    first["metadata"]["numberingListId"] = json!("restarted-list");
    let mut second = paragraph("root-next", "Root next", &numbering);
    second["metadata"]["numberingListId"] = json!("restarted-list");
    let bytes = export(&json!({ "kind": "document", "name": "nested-start", "blocks": [first, second] }));
    let xml = read_part(&bytes, "word/document.xml");
    assert!(xml.contains("SEQ office_seq_1 \\r 7 \\h"));
    let data = reopen(&bytes);
    let blocks = paragraphs(&data);
    assert_eq!(blocks.iter().find(|block| block["content"] == "Nested first").unwrap()["details"]["listLabel"], "0.7)");
    assert_eq!(blocks.iter().find(|block| block["content"] == "Root next").unwrap()["details"]["listLabel"], "[8]");
}

#[test]
fn native_and_consecutive_custom_tokens_keep_the_exact_pattern_and_visible_ordinal() {
    for (pattern, label) in [("001", "[003]"), ("A", "[C]"), ("α", "[γ]"), ("I", "[III]"), ("١", "[٣]")] {
        for consecutive in [false, true] {
            let mut numbering = definition();
            numbering["abstract"]["levels"][0]["format"] = json!("custom");
            numbering["abstract"]["levels"][0]["customFormat"] = json!(pattern);
            numbering["levels"][0]["format"] = json!("custom");
            numbering["levels"][0]["customFormat"] = json!(pattern);
            numbering["instance"]["levelStartOverrides"]["0"] = json!(3);
            if consecutive {
                numbering["abstract"]["consecutiveNumbering"] = json!(true);
                numbering["instance"]["consecutiveListId"] = json!(9);
            }
            let bytes = export(&json!({ "kind": "document", "name": "custom-token", "blocks": [paragraph("custom", "Custom item", &numbering)] }));
            assert!(read_part(&bytes, "word/numbering.xml").contains(&format!("w:format=\"{pattern}\"")));
            let data = reopen(&bytes);
            let block = paragraphs(&data).into_iter().find(|block| block["content"] == "Custom item").expect("custom marker is separate from the text");
            let num_id = block["details"]["numbering"]["numId"].as_i64().unwrap();
            assert_eq!(abstract_for_instance(&data, num_id)["levels"][0]["customFormat"], pattern);
            if consecutive { assert_eq!(block["details"]["listLabel"], label); }
        }
    }
}

#[test]
fn missing_or_unsupported_mapping_keeps_the_current_office_marker_and_body_once() {
    let mut numbering = definition();
    numbering["abstract"]["consecutiveNumbering"] = json!(true);
    numbering["instance"]["consecutiveListId"] = json!(9);
    let bytes = export(&json!({ "kind": "document", "name": "mapping-fallback", "blocks": [paragraph("item", "Item", &numbering)] }));
    let mapping = read_part(&bytes, "customXml/office-numbering.xml").replacen("version=\"1\"", "version=\"2\"", 1);
    for bytes in [rewrite_part(&bytes, "customXml/office-numbering.xml", None), rewrite_part(&bytes, "customXml/office-numbering.xml", Some(&mapping))] {
        let data = reopen(&bytes);
        let blocks = paragraphs(&data);
        let block = blocks.iter().find(|block| block["content"] == "[0] Item").expect("visible marker and original body are retained once");
        assert!(block["details"]["numbering"]["numId"].is_null());
        assert_eq!(block["details"]["officeNumberingMarker"]["text"], "[0] ");
        assert!(block["details"]["officeNumberingMarker"]["xml"].as_str().unwrap().contains("SEQ office_seq_1"));
    }
}

#[test]
fn incomplete_custom_formats_and_invalid_padding_fail_before_creating_a_file() {
    let mut custom = definition();
    custom["abstract"]["levels"][0]["format"] = json!("custom");
    let model = json!({ "kind": "document", "name": "unmapped-custom-format", "blocks": [paragraph("item", "Item", &custom)] });
    let error = office_writer::export_docx_bytes(&model.to_string()).expect_err("custom formats require their numbering token");
    assert_eq!(error.code, "INVALID_MODEL");
    for digits in [json!(-1), json!(1.5), json!(256)] {
        let mut numbering = definition();
        numbering["abstract"]["levels"][0]["minimumDigits"] = digits;
        let model = json!({ "kind": "document", "name": "invalid-padding", "blocks": [paragraph("item", "Item", &numbering)] });
        let error = office_writer::export_docx_bytes(&model.to_string()).expect_err("invalid padding cannot be rounded or dropped");
        assert_eq!(error.code, "INVALID_MODEL");
    }
}

fn consecutive_source() -> Vec<u8> {
    let mut numbering = definition();
    numbering["abstract"]["consecutiveNumbering"] = json!(true);
    numbering["instance"]["consecutiveListId"] = json!(9);
    export(&json!({ "kind": "document", "name": "preserved-marker-source", "blocks": [paragraph("item", "Item", &numbering)] }))
}

// Assemble the documented editable marker contract; the TypeScript normalizer is tested separately.
fn editable_preserved_paragraph(source: &Value) -> Value {
    let mut runs = Vec::<Value>::new();
    for (index, inline) in source["details"]["inlines"].as_array().expect("paragraph inlines are present").iter().enumerate() {
        if inline["type"] != "text" { continue; }
        let style = &inline["style"];
        let mut run = json!({ "id": format!("editable-run-{index}"), "text": inline["text"],
            "bold": style["bold"].as_bool().unwrap_or(false), "italic": style["italic"].as_bool().unwrap_or(false),
            "underline": style["underline"].as_bool().unwrap_or(false), "strike": style["strike"].as_bool().unwrap_or(false),
            "fontFamily": style["fontFamilyAscii"].as_str().or(style["resolvedFontFamilyAscii"].as_str()).or(style["fontFamily"].as_str()).unwrap_or("Aptos"),
            "fontSize": style["fontSizePt"].as_f64().filter(|value| *value > 0.0).unwrap_or(11.0), "color": style["color"].as_str().unwrap_or("#1f2937"),
            "metadata": { "link": inline["link"], "highlight": style["highlight"] } });
        if let Some(family) = style["fontFamilyEastAsia"].as_str().or(style["resolvedFontFamilyEastAsia"].as_str()) { run["eastAsiaFontFamily"] = json!(family); }
        if let Some(align) = style["verticalAlign"].as_str() {
            match align {
                "sup" | "super" | "superscript" => run["verticalAlign"] = json!("superscript"),
                "sub" | "subscript" => run["verticalAlign"] = json!("subscript"),
                _ => {},
            }
        }
        runs.push(run);
    }
    let mut marker = source["details"]["officeNumberingMarker"].clone();
    assert_eq!(marker["version"], 1);
    assert!(!source["details"]["numbering"]["numId"].as_i64().is_some_and(|id| id > 0), "fixture uses a preserved sequence prefix without native numbering");
    let text = marker["text"].as_str().unwrap().to_owned();
    assert!(runs.iter().filter_map(|run| run["text"].as_str()).collect::<String>().starts_with(&text));
    let mut remaining = text.len();
    let mut prefix = Vec::<Value>::new();
    for run in &runs {
        if remaining == 0 { break; }
        let raw = run["text"].as_str().unwrap();
        let length = remaining.min(raw.len());
        if length == 0 { continue; }
        let mut properties = serde_json::Map::new();
        for name in ["bold", "italic", "underline", "strike", "fontFamily", "eastAsiaFontFamily", "fontSize", "color", "verticalAlign", "href", "footnoteId", "endnoteId"] { properties.insert(name.to_owned(), run[name].clone()); }
        let mut metadata = serde_json::Map::new();
        for name in ["highlight", "verticalAlign", "href", "link", "note", "footnoteId", "endnoteId", "latexFormula"] { metadata.insert(name.to_owned(), run["metadata"][name].clone()); }
        properties.insert("metadata".to_owned(), Value::Object(metadata));
        let properties = Value::Object(properties);
        if let Some(last) = prefix.last_mut().filter(|last| last["properties"] == properties) { last["text"] = json!(format!("{}{}", last["text"].as_str().unwrap(), &raw[..length])); }
        else { prefix.push(json!({ "text": &raw[..length], "properties": properties })); }
        remaining -= length;
    }
    assert_eq!(remaining, 0);
    marker["baseline"] = json!({ "version": 1, "prefix": prefix,
        "numbering": { "kind": null, "level": null, "list": null, "checklist": null, "definition": null, "listId": null, "numbering": {} } });
    json!({ "id": "editable-marker", "type": "paragraph", "level": 0, "runs": runs, "style": {}, "metadata": { "numbering": {}, "officeNumberingMarker": marker, "source": source, "style": source["details"]["style"] } })
}

fn editable_fallback(bytes: &[u8]) -> Value {
    let data = reopen(bytes);
    let blocks = paragraphs(&data);
    assert_eq!(blocks.len(), 1);
    json!({ "kind": "document", "name": "editable-preserved-marker", "blocks": [editable_preserved_paragraph(blocks[0])], "metadata": { "source": data } })
}

const MARKER_SOURCE_STYLES: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="SimSun"/><w:sz w:val="24"/><w:lang w:val="en-US" w:eastAsia="zh-CN"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Original Normal"/><w:rPr><w:color w:val="112233"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="AmbientBase"><w:name w:val="Original Ambient Base"/><w:basedOn w:val="Normal"/><w:rPr><w:rFonts w:hAnsi="Georgia"/><w:sz w:val="28"/><w:lang w:eastAsia="ja-JP"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="SourceParagraph"><w:name w:val="Original Paragraph"/><w:basedOn w:val="AmbientBase"/><w:next w:val="Normal"/><w:link w:val="EndnoteReference"/><w:rPr><w:color w:val="224466"/></w:rPr></w:style>
  <w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="Original Marker Base"/><w:basedOn w:val="DefaultParagraphFont"/><w:rPr><w:rFonts w:ascii="Courier New"/><w:b/><w:color w:val="4477AA"/></w:rPr></w:style>
  <w:style w:type="character" w:styleId="EndnoteReference"><w:name w:val="Original Marker"/><w:basedOn w:val="FootnoteReference"/><w:link w:val="SourceParagraph"/><w:rPr><w:b/><w:sz w:val="34"/></w:rPr></w:style>
</w:styles>"#;

fn styled_marker_source(styles: &str, run_properties: &str, paragraph_style: Option<&str>) -> Vec<u8> {
    let source = consecutive_source();
    let mut document = read_part(&source, "word/document.xml");
    if let Some(style) = paragraph_style {
        assert!(document.contains("<w:pPr>"));
        document = document.replacen("<w:pPr>", &format!("<w:pPr><w:pStyle w:val=\"{style}\"/>"), 1);
    }
    let text = document.find("<w:t xml:space=\"preserve\">0</w:t>").expect("fixture has a cached numeral");
    let start = word_run_start(&document, text);
    let end = text + document[text..].find("</w:r>").unwrap() + "</w:r>".len();
    document.replace_range(start..end, &format!("<w:r><w:rPr>{run_properties}</w:rPr><w:t xml:space=\"preserve\">0</w:t></w:r>"));
    let source = rewrite_part(&source, "word/document.xml", Some(&document));
    let source = rewrite_part(&source, "word/styles.xml", Some(styles));
    rewrite_part(&source, "customXml/office-numbering.xml", None)
}

fn named_style<'a>(styles: &'a str, name: &str) -> &'a str {
    let name = styles.find(&format!("<w:name w:val=\"{name}\"")).expect("style name survives in the actual styles part");
    let start = styles[..name].rfind("<w:style ").unwrap();
    let end = name + styles[name..].find("</w:style>").unwrap() + "</w:style>".len();
    &styles[start..end]
}

fn xml_attribute<'a>(element: &'a str, name: &str) -> &'a str {
    let prefix = format!("{name}=\"");
    let start = element.find(&prefix).expect("XML attribute exists") + prefix.len();
    &element[start..start + element[start..].find('"').unwrap()]
}

fn word_run_start(document: &str, text: usize) -> usize {
    document[..text].rmatch_indices("<w:r").find(|(start, _)| {
        matches!(document.as_bytes().get(*start + "<w:r".len()), Some(b'>' | b' ' | b'\t' | b'\r' | b'\n'))
    }).map(|(start, _)| start).expect("text belongs to a Word run, including runs with namespace attributes")
}

fn cached_numeral_run(document: &str) -> &str {
    let text = document.find("<w:t xml:space=\"preserve\">0</w:t>").expect("cached numeral survives");
    let start = word_run_start(document, text);
    let end = text + document[text..].find("</w:r>").unwrap() + "</w:r>".len();
    &document[start..end]
}

#[test]
fn original_character_style_dependencies_and_ambient_properties_survive_repeated_docx_saves() {
    let source = styled_marker_source(MARKER_SOURCE_STYLES,
        "<w:rStyle w:val=\"EndnoteReference\"/><w:rFonts w:ascii=\"Verdana\"/><w:i w:val=\"1\"/><w:color w:val=\"AA5511\"/><w:lang w:bidi=\"ar-SA\"/>", Some("SourceParagraph"));
    let imported = reopen(&source);
    assert_eq!(imported["officeSource"]["version"], 1);
    assert_eq!(imported["officeSource"]["styles"]["partName"], "word/styles.xml");
    assert_eq!(imported["officeSource"]["styles"]["xml"], MARKER_SOURCE_STYLES);
    let mut model = editable_fallback(&source);
    let retained_source = model["metadata"]["source"].clone();
    for iteration in 0..3 {
        let bytes = export(&model);
        let styles = read_part(&bytes, "word/styles.xml");
        let marker = named_style(&styles, "Original Marker");
        let base = named_style(&styles, "Original Marker Base");
        let normal = named_style(&styles, "Original Normal");
        let marker_id = xml_attribute(marker, "w:styleId");
        let base_id = xml_attribute(base, "w:styleId");
        let normal_id = xml_attribute(normal, "w:styleId");
        assert_ne!(marker_id, "EndnoteReference");
        assert_ne!(base_id, "FootnoteReference");
        assert_ne!(normal_id, "Normal");
        assert!(marker.contains(&format!("<w:basedOn w:val=\"{base_id}\"/>")));
        assert!(marker.contains("<w:link w:val=\"SourceParagraph\"/>"));
        assert!(base.contains("<w:basedOn w:val=\"DefaultParagraphFont\"/>"));
        let paragraph = named_style(&styles, "Original Paragraph");
        assert!(paragraph.contains(&format!("<w:next w:val=\"{normal_id}\"/>")));
        assert!(paragraph.contains(&format!("<w:link w:val=\"{marker_id}\"/>")));
        assert!(named_style(&styles, "Original Ambient Base").contains(&format!("<w:basedOn w:val=\"{normal_id}\"/>")));
        assert!(!normal.contains("w:default="));
        for generated in ["FootnoteReference", "EndnoteReference"] {
            let generated = named_style(&styles, generated);
            assert!(generated.contains("<w:vertAlign w:val=\"superscript\"/>"));
            assert!(!generated.contains("4477AA"));
        }
        assert_eq!(styles.matches("w:default=\"1\"").count(), 1);
        assert_eq!(styles.matches("<w:docDefaults>").count(), 1);
        let document = read_part(&bytes, "word/document.xml");
        let numeral = cached_numeral_run(&document);
        for property in [format!("<w:rStyle w:val=\"{marker_id}\"/>"), "<w:b w:val=\"0\"/>".into(), "<w:i w:val=\"1\"/>".into(), "<w:sz w:val=\"34\"/>".into(), "<w:color w:val=\"AA5511\"/>".into(), "w:ascii=\"Verdana\"".into(), "w:hAnsi=\"Georgia\"".into(), "w:eastAsia=\"SimSun\"".into(), "w:val=\"en-US\"".into(), "w:eastAsia=\"ja-JP\"".into(), "w:bidi=\"ar-SA\"".into()] {
            assert!(numeral.contains(&property), "save {iteration}: {property} is present in {numeral}");
        }
        let parsed = reopen(&bytes);
        let block = paragraphs(&parsed)[0];
        assert_eq!(block["content"], "[0] Item");
        let numeral = block["details"]["inlines"].as_array().unwrap().iter().find(|inline| inline["text"] == "0").unwrap();
        assert_eq!(numeral["style"]["bold"], false);
        assert_eq!(numeral["style"]["italic"], true);
        assert_eq!(numeral["style"]["fontSizePt"].as_f64(), Some(17.0));
        assert_eq!(numeral["style"]["fontFamilyAscii"], "Verdana");
        assert_eq!(numeral["style"]["fontFamilyHAnsi"], "Georgia");
        assert_eq!(numeral["style"]["fontFamilyEastAsia"], "SimSun");
        assert_eq!(numeral["style"]["color"].as_str().unwrap().trim_start_matches('#').to_ascii_uppercase(), "AA5511");
        if iteration == 0 { assert_eq!(model["metadata"]["source"], retained_source); }
        model = editable_fallback(&bytes);
    }
}

#[test]
fn missing_invalid_and_cyclic_source_styles_keep_text_and_report_unrestored_context() {
    let source = styled_marker_source(MARKER_SOURCE_STYLES, "<w:rStyle w:val=\"EndnoteReference\"/>", Some("SourceParagraph"));
    let model = editable_fallback(&source);
    let mut cases = Vec::new();
    let mut absent = model.clone();
    absent["metadata"]["source"]["officeSource"] = Value::Null;
    cases.push(absent);
    for (version, styles) in [
        (2, MARKER_SOURCE_STYLES.to_owned()),
        (1, MARKER_SOURCE_STYLES.replacen("</w:styles>", "", 1)),
        (1, MARKER_SOURCE_STYLES.replacen("w:styleId=\"FootnoteReference\"", "w:styleId=\"MissingBase\"", 1)),
        (1, MARKER_SOURCE_STYLES.replacen("<w:basedOn w:val=\"DefaultParagraphFont\"/>", "<w:basedOn w:val=\"EndnoteReference\"/>", 1)),
        (1, MARKER_SOURCE_STYLES.replacen("<w:style w:type=\"character\" w:styleId=\"EndnoteReference\">", "<w:style xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"unretained\" w:type=\"character\" w:styleId=\"EndnoteReference\">", 1)),
    ] {
        let mut current = model.clone();
        current["metadata"]["source"]["officeSource"]["version"] = json!(version);
        current["metadata"]["source"]["officeSource"]["styles"]["xml"] = json!(styles);
        cases.push(current);
    }
    for current in cases {
        let retained = current["metadata"]["source"].clone();
        let (bytes, warnings) = office_writer::export_docx_bytes(&current.to_string()).unwrap();
        assert_eq!(warnings.len(), 1, "{warnings:?}");
        assert_eq!(warnings[0].code, "DOCX_NUMBERING_MARKER_STYLE_CONTEXT");
        assert_eq!(current["metadata"]["source"], retained);
        let document = read_part(&bytes, "word/document.xml");
        assert!(document.contains("SEQ office_seq_1 "));
        assert_eq!(reopen(&bytes)["blocks"][0]["content"], "[0] Item");
        assert!(!read_part(&bytes, "word/styles.xml").contains("Original Marker"));
    }
}

#[test]
fn absent_and_declared_empty_run_defaults_keep_their_source_size_after_docx_saves() {
    for (defaults, size) in [
        ("", 12.0),
        ("<w:docDefaults><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>", 12.0),
        ("<w:docDefaults><w:rPrDefault/></w:docDefaults>", 10.0),
        ("<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii=\"Consolas\"/></w:rPr></w:rPrDefault></w:docDefaults>", 10.0),
    ] {
        let styles = format!("<w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">{defaults}<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/></w:style><w:style w:type=\"character\" w:styleId=\"SparseMarker\"><w:name w:val=\"Sparse Marker\"/><w:basedOn w:val=\"DefaultParagraphFont\"/><w:rPr><w:color w:val=\"336699\"/></w:rPr></w:style></w:styles>");
        let source = styled_marker_source(&styles, "<w:rStyle w:val=\"SparseMarker\"/>", None);
        let parsed = reopen(&source);
        let original = paragraphs(&parsed)[0]["details"]["inlines"].as_array().unwrap().iter().find(|inline| inline["text"] == "0").unwrap();
        assert_eq!(original["style"]["fontSizePt"].as_f64(), Some(size));
        let mut model = editable_fallback(&source);
        for _ in 0..2 {
            let bytes = export(&model);
            let document = read_part(&bytes, "word/document.xml");
            assert!(cached_numeral_run(&document).contains(&format!("<w:sz w:val=\"{}\"/>", (size * 2.0) as u32)));
            assert!(read_part(&bytes, "word/styles.xml").contains("<w:sz w:val=\"22\"/>"));
            let parsed = reopen(&bytes);
            let numeral = paragraphs(&parsed)[0]["details"]["inlines"].as_array().unwrap().iter().find(|inline| inline["text"] == "0").unwrap();
            assert_eq!(numeral["style"]["fontSizePt"].as_f64(), Some(size));
            model = editable_fallback(&bytes);
        }
    }
}

const THEME_MARKER_STYLES: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="SimSun"/><w:sz w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="character" w:styleId="ThemeMarker"><w:name w:val="Theme Marker"/><w:basedOn w:val="DefaultParagraphFont"/><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="minorHAnsi" w:eastAsiaTheme="majorEastAsia"/><w:color w:val="FF0000" w:themeColor="accent1"/><w:sz w:val="30"/></w:rPr></w:style>
</w:styles>"#;

fn marker_theme_xml() -> String {
    let colors = [("dk1", "000000"), ("lt1", "FFFFFF"), ("dk2", "223344"), ("lt2", "EEEEEE"), ("accent1", "C0504D"), ("accent2", "4F81BD"), ("accent3", "9BBB59"), ("accent4", "8064A2"), ("accent5", "4BACC6"), ("accent6", "F79646"), ("hlink", "0000FF"), ("folHlink", "800080")]
        .into_iter().map(|(name, value)| format!("<a:{name}><a:srgbClr val=\"{value}\"/></a:{name}>")).collect::<String>();
    let fill = "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>";
    let lines = [6350, 12700, 19050].into_iter().map(|width| format!("<a:ln w=\"{width}\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\">{fill}<a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln>")).collect::<String>();
    format!("<?xml version=\"1.0\" encoding=\"UTF-8\"?><a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Marker Theme\"><a:themeElements><a:clrScheme name=\"Marker Colors\">{colors}</a:clrScheme><a:fontScheme name=\"Marker Fonts\"><a:majorFont><a:latin typeface=\"Georgia\"/><a:ea typeface=\"SimSun\"/><a:cs typeface=\"Arial\"/></a:majorFont><a:minorFont><a:latin typeface=\"Candara\"/><a:ea typeface=\"SimSun\"/><a:cs typeface=\"Arial\"/></a:minorFont></a:fontScheme><a:fmtScheme name=\"Marker Formats\"><a:fillStyleLst>{fills}</a:fillStyleLst><a:lnStyleLst>{lines}</a:lnStyleLst><a:effectStyleLst>{effects}</a:effectStyleLst><a:bgFillStyleLst>{fills}</a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>", fills = fill.repeat(3), effects = "<a:effectStyle><a:effectLst/></a:effectStyle>".repeat(3))
}

fn theme_marker_source() -> Vec<u8> {
    let source = styled_marker_source(THEME_MARKER_STYLES, "<w:rStyle w:val=\"ThemeMarker\"/>", None);
    let relationships = read_part(&source, "word/_rels/document.xml.rels")
        .replace("Target=\"styles.xml\"", "Target=\"../office-resources/styles/source.xml\"")
        .replace("</Relationships>", "<Relationship Id=\"rIdSourceTheme\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"../office-resources/themes/source.xml\"/></Relationships>");
    let content_types = read_part(&source, "[Content_Types].xml")
        .replace("PartName=\"/word/styles.xml\"", "PartName=\"/office-resources/styles/source.xml\"")
        .replace("</Types>", "<Override PartName=\"/office-resources/themes/source.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/></Types>");
    let source = rewrite_part(&source, "word/_rels/document.xml.rels", Some(&relationships));
    let source = rewrite_part(&source, "[Content_Types].xml", Some(&content_types));
    let source = rewrite_part(&source, "word/styles.xml", None);
    let source = rewrite_part(&source, "office-resources/styles/source.xml", Some(THEME_MARKER_STYLES));
    rewrite_part(&source, "office-resources/themes/source.xml", Some(&marker_theme_xml()))
}

#[test]
fn source_styles_and_theme_follow_internal_relationships_and_reopen_with_their_font_and_color() {
    let source = theme_marker_source();
    let parsed = reopen(&source);
    assert_eq!(parsed["officeSource"]["styles"]["partName"], "office-resources/styles/source.xml");
    assert_eq!(parsed["officeSource"]["theme"]["partName"], "office-resources/themes/source.xml");
    assert_eq!(parsed["officeSource"]["theme"]["xml"], marker_theme_xml());
    let mut model = editable_fallback(&source);
    for _ in 0..2 {
        let bytes = export(&model);
        let theme = read_part(&bytes, "word/theme/theme1.xml");
        assert!(theme.contains("Marker Theme"));
        assert!(theme.contains("typeface=\"Georgia\""));
        assert!(theme.contains("val=\"C0504D\""));
        assert!(read_part(&bytes, "word/_rels/document.xml.rels").contains("Target=\"theme/theme1.xml\""));
        assert!(read_part(&bytes, "[Content_Types].xml").contains("PartName=\"/word/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\""));
        let styles = read_part(&bytes, "word/styles.xml");
        assert!(named_style(&styles, "Theme Marker").contains("w:themeColor=\"accent1\""));
        let document = read_part(&bytes, "word/document.xml");
        let numeral = cached_numeral_run(&document);
        assert!(numeral.contains("w:asciiTheme=\"majorHAnsi\""));
        assert!(numeral.contains("w:hAnsiTheme=\"minorHAnsi\""));
        assert!(!numeral.contains("w:ascii=\"Calibri\""));
        assert!(!numeral.contains("w:hAnsi=\"Calibri\""));
        assert!(numeral.contains("w:themeColor=\"accent1\""));
        let parsed = reopen(&bytes);
        let block = paragraphs(&parsed)[0];
        assert_eq!(block["content"], "[0] Item");
        let numeral = block["details"]["inlines"].as_array().unwrap().iter().find(|inline| inline["text"] == "0").unwrap();
        assert_eq!(numeral["style"]["resolvedFontFamilyAscii"], "Georgia");
        assert_eq!(numeral["style"]["resolvedFontFamilyHAnsi"], "Candara");
        assert_eq!(numeral["style"]["resolvedFontFamilyEastAsia"], "SimSun");
        assert_eq!(numeral["style"]["fontSizePt"].as_f64(), Some(15.0));
        assert_eq!(numeral["style"]["color"].as_str().unwrap().trim_start_matches('#').to_ascii_uppercase(), "C0504D");
        assert_eq!(parsed["officeSource"]["theme"]["partName"], "word/theme/theme1.xml");
        model = editable_fallback(&bytes);
    }
}

#[test]
fn an_implicit_source_minor_theme_font_is_materialized_without_replacing_generated_defaults() {
    let mut styles = THEME_MARKER_STYLES.to_owned();
    let start = styles.find("<w:docDefaults>").unwrap();
    let end = styles.find("</w:docDefaults>").unwrap() + "</w:docDefaults>".len();
    styles.replace_range(start..end, "");
    styles = styles.replace("<w:rFonts w:asciiTheme=\"majorHAnsi\" w:hAnsiTheme=\"minorHAnsi\" w:eastAsiaTheme=\"majorEastAsia\"/>", "").replace("<w:sz w:val=\"30\"/>", "");
    let source = rewrite_part(&theme_marker_source(), "office-resources/styles/source.xml", Some(&styles));
    let mut model = editable_fallback(&source);
    for _ in 0..2 {
        let bytes = export(&model);
        let document = read_part(&bytes, "word/document.xml");
        let numeral = cached_numeral_run(&document);
        assert!(numeral.contains("w:asciiTheme=\"minorHAnsi\""));
        assert!(numeral.contains("<w:sz w:val=\"24\"/>"));
        assert!(read_part(&bytes, "word/styles.xml").contains("w:ascii=\"Aptos\""));
        let parsed = reopen(&bytes);
        let numeral = paragraphs(&parsed)[0]["details"]["inlines"].as_array().unwrap().iter().find(|inline| inline["text"] == "0").unwrap();
        assert_eq!(numeral["style"]["resolvedFontFamilyAscii"], "Candara");
        assert_eq!(numeral["style"]["fontSizePt"].as_f64(), Some(12.0));
        model = editable_fallback(&bytes);
    }
}

#[test]
fn missing_incomplete_and_unsafe_source_themes_report_context_without_writing_unsafe_resources() {
    let model = editable_fallback(&theme_marker_source());
    let mut themes = vec![Value::Null];
    for theme in [
        "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:themeElements><a:clrScheme/></a:themeElements></a:theme>".to_owned(),
        "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:themeElements><a:fontScheme/></a:themeElements></a:theme>".to_owned(),
        marker_theme_xml().replacen("<a:themeElements>", "<a:themeElements xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"unretained\">", 1),
        marker_theme_xml().replacen("<a:theme ", "<!DOCTYPE a:theme><a:theme ", 1),
        marker_theme_xml().replacen("name=\"Marker Theme\"", "name=\"Marker Theme\" name=\"Duplicate\"", 1),
    ] { themes.push(json!({ "partName": "office-resources/themes/source.xml", "xml": theme })); }
    for theme in themes {
        let mut current = model.clone();
        current["metadata"]["source"]["officeSource"]["theme"] = theme;
        let (bytes, warnings) = office_writer::export_docx_bytes(&current.to_string()).unwrap();
        assert_eq!(warnings.len(), 1, "{warnings:?}");
        assert_eq!(warnings[0].code, "DOCX_NUMBERING_MARKER_STYLE_CONTEXT");
        let mut archive = ZipArchive::new(Cursor::new(bytes.as_slice())).unwrap();
        assert!(matches!(archive.by_name("word/theme/theme1.xml"), Err(zip::result::ZipError::FileNotFound)));
        assert!(!read_part(&bytes, "word/_rels/document.xml.rels").contains("/relationships/theme"));
        assert!(read_part(&bytes, "word/document.xml").contains("SEQ office_seq_1 "));
        assert_eq!(paragraphs(&reopen(&bytes))[0]["content"], "[0] Item");
    }
}

fn comment(block_id: &str, start: usize, end: usize, message: &str) -> Value {
    json!({ "anchor": [{ "blockId": block_id, "startOffset": start, "endOffset": end }],
        "messages": [{ "id": "message", "authorName": "Author", "text": message, "createdAt": "2026-11-22T00:00:00Z" }], "resolved": false })
}

#[test]
fn missing_and_changed_mappings_replay_current_fields_once_after_the_body_is_edited() {
    let source = consecutive_source();
    let changed = read_part(&source, "word/document.xml")
        .replace("SEQ office_seq_1 \\r 0 \\h", "SEQ office_seq_1 \\r 5 \\h")
        .replacen(">0</w:t>", ">5</w:t>", 1);
    for (bytes, label) in [(rewrite_part(&source, "customXml/office-numbering.xml", None), "[0] "), (rewrite_part(&source, "word/document.xml", Some(&changed)), "[5] ")] {
        let mut model = editable_fallback(&bytes);
        model["blocks"][0]["runs"].as_array_mut().unwrap().last_mut().unwrap()["text"] = json!("Edited body");
        let saved = export(&model);
        let document = read_part(&saved, "word/document.xml");
        assert_eq!(document.matches("Edited body</w:t>").count(), 1);
        assert_eq!(document.matches("w:name=\"office_list_v1_").count(), 1);
        assert!(document.contains(if label == "[0] " { "SEQ office_seq_1 \\r 0 \\h" } else { "SEQ office_seq_1 \\r 5 \\h" }));
        let data = reopen(&saved);
        let block = paragraphs(&data).into_iter().next().unwrap();
        assert_eq!(block["content"], format!("{label}Edited body"));
        assert_eq!(block["details"]["officeNumberingMarker"]["text"], label);
    }
}

#[test]
fn editing_a_cached_label_or_its_character_style_preserves_current_text_without_replaying_stale_fields() {
    let source = consecutive_source();
    let bytes = rewrite_part(&source, "customXml/office-numbering.xml", None);
    for style_only in [false, true] {
        let mut model = editable_fallback(&bytes);
        if style_only { model["blocks"][0]["runs"][0]["color"] = json!("#123456"); }
        else { model["blocks"][0]["runs"] = json!([{ "id": "edited", "text": "[review] Item" }]); }
        let original = model["blocks"][0]["metadata"]["officeNumberingMarker"]["xml"].clone();
        let (saved, warnings) = office_writer::export_docx_bytes(&model.to_string()).unwrap();
        assert!(warnings.iter().any(|warning| warning.code == "DOCX_NUMBERING_MARKER_NOT_REPLAYED"));
        assert!(!read_part(&saved, "word/document.xml").contains("<w:fldChar"));
        assert_eq!(model["blocks"][0]["metadata"]["officeNumberingMarker"]["xml"], original);
        let data = reopen(&saved);
        assert_eq!(paragraphs(&data)[0]["content"], if style_only { "[0] Item" } else { "[review] Item" });
    }
}

#[test]
fn explicit_native_numbering_replaces_imported_fields_and_removes_only_the_unchanged_cached_label() {
    let bytes = rewrite_part(&consecutive_source(), "customXml/office-numbering.xml", None);
    let mut model = editable_fallback(&bytes);
    model["blocks"][0]["type"] = json!("numbered");
    model["blocks"][0]["list"] = json!({ "id": "replacement-native-list", "preset": "ORDER_LIST" });
    let (saved, warnings) = office_writer::export_docx_bytes(&model.to_string()).unwrap();
    assert_eq!(warnings.len(), 1);
    assert_eq!(warnings[0].code, "DOCX_NUMBERING_MARKER_REPLACED");
    let document = read_part(&saved, "word/document.xml");
    assert!(document.contains("<w:numPr>"));
    assert!(!document.contains("<w:fldChar"));
    let data = reopen(&saved);
    let block = paragraphs(&data).into_iter().next().unwrap();
    assert_eq!(block["content"], "Item");
    assert!(block["details"]["numbering"]["numId"].as_i64().is_some_and(|id| id > 0));
    assert!(block["details"]["officeNumberingMarker"].is_null());
}

#[test]
fn current_comment_ranges_replace_old_ids_inside_replayed_fields_and_reopen_at_visible_text_offsets() {
    let bytes = rewrite_part(&consecutive_source(), "customXml/office-numbering.xml", None);
    let mut old_model = editable_fallback(&bytes);
    old_model["comments"] = json!({ "old": comment("editable-marker", 2, 8, "Old comment") });
    let old = export(&old_model);
    let document = read_part(&old, "word/document.xml").replace("w:id=\"1\"", "w:id=\"77\"");
    let comments = read_part(&old, "word/comments.xml").replace("w:id=\"1\"", "w:id=\"77\"");
    let old = rewrite_part(&rewrite_part(&old, "word/document.xml", Some(&document)), "word/comments.xml", Some(&comments));
    let mut current = editable_fallback(&old);
    assert!(current["blocks"][0]["metadata"]["officeNumberingMarker"]["xml"].as_str().unwrap().contains("commentRangeStart w:id=\"77\""));
    current["comments"] = json!({ "current": comment("editable-marker", 1, 7, "Current comment") });
    let saved = export(&current);
    let document = read_part(&saved, "word/document.xml");
    assert!(!document.contains("w:id=\"77\""));
    assert_eq!(document.matches("commentRangeStart w:id=\"1\"").count(), 1);
    let data = reopen(&saved);
    assert_eq!(paragraphs(&data)[0]["content"], "[0] Item");
    let comments = data["comments"].as_array().unwrap();
    assert_eq!(comments.len(), 1);
    assert_eq!(comments[0]["text"], "Current comment");
    assert_eq!(comments[0]["targets"].as_array().unwrap().len(), 1);
    assert_eq!(comments[0]["targets"][0]["start"], 1);
    assert_eq!(comments[0]["targets"][0]["end"], 7);
}

#[test]
fn comments_started_in_a_previous_paragraph_preserve_the_next_annotated_sequence_label() {
    let mut numbering = definition();
    numbering["abstract"]["consecutiveNumbering"] = json!(true);
    numbering["instance"]["consecutiveListId"] = json!(9);
    let model = json!({ "kind": "document", "name": "cross-paragraph-comment", "blocks": [
        { "id": "lead", "type": "paragraph", "runs": [{ "text": "Lead" }] }, paragraph("item", "Item", &numbering)
    ], "comments": { "cross-paragraph": {
        "anchor": [{ "blockId": "lead", "startOffset": 0, "endOffset": 4 }, { "blockId": "item", "startOffset": 0, "endOffset": 4 }],
        "messages": [{ "authorName": "Author", "text": "Across the list label" }]
    } } });
    let bytes = export(&model);
    let envelope = serde_json::to_value(parse_office_bytes(&bytes, "cross-paragraph-comment.docx").unwrap()).unwrap();
    let data = &envelope["payload"]["data"];
    let block = paragraphs(data).into_iter().find(|block| block["content"] == "[0] Item").expect("the annotated label stays editable text");
    assert_eq!(block["details"]["officeNumberingMarker"]["text"], "[0] ");
    let targets = data["comments"][0]["targets"].as_array().unwrap();
    assert_eq!(targets.len(), 2);
    assert_eq!(targets[0]["start"], 0);
    assert_eq!(targets[0]["end"], 4);
    assert_eq!(targets[1]["start"], 0);
    assert_eq!(targets[1]["end"], 8);
    assert!(envelope["diagnostics"]["warnings"].as_array().unwrap().iter().any(|warning| warning.as_str().is_some_and(|warning| warning.contains("label intersects a comment"))));
}

#[test]
fn body_only_comments_keep_semantic_numbering_and_reopen_without_cached_label_offset_drift() {
    let mut numbering = definition();
    numbering["abstract"]["consecutiveNumbering"] = json!(true);
    numbering["instance"]["consecutiveListId"] = json!(9);
    let model = json!({ "kind": "document", "name": "body-only-comment", "blocks": [paragraph("item", "Item", &numbering)],
        "comments": { "body": comment("item", 1, 4, "Body comment") } });
    let data = reopen(&export(&model));
    assert_eq!(paragraphs(&data)[0]["content"], "Item");
    assert!(paragraphs(&data)[0]["details"]["numbering"]["numId"].as_i64().is_some_and(|id| id > 0));
    assert_eq!(data["comments"][0]["targets"][0]["start"], 1);
    assert_eq!(data["comments"][0]["targets"][0]["end"], 4);
}

#[test]
fn distinct_visible_cached_fragments_keep_their_own_styles_when_preserved_and_reexported() {
    let source = consecutive_source();
    let mut document = read_part(&source, "word/document.xml");
    let text = document.find("<w:t xml:space=\"preserve\">0</w:t>").unwrap();
    let start = word_run_start(&document, text);
    let end = text + document[text..].find("</w:r>").unwrap() + "</w:r>".len();
    document.replace_range(start..end, "<w:r><w:rPr><w:b/><w:color w:val=\"FF0000\"/><w:sz w:val=\"30\"/></w:rPr><w:t xml:space=\"preserve\">0</w:t></w:r>");
    let bytes = rewrite_part(&rewrite_part(&source, "word/document.xml", Some(&document)), "customXml/office-numbering.xml", None);
    for data in [reopen(&bytes), reopen(&export(&editable_fallback(&bytes)))] {
        let block = paragraphs(&data).into_iter().next().unwrap();
        assert_eq!(block["content"], "[0] Item");
        let inlines = block["details"]["inlines"].as_array().unwrap();
        let numeral = inlines.iter().find(|inline| inline["text"] == "0").expect("visible numeral is separate from the surrounding style");
        assert_eq!(numeral["style"]["color"].as_str().unwrap().trim_start_matches('#').to_ascii_uppercase(), "FF0000");
        assert_eq!(numeral["style"]["fontSizePt"].as_f64(), Some(15.0));
        let bracket = inlines.iter().find(|inline| inline["text"] == "[").unwrap();
        assert_eq!(bracket["style"]["fontSizePt"].as_f64(), Some(18.0));
    }
}

#[test]
fn an_empty_body_note_reference_survives_marker_replay_and_reopens_with_its_note_part() {
    let bytes = rewrite_part(&consecutive_source(), "customXml/office-numbering.xml", None);
    let mut model = editable_fallback(&bytes);
    let runs = model["blocks"][0]["runs"].as_array_mut().unwrap();
    runs.insert(runs.len() - 1, json!({ "id": "body-note", "text": "", "metadata": { "note": { "type": "footnote", "noteId": "n1" } } }));
    model["footnotes"] = json!({ "n1": { "id": "n1", "text": "Note text", "referenceBlockId": "editable-marker", "referenceRunId": "body-note" } });
    let bytes = export(&model);
    assert!(read_part(&bytes, "word/document.xml").contains("<w:footnoteReference w:id=\"1\"/>"));
    let data = reopen(&bytes);
    assert_eq!(paragraphs(&data)[0]["content"], "[0] Item");
    assert!(paragraphs(&data)[0]["details"]["inlines"].as_array().unwrap().iter().any(|inline| inline["noteReference"]["id"].as_i64() == Some(1)));
    assert_eq!(data["footnotes"][0]["text"].as_str().unwrap().trim(), "Note text");
}

#[test]
fn empty_note_references_at_body_comment_and_nested_control_boundaries_survive_replay_once() {
    let source = rewrite_part(&consecutive_source(), "customXml/office-numbering.xml", None);
    for with_controls in [false, true] {
        let mut model = editable_fallback(&source);
        let runs = model["blocks"][0]["runs"].as_array_mut().unwrap();
        runs.pop();
        runs.extend([
            json!({ "id": "start-note", "text": "", "metadata": { "note": { "type": "footnote", "noteId": "n1" } } }),
            json!({ "id": "body-first", "text": "It" }),
            json!({ "id": "middle-note", "text": "", "metadata": { "note": { "type": "footnote", "noteId": "n2" } } }),
            json!({ "id": "middle-endnote", "text": "", "metadata": { "note": { "type": "endnote", "noteId": "endnote" } } }),
            json!({ "id": "body-last", "text": "em" }),
            json!({ "id": "end-note", "text": "", "metadata": { "note": { "type": "footnote", "noteId": "n3" } } }),
        ]);
        model["footnotes"] = json!({ "n1": { "text": "Start note" }, "n2": { "text": "Middle note" }, "n3": { "text": "End note" } });
        model["endnotes"] = json!({ "endnote": { "text": "Endnote body" } });
        if with_controls {
            model["contentControls"] = json!([
                { "id": "body", "type": "richText", "blockId": "editable-marker", "startOffset": 4, "endOffset": 8 },
                { "id": "nested", "parentId": "body", "type": "richText", "blockId": "editable-marker", "startOffset": 6, "endOffset": 8 },
            ]);
        }
        model["comments"] = json!({ "body": comment("editable-marker", 6, 8, "Body note comment") });
        let bytes = export(&model);
        let document = read_part(&bytes, "word/document.xml");
        assert_eq!(document.matches("<w:footnoteReference ").count(), 3);
        assert_eq!(document.matches("<w:endnoteReference ").count(), 1);
        let start = document.find("<w:footnoteReference w:id=\"1\"/>").unwrap();
        let middle = document.find("<w:footnoteReference w:id=\"2\"/>").unwrap();
        let end = document.find("<w:footnoteReference w:id=\"3\"/>").unwrap();
        assert!(start < document.find("It</w:t>").unwrap());
        assert!(document.find("It</w:t>").unwrap() < middle && middle < document.find("em</w:t>").unwrap());
        assert!(document.find("em</w:t>").unwrap() < end);
        if with_controls {
            assert_eq!(document.matches("<w:sdt>").count(), 2);
            assert!(document.find("<w:sdt>").unwrap() < start);
            assert!(document.rfind("</w:sdt>").unwrap() < end);
        } else {
            assert!(document.find("<w:commentRangeStart w:id=\"1\"/>").unwrap() < middle);
            assert!(document.find("<w:commentRangeEnd w:id=\"1\"/>").unwrap() < end);
        }
        let data = reopen(&bytes);
        let block = paragraphs(&data).into_iter().next().unwrap();
        assert_eq!(block["content"], "[0] Item");
        assert_eq!(block["details"]["inlines"].as_array().unwrap().iter().filter(|inline| inline["noteReference"]["id"].as_i64().is_some()).count(), 4);
        assert_eq!(data["footnotes"].as_array().unwrap().len(), 3);
        assert_eq!(data["endnotes"].as_array().unwrap().len(), 1);
        assert_eq!(data["comments"][0]["targets"][0]["start"], 6);
        assert_eq!(data["comments"][0]["targets"][0]["end"], 8);
    }
}

#[test]
fn comment_and_control_slices_do_not_duplicate_a_reference_attached_to_nonempty_text() {
    for with_controls in [false, true] {
        let mut model = json!({ "kind": "document", "name": "split-note-reference", "blocks": [
            { "id": "body", "type": "paragraph", "runs": [{ "id": "run", "text": "Alpha", "metadata": { "note": { "type": "footnote", "noteId": "note" } } }] }
        ], "footnotes": { "note": { "text": "Single reference" } } });
        if with_controls {
            model["contentControls"] = json!([{ "id": "middle", "type": "richText", "blockId": "body", "startOffset": 1, "endOffset": 4 }]);
        }
        model["comments"] = json!({ "middle": comment("body", 2, 3, "Middle text") });
        let bytes = export(&model);
        assert_eq!(read_part(&bytes, "word/document.xml").matches("<w:footnoteReference ").count(), 1);
        let data = reopen(&bytes);
        let block = paragraphs(&data).into_iter().next().unwrap();
        assert_eq!(block["content"], "Alpha");
        assert_eq!(block["details"]["inlines"].as_array().unwrap().iter().filter(|inline| inline["noteReference"]["id"].as_i64() == Some(1)).count(), 1);
        assert_eq!(data["comments"][0]["targets"][0]["start"], 2);
        assert_eq!(data["comments"][0]["targets"][0]["end"], 3);
    }
}

#[test]
fn a_paragraph_containing_only_an_empty_note_reference_keeps_it_with_point_comments_or_controls() {
    for with_controls in [false, true] {
        let mut model = json!({ "kind": "document", "name": "empty-note-paragraph", "blocks": [
            { "id": "body", "type": "paragraph", "runs": [{ "id": "run", "text": "", "footnoteId": "note" }] }
        ], "footnotes": { "note": { "text": "Only note" } } });
        if with_controls {
            model["contentControls"] = json!([{ "id": "point", "type": "richText", "blockId": "body", "startOffset": 0, "endOffset": 0 }]);
        }
        model["comments"] = json!({ "point": comment("body", 0, 0, "Point comment") });
        let bytes = export(&model);
        assert_eq!(read_part(&bytes, "word/document.xml").matches("<w:footnoteReference ").count(), 1);
        let data = reopen(&bytes);
        assert_eq!(data["footnotes"][0]["text"].as_str().unwrap().trim(), "Only note");
        let block = paragraphs(&data).into_iter().next().unwrap();
        assert_eq!(block["content"], "");
        assert_eq!(block["details"]["inlines"].as_array().unwrap().iter().filter(|inline| inline["noteReference"]["id"].as_i64() == Some(1)).count(), 1);
        assert_eq!(data["comments"][0]["targets"][0]["start"], 0);
        assert_eq!(data["comments"][0]["targets"][0]["end"], 0);
    }
}

#[test]
fn collapsed_comment_order_does_not_extend_a_point_comment_over_following_text() {
    let model = json!({ "kind": "document", "name": "collapsed-comment", "blocks": [{ "id": "body", "type": "paragraph", "runs": [{ "text": "Alpha" }] }],
        "comments": { "a-point": comment("body", 0, 0, "Point comment"), "b-range": comment("body", 0, 5, "Text comment") } });
    let bytes = export(&model);
    let document = read_part(&bytes, "word/document.xml");
    let start = document.find("<w:commentRangeStart w:id=\"1\"/>").unwrap();
    let end = document.find("<w:commentRangeEnd w:id=\"1\"/>").unwrap();
    let reference = document.find("<w:commentReference w:id=\"1\"/>").unwrap();
    assert!(start < end && end < reference && reference < document.find("Alpha</w:t>").unwrap());
    let data = reopen(&bytes);
    assert_eq!(data["comments"][0]["targets"].as_array().unwrap().len(), 1);
    assert_eq!(data["comments"][0]["targets"][0]["start"], 0);
    assert_eq!(data["comments"][0]["targets"][0]["end"], 0);
    assert_eq!(data["comments"][1]["targets"][0]["start"], 0);
    assert_eq!(data["comments"][1]["targets"][0]["end"], 5);
}

fn comment_range(data: &Value, message: &str) -> (usize, usize) {
    let comment = data["comments"].as_array().unwrap().iter().find(|comment| comment["text"] == message).expect("comment message reopens");
    let targets = comment["targets"].as_array().unwrap();
    assert_eq!(targets.len(), 1, "{comment}");
    (targets[0]["start"].as_u64().unwrap() as usize, targets[0]["end"].as_u64().unwrap() as usize)
}

#[test]
fn deleted_cached_label_comments_reopen_as_points_alongside_body_controls() {
    let source = rewrite_part(&consecutive_source(), "customXml/office-numbering.xml", None);
    let mut model = editable_fallback(&source);
    model["blocks"][0]["runs"] = json!([{ "text": "Item" }]);
    model["blocks"][0]["type"] = json!("numbered");
    model["blocks"][0]["list"] = json!({ "id": "replacement-list", "preset": "ORDER_LIST" });
    model["blocks"][0]["metadata"].as_object_mut().unwrap().remove("officeNumberingMarker");
    model["comments"] = json!({ "label": comment("editable-marker", 0, 0, "Deleted label"), "body": comment("editable-marker", 0, 4, "Body comment") });
    model["contentControls"] = json!([{ "id": "body", "type": "richText", "blockId": "editable-marker", "startOffset": 0, "endOffset": 4 }]);
    let bytes = export(&model);
    let xml = read_part(&bytes, "word/document.xml");
    assert!(!xml.contains("<w:fldChar"));
    assert!(xml.contains("<w:numPr>"));
    assert_eq!(xml.matches("<w:sdt>").count(), 1);
    let data = reopen(&bytes);
    assert_eq!(paragraphs(&data)[0]["content"], "Item");
    assert_eq!(comment_range(&data, "Deleted label"), (0, 0));
    assert_eq!(comment_range(&data, "Body comment"), (0, 4));
}

#[test]
fn nested_controls_keep_inside_crossing_and_boundary_comments_in_one_paragraph() {
    for outer_type in ["richText", "plainText"] {
        let model = json!({ "kind": "document", "name": "combined-annotations", "blocks": [{ "id": "body", "type": "paragraph", "runs": [{ "text": "AlphaBeta" }] }],
            "contentControls": [
                { "id": "outer", "type": outer_type, "blockId": "body", "startOffset": 1, "endOffset": 8, "title": "Outer control" },
                { "id": "inner", "type": "richText", "parentId": "outer", "blockId": "body", "startOffset": 2, "endOffset": 6, "title": "Inner control" },
            ], "comments": {
                "inside": comment("body", 3, 5, "Inside nested control"), "enter": comment("body", 0, 4, "Across control start"),
                "exit": comment("body", 6, 9, "Across control end"), "whole": comment("body", 1, 8, "Exact outer range"),
                "point-start": comment("body", 1, 1, "Point at control start"), "point-end": comment("body", 8, 8, "Point at control end"),
            } });
        let bytes = export(&model);
        let xml = read_part(&bytes, "word/document.xml");
        assert_eq!(xml.matches("<w:sdt>").count(), 2);
        assert_eq!(xml.matches("<w:commentRangeStart ").count(), 6);
        assert_eq!(xml.matches("<w:commentRangeEnd ").count(), 6);
        assert_eq!(xml.matches("<w:commentReference ").count(), 6);
        assert!(xml.contains("w:alias w:val=\"Inner control\""));
        let envelope = serde_json::to_value(parse_office_bytes(&bytes, "combined-annotations.docx").unwrap()).unwrap();
        let data = &envelope["payload"]["data"];
        assert_eq!(paragraphs(data)[0]["content"], "AlphaBeta");
        for (message, expected) in [("Inside nested control", (3, 5)), ("Across control start", (0, 4)), ("Across control end", (6, 9)), ("Exact outer range", (1, 8)), ("Point at control start", (1, 1)), ("Point at control end", (8, 8))] {
            assert_eq!(comment_range(data, message), expected, "{outer_type}: {message}");
        }
        assert!(!envelope["diagnostics"]["warnings"].as_array().unwrap().iter().any(|warning| warning.as_str().is_some_and(|warning| warning.contains("SDT comment projection"))));
    }
}

#[test]
fn invalid_overlapping_and_unsupported_control_wrappers_preserve_current_comments() {
    for (controls, warning) in [
        (json!([{ "id": "bad-range", "type": "richText", "blockId": "body", "startOffset": 3, "endOffset": 20 }]), "DOCX_CONTENT_CONTROL_RANGE"),
        (json!([{ "id": "first", "type": "richText", "blockId": "body", "startOffset": 0, "endOffset": 4 }, { "id": "overlap", "type": "richText", "blockId": "body", "startOffset": 2, "endOffset": 5 }]), "DOCX_CONTENT_CONTROL_OVERLAP"),
        (json!([{ "id": "unknown", "type": "unsupported", "blockId": "body", "startOffset": 1, "endOffset": 4 }]), "DOCX_CONTENT_CONTROL_TYPE_UNSUPPORTED"),
    ] {
        let model = json!({ "kind": "document", "name": "control-fallback-comment", "blocks": [{ "id": "body", "type": "paragraph", "runs": [{ "text": "Alpha" }] }],
            "comments": { "range": comment("body", 1, 4, "Current range"), "point": comment("body", 2, 2, "Current point") }, "contentControls": controls });
        let (bytes, warnings) = office_writer::export_docx_bytes(&model.to_string()).unwrap();
        assert!(warnings.iter().any(|item| item.code == warning));
        assert!(!warnings.iter().any(|item| item.code == "DOCX_COMMENT_CONTENT_CONTROL_OVERLAP"));
        let data = reopen(&bytes);
        assert_eq!(paragraphs(&data)[0]["content"], "Alpha");
        assert_eq!(comment_range(&data, "Current range"), (1, 4));
        assert_eq!(comment_range(&data, "Current point"), (2, 2));
    }
}

#[test]
fn typed_control_comment_projection_keeps_cross_paragraph_and_table_ranges() {
    let model = json!({ "kind": "document", "name": "cross-control-comment", "blocks": [
        { "id": "lead", "type": "paragraph", "runs": [{ "text": "Lead" }] },
        { "id": "table", "type": "table", "metadata": { "table": { "rows": [{ "cells": [{ "blocks": [{ "id": "cell", "type": "paragraph", "runs": [{ "text": "Alpha" }] }] }] }] } } },
        { "id": "tail", "type": "paragraph", "runs": [{ "text": "Tail" }] },
    ], "contentControls": [{ "id": "typed", "type": "plainText", "blockId": "cell", "startOffset": 1, "endOffset": 4 }],
        "comments": { "cross": { "anchor": [{ "blockId": "lead", "startOffset": 1, "endOffset": 4 }, { "blockId": "cell", "startOffset": 0, "endOffset": 5 }, { "blockId": "tail", "startOffset": 0, "endOffset": 2 }], "messages": [{ "authorName": "Author", "text": "Across table control" }] } } });
    let data = reopen(&export(&model));
    let targets = data["comments"][0]["targets"].as_array().unwrap();
    assert_eq!(targets.len(), 3);
    assert_eq!(targets.iter().map(|target| (target["start"].as_u64().unwrap(), target["end"].as_u64().unwrap())).collect::<Vec<_>>(), vec![(1, 4), (0, 5), (0, 2)]);
}

fn clear_formatting_defaults() -> Value {
    json!({ "version": 1,
        "text": { "fontFamily": "Georgia", "eastAsiaFontFamily": "宋体", "fontSize": 15, "color": "auto",
            "bold": true, "italic": true, "underline": true, "strike": true, "verticalAlign": "superscript", "highlight": "#808000" },
        "paragraph": { "align": "right", "indent": 0, "indentStart": -12, "indentEnd": 9, "indentFirstLine": 4, "hanging": 18,
            "spacingRule": "exact", "lineHeight": 24, "spaceBefore": 0, "spaceAfter": 12, "wordWrap": false, "snapToGrid": true }
    })
}

#[test]
fn cleared_defaults_and_explicit_normal_overrides_survive_native_docx_reopen() {
    let defaults = clear_formatting_defaults();
    let mut cleared = defaults["text"].clone(); cleared["text"] = json!("Cleared body");
    let normal = json!({ "text": "Normal body", "fontFamily": "Arial", "fontSize": 11, "color": "auto", "bold": false, "italic": false, "underline": false, "strike": false });
    let mut custom = normal.clone(); custom["text"] = json!("Custom highlight");
    custom["highlight"] = json!("#ABCDEF"); custom["metadata"] = json!({ "highlight": "#FF0000" });
    let bytes = export(&json!({ "kind": "document", "name": "clear-formatting", "defaultStyle": defaults, "blocks": [
        { "id": "cleared", "type": "paragraph", "style": defaults["paragraph"], "runs": [cleared] },
        { "id": "normal", "type": "paragraph", "runs": [normal] },
        { "id": "custom", "type": "paragraph", "runs": [custom] },
    ] }));
    let styles = read_part(&bytes, "word/styles.xml");
    let declaration = styles.split("</w:docDefaults>").next().unwrap();
    assert!(declaration.contains("w:ascii=\"Georgia\" w:hAnsi=\"Georgia\" w:eastAsia=\"宋体\""));
    assert!(declaration.contains("<w:highlight w:val=\"darkYellow\"/>") && declaration.contains("<w:color w:val=\"auto\"/>"));
    assert!(declaration.contains("w:line=\"360\" w:lineRule=\"exact\""));
    let document = read_part(&bytes, "word/document.xml");
    assert!(document.contains("<w:b w:val=\"0\"/><w:bCs w:val=\"0\"/>") && document.contains("<w:i w:val=\"0\"/><w:iCs w:val=\"0\"/>"));
    assert!(document.contains("<w:u w:val=\"none\"/>") && document.contains("<w:strike w:val=\"0\"/>"));
    assert!(document.contains("<w:vertAlign w:val=\"baseline\"/>") && document.contains("<w:highlight w:val=\"none\"/><w:shd w:val=\"nil\"/>"));
    assert!(document.contains("w:fill=\"ABCDEF\""));
    let data = reopen(&bytes);
    let imported = &data["defaultStyle"];
    assert_eq!(imported["version"], 1);
    assert_eq!(imported["text"]["fontFamilyAscii"], "Georgia");
    assert_eq!(imported["text"]["fontFamilyEastAsia"], "宋体");
    assert_eq!(imported["text"]["fontSizePt"].as_f64(), Some(15.0));
    for mark in ["bold", "italic", "underline", "strike"] { assert_eq!(imported["text"][mark], true); }
    assert_eq!(imported["text"]["highlight"], "darkYellow");
    assert_eq!(imported["text"]["verticalAlign"], "superscript");
    assert_eq!(imported["paragraph"]["align"], "right");
    assert_eq!(imported["paragraph"]["spacing"]["lineTwips"], 360);
    assert_eq!(imported["paragraph"]["indent"]["leftTwips"], -180);
    assert_eq!(imported["paragraph"]["indent"]["hangingTwips"], 210);
    assert_eq!(imported["wordWrap"], false);
    assert_eq!(imported["paragraph"]["snapToGrid"], true);
    let blocks = paragraphs(&data);
    let normal = blocks.iter().find(|block| block["content"] == "Normal body").unwrap();
    let normal_style = &normal["details"]["inlines"][0]["style"];
    for mark in ["bold", "italic", "underline", "strike"] { assert_eq!(normal_style[mark], false); }
    assert_eq!(normal_style["verticalAlign"], "baseline");
    assert_eq!(normal_style["highlight"], "none");
    let custom = blocks.iter().find(|block| block["content"] == "Custom highlight").unwrap();
    assert_eq!(custom["details"]["inlines"][0]["style"]["backgroundColor"], "#ABCDEF");
}

#[test]
fn document_defaults_reexport_from_reader_dto_with_fonts_and_spacing_intact() {
    let defaults = clear_formatting_defaults();
    let model = json!({ "kind": "document", "name": "default-style-two-saves", "defaultStyle": defaults,
        "blocks": [{ "id": "body", "type": "paragraph", "style": defaults["paragraph"], "runs": [{ "text": "Default style survives two Office saves" }] }] });
    let first = reopen(&export(&model));
    let parsed = &first["defaultStyle"];
    let text = &parsed["text"];
    let paragraph = &parsed["paragraph"];
    let mut saved = model.clone();
    // The TypeScript normalizer has a separate contract test for the same DTO
    // mapping. This fixture verifies a second native package, not a JSON rename.
    saved["defaultStyle"] = json!({ "version": 1, "text": {
        "fontFamily": text["fontFamilyAscii"], "eastAsiaFontFamily": text["fontFamilyEastAsia"], "fontSize": text["fontSizePt"],
        "color": "auto", "bold": text["bold"], "italic": text["italic"], "underline": text["underline"], "strike": text["strike"],
        "highlight": text["highlight"], "verticalAlign": text["verticalAlign"],
    }, "paragraph": { "align": paragraph["align"], "indent": 0,
        "indentStart": paragraph["indent"]["leftTwips"].as_f64().unwrap() / 15.0,
        "indentEnd": paragraph["indent"]["rightTwips"].as_f64().unwrap() / 15.0,
        "hanging": paragraph["indent"]["hangingTwips"].as_f64().unwrap() / 15.0,
        "lineHeight": paragraph["spacing"]["lineTwips"].as_f64().unwrap() / 15.0, "spacingRule": paragraph["spacing"]["lineRule"],
        "spaceBefore": paragraph["spacing"]["beforeTwips"].as_f64().unwrap() / 15.0,
        "spaceAfter": paragraph["spacing"]["afterTwips"].as_f64().unwrap() / 15.0,
        "wordWrap": parsed["wordWrap"], "snapToGrid": paragraph["snapToGrid"],
    } });
    saved["metadata"] = json!({ "source": { "officeSource": first["officeSource"] } });
    let second = reopen(&export(&saved));
    assert_eq!(second["defaultStyle"], first["defaultStyle"]);
    assert_eq!(paragraphs(&second)[0]["content"], "Default style survives two Office saves");
}

#[test]
fn custom_default_shading_does_not_return_after_an_explicit_clear() {
    let mut defaults = clear_formatting_defaults(); defaults["text"]["highlight"] = json!("#ABCDEF");
    let data = reopen(&export(&json!({ "kind": "document", "name": "custom-default-shading", "defaultStyle": defaults,
        "blocks": [{ "id": "normal", "type": "paragraph", "runs": [{ "text": "Normal body", "fontFamily": "Arial", "fontSize": 11, "color": "auto", "bold": false, "italic": false, "underline": false, "strike": false }] }]
    })));
    assert_eq!(data["defaultStyle"]["text"]["backgroundColor"], "#ABCDEF");
    let style = &paragraphs(&data)[0]["details"]["inlines"][0]["style"];
    assert_eq!(style["highlight"], "none");
    assert!(style.get("backgroundColor").is_none(), "An explicit nil run must override the document's arbitrary RGB shading.");
}
