use std::io::{Cursor, Read, Write};

use serde_json::{Value, json};
use zip::{ZipArchive, ZipWriter, write::SimpleFileOptions};

use crate::parser::parse_office_bytes;

// Native DOC-02/DOC-03 interoperability drafts. No test, build or Office validation
// was run in this turn. The TypeScript DTO mapping has separate contracts.
const STYLES: &str = r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
 <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia" w:eastAsia="宋体"/><w:sz w:val="30"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:line="288" w:after="180"/></w:pPr></w:pPrDefault></w:docDefaults>
 <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Original Normal"/></w:style>
 <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="Original Story Paragraph"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:tabs><w:tab w:pos="720" w:val="right"/></w:tabs><w:jc w:val="center"/></w:pPr></w:style>
 <w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="Original Story Character"/><w:rPr><w:caps/><w:smallCaps/><w:spacing w:val="24"/><w:lang w:val="fr-FR"/><w:color w:val="FF0000" w:themeColor="accent1"/></w:rPr></w:style>
</w:styles>"#;
const HEADER: &str = r#"<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
 <w:p><w:pPr><w:pStyle w:val="Heading1"/><w:pBdr><w:bottom w:val="single" w:sz="4" w:color="112233"/></w:pBdr></w:pPr>
  <w:hyperlink r:id="rIdLink"><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/><w:rFonts w:ascii="Georgia"/><w:sz w:val="36"/><w:b w:val="0"/></w:rPr><w:t>A😀</w:t></w:r></w:hyperlink>
  <w:fldSimple w:instr=" DATE \@ &quot;yyyy/MM/dd&quot; "><w:r><w:rPr><w:rFonts w:ascii="Georgia"/><w:sz w:val="36"/><w:color w:val="FF0000"/><w:b w:val="0"/></w:rPr><w:t>2026/</w:t></w:r><w:hyperlink r:id="rIdLink"><w:r><w:rPr><w:rFonts w:ascii="Aptos"/><w:sz w:val="24"/><w:color w:val="0000FF"/><w:b/><w:u w:val="single"/></w:rPr><w:t>11/22</w:t></w:r></w:hyperlink></w:fldSimple>
  <w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve"> word</w:t></w:r>
 </w:p>
 <w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:sdt><w:sdtPr><w:id w:val="77"/><w:tag w:val="office_story_lock_v1_7461696c"/><w:lock w:val="contentLocked"/><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>Tail</w:t></w:r></w:sdtContent></w:sdt></w:p>
</w:hdr>"#;

fn export(model: &Value) -> Vec<u8> {
    let (bytes, warnings) = office_writer::export_docx_bytes(&model.to_string()).expect("story DOCX can be exported");
    assert!(warnings.is_empty(), "{warnings:?}");
    bytes
}

fn reopen(bytes: &[u8]) -> Value {
    let envelope = parse_office_bytes(bytes, "story-format.docx").expect("story DOCX can be reopened");
    serde_json::to_value(envelope).unwrap()["payload"]["data"].clone()
}

fn read_part(bytes: &[u8], part: &str) -> String {
    let mut archive = ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut text = String::new();
    archive.by_name(part).unwrap().read_to_string(&mut text).unwrap();
    text
}

fn rewrite_part(bytes: &[u8], part: &str, replacement: Option<&str>) -> Vec<u8> {
    let mut archive = ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
    let mut found = false;
    for index in 0..archive.len() {
        let mut source = archive.by_index(index).unwrap();
        let name = source.name().to_owned();
        if name == part { found = true; }
        if name == part && replacement.is_none() { continue; }
        let mut contents = Vec::new();
        if name == part { contents.extend_from_slice(replacement.unwrap().as_bytes()); }
        else { source.read_to_end(&mut contents).unwrap(); }
        writer.start_file(name, SimpleFileOptions::default()).unwrap();
        writer.write_all(&contents).unwrap();
    }
    if !found {
        if let Some(contents) = replacement {
            writer.start_file(part, SimpleFileOptions::default()).unwrap();
            writer.write_all(contents.as_bytes()).unwrap();
        }
    }
    writer.finish().unwrap().into_inner()
}

fn source_theme() -> String {
    let colors = [("dk1", "000000"), ("lt1", "FFFFFF"), ("dk2", "223344"), ("lt2", "EEEEEE"), ("accent1", "C0504D"), ("accent2", "4F81BD"), ("accent3", "9BBB59"), ("accent4", "8064A2"), ("accent5", "4BACC6"), ("accent6", "F79646"), ("hlink", "0000FF"), ("folHlink", "800080")]
        .into_iter().map(|(name, value)| format!("<a:{name}><a:srgbClr val=\"{value}\"/></a:{name}>")).collect::<String>();
    let fill = "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>";
    let lines = [6350, 12700, 19050].into_iter().map(|width| format!("<a:ln w=\"{width}\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\">{fill}<a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln>")).collect::<String>();
    format!("<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Story Theme\"><a:themeElements><a:clrScheme name=\"Story Colors\">{colors}</a:clrScheme><a:fontScheme name=\"Story Fonts\"><a:majorFont><a:latin typeface=\"Georgia\"/><a:ea typeface=\"宋体\"/><a:cs typeface=\"Arial\"/></a:majorFont><a:minorFont><a:latin typeface=\"Aptos\"/><a:ea typeface=\"宋体\"/><a:cs typeface=\"Arial\"/></a:minorFont></a:fontScheme><a:fmtScheme name=\"Story Formats\"><a:fillStyleLst>{fills}</a:fillStyleLst><a:lnStyleLst>{lines}</a:lnStyleLst><a:effectStyleLst>{effects}</a:effectStyleLst><a:bgFillStyleLst>{fills}</a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>", fills = fill.repeat(3), effects = "<a:effectStyle><a:effectLst/></a:effectStyle>".repeat(3))
}

fn source_package() -> Vec<u8> {
    let seed = export(&json!({ "kind": "document", "name": "source-story", "blocks": [{ "id": "body", "type": "paragraph", "runs": [{ "text": "body unchanged" }] }],
        "sectionHeaderFooters": { "first": { "header": "seed", "footer": "", "pageFields": [] } } }));
    let relations = read_part(&seed, "word/_rels/document.xml.rels").replace("</Relationships>", "<Relationship Id=\"rIdSourceTheme\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/></Relationships>");
    let types = read_part(&seed, "[Content_Types].xml").replace("</Types>", "<Override PartName=\"/word/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/></Types>");
    let source = rewrite_part(&seed, "word/styles.xml", Some(STYLES));
    let source = rewrite_part(&source, "word/header1.xml", Some(HEADER));
    let source = rewrite_part(&source, "word/_rels/header1.xml.rels", Some("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rIdLink\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"https://example.org/story\" TargetMode=\"External\"/></Relationships>"));
    let source = rewrite_part(&source, "word/_rels/document.xml.rels", Some(&relations));
    let source = rewrite_part(&source, "[Content_Types].xml", Some(&types));
    rewrite_part(&source, "word/theme/theme1.xml", Some(&source_theme()))
}

fn text_style(raw: &Value) -> Value {
    let family = raw["fontFamilyAscii"].as_str().filter(|value| !value.is_empty()).or_else(|| raw["resolvedFontFamilyAscii"].as_str()).unwrap_or("Aptos");
    let mut result = json!({ "fontFamily": family, "fontSize": raw["fontSizePt"].as_f64().unwrap_or(11.0), "color": raw["color"].as_str().unwrap_or("#1f2937"),
        "bold": raw["bold"].as_bool().unwrap_or(false), "italic": raw["italic"].as_bool().unwrap_or(false), "underline": raw["underline"].as_bool().unwrap_or(false), "strike": raw["strike"].as_bool().unwrap_or(false) });
    if let Some(family) = raw["fontFamilyEastAsia"].as_str().filter(|value| !value.is_empty()).or_else(|| raw["resolvedFontFamilyEastAsia"].as_str()) { result["eastAsiaFontFamily"] = json!(family); }
    if let Some(vertical) = raw["verticalAlign"].as_str().filter(|value| matches!(*value, "superscript" | "subscript")) { result["verticalAlign"] = json!(vertical); }
    if let Some(highlight) = raw["highlight"].as_str().filter(|value| !matches!(*value, "none" | "")) { result["highlight"] = json!(highlight); }
    result
}

fn paragraph_style(raw: &Value) -> Value {
    let rule = raw["spacing"]["lineRule"].as_str().unwrap_or("auto");
    let line = raw["spacing"]["lineTwips"].as_f64().map(|value| value / if rule == "auto" { 240.0 } else { 15.0 })
        .or_else(|| raw["spacing"]["lineMultiplier"].as_f64()).unwrap_or(1.5);
    let before = raw["spacing"]["beforeTwips"].as_f64().map(|value| value / 15.0).or_else(|| raw["spacing"]["beforePt"].as_f64().map(|value| value * 4.0 / 3.0)).unwrap_or(0.0);
    let after = raw["spacing"]["afterTwips"].as_f64().map(|value| value / 15.0).or_else(|| raw["spacing"]["afterPt"].as_f64().map(|value| value * 4.0 / 3.0)).unwrap_or(8.0);
    let mut result = json!({ "align": raw["align"].as_str().unwrap_or("left"), "indent": 0, "spacingRule": rule, "lineHeight": line, "spaceBefore": before, "spaceAfter": after });
    for (target, source) in [("indentStart", "leftTwips"), ("indentEnd", "rightTwips"), ("indentFirstLine", "firstLineTwips"), ("hanging", "hangingTwips")] {
        if let Some(value) = raw["indent"][source].as_f64() { result[target] = json!(value / 15.0); }
    }
    for field in ["wordWrap", "snapToGrid"] { if let Some(value) = raw[field].as_bool() { result[field] = json!(value); } }
    result
}

fn normalize_run(raw: &Value, id: String) -> Value {
    let mut run = text_style(&raw["style"]);
    run["id"] = json!(id); run["text"] = raw["text"].clone();
    run["metadata"] = json!({ "source": raw, "officeStoryStyleBaseline": text_style(&raw["style"]) });
    if let Some(link) = raw["link"].as_str() { run["metadata"]["link"] = json!(link); }
    run
}

fn parsed_story(data: &Value) -> &Value {
    &data["sectionLayout"][0]["headerSections"].as_array().unwrap().iter().find(|reference| reference["referenceType"] == "default").unwrap()["storyText"]
}

// Assemble the published editable story contract without invoking a second
// parser. Source provenance is retained, while only model fields form baseline.
fn editable_story(story: &Value, story_key: &str) -> (Value, Vec<Value>) {
    let runs = story["runs"].as_array().unwrap().iter().enumerate().map(|(index, run)| normalize_run(run, format!("story-run-{index}"))).collect::<Vec<_>>();
    let paragraphs = story["paragraphs"].as_array().unwrap().iter().map(|paragraph| {
        let mut style = paragraph_style(&paragraph["style"]);
        if let Some(wrap) = paragraph["wordWrap"].as_bool() { style["wordWrap"] = json!(wrap); }
        json!({ "startOffset": paragraph["startOffset"], "endOffset": paragraph["endOffset"], "style": style,
            "metadata": { "source": paragraph, "officeStoryParagraphStyleBaseline": style } })
    }).collect::<Vec<_>>();
    let fields = story["fields"].as_array().unwrap().iter().enumerate().map(|(field_index, field)| {
        let style = text_style(&field["textStyle"]);
        let cache = field["cachedRuns"].as_array().unwrap().iter().enumerate().map(|(index, run)| normalize_run(run, format!("field-{field_index}-run-{index}"))).collect::<Vec<_>>();
        json!({ "story": story_key, "field": field["field"], "offset": field["offset"], "instruction": field["instruction"], "cachedResult": field["cachedResult"], "textStyle": style,
            "metadata": { "source": field, "officeStoryStyleBaseline": style, "officeStoryFieldCache": { "version": 1, "runs": cache } } })
    }).collect::<Vec<_>>();
    let baseline_runs = runs.iter().map(|run| {
        let mut run = run.clone();
        if let Some(link) = run["metadata"]["link"].as_str().map(str::to_owned) { run["href"] = json!(link); }
        let object = run.as_object_mut().unwrap(); object.remove("id"); object.remove("metadata"); run
    }).collect::<Vec<_>>();
    let baseline_paragraphs = paragraphs.iter().map(|paragraph| { let mut paragraph = paragraph.clone(); paragraph.as_object_mut().unwrap().remove("metadata"); paragraph }).collect::<Vec<_>>();
    let baseline_fields = fields.iter().map(|field| {
        let mut field = field.clone(); field["cachedRuns"] = field["metadata"]["officeStoryFieldCache"].clone();
        let object = field.as_object_mut().unwrap(); object.remove("story"); object.remove("metadata"); field
    }).collect::<Vec<_>>();
    let locks = story.get("locks").cloned().unwrap_or_else(|| json!([]));
    let rich = json!({ "runs": runs, "paragraphs": paragraphs, "locks": locks,
        "metadata": { "officeSource": { "version": 1, "xml": story["sourceXml"], "relationshipsXml": story["relationshipsXml"], "unsupportedFeatures": story["unsupportedFeatures"],
            "baseline": { "text": story["text"], "runs": baseline_runs, "paragraphs": baseline_paragraphs, "fields": baseline_fields, "locks": locks } } } });
    (rich, fields)
}

fn editable_model(data: &Value) -> Value {
    let story = parsed_story(data);
    let (rich, fields) = editable_story(story, "header");
    json!({ "kind": "document", "name": "editable-story", "metadata": { "source": { "officeSource": data["officeSource"] } },
        "defaultStyle": { "version": 1, "text": text_style(&data["defaultStyle"]["text"]), "paragraph": paragraph_style(&data["defaultStyle"]["paragraph"]) },
        "blocks": [{ "id": "body", "type": "paragraph", "runs": [{ "text": "body unchanged" }] }],
        "sectionHeaderFooters": { "first": { "header": story["text"], "footer": "", "pageFields": fields, "storyContent": { "header": rich } } } })
}

// This test adapter reconstructs all six native references before the next
// save, so dormant content is read from parser output rather than the seed.
fn editable_all_stories(data: &Value) -> Value {
    let mut model = editable_model(data);
    let section = &mut model["sectionHeaderFooters"]["first"];
    let mut fields = Vec::new();
    for (container, area) in [("headerSections", "header"), ("footerSections", "footer")] {
        for reference in data["sectionLayout"][0][container].as_array().unwrap() {
            let story_key = match (area, reference["referenceType"].as_str().unwrap()) {
                ("header", "default") => "header", ("footer", "default") => "footer",
                ("header", "first") => "firstPageHeader", ("footer", "first") => "firstPageFooter",
                ("header", "even") => "evenPageHeader", ("footer", "even") => "evenPageFooter",
                other => panic!("Unexpected native story reference in the fixture: {other:?}"),
            };
            let story = &reference["storyText"];
            let (rich, story_fields) = editable_story(story, story_key);
            section[story_key] = story["text"].clone();
            section["storyContent"][story_key] = rich;
            fields.extend(story_fields);
        }
    }
    section["pageFields"] = json!(fields);
    section["differentFirstPage"] = json!(data["sectionLayout"][0]["sectionPropertiesXml"].as_str().unwrap().contains("titlePg"));
    section["differentOddEven"] = data["evenAndOddHeaders"].clone();
    model
}

fn edit_other_word(model: &mut Value, before: &str, after: &str) {
    let section = &mut model["sectionHeaderFooters"]["first"];
    let text = section["header"].as_str().unwrap().replace(before, after);
    section["header"] = json!(text);
    let run = section["storyContent"]["header"]["runs"].as_array_mut().unwrap().iter_mut().find(|run| run["text"] == before).unwrap();
    run["text"] = json!(after);
    assert_eq!(before.encode_utf16().count(), after.encode_utf16().count(), "the unrelated edit retains paragraph and lock offsets");
}

fn assert_mixed_cache(story: &Value) {
    assert_eq!(story["fields"][0]["cachedResult"], "2026/11/22");
    assert_eq!(story["fields"][0]["offset"], 3);
    let cache = story["fields"][0]["cachedRuns"].as_array().unwrap();
    assert_eq!(cache.len(), 2);
    assert_eq!(cache[0]["text"], "2026/"); assert_eq!(cache[1]["text"], "11/22");
    assert_eq!(cache[0]["style"]["fontFamilyAscii"], "Georgia"); assert_eq!(cache[0]["style"]["fontSizePt"].as_f64(), Some(18.0));
    assert_eq!(cache[0]["style"]["bold"], false); assert_eq!(cache[0]["style"]["color"], "#FF0000");
    assert_eq!(cache[1]["style"]["fontFamilyAscii"], "Aptos"); assert_eq!(cache[1]["style"]["fontSizePt"].as_f64(), Some(12.0));
    assert_eq!(cache[1]["style"]["bold"], true); assert_eq!(cache[1]["style"]["underline"], true); assert_eq!(cache[1]["style"]["color"], "#0000FF");
    assert_eq!(cache[1]["link"], "https://example.org/story");
}

#[test]
fn two_docx_saves_retain_unicode_mixed_field_cache_advanced_properties_and_locks_after_other_edits() {
    let first = reopen(&source_package());
    assert_eq!(first["officeSource"]["styles"]["xml"], STYLES);
    assert_mixed_cache(parsed_story(&first));
    let mut model = editable_model(&first);
    edit_other_word(&mut model, " word", " edit");
    for (round, expected) in [(0, "A😀 edit\nTail"), (1, "A😀 next\nTail")] {
        let bytes = export(&model);
        let data = reopen(&bytes);
        let story = parsed_story(&data);
        assert_eq!(story["text"], expected); assert_mixed_cache(story);
        assert_eq!(story["locks"], json!([{ "id": "tail", "startOffset": 10, "endOffset": 14, "lock": "contentLocked" }]));
        let first_run = &story["runs"][0];
        assert_eq!(first_run["text"], "A😀"); assert_eq!(first_run["link"], "https://example.org/story");
        let source_properties = first_run["sourceFormat"]["propertiesXml"].as_str().unwrap();
        for property in ["w:caps", "w:smallCaps", "w:spacing", "w:lang", "w:themeColor=\"accent1\""] { assert!(source_properties.contains(property), "{property} survives round {round}"); }
        assert!(source_properties.contains("fr-FR"));
        let paragraph_properties = story["paragraphs"][0]["sourceFormat"]["propertiesXml"].as_str().unwrap();
        for property in ["w:keepNext", "w:tabs", "w:pBdr"] { assert!(paragraph_properties.contains(property), "{property} survives round {round}"); }
        let styles = read_part(&bytes, "word/styles.xml");
        assert!(styles.contains("Original Story Paragraph")); assert!(styles.contains("Original Story Character"));
        assert!(read_part(&bytes, "word/theme/theme1.xml").contains("Story Theme"));
        assert!(read_part(&bytes, "word/_rels/header1.xml.rels").contains("https://example.org/story"));
        assert!(data["blocks"].as_array().unwrap().iter().any(|block| block["content"] == "body unchanged"));
        model = editable_model(&data);
        if round == 0 { edit_other_word(&mut model, " edit", " next"); }
    }
}

#[test]
fn story_clear_markers_suppress_source_format_and_retain_cached_hyperlinks_on_two_saves() {
    let mut model = editable_model(&reopen(&source_package()));
    let mut text = model["defaultStyle"]["text"].clone(); text["fontSize"] = json!(9);
    let paragraph = model["defaultStyle"]["paragraph"].clone();
    let section = &mut model["sectionHeaderFooters"]["first"];
    let run = &mut section["storyContent"]["header"]["runs"][0];
    for (key, value) in text.as_object().unwrap() { run[key] = value.clone(); }
    run["metadata"]["officeStoryTextFormatCleared"] = json!(true);
    section["pageFields"][0]["textStyle"] = text.clone();
    section["pageFields"][0]["metadata"]["officeStoryTextFormatCleared"] = json!(true);
    section["storyContent"]["header"]["paragraphs"][0]["style"] = paragraph;
    section["storyContent"]["header"]["paragraphs"][0]["metadata"]["officeStoryParagraphFormatCleared"] = json!(true);
    for _ in 0..2 {
        let data = reopen(&export(&model));
        let story = parsed_story(&data);
        assert_eq!(story["text"], "A😀 word\nTail");
        let first = &story["runs"][0]; assert_eq!(first["style"]["fontSizePt"].as_f64(), Some(9.0));
        assert_eq!(first["link"], "https://example.org/story");
        assert_eq!(story["fields"][0]["field"], "DATE"); assert_eq!(story["fields"][0]["cachedResult"], "2026/11/22");
        let cache = story["fields"][0]["cachedRuns"].as_array().unwrap(); assert_eq!(cache.len(), 2);
        assert_eq!(cache[0]["text"], "2026/"); assert_eq!(cache[1]["text"], "11/22");
        assert_eq!(cache[1]["link"], "https://example.org/story");
        for run in [first, &cache[0], &cache[1]] {
            assert_eq!(run["style"]["fontSizePt"].as_f64(), Some(9.0)); assert_eq!(run["style"]["bold"], false); assert_eq!(run["style"]["underline"], false);
            let properties = run["sourceFormat"]["propertiesXml"].as_str().unwrap();
            for property in ["w:caps", "w:smallCaps", "w:spacing", "w:rStyle", "w:themeColor"] { assert!(!properties.contains(property), "cleared source property must stay absent: {property}"); }
            // The writer's ordinary zh-CN default can materialize on a later
            // source replay; the cleared source character language cannot.
            assert!(!properties.contains("fr-FR"));
        }
        let properties = story["paragraphs"][0]["sourceFormat"]["propertiesXml"].as_str().unwrap();
        for property in ["w:keepNext", "w:tabs", "w:pBdr"] { assert!(!properties.contains(property)); }
        // A later source replay may materialize the normal paragraph style;
        // the cleared custom paragraph's advanced properties remain absent.
        assert_eq!(story["locks"][0]["id"], "tail");
        model = editable_model(&data);
    }
}

#[test]
fn cleared_field_cache_retains_later_mixed_local_edits_and_links_on_two_docx_saves() {
    let mut model = editable_model(&reopen(&source_package()));
    let mut cleared_style = model["defaultStyle"]["text"].clone();
    cleared_style["fontSize"] = json!(9);
    let field = &mut model["sectionHeaderFooters"]["first"]["pageFields"][0];
    field["textStyle"] = cleared_style.clone();
    field["metadata"]["officeStoryStyleBaseline"] = cleared_style.clone();
    field["metadata"]["officeStoryTextFormatCleared"] = json!(true);
    for run in field["metadata"]["officeStoryFieldCache"]["runs"].as_array_mut().unwrap() {
        for (key, value) in cleared_style.as_object().unwrap() { run[key] = value.clone(); }
    }
    let first = &mut field["metadata"]["officeStoryFieldCache"]["runs"][0];
    first["fontSize"] = json!(17); first["bold"] = json!(true); first["highlight"] = json!("yellow");
    let second = &mut field["metadata"]["officeStoryFieldCache"]["runs"][1];
    second["italic"] = json!(true);
    for _ in 0..2 {
        let data = reopen(&export(&model));
        let story = parsed_story(&data);
        assert_eq!(story["fields"][0]["field"], "DATE");
        assert_eq!(story["fields"][0]["cachedResult"], "2026/11/22");
        let cache = story["fields"][0]["cachedRuns"].as_array().unwrap();
        assert_eq!(cache.len(), 2);
        assert_eq!(cache[0]["text"], "2026/"); assert_eq!(cache[1]["text"], "11/22");
        assert_eq!(cache[0]["style"]["fontSizePt"].as_f64(), Some(17.0));
        assert_eq!(cache[0]["style"]["bold"], true); assert_eq!(cache[0]["style"]["highlight"], "yellow");
        assert_eq!(cache[1]["style"]["fontSizePt"].as_f64(), Some(9.0));
        assert_eq!(cache[1]["style"]["bold"], false); assert_eq!(cache[1]["style"]["italic"], true);
        assert_eq!(cache[1]["link"], "https://example.org/story");
        for run in cache {
            let properties = run["sourceFormat"]["propertiesXml"].as_str().unwrap();
            for property in ["w:caps", "w:smallCaps", "w:spacing", "w:rStyle", "w:themeColor"] {
                assert!(!properties.contains(property), "cleared source property returned: {property}");
            }
        }
        model = editable_model(&data);
    }
}

#[test]
fn changed_field_cache_and_unretained_style_resource_fail_without_silent_flattening() {
    let original = editable_model(&reopen(&source_package()));
    let mut cache = original.clone(); cache["sectionHeaderFooters"]["first"]["pageFields"][0]["cachedResult"] = json!("changed");
    assert!(office_writer::export_docx_bytes(&cache.to_string()).is_err());
    let missing_source = rewrite_part(&source_package(), "word/styles.xml", None);
    let missing = reopen(&missing_source);
    assert!(!missing["officeSource"]["resourceErrors"].as_array().unwrap().is_empty());
    let mut retained = original;
    retained["metadata"]["source"]["officeSource"] = missing["officeSource"].clone();
    assert!(office_writer::export_docx_bytes(&retained.to_string()).is_err());
}

#[test]
fn custom_date_time_instructions_and_refreshed_mixed_caches_survive_two_docx_saves() {
    let header = HEADER.replacen("<w:color w:val=\"FF0000\"/><w:b w:val=\"0\"/>",
        "<w:color w:val=\"FF0000\"/><w:b w:val=\"0\"/><w:lang w:val=\"fr-FR\"/>", 1);
    let source = reopen(&rewrite_part(&source_package(), "word/header1.xml", Some(&header)));
    assert_eq!(parsed_story(&source)["fields"][0]["cachedRuns"][0]["style"]["language"], "fr-FR");
    let mut model = editable_model(&source);
    let date_instruction = r#"DATE \@ "yyyy'/'MM'/'dd' \"Hi\"'""#;
    let time_instruction = r#"TIME \@ "hh':'mm':'ss' 'AM/PM" \*Lower"#;
    model["sectionHeaderFooters"]["first"]["pageFields"][0]["instruction"] = json!(date_instruction);
    let defaults = model["defaultStyle"]["text"].clone();
    let mut first_time_run = defaults.clone();
    first_time_run["id"] = json!("time-first"); first_time_run["text"] = json!("01:");
    first_time_run["bold"] = json!(true); first_time_run["color"] = json!("#00AA00");
    let mut last_time_run = defaults.clone();
    last_time_run["id"] = json!("time-last"); last_time_run["text"] = json!("04:05 pm");
    last_time_run["italic"] = json!(true); last_time_run["color"] = json!("#AA00AA");
    last_time_run["href"] = json!("https://example.org/time");
    let fields = model["sectionHeaderFooters"]["first"]["pageFields"].as_array_mut().unwrap();
    fields.push(json!({ "story": "header", "field": "TIME", "offset": 13, "instruction": time_instruction, "cachedResult": "01:04:05 pm",
        "textStyle": defaults, "metadata": { "officeStoryStyleBaseline": defaults, "officeStoryFieldCache": { "version": 1, "runs": [first_time_run, last_time_run] } } }));
    fields.push(json!({ "story": "header", "field": "DATE", "offset": 13, "instruction": "DATE \\h", "cachedResult": "opaque result", "textStyle": defaults }));
    // These cache values are explicit model-boundary fixtures, not a second
    // Rust date formatter. The TypeScript picture resolver has its own draft.
    for (round, suffix, hour, remainder) in [(0, "11/23 \"Hi\"", "01:", "04:05 pm"), (1, "11/24 \"Hi\"", "02:", "05:06 pm")] {
        let section = &mut model["sectionHeaderFooters"]["first"];
        section["pageFields"][0]["cachedResult"] = json!(format!("2026/{suffix}"));
        section["pageFields"][0]["metadata"]["officeStoryFieldCache"]["runs"][1]["text"] = json!(suffix);
        section["pageFields"][1]["cachedResult"] = json!(format!("{hour}{remainder}"));
        section["pageFields"][1]["metadata"]["officeStoryFieldCache"]["runs"][0]["text"] = json!(hour);
        section["pageFields"][1]["metadata"]["officeStoryFieldCache"]["runs"][1]["text"] = json!(remainder);
        let bytes = export(&model);
        let xml = read_part(&bytes, "word/header1.xml");
        assert!(xml.contains("\\&quot;Hi\\&quot;"), "escaped instruction quotes stay inside the XML attribute");
        let data = reopen(&bytes);
        let story = parsed_story(&data);
        assert_eq!(story["text"], if round == 0 { "A😀 word\nTail" } else { "A😀 edit\nTail" });
        assert_eq!(story["locks"], json!([{ "id": "tail", "startOffset": 10, "endOffset": 14, "lock": "contentLocked" }]));
        assert_eq!(story["fields"].as_array().unwrap().len(), 3);
        let date = &story["fields"][0];
        assert_eq!(date["instruction"], date_instruction); assert_eq!(date["offset"], 3);
        assert_eq!(date["cachedResult"], format!("2026/{suffix}"));
        assert_eq!(date["cachedRuns"][0]["text"], "2026/"); assert_eq!(date["cachedRuns"][1]["text"], suffix);
        assert_eq!(date["cachedRuns"][0]["style"]["color"], "#FF0000");
        assert_eq!(date["cachedRuns"][0]["style"]["language"], "fr-FR");
        assert_eq!(date["textStyle"]["language"], "fr-FR");
        assert_eq!(date["cachedRuns"][1]["style"]["color"], "#0000FF");
        assert_eq!(date["cachedRuns"][1]["link"], "https://example.org/story");
        let time = &story["fields"][1];
        assert_eq!(time["instruction"], time_instruction); assert_eq!(time["offset"], 13);
        assert_eq!(time["cachedResult"], format!("{hour}{remainder}"));
        assert_eq!(time["cachedRuns"][0]["text"], hour); assert_eq!(time["cachedRuns"][1]["text"], remainder);
        assert_eq!(time["cachedRuns"][0]["style"]["bold"], true); assert_eq!(time["cachedRuns"][0]["style"]["color"], "#00AA00");
        assert_eq!(time["cachedRuns"][1]["style"]["italic"], true); assert_eq!(time["cachedRuns"][1]["style"]["color"], "#AA00AA");
        assert_eq!(time["cachedRuns"][1]["link"], "https://example.org/time");
        assert_eq!(story["fields"][2]["instruction"], "DATE \\h"); assert_eq!(story["fields"][2]["cachedResult"], "opaque result");
        model = editable_model(&data);
        assert_eq!(model["sectionHeaderFooters"]["first"]["pageFields"][0]["metadata"]["officeStoryFieldCache"]["runs"][0]["metadata"]["source"]["style"]["language"], "fr-FR");
        if round == 0 { edit_other_word(&mut model, " word", " edit"); }
    }
}

#[test]
fn dormant_first_and_even_stories_reopen_before_their_options_are_enabled() {
    let text_style = json!({ "fontFamily": "Aptos", "fontSize": 11, "color": "#1f2937", "bold": false, "italic": false, "underline": false, "strike": false });
    let mut first_run = text_style.clone(); first_run["id"] = json!("cover-one"); first_run["text"] = json!("Cover");
    first_run["bold"] = json!(true); first_run["color"] = json!("#AA0000");
    let mut last_run = text_style.clone(); last_run["id"] = json!("cover-two"); last_run["text"] = json!("😀");
    last_run["italic"] = json!(true); last_run["color"] = json!("#0000AA"); last_run["href"] = json!("https://example.org/cover");
    let mut first_cache = first_run.clone(); first_cache["id"] = json!("cover-cache-one"); first_cache["text"] = json!("2026/");
    let mut last_cache = last_run.clone(); last_cache["id"] = json!("cover-cache-two"); last_cache["text"] = json!("11/22");
    let mut model = json!({ "kind": "document", "name": "dormant-stories", "blocks": [{ "id": "body", "type": "paragraph", "runs": [{ "text": "body unchanged" }] }],
        "sectionHeaderFooters": { "first": {
            "header": "Main header", "footer": "Main footer", "firstPageHeader": "Cover😀", "firstPageFooter": "Cover footer",
            "evenPageHeader": "Even header", "evenPageFooter": "Even footer", "differentFirstPage": false, "differentOddEven": false,
            "storyContent": { "firstPageHeader": { "runs": [first_run, last_run], "paragraphs": [{ "startOffset": 0, "endOffset": 7,
                "style": { "align": "center", "indent": 0, "spacingRule": "auto", "lineHeight": 1.5, "spaceBefore": 0, "spaceAfter": 8 } }] } },
            "pageFields": [
                { "story": "firstPageHeader", "field": "DATE", "offset": 0, "instruction": "DATE \\@ \"yyyy/MM/dd\"", "cachedResult": "2026/11/22",
                    "textStyle": text_style, "metadata": { "officeStoryStyleBaseline": text_style, "officeStoryFieldCache": { "version": 1, "runs": [first_cache, last_cache] } } },
                { "story": "evenPageFooter", "field": "TIME", "offset": 0, "instruction": "TIME \\@ \"HH:mm:ss\"", "cachedResult": "13:04:05" }
            ]
        } } });
    for (first, even) in [(false, false), (true, false), (true, true)] {
        model["sectionHeaderFooters"]["first"]["differentFirstPage"] = json!(first);
        model["sectionHeaderFooters"]["first"]["differentOddEven"] = json!(even);
        let bytes = export(&model);
        let data = reopen(&bytes);
        let section = &data["sectionLayout"][0];
        assert_eq!(section["headerSections"].as_array().unwrap().len(), 3);
        assert_eq!(section["footerSections"].as_array().unwrap().len(), 3);
        assert_eq!(section["sectionPropertiesXml"].as_str().unwrap().contains("titlePg"), first);
        assert_eq!(data["evenAndOddHeaders"], even);
        assert_eq!(read_part(&bytes, "word/settings.xml").contains("evenAndOddHeaders"), even);
        for (container, kind, expected) in [
            ("headerSections", "default", "Main header"), ("footerSections", "default", "Main footer"),
            ("headerSections", "first", "Cover😀"), ("footerSections", "first", "Cover footer"),
            ("headerSections", "even", "Even header"), ("footerSections", "even", "Even footer"),
        ] {
            let story = &section[container].as_array().unwrap().iter().find(|reference| reference["referenceType"] == kind).unwrap()["storyText"];
            assert_eq!(story["text"], expected);
            if container == "headerSections" && kind == "first" {
                assert_eq!(story["runs"][0]["style"]["bold"], true); assert_eq!(story["runs"][0]["style"]["color"], "#AA0000");
                assert_eq!(story["runs"][1]["style"]["italic"], true); assert_eq!(story["runs"][1]["link"], "https://example.org/cover");
                assert_eq!(story["paragraphs"][0]["style"]["align"], "center");
                assert_eq!(story["fields"][0]["instruction"], "DATE \\@ \"yyyy/MM/dd\"");
                assert_eq!(story["fields"][0]["cachedResult"], "2026/11/22");
                assert_eq!(story["fields"][0]["cachedRuns"][0]["text"], "2026/");
                assert_eq!(story["fields"][0]["cachedRuns"][1]["text"], "11/22");
                assert_eq!(story["fields"][0]["cachedRuns"][1]["link"], "https://example.org/cover");
            }
            if container == "footerSections" && kind == "even" {
                assert_eq!(story["fields"][0]["instruction"], "TIME \\@ \"HH:mm:ss\"");
                assert_eq!(story["fields"][0]["cachedResult"], "13:04:05");
            }
        }
        model = editable_all_stories(&data);
    }
}
