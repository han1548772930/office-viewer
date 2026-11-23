//! Native numbering-resource and shared table/story binding drafts, unexecuted.
use std::io::{Cursor, Read};
use serde_json::{Value, json};
use zip::ZipArchive;
use office_io_common::document_numbering_source::decode_media;
use crate::{document_format_sidecar::{Element, parse_xml}, parser::parse_office_bytes};

const WORD: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const PNG: &str = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=";

fn source() -> Value {
    json!({"version":1,"styles":{"partName":"word/styles.xml","xml":format!(r#"<w:styles xmlns:w="{WORD}"><w:style w:type="table" w:styleId="SourceGrid"><w:tblStylePr w:type="firstRow"><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:tblStylePr></w:style></w:styles>"#)},
        "numbering":{"partName":"word/numbering.xml","xml":format!(r##"<w:numbering xmlns:w="{WORD}" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:numPicBullet w:numPicBulletId="1"><w:pict><v:shapetype id="_x0000_t75"/><v:shape id="_x0000_i1001" type="#_x0000_t75" style="width:12px;height:12px"><v:imagedata r:id="sourcePicture"/></v:shape></w:pict></w:numPicBullet><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="4"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlPicBulletId w:val="1"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="9"/></w:lvlOverride></w:num></w:numbering>"##),
            "relationshipsXml":"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"sourcePicture\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/source.png\"/></Relationships>",
            "resources":[{"partName":"word/media/source.png","contentType":"image/png","dataBase64":PNG}]}})
}

fn model() -> Value {
    let header = "Header😀";
    json!({"kind":"document","name":"Numbered tables and stories","metadata":{"source":{"officeSource":source()}},
        "blocks":[{"id":"body-table","type":"table","runs":[],"metadata":{"table":{"columnWidths":[160],"headerRowCount":1,
            "style":{"nativeProperties":{"version":1,"propertiesXml":"<w:tblPr><w:tblStyle w:val=\"SourceGrid\"/></w:tblPr>","namespaces":{"w":WORD},"baseline":{}}},
            "rows":[{"id":"row","cells":[{"id":"cell","columnIndex":0,"rowSpan":1,"colSpan":1,"style":{},"blocks":[{"id":"cell-paragraph","type":"paragraph","runs":[{"text":"Table content"}]}]}]}]}}}],
        "sectionHeaderFooters":{"first":{"header":header,"storyContent":{"header":{"runs":[{"id":"header-run","text":header}],
            "paragraphs":[{"startOffset":0,"endOffset":header.encode_utf16().count(),"style":{},"metadata":{"source":{"sourceFormat":{"version":1,"propertiesXml":"<w:pPr><w:numPr><w:numId w:val=\"1\"/></w:numPr></w:pPr>","namespaces":{"w":WORD}}}}}]}}}}})
}

fn part(bytes: &[u8], path: &str) -> String {
    let mut archive = ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut xml = String::new(); archive.by_name(path).unwrap().read_to_string(&mut xml).unwrap(); xml
}
fn descendant<'a>(node: &'a Element, name: &str) -> Option<&'a Element> {
    if node.name == name { Some(node) } else { node.children.iter().find_map(|child| descendant(child, name)) }
}

#[test]
fn a_table_conditional_style_and_header_paragraph_share_one_numbering_instance_after_real_docx_reopen() {
    let input = model(); let original = input.clone();
    let (bytes, warnings) = office_writer::export_docx_bytes(&input.to_string()).unwrap();
    assert!(warnings.is_empty(), "{warnings:?}"); assert_eq!(input, original);
    let header = parse_xml(&part(&bytes, "word/header1.xml")).unwrap();
    let story_number = descendant(&header, "numId").unwrap().attribute("val").unwrap();
    let body = parse_xml(&part(&bytes, "word/document.xml")).unwrap();
    let style_id = descendant(&body, "tblStyle").unwrap().attribute("val").unwrap();
    let styles = parse_xml(&part(&bytes, "word/styles.xml")).unwrap();
    let style = styles.children.iter().find(|style| style.attribute("styleId") == Some(style_id)).unwrap();
    assert_eq!(descendant(style, "numId").unwrap().attribute("val"), Some(story_number));
    let numbering_xml = part(&bytes, "word/numbering.xml"); let numbering = parse_xml(&numbering_xml).unwrap();
    let instance = numbering.children.iter().find(|node| node.name == "num" && node.attribute("numId") == Some(story_number)).unwrap();
    assert_eq!(descendant(instance, "startOverride").unwrap().attribute("val"), Some("9"));
    let parsed = serde_json::to_value(parse_office_bytes(&bytes, "numbering-resources.docx").unwrap()).unwrap();
    let resource = &parsed["payload"]["data"]["officeSource"]["numbering"];
    assert_eq!(resource["xml"], numbering_xml);
    let media = resource["resources"].as_array().unwrap(); assert_eq!(media.len(), 1);
    assert_eq!(media[0]["partName"], "word/media/office-source-numbering-1.png");
    assert_eq!(decode_media(media[0]["dataBase64"].as_str().unwrap()).unwrap(), decode_media(PNG).unwrap());
    assert!(parsed["payload"]["data"]["officeSource"]["resourceErrors"].is_null());
}
