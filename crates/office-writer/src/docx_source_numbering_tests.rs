//! Native resource drafts for the final unified verification phase.
use std::{collections::BTreeSet, io::{Cursor, Read}};
use serde_json::{Value, json};
use zip::ZipArchive;
use super::{SourceStyles, Element, WORD_NAMESPACE, parse_xml};
use crate::opc::Package;
use office_io_common::document_numbering_source::decode_media;

const PNG: &str = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=";
const GENERATED_STYLES: &str = r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="table" w:styleId="Grid"/><w:style w:type="paragraph" w:styleId="Normal"/><w:style w:type="character" w:styleId="Strong"/></w:styles>"#;
const GENERATED_NUMBERING: &str = r##"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:numPicBullet w:numPicBulletId="1"><w:pict><v:shapetype id="_x0000_t75"/><v:shape id="_x0000_i1001" type="#_x0000_t75"><v:imagedata r:id="rIdNumberingPicture1"/></v:shape></w:pict></w:numPicBullet><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>"##;

fn source() -> Value {
    let styles = format!(r#"<w:styles xmlns:w="{WORD_NAMESPACE}">
      <w:style w:type="paragraph" w:styleId="Normal"><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:style>
      <w:style w:type="character" w:styleId="Strong"><w:rPr><w:b/></w:rPr></w:style>
      <w:style w:type="numbering" w:styleId="ListLink"><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:style>
      <w:style w:type="table" w:styleId="Grid"><w:tblStylePr w:type="firstRow"><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr></w:tblStylePr></w:style>
    </w:styles>"#);
    let numbering = format!(r##"<w:numbering xmlns:w="{WORD_NAMESPACE}" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <w:numPicBullet w:numPicBulletId="1"><w:pict><v:shapetype id="_x0000_t75" coordsize="21600,21600"/><v:shape id="_x0000_i1001" type="#_x0000_t75" style="width:12px;height:12px"><v:imagedata r:id="rIdNumberingPicture1"/></v:shape></w:pict></w:numPicBullet>
      <w:abstractNum w:abstractNumId="0"><w:styleLink w:val="ListLink"/><w:lvl w:ilvl="0"><w:start w:val="3"/><w:numFmt w:val="bullet"/><w:pStyle w:val="Normal"/><w:lvlPicBulletId w:val="1"/><w:lvlText w:val="•"/><w:rPr><w:rStyle w:val="Strong"/></w:rPr></w:lvl></w:abstractNum>
      <w:num w:numId="1"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="7"/><w:lvl w:ilvl="0"><w:start w:val="5"/><w:numFmt w:val="upperRoman"/><w:lvlText w:val="%1)"/><w:rPr><w:rStyle w:val="Strong"/></w:rPr></w:lvl></w:lvlOverride></w:num>
    </w:numbering>"##);
    json!({"version":1,"styles":{"partName":"word/styles.xml","xml":styles},"numbering":{
        "partName":"word/numbering.xml","xml":numbering,
        "relationshipsXml":"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rIdNumberingPicture1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/source.png\"/></Relationships>",
        "resources":[{"partName":"word/media/source.png","contentType":"image/png","dataBase64":PNG}]
    }})
}

fn registry(source: &Value) -> SourceStyles {
    let model = json!({"metadata":{"source":{"officeSource":source}}});
    let mut registry = SourceStyles::new(&model, GENERATED_STYLES.into()).unwrap();
    registry.reserve_numbering(GENERATED_NUMBERING).unwrap(); registry
}

fn table_properties() -> Element {
    parse_xml(&format!(r#"<w:tblPr xmlns:w="{WORD_NAMESPACE}"><w:tblStyle w:val="Grid"/></w:tblPr>"#), false).unwrap()
}

fn package() -> Package {
    let mut package = Package::new();
    package.add_binary("word/media/generated.png", "image/png", decode_media(PNG).unwrap()).unwrap();
    package.add_relationship(Some("word/numbering.xml"), "rIdNumberingPicture1",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", "media/generated.png", false).unwrap();
    package
}

fn numbered_style<'a>(root: &'a Element, id: &str) -> &'a Element {
    root.children().find(|style| style.attribute("styleId") == Some(id)).unwrap()
}

fn descendants<'a>(root: &'a Element, local: &str, output: &mut Vec<&'a Element>) {
    if root.local == local { output.push(root); }
    for child in root.children() { descendants(child, local, output); }
}

#[test]
fn table_style_numbering_imports_real_overrides_and_preserves_legal_binding_backlinks() {
    let source = source(); let before = source.clone();
    let mut registry = registry(&source); let mut property = table_properties();
    registry.restore_table_properties(&mut property).unwrap();
    let style_id = property.child("tblStyle").unwrap().attribute("val").unwrap();
    assert_ne!(style_id, "Grid");
    let styles = parse_xml(&registry.styles_xml().unwrap(), true).unwrap();
    let table = numbered_style(&styles, style_id);
    let num_id = table.child("tblStylePr").unwrap().child("pPr").unwrap().child("numPr").unwrap().child("numId").unwrap().attribute("val").unwrap();
    assert_ne!(num_id, "1");
    let mut package = package();
    let numbering_xml = registry.numbering_xml(&mut package).unwrap();
    let numbering = parse_xml(&numbering_xml, true).unwrap();
    let number = numbering.children().find(|node| node.is_word("num") && node.attribute("numId") == Some(num_id)).unwrap();
    let abstract_id = number.child("abstractNumId").unwrap().attribute("val").unwrap();
    assert_ne!(abstract_id, "0");
    let definition = numbering.children().find(|node| node.is_word("abstractNum") && node.attribute("abstractNumId") == Some(abstract_id)).unwrap();
    let link = definition.child("styleLink").unwrap().attribute("val").unwrap();
    assert_eq!(numbered_style(&styles, link).child("pPr").unwrap().child("numPr").unwrap().child("numId").unwrap().attribute("val"), Some(num_id));
    let level = definition.child("lvl").unwrap();
    let paragraph = level.child("pStyle").unwrap().attribute("val").unwrap();
    assert_ne!(paragraph, "Normal");
    assert_eq!(numbered_style(&styles, paragraph).child("pPr").unwrap().child("numPr").unwrap().child("numId").unwrap().attribute("val"), Some(num_id));
    let override_level = number.child("lvlOverride").unwrap();
    assert_eq!(override_level.child("startOverride").unwrap().attribute("val"), Some("7"));
    assert_eq!(override_level.child("lvl").unwrap().child("numFmt").unwrap().attribute("val"), Some("upperRoman"));
    let run_style = override_level.child("lvl").unwrap().child("rPr").unwrap().child("rStyle").unwrap().attribute("val").unwrap();
    assert_ne!(run_style, "Strong");
    assert!(numbered_style(&styles, run_style).child("rPr").unwrap().child("b").is_some());
    package.add_xml("word/numbering.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml", numbering_xml).unwrap();
    let bytes = package.finish().unwrap();
    let mut archive = ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut rels = String::new(); archive.by_name("word/_rels/numbering.xml.rels").unwrap().read_to_string(&mut rels).unwrap();
    assert!(rels.contains("rIdNumberingPicture1")); assert!(rels.contains("rIdOfficeSourceNumbering1"));
    let mut media = Vec::new(); archive.by_name("word/media/office-source-numbering-1.png").unwrap().read_to_end(&mut media).unwrap();
    assert_eq!(media, decode_media(PNG).unwrap());
    let mut shapes = Vec::new(); descendants(&numbering, "shape", &mut shapes);
    let ids = shapes.iter().map(|shape| shape.attributes.iter().find(|attribute| attribute.local == "id").unwrap().value.as_str()).collect::<BTreeSet<_>>();
    assert_eq!(ids.len(), 2); assert!(ids.contains("_x0000_i1001"));
    let imported_shape = shapes.iter().find(|shape| shape.attributes.iter().any(|attribute| attribute.local == "id" && attribute.value != "_x0000_i1001")).unwrap();
    assert!(!imported_shape.attributes.iter().find(|attribute| attribute.local == "type").unwrap().value.ends_with("_x0000_t75"));
    assert_eq!(source, before);
}

#[test]
fn numbered_story_defaults_use_the_same_imported_instance_as_table_styles() {
    let mut source = source();
    let xml = source["styles"]["xml"].as_str().unwrap().replace("</w:styles>", "<w:docDefaults><w:pPrDefault><w:pPr><w:numPr><w:numId w:val=\"1\"/></w:numPr></w:pPr></w:pPrDefault></w:docDefaults></w:styles>");
    source["styles"]["xml"] = json!(xml);
    let mut registry = registry(&source);
    let provenance = json!({"metadata":{"source":{"sourceFormat":{"version":1,"propertiesXml":"<w:pPr/>","namespaces":{"w":WORD_NAMESPACE}}}}});
    let story = parse_xml(&registry.story_paragraph_properties("", &json!({}), &provenance).unwrap(), false).unwrap();
    let story_num = story.child("numPr").unwrap().child("numId").unwrap().attribute("val").unwrap();
    let mut table = table_properties(); registry.restore_table_properties(&mut table).unwrap();
    let styles = parse_xml(&registry.styles_xml().unwrap(), true).unwrap();
    let table = numbered_style(&styles, table.child("tblStyle").unwrap().attribute("val").unwrap());
    assert_eq!(table.child("tblStylePr").unwrap().child("pPr").unwrap().child("numPr").unwrap().child("numId").unwrap().attribute("val"), Some(story_num));
}

#[test]
fn numbered_picture_direct_sources_and_implicit_office_shape_types_get_real_media_bindings() {
    let mut source = source();
    let xml = source["numbering"]["xml"].as_str().unwrap()
        .replace("<v:shapetype id=\"_x0000_t75\" coordsize=\"21600,21600\"/>", "")
        .replace("r:id=\"rIdNumberingPicture1\"", "src=\"media/source.png\"");
    source["numbering"]["xml"] = json!(xml); source["numbering"].as_object_mut().unwrap().remove("relationshipsXml");
    let mut registry = registry(&source); registry.restore_table_properties(&mut table_properties()).unwrap();
    let mut package = package();
    let xml = registry.numbering_xml(&mut package).unwrap();
    assert!(xml.contains("src=\"media/office-source-numbering-1.png\""));
    assert!(xml.contains("coordsize=\"21600,21600\""));
    assert!(package.contains_part("word/media/office-source-numbering-1.png"));
}

#[test]
fn missing_numbering_media_and_delegation_cycles_fail_without_rewriting_provenance() {
    let mut missing = source(); missing["numbering"]["resources"] = json!([]);
    let before = missing.clone(); let mut missing_registry = registry(&missing);
    missing_registry.restore_table_properties(&mut table_properties()).unwrap();
    assert!(missing_registry.numbering_xml(&mut package()).unwrap_err().to_string().contains("DOCX_NUMBERING_MEDIA_MISSING"));
    assert_eq!(missing, before);
    let mut cycle = source();
    let xml = cycle["numbering"]["xml"].as_str().unwrap().replace("<w:styleLink w:val=\"ListLink\"/>", "<w:numStyleLink w:val=\"ListLink\"/>");
    cycle["numbering"]["xml"] = json!(xml);
    let mut cycle_registry = registry(&cycle); let mut property = table_properties();
    assert!(cycle_registry.restore_table_properties(&mut property).unwrap_err().contains("DOCX_NUMBERING_CYCLE"));
    assert_eq!(property.child("tblStyle").unwrap().attribute("val"), Some("Grid"));
}

#[test]
fn delegated_numbering_styles_inherit_a_real_instance_without_being_treated_as_an_association_cycle() {
    let mut source = source();
    let styles = source["styles"]["xml"].as_str().unwrap()
        .replace("<w:style w:type=\"numbering\" w:styleId=\"ListLink\"><w:pPr><w:numPr><w:numId w:val=\"1\"/></w:numPr></w:pPr></w:style>",
            "<w:style w:type=\"numbering\" w:styleId=\"ListLink\"><w:basedOn w:val=\"DelegateBase\"/></w:style><w:style w:type=\"numbering\" w:styleId=\"DelegateBase\"><w:pPr><w:numPr><w:numId w:val=\"2\"/></w:numPr></w:pPr></w:style>");
    source["styles"]["xml"] = json!(styles);
    let numbering = source["numbering"]["xml"].as_str().unwrap()
        .replace("<w:styleLink w:val=\"ListLink\"/>", "<w:numStyleLink w:val=\"ListLink\"/>")
        .replace("</w:numbering>", "<w:abstractNum w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"11\"/><w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"%1.\"/></w:lvl></w:abstractNum><w:num w:numId=\"2\"><w:abstractNumId w:val=\"1\"/></w:num></w:numbering>");
    source["numbering"]["xml"] = json!(numbering);
    let original = source.clone();
    let mut registry = registry(&source);
    let mut property = table_properties(); registry.restore_table_properties(&mut property).unwrap();
    let styles = parse_xml(&registry.styles_xml().unwrap(), true).unwrap();
    let root = parse_xml(&registry.numbering_xml(&mut package()).unwrap(), true).unwrap();
    let delegate = root.children().find(|node| node.is_word("abstractNum") && node.child("numStyleLink").is_some()).unwrap();
    let linked_style = numbered_style(&styles, delegate.child("numStyleLink").unwrap().attribute("val").unwrap());
    let base = numbered_style(&styles, linked_style.child("basedOn").unwrap().attribute("val").unwrap());
    let num_id = base.child("pPr").unwrap().child("numPr").unwrap().child("numId").unwrap().attribute("val").unwrap();
    let instance = root.children().find(|node| node.is_word("num") && node.attribute("numId") == Some(num_id)).unwrap();
    let abstract_id = instance.child("abstractNumId").unwrap().attribute("val").unwrap();
    let definition = root.children().find(|node| node.is_word("abstractNum") && node.attribute("abstractNumId") == Some(abstract_id)).unwrap();
    assert_eq!(definition.child("lvl").unwrap().child("start").unwrap().attribute("val"), Some("11"));
    assert!(definition.child("numStyleLink").is_none());
    assert_eq!(source, original);
}

#[test]
fn numbering_suppression_needs_no_source_and_import_rejects_unreserved_generated_ids() {
    let mut registry = SourceStyles::new(&json!({}), GENERATED_STYLES.into()).unwrap();
    for id in ["0", "00"] {
        let mut property = parse_xml(&format!("<w:pPr xmlns:w=\"{WORD_NAMESPACE}\"><w:numPr><w:numId w:val=\"{id}\"/></w:numPr></w:pPr>"), false).unwrap();
        registry.restore_table_properties(&mut property).unwrap();
        assert_eq!(property.child("numPr").unwrap().child("numId").unwrap().attribute("val"), Some(id));
    }
    let mut unreserved = SourceStyles::new(&json!({"metadata":{"source":{"officeSource":source()}}}), GENERATED_STYLES.into()).unwrap();
    assert!(unreserved.restore_table_properties(&mut table_properties()).unwrap_err().contains("DOCX_NUMBERING_ORDER"));
}
