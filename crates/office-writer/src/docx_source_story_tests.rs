use std::collections::BTreeMap;

use serde_json::{Value, json};

use super::{SourceStyles, WORD_NAMESPACE, parse_xml};

// DOC-02 contract drafts. This turn did not compile or execute these tests.
const GENERATED: &str = r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Normal"/><w:style w:type="paragraph" w:styleId="Heading1"/><w:style w:type="character" w:styleId="Strong"/></w:styles>"#;

fn registry(styles: &str, theme: Option<&str>) -> SourceStyles {
    let model = json!({ "metadata": { "source": { "officeSource": {
        "version": 1, "styles": { "xml": styles }, "theme": theme.map(|xml| json!({ "xml": xml }))
    } } } });
    SourceStyles::new(&model, GENERATED.into()).expect("generated style registry can be created")
}

fn source_format(properties: &str, paragraph: Option<&str>) -> Value {
    json!({ "version": 1, "propertiesXml": properties, "paragraphStyleId": paragraph,
        "namespaces": { "w": WORD_NAMESPACE } })
}

fn run_provenance(properties: &str, baseline: &Value) -> Value {
    json!({ "metadata": { "source": { "sourceFormat": source_format(properties, None) }, "officeStoryStyleBaseline": baseline } })
}

#[test]
fn story_and_numbering_marker_share_remapped_style_ids_and_dependency_closure() {
    let mut styles = registry(r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Source Normal"/><w:pPr><w:keepNext/></w:pPr></w:style>
      <w:style w:type="paragraph" w:styleId="Heading1"><w:basedOn w:val="Normal"/><w:pPr><w:jc w:val="center"/></w:pPr><w:rPr><w:caps/></w:rPr></w:style>
      <w:style w:type="character" w:styleId="Strong"><w:rPr><w:b/><w:smallCaps/></w:rPr></w:style>
    </w:styles>"#, None);
    let mut story = parse_xml(&format!(r#"<w:hdr xmlns:w="{WORD_NAMESPACE}"><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:rPr><w:rStyle w:val="Strong"/></w:rPr><w:t>story</w:t></w:r></w:p></w:hdr>"#), true).unwrap();
    styles.restore_story(&mut story).expect("source story is restored");
    let paragraph = story.child("p").unwrap();
    let paragraph_id = paragraph.child("pPr").unwrap().child("pStyle").unwrap().attribute("val").unwrap().to_owned();
    let run_id = paragraph.child("r").unwrap().child("rPr").unwrap().child("rStyle").unwrap().attribute("val").unwrap().to_owned();
    assert_ne!(paragraph_id, "Heading1");
    assert_ne!(run_id, "Strong");
    assert!(paragraph.child("pPr").unwrap().child("keepNext").is_some());
    let mut marker = vec![parse_xml(&format!(r#"<w:r xmlns:w="{WORD_NAMESPACE}"><w:rPr><w:rStyle w:val="Strong"/></w:rPr><w:t>1.</w:t></w:r>"#), false).unwrap()];
    styles.restore_marker(&mut marker, &BTreeMap::from([("w".into(), WORD_NAMESPACE.into())]), &json!({ "metadata": { "style": { "styleId": "Heading1" } } })).unwrap();
    assert_eq!(marker[0].child("rPr").unwrap().child("rStyle").unwrap().attribute("val"), Some(run_id.as_str()));
    let result = parse_xml(&styles.styles_xml(), true).unwrap();
    let imported = result.children().find(|style| style.attribute("styleId") == Some(paragraph_id.as_str())).unwrap();
    let base = imported.child("basedOn").unwrap().attribute("val").unwrap();
    assert_ne!(base, "Normal");
    let normal = result.children().find(|style| style.attribute("styleId") == Some(base)).unwrap();
    assert!(normal.attribute("default").is_none());
    assert_eq!(normal.child("name").unwrap().attribute("val"), Some("Source Normal"));
    assert_eq!(result.children().filter(|style| style.attribute("styleId") == Some(run_id.as_str())).count(), 1);
}

#[test]
fn strict_and_alternate_prefix_theme_context_is_emitted_as_transitional_ooxml() {
    let mut styles = registry(r#"<s:styles xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main"><s:docDefaults><s:rPrDefault><s:rPr><s:rFonts s:asciiTheme="minorHAnsi"/><s:color s:themeColor="accent1"/></s:rPr></s:rPrDefault></s:docDefaults></s:styles>"#,
      Some(r#"<d:theme xmlns:d="http://purl.oclc.org/ooxml/drawingml/main" name="source"><d:themeElements><d:clrScheme name="source"><d:accent1><d:srgbClr val="112233"/></d:accent1></d:clrScheme><d:fontScheme name="source"><d:majorFont><d:latin typeface="Georgia"/></d:majorFont><d:minorFont><d:latin typeface="Aptos"/></d:minorFont></d:fontScheme></d:themeElements></d:theme>"#));
    let current = json!({ "fontFamily": "Aptos", "fontSize": 10, "color": "#112233", "bold": false });
    let mut provenance = run_provenance("<s:rPr><s:lang s:val=\"zh-CN\"/></s:rPr>", &current);
    provenance["metadata"]["source"]["sourceFormat"]["namespaces"] = json!({ "s": "http://purl.oclc.org/ooxml/wordprocessingml/main" });
    let xml = styles.story_run_properties("<w:b w:val=\"0\"/>", &current, &provenance).unwrap();
    let result = parse_xml(&xml, false).unwrap();
    assert!(result.is_word("rPr"));
    assert_eq!(result.child("rFonts").unwrap().attribute("asciiTheme"), Some("minorHAnsi"));
    assert_eq!(result.child("color").unwrap().attribute("themeColor"), Some("accent1"));
    assert_eq!(result.child("lang").unwrap().attribute("val"), Some("zh-CN"));
    assert!(!xml.contains("http://purl.oclc.org/ooxml/"));
    let theme_xml = styles.theme_xml().expect("referenced source theme is retained");
    let theme = parse_xml(theme_xml, true).unwrap();
    assert_eq!(theme.namespace, "http://schemas.openxmlformats.org/drawingml/2006/main");
    assert!(!theme_xml.contains("http://purl.oclc.org/ooxml/"));
}

#[test]
fn modeled_property_overlay_keeps_advanced_run_properties_and_clear_drops_them() {
    let mut styles = registry(&format!("<w:styles xmlns:w=\"{WORD_NAMESPACE}\"/>"), None);
    let baseline = json!({ "fontFamily": "Georgia", "fontSize": 18, "color": "#112233", "bold": false, "italic": false });
    let mut current = baseline.clone(); current["bold"] = json!(true);
    let mut provenance = run_provenance(r#"<w:rPr><w:rFonts w:ascii="Georgia" w:eastAsia="宋体"/><w:sz w:val="36"/><w:b w:val="0"/><w:caps/><w:smallCaps/><w:spacing w:val="24"/><w:lang w:val="zh-CN"/><w:color w:val="112233"/></w:rPr>"#, &baseline);
    let generated = r#"<w:rFonts w:ascii="Aptos"/><w:sz w:val="18"/><w:b w:val="1"/><w:color w:val="FF0000"/>"#;
    let xml = styles.story_run_properties(generated, &current, &provenance).unwrap();
    let preserved = parse_xml(&xml, false).unwrap();
    assert_eq!(preserved.child("b").unwrap().attribute("val"), Some("1"));
    assert_eq!(preserved.child("rFonts").unwrap().attribute("ascii"), Some("Georgia"));
    assert_eq!(preserved.child("sz").unwrap().attribute("val"), Some("36"));
    assert!(preserved.child("caps").is_some()); assert!(preserved.child("smallCaps").is_some());
    assert_eq!(preserved.child("spacing").unwrap().attribute("val"), Some("24"));
    provenance["metadata"]["officeStoryTextFormatCleared"] = json!(true);
    let cleared = parse_xml(&styles.story_run_properties(generated, &current, &provenance).unwrap(), false).unwrap();
    assert!(cleared.child("caps").is_none()); assert!(cleared.child("smallCaps").is_none());
    assert!(cleared.child("spacing").is_none()); assert!(cleared.child("lang").is_none());
    assert_eq!(cleared.child("rFonts").unwrap().attribute("ascii"), Some("Aptos"));
}

#[test]
fn paragraph_overlay_keeps_advanced_properties_but_clear_drops_source_paragraph_mark() {
    let mut styles = registry(&format!("<w:styles xmlns:w=\"{WORD_NAMESPACE}\"/>"), None);
    let baseline = json!({ "align": "center", "spaceBefore": 12, "spaceAfter": 6, "lineHeight": 1.25 });
    let mut current = baseline.clone(); current["align"] = json!("left");
    let mut provenance = json!({ "metadata": { "source": { "sourceFormat": source_format(r#"<w:pPr><w:jc w:val="center"/><w:keepNext/><w:spacing w:before="240" w:after="120"/><w:tabs><w:tab w:pos="720" w:val="right"/></w:tabs><w:rPr><w:smallCaps/></w:rPr></w:pPr>"#, None) }, "officeStoryParagraphStyleBaseline": baseline } });
    let generated = r#"<w:jc w:val="left"/><w:spacing w:before="0" w:after="0"/>"#;
    let preserved = parse_xml(&styles.story_paragraph_properties(generated, &current, &provenance).unwrap(), false).unwrap();
    assert_eq!(preserved.child("jc").unwrap().attribute("val"), Some("left"));
    assert_eq!(preserved.child("spacing").unwrap().attribute("before"), Some("240"));
    assert!(preserved.child("keepNext").is_some()); assert!(preserved.child("tabs").is_some());
    assert!(preserved.child("rPr").unwrap().child("smallCaps").is_some());
    provenance["metadata"]["officeStoryParagraphFormatCleared"] = json!(true);
    let cleared = parse_xml(&styles.story_paragraph_properties(generated, &current, &provenance).unwrap(), false).unwrap();
    assert!(cleared.child("keepNext").is_none()); assert!(cleared.child("tabs").is_none()); assert!(cleared.child("rPr").is_none());
    assert_eq!(cleared.child("spacing").unwrap().attribute("before"), Some("0"));
}

#[test]
fn equal_extension_prefixes_keep_independent_default_style_and_story_scopes() {
    const DEFAULT_EXTENSION: &str = "http://schemas.microsoft.com/office/word/2010/wordml";
    const STYLE_EXTENSION: &str = "http://schemas.microsoft.com/office/word/2012/wordml";
    const STORY_EXTENSION: &str = "http://schemas.microsoft.com/office/word/2018/wordml";
    let mut styles = registry(&format!(r#"<w:styles xmlns:w="{WORD_NAMESPACE}"><w:docDefaults xmlns:x="{DEFAULT_EXTENSION}"><w:rPrDefault><w:rPr><x:ligatures w:val="standard"/></w:rPr></w:rPrDefault></w:docDefaults><w:style xmlns:x="{STYLE_EXTENSION}" w:type="paragraph" w:styleId="Normal" w:default="1"><w:rPr><x:ligatures w:val="contextual"/></w:rPr></w:style></w:styles>"#), None);
    let mut story = parse_xml(&format!(r#"<w:hdr xmlns:w="{WORD_NAMESPACE}" xmlns:x="{STORY_EXTENSION}"><w:p><w:pPr><w:rPr><x:ligatures w:val="historical"/></w:rPr></w:pPr><w:r><w:rPr><x:ligatures w:val="historical"/></w:rPr><w:t>styled</w:t></w:r></w:p></w:hdr>"#), true).unwrap();
    styles.restore_story(&mut story).expect("inherited namespaces stay on each detached property");
    let result = parse_xml(&story.render(None), true).unwrap();
    let paragraph = result.child("p").unwrap();
    for properties in [paragraph.child("r").unwrap().child("rPr").unwrap(), paragraph.child("pPr").unwrap().child("rPr").unwrap()] {
        let namespaces = properties.children().filter(|property| property.local == "ligatures").map(|property| property.namespace.as_str()).collect::<Vec<_>>();
        assert!(namespaces.contains(&DEFAULT_EXTENSION)); assert!(namespaces.contains(&STYLE_EXTENSION)); assert!(namespaces.contains(&STORY_EXTENSION));
    }
}

#[test]
fn merged_tabs_borders_and_paragraph_marks_keep_independent_child_scopes() {
    const FIRST: &str = "http://schemas.microsoft.com/office/word/2010/wordml";
    const SECOND: &str = "http://schemas.microsoft.com/office/word/2012/wordml";
    let mut styles = registry(&format!(r#"<w:styles xmlns:w="{WORD_NAMESPACE}"><w:docDefaults><w:pPrDefault><w:pPr><w:tabs xmlns:x="{FIRST}"><w:tab w:pos="720" w:val="left" x:custom="default"/></w:tabs><w:pBdr xmlns:x="{FIRST}"><w:top w:val="single" x:custom="default"/></w:pBdr><w:rPr xmlns:x="{FIRST}"><w:caps/><x:ligatures w:val="standard"/></w:rPr></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:pPr><w:tabs xmlns:x="{SECOND}"><w:tab w:pos="1440" w:val="right" x:custom="style"/></w:tabs><w:pBdr xmlns:x="{SECOND}"><w:bottom w:val="single" x:custom="style"/></w:pBdr><w:rPr xmlns:x="{SECOND}"><w:smallCaps/><x:ligatures w:val="contextual"/></w:rPr></w:pPr></w:style></w:styles>"#), None);
    let mut story = parse_xml(&format!(r#"<w:hdr xmlns:w="{WORD_NAMESPACE}"><w:p><w:r><w:t>paragraph</w:t></w:r></w:p></w:hdr>"#), true).unwrap();
    styles.restore_story(&mut story).unwrap();
    let result = parse_xml(&story.render(None), true).unwrap();
    let paragraph = result.child("p").unwrap().child("pPr").unwrap();
    for container in ["tabs", "pBdr"] {
        let children = paragraph.child(container).unwrap().children().collect::<Vec<_>>();
        assert_eq!(children.len(), 2);
        assert!(children[0].attributes.iter().any(|attribute| attribute.namespace == FIRST && attribute.local == "custom"));
        assert!(children[1].attributes.iter().any(|attribute| attribute.namespace == SECOND && attribute.local == "custom"));
    }
    let mark = paragraph.child("rPr").unwrap();
    assert!(mark.child("caps").is_some()); assert!(mark.child("smallCaps").is_some());
    let scopes = mark.children().filter(|child| child.local == "ligatures").map(|child| child.namespace.as_str()).collect::<Vec<_>>();
    assert!(scopes.contains(&FIRST)); assert!(scopes.contains(&SECOND));
}

#[test]
fn malformed_source_envelope_and_unmapped_paragraph_numbering_are_explicit_errors() {
    let current = json!({ "fontFamily": "Aptos", "fontSize": 9, "bold": false });
    let provenance = run_provenance("<w:rPr><w:caps/></w:rPr>", &current);
    let mut malformed = registry("<w:styles>", None);
    assert!(malformed.story_run_properties("", &current, &provenance).is_err());
    let mut missing = SourceStyles::new(&json!({ "metadata": { "source": { "officeSource": { "version": 1, "resourceErrors": ["declared styles part is missing"] } } } }), GENERATED.into()).unwrap();
    assert!(missing.story_run_properties("", &current, &provenance).unwrap_err().contains("declared styles part is missing"));
    let mut legal_absence = SourceStyles::new(&json!({ "metadata": { "source": { "officeSource": { "version": 1, "resourceErrors": [] } } } }), GENERATED.into()).unwrap();
    assert!(legal_absence.story_run_properties("", &current, &provenance).is_ok());
    let mut numbering = registry(&format!(r#"<w:styles xmlns:w="{WORD_NAMESPACE}"><w:docDefaults><w:pPrDefault><w:pPr><w:numPr><w:numId w:val="99"/></w:numPr></w:pPr></w:pPrDefault></w:docDefaults></w:styles>"#), None);
    let paragraph = json!({ "metadata": { "source": { "sourceFormat": source_format("<w:pPr/>", None) }, "officeStoryParagraphStyleBaseline": {} } });
    assert!(numbering.story_paragraph_properties("", &json!({}), &paragraph).unwrap_err().contains("numbering IDs"));
}
