//! Unexecuted native table resource contracts for unified validation.
use serde_json::json;
use super::{SourceStyles, WORD_NAMESPACE, parse_xml};

fn registry(styles: &str, theme: Option<&str>) -> SourceStyles {
    let model = json!({"metadata":{"source":{"officeSource":{"version":1,"styles":{"xml":styles},"theme":theme.map(|xml| json!({"xml":xml}))}}}});
    SourceStyles::new(&model,format!("<w:styles xmlns:w=\"{WORD_NAMESPACE}\"><w:style w:type=\"table\" w:styleId=\"Grid\"/><w:style w:type=\"character\" w:styleId=\"Strong\"/></w:styles>")).unwrap()
}

#[test]
fn table_conditional_styles_dependencies_and_conflicting_ids_use_the_shared_registry() {
    let xml = format!("<w:styles xmlns:w=\"{WORD_NAMESPACE}\"><w:style w:type=\"table\" w:styleId=\"Base\"><w:tblPr><w:tblCellMar><w:left w:w=\"120\" w:type=\"dxa\"/></w:tblCellMar></w:tblPr></w:style><w:style w:type=\"character\" w:styleId=\"Strong\"><w:rPr><w:b/></w:rPr></w:style><w:style w:type=\"table\" w:styleId=\"Grid\"><w:basedOn w:val=\"Base\"/><w:tblStylePr w:type=\"firstRow\"><w:rPr><w:rStyle w:val=\"Strong\"/></w:rPr><w:tcPr><w:shd w:fill=\"123456\"/></w:tcPr></w:tblStylePr></w:style></w:styles>");
    let mut styles = registry(&xml,None);
    let mut properties = parse_xml(&format!("<w:tblPr xmlns:w=\"{WORD_NAMESPACE}\"><w:tblStyle w:val=\"Grid\"/></w:tblPr>"),true).unwrap();
    styles.restore_table_properties(&mut properties).unwrap();
    let mapped = properties.child("tblStyle").unwrap().attribute("val").unwrap();
    assert_ne!(mapped,"Grid");
    let output = parse_xml(&styles.styles_xml().unwrap(),true).unwrap();
    let table = output.children().find(|style| style.attribute("styleId") == Some(mapped)).unwrap();
    assert_eq!(table.child("basedOn").unwrap().attribute("val"),Some("Base"));
    let conditional = table.child("tblStylePr").unwrap();
    assert_eq!(conditional.child("tcPr").unwrap().child("shd").unwrap().attribute("fill"),Some("123456"));
    let strong = conditional.child("rPr").unwrap().child("rStyle").unwrap().attribute("val").unwrap();
    assert_ne!(strong,"Strong");
    assert!(output.children().any(|style| style.attribute("styleId") == Some(strong) && style.child("rPr").unwrap().child("b").is_some()));
}

#[test]
fn table_theme_references_require_a_retained_real_theme() {
    let theme = "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:themeElements><a:clrScheme name=\"Table colors\"><a:accent1><a:srgbClr val=\"2468AB\"/></a:accent1></a:clrScheme></a:themeElements></a:theme>";
    let source = format!("<w:styles xmlns:w=\"{WORD_NAMESPACE}\"/>");
    let property = format!("<w:tcPr xmlns:w=\"{WORD_NAMESPACE}\"><w:shd w:themeFill=\"accent1\"/></w:tcPr>");
    let mut missing = registry(&source,None);
    assert!(missing.restore_table_properties(&mut parse_xml(&property,true).unwrap()).is_err());
    let mut retained = registry(&source,Some(theme));
    retained.restore_table_properties(&mut parse_xml(&property,true).unwrap()).unwrap();
    assert!(retained.theme_xml().unwrap().contains("2468AB"));
}

#[test]
fn cyclic_or_missing_numbering_source_table_styles_are_rejected_with_the_source_intact() {
    for body in [
        "<w:style w:type=\"table\" w:styleId=\"A\"><w:basedOn w:val=\"B\"/></w:style><w:style w:type=\"table\" w:styleId=\"B\"><w:basedOn w:val=\"A\"/></w:style>",
        "<w:style w:type=\"table\" w:styleId=\"A\"><w:tblStylePr w:type=\"firstRow\"><w:pPr><w:numPr><w:numId w:val=\"8\"/></w:numPr></w:pPr></w:tblStylePr></w:style>",
    ] {
        let source = format!("<w:styles xmlns:w=\"{WORD_NAMESPACE}\">{body}</w:styles>");
        let mut styles = registry(&source,None);
        let mut property = parse_xml(&format!("<w:tblPr xmlns:w=\"{WORD_NAMESPACE}\"><w:tblStyle w:val=\"A\"/></w:tblPr>"),true).unwrap();
        assert!(styles.restore_table_properties(&mut property).is_err());
        assert_eq!(property.child("tblStyle").unwrap().attribute("val"),Some("A"));
    }
}
