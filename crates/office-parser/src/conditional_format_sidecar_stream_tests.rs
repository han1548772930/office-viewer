//! Unexecuted binary/source contracts for the conditional-format streaming boundary.

use super::*;
use super::tests::{exported_rules, part_xml, reopened_rules, themed_rule, worksheet_xml};
use std::io::{Cursor, Write};

fn expression(id: &str) -> Value {
    json!({ "id": id, "type": "expression", "ranges": [{ "startRow": 0, "startCol": 0, "endRow": 3, "endCol": 0 }],
        "operator": "", "formula1": "=A1>0", "style": { "bold": true }, "priority": 17, "stopIfTrue": false })
}

fn rewrite(bytes: &[u8], replacements: &[(&str, Option<String>)]) -> Vec<u8> {
    let mut archive = ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut output = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options = zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);
    let mut seen = HashSet::new();
    for index in 0..archive.len() {
        let mut entry = archive.by_index(index).unwrap(); let name = entry.name().to_owned(); seen.insert(name.clone());
        if let Some((_, replacement)) = replacements.iter().find(|(part, _)| *part == name) {
            if let Some(content) = replacement { output.start_file(&name, options).unwrap(); output.write_all(content.as_bytes()).unwrap(); }
        } else { output.start_file(&name, options).unwrap(); std::io::copy(&mut entry, &mut output).unwrap(); }
    }
    for (name, content) in replacements {
        if !seen.contains(*name) && let Some(content) = content {
            output.start_file(*name, options).unwrap(); output.write_all(content.as_bytes()).unwrap();
        }
    }
    output.finish().unwrap().into_inner()
}

fn with_groups(bytes: &[u8], groups: &[String]) -> Vec<u8> {
    let mut root = source_xml::parse(&worksheet_xml(bytes)).unwrap();
    root.replace_children(|node| node.is_main("conditionalFormatting"), Vec::new());
    let mut xml = root.render();
    let insertion = xml.find("<extLst").unwrap_or_else(|| xml.find("</worksheet>").unwrap());
    xml.insert_str(insertion, &groups.join(""));
    rewrite(bytes, &[("xl/worksheets/sheet1.xml", Some(xml))])
}

fn group(priority: usize, payload: &str) -> String {
    format!(r#"<conditionalFormatting sqref="A1:A4"><cfRule type="expression" priority="{priority}"><formula>A1&gt;0</formula><extLst><ext uri="vendor"><v:payload xmlns:v="urn:vendor">{payload}</v:payload></ext></extLst></cfRule></conditionalFormatting>"#)
}

fn read_failure(bytes: &[u8]) -> String {
    let sidecar = ConditionalFormatSidecar::read(bytes);
    assert!(sidecar.by_sheet.is_empty(), "a failed read must not publish any sheet rules");
    sidecar.failure.expect("declared source failure must be observable")
}

fn large_data() -> String {
    let mut xml = String::new();
    for row in 1..=90_000 { xml.push_str(&format!(r#"<row r="{row}"><c r="A{row}"><v>7</v></c></row>"#)); }
    let text = "x".repeat(24_000);
    for row in 90_001..=90_900 { xml.push_str(&format!(r#"<row r="{row}"><c r="A{row}" t="inlineStr"><is><t>{text}</t></is></c></row>"#)); }
    xml
}

fn replace_sheet_data(mut xml: String, data: &str) -> String {
    let start = xml.find("<sheetData>").unwrap() + "<sheetData>".len();
    let end = xml.find("</sheetData>").unwrap(); xml.replace_range(start..end, data);
    if let Some(start) = xml.find("<dimension ") {
        let end = xml[start..].find("/>").unwrap() + start + 2;
        xml.replace_range(start..end, "<dimension ref=\"A1:A90900\"/>");
    }
    xml
}

fn workbook_with_two_rules() -> Vec<u8> {
    office_writer::export_xlsx_bytes(&json!({ "kind": "workbook", "name": "Atomic CF sources", "sheets": [
        { "id": "first", "name": "First", "cells": { "0:0": { "row": 0, "col": 0, "value": 7 } }, "conditionalFormats": [expression("first-rule")] },
        { "id": "second", "name": "Second", "cells": { "0:0": { "row": 0, "col": 0, "value": 9 } }, "conditionalFormats": [expression("second-rule")] }
    ] }).to_string()).unwrap().0
}

#[test]
fn large_worksheet_keeps_late_base_and_linked_x14_sources_through_native_style_edit_and_reopen() {
    let original = json!({ "id": "large-linked-bar", "type": "dataBar", "ranges": [{ "startRow": 0, "startCol": 0, "endRow": 3, "endCol": 0 }],
        "operator": "", "formula1": "", "style": { "bold": true }, "priority": 17, "stopIfTrue": false,
        "config": { "min": { "type": "autoMin" }, "max": { "type": "autoMax" }, "positiveColor": "#336699", "nativeColor": "#AA3344" } });
    let bytes = exported_rules(json!([original]));
    let inflated_xml = replace_sheet_data(worksheet_xml(&bytes), &large_data());
    assert!(inflated_xml.len() > source_xml::XML_FRAGMENT_BYTE_LIMIT);
    let inflated = rewrite(&bytes, &[("xl/worksheets/sheet1.xml", Some(inflated_xml))]);
    let mut sidecar = ConditionalFormatSidecar::read(&inflated); assert!(sidecar.failure.is_none(), "{:?}", sidecar.failure);
    let mut rules = sidecar.take(0); assert_eq!(rules.len(), 1);
    assert_eq!(rules[0]["id"], "large-linked-bar"); assert_eq!(rules[0]["priority"], 17);
    assert!(rules[0]["officeSource"]["x14Xml"].as_str().unwrap().contains("autoMin"));
    assert!(rules[0]["officeSource"]["differentialFormatXml"].is_string());
    let before_link = linked_id(&source_xml::parse(rules[0]["officeSource"]["ruleXml"].as_str().unwrap()).unwrap());
    let parsed_large = reopened_rules(&inflated); assert_eq!(parsed_large[0]["id"], rules[0]["id"]);
    rules[0]["style"]["bold"] = json!(false);
    let reopened = reopened_rules(&exported_rules(json!(rules)));
    assert_eq!(reopened[0]["id"], "large-linked-bar"); assert_eq!(reopened[0]["priority"], 17); assert_eq!(reopened[0]["style"]["bold"], false);
    assert_eq!(linked_id(&source_xml::parse(reopened[0]["officeSource"]["ruleXml"].as_str().unwrap()).unwrap()), before_link);
    assert_eq!(reopened[0]["config"]["min"]["type"], "autoMin"); assert_eq!(reopened[0]["config"]["max"]["type"], "autoMax");
}

#[test]
fn large_styles_preserve_the_referenced_dxf_original_index_without_retaining_unrelated_records() {
    let bytes = exported_rules(json!([expression("indexed-style")]));
    let old_styles = source_xml::parse(&part_xml(&bytes, "xl/styles.xml")).unwrap();
    let dxf = old_styles.child("dxfs").unwrap().children[0].clone();
    let xfs = r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#.repeat(350_000);
    let unused = "<dxf><font><i/></font></dxf>".repeat(90_000);
    let styles = format!(r#"<styleSheet xmlns="{}"><cellXfs count="350000">{xfs}</cellXfs><dxfs count="90001">{unused}{}</dxfs></styleSheet>"#, source_xml::MAIN_NAMESPACE, dxf.render());
    assert!(styles.len() > source_xml::XML_FRAGMENT_BYTE_LIMIT);
    let worksheet = worksheet_xml(&bytes).replace("dxfId=\"0\"", "dxfId=\"90000\"");
    let edited = rewrite(&bytes, &[("xl/styles.xml", Some(styles)), ("xl/worksheets/sheet1.xml", Some(worksheet))]);
    let mut sidecar = ConditionalFormatSidecar::read(&edited); assert!(sidecar.failure.is_none(), "{:?}", sidecar.failure);
    let rules = sidecar.take(0); assert_eq!(rules.len(), 1); assert_eq!(rules[0]["style"]["bold"], true);
    assert!(rules[0]["officeSource"]["ruleXml"].as_str().unwrap().contains("dxfId=\"90000\""));
    let retained = source_xml::parse(rules[0]["officeSource"]["differentialFormatXml"].as_str().unwrap()).unwrap();
    assert_eq!(retained.fingerprint(), dxf.fingerprint());
}

#[test]
fn no_cf_does_not_read_unrelated_broken_or_missing_styles_and_theme_parts() {
    let bytes = exported_rules(json!([]));
    let rels = part_xml(&bytes, "xl/_rels/workbook.xml.rels").replace("</Relationships>", &format!(r#"<Relationship Id="rIdUnusedTheme" Type="{REL_NAMESPACE}/theme" Target="theme/unused.xml"/></Relationships>"#));
    for style in [None, Some("<broken-style".into())] {
        let edited = rewrite(&bytes, &[("xl/styles.xml", style), ("xl/_rels/workbook.xml.rels", Some(rels.clone())), ("xl/theme/unused.xml", Some("<broken-theme".into()))]);
        let mut sidecar = ConditionalFormatSidecar::read(&edited);
        assert!(sidecar.failure.is_none(), "{:?}", sidecar.failure); assert!(sidecar.warnings.is_empty()); assert!(sidecar.take(0).is_empty());
    }
}

#[test]
fn missing_or_malformed_referenced_differential_styles_fail_without_publishing_partial_rules() {
    let bytes = exported_rules(json!([expression("requires-style")]));
    for styles in [None, Some("<styleSheet><dxfs><dxf>".into()), Some("<styleSheet><dxfs/></styleSheet>".into())] {
        let edited = rewrite(&bytes, &[("xl/styles.xml", styles)]);
        let failure = read_failure(&edited);
        assert!(failure.starts_with("CF_SOURCE_REQUIRED_RESOURCE") || failure.starts_with("CF_SOURCE_XML_INVALID"), "{failure}");
        assert!(crate::parser::parse_office_bytes(&edited, "missing-cf-style.xlsx").is_err());
    }
}

#[test]
fn invalid_and_extreme_dxf_indices_reject_without_a_sparse_style_vector() {
    let bytes = exported_rules(json!([expression("invalid-style-index")]));
    for (index, class) in [("-1", "CF_SOURCE_XML_INVALID"), ("not-index", "CF_SOURCE_XML_INVALID"),
        ("4294967296", "CF_SOURCE_XML_INVALID"), ("4294967295", "CF_SOURCE_REQUIRED_RESOURCE")] {
        let xml = worksheet_xml(&bytes).replace("dxfId=\"0\"", &format!("dxfId=\"{index}\""));
        let edited = rewrite(&bytes, &[("xl/worksheets/sheet1.xml", Some(xml))]);
        assert!(read_failure(&edited).starts_with(class), "{index}");
    }
}

#[test]
fn a_required_theme_cannot_be_missing_malformed_or_in_the_wrong_namespace() {
    let bytes = exported_rules(json!([themed_rule()]));
    for theme in [None, Some("<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">".into()),
        Some("<theme xmlns=\"urn:wrong\"/>".into())] {
        let edited = rewrite(&bytes, &[("xl/theme/theme1.xml", theme)]);
        let failure = read_failure(&edited);
        assert!(failure.starts_with("CF_SOURCE_REQUIRED_RESOURCE") || failure.starts_with("CF_SOURCE_XML_INVALID"), "{failure}");
        assert!(crate::parser::parse_office_bytes(&edited, "required-cf-theme.xlsx").is_err());
    }
}

#[test]
fn rules_without_theme_references_do_not_read_a_damaged_unrelated_theme() {
    let bytes = with_groups(&exported_rules(json!([expression("rgb-only")])), &[group(1, r#"<v:n theme="vendor-only"/>"#)]);
    let rels = part_xml(&bytes, "xl/_rels/workbook.xml.rels").replace("</Relationships>", &format!(r#"<Relationship Id="rIdUnusedTheme" Type="{REL_NAMESPACE}/theme" Target="theme/unused.xml"/></Relationships>"#));
    let edited = rewrite(&bytes, &[("xl/_rels/workbook.xml.rels", Some(rels)), ("xl/theme/unused.xml", Some("<broken-theme".into()))]);
    let mut sidecar = ConditionalFormatSidecar::read(&edited); assert!(sidecar.failure.is_none(), "{:?}", sidecar.failure);
    assert_eq!(sidecar.take(0).len(), 1);
}

#[test]
fn a_later_sheet_failure_discards_complete_cf_sources_from_earlier_sheets() {
    let bytes = workbook_with_two_rules();
    let xml = part_xml(&bytes, "xl/worksheets/sheet2.xml").replace("sqref=\"A1:A4\"", "sqref=\"A0\"");
    let edited = rewrite(&bytes, &[("xl/worksheets/sheet2.xml", Some(xml))]);
    assert!(read_failure(&edited).starts_with("CF_SOURCE_XML_INVALID"));
}

#[test]
fn invalid_sqref_is_an_import_failure_and_not_an_empty_successful_rule_list() {
    let bytes = exported_rules(json!([expression("invalid-scope")]));
    let xml = worksheet_xml(&bytes).replace("sqref=\"A1:A4\"", "sqref=\"A0 A2:A4\"");
    let edited = rewrite(&bytes, &[("xl/worksheets/sheet1.xml", Some(xml))]);
    assert!(read_failure(&edited).starts_with("CF_SOURCE_XML_INVALID"));
    let failure = crate::parser::parse_office_bytes(&edited, "invalid-cf-scope.xlsx").unwrap_err();
    assert!(matches!(&failure.error.code, crate::model::ParseErrorCode::InvalidOfficeFile));
    assert!(failure.error.details.as_deref().unwrap().starts_with("CF_SOURCE_XML_INVALID"));
}

#[test]
fn valid_cf_before_a_corrupt_xml_suffix_still_rejects_the_complete_worksheet() {
    let bytes = exported_rules(json!([expression("early-valid")]));
    let xml = worksheet_xml(&bytes).replace("</worksheet>", "<sheetData><row></worksheet>");
    let edited = rewrite(&bytes, &[("xl/worksheets/sheet1.xml", Some(xml))]);
    assert!(read_failure(&edited).starts_with("CF_SOURCE_XML_INVALID"));
    assert!(crate::parser::parse_office_bytes(&edited, "corrupt-cf-suffix.xlsx").is_err());
}

#[test]
fn cf_byte_depth_and_element_limits_do_not_reclassify_as_valid_empty_metadata() {
    let original = exported_rules(json!([expression("fragment-limits")]));
    let bytes = with_groups(&original, &[group(1, &"x".repeat(2048))]);
    for part in [XmlSelectionLimits { fragment_bytes: 1024, ..XmlSelectionLimits::default() },
        XmlSelectionLimits { fragment_depth: 2, ..XmlSelectionLimits::default() },
        XmlSelectionLimits { fragment_elements: 2, ..XmlSelectionLimits::default() }] {
        let failure = read_sidecar_with_limits(&bytes, SidecarLimits { part, ..SidecarLimits::default() }).unwrap_err();
        assert!(failure.starts_with("CF_SOURCE_FRAGMENT_LIMIT"), "{failure}");
    }
}

#[test]
fn many_small_sources_spend_part_and_cross_sheet_workbook_resource_budgets_atomically() {
    let original = exported_rules(json!([expression("part-budget")]));
    let bytes = with_groups(&original, &[group(1, &"x".repeat(2048)), group(2, &"y".repeat(2048))]);
    let xml = worksheet_xml(&bytes);
    let selected = source_xml::extract(Cursor::new(xml.as_bytes()), XmlSelectionLimits::default(), worksheet_group).unwrap();
    let first_bytes = selected.fragments[0].rendered_size();
    let failure = read_sidecar_with_limits(&bytes, SidecarLimits { part: XmlSelectionLimits { total_bytes: first_bytes, ..XmlSelectionLimits::default() },
        ..SidecarLimits::default() }).unwrap_err();
    assert!(failure.starts_with("CF_SOURCE_TOTAL_LIMIT"));
    let original = workbook_with_two_rules();
    let first = with_groups(&original, &[group(1, &"x".repeat(2048))]);
    let second_xml = part_xml(&first, "xl/worksheets/sheet2.xml");
    let second = second_xml.replace("</sheetData>", &format!("</sheetData>{}", group(2, &"y".repeat(2048))));
    let both = rewrite(&first, &[("xl/worksheets/sheet2.xml", Some(second))]);
    let failure = read_sidecar_with_limits(&both, SidecarLimits { resource_bytes: first_bytes + 512, ..SidecarLimits::default() }).unwrap_err();
    assert!(failure.starts_with("CF_SOURCE_TOTAL_LIMIT")); assert!(failure.contains("sheet2.xml"), "{failure}");
    assert!(ConditionalFormatSidecar::read(&both).failure.is_none());
}

#[test]
fn duplicated_theme_and_dxf_provenance_spends_the_serialized_model_budget() {
    let first = themed_rule(); let mut second = first.clone(); second["id"] = json!("copied-rich-source");
    let bytes = exported_rules(json!([first, second]));
    let mut valid = ConditionalFormatSidecar::read(&bytes); assert!(valid.failure.is_none());
    let rules = valid.take(0); assert_eq!(rules.len(), 2);
    let first_size = serde_json::to_vec(&rules[0]).unwrap().len();
    let failure = read_sidecar_with_limits(&bytes, SidecarLimits { model_bytes: first_size + 2, ..SidecarLimits::default() }).unwrap_err();
    assert!(failure.starts_with("CF_SOURCE_TOTAL_LIMIT"));
}

#[test]
fn native_fragment_limit_errors_have_parser_failure_code_and_prefixed_details() {
    let original = exported_rules(json!([expression("native-fragment-limit")]));
    let bytes = with_groups(&original, &[group(1, &"x".repeat(source_xml::XML_FRAGMENT_BYTE_LIMIT))]);
    assert!(read_failure(&bytes).starts_with("CF_SOURCE_FRAGMENT_LIMIT"));
    let failure = crate::parser::parse_office_bytes(&bytes, "cf-fragment-limit.xlsx").unwrap_err();
    assert!(matches!(&failure.error.code, crate::model::ParseErrorCode::ParserFailure));
    assert!(failure.error.details.as_deref().unwrap().starts_with("CF_SOURCE_FRAGMENT_LIMIT"));
}

#[test]
fn native_part_element_totals_have_parser_failure_code_and_never_return_partial_rules() {
    let original = exported_rules(json!([expression("native-total-limit")]));
    let payload = "<v:n/>".repeat(175_000);
    let groups = (1..=6).map(|priority| group(priority, &payload)).collect::<Vec<_>>();
    let bytes = with_groups(&original, &groups);
    assert!(read_failure(&bytes).starts_with("CF_SOURCE_TOTAL_LIMIT"));
    let failure = crate::parser::parse_office_bytes(&bytes, "cf-total-limit.xlsx").unwrap_err();
    assert!(matches!(&failure.error.code, crate::model::ParseErrorCode::ParserFailure));
    assert!(failure.error.details.as_deref().unwrap().starts_with("CF_SOURCE_TOTAL_LIMIT"));
}
