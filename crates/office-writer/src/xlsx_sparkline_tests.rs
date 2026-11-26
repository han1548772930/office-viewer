//! Deferred draft contracts. No checks or tests were run in this implementation batch.
use super::*;
use serde_json::json;

fn spec(source: &str) -> SparklineSpec {
    serde_json::from_value(json!({ "type": "line", "sourceRange": source, "color": "#246FCE",
        "emptyCells": "connect", "axis": "individual", "groupId": "authored-group", "themeType": "custom",
        "points": { "highPoint": { "visible": true, "color": "#CC2244" },
            "firstPoint": { "visible": true, "color": "#102030" }, "lowPoint": { "visible": false, "color": "#556677" },
            "lastPoint": { "visible": false, "color": "#334455" }, "negativePoint": { "visible": false, "color": "rgba(0,0,255,0.5)" },
            "markersPoint": { "visible": true, "color": "#AABBCC" } },
        "axisOptions": { "visible": true, "reverse": true, "color": "#010203" },
        "extremumMin": { "type": "group", "value": -99.25 }, "extremumMax": { "type": "custom", "value": 9.123456789123456 },
        "nonNumericCells": "zero", "containHiddenCells": true, "lineWidth": 2.25, "showGradient": true, "showRadius": true })).unwrap()
}
fn cells(entries: &[(&str, SparklineSpec)]) -> serde_json::Map<String, Value> {
    entries.iter().map(|(target, spec)| ((*target).into(), json!({ "sparkline": spec }))).collect()
}
fn native_groups(root: &Element) -> &[Element] {
    &root.children.iter().find(|ext| ext.attribute_ns("", "uri") == Some(io::NATIVE_EXTENSION_URI))
        .unwrap().child("sparklineGroups").unwrap().children
}
fn local_records(root: &Element) -> Vec<SparklineModelRecord> {
    root.children.iter().find(|ext| ext.attribute_ns("", "uri") == Some(io::MODEL_EXTENSION_URI))
        .unwrap().child("sparklines").unwrap().children.iter()
        .map(|node| serde_json::from_str(&node.text).unwrap()).collect()
}
fn workbook(entries: &[(&str, SparklineSpec)]) -> super::super::WorkbookInput {
    serde_json::from_value(json!({ "kind": "workbook", "sheets": [{ "id": "sheet", "name": "Data", "cells": cells(entries) }] })).unwrap()
}
fn retained(color: &str, theme: Option<&str>, palette: Option<Vec<String>>) -> SparklineSpec {
    let group = source_xml::parse(&format!("<x14:sparklineGroup xmlns:x14=\"{}\" xmlns:xm=\"{}\"><x14:colorSeries {color}/><x14:sparklines><x14:sparkline><xm:f>Data!B1:E1</xm:f><xm:sqref>G1</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup>", source_xml::X14_NAMESPACE, source_xml::XM_NAMESPACE)).unwrap();
    let parsed_theme = theme.map(source_xml::parse).transpose().unwrap();
    io::decode_native_group(&group, "retained-group", parsed_theme.as_ref(), theme, palette.as_deref(), &mut Vec::new()).unwrap().settings
}

#[test]
fn one_group_emits_ordered_native_colors_exact_sources_and_independent_extrema() {
    let input = cells(&[("G1", spec("'O''Brien 数据!2026'!$B$1:E$1")), ("G2", spec("B2:E2"))]);
    let mut warnings = Vec::new(); let root = source_xml::parse(&write_xml("Data", &input, &mut warnings).unwrap()).unwrap();
    let groups = native_groups(&root); assert_eq!(groups.len(), 1);
    let group = &groups[0]; assert_eq!(group.attribute_ns("", "lineWeight"), Some("1.6875"));
    assert_eq!(group.attribute_ns("", "minAxisType"), Some("group")); assert_eq!(group.attribute_ns("", "manualMin"), Some("-99.25"));
    assert_eq!(group.attribute_ns("", "maxAxisType"), Some("custom")); assert_eq!(group.attribute_ns("", "manualMax"), Some("9.123456789123456"));
    assert_eq!(group.attribute_ns("", "negative"), Some("0")); assert_eq!(group.attribute_ns("", "markers"), Some("1"));
    assert_eq!(group.child("colorNegative").unwrap().attribute_ns("", "rgb"), Some("800000FF"));
    assert_eq!(group.children[..8].iter().map(|node| node.name.as_str()).collect::<Vec<_>>(), io::COLOR_NAMES);
    let members = &group.child("sparklines").unwrap().children;
    assert_eq!(members[0].child("f").unwrap().text, "'O''Brien 数据!2026'!$B$1:E$1");
    assert_eq!(members[1].child("f").unwrap().text, "'Data'!B2:E2");
    let records = local_records(&root); assert_eq!(records[0].spec, spec("'O''Brien 数据!2026'!$B$1:E$1"));
    assert!(records.iter().all(|record| record.native_signature.as_ref().is_some_and(|signature| signature.len() == 16)));
    assert!(warnings.iter().any(|warning| warning.code == "SPARKLINE_LOCAL_FEATURES"));
}

#[test]
fn pie_is_editable_local_metadata_without_fabricated_native_content() {
    let mut pie = spec("B1:E1"); pie.sparkline_type = "pie".into();
    let mut warnings = Vec::new(); let root = source_xml::parse(&write_xml("Data", &cells(&[("G1", pie.clone())]), &mut warnings).unwrap()).unwrap();
    assert!(!root.children.iter().any(|ext| ext.attribute_ns("", "uri") == Some(io::NATIVE_EXTENSION_URI)));
    let records = local_records(&root); assert_eq!(records.len(), 1); assert_eq!(records[0].native_signature, None); assert_eq!(records[0].spec, pie);
    assert!(warnings.iter().any(|warning| warning.code == "SPARKLINE_LOCAL_ONLY"));
}

#[test]
fn source_templates_with_incompatible_extra_metadata_split_only_the_native_projection() {
    let first = retained("rgb=\"FF246FCE\"", None, None); let mut second = first.clone(); second.source_range = "B2:E2".into();
    second.office_source.as_mut().unwrap().group_xml = second.office_source.as_ref().unwrap().group_xml.replace("<x14:sparklineGroup", "<x14:sparklineGroup futurePolicy=\"keep\"");
    let mut warnings = Vec::new(); let root = source_xml::parse(&write_xml("Data", &cells(&[("G1", first), ("G2", second)]), &mut warnings).unwrap()).unwrap();
    assert_eq!(native_groups(&root).len(), 2);
    assert!(local_records(&root).iter().all(|record| record.spec.group_id.as_deref() == Some("retained-group")));
    assert!(warnings.iter().any(|warning| warning.code == "SPARKLINE_NATIVE_GROUP_SPLIT"));
}

#[test]
fn unchanged_theme_tint_and_indexed_bindings_remain_author_colors() {
    let theme = include_str!("../../office-parser/tests/fixtures/sparklines/theme.xml");
    let first = retained("theme=\"4\" tint=\"0.5\"", Some(theme), None);
    let second = retained("indexed=\"2\"", None, Some(vec!["FF000000".into(), "FFFFFFFF".into(), "FF102030".into()]));
    let root = source_xml::parse(&write_xml("Data", &cells(&[("G1", first), ("G2", second)]), &mut Vec::new()).unwrap()).unwrap();
    assert!(native_groups(&root).iter().any(|group| group.child("colorSeries").unwrap().attribute_ns("", "theme") == Some("4")));
    assert!(native_groups(&root).iter().any(|group| group.child("colorSeries").unwrap().attribute_ns("", "indexed") == Some("2")));
}

#[test]
fn conflicting_workbook_themes_fail_before_any_package_is_returned() {
    let theme = include_str!("../../office-parser/tests/fixtures/sparklines/theme.xml"); let changed = theme.replace("val=\"FF0000\"", "val=\"00FF00\"");
    let first = retained("theme=\"4\"", Some(theme), None); let second = retained("theme=\"4\"", Some(&changed), None);
    assert!(merge_theme(&workbook(&[("G1", first), ("G2", second)]), None).unwrap_err().message.contains("SPARKLINE_THEME_CONFLICT"));
    let first = retained("theme=\"4\"", Some(theme), None);
    assert!(merge_theme(&workbook(&[("G1", first)]), Some(changed)).unwrap_err().message.contains("SPARKLINE_THEME_CONFLICT"));
}

#[test]
fn custom_palette_is_written_once_and_conflicting_or_changed_default_bindings_fail() {
    let palette = vec!["FF000000".into(), "FFFFFFFF".into(), "FF102030".into()];
    let first = retained("indexed=\"2\"", None, Some(palette.clone()));
    let styles = format!("<styleSheet xmlns=\"{}\"><fonts count=\"0\"/></styleSheet>", source_xml::MAIN_NAMESPACE);
    let output = apply_indexed_colors(&workbook(&[("G1", first.clone())]), styles.clone()).unwrap();
    let root = source_xml::parse(&output).unwrap(); assert_eq!(root.child("colors").unwrap().child("indexedColors").unwrap().children.len(), 3);
    let mut second = first.clone(); second.office_source.as_mut().unwrap().indexed_colors.as_mut().unwrap()[2] = "FF445566".into();
    assert!(apply_indexed_colors(&workbook(&[("G1", first.clone()), ("G2", second)]), styles.clone()).unwrap_err().message.contains("SPARKLINE_PALETTE_CONFLICT"));
    let default = retained("indexed=\"2\"", None, None);
    assert!(apply_indexed_colors(&workbook(&[("G1", first), ("G2", default)]), styles).unwrap_err().message.contains("SPARKLINE_PALETTE_CONFLICT"));
}

#[test]
fn noninteger_and_zero_widths_roundtrip_in_the_local_author_configuration() {
    for width in [0.0, 2.25, 7.5, 2112.0] {
        let mut input = spec("B1:E1"); input.line_width = Some(width);
        let root = source_xml::parse(&write_xml("Data", &cells(&[("G1", input)]), &mut Vec::new()).unwrap()).unwrap();
        assert_eq!(native_groups(&root)[0].attribute_ns("", "lineWeight").unwrap().parse::<f64>().unwrap(), width * 0.75);
        assert_eq!(local_records(&root)[0].spec.line_width, Some(width));
    }
}

#[test]
fn invalid_target_duplicate_canonical_target_and_bad_settings_fail_without_partial_xml() {
    assert!(write_xml("Data", &cells(&[("A0", spec("B1:E1"))]), &mut Vec::new()).is_err());
    assert!(write_xml("Data", &cells(&[("G1", spec("B1:E1")), ("$G$1", spec("B2:E2"))]), &mut Vec::new()).is_err());
    let mut invalid_spec = spec("B1:E1"); invalid_spec.line_width = Some(-1.0);
    assert!(write_xml("Data", &cells(&[("G1", spec("B1:E1")), ("G2", invalid_spec)]), &mut Vec::new()).is_err());
}
