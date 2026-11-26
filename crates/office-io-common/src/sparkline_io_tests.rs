//! Deferred draft contracts. None of these tests/checks/builds was executed in this batch.
use super::*;
use crate::conditional_format_xml::{self as xml, Element};

const SHEET: &str = include_str!("../../office-parser/tests/fixtures/sparklines/native-settings-sheet.xml");
const THEME: &str = include_str!("../../office-parser/tests/fixtures/sparklines/theme.xml");
fn groups() -> Vec<Element> { xml::parse(SHEET).unwrap().descendants("sparklineGroup") }
fn palette() -> Vec<String> { ["FF000000", "FFFFFFFF", "FF102030"].map(str::to_owned).into() }
fn decoded() -> NativeSparklineGroup {
    let theme = xml::parse(THEME).unwrap();
    decode_native_group(&groups()[0], "group-1", Some(&theme), Some(THEME), Some(&palette()), &mut Vec::new()).unwrap()
}

#[test]
fn references_preserve_foreign_sheet_quotes_absolute_flags_and_single_cell_sources() {
    assert_eq!(normalize_source_reference("='O''Brien 数据!2026'!$b$1:E$2").unwrap(), "'O''Brien 数据!2026'!$B$1:E$2");
    assert_eq!(source_formula("O'Brien", "$B1").unwrap(), "'O''Brien'!$B1");
    assert_eq!(normalize_source_reference("Data!D4:B2").unwrap(), "'Data'!D4:B2");
    assert_eq!(normalize_target("$XFD$1048576").as_deref(), Some("XFD1048576"));
    for source in ["A0", "A01", "XFE1", "A1048577", "A1:B2:C3", "'Open!A1", "A!B!C1", "A:A", "1:2"] {
        assert!(parse_source_reference(source).is_err(), "{source}");
    }
}

#[test]
fn native_omitted_settings_use_zero_and_point_weight_defaults_independently_of_editor_defaults() {
    let group = &groups()[1];
    let decoded = decode_native_group(group, "defaults", None, None, None, &mut Vec::new()).unwrap();
    assert_eq!(decoded.settings.empty_cells, "zero");
    assert_eq!(decoded.settings.line_width, Some(1.0));
    assert_eq!(decoded.settings.axis_options.as_ref().unwrap().visible, Some(false));
    assert!(decoded.settings.points.as_ref().unwrap().values().all(|point| point.visible == Some(false)));
}

#[test]
fn native_six_flags_are_independent_from_retained_disabled_marker_colors() {
    let decoded = decoded(); let spec = decoded.settings;
    assert_eq!(spec.negative_color, None);
    assert_eq!(spec.points.as_ref().unwrap()["negativePoint"].visible, Some(false));
    let color = spec.points.as_ref().unwrap()["negativePoint"].color.as_deref().unwrap();
    assert_eq!(css_argb(color).unwrap(), "800000FF");
    assert_eq!(spec.points.as_ref().unwrap()["highPoint"].visible, Some(true));
    assert_eq!(spec.points.as_ref().unwrap()["lastPoint"].visible, Some(false));
    assert_eq!(spec.points.as_ref().unwrap()["markersPoint"].visible, Some(true));
    assert_eq!(spec.office_source.as_ref().unwrap().resolved_colors.len(), 8);
}

#[test]
fn native_group_members_and_independent_axis_modes_survive_with_inactive_values() {
    let decoded = decoded();
    assert_eq!(decoded.members, [("G1".into(), "'O''Brien 数据!2026'!$B$1:E$1".into()), ("G2".into(), "'O''Brien 数据!2026'!$B$2:$E2".into())]);
    assert_eq!(decoded.settings.axis, "individual");
    assert_eq!(decoded.settings.extremum_min.as_ref().unwrap().mode, "group");
    assert_eq!(decoded.settings.extremum_min.as_ref().unwrap().value, Some(-99.25));
    assert_eq!(decoded.settings.extremum_max.as_ref().unwrap().mode, "custom");
    assert_eq!(decoded.settings.extremum_max.as_ref().unwrap().value, Some(9.123456789123456));
    assert!((decoded.settings.line_width.unwrap() - 2.0 / 3.0).abs() < 1e-14);
    assert!(decoded.settings.contain_hidden_cells.unwrap());
    assert!(decoded.settings.axis_options.as_ref().unwrap().reverse.unwrap());
}

#[test]
fn theme_tint_custom_indexed_palette_and_argb_alpha_resolve_without_flattening_the_source() {
    let spec = decoded().settings; let source = spec.office_source.as_ref().unwrap();
    assert_eq!(spec.color, "#FF8080");
    assert_eq!(spec.axis_options.as_ref().unwrap().color.as_deref(), Some("#102030"));
    assert!(source.group_xml.contains("theme=\"4\"")); assert!(source.group_xml.contains("tint=\"0.5\""));
    assert!(source.group_xml.contains("rgb=\"800000FF\"")); assert!(source.group_xml.contains("indexed=\"2\""));
    assert!(source.theme_xml.is_some()); assert_eq!(source.indexed_colors.as_ref().unwrap(), &palette());
    assert_eq!(css_argb("#12345680").unwrap(), "80123456");
}

#[test]
fn editing_series_color_replaces_only_that_native_binding_and_retains_date_axis() {
    let mut spec = decoded().settings; spec.color = "#010203".into();
    let native = native_group_for_spec(&spec).unwrap().unwrap();
    let series = native.children.iter().find(|node| node.is(xml::X14_NAMESPACE, "colorSeries")).unwrap();
    assert_eq!(series.attribute_ns("", "rgb"), Some("FF010203"));
    assert_eq!(series.attribute_ns("", "theme"), None); assert_eq!(series.attribute_ns("", "tint"), None);
    assert!(native.children.iter().any(|node| node.is(xml::XM_NAMESPACE, "f") && node.text == "'O''Brien 数据!2026'!$A$1:$D$1"));
    assert_eq!(native.attribute_ns("", "dateAxis"), Some("1"));
    assert_eq!(native.children.iter().find(|node| node.is(xml::X14_NAMESPACE, "colorNegative")).unwrap().attribute_ns("", "rgb"), Some("800000FF"));
}

#[test]
fn unedited_native_color_omissions_are_not_replaced_with_editor_fallback_colors() {
    let spec = decode_native_group(&groups()[1], "omitted", None, None, None, &mut Vec::new()).unwrap().settings;
    let native = native_group_for_spec(&spec).unwrap().unwrap();
    assert_eq!(native.children.iter().filter(|node| COLOR_NAMES.contains(&node.name.as_str())).count(), 1);
    assert_eq!(native.children[0].attribute_ns("", "rgb"), Some("400000FF"));
}

#[test]
fn signatures_ignore_aliases_attribute_order_default_spelling_and_member_order() {
    let original = groups()[2].render();
    let alias = original.replace("x14:", "s14:").replace("xmlns:x14", "xmlns:s14").replace("negative=\"1\"", "negative=\"true\"");
    let first = xml::parse(&original).unwrap(); let second = xml::parse(&alias).unwrap();
    assert_eq!(native_group_signature(&first, "K2").unwrap(), native_group_signature(&second, "K2").unwrap());
    let mut reordered = groups()[0].clone();
    let members = reordered.children.iter_mut().find(|node| node.is(xml::X14_NAMESPACE, "sparklines")).unwrap();
    members.children.reverse(); members.content = vec![xml::Content::Child(0), xml::Content::Child(1)];
    let theme = xml::parse(THEME).unwrap();
    assert_eq!(native_group_fingerprint(&groups()[0], Some(&theme), Some(&palette())).unwrap(), native_group_fingerprint(&reordered, Some(&theme), Some(&palette())).unwrap());
}

#[test]
fn changed_group_source_member_target_and_theme_colors_invalidate_native_signatures() {
    let original = groups()[0].clone(); let theme = xml::parse(THEME).unwrap();
    let before = native_group_fingerprint(&original, Some(&theme), Some(&palette())).unwrap();
    for changed in [original.render().replace("manualMax=\"9.123456789123456\"", "manualMax=\"12\""),
        original.render().replace("$B$1:E$1", "$C$1:E$1"), original.render().replace(">G2<", ">G3<")] {
        assert_ne!(before, native_group_fingerprint(&xml::parse(&changed).unwrap(), Some(&theme), Some(&palette())).unwrap());
    }
    let changed_theme = xml::parse(&THEME.replace("val=\"FF0000\"", "val=\"00FF00\"")).unwrap();
    assert_ne!(before, native_group_fingerprint(&original, Some(&changed_theme), Some(&palette())).unwrap());
}

#[test]
fn native_namespace_missing_resources_nonconforming_auto_and_duplicate_targets_are_rejected() {
    let group = groups()[0].clone(); let theme = xml::parse(THEME).unwrap();
    assert!(decode_native_group(&group, "bad", None, None, Some(&palette()), &mut Vec::new()).is_err());
    for changed in [group.render().replace("http://schemas.microsoft.com/office/spreadsheetml/2009/9/main", "urn:fake"),
        group.render().replace("theme=\"4\"", "theme=\"4\" auto=\"0\""), group.render().replace(">G2<", ">G1<")] {
        assert!(decode_native_group(&xml::parse(&changed).unwrap(), "bad", Some(&theme), Some(THEME), Some(&palette()), &mut Vec::new()).is_err());
    }
}

#[test]
fn canonical_dto_omits_absent_legacy_colors_and_preserves_full_office_source() {
    let spec = decoded().settings; let value = serde_json::to_value(&spec).unwrap();
    assert!(value.get("negativeColor").is_none()); assert!(value.get("lowPointColor").is_none());
    assert_eq!(value["points"]["negativePoint"]["visible"], false);
    assert_eq!(serde_json::from_value::<SparklineSpec>(value).unwrap(), spec);
    let mut source = spec.office_source.clone().unwrap(); source.version = 2; assert!(source.validate().is_err());
    source.version = 1; source.group_xml = " ".repeat(MAX_GROUP_SOURCE_BYTES + 1); assert!(source.validate().is_err());
}

#[test]
fn line_weight_and_pie_admission_do_not_discard_the_editable_configuration() {
    let mut spec = decoded().settings;
    for (pixels, points) in [(0.0, 0.0), (1.0, 0.75), (2112.0, 1584.0)] {
        spec.line_width = Some(pixels);
        let group = native_group_for_spec(&spec).unwrap().unwrap();
        assert_eq!(group.attribute_ns("", "lineWeight").unwrap().parse::<f64>().unwrap(), points);
    }
    spec.line_width = Some(2112.1); assert!(spec.validate().is_err());
    spec.line_width = Some(2.0); spec.sparkline_type = "pie".into(); spec.show_gradient = Some(true);
    assert!(native_group_for_spec(&spec).unwrap().is_none());
    assert!(serde_json::to_value(&spec).unwrap()["officeSource"].is_object());
}
