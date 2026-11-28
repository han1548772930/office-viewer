//! Deferred binary reopen contracts. These fixtures do not execute Office
//! recalculation; native cache edits below explicitly model its persisted output.

use std::io::{Cursor, Read, Write};
use office_io_common::sheet_formula_calendar_source::{self as source, FormulaCalendarSource};
use serde_json::{Value, json};
use crate::parser::parse_office_bytes;

fn fixture() -> Value {
    json!({ "kind": "workbook", "date1904": true, "activeSheetId": "data", "sheets": [
        { "id": "data", "name": "Data", "cells": {
            "0:0": { "row": 0, "col": 0, "value": -1401.5, "formula": null, "dataType": "number",
                "style": { "numberFormat": "yyyy-mm-dd hh:mm" },
                "metadata": { "chartDependencyDateSource": { "value": 60.5, "date1904": false, "targetValue": -1401.5, "targetDate1904": true, "targetFormula": null } } },
            "0:1": { "row": 0, "col": 1, "value": 29, "formula": "=DAY(A1)", "dataType": "formula", "cachedDataType": "number",
                "metadata": { "chartDependencyFormulaSource": { "version": 1, "formula": "=DAY(A1)", "date1904": false, "dateResult": false, "aliasSheetId": "data" } } },
            "4:0": { "row": 4, "col": 0, "value": "Amount", "dataType": "string" },
            "5:0": { "row": 5, "col": 0, "value": 12, "dataType": "number" }
        }, "metadata": { "chartDependencyImport": { "version": 1, "formulaAliases": {
            "sheets": { "original": "Data" }, "names": {}, "tables": { "oldtable": "DataTable" }
        } } } },
        { "id": "other", "name": "Other", "cells": {} }
    ], "tables": [{ "id": "stable-table", "name": "DataTable", "sheetId": "data",
        "range": { "startRow": 4, "startCol": 0, "endRow": 5, "endCol": 0 },
        "columns": [{ "id": "amount", "name": "Amount" }], "showHeaderRow": true, "showTotalsRow": false, "styleName": "TableStyleMedium2" }] })
}
fn export(snapshot: &Value) -> Vec<u8> { office_writer::export_xlsx_bytes(&snapshot.to_string()).unwrap().0 }
fn parsed(bytes: &[u8]) -> Value { serde_json::to_value(parse_office_bytes(bytes, "calendar-roundtrip.xlsx").unwrap()).unwrap() }
fn part(bytes: &[u8], name: &str) -> String {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).unwrap(); let mut text = String::new();
    archive.by_name(name).unwrap().read_to_string(&mut text).unwrap(); text
}
fn manifest(bytes: &[u8]) -> FormulaCalendarSource { source::from_xml(&part(bytes, source::PART_NAME)).unwrap() }
fn rewrite(bytes: &[u8], mut edit: impl FnMut(String, Vec<u8>) -> Option<(String, Vec<u8>)>) -> Vec<u8> {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options = zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);
    for index in 0..archive.len() {
        let mut entry = archive.by_index(index).unwrap(); let name = entry.name().unwrap().into_owned();
        let mut data = Vec::new(); entry.read_to_end(&mut data).unwrap();
        if let Some((name, data)) = edit(name, data) { writer.start_file(name, options).unwrap(); writer.write_all(&data).unwrap(); }
    }
    writer.finish().unwrap().into_inner()
}
fn edit_xml(bytes: &[u8], target: &str, edit: impl Fn(String) -> String) -> Vec<u8> {
    rewrite(bytes, |name, data| { let data = if name == target { edit(String::from_utf8(data).unwrap()).into_bytes() } else { data }; Some((name, data)) })
}
fn cell<'a>(reopened: &'a Value, sheet_name: &str, row: u64, col: u64) -> &'a Value {
    reopened["payload"]["data"]["formulaCalendarSource"]["sheets"].as_array().unwrap().iter()
        .find(|sheet| sheet["name"] == sheet_name).unwrap()["cells"].as_array().unwrap().iter()
        .find(|cell| cell["row"] == row && cell["col"] == col).unwrap()
}

#[test]
fn binary_reopen_admits_authored_formula_exact_serial_and_native_table_identity() {
    let bytes = export(&fixture()); let saved = manifest(&bytes); let reopened = parsed(&bytes);
    assert_eq!(cell(&reopened, "Data", 0, 1)["formula"], "=DAY(A1)");
    assert_eq!(cell(&reopened, "Data", 0, 1)["formulaSource"]["aliasSheetId"], "data");
    assert_eq!(cell(&reopened, "Data", 0, 0)["dateSource"]["value"], 60.5);
    assert_eq!(reopened["payload"]["data"]["tables"][0]["officeTableId"], saved.tables[0].office_table_id);
    assert!(reopened["payload"]["data"]["tables"][0]["officeSource"]["xml"].as_str().unwrap().contains("tableColumns"));
    assert!(reopened["payload"]["data"]["formulaCalendarSource"]["helpers"].as_array().unwrap().iter().all(|helper| helper["reclaim"].as_bool() == Some(true)));
    let native = reopened["payload"]["data"]["sheets"][0]["cells"].as_array().unwrap().iter().find(|cell| cell["row"] == 0 && cell["col"] == 1).unwrap()["formula"].as_str().unwrap();
    assert_eq!(native.trim_start_matches('='), saved.sheets[0].cells.iter().find(|cell| cell.col == 1).unwrap().projected_formula.as_ref().unwrap().trim_start_matches('='));
}

#[test]
fn native_renaming_preserves_alias_keys_and_rewrites_targets_by_sheet_and_table_id() {
    let bytes = export(&fixture());
    let bytes = rewrite(&bytes, |name, data| {
        if !name.ends_with(".xml") { return Some((name, data)); }
        let xml = String::from_utf8(data).unwrap();
        let xml = if name == source::PART_NAME { xml } else { xml.replace("name=\"Data\"", "name=\"Renamed\"")
            .replace("'Data'!", "'Renamed'!").replace("DataTable", "RenamedTable") };
        Some((name, xml.into_bytes()))
    });
    let reopened = parsed(&bytes); let restored = &reopened["payload"]["data"]["formulaCalendarSource"]["sheets"][0];
    assert_eq!(restored["name"], "Renamed");
    assert_eq!(restored["import"]["formulaAliases"]["sheets"]["original"], "Renamed");
    assert_eq!(restored["import"]["formulaAliases"]["tables"]["oldtable"], "RenamedTable");
    assert_eq!(cell(&reopened, "Renamed", 0, 1)["formula"], "=DAY(A1)");
}

#[test]
fn native_sheet_reordering_uses_ids_instead_of_manifest_array_positions() {
    let bytes = edit_xml(&export(&fixture()), "xl/workbook.xml", |xml| {
        let start = xml.find("<sheets>").unwrap() + "<sheets>".len(); let end = xml[start..].find("</sheets>").unwrap() + start;
        let mut sheets = Vec::new(); let mut cursor = start;
        while cursor < end { let length = xml[cursor..end].find("/>").unwrap() + 2; sheets.push(&xml[cursor..cursor + length]); cursor += length; }
        sheets.swap(0, 1); format!("{}{}{}", &xml[..start], sheets.concat(), &xml[end..])
    });
    let reopened = parsed(&bytes);
    assert_eq!(reopened["payload"]["data"]["sheets"][0]["name"], "Other");
    assert_eq!(cell(&reopened, "Data", 0, 1)["formula"], "=DAY(A1)");
}

#[test]
fn changing_only_native_formula_caches_does_not_invalidate_the_calculation_graph() {
    let original = export(&fixture());
    let bytes = rewrite(&original, |name, data| {
        let data = if name.starts_with("xl/worksheets/") && name.ends_with(".xml") {
            String::from_utf8(data).unwrap().replace("<v>29</v>", "<v>30</v>").into_bytes()
        } else { data }; Some((name, data))
    });
    assert_eq!(cell(&parsed(&bytes), "Data", 0, 1)["formula"], "=DAY(A1)");
}

#[test]
fn office_edited_author_formula_stays_native_and_unused_fresh_helpers_can_be_reclaimed() {
    let bytes = export(&fixture()); let saved = manifest(&bytes);
    let projected = saved.sheets[0].cells.iter().find(|cell| cell.col == 1).unwrap().projected_formula.as_deref().unwrap().trim_start_matches('=');
    let bytes = edit_xml(&bytes, "xl/worksheets/sheet1.xml", |xml| xml.replace(&format!("<f>{projected}</f>"), "<f>42+1</f>"));
    let reopened = parsed(&bytes); let sources = &reopened["payload"]["data"]["formulaCalendarSource"];
    assert!(sources["sheets"][0]["cells"].as_array().unwrap().iter().all(|cell| cell["col"] != 1));
    // A changed author no longer references the helper; other approved authors
    // still allow reclamation when the graph is fresh and otherwise unowned.
    assert!(reopened["payload"]["data"]["sheets"][0]["cells"].as_array().unwrap().iter().any(|cell| cell["formula"].as_str().is_some_and(|formula| formula.trim_start_matches('=') == "42+1")));
    assert!(sources["helpers"].as_array().unwrap().iter().all(|helper| helper["reclaim"].as_bool() == Some(true)));
}

#[test]
fn changed_helper_formula_blocks_authored_revival_and_helper_reclamation() {
    let bytes = export(&fixture()); let saved = manifest(&bytes);
    let helper_index = saved.helpers[0].office_sheet_id; let target = format!("xl/worksheets/sheet{helper_index}.xml");
    let bytes = edit_xml(&bytes, &target, |xml| {
        let start = xml.find("<f>").unwrap(); let end = xml[start..].find("</f>").unwrap() + start + 4;
        format!("{}<f>123</f>{}", &xml[..start], &xml[end..])
    });
    let reopened = parsed(&bytes); let source = &reopened["payload"]["data"]["formulaCalendarSource"];
    assert!(source["sheets"][0]["cells"].as_array().unwrap().iter().all(|cell| cell["col"] != 1));
    assert!(source["helpers"].as_array().unwrap().iter().all(|helper| helper["reclaim"].as_bool() != Some(true)));
}

#[test]
fn newly_authored_foreign_formula_prevents_owned_helper_removal() {
    let bytes = export(&fixture()); let helper = manifest(&bytes).helpers[0].name.clone();
    let bytes = edit_xml(&bytes, "xl/worksheets/sheet2.xml", |xml| xml.replace("</sheetData>", &format!("<row r=\"1\"><c r=\"A1\"><f>'{helper}'!A1</f><v>60.5</v></c></row></sheetData>")));
    let reopened = parsed(&bytes);
    assert_eq!(cell(&reopened, "Data", 0, 1)["formula"], "=DAY(A1)");
    assert!(reopened["payload"]["data"]["formulaCalendarSource"]["helpers"].as_array().unwrap().iter().all(|helper| helper["reclaim"].as_bool() != Some(true)));
}

#[test]
fn visible_helper_is_preserved_even_when_calculation_content_is_fresh() {
    let bytes = export(&fixture()); let saved = manifest(&bytes); let native_id = saved.helpers[0].office_sheet_id;
    let bytes = edit_xml(&bytes, "xl/workbook.xml", |xml| {
        let needle = format!("sheetId=\"{native_id}\""); let start = xml.find(&needle).unwrap(); let end = xml[start..].find("/>").unwrap() + start;
        format!("{}{}{}", &xml[..start], xml[start..end].replace(" state=\"veryHidden\"", ""), &xml[end..])
    });
    let reopened = parsed(&bytes);
    assert!(reopened["payload"]["data"]["formulaCalendarSource"]["helpers"].as_array().unwrap().iter().all(|helper| helper["reclaim"].as_bool() != Some(true)));
}

#[test]
fn changed_exact_date_value_does_not_inherit_the_old_fictitious_day() {
    let bytes = edit_xml(&export(&fixture()), "xl/worksheets/sheet1.xml", |xml| xml.replace("<v>-1401.5</v>", "<v>-1400.5</v>"));
    let reopened = parsed(&bytes);
    assert!(reopened["payload"]["data"]["formulaCalendarSource"]["sheets"][0]["cells"].as_array().unwrap().iter().all(|cell| cell["row"] != 0 || cell["col"] != 0));
}

#[test]
fn relocated_related_custom_xml_and_native_workbook_parts_are_discovered() {
    let bytes = rewrite(&export(&fixture()), |name, data| {
        let xml = String::from_utf8(data).unwrap();
        let relocated = if name == source::PART_NAME { "owned/calendar.xml".into() } else { name.replacen("xl/", "book/", 1) };
        let xml = xml.replace("/xl/", "/book/").replace("Target=\"xl/", "Target=\"book/")
            .replace("customXml/officeFormulaCalendars.xml", "owned/calendar.xml");
        Some((relocated, xml.into_bytes()))
    });
    assert_eq!(cell(&parsed(&bytes), "Data", 0, 1)["formula"], "=DAY(A1)");
}

#[test]
fn unrelated_or_missing_manifest_keeps_native_formulas_without_restoration() {
    let bytes = rewrite(&export(&fixture()), |name, data| if name == source::PART_NAME { None } else {
        let data = if name == "xl/_rels/workbook.xml.rels" {
            let mut xml = String::from_utf8(data).unwrap();
            let start = xml.find(&format!("Type=\"{}\"", source::RELATIONSHIP)).unwrap();
            let start = xml[..start].rfind("<Relationship ").unwrap(); let end = xml[start..].find("/>").unwrap() + start + 2;
            xml.replace_range(start..end, ""); xml.into_bytes()
        } else { data }; Some((name, data))
    });
    assert!(parsed(&bytes)["payload"]["data"].get("formulaCalendarSource").is_none());
}

#[test]
fn malformed_owned_manifest_is_an_explicit_parse_failure() {
    let bytes = edit_xml(&export(&fixture()), source::PART_NAME, |xml| xml.replace("&quot;version&quot;:1", "&quot;version&quot;:2").replace("\"version\":1", "\"version\":2"));
    assert!(parse_office_bytes(&bytes, "invalid-calendar.xlsx").is_err());
}
