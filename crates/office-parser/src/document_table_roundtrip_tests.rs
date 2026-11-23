//! Real DOCX writer -> Duke-backed parser drafts for unified validation.
//! These have been written but have not been run or independently opened in Office.
use std::io::{Cursor, Read, Write};

use office_io_common::document_table_source::{TABLE_SOURCE_NAMESPACE, TABLE_SOURCE_PART, WORD_NAMESPACE};
use serde_json::{Value, json};
use zip::{ZipArchive, ZipWriter, write::SimpleFileOptions};

use crate::{document_format_sidecar::{Element, parse_xml}, parser::parse_office_bytes};

fn cell(id: &str, text: &str) -> Value {
    json!({"id":id,"columnIndex":0,"rowSpan":1,"colSpan":1,"style":{},"blocks":[
        {"id":format!("{id}-paragraph"),"type":"paragraph","runs":[{"text":text}],"metadata":{}}
    ]})
}

fn model() -> Value {
    let mut merged = cell("stable-merged-cell", "Anchor");
    merged["rowSpan"] = json!(2);
    merged["style"] = json!({"fill":"#fff2cc","selectedDropdownOptionId":"chosen","authorCellSetting":{"value":7}});
    json!({"kind":"document","name":"Native tables","blocks":[{"id":"stable-table","type":"table","runs":[],"metadata":{
        "table":{"columnWidths":[60.25],"headerRowCount":0,"titleRow":false,
            "columnTypes":{"0":{"type":"dropdown","dropdownOptions":[{"id":"chosen","label":"Anchor","color":"#ff0000"}]}},
            "style":{"align":"center","indent":4.5,"layout":"fixed","authorTableSetting":"retained"},
            "rows":[{"id":"stable-first","height":31.25,"heightRule":"exact","allowSplit":false,"cells":[merged]},
                {"id":"stable-covered","heightRule":"auto","allowSplit":true,"cells":[]}]}
    }}]})
}

fn export(model: &Value) -> Vec<u8> {
    let (bytes, warnings) = office_writer::export_docx_bytes(&model.to_string()).unwrap();
    assert!(warnings.is_empty(), "{warnings:?}");
    bytes
}

fn reopen(bytes: &[u8]) -> Value {
    serde_json::to_value(parse_office_bytes(bytes, "native-tables.docx").unwrap()).unwrap()
}

fn part(bytes: &[u8], path: &str) -> String {
    let mut archive = ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut xml = String::new();
    archive.by_name(path).unwrap().read_to_string(&mut xml).unwrap();
    xml
}

fn rewrite_part(bytes: &[u8], path: &str, xml: &str) -> Vec<u8> {
    let mut archive = ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
    let mut found = false;
    for index in 0..archive.len() {
        let mut source = archive.by_index(index).unwrap();
        let name = source.name().to_owned();
        writer.start_file(&name, SimpleFileOptions::default()).unwrap();
        if name == path { found = true; writer.write_all(xml.as_bytes()).unwrap(); }
        else { std::io::copy(&mut source, &mut writer).unwrap(); }
    }
    assert!(found);
    writer.finish().unwrap().into_inner()
}

fn tables(envelope: &Value) -> Vec<&Value> {
    envelope["payload"]["data"]["blocks"].as_array().unwrap().iter().filter(|block| block["type"] == "table").collect()
}

fn native_table(xml: &str) -> Element {
    let mut root = parse_xml(xml).unwrap();
    let body = root.children.iter_mut().find(|node| node.name == "body").unwrap();
    let index = body.children.iter().position(|node| node.name == "tbl").unwrap();
    body.children.remove(index)
}

fn without_bookmarks(xml: &str) -> String {
    let tree = parse_xml(xml).unwrap();
    fn collect(node: &Element, ranges: &mut Vec<(usize, usize)>) {
        if matches!(node.name.as_str(), "bookmarkStart" | "bookmarkEnd") { ranges.push((node.start, node.end)); }
        else { for child in &node.children { collect(child, ranges); } }
    }
    let mut ranges = Vec::new(); collect(&tree, &mut ranges);
    ranges.sort_unstable_by(|left, right| right.0.cmp(&left.0));
    let mut result = xml.to_owned();
    for (start, end) in ranges { result.replace_range(start..end, ""); }
    result
}

#[test]
fn native_covered_rows_keep_authored_ids_column_types_and_precise_geometry() {
    let bytes = export(&model());
    assert!(part(&bytes, TABLE_SOURCE_PART).contains(TABLE_SOURCE_NAMESPACE));
    let envelope = reopen(&bytes); let table = tables(&envelope)[0];
    let details = &table["details"]; let source = &details["officeTable"];
    assert_eq!(details["rowCount"], 2);
    assert_eq!(source["tableId"], "stable-table");
    assert_eq!(source["nativeMatches"], true);
    assert_eq!(source["topologyMatches"], true);
    assert_eq!(source["data"]["columnWidths"], json!([60.25]));
    assert_eq!(source["data"]["rows"][1]["id"], "stable-covered");
    assert_eq!(source["data"]["rows"][1]["cells"], json!([]));
    assert_eq!(source["data"]["columnTypes"]["0"]["dropdownOptions"][0]["id"], "chosen");
    assert_eq!(details["nativeTableProperties"]["rows"].as_array().unwrap().len(), 2);
    let cells = table["children"].as_array().unwrap();
    assert_eq!(cells[0]["details"]["rowSpan"], 2);
    assert_eq!(cells[1]["details"]["style"]["vMergeContinuation"], true);
    assert_eq!(details["nativeTableProperties"]["rows"][0]["cells"][0]["nativeProperties"]["version"], 1);
}

#[test]
fn legacy_horizontal_members_and_vertical_continuations_form_one_real_logical_anchor() {
    let mut input = model();
    let table = &mut input["blocks"][0]["metadata"]["table"];
    let mut left = cell("left", "Left"); left["rowSpan"] = json!(2);
    let mut right = cell("right", "Right😀"); right["rowSpan"] = json!(2); right["columnIndex"] = json!(1);
    table["columnWidths"] = json!([60,80]); table["columnTypes"] = json!({});
    table["rows"][0]["cells"] = json!([left,right]);
    let bytes = export(&input);
    let xml = part(&bytes, "word/document.xml");
    let native = native_table(&xml);
    let mut insertions = Vec::new();
    for row in native.children.iter().filter(|node| node.name == "tr") {
        for (column, cell) in row.children.iter().filter(|node| node.name == "tc").enumerate() {
            let properties = cell.child("tcPr").unwrap();
            let position = properties.child("vMerge").unwrap().start;
            insertions.push((position, if column == 0 { "<w:hMerge w:val=\"restart\"/>" } else { "<w:hMerge/>" }));
        }
    }
    assert_eq!(insertions.len(), 4);
    insertions.sort_unstable_by(|left, right| right.0.cmp(&left.0));
    let mut legacy = xml;
    for (position, tag) in insertions { legacy.insert_str(position, tag); }
    let parsed = reopen(&rewrite_part(&bytes, "word/document.xml", &legacy));
    let table = tables(&parsed)[0];
    assert_eq!(table["details"]["officeTable"]["topologyMatches"], false);
    let cells = table["children"].as_array().unwrap();
    assert_eq!(cells.len(), 2);
    assert_eq!(cells[0]["details"]["columnSpan"], 2);
    assert_eq!(cells[0]["details"]["rowSpan"], 2);
    assert_eq!(cells[0]["content"], "Left\nRight😀");
    assert_eq!(cells[1]["details"]["style"]["vMergeContinuation"], true);
    assert_eq!(cells[1]["details"]["columnSpan"], 2);
    let source = &table["details"]["nativeTableProperties"]["rows"][0]["cells"][0]["nativeProperties"];
    assert_eq!(source["rewriteProperties"], json!(["tcW","tcBorders"]));
    assert_eq!(source["sourceHorizontalMerge"]["members"].as_array().unwrap().len(), 2);
    assert!(source["sourceHorizontalMerge"]["members"][1]["nativeProperties"]["propertiesXml"].as_str().unwrap().contains("hMerge"));
}

#[test]
fn an_external_text_edit_keeps_the_authored_settings_current() {
    let bytes = export(&model());
    let xml = part(&bytes, "word/document.xml");
    assert!(xml.contains(">Anchor</w:t>"));
    let edited = rewrite_part(&bytes, "word/document.xml", &xml.replace(">Anchor</w:t>", ">External 文本😀</w:t>"));
    let envelope = reopen(&edited); let table = tables(&envelope)[0];
    assert!(table["content"].as_str().unwrap().contains("External 文本😀"));
    assert_eq!(table["details"]["officeTable"]["nativeMatches"], true);
    assert_eq!(table["details"]["officeTable"]["data"]["rows"][0]["cells"][0]["style"]["selectedDropdownOptionId"], "chosen");
}

#[test]
fn external_native_width_and_row_height_changes_invalidate_only_the_native_baseline() {
    let bytes = export(&model());
    let xml = part(&bytes, "word/document.xml");
    assert!(xml.contains("<w:gridCol w:w=\"904\"/>"));
    let changed = xml.replace("<w:gridCol w:w=\"904\"/>", "<w:gridCol w:w=\"1200\"/>")
        .replace("w:val=\"469\"", "w:val=\"600\"");
    let envelope = reopen(&rewrite_part(&bytes, "word/document.xml", &changed));
    let table = tables(&envelope)[0];
    assert_eq!(table["details"]["officeTable"]["nativeMatches"], false);
    assert_eq!(table["details"]["officeTable"]["topologyMatches"], true);
    assert_eq!(table["details"]["style"]["columnWidthsTwips"], json!([1200]));
    assert_eq!(table["details"]["officeTable"]["data"]["style"]["authorTableSetting"], "retained");
}

#[test]
fn removing_a_physical_row_retains_old_cell_metadata_only_as_stale_source() {
    let bytes = export(&model());
    let mut xml = part(&bytes, "word/document.xml");
    let table = native_table(&xml);
    let last = table.children.iter().filter(|node| node.name == "tr").last().unwrap();
    xml.replace_range(last.start..last.end, "");
    let envelope = reopen(&rewrite_part(&bytes, "word/document.xml", &xml));
    let table = tables(&envelope)[0];
    assert_eq!(table["details"]["rowCount"], 1);
    assert_eq!(table["details"]["officeTable"]["nativeMatches"], false);
    assert_eq!(table["details"]["officeTable"]["topologyMatches"], false);
    assert_eq!(table["details"]["officeTable"]["data"]["rows"].as_array().unwrap().len(), 2);
}

#[test]
fn a_copied_native_source_key_does_not_duplicate_the_authored_table_identity() {
    let bytes = export(&model());
    let xml = part(&bytes, "word/document.xml");
    let table = native_table(&xml);
    let duplicate = without_bookmarks(&xml[table.start..table.end]);
    let mut edited = xml.clone(); edited.insert_str(table.end, &duplicate);
    let envelope = reopen(&rewrite_part(&bytes, "word/document.xml", &edited));
    let tables = tables(&envelope);
    assert_eq!(tables.len(), 2);
    assert_eq!(tables[0]["details"]["officeTable"]["tableId"], "stable-table");
    assert!(tables[1]["details"].get("officeTable").is_none());
    assert_ne!(tables[0]["id"], tables[1]["id"]);
    assert!(envelope.to_string().contains("duplicate source key"));
}

#[test]
fn nested_table_sources_and_synthetic_terminal_paragraphs_remain_distinct() {
    let mut outer = cell("outer-cell", "");
    outer["blocks"] = json!([{"id":"nested-table","type":"table","runs":[],"metadata":{"table":{
        "rows":[{"id":"nested-row","cells":[cell("nested-cell","Nested text")]}]}
    }}]);
    let model = json!({"kind":"document","blocks":[{"id":"outer-table","type":"table","runs":[],"metadata":{"table":{
        "columnWidths":[100],"rows":[{"id":"outer-row","cells":[outer]}]}
    }}]});
    let envelope = reopen(&export(&model));
    let outer = tables(&envelope)[0];
    let children = outer["children"][0]["children"].as_array().unwrap();
    let nested = children.iter().find(|child| child["type"] == "table").unwrap();
    assert_eq!(outer["details"]["officeTable"]["tableId"], "outer-table");
    assert_eq!(nested["details"]["officeTable"]["tableId"], "nested-table");
    assert_eq!(outer["details"]["nativeTableProperties"]["rows"][0]["cells"][0]["syntheticTail"], true);
    assert_eq!(nested["details"]["nativeTableProperties"]["rows"][0]["cells"][0]["syntheticTail"], false);
}

#[test]
fn external_table_xml_without_an_author_extension_still_retains_native_properties() {
    let xml = format!("<w:document xmlns:w=\"{WORD_NAMESPACE}\" xmlns:future=\"urn:test:future\"><w:body><w:tbl><w:tblPr><w:jc w:val=\"right\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"1500\"/></w:tblGrid><w:tr><w:trPr/><w:tc><w:tcPr><future:setting future:val=\"kept\"/></w:tcPr><w:p><w:r><w:t>Native</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>");
    let package = duke_pages::create_minimal_docx_package(Some(&xml));
    let bytes = duke_pages::package_to_bytes(&package).unwrap();
    let envelope = reopen(&bytes); let table = tables(&envelope)[0];
    assert!(table["details"].get("officeTable").is_none());
    let source = &table["details"]["nativeTableProperties"]["rows"][0]["cells"][0]["nativeProperties"];
    assert!(source["propertiesXml"].as_str().unwrap().contains("setting"));
    assert!(source["namespaces"].as_object().unwrap().values().any(|value| value == "urn:test:future"));
}

#[test]
fn native_continuation_field_results_reopen_as_current_continuation_paragraphs() {
    let mut document = model();
    document["blocks"][0]["metadata"]["table"]["rows"][0]["cells"][0]["style"]["nativeContinuationContent"] = json!({"1":{
        "version":1,"ownerCellId":"stable-merged-cell","rowId":"stable-covered","colSpan":1,"blocks":[
            {"id":"continued-paragraph","type":"paragraph","runs":[{"text":"2026"}],"metadata":{"documentTextFields":{"version":1,"fields":[
                {"id":"continued-native-date","field":"DATE","instruction":"DATE","startOffset":0,"endOffset":4,"locked":true}
            ]}}}
        ]
    }});
    let envelope = reopen(&export(&document));
    let table = tables(&envelope)[0];
    let continuation = table["children"].as_array().unwrap().iter().find(|cell| cell["details"]["style"]["vMergeContinuation"] == true).unwrap();
    assert!(continuation["content"].as_str().unwrap().contains("2026"));
    let paragraph = continuation["children"].as_array().unwrap().iter().find(|block|
        block["details"]["documentTextFields"]["fields"][0]["id"] == "continued-native-date").unwrap();
    let field = &paragraph["details"]["documentTextFields"]["fields"][0];
    assert_eq!(field["id"], "continued-native-date");
    assert_eq!(field["cachedResult"], "2026");
    assert_eq!(field["locked"], true);
    assert_eq!(table["details"]["officeTable"]["nativeMatches"], true);
    assert!(table["details"]["officeTable"]["data"]["rows"][0]["cells"][0]["style"].get("nativeContinuationContent").is_none());
}
