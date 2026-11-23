//! Drafts for the unified validation phase; no checks were executed while writing.
use std::io::{Cursor, Read};
use serde_json::{Value, json};
use office_io_common::{chart_xml, document_table_source::{TABLE_SOURCE_NAMESPACE, TABLE_SOURCE_PART, WORD_NAMESPACE}};
use crate::export_docx_bytes;

fn model(table: Value) -> Value {
    json!({"kind":"document","name":"Table contracts","blocks":[{"id":"table-main","type":"table","runs":[],"metadata":{"table":table}}]})
}

fn cell(id: &str, text: &str) -> Value {
    json!({"id":id,"rowSpan":1,"colSpan":1,"style":{},"blocks":[{"id":format!("{id}-paragraph"),"type":"paragraph","runs":[{"text":text}],"metadata":{}}]})
}

fn part(bytes: &[u8], path: &str) -> String {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut xml = String::new();
    archive.by_name(path).unwrap().read_to_string(&mut xml).unwrap();
    xml
}

fn export(table: Value) -> Vec<u8> {
    let (bytes, warnings) = export_docx_bytes(&model(table).to_string()).unwrap();
    assert!(warnings.is_empty(), "{warnings:?}");
    bytes
}

fn native_table(xml: &str) -> chart_xml::Node {
    chart_xml::parse(xml).unwrap().child("body").unwrap().child("tbl").unwrap().clone()
}

#[test]
fn actual_covered_rows_and_legacy_missing_column_indices_share_one_native_grid() {
    let mut first = cell("merged", "Merged"); first["rowSpan"] = json!(2);
    let bytes = export(json!({"rows":[{"id":"r0","cells":[first,cell("top-right","Top")]},{"id":"r1","cells":[cell("bottom-right","Bottom")]}]}));
    let xml = part(&bytes, "word/document.xml");
    let native = native_table(&xml);
    let rows = native.children_named("tr").collect::<Vec<_>>();
    assert_eq!(rows.len(), 2);
    assert_eq!(native.child("tblGrid").unwrap().children.len(), 2);
    assert_eq!(rows[0].children_named("tc").next().unwrap().child("tcPr").unwrap().child("vMerge").unwrap().attribute_ns(WORD_NAMESPACE,"val"),Some("restart"));
    assert_eq!(rows[1].children_named("tc").next().unwrap().child("tcPr").unwrap().child("vMerge").unwrap().attribute_ns(WORD_NAMESPACE,"val"),Some("continue"));
    assert!(xml.contains("Bottom"));
    let source = chart_xml::parse(&part(&bytes,TABLE_SOURCE_PART)).unwrap();
    let data: Value = serde_json::from_str(&source.children[0].children[0].text).unwrap();
    assert_eq!(data["rows"][1]["cells"][0]["columnIndex"],1);
    assert_eq!(data["rows"][0]["cells"][0]["rowSpan"],2);
    assert!(data["rows"][0]["cells"][0].get("blocks").is_none());
}

#[test]
fn explicit_sizes_header_height_and_split_settings_are_native_properties() {
    let bytes = export(json!({"columnWidths":[40.25,81.75],"headerRowCount":1,"titleRow":false,
        "style":{"align":"right","indent":12,"width":600,"layout":"auto","cellMargin":{"top":2,"right":3,"bottom":4,"left":5}},
        "rows":[{"id":"r0","height":41.5,"heightRule":"exact","allowSplit":false,"cells":[cell("a","A"),cell("b","B")]},
            {"id":"r1","height":30,"heightRule":"auto","allowSplit":true,"cells":[cell("c","C"),cell("d","D")]}]}));
    let table = native_table(&part(&bytes,"word/document.xml"));
    let properties = table.child("tblPr").unwrap();
    assert_eq!(properties.child("tblW").unwrap().attribute_ns(WORD_NAMESPACE,"w"),Some("1830"));
    assert_eq!(properties.child("jc").unwrap().attribute_ns(WORD_NAMESPACE,"val"),Some("right"));
    assert_eq!(properties.child("tblLayout").unwrap().attribute_ns(WORD_NAMESPACE,"type"),Some("autofit"));
    assert_eq!(properties.child("tblCellMar").unwrap().child("left").unwrap().attribute_ns(WORD_NAMESPACE,"w"),Some("75"));
    let rows = table.children_named("tr").collect::<Vec<_>>();
    let first = rows[0].child("trPr").unwrap();
    assert_eq!(first.child("trHeight").unwrap().attribute_ns(WORD_NAMESPACE,"hRule"),Some("exact"));
    assert_eq!(first.child("cantSplit").unwrap().attribute_ns(WORD_NAMESPACE,"val"),Some("1"));
    assert_eq!(first.child("tblHeader").unwrap().attribute_ns(WORD_NAMESPACE,"val"),Some("1"));
    assert_eq!(rows[1].child("trPr").unwrap().child("trHeight").unwrap().attribute_ns(WORD_NAMESPACE,"hRule"),Some("auto"));
    assert_eq!(rows[1].child("trPr").unwrap().child("tblHeader").unwrap().attribute_ns(WORD_NAMESPACE,"val"),Some("0"));
}

#[test]
fn canonical_horizontal_spans_rewrite_member_width_and_keep_original_member_provenance() {
    let mut merged = cell("combined","A\nB"); merged["colSpan"] = json!(2); merged["columnIndex"] = json!(0);
    let source = json!({"version":1,"rows":[{"rowId":"r0","members":[{"cellId":"a","columnSpan":1},{"cellId":"b","columnSpan":1}]}]});
    merged["style"] = json!({"nativeHorizontalMergeSource":source,"nativeProperties":{
        "version":1,"namespaces":{"w":WORD_NAMESPACE},
        "propertiesXml":"<w:tcPr><w:tcW w:w=\"600\" w:type=\"dxa\"/><w:hMerge w:val=\"restart\"/></w:tcPr>",
        "baseline":{"tcW":2100},"rewriteProperties":["tcW","tcBorders"]}});
    let bytes = export(json!({"columnWidths":[60,80],"rows":[{"id":"r0","cells":[merged]}]}));
    let table = native_table(&part(&bytes,"word/document.xml"));
    let properties = table.child("tr").unwrap().child("tc").unwrap().child("tcPr").unwrap();
    assert_eq!(properties.child("tcW").unwrap().attribute_ns(WORD_NAMESPACE,"w"),Some("2100"));
    assert_eq!(properties.child("gridSpan").unwrap().attribute_ns(WORD_NAMESPACE,"val"),Some("2"));
    assert!(properties.child("hMerge").is_none());
    let sidecar = chart_xml::parse(&part(&bytes,TABLE_SOURCE_PART)).unwrap();
    let data: Value = serde_json::from_str(&sidecar.children[0].children[0].text).unwrap();
    assert_eq!(data["rows"][0]["cells"][0]["style"]["nativeHorizontalMergeSource"],source);
}

#[test]
fn nested_tables_use_the_parent_cell_content_width_and_have_a_marked_terminal_paragraph() {
    let mut parent = cell("parent", "");
    parent["blocks"] = json!([{"id":"nested","type":"table","metadata":{"table":{"rows":[{"id":"nested-row","cells":[cell("x","X"),cell("y","Y")]}]}}}]);
    let bytes = export(json!({"columnWidths":[60],"rows":[{"id":"r0","cells":[parent]}]}));
    let table = native_table(&part(&bytes,"word/document.xml"));
    let parent = table.child("tr").unwrap().child("tc").unwrap();
    let nested = parent.child("tbl").unwrap();
    assert_eq!(nested.child("tblPr").unwrap().child("tblW").unwrap().attribute_ns(WORD_NAMESPACE,"w"),Some("660"));
    assert_eq!(nested.child("tblGrid").unwrap().children.len(),2);
    assert_eq!(parent.children.last().unwrap().attribute_ns(TABLE_SOURCE_NAMESPACE,"tableEnd"),Some("1"));
    let source = chart_xml::parse(&part(&bytes,TABLE_SOURCE_PART)).unwrap();
    assert_eq!(source.children.len(),2);
    assert!(source.children.iter().any(|table| table.attrs.get("tableId").map(String::as_str) == Some("nested")));
}

#[test]
fn malformed_grids_and_unversioned_continuation_content_are_rejected_without_fabricating_cells() {
    let mut overlap = cell("b","B"); overlap["columnIndex"] = json!(0);
    let mut span = cell("span","Span"); span["rowSpan"] = json!(2);
    let mut hole = cell("hole","Hole"); hole["columnIndex"] = json!(1);
    let mut preserved = cell("source","Source");
    preserved["style"]["nativeContinuationContent"] = json!({"1":{"content":"must remain intact"}});
    for table in [
        json!({"rows":[{"id":"r","cells":[cell("a","A"),overlap]}]}),
        json!({"rows":[{"id":"r","cells":[span]}]}),
        json!({"rows":[{"id":"r","cells":[hole]}]}),
        json!({"columnWidths":[0],"rows":[{"id":"r","cells":[cell("zero","Z")]}]}),
        json!({"rows":[{"id":"r","cells":[preserved]}]}),
    ] { assert!(export_docx_bytes(&model(table).to_string()).is_err()); }
}

#[test]
fn independent_continuation_contents_use_their_stable_row_and_shared_field_control_note_and_numbering_writers() {
    let mut merged = cell("merged-content", "Anchor");
    merged["rowSpan"] = json!(3);
    merged["style"]["nativeContinuationContent"] = json!({"1":{"version":1,"ownerCellId":"merged-content","rowId":"stable-last","colSpan":1,
        "blocks":[
            {"id":"continued-field","type":"paragraph","runs":[{"text":"26"}],"metadata":{"documentTextFields":{"version":1,"fields":[
                {"id":"continued-date","field":"DATE","instruction":"DATE","startOffset":0,"endOffset":2,"locked":true}]}}},
            {"id":"continued-control","type":"paragraph","runs":[{"text":"Control","footnoteId":"continued-note"}],"metadata":{}},
            {"id":"continued-list","type":"numbered","level":0,"runs":[{"text":"Continued list"}],"metadata":{}}
        ]}});
    let mut document = model(json!({"rows":[{"id":"first","cells":[merged]},{"id":"inserted","cells":[]},{"id":"stable-last","cells":[]}]}));
    document["contentControls"] = json!([{"id":"continued-sdt","type":"text","blockId":"continued-control","startOffset":0,"endOffset":7}]);
    document["footnotes"] = json!({"continued-note":{"body":{"text":"Continuation footnote"}}});
    let (bytes, warnings) = export_docx_bytes(&document.to_string()).unwrap();
    assert!(warnings.is_empty(), "{warnings:?}");
    let native = native_table(&part(&bytes,"word/document.xml"));
    let rows = native.children_named("tr").collect::<Vec<_>>();
    assert!(rows[1].child("tc").unwrap().children.iter().all(|node| node.name != "sdt"));
    let continued = rows[2].child("tc").unwrap();
    assert!(continued.descendants("instrText").iter().any(|node| node.text.trim() == "DATE"));
    assert!(continued.descendants("sdt").len() == 1);
    assert!(continued.descendants("footnoteReference").iter().any(|node| node.attribute_ns(WORD_NAMESPACE,"id") == Some("1")));
    assert!(continued.descendants("numPr").len() == 1);
    assert!(part(&bytes,"word/footnotes.xml").contains("Continuation footnote"));
    assert!(part(&bytes,"word/numbering.xml").contains("w:numId=\"1\""));
    assert!(part(&bytes,"customXml/office-fields.xml").contains("continued-date"));
    let source = chart_xml::parse(&part(&bytes,TABLE_SOURCE_PART)).unwrap();
    let data: Value = serde_json::from_str(&source.children[0].children[0].text).unwrap();
    assert!(data["rows"][0]["cells"][0]["style"].get("nativeContinuationContent").is_none());
}

#[test]
fn stale_or_conflicting_continuation_contents_are_rejected_before_resources_are_written() {
    let source = json!({"version":1,"ownerCellId":"merged","rowId":"last","colSpan":1,"blocks":[
        {"id":"continuation","type":"paragraph","runs":[{"text":"Keep"}]}]});
    for invalid in [
        json!({"1":{"version":2,"ownerCellId":"merged","rowId":"last","colSpan":1,"blocks":[]}}),
        json!({"1":{"version":1,"ownerCellId":"other","rowId":"last","colSpan":1,"blocks":[]}}),
        json!({"1":{"version":1,"ownerCellId":"merged","rowId":"removed","colSpan":1,"blocks":[]}}),
        json!({"1":source.clone(),"2":source.clone()}),
    ] {
        let mut merged = cell("merged","Anchor"); merged["rowSpan"] = json!(2);
        merged["style"]["nativeContinuationContent"] = invalid;
        let document = model(json!({"rows":[{"id":"first","cells":[merged]},{"id":"last","cells":[]}]}));
        assert!(export_docx_bytes(&document.to_string()).is_err());
    }
}

#[test]
fn unknown_native_properties_and_row_identity_bound_continuations_survive_a_row_insert() {
    let mut merged = cell("merged", "Text"); merged["rowSpan"] = json!(3);
    merged["style"]["nativeProperties"] = json!({"version":1,"propertiesXml":"<w:tcPr><w:noWrap/><future:flag w:val=\"1\" future:val=\"2\"/></w:tcPr>",
        "namespaces":{"w":WORD_NAMESPACE,"future":"urn:test:future"},"baseline":{}});
    merged["style"]["nativeContinuations"] = json!({"1":{"version":1,"ownerCellId":"merged","rowId":"original-last","colSpan":1,
        "propertiesXml":"<w:tcPr><w:textDirection w:val=\"tbRl\"/></w:tcPr>","namespaces":{"w":WORD_NAMESPACE},"baseline":{}}});
    let bytes = export(json!({"rows":[{"id":"r0","cells":[merged]},{"id":"inserted","cells":[]},{"id":"original-last","cells":[]}]}));
    let table = native_table(&part(&bytes,"word/document.xml"));
    let rows = table.children_named("tr").collect::<Vec<_>>();
    assert!(rows[0].child("tc").unwrap().child("tcPr").unwrap().child("noWrap").is_some());
    let future = rows[0].child("tc").unwrap().child("tcPr").unwrap().child("flag").unwrap();
    assert_eq!(future.attribute_ns("urn:test:future","val"),Some("2"));
    assert_eq!(future.attribute_ns(WORD_NAMESPACE,"val"),Some("1"));
    assert!(rows[1].child("tc").unwrap().child("tcPr").unwrap().child("textDirection").is_none());
    assert_eq!(rows[2].child("tc").unwrap().child("tcPr").unwrap().child("textDirection").unwrap().attribute_ns(WORD_NAMESPACE,"val"),Some("tbRl"));
}
