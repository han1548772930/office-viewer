//! Deferred native XML contracts; independent Office reopening is still required.
use super::*;
use serde_json::{Value, json};

fn native_source() -> String {
    format!(r#"<table xmlns="{}" xmlns:ov="urn:table-test" id="31" name="数据表" displayName="数据表" ref="B3:D8" headerRowCount="1" totalsRowCount="1" totalsRowShown="1" ov:author="retained">
      <autoFilter ref="B3:D7"><filterColumn colId="2" hiddenButton="1" showButton="0"><filters><filter val="old"/></filters><extLst><ext uri="column-extra"><ov:payload>keep-column</ov:payload></ext></extLst></filterColumn>
        <sortState ref="B4:D7" caseSensitive="1"><sortCondition ref="D4:D7" descending="1" customList="high,low"/></sortState><extLst><ext uri="filter-extra"><ov:payload>keep-filter</ov:payload></ext></extLst></autoFilter>
      <sortState ref="B4:D7" sortMethod="stroke"><sortCondition ref="C4:C7" descending="1"/><sortCondition ref="D4:D7" descending="0"/></sortState>
      <tableColumns count="3"><tableColumn id="1" name="甲"/><tableColumn id="2" name="乙" ov:deleted="yes"><calculatedColumnFormula>99</calculatedColumnFormula></tableColumn>
        <tableColumn id="3" name="丙" totalsRowFunction="sum" totalsRowLabel="汇总"><calculatedColumnFormula array="1" ov:formula="retained">[@甲]*2</calculatedColumnFormula><totalsRowFormula>SUM([丙])</totalsRowFormula><extLst><ext uri="field-extra"><ov:payload>keep-field</ov:payload></ext></extLst></tableColumn></tableColumns>
      <tableStyleInfo name="TableStyleMedium4" showFirstColumn="1" showLastColumn="0" showRowStripes="0" showColumnStripes="1"/>
      <extLst><ext uri="table-extra"><ov:payload>keep-table</ov:payload></ext></extLst></table>"#, native_xml::MAIN_NAMESPACE)
}

fn input() -> Value {
    json!({ "id": "stable-table", "name": "数据表", "sheetId": "data", "officeTableId": 31,
        "range": { "startRow": 9, "startCol": 5, "endRow": 17, "endCol": 7 },
        "showHeaderRow": true, "showTotalsRow": true, "styleName": "TableStyleMedium4",
        "columns": [ { "id": "stable-c", "name": "丙", "officeColumnId": 3 },
            { "id": "stable-a", "name": "甲", "officeColumnId": 1 }, { "id": "new-column", "name": "新增" } ],
        "officeSource": { "partName": "xl/tables/imported.xml", "xml": native_source() } })
}

fn prepared(value: Value) -> TableInput {
    let mut table: TableInput = serde_json::from_value(value).unwrap();
    prepare_identities(std::slice::from_mut(&mut table), &[]).unwrap();
    validate(std::slice::from_ref(&table), &table).unwrap();
    table
}

fn export(value: Value) -> Element {
    let table = prepared(value);
    let sheet = SheetInput { id: "data".into(), name: "Data".into(), ..SheetInput::default() };
    let output = write(&table, table.office_table_id.unwrap(), &sheet, &mut StyleRegistry::new(), &mut Vec::new()).unwrap();
    native_xml::parse(&output).unwrap()
}

fn attr<'a>(node: &'a Element, name: &str) -> Option<&'a str> { node.attrs.get(name).map(String::as_str) }

#[test]
fn new_columns_do_not_inherit_deleted_native_column_metadata() {
    let table = prepared(input());
    assert_eq!(table.office_table_id, Some(31));
    assert_eq!(table.columns.iter().map(|column| column.office_column_id).collect::<Vec<_>>(), [Some(3), Some(1), Some(4)]);
    let root = export(input()); let columns = source::columns(&root).unwrap();
    assert_eq!(attr(columns[2], "name"), Some("新增"));
    assert_eq!(attr(columns[2], "ov:deleted"), None);
    assert!(columns[2].child("calculatedColumnFormula").is_none());
}

#[test]
fn native_formulas_style_flags_and_unknown_extensions_follow_their_columns() {
    let root = export(input()); let columns = source::columns(&root).unwrap();
    assert_eq!(root.attribute_ns("urn:table-test", "author"), Some("retained"));
    assert_eq!(columns[0].child("calculatedColumnFormula").unwrap().text, "[@甲]*2");
    assert_eq!(attr(columns[0].child("calculatedColumnFormula").unwrap(), "array"), Some("1"));
    assert_eq!(columns[0].child("calculatedColumnFormula").unwrap().attribute_ns("urn:table-test", "formula"), Some("retained"));
    assert!(columns[0].render().contains("keep-field"));
    assert!(root.render().contains("keep-table"));
    let style = root.child("tableStyleInfo").unwrap();
    assert_eq!(attr(style, "showFirstColumn"), Some("1"));
    assert_eq!(attr(style, "showRowStripes"), Some("0"));
    assert_eq!(attr(style, "showColumnStripes"), Some("1"));
}

#[test]
fn explicit_null_removes_native_formulas_and_totals_without_erasing_extensions() {
    let mut value = input();
    for key in ["formula", "totalsRowFormula", "totalsRowLabel", "totalsRowFunction"] { value["columns"][0][key] = Value::Null; }
    let root = export(value); let column = source::columns(&root).unwrap()[0];
    assert!(column.child("calculatedColumnFormula").is_none()); assert!(column.child("totalsRowFormula").is_none());
    assert!(attr(column, "totalsRowLabel").is_none()); assert!(attr(column, "totalsRowFunction").is_none());
    assert!(column.render().contains("keep-field"));
}

#[test]
fn filter_ranges_exclude_totals_and_native_column_positions_follow_identity() {
    let root = export(input()); let filter = root.child("autoFilter").unwrap();
    assert_eq!(attr(filter, "ref"), Some("F10:H17"));
    let column = filter.child("filterColumn").unwrap();
    assert_eq!(attr(column, "colId"), Some("0"));
    assert_eq!(attr(column.child("filters").unwrap().child("filter").unwrap(), "val"), Some("old"));
    assert!(filter.render().contains("keep-column")); assert!(filter.render().contains("keep-filter"));
}

#[test]
fn replacing_typed_predicates_keeps_native_filter_extensions_and_button_state() {
    let mut value = input();
    value["autoFilter"] = json!({ "range": { "startRow": 9, "startCol": 5, "endRow": 17, "endCol": 7 },
        "columns": { "5": { "kind": "condition", "operator": "greaterThan", "value": 10 } }, "filteredRows": [] });
    value["columns"][0]["showFilterButton"] = json!(true);
    let root = export(value); let filter = root.child("autoFilter").unwrap(); let column = filter.child("filterColumn").unwrap();
    assert_eq!(attr(column, "colId"), Some("0")); assert!(column.child("filters").is_none());
    assert_eq!(attr(column, "hiddenButton"), Some("0")); assert_eq!(attr(column, "showButton"), Some("1"));
    assert_eq!(attr(column.child("customFilters").unwrap().child("customFilter").unwrap(), "val"), Some("10"));
    assert!(column.render().contains("keep-column")); assert!(filter.render().contains("keep-filter"));
    assert!(filter.child("sortState").is_some());
}

#[test]
fn table_and_filter_sort_keys_track_data_bounds_and_removed_columns() {
    let root = export(input()); let state = root.child("sortState").unwrap();
    assert_eq!(attr(state, "ref"), Some("F11:H17")); assert_eq!(attr(state, "sortMethod"), Some("stroke"));
    let keys = state.children.iter().filter(|child| source::spreadsheet(child, "sortCondition")).collect::<Vec<_>>();
    assert_eq!(keys.len(), 1); assert_eq!(attr(keys[0], "ref"), Some("F11:F17"));
    let state = root.child("autoFilter").unwrap().child("sortState").unwrap();
    assert_eq!(attr(state, "caseSensitive"), Some("1"));
    let key = state.child("sortCondition").unwrap();
    assert_eq!(attr(key, "customList"), Some("high,low")); assert_eq!(attr(key, "descending"), Some("1"));
    assert_eq!(attr(key, "ref"), Some("F11:F17"));
}

#[test]
fn horizontal_sort_conditions_keep_their_row_and_format_metadata_after_table_moves() {
    let mut value = input(); let mut source = source::parse_xml(&native_source()).unwrap();
    let state = source.child_mut("sortState").unwrap(); state.set_attr("columnSort", "1");
    let mut key = Element::new(native_xml::MAIN_NAMESPACE, "sortCondition");
    key.set_attr("ref", "B5:D5"); key.set_attr("sortBy", "cellColor"); key.set_attr("dxfId", "4");
    state.replace_children(|child| source::spreadsheet(child, "sortCondition"), vec![key]);
    value["officeSource"]["xml"] = json!(source.render());
    let root = export(value); let state = root.child("sortState").unwrap(); let key = state.child("sortCondition").unwrap();
    assert_eq!(attr(state, "columnSort"), Some("1")); assert_eq!(attr(key, "ref"), Some("F12:H12"));
    assert_eq!(attr(key, "sortBy"), Some("cellColor")); assert_eq!(attr(key, "dxfId"), Some("4"));
}

#[test]
fn native_partial_sort_regions_do_not_grow_to_include_new_columns() {
    let mut value = input(); let mut source = source::parse_xml(&native_source()).unwrap();
    source.child_mut("sortState").unwrap().set_attr("ref", "C4:D7"); value["officeSource"]["xml"] = json!(source.render());
    let root = export(value); let state = root.child("sortState").unwrap();
    assert_eq!(attr(state, "ref"), Some("F11:F17")); assert_eq!(attr(state.child("sortCondition").unwrap(), "ref"), Some("F11:F17"));
}

#[test]
fn header_only_tables_do_not_emit_a_sort_over_header_cells() {
    let mut value = input(); value["showTotalsRow"] = json!(false);
    value["range"]["endRow"] = value["range"]["startRow"].clone();
    let root = export(value); assert!(root.child("sortState").is_none());
    assert!(root.child("autoFilter").unwrap().child("sortState").is_none());
}

#[test]
fn clearing_filter_does_not_erase_the_independent_table_sort() {
    let mut value = input(); value["autoFilter"] = Value::Null;
    let root = export(value); assert!(root.child("autoFilter").is_none()); assert!(root.child("sortState").is_some());
}

#[test]
fn hiding_totals_retains_the_native_ever_shown_history_and_the_final_data_row() {
    let mut value = input(); value["showTotalsRow"] = json!(false);
    let root = export(value);
    assert_eq!(attr(&root, "totalsRowCount"), Some("0")); assert_eq!(attr(&root, "totalsRowShown"), Some("1"));
    assert_eq!(attr(root.child("autoFilter").unwrap(), "ref"), Some("F10:H18"));
    assert_eq!(attr(root.child("sortState").unwrap(), "ref"), Some("F11:H18"));
}

#[test]
fn strict_native_namespaces_convert_without_modifying_extension_namespaces() {
    let mut value = input(); value["officeSource"]["xml"] = json!(native_source().replace(native_xml::MAIN_NAMESPACE, native_xml::STRICT_MAIN_NAMESPACE));
    let root = export(value); assert_eq!(root.namespace, native_xml::MAIN_NAMESPACE);
    assert_eq!(source::columns(&root).unwrap()[0].namespace, native_xml::MAIN_NAMESPACE);
    assert!(root.descendants("payload").iter().all(|node| node.namespace == "urn:table-test"));
}

#[test]
fn later_requested_native_table_ids_are_reserved_before_new_ids_are_assigned() {
    let mut new = input(); new["id"] = json!("new-table"); new["name"] = json!("新表"); new.as_object_mut().unwrap().remove("officeSource"); new.as_object_mut().unwrap().remove("officeTableId");
    let mut tables: Vec<TableInput> = serde_json::from_value(json!([new, input()])).unwrap();
    prepare_identities(&mut tables, &[1]).unwrap();
    assert_eq!(tables[0].office_table_id, Some(2)); assert_eq!(tables[1].office_table_id, Some(31));
}

#[test]
fn calculation_mirrors_have_no_author_formula_source_or_native_identity() {
    let mut table = prepared(input()); table.id = "calendar-mirror".into(); clear_calculation_source(&mut table);
    assert!(table.office_source.is_none()); assert!(table.office_table_id.is_none());
    assert!(matches!(table.auto_filter, Some(None)));
    for (index, column) in table.columns.iter().enumerate() {
        assert_eq!(column.id, format!("calendar-mirror:column:{}", index + 1));
        assert!(column.office_column_id.is_none()); assert!(matches!(column.formula, Some(None)));
        assert!(matches!(column.totals_row_formula, Some(None))); assert!(matches!(column.totals_row_function, Some(None)));
    }
}

#[test]
fn out_of_table_sort_references_fail_before_native_xml_is_emitted() {
    let mut value = input(); value["officeSource"]["xml"] = json!(native_source().replace("ref=\"D4:D7\"", "ref=\"X4:X7\""));
    let table = prepared(value); let sheet = SheetInput::default();
    let error = write(&table, 31, &sheet, &mut StyleRegistry::new(), &mut Vec::new()).unwrap_err();
    assert!(error.to_string().contains("TABLE_SOURCE_INVALID"));
}
