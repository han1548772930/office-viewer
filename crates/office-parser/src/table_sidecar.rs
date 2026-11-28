//! Namespace-aware table identities resolved through the actual OPC package.

use std::{collections::BTreeSet, io::Cursor};
use office_io_common::{conditional_format_xml::{self as native_xml, Element}, sheet_table_source};
use quick_xml::{NsReader, XmlVersion, events::Event, name::ResolveResult};
use serde_json::{Value, json};
use zip::ZipArchive;
use crate::xlsx_parts;

const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_REL_NS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";

#[derive(Debug, Default)]
pub(crate) struct TableSidecar {
    pub(crate) tables: Vec<Value>,
    pub(crate) warnings: Vec<String>,
    pub(crate) failure: Option<String>,
    colors: crate::sheet_filter_color::FilterColors,
}

impl TableSidecar {
    pub(crate) fn read(bytes: &[u8]) -> Self {
        match read_tables(bytes) {
            Ok(sidecar) => sidecar,
            Err(error) => Self { failure: Some(error), ..Self::default() },
        }
    }

    /// Reuse the worksheet predicate adapter with actual typed source cells.
    /// Unprojected date/extension criteria remain in the complete table XML.
    pub(crate) fn restore_filters(&mut self, sheets: &mut [crate::model::SheetPayload], native_hidden_rows: &std::collections::BTreeMap<String, Vec<u32>>) {
        for table in &mut self.tables {
            let Some(native) = table.as_object_mut().and_then(|table| table.remove("nativeAutoFilter")) else { continue; };
            let Some(sheet_index) = table.get("sheetName").and_then(Value::as_str).and_then(|name| sheets.iter().position(|sheet| sheet.name == name)) else { continue; };
            let sheet = &sheets[sheet_index];
            let visible_rows = native_hidden_rows.get(&sheet.name).map(Vec::as_slice).unwrap_or(&sheet.hidden_rows);
            let saved = table.as_object_mut().and_then(|table| table.remove("projectAutoFilterState"));
            let result = serde_json::from_value::<office_io_common::OfficeFilterNode>(native).map_err(|error| error.to_string())
                .and_then(|node| {
                    if let Some(saved) = &saved {
                        let fingerprint = office_io_common::office_filter_fingerprint(Some(&node), |node| self.colors.signature(node))?;
                        if saved.get("fingerprint").and_then(Value::as_str) == Some(fingerprint.as_str()) {
                            let filter: crate::sheet_filter_import::AutoFilterState = serde_json::from_value(saved["state"].clone()).map_err(|error| error.to_string())?;
                            crate::sheet_filter_import::validate(&filter)?; return Ok(filter);
                        }
                    }
                    crate::sheet_filter_import::standard_filter_with_colors(&node, &sheet.cells, visible_rows, &self.colors)
                })
                .and_then(|filter| serde_json::to_value(filter).map_err(|error| error.to_string()));
            match result {
                Ok(filter) => { if let Some(table) = table.as_object_mut() { table.insert("autoFilter".into(), filter); } }
                Err(error) => {
                    self.warnings.push(format!("TABLE_AUTOFILTER_RETAINED: {}: {error}", table.get("name").and_then(Value::as_str).unwrap_or("table")));
                    // Exact axis metadata must not discard actual Office hides
                    // when a color resource or predicate cannot be projected.
                    let first = table["range"]["startRow"].as_u64().unwrap_or(0);
                    let last = table["range"]["endRow"].as_u64().unwrap_or(0);
                    let first_column = table["range"]["startCol"].as_u64().unwrap_or(0);
                    let last_column = table["range"]["endCol"].as_u64().unwrap_or(0);
                    let sheet = &mut sheets[sheet_index];
                    if let Some(rows) = native_hidden_rows.get(&sheet.name) {
                        let merged = sheet.merged_ranges.iter().filter(|merge| u64::from(merge.first_row) <= last && u64::from(merge.last_row) > first
                            && u64::from(merge.first_col) <= last_column && u64::from(merge.last_col) >= first_column)
                            .map(|merge| (merge.first_row, merge.last_row)).collect::<Vec<_>>();
                        sheet.hidden_rows.extend(rows.iter().copied().filter(|row| u64::from(*row) > first
                            && (u64::from(*row) <= last || merged.iter().any(|(first, last)| *first <= *row && *row <= *last))));
                        sheet.hidden_rows.sort_unstable(); sheet.hidden_rows.dedup();
                    }
                }
            }
        }
    }
}

fn sheet_namespace(node: &Element) -> bool {
    matches!(node.namespace.as_str(), native_xml::MAIN_NAMESPACE | native_xml::STRICT_MAIN_NAMESPACE)
}

fn read_tables(bytes: &[u8]) -> Result<TableSidecar, String> {
    let mut archive = ZipArchive::new(Cursor::new(bytes)).map_err(|error| format!("TABLE_SOURCE_INVALID: {error}"))?;
    let mut tables = Vec::new(); let mut warnings = Vec::new(); let mut native_ids = BTreeSet::new(); let mut names = BTreeSet::new();
    let parts = xlsx_parts::worksheet_parts(&mut archive)?;
    let mut retained_bytes = 0usize;
    for (sheet_index, sheet) in parts.iter().enumerate() {
        let sheet_xml = xlsx_parts::required_xml_part(&mut archive, &sheet.path)?;
        let identifiers = table_part_ids(&sheet_xml)?;
        if identifiers.is_empty() { continue; }
        let rels = xlsx_parts::required_xml_part(&mut archive, &xlsx_parts::relationship_part_path(&sheet.path))?;
        let rels = xlsx_parts::relationships(&rels)?;
        for id in identifiers {
            let rel = rels.iter().find(|rel| rel.id == id).filter(|rel| rel.kind == "table" && !rel.external)
                .ok_or_else(|| "TABLE_SOURCE_INVALID: missing internal table relationship".to_owned())?;
            let path = xlsx_parts::resolve_target(&sheet.path, &rel.target).ok_or_else(|| "TABLE_SOURCE_INVALID: invalid table target".to_owned())?;
            let xml = xlsx_parts::required_xml_part(&mut archive, &path)?;
            retained_bytes = retained_bytes.checked_add(xml.len()).filter(|bytes| *bytes <= 64 * 1024 * 1024)
                .ok_or_else(|| "TABLE_SOURCE_LIMIT: table XML exceeds 64 MiB".to_owned())?;
            let table = sheet_table_source::parse_xml(&xml)?;
            if table.name != "table" || !sheet_namespace(&table) { return Err("TABLE_SOURCE_INVALID: invalid table root or namespace".into()); }
            let native_id = table.attrs.get("id").and_then(|id| id.parse::<u32>().ok()).filter(|id| *id > 0)
                .ok_or_else(|| "TABLE_SOURCE_INVALID: missing native table ID".to_owned())?;
            let name = table.attrs.get("displayName").or_else(|| table.attrs.get("name")).filter(|name| !name.is_empty())
                .ok_or_else(|| "TABLE_SOURCE_INVALID: missing table name".to_owned())?;
            if !native_ids.insert(native_id) || !names.insert(name.to_lowercase()) { return Err("TABLE_SOURCE_INVALID: duplicate table identity".into()); }
            let range = table.attrs.get("ref").and_then(|reference| parse_range(reference))
                .ok_or_else(|| "TABLE_SOURCE_INVALID: invalid table range".to_owned())?;
            let column_root = table.children.iter().find(|node| node.name == "tableColumns" && sheet_namespace(node))
                .ok_or_else(|| "TABLE_SOURCE_INVALID: missing table columns".to_owned())?;
            let mut columns = Vec::new(); let mut column_ids = BTreeSet::new(); let mut column_names = BTreeSet::new();
            for column in column_root.children.iter().filter(|node| node.name == "tableColumn" && sheet_namespace(node)) {
                let id = column.attrs.get("id").and_then(|id| id.parse::<u32>().ok()).filter(|id| *id > 0)
                    .ok_or_else(|| "TABLE_SOURCE_INVALID: invalid table column ID".to_owned())?;
                let caption = column.attrs.get("name").filter(|caption| !caption.is_empty())
                    .ok_or_else(|| "TABLE_SOURCE_INVALID: missing table column name".to_owned())?;
                if !column_ids.insert(id) || !column_names.insert(caption.to_lowercase()) { return Err("TABLE_SOURCE_INVALID: duplicate table column".into()); }
                let calculated = column.children.iter().find(|node| node.name == "calculatedColumnFormula" && sheet_namespace(node));
                let totals_formula = column.children.iter().find(|node| node.name == "totalsRowFormula" && sheet_namespace(node));
                let filter = table.child("autoFilter").and_then(|filter| filter.children.iter().find(|node| node.name == "filterColumn" && sheet_namespace(node)
                    && node.attrs.get("colId").and_then(|value| value.parse::<usize>().ok()) == Some(columns.len())));
                let show_filter_button = filter.map(|filter| Ok::<_, String>(!sheet_table_source::flag(filter, "hiddenButton")?.unwrap_or(false)
                    && sheet_table_source::flag(filter, "showButton")?.unwrap_or(true))).transpose()?.unwrap_or(true);
                let mut value = json!({ "id": format!("{name}-column-{id}"), "name": caption, "officeColumnId": id,
                    "formula": calculated.map(|formula| format!("={}", formula.text.trim_start_matches('='))),
                    "totalsRowFormula": totals_formula.map(|formula| format!("={}", formula.text.trim_start_matches('='))),
                    "totalsRowFunction": column.attrs.get("totalsRowFunction"), "totalsRowLabel": column.attrs.get("totalsRowLabel"),
                    "showFilterButton": show_filter_button });
                if let Some(calculated) = calculated {
                    if let Some(array) = sheet_table_source::flag(calculated, "array")? { value["formulaIsArray"] = json!(array); }
                }
                columns.push(value);
            }
            if columns.len() != (range[3] - range[1] + 1) as usize { return Err("TABLE_SOURCE_INVALID: column count does not match table range".into()); }
            let style_name = table.children.iter().find(|node| node.name == "tableStyleInfo" && sheet_namespace(node))
                .and_then(|style| style.attrs.get("name")).cloned().unwrap_or_else(|| "TableStyleMedium2".into());
            let row_count = |name: &str, default: u32| -> Result<u32, String> {
                table.attrs.get(name).map_or(Ok(default), |value| value.parse::<u32>()
                    .ok().filter(|count| *count <= 1).ok_or_else(|| format!("TABLE_SOURCE_INVALID: invalid {name}")))
            };
            let header_count = row_count("headerRowCount", 1)?;
            let totals_count = row_count("totalsRowCount", 0)?;
            if header_count + totals_count > range[2] - range[0] + 1 { return Err("TABLE_SOURCE_INVALID: header and totals rows exceed the table range".into()); }
            // totalsRowShown records whether totals have ever been shown. The
            // current geometry is determined only by totalsRowCount (default 0).
            sheet_table_source::flag(&table, "totalsRowShown")?;
            let header = header_count == 1;
            let totals = totals_count == 1;
            let mut value = json!({ "id": name, "name": name, "officeTableId": native_id,
                "sheetName": sheet.name, "sheetIndex": sheet_index, "officeSheetId": sheet.office_sheet_id,
                "range": { "startRow": range[0], "startCol": range[1], "endRow": range[2], "endCol": range[3] },
                "columns": columns, "showHeaderRow": header, "showTotalsRow": totals, "styleName": style_name,
                "officeSource": { "partName": path, "xml": xml } });
            if let Some(style) = table.children.iter().find(|node| node.name == "tableStyleInfo" && sheet_namespace(node)) {
                for name in ["showFirstColumn", "showLastColumn", "showRowStripes", "showColumnStripes"] {
                    if let Some(flag) = sheet_table_source::flag(style, name)? { value[name] = json!(flag); }
                }
            }
            if let Some(filter) = table.child("autoFilter") { value["nativeAutoFilter"] = serde_json::to_value(filter_node(filter)).map_err(|error| error.to_string())?; }
            else { value["autoFilter"] = Value::Null; }
            if let Some(extensions) = table.child("extLst") {
                use office_io_common::filter_color::{TABLE_FILTER_NAMESPACE, TABLE_FILTER_EXTENSION_URI};
                let snapshots = extensions.children.iter().filter(|node| node.is_main("ext") && node.attribute_ns("", "uri") == Some(TABLE_FILTER_EXTENSION_URI)).collect::<Vec<_>>();
                if snapshots.len() == 1 {
                    let states = snapshots[0].children.iter().filter(|node| node.is(TABLE_FILTER_NAMESPACE, "filterState")).collect::<Vec<_>>();
                    if states.len() == 1 && states[0].attribute_ns("", "version") == Some("1") && states[0].children.is_empty() {
                        let snapshot = states[0];
                        match serde_json::from_str::<crate::sheet_filter_import::AutoFilterState>(&snapshot.text).map_err(|error| error.to_string())
                            .and_then(|state| { crate::sheet_filter_import::validate(&state)?; serde_json::to_value(state).map_err(|error| error.to_string()) }) {
                            Ok(state) => value["projectAutoFilterState"] = json!({ "fingerprint": snapshot.attribute_ns("", "fingerprint"), "state": state }),
                            Err(error) => warnings.push(format!("TABLE_AUTOFILTER_SNAPSHOT_INVALID: {name}: {error}; using current Office criteria")),
                        }
                    }
                }
            }
            tables.push(value);
            if tables.len() > 24_576 { return Err("TABLE_SOURCE_LIMIT: too many table identities".into()); }
        }
    }
    let nodes = tables.iter().filter_map(|table| table.get("nativeAutoFilter")).cloned()
        .map(serde_json::from_value::<office_io_common::OfficeFilterNode>).collect::<Result<Vec<_>, _>>().map_err(|error| error.to_string())?;
    let colors = match crate::sheet_filter_color::referenced_ids(nodes.iter()).and_then(|ids| crate::sheet_filter_color::FilterColors::read(&mut archive, &ids)) {
        Ok(colors) => colors,
        Err(error) => { warnings.push(error.clone()); crate::sheet_filter_color::FilterColors::failed(error) }
    };
    Ok(TableSidecar { tables, warnings, colors, failure: None })
}

fn filter_node(node: &Element) -> office_io_common::OfficeFilterNode {
    office_io_common::OfficeFilterNode::from_element(node)
}

fn table_part_ids(xml: &str) -> Result<Vec<String>, String> {
    let mut reader = NsReader::from_str(xml); let mut path = Vec::new(); let mut root = false;
    let mut identifiers = Vec::new(); let mut seen = BTreeSet::new();
    loop {
        let event = reader.read_event().map_err(|error| error.to_string())?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                let name = element.local_name(); let namespace = reader.resolver().resolve_element(element.name()).0;
                let native = xlsx_parts::spreadsheet_namespace(namespace);
                if path.is_empty() {
                    if root || name.as_ref() != "worksheet" || !native { return Err("TABLE_SOURCE_INVALID: invalid worksheet root".into()); }
                    root = true;
                }
                if path.len() == 2 && path[0] == b"worksheet" && path[1] == b"tableParts" && name.as_ref() == "tablePart" && native {
                    let mut id = None;
                    for attribute in element.attributes() {
                        let attribute = attribute.map_err(|error| error.to_string())?;
                        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
                        if local.as_ref() == "id" && matches!(namespace, ResolveResult::Bound(value) if value.as_ref() == REL_NS || value.as_ref() == STRICT_REL_NS) {
                            if id.is_some() { return Err("TABLE_SOURCE_INVALID: duplicate relationship attribute".into()); }
                            id = Some(attribute.normalized_value(XmlVersion::Implicit1_0).map_err(|error| error.to_string())?.into_owned());
                        }
                    }
                    let id = id.filter(|id| !id.is_empty()).ok_or_else(|| "TABLE_SOURCE_INVALID: missing table relationship ID".to_owned())?;
                    if !seen.insert(id.clone()) { return Err("TABLE_SOURCE_INVALID: duplicate table part reference".into()); }
                    identifiers.push(id);
                }
                if matches!(event, Event::Start(_)) { path.push(name.as_ref().as_bytes().to_vec()); }
                if path.len() > 128 { return Err("TABLE_SOURCE_LIMIT: worksheet nesting exceeds 128".into()); }
            }
            Event::End(_) => { path.pop().ok_or_else(|| "TABLE_SOURCE_INVALID: unmatched worksheet end tag".to_owned())?; }
            Event::DocType(_) => return Err("TABLE_SOURCE_INVALID: worksheet DTD is unsupported".into()),
            Event::Eof => break,
            _ => {}
        }
    }
    if !root || !path.is_empty() { return Err("TABLE_SOURCE_INVALID: incomplete worksheet XML".into()); }
    Ok(identifiers)
}

fn parse_range(reference: &str) -> Option<[u32; 4]> {
    sheet_table_source::range(reference)
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn table_relationship_attributes_use_the_native_namespace_and_direct_location() {
        let xml = format!("<s:worksheet xmlns:s=\"http://purl.oclc.org/ooxml/spreadsheetml/main\" xmlns:r=\"{STRICT_REL_NS}\"><s:tableParts><s:tablePart r:id=\"table-2\"/></s:tableParts></s:worksheet>");
        assert_eq!(table_part_ids(&xml).unwrap(), ["table-2"]);
        assert!(table_part_ids(&xml.replace(STRICT_REL_NS, "urn:unrelated")).is_err());
        assert!(table_part_ids(&xml.replace("<s:tableParts>", "<s:extLst><s:tableParts>").replace("</s:tableParts>", "</s:tableParts></s:extLst>")).unwrap().is_empty());
    }
    #[test]
    fn table_ranges_reject_reversal_and_native_axis_overflow() {
        assert_eq!(parse_range("A1:XFD1048576"), Some([0, 0, 1_048_575, 16_383]));
        for range in ["A0:B2", "B2:A1", "A1:XFE2", "A1:B1048577", "A1:", "1:B2"] { assert!(parse_range(range).is_none(), "{range}"); }
    }

    fn exported_native_table() -> Vec<u8> {
        let xml = format!("<table xmlns=\"{}\" id=\"31\" name=\"SalesData\" displayName=\"SalesData\" ref=\"A1:B3\" totalsRowCount=\"0\" totalsRowShown=\"1\"><tableColumns count=\"2\"><tableColumn id=\"7\" name=\"Item\"/><tableColumn id=\"9\" name=\"Amount\"><calculatedColumnFormula array=\"1\">2+3</calculatedColumnFormula><totalsRowFormula>SUM([Amount])</totalsRowFormula></tableColumn></tableColumns><tableStyleInfo name=\"TableStyleMedium4\" showFirstColumn=\"1\" showRowStripes=\"0\"/></table>", native_xml::MAIN_NAMESPACE);
        office_writer::export_xlsx_bytes(&json!({ "kind": "workbook", "sheets": [{ "id": "data", "name": "Data", "cells": {
            "A1": { "value": "Item" }, "B1": { "value": "Amount" }, "A2": { "value": "A" }, "B2": { "value": 5 },
            "A3": { "value": "B" }, "B3": { "value": 12 } } }], "tables": [{ "id": "stable-table", "name": "SalesData", "sheetId": "data", "officeTableId": 31,
            "range": { "startRow": 0, "startCol": 0, "endRow": 2, "endCol": 1 }, "showHeaderRow": true, "showTotalsRow": false,
            "columns": [{ "id": "item", "name": "Item", "officeColumnId": 7 }, { "id": "amount", "name": "Amount", "officeColumnId": 9 }],
            "officeSource": { "partName": "xl/tables/original.xml", "xml": xml } }] }).to_string()).unwrap().0
    }

    #[test]
    fn binary_reopen_preserves_formula_attributes_and_does_not_mistake_totals_history_for_a_row() {
        let bytes = exported_native_table();
        let envelope = crate::parser::parse_office_bytes(&bytes, "native-table.xlsx").unwrap();
        let envelope = serde_json::to_value(envelope).unwrap(); let table = &envelope["payload"]["data"]["tables"][0];
        assert_eq!(table["officeTableId"], 31); assert_eq!(table["showTotalsRow"], false);
        assert_eq!(table["range"]["endRow"], 2); assert_eq!(table["columns"][1]["officeColumnId"], 9);
        assert_eq!(table["columns"][1]["formula"], "=2+3"); assert_eq!(table["columns"][1]["formulaIsArray"], true);
        assert_eq!(table["columns"][1]["totalsRowFormula"], "=SUM([Amount])");
        assert_eq!(table["showFirstColumn"], true); assert_eq!(table["showRowStripes"], false);
        assert!(table["officeSource"]["xml"].as_str().unwrap().contains("totalsRowShown=\"1\""));
        let manifest = &envelope["payload"]["data"]["formulaCalendarSource"]["tables"][0];
        assert_eq!(manifest["tableId"], "stable-table"); assert_eq!(manifest["columns"][1]["columnId"], "amount");
    }
}
