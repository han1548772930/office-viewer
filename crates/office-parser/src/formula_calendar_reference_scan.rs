//! Conservative ownership checks before discarding native calculation sheets.
//! An incomplete scan keeps the whole family and reports the remaining evidence.

use std::{collections::{BTreeMap, BTreeSet}, io::Cursor};
use office_io_common::{chart_xml::{self, Node}, formula_references::{identifiers, references as formula_references, sheet_qualifiers}};
use office_io_common::sheet_formula_calendar_source::{HelperSource, TableSource};
use zip::ZipArchive;
use crate::{model::WorkbookPayload, xlsx_parts::{self, WorkbookParts}};

pub(super) fn can_reclaim(archive: &mut ZipArchive<Cursor<&[u8]>>, parts: &WorkbookParts, source_part: &str,
    workbook: &WorkbookPayload, helpers: &[(HelperSource, bool)], tables: &[TableSource],
    approved: &BTreeMap<String, BTreeSet<(u32, u32)>>, warnings: &mut Vec<String>) -> Result<bool, String> {
    let names = helpers.iter().map(|(helper, _)| helper.name.to_lowercase()).collect::<BTreeSet<_>>();
    let helper_ids = helpers.iter().map(|(helper, _)| helper.sheet_id.as_str()).collect::<BTreeSet<_>>();
    let table_names = tables.iter().filter(|table| helper_ids.contains(table.sheet_id.as_str())).map(|table| table.name.to_lowercase()).collect::<BTreeSet<_>>();
    let paths = parts.sheets.iter().filter(|part| names.contains(&part.name.to_lowercase())).map(|part| part.path.as_str()).collect::<BTreeSet<_>>();
    let table_paths = workbook.tables.iter().filter(|table| table.get("name").and_then(serde_json::Value::as_str).is_some_and(|name| table_names.contains(&name.to_lowercase())))
        .filter_map(|table| table.get("officeSource").and_then(|source| source.get("partName")).and_then(serde_json::Value::as_str)).collect::<BTreeSet<_>>();
    let sheet_order = parts.sheets.iter().enumerate().map(|(index, part)| (part.name.to_lowercase(), index)).collect::<BTreeMap<_, _>>();
    let references = |text: &str| resource_referenced(text, &names, &table_names, true)
        || formula_references(text).iter().filter(|reference| reference.prefix.contains(':')).any(|reference| {
            let endpoints = sheet_qualifiers(&reference.prefix);
            endpoints.first().zip(endpoints.last()).and_then(|(first, last)| sheet_order.get(&first.to_lowercase()).zip(sheet_order.get(&last.to_lowercase())))
                .is_some_and(|(first, last)| names.iter().filter_map(|name| sheet_order.get(name)).any(|index| index >= first.min(last) && index <= first.max(last)))
        });
    for sheet in &workbook.sheets {
        let helper = names.contains(&sheet.name.to_lowercase());
        if helper {
            if !sheet.visible.eq_ignore_ascii_case("veryHidden") || !sheet.charts.is_empty() || !sheet.data_validations.is_empty()
                || !sheet.conditional_formats.is_empty() || !sheet.hyperlinks.is_empty() || sheet.cells.iter().any(|cell| cell.sparkline.is_some()) { return Ok(false); }
        }
    }
    if workbook.defined_names.iter().any(|name| references(&name.refers_to))
        || workbook.local_defined_names.iter().filter(|name| !names.contains(&name.sheet.to_lowercase())).any(|name| references(&name.refers_to)) { return Ok(false); }
    for path in &paths {
        let rels_path = xlsx_parts::relationship_part_path(path);
        if archive.by_name(&rels_path).is_err() { continue; }
        let rels = xlsx_parts::required_xml_part(archive, &rels_path)?;
        for rel in xlsx_parts::relationships(&rels)? {
            let target = xlsx_parts::resolve_target(path, &rel.target);
            if rel.external || rel.kind != "table" || target.as_deref().is_none_or(|target| !table_paths.contains(target)) { return Ok(false); }
        }
    }
    let xml_paths = archive.file_names().filter(|path| path.ends_with(".xml")).map(str::to_owned).collect::<Vec<_>>();
    let mut bytes = 0usize;
    for path in xml_paths {
        if path == source_part || paths.contains(path.as_str()) || table_paths.contains(path.as_str())
            || path == "[Content_Types].xml" || path.ends_with("/styles.xml") || path.ends_with("/sharedStrings.xml")
            || path.starts_with("docProps/") { continue; }
        let xml = match xlsx_parts::required_xml_part(archive, &path) {
            Ok(xml) => xml,
            Err(error) => { warnings.push(format!("XLSX_FORMULA_CALENDAR_HELPERS_RETAINED: reference scan could not read {path}: {error}")); return Ok(false); }
        };
        bytes = match bytes.checked_add(xml.len()).filter(|bytes| *bytes <= 64 * 1024 * 1024) {
            Some(bytes) => bytes,
            None => { warnings.push("XLSX_FORMULA_CALENDAR_HELPERS_RETAINED: reference scan reached its 64 MiB XML budget".into()); return Ok(false); }
        };
        let root = match chart_xml::parse(&xml) {
            Ok(root) => root,
            Err(error) => { warnings.push(format!("XLSX_FORMULA_CALENDAR_HELPERS_RETAINED: reference scan could not inspect {path}: {error}")); return Ok(false); }
        };
        if foreign_node(&root, approved.get(&path), None, path == parts.workbook_path, &references) { return Ok(false); }
    }
    Ok(true)
}

fn foreign_node(node: &Node, approved: Option<&BTreeSet<(u32, u32)>>, cell: Option<(u32, u32)>, workbook: bool,
    references: &impl Fn(&str) -> bool) -> bool {
    // Names are inspected through their native scopes above. Sheet declarations
    // establish identities; their captions are not references to calculation data.
    if workbook && matches!(node.name.as_str(), "sheets" | "definedNames") { return false; }
    let cell = if node.name == "c" { node.attrs.get("r").and_then(|address| coordinate(address)) } else { cell };
    if node.name == "f" && cell.is_some_and(|cell| approved.is_some_and(|approved| approved.contains(&cell))) { return false; }
    if !matches!(node.name.as_str(), "t" | "v") && references(&node.text) { return true; }
    if node.attrs.values().any(|value| references(value)) { return true; }
    node.children.iter().any(|child| foreign_node(child, approved, cell, workbook, references))
}

fn coordinate(address: &str) -> Option<(u32, u32)> {
    let split = address.find(|character: char| character.is_ascii_digit())?;
    let mut col = 0u32;
    for byte in address[..split].bytes() {
        if !byte.is_ascii_alphabetic() { return None; }
        col = col.checked_mul(26)?.checked_add(u32::from(byte.to_ascii_uppercase() - b'A') + 1)?;
    }
    Some((address[split..].parse::<u32>().ok()?.checked_sub(1)?, col.checked_sub(1)?))
}

fn resource_referenced(text: &str, sheets: &BTreeSet<String>, tables: &BTreeSet<String>, strings: bool) -> bool {
    let lower = text.to_lowercase();
    // A dynamic reference can read a helper through arbitrary cells or names.
    // Without evaluating an unowned formula, absence cannot be proven.
    if lower.contains("indirect(") || lower.contains("indirect (") { return true; }
    if sheets.contains(&lower) || tables.contains(&lower) { return true; }
    if sheet_qualifiers(text).iter().any(|name| sheets.contains(&name.to_lowercase()))
        || identifiers(text).iter().any(|name| tables.contains(&name.to_lowercase())) { return true; }
    if strings {
        let mut chars = text.chars().peekable();
        while let Some(character) = chars.next() {
            if character != '"' { continue; }
            let mut literal = String::new();
            while let Some(character) = chars.next() {
                if character == '"' {
                    if chars.peek() == Some(&'"') { chars.next(); literal.push('"'); }
                    else { break; }
                } else { literal.push(character); }
            }
            if resource_referenced(&literal, sheets, tables, false) { return true; }
        }
    }
    false
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn external_references_and_indirect_literals_keep_owned_calculation_data() {
        let sheets = BTreeSet::from(["calc".into()]); let tables = BTreeSet::from(["calctable".into()]);
        for formula in ["Calc!A1", "'Calc'!Tax", "SUM(CalcTable[Amount])", "INDIRECT(\"'Calc'!A1\")"] {
            assert!(resource_referenced(formula, &sheets, &tables, true), "{formula}");
        }
        assert!(!resource_referenced("[External.xlsx]Calc!A1", &sheets, &tables, true));
        assert!(!resource_referenced("=SUM(Author!A1:A9)", &sheets, &tables, true));
    }
    #[test]
    fn only_the_approved_cell_formula_is_ignored_and_other_resources_stay_visible_to_the_scan() {
        let root = chart_xml::parse("<worksheet><sheetData><row><c r=\"A1\"><f>Calc!A1</f></c><c r=\"A2\"><f>Calc!A2</f></c></row></sheetData></worksheet>").unwrap();
        let approved = BTreeSet::from([(0, 0)]);
        assert!(foreign_node(&root, Some(&approved), None, false, &|text| text.contains("Calc!")));
        let approved = BTreeSet::from([(0, 0), (1, 0)]);
        assert!(!foreign_node(&root, Some(&approved), None, false, &|text| text.contains("Calc!")));
    }
}
