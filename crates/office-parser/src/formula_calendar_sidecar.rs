//! Recover authored formulas only after native identities and calculation views
//! prove freshness. Native Office edits remain authoritative when a guard fails.

use std::{collections::{BTreeMap, BTreeSet}, io::{BufReader, Cursor, Read}};
use office_io_common::{excel_date::convert_serial, formula_references::{canonical_formula, rewrite_sheet_qualifiers, rewrite_table_identifiers}};
use office_io_common::sheet_formula_calendar_source::{self as source, CalculationFingerprint, CellSource, FormulaCalendarSource, HelperSource};
use quick_xml::{NsReader, events::Event, name::ResolveResult};
use serde_json::{Value, json};
use zip::ZipArchive;
use crate::{model::{CellPayload, SheetPayload, WorkbookPayload}, xlsx_parts::{self, WorkbookParts}};

#[path = "formula_calendar_reference_scan.rs"]
mod reference_scan;

type Identities = BTreeMap<String, String>;
type Coordinates = BTreeMap<(u32, u32), usize>;

pub(crate) fn restore(bytes: &[u8], workbook: &mut WorkbookPayload) -> Result<Vec<String>, String> {
    let mut archive = ZipArchive::new(Cursor::new(bytes)).map_err(|error| format!("FORMULA_CALENDAR_SOURCE_INVALID: {error}"))?;
    let parts = xlsx_parts::workbook_parts(&mut archive)?;
    let Some((part_name, mut saved)) = read_source(&mut archive, &parts)? else { return Ok(Vec::new()); };
    if saved.date1904 != workbook.date1904 {
        return Ok(vec!["XLSX_FORMULA_CALENDAR_NATIVE_EDITS: workbook calendar changed; retained native formulas and calculation sheets".into()]);
    }
    let sheet_indices = parts.sheets.iter().filter_map(|part| workbook.sheets.iter().position(|sheet| sheet.name == part.name)
        .map(|index| (part.office_sheet_id, index))).collect::<BTreeMap<_, _>>();
    let expected_sheets = saved.sheets.iter().map(|sheet| (sheet.name.to_lowercase(), sheet.sheet_id.clone()))
        .chain(saved.helpers.iter().map(|helper| (helper.name.to_lowercase(), helper.sheet_id.clone()))).collect::<Identities>();
    let observed_sheets = saved.sheets.iter().map(|sheet| (sheet.office_sheet_id, &sheet.sheet_id))
        .chain(saved.helpers.iter().map(|helper| (helper.office_sheet_id, &helper.sheet_id)))
        .filter_map(|(native, id)| sheet_indices.get(&native).map(|index| (workbook.sheets[*index].name.to_lowercase(), id.clone())))
        .collect::<Identities>();
    let expected_tables = saved.tables.iter().map(|table| (table.name.to_lowercase(), table.table_id.clone())).collect::<Identities>();
    let table_indices = workbook.tables.iter().enumerate().filter_map(|(index, table)| table.get("officeTableId").and_then(Value::as_u64).map(|id| (id, index)))
        .collect::<BTreeMap<_, _>>();
    let observed_tables = saved.tables.iter().filter_map(|table| table_indices.get(&u64::from(table.office_table_id))
        .and_then(|index| workbook.tables[*index].get("name").and_then(Value::as_str)).map(|name| (name.to_lowercase(), table.table_id.clone())))
        .collect::<Identities>();
    let renamed_sheets = saved.sheets.iter().map(|sheet| (sheet.name.to_lowercase(), sheet_indices.get(&sheet.office_sheet_id)
        .map(|index| workbook.sheets[*index].name.clone()).unwrap_or_else(|| "#REF!".into()))).collect::<Identities>();
    let renamed_tables = saved.tables.iter().map(|table| (table.name.to_lowercase(), table_indices.get(&u64::from(table.office_table_id))
        .and_then(|index| workbook.tables[*index].get("name").and_then(Value::as_str)).unwrap_or("#REF!").to_owned())).collect::<Identities>();
    let helper_names = saved.helpers.iter().enumerate().map(|(index, helper)| (helper.name.as_str(), index)).collect::<BTreeMap<_, _>>();
    let coordinates = workbook.sheets.iter().map(|sheet| sheet.cells.iter().enumerate().map(|(index, cell)| ((cell.row, cell.col), index)).collect::<Coordinates>())
        .collect::<Vec<_>>();
    let mut fresh = Vec::new();
    for helper in &saved.helpers {
        let matches = if let Some(index) = sheet_indices.get(&helper.office_sheet_id) {
            helper_fingerprint(&workbook.sheets[*index], workbook, &saved, &observed_sheets, &observed_tables)? == helper.fingerprint
        } else { false };
        fresh.push(matches);
    }
    // Mirrors may depend on either calendar and the shared INDIRECT catalog.
    // Until every member is fresh, retain the native family rather than revive
    // an authored expression from a calculation graph with changed members.
    let family_fresh = fresh.iter().all(|fresh| *fresh);
    let mut restored = 0usize; let mut stale = 0usize;
    let mut approved = BTreeMap::<String, BTreeSet<(u32, u32)>>::new();
    let mut restored_sheets = Vec::new();
    for sheet in &saved.sheets {
        let Some(index) = sheet_indices.get(&sheet.office_sheet_id) else { stale += sheet.cells.len(); continue; };
        let native = &workbook.sheets[*index]; let mut descriptor = sheet.clone(); descriptor.name = native.name.clone(); descriptor.cells.clear();
        if let Some(import) = &mut descriptor.import { rename_alias_targets(import, &renamed_sheets, &renamed_tables); }
        for cell in &sheet.cells {
            let Some(native_index) = coordinates[*index].get(&(cell.row, cell.col)) else { stale += 1; continue; };
            let actual = &native.cells[*native_index];
            let mut restored_cell = cell.clone();
            let authored = cell.formula.as_ref().map(|formula| equal_formula(&rewrite_table_identifiers(&rewrite_sheet_qualifiers(formula, &renamed_sheets), &renamed_tables)));
            if let Some(marker) = &cell.formula_source {
                let helper = cell.calculation_sheet.as_deref().and_then(|name| helper_names.get(name)).and_then(|index| saved.helpers.get(*index));
                let calculation = helper.and_then(|helper| sheet_indices.get(&helper.office_sheet_id).map(|index| (helper, *index)))
                    .and_then(|(helper, index)| coordinates[index].get(&(cell.row, cell.col)).map(|cell_index| (helper, &workbook.sheets[index].cells[*cell_index])));
                let author_matches = formulas_match(actual.formula.as_deref(), cell.projected_formula.as_deref(), &observed_sheets, &observed_tables, &expected_sheets, &expected_tables);
                let calculation_matches = calculation.is_some_and(|(_, calculation)| formulas_match(calculation.formula.as_deref(), cell.calculation_formula.as_deref(),
                    &observed_sheets, &observed_tables, &expected_sheets, &expected_tables));
                let alias_owner_exists = marker.get("aliasSheetId").and_then(Value::as_str).is_none_or(|id| observed_sheets.values().any(|candidate| candidate == id));
                if !family_fresh || !author_matches || !calculation_matches || !alias_owner_exists { stale += 1; continue; }
                let (helper, calculation) = calculation.ok_or_else(|| "FORMULA_CALENDAR_SOURCE_INVALID: missing validated calculation cell".to_owned())?;
                restored_cell.formula = authored.clone();
                restored_cell.projected_formula = actual.formula.as_ref().map(|formula| equal_formula(formula));
                restored_cell.calculation_sheet = Some(workbook.sheets[sheet_indices[&helper.office_sheet_id]].name.clone());
                restored_cell.calculation_formula = calculation.formula.as_ref().map(|formula| equal_formula(formula));
                if let Some(marker) = restored_cell.formula_source.as_mut().and_then(Value::as_object_mut) {
                    marker.insert("formula".into(), json!(authored));
                }
                restored_cell.date_source = if marker.get("dateResult").and_then(Value::as_bool) == Some(true) {
                    matching_calculation_date(calculation, actual, helper, workbook.date1904, authored.as_deref())
                } else { None };
                if let Some(part) = parts.sheets.iter().find(|part| part.office_sheet_id == sheet.office_sheet_id) {
                    approved.entry(part.path.clone()).or_default().insert((cell.row, cell.col));
                }
                restored += 1;
            } else if cell.date_source.is_some() {
                if !formulas_match_or_blank(actual.formula.as_deref(), cell.formula.as_deref(), &observed_sheets, &observed_tables, &expected_sheets, &expected_tables)
                    || !matching_retained_date(cell, actual, workbook.date1904) { stale += 1; continue; }
                restored_cell.formula = actual.formula.as_ref().map(|formula| equal_formula(formula));
                if let Some(marker) = restored_cell.date_source.as_mut().and_then(Value::as_object_mut) {
                    marker.insert("targetValue".into(), actual.value.clone());
                    marker.insert("targetDate1904".into(), json!(workbook.date1904));
                    marker.insert("targetFormula".into(), json!(restored_cell.formula));
                }
                restored += 1;
            } else { continue; }
            descriptor.cells.push(restored_cell);
        }
        restored_sheets.push(descriptor);
    }
    drop(helper_names);
    let mut helpers = Vec::new();
    for (index, mut helper) in saved.helpers.into_iter().enumerate() {
        let Some(native_index) = sheet_indices.get(&helper.office_sheet_id) else { continue; };
        if helper.source_sheet_id.as_ref().is_some_and(|id| !restored_sheets.iter().any(|sheet| &sheet.sheet_id == id)) { continue; }
        helper.name = workbook.sheets[*native_index].name.clone(); helper.reclaim = false;
        helpers.push((helper, fresh[index]));
    }
    let mut retained_tables = Vec::new();
    for mut table in saved.tables {
        let Some(index) = table_indices.get(&u64::from(table.office_table_id)) else { continue; };
        let Some(name) = workbook.tables[*index].get("name").and_then(Value::as_str) else { continue; };
        let Some(owner_name) = workbook.tables[*index].get("sheetName").and_then(Value::as_str) else { continue; };
        let Some(owner_id) = observed_sheets.get(&owner_name.to_lowercase()) else { continue; };
        if !restored_sheets.iter().any(|sheet| &sheet.sheet_id == owner_id) && !helpers.iter().any(|(helper, _)| &helper.sheet_id == owner_id) { continue; }
        table.name = name.to_owned(); table.sheet_id = owner_id.clone(); retained_tables.push(table);
    }
    let mut warnings = Vec::new();
    let can_reclaim = if family_fresh {
        reference_scan::can_reclaim(&mut archive, &parts, &part_name, workbook, &helpers, &retained_tables, &approved, &mut warnings)?
    } else { false };
    for (helper, is_fresh) in &mut helpers { helper.reclaim = can_reclaim && *is_fresh; }
    saved = FormulaCalendarSource { version: 1, date1904: workbook.date1904, sheets: restored_sheets,
        helpers: helpers.into_iter().map(|(helper, _)| helper).collect(), tables: retained_tables };
    saved.validate()?;
    if restored > 0 { warnings.push(format!("XLSX_FORMULA_CALENDAR_RESTORED: retained provenance for {restored} native-matching cells")); }
    if stale > 0 { warnings.push(format!("XLSX_FORMULA_CALENDAR_NATIVE_EDITS: kept native formulas/values for {stale} unmatched source cells")); }
    workbook.formula_calendar_source = Some(saved);
    Ok(warnings)
}

fn equal_formula(formula: &str) -> String { format!("={}", formula.trim_start_matches('=')) }

fn formulas_match(actual: Option<&str>, expected: Option<&str>, observed_sheets: &Identities, observed_tables: &Identities,
    expected_sheets: &Identities, expected_tables: &Identities) -> bool {
    actual.zip(expected).is_some_and(|(actual, expected)| canonical_formula(actual, observed_sheets, observed_tables) == canonical_formula(expected, expected_sheets, expected_tables))
}
fn formulas_match_or_blank(actual: Option<&str>, expected: Option<&str>, observed_sheets: &Identities, observed_tables: &Identities,
    expected_sheets: &Identities, expected_tables: &Identities) -> bool {
    actual.is_none() && expected.is_none() || formulas_match(actual, expected, observed_sheets, observed_tables, expected_sheets, expected_tables)
}

fn matching_retained_date(saved: &CellSource, actual: &CellPayload, native1904: bool) -> bool {
    let Some(marker) = &saved.date_source else { return false; };
    let Some((value, source1904)) = source::finite_value(marker.get("value")).zip(marker.get("date1904").and_then(Value::as_bool)) else { return false; };
    let target = source::finite_value(marker.get("targetValue")).unwrap_or_else(|| convert_serial(value, source1904, native1904));
    actual.value.as_f64() == Some(target)
}

fn matching_calculation_date(calculation: &CellPayload, actual: &CellPayload, helper: &HelperSource, native1904: bool, formula: Option<&str>) -> Option<Value> {
    let value = calculation.value.as_f64()?; let target = actual.value.as_f64()?;
    (convert_serial(value, helper.date1904, native1904) == target).then(|| json!({ "value": value, "date1904": helper.date1904,
        "targetValue": target, "targetDate1904": native1904, "targetFormula": formula }))
}

fn rename_alias_targets(value: &mut Value, sheets: &Identities, tables: &Identities) {
    let Some(aliases) = value.get_mut("formulaAliases").and_then(Value::as_object_mut) else { return; };
    for (key, names) in [("sheets", sheets), ("nameSheets", sheets), ("tables", tables)] {
        for target in aliases.get_mut(key).and_then(Value::as_object_mut).into_iter().flat_map(|values| values.values_mut()) {
            if let Some(before) = target.as_str() && let Some(after) = names.get(&before.to_lowercase()) { *target = json!(after); }
        }
    }
}

fn helper_fingerprint(sheet: &SheetPayload, workbook: &WorkbookPayload, saved: &FormulaCalendarSource, sheets: &Identities, tables: &Identities) -> Result<String, String> {
    let mut fingerprint = CalculationFingerprint::default();
    for cell in &sheet.cells {
        if cell.formula.is_none() && cell.value.is_null() { continue; }
        fingerprint.cell(cell.row, cell.col, cell.formula.as_deref(), &cell.value, sheets, tables)?;
    }
    for range in &sheet.merged_ranges { fingerprint.merged_range((range.first_row, range.first_col, range.last_row, range.last_col))?; }
    for name in workbook.local_defined_names.iter().filter(|name| name.sheet.to_lowercase() == sheet.name.to_lowercase()) {
        fingerprint.defined_name(&name.name, &name.refers_to, sheets, tables)?;
    }
    for table in workbook.tables.iter().filter(|table| table.get("sheetName").and_then(Value::as_str).is_some_and(|name| name.to_lowercase() == sheet.name.to_lowercase())) {
        let native_id = table.get("officeTableId").and_then(Value::as_u64).ok_or_else(|| "FORMULA_CALENDAR_SOURCE_INVALID: helper table has no native identity".to_owned())?;
        let foreign_id = format!("foreign-native-table:{native_id}");
        let id = saved.tables.iter().find(|candidate| u64::from(candidate.office_table_id) == native_id).map(|table| table.table_id.as_str())
            .unwrap_or(&foreign_id);
        let range = table.get("range").ok_or_else(|| "FORMULA_CALENDAR_SOURCE_INVALID: helper table has no range".to_owned())?;
        let index = |key| range.get(key).and_then(Value::as_u64).and_then(|value| u32::try_from(value).ok())
            .ok_or_else(|| "FORMULA_CALENDAR_SOURCE_INVALID: helper table range is invalid".to_owned());
        let columns = table.get("columns").and_then(Value::as_array).ok_or_else(|| "FORMULA_CALENDAR_SOURCE_INVALID: missing helper table columns".to_owned())?
            .iter().map(|column| column.get("name").and_then(Value::as_str).map(str::to_owned).ok_or_else(|| "FORMULA_CALENDAR_SOURCE_INVALID: invalid helper table column".to_owned())).collect::<Result<Vec<_>, _>>()?;
        fingerprint.table(id, (index("startRow")?, index("startCol")?, index("endRow")?, index("endCol")?), &columns,
            table.get("showHeaderRow").and_then(Value::as_bool).unwrap_or(true), table.get("showTotalsRow").and_then(Value::as_bool).unwrap_or(false))?;
    }
    Ok(fingerprint.finish())
}

fn read_source(archive: &mut ZipArchive<Cursor<&[u8]>>, parts: &WorkbookParts) -> Result<Option<(String, FormulaCalendarSource)>, String> {
    let mut result = None;
    for part in &parts.custom_xml {
        let entry = archive.by_name(part).map_err(|error| format!("FORMULA_CALENDAR_SOURCE_INVALID: missing custom XML part {part}: {error}"))?;
        let mut reader = NsReader::from_reader(BufReader::new(entry.take(64 * 1024))); let mut buffer = Vec::new(); let mut owned = false;
        loop {
            let event = reader.read_event_into(&mut buffer).map_err(|error| error.to_string())?;
            match event {
                Event::Start(ref root) | Event::Empty(ref root) => {
                    owned = root.local_name().as_ref() == "formulaCalendars" && matches!(reader.resolver().resolve_element(root.name()).0,
                        ResolveResult::Bound(namespace) if namespace.as_ref() == source::NAMESPACE);
                    break;
                }
                Event::DocType(_) => break,
                Event::Eof => break,
                _ => {}
            }
            if reader.buffer_position() > 64 * 1024 { break; }
            buffer.clear();
        }
        drop(reader);
        if !owned { continue; }
        if result.is_some() { return Err("FORMULA_CALENDAR_SOURCE_INVALID: duplicate related source manifests".into()); }
        let xml = xlsx_parts::required_xml_part(archive, part)?;
        result = Some((part.clone(), source::from_xml(&xml)?));
    }
    Ok(result)
}
