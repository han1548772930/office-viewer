use std::collections::{BTreeMap, BTreeSet};
use serde::{Deserialize, Deserializer};
use office_io_common::{conditional_format_xml::{self as native_xml, Element}, sheet_table_source::{self as source, OfficeTableSource}};

use crate::{ExportError, ExportWarning, xml};
use super::{RangeInput, SheetInput, range_ref};

#[derive(Debug, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct TableInput {
    pub(crate) id: String,
    pub(crate) name: String,
    pub(crate) sheet_id: String,
    pub(crate) range: RangeInput,
    #[serde(default)]
    pub(crate) columns: Vec<TableColumnInput>,
    #[serde(default = "default_true")]
    pub(crate) show_header_row: bool,
    #[serde(default)]
    pub(crate) show_totals_row: bool,
    #[serde(default = "default_style")]
    pub(crate) style_name: String,
    #[serde(default)]
    #[serde(deserialize_with = "present_optional")]
    pub(super) auto_filter: Option<Option<super::auto_filter::AutoFilterInput>>,
    #[serde(default)]
    pub(crate) office_table_id: Option<u32>,
    #[serde(default)]
    pub(crate) office_source: Option<OfficeTableSource>,
    #[serde(default)]
    show_first_column: Option<bool>,
    #[serde(default)]
    show_last_column: Option<bool>,
    #[serde(default)]
    show_row_stripes: Option<bool>,
    #[serde(default)]
    show_column_stripes: Option<bool>,
}

#[derive(Debug, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct TableColumnInput {
    pub(crate) id: String,
    pub(crate) name: String,
    #[serde(default)]
    pub(crate) office_column_id: Option<u32>,
    #[serde(default, alias = "calculatedColumnFormula", deserialize_with = "present_optional")]
    pub(crate) formula: Option<Option<String>>,
    #[serde(default)]
    pub(crate) formula_is_array: Option<bool>,
    #[serde(default, deserialize_with = "present_optional")]
    pub(crate) totals_row_label: Option<Option<String>>,
    #[serde(default, deserialize_with = "present_optional")]
    pub(crate) totals_row_function: Option<Option<String>>,
    #[serde(default, deserialize_with = "present_optional")]
    pub(crate) totals_row_formula: Option<Option<String>>,
    #[serde(default)]
    show_filter_button: Option<bool>,
}

fn present_optional<'de, D: Deserializer<'de>, T: Deserialize<'de>>(deserializer: D) -> Result<Option<Option<T>>, D::Error> {
    Option::<T>::deserialize(deserializer).map(Some)
}

/// Reserve every requested native identity before allocating replacements, so a
/// new table cannot consume a later imported table's unchanged native ID.
fn allocate_ids(preferred: &[Option<u32>], reserved: &BTreeSet<u32>, protected: &BTreeSet<u32>) -> Result<Vec<u32>, ExportError> {
    let mut occupied = reserved.clone();
    occupied.extend(protected);
    occupied.extend(preferred.iter().flatten().copied());
    let mut claimed = reserved.clone(); let mut next = 1u32;
    preferred.iter().map(|preferred| {
        if let Some(id) = preferred {
            if *id == 0 { return Err(ExportError::invalid("TABLE_SOURCE_INVALID: native identities must be positive")); }
            if claimed.insert(*id) { return Ok(*id); }
        }
        while occupied.contains(&next) {
            next = next.checked_add(1).ok_or_else(|| ExportError::invalid("TABLE_SOURCE_LIMIT: native identity space is exhausted"))?;
        }
        let id = next; occupied.insert(id); claimed.insert(id); Ok(id)
    }).collect()
}

pub(crate) fn prepare_identities(tables: &mut [TableInput], reserved: &[u32]) -> Result<(), ExportError> {
    let mut budget = source::TableSourceBudget::default(); let mut preferred = Vec::with_capacity(tables.len());
    let mut names = BTreeSet::new(); let mut ids = BTreeSet::new();
    for table in tables.iter_mut() {
        if !ids.insert(table.id.clone()) || !names.insert(table.name.to_lowercase()) {
            return Err(ExportError::invalid("TABLE_SOURCE_INVALID: duplicate domain table ID or name"));
        }
        budget.admit(table.office_source.as_ref()).map_err(ExportError::invalid)?;
        let root = table.office_source.as_ref().map(source::read).transpose().map_err(ExportError::invalid)?;
        let original_columns = root.as_ref().map(source::columns).transpose().map_err(ExportError::invalid)?.unwrap_or_default();
        preferred.push(table.office_table_id.or_else(|| root.as_ref().and_then(|root| source::positive_id(root, "id").ok())));
        let column_ids = table.columns.iter().map(|column| column.office_column_id.or_else(|| original_columns.iter()
            .find(|source| source.attrs.get("name").is_some_and(|name| name.to_lowercase() == column.name.to_lowercase()))
            .and_then(|source| source::positive_id(source, "id").ok()))).collect::<Vec<_>>();
        // A removed native column still owns its source XML. New columns must
        // not acquire that identity and inherit its formulas or extensions.
        let protected = original_columns.iter().map(|column| source::positive_id(column, "id"))
            .collect::<Result<BTreeSet<_>, _>>().map_err(ExportError::invalid)?;
        for (column, id) in table.columns.iter_mut().zip(allocate_ids(&column_ids, &BTreeSet::new(), &protected)?) { column.office_column_id = Some(id); }
    }
    let reserved = reserved.iter().copied().collect::<BTreeSet<_>>();
    for (table, id) in tables.iter_mut().zip(allocate_ids(&preferred, &reserved, &BTreeSet::new())?) { table.office_table_id = Some(id); }
    Ok(())
}

/// Mirror cells already contain the authoritative calendar expressions. Native
/// calculated/totals formulas must not replace those expressions in Office.
pub(crate) fn clear_calculation_source(table: &mut TableInput) {
    table.office_table_id = None;
    table.office_source = None;
    table.auto_filter = Some(None);
    for (index, column) in table.columns.iter_mut().enumerate() {
        column.id = format!("{}:column:{}", table.id, index + 1);
        column.office_column_id = None;
        column.formula = Some(None);
        column.formula_is_array = None;
        column.totals_row_formula = Some(None);
        column.totals_row_function = Some(None);
        column.totals_row_label = Some(None);
        column.show_filter_button = None;
    }
}

pub(crate) fn validate(tables: &[TableInput], table: &TableInput) -> Result<(), ExportError> {
    let Some((start_row, start_col, end_row, end_col)) = bounds(&table.range) else {
        return Err(ExportError::invalid(format!("table {} has an invalid range", table.id)));
    };
    if start_row > end_row || start_col > end_col || end_row > super::MAX_ROW_INDEX || end_col > super::MAX_COLUMN_INDEX
        || table.columns.len() != (end_col - start_col + 1) as usize
        || end_row - start_row + 1 < u32::from(table.show_header_row) + u32::from(table.show_totals_row) {
        return Err(ExportError::invalid(format!("table {} has invalid header or column definitions", table.id)));
    }
    if !valid_name(&table.name) {
        return Err(ExportError::invalid(format!("table {} has an invalid name", table.id)));
    }
    let mut names = BTreeSet::new(); let mut ids = BTreeSet::new();
    if table.columns.iter().any(|column| column.id.is_empty() || !ids.insert(column.id.clone())
        || column.name.trim().is_empty() || column.name.encode_utf16().count() > 255 || !names.insert(column.name.to_lowercase())) {
        return Err(ExportError::invalid(format!("table {} has blank or duplicate column names", table.id)));
    }
    if tables.iter().filter(|other| other.id != table.id && other.sheet_id == table.sheet_id).any(|other| {
        bounds(&other.range).is_some_and(|(r1, c1, r2, c2)| start_row <= r2 && end_row >= r1 && start_col <= c2 && end_col >= c1)
    }) {
        return Err(ExportError::invalid(format!("table {} overlaps another table", table.id)));
    }
    for column in &table.columns {
        for formula in [&column.formula, &column.totals_row_formula].into_iter().flatten().flatten() {
            if formula.len() > 32_768 || !valid_text(formula) { return Err(ExportError::invalid("TABLE_SOURCE_INVALID: invalid table formula")); }
        }
        if column.totals_row_function.as_ref().and_then(Option::as_deref).is_some_and(|function| !matches!(function,
            "none" | "sum" | "min" | "max" | "average" | "count" | "countNums" | "stdDev" | "var" | "custom")) {
            return Err(ExportError::invalid("TABLE_SOURCE_INVALID: invalid totals-row function"));
        }
    }
    if let Some(Some(filter)) = &table.auto_filter {
        super::auto_filter::validate_filter(filter, None, None, false)?;
        let filter_bounds = super::range_bounds(&filter.range).ok_or_else(|| ExportError::invalid("TABLE_SOURCE_INVALID: invalid table filter range"))?;
        if !table.show_header_row || filter_bounds.0 != start_row || filter_bounds.1 != start_col || filter_bounds.3 != end_col
            || ![end_row, end_row - u32::from(table.show_totals_row)].contains(&filter_bounds.2) {
            return Err(ExportError::invalid("TABLE_SOURCE_INVALID: filter must cover the table header and data columns"));
        }
    }
    Ok(())
}

pub(crate) fn write(table: &TableInput, id: u32, sheet: &SheetInput, warnings: &mut Vec<ExportWarning>) -> Result<String, ExportError> {
    let reference = range_ref(&table.range).ok_or_else(|| ExportError::invalid("table range is invalid"))?;
    let mut root = table.office_source.as_ref().map(source::read).transpose().map_err(ExportError::invalid)?
        .unwrap_or_else(|| Element::new(native_xml::MAIN_NAMESPACE, "table"));
    normalize_namespaces(&mut root);
    let original_order = source::columns(&root).unwrap_or_default().into_iter()
        .map(|column| source::positive_id(column, "id")).collect::<Result<Vec<_>, _>>().map_err(ExportError::invalid)?;
    let original_columns = source::columns(&root).unwrap_or_default().into_iter()
        .map(|column| source::positive_id(column, "id").map(|id| (id, column.clone()))).collect::<Result<BTreeMap<_, _>, _>>().map_err(ExportError::invalid)?;
    for (name, value) in [("id", id.to_string()), ("name", table.name.clone()), ("displayName", table.name.clone()),
        ("ref", reference), ("headerRowCount", u8::from(table.show_header_row).to_string()),
        ("totalsRowCount", u8::from(table.show_totals_row).to_string()), ("totalsRowShown", u8::from(table.show_totals_row).to_string())] {
        root.set_attr(name, value);
    }
    let mut columns = root.child("tableColumns").cloned().unwrap_or_else(|| Element::new(native_xml::MAIN_NAMESPACE, "tableColumns"));
    columns.set_attr("count", table.columns.len().to_string());
    let mut replacement = Vec::with_capacity(table.columns.len());
    for column in &table.columns {
        let id = column.office_column_id.ok_or_else(|| ExportError::invalid("TABLE_SOURCE_INVALID: column has no prepared native ID"))?;
        let mut node = original_columns.get(&id).cloned().unwrap_or_else(|| Element::new(native_xml::MAIN_NAMESPACE, "tableColumn"));
        node.set_attr("id", id.to_string()); node.set_attr("name", &column.name);
        patch_formula(&mut node, "calculatedColumnFormula", &column.formula, column.formula_is_array);
        patch_formula(&mut node, "totalsRowFormula", &column.totals_row_formula, None);
        patch_attribute(&mut node, "totalsRowLabel", &column.totals_row_label);
        patch_attribute(&mut node, "totalsRowFunction", &column.totals_row_function);
        node.order_children(&["calculatedColumnFormula", "totalsRowFormula", "xmlColumnPr", "extLst"]);
        replacement.push(node);
    }
    columns.replace_children(|child| source::spreadsheet(child, "tableColumn"), replacement);
    root.replace_children(|child| source::spreadsheet(child, "tableColumns"), vec![columns]);
    let mut style = root.child("tableStyleInfo").cloned().unwrap_or_else(|| Element::new(native_xml::MAIN_NAMESPACE, "tableStyleInfo"));
    style.set_attr("name", &table.style_name);
    for (name, flag, default) in [("showFirstColumn", table.show_first_column, false), ("showLastColumn", table.show_last_column, false),
        ("showRowStripes", table.show_row_stripes, true), ("showColumnStripes", table.show_column_stripes, false)] {
        if flag.is_some() || !style.attrs.contains_key(name) { style.set_attr(name, u8::from(flag.unwrap_or(default)).to_string()); }
    }
    root.replace_children(|child| source::spreadsheet(child, "tableStyleInfo"), vec![style]);
    patch_filter(&mut root, table, &original_order, sheet, warnings)?;
    root.order_children(&["autoFilter", "sortState", "tableColumns", "tableStyleInfo", "extLst"]);
    root.validate_fragment().map_err(|error| ExportError::invalid(format!("TABLE_SOURCE_LIMIT: {error}")))?;
    Ok(format!("{}{}", xml::DECL, root.render()))
}

fn patch_attribute(node: &mut Element, name: &str, value: &Option<Option<String>>) {
    if let Some(value) = value { match value { Some(value) => node.set_attr(name, value), None => node.remove_attr(name) } }
}

fn patch_formula(node: &mut Element, name: &str, value: &Option<Option<String>>, array: Option<bool>) {
    if value.is_none() && array.is_none() { return; }
    if matches!(value, Some(None)) || value.as_ref().and_then(Option::as_ref).is_some_and(String::is_empty) {
        node.replace_children(|child| source::spreadsheet(child, name), vec![]); return;
    }
    let mut formula = node.child(name).cloned().unwrap_or_else(|| Element::new(native_xml::MAIN_NAMESPACE, name));
    if let Some(Some(value)) = value { formula.set_text(value.trim_start_matches('=')); }
    if let Some(array) = array { formula.set_attr("array", u8::from(array).to_string()); }
    if !formula.text.is_empty() { node.replace_children(|child| source::spreadsheet(child, name), vec![formula]); }
}

fn patch_filter(root: &mut Element, table: &TableInput, original_order: &[u32], sheet: &SheetInput, warnings: &mut Vec<ExportWarning>) -> Result<(), ExportError> {
    let (r1, c1, r2, c2) = bounds(&table.range).ok_or_else(|| ExportError::invalid("invalid table filter range"))?;
    if !table.show_header_row || matches!(table.auto_filter, Some(None)) {
        root.replace_children(|child| source::spreadsheet(child, "autoFilter"), vec![]); return Ok(());
    }
    let range = RangeInput { start_row: Some(r1), start_col: Some(c1), end_row: Some(r2 - u32::from(table.show_totals_row)), end_col: Some(c2), ..RangeInput::default() };
    let mut filter = if let Some(Some(filter)) = &table.auto_filter {
        let mut filter = filter.clone(); filter.range = range.clone();
        let xml = super::auto_filter::write_filter_xml(sheet, &filter, warnings)?;
        native_xml::parse(&xml.replacen("<autoFilter", &format!("<autoFilter xmlns=\"{}\"", native_xml::MAIN_NAMESPACE), 1)).map_err(ExportError::invalid)?
    } else { root.child("autoFilter").cloned().unwrap_or_else(|| Element::new(native_xml::MAIN_NAMESPACE, "autoFilter")) };
    filter.set_attr("ref", range_ref(&range).ok_or_else(|| ExportError::invalid("invalid table data range"))?);
    // Native colId is the column's position, not its stable OOXML column ID.
    // Only retained native criteria need this remap; typed criteria use current positions.
    if table.auto_filter.is_none() && !original_order.is_empty() {
        let order = root.child("tableColumns").map(|columns| columns.children.iter().filter_map(|column| source::positive_id(column, "id").ok()).collect::<Vec<_>>()).unwrap_or_default();
        let mut children = Vec::new();
        for child in &filter.children {
            if !source::spreadsheet(child, "filterColumn") { children.push(child.clone()); continue; }
            let old_index = child.attrs.get("colId").and_then(|value| value.parse::<usize>().ok());
            let native_id = old_index.and_then(|index| original_order.get(index));
            if let Some(index) = native_id.and_then(|id| order.iter().position(|current| current == id)) {
                let mut child = child.clone(); child.set_attr("colId", index.to_string()); children.push(child);
            }
        }
        filter.replace_children(|_| true, children);
    }
    for (index, column) in table.columns.iter().enumerate() {
        let Some(show) = column.show_filter_button else { continue; };
        let position = filter.children.iter().position(|child| source::spreadsheet(child, "filterColumn") && child.attrs.get("colId") == Some(&index.to_string()));
        if let Some(position) = position { filter.children[position].set_attr("hiddenButton", u8::from(!show).to_string()); }
        else if !show { let mut column = Element::new(native_xml::MAIN_NAMESPACE, "filterColumn"); column.set_attr("colId", index.to_string()); column.set_attr("hiddenButton", "1"); filter.push(column); }
    }
    root.replace_children(|child| source::spreadsheet(child, "autoFilter"), vec![filter]);
    Ok(())
}

fn normalize_namespaces(node: &mut Element) {
    if node.namespace == native_xml::STRICT_MAIN_NAMESPACE { node.namespace = native_xml::MAIN_NAMESPACE.into(); }
    for value in node.attrs.values_mut().chain(node.attribute_namespaces.values_mut()) {
        if value == native_xml::STRICT_MAIN_NAMESPACE { *value = native_xml::MAIN_NAMESPACE.into(); }
    }
    for child in &mut node.children { normalize_namespaces(child); }
}

fn valid_text(value: &str) -> bool { value.chars().all(|value| value >= '\u{20}' && !matches!(value, '\u{fffe}' | '\u{ffff}') || matches!(value, '\t' | '\r' | '\n')) }

fn valid_name(value: &str) -> bool {
    if value.is_empty() || value.encode_utf16().count() > 255 || !value.chars().next().is_some_and(|c| c.is_alphabetic() || c == '_')
        || value.chars().any(|c| c.is_whitespace() || c.is_control() || c.is_ascii() && !c.is_ascii_alphanumeric() && !matches!(c, '_' | '.')) { return false; }
    let lower = value.to_lowercase();
    if matches!(lower.as_str(), "r" | "c") { return false; }
    let letters = value.bytes().take_while(u8::is_ascii_alphabetic).count();
    if (1..=3).contains(&letters) && !value[letters..].is_empty() && value[letters..].bytes().all(|c| c.is_ascii_digit()) { return false; }
    if let Some(rest) = lower.strip_prefix('r') {
        if let Some((row, column)) = rest.split_once('c') { if !row.is_empty() && !column.is_empty() && row.bytes().chain(column.bytes()).all(|c| c.is_ascii_digit()) { return false; } }
    }
    valid_text(value)
}

fn default_true() -> bool { true }
fn default_style() -> String { "TableStyleMedium2".to_owned() }
fn bounds(range: &RangeInput) -> Option<(u32, u32, u32, u32)> {
    Some((range.start_row.or(range.first_row)?, range.start_col.or(range.first_col)?, range.end_row.or(range.last_row)?, range.end_col.or(range.last_col)?))
}
