//! Native table XML remains independent of the editable table/column identities.
//! Import and export share the same namespace, fragment and workbook budgets.

use std::collections::BTreeSet;
use serde::{Deserialize, Serialize};
use crate::conditional_format_xml::{self as xml, Element};

pub const WORKBOOK_BYTE_LIMIT: usize = 64 * 1024 * 1024;
pub const TABLE_LIMIT: usize = 24_576;

#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct OfficeTableSource {
    pub part_name: String,
    pub xml: String,
}

pub fn spreadsheet(node: &Element, name: &str) -> bool {
    node.name == name && matches!(node.namespace.as_str(), xml::MAIN_NAMESPACE | xml::STRICT_MAIN_NAMESPACE)
}

pub fn positive_id(node: &Element, name: &str) -> Result<u32, String> {
    node.attrs.get(name).and_then(|value| value.parse::<u32>().ok()).filter(|value| *value > 0)
        .ok_or_else(|| format!("TABLE_SOURCE_INVALID: invalid native {name}"))
}

pub fn flag(node: &Element, name: &str) -> Result<Option<bool>, String> {
    match node.attrs.get(name).map(String::as_str) {
        None => Ok(None), Some("1" | "true") => Ok(Some(true)), Some("0" | "false") => Ok(Some(false)),
        _ => Err(format!("TABLE_SOURCE_INVALID: invalid {name} flag")),
    }
}

/// Native table, filter and sort references share the same bounded A1 geometry.
/// These are worksheet-local references, never sheet-qualified or multi-area.
pub fn range(reference: &str) -> Option<[u32; 4]> {
    fn cell(input: &str) -> Option<(u32, u32)> {
        let input = input.strip_prefix('$').unwrap_or(input);
        let split = input.bytes().take_while(u8::is_ascii_alphabetic).count();
        if !(1..=3).contains(&split) { return None; }
        let (letters, digits) = input.split_at(split);
        let digits = digits.strip_prefix('$').unwrap_or(digits);
        if digits.is_empty() || !digits.bytes().all(|byte| byte.is_ascii_digit()) { return None; }
        let mut column = 0u32;
        for byte in letters.bytes() { column = column.checked_mul(26)?.checked_add(u32::from(byte.to_ascii_uppercase() - b'A') + 1)?; }
        let row = digits.parse::<u32>().ok()?.checked_sub(1)?;
        let column = column.checked_sub(1)?;
        (row <= 1_048_575 && column <= 16_383).then_some((row, column))
    }
    let (first, last) = reference.split_once(':').unwrap_or((reference, reference));
    let (r1, c1) = cell(first)?; let (r2, c2) = cell(last)?;
    (r1 <= r2 && c1 <= c2).then_some([r1, c1, r2, c2])
}

pub fn columns(root: &Element) -> Result<Vec<&Element>, String> {
    let roots = root.children.iter().filter(|child| spreadsheet(child, "tableColumns")).collect::<Vec<_>>();
    if roots.len() != 1 { return Err("TABLE_SOURCE_INVALID: missing or duplicate table columns".into()); }
    let columns = roots[0].children.iter().filter(|child| spreadsheet(child, "tableColumn")).collect::<Vec<_>>();
    if columns.is_empty() || columns.len() > 16_384 { return Err("TABLE_SOURCE_INVALID: invalid table column count".into()); }
    let mut ids = BTreeSet::new(); let mut names = BTreeSet::new();
    for column in &columns {
        let id = positive_id(column, "id")?;
        let name = column.attrs.get("name").filter(|name| !name.is_empty())
            .ok_or_else(|| "TABLE_SOURCE_INVALID: missing column name".to_owned())?;
        if !ids.insert(id) || !names.insert(name.to_lowercase()) { return Err("TABLE_SOURCE_INVALID: duplicate table column identity".into()); }
        for name in ["calculatedColumnFormula", "totalsRowFormula"] {
            if column.children.iter().filter(|child| spreadsheet(child, name)).count() > 1 {
                return Err("TABLE_SOURCE_INVALID: duplicate table column formula".into());
            }
        }
    }
    Ok(columns)
}

pub fn parse_xml(source: &str) -> Result<Element, String> {
    let root = xml::parse(source).map_err(|error| format!("TABLE_SOURCE_INVALID: {error}"))?;
    if !spreadsheet(&root, "table") { return Err("TABLE_SOURCE_INVALID: invalid table root or namespace".into()); }
    positive_id(&root, "id")?;
    columns(&root)?;
    for name in ["autoFilter", "sortState", "tableStyleInfo"] {
        if root.children.iter().filter(|child| spreadsheet(child, name)).count() > 1 {
            return Err(format!("TABLE_SOURCE_INVALID: duplicate {name}"));
        }
    }
    Ok(root)
}

pub fn read(source: &OfficeTableSource) -> Result<Element, String> {
    if source.part_name.is_empty() || source.part_name.len() > 4096 || source.part_name.contains('\\')
        || source.part_name.split('/').any(|segment| segment.is_empty() || matches!(segment, "." | "..")) {
        return Err("TABLE_SOURCE_INVALID: invalid table source part name".into());
    }
    parse_xml(&source.xml)
}

#[derive(Default)]
pub struct TableSourceBudget { bytes: usize, tables: usize }
impl TableSourceBudget {
    pub fn admit(&mut self, source: Option<&OfficeTableSource>) -> Result<(), String> {
        self.tables = self.tables.checked_add(1).filter(|tables| *tables <= TABLE_LIMIT)
            .ok_or_else(|| "TABLE_SOURCE_LIMIT: too many table identities".to_owned())?;
        if let Some(source) = source {
            if source.xml.len() > xml::XML_FRAGMENT_BYTE_LIMIT { return Err("TABLE_SOURCE_LIMIT: native table part exceeds 16 MiB".into()); }
            self.bytes = self.bytes.checked_add(source.xml.len()).filter(|bytes| *bytes <= WORKBOOK_BYTE_LIMIT)
                .ok_or_else(|| "TABLE_SOURCE_LIMIT: native table parts exceed 64 MiB".to_owned())?;
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    fn source() -> String { format!("<s:table xmlns:s=\"{}\" id=\"31\"><s:tableColumns count=\"1\"><s:tableColumn id=\"9\" name=\"税额\"/></s:tableColumns></s:table>", xml::STRICT_MAIN_NAMESPACE) }
    #[test]
    fn native_table_and_column_ids_use_their_resolved_namespaces() {
        let root = parse_xml(&source()).unwrap();
        assert_eq!(positive_id(&root, "id").unwrap(), 31);
        assert_eq!(positive_id(columns(&root).unwrap()[0], "id").unwrap(), 9);
        assert!(parse_xml(&source().replace(xml::STRICT_MAIN_NAMESPACE, "urn:unrelated")).is_err());
    }
    #[test]
    fn duplicate_columns_and_unrelated_table_column_roots_are_rejected() {
        assert!(parse_xml(&source().replace("</s:tableColumns>", "<s:tableColumn id=\"9\" name=\"Other\"/></s:tableColumns>")).is_err());
        assert!(parse_xml(&source().replace("s:tableColumns", "x:tableColumns").replace("id=\"31\"", "xmlns:x=\"urn:other\" id=\"31\"")).is_err());
    }
    #[test]
    fn native_source_admission_precedes_xml_parsing() {
        let mut budget = TableSourceBudget::default();
        let source = OfficeTableSource { part_name: "xl/tables/table1.xml".into(), xml: " ".repeat(xml::XML_FRAGMENT_BYTE_LIMIT + 1) };
        assert!(budget.admit(Some(&source)).unwrap_err().starts_with("TABLE_SOURCE_LIMIT:"));
    }

    #[test]
    fn local_references_are_bounded_and_preserve_absolute_a1_geometry() {
        assert_eq!(range("$a$1:$XFD$1048576"), Some([0, 0, 1_048_575, 16_383]));
        assert_eq!(range("B2"), Some([1, 1, 1, 1]));
        for value in ["A0", "B2:A1", "A1:XFE2", "A1:A1048577", "A1:B2:C3", "Sheet!A1", "A1 B2", "A$$1"] {
            assert_eq!(range(value), None, "{value}");
        }
    }
}
