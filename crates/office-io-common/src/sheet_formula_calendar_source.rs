//! Authored formulas and exact serials beside standard XLSX calculation views.
//! This part supplies provenance only; Office calculates ordinary cell formulas.

use serde::{Deserialize, Serialize};
use serde_json::Value;
use std::collections::{BTreeMap, BTreeSet};

pub const NAMESPACE: &str = "urn:office-studio:formula-calendar:1";
pub const PART_NAME: &str = "customXml/officeFormulaCalendars.xml";
pub const RELATIONSHIP: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";
pub const MAX_BYTES: usize = 16 * 1024 * 1024;
pub const MAX_CELLS: usize = 250_000;

#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct FormulaCalendarSource {
    pub version: u32,
    pub date1904: bool,
    pub sheets: Vec<SheetSource>,
    pub helpers: Vec<HelperSource>,
    #[serde(default)]
    pub tables: Vec<TableSource>,
}
#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct SheetSource {
    pub name: String,
    pub sheet_id: String,
    pub office_sheet_id: u32,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub import: Option<Value>,
    pub cells: Vec<CellSource>,
}
#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct CellSource {
    pub row: u32,
    pub col: u32,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub formula: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub projected_formula: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub formula_source: Option<Value>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub date_source: Option<Value>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub calculation_sheet: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub calculation_formula: Option<String>,
}
#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct HelperSource {
    pub name: String,
    pub sheet_id: String,
    pub office_sheet_id: u32,
    pub date1904: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub source_sheet_id: Option<String>,
    pub fingerprint: String,
    /// Computed by the native reader after freshness and foreign-reference
    /// checks. The writer never trusts this bit from a previous import.
    #[serde(default, skip_serializing_if = "is_false")]
    pub reclaim: bool,
}
fn is_false(value: &bool) -> bool { !*value }

/// Content-dependent guard for a calculation worksheet. Formula caches are
/// excluded so recalculation after editing author cells remains fresh. Source
/// identities replace physical sheet/table names, allowing an Office rename.
/// FNV is an accidental-staleness guard, not a security signature.
#[derive(Default)]
pub struct CalculationFingerprint { entries: BTreeMap<String, u64> }
impl CalculationFingerprint {
    fn insert(&mut self, key: String, value: &str) -> Result<(), String> {
        if self.entries.len() >= MAX_CELLS + 32_768 { return Err("FORMULA_CALENDAR_SOURCE_LIMIT: calculation fingerprint exceeds its entry budget".into()); }
        if self.entries.insert(key, hash(value.as_bytes())).is_some() { return Err("FORMULA_CALENDAR_SOURCE_INVALID: duplicate calculation resource".into()); }
        Ok(())
    }
    pub fn cell(&mut self, row: u32, col: u32, formula: Option<&str>, value: &Value,
        sheets: &BTreeMap<String, String>, tables: &BTreeMap<String, String>) -> Result<(), String> {
        if row > 1_048_575 || col > 16_383 { return Err("FORMULA_CALENDAR_SOURCE_INVALID: invalid calculation coordinate".into()); }
        let text = if let Some(formula) = formula {
            format!("f:{}", crate::formula_references::canonical_formula(formula, sheets, tables))
        } else { format!("v:{}", serde_json::to_string(value).map_err(|error| error.to_string())?) };
        self.insert(format!("c:{row}:{col}"), &text)
    }
    pub fn merged_range(&mut self, bounds: (u32, u32, u32, u32)) -> Result<(), String> {
        self.insert(format!("m:{},{},{},{}", bounds.0, bounds.1, bounds.2, bounds.3), "merge")
    }
    pub fn defined_name(&mut self, name: &str, formula: &str, sheets: &BTreeMap<String, String>, tables: &BTreeMap<String, String>) -> Result<(), String> {
        self.insert(format!("n:{}", name.to_lowercase()), &crate::formula_references::canonical_formula(formula, sheets, tables))
    }
    pub fn table(&mut self, id: &str, bounds: (u32, u32, u32, u32), columns: &[String], header: bool, totals: bool) -> Result<(), String> {
        self.insert(format!("t:{id}"), &serde_json::to_string(&(bounds, columns, header, totals)).map_err(|error| error.to_string())?)
    }
    pub fn finish(&self) -> String {
        let mut value = 0xcbf29ce484222325u64;
        for (key, contribution) in &self.entries {
            for byte in key.bytes().chain([0]).chain(contribution.to_le_bytes()) { value ^= u64::from(byte); value = value.wrapping_mul(0x100000001b3); }
        }
        format!("{value:016x}")
    }
}
fn hash(bytes: &[u8]) -> u64 {
    bytes.iter().fold(0xcbf29ce484222325, |value, byte| (value ^ u64::from(*byte)).wrapping_mul(0x100000001b3))
}
#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct TableSource {
    pub name: String,
    pub table_id: String,
    pub office_table_id: u32,
    pub sheet_id: String,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub columns: Vec<TableColumnSource>,
}
#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct TableColumnSource {
    pub column_id: String,
    pub office_column_id: u32,
}

impl FormulaCalendarSource {
    pub fn validate(&self) -> Result<(), String> {
        if self.version != 1 || self.sheets.len() > 16_384 || self.helpers.len() > 2049 { return Err("FORMULA_CALENDAR_SOURCE_INVALID: version or sheet count is invalid".into()); }
        let mut names = BTreeSet::new(); let mut ids = BTreeSet::new(); let mut native_ids = BTreeSet::new(); let mut cells = 0usize;
        for sheet in &self.sheets {
            if sheet.name.is_empty() || sheet.name.chars().count() > 31 || sheet.sheet_id.is_empty() || sheet.sheet_id.len() > 512
                || sheet.office_sheet_id == 0 || !native_ids.insert(sheet.office_sheet_id)
                || !names.insert(sheet.name.to_lowercase()) || !ids.insert(sheet.sheet_id.as_str()) {
                return Err("FORMULA_CALENDAR_SOURCE_INVALID: duplicate or invalid worksheet identity".into());
            }
            let mut coordinates = BTreeSet::new();
            for cell in &sheet.cells {
                cells += 1;
                if cells > MAX_CELLS { return Err("FORMULA_CALENDAR_SOURCE_LIMIT: too many authored cells".into()); }
                if cell.row > 1_048_575 || cell.col > 16_383 || !coordinates.insert((cell.row, cell.col))
                    || cell.formula.as_ref().is_some_and(|value| value.len() > 32_768)
                    || cell.projected_formula.as_ref().is_some_and(|value| value.len() > 32_768)
                    || cell.calculation_formula.as_ref().is_some_and(|value| value.len() > 32_768) {
                    return Err("FORMULA_CALENDAR_SOURCE_INVALID: invalid authored cell".into());
                }
                if let Some(marker) = &cell.formula_source {
                    if !marker.is_object() || marker.get("version").and_then(Value::as_u64) != Some(1)
                        || marker.get("formula").and_then(Value::as_str) != cell.formula.as_deref()
                        || cell.formula.is_none() || cell.projected_formula.is_none() || cell.calculation_formula.is_none() || cell.calculation_sheet.is_none()
                        || marker.get("date1904").and_then(Value::as_bool).is_none() || marker.get("dateResult").and_then(Value::as_bool).is_none()
                        || marker.get("aliasSheetId").is_some_and(|value| value.as_str().is_none_or(|value| value.is_empty() || value.len() > 512)) {
                        return Err("FORMULA_CALENDAR_SOURCE_INVALID: invalid formula provenance".into());
                    }
                }
                if let Some(marker) = &cell.date_source {
                    if !marker.is_object() || finite_value(marker.get("value")).is_none()
                        || marker.get("date1904").and_then(Value::as_bool).is_none()
                        || marker.get("targetDate1904").is_some_and(|value| value.as_bool() != Some(self.date1904))
                        || marker.get("targetValue").is_some_and(|value| finite_value(Some(value)).is_none())
                        || marker.get("targetFormula").is_some_and(|value| !value.is_null() && value.as_str() != cell.formula.as_deref()) {
                        return Err("FORMULA_CALENDAR_SOURCE_INVALID: invalid exact-date provenance".into());
                    }
                }
            }
            if let Some(import) = &sheet.import { validate_import(import)?; }
        }
        for helper in &self.helpers {
            if helper.name.is_empty() || helper.name.chars().count() > 31 || helper.sheet_id.is_empty() || helper.sheet_id.len() > 512
                || helper.office_sheet_id == 0 || !native_ids.insert(helper.office_sheet_id)
                || helper.fingerprint.len() != 16 || !helper.fingerprint.bytes().all(|byte| byte.is_ascii_hexdigit())
                || !names.insert(helper.name.to_lowercase()) || !ids.insert(helper.sheet_id.as_str()) {
                return Err("FORMULA_CALENDAR_SOURCE_INVALID: duplicate or invalid helper identity".into());
            }
        }
        for helper in &self.helpers {
            if helper.source_sheet_id.as_ref().is_some_and(|id| !self.sheets.iter().any(|sheet| &sheet.sheet_id == id)) {
                return Err("FORMULA_CALENDAR_SOURCE_INVALID: helper owner is missing".into());
            }
        }
        let helpers_by_name = self.helpers.iter().map(|helper| (helper.name.as_str(), helper)).collect::<BTreeMap<_, _>>();
        for sheet in &self.sheets { for cell in &sheet.cells {
            if let Some(name) = &cell.calculation_sheet {
                let helper = helpers_by_name.get(name.as_str())
                    .ok_or_else(|| "FORMULA_CALENDAR_SOURCE_INVALID: calculation worksheet is missing".to_owned())?;
                if helper.source_sheet_id.as_ref() != Some(&sheet.sheet_id)
                    || cell.formula_source.as_ref().and_then(|value| value.get("date1904")).and_then(Value::as_bool) != Some(helper.date1904) {
                    return Err("FORMULA_CALENDAR_SOURCE_INVALID: calculation worksheet has a different owner or calendar".into());
                }
            }
        } }
        if self.tables.len() > 24_576 { return Err("FORMULA_CALENDAR_SOURCE_LIMIT: too many table identities".into()); }
        let mut table_names = BTreeSet::new(); let mut table_ids = BTreeSet::new(); let mut native_table_ids = BTreeSet::new();
        for table in &self.tables {
            if table.name.is_empty() || table.name.encode_utf16().count() > 255 || table.table_id.is_empty() || table.table_id.len() > 512
                || table.office_table_id == 0 || !native_table_ids.insert(table.office_table_id)
                || !table_names.insert(table.name.to_lowercase()) || !table_ids.insert(table.table_id.as_str()) || !ids.contains(table.sheet_id.as_str()) {
                return Err("FORMULA_CALENDAR_SOURCE_INVALID: invalid table identity".into());
            }
            if table.columns.len() > 16_384 { return Err("FORMULA_CALENDAR_SOURCE_LIMIT: too many table column identities".into()); }
            let mut columns = BTreeSet::new(); let mut native_columns = BTreeSet::new();
            for column in &table.columns {
                if column.column_id.is_empty() || column.column_id.len() > 512 || column.office_column_id == 0
                    || !columns.insert(column.column_id.as_str()) || !native_columns.insert(column.office_column_id) {
                    return Err("FORMULA_CALENDAR_SOURCE_INVALID: duplicate or invalid table column identity".into());
                }
            }
        }
        Ok(())
    }
    pub fn to_xml(&self) -> Result<String, String> {
        self.validate()?;
        let text = serde_json::to_string(self).map_err(|error| error.to_string())?;
        if text.len() > MAX_BYTES { return Err("FORMULA_CALENDAR_SOURCE_LIMIT: provenance exceeds 16 MiB".into()); }
        let escaped = text.replace('&', "&amp;").replace('<', "&lt;").replace('>', "&gt;");
        let xml = format!("<?xml version=\"1.0\" encoding=\"UTF-8\"?><formulaCalendars xmlns=\"{NAMESPACE}\" version=\"1\">{escaped}</formulaCalendars>");
        if xml.len() > MAX_BYTES { return Err("FORMULA_CALENDAR_SOURCE_LIMIT: XML exceeds 16 MiB".into()); }
        Ok(xml)
    }
}

pub fn finite_value(value: Option<&Value>) -> Option<f64> {
    value?.as_f64().or_else(|| value?.as_str().filter(|value| !value.trim().is_empty()).and_then(|value| value.parse().ok()))
        .filter(|value| value.is_finite())
}
fn validate_import(value: &Value) -> Result<(), String> {
    if !value.is_object() || value.get("version").and_then(Value::as_u64) != Some(1) {
        return Err("FORMULA_CALENDAR_SOURCE_INVALID: invalid import provenance".into());
    }
    if let Some(aliases) = value.get("formulaAliases") {
        if !aliases.is_object() { return Err("FORMULA_CALENDAR_SOURCE_INVALID: invalid formula alias catalog".into()); }
        let mut count = 0usize;
        for key in ["sheets", "names", "tables", "nameSheets"] {
            let Some(values) = aliases.get(key) else { continue; };
            let values = values.as_object().ok_or_else(|| "FORMULA_CALENDAR_SOURCE_INVALID: invalid formula alias map".to_owned())?;
            for (name, target) in values {
                count += 1;
                if count > 32_768 || name.is_empty() || name.len() > 2048 || target.as_str().is_none_or(|value| value.len() > 2048) {
                    return Err("FORMULA_CALENDAR_SOURCE_INVALID: invalid formula alias entry".into());
                }
            }
        }
    }
    Ok(())
}

pub fn from_xml(xml: &str) -> Result<FormulaCalendarSource, String> {
    if xml.len() > MAX_BYTES { return Err("FORMULA_CALENDAR_SOURCE_LIMIT: XML exceeds 16 MiB".into()); }
    let root = crate::chart_xml::parse(xml)?;
    if root.name != "formulaCalendars" || root.namespace.as_deref() != Some(NAMESPACE)
        || root.attrs.get("version").map(String::as_str) != Some("1") || !root.children.is_empty() {
        return Err("FORMULA_CALENDAR_SOURCE_INVALID: XML root or namespace is invalid".into());
    }
    let source: FormulaCalendarSource = serde_json::from_str(&root.text).map_err(|error| format!("FORMULA_CALENDAR_SOURCE_INVALID: {error}"))?;
    source.validate()?; Ok(source)
}
