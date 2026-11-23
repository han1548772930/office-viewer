use serde::Deserialize;

use crate::{ExportError, xml};
use super::{RangeInput, esc, range_ref};

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
    pub(super) auto_filter: Option<super::auto_filter::AutoFilterInput>,
}

#[derive(Debug, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct TableColumnInput {
    pub(crate) id: String,
    pub(crate) name: String,
}

pub(crate) fn validate(tables: &[TableInput], table: &TableInput) -> Result<(), ExportError> {
    let Some((start_row, start_col, end_row, end_col)) = bounds(&table.range) else {
        return Err(ExportError::invalid(format!("table {} has an invalid range", table.id)));
    };
    if start_row >= end_row || table.columns.len() != (end_col - start_col + 1) as usize {
        return Err(ExportError::invalid(format!("table {} has invalid header or column definitions", table.id)));
    }
    if table.name.is_empty() || !table.name.chars().next().is_some_and(|c| c.is_ascii_alphabetic() || c == '_')
        || !table.name.chars().all(|c| c.is_ascii_alphanumeric() || c == '_' || c == '.') {
        return Err(ExportError::invalid(format!("table {} has an invalid name", table.id)));
    }
    let mut names = std::collections::BTreeSet::new();
    if table.columns.iter().any(|column| column.name.trim().is_empty() || !names.insert(column.name.to_lowercase())) {
        return Err(ExportError::invalid(format!("table {} has blank or duplicate column names", table.id)));
    }
    if tables.iter().filter(|other| other.id != table.id && other.sheet_id == table.sheet_id).any(|other| {
        bounds(&other.range).is_some_and(|(r1, c1, r2, c2)| start_row <= r2 && end_row >= r1 && start_col <= c2 && end_col >= c1)
    }) {
        return Err(ExportError::invalid(format!("table {} overlaps another table", table.id)));
    }
    Ok(())
}

pub(crate) fn write(table: &TableInput, id: usize) -> Result<String, ExportError> {
    let reference = range_ref(&table.range).ok_or_else(|| ExportError::invalid("table range is invalid"))?;
    let mut out = format!("{}<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" id=\"{id}\" name=\"{}\" displayName=\"{}\" ref=\"{}\" headerRowCount=\"{}\" totalsRowShown=\"{}\">", xml::DECL, esc(&table.name), esc(&table.name), reference, if table.show_header_row { 1 } else { 0 }, if table.show_totals_row { 1 } else { 0 });
    out.push_str(&format!("<autoFilter ref=\"{}\"/>", reference));
    out.push_str(&format!("<tableColumns count=\"{}\">", table.columns.len()));
    for (index, column) in table.columns.iter().enumerate() {
        out.push_str(&format!("<tableColumn id=\"{}\" name=\"{}\"/>", index + 1, esc(&column.name)));
    }
    out.push_str("</tableColumns>");
    out.push_str(&format!("<tableStyleInfo name=\"{}\" showFirstColumn=\"0\" showLastColumn=\"0\" showRowStripes=\"1\" showColumnStripes=\"0\"/></table>", esc(&table.style_name)));
    Ok(out)
}

fn default_true() -> bool { true }
fn default_style() -> String { "TableStyleMedium2".to_owned() }
fn bounds(range: &RangeInput) -> Option<(u32, u32, u32, u32)> {
    Some((range.start_row.or(range.first_row)?, range.start_col.or(range.first_col)?, range.end_row.or(range.last_row)?, range.end_col.or(range.last_col)?))
}
