//! One compact anchor grid drives native rows, merges and nested DOCX tables.
use std::collections::{BTreeSet, HashSet};

use office_io_common::document_table_source::{MAX_TABLE_COUNT, MAX_TABLE_SOURCE_BYTES, TABLE_SOURCE_NAMESPACE,
    TABLE_SOURCE_PART, TABLE_SOURCE_TARGET, TableSourceEntry, WORD_NAMESPACE, native_identity, validate_entries};
use serde_json::{Map, Value, json};

use super::{WriteContext, cell_blocks, get, xml_attribute, xml_text};
use crate::{ExportError, xml};

#[path = "docx_table_styles.rs"]
mod styles;

const MAX_ROWS: usize = 10_000;
const MAX_COLUMNS: usize = 512;
const MAX_GRID_SLOTS: usize = 1_000_000;
const MAX_NESTING: usize = 32;

#[derive(Default)]
pub(super) struct TableWriteState {
    depth: usize,
    count: usize,
    grid_slots: usize,
    source_bytes: usize,
    ids: BTreeSet<String>,
    entries: Vec<TableSourceEntry>,
}

#[derive(Clone, Copy)]
struct Anchor<'a> { cell: &'a Value, row: usize, column: usize, row_span: usize, col_span: usize, cell_index: usize }
struct Grid<'a> { rows: Vec<Vec<Anchor<'a>>>, column_count: usize }

pub(super) struct ContentGroup<'a> { pub blocks: &'a [Value] }

fn integer(value: Option<&Value>, fallback: Option<usize>, label: &str) -> Result<usize, ExportError> {
    match value.filter(|value| !value.is_null()) {
        None => fallback.ok_or_else(|| ExportError::invalid(format!("DOCX table {label} is missing"))),
        Some(value) => value.as_u64().and_then(|value| usize::try_from(value).ok()).ok_or_else(|| ExportError::invalid(format!("DOCX table {label} must be a nonnegative integer"))),
    }
}

fn identity(value: &Value, fallback: String, ids: &mut HashSet<String>, label: &str) -> Result<String, ExportError> {
    let id = match value.get("id") { None => fallback, Some(Value::String(id)) if !id.trim().is_empty() && id.len() <= 4096 => id.clone(),
        Some(_) => return Err(ExportError::invalid(format!("DOCX table {label} ID is invalid"))) };
    if !ids.insert(id.clone()) { return Err(ExportError::invalid(format!("DOCX table {label} IDs are duplicated"))); }
    Ok(id)
}

fn grid(table: &Value) -> Result<Grid<'_>, ExportError> {
    let rows = get(table, "rows").as_array().filter(|rows| !rows.is_empty()).ok_or_else(|| ExportError::invalid("DOCX table requires actual rows"))?;
    if rows.len() > MAX_ROWS { return Err(ExportError::invalid("DOCX_TABLE_RESOURCE_LIMIT: table exceeds 10,000 actual rows")); }
    let mut output = vec![Vec::<Anchor<'_>>::new(); rows.len()];
    let mut column_count = 0usize;
    let mut row_ids = HashSet::new();
    let mut cell_ids = HashSet::new();
    for (row_index, row) in rows.iter().enumerate() {
        identity(row, format!("row-{row_index}"), &mut row_ids, "row")?;
        let cells = get(row, "cells").as_array().ok_or_else(|| ExportError::invalid("DOCX table row requires a cells array, including covered empty rows"))?;
        let mut cursor = 0usize;
        for (cell_index, cell) in cells.iter().enumerate() {
            identity(cell, format!("row-{row_index}-cell-{cell_index}"), &mut cell_ids, "cell")?;
            if cell.get("blocks").is_some_and(|blocks| !blocks.is_array()) { return Err(ExportError::invalid("DOCX table cell blocks must be an array")); }
            let col_span = integer(cell.get("colSpan"), Some(1), "colSpan")?;
            let row_span = integer(cell.get("rowSpan"), Some(1), "rowSpan")?;
            if col_span == 0 || row_span == 0 || row_span > rows.len() - row_index {
                return Err(ExportError::invalid("DOCX table spans must be positive and stay within its actual rows"));
            }
            output[row_index].sort_by_key(|cell| cell.column);
            let column = if let Some(column) = cell.get("columnIndex").filter(|value| !value.is_null()) {
                integer(Some(column), None, "columnIndex")?
            } else {
                let mut column = cursor;
                for owner in &output[row_index] {
                    if owner.column + owner.col_span <= column { continue; }
                    if column.checked_add(col_span).is_some_and(|end| end <= owner.column) { break; }
                    column = owner.column + owner.col_span;
                }
                column
            };
            let end = column.checked_add(col_span).filter(|end| *end <= MAX_COLUMNS)
                .ok_or_else(|| ExportError::invalid("DOCX_TABLE_RESOURCE_LIMIT: table exceeds 512 logical columns"))?;
            column_count = column_count.max(end);
            if rows.len().checked_mul(column_count).filter(|slots| *slots <= MAX_GRID_SLOTS).is_none() {
                return Err(ExportError::invalid("DOCX_TABLE_RESOURCE_LIMIT: table exceeds its logical grid budget"));
            }
            let anchor = Anchor { cell, row: row_index, column, row_span, col_span, cell_index };
            for row in &mut output[row_index..row_index + row_span] {
                if row.iter().any(|owner| owner.column < end && column < owner.column + owner.col_span) {
                    return Err(ExportError::invalid("DOCX table contains overlapping anchor spans; the source was not shifted or clamped"));
                }
                row.push(anchor);
            }
            cursor = end;
        }
    }
    if column_count == 0 { return Err(ExportError::invalid("DOCX table requires at least one logical column")); }
    for row in &mut output {
        row.sort_by_key(|cell| cell.column);
        let mut cursor = 0;
        for anchor in row.iter() {
            if anchor.column != cursor { return Err(ExportError::invalid("DOCX table contains an uncovered grid interval; no cells were fabricated")); }
            cursor += anchor.col_span;
        }
        if cursor != column_count { return Err(ExportError::invalid("DOCX table rows must cover the same logical columns")); }
    }
    for (row_index, cells) in output.iter().enumerate() {
        for owner in cells.iter().filter(|cell| cell.row == row_index) { validate_continuation_content(rows, owner)?; }
    }
    Ok(Grid { rows: output, column_count })
}

fn validate_continuation_content(rows: &[Value], owner: &Anchor<'_>) -> Result<(), ExportError> {
    let Some(content) = get(owner.cell, "style").get("nativeContinuationContent").filter(|value| !value.is_null()) else { return Ok(()); };
    let content = content.as_object().ok_or_else(|| ExportError::invalid("DOCX table native continuation content must be a mapping"))?;
    let mut row_ids = HashSet::new();
    for (offset, source) in content {
        if offset.parse::<usize>().ok().filter(|offset| *offset > 0 && *offset < MAX_ROWS).is_none()
            || get(source, "version").as_u64() != Some(1) || !get(source, "blocks").is_array()
            || get(source, "ownerCellId").as_str().is_none() || get(source, "ownerCellId").as_str() != get(owner.cell, "id").as_str()
            || get(source, "colSpan").as_u64() != Some(owner.col_span as u64) {
            return Err(ExportError::invalid("DOCX_TABLE_CONTINUATION_CONTENT_INVALID: continuation content has an invalid version, owner, span or block array"));
        }
        let row_id = get(source, "rowId").as_str().filter(|id| !id.trim().is_empty()).ok_or_else(|| ExportError::invalid("DOCX continuation content requires a stable physical row ID"))?;
        if !row_ids.insert(row_id) { return Err(ExportError::invalid("DOCX continuation content has duplicate physical row IDs")); }
        let row_index = rows.iter().position(|row| get(row, "id").as_str() == Some(row_id));
        if row_index.is_none_or(|row| row <= owner.row || row >= owner.row + owner.row_span) {
            return Err(ExportError::invalid("DOCX_TABLE_CONTINUATION_CONTENT_STALE: continuation content no longer belongs to a covered physical row"));
        }
    }
    Ok(())
}

fn continuation_blocks<'a>(table: &Value, owner: &'a Value, row_index: usize) -> &'a [Value] {
    let row_id = get(table, "rows").as_array().and_then(|rows| rows.get(row_index)).and_then(|row| get(row, "id").as_str());
    get(get(owner, "style"), "nativeContinuationContent").as_object().into_iter().flat_map(|content| content.values())
        .find(|source| row_id.is_some() && get(source, "rowId").as_str() == row_id)
        .and_then(|source| get(source, "blocks").as_array()).map(Vec::as_slice).unwrap_or_default()
}

/// Numbering, fields, notes, controls and writing use the same physical row and
/// logical column order, including independent vertical-continuation contents.
pub(super) fn content_groups(table: &Value) -> Result<Vec<ContentGroup<'_>>, ExportError> {
    let grid = grid(table)?;
    let mut groups = Vec::new();
    for (row_index, cells) in grid.rows.iter().enumerate() {
        for owner in cells {
            let blocks = if owner.row == row_index { cell_blocks(owner.cell) } else { continuation_blocks(table, owner.cell, row_index) };
            if !blocks.is_empty() { groups.push(ContentGroup { blocks }); }
        }
    }
    Ok(groups)
}

pub(super) fn validate_document_tables(blocks: &[Value]) -> Result<(), ExportError> {
    fn visit(blocks: &[Value], depth: usize, count: &mut usize, slots: &mut usize) -> Result<(), ExportError> {
        for block in blocks {
            if get(block, "type").as_str() != Some("table") { continue; }
            *count += 1;
            if depth >= MAX_NESTING || *count > MAX_TABLE_COUNT { return Err(ExportError::invalid("DOCX_TABLE_RESOURCE_LIMIT: table nesting/count budget exceeded")); }
            let table = get(get(block, "metadata"), "table");
            let grid = grid(table)?;
            *slots = slots.checked_add(grid.rows.len().saturating_mul(grid.column_count)).filter(|slots| *slots <= MAX_GRID_SLOTS)
                .ok_or_else(|| ExportError::invalid("DOCX_TABLE_RESOURCE_LIMIT: document tables exceed their combined grid budget"))?;
            for group in content_groups(table)? { visit(group.blocks, depth + 1, count, slots)?; }
        }
        Ok(())
    }
    let mut count = 0;
    let mut slots = 0;
    visit(blocks, 0, &mut count, &mut slots)
}

fn widths(table: &Value, columns: usize, available: i64) -> Result<Vec<i64>, ExportError> {
    if let Some(widths) = table.get("columnWidths").filter(|value| !value.is_null()) {
        let values = widths.as_array().ok_or_else(|| ExportError::invalid("DOCX table columnWidths must be an array"))?;
        if !values.is_empty() {
            if values.len() != columns { return Err(ExportError::invalid("DOCX table column widths do not match its logical grid")); }
            return values.iter().map(|value| {
                let width = value.as_f64().filter(|width| width.is_finite() && *width > 0.0 && *width * 15.0 <= 2_147_483_647.0)
                    .ok_or_else(|| ExportError::invalid("DOCX table column widths require positive finite CSS pixels"))?;
                let width = (width * 15.0).round() as i64;
                if width == 0 { return Err(ExportError::invalid("DOCX table column width is smaller than one native twip")); }
                Ok(width)
            }).collect();
        }
    }
    let total = styles::number(get(table, "style"), "width", true)?.map(|width| (width * 15.0).round() as i64).unwrap_or(available);
    if total < columns as i64 { return Err(ExportError::invalid("DOCX table width is smaller than its native logical columns")); }
    Ok((0..columns).map(|index| total / columns as i64 + i64::from((index as i64) < total % columns as i64)).collect())
}

fn data(table: &Value, table_id: &str, grid: &Grid<'_>) -> Value {
    let mut source = table.as_object().map(|table| table.iter().filter(|(key, _)| key.as_str() != "rows").map(|(key, value)| (key.clone(), value.clone())).collect::<Map<_, _>>()).unwrap_or_default();
    fn remove_native_source(properties: &mut Map<String, Value>) {
        if let Some(style) = properties.get_mut("style").and_then(Value::as_object_mut) {
            for key in ["nativeProperties", "nativeContinuations", "nativeContinuationContent"] { style.remove(key); }
        }
    }
    remove_native_source(&mut source);
    let rows = get(table, "rows").as_array().unwrap_or_else(|| unreachable!("validated table rows"));
    let mut output_rows = Vec::with_capacity(rows.len());
    for (row_index, row) in rows.iter().enumerate() {
        let row_id = get(row, "id").as_str().map(str::to_owned).unwrap_or_else(|| format!("{table_id}-row-{row_index}"));
        let mut output = row.as_object().map(|row| row.iter().filter(|(key, _)| key.as_str() != "cells").map(|(key, value)| (key.clone(), value.clone())).collect::<Map<_, _>>()).unwrap_or_default();
        remove_native_source(&mut output);
        output.insert("id".into(), json!(row_id));
        let cells = grid.rows[row_index].iter().filter(|anchor| anchor.row == row_index).map(|anchor| {
            let mut cell = anchor.cell.as_object().map(|cell| cell.iter().filter(|(key, _)| key.as_str() != "blocks").map(|(key, value)| (key.clone(), value.clone())).collect::<Map<_, _>>()).unwrap_or_default();
            remove_native_source(&mut cell);
            cell.insert("id".into(), json!(get(anchor.cell, "id").as_str().map(str::to_owned).unwrap_or_else(|| format!("{row_id}-cell-{}", anchor.cell_index))));
            cell.insert("columnIndex".into(), json!(anchor.column));
            cell.insert("rowSpan".into(), json!(anchor.row_span)); cell.insert("colSpan".into(), json!(anchor.col_span));
            Value::Object(cell)
        }).collect::<Vec<_>>();
        output.insert("cells".into(), json!(cells)); output_rows.push(Value::Object(output));
    }
    source.insert("rows".into(), json!(output_rows)); Value::Object(source)
}

pub(super) fn write(context: &mut WriteContext<'_>, block: &Value, available: Option<i64>) -> Result<String, ExportError> {
    if context.tables.depth >= MAX_NESTING || context.tables.count >= MAX_TABLE_COUNT {
        return Err(ExportError::invalid("DOCX_TABLE_RESOURCE_LIMIT: table nesting/count budget exceeded"));
    }
    let table = get(get(block, "metadata"), "table");
    let grid = grid(table)?;
    styles::validate_column_types(table, grid.column_count)?;
    let rows = get(table, "rows").as_array().ok_or_else(|| ExportError::invalid("DOCX table rows are missing"))?;
    let header_rows = integer(table.get("headerRowCount"), Some(0), "headerRowCount")?;
    if header_rows > rows.len() { return Err(ExportError::invalid("DOCX table headerRowCount exceeds its actual rows")); }
    let available = available.unwrap_or(i64::from(context.page.width.saturating_sub(context.page.left + context.page.right).max(1)));
    let widths = widths(table, grid.column_count, available)?;
    let width = widths.iter().try_fold(0i64, |sum, value| sum.checked_add(*value)).filter(|width| *width <= 2_147_483_647)
        .ok_or_else(|| ExportError::invalid("DOCX table width exceeds the native range"))?;
    let slots = grid.rows.len().checked_mul(grid.column_count).ok_or_else(|| ExportError::invalid("DOCX table grid size overflow"))?;
    context.tables.grid_slots = context.tables.grid_slots.checked_add(slots).filter(|slots| *slots <= MAX_GRID_SLOTS)
        .ok_or_else(|| ExportError::invalid("DOCX_TABLE_RESOURCE_LIMIT: document tables exceed their combined grid budget"))?;
    context.tables.count += 1;
    let key = format!("table-{:08x}", context.tables.count);
    let id = match block.get("id") {
        None => format!("document-{key}"), Some(Value::String(id)) if !id.trim().is_empty() && id.len() <= 4096 => id.clone(),
        Some(_) => return Err(ExportError::invalid("DOCX table block ID is invalid")),
    };
    if !context.tables.ids.insert(id.clone()) { return Err(ExportError::invalid("DOCX table block IDs are duplicated")); }
    context.tables.depth += 1;
    let result = (|| {
        let properties = styles::table(context, table, width, &widths)?;
        let columns = widths.iter().map(|width| format!("<w:gridCol w:w=\"{width}\"/>")).collect::<String>();
        let mut contents = String::new();
        for (row_index, cells) in grid.rows.iter().enumerate() {
            contents.push_str("<w:tr>");
                contents.push_str(&styles::row(context, table, &rows[row_index], row_index, header_rows)?);
            for anchor in cells {
                let width = widths[anchor.column..anchor.column + anchor.col_span].iter().sum::<i64>();
                let continuation = anchor.row != row_index;
                let properties = styles::cell(context, table, anchor.cell, width, anchor.col_span, anchor.row_span, continuation, row_index)?;
                contents.push_str(&format!("<w:tc>{properties}"));
                let margins = styles::margins(table, Some(anchor.cell))?;
                let inner_width = (width - ((margins[1] + margins[3]) * 15.0).round() as i64).max(1);
                let blocks = if continuation { continuation_blocks(table, anchor.cell, row_index) } else { cell_blocks(anchor.cell) };
                for block in blocks {
                    match get(block, "type").as_str() {
                        Some("table") => contents.push_str(&write(context, block, Some(inner_width))?),
                        Some("sectionBreak") => return Err(ExportError::invalid("DOCX section breaks cannot be written inside a table cell")),
                        _ => contents.push_str(&context.block(block)?),
                    }
                }
                // A native cell always ends in a paragraph. The marker lets
                // import distinguish a format-required tail from authored text.
                if blocks.is_empty() || blocks.last().is_some_and(|block| get(block, "type").as_str() == Some("table")) { contents.push_str("<w:p ov:tableEnd=\"1\"/>"); }
                contents.push_str("</w:tc>");
            }
            contents.push_str("</w:tr>");
        }
        let result = format!("<w:tbl xmlns:w=\"{WORD_NAMESPACE}\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:ov=\"{TABLE_SOURCE_NAMESPACE}\" mc:Ignorable=\"ov\" ov:key=\"{key}\">{properties}<w:tblGrid>{columns}</w:tblGrid>{contents}</w:tbl>");
        let identity = native_identity(&result).map_err(ExportError::invalid)?;
        let source_data = data(table, &id, &grid);
        let entry = TableSourceEntry { key, table_id: id, native_signature: identity.native_signature,
            topology_signature: identity.topology_signature, data: source_data };
        let bytes = serde_json::to_vec(&entry).map_err(|error| ExportError::invalid(error.to_string()))?.len();
        context.tables.source_bytes = context.tables.source_bytes.checked_add(bytes).filter(|bytes| *bytes <= MAX_TABLE_SOURCE_BYTES)
            .ok_or_else(|| ExportError::invalid("DOCX_TABLE_SOURCE_RESOURCE_LIMIT: document table metadata exceeds 16 MiB"))?;
        context.tables.entries.push(entry);
        Ok(result)
    })();
    context.tables.depth -= 1;
    result
}

pub(super) fn write_sources(context: &mut WriteContext<'_>) -> Result<(), ExportError> {
    if context.tables.entries.is_empty() { return Ok(()); }
    validate_entries(&context.tables.entries).map_err(ExportError::invalid)?;
    let mut body = String::new();
    for entry in &context.tables.entries {
        let data = serde_json::to_string(&entry.data).map_err(|error| ExportError::invalid(error.to_string()))?;
        body.push_str(&format!("<table key=\"{}\" tableId=\"{}\" nativeSignature=\"{}\" topologySignature=\"{}\"><data>{}</data></table>",
            xml_attribute(&entry.key), xml_attribute(&entry.table_id), entry.native_signature, entry.topology_signature, xml_text(&data)));
    }
    let result = format!("{}<tables xmlns=\"{TABLE_SOURCE_NAMESPACE}\" version=\"1\">{body}</tables>", xml::DECL);
    if result.len() > MAX_TABLE_SOURCE_BYTES { return Err(ExportError::invalid("DOCX_TABLE_SOURCE_RESOURCE_LIMIT: escaped table source XML exceeds 16 MiB")); }
    context.package.add_xml(TABLE_SOURCE_PART, "application/xml", result)?;
    context.relationship("http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml", TABLE_SOURCE_TARGET, false)?;
    Ok(())
}

#[cfg(test)]
#[path = "docx_table_tests.rs"]
mod tests;
