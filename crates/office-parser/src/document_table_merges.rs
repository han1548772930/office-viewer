//! Canonicalize legacy hMerge into a real logical grid on the Duke path.
//! Physical member paragraphs keep their IDs and source refs. Original member
//! properties are provenance; vertical continuation ownership remains separate.
use std::collections::BTreeMap;

use duke_pages::{BlockDetails, BlockType, DocumentBlock};
use serde_json::{Map, Value, json};

const MAX_ROWS: usize = 10_000;
const MAX_COLUMNS: usize = 512;
const MAX_GRID_SLOTS: usize = 1_000_000;

#[derive(Debug)]
struct CellPlan {
    column: usize,
    span: usize,
    members: Vec<Value>,
    vertical: Option<String>,
    row_span: usize,
}

fn integer(value: &Value, name: &str) -> Result<usize, String> {
    value.get(name).and_then(Value::as_u64).and_then(|value| usize::try_from(value).ok())
        .ok_or_else(|| format!("DOCX_TABLE_HMERGE_INVALID: {name} must be a nonnegative integer"))
}

fn merge(value: &Value, name: &str) -> Result<Option<String>, String> {
    match value.get(name).filter(|value| !value.is_null()) {
        None => Ok(None),
        Some(Value::String(value)) if matches!(value.as_str(), "restart" | "continue") => Ok(Some(value.clone())),
        _ => Err(format!("DOCX_TABLE_HMERGE_INVALID: {name} must be restart or continue")),
    }
}

fn plan(native: &Value) -> Result<Vec<Vec<CellPlan>>, String> {
    let rows = native.get("rows").and_then(Value::as_array).filter(|rows| !rows.is_empty() && rows.len() <= MAX_ROWS)
        .ok_or("DOCX_TABLE_RESOURCE_LIMIT: legacy merged tables require 1–10,000 physical rows")?;
    let mut output = Vec::<Vec<CellPlan>>::with_capacity(rows.len());
    let mut columns = 0usize;
    let mut active = BTreeMap::<(usize, usize), (usize, usize)>::new();
    for (row_index, row) in rows.iter().enumerate() {
        let cells = row.get("cells").and_then(Value::as_array).ok_or("DOCX_TABLE_HMERGE_INVALID: native row cells are missing")?;
        let mut planned = Vec::<CellPlan>::new();
        for cell in cells {
            let column = integer(cell, "columnIndex")?;
            let span = integer(cell, "columnSpan")?;
            let end = column.checked_add(span).filter(|end| span > 0 && *end <= MAX_COLUMNS)
                .ok_or("DOCX_TABLE_RESOURCE_LIMIT: legacy merged cell exceeds 512 logical columns")?;
            columns = columns.max(end);
            if rows.len().checked_mul(columns).is_none_or(|slots| slots > MAX_GRID_SLOTS) {
                return Err("DOCX_TABLE_RESOURCE_LIMIT: legacy merged grid exceeds 1,000,000 slots".into());
            }
            let horizontal = merge(cell, "horizontalMerge")?;
            let vertical = merge(cell, "verticalMerge")?;
            if horizontal.as_deref() == Some("continue") {
                let preceding = planned.last_mut().filter(|preceding| preceding.members[0]["horizontalMerge"] == "restart")
                    .ok_or("DOCX_TABLE_HMERGE_INVALID: horizontal continuation has no restart in its physical row")?;
                if preceding.column + preceding.span != column || preceding.vertical != vertical {
                    return Err("DOCX_TABLE_HMERGE_INVALID: horizontal members are not contiguous with the same vertical merge state".into());
                }
                preceding.span += span;
                preceding.members.push(cell.clone());
            } else {
                if planned.last().is_some_and(|preceding| preceding.column + preceding.span > column) {
                    return Err("DOCX_TABLE_HMERGE_INVALID: native cells overlap".into());
                }
                planned.push(CellPlan { column, span, members: vec![cell.clone()], vertical, row_span: 1 });
            }
        }
        let mut next_active = BTreeMap::new();
        for (cell_index, cell) in planned.iter().enumerate() {
            let key = (cell.column, cell.span);
            match cell.vertical.as_deref() {
                Some("restart") => { next_active.insert(key, (row_index, cell_index)); }
                Some("continue") => {
                    let owner = *active.get(&key).ok_or("DOCX_TABLE_VMERGE_INVALID: vertical continuation has no matching logical interval in the preceding row")?;
                    output[owner.0][owner.1].row_span += 1;
                    next_active.insert(key, owner);
                }
                _ => {},
            }
        }
        active = next_active;
        output.push(planned);
    }
    Ok(output)
}

fn details(block: &mut DocumentBlock) -> &mut Map<String, Value> {
    if !matches!(&block.details, BlockDetails::Raw(Value::Object(_))) {
        block.details = BlockDetails::Raw(block.details.to_json_value());
    }
    match &mut block.details {
        BlockDetails::Raw(Value::Object(value)) => value,
        _ => unreachable!("table cell details are objects"),
    }
}

fn insert_cell(cells: &mut BTreeMap<(usize, usize), DocumentBlock>, cell: DocumentBlock) -> Result<(), String> {
    if cell.r#type != BlockType::TableCell { return Err("DOCX_TABLE_HMERGE_INVALID: a table row contains a non-cell block".into()); }
    let detail = cell.details.to_json_value();
    let key = (integer(&detail, "rowIndex")?, integer(&detail, "columnIndex")?);
    if cells.insert(key, cell).is_some() { return Err("DOCX_TABLE_HMERGE_INVALID: extracted cells have duplicate coordinates".into()); }
    Ok(())
}

/// Only legacy horizontal tables need this overlay. Other source grids and
/// their current provenance are left to the existing native adapters.
pub(crate) fn restore(table: &mut DocumentBlock, native: &mut Value) -> Result<(), String> {
    let horizontal = native.get("rows").and_then(Value::as_array).into_iter().flatten()
        .flat_map(|row| row.get("cells").and_then(Value::as_array).into_iter().flatten())
        .any(|cell| cell.get("horizontalMerge").is_some_and(|value| !value.is_null()));
    if !horizontal { return Ok(()); }
    let planned = plan(native)?;
    let mut cells = BTreeMap::new();
    let mut heads = BTreeMap::new();
    for mut child in std::mem::take(&mut table.children) {
        if child.r#type == BlockType::TableHead {
            let row = integer(&child.details.to_json_value(), "rowIndex")?;
            for cell in std::mem::take(&mut child.children) { insert_cell(&mut cells, cell)?; }
            if heads.insert(row, child).is_some() { return Err("DOCX_TABLE_HMERGE_INVALID: duplicate header rows".into()); }
        } else { insert_cell(&mut cells, child)?; }
    }
    let mut row_texts = Vec::with_capacity(planned.len());
    for (row_index, row) in planned.into_iter().enumerate() {
        let mut row_cells = Vec::with_capacity(row.len());
        let mut native_cells = Vec::with_capacity(row.len());
        for group in row {
            let mut anchor = None::<DocumentBlock>;
            let mut members = Vec::new();
            let mut right_border = None;
            for member in &group.members {
                let column = integer(member, "columnIndex")?;
                let mut cell = cells.remove(&(row_index, column))
                    .ok_or("DOCX_TABLE_HMERGE_INVALID: a physical member could not be mapped to an extracted cell")?;
                let original = cell.details.to_json_value();
                if integer(&original, "columnSpan")? != integer(member, "columnSpan")? {
                    return Err("DOCX_TABLE_HMERGE_INVALID: extracted and native horizontal member spans differ".into());
                }
                if member["syntheticTail"] == true && cell.children.last().is_some_and(|last|
                    last.content.is_empty() && last.children.is_empty() && matches!(&last.details, BlockDetails::Paragraph(_))) {
                    cell.children.pop();
                }
                members.push(json!({"cellId":cell.id,"columnIndex":column,"columnSpan":member["columnSpan"],
                    "blockIds":cell.children.iter().map(|block| &block.id).collect::<Vec<_>>(),
                    "details":original,"nativeProperties":member["nativeProperties"]}));
                right_border = original.get("style").and_then(|style| style.get("borders")).and_then(|borders| borders.get("right")).cloned();
                if let Some(anchor) = &mut anchor {
                    for block in &mut cell.children { block.parent_block_id = Some(anchor.id.clone()); }
                    anchor.children.append(&mut cell.children);
                    for source in cell.source_refs {
                        if !anchor.source_refs.contains(&source) { anchor.source_refs.push(source); }
                    }
                } else { anchor = Some(cell); }
            }
            let mut anchor = anchor.ok_or("DOCX_TABLE_HMERGE_INVALID: a logical cell has no physical members")?;
            let current = details(&mut anchor);
            current.insert("columnSpan".into(), json!(group.span));
            current.insert("rowSpan".into(), json!(group.row_span));
            let style = current.entry("style").or_insert_with(|| json!({}));
            if !style.is_object() { *style = json!({}); }
            style["gridSpan"] = json!(group.span);
            style["rowSpan"] = json!(group.row_span);
            style["vMergeContinuation"] = json!(group.vertical.as_deref() == Some("continue"));
            let mut source = group.members[0].clone();
            source["columnSpan"] = json!(group.span);
            if group.members.len() > 1 {
                if let Some(border) = right_border {
                    if !style["borders"].is_object() { style["borders"] = json!({}); }
                    style["borders"]["right"] = border;
                }
                source["nativeProperties"]["sourceHorizontalMerge"] = json!({"version":1,"members":members});
                source["nativeProperties"]["rewriteProperties"] = json!(["tcW","tcBorders"]);
            }
            // Every format-required member tail was handled above. Any final
            // paragraph here is authored and must survive normalization.
            source["syntheticTail"] = json!(false);
            anchor.content = anchor.children.iter().map(|block| block.content.as_str()).collect::<Vec<_>>().join("\n");
            native_cells.push(source);
            row_cells.push(anchor);
        }
        let text = row_cells.iter().map(|cell| cell.content.as_str()).collect::<Vec<_>>().join("\t");
        row_texts.push(text.clone());
        native["rows"][row_index]["cells"] = json!(native_cells);
        if let Some(mut head) = heads.remove(&row_index) {
            head.content = text; head.children = row_cells; table.children.push(head);
        } else { table.children.extend(row_cells); }
    }
    if !cells.is_empty() || !heads.is_empty() { return Err("DOCX_TABLE_HMERGE_INVALID: extracted table members remain unmapped".into()); }
    table.content = row_texts.join("\n");
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use duke_pages::DocModel;
    use office_io_common::document_table_source::WORD_NAMESPACE;

    fn cell(column: usize, span: usize, horizontal: Option<&str>, vertical: Option<&str>) -> Value {
        json!({"columnIndex":column,"columnSpan":span,"horizontalMerge":horizontal,"verticalMerge":vertical,
            "syntheticTail":false,"nativeProperties":{"version":1,"propertiesXml":"<w:tcPr/>","namespaces":{"w":WORD_NAMESPACE}}})
    }

    #[test]
    fn mixed_grid_spans_fold_horizontal_runs_and_recompute_vertical_ownership_once_per_row() {
        let native = json!({"rows":[
            {"cells":[cell(0, 1, Some("restart"), Some("restart")),cell(1, 2, Some("continue"), Some("restart"))]},
            {"cells":[cell(0, 3, None, Some("continue"))]},
            {"cells":[cell(0, 1, Some("restart"), Some("continue")),cell(1, 2, Some("continue"), Some("continue"))]}
        ]});
        let planned = plan(&native).unwrap();
        assert_eq!(planned[0].len(), 1);
        assert_eq!(planned[0][0].span, 3);
        assert_eq!(planned[0][0].row_span, 3);
        assert_eq!(planned[1][0].vertical.as_deref(), Some("continue"));
        assert_eq!(planned[2][0].members.len(), 2);
    }

    #[test]
    fn orphan_horizontal_members_and_changed_vertical_intervals_are_rejected() {
        let orphan = json!({"rows":[{"cells":[cell(0, 1, Some("continue"), None)]}]});
        assert!(plan(&orphan).unwrap_err().contains("no restart"));
        let interrupted = json!({"rows":[{"cells":[cell(0, 1, Some("restart"), None),cell(1, 1, None, None),cell(2, 1, Some("continue"), None)]}]});
        assert!(plan(&interrupted).unwrap_err().contains("no restart"));
        let vertical = json!({"rows":[
            {"cells":[cell(0, 1, Some("restart"), Some("restart")),cell(1, 1, Some("continue"), Some("restart"))]},
            {"cells":[cell(0, 1, None, Some("continue")),cell(1, 1, None, Some("continue"))]}
        ]});
        assert!(plan(&vertical).unwrap_err().contains("no matching logical interval"));
        let oversized = json!({"rows":[{"cells":[cell(511, 2, Some("restart"), None)]}]});
        assert!(plan(&oversized).unwrap_err().contains("512 logical columns"));
    }

    #[test]
    fn real_extracted_header_members_keep_paragraph_ids_order_and_provenance() {
        let mut source: DocModel = serde_json::from_value(json!({"nodes":[{"type":"table","blockId":"table","rows":[
            {"style":{"isHeader":true},"cells":[
                {"nodes":[{"type":"paragraph","blockId":"first","children":[{"type":"text","text":"A"}]}]},
                {"nodes":[{"type":"paragraph","blockId":"second","children":[{"type":"text","text":"B😀"}]},
                    {"type":"paragraph","blockId":"format-tail","children":[]}]}
            ]}
        ]}],"metadata":{"sourceParts":1,"warnings":[],"headerSections":[],"footerSections":[],"paragraphStyles":[]}})).unwrap();
        let mut document = crate::document_block_extraction::extract(&mut source);
        let mut native = json!({"version":1,"rows":[{"cells":[cell(0, 1, Some("restart"), None),cell(1, 1, Some("continue"), None)]}]});
        native["rows"][0]["cells"][1]["syntheticTail"] = json!(true);
        let table = &mut document.blocks[0];
        restore(table, &mut native).unwrap();
        assert_eq!(table.children[0].r#type, BlockType::TableHead);
        let head = &table.children[0];
        assert_eq!(head.children.len(), 1);
        let owner = &head.children[0];
        assert_eq!(owner.details.to_json_value()["columnSpan"], 2);
        assert_eq!(owner.children.iter().map(|block| block.id.as_str()).collect::<Vec<_>>(), ["first","second"]);
        assert_eq!(owner.content, "A\nB😀");
        assert!(owner.children.iter().all(|block| block.parent_block_id.as_deref() == Some(owner.id.as_str())));
        let provenance = &native["rows"][0]["cells"][0]["nativeProperties"]["sourceHorizontalMerge"];
        assert_eq!(provenance["version"], 1);
        assert_eq!(provenance["members"][1]["blockIds"], json!(["second"]));
        assert_eq!(native["rows"][0]["cells"][0]["nativeProperties"]["rewriteProperties"], json!(["tcW","tcBorders"]));
    }
}
