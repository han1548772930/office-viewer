//! Validate control ownership across paragraphs, physical rows and real story
//! hosts before independent writers assemble their native SDT wrappers.
use std::collections::{HashMap, HashSet};

use serde_json::Value;

use super::{ContentControl, control_pictures, control_writer, get, note_writer::{NoteBoundary, NotePositions}, slice_utf16, table_writer};
use crate::ExportError;

struct BlockAddress<'a> {
    block: &'a Value,
    container: &'a [Value],
    host: usize,
    start: usize,
    end: usize,
    row_spans: Vec<(usize, usize)>,
    note_positions: NotePositions,
}
enum RegionKind {
    Inline { block_id: String, start: NoteBoundary, end: NoteBoundary },
    Blocks,
    Rows { table_id: String, start: usize, end: usize },
}
struct Region { host: usize, start: NoteBoundary, end: NoteBoundary, owned: HashSet<String>, kind: RegionKind }

fn advance(cursor: &mut usize, count: usize) -> Result<(), ExportError> {
    *cursor = cursor.checked_add(count).ok_or_else(|| ExportError::invalid("DOCX SDT physical position capacity exceeded"))?;
    Ok(())
}

fn index_blocks<'a>(blocks: &'a [Value], host: usize, cursor: &mut usize, addresses: &mut HashMap<String, BlockAddress<'a>>) -> Result<(), ExportError> {
    for block in blocks {
        let start = *cursor;
        advance(cursor, 1)?;
        let mut row_spans = Vec::new();
        if get(block, "type").as_str() == Some("table") {
            let table = get(get(block, "metadata"), "table");
            let rows = get(table, "rows").as_array().ok_or_else(|| ExportError::invalid("DOCX SDT table has no physical rows"))?;
            let mut groups = table_writer::content_groups(table)?.into_iter().peekable();
            for row_index in 0..rows.len() {
                let row_start = *cursor;
                advance(cursor, 1)?;
                while groups.peek().is_some_and(|group| group.row_index == row_index) {
                    let group = groups.next().expect("The physical row content group was inspected");
                    index_blocks(group.blocks, host, cursor, addresses)?;
                }
                advance(cursor, 1)?;
                row_spans.push((row_start, *cursor));
            }
            if groups.next().is_some() { return Err(ExportError::invalid("DOCX SDT table content exceeds its physical rows")); }
        } else {
            for run in get(block, "runs").as_array().into_iter().flatten() {
                if let Some(text) = get(run, "text").as_str() { advance(cursor, text.encode_utf16().count())?; }
            }
        }
        advance(cursor, 1)?;
        if let Some(id) = get(block, "id").as_str() {
            let runs = get(block, "runs").as_array().map_or(&[][..], Vec::as_slice);
            let note_positions = NotePositions::from_runs(runs)?;
            if addresses.insert(id.to_owned(), BlockAddress { block, container: blocks, host, start, end: *cursor, row_spans, note_positions }).is_some() {
                return Err(ExportError::invalid("DOCX SDT input blocks must have unique physical ownership"));
            }
        }
    }
    Ok(())
}

fn collect_owned(blocks: &[Value], owned: &mut HashSet<String>) -> Result<(), ExportError> {
    for block in blocks {
        if let Some(id) = get(block, "id").as_str() { owned.insert(id.to_owned()); }
        if get(block, "type").as_str() == Some("table") {
            for group in table_writer::content_groups(get(get(block, "metadata"), "table"))? { collect_owned(group.blocks, owned)?; }
        }
    }
    Ok(())
}

fn region(control: &ContentControl, addresses: &HashMap<String, BlockAddress<'_>>) -> Result<Region, ExportError> {
    let fallback_id = get(&control.value, "blockId").as_str().ok_or_else(|| ExportError::invalid("DOCX SDT requires an input anchor"))?;
    let fallback = addresses.get(fallback_id).ok_or_else(|| ExportError::invalid("DOCX SDT input anchor does not exist"))?;
    if !matches!(get(fallback.block, "type").as_str(), Some("paragraph" | "heading" | "bullet" | "numbered" | "checklist" | "quote")) {
        return Err(ExportError::invalid("DOCX SDT input anchor must be an editable text paragraph"));
    }
    let text = get(fallback.block, "runs").as_array().into_iter().flatten().filter_map(|run| get(run, "text").as_str()).collect::<String>();
    if slice_utf16(&text, control.start, control.end).is_none() {
        return Err(ExportError::invalid("DOCX SDT input anchor range exceeds its text or splits a UTF-16 surrogate pair"));
    }
    let start_rank = fallback.note_positions.boundary_rank(control.start, control.start_note_reference_count)?;
    let end_rank = fallback.note_positions.boundary_rank(control.end, control.end_note_reference_count)?;
    if (control.start, start_rank) > (control.end, end_rank) { return Err(ExportError::invalid("DOCX SDT note-reference endpoints are reversed")); }
    let mut owned = HashSet::new();
    let (host, span_start, span_end, kind) = if get(&control.value, "placement").as_str() == Some("row") {
        let table_id = get(get(&control.value, "tableRowRange"), "tableId").as_str().ok_or_else(|| ExportError::invalid("DOCX row SDT requires a table"))?;
        let address = addresses.get(table_id).ok_or_else(|| ExportError::invalid("DOCX row SDT table does not exist"))?;
        if get(address.block, "type").as_str() != Some("table") { return Err(ExportError::invalid("DOCX row SDT must refer to a real table block")); }
        let table = get(get(address.block, "metadata"), "table");
        let rows = get(table, "rows").as_array().ok_or_else(|| ExportError::invalid("DOCX row SDT table has no physical rows"))?;
        let (start, end) = control_writer::row_range(control, table_id, rows)?;
        for group in table_writer::content_groups(table)?.into_iter().filter(|group| group.row_index >= start && group.row_index < end) {
            collect_owned(group.blocks, &mut owned)?;
        }
        (address.host, (address.row_spans[start].0, 0), (address.row_spans[end - 1].1, 0), RegionKind::Rows { table_id: table_id.into(), start, end })
    } else if get(&control.value, "placement").as_str() == Some("block") || !get(&control.value, "blockRange").is_null() {
        let first_id = get(get(&control.value, "blockRange"), "blockIds").as_array().and_then(|ids| ids.first()).and_then(Value::as_str).unwrap_or(fallback_id);
        let address = addresses.get(first_id).ok_or_else(|| ExportError::invalid("DOCX block SDT first block does not exist"))?;
        let (start, end) = control_writer::block_range(control, address.container)?;
        collect_owned(&address.container[start..end], &mut owned)?;
        let last_id = get(&address.container[end - 1], "id").as_str().ok_or_else(|| ExportError::invalid("DOCX block SDT final block has no identity"))?;
        let last = addresses.get(last_id).ok_or_else(|| ExportError::invalid("DOCX block SDT final block has no physical address"))?;
        (address.host, (address.start, 0), (last.end, 0), RegionKind::Blocks)
    } else {
        if !get(&control.value, "tableRowRange").is_null() { return Err(ExportError::invalid("DOCX SDT row ownership requires row placement")); }
        owned.insert(fallback_id.to_owned());
        (fallback.host, (fallback.start + 1 + control.start, start_rank), (fallback.start + 1 + control.end, end_rank),
            RegionKind::Inline { block_id: fallback_id.into(), start: (control.start, start_rank), end: (control.end, end_rank) })
    };
    if fallback.host != host || !owned.contains(fallback_id) {
        return Err(ExportError::invalid("DOCX structural SDT input anchor is outside its owned blocks or physical rows"));
    }
    Ok(Region { host, start: span_start, end: span_end, owned, kind })
}

fn contains(parent: &Region, child: &Region) -> bool {
    if parent.host != child.host || parent.start > child.start || child.end > parent.end || !child.owned.is_subset(&parent.owned) { return false; }
    match (&parent.kind, &child.kind) {
        (RegionKind::Inline { block_id: a, start: from, end: to }, RegionKind::Inline { block_id: b, start, end }) => a == b && from <= start && end <= to,
        (RegionKind::Inline { .. }, _) => false,
        (RegionKind::Rows { table_id: a, start: from, end: to }, RegionKind::Rows { table_id: b, start, end }) if a == b => from <= start && end <= to,
        _ => true,
    }
}

pub(super) fn validate(groups: &[&[Value]], controls: &HashMap<String, Vec<ContentControl>>) -> Result<(), ExportError> {
    let mut addresses = HashMap::new();
    for (host, blocks) in groups.iter().enumerate() {
        let mut cursor = 0;
        index_blocks(blocks, host, &mut cursor, &mut addresses)?;
    }
    let mut regions = HashMap::new();
    let mut pictures = HashMap::new();
    for control in controls.values().flatten() {
        let region = region(control, &addresses)?;
        if control.kind == "picture" {
            let owned = region.owned.iter().filter_map(|id| addresses.get(id).map(|address| address.block)).collect::<Vec<_>>();
            if let Some(image) = control_pictures::validate(control, &owned)? {
                if pictures.insert(image, control.id.as_str()).is_some() {
                    return Err(ExportError::invalid("DOCX picture images must have one scalar content-control owner"));
                }
            }
        }
        regions.insert(control.id.as_str(), region);
    }
    for control in controls.values().flatten() {
        if let Some(parent_id) = get(&control.value, "parentId").as_str() {
            let parent = regions.get(parent_id).ok_or_else(|| ExportError::invalid("DOCX SDT has no physical parent region"))?;
            if !contains(parent, &regions[control.id.as_str()]) {
                return Err(ExportError::invalid("DOCX SDT parent must contain the child's real region in the same story host"));
            }
        }
    }
    // One physical interval sweep catches crossings between inline, block and
    // row placements. Structural delimiters distinguish empty paragraphs and
    // rows without building an editor model or comparing every pair of SDTs.
    let mut ordered = regions.iter().collect::<Vec<_>>();
    ordered.sort_by(|(a_id, a), (b_id, b)| a.host.cmp(&b.host).then(a.start.cmp(&b.start)).then(b.end.cmp(&a.end)).then(a_id.cmp(b_id)));
    let mut stack: Vec<&Region> = Vec::new();
    for (_, region) in ordered {
        while stack.last().is_some_and(|parent| parent.host != region.host || region.start >= parent.end) { stack.pop(); }
        if stack.last().is_some_and(|parent| region.end > parent.end) {
            return Err(ExportError::invalid("DOCX SDT physical ranges cross between inline, block or row placements"));
        }
        if region.start < region.end { stack.push(region); }
    }
    Ok(())
}
