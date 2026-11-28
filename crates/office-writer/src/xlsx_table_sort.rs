//! Native sort metadata follows table geometry and stable column identities.
//! Applying a sort remains a model command; export never reorders source cells.

use office_io_common::{conditional_format_xml::Element, sheet_table_source as source};
use std::collections::BTreeMap;
use crate::ExportError;
use super::{TableInput, bounds};

pub(super) struct SourceGeometry {
    range: [u32; 4],
    header: u32,
    totals: u32,
}

impl SourceGeometry {
    pub(super) fn read(root: &Element) -> Result<Option<Self>, ExportError> {
        let Some(reference) = root.attrs.get("ref") else { return Ok(None); };
        let range = source::range(reference).ok_or_else(|| invalid("invalid source table reference"))?;
        let count = |name: &str, default: u32| -> Result<u32, ExportError> {
            root.attrs.get(name).map_or(Ok(default), |value| value.parse::<u32>().map_err(|_| invalid("invalid source table row count")))
        };
        let header = count("headerRowCount", 1)?;
        let totals = count("totalsRowCount", 0)?;
        if header > 1 || totals > 1 || header + totals > range[2] - range[0] + 1 { return Err(invalid("invalid source table row counts")); }
        Ok(Some(Self { range, header, totals }))
    }
}

fn invalid(detail: &str) -> ExportError { ExportError::invalid(format!("TABLE_SOURCE_INVALID: {detail}")) }
fn inside(inner: [u32; 4], outer: [u32; 4]) -> bool {
    inner[0] >= outer[0] && inner[1] >= outer[1] && inner[2] <= outer[2] && inner[3] <= outer[3]
}
fn reference(range: [u32; 4]) -> String {
    format!("{}{}:{}{}", super::super::column_name(range[1]), range[0] + 1, super::super::column_name(range[3]), range[2] + 1)
}

struct Projection<'a> {
    source: &'a SourceGeometry,
    current: [u32; 4],
    header: u32,
    totals: u32,
    previous_columns: &'a [u32],
    column_positions: BTreeMap<u32, usize>,
}

impl Projection<'_> {
    fn rows(&self, start: u32, end: u32) -> Option<(u32, u32)> {
        let before = self.source.range;
        let old_data_start = before[0] + self.source.header;
        let old_data_end = before[2].checked_sub(self.source.totals)?;
        let data_start = self.current[0] + self.header;
        let data_end = self.current[2].checked_sub(self.totals)?;
        let start = if start == old_data_start { data_start } else { self.current[0].checked_add(start.checked_sub(before[0])?)? };
        let end = if end == old_data_end { data_end }
            else if end == before[2] { self.current[2] }
            else { self.current[0].checked_add(end.checked_sub(before[0])?)?.min(data_end) };
        (start <= end && start >= self.current[0] && end <= self.current[2]).then_some((start, end))
    }

    fn columns(&self, start: u32, end: u32) -> Result<Option<(u32, u32)>, ExportError> {
        if start == self.source.range[1] && end == self.source.range[3] {
            return Ok(Some((self.current[1], self.current[3])));
        }
        let mut positions = self.previous_columns[(start - self.source.range[1]) as usize..=(end - self.source.range[1]) as usize]
            .iter().filter_map(|id| self.column_positions.get(id).copied()).collect::<Vec<_>>();
        if positions.is_empty() { return Ok(None); }
        positions.sort_unstable();
        if positions.windows(2).any(|pair| pair[1] != pair[0] + 1) { return Err(invalid("native sort region became discontiguous after column edits")); }
        Ok(Some((self.current[1] + positions[0] as u32, self.current[1] + positions[positions.len() - 1] as u32)))
    }

    fn range(&self, node: &Element) -> Result<[u32; 4], ExportError> {
        let range = node.attrs.get("ref").and_then(|value| source::range(value)).ok_or_else(|| invalid("invalid native sort reference"))?;
        if !inside(range, self.source.range) { return Err(invalid("native sort reference lies outside its source table")); }
        Ok(range)
    }

    fn sort(&self, state: &Element) -> Result<Option<Element>, ExportError> {
        let previous = self.range(state)?;
        if self.current[2].checked_sub(self.totals).is_none_or(|end| self.current[0] + self.header > end) {
            return Ok(None);
        }
        let Some((r1, r2)) = self.rows(previous[0], previous[2]) else { return Ok(None); };
        let Some((c1, c2)) = self.columns(previous[1], previous[3])? else { return Ok(None); };
        let current = [r1, c1, r2, c2];
        let column_sort = source::flag(state, "columnSort").map_err(ExportError::invalid)?.unwrap_or(false);
        let mut next = state.clone(); let mut children = Vec::new(); let mut conditions = 0;
        for child in &state.children {
            if !source::spreadsheet(child, "sortCondition") { children.push(child.clone()); continue; }
            let before = self.range(child)?;
            if !inside(before, previous) || (!column_sort && before[1] != before[3]) || (column_sort && before[0] != before[2]) {
                return Err(invalid("native sort key must be one row or column inside the sort region"));
            }
            let mapped = if column_sort {
                let Some(row) = self.current[0].checked_add(before[0] - self.source.range[0]).filter(|row| *row >= r1 && *row <= r2) else { continue; };
                [row, c1, row, c2]
            } else {
                let id = self.previous_columns[(before[1] - self.source.range[1]) as usize];
                let Some(index) = self.column_positions.get(&id) else { continue; };
                let col = self.current[1] + *index as u32;
                if col < c1 || col > c2 { return Err(invalid("native sort key moved outside the sort region")); }
                [r1, col, r2, col]
            };
            let mut child = child.clone(); child.set_attr("ref", reference(mapped)); children.push(child); conditions += 1;
        }
        if conditions == 0 { return Ok(None); }
        next.set_attr("ref", reference(current));
        next.replace_children(|_| true, children);
        Ok(Some(next))
    }

    fn container(&self, node: &mut Element) -> Result<(), ExportError> {
        let replacement = node.children.iter().filter(|child| source::spreadsheet(child, "sortState"))
            .map(|state| self.sort(state)).collect::<Result<Vec<_>, _>>()?.into_iter().flatten().collect();
        node.replace_children(|child| source::spreadsheet(child, "sortState"), replacement);
        Ok(())
    }
}

pub(super) fn patch(root: &mut Element, geometry: Option<&SourceGeometry>, table: &TableInput, original_order: &[u32]) -> Result<(), ExportError> {
    let Some(geometry) = geometry else { return Ok(()); };
    if original_order.len() != (geometry.range[3] - geometry.range[1] + 1) as usize { return Err(invalid("source column count differs from the native table range")); }
    let (r1, c1, r2, c2) = bounds(&table.range).ok_or_else(|| invalid("invalid current table range"))?;
    let column_positions = table.columns.iter().enumerate().map(|(index, column)| column.office_column_id
        .map(|id| (id, index)).ok_or_else(|| invalid("unprepared column identity"))).collect::<Result<_, _>>()?;
    let projection = Projection { source: geometry, current: [r1, c1, r2, c2], header: u32::from(table.show_header_row),
        totals: u32::from(table.show_totals_row), previous_columns: original_order, column_positions };
    projection.container(root)?;
    if let Some(filter) = root.child_mut("autoFilter") { projection.container(filter)?; }
    Ok(())
}
