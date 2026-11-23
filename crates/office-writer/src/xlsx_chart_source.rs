//! Snapshot visibility uses the same manual, size, filter and outline reasons
//! as the grid. Intervals remain intervals even for a whole-axis collapse.

use std::collections::BTreeSet;
use crate::ExportError;
use super::{ChartInput, Context, WorkbookInput};

pub(super) struct SourceView {
    pub points: Vec<usize>,
    pub fields: Option<BTreeSet<usize>>,
}

pub(super) fn read(workbook: &WorkbookInput, chart: &ChartInput, source: &Context<'_>) -> Result<SourceView, ExportError> {
    let mut points = Vec::new();
    points.try_reserve(source.point_count).map_err(|_| ExportError::invalid("CHART_HELPER_RESOURCE_LIMIT: source visibility exceeds its point budget"))?;
    if !chart.options.plot_visible_only.unwrap_or(true) {
        points.extend(0..source.point_count);
        return Ok(SourceView { points, fields: None });
    }
    let axis = super::super::dimensions::prepare(source.sheet)?;
    let mut hidden_rows = super::super::auto_filter::hidden_rows(source.sheet);
    for table in workbook.tables.iter().filter(|table| table.sheet_id == source.sheet.id) {
        for row in table.auto_filter.iter().flat_map(|filter| &filter.filtered_rows) { hidden_rows.insert(*row); }
    }
    let collapsed = source.sheet.outline.iter().flat_map(|outline| &outline.groups).filter(|group| group.collapsed).collect::<Vec<_>>();
    let row_visible = |row: u32| !hidden_rows.contains(&row) && !axis.manual_row_hidden(row) && axis.row_size(row).height > 0.0
        && !collapsed.iter().any(|group| group.axis == "row" && row >= group.start && row <= group.end);
    let column_visible = |column: u32| !axis.column_hidden(column)
        && !collapsed.iter().any(|group| group.axis == "column" && column >= group.start && column <= group.end);
    let mut fields = BTreeSet::new();
    for field in 0..source.field_count {
        let (row, column) = source.coordinates(chart, field, None)?;
        if if chart.source.orientation == "columns" { column_visible(column) } else { row_visible(row) } { fields.insert(field); }
    }
    for point in 0..source.point_count {
        let (row, column) = source.coordinates(chart, 0, Some(point))?;
        if if chart.source.orientation == "columns" { row_visible(row) } else { column_visible(column) } { points.push(point); }
    }
    Ok(SourceView { points, fields: Some(fields) })
}
