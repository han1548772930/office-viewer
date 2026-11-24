//! Shared multi-series waterfall geometry. Input is borrowed from the retained
//! numeric arena; the output is filled in place without per-point control DTOs.
//! Fixed reference: engine-chart 1.6.0 ES ake/oke/ske (series/category order).

use super::{ChartSeries, ChartStatisticsResult, Input, JobSummary, WaterfallStackType, MAX_CHART_SERIES, MAX_ROWS};
use super::math::finite;
use crate::ComputeError;

const TOTAL: u32 = 1;
const SYNTHETIC: u32 = 2;

pub(super) fn record_counts(rows: usize, series: usize, stack: WaterfallStackType, subtotal: bool) -> Result<(usize, usize), ComputeError> {
    if rows > MAX_ROWS || series == 0 || series > MAX_CHART_SERIES { return Err(ComputeError::memory()); }
    if rows == 0 { return Ok((0, 0)); }
    let source = rows.checked_mul(series).ok_or_else(ComputeError::memory)?;
    let extra = if subtotal { if stack == WaterfallStackType::Stacked { 1 } else { series } } else { 0 };
    let points = source.checked_add(extra).ok_or_else(ComputeError::memory)?;
    u32::try_from(points).map_err(|_| ComputeError::memory())?;
    let connectors = if stack == WaterfallStackType::Stacked { rows - 1 } else { source - 1 };
    Ok((points, connectors))
}

pub(super) fn output_capacity(rows: usize, series: usize, stack: WaterfallStackType, subtotal: bool) -> Result<(usize, usize), ComputeError> {
    let (points, connectors) = record_counts(rows, series, stack, subtotal)?;
    let values = points.checked_mul(3).and_then(|size| size.checked_add(connectors.checked_mul(2)?)).and_then(|size| size.checked_add(1));
    let indices = points.checked_mul(4).and_then(|size| size.checked_add(connectors.checked_mul(2)?)).and_then(|size| size.checked_add(2));
    Ok((values.ok_or_else(ComputeError::memory)?, indices.ok_or_else(ComputeError::memory)?))
}

fn value(input: &Input<'_>, series: &ChartSeries, row: usize) -> Result<(f64, bool), ComputeError> {
    let index = series.column as usize * input.rows + row;
    if input.states[index] == 1 { Ok((finite(input.values[index], "Waterfall value")?, true)) }
    else { Ok((0.0, false)) }
}

fn is_total(input: &Input<'_>, column: Option<u32>, row: usize) -> Result<bool, ComputeError> {
    let Some(column) = column else { return Ok(false); };
    let index = column as usize * input.rows + row;
    match input.states[index] {
        0 => Ok(false),
        1 if input.values[index] == 0.0 => Ok(false),
        1 if input.values[index] == 1.0 => Ok(true),
        _ => Err(ComputeError::invalid("Waterfall total flags must be empty or exact numeric 0/1")),
    }
}

fn point(output: &mut ChartStatisticsResult, values: usize, indices: usize, record: usize,
    identity: [u32; 4], start: f64, end: f64, value: f64) -> Result<(), ComputeError> {
    finite(end - start, "Waterfall span")?;
    output.values[values + 1 + record * 3..values + 1 + record * 3 + 3].copy_from_slice(&[start, end, value]);
    output.indices[indices + 2 + record * 4..indices + 2 + record * 4 + 4].copy_from_slice(&identity);
    Ok(())
}

pub(super) fn layout(input: &Input<'_>, series: &[ChartSeries], selected: &[u32], stack: WaterfallStackType,
    subtotal: bool, total_columns: Option<&[Option<u32>]>, output: &mut ChartStatisticsResult) -> Result<JobSummary, ComputeError> {
    let (points, connectors) = record_counts(input.rows, selected.len(), stack, subtotal)?;
    let (value_count, index_count) = output_capacity(input.rows, selected.len(), stack, subtotal)?;
    let values = output.values.len();
    let indices = output.indices.len();
    // ChartStatistics::new reserved these regions under the same budget. No
    // view/lease is exposed until the task completes and ownership is consumed.
    output.values.resize(values + value_count, 0.0);
    output.indices.resize(indices + index_count, 0);
    output.indices[indices] = points as u32;
    output.indices[indices + 1] = connectors as u32;
    if input.rows == 0 { return Ok(JobSummary::default()); }
    let mut running = 0.0;
    let mut count = 0;
    match stack {
        WaterfallStackType::Arrangement => {
            let stride = input.rows + usize::from(subtotal);
            for (position, &index) in selected.iter().enumerate() {
                let source = &series[index as usize];
                let total_column = total_columns.and_then(|columns| columns[position]);
                for row in 0..input.rows {
                    let (value, numeric) = value(input, source, row)?;
                    count += usize::from(numeric);
                    let total = is_total(input, total_column, row)?;
                    let start = if total { 0.0 } else { running };
                    let end = if total { value } else { finite(running + value, "Waterfall running total")? };
                    running = end;
                    let record = position * stride + row;
                    point(output, values, indices, record, [source.series_position, row as u32, record as u32, u32::from(total)], start, end, value)?;
                }
            }
            // The fixed renderer uses the final cumulative for each inserted
            // subtotal. They are display records and do not reset the baseline.
            if subtotal {
                for (position, &index) in selected.iter().enumerate() {
                    let record = position * stride + input.rows;
                    point(output, values, indices, record, [series[index as usize].series_position, input.rows as u32, record as u32, TOTAL | SYNTHETIC], 0.0, running, running)?;
                }
            }
            for connector in 0..connectors {
                let from = connector / input.rows * stride + connector % input.rows;
                let next = connector + 1;
                let to = next / input.rows * stride + next % input.rows;
                let from_value = output.values[values + 1 + from * 3 + 1];
                let total = output.indices[indices + 2 + to * 4 + 3] & TOTAL != 0;
                let to_value = output.values[values + 1 + to * 3 + usize::from(total)];
                let value_offset = values + 1 + points * 3 + connector * 2;
                output.values[value_offset..value_offset + 2].copy_from_slice(&[from_value, to_value]);
                let index_offset = indices + 2 + points * 4 + connector * 2;
                output.indices[index_offset..index_offset + 2].copy_from_slice(&[from as u32, to as u32]);
            }
        }
        WaterfallStackType::Stacked => {
            for row in 0..input.rows {
                let mut cursor = running;
                for (position, &index) in selected.iter().enumerate() {
                    if is_total(input, total_columns.and_then(|columns| columns[position]), row)? {
                        return Err(ComputeError::invalid("Stacked waterfall cannot contain explicit total roles"));
                    }
                    let (value, _) = value(input, &series[index as usize], row)?;
                    if value < 0.0 { cursor = finite(cursor + value, "Waterfall stacked baseline")?; }
                }
                for (position, &index) in selected.iter().enumerate() {
                    let source = &series[index as usize];
                    let (value, numeric) = value(input, source, row)?;
                    count += usize::from(numeric);
                    let top = finite(cursor + value.abs(), "Waterfall stacked span")?;
                    let (start, end) = if value < 0.0 { (top, cursor) } else { (cursor, top) };
                    point(output, values, indices, row * selected.len() + position, [source.series_position, row as u32, row as u32, 0], start, end, value)?;
                    cursor = top;
                    running = finite(running + value, "Waterfall running total")?;
                }
                if row < connectors {
                    let value_offset = values + 1 + points * 3 + row * 2;
                    output.values[value_offset..value_offset + 2].copy_from_slice(&[running, running]);
                    let index_offset = indices + 2 + points * 4 + row * 2;
                    output.indices[index_offset..index_offset + 2].copy_from_slice(&[row as u32, row as u32 + 1]);
                }
            }
            if subtotal {
                point(output, values, indices, points - 1, [u32::MAX, input.rows as u32, input.rows as u32, TOTAL | SYNTHETIC], 0.0, running, running)?;
            }
        }
    }
    output.values[values] = running;
    Ok(JobSummary { count, available: true, ..Default::default() })
}

#[cfg(test)]
#[path = "waterfall_tests.rs"]
mod tests;
