use super::{ChartSeries, ChartStatisticsResult, ErrorDirection, ErrorSides, ErrorValueType, Input, JobSummary};
use super::math::{finite, moments, product};
use crate::ComputeError;

pub(super) fn error_bars(input: &Input<'_>, series: &ChartSeries, direction: ErrorDirection,
    sides: ErrorSides, value_type: ErrorValueType, factor: f64, plus_column: Option<u32>,
    minus_column: Option<u32>, output: &mut ChartStatisticsResult) -> Result<JobSummary, ComputeError> {
    let plus_enabled = !matches!(sides, ErrorSides::Minus);
    let minus_enabled = !matches!(sides, ErrorSides::Plus);
    let samples = (0..input.rows).filter_map(|row| input.pair(series, row, 0))
        .map(|(x, y)| if matches!(direction, ErrorDirection::X) { x } else { y });
    let count = samples.clone().count();
    if count == 0 { return Ok(JobSummary::default()); }
    let statistical_error = if matches!(value_type, ErrorValueType::StdDev | ErrorValueType::StdErr)
        && factor != 0.0 && count > 1 {
        let stats = moments(samples)?;
        let multiplier = if matches!(value_type, ErrorValueType::StdDev) {
            count as f64 / (count - 1) as f64
        } else { 1.0 / (count - 1) as f64 };
        product(&[stats.scale, (stats.normalized_variance * multiplier).sqrt(), factor], "Statistical error")?
    } else { 0.0 };
    for row in 0..input.rows {
        let Some((x, y)) = input.pair(series, row, 0) else { continue; };
        let target = if matches!(direction, ErrorDirection::X) { x } else { y };
        let error = match value_type {
            ErrorValueType::FixedVal => factor,
            ErrorValueType::Percentage => product(&[target.abs(), factor, 0.01], "Percentage error")?,
            _ => statistical_error,
        };
        let authored = |enabled: bool, column: Option<u32>| {
            if !enabled { Some(0.0) }
            else if matches!(value_type, ErrorValueType::Cust) { column.and_then(|column| input.sample(column, row)) }
            else { Some(error) }
        };
        let (Some(plus), Some(minus)) = (authored(plus_enabled, plus_column), authored(minus_enabled, minus_column))
            else { continue; };
        if plus < 0.0 || minus < 0.0 { return Err(ComputeError::invalid("Custom error values cannot be negative")); }
        if plus_enabled { finite(target + plus, "Error upper bound")?; }
        if minus_enabled { finite(target - minus, "Error lower bound")?; }
        output.values.extend_from_slice(&[x, y, minus, plus]);
        output.indices.push(row as u32);
    }
    Ok(JobSummary { count, available: true, ..Default::default() })
}
