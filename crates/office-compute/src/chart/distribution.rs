use super::{ChartSeries, ChartStatisticsResult, HistogramBinning, Input, JobSummary, QuartileMethod};
use super::math::{finite, interpolate, moments, reserve};
use crate::ComputeError;

pub(super) fn histogram(input: &Input<'_>, series: &ChartSeries, binning: Option<&HistogramBinning>,
    underflow: Option<f64>, overflow: Option<f64>, output: &mut ChartStatisticsResult) -> Result<JobSummary, ComputeError> {
    let mut regular = Vec::new();
    reserve(&mut regular, input.rows)?;
    let mut count: usize = 0;
    let mut under_count: u32 = 0;
    let mut over_count: u32 = 0;
    for value in input.samples(series.column) {
        count += 1;
        if underflow.is_some_and(|threshold| value <= threshold) { under_count += 1; }
        else if overflow.is_some_and(|threshold| value > threshold) { over_count += 1; }
        else { regular.push(value); }
    }
    if count == 0 { return Ok(JobSummary::default()); }
    if let Some(threshold) = underflow {
        output.values.extend_from_slice(&[0.0, threshold]);
        output.indices.extend_from_slice(&[0, under_count]);
    }
    if !regular.is_empty() {
        let mut min = regular[0];
        let mut max = min;
        for &value in &regular { min = min.min(value); max = max.max(value); }
        let mut lower = underflow.unwrap_or(min);
        let mut upper = overflow.unwrap_or(max);
        if lower == upper {
            let padding = match binning {
                Some(HistogramBinning::Width { width }) => *width,
                _ => 1.0_f64.max(lower.abs() * 0.1),
            };
            let after = upper + padding;
            let before = lower - padding;
            if overflow.is_none() && after.is_finite() && after > upper { upper = after; }
            else if before.is_finite() && before < lower { lower = before; }
            else { return Err(ComputeError::invalid("Histogram width cannot form a representable interval")); }
        }
        let stats = moments(regular.iter().copied())?;
        let width = match binning { Some(HistogramBinning::Width { width }) => *width, _ => 0.0 };
        let scale = lower.abs().max(upper.abs()).max(stats.scale).max(width);
        let range = upper / scale - lower / scale;
        let scott = 3.5 * stats.normalized_variance.sqrt() * (stats.scale / scale)
            / (regular.len() as f64).cbrt();
        let default_count = if range == 0.0 { 1.0 } else {
            (range / if scott > 0.0 { scott } else { range / (regular.len() as f64).sqrt() }).ceil().max(1.0)
        };
        let bin_count = match binning {
            Some(HistogramBinning::Count { count }) => *count as f64,
            Some(HistogramBinning::Width { width }) => (range / (*width / scale)).ceil().max(1.0),
            None if min == max => 1.0,
            None => default_count,
        };
        if !bin_count.is_finite() || !(1.0..=10_000.0).contains(&bin_count) {
            return Err(ComputeError::invalid("Histogram count must be 1..10000"));
        }
        let bin_count = bin_count as usize;
        let mut boundaries = Vec::new();
        reserve(&mut boundaries, bin_count + 1)?;
        for index in 0..=bin_count {
            let boundary = if index == 0 { lower } else {
                match binning {
                    Some(HistogramBinning::Width { .. }) if index == bin_count && overflow.is_some() => upper,
                    Some(HistogramBinning::Width { width }) => finite(
                        (lower / scale + index as f64 * (*width / scale)) * scale, "Histogram boundary")?,
                    _ => interpolate(lower, upper, index as f64 / bin_count as f64)?,
                }
            };
            if boundaries.last().is_some_and(|previous| *previous >= boundary) {
                return Err(ComputeError::invalid("Histogram boundaries are not strictly increasing at binary64 precision"));
            }
            boundaries.push(boundary);
        }
        let mut counts = Vec::new();
        reserve(&mut counts, bin_count)?;
        counts.resize(bin_count, 0_u32);
        for value in regular {
            let mut first = 0;
            let mut last = bin_count - 1;
            while first < last {
                let middle = (first + last) / 2;
                if value <= boundaries[middle + 1] { last = middle; } else { first = middle + 1; }
            }
            counts[first] += 1;
        }
        for index in 0..bin_count {
            output.values.extend_from_slice(&[boundaries[index], boundaries[index + 1]]);
            output.indices.extend_from_slice(&[1, counts[index]]);
        }
    }
    if let Some(threshold) = overflow {
        output.values.extend_from_slice(&[threshold, 0.0]);
        output.indices.extend_from_slice(&[2, over_count]);
    }
    Ok(JobSummary { count, available: true, ..Default::default() })
}

pub(super) fn boxplot(input: &Input<'_>, series: &ChartSeries, method: QuartileMethod,
    multiplier: f64, output: &mut ChartStatisticsResult) -> Result<JobSummary, ComputeError> {
    let mut sorted = Vec::new();
    reserve(&mut sorted, input.rows)?;
    sorted.extend(input.samples(series.column));
    if sorted.is_empty() { return Ok(JobSummary::default()); }
    sorted.sort_unstable_by(|a, b| a.partial_cmp(b).expect("only finite samples"));
    let quartile = |fraction: f64| -> Result<f64, ComputeError> {
        let position = match method {
            QuartileMethod::Exclusive => fraction * (sorted.len() + 1) as f64 - 1.0,
            QuartileMethod::Inclusive => fraction * (sorted.len() - 1) as f64,
        };
        let lower = position.floor().clamp(0.0, (sorted.len() - 1) as f64) as usize;
        let upper = position.ceil().clamp(0.0, (sorted.len() - 1) as f64) as usize;
        interpolate(sorted[lower], sorted[upper], (position - position.floor()).clamp(0.0, 1.0))
    };
    let q1 = quartile(0.25)?;
    let median = quartile(0.5)?;
    let q3 = quartile(0.75)?;
    let scale = q1.abs().max(q3.abs()).max(sorted[0].abs()).max(sorted[sorted.len() - 1].abs());
    let scale = if scale == 0.0 { 1.0 } else { scale };
    let distance = (q3 / scale - q1 / scale) * multiplier;
    let within = |value: f64| value / scale >= q1 / scale - distance && value / scale <= q3 / scale + distance;
    let min = sorted.iter().copied().find(|&value| within(value)).unwrap_or(q1);
    let max = sorted.iter().rev().copied().find(|&value| within(value)).unwrap_or(q3);
    output.values.extend_from_slice(&[min, q1, median, q3, max]);
    output.values.extend(sorted.iter().copied().filter(|&value| !within(value)));
    Ok(JobSummary { count: sorted.len(), available: true, ..Default::default() })
}

pub(super) fn pareto(input: &Input<'_>, series: &ChartSeries, include_zero: bool,
    output: &mut ChartStatisticsResult) -> Result<JobSummary, ComputeError> {
    struct Group { category: u32, first: u32, value: f64 }
    let mut groups = Vec::new();
    reserve(&mut groups, input.category_count)?;
    groups.extend((0..input.category_count).map(|category| Group { category: category as u32, first: u32::MAX, value: 0.0 }));
    let mut count = 0;
    for row in 0..input.rows {
        let Some(value) = input.sample(series.column, row) else { continue; };
        if value < 0.0 { return Err(ComputeError::invalid("Pareto weights cannot be negative")); }
        count += 1;
        let group = &mut groups[input.categories[row] as usize];
        group.first = group.first.min(row as u32);
        group.value = finite(group.value + value, "Pareto category sum")?;
    }
    groups.retain(|group| group.first != u32::MAX && (include_zero || group.value > 0.0));
    groups.sort_unstable_by(|a, b| b.value.partial_cmp(&a.value).expect("finite group sum").then(a.first.cmp(&b.first)));
    let scale = groups.first().filter(|group| group.value != 0.0).map(|group| group.value).unwrap_or(1.0);
    let total: f64 = groups.iter().map(|group| group.value / scale).sum();
    let mut cumulative = 0.0;
    for (index, group) in groups.iter().enumerate() {
        cumulative += group.value / scale;
        let percent = if total == 0.0 { 0.0 } else if index + 1 == groups.len() { 100.0 }
            else { (cumulative / total * 100.0).min(100.0) };
        output.values.extend_from_slice(&[group.value, percent]);
        output.indices.push(group.category);
    }
    Ok(JobSummary { count, available: !groups.is_empty(), ..Default::default() })
}

pub(super) fn waterfall(input: &Input<'_>, series: &ChartSeries, total_column: Option<u32>,
    output: &mut ChartStatisticsResult) -> Result<JobSummary, ComputeError> {
    let mut running = 0.0;
    let mut count = 0;
    for row in 0..input.rows {
        let index = series.column as usize * input.rows + row;
        let value = if input.states[index] == 1 {
            count += 1;
            finite(input.values[index], "Waterfall value")?
        } else { 0.0 };
        let is_total = match total_column {
            None => false,
            Some(column) => match input.sample(column, row) {
                Some(0.0) => false,
                Some(1.0) => true,
                _ => return Err(ComputeError::invalid("Waterfall total flags must be exact numeric 0/1")),
            },
        };
        let start = if is_total { 0.0 } else { running };
        let end = if is_total { value } else { finite(running + value, "Waterfall running total")? };
        running = end;
        output.values.extend_from_slice(&[start, end, value]);
        output.indices.extend_from_slice(&[row as u32, u32::from(is_total)]);
    }
    Ok(JobSummary { count, available: input.rows > 0, ..Default::default() })
}
