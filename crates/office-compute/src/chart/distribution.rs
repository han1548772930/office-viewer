use super::{ChartSeries, ChartStatisticsResult, HistogramBinning, Input, JobSummary, QuartileMethod, WhiskerMethod};
use super::math::{finite, interpolate, moments, reserve};
use crate::ComputeError;

pub(super) const MAX_HISTOGRAM_BINS: usize = 1000;

pub(super) fn histogram(input: &Input<'_>, series: &ChartSeries, binning: Option<&HistogramBinning>,
    underflow: Option<f64>, overflow: Option<f64>, output: &mut ChartStatisticsResult) -> Result<JobSummary, ComputeError> {
    // Ineffective finite author values are not rewritten. The semantic runtime
    // ignores an inverted/equal threshold pair and falls back to Scott Auto.
    let (underflow, overflow) = if underflow.zip(overflow).is_some_and(|(lower, upper)| lower >= upper) {
        (None, None)
    } else { (underflow, overflow) };
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
    let has_regular = !regular.is_empty();
    if let Some(threshold) = underflow.filter(|_| has_regular || under_count > 0) {
        output.values.extend_from_slice(&[0.0, threshold]);
        output.indices.extend_from_slice(&[0, under_count]);
    }
    if has_regular {
        let mut min = regular[0];
        let mut max = min;
        for &value in &regular { min = min.min(value); max = max.max(value); }
        let edge_count = usize::from(underflow.is_some()) + usize::from(overflow.is_some());
        let capacity = MAX_HISTOGRAM_BINS - edge_count;
        let mut boundaries = None;
        if let Some(HistogramBinning::Width { width }) = binning {
            if width.is_finite() && *width > 0.0 {
                let lower = underflow.unwrap_or(min);
                let upper = overflow.unwrap_or(if min == max { min + *width } else { max });
                let scale = lower.abs().max(upper.abs()).max(*width);
                let count = ((upper / scale - lower / scale) / (*width / scale)).ceil().max(1.0);
                if upper.is_finite() && upper > lower && count.is_finite() && count <= capacity as f64 {
                    boundaries = histogram_boundaries(lower, upper, count as usize, Some(*width))?;
                }
            }
        }
        let boundaries = match boundaries {
            Some(boundaries) => boundaries,
            None if min == max => vec![min, max],
            None => {
                let requested = match binning {
                    Some(HistogramBinning::Count { count }) if count.is_finite() && count.fract() == 0.0
                        && *count >= 1.0 && *count <= MAX_HISTOGRAM_BINS as f64 && *count > edge_count as f64 =>
                        Some(*count as usize - edge_count),
                    _ => None,
                };
                let count = match requested {
                    Some(count) => count,
                    None => histogram_auto_count(&regular, min, max, capacity)?,
                };
                histogram_boundaries(min, max, count, None)?.ok_or_else(||
                    ComputeError::invalid("Histogram boundaries are not strictly increasing at binary64 precision"))?
            }
        };
        let bin_count = boundaries.len() - 1;
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
    if let Some(threshold) = overflow.filter(|_| has_regular || over_count > 0) {
        output.values.extend_from_slice(&[threshold, 0.0]);
        output.indices.extend_from_slice(&[2, over_count]);
    }
    Ok(JobSummary { count, available: true, ..Default::default() })
}

fn histogram_auto_count(samples: &[f64], min: f64, max: f64, capacity: usize) -> Result<usize, ComputeError> {
    let stats = moments(samples.iter().copied())?;
    let scale = min.abs().max(max.abs()).max(stats.scale);
    let range = max / scale - min / scale;
    let width = 3.5 * (stats.normalized_variance * (samples.len() as f64 / (samples.len() - 1) as f64)).sqrt()
        * (stats.scale / scale) / (samples.len() as f64).cbrt();
    Ok(if width.is_finite() && width > 0.0 { (range / width).ceil().max(1.0).min(capacity as f64) as usize } else { 1 })
}

fn histogram_boundaries(lower: f64, upper: f64, count: usize, width: Option<f64>) -> Result<Option<Vec<f64>>, ComputeError> {
    let mut boundaries = Vec::new();
    reserve(&mut boundaries, count + 1)?;
    boundaries.push(lower);
    let scale = lower.abs().max(upper.abs()).max(width.unwrap_or(0.0));
    for index in 1..=count {
        let boundary = if index == count { upper } else {
            match width {
                Some(width) => ((lower / scale + index as f64 * (width / scale)) * scale).min(upper),
                None => interpolate(lower, upper, index as f64 / count as f64)?,
            }
        };
        if !boundary.is_finite() || boundary <= *boundaries.last().expect("initial lower bound") { return Ok(None); }
        boundaries.push(boundary);
    }
    Ok(Some(boundaries))
}

pub(super) fn boxplot(input: &Input<'_>, series: &ChartSeries, method: QuartileMethod,
    whiskers: WhiskerMethod, multiplier: f64, output: &mut ChartStatisticsResult) -> Result<JobSummary, ComputeError> {
    let mut sorted = Vec::new();
    reserve(&mut sorted, input.rows)?;
    sorted.extend((0..input.rows).filter_map(|row| input.sample(series.column, row).map(|value| (value, row as u32))));
    if sorted.is_empty() { return Ok(JobSummary::default()); }
    sorted.sort_unstable_by(|a, b| a.0.partial_cmp(&b.0).expect("only finite samples").then(a.1.cmp(&b.1)));
    if matches!(method, QuartileMethod::MedianOfHalves) && sorted.len() < 2 {
        return Ok(JobSummary { count: sorted.len(), ..Default::default() });
    }
    let half_median = |start: usize, end: usize| -> Result<f64, ComputeError> {
        let length = end - start;
        let middle = start + length / 2;
        if length % 2 == 1 { Ok(sorted[middle].0) } else { interpolate(sorted[middle - 1].0, sorted[middle].0, 0.5) }
    };
    let quartile = |fraction: f64| -> Result<f64, ComputeError> {
        if matches!(method, QuartileMethod::MedianOfHalves) {
            return if fraction == 0.25 { half_median(0, sorted.len() / 2) }
                else if fraction == 0.75 { half_median((sorted.len() + 1) / 2, sorted.len()) }
                else { half_median(0, sorted.len()) };
        }
        let position = match method {
            QuartileMethod::Exclusive => fraction * (sorted.len() + 1) as f64 - 1.0,
            QuartileMethod::Inclusive => fraction * (sorted.len() - 1) as f64,
            QuartileMethod::MedianOfHalves => unreachable!("handled median-of-halves above"),
        };
        let lower = position.floor().clamp(0.0, (sorted.len() - 1) as f64) as usize;
        let upper = position.ceil().clamp(0.0, (sorted.len() - 1) as f64) as usize;
        interpolate(sorted[lower].0, sorted[upper].0, (position - position.floor()).clamp(0.0, 1.0))
    };
    let q1 = quartile(0.25)?;
    let median = quartile(0.5)?;
    let q3 = quartile(0.75)?;
    if matches!(whiskers, WhiskerMethod::Range) {
        output.values.extend_from_slice(&[sorted[0].0, q1, median, q3, sorted[sorted.len() - 1].0]);
        return Ok(JobSummary { count: sorted.len(), available: true, ..Default::default() });
    }
    let scale = q1.abs().max(q3.abs()).max(sorted[0].0.abs()).max(sorted[sorted.len() - 1].0.abs());
    let scale = if scale == 0.0 { 1.0 } else { scale };
    let distance = (q3 / scale - q1 / scale) * multiplier;
    let within = |value: f64| value / scale >= q1 / scale - distance && value / scale <= q3 / scale + distance;
    let min = sorted.iter().find(|sample| within(sample.0)).map(|sample| sample.0).unwrap_or(q1);
    let max = sorted.iter().rev().find(|sample| within(sample.0)).map(|sample| sample.0).unwrap_or(q3);
    output.values.extend_from_slice(&[min, q1, median, q3, max]);
    for &(value, source_index) in sorted.iter().filter(|sample| !within(sample.0)) {
        output.values.push(value);
        output.indices.push(source_index);
    }
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
        count += 1;
        let group = &mut groups[input.categories[row] as usize];
        group.first = group.first.min(row as u32);
        group.value = finite(group.value + value, "Pareto category sum")?;
    }
    groups.retain(|group| group.first != u32::MAX && if include_zero { group.value >= 0.0 } else { group.value > 0.0 });
    groups.sort_unstable_by(|a, b| b.value.partial_cmp(&a.value).expect("finite group sum").then(a.first.cmp(&b.first)));
    let scale = groups.first().filter(|group| group.value != 0.0).map(|group| group.value).unwrap_or(1.0);
    // Preserve the fixed converter's binary64 rounding when the raw sum is
    // finite. Scaling only handles overflow; scaling every input changes 6/10.
    let raw_total: f64 = groups.iter().map(|group| group.value).sum();
    let scaled = !raw_total.is_finite();
    let total = if scaled { groups.iter().map(|group| group.value / scale).sum() } else { raw_total };
    let mut cumulative = 0.0;
    for (index, group) in groups.iter().enumerate() {
        cumulative += if scaled { group.value / scale } else { group.value };
        let percent = if total == 0.0 { 0.0 } else if index + 1 == groups.len() { 100.0 }
            else { ((cumulative / total * 10_000.0).floor() / 100.0).min(100.0) };
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
