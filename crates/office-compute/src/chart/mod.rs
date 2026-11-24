mod distribution;
mod hierarchy;
mod errors;
mod math;
mod trendline;
mod waterfall;
#[cfg(test)]
mod tests;

use crate::{ComputeError, MAX_ELEMENTS, MAX_ROWS};
use serde::{Deserialize, Serialize};
use std::collections::HashSet;
use math::reserve;

pub const MAX_CHART_COLUMNS: usize = 1024;
pub const MAX_CHART_SERIES: usize = 1024;
pub const MAX_CHART_JOBS: usize = 4096;
pub const CHART_RUNTIME_ALLOWANCE: usize = 8 * 1024 * 1024;
pub const CHART_SCRATCH_ALLOWANCE: usize = 128 * 1024;

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct ChartSeries {
    pub field: u32,
    pub series_position: u32,
    pub column: u32,
    pub x_column: Option<u32>,
}

#[derive(Debug, Clone, Copy, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum QuartileMethod { MedianOfHalves, Inclusive, Exclusive }

#[derive(Debug, Clone, Copy, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum WhiskerMethod { Range, Iqr }

// Packets written before the range mode was introduced used IQR whiskers.
impl Default for WhiskerMethod { fn default() -> Self { Self::Iqr } }

#[derive(Debug, Clone, Copy, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum TrendType { Exp, Linear, Log, MovingAvg, Poly, Pow }

#[derive(Debug, Clone, Copy, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum ErrorDirection { X, Y }

#[derive(Debug, Clone, Copy, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum ErrorSides { Both, Plus, Minus }

#[derive(Debug, Clone, Copy, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum ErrorValueType { FixedVal, Percentage, StdDev, StdErr, Cust }

#[derive(Debug, Clone, Copy, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum WaterfallStackType { Arrangement, Stacked }

#[derive(Debug, Clone, Deserialize)]
#[serde(tag = "mode", rename_all = "camelCase", deny_unknown_fields)]
pub enum HistogramBinning { Count { count: f64 }, Width { width: f64 } }

#[derive(Debug, Clone, Deserialize)]
#[serde(tag = "kind", rename_all = "camelCase", rename_all_fields = "camelCase", deny_unknown_fields)]
pub enum ChartJob {
    Histogram { series_index: u32, slot: u32, binning: Option<HistogramBinning>,
        underflow_threshold: Option<f64>, overflow_threshold: Option<f64> },
    Boxplot { series_index: u32, slot: u32, quartile_method: QuartileMethod,
        #[serde(default)] whisker_method: WhiskerMethod, whisker_multiplier: f64 },
    Pareto { series_index: u32, slot: u32, include_zero_values: bool },
    Hierarchy { series_index: u32, slot: u32, hierarchy_columns: Vec<u32>, label_columns: Vec<u32>, key_count: u32, label_count: u32 },
    /// Total/change flags share the columnar numeric input, avoiding a large
    /// per-point control DTO. Only exact 0/1 flags are accepted.
    Waterfall { series_index: u32, slot: u32, total_column: Option<u32> },
    WaterfallLayout { series_index: u32, slot: u32, series_indices: Vec<u32>, stack_type: WaterfallStackType,
        use_subtotal: bool, total_columns: Option<Vec<Option<u32>>> },
    Trendline { series_index: u32, slot: u32, #[serde(rename = "type")] trend_type: TrendType,
        order: Option<u32>, period: Option<u32> },
    ErrorBars { series_index: u32, slot: u32, direction: ErrorDirection,
        #[serde(rename = "type")] sides: ErrorSides, value_type: ErrorValueType,
        value: Option<f64>, plus_column: Option<u32>, minus_column: Option<u32> },
}

impl ChartJob {
    pub fn key(&self) -> (&'static str, u32, u32) {
        match self {
            Self::Histogram { series_index, slot, .. } => ("histogram", *series_index, *slot),
            Self::Boxplot { series_index, slot, .. } => ("boxplot", *series_index, *slot),
            Self::Pareto { series_index, slot, .. } => ("pareto", *series_index, *slot),
            Self::Hierarchy { series_index, slot, .. } => ("hierarchy", *series_index, *slot),
            Self::Waterfall { series_index, slot, .. } => ("waterfall", *series_index, *slot),
            Self::WaterfallLayout { series_index, slot, .. } => ("waterfallLayout", *series_index, *slot),
            Self::Trendline { series_index, slot, .. } => ("trendline", *series_index, *slot),
            Self::ErrorBars { series_index, slot, .. } => ("errorBars", *series_index, *slot),
        }
    }

    pub fn output_capacity(&self, rows: usize) -> Result<(usize, usize), ComputeError> {
        if rows > MAX_ROWS { return Err(ComputeError::memory()); }
        Ok(match self {
            Self::Histogram { .. } => (distribution::MAX_HISTOGRAM_BINS * 2, distribution::MAX_HISTOGRAM_BINS * 2),
            Self::Boxplot { .. } => (rows + 5, rows),
            Self::Pareto { .. } => (rows * 2, rows),
            Self::Hierarchy { hierarchy_columns, .. } => {
                let nodes = rows.checked_mul(hierarchy_columns.len()).ok_or_else(ComputeError::memory)?;
                (nodes.checked_mul(2).ok_or_else(ComputeError::memory)?,
                    nodes.checked_mul(4).and_then(|size| size.checked_add(rows + 2)).ok_or_else(ComputeError::memory)?)
            }
            Self::Waterfall { .. } => (rows * 3, rows * 2),
            Self::WaterfallLayout { series_indices, stack_type, use_subtotal, .. } => {
                return waterfall::output_capacity(rows, series_indices.len(), *stack_type, *use_subtotal);
            }
            Self::Trendline { trend_type: TrendType::MovingAvg, .. } => (rows * 2, 0),
            Self::Trendline { .. } => (33 * 2, 0),
            Self::ErrorBars { .. } => (rows * 4, rows),
        })
    }

    fn scratch_capacity(&self, rows: usize) -> usize {
        match self {
            Self::Histogram { .. } => rows * 8 + (distribution::MAX_HISTOGRAM_BINS + 1) * 12,
            Self::Boxplot { .. } => rows * 16,
            Self::Pareto { .. } => rows * 16,
            Self::Hierarchy { hierarchy_columns, .. } => rows * hierarchy_columns.len() * 256 + (hierarchy_columns.len() + 1) * 64,
            // Each of the two stacks reserves at most period nodes; a node is
            // <=56 bytes natively and <=48 on wasm32.
            Self::Trendline { trend_type: TrendType::MovingAvg, period, .. } =>
                rows.min(period.unwrap_or(2) as usize) * 112,
            Self::Trendline { .. } => rows * 16 + 4096,
            Self::Waterfall { .. } | Self::WaterfallLayout { .. } | Self::ErrorBars { .. } => 0,
        }
    }
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct ChartStatisticsParams {
    pub row_count: u32,
    pub column_count: u32,
    pub category_count: u32,
    pub series: Vec<ChartSeries>,
    pub jobs: Vec<ChartJob>,
}

pub struct ChartMemoryEstimate {
    pub input_bytes: usize,
    pub output_bytes: usize,
    pub estimated_peak_bytes: usize,
    pub value_capacity: usize,
    pub index_capacity: usize,
}

impl ChartStatisticsParams {
    pub fn validate_shape(&self, values: usize, states: usize, categories: usize) -> Result<(), ComputeError> {
        let rows = self.row_count as usize;
        let columns = self.column_count as usize;
        if rows > MAX_ROWS || columns > MAX_CHART_COLUMNS || self.series.len() > MAX_CHART_SERIES
            || self.jobs.len() > MAX_CHART_JOBS {
            return Err(ComputeError::memory());
        }
        let length = rows.checked_mul(columns).ok_or_else(ComputeError::memory)?;
        if length > MAX_ELEMENTS { return Err(ComputeError::memory()); }
        if columns == 0 || values != length || states != length || categories != rows
            || self.category_count > self.row_count || (rows == 0 && self.category_count != 0)
            || (rows > 0 && self.category_count == 0) {
            return Err(ComputeError::invalid("Invalid chart column/row/category dimensions"));
        }
        let mut positions = HashSet::new();
        for series in &self.series {
            if series.field > 1_048_575 || series.series_position as usize >= MAX_CHART_SERIES || !positions.insert(series.series_position)
                || series.column >= self.column_count || series.x_column.is_some_and(|column| column >= self.column_count) {
                return Err(ComputeError::invalid("Chart series positions must be unique; fields/columns must be in range"));
            }
        }
        let mut slots = HashSet::new();
        for job in &self.jobs {
            let (kind, series, slot) = job.key();
            if series as usize >= self.series.len() || slot as usize >= MAX_CHART_JOBS
                || !slots.insert((kind, series, slot)) {
                return Err(ComputeError::invalid("Chart job series/slot is missing or duplicated"));
            }
            match job {
                ChartJob::Histogram { binning, underflow_threshold, overflow_threshold, .. } => {
                    match binning {
                        Some(HistogramBinning::Count { count }) if !count.is_finite() =>
                            return Err(ComputeError::invalid("Histogram count must be finite")),
                        Some(HistogramBinning::Width { width }) if !width.is_finite() =>
                            return Err(ComputeError::invalid("Histogram width must be finite")),
                        _ => {}
                    }
                    if underflow_threshold.is_some_and(|value| !value.is_finite())
                        || overflow_threshold.is_some_and(|value| !value.is_finite()) {
                        return Err(ComputeError::invalid("Histogram thresholds must be finite"));
                    }
                }
                ChartJob::Boxplot { whisker_multiplier, .. } => {
                    if !whisker_multiplier.is_finite() || !(0.0..=100.0).contains(whisker_multiplier) {
                        return Err(ComputeError::invalid("Boxplot whisker multiplier must be 0..100"));
                    }
                }
                ChartJob::Hierarchy { hierarchy_columns, label_columns, key_count, label_count, .. } => {
                    let mut selected = HashSet::new();
                    if hierarchy_columns.is_empty() || hierarchy_columns.len() > MAX_CHART_COLUMNS || hierarchy_columns.len() != label_columns.len()
                        || hierarchy_columns.iter().chain(label_columns.iter()).chain(std::iter::once(&self.series[series as usize].column))
                            .any(|column| *column >= self.column_count || !selected.insert(*column))
                        || *key_count as usize > rows * hierarchy_columns.len() || *label_count as usize > rows * hierarchy_columns.len() {
                        return Err(ComputeError::invalid("Hierarchy key/label/value columns are missing, duplicated or out of range"));
                    }
                }
                ChartJob::Waterfall { total_column, .. } => {
                    if total_column.is_some_and(|column| column >= self.column_count) {
                        return Err(ComputeError::invalid("Waterfall total flag column is out of range"));
                    }
                }
                ChartJob::WaterfallLayout { series_indices, total_columns, .. } => {
                    let mut selected = HashSet::new();
                    if series_indices.is_empty() || series_indices.len() > MAX_CHART_SERIES || series_indices.first() != Some(&series)
                        || series_indices.iter().any(|index| *index as usize >= self.series.len() || !selected.insert(*index))
                        || total_columns.as_ref().is_some_and(|columns| columns.len() != series_indices.len()
                            || columns.iter().any(|column| column.is_some_and(|column| column >= self.column_count))) {
                        return Err(ComputeError::invalid("Waterfall layout series/total columns are missing, duplicated or out of range"));
                    }
                }
                ChartJob::Trendline { order, period, .. } => {
                    if order.is_some_and(|value| !(2..=6).contains(&value))
                        || period.is_some_and(|value| !(2..=100_000).contains(&value)) {
                        return Err(ComputeError::invalid("Trendline order/period is out of range"));
                    }
                }
                ChartJob::ErrorBars { direction, sides, value_type, value, plus_column, minus_column, .. } => {
                    if value.is_some_and(|number| !number.is_finite() || number < 0.0) {
                        return Err(ComputeError::invalid("Error value must be nonnegative and finite"));
                    }
                    if matches!(direction, ErrorDirection::X) && self.series[series as usize].x_column.is_none() {
                        return Err(ComputeError::invalid("X error bars require a numeric X column"));
                    }
                    if matches!(value_type, ErrorValueType::Cust) {
                        let selected = [
                            (!matches!(sides, ErrorSides::Minus), plus_column),
                            (!matches!(sides, ErrorSides::Plus), minus_column),
                        ];
                        for (enabled, column) in selected {
                            if enabled && column.is_none_or(|column| column >= self.column_count) {
                                return Err(ComputeError::invalid("Enabled custom error columns must be present and in range"));
                            }
                        }
                    }
                }
                ChartJob::Pareto { .. } => {}
            }
        }
        Ok(())
    }

    /// JS inputs + WASM retained inputs + WASM/JS output regions + largest
    /// job scratch + bounded control/runtime allowances. This matches tasks.ts.
    pub fn memory_estimate(&self) -> Result<ChartMemoryEstimate, ComputeError> {
        let rows = self.row_count as usize;
        if rows > MAX_ROWS || self.column_count as usize > MAX_CHART_COLUMNS
            || self.series.len() > MAX_CHART_SERIES || self.jobs.len() > MAX_CHART_JOBS {
            return Err(ComputeError::memory());
        }
        let elements = rows.checked_mul(self.column_count as usize).ok_or_else(ComputeError::memory)?;
        self.validate_shape(elements, elements, rows)?;
        let input_bytes = rows.checked_mul(self.column_count as usize).and_then(|value| value.checked_mul(9))
            .and_then(|value| value.checked_add(rows.checked_mul(4)?)).ok_or_else(ComputeError::memory)?;
        let mut value_capacity: usize = 0;
        let mut index_capacity: usize = 0;
        for job in &self.jobs {
            let (values, indices) = job.output_capacity(rows)?;
            value_capacity = value_capacity.checked_add(values).ok_or_else(ComputeError::memory)?;
            index_capacity = index_capacity.checked_add(indices).ok_or_else(ComputeError::memory)?;
        }
        let output_bytes = value_capacity.checked_mul(8)
            .and_then(|value| value.checked_add(index_capacity.checked_mul(4)?)).ok_or_else(ComputeError::memory)?;
        let scratch = self.jobs.iter().map(|job| job.scratch_capacity(rows)).max().unwrap_or(0)
            .checked_add(CHART_SCRATCH_ALLOWANCE).ok_or_else(ComputeError::memory)?;
        let controls = self.series.len().checked_mul(32)
            .and_then(|value| value.checked_add(self.jobs.len().checked_mul(512)?)).ok_or_else(ComputeError::memory)?;
        let controls = self.jobs.iter().try_fold(controls, |total, job| {
            let extra = match job {
                ChartJob::WaterfallLayout { series_indices, total_columns, .. } => series_indices.len().checked_mul(4)
                    .and_then(|size| size.checked_add(total_columns.as_ref().map_or(0, |columns| columns.len()).checked_mul(8)?)),
                ChartJob::Hierarchy { hierarchy_columns, .. } => hierarchy_columns.len().checked_mul(8),
                _ => Some(0),
            }.ok_or_else(ComputeError::memory)?;
            total.checked_add(extra).ok_or_else(ComputeError::memory)
        })?;
        let estimated_peak_bytes = input_bytes.checked_mul(2)
            .and_then(|value| value.checked_add(output_bytes.checked_mul(2)?))
            .and_then(|value| value.checked_add(scratch))
            .and_then(|value| value.checked_add(controls))
            .and_then(|value| value.checked_add(CHART_RUNTIME_ALLOWANCE)).ok_or_else(ComputeError::memory)?;
        Ok(ChartMemoryEstimate { input_bytes, output_bytes, estimated_peak_bytes, value_capacity, index_capacity })
    }
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ChartTrendlineFit {
    pub coefficients: Vec<f64>,
    pub x_scale: f64,
    pub x_mean: f64,
    pub x_spread: f64,
    pub y_scale: f64,
    pub y_mean: f64,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ChartSegment {
    pub kind: &'static str,
    pub job_index: usize,
    pub field: u32,
    pub series_position: u32,
    pub slot: u32,
    pub value_offset: usize,
    pub value_length: usize,
    pub index_offset: usize,
    pub index_length: usize,
    pub valid_sample_count: usize,
    pub available: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fit: Option<ChartTrendlineFit>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub r2: Option<f64>,
}

pub struct ChartStatisticsResult {
    pub values: Vec<f64>,
    pub indices: Vec<u32>,
    pub segments: Vec<ChartSegment>,
}

#[derive(Default)]
pub(super) struct JobSummary {
    count: usize,
    available: bool,
    fit: Option<ChartTrendlineFit>,
    r2: Option<f64>,
}

pub(super) struct Input<'a> {
    values: &'a [f64],
    states: &'a [u8],
    categories: &'a [u32],
    rows: usize,
    category_count: usize,
}

impl Input<'_> {
    fn sample(&self, column: u32, row: usize) -> Option<f64> {
        let index = column as usize * self.rows + row;
        (self.states[index] == 1 && self.values[index].is_finite()).then_some(self.values[index])
    }

    fn samples(&self, column: u32) -> impl Iterator<Item = f64> + Clone + '_ {
        (0..self.rows).filter_map(move |row| self.sample(column, row))
    }

    fn pair(&self, series: &ChartSeries, row: usize, ordinal_base: usize) -> Option<(f64, f64)> {
        let y = self.sample(series.column, row)?;
        let x = match series.x_column {
            Some(column) => self.sample(column, row)?,
            None => (row + ordinal_base) as f64,
        };
        Some((x, y))
    }
}

/// Whole jobs are the progress unit. A single job may include a synchronous
/// sort/fit; the Worker/client kill-and-recreate path bounds cancellation there.
pub struct ChartStatistics {
    values: Vec<f64>,
    states: Vec<u8>,
    categories: Vec<u32>,
    params: ChartStatisticsParams,
    result: Option<ChartStatisticsResult>,
    position: usize,
    failed: bool,
}

impl ChartStatistics {
    pub fn new(values: Vec<f64>, states: Vec<u8>, categories: Vec<u32>,
        params: ChartStatisticsParams, memory_budget: usize) -> Result<Self, ComputeError> {
        params.validate_shape(values.len(), states.len(), categories.len())?;
        crate::error::validate_states(&states)?;
        if categories.iter().any(|&category| category >= params.category_count) {
            return Err(ComputeError::invalid("Chart category dictionary index is out of range"));
        }
        let memory = params.memory_estimate()?;
        if memory.estimated_peak_bytes > memory_budget { return Err(ComputeError::memory()); }
        let mut output = Vec::new();
        let mut indices = Vec::new();
        let mut segments = Vec::new();
        reserve(&mut output, memory.value_capacity)?;
        reserve(&mut indices, memory.index_capacity)?;
        reserve(&mut segments, params.jobs.len())?;
        Ok(Self { values, states, categories, params, result: Some(ChartStatisticsResult {
            values: output, indices, segments }), position: 0, failed: false })
    }

    pub fn step(&mut self, jobs: u32) -> Result<bool, ComputeError> {
        if self.failed || self.result.is_none() {
            return Err(ComputeError::new("COMPUTE_RESULT_CONSUMED", "Chart task was consumed or failed"));
        }
        let end = (self.position + jobs.clamp(1, 8) as usize).min(self.params.jobs.len());
        while self.position < end {
            let job = &self.params.jobs[self.position];
            let (kind, series_index, slot) = job.key();
            let series = &self.params.series[series_index as usize];
            let input = Input { values: &self.values, states: &self.states, categories: &self.categories,
                rows: self.params.row_count as usize, category_count: self.params.category_count as usize };
            let result = self.result.as_mut().expect("result owner checked above");
            let value_offset = result.values.len();
            let index_offset = result.indices.len();
            let summary = match job {
                ChartJob::Histogram { binning, underflow_threshold, overflow_threshold, .. } =>
                    distribution::histogram(&input, series, binning.as_ref(), *underflow_threshold, *overflow_threshold, result),
                ChartJob::Boxplot { quartile_method, whisker_method, whisker_multiplier, .. } =>
                    distribution::boxplot(&input, series, *quartile_method, *whisker_method, *whisker_multiplier, result),
                ChartJob::Pareto { include_zero_values, .. } =>
                    distribution::pareto(&input, series, *include_zero_values, result),
                ChartJob::Hierarchy { hierarchy_columns, label_columns, key_count, label_count, .. } =>
                    hierarchy::build(&input, series, hierarchy_columns, label_columns, *key_count, *label_count, result),
                ChartJob::Waterfall { total_column, .. } =>
                    distribution::waterfall(&input, series, *total_column, result),
                ChartJob::WaterfallLayout { series_indices, stack_type, use_subtotal, total_columns, .. } =>
                    waterfall::layout(&input, &self.params.series, series_indices, *stack_type, *use_subtotal, total_columns.as_deref(), result),
                ChartJob::Trendline { trend_type, order, period, .. } =>
                    trendline::trendline(&input, series, *trend_type, *order, *period, result),
                ChartJob::ErrorBars { direction, sides, value_type, value, plus_column, minus_column, .. } =>
                    errors::error_bars(&input, series, *direction, *sides, *value_type,
                        value.unwrap_or(1.0), *plus_column, *minus_column, result),
            };
            let summary = match summary {
                Ok(summary) => summary,
                Err(error) => { self.failed = true; return Err(error); }
            };
            result.segments.push(ChartSegment { kind, job_index: self.position, field: series.field, series_position: series.series_position, slot,
                value_offset, value_length: result.values.len() - value_offset,
                index_offset, index_length: result.indices.len() - index_offset,
                valid_sample_count: summary.count, available: summary.available, fit: summary.fit, r2: summary.r2 });
            self.position += 1;
        }
        Ok(self.position == self.params.jobs.len())
    }

    pub fn progress(&self) -> f64 {
        if self.params.jobs.is_empty() { 1.0 } else { self.position as f64 / self.params.jobs.len() as f64 }
    }

    pub fn take_result(&mut self) -> Result<ChartStatisticsResult, ComputeError> {
        if self.failed { return Err(ComputeError::new("COMPUTE_RESULT_CONSUMED", "Chart task failed")); }
        if self.position != self.params.jobs.len() {
            return Err(ComputeError::new("COMPUTE_NOT_READY", "Chart statistics have not finished"));
        }
        self.result.take().ok_or_else(|| ComputeError::new("COMPUTE_RESULT_CONSUMED", "Chart result was already consumed"))
    }
}
