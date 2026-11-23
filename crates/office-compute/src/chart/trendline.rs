use super::{ChartSeries, ChartStatisticsResult, ChartTrendlineFit, Input, JobSummary, TrendType};
use super::math::{CompensatedSum, finite, interpolate, moments, reserve};
use crate::ComputeError;

fn transformed_x(value: f64, kind: TrendType) -> f64 {
    if matches!(kind, TrendType::Log | TrendType::Pow) { value.ln() } else { value }
}

fn transformed_y(value: f64, kind: TrendType) -> f64 {
    if matches!(kind, TrendType::Exp | TrendType::Pow) { value.ln() } else { value }
}

pub(super) fn trendline(input: &Input<'_>, series: &ChartSeries, kind: TrendType,
    order: Option<u32>, period: Option<u32>, output: &mut ChartStatisticsResult) -> Result<JobSummary, ComputeError> {
    if matches!(kind, TrendType::MovingAvg) { return moving_average(input, series, period.unwrap_or(2) as usize, output); }
    let mut pairs = Vec::new();
    reserve(&mut pairs, input.rows)?;
    for row in 0..input.rows {
        let Some((x, y)) = input.pair(series, row, 1) else { continue; };
        if matches!(kind, TrendType::Log | TrendType::Pow) && x <= 0.0
            || matches!(kind, TrendType::Exp | TrendType::Pow) && y <= 0.0 { continue; }
        pairs.push((x, y));
    }
    let count = pairs.len();
    let unavailable = || JobSummary { count, ..Default::default() };
    let degree = if matches!(kind, TrendType::Poly) { order.unwrap_or(2) as usize } else { 1 };
    if count <= degree { return Ok(unavailable()); }
    let xm = moments(pairs.iter().map(|&(x, _)| transformed_x(x, kind)))?;
    let ym = moments(pairs.iter().map(|&(_, y)| transformed_y(y, kind)))?;
    let spread = pairs.iter().map(|&(x, _)| (transformed_x(x, kind) / xm.scale - xm.normalized_mean).abs())
        .fold(0.0_f64, f64::max);
    if spread == 0.0 { return Ok(unavailable()); }
    let Some(coefficients) = polynomial_fit(&pairs, kind, xm.scale, xm.normalized_mean, spread,
        ym.scale, ym.normalized_mean, degree) else { return Ok(unavailable()); };
    let predict_scaled = |x: f64| {
        let argument = (transformed_x(x, kind) / xm.scale - xm.normalized_mean) / spread;
        ym.normalized_mean + coefficients.iter().rev().fold(0.0, |sum, coefficient| sum * argument + coefficient)
    };
    let exponential = matches!(kind, TrendType::Exp | TrendType::Pow);
    let min_x = pairs.iter().map(|&(x, _)| x).fold(f64::INFINITY, f64::min);
    let max_x = pairs.iter().map(|&(x, _)| x).fold(f64::NEG_INFINITY, f64::max);
    for index in 0..33 {
        let x = interpolate(min_x, max_x, index as f64 / 32.0)?;
        let normalized = predict_scaled(x);
        let y = if exponential { (normalized * ym.scale).exp() } else { normalized * ym.scale };
        if y.is_finite() { output.values.extend_from_slice(&[x, y]); }
    }
    let original = if exponential { moments(pairs.iter().map(|&(_, y)| y))? } else { ym };
    let relative_prediction = |x: f64| if exponential {
        (predict_scaled(x) * ym.scale - original.scale.ln()).exp()
    } else { predict_scaled(x) };
    let r2 = r_squared(&pairs, original.scale, relative_prediction)?;
    Ok(JobSummary { count, available: true, r2,
        fit: Some(ChartTrendlineFit { coefficients, x_scale: xm.scale, x_mean: xm.normalized_mean,
            x_spread: spread, y_scale: ym.scale, y_mean: ym.normalized_mean }) })
}

fn polynomial_fit(pairs: &[(f64, f64)], kind: TrendType, x_scale: f64, x_mean: f64,
    spread: f64, y_scale: f64, y_mean: f64, degree: usize) -> Option<Vec<f64>> {
    let count = degree + 1;
    let mut powers = [CompensatedSum::default(); 13];
    let mut targets = [CompensatedSum::default(); 7];
    for &(x, y) in pairs {
        let x = (transformed_x(x, kind) / x_scale - x_mean) / spread;
        let y = transformed_y(y, kind) / y_scale - y_mean;
        let mut power = 1.0;
        for index in 0..=degree * 2 {
            powers[index].add(power);
            if index < count { targets[index].add(power * y); }
            power *= x;
        }
    }
    let mut matrix = [[0.0; 8]; 7];
    for row in 0..count {
        for column in 0..count { matrix[row][column] = powers[row + column].value() / pairs.len() as f64; }
        matrix[row][count] = targets[row].value() / pairs.len() as f64;
    }
    for column in 0..count {
        let mut pivot = column;
        for row in column + 1..count {
            if matrix[row][column].abs() > matrix[pivot][column].abs() { pivot = row; }
        }
        if matrix[pivot][column].abs() < 1e-12 { return None; }
        matrix.swap(column, pivot);
        let divisor = matrix[column][column];
        for entry in column..=count { matrix[column][entry] /= divisor; }
        for row in 0..count {
            if row == column { continue; }
            let multiplier = matrix[row][column];
            for entry in column..=count { matrix[row][entry] -= multiplier * matrix[column][entry]; }
        }
    }
    let coefficients: Vec<f64> = (0..count).map(|row| matrix[row][count]).collect();
    coefficients.iter().all(|value| value.is_finite()).then_some(coefficients)
}

fn r_squared(pairs: &[(f64, f64)], scale: f64, predict: impl Fn(f64) -> f64) -> Result<Option<f64>, ComputeError> {
    let mean = moments(pairs.iter().map(|&(_, y)| y / scale))?.mean;
    let mut total_scale = 0.0_f64;
    let mut residual_scale = 0.0_f64;
    for &(x, y) in pairs {
        let prediction = predict(x);
        let actual = y / scale;
        let residual = actual - prediction;
        if !prediction.is_finite() || !residual.is_finite() { return Ok(None); }
        total_scale = total_scale.max((actual - mean).abs());
        residual_scale = residual_scale.max(residual.abs());
    }
    if total_scale == 0.0 { return Ok(Some(if residual_scale <= 32.0 * f64::EPSILON { 1.0 } else { 0.0 })); }
    if residual_scale == 0.0 { return Ok(Some(1.0)); }
    let mut total = CompensatedSum::default();
    let mut residual = CompensatedSum::default();
    for &(x, y) in pairs {
        let actual = y / scale;
        total.add(((actual - mean) / total_scale).powi(2));
        residual.add(((actual - predict(x)) / residual_scale).powi(2));
    }
    let ratio = (2.0 * (residual_scale.ln() - total_scale.ln()) + residual.value().ln() - total.value().ln()).exp();
    let r2 = 1.0 - ratio;
    Ok(r2.is_finite().then_some(r2))
}

#[derive(Clone, Copy, Default)]
struct WindowAggregate { scale: f64, total: f64, correction: f64, count: usize, invalid: usize }

impl WindowAggregate {
    fn single(value: f64, valid: bool) -> Self {
        let scale = value.abs();
        Self { scale, total: if scale == 0.0 { 0.0 } else { value / scale }, correction: 0.0,
            count: 1, invalid: usize::from(!valid) }
    }

    fn merge(first: Self, second: Self) -> Self {
        let scale = first.scale.max(second.scale);
        let mut result = Self { scale, count: first.count + second.count, invalid: first.invalid + second.invalid,
            ..Default::default() };
        if scale == 0.0 { return result; }
        for value in [first.total * (first.scale / scale), first.correction * (first.scale / scale),
            second.total * (second.scale / scale), second.correction * (second.scale / scale)] {
            let next = result.total + value;
            result.correction += if result.total.abs() >= value.abs() {
                (result.total - next) + value
            } else { (value - next) + result.total };
            result.total = next;
        }
        result
    }

    fn mean(self) -> Result<f64, ComputeError> {
        finite(((self.total + self.correction) / self.count as f64).clamp(-1.0, 1.0) * self.scale, "Moving average")
    }
}

#[derive(Clone, Copy)]
struct WindowNode { value: f64, valid: bool, aggregate: WindowAggregate }

/// The two stacks maintain scaled compensated aggregates at every prefix.
/// Removing a large value reveals a freshly scaled suffix; a small surviving
/// window never depends on a global scale that can underflow its values.
struct MeanWindow { left: Vec<WindowNode>, right: Vec<WindowNode> }

impl MeanWindow {
    fn new(capacity: usize) -> Result<Self, ComputeError> {
        let mut left = Vec::new();
        let mut right = Vec::new();
        reserve(&mut left, capacity)?;
        reserve(&mut right, capacity)?;
        Ok(Self { left, right })
    }

    fn len(&self) -> usize { self.left.len() + self.right.len() }

    fn push(&mut self, value: f64, valid: bool) {
        let aggregate = WindowAggregate::merge(self.right.last().map(|node| node.aggregate).unwrap_or_default(),
            WindowAggregate::single(value, valid));
        self.right.push(WindowNode { value, valid, aggregate });
    }

    fn pop(&mut self) {
        if self.left.is_empty() {
            while let Some(node) = self.right.pop() {
                let aggregate = WindowAggregate::merge(WindowAggregate::single(node.value, node.valid),
                    self.left.last().map(|node| node.aggregate).unwrap_or_default());
                self.left.push(WindowNode { aggregate, ..node });
            }
        }
        self.left.pop();
    }

    fn aggregate(&self) -> WindowAggregate {
        WindowAggregate::merge(self.left.last().map(|node| node.aggregate).unwrap_or_default(),
            self.right.last().map(|node| node.aggregate).unwrap_or_default())
    }
}

fn moving_average(input: &Input<'_>, series: &ChartSeries, period: usize,
    output: &mut ChartStatisticsResult) -> Result<JobSummary, ComputeError> {
    if input.rows < period {
        return Ok(JobSummary { count: (0..input.rows).filter(|&row| input.pair(series, row, 1).is_some()).count(),
            ..Default::default() });
    }
    let mut window = MeanWindow::new(period)?;
    let mut count = 0;
    for row in 0..input.rows {
        let pair = input.pair(series, row, 1);
        if pair.is_some() { count += 1; }
        if window.len() == period { window.pop(); }
        window.push(pair.map(|(_, y)| y).unwrap_or(0.0), pair.is_some());
        let aggregate = window.aggregate();
        if aggregate.count == period && aggregate.invalid == 0 {
            let x = pair.expect("valid window includes current point").0;
            output.values.extend_from_slice(&[x, aggregate.mean()?]);
        }
    }
    Ok(JobSummary { count, available: true, ..Default::default() })
}
