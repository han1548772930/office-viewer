use super::*;

fn run(values: Vec<f64>, rows: u32, columns: u32, x: Option<u32>, categories: Vec<u32>,
    category_count: u32, jobs: Vec<ChartJob>) -> ChartStatisticsResult {
    let params = ChartStatisticsParams { row_count: rows, column_count: columns, category_count,
        series: vec![ChartSeries { field: 1, series_position: 0, column: 0, x_column: x }], jobs };
    let states = vec![1; values.len()];
    let mut task = ChartStatistics::new(values, states, categories, params, usize::MAX).unwrap();
    while !task.step(1).unwrap() {}
    let result = task.take_result().unwrap();
    assert_eq!(task.take_result().err().unwrap().code, "COMPUTE_RESULT_CONSUMED");
    result
}

fn close(actual: f64, expected: f64) {
    assert!((actual - expected).abs() <= 1e-10 * expected.abs().max(1.0), "{actual} != {expected}");
}

#[test]
fn histogram_thresholds_and_right_closed_intervals() {
    let result = run(vec![0.0, 1.0, 2.0, 3.0, 4.0], 5, 1, None, vec![0, 1, 2, 3, 4], 5,
        vec![ChartJob::Histogram { series_index: 0, slot: 0, binning: Some(HistogramBinning::Count { count: 2 }),
            underflow_threshold: Some(1.0), overflow_threshold: Some(3.0) }]);
    assert_eq!(result.values, vec![0.0, 1.0, 1.0, 2.0, 2.0, 3.0, 3.0, 0.0]);
    assert_eq!(result.indices, vec![0, 2, 1, 1, 1, 1, 2, 1]);
    assert_eq!(result.segments[0].valid_sample_count, 5);
}

#[test]
fn constant_histogram_retains_authored_count_and_finite_extreme_bounds() {
    let result = run(vec![7.0; 3], 3, 1, None, vec![0, 1, 2], 3,
        vec![ChartJob::Histogram { series_index: 0, slot: 0, binning: Some(HistogramBinning::Count { count: 3 }),
            underflow_threshold: None, overflow_threshold: None }]);
    assert_eq!(result.indices, vec![1, 3, 1, 0, 1, 0]);
    assert_eq!(result.values[0], 7.0);
    assert_eq!(result.values[5], 8.0);
    let result = run(vec![f64::MAX], 1, 1, None, vec![0], 1,
        vec![ChartJob::Histogram { series_index: 0, slot: 0, binning: None,
            underflow_threshold: None, overflow_threshold: None }]);
    assert!(result.values[0] < result.values[1]);
    assert_eq!(result.values[1], f64::MAX);
}

#[test]
fn boxplot_quartile_methods_and_outliers() {
    let result = run(vec![1.0, 2.0, 3.0, 4.0], 4, 1, None, vec![0, 1, 2, 3], 4,
        vec![ChartJob::Boxplot { series_index: 0, slot: 0, quartile_method: QuartileMethod::Exclusive, whisker_multiplier: 1.5 },
            ChartJob::Boxplot { series_index: 0, slot: 1, quartile_method: QuartileMethod::Inclusive, whisker_multiplier: 0.0 }]);
    assert_eq!(&result.values[..5], &[1.0, 1.25, 2.5, 3.75, 4.0]);
    assert_eq!(&result.values[5..], &[2.0, 1.75, 2.5, 3.25, 3.0, 1.0, 4.0]);
}

#[test]
fn pareto_groups_duplicate_categories_and_orders_ties_by_first_source_row() {
    let result = run(vec![2.0, 3.0, 1.0, 0.0], 4, 1, None, vec![0, 1, 0, 2], 3,
        vec![ChartJob::Pareto { series_index: 0, slot: 0, include_zero_values: true }]);
    assert_eq!(result.indices, vec![0, 1, 2]);
    assert_eq!(result.values, vec![3.0, 50.0, 3.0, 100.0, 0.0, 100.0]);
}

#[test]
fn waterfall_total_role_resets_the_running_baseline() {
    let result = run(vec![5.0, -2.0, 10.0, 3.0, 0.0, 0.0, 1.0, 0.0], 4, 2, None, vec![0, 1, 2, 3], 4,
        vec![ChartJob::Waterfall { series_index: 0, slot: 0, total_column: Some(1) }]);
    assert_eq!(result.values, vec![0.0, 5.0, 5.0, 5.0, 3.0, -2.0, 0.0, 10.0, 10.0, 10.0, 13.0, 3.0]);
    assert_eq!(result.indices, vec![0, 0, 1, 0, 2, 1, 3, 0]);
}

#[test]
fn standard_error_uses_only_complete_xy_pairs() {
    let result = run(vec![2.0, f64::MAX, 4.0, 0.0, f64::NAN, 2.0], 3, 2, Some(1), vec![0, 1, 2], 3,
        vec![ChartJob::ErrorBars { series_index: 0, slot: 0, direction: ErrorDirection::Y, sides: ErrorSides::Both,
            value_type: ErrorValueType::StdErr, value: Some(1.0), plus_column: None, minus_column: None }]);
    assert_eq!(result.indices, vec![0, 2]);
    assert_eq!(result.values, vec![0.0, 2.0, 1.0, 1.0, 2.0, 4.0, 1.0, 1.0]);
    assert_eq!(result.segments[0].valid_sample_count, 2);
}

#[test]
fn unused_error_directions_and_zero_factors_do_not_compute_an_overflowing_deviation() {
    let result = run(vec![-f64::MAX, f64::MAX], 2, 1, None, vec![0, 1], 2,
        vec![ChartJob::ErrorBars { series_index: 0, slot: 0, direction: ErrorDirection::Y, sides: ErrorSides::Plus,
            value_type: ErrorValueType::StdDev, value: Some(0.0), plus_column: None, minus_column: None }]);
    assert_eq!(result.values[2], 0.0);
    assert_eq!(result.values[3], 0.0);
    assert_eq!(result.values[6], 0.0);
    assert_eq!(result.values[7], 0.0);
    let result = run(vec![f64::MAX], 1, 1, None, vec![0], 1,
        vec![ChartJob::ErrorBars { series_index: 0, slot: 0, direction: ErrorDirection::Y, sides: ErrorSides::Minus,
            value_type: ErrorValueType::Percentage, value: Some(1.0), plus_column: None, minus_column: None }]);
    assert!(result.values[2].is_finite());
    assert_eq!(result.values[3], 0.0);
}

#[test]
fn custom_error_missing_values_omit_intervals_and_disabled_columns_are_unused() {
    let result = run(vec![2.0, 3.0, 1.0, f64::NAN], 2, 2, None, vec![0, 1], 2,
        vec![ChartJob::ErrorBars { series_index: 0, slot: 0, direction: ErrorDirection::Y, sides: ErrorSides::Plus,
            value_type: ErrorValueType::Cust, value: None, plus_column: Some(1), minus_column: Some(999) }]);
    assert_eq!(result.indices, vec![0]);
    assert_eq!(result.values, vec![0.0, 2.0, 0.0, 1.0]);
}

#[test]
fn linear_and_polynomial_fits_return_normalized_metadata_and_original_y_r_squared() {
    let result = run(vec![2.0, 4.0, 6.0], 3, 1, None, vec![0, 1, 2], 3,
        vec![ChartJob::Trendline { series_index: 0, slot: 3, trend_type: TrendType::Linear, order: None, period: None }]);
    close(result.values[1], 2.0);
    close(*result.values.last().unwrap(), 6.0);
    close(result.segments[0].r2.unwrap(), 1.0);
    assert_eq!(result.segments[0].slot, 3);
    assert_eq!(result.segments[0].fit.as_ref().unwrap().coefficients.len(), 2);
    let result = run(vec![0.0, 1.0, 4.0, 9.0, 0.0, 1.0, 2.0, 3.0], 4, 2, Some(1), vec![0, 1, 2, 3], 4,
        vec![ChartJob::Trendline { series_index: 0, slot: 0, trend_type: TrendType::Poly, order: Some(2), period: None }]);
    close(result.values[1], 0.0);
    close(*result.values.last().unwrap(), 9.0);
}

#[test]
fn moving_average_recovers_tiny_windows_after_large_values_leave() {
    let result = run(vec![f64::MAX, f64::MAX, 1e-200, 1e-200], 4, 1, None, vec![0, 1, 2, 3], 4,
        vec![ChartJob::Trendline { series_index: 0, slot: 0, trend_type: TrendType::MovingAvg, order: None, period: Some(2) }]);
    assert_eq!(result.values[1], f64::MAX);
    assert_eq!(*result.values.last().unwrap(), 1e-200);
    let result = run(vec![1.0, 2.0, 3.0, 1.0, f64::NAN, 3.0], 3, 2, Some(1), vec![0, 1, 2], 3,
        vec![ChartJob::Trendline { series_index: 0, slot: 0, trend_type: TrendType::MovingAvg, order: None, period: Some(2) }]);
    assert!(result.values.is_empty());
}

#[test]
fn partial_failure_and_budget_exhaustion_never_publish_a_prefix() {
    let params = ChartStatisticsParams { row_count: 1, column_count: 1, category_count: 1,
        series: vec![ChartSeries { field: 1, series_position: 0, column: 0, x_column: None }],
        jobs: vec![ChartJob::Pareto { series_index: 0, slot: 0, include_zero_values: false }] };
    let memory = params.memory_estimate().unwrap();
    assert_eq!(ChartStatistics::new(vec![1.0], vec![1], vec![0], params.clone(),
        memory.estimated_peak_bytes - 1).err().unwrap().code, "COMPUTE_MEMORY_LIMIT");
    let mut task = ChartStatistics::new(vec![-1.0], vec![1], vec![0], params, usize::MAX).unwrap();
    assert_eq!(task.step(1).unwrap_err().code, "COMPUTE_INVALID_INPUT");
    assert_eq!(task.take_result().err().unwrap().code, "COMPUTE_RESULT_CONSUMED");
}

#[test]
fn finite_product_rounds_the_final_product_without_intermediate_overflow() {
    let value = math::product(&[f64::MAX, 100.0, 0.001], "test").unwrap();
    assert!(value.is_finite());
    close(value / f64::MAX, 0.1);
    assert_eq!(math::product(&[f64::MAX, f64::MAX, 0.0], "test").unwrap(), 0.0);
}

#[test]
fn repeated_source_y_fields_use_distinct_x_samples_and_keep_the_shared_authored_slot() {
    let params = ChartStatisticsParams { row_count: 3, column_count: 3, category_count: 3,
        series: vec![
            ChartSeries { field: 7, series_position: 0, column: 0, x_column: Some(1) },
            ChartSeries { field: 7, series_position: 1, column: 0, x_column: Some(2) },
        ], jobs: vec![
            ChartJob::Trendline { series_index: 0, slot: 0, trend_type: TrendType::Linear, order: None, period: None },
            ChartJob::Trendline { series_index: 1, slot: 0, trend_type: TrendType::Linear, order: None, period: None },
        ] };
    let mut task = ChartStatistics::new(vec![2.0, 4.0, 6.0, 1.0, 2.0, 3.0, 10.0, 20.0, 30.0],
        vec![1; 9], vec![0, 1, 2], params.clone(), usize::MAX).unwrap();
    while !task.step(1).unwrap() {}
    let result = task.take_result().unwrap();
    assert_eq!(result.segments.iter().map(|segment| (segment.field, segment.series_position, segment.slot)).collect::<Vec<_>>(),
        vec![(7, 0, 0), (7, 1, 0)]);
    for (segment, first_x, last_x) in [(&result.segments[0], 1.0, 3.0), (&result.segments[1], 10.0, 30.0)] {
        let start = segment.value_offset as usize;
        let last = start + segment.value_length as usize - 2;
        close(result.values[start], first_x); close(result.values[last], last_x);
        close(result.values[start + 1], 2.0); close(result.values[last + 1], 6.0);
    }
    let mut duplicated_position = params.clone();
    duplicated_position.series[1].series_position = 0;
    assert_eq!(duplicated_position.validate_shape(9, 9, 3).unwrap_err().code, "COMPUTE_INVALID_INPUT");
    let mut duplicated_job = params;
    duplicated_job.jobs.push(duplicated_job.jobs[0].clone());
    assert_eq!(duplicated_job.validate_shape(9, 9, 3).unwrap_err().code, "COMPUTE_INVALID_INPUT");
}
