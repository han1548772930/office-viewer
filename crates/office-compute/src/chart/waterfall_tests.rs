//! Draft multi-series/role/ownership contracts; run in the unified test stage.

use super::*;
use super::super::{ChartJob, ChartStatistics, ChartStatisticsParams};

fn params(stack_type: WaterfallStackType, use_subtotal: bool, total_columns: Option<Vec<Option<u32>>>) -> ChartStatisticsParams {
    ChartStatisticsParams { row_count: 2, column_count: 4, category_count: 2,
        series: vec![ChartSeries { field: 1, series_position: 7, column: 0, x_column: None },
            ChartSeries { field: 2, series_position: 2, column: 1, x_column: None }],
        jobs: vec![ChartJob::WaterfallLayout { series_index: 1, slot: 0, series_indices: vec![1, 0], stack_type, use_subtotal, total_columns }] }
}

fn complete(values: Vec<f64>, states: Vec<u8>, params: ChartStatisticsParams) -> ChartStatisticsResult {
    let mut task = ChartStatistics::new(values, states, vec![0, 1], params, 256 * 1024 * 1024).unwrap();
    assert!(task.step(1).unwrap());
    let result = task.take_result().unwrap();
    assert_eq!(task.take_result().err().unwrap().code, "COMPUTE_RESULT_CONSUMED");
    result
}

#[test]
fn arrangement_shares_cumulative_resets_totals_and_skips_synthetic_connector_slots() {
    let params = params(WaterfallStackType::Arrangement, true, Some(vec![None, Some(3)]));
    let result = complete(vec![10.0, -3.0, 5.0, 2.0, 0.0, 0.0, 1.0, 0.0], vec![1; 8], params);
    assert_eq!(result.values, [7.0, 0.0,5.0,5.0, 5.0,7.0,2.0, 0.0,7.0,7.0,
        0.0,10.0,10.0, 10.0,7.0,-3.0, 0.0,7.0,7.0, 5.0,5.0, 7.0,10.0, 10.0,10.0]);
    assert_eq!(result.indices, [6,3, 2,0,0,0, 2,1,1,0, 2,2,2,3, 7,0,3,1, 7,1,4,0, 7,2,5,3, 0,1, 1,3, 3,4]);
    assert_eq!(result.segments[0].kind, "waterfallLayout");
    assert_eq!(result.segments[0].field, 2);
    assert_eq!(result.segments[0].series_position, 2);
    assert_eq!(result.segments[0].valid_sample_count, 4);
}

#[test]
fn stacked_orders_signed_spans_by_category_and_emits_one_global_subtotal() {
    let mut params = params(WaterfallStackType::Stacked, true, None);
    if let ChartJob::WaterfallLayout { series_index, series_indices, .. } = &mut params.jobs[0] { *series_index = 0; *series_indices = vec![0, 1]; }
    let result = complete(vec![5.0,-4.0, -2.0,3.0, 0.0,0.0, 0.0,0.0], vec![1;8], params);
    assert_eq!(result.values, [2.0, -2.0,3.0,5.0, 5.0,3.0,-2.0, 3.0,-1.0,-4.0, 3.0,6.0,3.0, 0.0,2.0,2.0, 3.0,3.0]);
    assert_eq!(result.indices, [5,1, 7,0,0,0, 2,0,0,0, 7,1,1,0, 2,1,1,0, u32::MAX,2,2,3, 0,1]);
}

#[test]
fn omitting_subtotals_preserves_source_point_identity_and_cross_series_cumulative() {
    let result = complete(vec![10.0,-3.0,5.0,2.0,0.0,0.0,0.0,0.0], vec![1;8], params(WaterfallStackType::Arrangement, false, None));
    assert_eq!(result.indices[..2], [4,3]);
    assert_eq!(result.values[0], 14.0);
    assert_eq!(result.values[1 + 2 * 3..1 + 2 * 3 + 3], [7.0,17.0,10.0]);
    assert!((0..4).all(|index| result.indices[2 + index * 4 + 3] == 0));
}

#[test]
fn empty_values_and_empty_total_flags_are_changes_with_zero_numeric_samples() {
    let result = complete(vec![0.0;8], vec![0;8], params(WaterfallStackType::Arrangement, true, Some(vec![Some(2),Some(3)])));
    assert_eq!(result.indices[..2], [6,3]);
    assert!(result.values.iter().all(|value| *value == 0.0));
    assert!(result.segments[0].available);
    assert_eq!(result.segments[0].valid_sample_count, 0);
}

#[test]
fn zero_rows_return_only_empty_headers_and_no_automatic_records() {
    let mut params = params(WaterfallStackType::Stacked, true, None);
    params.row_count = 0; params.category_count = 0;
    let mut task = ChartStatistics::new(vec![],vec![],vec![],params,256 * 1024 * 1024).unwrap();
    assert!(task.step(1).unwrap());
    let result = task.take_result().unwrap();
    assert_eq!(result.values, [0.0]);
    assert_eq!(result.indices, [0,0]);
    assert!(!result.segments[0].available);
}

#[test]
fn explicit_stacked_totals_invalid_flags_and_numeric_overflow_fail_atomically() {
    for (stack, values, expected) in [
        (WaterfallStackType::Stacked, vec![1.0,2.0,3.0,4.0,1.0,0.0,0.0,0.0], "COMPUTE_INVALID_INPUT"),
        (WaterfallStackType::Arrangement, vec![1.0,2.0,3.0,4.0,2.0,0.0,0.0,0.0], "COMPUTE_INVALID_INPUT"),
        (WaterfallStackType::Arrangement, vec![f64::MAX,f64::MAX,0.0,0.0,0.0,0.0,0.0,0.0], "COMPUTE_NUMERIC_RANGE"),
    ] {
        let mut task = ChartStatistics::new(values, vec![1;8], vec![0,1], params(stack,true,Some(vec![Some(2),None])),256 * 1024 * 1024).unwrap();
        assert_eq!(task.step(1).unwrap_err().code, expected);
        assert_eq!(task.take_result().err().unwrap().code, "COMPUTE_RESULT_CONSUMED");
    }
}

#[test]
fn malformed_series_identity_and_total_column_shapes_are_rejected_before_allocation() {
    for patch in 0..4 {
        let mut params = params(WaterfallStackType::Arrangement,true,None);
        if let ChartJob::WaterfallLayout { series_index, series_indices, total_columns, .. } = &mut params.jobs[0] {
            match patch { 0 => *series_indices = vec![1,1], 1 => *series_index = 0, 2 => *series_indices = vec![1,2], _ => *total_columns = Some(vec![None]) }
        }
        assert_eq!(params.memory_estimate().err().unwrap().code, "COMPUTE_INVALID_INPUT");
    }
    assert!(record_counts(MAX_ROWS,MAX_CHART_SERIES,WaterfallStackType::Arrangement,true).is_err());
}

#[test]
fn admission_counts_both_output_buffers_and_variable_control_arrays() {
    let params = params(WaterfallStackType::Arrangement,true,Some(vec![None,Some(3)]));
    let memory = params.memory_estimate().unwrap();
    assert_eq!((memory.value_capacity,memory.index_capacity), (25,32));
    assert_eq!(memory.output_bytes, 25 * 8 + 32 * 4);
    assert_eq!(memory.estimated_peak_bytes, memory.input_bytes * 2 + memory.output_bytes * 2
        + super::super::CHART_SCRATCH_ALLOWANCE + super::super::CHART_RUNTIME_ALLOWANCE + 2 * 32 + 512 + 2 * 4 + 2 * 8);
}
