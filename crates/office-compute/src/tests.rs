use super::{ComputeError, NumericAggregate, NumericFilter, StableNumericSort};

#[test]
fn aggregate_tracks_numeric_empty_error_and_non_finite_inputs() {
    let mut aggregate = NumericAggregate::default();
    aggregate
        .push(&[1.0, 3.0, f64::NAN, 9.0, 0.0], &[1, 1, 1, 3, 0])
        .unwrap();

    let result = aggregate.statistics();
    assert_eq!(result.numeric_count, 2);
    assert_eq!(result.non_empty_count, 4);
    assert_eq!(result.empty_count, 1);
    assert_eq!(result.error_count, 1);
    assert_eq!(result.invalid_numeric_count, 1);
    assert_eq!(result.sum, Some(4.0));
    assert_eq!(result.mean, Some(2.0));
    assert_eq!(result.min, Some(1.0));
    assert_eq!(result.max, Some(3.0));
    assert_eq!(result.population_variance, Some(1.0));
    assert_eq!(result.sample_variance, Some(2.0));
    assert_eq!(result.sample_standard_deviation, Some(2.0_f64.sqrt()));
    assert!(!result.overflowed);
}

#[test]
fn aggregate_rejects_invalid_chunk_without_partially_mutating_state() {
    let mut aggregate = NumericAggregate::default();
    aggregate.push(&[2.0, 4.0], &[1, 1]).unwrap();
    let before = aggregate.statistics();

    assert_eq!(
        aggregate.push(&[10.0, 20.0], &[1, 9]),
        Err(ComputeError::invalid(
            "Cell states must be 0=empty, 1=number, 2=other, 3=error"
        ))
    );
    assert_eq!(aggregate.statistics(), before);
}

#[test]
fn empty_and_overflowed_statistics_keep_null_and_zero_distinct() {
    let empty = NumericAggregate::default().statistics();
    assert_eq!(empty.sum, Some(0.0));
    assert_eq!(empty.mean, None);
    assert_eq!(empty.population_variance, None);
    assert!(!empty.overflowed);
    let mut huge = NumericAggregate::default();
    huge.push(&[f64::MAX, f64::MAX], &[1, 1]).unwrap();
    let result = huge.statistics();
    assert_eq!(result.numeric_count, 2);
    assert_eq!(result.sum, None);
    assert_eq!(result.min, Some(f64::MAX));
    assert!(result.overflowed);
}

#[test]
fn stable_multikey_sort_is_incremental_and_keeps_equal_rows_in_source_order() {
    let rows = 6;
    let mut sort = StableNumericSort::new(
        vec![2.0, 1.0, 2.0, 1.0, 2.0, 1.0, 9.0, 3.0, 9.0, 3.0, 8.0, 2.0],
        vec![1; rows * 2],
        rows,
        vec![0, 0],
        true,
    )
    .unwrap();

    assert!(!sort.is_done());
    while !sort.step(1) {}
    assert_eq!(sort.take_result().unwrap(), vec![5, 1, 3, 4, 0, 2]);
}

#[test]
fn descending_sort_keeps_empty_and_error_cells_last_and_stable() {
    let mut sort = StableNumericSort::new(
        vec![1.0, 0.0, 3.0, 0.0, 2.0],
        vec![1, 0, 1, 3, 1],
        5,
        vec![1],
        true,
    )
    .unwrap();

    while !sort.step(2) {}
    assert_eq!(sort.take_result().unwrap(), vec![2, 4, 0, 1, 3]);
}

#[test]
fn numeric_filter_supports_all_any_and_error_empty_predicates() {
    let values = vec![1.0, 2.0, 3.0, 4.0, 0.0, 0.0, 0.0, 1.0, 1.0, 0.0];
    let states = vec![1, 1, 1, 1, 0, 1, 1, 1, 1, 3];
    let operators = vec![6, 4];
    let lower = vec![2.0, 0.0];
    let upper = vec![4.0, 0.0];

    let mut all = NumericFilter::new(
        values.clone(),
        states.clone(),
        5,
        operators.clone(),
        lower.clone(),
        upper.clone(),
        true,
    )
    .unwrap();
    while !all.step(1) {}
    assert_eq!(all.take_result().unwrap(), vec![2, 3]);

    let mut any =
        NumericFilter::new(values, states.clone(), 5, operators, lower, upper, false).unwrap();
    while !any.step(3) {}
    assert_eq!(any.take_result().unwrap(), vec![1, 2, 3]);

    let mut empty_or_error = NumericFilter::new(
        vec![0.0; 10],
        states,
        5,
        vec![8, 10],
        vec![0.0; 2],
        vec![0.0; 2],
        false,
    )
    .unwrap();
    while !empty_or_error.step(10) {}
    assert_eq!(empty_or_error.take_result().unwrap(), vec![4]);
}
