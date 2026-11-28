use super::*;

const MEMORY: usize = 256 * 1024 * 1024;
fn bytes(units: &[u16]) -> Vec<u8> { units.iter().flat_map(|unit| unit.to_le_bytes()).collect() }
fn input(labels: &[Vec<u16>], query: &[u16], start: u32, count: u32) -> (Vec<u32>, Vec<u8>, Vec<u8>, ValidationListQueryParams) {
    let mut offsets = vec![0];
    let mut text = Vec::new();
    for label in labels { text.extend(bytes(label)); offsets.push(text.len() as u32); }
    let query = bytes(query);
    let params = ValidationListQueryParams { row_count: labels.len() as u32, text_byte_length: text.len() as u32,
        query_byte_length: query.len() as u32, start, count };
    (offsets, text, query, params)
}
fn run(labels: &[Vec<u16>], query: &[u16], start: u32, count: u32, budget: u32) -> ValidationListQueryResult {
    let (offsets, text, query, params) = input(labels, query, start, count);
    let mut task = ValidationListQuery::new(offsets, text, query, params, MEMORY).unwrap();
    let mut previous = 0.0;
    for _ in 0..200_000 {
        let done = task.step(budget).unwrap();
        let progress = task.progress();
        assert!(progress >= previous && progress <= 1.0);
        previous = progress;
        if done { return task.take_result().unwrap(); }
    }
    panic!("List query did not finish within the small-fixture work bound");
}
fn labels(values: &[&str]) -> Vec<Vec<u16>> { values.iter().map(|value| value.encode_utf16().collect()).collect() }

#[test]
fn kmp_matches_an_independent_small_window_oracle_for_every_step_budget() {
    let source = labels(&["", "aaaaab", "xaaaaby", "abababab", "aa", "汉字\nemoji👩‍💻", "a b", " a "]);
    for query in ["", "a", "aaaaab", "ababab", "字\n", "👩‍💻", " ", "missing"] {
        let query: Vec<u16> = query.encode_utf16().collect();
        let all: Vec<u32> = source.iter().enumerate().filter(|(_, text)| query.is_empty()
            || text.windows(query.len()).any(|window| window == query)).map(|(index, _)| index as u32).collect();
        for (start, count) in [(0, 8), (1, 2), (6, 3), (u32::MAX, 1)] {
            for budget in [1, 2, 17, 4096] {
                let result = run(&source, &query, start, count, budget);
                assert_eq!(result.total, all.len() as u32);
                assert_eq!(result.indices, all.iter().skip(start as usize).take(count as usize).copied().collect::<Vec<_>>());
            }
        }
    }
}

#[test]
fn utf16_identity_preserves_lone_surrogates_and_never_matches_across_rows() {
    let source = vec![vec![0xd800, b'x' as u16], vec![0xdc00], vec![0xd83d, 0xde00], vec![b'a' as u16], vec![b'b' as u16]];
    assert_eq!(run(&source, &[0xd800], 0, 10, 1).indices, vec![0]);
    assert_eq!(run(&source, &[0xdc00], 0, 10, 1).indices, vec![1]);
    assert_eq!(run(&source, &[0xde00], 0, 10, 1).indices, vec![2]);
    assert_eq!(run(&source, &[b'a' as u16, b'b' as u16], 0, 10, 1).total, 0);
}

#[test]
fn offset_validation_is_chunked_and_rejects_odd_decreasing_and_out_of_bounds_offsets() {
    for offsets in [vec![0, 1, 4], vec![0, 6, 4], vec![0, 4, 2, 4]] {
        let params = ValidationListQueryParams { row_count: offsets.len() as u32 - 1, text_byte_length: 4,
            query_byte_length: 2, start: 0, count: 10 };
        let mut task = ValidationListQuery::new(offsets, vec![97, 0, 98, 0], vec![97, 0], params, MEMORY).unwrap();
        let error = loop { match task.step(1) { Err(error) => break error, Ok(false) => {}, Ok(true) => panic!("Invalid offsets accepted") } };
        assert_eq!(error.code, "COMPUTE_INVALID_INPUT");
    }
}

#[test]
fn empty_inputs_late_windows_and_independent_admission_counts_are_exact() {
    assert_eq!(run(&[], &[], 0, 1, 1), ValidationListQueryResult { indices: vec![], total: 0 });
    let (offsets, text, query, params) = input(&labels(&["alpha", "beta"]), &[b'a' as u16], 1, 1);
    let estimate = params.memory_estimate().unwrap();
    assert_eq!(estimate.input_bytes, 12 + 18 + 2); assert_eq!(estimate.output_bytes, 4);
    assert_eq!(estimate.estimated_peak_bytes, estimate.input_bytes * 2 + 4 + 8 + 8 * 1024 * 1024);
    assert!(ValidationListQuery::new(offsets, text, query, params, estimate.estimated_peak_bytes - 1).is_err());
    let params = ValidationListQueryParams { row_count: 2, text_byte_length: 18, query_byte_length: 2, start: u32::MAX, count: 10 };
    assert_eq!(params.memory_estimate().unwrap().output_bytes, 0);
}

#[test]
fn unfinished_zero_budget_and_double_consumption_cannot_expose_a_partial_result() {
    let (offsets, text, query, params) = input(&labels(&["alpha", "beta"]), &[b'a' as u16], 0, 2);
    let mut task = ValidationListQuery::new(offsets, text, query, params, MEMORY).unwrap();
    assert_eq!(task.take_result().unwrap_err().code, "COMPUTE_NOT_READY");
    assert_eq!(task.step(0).unwrap_err().code, "COMPUTE_INVALID_INPUT");
    while !task.step(1).unwrap() {}
    assert_eq!(task.take_result().unwrap().indices, vec![0, 1]);
    assert_eq!(task.take_result().unwrap_err().code, "COMPUTE_RESULT_CONSUMED");
    assert_eq!(task.step(1).unwrap_err().code, "COMPUTE_RESULT_CONSUMED");
}

#[test]
fn shape_checks_reject_odd_byte_lengths_and_excessive_row_query_or_window_limits() {
    let base = ValidationListQueryParams { row_count: 1, text_byte_length: 2, query_byte_length: 2, start: 0, count: 1 };
    assert!(base.validate_shape(2, 2, 2).is_ok());
    for params in [ValidationListQueryParams { text_byte_length: 1, ..base.clone() },
        ValidationListQueryParams { query_byte_length: 1, ..base.clone() }, ValidationListQueryParams { count: 0, ..base.clone() },
        ValidationListQueryParams { row_count: MAX_ROWS as u32 + 1, ..base.clone() },
        ValidationListQueryParams { query_byte_length: (MAX_ELEMENTS as u32 + 1) * 2, ..base.clone() }] {
        assert!(params.validate_shape(params.row_count as usize + 1, params.text_byte_length as usize, params.query_byte_length as usize).is_err());
    }
}
