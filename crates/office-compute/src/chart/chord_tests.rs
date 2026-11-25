use super::*;

// P28 analytical-fixture drafts. They are deliberately not executed during
// the implementation phase; browser buffer ownership is in wasm_chart.rs.
#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct ChordLayoutFixture { start_angle: f64, clockwise: bool, pad_angle: f64, min_angle: f64 }
#[derive(Deserialize)]
#[serde(untagged)]
enum Golden { Number(f64), Pi([i32; 2]) }
#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct ChordFixture {
    name: String, label_count: u32, layout: ChordLayoutFixture,
    rows: Vec<(Option<f64>, Option<u32>, Option<u32>)>, values: Vec<Golden>, indices: Vec<u32>,
    valid_sample_count: usize, available: bool,
}
#[derive(Deserialize)]
struct ChordFixtures { cases: Vec<ChordFixture> }

fn fixtures() -> Vec<ChordFixture> {
    serde_json::from_str::<ChordFixtures>(include_str!("../../../../tests/fixtures/chart-chord.json")).unwrap().cases
}
fn task(fixture: &ChordFixture, budget: usize) -> Result<ChartStatistics, ComputeError> {
    let rows = fixture.rows.len();
    let mut values = vec![0.0; rows * 3];
    let mut states = vec![0; values.len()];
    for (row, (weight, source, target)) in fixture.rows.iter().enumerate() {
        for (column, value) in [*weight, source.map(|value| value as f64), target.map(|value| value as f64)].into_iter().enumerate() {
            if let Some(value) = value { values[column * rows + row] = value; states[column * rows + row] = 1; }
        }
    }
    ChartStatistics::new(values, states, vec![0; rows], ChartStatisticsParams {
        row_count: rows as u32, column_count: 3, category_count: u32::from(rows > 0),
        series: vec![ChartSeries { field: 2, series_position: 0, column: 0, x_column: None }],
        jobs: vec![ChartJob::Chord { series_index: 0, slot: 0, source_column: 1, target_column: 2, label_count: fixture.label_count,
            start_angle: fixture.layout.start_angle, clockwise: fixture.layout.clockwise,
            pad_angle: fixture.layout.pad_angle, min_angle: fixture.layout.min_angle }],
    }, budget)
}

#[test]
fn chord_matches_independent_portable_angles_values_ids_and_single_result_owner() {
    for fixture in fixtures() {
        let mut compute = task(&fixture, usize::MAX).unwrap();
        assert_eq!(compute.take_result().err().unwrap().code, "COMPUTE_NOT_READY");
        assert!(compute.step(1).unwrap(), "{}", fixture.name);
        let output = compute.take_result().unwrap();
        assert_eq!(output.indices, fixture.indices, "{}", fixture.name);
        assert_eq!(output.values.len(), fixture.values.len());
        for (actual, expected) in output.values.iter().zip(&fixture.values) {
            match expected {
                Golden::Number(value) => assert_eq!(actual, value, "{}", fixture.name),
                Golden::Pi([numerator, denominator]) => {
                    let expected = *numerator as f64 / *denominator as f64 * std::f64::consts::PI;
                    assert!((actual - expected).abs() <= 1e-9, "{}: {actual} != {expected}", fixture.name);
                }
            }
        }
        assert_eq!(output.segments[0].valid_sample_count, fixture.valid_sample_count);
        assert_eq!(output.segments[0].available, fixture.available);
        assert_eq!(compute.take_result().err().unwrap().code, "COMPUTE_RESULT_CONSUMED");
    }
}

#[test]
fn chord_rejects_dictionary_state_fractional_label_overlap_and_insufficient_memory() {
    let fixture = fixtures().remove(0);
    assert_eq!(task(&fixture, 1).err().unwrap().code, "COMPUTE_MEMORY_LIMIT");
    for (state, value) in [(2, 0.0), (1, 0.5), (1, 2.0), (1, f64::INFINITY)] {
        let mut compute = task(&fixture, usize::MAX).unwrap();
        let offset = fixture.rows.len();
        compute.states[offset] = state; compute.values[offset] = value;
        assert_eq!(compute.step(1).err().unwrap().code, "COMPUTE_INVALID_INPUT");
        assert_eq!(compute.take_result().err().unwrap().code, "COMPUTE_RESULT_CONSUMED");
    }
    let mut compute = task(&fixture, usize::MAX).unwrap();
    if let ChartJob::Chord { source_column, .. } = &mut compute.params.jobs[0] { *source_column = 0; }
    assert!(compute.params.validate_shape(compute.values.len(), compute.states.len(), compute.categories.len()).is_err());
}

#[test]
fn chord_rejects_error_weight_states_and_unknown_bulk_states_before_publishing() {
    let fixture = fixtures().remove(0);
    let mut error_weight = task(&fixture, usize::MAX).unwrap();
    error_weight.states[0] = 3;
    assert_eq!(error_weight.step(1).err().unwrap().code, "COMPUTE_INVALID_INPUT");
    assert_eq!(error_weight.take_result().err().unwrap().code, "COMPUTE_RESULT_CONSUMED");
    for offset in [0, fixture.rows.len(), fixture.rows.len() * 2] {
        let mut unknown = task(&fixture, usize::MAX).unwrap();
        unknown.states[offset] = 255;
        assert_eq!(ChartStatistics::new(unknown.values, unknown.states, unknown.categories, unknown.params,
            usize::MAX).err().unwrap().code, "COMPUTE_INVALID_INPUT");
    }
}

#[test]
fn chord_skips_nonfinite_weights_and_never_emits_partial_overflow_results() {
    let fixtures = fixtures();
    let mut skipped = task(&fixtures[0], usize::MAX).unwrap();
    skipped.values[0] = f64::INFINITY;
    assert!(skipped.step(1).unwrap());
    let result = skipped.take_result().unwrap();
    assert_eq!(result.segments[0].valid_sample_count, 1); assert_eq!(result.values[0], 3.0);
    let mut link_overflow = task(&fixtures[0], usize::MAX).unwrap();
    link_overflow.values[0] = f64::MAX; link_overflow.values[1] = f64::MAX;
    assert_eq!(link_overflow.step(1).err().unwrap().code, "COMPUTE_NUMERIC_RANGE");
    assert_eq!(link_overflow.take_result().err().unwrap().code, "COMPUTE_RESULT_CONSUMED");
    let mut node_overflow = task(&fixtures[4], usize::MAX).unwrap();
    node_overflow.values[3] = 0.0;
    assert_eq!(node_overflow.step(1).err().unwrap().code, "COMPUTE_NUMERIC_RANGE");
    assert_eq!(node_overflow.take_result().err().unwrap().code, "COMPUTE_RESULT_CONSUMED");
}

#[test]
fn chord_extreme_finite_angle_budgets_close_both_directions_without_dense_matrices() {
    for clockwise in [true, false] {
        for (padding, minimum) in [(2.0, 10.0), (720.0, 0.0), (1e308, 1e308)] {
            let mut compute = task(&fixtures()[5], usize::MAX).unwrap();
            if let ChartJob::Chord { start_angle, clockwise: direction, pad_angle, min_angle, .. } = &mut compute.params.jobs[0] {
                *start_angle = 1e308; *direction = clockwise; *pad_angle = padding; *min_angle = minimum;
            }
            assert!(compute.step(1).unwrap());
            let output = compute.take_result().unwrap();
            assert!(output.values.iter().all(|value| value.is_finite()));
            let count = output.indices[0] as usize;
            assert!(count > 0);
            let direction = if clockwise { 1.0 } else { -1.0 };
            // Independent closed forms for these three authored budgets:
            // normal pads, pads filling a circle, or minimum arcs filling it.
            let expected_padding = if padding == 2.0 { std::f64::consts::PI / 90.0 }
                else if minimum == 0.0 { std::f64::consts::TAU / count as f64 } else { 0.0 };
            for node in 0..count {
                let start = output.values[node * 3 + 1];
                let end = output.values[node * 3 + 2];
                let span = (end - start) * direction;
                assert!(span >= -1e-10);
                if node + 1 < count {
                    let next_start = output.values[(node + 1) * 3 + 1];
                    assert!((next_start - end - expected_padding * direction).abs() <= 1e-9);
                }
            }
            let first_start = output.values[1];
            let final_end = output.values[count * 3 - 1];
            assert!((final_end + expected_padding * direction - first_start
                - std::f64::consts::TAU * direction).abs() <= 1e-9);
        }
    }
}
