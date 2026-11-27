use super::*;

// P26 shared-fixture drafts. These tests have not been executed in this phase.
#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct Fixture { name: String, depth: usize, key_count: u32, rows: Vec<Row>, values: Vec<f64>, indices: Vec<u32>, valid_sample_count: usize, available: bool }
#[derive(Deserialize)]
struct Row { weight: Option<f64>, keys: Vec<Option<u32>> }
#[derive(Deserialize)]
struct Fixtures { cases: Vec<Fixture> }

fn task(fixture: &Fixture, budget: usize) -> Result<ChartStatistics, ComputeError> {
    let rows = fixture.rows.len();
    let columns = fixture.depth * 2 + 1;
    let mut values = vec![0.0; rows * columns];
    let mut states = vec![0; values.len()];
    for (row_index, row) in fixture.rows.iter().enumerate() {
        if let Some(weight) = row.weight { values[row_index] = weight; states[row_index] = 1; }
        for (depth, key) in row.keys.iter().enumerate() {
            if let Some(key) = key {
                for column in [depth + 1, fixture.depth + depth + 1] {
                    let offset = column * rows + row_index; values[offset] = *key as f64; states[offset] = 1;
                }
            }
        }
    }
    ChartStatistics::new(values, states, vec![0; rows], ChartStatisticsParams { row_count: rows as u32, column_count: columns as u32, category_count: u32::from(rows > 0),
        series: vec![ChartSeries { field: fixture.depth as u32, series_position: 0, column: 0, x_column: None }],
        jobs: vec![ChartJob::Hierarchy { series_index: 0, slot: 0, hierarchy_columns: (1..=fixture.depth as u32).collect(),
            label_columns: (fixture.depth as u32 + 1..=fixture.depth as u32 * 2).collect(), key_count: fixture.key_count, label_count: fixture.key_count }] }, budget)
}

fn fixtures() -> Vec<Fixture> {
    let source = include_str!("../../../../tests/fixtures/chart-hierarchy.json");
    serde_json::from_str::<Fixtures>(source).unwrap().cases
}

#[test]
fn hierarchy_matches_shared_portable_policy_packets_and_has_one_result_owner() {
    for fixture in fixtures() {
        let mut compute = task(&fixture, usize::MAX).unwrap();
        assert_eq!(compute.take_result().err().unwrap().code, "COMPUTE_NOT_READY");
        assert!(compute.step(1).unwrap(), "{}", fixture.name);
        let output = compute.take_result().unwrap();
        assert_eq!(output.values, fixture.values, "{}", fixture.name);
        assert_eq!(output.indices, fixture.indices, "{}", fixture.name);
        assert_eq!(output.segments[0].valid_sample_count, fixture.valid_sample_count, "{}", fixture.name);
        assert_eq!(output.segments[0].available, fixture.available, "{}", fixture.name);
        assert_eq!(compute.take_result().err().unwrap().code, "COMPUTE_RESULT_CONSUMED");
    }
}

#[test]
fn hierarchy_rejects_insufficient_memory_and_mismatched_token_states() {
    let fixture = fixtures().remove(0);
    assert_eq!(task(&fixture, 1).err().unwrap().code, "COMPUTE_MEMORY_LIMIT");
    let mut compute = task(&fixture, usize::MAX).unwrap();
    let rows = fixture.rows.len();
    compute.states[rows] = 0;
    assert_eq!(compute.step(1).err().unwrap().code, "COMPUTE_INVALID_INPUT");
    assert_eq!(compute.take_result().err().unwrap().code, "COMPUTE_RESULT_CONSUMED");
}

#[test]
fn hierarchy_rejects_overlapping_columns_and_nonfinite_signed_aggregation() {
    let fixture = fixtures().remove(0);
    let mut compute = task(&fixture, usize::MAX).unwrap();
    if let ChartJob::Hierarchy { label_columns, .. } = &mut compute.params.jobs[0] { label_columns[0] = 1; }
    assert!(compute.params.validate_shape(compute.values.len(), compute.states.len(), compute.categories.len()).is_err());
    let fixture = Fixture { name: "overflow".into(), depth: 1, key_count: 1,
        rows: vec![Row { weight: Some(f64::MAX), keys: vec![Some(0)] }, Row { weight: Some(f64::MAX), keys: vec![Some(0)] }],
        values: vec![], indices: vec![], valid_sample_count: 0, available: false };
    let mut compute = task(&fixture, usize::MAX).unwrap();
    assert_eq!(compute.step(1).err().unwrap().code, "COMPUTE_NUMERIC_RANGE");
}
