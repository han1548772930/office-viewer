use super::*;

fn task(layout: RelationLayout, nodes: &[f64], seeds: &[f64], links: &[(f64, f64, f64)], weighted: bool) -> ChartStatistics {
    let rows = nodes.len().max(links.len());
    let mut values = vec![0.0; rows * 5]; let mut states = vec![0; rows * 5];
    for (index, &value) in nodes.iter().enumerate() {
        values[index] = value; states[index] = 1;
        values[rows + index] = seeds[index]; states[rows + index] = 1;
    }
    for (index, &(source, target, value)) in links.iter().enumerate() {
        for (column, item) in [(2, source), (3, target), (4, value)] {
            values[column * rows + index] = item; states[column * rows + index] = 1;
        }
    }
    ChartStatistics::new(values, states, vec![0; rows], ChartStatisticsParams {
        row_count: rows as u32, column_count: 5, category_count: u32::from(rows > 0),
        series: vec![ChartSeries { field: 2, series_position: 0, column: 0, x_column: None }],
        jobs: vec![ChartJob::RelationLayout { series_index: 0, slot: 0, seed_column: 1,
            source_column: 2, target_column: 3, link_value_column: 4, node_count: nodes.len() as u32,
            link_count: links.len() as u32, width: 600.0, height: 400.0, layout,
            use_value_as_symbol_size: weighted, symbol_size: if weighted { 32.0 } else { 10.0 },
            repulsion: 5000.0, gravity: 0.1, edge_length: 80.0 }],
    }, usize::MAX).unwrap()
}
fn complete(mut task: ChartStatistics) -> ChartStatisticsResult { assert!(task.step(1).unwrap()); task.take_result().unwrap() }

#[test]
fn relation_none_pins_five_incident_ranked_nodes_and_differs_from_circular() {
    let nodes = [1.0; 7]; let seeds = [65.0, 66.0, 67.0, 68.0, 69.0, 70.0, 71.0];
    let links = [(0.0, 1.0, 5.0), (0.0, 2.0, 4.0), (0.0, 3.0, 3.0), (0.0, 4.0, 2.0), (0.0, 5.0, 1.0)];
    let first = complete(task(RelationLayout::None, &nodes, &seeds, &links, false));
    let repeated = complete(task(RelationLayout::None, &nodes, &seeds, &links, false));
    assert_eq!(first.values, repeated.values); assert_eq!(first.indices, (0..7).collect::<Vec<_>>());
    assert!((first.values[0] - 300.0).abs() < 1e-9);
    assert!((first.values[1] - (200.0 - 400.0 / 3.0)).abs() < 1e-9);
    assert_ne!(first.values, complete(task(RelationLayout::Circular, &nodes, &seeds, &links, false)).values);
}

#[test]
fn relation_symbol_sizes_are_linear_and_zero_only_graphs_remain_finite() {
    let output = complete(task(RelationLayout::Circular, &[2.0, 8.0, 0.0], &[1.0, 2.0, 3.0], &[], true));
    assert_eq!(output.values[2], 8.0); assert_eq!(output.values[5], 32.0); assert_eq!(output.values[8], 0.0);
    let zero = complete(task(RelationLayout::Circular, &[0.0, 0.0], &[1.0, 2.0], &[(0.0, 1.0, 0.0)], true));
    assert!(zero.values.iter().all(|value| value.is_finite())); assert_eq!(zero.values[2], 10.0);
}

#[test]
fn relation_force_is_deterministic_bounded_and_reacts_to_force_parameters() {
    let nodes = [4.0; 8]; let seeds = [65.0; 8]; let links = [(0.0, 1.0, 4.0), (1.0, 7.0, 1.0), (7.0, 7.0, 0.0)];
    let first = complete(task(RelationLayout::Force, &nodes, &seeds, &links, false));
    assert_eq!(first.values, complete(task(RelationLayout::Force, &nodes, &seeds, &links, false)).values);
    for node in first.values.chunks_exact(3) { assert!(node[0] >= 7.0 && node[0] <= 593.0 && node[1] >= 7.0 && node[1] <= 393.0); }
    let mut altered = task(RelationLayout::Force, &nodes, &seeds, &links, false);
    if let ChartJob::RelationLayout { repulsion, gravity, .. } = &mut altered.params.jobs[0] { *repulsion = 0.0; *gravity = 0.4; }
    assert_ne!(first.values, complete(altered).values);
}

#[test]
fn relation_rejects_invalid_bulk_and_consumed_results() {
    for links in [vec![(0.5, 1.0, 1.0)], vec![(0.0, 2.0, 1.0)], vec![(0.0, 1.0, -1.0)]] {
        let mut compute = task(RelationLayout::Force, &[1.0, 1.0], &[65.0, 66.0], &links, false);
        assert_eq!(compute.step(1).err().unwrap().code, "COMPUTE_INVALID_INPUT");
        assert_eq!(compute.take_result().err().unwrap().code, "COMPUTE_RESULT_CONSUMED");
    }
    let mut compute = task(RelationLayout::None, &[1.0], &[65.0], &[], false);
    assert_eq!(compute.take_result().err().unwrap().code, "COMPUTE_NOT_READY");
    assert!(compute.step(1).unwrap()); assert!(compute.take_result().is_ok());
    assert_eq!(compute.take_result().err().unwrap().code, "COMPUTE_RESULT_CONSUMED");
}
