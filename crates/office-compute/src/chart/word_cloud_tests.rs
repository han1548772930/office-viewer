use super::*;

fn task(glyphs: Vec<WordCloudGlyph>, mask: &[f64], glyph_cells: &[f64], width: u32, height: u32) -> ChartStatistics {
    let rows = mask.len().max(glyph_cells.len());
    let mut values = vec![0.0; rows * 2];
    let mut states = vec![0; rows * 2];
    for (row, value) in glyph_cells.iter().enumerate() { values[row] = *value; states[row] = 1; }
    for (row, value) in mask.iter().enumerate() { values[rows + row] = *value; states[rows + row] = 1; }
    ChartStatistics::new(values, states, vec![0; rows], ChartStatisticsParams {
        row_count: rows as u32, column_count: 2, category_count: 1,
        series: vec![ChartSeries { field: 1, series_position: 0, column: 0, x_column: None }],
        jobs: vec![ChartJob::WordCloud { series_index: 0, slot: 0, glyph_column: 0, mask_column: 1,
            width, height, glyphs, max_attempts: 2400, min_font_size: 16.0 }],
    }, usize::MAX).unwrap()
}

#[test]
fn word_cloud_uses_actual_occupied_cells_and_never_overlaps_or_leaves_mask() {
    let glyphs = (0..3).map(|index| WordCloudGlyph { width: 1, height: 1,
        offset: index, length: 1, font_size: 16.0, weight: (index + 1) as f64 }).collect();
    let mut compute = task(glyphs, &[1.0; 9], &[0.0; 3], 3, 3);
    assert_eq!(compute.take_result().err().unwrap().code, "COMPUTE_NOT_READY");
    assert!(compute.step(1).unwrap());
    let output = compute.take_result().unwrap();
    assert_eq!(output.indices.len(), 3);
    let mut cells = HashSet::new();
    for coordinates in output.values.chunks_exact(3) {
        assert_eq!(coordinates[2], 1.0);
        assert!(coordinates[0] >= 0.5 && coordinates[0] <= 2.5);
        assert!(coordinates[1] >= 0.5 && coordinates[1] <= 2.5);
        assert!(cells.insert((coordinates[0] as u32, coordinates[1] as u32)));
    }
    assert_eq!(compute.take_result().err().unwrap().code, "COMPUTE_RESULT_CONSUMED");
}

#[test]
fn word_cloud_rejects_non_binary_masks_and_duplicate_or_out_of_range_glyph_cells() {
    for cells in [vec![0.0, 0.0], vec![0.0, 4.0], vec![0.0, 0.5]] {
        let mut compute = task(vec![WordCloudGlyph { width: 2, height: 2, offset: 0,
            length: 2, font_size: 16.0, weight: 1.0 }], &[1.0; 9], &cells, 3, 3);
        assert_eq!(compute.step(1).err().unwrap().code, "COMPUTE_INVALID_INPUT");
        assert_eq!(compute.take_result().err().unwrap().code, "COMPUTE_RESULT_CONSUMED");
    }
    let mut compute = task(vec![WordCloudGlyph { width: 1, height: 1, offset: 0,
        length: 1, font_size: 16.0, weight: 1.0 }], &[0.5; 9], &[0.0], 3, 3);
    assert_eq!(compute.step(1).err().unwrap().code, "COMPUTE_INVALID_INPUT");
}

#[test]
fn word_cloud_reports_missing_placements_without_publishing_outside_mask() {
    let glyphs = (0..2).map(|index| WordCloudGlyph { width: 1, height: 1,
        offset: index, length: 1, font_size: 16.0, weight: 1.0 }).collect();
    let mut compute = task(glyphs, &[1.0], &[0.0; 2], 1, 1);
    assert!(compute.step(1).unwrap());
    let result = compute.take_result().unwrap();
    assert_eq!(result.indices.len(), 1); assert_eq!(result.values, vec![0.5, 0.5, 1.0]);
}
