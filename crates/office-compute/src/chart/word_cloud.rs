use super::{math::reserve, ChartStatisticsResult, Input, JobSummary, WordCloudGlyph};
use crate::ComputeError;

fn next_random(seed: &mut u32) -> f64 {
    *seed ^= *seed << 13;
    *seed ^= *seed >> 17;
    *seed ^= *seed << 5;
    *seed as f64 / (u32::MAX as f64 + 1.0)
}

/// Raster cells, mask and occupied cells share one coordinate system. Scaling
/// visits every occupied source cell, so narrow glyph strokes cannot disappear
/// merely because a bounding-box estimate missed them.
fn visit_cells(input: &Input<'_>, column: u32, glyph: &WordCloudGlyph, scale: f64,
    left: usize, top: usize, width: usize, mut visit: impl FnMut(usize) -> bool) -> bool {
    for row in glyph.offset as usize..glyph.offset as usize + glyph.length as usize {
        // The complete input was checked before any placement was published.
        let cell = input.values[column as usize * input.rows + row] as usize;
        let x = (cell % glyph.width as usize) as f64;
        let y = (cell / glyph.width as usize) as f64;
        let at = (top + (y * scale).floor() as usize) * width + left + (x * scale).floor() as usize;
        if !visit(at) { return false; }
    }
    true
}

pub(super) fn layout(input: &Input<'_>, glyph_column: u32, mask_column: u32,
    width: u32, height: u32, glyphs: &[WordCloudGlyph], max_attempts: u32,
    min_font_size: f64, result: &mut ChartStatisticsResult) -> Result<JobSummary, ComputeError> {
    let width = width as usize;
    let height = height as usize;
    let cells = width * height;
    let mut mask = Vec::new();
    let mut occupied = Vec::new();
    reserve(&mut mask, cells)?;
    reserve(&mut occupied, cells)?;
    occupied.resize(cells, 0u8);
    let mut center_x = 0.0;
    let mut center_y = 0.0;
    let mut allowed = 0usize;
    for row in 0..cells {
        let value = input.sample(mask_column, row)
            .ok_or_else(|| ComputeError::invalid("Word cloud mask cells require numeric 0/1 states"))?;
        if value != 0.0 && value != 1.0 {
            return Err(ComputeError::invalid("Word cloud mask cells must be exact 0/1"));
        }
        mask.push(value as u8);
        if value == 1.0 {
            center_x += (row % width) as f64;
            center_y += (row / width) as f64;
            allowed += 1;
        }
    }
    for glyph in glyphs {
        let mut previous = None;
        for row in glyph.offset as usize..glyph.offset as usize + glyph.length as usize {
            let value = input.sample(glyph_column, row)
                .ok_or_else(|| ComputeError::invalid("Word cloud glyph cells require numeric states"))?;
            if value < 0.0 || value.fract() != 0.0 || value >= glyph.width as f64 * glyph.height as f64
                || previous.is_some_and(|previous| value <= previous) {
                return Err(ComputeError::invalid("Word cloud glyph cells must be sorted unique in-range indices"));
            }
            previous = Some(value);
        }
    }
    if allowed == 0 || glyphs.is_empty() { return Ok(JobSummary::default()); }
    center_x = center_x / allowed as f64 + 0.5;
    center_y = center_y / allowed as f64 + 0.5;
    let mut order = Vec::new();
    reserve(&mut order, glyphs.len())?;
    order.extend(0..glyphs.len());
    order.sort_unstable_by(|&first, &second| glyphs[second].weight.total_cmp(&glyphs[first].weight)
        .then_with(|| first.cmp(&second)));
    let mut seed = 1u32;
    let mut count = 0;
    for index in order {
        let glyph = &glyphs[index];
        if glyph.length == 0 { continue; }
        let phase = next_random(&mut seed) * std::f64::consts::TAU;
        let minimum = (min_font_size / glyph.font_size).clamp(0.0, 1.0);
        let mut scale = 1.0;
        'shrink: loop {
            let glyph_width = (glyph.width as f64 * scale).ceil() as usize;
            let glyph_height = (glyph.height as f64 * scale).ceil() as usize;
            if glyph_width <= width && glyph_height <= height {
                for attempt in 0..max_attempts {
                    let radius = (attempt as f64 / max_attempts as f64).sqrt();
                    let angle = phase + attempt as f64 * 0.47;
                    let x = center_x + angle.cos() * radius * width as f64 * 0.5 - glyph_width as f64 * 0.5;
                    let y = center_y + angle.sin() * radius * height as f64 * 0.5 - glyph_height as f64 * 0.5;
                    if x < 0.0 || y < 0.0 { continue; }
                    let left = x.round() as usize;
                    let top = y.round() as usize;
                    if left + glyph_width > width || top + glyph_height > height { continue; }
                    if !visit_cells(input, glyph_column, glyph, scale, left, top, width,
                        |at| mask[at] == 1 && occupied[at] == 0) { continue; }
                    visit_cells(input, glyph_column, glyph, scale, left, top, width,
                        |at| { occupied[at] = 1; true });
                    result.values.extend_from_slice(&[left as f64 + glyph.width as f64 * scale * 0.5,
                        top as f64 + glyph.height as f64 * scale * 0.5, scale]);
                    result.indices.push(index as u32);
                    count += 1;
                    break 'shrink;
                }
            }
            if scale <= minimum { break; }
            // Fixed public zoom behavior shrinks by .8, with no enlargement.
            scale = (scale * 0.8).max(minimum);
        }
    }
    Ok(JobSummary { count, available: count > 0, ..JobSummary::default() })
}
