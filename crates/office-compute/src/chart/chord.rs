use super::{ChartSeries, ChartStatisticsResult, Input, JobSummary};
use super::math::{finite, reserve};
use crate::ComputeError;
use std::collections::HashMap;

struct Node { label: u32, value: f64, span: f64, start: f64, end: f64 }
struct Link { source: usize, target: usize, value: f64, source_start: f64, source_end: f64, target_start: f64, target_end: f64 }
struct Angles { start: f64, direction: f64, padding: f64, minimum: f64, available: f64 }

fn label(input: &Input<'_>, column: u32, row: usize, count: u32) -> Result<Option<u32>, ComputeError> {
    let at = column as usize * input.rows + row;
    if input.states[at] == 0 { return Ok(None); }
    let value = input.values[at];
    if input.states[at] != 1 || !value.is_finite() || value < 0.0 || value.fract() != 0.0 || value >= count as f64 {
        return Err(ComputeError::invalid("Chord endpoint dictionary index or state is invalid"));
    }
    Ok(Some(value as u32))
}

fn node_at(label: u32, lookup: &mut HashMap<u32, usize>, nodes: &mut Vec<Node>) -> Result<usize, ComputeError> {
    if let Some(index) = lookup.get(&label) { return Ok(*index); }
    lookup.try_reserve(1).map_err(|_| ComputeError::memory())?;
    let index = nodes.len();
    nodes.push(Node { label, value: 0.0, span: 0.0, start: 0.0, end: 0.0 });
    lookup.insert(label, index);
    Ok(index)
}

fn parameters(start_angle: f64, clockwise: bool, pad_angle: f64, min_angle: f64, nodes: usize) -> Angles {
    let start = -(start_angle % 360.0) / 180.0 * std::f64::consts::PI;
    let direction = if clockwise { 1.0 } else { -1.0 };
    if nodes == 0 { return Angles { start, direction, padding: 0.0, minimum: 0.0, available: 0.0 }; }
    let maximum = std::f64::consts::TAU / nodes as f64;
    let mut minimum = maximum.min(min_angle / 180.0 * std::f64::consts::PI);
    let mut padding = pad_angle / 180.0 * std::f64::consts::PI;
    if padding >= maximum { padding = (maximum - minimum).max(0.0); }
    if padding + minimum >= maximum { minimum = (maximum - padding).max(0.0); }
    let available = (std::f64::consts::TAU - padding * nodes as f64).max(0.0);
    minimum = minimum.min(available / nodes as f64);
    Angles { start, direction, padding, minimum, available }
}

fn compensated_add(sum: &mut f64, correction: &mut f64, value: f64) {
    let addend = value - *correction;
    let next = *sum + addend;
    *correction = next - *sum - addend;
    *sum = next;
}

/// Fixed display-label IDs own identity. Undirected reverse duplicates retain
/// the first endpoint order; all sums are finite and nodes are aggregated after
/// links. Strings/formatters remain with the owner, not in a large control DTO.
/// This sparse O(rows + nodes + links) job uses the existing whole-job Worker
/// cancellation path (terminate/recreate), not per-node async WASM stepping.
pub(super) fn build(input: &Input<'_>, series: &ChartSeries, source_column: u32, target_column: u32,
    label_count: u32, start_angle: f64, clockwise: bool, pad_angle: f64, min_angle: f64,
    output: &mut ChartStatisticsResult) -> Result<JobSummary, ComputeError> {
    let mut nodes = Vec::new();
    let mut links: Vec<Link> = Vec::new();
    reserve(&mut nodes, input.rows.checked_mul(2).ok_or_else(ComputeError::memory)?)?;
    reserve(&mut links, input.rows)?;
    let mut node_ids = HashMap::new();
    let mut link_ids: HashMap<(usize, usize), usize> = HashMap::new();
    let mut count = 0;
    for row in 0..input.rows {
        let first = label(input, source_column, row, label_count)?;
        let second = label(input, target_column, row, label_count)?;
        if input.states[series.column as usize * input.rows + row] > 2 {
            return Err(ComputeError::invalid("Chord weight state is invalid"));
        }
        let Some(value) = input.sample(series.column, row) else { continue; };
        let (Some(first), Some(second)) = (first, second) else { continue; };
        if value <= 0.0 || first == second { continue; }
        let source = node_at(first, &mut node_ids, &mut nodes)?;
        let target = node_at(second, &mut node_ids, &mut nodes)?;
        let key = (source.min(target), source.max(target));
        if let Some(index) = link_ids.get(&key) {
            links[*index].value = finite(links[*index].value + value, "Chord link aggregate")?;
        } else {
            link_ids.try_reserve(1).map_err(|_| ComputeError::memory())?;
            link_ids.insert(key, links.len());
            links.push(Link { source, target, value, source_start: 0.0, source_end: 0.0, target_start: 0.0, target_end: 0.0 });
        }
        count += 1;
    }
    for link in &links {
        nodes[link.source].value = finite(nodes[link.source].value + link.value, "Chord node incident aggregate")?;
        nodes[link.target].value = finite(nodes[link.target].value + link.value, "Chord node incident aggregate")?;
    }
    let angles = parameters(start_angle, clockwise, pad_angle, min_angle, nodes.len());
    let scale = nodes.iter().fold(0.0_f64, |maximum, node| maximum.max(node.value));
    let total: f64 = nodes.iter().map(|node| node.value / scale).sum();
    let mut deficit = 0.0;
    let mut surplus = 0.0;
    for node in &mut nodes {
        node.span = angles.available * ((node.value / scale) / total);
        deficit += (angles.minimum - node.span).max(0.0);
        surplus += (node.span - angles.minimum).max(0.0);
    }
    // A proportional surplus reduction preserves the min-angle floor and
    // complete sweep even with tiny values or a globally overflowing sum.
    // ECharts' borrowing order is not an authoritative second local engine.
    let fraction = if surplus > 0.0 { (deficit / surplus).min(1.0) } else { 0.0 };
    let mut allocated = 0.0;
    let mut correction = 0.0;
    for node in &mut nodes {
        node.span = if node.span < angles.minimum { angles.minimum } else { node.span - (node.span - angles.minimum) * fraction };
        compensated_add(&mut allocated, &mut correction, node.span);
    }
    if let Some(last) = nodes.last_mut() { last.span = (last.span + angles.available - allocated).max(0.0); }
    let mut offsets = Vec::new();
    let mut consumed_values = Vec::new();
    reserve(&mut offsets, nodes.len())?;
    reserve(&mut consumed_values, nodes.len())?;
    consumed_values.resize(nodes.len(), 0.0);
    let mut consumed = 0.0;
    correction = 0.0;
    for node in &mut nodes {
        node.start = angles.start + consumed * angles.direction;
        node.end = node.start + node.span * angles.direction;
        offsets.push(node.start);
        compensated_add(&mut consumed, &mut correction, node.span + angles.padding);
    }
    for link in &mut links {
        let source = &nodes[link.source];
        let target = &nodes[link.target];
        link.source_start = offsets[link.source];
        link.target_start = offsets[link.target];
        consumed_values[link.source] = finite(consumed_values[link.source] + link.value, "Chord endpoint consumed value")?;
        consumed_values[link.target] = finite(consumed_values[link.target] + link.value, "Chord endpoint consumed value")?;
        link.source_end = source.start + source.span * (consumed_values[link.source] / source.value) * angles.direction;
        link.target_end = target.start + target.span * (consumed_values[link.target] / target.value) * angles.direction;
        offsets[link.source] = link.source_end;
        offsets[link.target] = link.target_end;
    }
    output.indices.push(nodes.len() as u32);
    output.indices.push(links.len() as u32);
    for node in &nodes {
        output.values.extend([node.value, node.start, node.end]);
        output.indices.push(node.label);
    }
    for link in &links {
        output.values.extend([link.value, link.source_start, link.source_end, link.target_start, link.target_end]);
        output.indices.extend([link.source as u32, link.target as u32]);
    }
    Ok(JobSummary { count, available: !links.is_empty(), ..Default::default() })
}
