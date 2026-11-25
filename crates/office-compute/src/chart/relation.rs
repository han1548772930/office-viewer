use super::{Input, JobSummary, RelationLayout, ChartStatisticsResult};
use super::math::{finite, reserve};
use crate::ComputeError;
use std::f64::consts::{PI, TAU};

#[derive(Clone, Copy)]
struct Node { x: f64, y: f64, vx: f64, vy: f64, value: f64, size: f64, seed: u32, fixed: bool }
#[derive(Clone, Copy)]
struct Link { source: usize, target: usize, value: f64 }
#[derive(Clone, Copy)]
struct Cell {
    left: f64, top: f64, right: f64, bottom: f64,
    x: f64, y: f64, start: usize, end: usize, children: [usize; 4],
}
const MISSING: usize = usize::MAX;

fn sample(input: &Input<'_>, column: u32, row: usize, label: &str) -> Result<f64, ComputeError> {
    input.sample(column, row).ok_or_else(|| ComputeError::invalid(format!("Relation {label} must have a finite numeric sample")))
}
fn endpoint(value: f64, count: usize) -> Result<usize, ComputeError> {
    if value < 0.0 || value.fract() != 0.0 || value >= count as f64 {
        return Err(ComputeError::invalid("Relation endpoint index is out of range"));
    }
    Ok(value as usize)
}

/// Tight child bounds compress empty quadtree levels. Every internal cell has
/// at least two children, bounding storage by 2*N; source names never enter the
/// numeric kernel. Coincident points use a deterministic aggregate at depth 24.
fn tree_cell(nodes: &[Node], order: &mut [usize], scratch: &mut [usize], cells: &mut Vec<Cell>,
    start: usize, end: usize, depth: usize) -> Result<usize, ComputeError> {
    let mut left = f64::INFINITY;
    let mut top = f64::INFINITY;
    let mut right = f64::NEG_INFINITY;
    let mut bottom = f64::NEG_INFINITY;
    let mut x = 0.0;
    let mut y = 0.0;
    for (offset, &index) in order[start..end].iter().enumerate() {
        let node = nodes[index];
        left = left.min(node.x); right = right.max(node.x);
        top = top.min(node.y); bottom = bottom.max(node.y);
        x += (node.x - x) / (offset + 1) as f64;
        y += (node.y - y) / (offset + 1) as f64;
    }
    finite(x, "Relation tree centroid")?; finite(y, "Relation tree centroid")?;
    let at = cells.len();
    cells.push(Cell { left, top, right, bottom, x, y, start, end, children: [MISSING; 4] });
    if end - start <= 4 || depth >= 24 || (right - left).max(bottom - top) <= 1e-6 { return Ok(at); }
    let mid_x = left / 2.0 + right / 2.0;
    let mid_y = top / 2.0 + bottom / 2.0;
    let quadrant = |node: Node| usize::from(node.x >= mid_x) + usize::from(node.y >= mid_y) * 2;
    let mut counts = [0; 4];
    for &index in &order[start..end] { counts[quadrant(nodes[index])] += 1; }
    if counts.iter().filter(|&&count| count > 0).count() < 2 { return Ok(at); }
    let mut offsets = [start; 4];
    for index in 1..4 { offsets[index] = offsets[index - 1] + counts[index - 1]; }
    let mut positions = offsets;
    for &index in &order[start..end] {
        let part = quadrant(nodes[index]);
        scratch[positions[part]] = index; positions[part] += 1;
    }
    order[start..end].copy_from_slice(&scratch[start..end]);
    for part in 0..4 {
        if counts[part] > 0 {
            let child = tree_cell(nodes, order, scratch, cells,
                offsets[part], offsets[part] + counts[part], depth + 1)?;
            cells[at].children[part] = child;
        }
    }
    Ok(at)
}

fn push_repulsion(node: Node, x: f64, y: f64, mass: f64, repulsion: f64,
    index: usize, target: &mut (f64, f64)) {
    let mut dx = node.x - x;
    let mut dy = node.y - y;
    if dx.abs() + dy.abs() < 1e-9 {
        let angle = ((node.seed as usize).wrapping_add(index.wrapping_mul(2654435761)) % 8192) as f64 * TAU / 8192.0;
        dx = angle.cos() * 0.1; dy = angle.sin() * 0.1;
    }
    let distance = dx.hypot(dy) + 0.1;
    let amount = (repulsion / distance / distance * mass).min(100.0);
    target.0 += amount * dx / distance; target.1 += amount * dy / distance;
}

fn repulsion_force(nodes: &[Node], order: &[usize], cells: &[Cell], index: usize,
    theta: f64, repulsion: f64, stack: &mut Vec<usize>) -> (f64, f64) {
    let node = nodes[index];
    let mut result = (0.0, 0.0);
    stack.clear(); stack.push(0);
    while let Some(at) = stack.pop() {
        let cell = cells[at];
        let leaf = cell.children[0] == MISSING && cell.children[1] == MISSING
            && cell.children[2] == MISSING && cell.children[3] == MISSING;
        if leaf && cell.end - cell.start <= 4 {
            for &other in &order[cell.start..cell.end] {
                if other != index { push_repulsion(node, nodes[other].x, nodes[other].y, 1.0, repulsion, index, &mut result); }
            }
            continue;
        }
        let contains = node.x >= cell.left && node.x <= cell.right && node.y >= cell.top && node.y <= cell.bottom;
        let distance = (cell.x - node.x).hypot(cell.y - node.y) + 0.1;
        if leaf || !contains && (cell.right - cell.left).max(cell.bottom - cell.top) / distance < theta {
            let mut mass = (cell.end - cell.start) as f64;
            let mut x = cell.x;
            let mut y = cell.y;
            if contains {
                mass -= 1.0;
                if mass <= 0.0 { continue; }
                x += (x - node.x) / mass; y += (y - node.y) / mass;
            }
            push_repulsion(node, x, y, mass, repulsion, index, &mut result);
        } else {
            for &child in cell.children.iter().rev() { if child != MISSING { stack.push(child); } }
        }
    }
    result
}

fn bounded_axis(value: f64, length: f64, inset: f64) -> f64 {
    if inset * 2.0 >= length { length / 2.0 } else { value.clamp(inset, length - inset) }
}

fn initialize(nodes: &mut [Node], links: &[Link], width: f64, height: f64,
    layout: RelationLayout) -> Result<(), ComputeError> {
    let count = nodes.len();
    if layout != RelationLayout::None {
        let largest = nodes.iter().map(|node| node.size).fold(0.0, f64::max);
        let radius = (width.min(height) / 2.0 - largest / 2.0 - 8.0).max(0.0);
        for (index, node) in nodes.iter_mut().enumerate() {
            let angle = -PI / 2.0 + index as f64 * TAU / count as f64;
            node.x = width / 2.0 + if count == 1 { 0.0 } else { radius * angle.cos() };
            node.y = height / 2.0 + if count == 1 { 0.0 } else { radius * angle.sin() };
        }
        return Ok(());
    }
    let mut incident = Vec::new(); reserve(&mut incident, count)?; incident.resize(count, 0.0);
    for link in links {
        incident[link.source] = finite(incident[link.source] + link.value, "Relation incident weight")?;
        // The upstream None layout ranks self-links twice; data node values
        // count them once. Keep those two contracts separate.
        incident[link.target] = finite(incident[link.target] + link.value, "Relation incident weight")?;
    }
    let mut rank = Vec::new(); reserve(&mut rank, count)?; rank.extend(0..count);
    rank.sort_unstable_by(|&first, &second| incident[second].total_cmp(&incident[first]).then(first.cmp(&second)));
    for (position, &index) in rank.iter().take(5).enumerate() {
        let angle = position as f64 * TAU / 5.0 - PI / 2.0;
        nodes[index].x = width / 2.0 + width.min(height) / 3.0 * angle.cos();
        nodes[index].y = height / 2.0 + width.min(height) / 3.0 * angle.sin();
        nodes[index].fixed = true;
    }
    for node in nodes.iter_mut().filter(|node| !node.fixed) {
        node.x = bounded_axis((node.seed & 0xff_ffff) as f64 % width, width, 50.0);
        node.y = bounded_axis(((node.seed >> 12) & 0xff_ffff) as f64 % height, height, 50.0);
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
pub(super) fn layout(input: &Input<'_>, value_column: u32, seed_column: u32, source_column: u32,
    target_column: u32, link_value_column: u32, node_count: u32, link_count: u32, width: f64, height: f64,
    layout: RelationLayout, use_value_as_symbol_size: bool, symbol_size: f64, repulsion: f64,
    gravity: f64, edge_length: f64, result: &mut ChartStatisticsResult) -> Result<JobSummary, ComputeError> {
    let count = node_count as usize;
    if count == 0 { return Ok(JobSummary::default()); }
    let mut nodes = Vec::new(); reserve(&mut nodes, count)?;
    let mut links = Vec::new(); reserve(&mut links, link_count as usize)?;
    let mut maximum: f64 = 0.0;
    for index in 0..count {
        let value = sample(input, value_column, index, "node value")?;
        let seed = sample(input, seed_column, index, "name seed")?;
        if value < 0.0 || seed < 0.0 || seed > 2_147_483_648.0 || seed.fract() != 0.0 {
            return Err(ComputeError::invalid("Relation node values must be nonnegative; name seeds must be UTF-16 hash integers"));
        }
        maximum = maximum.max(value);
        nodes.push(Node { x: 0.0, y: 0.0, vx: 0.0, vy: 0.0, value, size: symbol_size, seed: seed as u32, fixed: false });
    }
    for index in 0..link_count as usize {
        let source = endpoint(sample(input, source_column, index, "link source")?, count)?;
        let target = endpoint(sample(input, target_column, index, "link target")?, count)?;
        let value = sample(input, link_value_column, index, "link weight")?;
        if value < 0.0 { return Err(ComputeError::invalid("Relation link weights must be nonnegative")); }
        links.push(Link { source, target, value });
    }
    for node in &mut nodes {
        if use_value_as_symbol_size {
            // A zero-only graph remains editable/visible instead of dividing
            // 0 by 0. Mixed zero nodes retain their exact zero-sized symbols.
            node.size = if maximum == 0.0 { symbol_size.min(10.0) } else { node.value / maximum * symbol_size };
        }
    }
    initialize(&mut nodes, &links, width, height, layout)?;
    if layout != RelationLayout::Circular && count > 1 {
        let mut order = Vec::new(); reserve(&mut order, count)?; order.extend(0..count);
        let mut scratch = Vec::new(); reserve(&mut scratch, count)?; scratch.resize(count, 0);
        let cell_capacity = count.checked_mul(2).ok_or_else(ComputeError::memory)?;
        let mut cells = Vec::new(); reserve(&mut cells, cell_capacity)?;
        let mut stack = Vec::new(); reserve(&mut stack, cell_capacity)?;
        let mut forces = Vec::new(); reserve(&mut forces, count)?; forces.resize(count, (0.0, 0.0));
        let iterations = if layout == RelationLayout::None { 10 } else { 240 };
        for iteration in 0..iterations {
            cells.clear();
            tree_cell(&nodes, &mut order, &mut scratch, &mut cells, 0, count, 0)?;
            let fixed = layout == RelationLayout::None;
            let alpha = if fixed { 1.0 } else { (1.0 - iteration as f64 / iterations as f64).max(0.05) };
            for index in 0..count {
                forces[index] = if nodes[index].fixed { (0.0, 0.0) } else {
                    repulsion_force(&nodes, &order, &cells, index, if fixed { 0.1 } else { 0.65 },
                        if fixed { 5000.0 } else { repulsion }, &mut stack)
                };
                if !fixed {
                    forces[index].0 += (width / 2.0 - nodes[index].x) * gravity * 0.1;
                    forces[index].1 += (height / 2.0 - nodes[index].y) * gravity * 0.1;
                }
            }
            for link in &links {
                if link.source == link.target { continue; }
                let first = nodes[link.source]; let second = nodes[link.target];
                let dx = second.x - first.x; let dy = second.y - first.y;
                let distance = dx.hypot(dy) + 0.1;
                let amount = (distance - if fixed { 50.0 } else { edge_length }) * if fixed { 0.0001 } else { 0.025 };
                if !first.fixed { forces[link.source].0 += amount * dx / distance; forces[link.source].1 += amount * dy / distance; }
                if !second.fixed { forces[link.target].0 -= amount * dx / distance; forces[link.target].1 -= amount * dy / distance; }
            }
            let mut moved: f64 = 0.0;
            for (index, node) in nodes.iter_mut().enumerate() {
                if node.fixed { continue; }
                let fx = finite(forces[index].0, "Relation X force")?;
                let fy = finite(forces[index].1, "Relation Y force")?;
                node.vx = finite(node.vx + fx * alpha, "Relation X velocity")?;
                node.vy = finite(node.vy + fy * alpha, "Relation Y velocity")?;
                let dx = if fixed { node.vx } else { node.vx.clamp(-12.0, 12.0) * alpha };
                let dy = if fixed { node.vy } else { node.vy.clamp(-12.0, 12.0) * alpha };
                let x = finite(node.x + dx, "Relation X position")?;
                let y = finite(node.y + dy, "Relation Y position")?;
                node.x = if fixed { x } else { bounded_axis(x, width, node.size / 2.0 + 2.0) };
                node.y = if fixed { y } else { bounded_axis(y, height, node.size / 2.0 + 2.0) };
                node.vx *= if fixed { 0.6 } else { 0.85 };
                node.vy *= if fixed { 0.6 } else { 0.85 };
                moved = moved.max(dx.abs().max(dy.abs()));
            }
            if !fixed && iteration > 80 && moved < 0.01 { break; }
        }
    }
    for (index, node) in nodes.iter().enumerate() {
        result.values.extend([node.x, node.y, node.size]); result.indices.push(index as u32);
    }
    Ok(JobSummary { count, available: true, ..JobSummary::default() })
}
