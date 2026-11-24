use super::{ChartSeries, ChartStatisticsResult, Input, JobSummary};
use super::math::{finite, reserve};
use crate::ComputeError;
use std::collections::HashMap;

const ROOT: u32 = u32::MAX;

#[derive(Clone, Copy)]
struct Token { key: u32, label: u32 }

struct Node { parent: usize, key: u32, label: u32, self_value: f64, value: f64, children: Vec<usize> }

fn token(input: &Input<'_>, key_column: u32, label_column: u32, row: usize,
    key_count: u32, label_count: u32) -> Result<Option<Token>, ComputeError> {
    let key_offset = key_column as usize * input.rows + row;
    let label_offset = label_column as usize * input.rows + row;
    let key_state = input.states[key_offset];
    let label_state = input.states[label_offset];
    if key_state == 0 && label_state == 0 { return Ok(None); }
    let key = input.values[key_offset];
    let label = input.values[label_offset];
    if key_state != 1 || label_state != 1 || !key.is_finite() || key < 0.0 || key.fract() != 0.0 || key >= key_count as f64
        || !label.is_finite() || label < 0.0 || label.fract() != 0.0 || label >= label_count as f64 {
        return Err(ComputeError::invalid("Hierarchy token state or dictionary index is invalid"));
    }
    Ok(Some(Token { key: key as u32, label: label as u32 }))
}

/// Renderer-independent fixed Ua policy. IDs and source number formatting are
/// reconstructed from retained dictionaries by the owner after transfer.
/// A hierarchy is one existing chart job; cancellation during its sort/aggregate
/// uses the Worker's established terminate-and-recreate path.
pub(super) fn build(input: &Input<'_>, series: &ChartSeries, hierarchy_columns: &[u32], label_columns: &[u32],
    key_count: u32, label_count: u32, output: &mut ChartStatisticsResult) -> Result<JobSummary, ComputeError> {
    let maximum_nodes = input.rows.checked_mul(hierarchy_columns.len()).ok_or_else(ComputeError::memory)?;
    let mut nodes = Vec::new();
    reserve(&mut nodes, maximum_nodes + 1)?;
    nodes.push(Node { parent: 0, key: 0, label: 0, self_value: 0.0, value: 0.0, children: Vec::new() });
    let mut lookup = HashMap::new();
    let mut cache = Vec::new();
    reserve(&mut cache, hierarchy_columns.len())?;
    cache.resize(hierarchy_columns.len(), None::<Token>);
    let mut tokens = Vec::new();
    reserve(&mut tokens, cache.len())?;
    let mut path = Vec::new();
    reserve(&mut path, cache.len())?;
    let mut count = 0;
    for row in 0..input.rows {
        tokens.clear(); path.clear();
        for depth in 0..cache.len() {
            tokens.push(token(input, hierarchy_columns[depth], label_columns[depth], row, key_count, label_count)?);
        }
        let mut rejected = false;
        for depth in 0..cache.len() {
            let Some(current) = tokens[depth] else {
                if let Some(carried) = cache[depth] { path.push(carried); continue; }
                rejected = depth == 0 || tokens[depth + 1..].iter().any(Option::is_some);
                break;
            };
            if cache[depth].map(|previous| previous.key) != Some(current.key) {
                for descendant in cache.iter_mut().skip(depth + 1) { *descendant = None; }
            }
            cache[depth] = Some(current); path.push(current);
        }
        // Invalid weights still affect the scoped hierarchy of later rows.
        let Some(value) = input.sample(series.column, row) else { continue; };
        if rejected || path.is_empty() { continue; }
        count += 1;
        let mut parent = 0;
        for current in &path {
            let identity = (parent, current.key);
            let child = if let Some(child) = lookup.get(&identity) { *child } else {
                let child = nodes.len();
                if child > maximum_nodes { return Err(ComputeError::memory()); }
                lookup.try_reserve(1).map_err(|_| ComputeError::memory())?;
                reserve(&mut nodes[parent].children, 1)?;
                nodes.push(Node { parent, key: current.key, label: current.label, self_value: 0.0, value: 0.0, children: Vec::new() });
                nodes[parent].children.push(child); lookup.insert(identity, child); child
            };
            parent = child;
        }
        nodes[parent].self_value = finite(nodes[parent].self_value + value, "Hierarchy path self value")?;
    }
    // Creation is parent-before-child, so reverse creation is a postorder.
    for index in (1..nodes.len()).rev() {
        let mut children = std::mem::take(&mut nodes[index].children);
        children.retain(|child| nodes[*child].value > 0.0);
        children.sort_by(|left, right| nodes[*right].value.total_cmp(&nodes[*left].value).then_with(|| left.cmp(right)));
        let mut child_value = 0.0;
        for child in &children { child_value = finite(child_value + nodes[*child].value, "Hierarchy child total")?; }
        nodes[index].value = finite(nodes[index].self_value + child_value, "Hierarchy node total")?;
        nodes[index].children = children;
    }
    let mut source_roots = Vec::new();
    reserve(&mut source_roots, nodes[0].children.len())?;
    source_roots.extend(nodes[0].children.iter().copied().filter(|root| nodes[*root].value > 0.0));
    let mut roots = source_roots.clone();
    roots.sort_by(|left, right| nodes[*right].value.total_cmp(&nodes[*left].value).then_with(|| left.cmp(right)));
    let mut preorder = Vec::new();
    reserve(&mut preorder, nodes.len())?;
    let mut stack = Vec::new();
    reserve(&mut stack, nodes.len())?;
    stack.extend(roots.iter().rev().copied());
    let mut output_index = Vec::new();
    reserve(&mut output_index, nodes.len())?;
    output_index.resize(nodes.len(), ROOT);
    while let Some(index) = stack.pop() {
        output_index[index] = preorder.len() as u32; preorder.push(index);
        stack.extend(nodes[index].children.iter().rev().copied());
    }
    output.indices.extend([preorder.len() as u32, roots.len() as u32]);
    for original in &preorder {
        let node = &nodes[*original];
        output.values.extend([node.value, node.self_value]);
        output.indices.extend([if node.parent == 0 { ROOT } else { output_index[node.parent] }, node.key, node.label, (*original - 1) as u32]);
    }
    output.indices.extend(source_roots.iter().map(|root| output_index[*root]));
    Ok(JobSummary { count, available: !roots.is_empty(), ..Default::default() })
}
