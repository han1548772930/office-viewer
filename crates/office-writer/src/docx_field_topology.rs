//! A field's result range alone cannot distinguish an empty child from an
//! adjacent root, or equal-range parent/child fields. Preserve explicit source
//! parent identity and use geometry only for fields without that provenance.
use std::collections::HashMap;

use serde_json::Value;

use super::Field;
use crate::ExportError;

const MAX_FIELD_DEPTH: usize = 512;

pub(super) struct FieldTopology {
    pub depths: Vec<usize>,
    /// false opens the field; true closes it. DFS order is authoritative when
    /// multiple delimiters share the same paragraph/UTF-16 position.
    pub events: Vec<(usize, bool)>,
}

fn depths(parents: &[Option<usize>]) -> Result<Vec<usize>, ExportError> {
    let mut state = vec![0u8; parents.len()];
    let mut result = vec![0usize; parents.len()];
    for first in 0..parents.len() {
        if state[first] == 2 { continue; }
        let mut path = Vec::new();
        let mut current = Some(first);
        while let Some(index) = current {
            if state[index] == 2 { break; }
            if state[index] == 1 { return Err(ExportError::invalid("正文字段父身份形成循环。")); }
            state[index] = 1;
            path.push(index);
            if path.len() > MAX_FIELD_DEPTH + 1 { return Err(ExportError::invalid("正文字段嵌套层级超出 512。")); }
            current = parents[index];
        }
        while let Some(index) = path.pop() {
            result[index] = parents[index].map_or(0, |parent| result[parent] + 1);
            if result[index] > MAX_FIELD_DEPTH { return Err(ExportError::invalid("正文字段嵌套层级超出 512。")); }
            state[index] = 2;
        }
    }
    Ok(result)
}

pub(super) fn resolve(fields: &[Field]) -> Result<FieldTopology, ExportError> {
    let by_id: HashMap<&str, usize> = fields.iter().enumerate().map(|(index, field)|
        (field.value["id"].as_str().unwrap_or_default(), index)).collect();
    let mut parents = vec![None; fields.len()];
    let mut infer = vec![true; fields.len()];
    for (index, field) in fields.iter().enumerate() {
        let source = field.value.get("metadata").and_then(|metadata| metadata.get("source"));
        if source.is_some_and(|source| !source.is_null() && !source.is_object()) {
            return Err(ExportError::invalid("正文字段 source 必须为对象。"));
        }
        match source.and_then(|source| source.get("parentFieldId")) {
            None => {}
            Some(Value::Null) => infer[index] = false,
            Some(Value::String(id)) if !id.trim().is_empty() && id.len() <= 4096 => {
                infer[index] = false;
                let parent = *by_id.get(id.as_str()).ok_or_else(|| ExportError::invalid("正文字段父身份不存在。"))?;
                if parent == index { return Err(ExportError::invalid("正文字段不能以自身为父字段。")); }
                if fields[parent].start > field.start || fields[parent].end < field.end {
                    return Err(ExportError::invalid("正文字段结果范围超出显式父字段。"));
                }
                parents[index] = Some(parent);
            }
            Some(_) => return Err(ExportError::invalid("正文字段 parentFieldId 必须为有效标识或 null。")),
        }
    }
    let explicit_depths = depths(&parents)?;
    let mut order: Vec<usize> = (0..fields.len()).collect();
    order.sort_by_key(|&index| (fields[index].start, std::cmp::Reverse(fields[index].end), explicit_depths[index],
        fields[index].begin_sequence.unwrap_or(u64::from(fields[index].bookmark)), fields[index].bookmark));
    let mut enclosing: Vec<usize> = Vec::new();
    for index in order {
        let field = &fields[index];
        while enclosing.last().is_some_and(|&parent| fields[parent].end <= field.start) { enclosing.pop(); }
        if infer[index] {
            if let Some(&parent) = enclosing.last() {
                if fields[parent].end < field.end { return Err(ExportError::invalid("正文字段结果范围部分交叉。")); }
                parents[index] = Some(parent);
            }
        }
        if field.start < field.end { enclosing.push(index); }
    }
    let depths = depths(&parents)?;
    let root = fields.len();
    let mut children = vec![Vec::new(); fields.len() + 1];
    for (index, parent) in parents.iter().enumerate() { children[parent.unwrap_or(root)].push(index); }
    for siblings in &mut children {
        // An empty sibling at a shared start precedes the nonempty sibling;
        // an empty child at its parent's end is instead ordered by the tree.
        siblings.sort_by_key(|&index| (fields[index].start, fields[index].start != fields[index].end,
            fields[index].begin_sequence.unwrap_or(u64::from(fields[index].bookmark)), fields[index].bookmark));
        for pair in siblings.windows(2) {
            if fields[pair[0]].end > fields[pair[1]].start {
                return Err(ExportError::invalid("正文字段兄弟结果范围交叉或声明了错误的父身份。"));
            }
        }
    }
    let mut events = Vec::with_capacity(fields.len().saturating_mul(2));
    let mut pending: Vec<(usize, bool)> = children[root].iter().rev().map(|&index| (index, false)).collect();
    while let Some((index, closing)) = pending.pop() {
        events.push((index, closing));
        if !closing {
            pending.push((index, true));
            pending.extend(children[index].iter().rev().map(|&child| (child, false)));
        }
    }
    Ok(FieldTopology { depths, events })
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    fn field(id: &str, start: usize, end: usize, parent: Option<Value>, sequence: u32) -> Field {
        let mut value = json!({ "id": id, "metadata": { "source": {} } });
        if let Some(parent) = parent { value["metadata"]["source"]["parentFieldId"] = parent; }
        Field { value, start, end, first: 0, last: 0, local_start: start, local_end: end,
            bookmark: sequence, bookmark_name: id.into(), begin_sequence: None }
    }

    fn event_ids(fields: &[Field]) -> Vec<(&str, bool)> {
        resolve(fields).unwrap().events.into_iter().map(|(index, closing)|
            (fields[index].value["id"].as_str().unwrap(), closing)).collect()
    }

    #[test]
    fn empty_children_at_both_result_edges_stay_inside_their_parent() {
        let fields = vec![field("parent", 1, 4, Some(Value::Null), 0),
            field("first", 1, 1, Some(json!("parent")), 1), field("last", 4, 4, Some(json!("parent")), 2)];
        assert_eq!(event_ids(&fields), vec![("parent", false), ("first", false), ("first", true),
            ("last", false), ("last", true), ("parent", true)]);
        assert_eq!(resolve(&fields).unwrap().depths, vec![0, 1, 1]);
    }

    #[test]
    fn explicit_parent_identity_orders_equal_and_empty_ranges_even_when_child_is_listed_first() {
        for (start, end) in [(0, 0), (1, 4)] {
            let fields = vec![field("inner", start, end, Some(json!("outer")), 0),
                field("outer", start, end, Some(Value::Null), 1)];
            assert_eq!(event_ids(&fields), vec![("outer", false), ("inner", false), ("inner", true), ("outer", true)]);
        }
    }

    #[test]
    fn explicit_empty_roots_remain_adjacent_and_new_fields_can_infer_geometric_parents() {
        let fields = vec![field("empty", 0, 0, Some(Value::Null), 1), field("outer", 0, 4, None, 0),
            field("inner", 1, 3, None, 2), field("next", 4, 5, None, 3)];
        assert_eq!(event_ids(&fields), vec![("empty", false), ("empty", true), ("outer", false),
            ("inner", false), ("inner", true), ("outer", true), ("next", false), ("next", true)]);
    }

    #[test]
    fn invalid_parent_identity_cycle_containment_and_overlapping_explicit_roots_are_rejected() {
        for fields in [
            vec![field("a", 0, 1, Some(json!("missing")), 0)],
            vec![field("a", 0, 1, Some(json!("a")), 0)],
            vec![field("a", 0, 1, Some(json!("b")), 0), field("b", 0, 1, Some(json!("a")), 1)],
            vec![field("a", 0, 1, Some(Value::Null), 0), field("b", 0, 2, Some(json!("a")), 1)],
            vec![field("a", 0, 4, Some(Value::Null), 0), field("b", 1, 3, Some(Value::Null), 1)],
        ] { assert!(resolve(&fields).is_err()); }
    }
}
