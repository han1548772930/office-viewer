//! A story field is one editor input object, but its cached result can contain
//! nested native fields. Their offsets refer to the parent's UTF-16 cache.
//! Split the current styled cache at those boundaries instead of flattening it.
use serde_json::Value;

use super::{ExportError, Token, WriteContext, field_cache_style, field_cached_runs, get, offset, render_styled_text, slice_utf16, xml_text};

struct NestedField<'a> {
    instruction: &'a str,
    start: usize,
    end: usize,
    locked: bool,
    children: Vec<NestedField<'a>>,
}

struct CacheRun<'a> {
    start: usize,
    end: usize,
    style: Value,
    provenance: Option<&'a Value>,
}

fn flag(value: &Value, key: &str) -> Result<bool, ExportError> {
    match value.get(key) {
        None | Some(Value::Null) => Ok(false),
        Some(Value::Bool(value)) => Ok(*value),
        _ => Err(ExportError::invalid(format!("页眉页脚字段 {key} 标记无效。"))),
    }
}

pub(super) fn is_locked(field: &Value) -> Result<bool, ExportError> {
    let declared = flag(field, "locked")?;
    let source = get(get(field, "metadata"), "source");
    // An explicit editable lock is authoritative. Source is only a default.
    Ok(if field.get("locked").is_some_and(|value| !value.is_null()) { declared } else { flag(source, "locked")? })
}

fn children<'a>(declared: Option<&'a Value>, cached: &str, parent_start: usize, parent_end: usize, depth: usize, resolved: bool) -> Result<Vec<NestedField<'a>>, ExportError> {
    let Some(declared) = declared.filter(|value| !value.is_null()) else { return Ok(Vec::new()); };
    if depth > 256 { return Err(ExportError::invalid("页眉页脚嵌套字段层级超出安全范围。")); }
    let declared = declared.as_array().ok_or_else(|| ExportError::invalid("页眉页脚嵌套字段集合无效。"))?;
    let parent_length = parent_end - parent_start;
    let mut result = Vec::with_capacity(declared.len());
    let mut cursor = parent_start;
    for field in declared {
        let name = get(field, "field").as_str().filter(|name| !name.trim().is_empty()).ok_or_else(|| ExportError::invalid("页眉页脚嵌套字段类型缺失。"))?;
        let instruction = get(field, "instruction").as_str().filter(|instruction| !instruction.trim().is_empty()).ok_or_else(|| ExportError::invalid("页眉页脚嵌套字段指令缺失。"))?;
        if !instruction.split_whitespace().next().is_some_and(|first| first.eq_ignore_ascii_case(name)) {
            return Err(ExportError::invalid("页眉页脚嵌套字段指令与类型不一致。"));
        }
        let from = offset(field, "startOffset")?;
        let to = offset(field, "endOffset")?;
        if from > to || to > parent_length { return Err(ExportError::invalid("页眉页脚嵌套字段超出父字段缓存范围。")); }
        let start = parent_start.checked_add(from).ok_or_else(|| ExportError::invalid("页眉页脚嵌套字段位置溢出。"))?;
        let end = parent_start.checked_add(to).ok_or_else(|| ExportError::invalid("页眉页脚嵌套字段位置溢出。"))?;
        if start < cursor || slice_utf16(cached, start, end).is_none() {
            return Err(ExportError::invalid("页眉页脚嵌套字段交叉重叠、顺序无效或拆分了 UTF-16 代理对。"));
        }
        let nested = children(field.get("nestedFields"), cached, start, end, depth + 1, resolved)?;
        if flag(field, "containsNestedFields")? && nested.is_empty() && resolved {
            return Err(ExportError::invalid("页眉页脚嵌套字段代码未完整映射；来源已保留，不能丢弃后写出。"));
        }
        result.push(NestedField { instruction, start, end, locked: flag(field, "locked")?, children: nested });
        cursor = end;
    }
    Ok(result)
}

fn prepare(field: &Value, resolved: bool) -> Result<Vec<NestedField<'_>>, ExportError> {
    let cached = get(field, "cachedResult").as_str().unwrap_or_default();
    let source = get(get(field, "metadata"), "source");
    let nested = children(source.get("nestedFields"), cached, 0, cached.encode_utf16().count(), 0, resolved)?;
    let contains = flag(field, "containsNestedFields")? || flag(source, "containsNestedFields")?;
    if resolved && contains && nested.is_empty() {
        return Err(ExportError::invalid("页眉页脚嵌套字段代码未完整映射；来源已保留，不能丢弃后写出。"));
    }
    Ok(nested)
}

pub(super) fn validate(field: &Value) -> Result<(), ExportError> {
    is_locked(field)?;
    // Unchanged unsupported source can still be retained verbatim. A changed
    // source must contain a complete nested-field mapping before rendering.
    prepare(field, false)?;
    Ok(())
}

fn literal(context: &mut WriteContext<'_>, path: &str, cached: &str, spans: &[CacheRun<'_>], from: usize, to: usize, next_relation: &mut u32) -> Result<String, ExportError> {
    let mut result = String::new();
    if from == to {
        let run = spans.iter().rev().find(|run| run.start <= from && from <= run.end).ok_or_else(|| ExportError::invalid("页眉页脚空字段缓存缺少格式范围。"))?;
        return render_styled_text(context, path, &run.style, run.provenance.unwrap_or(&run.style), "", next_relation);
    }
    for run in spans.iter().filter(|run| run.start < to && run.end > from) {
        let start = from.max(run.start);
        let end = to.min(run.end);
        let part = slice_utf16(cached, start, end).ok_or_else(|| ExportError::invalid("页眉页脚字段缓存格式拆分了 UTF-16 代理对。"))?;
        result.push_str(&render_styled_text(context, path, &run.style, run.provenance.unwrap_or(&run.style), part, next_relation)?);
    }
    Ok(result)
}

fn range(context: &mut WriteContext<'_>, path: &str, cached: &str, spans: &[CacheRun<'_>], from: usize, to: usize, nested: &[NestedField<'_>], next_relation: &mut u32) -> Result<String, ExportError> {
    let mut result = String::new();
    let mut cursor = from;
    for field in nested {
        if cursor < field.start { result.push_str(&literal(context, path, cached, spans, cursor, field.start, next_relation)?); }
        let lock = if field.locked { " w:fldLock=\"true\"" } else { "" };
        result.push_str(&format!("<w:r><w:fldChar w:fldCharType=\"begin\"{lock}/></w:r><w:r><w:instrText xml:space=\"preserve\"> {} </w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>", xml_text(field.instruction)));
        result.push_str(&range(context, path, cached, spans, field.start, field.end, &field.children, next_relation)?);
        result.push_str("<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>");
        cursor = field.end;
    }
    if cursor < to || nested.is_empty() && from == to { result.push_str(&literal(context, path, cached, spans, cursor, to, next_relation)?); }
    Ok(result)
}

pub(super) fn render(context: &mut WriteContext<'_>, path: &str, token: &Token<'_>, next_relation: &mut u32) -> Result<String, ExportError> {
    let field = token.field.ok_or_else(|| ExportError::invalid("页眉页脚字段缓存缺少字段定义。"))?;
    let cached = get(field, "cachedResult").as_str().unwrap_or_default();
    let nested = prepare(field, true)?;
    let mut spans = Vec::new();
    let mut offset = 0;
    for run in field_cached_runs(field)?.into_iter().flatten() {
        let end = offset + get(run, "text").as_str().unwrap_or_default().encode_utf16().count();
        spans.push(CacheRun { start: offset, end, style: field_cache_style(field, run, token.style)?, provenance: None });
        offset = end;
    }
    if spans.is_empty() { spans.push(CacheRun { start: 0, end: cached.encode_utf16().count(), style: token.style.clone(), provenance: Some(field) }); }
    range(context, path, cached, &spans, 0, cached.encode_utf16().count(), &nested, next_relation)
}

#[cfg(test)]
mod tests {
    use serde_json::json;
    use super::*;

    #[test]
    fn nested_utf16_ranges_are_relative_to_each_parent_cache() {
        let field = json!({ "cachedResult": "a😀bc", "metadata": { "source": { "nestedFields": [
            { "field": "REF", "instruction": "REF Inner", "startOffset": 1, "endOffset": 5, "containsNestedFields": true,
              "nestedFields": [{ "field": "DATE", "instruction": "DATE", "startOffset": 2, "endOffset": 3 }] }
        ] } } });
        let fields = prepare(&field, true).unwrap();
        assert_eq!((fields[0].start, fields[0].end), (1, 5));
        assert_eq!((fields[0].children[0].start, fields[0].children[0].end), (3, 4));
    }

    #[test]
    fn invalid_surrogate_and_sibling_overlap_are_rejected() {
        for nested in [
            json!([{ "field": "DATE", "instruction": "DATE", "startOffset": 2, "endOffset": 3 }]),
            json!([{ "field": "DATE", "instruction": "DATE", "startOffset": 0, "endOffset": 3 }, { "field": "TIME", "instruction": "TIME", "startOffset": 1, "endOffset": 4 }]),
        ] {
            let field = json!({ "cachedResult": "a😀b", "metadata": { "source": { "nestedFields": nested } } });
            assert!(validate(&field).is_err());
        }
    }

    #[test]
    fn editable_lock_can_override_source_lock() {
        let field = json!({ "locked": false, "metadata": { "source": { "locked": true } } });
        assert!(!is_locked(&field).unwrap());
    }
}
