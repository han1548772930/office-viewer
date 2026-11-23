//! Native Word fields share paragraph UTF-16 ranges with comments and SDTs.
//! The current runs are authoritative results; cachedResult is import provenance.
use std::collections::{BTreeSet, HashMap};

use serde_json::{Value, json};

use super::{WriteContext, get, slice_utf16, source_xml, xml, xml_attribute};
use crate::ExportError;

const IDENTITY_PART: &str = "customXml/office-fields.xml";
const IDENTITY_NAMESPACE: &str = "urn:office-studio:document-fields:1";

pub(super) struct Boundary {
    pub offset: usize,
    pub priority: u8,
    pub order: i64,
    pub xml: String,
}

struct Paragraph<'a> { block: &'a Value, id: String, text: String, absolute: usize }
struct Field {
    value: Value,
    start: usize,
    end: usize,
    first: usize,
    last: usize,
    local_start: usize,
    local_end: usize,
    bookmark: u32,
    bookmark_name: String,
}

#[derive(Clone)]
struct LanguageRange { start: usize, end: usize, depth: usize, language: String }

#[derive(Default)]
pub(super) struct TextFieldPlan {
    by_block: HashMap<String, Vec<Boundary>>,
    languages: HashMap<String, Vec<LanguageRange>>,
    identities: Vec<(String, String)>,
    pub has_fields: bool,
}

fn offset(value: &Value, name: &str) -> Result<usize, ExportError> {
    get(value, name).as_u64().filter(|value| *value <= 9_007_199_254_740_991).and_then(|value| usize::try_from(value).ok())
        .ok_or_else(|| ExportError::invalid(format!("正文字段 {name} 必须为非负安全整数。")))
}

fn collect<'a>(blocks: &'a [Value], output: &mut Vec<Paragraph<'a>>, absolute: &mut usize) -> Result<(), ExportError> {
    for block in blocks {
        let metadata = get(block, "metadata");
        if !get(metadata, "nativeFieldSource").is_null() {
            return Err(ExportError::invalid("正文包含未映射的原生字段；来源已保留，需补齐范围映射后才能导出 DOCX。"));
        }
        if get(block, "type").as_str() == Some("table") {
            for row in get(get(metadata, "table"), "rows").as_array().into_iter().flatten() {
                for cell in get(row, "cells").as_array().into_iter().flatten() {
                    collect(get(cell, "blocks").as_array().map(Vec::as_slice).unwrap_or_default(), output, absolute)?;
                }
            }
        } else if !matches!(get(block, "type").as_str(), Some("image" | "divider" | "pageBreak" | "columnBreak" | "sectionBreak")) {
            let id = get(block, "id").as_str().map(str::to_owned).unwrap_or_else(|| format!("__office-field-paragraph-{}", output.len()));
            let text = get(block, "runs").as_array().into_iter().flatten().map(|run| get(run, "text").as_str().unwrap_or_default()).collect::<String>();
            let length = text.encode_utf16().count();
            output.push(Paragraph { block, id, text, absolute: *absolute });
            *absolute = absolute.checked_add(length).and_then(|value| value.checked_add(1)).ok_or_else(|| ExportError::invalid("正文字段范围超出索引容量。"))?;
        } else if get(metadata, "documentTextFields").get("fields").and_then(Value::as_array).is_some_and(|fields| !fields.is_empty()) {
            return Err(ExportError::invalid("非文本对象不能持有正文字段结果范围。"));
        }
    }
    Ok(())
}

fn boolean(value: &Value, name: &str) -> Result<bool, ExportError> {
    match value.get(name) {
        Some(Value::Bool(value)) => Ok(*value),
        None => Ok(false),
        _ => Err(ExportError::invalid(format!("正文字段 {name} 必须为布尔值。"))),
    }
}

fn instruction(context: &mut WriteContext<'_>, block: &Value, field: &Value) -> Result<String, ExportError> {
    let name = get(field, "field").as_str().unwrap_or_default();
    let current = get(field, "instruction").as_str().filter(|value| !value.trim().is_empty()).unwrap_or(name).trim();
    let source = get(get(field, "metadata"), "source");
    let preserved = get(source, "instructionXml").as_str().filter(|value| !value.is_empty());
    if let Some(preserved) = preserved.filter(|_| get(source, "instruction").as_str().is_some_and(|value| value.trim() == current)) {
        let namespaces = get(source, "instructionNamespaces").as_object().into_iter().flatten().map(|(prefix, namespace)| {
            let namespace = namespace.as_str().ok_or_else(|| ExportError::invalid("字段指令命名空间无效。"))?;
            if prefix.is_empty() || prefix == "w" || prefix == "r" { return Ok(String::new()); }
            if !prefix.bytes().all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b'.')) { return Err(ExportError::invalid("字段指令命名空间前缀无效。")); }
            Ok(format!(" xmlns:{prefix}=\"{}\"", xml_attribute(namespace)))
        }).collect::<Result<String, ExportError>>()?;
        let mut root = source_xml::parse_xml(&format!("<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"{namespaces}>{preserved}</w:p>"), true).map_err(ExportError::invalid)?;
        context.source_styles.restore_story(&mut root).map_err(ExportError::invalid)?;
        // Original code run styles are resolved through the same source-style
        // adapter as story fields, including language and theme references.
        return Ok(root.children().filter(|element| !element.is_word("pPr")).map(|element| element.render(None)).collect());
    }
    let _ = block;
    Ok(format!("<w:r><w:instrText xml:space=\"preserve\"> {} </w:instrText></w:r>", super::xml_text(current)))
}

impl TextFieldPlan {
    pub fn prepare(context: &mut WriteContext<'_>, blocks: &[Value]) -> Result<Self, ExportError> {
        fn has_fields(blocks: &[Value]) -> bool {
            blocks.iter().any(|block| {
                let metadata = get(block, "metadata");
                metadata.get("documentTextFields").is_some() || metadata.get("nativeFieldSource").is_some()
                    || get(get(metadata, "table"), "rows").as_array().into_iter().flatten().any(|row| get(row, "cells").as_array().into_iter().flatten().any(|cell| has_fields(get(cell, "blocks").as_array().map(Vec::as_slice).unwrap_or_default())))
            })
        }
        if !has_fields(blocks) { return Ok(Self::default()); }
        let mut paragraphs = Vec::new(); let mut absolute = 0;
        collect(blocks, &mut paragraphs, &mut absolute)?;
        let mut by_id = HashMap::new();
        for (index, paragraph) in paragraphs.iter().enumerate() {
            if by_id.insert(paragraph.id.clone(), index).is_some() { return Err(ExportError::invalid("正文字段段落标识重复。")); }
        }
        let mut fields = Vec::new(); let mut ids = BTreeSet::new();
        for (index, paragraph) in paragraphs.iter().enumerate() {
            let Some(collection) = get(paragraph.block, "metadata").get("documentTextFields") else { continue; };
            if get(collection, "version").as_u64() != Some(1) { return Err(ExportError::invalid("正文字段集合版本无效。")); }
            let entries = get(collection, "fields").as_array().ok_or_else(|| ExportError::invalid("正文字段集合必须为数组。"))?;
            for entry in entries {
                let id = get(entry, "id").as_str().filter(|id| !id.trim().is_empty()).ok_or_else(|| ExportError::invalid("正文字段标识缺失。"))?;
                if !ids.insert(id.to_owned()) { return Err(ExportError::invalid("正文字段标识重复。")); }
                let name = get(entry, "field").as_str().filter(|name| !name.trim().is_empty()).ok_or_else(|| ExportError::invalid("正文字段类型缺失。"))?;
                for key in ["instruction", "cachedResult", "language"] { if entry.get(key).is_some_and(|value| !value.is_string()) { return Err(ExportError::invalid(format!("字段 {id:?} 的 {key} 无效。"))); } }
                boolean(entry, "locked")?; boolean(entry, "containsNestedFields")?;
                if entry.get("metadata").is_some_and(|value| !value.is_object()) { return Err(ExportError::invalid("字段来源必须为对象。")); }
                let source = get(get(entry, "metadata"), "source");
                if get(source, "nestedInInstruction").as_bool() == Some(true) || get(source, "hasInstructionText").as_bool() == Some(true) {
                    return Err(ExportError::invalid("字段指令区包含嵌套结果；当前来源已保留，需补齐隐藏代码文本的范围映射后才能导出。"));
                }
                let last = match entry.get("endBlockId") {
                    None => index,
                    Some(Value::String(id)) if !id.trim().is_empty() => *by_id.get(id).ok_or_else(|| ExportError::invalid("跨段字段的末段不存在。"))?,
                    _ => return Err(ExportError::invalid("跨段字段末段标识无效。")),
                };
                let from = offset(entry, "startOffset")?; let to = offset(entry, "endOffset")?;
                if slice_utf16(&paragraph.text, from, from).is_none() || slice_utf16(&paragraphs[last].text, to, to).is_none() { return Err(ExportError::invalid("正文字段结果范围无效或拆分了 UTF-16 代理对。")); }
                let start = paragraph.absolute + from; let end = paragraphs[last].absolute + to;
                if end < start { return Err(ExportError::invalid("正文字段结果终点早于起点。")); }
                let bookmark = context.sequences.next_bookmark;
                context.sequences.next_bookmark = bookmark.checked_add(1).ok_or_else(|| ExportError::invalid("字段书签数量超出范围。"))?;
                let bookmark_name = format!("office_field_v1_{bookmark:08x}");
                fields.push(Field { value: entry.clone(), start, end, first: index, last, local_start: from, local_end: to, bookmark, bookmark_name });
                let _ = name;
            }
        }
        fields.sort_by_key(|field| (field.start, std::cmp::Reverse(field.end), field.bookmark));
        let mut enclosing: Vec<usize> = Vec::new(); let mut depths = Vec::new();
        for (index, field) in fields.iter().enumerate() {
            while enclosing.last().is_some_and(|&parent| fields[parent].end <= field.start) { enclosing.pop(); }
            if enclosing.last().is_some_and(|&parent| field.end > fields[parent].end) { return Err(ExportError::invalid("正文字段结果范围部分交叉。")); }
            depths.push(enclosing.len());
            if field.start < field.end { enclosing.push(index); }
        }
        let mut result = Self { has_fields: !fields.is_empty(), ..Self::default() };
        for (index, field) in fields.iter().enumerate() {
            let id = get(&field.value, "id").as_str().unwrap_or_default();
            let locked = boolean(&field.value, "locked")?;
            let lock_xml = if locked { " w:fldLock=\"1\"" } else { "" };
            let code = instruction(context, paragraphs[field.first].block, &field.value)?;
            let opening = format!("<w:bookmarkStart w:id=\"{}\" w:name=\"{}\"/><w:r><w:fldChar w:fldCharType=\"begin\"{lock_xml}/></w:r>{code}<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>", field.bookmark, field.bookmark_name);
            let closing = format!("<w:r><w:fldChar w:fldCharType=\"end\"/></w:r><w:bookmarkEnd w:id=\"{}\"/>", field.bookmark);
            if field.start == field.end {
                result.by_block.entry(paragraphs[field.first].id.clone()).or_default().push(Boundary { offset: field.local_start, priority: 4, order: index as i64, xml: format!("{opening}{closing}") });
            } else {
                result.by_block.entry(paragraphs[field.first].id.clone()).or_default().push(Boundary { offset: field.local_start, priority: 5, order: index as i64, xml: opening });
                result.by_block.entry(paragraphs[field.last].id.clone()).or_default().push(Boundary { offset: field.local_end, priority: 0, order: -(index as i64), xml: closing });
            }
            result.identities.push((field.bookmark_name.clone(), id.to_owned()));
            if let Some(language) = get(&field.value, "language").as_str().filter(|value| !value.is_empty()) {
                for paragraph in paragraphs.iter().take(field.last + 1).skip(field.first) {
                    let start = field.start.saturating_sub(paragraph.absolute).min(paragraph.text.encode_utf16().count());
                    let end = field.end.saturating_sub(paragraph.absolute).min(paragraph.text.encode_utf16().count());
                    result.languages.entry(paragraph.id.clone()).or_default().push(LanguageRange { start, end, depth: depths[index], language: language.to_owned() });
                }
            }
        }
        Ok(result)
    }

    pub fn take_boundaries(&mut self, block_id: &str) -> Vec<Boundary> { self.by_block.remove(block_id).unwrap_or_default() }

    pub fn annotate(&self, block_id: &str, runs: &[Value], prefix_length: usize) -> Result<Option<Vec<Value>>, ExportError> {
        let Some(languages) = self.languages.get(block_id) else { return Ok(None); };
        if !runs.iter().any(|run| get(get(run, "metadata"), "officeTextFieldStyleOutsideResult").as_bool() == Some(true)) { return Ok(None); }
        let mut result = Vec::new(); let mut offset = prefix_length;
        for run in runs {
            let text = get(run, "text").as_str().unwrap_or_default(); let end = offset + text.encode_utf16().count();
            if get(get(run, "metadata"), "officeTextFieldStyleOutsideResult").as_bool() != Some(true) || offset == end { result.push(run.clone()); offset = end; continue; }
            let mut points = BTreeSet::from([offset, end]);
            for range in languages { if range.start > offset && range.start < end { points.insert(range.start); } if range.end > offset && range.end < end { points.insert(range.end); } }
            let points = points.into_iter().collect::<Vec<_>>();
            for pair in points.windows(2) {
                let mut clipped = run.clone(); clipped["text"] = json!(slice_utf16(text, pair[0] - offset, pair[1] - offset).ok_or_else(|| ExportError::invalid("字段结果语言范围拆分了 UTF-16 代理对。"))?);
                if let Some(language) = languages.iter().filter(|range| range.start <= pair[0] && range.end >= pair[1]).max_by_key(|range| range.depth) {
                    clipped["metadata"]["officeFieldResultLanguage"] = json!(language.language);
                }
                if pair[0] != offset { if let Some(object) = clipped.as_object_mut() { object.remove("footnoteId"); object.remove("endnoteId"); } }
                result.push(clipped);
            }
            offset = end;
        }
        Ok(Some(result))
    }

    pub fn write_identities(&self, context: &mut WriteContext<'_>) -> Result<(), ExportError> {
        if self.identities.is_empty() { return Ok(()); }
        let entries = self.identities.iter().map(|(key, id)| format!("<field key=\"{}\" id=\"{}\"/>", xml_attribute(key), xml_attribute(id))).collect::<String>();
        context.package.add_xml(IDENTITY_PART, "application/xml", format!("{}<fields xmlns=\"{IDENTITY_NAMESPACE}\" version=\"1\">{entries}</fields>", xml::DECL))?;
        context.relationship("http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml", "../customXml/office-fields.xml", false)?;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn inherited_result_language_only_overrides_its_field_span() {
        let plan = TextFieldPlan { languages: HashMap::from([("p".into(), vec![LanguageRange { start: 1, end: 3, depth: 0, language: "fr-FR".into() }])]), ..TextFieldPlan::default() };
        let runs = vec![json!({ "text": "a😀b", "metadata": { "officeTextFieldStyleOutsideResult": true } })];
        let result = plan.annotate("p", &runs, 0).unwrap().unwrap();
        assert_eq!(result.iter().map(|run| run["text"].as_str().unwrap()).collect::<String>(), "a😀b");
        assert_eq!(result[1]["metadata"]["officeFieldResultLanguage"], "fr-FR");
        assert!(result[0]["metadata"].get("officeFieldResultLanguage").is_none());
        assert!(result[2]["metadata"].get("officeFieldResultLanguage").is_none());
    }
}
