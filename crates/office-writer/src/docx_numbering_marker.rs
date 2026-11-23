use std::collections::{BTreeMap, HashMap, HashSet};

use serde_json::{Map, Value, json};

use super::{ContentControl, get, integer, numbering::numbering_level};
use super::source_xml::{Content, Element, WORD_NAMESPACE, XML_NAMESPACE, parse_xml, valid_name, valid_xml_characters};
use super::source_styles::SourceStyles;
use crate::{ExportError, ExportWarning, xml};

const BOOKMARK_PREFIX: &str = "office_list_v1_";
const XML_LIMIT: usize = 16 * 1024 * 1024;

#[derive(Default)]
struct FieldState { separated: bool, code: String }

#[derive(Clone)]
struct Marker {
    elements: Vec<Element>,
    namespaces: BTreeMap<String, String>,
    text: String,
    sequence_names: HashSet<String>,
}

pub(crate) struct PreservedParagraph {
    marker: Option<Marker>,
    pub prefix_length: usize,
}

impl PreservedParagraph {
    pub fn replays(&self) -> bool { self.marker.is_some() }

    pub fn render(&self, events: &[(usize, String)]) -> Result<String, ExportError> {
        let Some(marker) = &self.marker else { return Ok(events.iter().map(|(_, text)| text.as_str()).collect()); };
        let mut result = String::new();
        let mut event_index = 0;
        let mut offset = 0;
        let mut fields = Vec::<FieldState>::new();
        emit_events(events, &mut event_index, 0, &mut result);
        for element in &marker.elements {
            if element.local != "r" { result.push_str(&element.render(Some(&marker.namespaces))); continue; }
            let properties = element.children().find(|element| element.local == "rPr").map(|element| element.render(None)).unwrap_or_default();
            let mut buffer = String::new();
            for child in element.children().filter(|element| element.local != "rPr") {
                let visible = visible_text(child, &fields);
                update_fields(child, &mut fields, &mut HashSet::new()).map_err(ExportError::invalid)?;
                let Some(text) = visible else { buffer.push_str(&child.render(None)); continue; };
                let end = offset + text.encode_utf16().count();
                let mut start_byte = 0;
                let mut start_offset = offset;
                while let Some((event_offset, _)) = events.get(event_index).filter(|(at, _)| *at < end) {
                    if *event_offset < start_offset { return Err(ExportError::invalid("编号域中的批注位置无效。")); }
                    let end_byte = utf16_byte(&text, *event_offset - offset).ok_or_else(|| ExportError::invalid("编号域中的批注不能分割 UTF-16 字符。"))?;
                    if end_byte > start_byte { buffer.push_str(&render_visible(child, &text[start_byte..end_byte])); }
                    flush_run(element, &properties, &mut buffer, &marker.namespaces, &mut result);
                    emit_events(events, &mut event_index, *event_offset, &mut result);
                    start_byte = end_byte;
                    start_offset = *event_offset;
                }
                if start_byte < text.len() { buffer.push_str(&render_visible(child, &text[start_byte..])); }
                offset = end;
            }
            flush_run(element, &properties, &mut buffer, &marker.namespaces, &mut result);
        }
        emit_events(events, &mut event_index, offset, &mut result);
        if offset != self.prefix_length || event_index != events.len() { return Err(ExportError::invalid("编号域可见结果与批注范围不一致。")); }
        Ok(result)
    }

    pub fn body_runs(&self, runs: &[Value]) -> Result<Vec<Value>, ExportError> {
        let mut remaining = self.prefix_length;
        let mut result = Vec::new();
        for run in runs {
            if remaining == 0 { result.push(run.clone()); continue; }
            let text = get(run, "text").as_str().unwrap_or_default();
            let length = text.encode_utf16().count();
            if remaining >= length { remaining -= length; continue; }
            let byte = utf16_byte(text, remaining).ok_or_else(|| ExportError::invalid("编号缓存前缀不能分割 UTF-16 字符。"))?;
            let mut run = run.clone();
            run["text"] = json!(&text[byte..]);
            result.push(run);
            remaining = 0;
        }
        if remaining != 0 { return Err(ExportError::invalid("编号域的缓存前缀超出了段落内容。")); }
        Ok(result)
    }
}

#[derive(Default)]
pub(crate) struct PreservedMarkerPlan {
    pub paragraphs: HashMap<usize, PreservedParagraph>,
    pub next_bookmark: u32,
    pub sequence_prefix: String,
    pub bookmark_names: HashSet<String>,
    pub warnings: Vec<ExportWarning>,
}

fn array(value: &Value) -> &[Value] { value.as_array().map_or(&[], Vec::as_slice) }

pub(super) fn numbering_contract(block: &Value) -> Value {
    let metadata = get(block, "metadata");
    let listed = matches!(get(block, "type").as_str(), Some("numbered" | "bullet" | "checklist"));
    let has_numbering = listed || matches!(get(block, "type").as_str(), Some("paragraph" | "heading" | "quote")) &&
        (get(block, "list").is_object() || get(metadata, "numberingDefinition").is_object() || integer(get(get(metadata, "numbering"), "numId")).is_some_and(|id| id > 0));
    let kind = if listed { get(block, "type").clone() } else if has_numbering { json!("numbered") } else { Value::Null };
    json!({ "kind": kind, "level": has_numbering.then(|| numbering_level(block)),
        "list": get(block, "list"), "checklist": get(block, "checklist"),
        "definition": get(metadata, "numberingDefinition"), "listId": get(metadata, "numberingListId"), "numbering": get(metadata, "numbering") })
}

fn appearance(run: &Value) -> Value {
    let mut properties = Map::new();
    for name in ["bold", "italic", "underline", "strike", "fontFamily", "eastAsiaFontFamily", "fontSize", "color", "verticalAlign", "href", "footnoteId", "endnoteId"] {
        properties.insert(name.to_owned(), get(run, name).clone());
    }
    let mut metadata = Map::new();
    for name in ["highlight", "verticalAlign", "href", "link", "note", "footnoteId", "endnoteId", "latexFormula"] {
        metadata.insert(name.to_owned(), get(get(run, "metadata"), name).clone());
    }
    properties.insert("metadata".to_owned(), Value::Object(metadata));
    Value::Object(properties)
}

pub(super) fn prefix_contract(block: &Value, text: &str) -> Option<Value> {
    let runs = array(get(block, "runs"));
    let current = runs.iter().filter_map(|run| get(run, "text").as_str()).collect::<String>();
    if !current.starts_with(text) { return None; }
    let mut remaining = text.len();
    let mut prefix = Vec::<Value>::new();
    for run in runs {
        if remaining == 0 { break; }
        let metadata = get(run, "metadata");
        if !get(run, "footnoteId").is_null() || !get(run, "endnoteId").is_null() || ["footnoteId", "endnoteId", "note", "latexFormula"].iter().any(|name| !get(metadata, name).is_null()) { return None; }
        let text = get(run, "text").as_str().unwrap_or_default();
        let length = remaining.min(text.len());
        if !text.is_char_boundary(length) { return None; }
        if length == 0 { continue; }
        let properties = appearance(run);
        if let Some(previous) = prefix.last_mut().filter(|part| get(part, "properties") == &properties) {
            previous["text"] = json!(format!("{}{}", get(previous, "text").as_str().unwrap_or_default(), &text[..length]));
        } else { prefix.push(json!({ "text": &text[..length], "properties": properties })); }
        remaining -= length;
    }
    (remaining == 0).then(|| json!(prefix))
}

fn parse_fragment(marker: &Value) -> Result<Marker, String> {
    let source = get(marker, "xml").as_str().ok_or("编号域缺少原始 XML。")?;
    if source.len() > XML_LIMIT { return Err("编号域 XML 超过 16 MiB。".into()); }
    let mut namespaces = BTreeMap::from([("w".to_owned(), WORD_NAMESPACE.to_owned())]);
    if let Some(values) = get(marker, "namespaces").as_object() {
        if values.len() > 128 { return Err("编号域 XML 命名空间过多。".into()); }
        for (prefix, namespace) in values {
            let value = namespace.as_str().ok_or("编号域 XML 命名空间无效。")?;
            if (!prefix.is_empty() && !valid_name(prefix)) || prefix == "xmlns" || value.len() > 2048 || value.chars().any(char::is_control) || (prefix == "xml" && value != XML_NAMESPACE) {
                return Err("编号域 XML 命名空间无效。".into());
            }
            namespaces.insert(prefix.clone(), value.to_owned());
        }
    }
    let mut wrapper = "<officeMarker".to_owned();
    for (prefix, value) in &namespaces { wrapper.push_str(&xml::attr(&if prefix.is_empty() { "xmlns".to_owned() } else { format!("xmlns:{prefix}") }, value)); }
    wrapper.push('>');
    wrapper.push_str(source);
    wrapper.push_str("</officeMarker>");
    let root = parse_xml(&wrapper, false)?;
    let elements = root.children().cloned().collect::<Vec<_>>();
    if root.local != "officeMarker" || root.content.iter().any(|content| matches!(content, Content::Text(text) if !text.trim().is_empty())) { return Err("编号域 XML 片段无效。".into()); }
    let mut result = Marker { elements, namespaces, text: String::new(), sequence_names: HashSet::new() };
    validate_marker(&mut result, marker)?;
    remove_imported_comments(&mut result.elements);
    Ok(result)
}

fn validate_attributes(element: &Element) -> Result<(), String> {
    for attribute in &element.attributes {
        if !(attribute.namespace.is_empty() || attribute.namespace == WORD_NAMESPACE || attribute.namespace == XML_NAMESPACE || attribute.namespace == "http://www.w3.org/2000/xmlns/" || attribute.namespace.starts_with("http://schemas.microsoft.com/office/word/")) {
            return Err("编号域依赖尚未保留的 Office 关系或扩展属性。".into());
        }
        if !valid_xml_characters(&attribute.value) { return Err("编号域含无效 XML 字符。".into()); }
    }
    Ok(())
}

fn validate_properties(element: &Element) -> Result<(), String> {
    if element.namespace != WORD_NAMESPACE && !element.namespace.starts_with("http://schemas.microsoft.com/office/word/") { return Err("编号域含不支持的字符属性命名空间。".into()); }
    validate_attributes(element)?;
    for child in element.children() { validate_properties(child)?; }
    Ok(())
}

fn is_comment(element: &Element) -> bool {
    element.namespace == WORD_NAMESPACE && matches!(element.local.as_str(), "commentRangeStart" | "commentRangeEnd" | "commentReference")
}

fn validate_comment(element: &Element) -> Result<(), String> {
    if !element.attribute("id").and_then(|id| id.parse::<i64>().ok()).is_some_and(|id| id >= 0) || element.children().next().is_some() || !element.text().trim().is_empty() {
        return Err("编号域中的原始批注标记无效。".into());
    }
    Ok(())
}

/** Only the current document comment model owns IDs and anchors in the new package. */
fn remove_imported_comments(elements: &mut Vec<Element>) {
    fn strip(element: &mut Element) {
        element.content.retain(|content| !matches!(content, Content::Element(child) if is_comment(child)));
        for content in &mut element.content { if let Content::Element(child) = content { strip(child); } }
    }
    for element in elements.iter_mut() {
        strip(element);
    }
    elements.retain(|element| !is_comment(element) && (element.local != "r" || element.children().any(|child| child.local != "rPr")));
}

fn validate_marker(marker: &mut Marker, source: &Value) -> Result<(), String> {
    let name = get(source, "bookmark").as_str().filter(|name| name.starts_with(BOOKMARK_PREFIX) && name.len() <= 40 && valid_name(name)).ok_or("编号域书签名称无效。")?;
    let id = integer(get(source, "bookmarkId")).filter(|id| *id >= 0).ok_or("编号域书签 ID 无效。")?.to_string();
    let Some(first) = marker.elements.first() else { return Err("编号域 XML 为空。".into()); };
    let last = marker.elements.last().ok_or("编号域缺少结束书签。")?;
    if first.local != "bookmarkStart" || first.attribute("name") != Some(name) || first.attribute("id") != Some(id.as_str()) || last.local != "bookmarkEnd" || last.attribute("id") != Some(id.as_str()) { return Err("编号域书签与导入记录不匹配。".into()); }
    let mut fields = Vec::<FieldState>::new();
    for (index, element) in marker.elements.iter().enumerate() {
        if element.namespace != WORD_NAMESPACE { return Err("编号域含不支持的 XML 命名空间。".into()); }
        validate_attributes(element)?;
        if index == 0 || index + 1 == marker.elements.len() {
            if element.children().next().is_some() || !element.text().trim().is_empty() { return Err("编号域书签不能包含内容。".into()); }
            continue;
        }
        if is_comment(element) { validate_comment(element)?; continue; }
        if element.local != "r" || !element.text().trim().is_empty() { return Err("编号域含不支持的段落内容。".into()); }
        let mut properties = false;
        for child in element.children() {
            if child.namespace != WORD_NAMESPACE { return Err("编号域含不支持的 XML 内容。".into()); }
            validate_attributes(child)?;
            if child.local == "rPr" {
                if properties { return Err("编号域文本运行含重复字符属性。".into()); }
                properties = true;
                validate_properties(child)?;
                continue;
            }
            if is_comment(child) { validate_comment(child)?; continue; }
            if child.children().next().is_some() { return Err("编号域含不支持的嵌套内容。".into()); }
            if let Some(text) = visible_text(child, &fields) { marker.text.push_str(&text); }
            update_fields(child, &mut fields, &mut marker.sequence_names)?;
        }
    }
    if !fields.is_empty() || get(source, "text").as_str() != Some(marker.text.as_str()) { return Err("编号域未闭合或可见结果与导入记录不匹配。".into()); }
    Ok(())
}

fn visible_text(element: &Element, fields: &[FieldState]) -> Option<String> {
    if !fields.iter().all(|field| field.separated) { return None; }
    match element.local.as_str() {
        "t" => Some(element.text().replace("\r\n", " ").replace(['\r', '\n'], " ")), "tab" => Some("\t".into()), "br" | "cr" => Some("\n".into()),
        "noBreakHyphen" => Some("\u{2011}".into()), "softHyphen" => Some("\u{ad}".into()), _ => None,
    }
}

fn update_fields(element: &Element, fields: &mut Vec<FieldState>, names: &mut HashSet<String>) -> Result<(), String> {
    match element.local.as_str() {
        "fldChar" => match element.attribute("fldCharType") {
            Some("begin") => fields.push(FieldState::default()),
            Some("separate") => {
                let field = fields.last_mut().ok_or("编号域缺少开始符。")?;
                if field.separated { return Err("编号域含重复结果分隔符。".into()); }
                field.separated = true;
            },
            Some("end") => {
                let field = fields.pop().ok_or("编号域缺少开始符。")?;
                if !field.separated { return Err("编号域缺少结果分隔符。".into()); }
                let mut tokens = field.code.trim_start().splitn(2, char::is_whitespace);
                if tokens.next().is_some_and(|token| token.eq_ignore_ascii_case("SEQ")) {
                    let remaining = tokens.next().unwrap_or_default().trim_start();
                    let name = if let Some(quoted) = remaining.strip_prefix('"') { quoted.split('"').next().unwrap_or_default() } else { remaining.split_whitespace().next().unwrap_or_default() };
                    if name.is_empty() || name.len() > 256 || name.starts_with('\\') { return Err("编号域缺少 SEQ 标识符。".into()); }
                    names.insert(name.to_ascii_lowercase());
                }
            },
            _ => return Err("编号域含未知字段分隔符。".into()),
        },
        "instrText" => fields.last_mut().filter(|field| !field.separated).ok_or("编号域指令位于字段外。")?.code.push_str(&element.text()),
        "t" | "tab" | "br" | "cr" | "noBreakHyphen" | "softHyphen" => {},
        _ => return Err(format!("编号域含不支持的内容 {}。", element.local)),
    }
    Ok(())
}

fn utf16_byte(text: &str, offset: usize) -> Option<usize> {
    let mut current = 0;
    for (byte, character) in text.char_indices() {
        if current == offset { return Some(byte); }
        current += character.len_utf16();
    }
    (current == offset).then_some(text.len())
}

fn render_visible(element: &Element, text: &str) -> String {
    if element.local != "t" { return element.render(None); }
    let mut element = element.clone();
    element.empty = false;
    element.content = vec![Content::Text(text.to_owned())];
    element.render(None)
}

fn flush_run(element: &Element, properties: &str, buffer: &mut String, namespaces: &BTreeMap<String, String>, result: &mut String) {
    if buffer.is_empty() { return; }
    result.push_str(&element.opening(Some(namespaces), false));
    result.push_str(properties);
    result.push_str(buffer);
    result.push_str(&format!("</{}>", element.name));
    buffer.clear();
}

fn emit_events(events: &[(usize, String)], index: &mut usize, offset: usize, result: &mut String) {
    while let Some((_, text)) = events.get(*index).filter(|(at, _)| *at == offset) { result.push_str(text); *index += 1; }
}

fn source_numbering_matches(marker: &Value, baseline: &Value) -> bool {
    let Some(source) = marker.get("sourceNumbering") else { return true; };
    let numbering = get(get(baseline, "numbering"), "numbering");
    match integer(get(source, "numId")) {
        Some(id) if id > 0 => integer(get(numbering, "numId")) == Some(id) && integer(get(source, "ilvl")) == integer(get(numbering, "ilvl")),
        Some(0) | None => !integer(get(numbering, "numId")).is_some_and(|id| id > 0),
        Some(_) => false,
    }
}

pub(super) fn prepare_preserved_markers(blocks: &[Value], controls: &HashMap<String, Vec<ContentControl>>, styles: &mut SourceStyles) -> Result<PreservedMarkerPlan, ExportError> {
    let mut plan = PreservedMarkerPlan { next_bookmark: 1, sequence_prefix: "office_seq_".into(), ..Default::default() };
    let mut names = HashSet::<String>::new();
    let mut sequence_names = HashSet::<String>::new();
    fn visit(blocks: &[Value], controls: &HashMap<String, Vec<ContentControl>>, styles: &mut SourceStyles, plan: &mut PreservedMarkerPlan, names: &mut HashSet<String>, sequence_names: &mut HashSet<String>) -> Result<(), ExportError> {
        for block in blocks {
            if get(block, "type").as_str() == Some("table") {
                for group in super::table_writer::content_groups(get(get(block, "metadata"), "table"))? { visit(group.blocks, controls, styles, plan, names, sequence_names)?; }
                continue;
            }
            let marker = get(get(block, "metadata"), "officeNumberingMarker");
            if marker.is_null() { continue; }
            let id = get(block, "id").as_str().unwrap_or_default();
            let reason = (|| -> Result<(), String> {
                let baseline = get(marker, "baseline");
                let text = get(marker, "text").as_str().ok_or("导入编号域缺少可见缓存。")?;
                if get(marker, "version").as_u64() != Some(1) || get(baseline, "version").as_u64() != Some(1) || !get(baseline, "numbering").is_object() || !get(baseline, "prefix").is_array() { return Err("导入编号域没有可核对的模型基线。".into()); }
                if array(get(baseline, "prefix")).iter().filter_map(|part| get(part, "text").as_str()).collect::<String>() != text { return Err("导入编号域的缓存基线不一致。".into()); }
                let prefix = prefix_contract(block, text).ok_or("编号缓存前缀已编辑。")?;
                let length = text.encode_utf16().count();
                if controls.get(id).is_some_and(|controls| controls.iter().any(|control| control.start < length)) { return Err("编号缓存与内容控件范围重叠。".into()); }
                if get(baseline, "numbering") != &numbering_contract(block) {
                    plan.paragraphs.insert(block as *const Value as usize, PreservedParagraph { marker: None, prefix_length: length });
                    plan.warnings.push(ExportWarning::new("DOCX_NUMBERING_MARKER_REPLACED", format!("段落 {id:?} 已更换编号；原始域和未编辑的缓存标签按当前编号模型替换。")));
                    return Ok(());
                }
                if &prefix != get(baseline, "prefix") { return Err("编号缓存的字符格式已编辑。".into()); }
                if !source_numbering_matches(marker, baseline) { return Err("编号域的原始段落编号与模型基线不一致。".into()); }
                if !matches!(get(block, "type").as_str(), Some("paragraph" | "heading" | "quote" | "numbered" | "bullet" | "checklist")) || get(get(get(block, "metadata"), "codeBlock"), "source").is_string() { return Err("当前段落内容已改变为其他节点类型。".into()); }
                let mut parsed = parse_fragment(marker)?;
                let bookmark_id = plan.next_bookmark;
                plan.next_bookmark = plan.next_bookmark.checked_add(1).ok_or("编号域书签 ID 超出上限。")?;
                let original_name = get(marker, "bookmark").as_str().unwrap_or_default();
                let name = if names.insert(original_name.to_ascii_lowercase()) { original_name.to_owned() } else {
                    let mut index = bookmark_id;
                    loop {
                        let name = format!("{BOOKMARK_PREFIX}copy_{index}");
                        if names.insert(name.to_ascii_lowercase()) { break name; }
                        index = index.checked_add(1).ok_or("编号域书签名称超出上限。")?;
                    }
                };
                for element in &mut parsed.elements {
                    if matches!(element.local.as_str(), "bookmarkStart" | "bookmarkEnd") {
                        for attribute in &mut element.attributes {
                            if attribute.namespace == WORD_NAMESPACE && attribute.local == "id" { attribute.value = bookmark_id.to_string(); }
                            if element.local == "bookmarkStart" && attribute.namespace == WORD_NAMESPACE && attribute.local == "name" { attribute.value = name.clone(); }
                        }
                    }
                }
                sequence_names.extend(parsed.sequence_names.iter().cloned());
                if let Err(reason) = styles.restore_marker(&mut parsed.elements, &parsed.namespaces, block) {
                    plan.warnings.push(ExportWarning::new("DOCX_NUMBERING_MARKER_STYLE_CONTEXT", format!("段落 {id:?} 的编号域样式上下文未完整恢复：{reason} 原始资源仍保留在文档模型中。")));
                }
                plan.paragraphs.insert(block as *const Value as usize, PreservedParagraph { marker: Some(parsed), prefix_length: length });
                Ok(())
            })();
            if let Err(reason) = reason { plan.warnings.push(ExportWarning::new("DOCX_NUMBERING_MARKER_NOT_REPLAYED", format!("段落 {id:?} 的原始编号域未重放：{reason} 当前文本保留，原始 XML 仍在文档模型中。"))); }
        }
        Ok(())
    }
    visit(blocks, controls, styles, &mut plan, &mut names, &mut sequence_names)?;
    plan.bookmark_names = names;
    let mut index = 0u32;
    while sequence_names.iter().any(|name| name.starts_with(&plan.sequence_prefix)) {
        index = index.checked_add(1).ok_or_else(|| ExportError::invalid("编号域计数器名称超出上限。"))?;
        plan.sequence_prefix = format!("office_seq_r{index}_");
    }
    Ok(plan)
}

#[cfg(test)]
mod tests {
    use std::io::{Cursor, Read};
    use quick_xml::{Reader, events::Event};
    use zip::ZipArchive;
    use super::*;

    fn field_marker(code: &str, result: &str) -> String {
        format!("<w:bookmarkStart w:id=\"99\" w:name=\"office_list_v1_fixture\"/><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText xml:space=\"preserve\">{}</w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:t xml:space=\"preserve\">{}</w:t></w:r><w:r><w:fldChar w:fldCharType=\"end\"/></w:r><w:r><w:t xml:space=\"preserve\"> </w:t></w:r><w:bookmarkEnd w:id=\"99\"/>", xml::escape(code), xml::escape(result))
    }

    fn cached_block(source: &str, text: &str) -> Value {
        let mut block = json!({ "id": "cached", "type": "paragraph", "level": 0,
            "runs": [{ "id": "cached-run", "text": format!("{text}Body") }], "metadata": {} });
        let baseline = json!({ "version": 1, "numbering": numbering_contract(&block), "prefix": prefix_contract(&block, text).unwrap() });
        block["metadata"]["officeNumberingMarker"] = json!({ "version": 1, "bookmark": "office_list_v1_fixture", "bookmarkId": 99,
            "xml": source, "text": text, "baseline": baseline });
        block
    }

    fn numbered_block() -> Value {
        let levels = json!([{ "ilvl": 0, "start": 1, "format": "decimal", "text": "%1.", "suffix": "space" }]);
        json!({ "id": "new-list", "type": "numbered", "level": 0, "runs": [{ "text": "New item" }], "metadata": {
            "numberingListId": "new-list", "numberingDefinition": { "abstract": { "abstractNumId": 7, "consecutiveNumbering": true, "levels": levels },
            "instance": { "numId": 9, "abstractNumId": 7, "consecutiveListId": 9 }, "levels": levels }
        } })
    }

    fn write(blocks: Vec<Value>) -> (Vec<u8>, Vec<ExportWarning>) {
        crate::export_docx_bytes(&json!({ "kind": "document", "name": "preserved-numbering", "blocks": blocks }).to_string()).expect("document exports")
    }

    fn part(bytes: &[u8], name: &str) -> String {
        let mut archive = ZipArchive::new(Cursor::new(bytes)).unwrap();
        let mut result = String::new();
        archive.by_name(name).unwrap().read_to_string(&mut result).unwrap();
        result
    }

    fn bookmarks(document: &str) -> Vec<(String, String)> {
        let mut reader = Reader::from_str(document);
        let mut result = Vec::new();
        loop {
            match reader.read_event().unwrap() {
                Event::Empty(tag) | Event::Start(tag) if tag.local_name().as_ref() == b"bookmarkStart" => {
                    let mut id = String::new();
                    let mut name = String::new();
                    for attribute in tag.attributes() {
                        let attribute = attribute.unwrap();
                        let value = String::from_utf8(attribute.value.to_vec()).unwrap();
                        match attribute.key.local_name().as_ref() { b"id" => id = value, b"name" => name = value, _ => {} }
                    }
                    result.push((id, name));
                }
                Event::Eof => break,
                _ => {},
            }
        }
        result
    }

    #[test]
    fn replayed_copies_and_new_numbering_have_distinct_bookmarks_and_sequence_names() {
        let original = cached_block(&field_marker(" SEQ office_seq_1 \\r 0 ", "0"), "0 ");
        let mut copied = original.clone();
        copied["id"] = json!("copied");
        let (bytes, warnings) = write(vec![original, copied, numbered_block()]);
        assert!(warnings.is_empty(), "{warnings:?}");
        let document = part(&bytes, "word/document.xml");
        let bookmarks = bookmarks(&document);
        assert_eq!(bookmarks.len(), 3);
        assert_eq!(bookmarks.iter().map(|(id, _)| id).collect::<HashSet<_>>().len(), 3);
        assert_eq!(bookmarks.iter().map(|(_, name)| name.to_ascii_lowercase()).collect::<HashSet<_>>().len(), 3);
        assert!(bookmarks.iter().all(|(_, name)| name.len() <= 40));
        assert_eq!(document.matches("SEQ office_seq_1 \\r 0").count(), 2);
        assert!(document.contains("SEQ office_seq_r1_1 \\r 1 \\h"));
        assert_eq!(document.matches("Body</w:t>").count(), 2);
        assert_eq!(document.matches("New item</w:t>").count(), 1);
    }

    #[test]
    fn only_an_initial_seq_command_reserves_an_identifier_including_quoted_and_case_insensitive_names() {
        for (code, expected) in [
            (" IF 1 = 1 \"SEQ office_seq_1\" \"other\" ", "SEQ office_seq_1 \\r 1 \\h"),
            (" seq \"OFFICE_SEQ_1\" \\c ", "SEQ office_seq_r1_1 \\r 1 \\h"),
        ] {
            let (bytes, warnings) = write(vec![cached_block(&field_marker(code, "0"), "0 "), numbered_block()]);
            assert!(warnings.is_empty(), "{warnings:?}");
            assert!(part(&bytes, "word/document.xml").contains(expected), "{code}");
        }
    }

    #[test]
    fn invalid_replay_xml_preserves_current_text_and_reports_a_diagnostic() {
        let original = field_marker(" SEQ office_seq_1 \\c ", "0");
        for invalid in [
            original.replace("<w:t xml:space=\"preserve\">0</w:t>", "<w:t>0</w:p>"),
            original.replace("w:", "q:"),
            original.replacen("<w:r>", "<w:r xmlns:w=\"urn:wrong-word-namespace\">", 1),
            original.replace(">0</w:t>", ">&unknown;</w:t>"),
            original.replace(">0</w:t>", ">&#0;</w:t>"),
            original.replace(">0</w:t>", ">&#xFFFE;</w:t>"),
            original.replacen("w:id=\"99\"", "w:id=\"99\" w:id=\"99\"", 1),
            original.replacen("w:id=\"99\"", "w:id=\"&#0;\"", 1),
            format!("<!DOCTYPE marker [<!ENTITY custom '0'>]>{original}"),
        ] {
            let (bytes, warnings) = write(vec![cached_block(&invalid, "0 ")]);
            assert!(warnings.iter().any(|warning| warning.code == "DOCX_NUMBERING_MARKER_NOT_REPLAYED"), "{invalid}");
            let document = part(&bytes, "word/document.xml");
            assert!(!document.contains("office_list_v1_fixture"));
            assert!(!document.contains("<w:fldChar"));
            assert!(document.contains("0 Body</w:t>"));
        }
        let mut wrong_namespace = cached_block(&original, "0 ");
        wrong_namespace["metadata"]["officeNumberingMarker"]["namespaces"] = json!({ "w": "urn:wrong" });
        let (_, warnings) = write(vec![wrong_namespace]);
        assert!(warnings.iter().any(|warning| warning.code == "DOCX_NUMBERING_MARKER_NOT_REPLAYED"));
    }

    #[test]
    fn valid_entities_replay_and_preserved_character_style_references_have_an_explicit_context_warning() {
        let source = field_marker(" SEQ office_seq_1 \\c ", "&")
            .replace("<w:r><w:t", "<w:r><w:rPr><w:rStyle w:val=\"CustomMarker\"/></w:rPr><w:t");
        let (bytes, warnings) = write(vec![cached_block(&source, "& ")]);
        assert_eq!(warnings.len(), 1);
        assert_eq!(warnings[0].code, "DOCX_NUMBERING_MARKER_STYLE_CONTEXT");
        let document = part(&bytes, "word/document.xml");
        assert!(document.contains("<w:rStyle w:val=\"CustomMarker\"/>"));
        assert!(document.contains("&amp;</w:t>"));
        assert!(!part(&bytes, "word/styles.xml").contains("CustomMarker"));
    }

    #[test]
    fn a_changed_source_numbering_contract_blocks_replay_while_explicit_new_numbering_replaces_the_cached_label() {
        let mut block = cached_block(&field_marker(" SEQ office_seq_1 \\c ", "0"), "0 ");
        block["metadata"]["officeNumberingMarker"]["sourceNumbering"] = json!({ "numId": 42, "ilvl": 0 });
        let (bytes, warnings) = write(vec![block.clone()]);
        assert!(warnings.iter().any(|warning| warning.code == "DOCX_NUMBERING_MARKER_NOT_REPLAYED"));
        assert!(!part(&bytes, "word/document.xml").contains("<w:fldChar"));
        block["type"] = json!("numbered");
        block["list"] = json!({ "id": "explicit-new-list", "preset": "ORDER_LIST" });
        let (bytes, warnings) = write(vec![block]);
        assert_eq!(warnings.len(), 1);
        assert_eq!(warnings[0].code, "DOCX_NUMBERING_MARKER_REPLACED");
        let document = part(&bytes, "word/document.xml");
        assert!(document.contains("<w:numPr>"));
        assert!(!document.contains("<w:fldChar"));
        assert!(!document.contains("0 Body"));
        assert_eq!(document.matches("Body</w:t>").count(), 1);
    }

    #[test]
    fn comment_boundaries_inside_a_replayed_surrogate_pair_reject_the_entire_export() {
        let block = cached_block(&field_marker(" SEQ office_seq_1 \\c ", "😀"), "😀 ");
        let model = json!({ "kind": "document", "name": "invalid-marker-comment", "blocks": [block], "comments": {
            "bad-boundary": { "anchor": [{ "blockId": "cached", "startOffset": 1, "endOffset": 2 }],
                "messages": [{ "authorName": "Author", "text": "Inside surrogate pair" }] }
        } });
        let error = crate::export_docx_bytes(&model.to_string()).expect_err("a split surrogate pair must not produce a file");
        assert_eq!(error.code, "INVALID_MODEL");
        assert!(error.message.contains("UTF-16"));
    }
}
