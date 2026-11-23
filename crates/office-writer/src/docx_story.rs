use std::collections::BTreeSet;

use serde_json::{Value, json};

use super::{REL_HYPERLINK, WriteContext, get, run_link, run_properties_xml_with_defaults, slice_utf16, source_xml, xml, xml_attribute, xml_text};
use super::default_style as default_styles;
use crate::ExportError;

#[path = "docx_story_field_cache.rs"]
mod field_cache;

pub(super) fn default_style(model: &Value) -> Value {
    let mut text = json!({ "fontFamily": "Aptos", "fontSize": 9, "color": "#1f2937", "bold": false, "italic": false, "underline": false, "strike": false });
    if let Some(declared) = get(get(model, "defaultStyle"), "text").as_object() {
        for (key, value) in declared { text[key] = value.clone(); }
    }
    text["fontSize"] = json!(9);
    json!({ "version": 1, "text": text, "paragraph": { "align": "center", "indent": 0, "lineHeight": 1.25, "spaceBefore": 0, "spaceAfter": 0 } })
}

fn snapshot(text: &str, content: &Value, fields: &[&Value]) -> Value {
    let runs = get(content, "runs").as_array().into_iter().flatten().map(|run| {
        let mut properties = run.as_object().cloned().unwrap_or_default();
        properties.remove("id"); properties.remove("metadata");
        if let Some(href) = run_link(run) { properties.insert("href".into(), json!(href)); }
        Value::Object(properties)
    }).collect::<Vec<_>>();
    let paragraphs = get(content, "paragraphs").as_array().into_iter().flatten().map(|paragraph| {
        let mut properties = paragraph.as_object().cloned().unwrap_or_default(); properties.remove("metadata"); Value::Object(properties)
    }).collect::<Vec<_>>();
    let fields = fields.iter().map(|field| {
        let mut properties = field.as_object().cloned().unwrap_or_default(); properties.remove("story"); properties.remove("metadata");
        // Cache runs are editable model data within metadata, unlike provenance.
        // They must participate in the unchanged-source snapshot comparison.
        if let Some(cache) = get(field, "metadata").get("officeStoryFieldCache") { properties.insert("cachedRuns".into(), cache.clone()); }
        Value::Object(properties)
    }).collect::<Vec<_>>();
    json!({ "text": text, "runs": runs, "paragraphs": paragraphs, "fields": fields, "locks": content.get("locks").cloned().unwrap_or_else(|| json!([])) })
}

fn attribute<'a>(node: &'a source_xml::Element, name: &str) -> Option<&'a str> {
    node.attributes.iter().find(|attribute| attribute.local == name).map(|attribute| attribute.value.as_str())
}

fn preserve_source(context: &mut WriteContext<'_>, path: &str, content_type: &str, root: &str, text: &str, fields: &[&Value], content: &Value) -> Result<bool, ExportError> {
    let source = get(get(content, "metadata"), "officeSource");
    if source.is_null() { return Ok(false); }
    if get(source, "version").as_u64() != Some(1) { return Err(ExportError::invalid("页眉页脚来源版本无效。")); }
    let formatting_cleared = get(content, "runs").as_array().into_iter().flatten().any(|run| get(get(run, "metadata"), "officeStoryTextFormatCleared").as_bool() == Some(true))
        || get(content, "paragraphs").as_array().into_iter().flatten().any(|paragraph| get(get(paragraph, "metadata"), "officeStoryParagraphFormatCleared").as_bool() == Some(true))
        || fields.iter().any(|field| get(get(field, "metadata"), "officeStoryTextFormatCleared").as_bool() == Some(true));
    let unchanged = !formatting_cleared && get(source, "baseline") == &snapshot(text, content, fields);
    let unsupported = get(source, "unsupportedFeatures").as_array().is_some_and(|features| !features.is_empty());
    if !unchanged {
        if unsupported { return Err(ExportError::invalid("此页眉页脚包含尚未接入可编辑 DOCX 写出的对象；来源数据已保留，不能丢弃后导出。")); }
        return Ok(false);
    }
    let xml = get(source, "xml").as_str().ok_or_else(|| ExportError::invalid("页眉页脚来源 XML 缺失。"))?;
    let mut node = source_xml::parse_xml(xml, true).map_err(ExportError::invalid)?;
    if !node.is_word(root) && !(node.local == root && node.namespace == "http://purl.oclc.org/ooxml/wordprocessingml/main") { return Err(ExportError::invalid("页眉页脚来源 XML 根节点无效。")); }
    context.source_styles.restore_story(&mut node).map_err(ExportError::invalid)?;
    let mut relationships = Vec::new();
    if let Some(xml) = get(source, "relationshipsXml").as_str() {
        let node = source_xml::parse_xml(xml, true).map_err(ExportError::invalid)?;
        if node.local != "Relationships" || node.namespace != "http://schemas.openxmlformats.org/package/2006/relationships" { return Err(ExportError::invalid("页眉页脚关系 XML 根节点无效。")); }
        for relation in node.children() {
            let id = attribute(relation, "Id").ok_or_else(|| ExportError::invalid("页眉页脚关系 ID 缺失。"))?;
            let kind = attribute(relation, "Type").ok_or_else(|| ExportError::invalid("页眉页脚关系类型缺失。"))?;
            let target = attribute(relation, "Target").ok_or_else(|| ExportError::invalid("页眉页脚关系目标缺失。"))?;
            if attribute(relation, "TargetMode") != Some("External") { return Err(ExportError::invalid("页眉页脚的内嵌资源写出尚未接入；来源关系已保留。")); }
            let kind = kind.strip_prefix("http://purl.oclc.org/ooxml/officeDocument/relationships/")
                .map(|suffix| format!("http://schemas.openxmlformats.org/officeDocument/2006/relationships/{suffix}")).unwrap_or_else(|| kind.to_owned());
            relationships.push((id.to_owned(), kind, target.to_owned()));
        }
    }
    for (id, kind, target) in relationships { context.package.add_relationship(Some(path), &id, &kind, &target, true)?; }
    context.package.add_xml(path, content_type, format!("{}{}", xml::DECL, node.render(None)))?;
    Ok(true)
}

fn offset(value: &Value, name: &str) -> Result<usize, ExportError> {
    get(value, name).as_u64().filter(|value| *value <= 9_007_199_254_740_991u64).and_then(|value| usize::try_from(value).ok())
        .ok_or_else(|| ExportError::invalid(format!("页眉页脚 {name} 偏移无效。")))
}

fn legacy_content(text: &str, defaults: &Value) -> Value {
    let mut run = get(defaults, "text").clone(); run["text"] = json!(text);
    let mut start = 0;
    let paragraphs = text.split('\n').map(|part| {
        let end = start + part.encode_utf16().count();
        let paragraph = json!({ "startOffset": start, "endOffset": end, "style": defaults["paragraph"] });
        start = end + 1; paragraph
    }).collect::<Vec<_>>();
    json!({ "runs": [run], "paragraphs": paragraphs })
}

struct Token<'a> { start: usize, end: usize, text: String, style: &'a Value, field: Option<&'a Value> }
struct PreparedParagraph<'a> { start: usize, end: usize, tokens: Vec<Token<'a>>, style: &'a Value, empty_style: &'a Value, provenance: &'a Value }
#[derive(Clone)]
struct Lock { id: String, start: usize, end: usize, kind: String }

fn hex_id(id: &str) -> String { id.as_bytes().iter().map(|byte| format!("{byte:02x}")).collect() }

fn open_lock(context: &mut WriteContext<'_>, lock: &Lock) -> Result<String, ExportError> {
    let numeric = context.next_story_control;
    context.next_story_control = numeric.checked_add(1).ok_or_else(|| ExportError::invalid("页眉页脚内容控件数量超出范围。"))?;
    Ok(format!("<w:sdt><w:sdtPr><w:id w:val=\"{numeric}\"/><w:tag w:val=\"office_story_lock_v1_{}\"/><w:lock w:val=\"{}\"/><w:text/></w:sdtPr><w:sdtContent>", hex_id(&lock.id), xml_attribute(&lock.kind)))
}

fn literal_tokens<'a>(tokens: &mut Vec<Token<'a>>, spans: &[(usize, usize, &'a Value)], from: usize, to: usize, buffer: &mut usize) -> Result<(), ExportError> {
    for &(start, end, style) in spans {
        let a = from.max(start);
        let b = to.min(end);
        if a >= b { continue; }
        let text = get(style, "text").as_str().unwrap_or_default();
        let value = slice_utf16(text, a - start, b - start).ok_or_else(|| ExportError::invalid("页眉页脚字符范围拆分了 UTF-16 代理对。"))?;
        tokens.push(Token { start: *buffer, end: *buffer + b - a, text: value.to_owned(), style, field: None });
        *buffer += b - a;
    }
    Ok(())
}

fn story_run_xml(context: &mut WriteContext<'_>, style: &Value, provenance: &Value, text: &str) -> Result<String, ExportError> {
    let generated = run_properties_xml_with_defaults(style, context.run_defaults);
    let properties = context.source_styles.story_run_properties(&generated, style, provenance).map_err(ExportError::invalid)?;
    Ok(text.replace("\r\n", "\n").replace('\r', "\n").split('\n').enumerate().map(|(index, part)| {
        let break_xml = if index == 0 { "" } else { "<w:br/>" };
        format!("<w:r>{properties}{break_xml}<w:t xml:space=\"preserve\">{}</w:t></w:r>", xml_text(part))
    }).collect())
}

fn render_styled_text(context: &mut WriteContext<'_>, path: &str, style: &Value, provenance: &Value, text: &str, next_relation: &mut u32) -> Result<String, ExportError> {
    let mut fragments = String::new();
    for (index, part) in text.split('\t').enumerate() {
        if index > 0 {
            let placeholder = story_run_xml(context, style, provenance, "")?;
            fragments.push_str(&placeholder.replace("<w:t xml:space=\"preserve\"></w:t>", "<w:tab/>"));
        }
        if !part.is_empty() || text.is_empty() { fragments.push_str(&story_run_xml(context, style, provenance, &part.replace('\u{2028}', "\n"))?); }
    }
    if let Some(href) = run_link(style) {
        if let Some(anchor) = href.strip_prefix('#') { return Ok(format!("<w:hyperlink w:anchor=\"{}\" w:history=\"1\">{fragments}</w:hyperlink>", xml_attribute(anchor))); }
        if !href.starts_with("https://") && !href.starts_with("http://") && !href.starts_with("mailto:") { return Err(ExportError::invalid("页眉页脚超链接目标不受支持。")); }
        let id = format!("rId{}", *next_relation);
        *next_relation = next_relation.checked_add(1).ok_or_else(|| ExportError::invalid("页眉页脚超链接数量超出范围。"))?;
        context.package.add_relationship(Some(path), &id, REL_HYPERLINK, &href, true)?;
        return Ok(format!("<w:hyperlink r:id=\"{id}\" w:history=\"1\">{fragments}</w:hyperlink>"));
    }
    Ok(fragments)
}

fn field_cached_runs(field: &Value) -> Result<Option<&Vec<Value>>, ExportError> {
    let Some(cache) = get(field, "metadata").get("officeStoryFieldCache") else { return Ok(None); };
    if get(cache, "version").as_u64() != Some(1) { return Err(ExportError::invalid("页眉页脚字段缓存版本无效。")); }
    let runs = get(cache, "runs").as_array().ok_or_else(|| ExportError::invalid("页眉页脚字段缓存运行无效。"))?;
    if runs.iter().any(|run| !get(run, "text").is_string() || !get(run, "id").is_string()
        || run.get("metadata").is_some_and(|metadata| !metadata.is_object())) { return Err(ExportError::invalid("页眉页脚字段缓存运行格式无效。")); }
    let text = runs.iter().map(|run| get(run, "text").as_str().unwrap_or_default()).collect::<String>();
    if text != get(field, "cachedResult").as_str().unwrap_or_default() { return Err(ExportError::invalid("页眉页脚字段缓存文本与格式范围不一致；需同时更新缓存运行。")); }
    let mut offset = 0;
    for run in runs {
        let end = offset + get(run, "text").as_str().unwrap_or_default().encode_utf16().count();
        if slice_utf16(&text, offset, end).is_none() { return Err(ExportError::invalid("页眉页脚字段缓存格式拆分了 UTF-16 代理对。")); }
        offset = end;
    }
    Ok(Some(runs))
}

fn field_cache_style(field: &Value, run: &Value, defaults: &Value) -> Result<Value, ExportError> {
    let cleared = get(get(field, "metadata"), "officeStoryTextFormatCleared").as_bool() == Some(true);
    let declared = get(field, "textStyle");
    if declared.is_null() && !cleared { return Ok(run.clone()); }
    let style = if declared.is_null() { defaults } else { declared };
    let baseline = get(get(field, "metadata"), "officeStoryStyleBaseline");
    let mut result = run.clone();
    for key in ["bold", "italic", "underline", "strike", "color", "fontFamily", "eastAsiaFontFamily", "fontSize", "highlight", "verticalAlign"] {
        // A clear marker removes source rPr; the materialized cache can then
        // receive local edits without becoming a new field-wide override.
        if !baseline.is_null() && get(style, key) == get(baseline, key) { continue; }
        if let Some(value) = style.get(key) { result[key] = value.clone(); }
        else if let Some(object) = result.as_object_mut() { object.remove(key); }
        if key == "highlight" {
            if let Some(metadata) = result.get_mut("metadata").and_then(Value::as_object_mut) { metadata.remove("highlight"); }
        }
    }
    if cleared {
        if get(&result, "metadata").is_null() { result["metadata"] = json!({}); }
        let metadata = result["metadata"].as_object_mut().ok_or_else(|| ExportError::invalid("页眉页脚字段缓存来源格式无效。"))?;
        for key in ["bold", "italic", "underline", "strike", "color", "fontFamily", "eastAsiaFontFamily", "fontSize", "highlight", "verticalAlign"] { metadata.remove(key); }
        metadata.insert("officeStoryTextFormatCleared".into(), json!(true));
    }
    Ok(result)
}

fn render_token(context: &mut WriteContext<'_>, path: &str, token: &Token<'_>, from: usize, to: usize, next_relation: &mut u32) -> Result<String, ExportError> {
    if let Some(field) = token.field {
        let name = get(field, "field").as_str().unwrap_or_default();
        let instruction = get(field, "instruction").as_str().filter(|value| !value.trim().is_empty()).unwrap_or(name);
        let runs = field_cache::render(context, path, token, next_relation)?;
        let lock = if field_cache::is_locked(field)? { " w:fldLock=\"true\"" } else { "" };
        return Ok(format!("<w:fldSimple w:instr=\" {} \"{lock}>{runs}</w:fldSimple>", xml_attribute(instruction)));
    }
    let text = slice_utf16(&token.text, from - token.start, to - token.start).ok_or_else(|| ExportError::invalid("页眉页脚内容锁拆分了 UTF-16 代理对。"))?;
    render_styled_text(context, path, token.style, token.style, text, next_relation)
}

fn render_paragraph(context: &mut WriteContext<'_>, path: &str, tokens: &[Token<'_>], locks: &[Lock], from: usize, to: usize, empty_style: &Value, next_relation: &mut u32) -> Result<String, ExportError> {
    let local = locks.iter().filter(|lock| lock.start <= to && lock.end >= from).map(|lock| Lock { start: lock.start.max(from), end: lock.end.min(to), ..lock.clone() }).collect::<Vec<_>>();
    let mut boundaries = BTreeSet::from([from, to]);
    for token in tokens { boundaries.extend([token.start, token.end]); }
    for lock in &local { boundaries.extend([lock.start, lock.end]); }
    let points = boundaries.into_iter().collect::<Vec<_>>();
    let mut result = String::new();
    for (index, point) in points.iter().copied().enumerate() {
        let mut closes = local.iter().filter(|lock| lock.start < lock.end && lock.end == point).collect::<Vec<_>>();
        closes.sort_by_key(|lock| (std::cmp::Reverse(lock.start), std::cmp::Reverse(lock.id.clone())));
        for _ in closes { result.push_str("</w:sdtContent></w:sdt>"); }
        for lock in local.iter().filter(|lock| lock.start == point && lock.end == point) {
            result.push_str(&open_lock(context, lock)?);
            result.push_str(&story_run_xml(context, empty_style, empty_style, "")?);
            result.push_str("</w:sdtContent></w:sdt>");
        }
        let mut opens = local.iter().filter(|lock| lock.start == point && lock.start < lock.end).collect::<Vec<_>>();
        opens.sort_by_key(|lock| (std::cmp::Reverse(lock.end), lock.id.clone()));
        for lock in opens { result.push_str(&open_lock(context, lock)?); }
        if let Some(&end) = points.get(index + 1) {
            for token in tokens.iter().filter(|token| token.start < end && token.end > point) {
                result.push_str(&render_token(context, path, token, point.max(token.start), end.min(token.end), next_relation)?);
            }
        }
    }
    if result.is_empty() { result = story_run_xml(context, empty_style, empty_style, "")?; }
    Ok(result)
}

pub(super) fn write_story(context: &mut WriteContext<'_>, path: &str, content_type: &str, story: &str, text: &str, fields: &[&Value], supplied: &Value) -> Result<(), ExportError> {
    let root = if story == "header" { "hdr" } else { "ftr" };
    let defaults = context.story_defaults.clone();
    let legacy = legacy_content(text, &defaults);
    let content = if supplied.is_null() { &legacy } else { supplied };
    let runs = get(content, "runs").as_array().ok_or_else(|| ExportError::invalid("页眉页脚字符格式范围缺失。"))?;
    let paragraphs = get(content, "paragraphs").as_array().filter(|paragraphs| !paragraphs.is_empty()).ok_or_else(|| ExportError::invalid("页眉页脚段落范围缺失。"))?;
    if runs.iter().map(|run| get(run, "text").as_str().unwrap_or_default()).collect::<String>() != text { return Err(ExportError::invalid("页眉页脚文本与格式范围不一致。")); }
    let total = text.encode_utf16().count();
    let mut spans = Vec::new(); let mut offset_units = 0;
    for run in runs {
        if !supplied.is_null() { default_styles::validate(&json!({ "version": 1, "text": run, "paragraph": defaults["paragraph"] }))?; }
        let end = offset_units + get(run, "text").as_str().unwrap_or_default().encode_utf16().count();
        if slice_utf16(text, offset_units, end).is_none() { return Err(ExportError::invalid("页眉页脚字符格式边界无效。")); }
        spans.push((offset_units, end, run)); offset_units = end;
    }
    for field in fields {
        let name = get(field, "field").as_str().unwrap_or_default();
        if !matches!(name, "PAGE" | "NUMPAGES" | "DATE" | "TIME") { return Err(ExportError::invalid("页眉页脚字段类型无效。")); }
        if let Some(instruction) = field.get("instruction") {
            if !instruction.as_str().is_some_and(|value| value.trim().is_empty() || value.split_whitespace().next().is_some_and(|key| key.eq_ignore_ascii_case(name))) { return Err(ExportError::invalid("页眉页脚字段指令与类型不一致。")); }
        }
        if !get(field, "textStyle").is_null() { default_styles::validate(&json!({ "version": 1, "text": field["textStyle"], "paragraph": defaults["paragraph"] }))?; }
        if field.get("cachedResult").is_some_and(|value| !value.is_string()) { return Err(ExportError::invalid("页眉页脚字段缓存无效。")); }
        if let Some(runs) = field_cached_runs(field)? {
            for run in runs { default_styles::validate(&json!({ "version": 1, "text": run, "paragraph": defaults["paragraph"] }))?; }
        }
        field_cache::validate(field)?;
        let position = if field.get("offset").is_some() { offset(field, "offset")? } else { total };
        if slice_utf16(text, position, position).is_none() { return Err(ExportError::invalid("页眉页脚字段位置无效。")); }
    }
    let mut next = 0;
    for (index, paragraph) in paragraphs.iter().enumerate() {
        let start = offset(paragraph, "startOffset")?;
        let end = offset(paragraph, "endOffset")?;
        let part = slice_utf16(text, start, end).ok_or_else(|| ExportError::invalid("页眉页脚段落位置无效。"))?;
        if start != next || part.contains('\n') || index + 1 == paragraphs.len() && end != total
            || index + 1 < paragraphs.len() && slice_utf16(text, end, end + 1) != Some("\n") { return Err(ExportError::invalid("页眉页脚段落范围与文本不一致。")); }
        default_styles::validate(&json!({ "version": 1, "text": defaults["text"], "paragraph": paragraph["style"] }))?;
        next = end + 1;
    }
    let mut locks = Vec::new();
    if let Some(declared) = content.get("locks") {
        let mut ids = BTreeSet::new();
        for lock in declared.as_array().ok_or_else(|| ExportError::invalid("页眉页脚内容锁无效。"))? {
            let id = get(lock, "id").as_str().filter(|id| !id.is_empty() && id.len() <= 1024).ok_or_else(|| ExportError::invalid("页眉页脚内容锁 ID 无效。"))?;
            let kind = get(lock, "lock").as_str().filter(|kind| matches!(*kind, "contentLocked" | "sdtLocked" | "sdtContentLocked")).ok_or_else(|| ExportError::invalid("页眉页脚内容锁类型无效。"))?;
            if !ids.insert(id) { return Err(ExportError::invalid("页眉页脚内容锁 ID 重复。")); }
            locks.push(Lock { id: id.to_owned(), start: offset(lock, "startOffset")?, end: offset(lock, "endOffset")?, kind: kind.to_owned() });
        }
    }
    // Prepare the complete input buffer and every paragraph before retaining
    // source XML or writing relationships. Lock coordinates include fields.
    let mut prepared = Vec::new(); let mut buffer_text = String::new();
    let mut buffer = 0; let mut field_index = 0;
    for (index, paragraph) in paragraphs.iter().enumerate() {
        if index > 0 { buffer += 1; buffer_text.push('\n'); }
        let from = offset(paragraph, "startOffset")?;
        let to = offset(paragraph, "endOffset")?;
        let buffer_start = buffer;
        let mut tokens = Vec::new(); let mut cursor = from;
        while let Some(field) = fields.get(field_index) {
            let position = field.get("offset").map(|_| offset(field, "offset")).transpose()?.unwrap_or(total);
            if position > to { break; }
            if position < cursor { return Err(ExportError::invalid("页眉页脚字段顺序无效。")); }
            literal_tokens(&mut tokens, &spans, cursor, position, &mut buffer)?;
            if field.get("offset").is_none() { tokens.push(Token { start: buffer, end: buffer + 1, text: " ".into(), style: &defaults["text"], field: None }); buffer += 1; }
            let style = field.get("textStyle").filter(|style| !style.is_null()).unwrap_or(&defaults["text"]);
            tokens.push(Token { start: buffer, end: buffer + 1, text: "\u{fffc}".into(), style, field: Some(field) }); buffer += 1;
            cursor = position; field_index += 1;
        }
        literal_tokens(&mut tokens, &spans, cursor, to, &mut buffer)?;
        let empty_style = spans.iter().rev().find(|(start, end, _)| *start <= from && from <= *end).map(|(_, _, run)| *run).unwrap_or(&defaults["text"]);
        for token in &tokens { buffer_text.push_str(&token.text); }
        prepared.push(PreparedParagraph { start: buffer_start, end: buffer, tokens, style: &paragraph["style"], empty_style, provenance: paragraph });
    }
    if field_index != fields.len() || buffer_text.encode_utf16().count() != buffer { return Err(ExportError::invalid("页眉页脚输入 buffer 与字段范围不一致。")); }
    if locks.iter().any(|lock| slice_utf16(&buffer_text, lock.start, lock.end).is_none()) { return Err(ExportError::invalid("页眉页脚内容锁超出范围或拆分了 UTF-16 代理对。")); }
    for (index, first) in locks.iter().enumerate() { for second in locks.iter().skip(index + 1) {
        if first.start < second.start && second.start < first.end && first.end < second.end || second.start < first.start && first.start < second.end && second.end < first.end { return Err(ExportError::invalid("页眉页脚内容锁交叉重叠。")); }
    } }
    if !supplied.is_null() && preserve_source(context, path, content_type, root, text, fields, content)? { return Ok(()); }
    let mut body = String::new(); let mut next_relation = 1;
    for paragraph in prepared {
        let runs = render_paragraph(context, path, &paragraph.tokens, &locks, paragraph.start, paragraph.end, paragraph.empty_style, &mut next_relation)?;
        let generated = default_styles::paragraph_properties(paragraph.style);
        let properties = context.source_styles.story_paragraph_properties(&generated, paragraph.style, paragraph.provenance).map_err(ExportError::invalid)?;
        body.push_str(&format!("<w:p>{properties}{runs}</w:p>"));
    }
    context.package.add_xml(path, content_type, format!("{}<w:{root} xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">{body}</w:{root}>", xml::DECL))
}
