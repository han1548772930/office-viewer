//! Borrow all actual editing hosts before preparing shared DOCX dependencies.
use std::collections::{BTreeSet, HashMap};

use serde_json::{Value, json};

use super::{ContentControl, NoteBundle, NoteKind, WriteContext, get, note_writer, slice_utf16, table_writer, text_fields};
use crate::ExportError;

pub(super) struct AuxiliaryContent<'a> {
    pub key: String,
    pub content: &'a Value,
    pub blocks: &'a [Value],
    pub page_field_objects: bool,
}

pub(super) fn content_blocks(content: &Value) -> Result<Option<&[Value]>, ExportError> {
    content.get("blocks").map(|blocks| blocks.as_array().map(Vec::as_slice)
        .ok_or_else(|| ExportError::invalid("辅助内容区块树必须为数组。"))).transpose()
}

pub(super) fn auxiliary_contents<'a>(model: &'a Value, body: &[Value], notes: &'a NoteBundle) -> Result<Vec<AuxiliaryContent<'a>>, ExportError> {
    let mut result = Vec::new();
    let mut section_ids = vec!["first"];
    section_ids.extend(body.iter().filter(|block| get(block, "type").as_str() == Some("sectionBreak")).filter_map(|block| get(block, "id").as_str()));
    for (index, section_id) in section_ids.into_iter().enumerate() {
        let section = get(get(model, "sectionHeaderFooters"), section_id);
        let legacy_linked = index > 0 && (section.is_null() || get(section, "linkToPrevious").as_bool() == Some(true));
        for story in ["header", "footer", "firstPageHeader", "firstPageFooter", "evenPageHeader", "evenPageFooter"] {
            if index > 0 && get(get(section, "storyLinkToPrevious"), story).as_bool().unwrap_or(legacy_linked) { continue; }
            let content = get(get(section, "storyContent"), story);
            if let Some(blocks) = content_blocks(content)? {
                result.push(AuxiliaryContent { key: format!("{section_id}:{story}"), content, blocks, page_field_objects: true });
            }
        }
    }
    for (kind, entries) in [(NoteKind::Footnote, &notes.footnotes), (NoteKind::Endnote, &notes.endnotes)] {
        for note in entries {
            let content = get(&note.value, "content");
            if let Some(blocks) = content_blocks(content)? {
                result.push(AuxiliaryContent { key: format!("{}:{}", kind.collection(), note.id), content, blocks, page_field_objects: false });
            }
        }
        let name = if kind == NoteKind::Footnote { "footnote" } else { "endnote" };
        for (index, (decoration, content)) in note_writer::decoration_contents(get(get(model, "noteSettings"), name), kind)?.into_iter().enumerate() {
            if let Some(blocks) = content_blocks(content)? {
                result.push(AuxiliaryContent { key: format!("{}:decoration:{decoration}:{index}", kind.collection()), content, blocks, page_field_objects: false });
            }
        }
    }
    Ok(result)
}

pub(super) fn validate_ownership(groups: &[&[Value]]) -> Result<(), ExportError> {
    fn visit(blocks: &[Value], ids: &mut BTreeSet<String>) -> Result<(), ExportError> {
        for block in blocks {
            if let Some(id) = get(block, "id").as_str() {
                if !ids.insert(id.to_owned()) { return Err(ExportError::invalid(format!("文档块 {id:?} 被多个位置或内容区共享。"))); }
            }
            if get(block, "type").as_str() == Some("table") {
                for group in table_writer::content_groups(get(get(block, "metadata"), "table"))? { visit(group.blocks, ids)?; }
            }
        }
        Ok(())
    }
    let mut ids = BTreeSet::new();
    for blocks in groups { visit(blocks, &mut ids)?; }
    Ok(())
}

#[derive(Clone, Copy)]
struct FieldRange { start: usize, end: usize }

fn page_field_ranges(paragraphs: &[text_fields::Paragraph<'_>]) -> Result<Vec<FieldRange>, ExportError> {
    let by_id = paragraphs.iter().map(|paragraph| (paragraph.id.as_str(), paragraph)).collect::<HashMap<_, _>>();
    let mut all = Vec::<(FieldRange, bool)>::new();
    for paragraph in paragraphs {
        for field in get(get(get(paragraph.block, "metadata"), "documentTextFields"), "fields").as_array().into_iter().flatten() {
            let last = get(field, "endBlockId").as_str().map(|id| by_id.get(id).copied()).unwrap_or(Some(paragraph))
                .ok_or_else(|| ExportError::invalid("内容锁中的字段末段不属于该内容区。"))?;
            let range = FieldRange { start: paragraph.absolute + offset(field, "startOffset")?, end: last.absolute + offset(field, "endOffset")? };
            let page = get(field, "field").as_str().is_some_and(|name| ["PAGE", "NUMPAGES", "DATE", "TIME"].iter().any(|candidate| name.eq_ignore_ascii_case(candidate)));
            all.push((range, page));
        }
    }
    all.sort_by_key(|(range, _)| (range.start, std::cmp::Reverse(range.end)));
    // Match the TypeScript projection: equal nonempty PAGE ranges share one
    // input object, while adjacent zero-length fields each retain an object.
    let mut result: Vec<FieldRange> = Vec::new();
    let mut outer: Option<FieldRange> = None;
    for (range, page) in all {
        let contained = outer.is_some_and(|outer| outer.start <= range.start && outer.end >= range.end
            && (outer.start < range.start || outer.end > range.end));
        let duplicate = range.start < range.end && result.last().is_some_and(|last| last.start == range.start && last.end == range.end);
        if page && !contained && !duplicate { result.push(range); }
        if outer.is_none_or(|outer| range.end > outer.end) { outer = Some(range); }
    }
    Ok(result)
}

fn offset(value: &Value, key: &str) -> Result<usize, ExportError> {
    get(value, key).as_u64().filter(|value| *value <= 9_007_199_254_740_991).and_then(|value| usize::try_from(value).ok())
        .ok_or_else(|| ExportError::invalid(format!("内容锁 {key} 必须为非负安全整数。")))
}

fn display_offset(offset: usize, ranges: &[FieldRange], ending: bool) -> Result<usize, ExportError> {
    let mut removed = 0i128;
    for range in ranges {
        let position = range.start as i128 - removed;
        if (offset as i128) < position { break; }
        if (offset as i128) == position { return Ok(if ending { range.end } else { range.start }); }
        removed += range.end as i128 - range.start as i128 - 1;
    }
    usize::try_from(offset as i128 + removed).map_err(|_| ExportError::invalid("内容锁的字段偏移超出容量。"))
}

/// Legacy bridge locks are converted to the same inline SDT ranges as block commands.
pub(super) fn append_locks(context: &mut WriteContext<'_>, hosts: &[AuxiliaryContent<'_>]) -> Result<(), ExportError> {
    for host in hosts {
        let Some(locks) = host.content.get("locks") else { continue; };
        let locks = locks.as_array().ok_or_else(|| ExportError::invalid("辅助内容区的锁必须为数组。"))?;
        if locks.is_empty() { continue; }
        let paragraphs = text_fields::story_paragraphs(host.blocks)?;
        let ranges = if host.page_field_objects { page_field_ranges(&paragraphs)? } else { Vec::new() };
        let mut ids = BTreeSet::new();
        let total = paragraphs.last().map_or(0, |paragraph| paragraph.absolute + paragraph.text.encode_utf16().count());
        let mut declared = Vec::new();
        for lock in locks {
            let id = get(lock, "id").as_str().filter(|id| !id.is_empty() && id.len() <= 1024).ok_or_else(|| ExportError::invalid("辅助内容区的锁标识无效。"))?;
            if !ids.insert(id) { return Err(ExportError::invalid("辅助内容区的锁标识重复。")); }
            let kind = get(lock, "lock").as_str().filter(|kind| matches!(*kind, "contentLocked" | "sdtLocked" | "sdtContentLocked"))
                .ok_or_else(|| ExportError::invalid("辅助内容区的锁类型无效。"))?;
            let input_start = offset(lock, "startOffset")?; let input_end = offset(lock, "endOffset")?;
            let start = display_offset(input_start, &ranges, false)?; let end = display_offset(input_end, &ranges, true)?;
            if input_end < input_start || end < start || end > total { return Err(ExportError::invalid("辅助内容区的锁范围无效。")); }
            declared.push((input_start, input_end));
            let mut mapped = false;
            for paragraph in &paragraphs {
                let length = paragraph.text.encode_utf16().count(); let to = paragraph.absolute + length;
                let selected = if start == end { start >= paragraph.absolute && start <= to } else { start <= to && end >= paragraph.absolute };
                if !selected { continue; }
                let from = start.saturating_sub(paragraph.absolute); let to = end.saturating_sub(paragraph.absolute).min(length);
                if slice_utf16(&paragraph.text, from, to).is_none() { return Err(ExportError::invalid("辅助内容区的锁拆分了 UTF-16 代理对。")); }
                let control_id = format!("story-lock:{}:{id}:{}", host.key, paragraph.id);
                let numeric_id = context.next_story_control;
                context.next_story_control = numeric_id.checked_add(1).ok_or_else(|| ExportError::invalid("辅助内容区内容控件数量超出容量。"))?;
                context.controls_by_block.entry(paragraph.id.clone()).or_default().push(ContentControl { id: control_id.clone(), kind: "richText".into(),
                    numeric_id, start: from, end: to, value: json!({ "id": control_id, "type": "richText", "blockId": paragraph.id,
                        "startOffset": from, "endOffset": to, "lock": kind, "tag": super::story_writer::lock_tag(id), "placement": "inline", "metadata": { "officeStoryLockSource": { "version": 1, "id": id, "hostId": host.key } } }) });
                mapped = true;
            }
            if !mapped { return Err(ExportError::invalid("辅助内容区的锁没有可写出的文字目标。")); }
        }
        for (index, (start, end)) in declared.iter().enumerate() {
            if declared.iter().skip(index + 1).any(|(other_start, other_end)| start < other_start && other_start < end && end < other_end
                || other_start < start && start < other_end && other_end < end) { return Err(ExportError::invalid("辅助内容区的锁交叉重叠。")); }
        }
    }
    for controls in context.controls_by_block.values_mut() { controls.sort_by_key(|control| (control.start, std::cmp::Reverse(control.end), control.id.clone())); }
    Ok(())
}
