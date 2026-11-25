//! Native and public special-note identities share one allocation plan.
use std::collections::{BTreeMap, BTreeSet};

use office_io_common::document_note_source::{self, NOTE_SAFE_INTEGER};
use serde_json::Value;

use super::{DECORATIONS, NoteKind, native_kind};
use super::super::{get, source_xml};
use crate::ExportError;

#[derive(Clone, Copy, Default)]
struct Source<'a> {
    native_id: Option<i64>,
    native_part: Option<&'a str>,
    public_id: Option<&'a str>,
}

fn source<'a>(content: &'a Value, kind: NoteKind, note_type: &str) -> Result<Source<'a>, ExportError> {
    if !content.is_object() { return Err(ExportError::invalid("注释分隔内容必须为对象。")); }
    let Some(metadata) = content.get("metadata") else { return Ok(Source::default()); };
    if !metadata.is_object() { return Err(ExportError::invalid("注释分隔内容元数据必须为对象。")); }
    let mut result = Source::default();
    if let Some(native) = metadata.get("decorationSource") {
        if !native.is_object() || get(native, "version").as_u64() != Some(1)
            || get(native, "kind").as_str() != Some(native_kind(kind).name()) || get(native, "type").as_str() != Some(note_type) {
            return Err(ExportError::invalid("原生注释分隔内容来源版本、类型或宿主无效。"));
        }
        let id = get(native, "noteId").as_i64().filter(|id| id.unsigned_abs() <= NOTE_SAFE_INTEGER)
            .ok_or_else(|| ExportError::invalid("原生注释分隔内容来源 ID 必须为安全整数。"))?;
        let part = get(native, "partName").as_str().filter(|part| !part.is_empty())
            .ok_or_else(|| ExportError::invalid("原生注释分隔内容缺少真实部件名。"))?;
        let xml = get(native, "sourceXml").as_str().ok_or_else(|| ExportError::invalid("原生注释分隔内容缺少来源 XML。"))?;
        let mut root = source_xml::parse_xml(xml, false).map_err(ExportError::invalid)?;
        source_xml::canonicalize_office_namespaces(&mut root).map_err(ExportError::invalid)?;
        if !root.is_word(native_kind(kind).name()) || root.attribute("type") != Some(note_type)
            || root.attribute("id").and_then(|id| id.parse::<i64>().ok()) != Some(id) {
            return Err(ExportError::invalid("原生注释分隔内容的 XML 与来源身份不一致。"));
        }
        result.native_id = Some(id);
        result.native_part = Some(part);
    }
    if let Some(public) = metadata.get("univerNoteDecorationSource") {
        result.public_id = Some(document_note_source::public_note_decoration_source_id(public, native_kind(kind), note_type).map_err(ExportError::invalid)?);
    }
    Ok(result)
}

fn dormant_type(content: &Value) -> Result<&'static str, ExportError> {
    let metadata = get(content, "metadata");
    let note_type = get(get(metadata, "univerNoteDecorationSource"), "type").as_str()
        .or_else(|| get(get(metadata, "decorationSource"), "type").as_str());
    DECORATIONS.into_iter().find(|candidate| Some(*candidate) == note_type)
        .ok_or_else(|| ExportError::invalid("休眠注释分隔内容缺少有效来源类型。"))
}

/** Every resource, field and control writer borrows this same deduplicated set. */
pub(in super::super) fn decoration_contents<'a>(settings: &'a Value, kind: NoteKind) -> Result<Vec<(&'static str, &'a Value)>, ExportError> {
    if !settings.is_null() && !settings.is_object() { return Err(ExportError::invalid("注释设置必须为对象。")); }
    let mut candidates = Vec::new();
    for key in DECORATIONS {
        if let Some(content) = settings.get(key) { candidates.push((key, content)); }
    }
    if let Some(dormant) = get(settings, "metadata").get("officeDormantNoteDecorations") {
        if !dormant.is_object() || get(dormant, "version").as_u64() != Some(1) {
            return Err(ExportError::invalid("休眠注释分隔内容版本无效。"));
        }
        let contents = get(dormant, "contents").as_array().filter(|contents| contents.len() <= 100_000)
            .ok_or_else(|| ExportError::invalid("休眠注释分隔内容必须为预算内的数组。"))?;
        for content in contents { candidates.push((dormant_type(content)?, content)); }
    }
    let mut result: Vec<(&'static str, &'a Value)> = Vec::new();
    let mut native = BTreeMap::<(&str, i64), usize>::new();
    let mut public = BTreeMap::<&str, usize>::new();
    for (key, content) in candidates {
        let provenance = source(content, kind, key)?;
        let active = settings.get(key).is_some_and(|active| std::ptr::eq(active, content));
        if !active && provenance.native_id.is_none() && provenance.public_id.is_none() {
            return Err(ExportError::invalid("休眠注释分隔内容需要原生或公开来源身份。"));
        }
        let native_key = provenance.native_part.zip(provenance.native_id);
        let previous_native = native_key.and_then(|key| native.get(&key).copied());
        let previous_public = provenance.public_id.and_then(|key| public.get(key).copied());
        if previous_native.is_some() && previous_public.is_some() && previous_native != previous_public {
            return Err(ExportError::invalid("一份注释分隔来源关联多个不同定义。"));
        }
        let index = if let Some(index) = previous_native.or(previous_public) {
            if result[index] != (key, content) {
                return Err(ExportError::invalid("相同注释分隔来源身份的内容冲突；不能静默丢弃或改写来源。"));
            }
            index
        } else {
            let index = result.len();
            result.push((key, content));
            index
        };
        if let Some(key) = native_key { native.insert(key, index); }
        if let Some(key) = provenance.public_id { public.insert(key, index); }
    }
    Ok(result)
}

pub(super) struct Decoration<'a> {
    pub id: i64,
    pub kind: &'static str,
    pub content: Option<&'a Value>,
    pub active: bool,
    pub public_id: Option<&'a str>,
}

fn allocate_negative(next: &mut i64, used: &mut BTreeSet<i64>) -> Result<i64, ExportError> {
    while used.contains(next) {
        *next = next.checked_sub(1).filter(|next| next.unsigned_abs() <= NOTE_SAFE_INTEGER)
            .ok_or_else(|| ExportError::invalid("注释分隔内容标识超出安全范围。"))?;
    }
    let id = *next;
    used.insert(id);
    Ok(id)
}

fn public_candidate(id: &str) -> i64 {
    // Stable across active/dormant order and unrelated additions. A saved note
    // also carries this public ID and its actual native ID on its next import.
    let hash = id.as_bytes().iter().fold(0xcbf29ce484222325u64, |hash, byte| (hash ^ u64::from(*byte)).wrapping_mul(0x100000001b3));
    -i64::try_from(2 + hash % (NOTE_SAFE_INTEGER - 1)).expect("the hash is bounded by the safe-integer constant")
}

pub(super) fn decorations(settings: &Value, kind: NoteKind) -> Result<Vec<Decoration<'_>>, ExportError> {
    let contents = decoration_contents(settings, kind)?;
    let prepared = contents.iter().map(|(key, content)| source(content, kind, key)).collect::<Result<Vec<_>, _>>()?;
    let mut used = BTreeSet::new();
    // Reserve every native identity before allocating any new public story.
    for provenance in &prepared {
        if let Some(id) = provenance.native_id {
            if !used.insert(id) { return Err(ExportError::invalid("不同来源部件的注释分隔定义使用相同数字 ID；请先解决来源冲突。")); }
        }
    }
    let mut result = Vec::new();
    let mut next = -2;
    let mut defaults = Vec::new();
    for (key, preferred) in [("separator", -1), ("continuationSeparator", 0)] {
        if settings.get(key).is_some() { continue; }
        let id = if used.insert(preferred) { preferred } else { allocate_negative(&mut next, &mut used)? };
        defaults.push(Decoration { id, kind: key, content: None, active: true, public_id: None });
    }
    let mut public_ids = BTreeMap::new();
    for id in prepared.iter().filter(|source| source.native_id.is_none()).filter_map(|source| source.public_id).collect::<BTreeSet<_>>() {
        let mut candidate = public_candidate(id);
        while used.contains(&candidate) {
            candidate = if candidate.unsigned_abs() == NOTE_SAFE_INTEGER { -2 } else { candidate - 1 };
        }
        used.insert(candidate);
        public_ids.insert(id, candidate);
    }
    for ((key, content), provenance) in contents.into_iter().zip(prepared) {
        let id = if let Some(id) = provenance.native_id { id }
            else if let Some(id) = provenance.public_id { public_ids[id] }
            else {
                let preferred = if key == "separator" { -1 } else if key == "continuationSeparator" { 0 } else { -2 };
                if used.insert(preferred) { preferred } else { allocate_negative(&mut next, &mut used)? }
            };
        let active = settings.get(key).is_some_and(|active| std::ptr::eq(active, content));
        result.push(Decoration { id, kind: key, content: Some(content), active, public_id: provenance.public_id });
    }
    result.extend(defaults);
    Ok(result)
}

pub(in super::super) fn reserved_ids(model: &Value, kind: NoteKind) -> Result<BTreeSet<i64>, ExportError> {
    decorations(get(get(model, "noteSettings"), native_kind(kind).name()), kind).map(|values| values.into_iter().map(|value| value.id).collect())
}
