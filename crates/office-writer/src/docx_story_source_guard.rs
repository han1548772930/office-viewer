//! A block-tree rewrite may replace only source structures mapped by the import
//! adapters. The legacy flat reader's feature list also contains native tables
//! and fields, so audit the actual XML instead of treating that list as a ban.
use std::collections::{BTreeSet, HashMap};

use serde_json::Value;

use super::super::{get, source_xml};
use crate::ExportError;

const REL_NAMESPACE: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_REL_NAMESPACE: &str = "http://schemas.openxmlformats.org/package/2006/relationships";

fn attribute<'a>(node: &'a source_xml::Element, namespace: &str, name: &str) -> Option<&'a str> {
    node.attributes.iter().find(|attribute| attribute.namespace == namespace && attribute.local == name).map(|attribute| attribute.value.as_str())
}

fn relationships(source: &Value) -> Result<HashMap<String, String>, ExportError> {
    let Some(xml) = source.get("relationshipsXml") else { return Ok(HashMap::new()); };
    let xml = xml.as_str().ok_or_else(|| ExportError::invalid("辅助内容区的来源关系 XML 无效。"))?;
    let root = source_xml::parse_xml(xml, true).map_err(ExportError::invalid)?;
    if root.namespace != PACKAGE_REL_NAMESPACE || root.local != "Relationships" { return Err(ExportError::invalid("辅助内容区的来源关系根节点无效。")); }
    let mut result = HashMap::new();
    for node in root.children() {
        if node.namespace != PACKAGE_REL_NAMESPACE || node.local != "Relationship" { return Err(ExportError::invalid("辅助内容区包含未知来源关系声明。")); }
        let id = attribute(node, "", "Id").filter(|id| !id.is_empty()).ok_or_else(|| ExportError::invalid("辅助内容区的来源关系 ID 缺失。"))?;
        let kind = attribute(node, "", "Type").ok_or_else(|| ExportError::invalid("辅助内容区的来源关系类型缺失。"))?;
        let target = attribute(node, "", "Target").ok_or_else(|| ExportError::invalid("辅助内容区的来源关系目标缺失。"))?;
        if kind.ends_with("/hyperlink") {
            if attribute(node, "", "TargetMode") != Some("External") { return Err(ExportError::invalid("辅助内容区的内部超链接关系尚未映射；来源已保留。")); }
            if result.insert(id.to_owned(), target.to_owned()).is_some() { return Err(ExportError::invalid("辅助内容区的来源超链接关系 ID 重复。")); }
        }
    }
    Ok(result)
}

fn mapped_lock(node: &source_xml::Element, content: &Value) -> bool {
    let Some(properties) = node.child("sdtPr") else { return false; };
    if properties.children().any(|property| property.namespace != source_xml::WORD_NAMESPACE
        || !matches!(property.local.as_str(), "id" | "lock" | "tag" | "text")) { return false; }
    // End properties have their own formatting semantics, beyond a bridge lock.
    if node.child("sdtEndPr").is_some() { return false; }
    let Some(kind @ ("contentLocked" | "sdtLocked" | "sdtContentLocked")) = properties.child("lock").and_then(|lock| lock.attribute("val")) else { return false; };
    let raw_id = properties.child("id").and_then(|id| id.attribute("val"));
    let encoded = properties.child("tag").and_then(|tag| tag.attribute("val")).and_then(|tag| tag.strip_prefix("office_story_lock_v1_"));
    let decoded = encoded.and_then(|encoded| {
        if encoded.len() > 2048 || encoded.len() % 2 != 0 || !encoded.is_ascii() { return None; }
        let bytes = (0..encoded.len()).step_by(2).map(|index| u8::from_str_radix(&encoded[index..index + 2], 16)).collect::<Result<Vec<_>, _>>().ok()?;
        String::from_utf8(bytes).ok()
    });
    let id = decoded.as_deref().or(raw_id);
    get(content, "locks").as_array().into_iter().flatten().any(|lock| get(lock, "id").as_str() == id && get(lock, "lock").as_str() == Some(kind))
}

#[derive(Default)]
struct Audit { gaps: BTreeSet<String>, has_native_tree_features: bool }

fn audit(node: &source_xml::Element, content: &Value, links: &HashMap<String, String>, result: &mut Audit, properties: bool) {
    if node.namespace != source_xml::WORD_NAMESPACE {
        if !properties { result.gaps.insert(node.name.clone()); }
    } else {
        match node.local.as_str() {
            "ins" | "del" | "moveFrom" | "moveTo" | "moveFromRangeStart" | "moveFromRangeEnd" | "moveToRangeStart" | "moveToRangeEnd"
            | "pPrChange" | "rPrChange" | "tblPrChange" | "tblGridChange" | "trPrChange" | "tcPrChange" | "cellIns" | "cellDel" | "cellMerge" => { result.gaps.insert("trackedChanges".into()); }
            "hdr" | "ftr" | "p" | "tr" | "tc" | "sdtContent" | "r" | "t" | "tab" | "br" | "cr" | "noBreakHyphen" | "softHyphen" | "sym" | "lastRenderedPageBreak" | "proofErr" => {},
            "tbl" | "fldSimple" | "fldChar" | "instrText" => { result.has_native_tree_features = true; }
            "pPr" | "rPr" | "tblPr" | "tblGrid" | "trPr" | "tcPr" => {
                // pPr/rPr and table provenance retain unmodeled properties through
                // the shared resource adapter. Revisions still require a model.
                for child in node.children() { audit(child, content, links, result, true); }
                return;
            }
            "sdt" => {
                if !mapped_lock(node, content) { result.gaps.insert("contentControls".into()); }
                for child in node.children().filter(|child| child.is_word("sdtContent")) { audit(child, content, links, result, false); }
                return;
            }
            "hyperlink" => {
                let target = node.attribute("anchor").map(|anchor| format!("#{anchor}"))
                    .or_else(|| attribute(node, REL_NAMESPACE, "id").and_then(|id| links.get(id).cloned()));
                if !target.is_some_and(|target| target.starts_with('#') || target.starts_with("https://") || target.starts_with("http://") || target.starts_with("mailto:")) {
                    result.gaps.insert("unresolvedHyperlink".into());
                }
            }
            "bookmarkStart" => {
                let generated = node.attribute("name").is_some_and(|name| name.starts_with("office_field_v1_") || name.starts_with("office_list_v1_")
                    || name.starts_with("office_code_v1_") || name.starts_with("office_code_plain_v1_"));
                if !generated { result.gaps.insert("nativeBookmarks".into()); }
            }
            "bookmarkEnd" => {},
            "commentRangeStart" | "commentRangeEnd" | "commentReference" => { result.gaps.insert("nativeCommentAnchors".into()); }
            _ if properties => {},
            _ => { result.gaps.insert(node.local.clone()); },
        }
    }
    for child in node.children() { audit(child, content, links, result, properties); }
}

pub(super) fn validate(content: &Value, root: &str) -> Result<(), ExportError> {
    let source = get(get(content, "metadata"), "officeSource");
    if source.is_null() { return Ok(()); }
    if get(source, "version").as_u64() != Some(1) { return Err(ExportError::invalid("辅助内容区的来源版本无效。")); }
    let xml = get(source, "xml").as_str().ok_or_else(|| ExportError::invalid("辅助内容区的来源 XML 缺失。"))?;
    let mut node = source_xml::parse_xml(xml, true).map_err(ExportError::invalid)?;
    source_xml::canonicalize_office_namespaces(&mut node).map_err(ExportError::invalid)?;
    if !node.is_word(root) { return Err(ExportError::invalid("辅助内容区的来源 XML 根节点无效。")); }
    let links = relationships(source)?;
    let mut result = Audit::default();
    audit(&node, content, &links, &mut result, false);
    if let Some(features) = source.get("unsupportedFeatures") {
        for feature in features.as_array().ok_or_else(|| ExportError::invalid("辅助内容区的来源特征清单无效。"))? {
            let feature = feature.as_str().ok_or_else(|| ExportError::invalid("辅助内容区的来源特征名称无效。"))?;
            if !matches!(feature, "tables" | "unknownFields" | "nestedFieldCodes" | "contentControls") { result.gaps.insert(feature.to_owned()); }
        }
    }
    if result.has_native_tree_features && get(get(get(content, "metadata"), "officeStoryBlockSource"), "version").as_u64() != Some(1) {
        result.gaps.insert("nativeBlockMapping".into());
    }
    if !result.gaps.is_empty() {
        return Err(ExportError::invalid(format!("此辅助内容区仍有未接入块树 DOCX 写出的来源结构：{}。来源已保留，需补齐映射后导出。", result.gaps.into_iter().collect::<Vec<_>>().join(", "))));
    }
    Ok(())
}

#[cfg(test)]
#[path = "docx_story_source_guard_tests.rs"]
mod tests;
