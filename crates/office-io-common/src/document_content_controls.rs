//! Native SDT addresses and authored-property freshness shared by DOCX IO.
//! Text and row positions are always read from the current native structure;
//! a sidecar can retain identity and non-native properties, never old content.
use std::collections::BTreeSet;

use serde::{Deserialize, Serialize};
use serde_json::{Value, json};

use crate::chart_xml::Node;

pub const WORD_NAMESPACE: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
pub const STRICT_WORD_NAMESPACE: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
pub const WORD_2010_NAMESPACE: &str = "http://schemas.microsoft.com/office/word/2010/wordml";
pub const WORD_2012_NAMESPACE: &str = "http://schemas.microsoft.com/office/word/2012/wordml";
pub const CONTROL_SOURCE_NAMESPACE: &str = "urn:office-studio:content-controls:1";
pub const CONTROL_SOURCE_PART: &str = "customXml/office-content-controls.xml";
pub const CONTROL_SOURCE_TARGET: &str = "../customXml/office-content-controls.xml";
pub const MAX_CONTROL_COUNT: usize = 100_000;
pub const MAX_CONTROL_SOURCE_BYTES: usize = 16 * 1024 * 1024;

pub fn word(node: &Node, name: &str) -> bool {
    node.name == name && matches!(node.namespace.as_deref(), Some(WORD_NAMESPACE | STRICT_WORD_NAMESPACE))
}

pub fn word_attribute<'a>(node: &'a Node, name: &str) -> Option<&'a str> {
    node.attribute_ns(WORD_NAMESPACE, name).or_else(|| node.attribute_ns(STRICT_WORD_NAMESPACE, name))
}

/// These SDTs are already owned by the formula/code/story-lock adapters. Their
/// marker must survive extraction and must not create an additional user control.
pub fn internal_feature_tag(tag: &str) -> bool {
    tag.starts_with("office-latex:v1:") || tag.starts_with("office-code:v1:")
        || tag.starts_with("office_story_lock_v1_")
}

/// Only native content wrappers are transparent. A descendant table cannot
/// contribute its rows/cells to the outer table's topology.
pub fn word_children<'a>(root: &'a Node, name: &str) -> Vec<&'a Node> {
    fn visit<'a>(root: &'a Node, name: &str, output: &mut Vec<&'a Node>) {
        for node in &root.children {
            if word(node, name) { output.push(node); }
            else if word(node, "sdt") {
                if let Some(content) = node.children.iter().find(|node| word(node, "sdtContent")) { visit(content, name, output); }
            } else if ["sdtContent", "customXml", "ins", "moveTo"].iter().any(|tag| word(node, tag)) {
                visit(node, name, output);
            }
        }
    }
    let mut output = Vec::new();
    visit(root, name, &mut output);
    output
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeControlTextRange {
    pub block_id: String,
    pub start_offset: usize,
    pub end_offset: usize,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub start_note_reference_count: Option<usize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub end_note_reference_count: Option<usize>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeControlRowRange {
    pub table_block_id: String,
    pub row_indexes: Vec<usize>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeDocumentContentControl {
    pub id: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub parent_id: Option<String>,
    pub properties: Value,
    pub text_ranges: Vec<NativeControlTextRange>,
    #[serde(default, skip_serializing_if = "is_false")]
    pub has_object_content: bool,
    #[serde(default, skip_serializing_if = "is_zero")]
    pub object_content_count: usize,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub block_ids: Vec<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub table_row_range: Option<NativeControlRowRange>,
    pub part_name: String,
    pub story: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub note_id: Option<i64>,
    pub source_xml: String,
}

fn is_false(value: &bool) -> bool { !*value }
fn is_zero(value: &usize) -> bool { *value == 0 }

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct ControlSourceEntry {
    pub native_id: String,
    pub id: String,
    pub properties_signature: String,
    pub properties: Value,
}

fn canonical_namespace(namespace: Option<&str>) -> &str {
    match namespace { Some(STRICT_WORD_NAMESPACE) => WORD_NAMESPACE, Some(value) => value, None => "" }
}

fn signature_node(node: &Node) -> Value {
    let mut attributes = node.expanded_attrs.iter().flat_map(|(namespace, attributes)| attributes.iter().map(move |(name, value)|
        (canonical_namespace(Some(namespace)), name.as_str(), value.as_str()))).collect::<Vec<_>>();
    attributes.sort_unstable();
    json!([canonical_namespace(node.namespace.as_deref()), node.name, attributes, node.text.trim(),
        node.children.iter().map(signature_node).collect::<Vec<_>>()])
}

pub fn properties_signature(properties: Option<&Node>, end_properties: Option<&Node>) -> Result<String, String> {
    let data = serde_json::to_vec(&json!([properties.map(signature_node), end_properties.map(signature_node)]))
        .map_err(|error| error.to_string())?;
    let mut hash = 0xcbf29ce484222325u64;
    for byte in data { hash ^= u64::from(byte); hash = hash.wrapping_mul(0x100000001b3); }
    Ok(format!("{hash:016x}"))
}

pub fn validate_source_entries(entries: &[ControlSourceEntry]) -> Result<(), String> {
    if entries.len() > MAX_CONTROL_COUNT { return Err("DOCX_SDT_RESOURCE_LIMIT: too many authored content controls".into()); }
    let mut ids = BTreeSet::new();
    let mut native_ids = BTreeSet::new();
    let mut bytes = 0usize;
    for entry in entries {
        if entry.id.trim().is_empty() || entry.id.len() > 4096 || !ids.insert(entry.id.as_str())
            || entry.native_id.parse::<u32>().ok().filter(|id| *id > 0).is_none() || !native_ids.insert(entry.native_id.as_str())
            || entry.properties_signature.len() != 16 || !entry.properties_signature.bytes().all(|byte| byte.is_ascii_hexdigit())
            || !entry.properties.is_object() {
            return Err("DOCX content-control source has invalid or duplicate identity/properties".into());
        }
        bytes = bytes.checked_add(serde_json::to_vec(entry).map_err(|error| error.to_string())?.len())
            .filter(|bytes| *bytes <= MAX_CONTROL_SOURCE_BYTES).ok_or("DOCX_SDT_RESOURCE_LIMIT: authored control properties exceed 16 MiB")?;
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn row_wrappers_are_transparent_but_nested_tables_and_extension_names_are_not() {
        let xml = format!("<w:tbl xmlns:w=\"{WORD_NAMESPACE}\" xmlns:x=\"urn:other\"><w:sdt><w:sdtPr><x:tr/></w:sdtPr><w:sdtContent><w:sdt><w:sdtContent><w:tr><w:tc><w:tbl><w:tr/></w:tbl></w:tc></w:tr></w:sdtContent></w:sdt></w:sdtContent></w:sdt><x:tr/><w:tr/></w:tbl>");
        let root = crate::chart_xml::parse(&xml).unwrap();
        assert_eq!(word_children(&root, "tr").len(), 2);
        assert_eq!(word_children(word_children(&root, "tr")[0], "tc").len(), 1);
    }

    #[test]
    fn native_property_freshness_ignores_prefixes_and_retains_unknown_namespaces() {
        let first = crate::chart_xml::parse(&format!("<w:sdtPr xmlns:w=\"{WORD_NAMESPACE}\"><w:id w:val=\"1\"/><w:tag w:val=\"x\"/></w:sdtPr>")).unwrap();
        let second = crate::chart_xml::parse(&format!("<q:sdtPr xmlns:q=\"{STRICT_WORD_NAMESPACE}\"><q:id q:val=\"1\"/><q:tag q:val=\"x\"/></q:sdtPr>")).unwrap();
        assert_eq!(properties_signature(Some(&first), None).unwrap(), properties_signature(Some(&second), None).unwrap());
        let changed = crate::chart_xml::parse(&format!("<w:sdtPr xmlns:w=\"{WORD_NAMESPACE}\" xmlns:x=\"urn:other\"><w:id w:val=\"1\"/><x:tag x:val=\"x\"/></w:sdtPr>")).unwrap();
        assert_ne!(properties_signature(Some(&first), None).unwrap(), properties_signature(Some(&changed), None).unwrap());
    }
}
