//! Native table geometry stays authoritative after external Office edits.
//! Text changes do not invalidate column types, IDs or precise authored sizes.
use std::collections::BTreeSet;

use serde::{Deserialize, Serialize};
use serde_json::{Value, json};

use crate::chart_xml::{Node, parse};

pub const TABLE_SOURCE_NAMESPACE: &str = "urn:office-studio:document-tables:1";
pub const TABLE_SOURCE_PART: &str = "customXml/office-tables.xml";
pub const TABLE_SOURCE_TARGET: &str = "../customXml/office-tables.xml";
pub const WORD_NAMESPACE: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT_WORD_NAMESPACE: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
pub const MAX_TABLE_SOURCE_BYTES: usize = 16 * 1024 * 1024;
pub const MAX_TABLE_COUNT: usize = 10_000;

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct TableSourceEntry {
    pub key: String,
    pub table_id: String,
    pub native_signature: String,
    pub topology_signature: String,
    pub data: Value,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NativeTableIdentity {
    pub key: Option<String>,
    pub native_signature: String,
    pub topology_signature: String,
}

fn word(node: &Node, name: &str) -> bool {
    node.name == name && matches!(node.namespace.as_deref(), Some(WORD_NAMESPACE | STRICT_WORD_NAMESPACE))
}

fn child<'a>(node: &'a Node, name: &str) -> Option<&'a Node> { node.children.iter().find(|node| word(node, name)) }

fn canonical_namespace(namespace: Option<&str>) -> &str {
    match namespace { Some(STRICT_WORD_NAMESPACE) => WORD_NAMESPACE, Some(value) => value, None => "" }
}

fn property(node: &Node) -> Value {
    let mut attributes = node.expanded_attrs.iter().flat_map(|(namespace, attributes)| attributes.iter().map(move |(key, value)| (
        canonical_namespace(Some(namespace)), key.as_str(), value.as_str(),
    ))).collect::<Vec<_>>();
    attributes.sort_unstable();
    json!([canonical_namespace(node.namespace.as_deref()), node.name, attributes, node.text.trim(),
        node.children.iter().map(property).collect::<Vec<_>>()])
}

fn signature(value: &Value) -> Result<String, String> {
    let data = serde_json::to_vec(value).map_err(|error| error.to_string())?;
    let mut hash = 0xcbf29ce484222325u64;
    for byte in data { hash ^= u64::from(byte); hash = hash.wrapping_mul(0x100000001b3); }
    Ok(format!("{hash:016x}"))
}

fn integer_property(node: Option<&Node>, name: &str, fallback: u64) -> Result<u64, String> {
    let Some(value) = node.and_then(|node| child(node, name)).and_then(|node| word_attribute(node, "val")) else { return Ok(fallback); };
    value.parse::<u64>().map_err(|_| format!("Native table {name} is not a nonnegative integer"))
}

fn word_attribute<'a>(node: &'a Node, name: &str) -> Option<&'a str> {
    node.attribute_ns(WORD_NAMESPACE, name).or_else(|| node.attribute_ns(STRICT_WORD_NAMESPACE, name))
}

fn merge_property(properties: Option<&Node>, name: &str) -> Result<Option<String>, String> {
    let merge = properties.and_then(|node| child(node, name)).map(|node| word_attribute(node, "val").unwrap_or("continue").to_owned());
    if merge.as_deref().is_some_and(|value| !matches!(value, "restart" | "continue")) { return Err(format!("Native table {name} is invalid")); }
    Ok(merge)
}

/// Accept a namespace-complete table fragment. Only this table's own properties
/// and topology participate; paragraphs and nested table contents are excluded.
pub fn native_identity(xml: &str) -> Result<NativeTableIdentity, String> {
    let root = parse(xml)?;
    if !word(&root, "tbl") { return Err("Native table identity requires a Word table root".into()); }
    let key = root.attribute_ns(TABLE_SOURCE_NAMESPACE, "key").map(str::to_owned);
    if key.as_deref().is_some_and(|key| !valid_key(key)) { return Err("Native table source key is invalid".into()); }
    let mut layout_rows = Vec::new();
    let mut topology_rows = Vec::new();
    for row in crate::document_content_controls::word_children(&root, "tr") {
        let row_properties = child(row, "trPr");
        let mut cell_layout = Vec::new();
        let mut cell_topology = Vec::new();
        for cell in crate::document_content_controls::word_children(row, "tc") {
            let properties = child(cell, "tcPr");
            let span = integer_property(properties, "gridSpan", 1)?;
            if span == 0 { return Err("Native table grid span must be positive".into()); }
            let merge = merge_property(properties, "vMerge")?;
            let horizontal_merge = merge_property(properties, "hMerge")?;
            cell_layout.push(properties.map(property));
            cell_topology.push(json!([span, merge, horizontal_merge]));
        }
        layout_rows.push(json!([row_properties.map(property), cell_layout]));
        topology_rows.push(json!([integer_property(row_properties, "gridBefore", 0)?, integer_property(row_properties, "gridAfter", 0)?, cell_topology]));
    }
    let grid = child(&root, "tblGrid");
    let column_count = grid.map(|grid| grid.children.iter().filter(|node| word(node, "gridCol")).count()).unwrap_or(0);
    Ok(NativeTableIdentity {
        key,
        native_signature: signature(&json!([child(&root, "tblPr").map(property), grid.map(property), layout_rows]))?,
        topology_signature: signature(&json!([column_count, topology_rows]))?,
    })
}

pub fn valid_key(key: &str) -> bool {
    key.strip_prefix("table-").is_some_and(|suffix| suffix.len() == 8 && suffix.bytes().all(|byte| byte.is_ascii_hexdigit()))
}

pub fn validate_entries(entries: &[TableSourceEntry]) -> Result<(), String> {
    if entries.len() > MAX_TABLE_COUNT { return Err("DOCX_TABLE_SOURCE_RESOURCE_LIMIT: too many table source entries".into()); }
    let mut keys = BTreeSet::new();
    let mut ids = BTreeSet::new();
    let mut bytes = 0usize;
    for entry in entries {
        if !valid_key(&entry.key) || !keys.insert(entry.key.as_str()) || entry.table_id.trim().is_empty()
            || entry.table_id.len() > 4096 || !ids.insert(entry.table_id.as_str()) || !entry.data.is_object()
            || [&entry.native_signature, &entry.topology_signature].iter().any(|value| value.len() != 16 || !value.bytes().all(|byte| byte.is_ascii_hexdigit())) {
            return Err("DOCX table source entry has invalid or duplicate identity, signature or data".into());
        }
        bytes = bytes.checked_add(serde_json::to_vec(entry).map_err(|error| error.to_string())?.len())
            .filter(|bytes| *bytes <= MAX_TABLE_SOURCE_BYTES).ok_or("DOCX_TABLE_SOURCE_RESOURCE_LIMIT: table source data exceeds 16 MiB")?;
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn table(prefix: &str, namespace: &str, text: &str, width: u32) -> String {
        format!("<{prefix}:tbl xmlns:{prefix}=\"{namespace}\" xmlns:ov=\"{TABLE_SOURCE_NAMESPACE}\" ov:key=\"table-00000001\"><{prefix}:tblPr><{prefix}:tblW {prefix}:w=\"{width}\" {prefix}:type=\"dxa\"/></{prefix}:tblPr><{prefix}:tblGrid><{prefix}:gridCol {prefix}:w=\"{width}\"/></{prefix}:tblGrid><{prefix}:tr><{prefix}:tc><{prefix}:tcPr/><{prefix}:p><{prefix}:r><{prefix}:t>{text}</{prefix}:t></{prefix}:r></{prefix}:p></{prefix}:tc></{prefix}:tr></{prefix}:tbl>")
    }

    #[test]
    fn native_text_edits_and_namespace_prefixes_do_not_invalidate_authored_metadata() {
        let first = native_identity(&table("w", WORD_NAMESPACE, "old", 1500)).unwrap();
        let second = native_identity(&table("z", STRICT_WORD_NAMESPACE, "new &amp; 中", 1500)).unwrap();
        assert_eq!(first, second);
        assert_eq!(first.key.as_deref(), Some("table-00000001"));
        let resized = native_identity(&table("w", WORD_NAMESPACE, "new", 1800)).unwrap();
        assert_ne!(first.native_signature, resized.native_signature);
        assert_eq!(first.topology_signature, resized.topology_signature);
    }

    #[test]
    fn an_attribute_from_another_namespace_cannot_claim_table_ownership() {
        let xml = table("w", WORD_NAMESPACE, "text", 1500).replace(TABLE_SOURCE_NAMESPACE, "urn:another-extension");
        assert!(native_identity(&xml).unwrap().key.is_none());
    }

    #[test]
    fn property_attribute_order_and_prefixes_are_equivalent_but_extension_changes_are_not() {
        let first = table("w", WORD_NAMESPACE, "text", 1500).replace("<w:tcPr/>", "<w:tcPr xmlns:x=\"urn:future\"><x:property x:val=\"2\" w:val=\"1\"/></w:tcPr>");
        let renamed = table("z", STRICT_WORD_NAMESPACE, "text", 1500).replace("<z:tcPr/>", "<z:tcPr xmlns:y=\"urn:future\"><y:property z:val=\"1\" y:val=\"2\"/></z:tcPr>");
        let changed = first.replace("x:val=\"2\"", "x:val=\"3\"");
        assert_eq!(native_identity(&first).unwrap(), native_identity(&renamed).unwrap());
        let changed = native_identity(&changed).unwrap();
        let original = native_identity(&first).unwrap();
        assert_ne!(original.native_signature, changed.native_signature);
        assert_eq!(original.topology_signature, changed.topology_signature);
    }

    #[test]
    fn legacy_horizontal_merge_changes_invalidate_cell_metadata_topology() {
        let first = table("w", WORD_NAMESPACE, "text", 1500);
        let merged = first.replace("<w:tcPr/>", "<w:tcPr><w:hMerge w:val=\"restart\"/></w:tcPr>");
        assert_ne!(native_identity(&first).unwrap().topology_signature, native_identity(&merged).unwrap().topology_signature);
        assert!(native_identity(&first.replace("<w:tcPr/>", "<w:tcPr><w:hMerge w:val=\"invalid\"/></w:tcPr>")).is_err());
    }

    #[test]
    fn only_expanded_word_rows_and_cells_participate_in_native_geometry() {
        let first = table("w", WORD_NAMESPACE, "text", 1500);
        let extended = first.replace("</w:tbl>", "<x:tr xmlns:x=\"urn:future\"><x:tc/></x:tr></w:tbl>");
        assert_eq!(native_identity(&first).unwrap(), native_identity(&extended).unwrap());
    }
}
