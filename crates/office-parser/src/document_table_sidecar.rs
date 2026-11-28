//! Keep native table properties and authored non-native metadata on the real
//! duke-pages extraction path, including covered rows and nested tables.
use std::{collections::{BTreeMap, HashMap, HashSet}, io::Cursor};

use duke_pages::{BlockDetails, BlockType, DocModel, DocumentBlock, StructuredDocument, TableNode};
use office_io_common::{chart_xml, document_table_source::{TABLE_SOURCE_NAMESPACE, TABLE_SOURCE_PART, TableSourceEntry, native_identity, validate_entries}};
use quick_xml::{Reader, events::Event};
use serde_json::{Map, Value, json};
use zip::ZipArchive;

use crate::document_format_sidecar::{Element, parse_xml, read_xml};

fn scope(node: &Element, inherited: &BTreeMap<String, String>) -> BTreeMap<String, String> {
    let mut result = inherited.clone();
    result.extend(node.namespaces.iter().map(|(prefix, namespace)| (prefix.clone(), namespace.clone())));
    result
}

fn escape(value: &str) -> String { quick_xml::escape::escape(value).into_owned() }

fn opening(xml: &str) -> Result<(usize, String, bool), String> {
    let mut reader = Reader::from_str(xml);
    loop {
        match reader.read_event().map_err(|error| error.to_string())? {
            Event::Start(tag) => return Ok((reader.buffer_position() as usize, String::from_utf8_lossy(tag.name().as_ref().as_bytes()).into_owned(), false)),
            Event::Empty(tag) => return Ok((reader.buffer_position() as usize, String::from_utf8_lossy(tag.name().as_ref().as_bytes()).into_owned(), true)),
            Event::Eof => return Err("Native table source has no opening element".into()),
            Event::DocType(_) => return Err("Native table source must not declare a document type".into()),
            _ => {}
        }
    }
}

fn complete_namespace(xml: &str, inherited: &BTreeMap<String, String>) -> Result<String, String> {
    let root = parse_xml(xml)?;
    let end = opening(xml)?.0.checked_sub(1).ok_or("Native table root has no opening delimiter")?;
    let offset = if xml.as_bytes().get(end.wrapping_sub(1)) == Some(&b'/') { end - 1 } else { end };
    let mut declarations = String::new();
    for (prefix, namespace) in inherited {
        if root.namespaces.contains_key(prefix) { continue; }
        if prefix.is_empty() { declarations.push_str(&format!(" xmlns=\"{}\"", escape(namespace))); }
        else { declarations.push_str(&format!(" xmlns:{prefix}=\"{}\"", escape(namespace))); }
    }
    Ok(format!("{}{declarations}{}", &xml[..offset], &xml[offset..]))
}

fn property(source: Option<&Element>, tag: &str, xml: &str, inherited: &BTreeMap<String, String>) -> Value {
    let Some(source) = source else {
        return json!({"version":1,"propertiesXml":format!("<w:{tag}/>"),"namespaces":inherited});
    };
    json!({"version":1,"propertiesXml":&xml[source.start..source.end],"namespaces":scope(source, inherited)})
}

fn word(node: &chart_xml::Node, name: &str) -> bool {
    node.name == name && matches!(node.namespace.as_deref(), Some(
        office_io_common::document_table_source::WORD_NAMESPACE | "http://purl.oclc.org/ooxml/wordprocessingml/main"))
}

fn word_child<'a>(node: &'a chart_xml::Node, name: &str) -> Option<&'a chart_xml::Node> {
    node.children.iter().find(|node| word(node, name))
}

fn word_attribute<'a>(node: &'a chart_xml::Node, name: &str) -> Option<&'a str> {
    node.attribute_ns(office_io_common::document_table_source::WORD_NAMESPACE, name)
        .or_else(|| node.attribute_ns("http://purl.oclc.org/ooxml/wordprocessingml/main", name))
}

// Both bounded readers preserve the same element order. Pair offsets with the
// expanded-name tree so an extension named tr/tc/tcPr cannot become a Word row
// or override a managed property merely by sharing its local name.
fn source_child<'a>(node: &'a Element, native: &chart_xml::Node, name: &str) -> Option<&'a Element> {
    node.children.iter().zip(&native.children).find(|(_, node)| word(node, name)).map(|(source, _)| source)
}

fn native_properties(xml: &str) -> Result<Value, String> {
    let table = parse_xml(xml)?;
    let native = chart_xml::parse(xml)?;
    if !word(&native, "tbl") { return Err("Native table properties require a Word table root".into()); }
    let namespaces = scope(&table, &BTreeMap::new());
    let mut rows = Vec::new();
    for (row, native_row, row_scope) in crate::document_native_xml::content_children(&table, &native, "tr", &namespaces) {
        let before = word_child(native_row, "trPr").and_then(|properties| word_child(properties, "gridBefore")).and_then(|value| word_attribute(value, "val")).unwrap_or("0")
            .parse::<usize>().map_err(|_| "Native table gridBefore is invalid")?;
        let mut column = before;
        let mut cells = Vec::new();
        for (cell, native_cell, cell_scope) in crate::document_native_xml::content_children(row, native_row, "tc", &row_scope) {
            let properties = source_child(cell, native_cell, "tcPr");
            let span = word_child(native_cell, "tcPr").and_then(|properties| word_child(properties, "gridSpan")).and_then(|span| word_attribute(span, "val")).unwrap_or("1")
                .parse::<usize>().ok().filter(|span| *span > 0).ok_or("Native table gridSpan is invalid")?;
            let synthetic_tail = native_cell.children.last().is_some_and(|paragraph| word(paragraph, "p")
                    && paragraph.attribute_ns(TABLE_SOURCE_NAMESPACE, "tableEnd") == Some("1"));
            let merge = |name| word_child(native_cell, "tcPr").and_then(|properties| word_child(properties, name))
                .map(|merge| word_attribute(merge, "val").unwrap_or("continue"));
            cells.push(json!({"columnIndex":column,"columnSpan":span,"syntheticTail":synthetic_tail,
                "horizontalMerge":merge("hMerge"),"verticalMerge":merge("vMerge"),
                "nativeProperties":property(properties, "tcPr", xml, &cell_scope)}));
            column = column.checked_add(span).ok_or("Native table column index overflow")?;
        }
        rows.push(json!({"nativeProperties":property(source_child(row, native_row, "trPr"), "trPr", xml, &row_scope),"cells":cells}));
    }
    let background = word_child(&native, "tblPr").and_then(|properties| word_child(properties, "shd")).and_then(|shading| word_attribute(shading, "fill"));
    Ok(json!({"version":1,"nativeProperties":property(source_child(&table, &native, "tblPr"), "tblPr", xml, &namespaces),"rows":rows,"backgroundColor":background}))
}

fn read_entries(archive: &mut ZipArchive<Cursor<&[u8]>>) -> Result<HashMap<String, TableSourceEntry>, String> {
    let Some(xml) = read_xml(archive, TABLE_SOURCE_PART)? else { return Ok(HashMap::new()); };
    let root = chart_xml::parse(&xml)?;
    if root.name != "tables" || root.namespace.as_deref() != Some(TABLE_SOURCE_NAMESPACE) || root.attrs.get("version").map(String::as_str) != Some("1") {
        return Err("DOCX table source mapping has an invalid root, namespace or version".into());
    }
    let mut entries = Vec::new();
    for table in root.children {
        if table.name != "table" || table.namespace.as_deref() != Some(TABLE_SOURCE_NAMESPACE) || table.children.len() != 1 {
            return Err("DOCX table source mapping has an unknown entry".into());
        }
        let data = &table.children[0];
        if data.name != "data" || data.namespace.as_deref() != Some(TABLE_SOURCE_NAMESPACE) || !data.children.is_empty() {
            return Err("DOCX table source mapping requires one JSON data node".into());
        }
        let required = |name: &str| table.attrs.get(name).cloned().ok_or_else(|| format!("DOCX table source entry requires {name}"));
        entries.push(TableSourceEntry { key: required("key")?, table_id: required("tableId")?, native_signature: required("nativeSignature")?,
            topology_signature: required("topologySignature")?, data: serde_json::from_str(&data.text).map_err(|error| format!("DOCX table source data is invalid: {error}"))? });
    }
    validate_entries(&entries)?;
    Ok(entries.into_iter().map(|entry| (entry.key.clone(), entry)).collect())
}

fn part_namespaces(archive: &mut ZipArchive<Cursor<&[u8]>>, part: &str) -> Result<BTreeMap<String, String>, String> {
    let Some(xml) = read_xml(archive, part)? else { return Ok(BTreeMap::new()); };
    let root = parse_xml(&xml)?;
    Ok(root.namespaces.into_iter().collect())
}

fn model_namespaces(source: &DocModel) -> Result<BTreeMap<String, String>, String> {
    let Some(xml) = source.metadata.document_open_tag.as_deref() else { return Ok(BTreeMap::new()); };
    let (end, name, empty) = opening(xml)?;
    let completed = if empty { xml[..end].to_owned() } else { format!("{}</{name}>", &xml[..end]) };
    let root = parse_xml(&completed)?;
    if root.name != "document" { return Err("Native document namespace context has the wrong root".into()); }
    Ok(root.namespaces.into_iter().collect())
}

fn details(block: &mut DocumentBlock) -> &mut Map<String, Value> {
    if !matches!(&block.details, BlockDetails::Raw(Value::Object(_))) {
        block.details = BlockDetails::Raw(block.details.to_json_value());
    }
    if let BlockDetails::Raw(Value::Object(details)) = &mut block.details { details }
    else { unreachable!("table extraction details are objects") }
}

/// Identity proof is independent of text; native geometry overrides old authored
/// geometry after an external resize, and a changed grid keeps stale data only
/// as explicit provenance rather than assigning it to unrelated columns.
pub(crate) fn restore(document: &mut StructuredDocument, source: &DocModel,
    mut archive: Option<&mut ZipArchive<Cursor<&[u8]>>>) -> Result<Vec<String>, String> {
    use crate::document_story_source::{SourceAddress, collect, stories};
    let tables = stories(source).iter().flat_map(|story| collect(story).tables).collect::<HashMap<_, _>>();
    let mut entries = HashMap::new();
    let mut part_scopes = HashMap::new();
    if let Some(archive) = archive.as_mut() {
        entries = read_entries(archive)?;
        for address in tables.keys() {
            if !part_scopes.contains_key(&address.part) {
                part_scopes.insert(address.part.clone(), part_namespaces(archive, &address.part)?);
            }
        }
    }
    // Duke normalizes the package vocabulary and hoists the actual namespace
    // bindings into document_open_tag. Its table source uses that vocabulary;
    // raw ZIP prefixes alone cannot describe renamed extension prefixes.
    let mut namespaces = model_namespaces(source)?;
    // Generated legacy fragments may have no document opening tag. Only the
    // standard prefixes emitted by the current Duke writer get a fallback;
    // unknown prefixes still fail the namespace-aware native XML parser.
    for (prefix, namespace) in [
        ("w", office_io_common::document_table_source::WORD_NAMESPACE),
        ("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"),
        ("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"),
        ("a", "http://schemas.openxmlformats.org/drawingml/2006/main"),
        ("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture"),
        ("w14", "http://schemas.microsoft.com/office/word/2010/wordml"),
        ("w15", "http://schemas.microsoft.com/office/word/2012/wordml"),
        ("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006"),
        ("v", "urn:schemas-microsoft-com:vml"),
        ("o", "urn:schemas-microsoft-com:office:office"),
        ("w10", "urn:schemas-microsoft-com:office:word"),
        ("m", "http://schemas.openxmlformats.org/officeDocument/2006/math"),
    ] { namespaces.entry(prefix.into()).or_insert_with(|| namespace.into()); }
    for scope in part_scopes.values_mut() {
        // The extracted fragments use Duke's canonical vocabulary; retain the
        // owning part's extra bindings without overriding that vocabulary.
        scope.extend(namespaces.iter().map(|(prefix, namespace)| (prefix.clone(), namespace.clone())));
    }
    let mut warnings = Vec::new();
    let mut used = HashSet::new();
    let mut used_sources = HashSet::new();
    fn visit(blocks: &mut [DocumentBlock], note_id: Option<i64>, tables: &HashMap<SourceAddress, &TableNode>, entries: &HashMap<String, TableSourceEntry>,
        namespaces: &BTreeMap<String, String>, scopes: &HashMap<String, BTreeMap<String, String>>, used: &mut HashSet<String>,
        used_sources: &mut HashSet<SourceAddress>, warnings: &mut Vec<String>) -> Result<(), String> {
        for block in blocks {
            let note_id = block.note_id.or(note_id);
            if block.r#type == BlockType::Table {
                let (address, source) = block.source_refs.iter().find_map(|reference| {
                    let address = SourceAddress::from_ref(reference, note_id);
                    tables.get(&address).map(|source| (address, *source))
                }).ok_or("A native table has no corresponding source in its package/story owner")?;
                used_sources.insert(address.clone());
                if source.block_id.as_deref().is_some_and(|id| id != block.id) { return Err("Native table extraction/source identities do not match".into()); }
                // Source row/cell styles are already resolved by duke-pages;
                // retain their actual physical row count even when no anchors
                // start in a row covered by preceding vertical merges.
                details(block).insert("rowCount".into(), json!(source.rows.len()));
                if let Some(xml) = &source.source_xml {
                    let xml = complete_namespace(xml, scopes.get(&address.part).unwrap_or(namespaces))?;
                    let identity = native_identity(&xml)?;
                    let mut properties = native_properties(&xml)?;
                    crate::document_table_merges::restore(block, &mut properties)?;
                    details(block).insert("nativeTableProperties".into(), properties);
                    if let Some(entry) = identity.key.as_ref().and_then(|key| entries.get(key)) {
                        if !used.insert(entry.key.clone()) {
                            warnings.push("A copied table has a duplicate source key; it keeps native identities rather than reusing the original table ID".into());
                        } else {
                            let topology_matches = identity.topology_signature == entry.topology_signature;
                            let native_matches = topology_matches && identity.native_signature == entry.native_signature;
                            details(block).insert("officeTable".into(), json!({"version":1,"tableId":entry.table_id,"nativeMatches":native_matches,
                                "topologyMatches":topology_matches,"data":entry.data}));
                            if !topology_matches { warnings.push(format!("Table {} changed its native grid; authored column/cell metadata was retained as stale source", entry.table_id)); }
                            else if !native_matches { warnings.push(format!("Table {} changed native geometry or styles; current native values replace its authored baseline", entry.table_id)); }
                        }
                    }
                }
            }
            visit(&mut block.children, note_id, tables, entries, namespaces, scopes, used, used_sources, warnings)?;
        }
        Ok(())
    }
    visit(&mut document.blocks, None, &tables, &entries, &namespaces, &part_scopes, &mut used, &mut used_sources, &mut warnings)?;
    if used_sources.len() != tables.len() { warnings.push("Some native story tables were not extracted; their source model remains retained".into()); }
    Ok(warnings)
}

#[cfg(test)]
mod tests {
    use super::*;
    use office_io_common::document_table_source::WORD_NAMESPACE;

    #[test]
    fn namespace_completion_uses_the_real_opening_event_and_retains_local_bindings() {
        let xml = "<?source table?><w:tbl xmlns:future=\"urn:local\" future:label=\"A > B\"><w:tblPr><future:setting/></w:tblPr><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>";
        let namespaces = BTreeMap::from([("w".into(), WORD_NAMESPACE.into()), ("future".into(), "urn:inherited".into())]);
        let completed = complete_namespace(xml, &namespaces).unwrap();
        let root = chart_xml::parse(&completed).unwrap();
        assert_eq!(root.attribute_ns("urn:local", "label"), Some("A > B"));
        assert_eq!(root.namespace.as_deref(), Some(WORD_NAMESPACE));
        assert!(root.attribute_ns("urn:inherited", "label").is_none());
        assert!(native_identity(&completed).is_ok());
    }

    #[test]
    fn only_word_properties_supply_grid_geometry_shading_and_synthetic_tails() {
        let xml = format!("<w:tbl xmlns:w=\"{WORD_NAMESPACE}\" xmlns:x=\"urn:future\" xmlns:ov=\"{TABLE_SOURCE_NAMESPACE}\"><x:tblPr><x:shd x:fill=\"bad\"/></x:tblPr><w:tblPr><w:shd x:fill=\"bad\" w:fill=\"112233\"/></w:tblPr><x:tr><x:tc/></x:tr><w:tr><x:trPr><x:gridBefore x:val=\"99\"/></x:trPr><w:trPr><w:gridBefore x:val=\"99\" w:val=\"1\"/></w:trPr><x:tc/><w:tc><x:tcPr/><w:tcPr><w:gridSpan x:val=\"99\" w:val=\"2\"/></w:tcPr><w:p x:tableEnd=\"1\"/></w:tc></w:tr></w:tbl>");
        let properties = native_properties(&xml).unwrap();
        assert_eq!(properties["backgroundColor"], "112233");
        let rows = properties["rows"].as_array().unwrap();
        assert_eq!(rows.len(), 1);
        let cells = rows[0]["cells"].as_array().unwrap();
        assert_eq!(cells.len(), 1);
        assert_eq!(cells[0]["columnIndex"], 1);
        assert_eq!(cells[0]["columnSpan"], 2);
        assert_eq!(cells[0]["syntheticTail"], false);
        assert!(cells[0]["nativeProperties"]["propertiesXml"].as_str().unwrap().contains("w:gridSpan"));
    }

    #[test]
    fn an_empty_property_fragment_gets_its_namespace_before_expanded_name_validation() {
        let completed = complete_namespace("<w:tbl/>", &BTreeMap::from([("w".into(), WORD_NAMESPACE.into())])).unwrap();
        assert_eq!(chart_xml::parse(&completed).unwrap().name, "tbl");
        assert!(complete_namespace("<!DOCTYPE tbl><w:tbl/>", &BTreeMap::new()).is_err());
    }
}
