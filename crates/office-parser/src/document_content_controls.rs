//! Read native SDT properties and current text/row addresses before frontend
//! normalization. Duke's simplified form-field projection is not the SDT model.
use std::{collections::{BTreeMap, HashMap, HashSet}, io::Cursor};

use duke_pages::{DocModel, DocumentBlock, StructuredDocument};
use office_io_common::{chart_xml::{self, Node}, document_content_controls::{self as contract,
    ControlSourceEntry, NativeControlRowRange, NativeControlTextRange, NativeDocumentContentControl,
    WORD_2010_NAMESPACE, WORD_2012_NAMESPACE, word, word_attribute}};
use serde_json::{Map, Value, json};
use zip::ZipArchive;

use crate::{document_format_sidecar::{Element, parse_xml, read_xml},
    document_native_xml::{complete_fragment, content_children, scope},
    document_story_source::{SourceAddress, SourceStory, collect, stories}};

#[derive(Default)]
pub(crate) struct Restoration {
    pub controls: Vec<NativeDocumentContentControl>,
    pub warnings: Vec<String>,
}

fn child<'a>(node: &'a Node, name: &str) -> Option<&'a Node> { node.children.iter().find(|node| word(node, name)) }
fn extension_child<'a>(node: &'a Node, namespace: &str, name: &str) -> Option<&'a Node> {
    node.children.iter().find(|node| node.name == name && node.namespace.as_deref() == Some(namespace))
}
fn on_off(value: Option<&str>) -> bool { !matches!(value, Some("0" | "false" | "off")) }
fn insert_string(output: &mut Map<String, Value>, key: &str, value: Option<&str>) {
    if let Some(value) = value { output.insert(key.into(), json!(value)); }
}
fn property_value<'a>(properties: &'a Node, name: &str) -> Option<&'a str> { word_attribute(child(properties, name)?, "val") }

fn raw_child(xml: &str, source: &Element, native: &Node, name: &str, namespaces: &BTreeMap<String, String>) -> Result<Option<String>, String> {
    source.children.iter().zip(&native.children).find(|(_, node)| word(node, name))
        .map(|(node, _)| complete_fragment(xml, node, &scope(node, namespaces))).transpose()
}

fn properties(xml: &str, source: &Element, native: &Node, namespaces: &BTreeMap<String, String>) -> Result<Value, String> {
    let mut output = Map::new();
    let Some(properties) = child(native, "sdtPr") else { output.insert("type".into(), json!("richText")); return Ok(Value::Object(output)); };
    let kind = if child(properties, "text").is_some() { "text" }
        else if extension_child(properties, WORD_2010_NAMESPACE, "checkbox").is_some() { "checkbox" }
        else if child(properties, "date").is_some() { "date" }
        else if child(properties, "dropDownList").is_some() { "dropDownList" }
        else if child(properties, "comboBox").is_some() { "comboBox" }
        else if child(properties, "picture").is_some() { "picture" }
        else if child(properties, "group").is_some() { "group" }
        else if extension_child(properties, WORD_2012_NAMESPACE, "repeatingSection").is_some() { "repeatingSection" }
        else if extension_child(properties, WORD_2012_NAMESPACE, "repeatingSectionItem").is_some() { "repeatingSectionItem" }
        else if child(properties, "docPartObj").is_some() { "docPartObj" }
        else { "richText" };
    output.insert("type".into(), json!(kind));
    for (key, name) in [("alias", "alias"), ("title", "alias"), ("tag", "tag"), ("lock", "lock")] {
        insert_string(&mut output, key, property_value(properties, name));
    }
    for name in ["appearance", "color"] {
        insert_string(&mut output, name, extension_child(properties, WORD_2012_NAMESPACE, name)
            .and_then(|node| node.attribute_ns(WORD_2012_NAMESPACE, "val")).or_else(|| property_value(properties, name)));
    }
    for (key, name) in [("temporary", "temporary"), ("showingPlaceholder", "showingPlcHdr")] {
        if let Some(node) = child(properties, name) { output.insert(key.into(), json!(on_off(word_attribute(node, "val")))); }
    }
    if let Some(node) = child(properties, "text") {
        output.insert("multiLine".into(), json!(word_attribute(node, "multiLine").is_some_and(|value| on_off(Some(value)))));
    }
    if let Some(node) = child(properties, "placeholder").and_then(|node| child(node, "docPart")) {
        output.insert("placeholderDocPart".into(), json!({"name":word_attribute(node, "val")}));
    }
    if let Some(binding) = child(properties, "dataBinding") {
        let mut value = Map::new();
        for key in ["prefixMappings", "xpath", "storeItemID"] { insert_string(&mut value, key, word_attribute(binding, key)); }
        output.insert("dataBinding".into(), Value::Object(value));
    }
    if let Some(checkbox) = extension_child(properties, WORD_2010_NAMESPACE, "checkbox") {
        let checked = extension_child(checkbox, WORD_2010_NAMESPACE, "checked")
            .is_some_and(|node| on_off(node.attribute_ns(WORD_2010_NAMESPACE, "val")));
        let mut value = Map::new(); value.insert("checked".into(), json!(checked));
        for name in ["checkedState", "uncheckedState"] {
            if let Some(state) = extension_child(checkbox, WORD_2010_NAMESPACE, name) {
                let mut attributes = Map::new();
                insert_string(&mut attributes, "value", state.attribute_ns(WORD_2010_NAMESPACE, "val"));
                insert_string(&mut attributes, "font", state.attribute_ns(WORD_2010_NAMESPACE, "font"));
                value.insert(name.into(), Value::Object(attributes));
            }
        }
        output.insert("checked".into(), json!(checked)); output.insert("checkbox".into(), Value::Object(value));
    }
    if let Some(date) = child(properties, "date") {
        insert_string(&mut output, "dateValue", word_attribute(date, "fullDate"));
        for (key, name) in [("dateFormat", "dateFormat"), ("dateLocale", "lid"), ("dateCalendar", "calendar"), ("dateStorage", "storeMappedDataAs")] {
            insert_string(&mut output, key, property_value(date, name));
        }
    }
    if let Some(list) = child(properties, "dropDownList").or_else(|| child(properties, "comboBox")) {
        let items = list.children.iter().filter(|node| word(node, "listItem")).map(|node| {
            let mut value = Map::new(); insert_string(&mut value, "displayText", word_attribute(node, "displayText"));
            value.insert("value".into(), json!(word_attribute(node, "value").unwrap_or_default())); Value::Object(value)
        }).collect::<Vec<_>>();
        output.insert("items".into(), json!(items));
    }
    if let Some(section) = extension_child(properties, WORD_2012_NAMESPACE, "repeatingSection") {
        let mut value = Map::new();
        insert_string(&mut value, "sectionTitle", extension_child(section, WORD_2012_NAMESPACE, "sectionTitle")
            .and_then(|node| node.attribute_ns(WORD_2012_NAMESPACE, "val")).or_else(|| section.attribute_ns(WORD_2012_NAMESPACE, "sectionTitle")));
        if let Some(locked) = extension_child(section, WORD_2012_NAMESPACE, "doNotAllowInsertDelete") {
            value.insert("doNotAllowInsertDelete".into(), json!(on_off(locked.attribute_ns(WORD_2012_NAMESPACE, "val"))));
        } else if let Some(locked) = section.attribute_ns(WORD_2012_NAMESPACE, "doNotAllowInsertDelete") {
            value.insert("doNotAllowInsertDelete".into(), json!(on_off(Some(locked))));
        }
        output.insert("repeatingSection".into(), Value::Object(value));
    }
    if let Some(part) = child(properties, "docPartObj") {
        let mut value = Map::new();
        insert_string(&mut value, "gallery", property_value(part, "docPartGallery"));
        insert_string(&mut value, "category", property_value(part, "docPartCategory"));
        if let Some(unique) = child(part, "docPartUnique") { value.insert("unique".into(), json!(on_off(word_attribute(unique, "val")))); }
        output.insert("docPart".into(), Value::Object(value));
    }
    if let Some(raw) = raw_child(xml, source, native, "sdtPr", namespaces)? { output.insert("rawSdtPr".into(), json!(raw)); }
    if let Some(raw) = raw_child(xml, source, native, "sdtEndPr", namespaces)? { output.insert("rawSdtEndPr".into(), json!(raw)); }
    Ok(Value::Object(output))
}

fn read_sources(archive: &mut ZipArchive<Cursor<&[u8]>>) -> Result<HashMap<String, ControlSourceEntry>, String> {
    let Some(xml) = read_xml(archive, contract::CONTROL_SOURCE_PART)? else { return Ok(HashMap::new()); };
    let root = chart_xml::parse(&xml)?;
    if root.name != "controls" || root.namespace.as_deref() != Some(contract::CONTROL_SOURCE_NAMESPACE) { return Ok(HashMap::new()); }
    if root.attrs.get("version").map(String::as_str) != Some("1") { return Err("Native SDT source mapping has an invalid version".into()); }
    let mut entries = Vec::new();
    for node in &root.children {
        if node.name != "control" || node.namespace != root.namespace || !node.children.is_empty() { return Err("Native SDT source mapping has an unknown entry".into()); }
        entries.push(serde_json::from_str::<ControlSourceEntry>(&node.text).map_err(|error| error.to_string())?);
    }
    contract::validate_source_entries(&entries)?;
    Ok(entries.into_iter().map(|entry| (entry.native_id.clone(), entry)).collect())
}

fn collect_blocks(blocks: &[DocumentBlock], inherited_note: Option<i64>, result: &mut HashMap<SourceAddress, (String, String)>) {
    for block in blocks {
        let note = block.note_id.or(inherited_note);
        for reference in &block.source_refs { result.entry(SourceAddress::from_ref(reference, note)).or_insert_with(|| (block.id.clone(), block.content.clone())); }
        collect_blocks(&block.children, note, result);
    }
}

fn raw_hosts<'a>(source: &'a Element, node: &'a Node, paragraphs: &mut Vec<(&'a Element, &'a Node)>, tables: &mut Vec<(&'a Element, &'a Node)>) {
    if word(node, "p") {
        let go_back = node.children.iter().any(|node| word(node, "bookmarkStart") && word_attribute(node, "name") == Some("_GoBack"))
            && !node.children.iter().any(|node| word(node, "r") || word(node, "drawing"));
        if !go_back { paragraphs.push((source, node)); }
        return;
    }
    if word(node, "tbl") { tables.push((source, node)); }
    if ["sdtPr", "sdtEndPr", "pPr", "rPr", "trPr", "tcPr", "tblPr", "del", "moveFrom", "drawing", "pict", "txbxContent"].iter().any(|tag| word(node, tag)) { return; }
    for (source, node) in source.children.iter().zip(&node.children) { raw_hosts(source, node, paragraphs, tables); }
}

fn story_root<'a>(root: &'a Element, native: &'a Node, story: &SourceStory<'_>) -> Option<(&'a Element, &'a Node)> {
    if story.story == "body" { return root.children.iter().zip(&native.children).find(|(_, node)| word(node, "body")); }
    if let Some(id) = story.note_id {
        let name = if story.story == "footnote" { "footnote" } else { "endnote" };
        return root.children.iter().zip(&native.children).find(|(_, node)| word(node, name) && word_attribute(node, "id").and_then(|value| value.parse::<i64>().ok()) == Some(id));
    }
    (story.story == "header" || story.story == "footer").then_some((root, native))
}

struct ParagraphCursor { block: String, text: String, expected: String, offset: usize, note_reference_count: usize }
struct Walker<'a> {
    xml: &'a str, story: &'a SourceStory<'a>, sources: &'a HashMap<String, ControlSourceEntry>, used_sources: &'a mut HashSet<String>,
    paragraphs: HashMap<usize, (String, String)>, tables: HashMap<usize, String>, rows: HashMap<usize, (String, usize)>,
    result: &'a mut Restoration, active: Vec<usize>, current: Option<ParagraphCursor>, table: Option<String>, invalid_blocks: HashSet<String>,
    source_bytes: &'a mut usize, ids: &'a mut HashSet<String>,
}

impl Walker<'_> {
    fn address(&mut self) {
        if let Some(current) = &self.current {
            for index in &self.active {
                let ranges = &mut self.result.controls[*index].text_ranges;
                if !ranges.iter().any(|range| range.block_id == current.block) {
                    ranges.push(NativeControlTextRange { block_id: current.block.clone(), start_offset: current.offset, end_offset: current.offset,
                        start_note_reference_count: Some(current.note_reference_count), end_note_reference_count: Some(current.note_reference_count) });
                }
            }
        }
    }

    fn text(&mut self, text: &str) -> Result<(), String> {
        self.address();
        if let Some(current) = &mut self.current {
            current.offset = current.offset.checked_add(text.encode_utf16().count()).ok_or("Native SDT UTF-16 offset overflow")?;
            current.text.push_str(text);
            for index in &self.active {
                if let Some(range) = self.result.controls[*index].text_ranges.iter_mut().find(|range| range.block_id == current.block) {
                    range.end_offset = current.offset; range.end_note_reference_count = Some(current.note_reference_count);
                }
            }
        }
        Ok(())
    }

    fn visit(&mut self, source: &Element, node: &Node, inherited: &BTreeMap<String, String>) -> Result<(), String> {
        let namespaces = scope(source, inherited);
        if ["drawing", "pict", "object"].iter().any(|tag| word(node, tag)) {
            for index in &self.active {
                self.result.controls[*index].has_object_content = true;
                self.result.controls[*index].object_content_count += 1;
            }
            return Ok(());
        }
        if ["sdtPr", "sdtEndPr", "pPr", "rPr", "trPr", "tcPr", "tblPr", "del", "moveFrom", "drawing", "pict", "object", "txbxContent", "instrText"].iter().any(|tag| word(node, tag)) { return Ok(()); }
        if word(node, "p") {
            let previous = self.current.take();
            self.current = self.paragraphs.get(&source.start).map(|(block, expected)| ParagraphCursor { block: block.clone(), expected: expected.clone(), text: String::new(), offset: 0, note_reference_count: 0 });
            self.address();
            for (source, node) in source.children.iter().zip(&node.children) { self.visit(source, node, &namespaces)?; }
            if let Some(current) = self.current.take() {
                if current.expected != current.text {
                    self.invalid_blocks.insert(current.block.clone());
                    self.result.warnings.push(format!("Native SDT text in {} differs from Duke's visible projection; the exact SDT XML and properties remain retained, without invented offsets", current.block));
                }
            }
            self.current = previous;
            return Ok(());
        }
        if word(node, "tbl") {
            let previous_table = self.table.take(); let previous_paragraph = self.current.take();
            self.table = self.tables.get(&source.start).cloned();
            for (source, node) in source.children.iter().zip(&node.children) { self.visit(source, node, &namespaces)?; }
            self.table = previous_table; self.current = previous_paragraph; return Ok(());
        }
        if word(node, "sdt") {
            let content = source.children.iter().zip(&node.children).find(|(_, node)| word(node, "sdtContent"));
            if child(node, "sdtPr").and_then(|node| property_value(node, "tag")).is_some_and(contract::internal_feature_tag) {
                if let Some((source, node)) = content { self.visit(source, node, &namespaces)?; }
                return Ok(());
            }
            if self.result.controls.len() >= contract::MAX_CONTROL_COUNT { return Err("DOCX_SDT_RESOURCE_LIMIT: more than 100,000 controls".into()); }
            let native_id = child(node, "sdtPr").and_then(|node| child(node, "id")).and_then(|node| word_attribute(node, "val"));
            let signature = contract::properties_signature(child(node, "sdtPr"), child(node, "sdtEndPr"))?;
            let authored = native_id.and_then(|id| self.sources.get(id)).filter(|entry| entry.properties_signature == signature && self.used_sources.insert(entry.native_id.clone()));
            let mut value = properties(self.xml, source, node, &namespaces)?;
            let mut id = authored.map(|entry| entry.id.clone()).unwrap_or_else(|| format!("docx-sdt:{}:{}:{}:{}", self.story.part, self.story.story,
                self.story.note_id.map(|id| id.to_string()).unwrap_or_default(), native_id.unwrap_or("anonymous")));
            if !self.ids.insert(id.clone()) { id.push_str(&format!(":{}", source.start)); if !self.ids.insert(id.clone()) { return Err("Native SDT identities collide".into()); } }
            if let Some(authored) = authored {
                // Current native properties win. Authored names and dormant bindings
                // are safe only when the generated property signature still matches.
                let mut merged = authored.properties.as_object().cloned().unwrap_or_default();
                merged.extend(value.as_object().into_iter().flat_map(|value| value.iter().map(|(key, value)| (key.clone(), value.clone()))));
                value = Value::Object(merged);
            }
            let row_range = self.table.as_ref().and_then(|table| {
                let (source, node) = content?;
                let row_indexes = content_children(source, node, "tr", &scope(source, &namespaces)).into_iter()
                    .filter_map(|(source, _, _)| self.rows.get(&source.start).filter(|(owner, _)| owner == table).map(|(_, index)| *index))
                    .collect::<Vec<_>>();
                (!row_indexes.is_empty()).then(|| NativeControlRowRange { table_block_id: table.clone(), row_indexes })
            });
            let block_ids = if self.current.is_none() && row_range.is_none() {
                content.map(|(source, node)| {
                    let namespaces = scope(source, &namespaces);
                    let mut blocks = content_children(source, node, "p", &namespaces);
                    blocks.extend(content_children(source, node, "tbl", &namespaces));
                    blocks.sort_by_key(|(source, _, _)| source.start);
                    blocks.iter().map(|(source, _, _)| self.paragraphs.get(&source.start).map(|(block, _)| block.clone())
                        .or_else(|| self.tables.get(&source.start).cloned())).collect::<Option<Vec<_>>>().unwrap_or_default()
                }).unwrap_or_default()
            } else { Vec::new() };
            if let Some(properties) = value.as_object_mut() { properties.insert("placement".into(), json!(if row_range.is_some() { "row" } else if self.current.is_some() { "inline" } else { "block" })); }
            let source_xml = complete_fragment(self.xml, source, &namespaces)?;
            *self.source_bytes = self.source_bytes.checked_add(source_xml.len() + serde_json::to_vec(&value).map_err(|error| error.to_string())?.len())
                .filter(|bytes| *bytes <= contract::MAX_CONTROL_SOURCE_BYTES).ok_or("DOCX_SDT_RESOURCE_LIMIT: retained control source exceeds 16 MiB")?;
            let index = self.result.controls.len();
            let parent_id = self.active.last().map(|index| self.result.controls[*index].id.clone());
            self.result.controls.push(NativeDocumentContentControl { id, parent_id,
                properties: value, text_ranges: Vec::new(), has_object_content: false, object_content_count: 0, block_ids, table_row_range: row_range, part_name: self.story.part.into(), story: self.story.story.into(),
                note_id: self.story.note_id, source_xml });
            self.active.push(index); self.address();
            if let Some((source, node)) = content { self.visit(source, node, &namespaces)?; }
            self.active.pop(); return Ok(());
        }
        if word(node, "footnoteReference") || word(node, "endnoteReference") {
            self.address();
            if let Some(current) = &mut self.current {
                current.note_reference_count = current.note_reference_count.checked_add(1).ok_or("Native note reference count overflow")?;
                for index in &self.active {
                    if let Some(range) = self.result.controls[*index].text_ranges.iter_mut().find(|range| range.block_id == current.block) {
                        range.end_note_reference_count = Some(current.note_reference_count);
                    }
                }
            }
            return Ok(());
        }
        if word(node, "t") { return self.text(&node.text); }
        if word(node, "tab") { return self.text("\t"); }
        if word(node, "cr") || word(node, "br") && !matches!(word_attribute(node, "type"), Some("page" | "column")) { return self.text("\n"); }
        if word(node, "noBreakHyphen") { return self.text("\u{2011}"); }
        if word(node, "softHyphen") { return self.text("\u{00ad}"); }
        if word(node, "sym") {
            let character = word_attribute(node, "char").and_then(|value| u32::from_str_radix(value, 16).ok()).and_then(char::from_u32);
            if let Some(character) = character { return self.text(&character.to_string()); }
        }
        for (source, node) in source.children.iter().zip(&node.children) { self.visit(source, node, &namespaces)?; }
        Ok(())
    }
}

pub(crate) fn restore(document: &StructuredDocument, source: &DocModel, archive: &mut ZipArchive<Cursor<&[u8]>>) -> Result<Restoration, String> {
    let sources = read_sources(archive)?;
    let mut result = Restoration::default(); let mut blocks = HashMap::new();
    collect_blocks(&document.blocks, None, &mut blocks);
    let mut source_bytes = 0usize; let mut ids = HashSet::new(); let mut used_sources = HashSet::new();
    for story in stories(source) {
        let Some(xml) = read_xml(archive, story.part)? else { continue; };
        let native = chart_xml::parse(&xml)?;
        if !native.descendants("sdt").iter().any(|node| word(node, "sdt")) { continue; }
        let root = parse_xml(&xml)?;
        let part_scope = scope(&root, &BTreeMap::new());
        let Some((root, native)) = story_root(&root, &native, &story) else {
            result.warnings.push(format!("Native SDT in {} {} needs an additional editable story mapping; its source model remains retained", story.part, story.story)); continue;
        };
        let mut raw_paragraphs = Vec::new(); let mut raw_tables = Vec::new();
        raw_hosts(root, native, &mut raw_paragraphs, &mut raw_tables);
        let source_nodes = collect(&story); let mut paragraphs = HashMap::new(); let mut tables = HashMap::new(); let mut rows = HashMap::new();
        if raw_paragraphs.len() == source_nodes.paragraphs.len() && raw_tables.len() == source_nodes.tables.len() {
            for ((raw, _), (address, _)) in raw_paragraphs.iter().zip(&source_nodes.paragraphs) {
                if let Some(block) = blocks.get(address) { paragraphs.insert(raw.start, block.clone()); }
            }
            for ((raw, native), (address, source)) in raw_tables.iter().zip(&source_nodes.tables) {
                if let Some((block, _)) = blocks.get(address) {
                    tables.insert(raw.start, block.clone());
                    let native_rows = content_children(raw, native, "tr", &BTreeMap::new());
                    if native_rows.len() == source.rows.len() { for (index, (row, _, _)) in native_rows.into_iter().enumerate() { rows.insert(row.start, (block.clone(), index)); } }
                }
            }
        } else { result.warnings.push(format!("Native SDT addresses in {} do not match extracted paragraph/table counts; exact source is retained without guessed ranges", story.part)); }
        let start = result.controls.len();
        let namespaces = scope(root, &part_scope);
        let mut walker = Walker { xml: &xml, story: &story, sources: &sources, used_sources: &mut used_sources, paragraphs, tables, rows, result: &mut result,
            active: Vec::new(), current: None, table: None, invalid_blocks: HashSet::new(), source_bytes: &mut source_bytes, ids: &mut ids };
        walker.visit(root, native, &namespaces)?;
        let invalid = std::mem::take(&mut walker.invalid_blocks);
        for control in &mut walker.result.controls[start..] { control.text_ranges.retain(|range| !invalid.contains(&range.block_id)); }
    }
    for entry in sources.values().filter(|entry| !used_sources.contains(&entry.native_id)) {
        result.warnings.push(format!("Authored content-control properties for {:?} are stale or no longer have a native wrapper; they were not applied to new content", entry.id));
    }
    Ok(result)
}
