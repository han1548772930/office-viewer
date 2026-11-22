use std::collections::{HashMap, HashSet};
use std::io::Cursor;
use std::ops::Range;

use duke_pages::{DocModel, DocNode, NumberingDefinitionSet, ParagraphChildNode, ParagraphNode, ParagraphNumbering, StructuredDocument, TableCellContentNode, TableNode};
use quick_xml::{Reader, XmlVersion, events::Event};
use serde_json::{Value, json};
use zip::ZipArchive;

use super::{Element, XML_ELEMENT_LIMIT, parse_xml, part_for_type, read_xml, restore_numbering};

const NAMESPACE: &str = "urn:office-studio:numbering:1";
const BOOKMARK_PREFIX: &str = "office_list_v1_";

#[derive(Default)]
pub(crate) struct NumberingRestoration {
    pub definitions: Option<Value>,
    pub warnings: Vec<String>,
    details: HashMap<String, ParagraphExtras>,
}

#[derive(Default)]
struct ParagraphExtras {
    list_id: Option<String>,
    label: Option<String>,
    preserved: Option<Value>,
}

fn array(value: &Value) -> &[Value] { value.as_array().map_or(&[], Vec::as_slice) }
fn object(value: &Value) -> Value { if value.is_object() { value.clone() } else { json!({}) } }

fn normalize_instruction(code: &str) -> String {
    let mut output = String::new();
    let mut quoted = false;
    let mut escaped = false;
    let mut space = false;
    for character in code.trim().chars() {
        if !quoted && character.is_whitespace() { space = true; continue; }
        if space && !output.is_empty() { output.push(' '); }
        space = false;
        output.push(character);
        if character == '"' && !escaped { quoted = !quoted; }
        escaped = character == '\\' && !escaped;
    }
    output
}

fn resolve_part(owner: &str, target: &str) -> Result<String, String> {
    if target.contains('\\') || target.contains(':') || target.contains('#') || target.contains('?') { return Err("The numbering mapping has an invalid package target.".into()); }
    let mut parts = if target.starts_with('/') { Vec::new() } else { owner.rsplit_once('/').map_or(Vec::new(), |(parent, _)| parent.split('/').map(str::to_owned).collect()) };
    for part in target.trim_start_matches('/').split('/') {
        match part {
            "" | "." => {},
            ".." => { parts.pop().ok_or("The numbering mapping target leaves the package.")?; },
            _ => parts.push(part.to_owned()),
        }
    }
    if parts.is_empty() { return Err("The numbering mapping has an empty package target.".into()); }
    Ok(parts.join("/"))
}

fn read_mapping(archive: &mut ZipArchive<Cursor<&[u8]>>) -> Result<Option<Value>, String> {
    let owner = part_for_type(archive, ".document.main+xml", "word/document.xml")?;
    let (parent, file) = owner.rsplit_once('/').unwrap_or(("", owner.as_str()));
    let relationships = if parent.is_empty() { format!("_rels/{file}.rels") } else { format!("{parent}/_rels/{file}.rels") };
    let Some(xml) = read_xml(archive, &relationships)? else { return Ok(None); };
    let root = parse_xml(&xml)?;
    let mut found = None;
    let mut count = 0;
    for relationship in &root.children {
        if relationship.name != "Relationship" || relationship.attribute("Type") != Some("http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml") || relationship.attribute("TargetMode") == Some("External") { continue; }
        count += 1;
        if count > 128 { return Err("The document has too many custom XML relationships.".into()); }
        let Some(target) = relationship.attribute("Target") else { continue; };
        let part = resolve_part(&owner, target)?;
        let Some(xml) = read_xml(archive, &part)? else { continue; };
        let root = parse_xml(&xml)?;
        let prefix = root.qualified_name.split_once(':').map_or("", |(prefix, _)| prefix);
        if root.name != "officeNumbering" || root.namespaces.get(prefix).map(String::as_str) != Some(NAMESPACE) { continue; }
        if found.is_some() { return Err("The document has more than one numbering mapping.".into()); }
        if root.attribute("version") != Some("1") { return Err("The numbering mapping uses an unsupported version.".into()); }
        let mapping = root.child("mapping").ok_or("The numbering mapping has no data.")?;
        if !mapping.children.is_empty() { return Err("The numbering mapping data contains unexpected XML elements.".into()); }
        let data: Value = serde_json::from_str(&mapping.text).map_err(|error| format!("The numbering mapping JSON is invalid: {error}"))?;
        if data["version"].as_u64() != Some(1) || !data["instances"].is_array() || !data["paragraphs"].is_array() || array(&data["instances"]).len() > XML_ELEMENT_LIMIT || array(&data["paragraphs"]).len() > XML_ELEMENT_LIMIT { return Err("The numbering mapping has an invalid schema or size.".into()); }
        found = Some(data);
    }
    Ok(found)
}

fn document_namespaces(archive: &mut ZipArchive<Cursor<&[u8]>>) -> Result<HashMap<String, String>, String> {
    let owner = part_for_type(archive, ".document.main+xml", "word/document.xml")?;
    let Some(xml) = read_xml(archive, &owner)? else { return Ok(HashMap::new()); };
    let mut reader = Reader::from_str(&xml);
    loop {
        match reader.read_event().map_err(|error| error.to_string())? {
            Event::Start(tag) | Event::Empty(tag) => {
                let mut namespaces = HashMap::new();
                for attribute in tag.attributes() {
                    let attribute = attribute.map_err(|error| error.to_string())?;
                    let name = String::from_utf8_lossy(attribute.key.as_ref());
                    let prefix = if name == "xmlns" { Some("") } else { name.strip_prefix("xmlns:") };
                    if let Some(prefix) = prefix {
                        let value = attribute.decoded_and_normalized_value_with(XmlVersion::Implicit1_0, reader.decoder(), 1, quick_xml::escape::resolve_xml_entity).map_err(|error| error.to_string())?;
                        namespaces.insert(prefix.to_owned(), value.into_owned());
                    }
                }
                return Ok(namespaces);
            }
            Event::Eof => return Ok(HashMap::new()),
            Event::DocType(_) => return Err("Office XML must not declare a document type.".into()),
            _ => {},
        }
    }
}

fn patch_raw_level(level: &mut Value, xml: &Element) {
    if let Some(format) = xml.child("numFmt") {
        if let Some(value) = format.attribute("val") { level["format"] = json!(value); }
        if let Some(pattern) = format.attribute("format") { level["customFormat"] = json!(pattern); }
    }
}

fn enhance_native_definitions(definitions: &mut Value, root: &Element) {
    for node in &root.children {
        if node.name == "abstractNum" {
            let Some(id) = node.integer("abstractNumId") else { continue; };
            let Some(definition) = definitions["abstracts"].as_array_mut().and_then(|values| values.iter_mut().find(|value| value["abstractNumId"].as_i64() == Some(id))) else { continue; };
            if let Some(kind) = node.child("multiLevelType").and_then(|child| child.attribute("val")) { definition["multiLevelType"] = json!(kind); }
            if let Some(name) = node.child("name").and_then(|child| child.attribute("val")) { definition["name"] = json!(name); }
            for xml in node.children.iter().filter(|child| child.name == "lvl") {
                if let Some(level) = definition["levels"].as_array_mut().and_then(|levels| levels.iter_mut().find(|level| level["ilvl"].as_i64() == xml.integer("ilvl"))) { patch_raw_level(level, xml); }
            }
        } else if node.name == "num" {
            let Some(id) = node.integer("numId") else { continue; };
            let Some(instance) = definitions["instances"].as_array_mut().and_then(|values| values.iter_mut().find(|value| value["numId"].as_i64() == Some(id))) else { continue; };
            for level in node.children.iter().filter(|child| child.name == "lvlOverride").filter_map(|child| child.child("lvl")) {
                if let Some(target) = instance["levelOverrides"].as_array_mut().and_then(|levels| levels.iter_mut().find(|value| value["ilvl"].as_i64() == level.integer("ilvl"))) { patch_raw_level(target, level); }
            }
        }
    }
}

fn canonical_xml(node: &Element) -> Value {
    let attributes = node.attributes.iter().map(|(key, value)| (key.clone(), json!(value))).collect::<serde_json::Map<_, _>>();
    json!({ "name": node.name, "attributes": attributes, "text": node.text.trim(), "children": node.children.iter().map(canonical_xml).collect::<Vec<_>>() })
}

fn native_contract_matches(root: Option<&Element>, expected: &Value, num_id: i64) -> bool {
    let (Some(root), Some(xml)) = (root, expected.as_str()) else { return false; };
    let Ok(expected) = parse_xml(xml) else { return false; };
    let Some(number) = root.children.iter().find(|node| node.name == "num" && node.integer("numId") == Some(num_id)) else { return false; };
    let Some(id) = number.child("abstractNumId").and_then(|node| node.integer("val")) else { return false; };
    let Some(base) = root.children.iter().find(|node| node.name == "abstractNum" && node.integer("abstractNumId") == Some(id)) else { return false; };
    expected.child("abstractNum").is_some_and(|node| canonical_xml(node) == canonical_xml(base)) && expected.child("num").is_some_and(|node| canonical_xml(node) == canonical_xml(number))
}

#[derive(Default)]
struct FieldState { code: String, separated: bool }

#[derive(Default)]
struct Marker {
    name: String,
    bookmark_id: i64,
    start: usize,
    end: usize,
    contract: Vec<Value>,
    raw_text: String,
    display_text: String,
    styles: Vec<Value>,
    visible_ranges: Vec<Range<usize>>,
    visible_nodes: HashSet<usize>,
    commented: bool,
}

fn style_signature(properties: &Element) -> Value {
    let fonts = properties.child("rFonts");
    let flag = |name: &str| properties.child(name).is_some_and(Element::on_off);
    json!({
        "font": fonts.and_then(|fonts| fonts.attribute("ascii")).unwrap_or("Arial"),
        "eastAsia": fonts.and_then(|fonts| fonts.attribute("eastAsia")).or_else(|| fonts.and_then(|fonts| fonts.attribute("ascii"))).unwrap_or("Arial"),
        "color": properties.child("color").and_then(|node| node.attribute("val")).unwrap_or("1F2937").to_ascii_uppercase(),
        "size": properties.child("sz").and_then(|node| node.integer("val")).unwrap_or(22),
        "bold": flag("b"), "italic": flag("i"), "strike": flag("strike"),
        "underline": properties.child("u").is_some_and(|node| !matches!(node.attribute("val"), Some("none" | "0" | "false" | "off"))),
    })
}

fn append_text(marker: &mut Marker, text: &str, node: &Element, fields: &[FieldState], comments: &HashSet<i64>) {
    let start = marker.raw_text.len();
    marker.raw_text.push_str(text);
    if fields.iter().all(|field| field.separated) {
        marker.display_text.push_str(text);
        marker.visible_ranges.push(start..marker.raw_text.len());
        marker.visible_nodes.insert(node.start);
        marker.commented |= !text.is_empty() && !comments.is_empty();
    }
    if fields.is_empty() && !text.is_empty() {
        if let Some(last) = marker.contract.last_mut().filter(|last| last["text"].is_string()) { last["text"] = json!(format!("{}{text}", last["text"].as_str().unwrap_or_default())); }
        else { marker.contract.push(json!({ "text": text })); }
    }
}

fn comment_boundary(node: &Element, comments: &mut HashSet<i64>) -> Result<(), String> {
    let id = node.integer("id").filter(|id| *id >= 0).ok_or("A numbering marker comment has an invalid ID.")?;
    match node.name.as_str() {
        "commentRangeStart" => { comments.insert(id); },
        "commentRangeEnd" => { comments.remove(&id); },
        _ => {},
    }
    Ok(())
}

fn track_comment_boundaries(node: &Element, comments: &mut HashSet<i64>) -> Result<(), String> {
    if matches!(node.name.as_str(), "commentRangeStart" | "commentRangeEnd") { comment_boundary(node, comments)?; }
    else if node.name != "pPr" { for child in &node.children { track_comment_boundaries(child, comments)?; } }
    Ok(())
}

fn marker_from_paragraph(root: &Element, inherited_comments: &HashSet<i64>) -> Result<Option<Marker>, String> {
    let starts = root.children.iter().enumerate().filter(|(_, child)| child.name == "bookmarkStart" && child.attribute("name").is_some_and(|name| name.starts_with(BOOKMARK_PREFIX))).collect::<Vec<_>>();
    if starts.is_empty() { return Ok(None); }
    if starts.len() != 1 { return Err("The paragraph has ambiguous numbering markers.".into()); }
    let (start_index, bookmark) = starts[0];
    let bookmark_id = bookmark.integer("id").ok_or("The numbering marker has no bookmark ID.")?;
    let (end_index, end) = root.children.iter().enumerate().skip(start_index + 1).find(|(_, child)| child.name == "bookmarkEnd" && child.integer("id") == Some(bookmark_id)).ok_or("The numbering marker has no matching bookmark end.")?;
    let mut marker = Marker { name: bookmark.attribute("name").unwrap_or_default().to_owned(), bookmark_id, start: bookmark.start, end: end.end, ..Default::default() };
    let mut fields = Vec::<FieldState>::new();
    let mut comments = inherited_comments.clone();
    for node in &root.children[..start_index] { track_comment_boundaries(node, &mut comments)?; }
    fn visit(node: &Element, marker: &mut Marker, fields: &mut Vec<FieldState>, comments: &mut HashSet<i64>) -> Result<(), String> {
        match node.name.as_str() {
            "rPr" => {},
            "commentRangeStart" | "commentRangeEnd" => comment_boundary(node, comments)?,
            "commentReference" => {
                comment_boundary(node, comments)?;
                marker.commented = true;
            },
            "fldChar" => match node.attribute("fldCharType") {
                Some("begin") => fields.push(FieldState::default()),
                Some("separate") => { let field = fields.last_mut().ok_or("A marker field has no beginning.")?; if field.separated { return Err("A marker field has more than one separator.".into()); } field.separated = true; },
                Some("end") => {
                    let field = fields.pop().ok_or("A marker field has no beginning.")?;
                    if !field.separated { return Err("A marker field has no result separator.".into()); }
                    let code = normalize_instruction(&field.code);
                    if let Some(parent) = fields.last_mut().filter(|parent| !parent.separated) { parent.code.push_str(&format!("{{{code}}}")); }
                    else if fields.is_empty() { marker.contract.push(json!({ "field": code })); }
                },
                _ => return Err("A marker field has an unknown delimiter.".into()),
            },
            "instrText" => {
                let field = fields.last_mut().filter(|field| !field.separated).ok_or("A marker instruction is outside its field.")?;
                field.code.push_str(&node.text);
            },
            "t" => append_text(marker, &node.text.replace("\r\n", " ").replace(['\r', '\n'], " "), node, fields, comments),
            "tab" => append_text(marker, "\t", node, fields, comments),
            "br" | "cr" => append_text(marker, "\n", node, fields, comments),
            "noBreakHyphen" => append_text(marker, "\u{2011}", node, fields, comments),
            "softHyphen" => append_text(marker, "\u{ad}", node, fields, comments),
            "r" => {
                if node.children.iter().any(|child| !matches!(child.name.as_str(), "rPr" | "commentReference")) {
                    if let Some(properties) = node.child("rPr") { marker.styles.push(style_signature(properties)); }
                    else { marker.styles.push(json!({})); }
                }
                for child in &node.children { visit(child, marker, fields, comments)?; }
            },
            _ => return Err(format!("The marker contains unsupported Office content: {}.", node.name)),
        }
        Ok(())
    }
    for node in &root.children[start_index + 1..end_index] { visit(node, &mut marker, &mut fields, &mut comments)?; }
    if !fields.is_empty() { return Err("A marker field is not closed.".into()); }
    Ok(Some(marker))
}

fn paragraph_numbering(root: &Element) -> Option<(i64, i64)> {
    let number = root.child("pPr")?.child("numPr")?;
    Some((number.child("numId")?.integer("val")?, number.child("ilvl").and_then(|node| node.integer("val")).unwrap_or(0)))
}

fn trim_marker_children(paragraph: &ParagraphNode, marker: &Marker) -> Option<(Vec<ParagraphChildNode>, usize)> {
    let mut children = paragraph.children.clone();
    let mut remaining = marker.raw_text.as_str();
    let mut removed = 0usize;
    for child in &mut children {
        if remaining.is_empty() { break; }
        let ParagraphChildNode::Text(run) = child else { return None; };
        if run.note_reference.is_some() { return None; }
        if run.text.is_empty() { removed += 1; continue; }
        let length = remaining.len().min(run.text.len());
        if !run.text.is_char_boundary(length) || !remaining.is_char_boundary(length) || run.text[..length] != remaining[..length] { return None; }
        remaining = &remaining[length..];
        if length == run.text.len() { removed += 1; }
        else { run.text = run.text[length..].to_owned(); break; }
    }
    if !remaining.is_empty() { return None; }
    children.drain(..removed);
    Some((children, removed))
}

fn visible_marker_children(paragraph: &ParagraphNode, marker: &Marker) -> Option<Vec<ParagraphChildNode>> {
    let mut result = Vec::new();
    let mut offset = 0;
    for child in &paragraph.children {
        if offset >= marker.raw_text.len() { break; }
        let ParagraphChildNode::Text(run) = child else { return None; };
        let end = (offset + run.text.len()).min(marker.raw_text.len());
        for visible in &marker.visible_ranges {
            let start = offset.max(visible.start);
            let finish = end.min(visible.end);
            if start >= finish { continue; }
            let mut visible_run = run.clone();
            visible_run.text = run.text.get(start - offset..finish - offset)?.to_owned();
            visible_run.note_reference = None;
            visible_run.field_instruction = None;
            result.push(ParagraphChildNode::Text(visible_run));
        }
        offset = end;
    }
    let text = result.iter().filter_map(|child| match child { ParagraphChildNode::Text(run) => Some(run.text.as_str()), _ => None }).collect::<String>();
    (offset == marker.raw_text.len() && text == marker.display_text).then_some(result)
}

fn opening_tag<'a>(xml: &'a str, node: &Element) -> &'a str {
    let mut quote = None;
    for (offset, byte) in xml.as_bytes()[node.start..node.end].iter().copied().enumerate() {
        match (quote, byte) {
            (Some(current), byte) if current == byte => quote = None,
            (None, b'\'' | b'"') => quote = Some(byte),
            (None, b'>') => return &xml[node.start..node.start + offset + 1],
            _ => {},
        }
    }
    ""
}

/** Review extraction uses source XML to enumerate runs; keep it aligned with the visible children. */
fn projected_marker_xml(xml: &str, root: &Element, marker: &Marker, keep_text: bool) -> String {
    fn project(xml: &str, node: &Element, marker: &Marker, keep_text: bool) -> String {
        if matches!(node.name.as_str(), "commentRangeStart" | "commentRangeEnd" | "commentReference") {
            return xml[node.start..node.end].to_owned();
        }
        if keep_text && marker.visible_nodes.contains(&node.start) { return xml[node.start..node.end].to_owned(); }
        if node.name != "r" { return String::new(); }
        let content = node.children.iter().filter(|node| node.name != "rPr").map(|child| project(xml, child, marker, keep_text)).collect::<String>();
        if content.is_empty() { return String::new(); }
        let properties = node.child("rPr").map(|properties| &xml[properties.start..properties.end]).unwrap_or_default();
        format!("{}{properties}{content}</{}>", opening_tag(xml, node), node.qualified_name)
    }
    root.children.iter().filter(|node| node.start > marker.start && node.end < marker.end).map(|node| project(xml, node, marker, keep_text)).collect()
}

fn update_source_positions(paragraph: &mut ParagraphNode, removed: usize, inserted: usize, bookmark: &str) {
    paragraph.bookmarks.retain(|value| value.name != bookmark);
    let shift = |index: usize| if index >= removed { index - removed + inserted } else { 0 };
    for bookmark in &mut paragraph.bookmarks { bookmark.child_index = shift(bookmark.child_index); }
    if let Some(comments) = &mut paragraph.comment_ranges {
        for comment in comments { comment.child_start = shift(comment.child_start); comment.child_end = shift(comment.child_end); }
    }
    paragraph.source_text_patch = None;
    paragraph.source_run_provenance = None;
}

fn visit_paragraphs(nodes: &mut [DocNode], visitor: &mut impl FnMut(&mut ParagraphNode, String)) {
    fn table(value: &mut TableNode, path: &str, visitor: &mut impl FnMut(&mut ParagraphNode, String)) {
        let root = value.block_id.clone().unwrap_or_else(|| path.to_owned());
        for (row_index, row) in value.rows.iter_mut().enumerate() { for (cell_index, cell) in row.cells.iter_mut().enumerate() { for (index, node) in cell.nodes.iter_mut().enumerate() {
            let path = format!("{root}/row/{row_index}/cell/{cell_index}/{index}");
            match node {
                TableCellContentNode::Paragraph(paragraph) => { let id = paragraph.block_id.clone().unwrap_or(path); visitor(paragraph, id); },
                TableCellContentNode::Table(value) => table(value, &path, visitor),
            }
        } } }
    }
    for (index, node) in nodes.iter_mut().enumerate() {
        let path = format!("document/{index}");
        match node {
            DocNode::Paragraph(paragraph) => { let id = paragraph.block_id.clone().unwrap_or(path); visitor(paragraph, id); },
            DocNode::Table(value) => table(value, &path, visitor),
        }
    }
}

fn prepare_definitions(current: &Value, mapping: &Value) -> Result<(Value, HashMap<i64, i64>), String> {
    let mut definitions = current.clone();
    if !definitions.is_object() { definitions = json!({ "abstracts": [], "instances": [] }); }
    let mut abstracts = array(&definitions["abstracts"]).to_vec();
    let mut instances = array(&definitions["instances"]).to_vec();
    let mut next_abstract = abstracts.iter().filter_map(|value| value["abstractNumId"].as_i64()).max().unwrap_or(0).saturating_add(1);
    let mut next_instance = instances.iter().filter_map(|value| value["numId"].as_i64()).max().unwrap_or(0).saturating_add(1);
    let mut instance_ids = HashMap::new();
    let mut abstract_ids = HashMap::new();
    for entry in array(&mapping["instances"]) {
        let id = entry["numId"].as_i64().filter(|id| *id > 0).ok_or("The mapping has an invalid numbering instance ID.")?;
        if instance_ids.insert(id, next_instance).is_some() { return Err("The mapping has duplicate numbering instance IDs.".into()); }
        next_instance = next_instance.checked_add(1).ok_or("The mapping numbering IDs overflow.")?;
    }
    let mut counter_ids = instance_ids.clone();
    for entry in array(&mapping["instances"]) {
        let id = entry["numId"].as_i64().unwrap_or_default();
        let definition = &entry["definition"];
        if !definition["abstract"].is_object() || !definition["instance"].is_object() { return Err("The mapping has an invalid numbering definition.".into()); }
        let mut base = object(&definition["abstract"]);
        if array(&base["levels"]).is_empty() { base["levels"] = definition["levels"].clone(); }
        if array(&base["levels"]).is_empty() || array(&base["levels"]).len() > 9 || array(&base["levels"]).iter().any(|level| !level["ilvl"].as_i64().is_some_and(|level| (0..=8).contains(&level))) { return Err("The mapping has invalid numbering levels.".into()); }
        let signature = base.to_string();
        let abstract_id = if let Some(id) = abstract_ids.get(&signature) { *id } else {
            let id = next_abstract;
            next_abstract = next_abstract.checked_add(1).ok_or("The mapping abstract IDs overflow.")?;
            abstract_ids.insert(signature, id);
            if base["sourceAbstractNumId"].is_null() { base["sourceAbstractNumId"] = base["abstractNumId"].clone(); }
            base["abstractNumId"] = json!(id);
            abstracts.push(base);
            id
        };
        let mut instance = definition["instance"].clone();
        if instance["sourceNumId"].is_null() { instance["sourceNumId"] = instance["numId"].clone(); }
        instance["numId"] = json!(instance_ids[&id]);
        instance["abstractNumId"] = json!(abstract_id);
        if definition["abstract"]["consecutiveNumbering"].as_bool() == Some(true) {
            let counter = entry["counterId"].as_i64().filter(|counter| *counter > 0).ok_or("The mapping has an invalid consecutive counter ID.")?;
            let counter = *counter_ids.entry(counter).or_insert_with(|| { let id = next_instance; next_instance = next_instance.saturating_add(1); id });
            instance["consecutiveListId"] = json!(counter);
        }
        instances.push(instance);
    }
    definitions["abstracts"] = json!(abstracts);
    definitions["instances"] = json!(instances);
    serde_json::from_value::<NumberingDefinitionSet>(definitions.clone()).map_err(|error| format!("The numbering mapping has invalid typed fields: {error}"))?;
    Ok((definitions, instance_ids))
}

fn preserve_office_marker(paragraph: &mut ParagraphNode, xml: &str, root: &Element, marker: &Marker, namespaces: &HashMap<String, String>, extras: &mut ParagraphExtras) -> bool {
    let Some((mut children, removed)) = trim_marker_children(paragraph, marker) else { return false; };
    let Some(mut visible) = visible_marker_children(paragraph, marker) else { return false; };
    let inserted = visible.len();
    visible.append(&mut children);
    let children = visible;
    paragraph.children = children;
    update_source_positions(paragraph, removed, inserted, &marker.name);
    paragraph.source_xml = Some(format!("{}{}{}", &xml[..marker.start], projected_marker_xml(xml, root, marker, true), &xml[marker.end..]));
    let mut namespaces = namespaces.clone();
    namespaces.extend(root.namespaces.clone());
    extras.preserved = Some(json!({ "version": 1, "bookmark": marker.name, "bookmarkId": marker.bookmark_id, "xml": &xml[marker.start..marker.end], "text": marker.display_text, "namespaces": namespaces, "sourceNumbering": paragraph_numbering(root).map(|(num_id, ilvl)| json!({ "numId": num_id, "ilvl": ilvl })) }));
    true
}

pub(crate) fn restore_numbering_with_extensions(source: &mut DocModel, archive: &mut ZipArchive<Cursor<&[u8]>>) -> NumberingRestoration {
    let mut result = NumberingRestoration { warnings: restore_numbering(source, archive), ..Default::default() };
    let mut native = source.metadata.numbering_definitions.as_ref().and_then(|definitions| serde_json::to_value(definitions).ok()).unwrap_or_else(|| json!({ "abstracts": [], "instances": [] }));
    let native_root = part_for_type(archive, ".numbering+xml", "word/numbering.xml").and_then(|part| read_xml(archive, &part)).and_then(|xml| xml.map(|xml| parse_xml(&xml)).transpose()).ok().flatten();
    if let Some(root) = &native_root { enhance_native_definitions(&mut native, root); }
    let mapping = match read_mapping(archive) {
        Ok(mapping) => mapping,
        Err(error) => { result.warnings.push(format!("DOCX numbering mapping was ignored: {error}")); None },
    };
    let prepared = mapping.as_ref().and_then(|mapping| match prepare_definitions(&native, mapping) {
        Ok(prepared) => Some(prepared),
        Err(error) => { result.warnings.push(format!("DOCX numbering mapping was ignored: {error}")); None },
    });
    let mut entries = HashMap::<String, &Value>::new();
    let mut ambiguous = HashMap::<String, usize>::new();
    if let Some(mapping) = &mapping {
        for entry in array(&mapping["paragraphs"]) {
            if let Some(name) = entry["bookmark"].as_str() { *ambiguous.entry(name.to_owned()).or_default() += 1; entries.insert(name.to_owned(), entry); }
        }
    }
    let instances = mapping.as_ref().map(|mapping| array(&mapping["instances"]).iter().filter_map(|entry| Some((entry["numId"].as_i64()?, entry))).collect::<HashMap<_, _>>()).unwrap_or_default();
    let mut actual_names = HashMap::<String, usize>::new();
    let mut has_markers = false;
    visit_paragraphs(&mut source.nodes, &mut |paragraph, _| {
        has_markers |= paragraph.source_xml.as_ref().is_some_and(|xml| xml.contains(BOOKMARK_PREFIX));
        for bookmark in &paragraph.bookmarks { if bookmark.name.starts_with(BOOKMARK_PREFIX) { *actual_names.entry(bookmark.name.clone()).or_default() += 1; } }
    });
    let namespaces = if has_markers { document_namespaces(archive).unwrap_or_else(|error| { result.warnings.push(format!("DOCX numbering marker namespaces could not be read: {error}")); HashMap::new() }) } else { HashMap::new() };
    let mut active_comments = HashSet::<i64>::new();
    visit_paragraphs(&mut source.nodes, &mut |paragraph, id| {
        if !has_markers { return; }
        let Some(original) = paragraph.source_xml.as_ref() else { return; };
        let has_marker = original.contains(BOOKMARK_PREFIX);
        if !has_marker && !original.contains("commentRangeStart") && !original.contains("commentRangeEnd") { return; }
        let xml = original.clone();
        let root = match parse_xml(&xml) { Ok(root) => root, Err(error) => { result.warnings.push(format!("DOCX numbering marker in {id} could not be read: {error}")); return; } };
        let inherited_comments = active_comments.clone();
        if let Err(error) = track_comment_boundaries(&root, &mut active_comments) { result.warnings.push(format!("DOCX numbering comment context in {id} could not be fully read: {error}")); }
        if !has_marker { return; }
        let marker = match marker_from_paragraph(&root, &inherited_comments) { Ok(Some(marker)) => marker, Ok(None) => return, Err(error) => { result.warnings.push(format!("DOCX numbering marker in {id} was retained: {error}")); return; } };
        let mut extras = ParagraphExtras::default();
        let restore = (|| {
            // A semantic list label has no text anchor. Retain annotated Office labels as fields.
            if marker.commented && !marker.display_text.is_empty() { return None; }
            let (definitions, ids) = prepared.as_ref()?;
            let entry = *entries.get(&marker.name)?;
            if ambiguous.get(&marker.name).copied().unwrap_or(0) != 1 || actual_names.get(&marker.name).copied().unwrap_or(0) != 1 { return None; }
            let num_id = entry["numId"].as_i64()?;
            let index = entry["ilvl"].as_i64().filter(|index| (0..=8).contains(index))?;
            let instance = *instances.get(&num_id)?;
            let mode = entry["mode"].as_str()?;
            let numbered = paragraph_numbering(&root).filter(|(id, _)| *id != 0);
            if mode == "sequence" { if numbered.is_some() { return None; } }
            else if matches!(mode, "native" | "sequence-native-marker") {
                if numbered != Some((num_id, index)) || !native_contract_matches(native_root.as_ref(), &instance["nativeXml"], num_id) { return None; }
            } else { return None; }
            if json!(marker.contract) != entry["contract"] { return None; }
            if mode != "native" && (marker.styles.is_empty() || marker.styles.iter().any(|style| *style != entry["markerStyle"])) { return None; }
            let new_id = *ids.get(&num_id)?;
            let (children, removed) = trim_marker_children(paragraph, &marker)?;
            if !definitions["instances"].as_array()?.iter().any(|instance| instance["numId"].as_i64() == Some(new_id)) { return None; }
            Some((new_id, index, children, removed, instance, mode))
        })();
        if let Some((num_id, index, children, removed, instance, mode)) = restore {
            paragraph.children = children;
            paragraph.style.get_or_insert_with(Default::default).numbering = Some(ParagraphNumbering { num_id, ilvl: index });
            update_source_positions(paragraph, removed, 0, &marker.name);
            paragraph.source_xml = Some(format!("{}{}{}", &xml[..marker.start], projected_marker_xml(&xml, &root, &marker, false), &xml[marker.end..]));
            extras.list_id = instance["listId"].as_str().filter(|name| !name.is_empty()).map(str::to_owned);
            if mode == "sequence" {
                let suffix = instance["definition"]["levels"].as_array().and_then(|levels| levels.iter().find(|level| level["ilvl"].as_i64() == Some(index))).and_then(|level| level["suffix"].as_str()).unwrap_or("tab");
                extras.label = Some(match suffix { "nothing" => marker.display_text.clone(), "space" => marker.display_text.strip_suffix(' ').unwrap_or(&marker.display_text).to_owned(), _ => marker.display_text.strip_suffix('\t').unwrap_or(&marker.display_text).to_owned() });
            }
        } else {
            let preserved = preserve_office_marker(paragraph, &xml, &root, &marker, &namespaces, &mut extras);
            let reason = if marker.commented && !marker.display_text.is_empty() {
                "the visible Office label intersects a comment and semantic labels have no editable text anchor"
            } else { "the numbering mapping is missing or differs from the current fields" };
            result.warnings.push(format!("DOCX numbering extension for {id} was not applied because {reason}; {}.", if preserved { "current Office fields and their visible result were retained" } else { "current paragraph content was retained" }));
        }
        result.details.insert(id, extras);
    });
    let definitions = prepared.map(|(definitions, _)| definitions).unwrap_or(native);
    if !array(&definitions["abstracts"]).is_empty() || !array(&definitions["instances"]).is_empty() {
        if let Ok(typed) = serde_json::from_value::<NumberingDefinitionSet>(definitions.clone()) { source.metadata.numbering_definitions = Some(typed); }
        result.definitions = Some(definitions);
    }
    result
}

pub(crate) fn restore_numbering_details(document: &mut StructuredDocument, restoration: &NumberingRestoration) {
    fn visit(blocks: &mut [duke_pages::DocumentBlock], details: &HashMap<String, ParagraphExtras>) {
        for block in blocks {
            let extra = block.source_refs.iter().filter(|source| source.story == "body").find_map(|source| details.get(&source.node_id).or_else(|| source.node_id.rfind("/body/").and_then(|index| details.get(&source.node_id[index..]))));
            if let Some(extra) = extra {
                let mut value = block.details.to_json_value();
                if value["kind"].as_str() == Some("paragraph") {
                    if let Some(id) = &extra.list_id { value["numberingListId"] = json!(id); }
                    if let Some(label) = &extra.label { value["listLabel"] = json!(label); }
                    if let Some(marker) = &extra.preserved { value["officeNumberingMarker"] = marker.clone(); }
                    block.details = value.into();
                }
            }
            visit(&mut block.children, details);
        }
    }
    visit(&mut document.blocks, &restoration.details);
}
