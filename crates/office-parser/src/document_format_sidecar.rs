use std::{collections::HashMap, io::{Cursor, Read}};

use duke_pages::{BlockDetails, DocModel, DocNode, DocumentBlock, NumberingDefinitionSet, NumberingLevelDefinition, ParagraphNode, StructuredDocument, TableCellContentNode, TableNode};
use quick_xml::{Reader, XmlVersion, events::Event};
use serde_json::{Value, json};
use zip::ZipArchive;

const XML_LIMIT: u64 = 16 * 1024 * 1024;
const XML_DEPTH_LIMIT: usize = 128;
const XML_ELEMENT_LIMIT: usize = 262_144;

use crate::model::DocumentOfficeXmlPartPayload;
#[path = "document_numbering_source.rs"]
mod numbering_source;

#[path = "document_numbering_extension.rs"]
mod numbering_extension;
pub(crate) use numbering_extension::{NumberingRestoration, restore_numbering_with_extensions, restore_numbering_details};
#[path = "document_comment_sidecar.rs"]
mod comments;
pub(crate) use comments::restore_sdt_comment_targets;
#[path = "document_default_style.rs"]
mod default_style;
pub(crate) use default_style::{canonical_resource, read_document_default_style, read_document_style_sheet};
#[path = "document_legacy_default_style.rs"]
mod legacy_default_style;
pub(crate) use legacy_default_style::{read_doc_default_style, read_odt_default_style};
#[path = "document_legacy_source_sidecar.rs"]
mod legacy_source_sidecar;
pub(crate) use legacy_source_sidecar::read_legacy_origins;

#[derive(Default)]
pub(crate) struct Element {
    pub(crate) name: String,
    pub(crate) qualified_name: String,
    pub(crate) attributes: HashMap<String, String>,
    pub(crate) namespaces: HashMap<String, String>,
    pub(crate) text: String,
    pub(crate) children: Vec<Element>,
    pub(crate) start: usize,
    pub(crate) end: usize,
}

impl Element {
    pub(crate) fn child(&self, name: &str) -> Option<&Self> { self.children.iter().find(|child| child.name == name) }
    pub(crate) fn attribute(&self, name: &str) -> Option<&str> { self.attributes.get(name).map(String::as_str) }
    fn integer(&self, name: &str) -> Option<i64> { self.attribute(name)?.parse().ok() }
    fn on_off(&self) -> bool { !matches!(self.attribute("val"), Some("0" | "false" | "off")) }
}

pub(crate) fn parse_xml(xml: &str) -> Result<Element, String> {
    if xml.len() as u64 > XML_LIMIT { return Err("Office XML exceeds the 16 MiB limit.".into()); }
    let mut reader = Reader::from_str(xml);
    let mut stack: Vec<Element> = Vec::new();
    let mut root = None;
    let mut elements = 0;
    loop {
        let position = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(|error| error.to_string())?;
        let empty = matches!(&event, Event::Empty(_));
        match event {
            Event::Start(tag) | Event::Empty(tag) => {
                elements += 1;
                if elements > XML_ELEMENT_LIMIT || stack.len() >= XML_DEPTH_LIMIT {
                    return Err("Office XML exceeds the element or nesting limit.".into());
                }
                let mut node = Element { name: String::from_utf8_lossy(tag.local_name().as_ref()).into_owned(), qualified_name: String::from_utf8_lossy(tag.name().as_ref()).into_owned(), start: position, end: reader.buffer_position() as usize, ..Default::default() };
                for attribute in tag.attributes() {
                    let attribute = attribute.map_err(|error| error.to_string())?;
                    let value = attribute.decoded_and_normalized_value_with(XmlVersion::Implicit1_0, reader.decoder(), 1, quick_xml::escape::resolve_xml_entity).map_err(|error| error.to_string())?;
                    let key = String::from_utf8_lossy(attribute.key.as_ref());
                    if key == "xmlns" { node.namespaces.insert(String::new(), value.to_string()); }
                    else if let Some(prefix) = key.strip_prefix("xmlns:") { node.namespaces.insert(prefix.to_owned(), value.to_string()); }
                    node.attributes.insert(String::from_utf8_lossy(attribute.key.local_name().as_ref()).into_owned(), value.into_owned());
                }
                if empty {
                    if let Some(parent) = stack.last_mut() { parent.children.push(node); }
                    else if root.replace(node).is_some() { return Err("XML has more than one root element.".into()); }
                } else { stack.push(node); }
            }
            Event::End(_) => {
                let mut node = stack.pop().ok_or("XML has an unmatched closing element.")?;
                node.end = reader.buffer_position() as usize;
                if let Some(parent) = stack.last_mut() { parent.children.push(node); }
                else if root.replace(node).is_some() { return Err("XML has more than one root element.".into()); }
            }
            Event::Text(value) => if let Some(node) = stack.last_mut() { node.text.push_str(&value.decode().map_err(|error| error.to_string())?); },
            Event::CData(value) => if let Some(node) = stack.last_mut() { node.text.push_str(&value.decode().map_err(|error| error.to_string())?); },
            Event::GeneralRef(value) => if let Some(node) = stack.last_mut() {
                if let Some(character) = value.resolve_char_ref().map_err(|error| error.to_string())? { node.text.push(character); }
                else {
                    let entity = value.decode().map_err(|error| error.to_string())?;
                    node.text.push_str(quick_xml::escape::resolve_xml_entity(&entity).ok_or("XML has an unsupported entity.")?);
                }
            },
            Event::Eof => {
                if !stack.is_empty() { return Err("XML has an unclosed element.".into()); }
                return root.ok_or_else(|| "XML has no root element.".into());
            }
            Event::DocType(_) => return Err("Office XML must not declare a document type.".into()),
            _ => {}
        }
    }
}

pub(crate) fn read_xml(archive: &mut ZipArchive<Cursor<&[u8]>>, part: &str) -> Result<Option<String>, String> {
    let entry = match archive.by_name(part.trim_start_matches('/')) {
        Ok(entry) => entry,
        Err(zip::result::ZipError::FileNotFound) => return Ok(None),
        Err(error) => return Err(error.to_string()),
    };
    if entry.size() > XML_LIMIT { return Err(format!("Office XML part {part} exceeds the 16 MiB limit.")); }
    let mut text = String::new();
    entry.take(XML_LIMIT + 1).read_to_string(&mut text).map_err(|error| error.to_string())?;
    if text.len() as u64 > XML_LIMIT { return Err(format!("Office XML part {part} exceeds the 16 MiB limit.")); }
    Ok(Some(text))
}

fn part_for_type(archive: &mut ZipArchive<Cursor<&[u8]>>, suffix: &str, fallback: &str) -> Result<String, String> {
    let Some(text) = read_xml(archive, "[Content_Types].xml")? else { return Ok(fallback.to_owned()); };
    let root = parse_xml(&text)?;
    Ok(root.children.iter().find(|node| node.name == "Override" && node.attribute("ContentType").is_some_and(|kind| kind.ends_with(suffix)))
        .and_then(|node| node.attribute("PartName")).map(|part| part.trim_start_matches('/').to_owned())
        .unwrap_or_else(|| fallback.to_owned()))
}

fn related_part(archive: &mut ZipArchive<Cursor<&[u8]>>, owner: &str, kind: &str) -> Result<Option<String>, String> {
    let (directory, file) = owner.rsplit_once('/').unwrap_or(("", owner));
    let relationships = if directory.is_empty() { format!("_rels/{file}.rels") } else { format!("{directory}/_rels/{file}.rels") };
    let Some(text) = read_xml(archive, &relationships)? else { return Ok(None); };
    let root = parse_xml(&text)?;
    let Some(relationship) = root.children.iter().find(|node| node.name == "Relationship" && node.attribute("Type").is_some_and(|value| value.rsplit('/').next() == Some(kind)) && node.attribute("TargetMode") != Some("External")) else { return Ok(None); };
    let target = relationship.attribute("Target").ok_or("Office source relationship has no target.")?;
    if target.contains(['\\', '\0', '?', '#']) || target.split('/').any(|component| component.contains(':')) { return Err("Office source relationship has an invalid package path.".into()); }
    let path = if target.starts_with('/') || directory.is_empty() { target.trim_start_matches('/').to_owned() } else { format!("{directory}/{target}") };
    let mut components = Vec::new();
    for component in path.split('/') {
        match component {
            "" | "." => {},
            ".." => { if components.pop().is_none() { return Err("Office source relationship escapes the package root.".into()); } },
            component => components.push(component),
        }
    }
    Ok(Some(components.join("/")))
}

pub(crate) fn read_office_source(archive: &mut ZipArchive<Cursor<&[u8]>>) -> (Option<crate::model::DocumentOfficeSourcePayload>, Vec<String>) {
    use crate::model::{DocumentOfficeSourcePayload, DocumentOfficeXmlPartPayload};
    let mut warnings = Vec::new();
    let main = match part_for_type(archive, ".document.main+xml", "word/document.xml") {
        Ok(part) => part,
        Err(error) => {
            warnings.push(format!("DOCX source style resources could not be located: {error}"));
            return (Some(DocumentOfficeSourcePayload { version: 1, styles: None, theme: None, numbering: None, resource_errors: warnings.clone(), origin: None, legacy_origins: Vec::new() }), warnings);
        },
    };
    let mut read = |kind: &str, suffix: &str, fallback: &str| {
        let read_part = (|| {
            let related = related_part(archive, &main, kind)?;
            let declared = if related.is_none() {
                read_xml(archive, "[Content_Types].xml")?.map(|xml| parse_xml(&xml)).transpose()?
                    .and_then(|root| root.children.into_iter().find(|node| node.name == "Override" && node.attribute("ContentType").is_some_and(|value| value.ends_with(suffix)))
                        .and_then(|node| node.attribute("PartName").map(|part| part.trim_start_matches('/').to_owned())))
            } else { None };
            let required = related.is_some() || declared.is_some();
            let part_name = related.or(declared).unwrap_or_else(|| fallback.to_owned());
            let xml = read_xml(archive, &part_name)?;
            if required && xml.is_none() { return Err(format!("Declared Office resource {part_name:?} is missing.")); }
            Ok(xml.map(|xml| DocumentOfficeXmlPartPayload { part_name, xml }))
        })();
        match read_part {
            Ok(part) => part,
            Err(error) => { warnings.push(format!("DOCX source {kind} resource could not be retained: {error}")); None },
        }
    };
    let styles = read("styles", ".styles+xml", "word/styles.xml");
    let theme = read("theme", ".theme+xml", "word/theme/theme1.xml");
    let numbering_part = read("numbering", ".numbering+xml", "word/numbering.xml");
    let numbering = numbering_part.map(|part| {
        let (resource, errors) = numbering_source::read(archive, part);
        warnings.extend(errors.into_iter().map(|error| format!("DOCX source numbering resource could not be retained: {error}")));
        resource
    });
    let source = Some(DocumentOfficeSourcePayload { version: 1, styles, theme, numbering, resource_errors: warnings.clone(), origin: None, legacy_origins: Vec::new() });
    (source, warnings)
}

fn picture_dimension(style: &str, property: &str) -> Option<f64> {
    for declaration in style.rsplit(';') {
        let Some((name, value)) = declaration.split_once(':') else { continue };
        if !name.trim().eq_ignore_ascii_case(property) { continue; }
        let value = value.trim().to_ascii_lowercase();
        for (unit, scale) in [("px", 1.0), ("pt", 4.0 / 3.0), ("in", 96.0), ("cm", 96.0 / 2.54), ("mm", 96.0 / 25.4)] {
            let Some(number) = value.strip_suffix(unit) else { continue };
            return number.trim().parse::<f64>().ok().map(|number| number * scale).filter(|number| number.is_finite() && *number > 0.0);
        }
        return None;
    }
    None
}

fn picture_dimensions(root: &Element) -> HashMap<i64, (Option<f64>, Option<f64>)> {
    root.children.iter().filter(|node| node.name == "numPicBullet").filter_map(|node| {
        let id = node.integer("numPicBulletId")?;
        let style = node.child("pict")?.child("shape")?.attribute("style")?;
        Some((id, (picture_dimension(style, "width"), picture_dimension(style, "height"))))
    }).collect()
}

fn patch_level(level: &mut NumberingLevelDefinition, xml: &Element, pictures: &HashMap<i64, (Option<f64>, Option<f64>)>, _warnings: &mut Vec<String>) {
    if let Some(start) = xml.child("start").and_then(|start| start.integer("val")).filter(|start| (0..=i64::from(i32::MAX)).contains(start)) {
        level.start = Some(start);
    }
    if let Some(pattern) = xml.child("numFmt").and_then(|format| format.attribute("format")) {
        let bytes = pattern.as_bytes();
        if !bytes.is_empty() && bytes.len() <= u8::MAX as usize && bytes.last() == Some(&b'1') && bytes[..bytes.len() - 1].iter().all(|byte| *byte == b'0') {
            level.format = Some("decimal".to_owned());
            level.minimum_digits = Some(bytes.len() as u8);
        }
    }
    if xml.child("rPr").and_then(|run| run.child("color")).and_then(|color| color.attribute("val")) == Some("auto") {
        level.bullet_color = Some("auto".to_owned());
        level.run_style.get_or_insert_with(Default::default).color = Some("auto".to_owned());
    }
    if let Some(picture) = &mut level.picture_bullet {
        if let Some((width, height)) = pictures.get(&picture.num_pic_bullet_id) {
            if width.is_some() { picture.width_px = *width; }
            if height.is_some() { picture.height_px = *height; }
        }
    }
}

fn patch_numbering(definitions: &mut NumberingDefinitionSet, root: &Element) -> Vec<String> {
    let mut warnings = Vec::new();
    let pictures = picture_dimensions(root);
    for node in &root.children {
        if node.name == "abstractNum" {
            let Some(abstract_id) = node.integer("abstractNumId") else { continue };
            let Some(definition) = definitions.abstracts.iter_mut().find(|value| value.abstract_num_id == abstract_id) else { continue };
            for xml in node.children.iter().filter(|child| child.name == "lvl") {
                if let Some(level) = definition.levels.iter_mut().find(|level| Some(level.ilvl) == xml.integer("ilvl")) { patch_level(level, xml, &pictures, &mut warnings); }
            }
        } else if node.name == "num" {
            let Some(num_id) = node.integer("numId") else { continue };
            let Some(instance) = definitions.instances.iter_mut().find(|value| value.num_id == num_id) else { continue };
            for override_xml in node.children.iter().filter(|child| child.name == "lvlOverride") {
                let Some(index) = override_xml.integer("ilvl").filter(|index| (0..=8).contains(index)) else { continue };
                if let Some(start) = override_xml.child("startOverride").and_then(|value| value.integer("val")).filter(|value| (0..=i64::from(i32::MAX)).contains(value)) {
                    instance.level_start_overrides.get_or_insert_with(HashMap::new).insert(index.to_string(), start);
                }
                if let Some(xml) = override_xml.child("lvl") {
                    if let Some(level) = instance.level_overrides.iter_mut().flatten().find(|level| level.ilvl == index) { patch_level(level, xml, &pictures, &mut warnings); }
                }
            }
        }
    }
    warnings
}

pub(crate) fn restore_numbering(source: &mut DocModel, archive: &mut ZipArchive<Cursor<&[u8]>>) -> Vec<String> {
    let Some(definitions) = &mut source.metadata.numbering_definitions else { return Vec::new() };
    let result = part_for_type(archive, ".numbering+xml", "word/numbering.xml")
        .and_then(|part| read_xml(archive, &part)).and_then(|text| text.map(|text| parse_xml(&text)).transpose());
    match result {
        Ok(Some(root)) => patch_numbering(definitions, &root),
        Ok(None) => Vec::new(),
        Err(error) => vec![format!("DOCX numbering properties could not be restored: {error}")],
    }
}

#[derive(Default)]
pub(crate) struct ParagraphWordWrap {
    defaults: Option<bool>,
    default_style: Option<String>,
    styles: HashMap<String, (Option<bool>, Option<String>)>,
}

impl ParagraphWordWrap {
    fn resolve(&self, style: &str) -> Option<bool> {
        let mut current = if style.is_empty() { self.default_style.as_deref().unwrap_or("Normal") } else { style };
        for _ in 0..self.styles.len() + 1 {
            let Some((value, parent)) = self.styles.get(current) else { return self.defaults };
            if value.is_some() { return *value; }
            let Some(parent) = parent else { return self.defaults };
            current = parent;
        }
        self.defaults
    }
}

pub(crate) fn read_word_wrap(archive: &mut ZipArchive<Cursor<&[u8]>>) -> Result<ParagraphWordWrap, String> {
    let part = part_for_type(archive, ".styles+xml", "word/styles.xml")?;
    let Some(text) = read_xml(archive, &part)? else { return Ok(ParagraphWordWrap::default()); };
    read_word_wrap_xml(&text)
}

pub(crate) fn read_word_wrap_xml(text: &str) -> Result<ParagraphWordWrap, String> {
    let root = parse_xml(text)?;
    let mut result = ParagraphWordWrap {
        defaults: root.child("docDefaults").and_then(|node| node.child("pPrDefault")).and_then(|node| node.child("pPr")).and_then(|node| node.child("wordWrap")).map(Element::on_off),
        ..Default::default()
    };
    for style in root.children.iter().filter(|node| node.name == "style" && node.attribute("type") == Some("paragraph")) {
        let Some(id) = style.attribute("styleId") else { continue };
        let value = style.child("pPr").and_then(|node| node.child("wordWrap")).map(Element::on_off);
        let parent = style.child("basedOn").and_then(|node| node.attribute("val")).map(str::to_owned);
        if matches!(style.attribute("default"), Some("1" | "true" | "on")) { result.default_style = Some(id.to_owned()); }
        result.styles.insert(id.to_owned(), (value, parent));
    }
    Ok(result)
}

#[derive(Default)]
struct ParagraphPatch {
    word_wrap: Option<bool>,
    text: String,
    automatic_ranges: Vec<(usize, usize)>,
    no_shading_ranges: Vec<(usize, usize)>,
}

fn paragraph_patch(root: &Element) -> ParagraphPatch {
    let mut result = ParagraphPatch { word_wrap: root.child("pPr").and_then(|node| node.child("wordWrap")).map(Element::on_off), ..Default::default() };
    fn visit(node: &Element, patch: &mut ParagraphPatch) {
        if node.name == "pPr" || node.name == "del" { return; }
        if node.name == "r" {
            let start = patch.text.len();
            fn append(node: &Element, text: &mut String) {
                match node.name.as_str() {
                    "rPr" | "instrText" | "delText" => {}
                    "t" => text.push_str(&node.text),
                    "tab" => text.push('\t'),
                    "br" | "cr" => text.push('\n'),
                    "noBreakHyphen" => text.push('\u{2011}'),
                    "softHyphen" => text.push('\u{ad}'),
                    _ => for child in &node.children { append(child, text); },
                }
            }
            append(node, &mut patch.text);
            if node.child("rPr").and_then(|run| run.child("color")).and_then(|color| color.attribute("val")) == Some("auto") && patch.text.len() > start {
                patch.automatic_ranges.push((start, patch.text.len()));
            }
            if node.child("rPr").and_then(|run| run.child("shd")).and_then(|shading| shading.attribute("val")) == Some("nil") && patch.text.len() > start {
                patch.no_shading_ranges.push((start, patch.text.len()));
            }
        } else { for child in &node.children { visit(child, patch); } }
    }
    visit(root, &mut result);
    result
}

fn object_field<'a>(value: &'a mut Value, name: &str) -> &'a mut Value {
    if !value.is_object() { *value = json!({}); }
    let field = value.as_object_mut().expect("object was initialized").entry(name).or_insert_with(|| json!({}));
    if !field.is_object() { *field = json!({}); }
    field
}

fn patch_paragraph(block: &mut DocumentBlock, patch: &ParagraphPatch, inherited_wrap: Option<bool>) -> bool {
    if patch.word_wrap.is_none() && inherited_wrap.is_none() && patch.automatic_ranges.is_empty() && patch.no_shading_ranges.is_empty() { return true; }
    let mut details = block.details.to_json_value();
    if let Some(wrap) = patch.word_wrap.or(inherited_wrap) { object_field(&mut details, "style")["wordWrap"] = json!(wrap); }
    let mut restored = true;
    if !patch.automatic_ranges.is_empty() || !patch.no_shading_ranges.is_empty() {
        if let Some(inlines) = details["inlines"].as_array_mut() {
            let visible = inlines.iter().filter_map(|inline| inline["text"].as_str()).collect::<String>();
            if visible == patch.text {
                let mut updated = Vec::new();
                let mut byte_offset = 0;
                for inline in inlines.iter() {
                    let Some(text) = inline["text"].as_str() else { updated.push(inline.clone()); continue };
                    let end = byte_offset + text.len();
                    let mut boundaries = vec![byte_offset, end];
                    for (start, finish) in patch.automatic_ranges.iter().chain(&patch.no_shading_ranges) {
                        if *start > byte_offset && *start < end { boundaries.push(*start); }
                        if *finish > byte_offset && *finish < end { boundaries.push(*finish); }
                    }
                    boundaries.sort_unstable();
                    boundaries.dedup();
                    let initial = inline["start"].as_u64().unwrap_or(0);
                    for pair in boundaries.windows(2) {
                        let mut value = inline.clone();
                        let start = pair[0] - byte_offset;
                        let finish = pair[1] - byte_offset;
                        value["text"] = json!(&text[start..finish]);
                        value["start"] = json!(initial + text[..start].encode_utf16().count() as u64);
                        value["end"] = json!(initial + text[..finish].encode_utf16().count() as u64);
                        if patch.automatic_ranges.iter().any(|(start, finish)| pair[0] >= *start && pair[1] <= *finish) { object_field(&mut value, "style")["color"] = json!("auto"); }
                        if patch.no_shading_ranges.iter().any(|(start, finish)| pair[0] >= *start && pair[1] <= *finish) {
                            // Duke's inherited background cannot represent an
                            // explicit w:shd nil. Remove it only on this range.
                            object_field(&mut value, "style").as_object_mut().expect("style was initialized").remove("backgroundColor");
                        }
                        updated.push(value);
                    }
                    if text.is_empty() { updated.push(inline.clone()); }
                    byte_offset = end;
                }
                *inlines = updated;
            } else { restored = false; }
        } else { restored = false; }
    }
    block.details = BlockDetails::Raw(details);
    restored
}

pub(crate) fn restore_paragraphs(document: &mut StructuredDocument, source: &DocModel, wrap: &ParagraphWordWrap) {
    let mut paragraphs = HashMap::<String, &ParagraphNode>::new();
    fn table<'a>(value: &'a TableNode, path: &str, output: &mut HashMap<String, &'a ParagraphNode>) {
        let root = value.block_id.as_deref().unwrap_or(path);
        for (row_index, row) in value.rows.iter().enumerate() {
            for (cell_index, cell) in row.cells.iter().enumerate() {
                for (node_index, node) in cell.nodes.iter().enumerate() {
                    let path = format!("{root}/row/{row_index}/cell/{cell_index}/{node_index}");
                    match node {
                        TableCellContentNode::Paragraph(value) => { output.insert(value.block_id.clone().unwrap_or(path), value); }
                        TableCellContentNode::Table(value) => table(value, &path, output),
                    }
                }
            }
        }
    }
    for (index, node) in source.nodes.iter().enumerate() {
        let path = format!("document/{index}");
        match node {
            DocNode::Paragraph(value) => { paragraphs.insert(value.block_id.clone().unwrap_or(path), value); }
            DocNode::Table(value) => table(value, &path, &mut paragraphs),
        }
    }
    fn visit(blocks: &mut [DocumentBlock], paragraphs: &HashMap<String, &ParagraphNode>, wrap: &ParagraphWordWrap, warnings: &mut Vec<String>) {
        for block in blocks {
            let paragraph = matches!(&block.details, BlockDetails::Paragraph(_))
                || matches!(&block.details, BlockDetails::Raw(details) if details["kind"].as_str() == Some("paragraph"));
            let source = paragraph.then(|| block.source_refs.iter().filter(|source| source.story == "body").find_map(|source| {
                paragraphs.get(&source.node_id).copied().or_else(|| source.node_id.rfind("/body/").and_then(|index| paragraphs.get(&source.node_id[index..]).copied()))
            })).flatten();
            if let Some(source) = source {
                if let Some(xml) = &source.source_xml {
                    match parse_xml(xml) {
                        Ok(root) => {
                            let style = source.style.as_ref().and_then(|style| style.style_id.as_deref()).unwrap_or("");
                            if !patch_paragraph(block, &paragraph_patch(&root), wrap.resolve(style)) { warnings.push(format!("DOCX automatic text color or explicit shading could not be mapped safely for paragraph {}.", block.id)); }
                        }
                        Err(error) => warnings.push(format!("DOCX paragraph properties could not be restored for {}: {error}", block.id)),
                    }
                }
            }
            visit(&mut block.children, paragraphs, wrap, warnings);
        }
    }
    visit(&mut document.blocks, &paragraphs, wrap, &mut document.warnings);
}

#[cfg(test)]
mod tests {
    use std::io::Write;
    use zip::{ZipWriter, write::SimpleFileOptions};
    use super::*;

    fn archive_bytes(parts: &[(&str, &str)]) -> Vec<u8> {
        let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
        for (name, text) in parts {
            writer.start_file(*name, SimpleFileOptions::default()).unwrap();
            writer.write_all(text.as_bytes()).unwrap();
        }
        writer.finish().unwrap().into_inner()
    }

    fn source_model(nodes: Value, numbering: Value) -> DocModel {
        serde_json::from_value(json!({
            "nodes": nodes,
            "metadata": { "sourceParts": 1, "warnings": [], "headerSections": [], "footerSections": [], "paragraphStyles": [], "numberingDefinitions": numbering }
        })).unwrap()
    }

    fn paragraph(text: &str, inlines: Value) -> DocumentBlock {
        serde_json::from_value(json!({
            "object": "block", "id": "paragraph", "type": "text", "content": text,
            "details": { "kind": "paragraph", "inlines": inlines, "retainedProperty": "keep" },
            "sourceRefs": [], "children": []
        })).unwrap()
    }

    #[test]
    fn numbering_restoration_preserves_zero_start_and_automatic_marker_color() {
        let mut source = source_model(json!([]), json!({
            "abstracts": [{ "abstractNumId": 7, "levels": [{ "ilvl": 0, "start": 1, "format": "decimal", "text": "%1." }] }],
            "instances": [{ "numId": 9, "abstractNumId": 7, "levelOverrides": [{ "ilvl": 0, "text": "[%1]" }] }]
        }));
        let bytes = archive_bytes(&[
            ("[Content_Types].xml", r#"<Types><Override PartName="/custom/list.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>"#),
            ("custom/list.xml", r#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="7"><w:lvl w:ilvl="0"><w:rPr><w:color w:val="auto"/></w:rPr></w:lvl></w:abstractNum><w:num w:numId="9"><w:lvlOverride w:ilvl="0"><w:startOverride w:val="0"/><w:lvl w:ilvl="0"><w:rPr><w:color w:val="auto"/></w:rPr></w:lvl></w:lvlOverride></w:num></w:numbering>"#),
        ]);
        let mut archive = ZipArchive::new(Cursor::new(bytes.as_slice())).unwrap();
        assert!(restore_numbering(&mut source, &mut archive).is_empty());
        let definitions = source.metadata.numbering_definitions.unwrap();
        assert_eq!(definitions.instances[0].level_start_overrides.as_ref().unwrap()["0"], 0);
        assert_eq!(definitions.abstracts[0].levels[0].bullet_color.as_deref(), Some("auto"));
        assert_eq!(definitions.abstracts[0].levels[0].run_style.as_ref().unwrap().color.as_deref(), Some("auto"));
        assert_eq!(definitions.instances[0].level_overrides.as_ref().unwrap()[0].bullet_color.as_deref(), Some("auto"));
    }

    #[test]
    fn numbering_restoration_reads_padded_decimal_tokens_and_reports_unmapped_custom_formats() {
        let mut source = source_model(json!([]), json!({
            "abstracts": [{ "abstractNumId": 7, "levels": [{ "ilvl": 0, "format": "custom" }, { "ilvl": 1, "format": "custom" }] }],
            "instances": [{ "numId": 9, "abstractNumId": 7, "levelOverrides": [{ "ilvl": 0 }] }]
        }));
        let bytes = archive_bytes(&[("word/numbering.xml", r#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="7"><w:lvl w:ilvl="0"><w:start w:val="0"/><w:numFmt w:val="custom" w:format="001"/></w:lvl><w:lvl w:ilvl="1"><w:numFmt w:val="custom" w:format="ア"/></w:lvl></w:abstractNum><w:num w:numId="9"><w:lvlOverride w:ilvl="0"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal" w:format="0001"/></w:lvl></w:lvlOverride></w:num></w:numbering>"#)]);
        let mut archive = ZipArchive::new(Cursor::new(bytes.as_slice())).unwrap();
        let warnings = restore_numbering(&mut source, &mut archive);
        assert_eq!(warnings.len(), 1);
        assert!(warnings[0].contains("ア"));
        let definitions = source.metadata.numbering_definitions.unwrap();
        assert_eq!(definitions.abstracts[0].levels[0].start, Some(0));
        assert_eq!(definitions.abstracts[0].levels[0].format.as_deref(), Some("decimal"));
        assert_eq!(definitions.abstracts[0].levels[0].minimum_digits, Some(3));
        assert_eq!(definitions.instances[0].level_overrides.as_ref().unwrap()[0].minimum_digits, Some(4));
    }

    #[test]
    fn picture_bullet_dimensions_preserve_fractional_pixels_and_convert_css_units() {
        let mut source = source_model(json!([]), json!({
            "abstracts": [{ "abstractNumId": 7, "levels": [{ "ilvl": 0, "format": "bullet", "pictureBulletId": 5,
                "pictureBullet": { "numPicBulletId": 5, "src": "data:image/png;base64,AA==", "widthPx": 1, "heightPx": 1 } }] }],
            "instances": []
        }));
        let bytes = archive_bytes(&[("word/numbering.xml", r##"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml"><w:numPicBullet w:numPicBulletId="5"><w:pict><v:shapetype id="_x0000_t75"/><v:shape type="#_x0000_t75" style="width:2pt;width:9.5px;height:7.125PT"/></w:pict></w:numPicBullet><w:abstractNum w:abstractNumId="7"><w:lvl w:ilvl="0"><w:lvlPicBulletId w:val="5"/></w:lvl></w:abstractNum></w:numbering>"##)]);
        let mut archive = ZipArchive::new(Cursor::new(bytes.as_slice())).unwrap();
        assert!(restore_numbering(&mut source, &mut archive).is_empty());
        let definitions = source.metadata.numbering_definitions.unwrap();
        let picture = definitions.abstracts[0].levels[0].picture_bullet.as_ref().unwrap();
        assert_eq!(picture.width_px, Some(9.5));
        assert_eq!(picture.height_px, Some(9.5));
        assert_eq!(picture_dimension("width:1in", "width"), Some(96.0));
        assert_eq!(picture_dimension("width:auto", "width"), None);
        assert_eq!(picture_dimension("width:-2px", "width"), None);
    }

    #[test]
    fn style_word_wrap_resolves_defaults_inheritance_direct_values_and_cycles() {
        let bytes = archive_bytes(&[
            ("[Content_Types].xml", r#"<Types><Override PartName="/custom/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>"#),
            ("custom/styles.xml", r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:pPrDefault><w:pPr><w:wordWrap w:val="0"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Default"><w:basedOn w:val="Base"/></w:style><w:style w:type="paragraph" w:styleId="Base"><w:pPr><w:wordWrap/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Child"><w:basedOn w:val="Base"/></w:style><w:style w:type="paragraph" w:styleId="Direct"><w:basedOn w:val="Base"/><w:pPr><w:wordWrap w:val="false"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="CycleA"><w:basedOn w:val="CycleB"/></w:style><w:style w:type="paragraph" w:styleId="CycleB"><w:basedOn w:val="CycleA"/></w:style></w:styles>"#),
        ]);
        let mut archive = ZipArchive::new(Cursor::new(bytes.as_slice())).unwrap();
        let wrap = read_word_wrap(&mut archive).unwrap();
        assert_eq!(wrap.resolve(""), Some(true));
        assert_eq!(wrap.resolve("Child"), Some(true));
        assert_eq!(wrap.resolve("Direct"), Some(false));
        assert_eq!(wrap.resolve("Missing"), Some(false));
        assert_eq!(wrap.resolve("CycleA"), Some(false));
    }

    #[test]
    fn automatic_color_splits_unicode_inlines_with_utf16_offsets_and_keeps_other_fields() {
        let root = parse_xml(r#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:rPr><w:color w:val="auto"/></w:rPr><w:t>A&amp;😀</w:t></w:r><w:r><w:t>B</w:t></w:r></w:p>"#).unwrap();
        let patch = paragraph_patch(&root);
        let mut block = paragraph("A&😀B", json!([{ "type": "text", "text": "A&😀B", "start": 0, "end": 5, "style": { "bold": true, "color": "#112233" }, "link": "https://example.test", "retainedInline": 7 }]));
        assert!(patch_paragraph(&mut block, &patch, None));
        let details = block.details.to_json_value();
        let inlines = details["inlines"].as_array().unwrap();
        assert_eq!(inlines.len(), 2);
        assert_eq!(inlines[0]["text"], "A&😀");
        assert_eq!(inlines[0]["start"], 0);
        assert_eq!(inlines[0]["end"], 4);
        assert_eq!(inlines[0]["style"]["color"], "auto");
        assert_eq!(inlines[1]["text"], "B");
        assert_eq!(inlines[1]["start"], 4);
        assert_eq!(inlines[1]["end"], 5);
        assert_eq!(inlines[1]["style"]["color"], "#112233");
        for inline in inlines {
            assert_eq!(inline["style"]["bold"], true);
            assert_eq!(inline["link"], "https://example.test");
            assert_eq!(inline["retainedInline"], 7);
        }
        assert_eq!(details["retainedProperty"], "keep");
        assert_eq!(block.content, "A&😀B");
    }

    #[test]
    fn explicit_nil_shading_splits_unicode_and_preserves_unselected_background_and_empty_entities() {
        let root = parse_xml(r#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>A</w:t></w:r><w:r><w:rPr><w:highlight w:val="none"/><w:shd w:val="nil"/></w:rPr><w:t>😀</w:t></w:r><w:r><w:t>B</w:t></w:r></w:p>"#).unwrap();
        let mut block = paragraph("A😀B", json!([
            { "type": "text", "text": "A😀B", "start": 0, "end": 4, "style": { "backgroundColor": "#ABCDEF", "highlight": "none", "bold": true }, "link": "https://example.test" },
            { "type": "noteReference", "text": "", "start": 2, "end": 2, "noteReference": { "id": "note-1" } },
        ]));
        assert!(patch_paragraph(&mut block, &paragraph_patch(&root), None));
        let details = block.details.to_json_value();
        let inlines = details["inlines"].as_array().unwrap();
        assert_eq!(inlines.len(), 4);
        assert_eq!(inlines[0]["style"]["backgroundColor"], "#ABCDEF");
        assert_eq!(inlines[1]["text"], "😀");
        assert_eq!((inlines[1]["start"].as_u64(), inlines[1]["end"].as_u64()), (Some(1), Some(3)));
        assert!(inlines[1]["style"].get("backgroundColor").is_none());
        assert_eq!(inlines[2]["style"]["backgroundColor"], "#ABCDEF");
        for inline in &inlines[..3] { assert_eq!(inline["style"]["bold"], true); assert_eq!(inline["link"], "https://example.test"); }
        assert_eq!(inlines[3]["noteReference"]["id"], "note-1");
        assert_eq!(details["retainedProperty"], "keep");
    }

    #[test]
    fn paragraph_mark_color_does_not_change_explicit_body_color() {
        let root = parse_xml(r#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pPr><w:rPr><w:color w:val="auto"/></w:rPr></w:pPr><w:r><w:rPr><w:color w:val="FF0000"/></w:rPr><w:t>Body</w:t></w:r></w:p>"#).unwrap();
        let patch = paragraph_patch(&root);
        assert!(patch.automatic_ranges.is_empty());
        let mut block = paragraph("Body", json!([{ "type": "text", "text": "Body", "start": 0, "end": 4, "style": { "color": "#FF0000" } }]));
        let before = block.clone();
        assert!(patch_paragraph(&mut block, &patch, None));
        assert_eq!(block, before);
    }

    #[test]
    fn mismatched_visible_text_keeps_color_and_reports_the_mapping_failure() {
        let root = parse_xml(r#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pPr><w:wordWrap/></w:pPr><w:r><w:rPr><w:color w:val="auto"/></w:rPr><w:t>Changed</w:t></w:r></w:p>"#).unwrap();
        let mut block = paragraph("Existing", json!([{ "type": "text", "text": "Existing", "start": 0, "end": 8, "style": { "color": "#112233" } }]));
        let before = block.details.to_json_value()["inlines"].clone();
        assert!(!patch_paragraph(&mut block, &paragraph_patch(&root), None));
        let details = block.details.to_json_value();
        assert_eq!(details["inlines"], before);
        assert_eq!(details["style"]["wordWrap"], true);
    }

    #[test]
    fn paragraph_restoration_uses_node_ids_and_does_not_patch_child_objects() {
        let source = source_model(json!([
            { "type": "paragraph", "children": [{ "type": "text", "text": "Same" }], "sourceXml": "<w:p><w:pPr><w:wordWrap/></w:pPr><w:r><w:t>Same</w:t></w:r></w:p>" },
            { "type": "paragraph", "children": [{ "type": "text", "text": "Same" }], "sourceXml": "<w:p><w:pPr><w:wordWrap w:val=\"0\"/></w:pPr><w:r><w:t>Same</w:t></w:r></w:p>" }
        ]), Value::Null);
        let mut document = duke_pages::extract_document_blocks(&source);
        let mut child = paragraph("Object", json!([]));
        child.details = BlockDetails::Raw(json!({ "kind": "image", "retained": "object" }));
        child.source_refs = document.blocks[0].source_refs.clone();
        let before = child.clone();
        document.blocks[0].children.push(child);
        restore_paragraphs(&mut document, &source, &ParagraphWordWrap::default());
        assert!(document.warnings.is_empty());
        assert_eq!(document.blocks[0].details.to_json_value()["style"]["wordWrap"], true);
        assert_eq!(document.blocks[1].details.to_json_value()["style"]["wordWrap"], false);
        assert_eq!(document.blocks[0].children[0], before);
    }

    #[test]
    fn malformed_xml_and_excessive_nesting_are_rejected() {
        assert!(parse_xml("<a><b></a>").is_err());
        assert!(parse_xml("<!DOCTYPE a [<!ENTITY x 'text'>]><a>&x;</a>").is_err());
        let nested = "<a>".repeat(XML_DEPTH_LIMIT + 1) + &"</a>".repeat(XML_DEPTH_LIMIT + 1);
        assert!(parse_xml(&nested).is_err());
        let bytes = archive_bytes(&[("word/styles.xml", "<w:styles>")]);
        let mut archive = ZipArchive::new(Cursor::new(bytes.as_slice())).unwrap();
        assert!(read_word_wrap(&mut archive).is_err());
    }
}
