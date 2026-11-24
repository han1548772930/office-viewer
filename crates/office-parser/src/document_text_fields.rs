//! Restore Word field delimiters without making cached text the field definition.
//! Positions are the same half-open UTF-16 offsets as the extracted paragraphs.
use std::{collections::{BTreeMap, HashMap}, io::Cursor};

use duke_pages::{BlockDetails, DocModel, DocumentBlock, ParagraphChildNode, ParagraphNode, StructuredDocument};
use serde_json::{Value, json};
use zip::ZipArchive;

use crate::document_format_sidecar::{Element, parse_xml, read_xml};

#[derive(Clone, Copy, Debug)]
struct Position { paragraph: usize, offset: usize, note_reference_count: usize }

struct Paragraph<'a> { source: &'a ParagraphNode, source_id: String, block_id: String, text: String }
struct RunSpan { start: usize, end: usize, format: Value }
struct Projection { text: String, offset: usize, spans: Vec<RunSpan>, note_reference_count: usize }

struct Pending {
    id: String,
    begin_sequence: usize,
    instruction: String,
    instruction_xml: String,
    begin: Position,
    result: Option<Position>,
    locked: bool,
    nested: bool,
    in_instruction: bool,
    language: Option<String>,
    parent_id: Option<String>,
    namespaces: BTreeMap<String, String>,
    has_instruction_text: bool,
}

struct Field { pending: Pending, end: Position }

#[derive(Default)]
pub(crate) struct FieldRestoration {
    pub warnings: Vec<String>,
    /// A whole-body provenance record is necessary if an extracted block is
    /// absent. Retain it once in the payload, including an entirely empty body.
    pub source: Option<Value>,
}

#[derive(Default)]
struct Reader {
    pending: Vec<Pending>,
    fields: Vec<Field>,
    next_id: usize,
    tagged_identity: Option<(usize, String)>,
    identities: HashMap<String, String>,
    id_prefix: String,
}

fn namespaces(node: &Element, inherited: &BTreeMap<String, String>) -> BTreeMap<String, String> {
    let mut scope = inherited.clone();
    scope.extend(node.namespaces.iter().map(|(key, value)| (key.clone(), value.clone())));
    scope
}

pub(crate) fn read_identities(archive: &mut ZipArchive<Cursor<&[u8]>>) -> Result<HashMap<String, String>, String> {
    let Some(xml) = read_xml(archive, "customXml/office-fields.xml")? else { return Ok(HashMap::new()); };
    let root = parse_xml(&xml)?;
    let namespace = |node: &Element, inherited: &HashMap<String, String>| {
        let prefix = node.qualified_name.split_once(':').map_or("", |(prefix, _)| prefix);
        node.namespaces.get(prefix).or_else(|| inherited.get(prefix)).cloned()
    };
    if root.name != "fields" || root.attribute("version") != Some("1")
        || namespace(&root, &HashMap::new()).as_deref() != Some("urn:office-studio:document-fields:1") {
        return Err("Native field identity mapping has an invalid version, namespace or root.".into());
    }
    let mut identities = HashMap::new(); let mut ids = std::collections::HashSet::new();
    for node in &root.children {
        if node.name != "field" || namespace(node, &root.namespaces).as_deref() != Some("urn:office-studio:document-fields:1") {
            return Err("Native field identity mapping has an unknown entry or namespace.".into());
        }
        let key = node.attribute("key").filter(|key| key.strip_prefix("office_field_v1_").is_some_and(|suffix| suffix.len() == 8 && suffix.bytes().all(|byte| byte.is_ascii_hexdigit()))).ok_or("Native field bookmark key is invalid.")?;
        let id = node.attribute("id").filter(|id| !id.trim().is_empty() && id.len() <= 4096).ok_or("Native field identity is invalid.")?;
        if identities.insert(key.to_owned(), id.to_owned()).is_some() || !ids.insert(id.to_owned()) { return Err("Native field identity mapping has duplicate keys or identities.".into()); }
    }
    Ok(identities)
}

impl Reader {
    fn begin(&mut self, at: Position, locked: bool, simple: Option<&str>) {
        self.next_id += 1;
        for ancestor in &mut self.pending { ancestor.nested = true; }
        let id = self.tagged_identity.take().filter(|(offset, _)| *offset == at.offset).map(|(_, id)| id)
            .unwrap_or_else(|| format!("{}{}", if self.id_prefix.is_empty() { "word-field-" } else { &self.id_prefix }, self.next_id));
        self.pending.push(Pending { id, begin_sequence: self.next_id, instruction: simple.unwrap_or_default().to_owned(), instruction_xml: String::new(),
            begin: at, result: simple.map(|_| at), locked, nested: false,
            in_instruction: self.pending.iter().any(|field| field.result.is_none()), language: None,
            parent_id: self.pending.last().map(|field| field.id.clone()), namespaces: BTreeMap::new(), has_instruction_text: false });
    }

    fn code(&mut self, xml: &str, properties: &str, except_last: bool) {
        let count = self.pending.len().saturating_sub(usize::from(except_last));
        for field in self.pending.iter_mut().take(count).filter(|field| field.result.is_none()) {
            field.instruction_xml.push_str("<w:r>");
            field.instruction_xml.push_str(properties);
            field.instruction_xml.push_str(xml);
            field.instruction_xml.push_str("</w:r>");
        }
    }

    fn end(&mut self, at: Position) -> Result<(), String> {
        let pending = self.pending.pop().ok_or("A native field end has no beginning.")?;
        if pending.instruction.trim().is_empty() { return Err("A native field has no instruction.".into()); }
        self.fields.push(Field { pending, end: at });
        Ok(())
    }

    fn visit(&mut self, node: &Element, xml: &str, paragraph: usize, projection: &mut Projection,
        inherited: &BTreeMap<String, String>, paragraph_style: Option<&str>, properties: &str) -> Result<(), String> {
        let scope = namespaces(node, inherited);
        for field in &mut self.pending { field.namespaces.extend(scope.iter().map(|(key, value)| (key.clone(), value.clone()))); }
        let at = Position { paragraph, offset: projection.offset, note_reference_count: projection.note_reference_count };
        match node.name.as_str() {
            "pPr" | "rPr" | "sdtPr" | "sdtEndPr" | "del" | "moveFrom" | "drawing" | "pict" | "object" | "txbxContent" | "oMath" | "oMathPara" => return Ok(()),
            "bookmarkStart" => {
                if let Some(id) = node.attribute("name").and_then(|name| self.identities.get(name)) { self.tagged_identity = Some((at.offset, id.clone())); }
                return Ok(());
            }
            "r" => {
                let properties = node.child("rPr").map(|property| &xml[property.start..property.end]).unwrap_or("");
                let mut run_scope = scope.clone();
                if let Some(property) = node.child("rPr") { run_scope = namespaces(property, &run_scope); }
                if node.children.iter().all(|child| matches!(child.name.as_str(), "rPr" | "lastRenderedPageBreak")) {
                    let mut format = json!({ "version": 1, "namespaces": run_scope });
                    if !properties.is_empty() { format["propertiesXml"] = json!(properties); }
                    if let Some(style) = paragraph_style { format["paragraphStyleId"] = json!(style); }
                    projection.spans.push(RunSpan { start: at.offset, end: at.offset, format });
                }
                for child in &node.children { self.visit(child, xml, paragraph, projection, &run_scope, paragraph_style, properties)?; }
                return Ok(());
            }
            "fldSimple" => {
                let instruction = node.attribute("instr").ok_or("A simple native field has no instruction.")?;
                if self.pending.iter().any(|field| field.result.is_none()) {
                    self.code(&xml[node.start..node.end], properties, false);
                    for ancestor in &mut self.pending { ancestor.nested = true; }
                    // Duke projects text within an instruction too. Preserve it and
                    // explicitly mark this case until an offset-remapping adapter exists.
                    for child in &node.children { self.visit(child, xml, paragraph, projection, &scope, paragraph_style, properties)?; }
                    return Ok(());
                }
                self.begin(at, node.attribute("fldLock").is_some_and(on_off), Some(instruction));
                for child in &node.children { self.visit(child, xml, paragraph, projection, &scope, paragraph_style, properties)?; }
                self.end(Position { paragraph, offset: projection.offset, note_reference_count: projection.note_reference_count })?;
                return Ok(());
            }
            "fldChar" => {
                match node.attribute("fldCharType") {
                    Some("begin") => {
                        self.code(&xml[node.start..node.end], properties, false);
                        self.begin(at, node.attribute("fldLock").is_some_and(on_off), None);
                    }
                    Some("separate") => {
                        self.code(&xml[node.start..node.end], properties, true);
                        let field = self.pending.last_mut().ok_or("A native field separator has no beginning.")?;
                        if field.result.replace(at).is_some() { return Err("A native field has two result separators.".into()); }
                    }
                    Some("end") => {
                        self.code(&xml[node.start..node.end], properties, true);
                        self.end(at)?;
                    }
                    _ => return Err("A native field delimiter has an unknown type.".into()),
                }
                return Ok(());
            }
            "instrText" => {
                let field = self.pending.last_mut().filter(|field| field.result.is_none()).ok_or("A native instruction is outside its field code.")?;
                field.instruction.push_str(&node.text);
                self.code(&xml[node.start..node.end], properties, false);
                return Ok(());
            }
            "t" | "tab" | "br" | "cr" | "noBreakHyphen" | "softHyphen" | "sym" | "footnoteReference" | "endnoteReference" => {
                self.code(&xml[node.start..node.end], properties, false);
                if matches!(node.name.as_str(), "footnoteReference" | "endnoteReference") {
                    projection.note_reference_count = projection.note_reference_count.checked_add(1).ok_or("Native note reference count overflow")?;
                }
                let wrapped = format!("<w:r>{}</w:r>", &xml[node.start..node.end]);
                let value = duke_pages::parse::parse_run_text_tokens(&wrapped).into_iter().map(|token| token.text).collect::<String>();
                let end = at.offset + value.encode_utf16().count();
                if !value.is_empty() {
                    for field in self.pending.iter_mut().filter(|field| field.result.is_none()) { field.has_instruction_text = true; }
                    let language = if properties.is_empty() { None } else { parse_xml(properties)?.child("lang").and_then(|lang| lang.attribute("val").or_else(|| lang.attribute("eastAsia"))).map(str::to_owned) };
                    for field in self.pending.iter_mut().filter(|field| field.result.is_some() && field.language.is_none()) { field.language = language.clone(); }
                    let mut format = json!({ "version": 1, "namespaces": scope });
                    if !properties.is_empty() { format["propertiesXml"] = json!(properties); }
                    if let Some(style) = paragraph_style { format["paragraphStyleId"] = json!(style); }
                    projection.spans.push(RunSpan { start: at.offset, end, format });
                    projection.text.push_str(&value);
                    projection.offset = end;
                }
                return Ok(());
            }
            _ => {},
        }
        for child in &node.children { self.visit(child, xml, paragraph, projection, &scope, paragraph_style, properties)?; }
        Ok(())
    }
}

fn on_off(value: &str) -> bool { !matches!(value, "0" | "false" | "off") }

fn paragraph_block(block: &DocumentBlock) -> bool {
    matches!(&block.details, BlockDetails::Paragraph(_)) || matches!(&block.details, BlockDetails::Raw(raw) if raw["kind"].as_str() == Some("paragraph"))
}

fn collect_blocks(blocks: &[DocumentBlock], note_id: Option<i64>, output: &mut HashMap<crate::document_story_source::SourceAddress, (String, String)>) {
    for block in blocks {
        let note_id = block.note_id.or(note_id);
        if paragraph_block(block) {
            for source in &block.source_refs { output.insert(crate::document_story_source::SourceAddress::from_ref(source, note_id), (block.id.clone(), block.content.clone())); }
        }
        collect_blocks(&block.children, note_id, output);
    }
}

fn slice_utf16(text: &str, start: usize, end: usize) -> Option<&str> {
    let mut units = 0; let mut from = None; let mut to = None;
    for (byte, ch) in text.char_indices() {
        if units == start { from = Some(byte); }
        if units == end { to = Some(byte); }
        units += ch.len_utf16();
    }
    if units == start { from = Some(text.len()); }
    if units == end { to = Some(text.len()); }
    (start <= end).then_some(())?;
    text.get(from?..to?)
}

fn source_language(paragraph: &ParagraphNode, at: usize) -> Option<String> {
    let mut offset = 0;
    for child in &paragraph.children {
        if let ParagraphChildNode::Text(run) = child {
            let end = offset + run.text.encode_utf16().count();
            if offset <= at && at < end {
                return run.style.as_ref().and_then(|style| style.language.as_ref().or(style.language_east_asia.as_ref()).or(style.language_bidi.as_ref())).cloned();
            }
            offset = end;
        }
    }
    paragraph.paragraph_mark_style.as_ref().and_then(|style| style.language.as_ref().or(style.language_east_asia.as_ref()).or(style.language_bidi.as_ref())).cloned()
}

fn field_value(field: Field, paragraphs: &[Paragraph<'_>]) -> Result<(String, Value), String> {
    let mut pending = field.pending;
    let start = pending.result.unwrap_or(pending.begin);
    let first = &paragraphs[start.paragraph]; let last = &paragraphs[field.end.paragraph];
    if pending.language.is_none() { pending.language = source_language(first.source, start.offset); }
    let mut cached = String::new();
    for (index, paragraph) in paragraphs.iter().enumerate().take(field.end.paragraph + 1).skip(start.paragraph) {
        if index != start.paragraph { cached.push('\n'); }
        let from = if index == start.paragraph { start.offset } else { 0 };
        let end = if index == field.end.paragraph { field.end.offset } else { paragraph.text.encode_utf16().count() };
        cached.push_str(slice_utf16(&paragraph.text, from, end).ok_or("A native field result splits a UTF-16 surrogate pair.")?);
    }
    let instruction = pending.instruction.trim().to_owned();
    let name = instruction.split_whitespace().next().unwrap_or_default().to_ascii_uppercase();
    let mut value = json!({ "id": pending.id, "field": name, "instruction": instruction, "cachedResult": cached,
        "startOffset": start.offset, "endOffset": field.end.offset,
        "startNoteReferenceCount": start.note_reference_count, "endNoteReferenceCount": field.end.note_reference_count,
        "locked": pending.locked, "containsNestedFields": pending.nested,
        "metadata": { "source": { "format": "docx", "instruction": instruction, "instructionXml": pending.instruction_xml,
            "nestedInInstruction": pending.in_instruction, "parentFieldId": pending.parent_id, "beginSequence": pending.begin_sequence,
            "instructionNamespaces": pending.namespaces, "hasInstructionText": pending.has_instruction_text,
            "beginNodeId": paragraphs[pending.begin.paragraph].source_id, "beginOffset": pending.begin.offset } } });
    if start.paragraph != field.end.paragraph { value["endBlockId"] = json!(last.block_id); }
    if let Some(language) = pending.language { value["language"] = json!(language); }
    Ok((first.block_id.clone(), value))
}

fn inline_sources(details: &mut Value, spans: &[RunSpan]) -> Result<(), String> {
    let Some(inlines) = details.get("inlines").and_then(Value::as_array) else { return Ok(()); };
    let mut result = Vec::new();
    for inline in inlines {
        if !matches!(inline["type"].as_str(), Some("text" | "form_field")) { result.push(inline.clone()); continue; }
        let Some(start) = inline["start"].as_u64().and_then(|value| usize::try_from(value).ok()) else { result.push(inline.clone()); continue; };
        let Some(end) = inline["end"].as_u64().and_then(|value| usize::try_from(value).ok()) else { result.push(inline.clone()); continue; };
        let Some(text) = inline["text"].as_str() else { result.push(inline.clone()); continue; };
        if inline["type"].as_str() == Some("form_field") {
            let mut value = inline.clone();
            // Internal formula/code/lock SDTs remain one typed field. Attach a
            // proven whole-field run format without splitting its identity.
            if let Some(span) = spans.iter().find(|span| span.start <= start && end <= span.end) { value["sourceFormat"] = span.format.clone(); }
            result.push(value); continue;
        }
        let mut boundaries = vec![start, end];
        for span in spans { if span.start > start && span.start < end { boundaries.push(span.start); } if span.end > start && span.end < end { boundaries.push(span.end); } }
        boundaries.sort_unstable(); boundaries.dedup();
        if start == end {
            let mut value = inline.clone();
            if let Some(span) = spans.iter().find(|span| span.start <= start && span.end >= end) { value["sourceFormat"] = span.format.clone(); }
            result.push(value); continue;
        }
        for pair in boundaries.windows(2) {
            let (from, to) = (pair[0], pair[1]);
            let part = slice_utf16(text, from - start, to - start).ok_or("A native run source splits a UTF-16 surrogate pair.")?;
            let mut clipped = inline.clone(); clipped["text"] = json!(part); clipped["start"] = json!(from); clipped["end"] = json!(to);
            if from != start { if let Some(object) = clipped.as_object_mut() { object.remove("noteReference"); } }
            if let Some(span) = spans.iter().find(|span| span.start <= from && span.end >= to) { clipped["sourceFormat"] = span.format.clone(); }
            result.push(clipped);
        }
    }
    details["inlines"] = json!(result);
    Ok(())
}

/// Ordinary run provenance uses the same visible-text reader as native fields.
/// Cross-paragraph fields are handled later by the complete story plan.
pub(crate) fn restore_run_formats(details: &mut Value, xml: &str, text: &str, inherited: &BTreeMap<String, String>) -> Result<(), String> {
    let root = parse_xml(xml)?;
    let style = root.child("pPr").and_then(|properties| properties.child("pStyle")).and_then(|style| style.attribute("val"));
    let mut reader = Reader::default();
    let mut projection = Projection { text: String::new(), offset: 0, spans: Vec::new(), note_reference_count: 0 };
    reader.visit(&root, xml, 0, &mut projection, inherited, style, "")?;
    if projection.text != text { return Err("Source run formatting does not match the extracted visible text.".into()); }
    inline_sources(details, &projection.spans)
}

/// Prepare every paragraph before publishing any field ranges. A failed mapping
/// retains source XML and a diagnostic instead of inventing valid offsets.
pub(crate) fn restore(document: &mut StructuredDocument, source: &DocModel, identities: &HashMap<String, String>) -> FieldRestoration {
    let mut result = FieldRestoration::default();
    let mut retained = Vec::new();
    for story in crate::document_story_source::stories(source) {
        let mut restored = restore_story(document, &story, identities);
        result.warnings.append(&mut restored.warnings);
        if let Some(source) = restored.source { retained.push(source); }
    }
    if retained.len() == 1 { result.source = retained.pop(); }
    else if !retained.is_empty() { result.source = Some(json!({ "version": 1, "format": "docx", "stories": retained })); }
    result
}

fn restore_story(document: &mut StructuredDocument, story: &crate::document_story_source::SourceStory<'_>, identities: &HashMap<String, String>) -> FieldRestoration {
    let sources = crate::document_story_source::collect(story).paragraphs;
    if !sources.iter().any(|(_, paragraph)| paragraph.source_xml.as_ref().is_some_and(|xml| xml.contains("fldChar") || xml.contains("fldSimple"))) { return FieldRestoration::default(); }
    let mut blocks = HashMap::new(); collect_blocks(&document.blocks, None, &mut blocks);
    let mut paragraphs = Vec::new();
    let mut missing = Vec::new();
    for (address, paragraph) in &sources {
        let Some((block_id, text)) = blocks.get(address) else { missing.push(address.node_id.clone()); continue; };
        paragraphs.push(Paragraph { source: paragraph, source_id: address.node_id.clone(), block_id: block_id.clone(), text: text.clone() });
    }
    if !missing.is_empty() {
        let error = format!("DOCX native field paragraphs {missing:?} could not be mapped to extracted blocks");
        let retained = sources.iter().map(|(address, paragraph)| json!({ "sourceNodeId": address.node_id, "xml": paragraph.source_xml })).collect::<Vec<_>>();
        return FieldRestoration { warnings: vec![format!("{error}; source definitions were retained in the document payload.")],
            source: Some(json!({ "version": 1, "format": "docx", "story": story.story, "partName": story.part, "noteId": story.note_id, "mappingError": error, "paragraphs": retained, "identities": identities })) };
    }
    let prepared: Result<HashMap<String, Value>, String> = (|| {
        let id_prefix = if story.story == "body" { "word-field-".into() } else { format!("word-field:{}:{}:{}:", story.story, story.part, story.note_id.map(|id| id.to_string()).unwrap_or_default()) };
        let mut reader = Reader { identities: identities.clone(), id_prefix, ..Reader::default() }; let mut projections = HashMap::new();
        for (index, paragraph) in paragraphs.iter().enumerate() {
            let xml = paragraph.source.source_xml.as_deref().ok_or("A native field paragraph has no source XML.")?;
            let root = parse_xml(xml)?;
            let scope = BTreeMap::from([("w".into(), "http://schemas.openxmlformats.org/wordprocessingml/2006/main".into()),
                ("r".into(), "http://schemas.openxmlformats.org/officeDocument/2006/relationships".into())]);
            let style = root.child("pPr").and_then(|properties| properties.child("pStyle")).and_then(|style| style.attribute("val"));
            let mut projection = Projection { text: String::new(), offset: 0, spans: Vec::new(), note_reference_count: 0 };
            reader.tagged_identity = None;
            reader.visit(&root, xml, index, &mut projection, &scope, style, "")?;
            if projection.text != paragraph.text { return Err(format!("Native field visible-text projection differs from paragraph {:?}.", paragraph.block_id)); }
            projections.insert(paragraph.block_id.clone(), projection);
        }
        if !reader.pending.is_empty() { return Err("A native field is not closed before the end of its story.".into()); }
        let mut fields = HashMap::<String, Vec<Value>>::new(); let mut ids = std::collections::HashSet::new();
        for field in reader.fields {
            let (block, field) = field_value(field, &paragraphs)?;
            if !ids.insert(field["id"].as_str().unwrap_or_default().to_owned()) { return Err("Native field identities are duplicated.".into()); }
            fields.entry(block).or_default().push(field);
        }
        let mut details = HashMap::new();
        fn prepare(blocks: &[DocumentBlock], note_id: Option<i64>, story: &crate::document_story_source::SourceStory<'_>, fields: &HashMap<String, Vec<Value>>, projections: &HashMap<String, Projection>, output: &mut HashMap<String, Value>) -> Result<(), String> {
            for block in blocks {
                let note_id = block.note_id.or(note_id);
                if belongs_to_story(block, note_id, story) && paragraph_block(block) && (projections.contains_key(&block.id) || fields.contains_key(&block.id)) {
                    let mut details = block.details.to_json_value();
                    if let Some(fields) = fields.get(&block.id) {
                        let marker = &details["officeNumberingMarker"];
                        let prefix = (marker["version"].as_u64() == Some(1)).then(|| marker["text"].as_str().map(|text| text.encode_utf16().count())).flatten();
                        let fields = fields.iter().filter(|field| !prefix.is_some_and(|prefix| field.get("endBlockId").is_none()
                            && field["startOffset"].as_u64().is_some_and(|start| start < prefix as u64)
                            && field["endOffset"].as_u64().is_some_and(|end| end <= prefix as u64))).collect::<Vec<_>>();
                        if !fields.is_empty() { details["documentTextFields"] = json!({ "version": 1, "fields": fields }); }
                    }
                    if let Some(projection) = projections.get(&block.id) { inline_sources(&mut details, &projection.spans)?; }
                    output.insert(block.id.clone(), details);
                }
                prepare(&block.children, note_id, story, fields, projections, output)?;
            }
            Ok(())
        }
        prepare(&document.blocks, None, story, &fields, &projections, &mut details)?;
        Ok(details)
    })();
    let mut warnings = Vec::new();
    fn visit(blocks: &mut [DocumentBlock], note_id: Option<i64>, story: &crate::document_story_source::SourceStory<'_>, prepared: &mut HashMap<String, Value>) {
        for block in blocks {
            let note_id = block.note_id.or(note_id);
            if belongs_to_story(block, note_id, story) {
                if let Some(details) = prepared.remove(&block.id) { block.details = BlockDetails::Raw(details); }
            }
            visit(&mut block.children, note_id, story, prepared);
        }
    }
    match prepared {
        Ok(mut details) => visit(&mut document.blocks, None, story, &mut details),
        Err(error) => {
            fn retain(blocks: &mut [DocumentBlock], note_id: Option<i64>, story: &crate::document_story_source::SourceStory<'_>, paragraphs: &[Paragraph<'_>], error: &str) {
                for block in blocks {
                    let note_id = block.note_id.or(note_id);
                    if let Some(paragraph) = paragraphs.iter().find(|paragraph| paragraph.block_id == block.id && belongs_to_story(block, note_id, story)) {
                        let mut details = block.details.to_json_value();
                        details["nativeFieldSource"] = json!({ "version": 1, "xml": paragraph.source.source_xml, "mappingError": error });
                        block.details = BlockDetails::Raw(details);
                    }
                    retain(&mut block.children, note_id, story, paragraphs, error);
                }
            }
            retain(&mut document.blocks, None, story, &paragraphs, &error);
            warnings.push(format!("DOCX native fields could not be mapped safely: {error}; source definitions were retained."));
        }
    }
    FieldRestoration { warnings, source: None }
}

fn belongs_to_story(block: &DocumentBlock, note_id: Option<i64>, story: &crate::document_story_source::SourceStory<'_>) -> bool {
    note_id == story.note_id && block.source_refs.iter().any(|reference| reference.part_name == story.part && reference.story == story.story)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn utf16_ranges_preserve_emoji_and_reject_half_surrogates() {
        assert_eq!(slice_utf16("a😀b", 1, 3), Some("😀"));
        assert_eq!(slice_utf16("a😀b", 2, 3), None);
    }

    #[test]
    fn field_reader_preserves_mapped_non_ascii_identity() {
        let mut reader = Reader { identities: HashMap::from([("office_field_v1_00000001".into(), "字段-α".into())]), ..Reader::default() };
        let xml = "<w:p><w:bookmarkStart w:name=\"office_field_v1_00000001\"/><w:fldSimple w:instr=\"DATE\"><w:r><w:t>2026</w:t></w:r></w:fldSimple></w:p>";
        let mut projection = Projection { text: String::new(), offset: 0, spans: Vec::new(), note_reference_count: 0 };
        reader.visit(&parse_xml(xml).unwrap(), xml, 0, &mut projection, &BTreeMap::new(), None, "").unwrap();
        assert_eq!(reader.fields[0].pending.id, "字段-α");
        assert_eq!(projection.text, "2026");
    }

    #[test]
    fn absent_extracted_paragraph_retains_complete_field_source_in_payload() {
        let xml = "<w:p><w:fldSimple w:instr=\"DATE\"><w:r><w:t/></w:r></w:fldSimple></w:p>";
        let source: DocModel = serde_json::from_value(json!({ "nodes": [{ "blockId": "empty-source", "type": "paragraph", "children": [], "sourceXml": xml }],
            "metadata": { "sourceParts": 1, "warnings": [], "headerSections": [], "footerSections": [], "paragraphStyles": [] } })).unwrap();
        let mut document: StructuredDocument = serde_json::from_value(json!({ "schemaVersion": 1, "parserVersion": "draft", "documentId": "doc", "blocks": [], "warnings": [], "geometryStatus": "unknown" })).unwrap();
        let restored = restore(&mut document, &source, &HashMap::new());
        let retained = restored.source.unwrap();
        assert_eq!(retained["paragraphs"][0]["sourceNodeId"], "empty-source");
        assert_eq!(retained["paragraphs"][0]["xml"], xml);
        assert!(!restored.warnings.is_empty());
        assert!(document.blocks.is_empty());
    }

    #[test]
    fn field_language_can_come_from_inherited_source_run_style() {
        let paragraph: ParagraphNode = serde_json::from_value(json!({ "type": "paragraph", "children": [
            { "type": "text", "text": "a😀b", "style": { "language": "de-DE" } }
        ] })).unwrap();
        assert_eq!(source_language(&paragraph, 1).as_deref(), Some("de-DE"));
    }

    #[test]
    fn bookmark_identity_part_requires_its_own_namespace() {
        use std::io::Write;
        use zip::{ZipWriter, write::SimpleFileOptions};
        for (namespace, valid) in [("urn:office-studio:document-fields:1", true), ("urn:unrelated", false)] {
            let xml = format!("<fields xmlns=\"{namespace}\" version=\"1\"><field key=\"office_field_v1_00000001\" id=\"字段\"/></fields>");
            let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
            writer.start_file("customXml/office-fields.xml", SimpleFileOptions::default()).unwrap();
            writer.write_all(xml.as_bytes()).unwrap();
            let bytes = writer.finish().unwrap().into_inner();
            let mut archive = ZipArchive::new(Cursor::new(bytes.as_slice())).unwrap();
            assert_eq!(read_identities(&mut archive).is_ok(), valid);
        }
    }
}
