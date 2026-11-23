use std::{cell::RefCell, collections::{BTreeMap, HashMap}, io::Cursor};

use duke_pages::{ParagraphStyle, TextStyle, parse::{ContentTypeLookup, ParseContext, ParsedStyleSheet, parse_paragraph_style, parse_run_style}};
use serde::Serialize;
use zip::ZipArchive;

use crate::document_format_sidecar::{Element, canonical_resource, parse_xml, read_xml};

#[derive(Debug, Default, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct StoryText {
    pub text: String,
    pub fields: Vec<StoryField>,
    pub runs: Vec<StoryRun>,
    pub paragraphs: Vec<StoryParagraph>,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub locks: Vec<StoryLock>,
    pub source_xml: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub relationships_xml: Option<String>,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub unsupported_features: Vec<String>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct StoryRun {
    pub text: String,
    pub style: Option<TextStyle>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub link: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub source_format: Option<StorySourceFormat>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct StorySourceFormat {
    pub version: u8,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub properties_xml: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub paragraph_style_id: Option<String>,
    /// Complete inherited bindings, not only declarations on rPr/pPr itself.
    pub namespaces: BTreeMap<String, String>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct StoryParagraph {
    /// Literal UTF-16 offsets. PAGE / NUMPAGES / DATE / TIME results are objects.
    pub start_offset: usize,
    pub end_offset: usize,
    pub style: Option<ParagraphStyle>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub word_wrap: Option<bool>,
    pub source_format: StorySourceFormat,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct StoryField {
    pub field: String,
    pub offset: usize,
    pub instruction: String,
    pub cached_result: String,
    pub text_style: Option<TextStyle>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub source_format: Option<StorySourceFormat>,
    /// Keep every cached run; the result is one input object, not one style.
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub cached_runs: Vec<StoryRun>,
    pub locked: bool,
    pub contains_nested_fields: bool,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub nested_fields: Vec<StoryNestedField>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct StoryNestedField {
    pub field: String,
    pub instruction: String,
    pub start_offset: usize,
    pub end_offset: usize,
    pub locked: bool,
    pub contains_nested_fields: bool,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub nested_fields: Vec<StoryNestedField>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct StoryLock {
    pub id: String,
    /// Input-buffer UTF-16 offsets, where a field occupies one U+FFFC.
    pub start_offset: usize,
    pub end_offset: usize,
    pub lock: String,
}

pub fn read_story(archive: &mut ZipArchive<Cursor<&[u8]>>, part: &str, sheet: &ParsedStyleSheet) -> Result<Option<StoryText>, String> {
    let Some(xml) = read_xml(archive, part)? else { return Ok(None); };
    let (directory, file) = part.trim_start_matches('/').rsplit_once('/').unwrap_or(("", part));
    let relations_part = if directory.is_empty() { format!("_rels/{file}.rels") } else { format!("{directory}/_rels/{file}.rels") };
    let relations_xml = read_xml(archive, &relations_part)?;
    let mut relationships = HashMap::new();
    if let Some(xml) = relations_xml.as_ref() {
        for node in parse_xml(xml)?.children {
            if node.name == "Relationship" && node.attribute("Type").is_some_and(|kind| kind.ends_with("/hyperlink")) {
                if let (Some(id), Some(target)) = (node.attribute("Id"), node.attribute("Target")) { relationships.insert(id.to_owned(), target.to_owned()); }
            }
        }
    }
    let mut story = parse_story_with_styles(&xml, sheet, relationships)?;
    story.relationships_xml = relations_xml;
    Ok(Some(story))
}

#[derive(Default)]
struct Pending {
    instruction: String, cached: Vec<StoryRun>, text_style: Option<TextStyle>, source_format: Option<StorySourceFormat>,
    separated: bool, locked: bool, contains_nested_fields: bool, nested_fields: Vec<StoryNestedField>,
}

struct StoryReader<'a> {
    story: StoryText,
    pending: Option<Pending>,
    parents: Vec<Pending>,
    literal_offset: usize,
    buffer_offset: usize,
    xml: &'a str,
    context: ParseContext<'a, &'a ParsedStyleSheet>,
}

impl StoryReader<'_> {
    fn unsupported(&mut self, kind: &str) {
        if !self.story.unsupported_features.iter().any(|value| value == kind) { self.story.unsupported_features.push(kind.to_owned()); }
    }

    fn run_style(&self, xml: &str, properties: Option<&Element>, paragraph_style: Option<&str>) -> Option<TextStyle> {
        let mut style = parse_run_style(xml, &self.context, paragraph_style);
        if properties.and_then(|properties| properties.child("shd")).and_then(|shading| shading.attribute("val")) == Some("nil") {
            if let Some(style) = style.as_mut() { style.background_color = None; }
        }
        style
    }

    fn source_format(&self, properties: Option<&Element>, paragraph_style: Option<&str>, inherited: &BTreeMap<String, String>) -> StorySourceFormat {
        let mut namespaces = inherited.clone();
        if let Some(properties) = properties { namespaces.extend(properties.namespaces.iter().map(|(prefix, namespace)| (prefix.clone(), namespace.clone()))); }
        StorySourceFormat { version: 1, properties_xml: properties.map(|node| self.xml[node.start..node.end].to_owned()),
            paragraph_style_id: paragraph_style.map(str::to_owned), namespaces }
    }

    fn append(&mut self, text: &str, style: Option<TextStyle>, link: Option<String>, source_format: Option<StorySourceFormat>) {
        let run = StoryRun { text: text.to_owned(), style, link, source_format };
        if let Some(pending) = &mut self.pending {
            if pending.separated { pending.cached.push(run); }
            return;
        }
        let length = text.encode_utf16().count();
        self.literal_offset += length;
        self.buffer_offset += length;
        self.story.text.push_str(text);
        self.story.runs.push(run);
    }

    fn finish_field(&mut self, pending: Pending) {
        let instruction = pending.instruction.trim();
        let name = instruction.split_whitespace().next().unwrap_or_default().to_ascii_uppercase();
        let cached_result = pending.cached.iter().map(|run| run.text.as_str()).collect::<String>();
        if let Some(mut parent) = self.parents.pop() {
            parent.contains_nested_fields = true;
            if parent.separated {
                let start_offset = parent.cached.iter().map(|run| run.text.encode_utf16().count()).sum::<usize>();
                parent.nested_fields.push(StoryNestedField { field: name, instruction: instruction.to_owned(), start_offset,
                    end_offset: start_offset + cached_result.encode_utf16().count(), locked: pending.locked,
                    contains_nested_fields: pending.contains_nested_fields, nested_fields: pending.nested_fields });
                parent.cached.extend(pending.cached);
            } else { self.unsupported("nestedFieldCodes"); }
            self.pending = Some(parent);
            return;
        }
        if matches!(name.as_str(), "PAGE" | "NUMPAGES" | "DATE" | "TIME") {
            self.story.fields.push(StoryField { field: name, offset: self.literal_offset, instruction: instruction.to_owned(), cached_result,
                text_style: pending.cached.first().and_then(|run| run.style.clone()).or(pending.text_style).or_else(|| self.context.styles().default_run_style.clone()),
                source_format: pending.cached.first().and_then(|run| run.source_format.clone()).or(pending.source_format), cached_runs: pending.cached,
                locked: pending.locked, contains_nested_fields: pending.contains_nested_fields, nested_fields: pending.nested_fields });
            self.buffer_offset += 1;
        } else {
            self.unsupported("unknownFields");
            for run in pending.cached { self.append(&run.text, run.style, run.link, run.source_format); }
        }
    }

    fn lock_properties(&mut self, node: &Element) -> Result<Option<(String, String, bool)>, String> {
        let Some(properties) = node.child("sdtPr") else { self.unsupported("contentControls"); return Ok(None); };
        if properties.children.iter().any(|property| !matches!(property.name.as_str(), "id" | "lock" | "tag" | "text")) { self.unsupported("contentControls"); }
        let Some(lock @ ("contentLocked" | "sdtLocked" | "sdtContentLocked")) = properties.child("lock").and_then(|lock| lock.attribute("val")) else { self.unsupported("contentControls"); return Ok(None); };
        let tag = properties.child("tag").and_then(|tag| tag.attribute("val"));
        let encoded = tag.and_then(|tag| tag.strip_prefix("office_story_lock_v1_"));
        let id = if let Some(encoded) = encoded {
            if encoded.len() > 2048 || encoded.len() % 2 != 0 || !encoded.is_ascii() { return Err("A story lock tag is malformed.".into()); }
            let bytes = (0..encoded.len()).step_by(2).map(|index| u8::from_str_radix(&encoded[index..index + 2], 16).map_err(|_| "A story lock tag is malformed.".to_owned())).collect::<Result<Vec<_>, _>>()?;
            String::from_utf8(bytes).map_err(|_| "A story lock tag has invalid UTF-8.".to_owned())?
        } else {
            properties.child("id").and_then(|id| id.attribute("val")).map(str::to_owned).unwrap_or_else(|| format!("story-lock-{}", node.start))
        };
        if id.is_empty() { return Err("A story lock has an empty identity.".into()); }
        Ok(Some((id, lock.to_owned(), encoded.is_some())))
    }

    fn push_lock(&mut self, id: String, kind: String, start: usize, end: usize, grouped: bool) -> Result<(), String> {
        if let Some(lock) = self.story.locks.iter_mut().find(|lock| lock.id == id) {
            if !grouped || lock.lock != kind { return Err("A story has duplicate or inconsistent lock identities.".into()); }
            lock.start_offset = lock.start_offset.min(start); lock.end_offset = lock.end_offset.max(end);
        } else { self.story.locks.push(StoryLock { id, start_offset: start, end_offset: end, lock: kind }); }
        Ok(())
    }

    fn visit_block(&mut self, node: &Element, inherited: &BTreeMap<String, String>) -> Result<(), String> {
        let scope = namespace_scope(node, inherited);
        match node.name.as_str() {
            "p" => {
                if self.pending.is_some() { return Err("A story field crosses a paragraph boundary.".into()); }
                if !self.story.paragraphs.is_empty() { self.append("\n", self.context.styles().default_run_style.clone(), None, None); }
                let start_offset = self.literal_offset;
                let properties = node.child("pPr");
                let style_id = properties.and_then(|properties| properties.child("pStyle")).and_then(|style| style.attribute("val"));
                let style = parse_paragraph_style(&self.xml[node.start..node.end], &self.context);
                let word_wrap = properties.and_then(|properties| properties.child("wordWrap")).map(|property| !matches!(property.attribute("val"), Some("0" | "off" | "false")));
                let previous_runs = self.story.runs.len();
                for child in &node.children { self.visit_inline(child, style_id, None, &scope)?; }
                if self.pending.is_some() { return Err("A story field is not closed.".into()); }
                if self.story.runs.len() == previous_runs {
                    let run = properties.and_then(|properties| properties.child("rPr")).map(|properties| format!("<w:r>{}</w:r>", &self.xml[properties.start..properties.end])).unwrap_or_else(|| "<w:r/>".into());
                    let run_properties = properties.and_then(|properties| properties.child("rPr"));
                    let mark_scope = properties.map(|properties| namespace_scope(properties, &scope)).unwrap_or_else(|| scope.clone());
                    let source_format = self.source_format(run_properties, style_id, &mark_scope);
                    self.append("", self.run_style(&run, run_properties, style_id), None, Some(source_format));
                }
                let source_format = self.source_format(properties, style_id, &scope);
                self.story.paragraphs.push(StoryParagraph { start_offset, end_offset: self.literal_offset, style, word_wrap, source_format });
            }
            "sdt" => {
                let lock = self.lock_properties(node)?;
                let start = self.buffer_offset + usize::from(!self.story.paragraphs.is_empty());
                if let Some(content) = node.child("sdtContent") {
                    let scope = namespace_scope(content, &scope);
                    for child in &content.children { self.visit_block(child, &scope)?; }
                }
                if let Some((id, kind, grouped)) = lock { self.push_lock(id, kind, start.min(self.buffer_offset), self.buffer_offset, grouped)?; }
            }
            "tbl" => { self.unsupported("tables"); for child in &node.children { self.visit_block(child, &scope)?; } }
            "tr" | "tc" | "sdtContent" | "ins" | "customXml" => for child in &node.children { self.visit_block(child, &scope)?; },
            "tblPr" | "trPr" | "tcPr" | "tblGrid" | "bookmarkStart" | "bookmarkEnd" | "proofErr" => {},
            "del" | "moveFrom" => self.unsupported("trackedChanges"),
            _ => { self.unsupported(&node.name); for child in &node.children { self.visit_block(child, &scope)?; } }
        }
        Ok(())
    }

    fn visit_inline(&mut self, node: &Element, paragraph_style: Option<&str>, link: Option<String>, inherited: &BTreeMap<String, String>) -> Result<(), String> {
        let scope = namespace_scope(node, inherited);
        match node.name.as_str() {
            "r" => {
                let style = self.run_style(&self.xml[node.start..node.end], node.child("rPr"), paragraph_style);
                let source_format = self.source_format(node.child("rPr"), paragraph_style, &scope);
                if node.children.iter().all(|child| matches!(child.name.as_str(), "rPr" | "lastRenderedPageBreak")) { self.append("", style.clone(), link.clone(), Some(source_format.clone())); }
                for child in &node.children {
                    match child.name.as_str() {
                        "t" => self.append(&child.text, style.clone(), link.clone(), Some(source_format.clone())),
                        "tab" => self.append("\t", style.clone(), link.clone(), Some(source_format.clone())),
                        "br" | "cr" => self.append("\u{2028}", style.clone(), link.clone(), Some(source_format.clone())),
                        "instrText" => self.pending.as_mut().ok_or("A story has an instruction outside a field.")?.instruction.push_str(&child.text),
                        "fldChar" => match child.attribute("fldCharType") {
                            Some("begin") => {
                                if let Some(mut parent) = self.pending.take() { parent.contains_nested_fields = true; self.parents.push(parent); }
                                self.pending = Some(Pending { text_style: style.clone(), source_format: Some(source_format.clone()),
                                    locked: child.attribute("fldLock").is_some_and(|value| !matches!(value, "0" | "off" | "false")), ..Default::default() });
                            },
                            Some("separate") => {
                                let pending = self.pending.as_mut().ok_or("A story field separator has no beginning.")?;
                                if pending.separated { return Err("A story field has two result separators.".into()); }
                                pending.separated = true;
                            },
                            Some("end") => { let pending = self.pending.take().ok_or("A story field has no beginning.")?; self.finish_field(pending); },
                            _ => {},
                        },
                        "rPr" | "lastRenderedPageBreak" => {},
                        _ => self.unsupported(&child.name),
                    }
                }
            }
            "fldSimple" => {
                if let Some(mut parent) = self.pending.take() { parent.contains_nested_fields = true; self.parents.push(parent); }
                self.pending = Some(Pending { instruction: node.attribute("instr").ok_or("A simple story field has no instruction.")?.to_owned(),
                    text_style: parse_run_style("<w:r/>", &self.context, paragraph_style), source_format: Some(self.source_format(None, paragraph_style, &scope)),
                    separated: true, locked: node.attribute("fldLock").is_some_and(|value| !matches!(value, "0" | "off" | "false")), ..Default::default() });
                for child in &node.children { self.visit_inline(child, paragraph_style, link.clone(), &scope)?; }
                let pending = self.pending.take().ok_or("A simple story field has an invalid nested delimiter.")?;
                self.finish_field(pending);
            }
            "hyperlink" => {
                let target = node.attribute("anchor").map(|anchor| format!("#{anchor}")).or_else(|| node.attribute("id").and_then(|id| self.context.relationships.get(id).cloned()));
                if target.is_none() { self.unsupported("unresolvedHyperlink"); }
                for child in &node.children { self.visit_inline(child, paragraph_style, target.clone(), &scope)?; }
            }
            "sdt" => {
                let lock = self.lock_properties(node)?;
                let inside_field = self.pending.is_some();
                if inside_field { self.unsupported("fieldContentControls"); }
                let start = self.buffer_offset;
                if let Some(content) = node.child("sdtContent") {
                    let scope = namespace_scope(content, &scope);
                    for child in &content.children { self.visit_inline(child, paragraph_style, link.clone(), &scope)?; }
                }
                if let Some((id, kind, grouped)) = lock.filter(|_| !inside_field) { self.push_lock(id, kind, start, self.buffer_offset, grouped)?; }
            }
            "sdtContent" | "ins" | "customXml" => for child in &node.children { self.visit_inline(child, paragraph_style, link.clone(), &scope)?; },
            "pPr" | "bookmarkStart" | "bookmarkEnd" | "proofErr" | "commentRangeStart" | "commentRangeEnd" => {},
            "del" | "moveFrom" => self.unsupported("trackedChanges"),
            _ => self.unsupported(&node.name),
        }
        Ok(())
    }
}

fn namespace_scope(node: &Element, inherited: &BTreeMap<String, String>) -> BTreeMap<String, String> {
    let mut result = inherited.clone();
    result.extend(node.namespaces.iter().map(|(prefix, namespace)| (prefix.clone(), namespace.clone())));
    result
}

fn parse_story_with_styles(xml: &str, sheet: &ParsedStyleSheet, relationships: HashMap<String, String>) -> Result<StoryText, String> {
    let expected = match parse_xml(xml)?.name.as_str() { "hdr" => "w:hdr", "ftr" => "w:ftr", _ => return Err("Story XML must have a Word header/footer root.".into()) };
    let canonical = canonical_resource(xml, expected)?;
    let root = parse_xml(&canonical)?;
    let parts = HashMap::new();
    let assets = HashMap::new();
    let context = ParseContext { relationships, content_types: ContentTypeLookup::default(), parts: &parts, binary_assets: &assets, style_sheet: sheet, warnings: RefCell::new(Vec::new()) };
    let mut reader = StoryReader { story: StoryText { source_xml: xml.to_owned(), ..Default::default() }, pending: None, parents: Vec::new(), literal_offset: 0, buffer_offset: 0, xml: &canonical, context };
    let scope = namespace_scope(&root, &BTreeMap::new());
    for node in &root.children { reader.visit_block(node, &scope)?; }
    if reader.pending.is_some() { return Err("A story field is not closed.".into()); }
    if reader.story.paragraphs.is_empty() {
        reader.append("", sheet.default_run_style.clone(), None, Some(reader.source_format(None, None, &scope)));
        let source_format = reader.source_format(None, None, &scope);
        reader.story.paragraphs.push(StoryParagraph { start_offset: 0, end_offset: 0, style: sheet.default_paragraph_style.clone(), word_wrap: None, source_format });
    }
    Ok(reader.story)
}

#[cfg(test)]
pub(crate) fn parse_story(xml: &str) -> Option<StoryText> {
    parse_story_with_styles(xml, &ParsedStyleSheet::empty(), HashMap::new()).ok()
}
