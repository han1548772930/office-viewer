use base64::{engine::general_purpose::STANDARD, Engine as _};
use serde_json::Value;
use std::collections::{BTreeMap, BTreeSet, HashMap, HashSet};

use crate::{opc::Package, xml, ExportError, ExportWarning};

#[path = "docx_source_xml.rs"]
mod source_xml;
#[path = "docx_source_styles.rs"]
mod source_styles;
#[path = "docx_legacy_source.rs"]
mod legacy_source;
#[path = "docx_default_style.rs"]
mod default_style;
use default_style::RunDefaultEffects;
#[path = "docx_story.rs"]
mod story_writer;
#[path = "docx_text_fields.rs"]
mod text_fields;
use text_fields::TextFieldPlan;
#[path = "docx_table.rs"]
mod table_writer;
#[path = "docx_block_hosts.rs"]
mod block_hosts;
#[path = "docx_notes.rs"]
mod note_writer;
#[path = "docx_custom_xml.rs"]
mod custom_xml;
#[path = "docx_content_controls.rs"]
mod control_writer;
#[path = "docx_content_control_ranges.rs"]
mod control_ranges;
#[path = "docx_content_control_pictures.rs"]
mod control_pictures;

#[path = "docx_numbering.rs"]
mod numbering;
use numbering::{assign_list_groups, numbering_left, numbering_level, prepare_sequence_groups, write_numbering_xml, NumberingInstance, SequencePlan};
#[path = "docx_numbering_marker.rs"]
mod numbering_marker;
use numbering_marker::{prepare_preserved_marker_groups, PreservedMarkerPlan, PreservedParagraph};
#[cfg(test)]
use numbering::{assign_lists, build_numbering_xml, prepare_sequences};
#[cfg(test)]
use numbering_marker::prepare_preserved_markers;

const REL_OFFICE_DOCUMENT: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const REL_CORE_PROPERTIES: &str =
    "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
const REL_STYLES: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
const REL_THEME: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
const REL_NUMBERING: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
const REL_HEADER: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
const REL_FOOTER: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
const REL_SETTINGS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
const REL_IMAGE: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const REL_HYPERLINK: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
const REL_FOOTNOTES: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";
const REL_ENDNOTES: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes";
const REL_COMMENTS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const REL_COMMENTS_EXTENDED: &str =
    "http://schemas.microsoft.com/office/2011/relationships/commentsExtended";
const DOCUMENT_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const STYLES_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
const NUMBERING_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";
const FOOTNOTES_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml";
const ENDNOTES_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml";
const COMMENTS_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
const COMMENTS_EXTENDED_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml";
const HEADER_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";
const FOOTER_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";
const SETTINGS_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";
const EMU_PER_PIXEL: f64 = 9_525.0;
const TWIPS_PER_PIXEL: f64 = 15.0;
const CONTENT_NAMESPACES: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture""#;

#[derive(Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
enum NoteKind {
    Footnote,
    Endnote,
}

impl NoteKind {
    fn field(self) -> &'static str {
        match self {
            Self::Footnote => "footnoteId",
            Self::Endnote => "endnoteId",
        }
    }

    fn collection(self) -> &'static str {
        match self {
            Self::Footnote => "footnotes",
            Self::Endnote => "endnotes",
        }
    }

    fn reference_tag(self) -> &'static str {
        match self {
            Self::Footnote => "footnoteReference",
            Self::Endnote => "endnoteReference",
        }
    }

    fn body_reference_tag(self) -> &'static str {
        match self {
            Self::Footnote => "footnoteRef",
            Self::Endnote => "endnoteRef",
        }
    }

    fn style(self) -> &'static str {
        match self {
            Self::Footnote => "FootnoteText",
            Self::Endnote => "EndnoteText",
        }
    }
}

#[derive(Clone)]
struct NoteEntry {
    id: i64,
    value: Value,
}

#[derive(Clone)]
struct CommentEntry {
    id: u32,
    paragraph_id: String,
    parent_paragraph_id: Option<String>,
    author: String,
    date: String,
    text: String,
    resolved: bool,
}

#[derive(Clone)]
struct CommentAnchor {
    id: u32,
    start: usize,
    end: usize,
    starts_here: bool,
    ends_here: bool,
}

#[derive(Default)]
struct CommentEventCursor {
    events: Vec<(usize, String)>,
    next: usize,
}

struct NoteBundle {
    footnotes: Vec<NoteEntry>,
    endnotes: Vec<NoteEntry>,
    ids: HashMap<(NoteKind, String), i64>,
    orphaned: Vec<(NoteKind, String)>,
    reference_properties: HashMap<(NoteKind, String), note_writer::ReferenceProperties>,
}

#[derive(Clone)]
struct ContentControl {
    id: String,
    kind: String,
    numeric_id: u32,
    start: usize,
    end: usize,
    value: Value,
}

#[derive(Clone, Copy)]
struct PageGeometry {
    width: u32,
    height: u32,
    top: u32,
    right: u32,
    bottom: u32,
    left: u32,
}

#[derive(Default)]
struct SectionHeaderFooterRefs {
    references: Vec<(String, String)>,
    different_first_page: bool,
    header_distance: u32,
    footer_distance: u32,
    note_properties_xml: String,
}

struct WriteContext<'a> {
    package: &'a mut Package,
    warnings: Vec<ExportWarning>,
    list_ids: HashMap<usize, u32>,
    numbering: Vec<NumberingInstance>,
    sequences: SequencePlan,
    preserved_markers: PreservedMarkerPlan,
    note_ids: HashMap<(NoteKind, String), i64>,
    note_reference_properties: HashMap<(NoteKind, String), note_writer::ReferenceProperties>,
    comment_entries: Vec<CommentEntry>,
    comment_anchors: HashMap<String, Vec<CommentAnchor>>,
    controls_by_block: HashMap<String, Vec<ContentControl>>,
    control_sources: control_writer::ControlWriteState,
    used_controls: BTreeSet<String>,
    next_relationship: u32,
    relationship_owner: String,
    next_image: u32,
    page: PageGeometry,
    section_header_footer_refs: HashMap<String, SectionHeaderFooterRefs>,
    current_section: String,
    run_defaults: RunDefaultEffects,
    story_defaults: Value,
    next_story_control: u32,
    source_styles: source_styles::SourceStyles,
    text_fields: TextFieldPlan,
    tables: table_writer::TableWriteState,
    custom_xml_stores: BTreeSet<String>,
}

impl<'a> WriteContext<'a> {
    fn relationship(
        &mut self,
        kind: &str,
        target: &str,
        external: bool,
    ) -> Result<String, ExportError> {
        let id = format!("rId{}", self.next_relationship);
        self.next_relationship += 1;
        self.package
            .add_relationship(Some(&self.relationship_owner), &id, kind, target, external)?;
        Ok(id)
    }

    fn with_relationship_owner<T>(&mut self, owner: &str, write: impl FnOnce(&mut Self) -> Result<T, ExportError>) -> Result<T, ExportError> {
        let previous = std::mem::replace(&mut self.relationship_owner, owner.to_owned());
        let result = write(self);
        self.relationship_owner = previous;
        result
    }

    fn warn(&mut self, code: &'static str, message: impl Into<String>) {
        self.warnings.push(ExportWarning::new(code, message));
    }

    fn image(&mut self, block: &Value) -> Result<Option<String>, ExportError> {
        let metadata = get(block, "metadata");
        let image = get(metadata, "image");
        let src = string(get(image, "src"))
            .or_else(|| string(get(metadata, "src")))
            .unwrap_or_default();
        let alt = string(get(image, "alt"))
            .or_else(|| string(get(metadata, "alt")))
            .unwrap_or_default();
        let embedded = match decode_data_image(&src) {
            Ok(image) => Some(image),
            Err(_) if control_pictures::linked_image(&src) => None,
            Err(reason) => {
                self.warn(
                    "DOCX_IMAGE_UNSUPPORTED",
                    format!("Image was exported as alt text because {reason}."),
                );
                return Ok(None);
            }
        };

        let width = positive(get(image, "width"), 320.0);
        let height = positive(get(image, "height"), 180.0);
        let content_width_twips = self
            .page
            .width
            .saturating_sub(self.page.left + self.page.right)
            .max(720);
        let max_width_emu = u64::from(content_width_twips) * 635;
        let mut width_emu = (width * EMU_PER_PIXEL).round().max(1.0) as u64;
        let mut height_emu = (height * EMU_PER_PIXEL).round().max(1.0) as u64;
        if width_emu > max_width_emu {
            height_emu = height_emu.saturating_mul(max_width_emu) / width_emu;
            width_emu = max_width_emu;
        }

        let index = self.next_image;
        self.next_image += 1;
        let (relationship, reference, extension) = if let Some((content_type, extension, bytes)) = embedded {
            let path = format!("word/media/image{index}.{extension}");
            self.package.add_binary(&path, content_type, bytes)?;
            let target = format!("media/image{index}.{extension}");
            (self.relationship(REL_IMAGE, &target, false)?, "embed", extension)
        } else { (self.relationship(REL_IMAGE, &src, true)?, "link", "linked") };
        let alt = xml_attribute(&alt);
        Ok(Some(format!(
            "<w:r><w:drawing><wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\"><wp:extent cx=\"{width_emu}\" cy=\"{height_emu}\"/><wp:effectExtent l=\"0\" t=\"0\" r=\"0\" b=\"0\"/><wp:docPr id=\"{index}\" name=\"Image {index}\" descr=\"{alt}\"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect=\"1\"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><pic:pic><pic:nvPicPr><pic:cNvPr id=\"{index}\" name=\"Image {index}.{extension}\" descr=\"{alt}\"/><pic:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></pic:cNvPicPr></pic:nvPicPr><pic:blipFill><a:blip r:{reference}=\"{relationship}\"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{width_emu}\" cy=\"{height_emu}\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"
        )))
    }

    fn runs(&mut self, runs: &[Value]) -> Result<String, ExportError> {
        let mut result = String::new();
        for run in runs {
            let text = string(get(run, "text")).unwrap_or_default();
            let formula = get(get(run, "metadata"), "latexFormula");
            let formula_source = string(get(formula, "latex"));
            let formula_mode = string(get(formula, "mode"));
            if let (Some(source), Some(mode)) = (
                formula_source,
                formula_mode.filter(|mode| mode == "inline" || mode == "block"),
            ) {
                let source_run = run_with_text(run, &source);
                let fragments = self.text_run_xml(&source_run, &source)?;
                let tag = xml_attribute(&format!("office-latex:v1:{mode}"));
                result.push_str(&format!("<w:sdt><w:sdtPr><w:tag w:val=\"{tag}\"/><w:text/></w:sdtPr><w:sdtContent>{fragments}</w:sdtContent></w:sdt>"));
                self.warn(
                    "DOCX_LATEX_SOURCE_FALLBACK",
                    "LaTeX source and mode were preserved in a tagged content control. Word will display the source because this writer does not convert LaTeX to native OfficeMath.",
                );
                continue;
            }
            let (fragments, has_note_reference) = if let Some((kind, source_id)) =
                note_reference(run)
            {
                if let Some(id) = self.note_ids.get(&(kind, source_id.clone())).copied() {
                    let properties = self.note_reference_properties.get(&(kind, source_id.clone())).cloned().unwrap_or_default();
                    (note_writer::reference_run(self, run, kind, id, &properties)?, true)
                } else {
                    self.warn(
                        "DOCX_NOTE_MISSING",
                        format!("A {0} reference to note {source_id:?} has no matching note body; its marker text was preserved.", kind.collection().trim_end_matches('s')),
                    );
                    (self.text_run_xml(run, &text)?, false)
                }
            } else {
                (self.text_run_xml(run, &text)?, false)
            };
            if has_note_reference {
                result.push_str(&fragments);
                continue;
            }
            let href = run_link(run);
            if let Some(href) = href {
                if let Some(bookmark) = href.strip_prefix('#') {
                    result.push_str(&format!(
                        "<w:hyperlink w:anchor=\"{}\" w:history=\"1\">{fragments}</w:hyperlink>",
                        xml_attribute(bookmark)
                    ));
                } else if href.starts_with("https://")
                    || href.starts_with("http://")
                    || href.starts_with("mailto:")
                {
                    let relationship = self.relationship(REL_HYPERLINK, &href, true)?;
                    result.push_str(&format!(
                        "<w:hyperlink r:id=\"{relationship}\" w:history=\"1\">{fragments}</w:hyperlink>"
                    ));
                } else {
                    self.warn(
                        "DOCX_LINK_UNSUPPORTED",
                        "A link with an unsupported target was exported as plain text.",
                    );
                    result.push_str(&fragments);
                }
            } else {
                result.push_str(&fragments);
            }
        }
        if result.is_empty() {
            result.push_str("<w:r><w:t></w:t></w:r>");
        }
        Ok(result)
    }

    fn run_properties(&mut self, run: &Value) -> Result<source_xml::Element, ExportError> {
        let generated = run_properties_xml_with_defaults(run, self.run_defaults);
        let mut provenance = run.clone();
        if let Some(baseline) = get(get(run, "metadata"), "officeTextStyleBaseline").as_object() {
            if get(get(&provenance, "metadata"), "officeStoryStyleBaseline").is_null() {
                provenance["metadata"]["officeStoryStyleBaseline"] = Value::Object(baseline.clone());
            }
        }
        let properties = self.source_styles.story_run_properties(&generated, run, &provenance).map_err(ExportError::invalid)?;
        let mut properties = source_xml::parse_xml(&properties, false).map_err(ExportError::invalid)?;
        if let Some(language) = get(get(run, "metadata"), "officeFieldResultLanguage").as_str() {
            let mut element = properties.child("lang").cloned().unwrap_or_else(|| source_xml::Element::word("lang"));
            element.set_word_attribute("val", language);
            properties.content.retain(|content| !matches!(content, source_xml::Content::Element(element) if element.is_word("lang")));
            properties.content.push(source_xml::Content::Element(element));
        }
        Ok(properties)
    }

    fn text_run_xml(&mut self, run: &Value, text: &str) -> Result<String, ExportError> {
        let properties = self.run_properties(run)?.render(None);
        let normalized = text.replace("\r\n", "\n").replace('\r', "\n").replace('\u{2028}', "\n");
        let mut content = String::new();
        for (line, text) in normalized.split('\n').enumerate() {
            if line > 0 { content.push_str("<w:br/>"); }
            for (column, text) in text.split('\t').enumerate() {
                if column > 0 { content.push_str("<w:tab/>"); }
                content.push_str(&format!("<w:t xml:space=\"preserve\">{}</w:t>", xml_text(text)));
            }
        }
        Ok(format!("<w:r>{properties}{content}</w:r>"))
    }

    fn runs_with_comments(&mut self, block: &Value, runs: &[Value], preserved: Option<&PreservedParagraph>) -> Result<String, ExportError> {
        let prefix_length = preserved.map_or(0, |paragraph| paragraph.prefix_length);
        let block_id = string(get(block, "id")).unwrap_or_default();
        let annotated = self.text_fields.annotate(&block_id, runs, prefix_length)?;
        let runs = annotated.as_deref().unwrap_or(runs);
        let marks = self.comment_anchors.get(&block_id).cloned().unwrap_or_default();
        let mut controls = self.controls_by_block.get(&block_id).cloned().unwrap_or_default().into_iter()
            .filter(|control| get(&control.value, "placement").as_str() != Some("block") && get(&control.value, "blockRange").is_null()).collect::<Vec<_>>();
        for control in &mut controls {
            control.start = control.start.checked_sub(prefix_length).ok_or_else(|| ExportError::invalid("内容控件与重放编号域重叠。"))?;
            control.end = control.end.checked_sub(prefix_length).ok_or_else(|| ExportError::invalid("内容控件与重放编号域重叠。"))?;
        }
        let body_length = runs.iter().map(|run| string(get(run, "text")).unwrap_or_default().encode_utf16().count()).sum::<usize>();
        let total = prefix_length + body_length;
        let mut events = Vec::<(usize, u8, i64, String)>::new();
        for event in self.text_fields.take_boundaries(&block_id) {
            if event.offset < prefix_length {
                return Err(ExportError::invalid("正文字段结果与原始编号域重叠；不能重复写出两个权威字段。"));
            }
            events.push((event.offset, event.priority, event.order, event.xml));
        }
        for mark in marks {
            if mark.start > mark.end || mark.end > total {
                self.warn("DOCX_COMMENT_RANGE_INVALID", format!("Comment range for comment {} exceeds paragraph {block_id:?} and was omitted.", mark.id));
                continue;
            }
            if mark.starts_here {
                events.push((mark.start, 2, i64::from(mark.id), format!("<w:commentRangeStart w:id=\"{}\"/>", mark.id)));
            }
            if mark.ends_here {
                let collapsed = mark.starts_here && mark.start == mark.end;
                events.push((mark.end, if collapsed { 3 } else { 0 }, i64::from(mark.id), format!("<w:commentRangeEnd w:id=\"{}\"/>", mark.id)));
                events.push((mark.end, if collapsed { 4 } else { 1 }, i64::from(mark.id), format!("<w:r><w:commentReference w:id=\"{}\"/></w:r>", mark.id)));
            }
        }
        if get(get(block, "metadata"), "codeBlock").get("source").is_some() {
            let source = get(get(block, "metadata"), "codeBlock")["source"].as_str().ok_or_else(|| ExportError::invalid("代码块源文本无效。"))?;
            let actual = runs.iter().map(|run| get(run, "text").as_str().unwrap_or_default()).collect::<String>();
            if actual != source { return Err(ExportError::invalid("代码块源文本与格式范围不一致。")); }
            let mut offset = prefix_length;
            for run in runs {
                let length = get(run, "text").as_str().unwrap_or_default().encode_utf16().count();
                if length > 0 && get(get(run, "metadata"), "codeSyntaxHighlight").as_bool() == Some(false) {
                    let id = self.sequences.next_bookmark;
                    self.sequences.next_bookmark = id.checked_add(1).ok_or_else(|| ExportError::invalid("Too many paragraph bookmarks."))?;
                    let name = format!("office_code_plain_v1_{}_{}", offset - prefix_length, offset - prefix_length + length);
                    events.push((offset, 2, i64::from(id), format!("<w:bookmarkStart w:id=\"{id}\" w:name=\"{name}\"/>")));
                    events.push((offset + length, 0, i64::from(id), format!("<w:bookmarkEnd w:id=\"{id}\"/>")));
                }
                offset += length;
            }
        }
        events.sort_by_key(|event| (event.0, event.1, event.2));
        let mut rendered = String::new();
        let prefix_events = if preserved.is_some() { events.partition_point(|event| event.0 <= prefix_length) } else { 0 };
        if let Some(paragraph) = preserved {
            let events = events[..prefix_events].iter().map(|event| (event.0, event.3.clone())).collect::<Vec<_>>();
            rendered.push_str(&paragraph.render(&events)?);
        }
        let mut comments = CommentEventCursor {
            events: events.into_iter().skip(prefix_events).map(|(offset, _, _, marker)| (offset - prefix_length, marker)).collect(),
            next: 0,
        };
        rendered.push_str(&self.render_control_range(runs, &controls, 0, body_length, None, 0, true, &mut comments)?);
        if rendered.is_empty() {
            rendered.push_str("<w:r><w:t></w:t></w:r>");
        }
        Ok(rendered)
    }

    fn render_control_range(
        &mut self,
        runs: &[Value],
        controls: &[ContentControl],
        start: usize,
        end: usize,
        parent_id: Option<&str>,
        depth: usize,
        include_end_objects: bool,
        comments: &mut CommentEventCursor,
    ) -> Result<String, ExportError> {
        if depth > 64 {
            return Err(ExportError::invalid(
                "Content control nesting exceeds the maximum depth of 64.",
            ));
        }
        let mut children = controls
            .iter()
            .filter(|control| {
                let candidate_parent = string(get(&control.value, "parentId"));
                let local_parent = candidate_parent.as_deref().filter(|id| controls.iter().any(|candidate| candidate.id == *id));
                local_parent == parent_id
            })
            .cloned()
            .collect::<Vec<_>>();
        children.sort_by_key(|control| (control.start, control.end, control.id.clone()));
        let mut cursor = start;
        let mut result = String::new();
        for control in children {
            if control.start < start || control.end > end || control.start > control.end {
                self.warn(
                    "DOCX_CONTENT_CONTROL_RANGE",
                    format!("Content control {:?} has a range outside its paragraph or parent; its wrapper was omitted.", control.id),
                );
                continue;
            }
            if control.start < cursor {
                self.warn(
                    "DOCX_CONTENT_CONTROL_OVERLAP",
                    format!(
                        "Content control {:?} overlaps a sibling control and was omitted.",
                        control.id
                    ),
                );
                continue;
            }
            result.push_str(&self.render_annotated_run_range(runs, cursor, control.start, false, comments)?);
            let content = self.render_control_range(
                runs,
                controls,
                control.start,
                control.end,
                Some(&control.id),
                depth + 1,
                false,
                comments,
            )?;
            if let Some(wrapped) = content_control_xml(&control, &content)? {
                if custom_xml::active_binding_store(get(&control.value, "dataBinding"))
                    .is_some_and(|id| !self.custom_xml_stores.contains(&id.to_lowercase())) {
                    self.warn(
                        "DOCX_SDT_DATA_BINDING_UNSUPPORTED",
                        format!("Content control {:?} references a custom XML binding, but no matching custom XML store is present in the document model.", control.id),
                    );
                }
                result.push_str(&control_writer::record(self, &control, wrapped)?);
            } else {
                self.warn(
                    "DOCX_CONTENT_CONTROL_TYPE_UNSUPPORTED",
                    format!("Content control {:?} has unsupported inline type {:?}; its text was preserved without the wrapper.", control.id, control.kind),
                );
                result.push_str(&content);
            }
            cursor = control.end;
        }
        result.push_str(&self.render_annotated_run_range(runs, cursor, end, include_end_objects, comments)?);
        Ok(result)
    }

    fn render_annotated_run_range(
        &mut self,
        runs: &[Value],
        start: usize,
        end: usize,
        include_end_objects: bool,
        comments: &mut CommentEventCursor,
    ) -> Result<String, ExportError> {
        let mut cursor = start;
        let mut result = String::new();
        while let Some((offset, marker)) = comments.events.get(comments.next) {
            // Each boundary belongs to its following interval. The paragraph's
            // final interval also owns ending comments and zero-length objects.
            if *offset > end || *offset == end && !include_end_objects { break; }
            if *offset < cursor {
                return Err(ExportError::invalid("Comment events are not ordered within the content control ranges."));
            }
            result.push_str(&self.render_run_range(runs, cursor, *offset, false)?);
            result.push_str(marker);
            cursor = *offset;
            comments.next += 1;
        }
        result.push_str(&self.render_run_range(runs, cursor, end, include_end_objects)?);
        Ok(result)
    }

    fn render_run_range(
        &mut self,
        runs: &[Value],
        start: usize,
        end: usize,
        include_end_objects: bool,
    ) -> Result<String, ExportError> {
        if start == end && !include_end_objects {
            return Ok(String::new());
        }
        let mut result = Vec::new();
        let mut offset = 0usize;
        for run in runs {
            let text = string(get(run, "text")).unwrap_or_default();
            let length = text.encode_utf16().count();
            let run_end = offset.saturating_add(length);
            let overlap_start = start.max(offset);
            let overlap_end = end.min(run_end);
            // Zero-length objects belong to the following text interval. Only the
            // paragraph's final interval owns objects at its ending boundary.
            if length == 0 && offset >= start && (offset < end || include_end_objects && offset == end) {
                result.push(run.clone());
            } else if overlap_start < overlap_end {
                let Some(part) = slice_utf16(&text, overlap_start - offset, overlap_end - offset)
                else {
                    return Err(ExportError::invalid(
                        "A content control boundary splits a UTF-16 surrogate pair.",
                    ));
                };
                let mut clipped = run.clone();
                if let Some(object) = clipped.as_object_mut() {
                    object.insert("text".to_owned(), Value::String(part.to_owned()));
                    if overlap_start > offset {
                        // A reference attached to text is emitted at its original
                        // run boundary, once, even when comments split the run.
                        object.remove("footnoteId");
                        object.remove("endnoteId");
                        if let Some(metadata) = object.get_mut("metadata").and_then(Value::as_object_mut) {
                            metadata.remove("note");
                            metadata.remove("footnoteId");
                            metadata.remove("endnoteId");
                        }
                    }
                }
                result.push(clipped);
            }
            offset = run_end;
            if offset > end || offset == end && !include_end_objects {
                break;
            }
        }
        if result.is_empty() { Ok(String::new()) } else { self.runs(&result) }
    }

    fn blocks(&mut self, blocks: &[Value], available: Option<i64>) -> Result<String, ExportError> {
        let mut contents = blocks.iter().map(|block| {
            if available.is_some() && get(block, "type").as_str() == Some("sectionBreak") {
                return Err(ExportError::invalid("DOCX section breaks cannot be written inside a table cell"));
            }
            if get(block, "type").as_str() == Some("table") { table_writer::write(self, block, available) }
            else { self.block(block) }
        }).collect::<Result<Vec<_>, _>>()?;
        control_pictures::merge_contents(blocks, &mut contents, &self.controls_by_block)?;
        control_writer::wrap_blocks(self, blocks, &contents)
    }

    fn block(&mut self, block: &Value) -> Result<String, ExportError> {
        match string(get(block, "type")).unwrap_or_default().as_str() {
            "table" => self.table(block),
            "image" => {
                let image = self.image(block)?;
                let contents = image.unwrap_or_else(|| {
                    let metadata = get(block, "metadata");
                    let image = get(metadata, "image");
                    let alt = string(get(image, "alt"))
                        .or_else(|| string(get(metadata, "alt")))
                        .unwrap_or_else(|| "Image unavailable".to_owned());
                    format!(
                        "<w:r><w:t xml:space=\"preserve\">{}</w:t></w:r>",
                        xml_text(&format!("[Image: {alt}]"))
                    )
                });
                Ok(format!("<w:p>{contents}</w:p>"))
            }
            "divider" => Ok(self.paragraph(block, true)?),
            "pageBreak" => Ok("<w:p><w:r><w:br w:type=\"page\"/></w:r></w:p>".to_owned()),
            "columnBreak" => Ok("<w:p><w:r><w:br w:type=\"column\"/></w:r></w:p>".to_owned()),
            "sectionBreak" => {
                if self.relationship_owner != "word/document.xml" { return Err(ExportError::invalid("辅助内容区不能包含正文分节符。")); }
                let value = string(get(get(block, "sectionBreak"), "type"))
                    .unwrap_or_else(|| "nextPage".to_owned());
                if !matches!(
                    value.as_str(),
                    "nextPage" | "continuous" | "nextColumn" | "evenPage" | "oddPage"
                ) {
                    return Err(ExportError::invalid(format!(
                        "Unsupported section break type {value:?}."
                    )));
                }
                let section = self.current_section.clone();
                let refs = self.section_header_footer_refs.get(&section);
                let properties = section_properties_with_headers(self.page, Some(&value), refs);
                self.current_section = string(get(block, "id")).unwrap_or_default();
                Ok(format!("<w:p><w:pPr>{properties}</w:pPr></w:p>"))
            }
            _ => self.paragraph(block, false),
        }
    }

    fn paragraph(&mut self, block: &Value, divider: bool) -> Result<String, ExportError> {
        let id = block as *const Value as usize;
        let preserved = self.preserved_markers.paragraphs.remove(&id);
        let sequence = self.sequences.paragraphs.get(&id).cloned();
        let mut properties = String::new();
        let block_type = string(get(block, "type")).unwrap_or_default();
        let code_block = get(get(block, "metadata"), "codeBlock");
        let is_code_block = string(get(code_block, "source")).is_some();
        if is_code_block {
            properties.push_str("<w:pStyle w:val=\"OfficeCodeBlock\"/><w:keepLines/>");
        } else if block_type == "heading" {
            let level = integer(get(block, "level")).unwrap_or(1).clamp(1, 6);
            properties.push_str(&format!(
                "<w:pStyle w:val=\"Heading{level}\"/><w:keepNext/>"
            ));
        } else if block_type == "quote" {
            properties.push_str("<w:pStyle w:val=\"Quote\"/>");
        } else if block_type == "paragraph" {
            match get(get(block, "metadata"), "namedStyleType").as_str() {
                Some("title") => properties.push_str("<w:pStyle w:val=\"Title\"/>"),
                Some("subtitle") => properties.push_str("<w:pStyle w:val=\"Subtitle\"/>"),
                _ => {}
            }
        }
        if let Some(num_id) = self.list_ids.get(&id).filter(|_| sequence.as_ref().is_none_or(|sequence| !sequence.sequence || sequence.native_marker)) {
            let level = numbering_level(block);
            properties.push_str(&format!(
                "<w:numPr><w:ilvl w:val=\"{level}\"/><w:numId w:val=\"{num_id}\"/></w:numPr>"
            ));
        }

        let style = get(block, "style");
        let spacing_before = twips(nonnegative(get(style, "spaceBefore"), 0.0));
        let spacing_after = twips(nonnegative(get(style, "spaceAfter"), 8.0));
        let line_rule = match get(style, "spacingRule").as_str() {
            Some("exact") => "exact",
            Some("atLeast") => "atLeast",
            _ => "auto",
        };
        let line = (positive(get(style, "lineHeight"), if line_rule == "auto" { 1.5 } else { 20.0 })
            * if line_rule == "auto" { 240.0 } else { TWIPS_PER_PIXEL })
            .round()
            .clamp(1.0, 31_680.0) as u32;
        properties.push_str(&format!("<w:spacing w:before=\"{spacing_before}\" w:after=\"{spacing_after}\" w:line=\"{line}\" w:lineRule=\"{line_rule}\"/>"));
        let inherited_sequence = sequence.as_ref().filter(|sequence| sequence.sequence);
        let left = get(style, "indentStart").as_f64()
            .or_else(|| get(style, "indent").as_f64().filter(|value| *value != 0.0).map(|value| {
                let inherited = inherited_sequence.map(|sequence| sequence.left).unwrap_or_else(|| self.list_ids.get(&id).map(|num_id| numbering_left(&self.numbering, *num_id, numbering_level(block))).unwrap_or(0.0));
                inherited + value * 24.0
            })).or_else(|| inherited_sequence.map(|sequence| sequence.left));
        let right = get(style, "indentEnd").as_f64().or_else(|| inherited_sequence.map(|sequence| sequence.right));
        let first_line = get(style, "indentFirstLine").as_f64().or_else(|| inherited_sequence.map(|sequence| sequence.first_line));
        let hanging = get(style, "hanging").as_f64().or_else(|| inherited_sequence.map(|sequence| sequence.hanging));
        if inherited_sequence.is_some_and(|sequence| sequence.tab && !sequence.native_marker) {
            properties.push_str(&format!("<w:tabs><w:tab w:val=\"left\" w:pos=\"{}\"/></w:tabs>", signed_twips(left.unwrap_or(0.0))));
        }
        if left.is_some() || right.is_some() || first_line.is_some() || hanging.is_some() {
            let mut indent = String::new();
            if let Some(left) = left {
                indent.push_str(&format!(" w:left=\"{}\"", signed_twips(left)));
            }
            if let Some(right) = right {
                indent.push_str(&format!(" w:right=\"{}\"", signed_twips(right)));
            }
            if first_line.is_some() || hanging.is_some() {
                let first = signed_twips(first_line.unwrap_or(0.0) - hanging.unwrap_or(0.0));
                let (property, value) = if first < 0 { ("hanging", -first) } else { ("firstLine", first) };
                indent.push_str(&format!(" w:{property}=\"{value}\""));
            }
            properties.push_str(&format!("<w:ind{indent}/>"));
        }
        for (field, tag) in [("wordWrap", "wordWrap"), ("snapToGrid", "snapToGrid")] {
            if let Some(value) = get(style, field).as_bool() {
                properties.push_str(&format!("<w:{tag} w:val=\"{}\"/>", u8::from(value)));
            }
        }
        let align = match string(get(style, "align")).unwrap_or_default().as_str() {
            "center" => "center",
            "right" => "right",
            "justify" => "both",
            _ => "left",
        };
        properties.push_str(&format!("<w:jc w:val=\"{align}\"/>"));
        if divider {
            properties.push_str("<w:pBdr><w:bottom w:val=\"single\" w:sz=\"6\" w:space=\"1\" w:color=\"AAB3BF\"/></w:pBdr>");
        }
        let paragraph_properties = self.source_styles.block_paragraph_properties(&properties, block).map_err(ExportError::invalid)?;
        let source_runs = field(block, "runs")
            .and_then(Value::as_array)
            .map_or(&[][..], Vec::as_slice);
        let runs = source_runs;
        let body_runs = preserved.as_ref().map(|paragraph| paragraph.body_runs(runs)).transpose()?;
        let runs = body_runs.as_deref().unwrap_or(runs);
        let paragraph = format!(
            "<w:p>{paragraph_properties}{}{}</w:p>",
            sequence.as_ref().map(|sequence| sequence.prefix.as_str()).unwrap_or_default(),
            self.runs_with_comments(block, runs, preserved.as_ref())?
        );
        if !is_code_block {
            return Ok(paragraph);
        }
        let language = string(get(code_block, "language")).unwrap_or_else(|| "plain".to_owned());
        let line_numbers = get(code_block, "lineNumbers").as_bool() == Some(true);
        let bookmark_name = xml_attribute(&format!(
            "office_code_v1_{language}_{}",
            u8::from(line_numbers)
        ));
        let bookmark_id = self.sequences.next_bookmark;
        self.sequences.next_bookmark = self.sequences.next_bookmark.checked_add(1).ok_or_else(|| ExportError::invalid("Too many paragraph bookmarks."))?;
        let bookmark_start =
            format!("<w:bookmarkStart w:id=\"{bookmark_id}\" w:name=\"{bookmark_name}\"/>");
        let bookmark_end = format!("<w:bookmarkEnd w:id=\"{bookmark_id}\"/>");
        let paragraph = if paragraph.contains("</w:pPr>") {
            paragraph.replacen("</w:pPr>", &format!("</w:pPr>{bookmark_start}"), 1)
        } else {
            paragraph.replacen("<w:p>", &format!("<w:p>{bookmark_start}"), 1)
        };
        Ok(paragraph.replacen("</w:p>", &format!("{bookmark_end}</w:p>"), 1))
    }

    fn table(&mut self, block: &Value) -> Result<String, ExportError> {
        table_writer::write(self, block, None)
    }
}

fn cell_blocks(cell: &Value) -> &[Value] {
    field(cell, "blocks").and_then(Value::as_array).map_or(&[][..], Vec::as_slice)
}

fn note_reference(run: &Value) -> Option<(NoteKind, String)> {
    let metadata = get(run, "metadata");
    for kind in [NoteKind::Footnote, NoteKind::Endnote] {
        let source = string(get(run, kind.field()))
            .or_else(|| get(run, kind.field()).as_i64().map(|id| id.to_string()))
            .or_else(|| string(get(metadata, kind.field())))
            .or_else(|| {
                get(metadata, kind.field())
                    .as_i64()
                    .map(|id| id.to_string())
            });
        if let Some(source) = source.filter(|source| !source.trim().is_empty()) {
            return Some((kind, source));
        }
    }
    for note in [get(metadata, "note"), get(metadata, "noteReference")] {
        let kind = match get(note, "type").as_str().or_else(|| get(note, "kind").as_str()) {
            Some("footnote") => NoteKind::Footnote,
            Some("endnote") => NoteKind::Endnote,
            _ => continue,
        };
        let id = string(get(note, "noteId")).or_else(|| get(note, "noteId").as_i64().map(|id| id.to_string()))
            .or_else(|| string(get(note, "id"))).or_else(|| get(note, "id").as_i64().map(|id| id.to_string()));
        if let Some(id) = id.filter(|id| !id.trim().is_empty()) { return Some((kind, id)); }
    }
    None
}

fn note_sources<'a>(model: &'a Value, kind: NoteKind) -> Option<&'a Value> {
    let collection = kind.collection();
    field(model, collection)
        .or_else(|| field(get(model, "metadata"), collection))
        .or_else(|| field(get(model, "notes"), collection))
}

fn collect_note_entries(model: &Value, kind: NoteKind) -> Vec<(String, Value)> {
    let Some(source) = note_sources(model, kind) else {
        return Vec::new();
    };
    match source {
        Value::Array(entries) => entries
            .iter()
            .enumerate()
            .filter_map(|(index, entry)| {
                let id = string(get(entry, "noteId"))
                    .or_else(|| string(get(entry, "id")))
                    .or_else(|| get(entry, "noteId").as_i64().map(|id| id.to_string()))
                    .or_else(|| get(entry, "id").as_i64().map(|id| id.to_string()))
                    .unwrap_or_else(|| (index + 1).to_string());
                (!id.trim().is_empty()).then(|| (id, entry.clone()))
            })
            .collect(),
        Value::Object(entries) => entries
            .iter()
            .map(|(id, entry)| (id.clone(), entry.clone()))
            .collect(),
        _ => Vec::new(),
    }
}

fn collect_references(blocks: &[Value], found: &mut BTreeSet<(NoteKind, String)>) -> Result<(), ExportError> {
    for block in blocks {
        if let Some(runs) = field(block, "runs").and_then(Value::as_array) {
            for run in runs {
                if let Some(reference) = note_reference(run) {
                    found.insert(reference);
                }
            }
        }
        if get(block, "type").as_str() == Some("table") {
            for group in table_writer::content_groups(get(get(block, "metadata"), "table"))? { collect_references(group.blocks, found)?; }
        }
    }
    Ok(())
}

fn collect_notes(model: &Value, blocks: &[Value]) -> Result<NoteBundle, ExportError> {
    let mut notes: [Vec<NoteEntry>; 2] = [Vec::new(), Vec::new()];
    let mut ids = HashMap::new();
    let mut next_id = [1i64, 1i64];
    let mut used_ids = [note_writer::reserved_ids(&model, NoteKind::Footnote)?, note_writer::reserved_ids(&model, NoteKind::Endnote)?];
    let mut reference_properties = HashMap::new();
    for kind in [NoteKind::Footnote, NoteKind::Endnote] {
        let index = if kind == NoteKind::Endnote { 1 } else { 0 };
        for (source_id, value) in collect_note_entries(model, kind) {
            if ids.contains_key(&(kind, source_id.clone())) {
                continue;
            }
            let preferred = get(get(get(&value, "metadata"), "officeNoteSource"), "nativeId").as_i64()
                .filter(|id| *id > 0 && id.unsigned_abs() <= office_io_common::document_note_source::NOTE_SAFE_INTEGER && !used_ids[index].contains(id));
            let id = if let Some(id) = preferred { used_ids[index].insert(id); id }
                else { note_writer::allocate_id(&mut next_id[index], &mut used_ids[index])? };
            ids.insert((kind, source_id.clone()), id);
            reference_properties.insert((kind, source_id), note_writer::ReferenceProperties::read(&value)?);
            notes[index].push(NoteEntry { id, value });
        }
    }

    let mut bundle = NoteBundle { footnotes: std::mem::take(&mut notes[0]), endnotes: std::mem::take(&mut notes[1]), ids, orphaned: Vec::new(), reference_properties };
    let mut references = BTreeSet::new();
    collect_references(blocks, &mut references)?;
    for host in block_hosts::auxiliary_contents(model, blocks, &bundle)? { collect_references(host.blocks, &mut references)?; }
    for (kind, source_id) in references {
        if bundle.ids.contains_key(&(kind, source_id.clone())) {
            continue;
        }
        let index = if kind == NoteKind::Endnote { 1 } else { 0 };
        let id = note_writer::allocate_id(&mut next_id[index], &mut used_ids[index])?;
        bundle.ids.insert((kind, source_id.clone()), id);
        let entries = if kind == NoteKind::Footnote { &mut bundle.footnotes } else { &mut bundle.endnotes };
        entries.push(NoteEntry {
            id,
            value: Value::Null,
        });
        bundle.orphaned.push((kind, source_id));
    }
    Ok(bundle)
}

fn collect_content_controls(
    model: &Value,
    blocks: &[Value],
) -> Result<HashMap<String, Vec<ContentControl>>, ExportError> {
    collect_content_controls_for_groups(model, &[blocks])
}

fn collect_content_controls_for_groups(model: &Value, groups: &[&[Value]]) -> Result<HashMap<String, Vec<ContentControl>>, ExportError> {
    let Some(items) = field(model, "contentControls") else {
        return Ok(HashMap::new());
    };
    let items = items.as_array().ok_or_else(|| ExportError::invalid("DOCX content controls must be an array"))?;
    if items.len() > office_io_common::document_content_controls::MAX_CONTROL_COUNT { return Err(ExportError::invalid("DOCX_SDT_RESOURCE_LIMIT: too many content controls")); }
    let mut block_ids = BTreeSet::new();
    for blocks in groups { collect_block_ids(blocks, &mut block_ids)?; }
    let mut controls = HashMap::<String, Vec<ContentControl>>::new();
    let mut identities = BTreeSet::new();
    for (index, value) in items.iter().enumerate() {
        let id = string(get(value, "id")).unwrap_or_default();
        let block_id = string(get(value, "blockId")).unwrap_or_default();
        let kind = string(get(value, "type")).unwrap_or_default();
        let start = get(value, "startOffset")
            .as_u64()
            .and_then(|value| usize::try_from(value).ok());
        let end = get(value, "endOffset")
            .as_u64()
            .and_then(|value| usize::try_from(value).ok());
        let (Some(start), Some(end)) = (start, end) else {
            return Err(ExportError::invalid("DOCX content control requires nonnegative UTF-16 offsets"));
        };
        if id.trim().is_empty() || id.len() > 4096 || !identities.insert(id.clone()) || block_id.trim().is_empty() || !block_ids.contains(&block_id) || start > end {
            return Err(ExportError::invalid("DOCX content control has an invalid/duplicate identity or input host"));
        }
        let numeric_id = u32::try_from(index + 1).map_err(|_| ExportError::invalid("DOCX native content-control ID capacity exceeded"))?;
        let key = if get(value, "placement").as_str() == Some("row") {
            if !get(value, "blockRange").is_null() { return Err(ExportError::invalid("DOCX row SDT cannot also own a block range")); }
            let table_id = get(get(value, "tableRowRange"), "tableId").as_str().filter(|id| block_ids.contains(*id))
                .ok_or_else(|| ExportError::invalid("DOCX row SDT requires an existing native table"))?;
            table_id.to_owned()
        } else if !get(value, "blockRange").is_null() {
            if get(value, "placement").as_str() != Some("block") || !get(value, "tableRowRange").is_null() { return Err(ExportError::invalid("DOCX block SDT has conflicting structural ownership")); }
            get(get(value, "blockRange"), "blockIds").as_array().and_then(|ids| ids.first()).and_then(Value::as_str)
                .filter(|id| block_ids.contains(*id)).ok_or_else(|| ExportError::invalid("DOCX block SDT requires an existing first block"))?.to_owned()
        } else { block_id };
        controls.entry(key).or_default().push(ContentControl {
            id,
            kind,
            numeric_id,
            start,
            end,
            value: value.clone(),
        });
    }
    let by_id = controls.values().flatten().map(|control| (control.id.as_str(), control)).collect::<HashMap<_, _>>();
    for control in by_id.values() {
        let mut parent = get(&control.value, "parentId").as_str();
        let mut chain = HashSet::from([control.id.as_str()]);
        while let Some(id) = parent {
            if !chain.insert(id) { return Err(ExportError::invalid("DOCX content controls contain a parent cycle")); }
            let ancestor = by_id.get(id).ok_or_else(|| ExportError::invalid("DOCX content control has a missing parent"))?;
            parent = get(&ancestor.value, "parentId").as_str();
        }
    }
    for block_controls in controls.values_mut() {
        block_controls.sort_by_key(|control| {
            (
                control.start,
                std::cmp::Reverse(control.end),
                control.id.clone(),
            )
        });
    }
    control_ranges::validate(groups, &controls)?;
    Ok(controls)
}

fn collect_block_ids(blocks: &[Value], found: &mut BTreeSet<String>) -> Result<(), ExportError> {
    for block in blocks {
        if let Some(id) = string(get(block, "id")).filter(|id| !id.trim().is_empty()) {
            found.insert(id);
        }
        if get(block, "type").as_str() == Some("table") {
            for group in table_writer::content_groups(get(get(block, "metadata"), "table"))? { collect_block_ids(group.blocks, found)?; }
        }
    }
    Ok(())
}

fn slice_utf16(text: &str, start: usize, end: usize) -> Option<&str> {
    if start > end {
        return None;
    }
    let byte_at = |offset: usize| {
        if offset == 0 {
            return Some(0);
        }
        let mut units = 0usize;
        for (byte, character) in text.char_indices() {
            if units == offset {
                return Some(byte);
            }
            units = units.checked_add(character.len_utf16())?;
            if units > offset {
                return None;
            }
        }
        (units == offset).then_some(text.len())
    };
    Some(&text[byte_at(start)?..byte_at(end)?])
}

fn content_control_xml(control: &ContentControl, content: &str) -> Result<Option<String>, ExportError> {
    let value = &control.value;
    let title = string(get(value, "title")).or_else(|| string(get(value, "alias")));
    let tag = string(get(value, "tag"));
    let placeholder = string(get(value, "placeholder"));
    let lock = match string(get(value, "lock")).as_deref() {
        Some("controlLocked" | "sdtLocked") => Some("sdtLocked"),
        Some("contentLocked") => Some("contentLocked"),
        Some("sdtContentLocked") => Some("sdtContentLocked"),
        _ => None,
    };
    let mut properties = format!("<w:id w:val=\"{}\"/>", control.numeric_id);
    if let Some(title) = title {
        properties.push_str(&format!("<w:alias w:val=\"{}\"/>", xml_attribute(&title)));
    }
    if let Some(tag) = tag {
        properties.push_str(&format!("<w:tag w:val=\"{}\"/>", xml_attribute(&tag)));
    }
    if let Some(lock) = lock {
        properties.push_str(&format!("<w:lock w:val=\"{lock}\"/>"));
    }
    let binding = get(value, "dataBinding");
    if let (Some(xpath), Some(store)) = (get(binding, "xpath").as_str().filter(|value| !value.is_empty()), custom_xml::active_binding_store(binding)) {
        properties.push_str(&format!("<w:dataBinding{}{}{} />", xml::attr("w:xpath", xpath), xml::attr("w:storeItemID", store),
            get(binding, "prefixMappings").as_str().map(|value| xml::attr("w:prefixMappings", value)).unwrap_or_default()));
    }
    match control.kind.as_str() {
        "plainText" | "text" => {
            if get(value, "multiLine").as_bool() == Some(true)
                || get(get(value, "text"), "multiLine").as_bool() == Some(true)
            {
                properties.push_str("<w:text w:multiLine=\"1\"/>");
            } else {
                properties.push_str("<w:text/>");
            }
        }
        "richText" => {}
        "checkbox" => {
            let checked = get(get(value, "checkbox"), "checked")
                .as_bool()
                .or_else(|| get(value, "checked").as_bool())
                .unwrap_or(false);
            let checked_value = string(get(get(get(value, "checkbox"), "checkedState"), "value"))
                .unwrap_or_else(|| "2612".to_owned());
            let checked_font = string(get(get(get(value, "checkbox"), "checkedState"), "font"))
                .unwrap_or_else(|| "MS Gothic".to_owned());
            let unchecked_value =
                string(get(get(get(value, "checkbox"), "uncheckedState"), "value"))
                    .unwrap_or_else(|| "2610".to_owned());
            let unchecked_font = string(get(get(get(value, "checkbox"), "uncheckedState"), "font"))
                .unwrap_or_else(|| "MS Gothic".to_owned());
            properties.push_str(&format!(
                "<w14:checkbox><w14:checked w14:val=\"{}\"/><w14:checkedState w14:val=\"{}\" w14:font=\"{}\"/><w14:uncheckedState w14:val=\"{}\" w14:font=\"{}\"/></w14:checkbox>",
                u8::from(checked),
                xml_attribute(&checked_value),
                xml_attribute(&checked_font),
                xml_attribute(&unchecked_value),
                xml_attribute(&unchecked_font)
            ));
        }
        "date" => {
            let date_settings = get(value, "date");
            let date = string(get(value, "dateValue"))
                .or_else(|| string(get(date_settings, "fullDate")))
                .or_else(|| string(get(value, "value")));
            if let Some(date) = date.filter(|date| valid_iso_date(date) || valid_iso_datetime(date))
            {
                let full_date = if valid_iso_date(&date) {
                    format!("{date}T00:00:00Z")
                } else {
                    date.clone()
                };
                let format = string(get(value, "dateFormat"))
                    .or_else(|| string(get(date_settings, "format")))
                    .unwrap_or_else(|| "yyyy-MM-dd".to_owned());
                let locale = string(get(value, "dateLocale"))
                    .or_else(|| string(get(date_settings, "locale")))
                    .unwrap_or_else(|| "en-US".to_owned());
                let calendar = string(get(value, "dateCalendar"))
                    .or_else(|| string(get(date_settings, "calendar")));
                let storage = string(get(value, "dateStorage"))
                    .or_else(|| string(get(date_settings, "storage")))
                    .unwrap_or_else(|| "dateTime".to_owned());
                properties.push_str(&format!(
                    "<w:date w:fullDate=\"{}\"><w:dateFormat w:val=\"{}\"/><w:lid w:val=\"{}\"/>{}<w:storeMappedDataAs w:val=\"{}\"/></w:date>",
                    xml_attribute(&full_date),
                    xml_attribute(&format),
                    xml_attribute(&locale),
                    calendar.map(|value| format!("<w:calendar w:val=\"{}\"/>", xml_attribute(&value))).unwrap_or_default(),
                    xml_attribute(&storage)
                ));
            } else {
                properties.push_str(
                    "<w:date><w:dateFormat w:val=\"yyyy-MM-dd\"/><w:lid w:val=\"en-US\"/></w:date>",
                );
            }
        }
        "dropdown" | "dropDownList" | "comboBox" => {
            let tag = if control.kind == "dropdown" || control.kind == "dropDownList" {
                "dropDownList"
            } else {
                "comboBox"
            };
            let mut items = String::new();
            if let Some(values) = field(value, "items")
                .or_else(|| field(value, "listItems"))
                .and_then(Value::as_array)
            {
                for item in values {
                    let item_value = string(get(item, "value")).unwrap_or_default();
                    let display = string(get(item, "displayText"))
                        .or_else(|| string(get(item, "display")))
                        .unwrap_or_else(|| item_value.clone());
                    items.push_str(&format!(
                        "<w:listItem w:displayText=\"{}\" w:value=\"{}\"/>",
                        xml_attribute(&display),
                        xml_attribute(&item_value)
                    ));
                }
            }
            properties.push_str(&format!("<w:{tag}>{items}</w:{tag}>"));
        }
        "picture" => properties.push_str("<w:picture/>"),
        "group" => properties.push_str("<w:group/>"),
        "repeatingSection" => {
            let section = get(value, "repeatingSection");
            let title = string(get(section, "sectionTitle"));
            let lock_insert = get(section, "doNotAllowInsertDelete").as_bool() == Some(true);
            properties.push_str(&format!(
                "<w15:repeatingSection>{}{}</w15:repeatingSection>",
                title
                    .map(|value| format!("<w15:sectionTitle w15:val=\"{}\"/>", xml_attribute(&value)))
                    .unwrap_or_default(),
                if lock_insert {
                    "<w15:doNotAllowInsertDelete/>"
                } else {
                    ""
                }
            ));
        }
        "repeatingSectionItem" => properties.push_str("<w15:repeatingSectionItem/>"),
        "docPartObj" => {
            let doc_part = get(value, "docPart");
            let gallery = string(get(doc_part, "gallery"));
            let category = string(get(doc_part, "category"));
            let unique = get(doc_part, "unique").as_bool() == Some(true);
            properties.push_str(&format!(
                "<w:docPartObj>{}{}{}</w:docPartObj>",
                gallery
                    .map(|value| format!("<w:docPartGallery w:val=\"{}\"/>", xml_attribute(&value)))
                    .unwrap_or_default(),
                category
                    .map(|value| format!(
                        "<w:docPartCategory w:val=\"{}\"/>",
                        xml_attribute(&value)
                    ))
                    .unwrap_or_default(),
                if unique { "<w:docPartUnique/>" } else { "" }
            ));
        }
        _ => return Ok(None),
    }
    if let Some(appearance) = string(get(value, "appearance")) {
        properties.push_str(&format!(
            "<w15:appearance w15:val=\"{}\"/>",
            xml_attribute(&appearance)
        ));
    }
    if let Some(color) = string(get(value, "color")).and_then(color) {
        properties.push_str(&format!("<w15:color w15:val=\"{color}\"/>"));
    }
    if get(value, "temporary").as_bool() == Some(true) {
        properties.push_str("<w:temporary/>");
    }
    if get(value, "showingPlaceholder").as_bool() == Some(true) {
        properties.push_str("<w:showingPlcHdr/>");
    }
    let placeholder_doc_part = get(value, "placeholderDocPart");
    if let Some(name) = string(get(placeholder_doc_part, "name")) {
        properties.push_str(&format!(
            "<w:placeholder><w:docPart w:val=\"{}\"/></w:placeholder>",
            xml_attribute(&name)
        ));
    }
    let content = if content.is_empty() {
        placeholder
            .map(|text| {
                format!(
                    "<w:r><w:t xml:space=\"preserve\">{}</w:t></w:r>",
                    xml_text(&text)
                )
            })
            .unwrap_or_else(|| "<w:r><w:t></w:t></w:r>".to_owned())
    } else {
        content.to_owned()
    };
    let (properties, end_properties) = control_writer::merge_properties(control, &properties)?;
    Ok(Some(format!("<w:sdt {CONTENT_NAMESPACES}>{properties}{end_properties}<w:sdtContent>{content}</w:sdtContent></w:sdt>")))
}

fn valid_iso_date(value: &str) -> bool {
    let parts = value.split('-').collect::<Vec<_>>();
    if parts.len() != 3
        || parts
            .iter()
            .any(|part| part.is_empty() || !part.bytes().all(|byte| byte.is_ascii_digit()))
    {
        return false;
    }
    let (Ok(year), Ok(month), Ok(day)) = (
        parts[0].parse::<u32>(),
        parts[1].parse::<u32>(),
        parts[2].parse::<u32>(),
    ) else {
        return false;
    };
    if parts[0].len() != 4
        || parts[1].len() != 2
        || parts[2].len() != 2
        || !(1..=12).contains(&month)
    {
        return false;
    }
    let leap = year % 4 == 0 && (year % 100 != 0 || year % 400 == 0);
    let days = match month {
        2 if leap => 29,
        2 => 28,
        4 | 6 | 9 | 11 => 30,
        _ => 31,
    };
    (1..=days).contains(&day)
}

fn valid_iso_datetime(value: &str) -> bool {
    value.len() >= 19
        && value.as_bytes().get(10) == Some(&b'T')
        && value.get(..10).is_some_and(valid_iso_date)
        && value.chars().skip(11).all(|character| {
            character.is_ascii_digit() || matches!(character, ':' | '.' | '+' | '-' | 'Z')
        })
        && value
            .chars()
            .any(|character| character == 'Z' || character == '+' || character == '-')
}

fn note_text(value: &Value) -> Option<String> {
    if let Some(text) = value.as_str() {
        return Some(text.to_owned());
    }
    for candidate in [
        get(value, "text"),
        get(value, "content"),
        get(value, "body"),
    ] {
        if let Some(text) = candidate.as_str() {
            return Some(text.to_owned());
        }
        if let Some(text) = string(get(candidate, "text")) {
            return Some(text);
        }
        if let Some(stream) = string(get(candidate, "dataStream")) {
            return Some(stream.trim_end_matches(['\r', '\n']).to_owned());
        }
    }
    None
}

fn collect_comments(
    model: &Value,
) -> Result<(Vec<CommentEntry>, HashMap<String, Vec<CommentAnchor>>, Vec<ExportWarning>), ExportError> {
    let mut entries = Vec::new();
    let mut anchors = HashMap::<String, Vec<CommentAnchor>>::new();
    let mut warnings = Vec::new();
    let Some(comments) = field(model, "comments").and_then(Value::as_object) else {
        return Ok((entries, anchors, warnings));
    };
    let mut next_id = 1u32;
    for comment in comments.values() {
        let messages = field(comment, "messages").and_then(Value::as_array).map_or(&[][..], Vec::as_slice);
        if messages.is_empty() {
            warnings.push(ExportWarning::new("DOCX_COMMENT_EMPTY", "A comment thread without messages was not written."));
            continue;
        }
        let root_id = next_id;
        let root_para_id = format!("{root_id:08X}");
        for (message_index, message) in messages.iter().enumerate() {
            let id = next_id;
            next_id = next_id.checked_add(1).ok_or_else(|| ExportError::invalid("The DOCX comment count exceeds the supported ID range."))?;
            let paragraph_id = format!("{id:08X}");
            let author = string(get(message, "authorName")).unwrap_or_else(|| "Unknown author".to_owned());
            let date = string(get(message, "createdAt")).unwrap_or_else(|| "1970-01-01T00:00:00Z".to_owned());
            let text = string(get(message, "text")).unwrap_or_default();
            entries.push(CommentEntry {
                id,
                paragraph_id: paragraph_id.clone(),
                parent_paragraph_id: (message_index > 0).then(|| root_para_id.clone()),
                author,
                date,
                text,
                resolved: message_index == 0 && get(comment, "resolved").as_bool() == Some(true),
            });
            if message_index == 0 {
                let ranges = field(comment, "anchor").and_then(Value::as_array).map_or(&[][..], Vec::as_slice);
                if ranges.is_empty() {
                    warnings.push(ExportWarning::new("DOCX_COMMENT_UNANCHORED", "A thread comment without a text anchor was written without a DOCX range marker."));
                }
                for (range_index, range) in ranges.iter().enumerate() {
                    let block_id = string(get(range, "blockId")).unwrap_or_default();
                    let Some(start) = integer(get(range, "startOffset")).and_then(|value| usize::try_from(value).ok()) else {
                        warnings.push(ExportWarning::new("DOCX_COMMENT_RANGE_INVALID", "A comment range with an invalid start offset was omitted."));
                        continue;
                    };
                    let Some(end) = integer(get(range, "endOffset")).and_then(|value| usize::try_from(value).ok()) else {
                        warnings.push(ExportWarning::new("DOCX_COMMENT_RANGE_INVALID", "A comment range with an invalid end offset was omitted."));
                        continue;
                    };
                    if block_id.is_empty() || start > end {
                        warnings.push(ExportWarning::new("DOCX_COMMENT_RANGE_INVALID", "A comment range without a block ID or with reversed offsets was omitted."));
                        continue;
                    }
                    anchors.entry(block_id).or_default().push(CommentAnchor {
                        id: root_id,
                        start,
                        end,
                        starts_here: range_index == 0,
                        ends_here: range_index + 1 == ranges.len(),
                    });
                }
            }
        }
    }
    Ok((entries, anchors, warnings))
}

fn comments_xml(comments: &[CommentEntry]) -> String {
    let mut body = String::new();
    for comment in comments {
        body.push_str(&format!(
            "<w:comment w:id=\"{}\" w:author=\"{}\" w:date=\"{}\"><w:p w14:paraId=\"{}\"><w:r><w:t xml:space=\"preserve\">{}</w:t></w:r></w:p></w:comment>",
            comment.id,
            xml_attribute(&comment.author),
            xml_attribute(&comment.date),
            comment.paragraph_id,
            xml_text(&comment.text)
        ));
    }
    format!("{}<w:comments xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\">{body}</w:comments>", xml::DECL)
}

fn comments_extended_xml(comments: &[CommentEntry]) -> String {
    let entries = comments.iter().map(|comment| {
        let parent = comment.parent_paragraph_id.as_ref().map_or(String::new(), |id| format!(" w15:paraIdParent=\"{id}\""));
        let resolved = if comment.resolved { "1" } else { "0" };
        format!("<w15:commentEx w15:paraId=\"{}\"{parent} w15:done=\"{resolved}\"/>", comment.paragraph_id)
    }).collect::<String>();
    format!("{}<w15:commentsEx xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\">{entries}</w15:commentsEx>", xml::DECL)
}

pub(crate) fn write(data: &str, package: &mut Package) -> Result<Vec<ExportWarning>, ExportError> {
    let model: Value = serde_json::from_str(data)
        .map_err(|error| ExportError::invalid(format!("Invalid document JSON: {error}")))?;
    if string(get(&model, "kind")).as_deref() != Some("document") {
        return Err(ExportError::invalid(
            "The DOCX writer requires a document model.",
        ));
    }
    let blocks = field(&model, "blocks")
        .and_then(Value::as_array)
        .ok_or_else(|| ExportError::invalid("The document model is missing its blocks array."))?;
    if !get(get(get(&model, "metadata"), "source"), "nativeFieldSource").is_null() {
        return Err(ExportError::invalid("正文包含尚未映射的原生字段段落；完整来源已保留，需补齐范围映射后才能写出 DOCX。"));
    }
    let native_controls = get(get(&model, "metadata"), "nativeContentControlSource");
    if !native_controls.is_null() {
        let unmapped = get(native_controls, "unmapped").as_array().ok_or_else(|| ExportError::invalid("DOCX native content-control provenance has an invalid retained range list"))?;
        if !unmapped.is_empty() { return Err(ExportError::invalid("DOCX_SDT_UNMAPPED: native content-control structure is retained; complete its editable range mapping before DOCX export")); }
    }
    let legacy_source = legacy_source::prepare(&model)?;
    let custom_xml = custom_xml::prepare(&model)?;
    let generated_styles = default_style::styles_xml(&model, styles_xml())?;
    let run_defaults = RunDefaultEffects::for_model(&model);
    let notes = collect_notes(&model, blocks)?;
    let auxiliary = block_hosts::auxiliary_contents(&model, blocks, &notes)?;
    let mut block_groups = vec![blocks.as_slice()];
    block_groups.extend(auxiliary.iter().map(|host| host.blocks));
    block_hosts::validate_ownership(&block_groups)?;
    for blocks in &block_groups { table_writer::validate_document_tables(blocks)?; }
    let (comment_entries, comment_anchors, comment_warnings) = collect_comments(&model)?;
    let controls_by_block = collect_content_controls_for_groups(&model, &block_groups)?;
    let declared_control_count = field(&model, "contentControls")
        .and_then(Value::as_array)
        .map_or(0, Vec::len);
    let page = page_geometry(get(&model, "page"))?;
    let section_header_footers = get(&model, "sectionHeaderFooters");
    let title = string(get(&model, "name")).unwrap_or_else(|| "Document".to_owned());
    let core = core_properties(&model, &title);
    package.add_relationship(
        None,
        "rId1",
        REL_OFFICE_DOCUMENT,
        "word/document.xml",
        false,
    )?;
    package.add_relationship(
        None,
        "rId2",
        REL_CORE_PROPERTIES,
        "docProps/core.xml",
        false,
    )?;
    package.add_xml(
        "docProps/core.xml",
        "application/vnd.openxmlformats-package.core-properties+xml",
        core,
    )?;
    package.add_relationship(
        Some("word/document.xml"),
        "rId1",
        REL_STYLES,
        "styles.xml",
        false,
    )?;
    package.add_relationship(
        Some("word/document.xml"),
        "rId2",
        REL_NUMBERING,
        "numbering.xml",
        false,
    )?;

    let mut context = WriteContext {
        package,
        warnings: Vec::new(),
        list_ids: HashMap::new(),
        numbering: Vec::new(),
        sequences: SequencePlan { next_bookmark: 1, ..Default::default() },
        preserved_markers: PreservedMarkerPlan::default(),
        note_ids: notes.ids.clone(),
        note_reference_properties: notes.reference_properties.clone(),
        comment_entries,
        comment_anchors,
        controls_by_block,
        control_sources: control_writer::ControlWriteState::default(),
        used_controls: BTreeSet::new(),
        next_relationship: 3,
        relationship_owner: "word/document.xml".to_owned(),
        next_image: 1,
        page,
        section_header_footer_refs: HashMap::new(),
        current_section: "first".to_owned(),
        run_defaults,
        story_defaults: story_writer::default_style(&model),
        next_story_control: 1_073_741_824,
        source_styles: source_styles::SourceStyles::new(&model, generated_styles)?,
        text_fields: TextFieldPlan::default(),
        tables: table_writer::TableWriteState::default(),
        custom_xml_stores: custom_xml.stores.clone(),
    };
    custom_xml.write(&mut context)?;
    if let Some(bytes) = legacy_source {
        use office_io_common::document_legacy_source::{LEGACY_SOURCE_CONTENT_TYPE, LEGACY_SOURCE_PART, LEGACY_SOURCE_RELATIONSHIP, LEGACY_SOURCE_TARGET};
        context.package.add_binary(LEGACY_SOURCE_PART, LEGACY_SOURCE_CONTENT_TYPE, bytes)?;
        context.relationship(LEGACY_SOURCE_RELATIONSHIP, LEGACY_SOURCE_TARGET, false)?;
    }
    let (list_ids, numbering) = assign_list_groups(&block_groups)?;
    context.list_ids = list_ids;
    context.numbering = numbering;
    let numbering_xml = write_numbering_xml(&context.numbering, context.package, &mut context.warnings)?;
    // Source stories may import numbered paragraph styles. Reserve all generated
    // list/picture IDs before any source consumer is allowed to allocate IDs.
    context.source_styles.reserve_numbering(&numbering_xml).map_err(ExportError::invalid)?;
    block_hosts::append_locks(&mut context, &auxiliary)?;
    context.preserved_markers = prepare_preserved_marker_groups(&block_groups, &context.controls_by_block, &mut context.source_styles)?;
    context.warnings.append(&mut context.preserved_markers.warnings);
    context.sequences = prepare_sequence_groups(&block_groups, &context.list_ids, &context.numbering, &numbering_xml, &context.preserved_markers)?;
    context.warnings.append(&mut context.sequences.warnings);
    context.text_fields = TextFieldPlan::prepare_groups(&mut context, &block_groups)?;
    let has_text_fields = context.text_fields.has_fields;
    let has_different_odd_even =
        prepare_section_header_footers(&mut context, section_header_footers, get(&model, "sectionSetup"), get(&model, "noteSettings"), blocks)?;
    let mut body = context.blocks(blocks, None)?;
    if let Some(mapping) = context.sequences.mapping_xml.take() {
        context.package.add_xml("customXml/office-numbering.xml", "application/xml", mapping)?;
        context.relationship("http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml", "../customXml/office-numbering.xml", false)?;
        context.package.add_xml(
            "customXml/office-numbering-properties.xml",
            "application/vnd.openxmlformats-officedocument.customXmlProperties+xml",
            format!("{}<ds:datastoreItem ds:itemID=\"{{8C2185D8-EB78-4A0E-8981-03B03F7D7A91}}\" xmlns:ds=\"http://schemas.openxmlformats.org/officeDocument/2006/customXml\"><ds:schemaRefs><ds:schemaRef ds:uri=\"urn:office-studio:numbering:1\"/></ds:schemaRefs></ds:datastoreItem>", xml::DECL),
        )?;
        context.package.add_relationship(
            Some("customXml/office-numbering.xml"), "rId1",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps",
            "office-numbering-properties.xml", false,
        )?;
    }
    context.warnings.extend(comment_warnings);
    for (kind, source_id) in &notes.orphaned {
        context.warn(
            "DOCX_NOTE_MISSING",
            format!(
                "A {0} reference to note {source_id:?} has no matching note body; an empty note was written to keep the reference valid.",
                kind.collection().trim_end_matches('s')
            ),
        );
    }
    if !notes.footnotes.is_empty() || !get(get(&model, "noteSettings"), "footnote").is_null() {
        let xml = note_writer::notes_xml(&mut context, NoteKind::Footnote, &notes.footnotes, get(get(&model, "noteSettings"), "footnote"))?;
        context.package.add_xml(
            "word/footnotes.xml",
            FOOTNOTES_CONTENT_TYPE,
            xml,
        )?;
        context.relationship(REL_FOOTNOTES, "footnotes.xml", false)?;
    }
    if !notes.endnotes.is_empty() || !get(get(&model, "noteSettings"), "endnote").is_null() {
        let xml = note_writer::notes_xml(&mut context, NoteKind::Endnote, &notes.endnotes, get(get(&model, "noteSettings"), "endnote"))?;
        context.package.add_xml(
            "word/endnotes.xml",
            ENDNOTES_CONTENT_TYPE,
            xml,
        )?;
        context.relationship(REL_ENDNOTES, "endnotes.xml", false)?;
    }
    if !context.comment_entries.is_empty() {
        let entries = context.comment_entries.clone();
        context.package.add_xml(
            "word/comments.xml",
            COMMENTS_CONTENT_TYPE,
            comments_xml(&entries),
        )?;
        context.package.add_xml(
            "word/commentsExtended.xml",
            COMMENTS_EXTENDED_CONTENT_TYPE,
            comments_extended_xml(&entries),
        )?;
        context.relationship(REL_COMMENTS, "comments.xml", false)?;
        context.relationship(REL_COMMENTS_EXTENDED, "commentsExtended.xml", false)?;
    }
    table_writer::write_sources(&mut context)?;
    control_writer::write_sources(&mut context)?;
    let fields = std::mem::take(&mut context.text_fields);
    fields.write_identities(&mut context)?;
    let declared_control_ids = get(&model, "contentControls").as_array().into_iter().flatten()
        .filter_map(|control| get(control, "id").as_str()).collect::<BTreeSet<_>>();
    let written_controls = context.used_controls.iter().filter(|id| declared_control_ids.contains(id.as_str())).count();
    if written_controls < declared_control_count {
        return Err(ExportError::invalid(format!("DOCX_CONTENT_CONTROL_UNWRITTEN: {} content control(s) have no complete native wrapper; export cannot discard their semantics", declared_control_count - written_controls)));
    }
    if blocks.is_empty() {
        body.push_str("<w:p><w:r><w:t></w:t></w:r></w:p>");
    }
    let final_section = context.current_section.clone();
    let final_refs = context.section_header_footer_refs.get(&final_section);
    let document_xml = document_xml_with_headers(&body, page, final_refs);
    context
        .package
        .add_xml("word/document.xml", DOCUMENT_CONTENT_TYPE, document_xml)?;
    let has_story_fields = section_header_footers.as_object().is_some_and(|sections| sections.values().any(|s| s["pageFields"].as_array().is_some_and(|fields| !fields.is_empty())));
    if has_different_odd_even || has_story_fields || has_text_fields || !notes.footnotes.is_empty() || !notes.endnotes.is_empty() || !get(&model, "noteSettings").is_null() {
        let notes_settings = note_writer::document_properties(get(&model, "noteSettings"), &notes)?;
        context
            .package
            .add_xml("word/settings.xml", SETTINGS_CONTENT_TYPE, settings_xml(has_different_odd_even, &notes_settings))?;
        context.relationship(REL_SETTINGS, "settings.xml", false)?;
    }
    // Package global resources after body, tables, stories, notes and comments
    // have all had the opportunity to import their native dependency closures.
    let numbering_xml = context.source_styles.numbering_xml(context.package)?;
    context.package.add_xml("word/numbering.xml", NUMBERING_CONTENT_TYPE, numbering_xml)?;
    context.package.add_xml("word/styles.xml", STYLES_CONTENT_TYPE, context.source_styles.styles_xml().map_err(ExportError::invalid)?)?;
    if let Some(theme) = context.source_styles.theme_xml().map(str::to_owned) {
        context.package.add_xml("word/theme/theme1.xml", "application/vnd.openxmlformats-officedocument.theme+xml", theme)?;
        context.relationship(REL_THEME, "theme/theme1.xml", false)?;
    }
    Ok(context.warnings)
}

fn styles_xml() -> String {
    let mut headings = String::new();
    for (id, size, extra) in [("Title", 52, "<w:b/>"), ("Subtitle", 30, "<w:color w:val=\"999999\"/>")] {
        headings.push_str(&format!("<w:style w:type=\"paragraph\" w:styleId=\"{id}\"><w:name w:val=\"{id}\"/><w:basedOn w:val=\"Normal\"/><w:next w:val=\"Normal\"/><w:qFormat/><w:rPr>{extra}<w:sz w:val=\"{size}\"/><w:szCs w:val=\"{size}\"/></w:rPr></w:style>"));
    }
    for level in 1..=6 {
        let size = match level {
            1 => 32,
            2 => 28,
            3 => 24,
            _ => 22,
        };
        let outline = level - 1;
        headings.push_str(&format!(
            "<w:style w:type=\"paragraph\" w:styleId=\"Heading{level}\"><w:name w:val=\"heading {level}\"/><w:basedOn w:val=\"Normal\"/><w:next w:val=\"Normal\"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before=\"240\" w:after=\"120\"/></w:pPr><w:rPr><w:b/><w:sz w:val=\"{size}\"/><w:szCs w:val=\"{size}\"/></w:rPr><w:outlineLvl w:val=\"{outline}\"/></w:style>"
        ));
    }
    format!(
        "{}<w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii=\"Aptos\" w:hAnsi=\"Aptos\" w:eastAsia=\"Aptos\"/><w:sz w:val=\"22\"/><w:szCs w:val=\"22\"/><w:lang w:val=\"zh-CN\"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after=\"120\" w:line=\"360\" w:lineRule=\"auto\"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/><w:qFormat/></w:style><w:style w:type=\"paragraph\" w:styleId=\"Quote\"><w:name w:val=\"Quote\"/><w:basedOn w:val=\"Normal\"/><w:pPr><w:ind w:left=\"360\"/><w:pBdr><w:left w:val=\"single\" w:sz=\"12\" w:space=\"8\" w:color=\"94A3B8\"/></w:pBdr></w:pPr></w:style><w:style w:type=\"paragraph\" w:styleId=\"OfficeCodeBlock\"><w:name w:val=\"Office Code Block\"/><w:basedOn w:val=\"Normal\"/><w:pPr><w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"F1F4F8\"/><w:spacing w:before=\"120\" w:after=\"120\" w:line=\"288\" w:lineRule=\"auto\"/></w:pPr><w:rPr><w:rFonts w:ascii=\"Consolas\" w:hAnsi=\"Consolas\"/><w:sz w:val=\"20\"/><w:szCs w:val=\"20\"/></w:rPr></w:style><w:style w:type=\"paragraph\" w:styleId=\"FootnoteText\"><w:name w:val=\"footnote text\"/><w:basedOn w:val=\"Normal\"/><w:rPr><w:sz w:val=\"20\"/><w:szCs w:val=\"20\"/></w:rPr></w:style><w:style w:type=\"paragraph\" w:styleId=\"EndnoteText\"><w:name w:val=\"endnote text\"/><w:basedOn w:val=\"Normal\"/><w:rPr><w:sz w:val=\"20\"/><w:szCs w:val=\"20\"/></w:rPr></w:style><w:style w:type=\"character\" w:styleId=\"FootnoteReference\"><w:name w:val=\"FootnoteReference\"/><w:basedOn w:val=\"DefaultParagraphFont\"/><w:rPr><w:vertAlign w:val=\"superscript\"/></w:rPr></w:style><w:style w:type=\"character\" w:styleId=\"EndnoteReference\"><w:name w:val=\"EndnoteReference\"/><w:basedOn w:val=\"DefaultParagraphFont\"/><w:rPr><w:vertAlign w:val=\"superscript\"/></w:rPr></w:style>{headings}</w:styles>",
        xml::DECL
    )
}

fn document_xml(body: &str, page: PageGeometry) -> String {
    document_xml_with_headers(body, page, None)
}

fn document_xml_with_headers(
    body: &str,
    page: PageGeometry,
    header_footer_refs: Option<&SectionHeaderFooterRefs>,
) -> String {
    format!(
        "{}<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><w:body>{body}{}</w:body></w:document>",
        xml::DECL,
        section_properties_with_headers(page, None, header_footer_refs)
    )
}

fn section_properties_with_headers(
    page: PageGeometry,
    section_type: Option<&str>,
    header_footer_refs: Option<&SectionHeaderFooterRefs>,
) -> String {
    let orientation = if page.width > page.height {
        " w:orient=\"landscape\""
    } else {
        ""
    };
    let section_type =
        section_type.map_or_else(String::new, |value| format!("<w:type w:val=\"{value}\"/>"));
    let mut header_footer = String::new();
    if let Some(section_refs) = header_footer_refs {
        for (kind, relationship) in &section_refs.references {
            header_footer.push_str(&format!(
                "<w:{}Reference w:type=\"{}\" r:id=\"{}\"/>",
                if kind.starts_with("header:") {
                    "header"
                } else {
                    "footer"
                },
                kind.split_once(':')
                    .map_or("default", |(_, page_type)| page_type),
                xml_attribute(relationship),
            ));
        }
        header_footer.push_str(&section_refs.note_properties_xml);
    }
    let title_page = if header_footer_refs.is_some_and(|r| r.different_first_page) { "<w:titlePg/>" } else { "" };
    format!(
        "<w:sectPr>{header_footer}{section_type}<w:pgSz w:w=\"{}\" w:h=\"{}\"{orientation}/><w:pgMar w:top=\"{}\" w:right=\"{}\" w:bottom=\"{}\" w:left=\"{}\" w:header=\"{}\" w:footer=\"{}\" w:gutter=\"0\"/>{title_page}</w:sectPr>",
        page.width, page.height, page.top, page.right, page.bottom, page.left,
        header_footer_refs.map_or(720, |r| r.header_distance), header_footer_refs.map_or(720, |r| r.footer_distance),
    )
}

fn prepare_section_header_footers(
    context: &mut WriteContext<'_>,
    sections: &Value,
    first_setup: &Value,
    note_settings: &Value,
    blocks: &[Value],
) -> Result<bool, ExportError> {
    let mut section_ids = vec![("first".to_owned(), first_setup)];
    section_ids.extend(
        blocks
            .iter()
            .filter(|block| string(get(block, "type")).as_deref() == Some("sectionBreak"))
            .filter_map(|block| Some((string(get(block, "id"))?, get(get(block, "sectionBreak"), "setup")))),
    );
    let mut has_different_odd_even = false;
    let mut next_part = 1usize;
    for (section_index, (section_id, setup)) in section_ids.iter().enumerate() {
        let value = get(sections, section_id);
        let legacy_linked = section_index > 0
            && (value.is_null() || get(value, "linkToPrevious").as_bool() == Some(true));
        let story_links = get(value, "storyLinkToPrevious");
        let story_names = ["header", "footer", "firstPageHeader", "firstPageFooter", "evenPageHeader", "evenPageFooter"];
        if !story_links.is_null() && (story_links.as_object().is_none() || story_links.as_object().is_some_and(|links|
            links.iter().any(|(key, linked)| !story_names.contains(&key.as_str()) || linked.as_bool().is_none()
                || section_index == 0 && linked.as_bool() == Some(true)))) {
            return Err(ExportError::invalid("Invalid per-story header/footer associations"));
        }
        if section_index == 0 && get(value, "linkToPrevious").as_bool() == Some(true) {
            return Err(ExportError::invalid("The first section cannot link to a previous section"));
        }
        let different_first_page = get(value, "differentFirstPage").as_bool() == Some(true);
        let different_odd_even = get(value, "differentOddEven").as_bool() == Some(true);
        has_different_odd_even |= different_odd_even;
        let mut refs = SectionHeaderFooterRefs {
            references: Vec::new(),
            different_first_page,
            header_distance: twips(get(setup, "marginHeader").as_f64().unwrap_or(48.0)),
            footer_distance: twips(get(setup, "marginFooter").as_f64().unwrap_or(48.0)),
            note_properties_xml: note_writer::section_properties(setup, note_settings)?,
        };
        for (story, page_type, field_name, relationship_type, content_type) in [
            ("header", "default", "header", REL_HEADER, HEADER_CONTENT_TYPE),
            ("footer", "default", "footer", REL_FOOTER, FOOTER_CONTENT_TYPE),
            ("header", "first", "firstPageHeader", REL_HEADER, HEADER_CONTENT_TYPE),
            ("footer", "first", "firstPageFooter", REL_FOOTER, FOOTER_CONTENT_TYPE),
            ("header", "even", "evenPageHeader", REL_HEADER, HEADER_CONTENT_TYPE),
            ("footer", "even", "evenPageFooter", REL_FOOTER, FOOTER_CONTENT_TYPE),
        ] {
            let linked = section_index > 0 && get(story_links, field_name).as_bool().unwrap_or(legacy_linked);
            if linked { continue; }
            // Dormant first/even variants are native parts too; toggling their option must not discard content.
            let text = string(get(value, field_name)).unwrap_or_default();
            let page_fields = page_fields_for_story(value, field_name);
            add_header_footer_part(context, &mut refs, &mut next_part, story, page_type, &text, &page_fields,
                get(get(value, "storyContent"), field_name), relationship_type, content_type)?;
        }
        context
            .section_header_footer_refs
            .insert(section_id.clone(), refs);
    }
    Ok(has_different_odd_even)
}

fn add_header_footer_part(
    context: &mut WriteContext<'_>,
    refs: &mut SectionHeaderFooterRefs,
    next_part: &mut usize,
    story: &str,
    page_type: &str,
    text: &str,
    page_fields: &[&Value],
    content: &Value,
    relationship_type: &str,
    content_type: &str,
) -> Result<(), ExportError> {
    let path = format!("word/{story}{}.xml", *next_part);
    *next_part += 1;
    story_writer::write_story(context, &path, content_type, story, text, page_fields, content)?;
    let target = path.strip_prefix("word/").unwrap_or(&path);
    let relationship = context.relationship(relationship_type, target, false)?;
    refs.references
        .push((format!("{story}:{page_type}"), relationship));
    Ok(())
}

fn page_fields_for_story<'a>(section: &'a Value, story: &str) -> Vec<&'a Value> {
    let mut fields: Vec<_> = field(section, "pageFields")
        .and_then(Value::as_array)
        .into_iter()
        .flatten()
        .filter(|entry| string(get(entry, "story")).as_deref() == Some(story))
        .filter(|entry| matches!(entry["field"].as_str(), Some("PAGE" | "NUMPAGES" | "DATE" | "TIME")))
        .collect();
    fields.sort_by_key(|entry| entry["offset"].as_u64().unwrap_or(u64::MAX));
    fields
}

fn settings_xml(different_odd_even: bool, note_properties: &str) -> String {
    let variants = if different_odd_even { "<w:evenAndOddHeaders/>" } else { "" };
    format!(
        "{}<w:settings {CONTENT_NAMESPACES}>{variants}{note_properties}<w:updateFields w:val=\"true\"/></w:settings>",
        xml::DECL
    )
}

fn core_properties(model: &Value, title: &str) -> String {
    let creator = string(get(get(model, "metadata"), "creator"))
        .unwrap_or_else(|| "Office Viewer".to_owned());
    format!(
        "{}<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"><dc:title>{}</dc:title><dc:creator>{}</dc:creator><cp:lastModifiedBy>{}</cp:lastModifiedBy></cp:coreProperties>",
        xml::DECL,
        xml_text(title),
        xml_text(&creator),
        xml_text(&creator)
    )
}

fn run_properties_xml(run: &Value) -> String {
    let mut properties = String::new();
    let font = string(get(run, "fontFamily")).unwrap_or_else(|| "Aptos".to_owned());
    let font = font
        .split(',')
        .next()
        .unwrap_or("Aptos")
        .trim()
        .trim_matches(['\'', '"']);
    let family = xml_attribute(font);
    let east_asia = string(get(run, "eastAsiaFontFamily"))
        .filter(|name| !name.trim().is_empty())
        .map(|name| xml_attribute(&name))
        .unwrap_or_else(|| family.clone());
    properties.push_str(&format!(
        "<w:rFonts w:ascii=\"{family}\" w:hAnsi=\"{family}\" w:eastAsia=\"{east_asia}\"/>"
    ));
    for (field, tags) in [("bold", &["b", "bCs"][..]), ("italic", &["i", "iCs"][..]), ("strike", &["strike"][..])] {
        if let Some(enabled) = get(run, field).as_bool() {
            for tag in tags {
                if enabled { properties.push_str(&format!("<w:{tag}/>")); }
                else { properties.push_str(&format!("<w:{tag} w:val=\"0\"/>")); }
            }
        }
    }
    if let Some(color) = string(get(run, "color")).and_then(text_color) {
        properties.push_str(&format!("<w:color w:val=\"{color}\"/>"));
    }
    let size = (positive(get(run, "fontSize"), 11.0) * 2.0)
        .round()
        .clamp(2.0, 3_276.0) as u32;
    properties.push_str(&format!(
        "<w:sz w:val=\"{size}\"/><w:szCs w:val=\"{size}\"/>"
    ));
    if let Some(enabled) = get(run, "underline").as_bool() {
        properties.push_str(if enabled { "<w:u w:val=\"single\"/>" } else { "<w:u w:val=\"none\"/>" });
    }
    if let Some(highlight) = string(get(run, "highlight"))
        .or_else(|| string(get(get(run, "metadata"), "highlight")))
        .and_then(default_style::highlight_properties)
    {
        properties.push_str(&highlight);
    }
    let vertical_align = string(get(run, "verticalAlign"))
        .or_else(|| string(get(get(run, "metadata"), "verticalAlign")))
        .unwrap_or_default();
    match vertical_align.as_str() {
        "superscript" | "sup" => properties.push_str("<w:vertAlign w:val=\"superscript\"/>"),
        "subscript" | "sub" => properties.push_str("<w:vertAlign w:val=\"subscript\"/>"),
        _ => {}
    }
    properties
}

fn run_xml(run: &Value, text: &str) -> String {
    run_xml_with_defaults(run, text, RunDefaultEffects::default())
}

fn run_xml_with_defaults(run: &Value, text: &str, defaults: RunDefaultEffects) -> String {
    let properties = format!("<w:rPr>{}</w:rPr>", run_properties_xml_with_defaults(run, defaults));
    text.replace("\r\n", "\n")
        .replace('\r', "\n")
        .split('\n')
        .enumerate()
        .map(|(index, part)| {
            let break_xml = if index == 0 { "" } else { "<w:br/>" };
            format!(
                "<w:r>{properties}{break_xml}<w:t xml:space=\"preserve\">{}</w:t></w:r>",
                xml_text(part)
            )
        })
        .collect()
}

fn run_properties_xml_with_defaults(run: &Value, defaults: RunDefaultEffects) -> String {
    let mut properties = run_properties_xml(run);
    let vertical_align = string(get(run, "verticalAlign"))
        .or_else(|| string(get(get(run, "metadata"), "verticalAlign")))
        .unwrap_or_default();
    if defaults.vertical_alignment && !matches!(vertical_align.as_str(), "superscript" | "sup" | "subscript" | "sub") {
        // A missing alignment in the normalized model means baseline. Omitting
        // rPr here would inherit a superscript/subscript document default.
        properties.push_str("<w:vertAlign w:val=\"baseline\"/>");
    }
    let highlight = string(get(run, "highlight"))
        .or_else(|| string(get(get(run, "metadata"), "highlight")))
        .and_then(default_style::highlight_properties);
    if defaults.highlight && highlight.is_none() {
        properties.push_str("<w:highlight w:val=\"none\"/><w:shd w:val=\"nil\"/>");
    }
    properties
}

fn run_with_text(run: &Value, text: &str) -> Value {
    let mut value = run.clone();
    if let Some(object) = value.as_object_mut() {
        object.insert("text".to_owned(), Value::String(text.to_owned()));
    }
    value
}

fn run_link(run: &Value) -> Option<String> {
    string(get(run, "href"))
        .or_else(|| string(get(get(run, "metadata"), "href")))
        .or_else(|| string(get(get(run, "metadata"), "link")))
        .map(|value| value.trim().to_owned())
        .filter(|value| !value.is_empty())
}

fn page_geometry(value: &Value) -> Result<PageGeometry, ExportError> {
    let width = positive(get(value, "width"), 794.0);
    let height = positive(get(value, "height"), 1_123.0);
    let max_twips = 31_680.0;
    let page_width = (width * TWIPS_PER_PIXEL).round();
    let page_height = (height * TWIPS_PER_PIXEL).round();
    if !(720.0..=max_twips).contains(&page_width) || !(720.0..=max_twips).contains(&page_height) {
        return Err(ExportError::invalid(
            "Document page dimensions must fit Word's 0.5 to 22 inch page range.",
        ));
    }
    let top = twips(positive(get(value, "marginTop"), 72.0));
    let right = twips(positive(get(value, "marginRight"), 72.0));
    let bottom = twips(positive(get(value, "marginBottom"), 72.0));
    let left = twips(positive(get(value, "marginLeft"), 72.0));
    let page = PageGeometry {
        width: page_width as u32,
        height: page_height as u32,
        top,
        right,
        bottom,
        left,
    };
    if page.left + page.right >= page.width || page.top + page.bottom >= page.height {
        return Err(ExportError::invalid(
            "Document margins must leave a positive page content area.",
        ));
    }
    Ok(page)
}

fn decode_data_image(source: &str) -> Result<(&'static str, &'static str, Vec<u8>), String> {
    let Some(value) = source.strip_prefix("data:") else {
        return Err(
            if source.starts_with("http:") || source.starts_with("https:") {
                "the image is an external URL"
            } else if source.starts_with("blob:") {
                "the image is a temporary blob URL"
            } else {
                "the source is not an embedded raster data URI"
            }
            .to_owned(),
        );
    };
    let Some((metadata, encoded)) = value.split_once(',') else {
        return Err("the data URI is malformed".to_owned());
    };
    let mut parts = metadata.split(';');
    let mime = parts.next().unwrap_or_default().to_ascii_lowercase();
    if !parts.any(|part| part.eq_ignore_ascii_case("base64")) {
        return Err("only base64 data URIs can be embedded".to_owned());
    }
    let (content_type, extension) = match mime.as_str() {
        "image/png" => ("image/png", "png"),
        "image/jpeg" => ("image/jpeg", "jpg"),
        "image/gif" => ("image/gif", "gif"),
        "image/bmp" => ("image/bmp", "bmp"),
        "image/svg+xml" => {
            return Err("vector SVG images are not embedded by this writer".to_owned());
        }
        _ => {
            return Err(format!(
                "the image format {mime:?} is not supported for embedding"
            ));
        }
    };
    let bytes = STANDARD
        .decode(encoded)
        .map_err(|_| "the image base64 data is invalid".to_owned())?;
    let signature_ok = match extension {
        "png" => bytes.starts_with(b"\x89PNG\r\n\x1a\n"),
        "jpg" => bytes.starts_with(b"\xff\xd8\xff"),
        "gif" => bytes.starts_with(b"GIF87a") || bytes.starts_with(b"GIF89a"),
        "bmp" => bytes.starts_with(b"BM"),
        _ => false,
    };
    if !signature_ok {
        return Err("the image content does not match its declared format".to_owned());
    }
    Ok((content_type, extension, bytes))
}

fn get<'a>(value: &'a Value, key: &str) -> &'a Value {
    value.get(key).unwrap_or(&Value::Null)
}

fn field<'a>(value: &'a Value, key: &str) -> Option<&'a Value> {
    value.get(key)
}

fn string(value: &Value) -> Option<String> {
    value.as_str().map(str::to_owned)
}

fn integer(value: &Value) -> Option<i64> {
    value.as_i64().or_else(|| {
        value
            .as_f64()
            .filter(|value| value.is_finite())
            .map(|value| value.round() as i64)
    })
}

fn positive(value: &Value, fallback: f64) -> f64 {
    value
        .as_f64()
        .filter(|value| value.is_finite() && *value > 0.0)
        .unwrap_or(fallback)
}

fn twips(pixels: f64) -> u32 {
    (pixels * TWIPS_PER_PIXEL).round().clamp(0.0, 31_680.0) as u32
}

fn signed_twips(pixels: f64) -> i32 {
    (pixels * TWIPS_PER_PIXEL).round().clamp(-31_680.0, 31_680.0) as i32
}

fn nonnegative(value: &Value, fallback: f64) -> f64 {
    value.as_f64().filter(|value| value.is_finite() && *value >= 0.0).unwrap_or(fallback)
}

fn text_color(value: String) -> Option<String> {
    if value.eq_ignore_ascii_case("auto") { Some("auto".to_owned()) } else { color(value) }
}

fn color(value: String) -> Option<String> {
    let value = value.strip_prefix('#').unwrap_or(&value);
    (value.len() == 6 && value.chars().all(|char| char.is_ascii_hexdigit()))
        .then(|| value.to_ascii_uppercase())
}

fn xml_text(value: &str) -> String {
    let valid = value
        .chars()
        .filter(|character| {
            matches!(*character, '\t' | '\n' | '\r')
                || ('\u{20}'..='\u{D7FF}').contains(character)
                || ('\u{E000}'..='\u{FFFD}').contains(character)
                || ('\u{10000}'..='\u{10FFFF}').contains(character)
        })
        .collect::<String>();
    xml::escape(&valid).into_owned()
}

fn xml_attribute(value: &str) -> String {
    xml_text(&value.replace(['\r', '\n', '\t'], " "))
}

#[cfg(test)]
mod tests {
    use std::io::Read;

    use crate::opc::Package;

    use super::{
        build_numbering_xml, decode_data_image, document_xml, page_geometry, run_xml, styles_xml,
        write, PageGeometry,
    };
    use serde_json::json;

    #[test]
    fn page_dimensions_and_margins_use_word_twips() {
        let page = page_geometry(&json!({"width": 794, "height": 1123, "marginTop": 72, "marginRight": 72, "marginBottom": 72, "marginLeft": 72})).unwrap();
        assert_eq!(page.width, 11_910);
        assert_eq!(page.height, 16_845);
        assert_eq!(page.left, 1_080);
    }

    #[test]
    fn page_geometry_rejects_dimensions_that_cannot_fit_word() {
        let result = page_geometry(&json!({"width": 5000, "height": 1123}));
        assert!(result.is_err());
    }

    #[test]
    fn run_writer_preserves_whitespace_and_rich_marks() {
        let output = run_xml(
            &json!({"text":"  A&B\nC ","bold":true,"italic":true,"underline":true,"fontSize":12,"fontFamily":"Aptos","color":"#FF0088"}),
            "  A&B\nC ",
        );
        assert!(output.contains("<w:b/>"));
        assert!(output.contains("<w:i/>"));
        assert!(output.contains("<w:u w:val=\"single\"/>"));
        assert!(output.contains("<w:color w:val=\"FF0088\"/>"));
        assert!(output.contains("xml:space=\"preserve\">  A&amp;B</w:t>"));
        assert!(output.contains("<w:br/>"));
    }

    #[test]
    fn code_blocks_and_latex_runs_write_versioned_roundtrip_tags() {
        let model = json!({
            "kind": "document",
            "blocks": [
                { "id": "code", "type": "paragraph", "runs": [{ "text": "const x = 1\\nreturn x", "fontFamily": "Consolas", "fontSize": 10 }], "metadata": { "codeBlock": { "source": "const x = 1\\nreturn x", "language": "typescript", "lineNumbers": true } } },
                { "id": "math", "type": "paragraph", "runs": [{ "text": "\u{FFFC}", "metadata": { "latexFormula": { "latex": "x^2", "mode": "inline", "version": 1 } } }], "metadata": {} }
            ]
        });
        let mut package = Package::new();
        let warnings = write(&model.to_string(), &mut package).unwrap();
        assert_eq!(warnings.len(), 1, "{warnings:?}");
        assert_eq!(warnings[0].code, "DOCX_LATEX_SOURCE_FALLBACK");
        let bytes = package.finish().unwrap();
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .unwrap()
            .read_to_string(&mut document)
            .unwrap();
        let mut styles = String::new();
        archive
            .by_name("word/styles.xml")
            .unwrap()
            .read_to_string(&mut styles)
            .unwrap();
        assert!(document.contains("office_code_v1_typescript_1"));
        assert!(document.contains("<w:pStyle w:val=\"OfficeCodeBlock\"/>"));
        assert!(document.contains("<w:rFonts w:ascii=\"Consolas\""));
        assert!(document.contains("office-latex:v1:inline"));
        assert!(document.contains("<w:t xml:space=\"preserve\">x^2</w:t>"));
        assert!(styles.contains("w:styleId=\"OfficeCodeBlock\""));
    }

    #[test]
    fn section_headers_and_footers_write_parts_links_and_page_variants() {
        let model = json!({
            "kind": "document",
            "page": { "width": 794, "height": 1123 },
            "blocks": [
                { "id": "body-1", "type": "paragraph", "runs": [{ "text": "First" }] },
                { "id": "section-2", "type": "sectionBreak", "sectionBreak": { "type": "nextPage" }, "runs": [] },
                { "id": "body-2", "type": "paragraph", "runs": [{ "text": "Second" }] },
                { "id": "section-3", "type": "sectionBreak", "sectionBreak": { "type": "continuous", "setup": {"marginHeader": 25.5, "marginFooter": 0} }, "runs": [] },
                { "id": "body-3", "type": "paragraph", "runs": [{ "text": "Third" }] }
            ],
            "sectionHeaderFooters": {
                "first": {
                    "header": "A&B", "footer": "First footer",
                    "firstPageHeader": "Dormant cover header", "firstPageFooter": "Dormant cover footer",
                    "evenPageHeader": "Dormant even header", "evenPageFooter": "Dormant even footer"
                },
                "section-2": {
                    "header": "Second header",
                    "footer": "Second footer",
                    "firstPageHeader": "Cover header",
                    "firstPageFooter": "Cover footer",
                    "evenPageHeader": "Even header",
                    "evenPageFooter": "Even footer",
                    "differentFirstPage": true,
                    "differentOddEven": true,
                    "linkToPrevious": false
                },
                "section-3": { "linkToPrevious": true }
            }
        });
        let mut package = Package::new();
        let warnings = write(&model.to_string(), &mut package).unwrap();
        assert!(warnings.is_empty());
        let bytes = package.finish().unwrap();
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
        let read_part = |archive: &mut zip::ZipArchive<std::io::Cursor<Vec<u8>>>, path: &str| {
            let mut content = String::new();
            archive
                .by_name(path)
                .unwrap()
                .read_to_string(&mut content)
                .unwrap();
            content
        };
        let document = read_part(&mut archive, "word/document.xml");
        let first_header = read_part(&mut archive, "word/header1.xml");
        let dormant_first_header = read_part(&mut archive, "word/header3.xml");
        let dormant_even_header = read_part(&mut archive, "word/header5.xml");
        let section_two_first_header = read_part(&mut archive, "word/header9.xml");
        let even_header = read_part(&mut archive, "word/header11.xml");
        let relationships = read_part(&mut archive, "word/_rels/document.xml.rels");
        let settings = read_part(&mut archive, "word/settings.xml");

        assert!(first_header.contains("A&amp;B"));
        assert!(document.contains("w:header=\"383\" w:footer=\"0\""));
        assert!(first_header.starts_with("<?xml"));
        assert!(dormant_first_header.contains("Dormant cover header"));
        assert!(dormant_even_header.contains("Dormant even header"));
        assert!(section_two_first_header.contains("Cover header"));
        assert!(even_header.contains("Even header"));
        assert!(settings.contains("<w:evenAndOddHeaders/>"));
        assert!(relationships.contains("Target=\"header1.xml\""));
        assert_eq!(document.matches("<w:titlePg/>").count(), 1);
        assert!(document.contains("<w:headerReference w:type=\"default\""));
        assert!(document.contains("<w:headerReference w:type=\"first\""));
        assert!(document.contains("<w:headerReference w:type=\"even\""));
        assert_eq!(document.matches("<w:headerReference w:type=\"first\"").count(), 2);
        assert_eq!(document.matches("<w:headerReference w:type=\"even\"").count(), 2);
        assert_eq!(document.matches("<w:sectPr>").count(), 3);
    }

    #[test]
    fn run_writer_exports_vertical_alignment_and_explicit_baseline_clears_legacy_metadata() {
        let superscript = run_xml(&json!({"text":"2", "verticalAlign":"superscript"}), "2");
        let subscript = run_xml(&json!({"text":"2", "verticalAlign":"subscript"}), "2");
        let legacy = run_xml(
            &json!({"text":"2", "metadata":{"verticalAlign":"sup"}}),
            "2",
        );
        let baseline = run_xml(
            &json!({"text":"2", "verticalAlign":"baseline", "metadata":{"verticalAlign":"superscript"}}),
            "2",
        );

        assert!(superscript.contains("<w:vertAlign w:val=\"superscript\"/>"));
        assert!(subscript.contains("<w:vertAlign w:val=\"subscript\"/>"));
        assert!(legacy.contains("<w:vertAlign w:val=\"superscript\"/>"));
        assert!(!baseline.contains("<w:vertAlign"));
    }

    #[test]
    fn package_xml_has_page_setup_and_standard_styles() {
        let page = PageGeometry {
            width: 11910,
            height: 16845,
            top: 1080,
            right: 1080,
            bottom: 1080,
            left: 1080,
        };
        let document = document_xml("<w:p/>", page);
        assert!(document.contains("<w:pgSz w:w=\"11910\" w:h=\"16845\""));
        assert!(document.contains("<w:pgMar w:top=\"1080\""));
        assert!(styles_xml().contains("w:styleId=\"Heading1\""));
        assert!(build_numbering_xml(&[]).contains("<w:abstractNum"));
    }

    #[test]
    fn section_and_column_breaks_write_native_word_boundaries() {
        let break_types = [
            "nextPage",
            "continuous",
            "nextColumn",
            "evenPage",
            "oddPage",
        ];
        let mut blocks = break_types
            .iter()
            .map(|break_type| {
                json!({
                    "type": "sectionBreak",
                    "sectionBreak": { "type": break_type },
                    "runs": []
                })
            })
            .collect::<Vec<_>>();
        blocks.push(json!({ "type": "columnBreak", "runs": [] }));
        let model = json!({ "kind": "document", "blocks": blocks });
        let mut package = Package::new();
        let warnings = write(&model.to_string(), &mut package).unwrap();
        assert!(warnings.is_empty());
        let bytes = package.finish().unwrap();
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .unwrap()
            .read_to_string(&mut document)
            .unwrap();

        for break_type in break_types {
            assert!(document.contains(&format!("<w:type w:val=\"{break_type}\"/>")));
        }
        assert!(document.contains("<w:br w:type=\"column\"/>"));
        assert_eq!(
            document.matches("<w:sectPr>").count(),
            break_types.len() + 1
        );
    }

    #[test]
    fn section_break_writer_rejects_unrecognized_types() {
        let model = json!({
            "kind": "document",
            "blocks": [{ "type": "sectionBreak", "sectionBreak": { "type": "invalid" }, "runs": [] }]
        });
        let mut package = Package::new();
        assert!(write(&model.to_string(), &mut package).is_err());
    }

    #[test]
    fn checklist_paragraphs_write_checked_state_as_native_word_numbering_symbols() {
        let model = json!({
            "kind": "document",
            "blocks": [
                {"id": "unchecked", "type": "checklist", "checklist": {"listId": "tasks", "checked": false}, "runs": [{"text": "Open the document"}]},
                {"id": "checked", "type": "checklist", "checklist": {"listId": "tasks", "checked": true}, "runs": [{"text": "Review the checklist"}]}
            ]
        });
        let mut package = Package::new();
        let warnings = write(&model.to_string(), &mut package).unwrap();
        assert!(warnings.is_empty(), "{warnings:?}");
        let bytes = package.finish().unwrap();
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .unwrap()
            .read_to_string(&mut document)
            .unwrap();
        let mut numbering = String::new();
        archive
            .by_name("word/numbering.xml")
            .unwrap()
            .read_to_string(&mut numbering)
            .unwrap();

        assert!(document.contains("<w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr>"));
        assert!(document.contains("<w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"2\"/></w:numPr>"));
        assert!(numbering.contains("w:abstractNumId=\"2\""));
        assert!(numbering.contains("w:abstractNumId=\"3\""));
        assert!(numbering.contains("w:lvlText w:val=\"&#x2610;\""));
        assert!(numbering.contains("w:lvlText w:val=\"&#x2611;\""));
        assert!(numbering.contains("w:ascii=\"Segoe UI Symbol\""));
        assert!(document.contains("Open the document"));
        assert!(document.contains("Review the checklist"));
    }

    #[test]
    fn data_uri_image_requires_matching_raster_signature() {
        let encoded = "data:image/png;base64,iVBORw0KGgo=";
        let image = decode_data_image(encoded).unwrap();
        assert_eq!(image.0, "image/png");
        assert_eq!(image.1, "png");
        assert_eq!(image.2, b"\x89PNG\r\n\x1a\n");
        assert!(decode_data_image("data:image/svg+xml;base64,PHN2Zy8+").is_err());
        assert!(decode_data_image("data:image/png;base64,YWJj").is_err());
    }

    #[test]
    fn table_row_span_keeps_its_declared_covered_row() {
        let model = json!({
            "kind": "document",
            "name": "Merged table",
            "blocks": [{
                "type": "table",
                "metadata": {"table": {"rows": [{
                    "cells": [{
                        "rowSpan": 2,
                        "blocks": [{"type": "paragraph", "runs": [{"text": "Merged"}]}]
                    }]
                }, { "cells": [] }]}}
            }]
        });
        let mut package = Package::new();
        write(&model.to_string(), &mut package).unwrap();
        let bytes = package.finish().unwrap();
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .unwrap()
            .read_to_string(&mut document)
            .unwrap();
        assert_eq!(document.matches("<w:tr>").count(), 2);
        assert!(document.contains("<w:vMerge w:val=\"restart\"/>"));
        assert!(document.contains("<w:vMerge w:val=\"continue\"/>"));
    }

    #[test]
    fn footnotes_and_endnotes_are_written_as_related_opc_parts() {
        let model = json!({
            "kind": "document",
            "name": "Notes",
            "blocks": [{
                "type": "paragraph",
                "runs": [
                    {"text": "First", "footnoteId": "f-1"},
                    {"text": " and ", "bold": true},
                    {"text": "second", "metadata": {"endnoteId": "e-1"}}
                ]
            }],
            "footnotes": {
                "f-1": {"body": {"runs": [{"text": "Footnote body", "italic": true}]}}
            },
            "metadata": {
                "endnotes": {
                    "e-1": {"body": {"dataStream": "Endnote body\r\n"}}
                }
            }
        });
        let mut package = Package::new();
        let warnings = write(&model.to_string(), &mut package).unwrap();
        assert!(warnings.is_empty());
        let bytes = package.finish().unwrap();
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .unwrap()
            .read_to_string(&mut document)
            .unwrap();
        let mut footnotes = String::new();
        archive
            .by_name("word/footnotes.xml")
            .unwrap()
            .read_to_string(&mut footnotes)
            .unwrap();
        let mut endnotes = String::new();
        archive
            .by_name("word/endnotes.xml")
            .unwrap()
            .read_to_string(&mut endnotes)
            .unwrap();
        let mut relationships = String::new();
        archive
            .by_name("word/_rels/document.xml.rels")
            .unwrap()
            .read_to_string(&mut relationships)
            .unwrap();
        let mut content_types = String::new();
        archive
            .by_name("[Content_Types].xml")
            .unwrap()
            .read_to_string(&mut content_types)
            .unwrap();

        assert!(document.contains("<w:footnoteReference w:id=\"1\"/>"));
        assert!(document.contains("<w:endnoteReference w:id=\"1\"/>"));
        assert!(footnotes.contains("w:type=\"separator\" w:id=\"-1\""));
        assert!(footnotes.contains("Footnote body"));
        assert!(footnotes.contains("<w:footnote w:id=\"1\">"));
        assert!(endnotes.contains("Endnote body"));
        assert!(endnotes.contains("<w:endnote w:id=\"1\">"));
        assert!(relationships.contains(&format!(
            "Type=\"{}\" Target=\"footnotes.xml\"",
            super::REL_FOOTNOTES
        )));
        assert!(relationships.contains(&format!(
            "Type=\"{}\" Target=\"endnotes.xml\"",
            super::REL_ENDNOTES
        )));
        assert!(content_types.contains(super::FOOTNOTES_CONTENT_TYPE));
        assert!(content_types.contains(super::ENDNOTES_CONTENT_TYPE));
    }

    #[test]
    fn table_cell_note_references_use_the_same_note_parts() {
        let model = json!({
            "kind": "document",
            "blocks": [{
                "type": "table",
                "metadata": {"table": {"rows": [{"cells": [{"blocks": [{"id": "cell-p", "type": "paragraph", "runs": [{"text": "cell", "footnoteId": "cell-note"}]}]}]}]}}
            }],
            "footnotes": {"cell-note": {"body": {"text": "Cell footnote"}}},
            "contentControls": [{"id": "cell-control", "type": "text", "blockId": "cell-p", "startOffset": 0, "endOffset": 4}]
        });
        let mut package = Package::new();
        let warnings = write(&model.to_string(), &mut package).unwrap();
        assert!(warnings.is_empty(), "{warnings:?}");
        let bytes = package.finish().unwrap();
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .unwrap()
            .read_to_string(&mut document)
            .unwrap();
        let mut footnotes = String::new();
        archive
            .by_name("word/footnotes.xml")
            .unwrap()
            .read_to_string(&mut footnotes)
            .unwrap();
        assert!(document.contains("<w:footnoteReference w:id=\"1\"/>"));
        assert!(document.contains("<w:sdt><w:sdtPr><w:id w:val=\"1\"/><w:text/></w:sdtPr>"));
        assert!(footnotes.contains("Cell footnote"));
    }

    #[test]
    fn dangling_note_reference_gets_an_empty_definition_and_warning() {
        let model = json!({
            "kind": "document",
            "blocks": [{"type": "paragraph", "runs": [{"text": "marker", "footnoteId": "missing"}]}]
        });
        let mut package = Package::new();
        let warnings = write(&model.to_string(), &mut package).unwrap();
        assert!(warnings
            .iter()
            .any(|warning| warning.code == "DOCX_NOTE_MISSING"));
        let bytes = package.finish().unwrap();
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .unwrap()
            .read_to_string(&mut document)
            .unwrap();
        let mut footnotes = String::new();
        archive
            .by_name("word/footnotes.xml")
            .unwrap()
            .read_to_string(&mut footnotes)
            .unwrap();
        assert!(document.contains("<w:footnoteReference w:id=\"1\"/>"));
        assert!(footnotes.contains("<w:footnote w:id=\"1\">"));
        assert!(footnotes.contains("<w:footnoteRef/>"));
    }

    #[test]
    fn content_controls_write_native_sdt_properties_and_preserve_utf16_ranges() {
        let model = json!({
            "kind": "document",
            "blocks": [
                {"id": "text", "type": "paragraph", "runs": [{"text": "Alpha 🌏 beta"}]},
                {"id": "check", "type": "paragraph", "runs": [{"text": "On"}]},
                {"id": "list", "type": "paragraph", "runs": [{"text": "red"}]}
            ],
            "contentControls": [
                {"id": "plain", "type": "text", "blockId": "text", "startOffset": 6, "endOffset": 8, "title": "Planet", "alias": "World", "tag": "planet-value", "lock": "contentLocked"},
                {"id": "checkbox", "type": "checkbox", "blockId": "check", "startOffset": 0, "endOffset": 2, "checked": true},
                {"id": "dropdown", "type": "dropDownList", "blockId": "list", "startOffset": 0, "endOffset": 3, "value": "r", "items": [{"displayText": "Red", "value": "r"}, {"displayText": "Blue & green", "value": "b"}]},
                {"id": "group", "type": "group", "blockId": "text", "startOffset": 0, "endOffset": 5},
                {"id": "nested", "parentId": "group", "type": "text", "blockId": "text", "startOffset": 0, "endOffset": 5, "alias": "Group label"}
            ]
        });
        let mut package = Package::new();
        let warnings = write(&model.to_string(), &mut package).unwrap();
        assert!(warnings.is_empty(), "{warnings:?}");
        let bytes = package.finish().unwrap();
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .unwrap()
            .read_to_string(&mut document)
            .unwrap();

        assert!(
            document.contains("xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\"")
        );
        assert!(document.contains("<w:alias w:val=\"Planet\"/>"));
        assert!(document.contains("<w:alias w:val=\"Group label\"/>"));
        assert!(document.contains("<w:group/>"));
        assert!(document.contains("<w:tag w:val=\"planet-value\"/>"));
        assert!(document.contains("<w:lock w:val=\"contentLocked\"/>"));
        assert!(document.contains("<w:text/>"));
        assert!(document.contains("<w:t xml:space=\"preserve\">Alpha</w:t>"));
        assert!(document.contains("<w:t xml:space=\"preserve\"> </w:t>"));
        assert!(document.contains("<w:t xml:space=\"preserve\">🌏</w:t>"));
        assert!(document.contains("<w:t xml:space=\"preserve\"> beta</w:t>"));
        assert!(document.contains("<w14:checked w14:val=\"1\"/>"));
        assert!(document.contains("<w:dropDownList><w:listItem w:displayText=\"Red\" w:value=\"r\"/><w:listItem w:displayText=\"Blue &amp; green\" w:value=\"b\"/></w:dropDownList>"));
        assert_eq!(document.matches("<w:sdt>").count(), 5);
    }

    #[test]
    fn content_control_boundary_cannot_split_a_utf16_surrogate_pair() {
        let model = json!({
            "kind": "document",
            "blocks": [{"id": "text", "type": "paragraph", "runs": [{"text": "🌏"}]}],
            "contentControls": [{"id": "half", "type": "plainText", "blockId": "text", "startOffset": 1, "endOffset": 2}]
        });
        let mut package = Package::new();
        let result = write(&model.to_string(), &mut package);
        assert!(result.is_err());
    }
}
