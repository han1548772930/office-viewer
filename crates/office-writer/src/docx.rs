use base64::{engine::general_purpose::STANDARD, Engine as _};
use serde_json::Value;
use std::collections::{BTreeMap, BTreeSet, HashMap};

use crate::{opc::Package, xml, ExportError, ExportWarning};

const REL_OFFICE_DOCUMENT: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const REL_CORE_PROPERTIES: &str =
    "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
const REL_STYLES: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
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
const HEADER_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";
const FOOTER_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";
const SETTINGS_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";
const MAX_TABLE_SPAN: usize = 512;
const EMU_PER_PIXEL: f64 = 9_525.0;
const TWIPS_PER_PIXEL: f64 = 15.0;

#[derive(Clone, Copy, PartialEq, Eq)]
enum ListKind {
    Bullet,
    Numbered,
    ChecklistUnchecked,
    ChecklistChecked,
}

struct NumberingInstance {
    id: u32,
    kind: ListKind,
}

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

struct NoteBundle {
    footnotes: Vec<NoteEntry>,
    endnotes: Vec<NoteEntry>,
    ids: HashMap<(NoteKind, String), i64>,
    orphaned: Vec<(NoteKind, String)>,
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
}

struct WriteContext<'a> {
    package: &'a mut Package,
    warnings: Vec<ExportWarning>,
    list_ids: HashMap<usize, u32>,
    numbering: Vec<NumberingInstance>,
    note_ids: HashMap<(NoteKind, String), i64>,
    controls_by_block: HashMap<String, Vec<ContentControl>>,
    used_controls: BTreeSet<String>,
    next_relationship: u32,
    next_image: u32,
    page: PageGeometry,
    section_header_footer_refs: HashMap<String, SectionHeaderFooterRefs>,
    current_section: String,
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
            .add_relationship(Some("word/document.xml"), &id, kind, target, external)?;
        Ok(id)
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
        let (content_type, extension, bytes) = match decode_data_image(&src) {
            Ok(image) => image,
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
        let path = format!("word/media/image{index}.{extension}");
        self.package.add_binary(&path, content_type, bytes)?;
        let target = format!("media/image{index}.{extension}");
        let relationship = self.relationship(REL_IMAGE, &target, false)?;
        let alt = xml_attribute(&alt);
        Ok(Some(format!(
            "<w:r><w:drawing><wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\"><wp:extent cx=\"{width_emu}\" cy=\"{height_emu}\"/><wp:effectExtent l=\"0\" t=\"0\" r=\"0\" b=\"0\"/><wp:docPr id=\"{index}\" name=\"Image {index}\" descr=\"{alt}\"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect=\"1\"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><pic:pic><pic:nvPicPr><pic:cNvPr id=\"{index}\" name=\"Image {index}.{extension}\" descr=\"{alt}\"/><pic:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></pic:cNvPicPr></pic:nvPicPr><pic:blipFill><a:blip r:embed=\"{relationship}\"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{width_emu}\" cy=\"{height_emu}\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"
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
                let fragments = run_xml(&source_run, &source);
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
                if let Some(id) = self.note_ids.get(&(kind, source_id.clone())) {
                    let tag = kind.reference_tag();
                    let style = match kind {
                        NoteKind::Footnote => "FootnoteReference",
                        NoteKind::Endnote => "EndnoteReference",
                    };
                    (
                        format!(
                            "<w:r><w:rPr><w:rStyle w:val=\"{style}\"/></w:rPr><w:{tag} w:id=\"{id}\"/></w:r>"
                        ),
                        true,
                    )
                } else {
                    self.warn(
                        "DOCX_NOTE_MISSING",
                        format!("A {0} reference to note {source_id:?} has no matching note body; its marker text was preserved.", kind.collection().trim_end_matches('s')),
                    );
                    (run_xml(run, &text), false)
                }
            } else {
                (run_xml(run, &text), false)
            };
            if has_note_reference {
                result.push_str(&fragments);
                if !text.is_empty() {
                    result.push_str(&run_xml(run, &text));
                }
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

    fn runs_with_controls(&mut self, block: &Value, runs: &[Value]) -> Result<String, ExportError> {
        let Some(block_id) = string(get(block, "id")) else {
            return self.runs(runs);
        };
        let Some(controls) = self.controls_by_block.get(&block_id).cloned() else {
            return self.runs(runs);
        };
        if controls.is_empty() {
            return self.runs(runs);
        }
        let total = runs
            .iter()
            .map(|run| {
                string(get(run, "text"))
                    .unwrap_or_default()
                    .encode_utf16()
                    .count()
            })
            .sum::<usize>();
        let mut result = self.render_control_range(runs, &controls, 0, total, None, 0)?;
        if result.is_empty() {
            result.push_str("<w:r><w:t></w:t></w:r>");
        }
        Ok(result)
    }

    fn render_control_range(
        &mut self,
        runs: &[Value],
        controls: &[ContentControl],
        start: usize,
        end: usize,
        parent_id: Option<&str>,
        depth: usize,
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
                candidate_parent.as_deref() == parent_id
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
            result.push_str(&self.render_run_range(runs, cursor, control.start)?);
            let content = self.render_control_range(
                runs,
                controls,
                control.start,
                control.end,
                Some(&control.id),
                depth + 1,
            )?;
            if let Some(wrapped) = content_control_xml(&control, &content) {
                if get(&control.value, "rawSdtPr")
                    .as_str()
                    .is_some_and(|value| !value.is_empty())
                    || get(&control.value, "rawSdtEndPr")
                        .as_str()
                        .is_some_and(|value| !value.is_empty())
                {
                    self.warn(
                        "DOCX_SDT_RAW_PROPERTIES_UNSUPPORTED",
                        format!("Content control {:?} contains preserved raw OOXML properties that this writer cannot safely merge into the generated SDT.", control.id),
                    );
                }
                if !get(&control.value, "dataBinding").is_null() {
                    self.warn(
                        "DOCX_SDT_DATA_BINDING_UNSUPPORTED",
                        format!("Content control {:?} references a custom XML binding, but no matching custom XML store is present in the document model.", control.id),
                    );
                }
                result.push_str(&wrapped);
                self.used_controls.insert(control.id);
            } else {
                self.warn(
                    "DOCX_CONTENT_CONTROL_TYPE_UNSUPPORTED",
                    format!("Content control {:?} has unsupported inline type {:?}; its text was preserved without the wrapper.", control.id, control.kind),
                );
                result.push_str(&content);
            }
            cursor = control.end;
        }
        result.push_str(&self.render_run_range(runs, cursor, end)?);
        Ok(result)
    }

    fn render_run_range(
        &mut self,
        runs: &[Value],
        start: usize,
        end: usize,
    ) -> Result<String, ExportError> {
        if start == end {
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
            if overlap_start < overlap_end {
                let Some(part) = slice_utf16(&text, overlap_start - offset, overlap_end - offset)
                else {
                    return Err(ExportError::invalid(
                        "A content control boundary splits a UTF-16 surrogate pair.",
                    ));
                };
                let mut clipped = run.clone();
                if let Some(object) = clipped.as_object_mut() {
                    object.insert("text".to_owned(), Value::String(part.to_owned()));
                }
                result.push(clipped);
            }
            offset = run_end;
            if offset >= end {
                break;
            }
        }
        self.runs(&result)
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
        }
        if let Some(num_id) = self.list_ids.get(&id) {
            let level = integer(get(block, "level")).unwrap_or(0).clamp(0, 8);
            properties.push_str(&format!(
                "<w:numPr><w:ilvl w:val=\"{level}\"/><w:numId w:val=\"{num_id}\"/></w:numPr>"
            ));
        }

        let style = get(block, "style");
        let spacing_before = twips(positive(get(style, "spaceBefore"), 0.0));
        let spacing_after = twips(positive(get(style, "spaceAfter"), 8.0));
        let line = (positive(get(style, "lineHeight"), 1.5) * 240.0)
            .round()
            .clamp(1.0, 2_400.0) as u32;
        properties.push_str(&format!("<w:spacing w:before=\"{spacing_before}\" w:after=\"{spacing_after}\" w:line=\"{line}\" w:lineRule=\"auto\"/>"));
        let indent = twips(positive(get(style, "indent"), 0.0));
        if indent > 0 {
            properties.push_str(&format!("<w:ind w:left=\"{indent}\"/>"));
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
        let paragraph_properties = if properties.is_empty() {
            String::new()
        } else {
            format!("<w:pPr>{properties}</w:pPr>")
        };
        let source_runs = field(block, "runs")
            .and_then(Value::as_array)
            .map_or(&[][..], Vec::as_slice);
        let code_runs = is_code_block.then(|| {
            source_runs
                .iter()
                .map(|run| {
                    let mut run = run.clone();
                    if let Some(properties) = run.as_object_mut() {
                        properties.insert(
                            "fontFamily".to_owned(),
                            Value::String("Consolas".to_owned()),
                        );
                    }
                    run
                })
                .collect::<Vec<_>>()
        });
        let runs = code_runs.as_deref().unwrap_or(source_runs);
        let paragraph = format!(
            "<w:p>{paragraph_properties}{}</w:p>",
            self.runs_with_controls(block, runs)?
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
        let bookmark_id = u32::try_from(id & 0x7fff_ffff).unwrap_or_default();
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
        let table = get(get(block, "metadata"), "table");
        let source_rows = field(table, "rows")
            .and_then(Value::as_array)
            .map_or(&[][..], Vec::as_slice);
        let header_rows = integer(get(table, "headerRowCount")).unwrap_or(0).max(0) as usize;
        let mut future_merges: BTreeMap<usize, Vec<(usize, usize)>> = BTreeMap::new();
        let mut rendered_rows = Vec::new();
        let mut total_columns = 1usize;
        let mut row_count = source_rows.len();

        let mut row_index = 0usize;
        while row_index < row_count {
            if rendered_rows.len() >= 10_000 {
                return Err(ExportError::unsupported(
                    "A DOCX table exceeds the 10,000 row safety limit.",
                ));
            }
            let continuations = future_merges.remove(&row_index).unwrap_or_default();
            let source_cells = field(source_rows.get(row_index).unwrap_or(&Value::Null), "cells")
                .and_then(Value::as_array)
                .map_or(&[][..], Vec::as_slice);
            let mut items: Vec<TableItem<'_>> = continuations
                .into_iter()
                .map(|(column, span)| TableItem {
                    column,
                    span,
                    row_span: 0,
                    cell: None,
                    continuation: true,
                })
                .collect();
            let mut next_column = 0usize;
            for cell in source_cells {
                let span = bounded_span(integer(get(cell, "colSpan")).unwrap_or(1), self);
                let row_span = bounded_span(integer(get(cell, "rowSpan")).unwrap_or(1), self);
                let explicit_column = integer(get(cell, "columnIndex"))
                    .filter(|column| *column >= 0)
                    .map(|column| column as usize);
                let mut column = explicit_column.unwrap_or_else(|| {
                    while items
                        .iter()
                        .any(|item| column_intersects(next_column, item.column, item.span))
                    {
                        next_column += 1;
                    }
                    next_column
                });
                while items
                    .iter()
                    .any(|item| column_intersects(column, item.column, item.span))
                {
                    column += 1;
                    self.warn(
                        "DOCX_TABLE_OVERLAP",
                        "Overlapping table cells were shifted to the next available column.",
                    );
                }
                next_column = column.saturating_add(span);
                items.push(TableItem {
                    column,
                    span,
                    row_span,
                    cell: Some(cell),
                    continuation: false,
                });
            }
            items.sort_by_key(|item| item.column);

            let mut row_xml = String::from("<w:tr>");
            if row_index < header_rows {
                row_xml.push_str("<w:trPr><w:tblHeader/></w:trPr>");
            }
            let mut column = 0usize;
            for item in items {
                if item.column < column {
                    continue;
                }
                while column < item.column {
                    row_xml.push_str(&table_cell_xml(self, &[], 1, 1, false, None)?);
                    column += 1;
                }
                let span = item.span.max(1);
                let merge = !item.continuation && item.row_span > 1;
                let cell = if let Some(cell) = item.cell {
                    table_cell_xml(
                        self,
                        cell_blocks(cell),
                        span,
                        item.row_span,
                        merge,
                        Some(cell),
                    )?
                } else {
                    table_cell_xml(self, &[], span, 0, true, None)?
                };
                row_xml.push_str(&cell);
                if merge {
                    for offset in 1..item.row_span {
                        let target_row = row_index + offset;
                        future_merges
                            .entry(target_row)
                            .or_default()
                            .push((column, span));
                    }
                    row_count = row_count.max(row_index + item.row_span);
                }
                column = column.saturating_add(span);
            }
            total_columns = total_columns.max(column);
            row_xml.push_str("</w:tr>");
            rendered_rows.push(row_xml);
            row_index += 1;
        }

        let columns = total_columns.clamp(1, MAX_TABLE_SPAN);
        let widths = self.table_column_widths(table, columns);
        let grid = widths
            .iter()
            .map(|width| format!("<w:gridCol w:w=\"{width}\"/>"))
            .collect::<String>();
        let mut rows = String::new();
        for row in rendered_rows {
            rows.push_str(&row);
        }
        let width = widths.iter().copied().sum::<u32>();
        Ok(format!(
            "<w:tbl><w:tblPr><w:tblW w:w=\"{width}\" w:type=\"dxa\"/><w:tblLayout w:type=\"fixed\"/><w:tblBorders><w:top w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"AAB3BF\"/><w:left w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"AAB3BF\"/><w:bottom w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"AAB3BF\"/><w:right w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"AAB3BF\"/><w:insideH w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"AAB3BF\"/><w:insideV w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"AAB3BF\"/></w:tblBorders><w:tblCellMar><w:top w:w=\"80\" w:type=\"dxa\"/><w:left w:w=\"100\" w:type=\"dxa\"/><w:bottom w:w=\"80\" w:type=\"dxa\"/><w:right w:w=\"100\" w:type=\"dxa\"/></w:tblCellMar></w:tblPr><w:tblGrid>{grid}</w:tblGrid>{rows}</w:tbl>"
        ))
    }

    fn table_column_widths(&self, table: &Value, columns: usize) -> Vec<u32> {
        let available = self
            .page
            .width
            .saturating_sub(self.page.left + self.page.right)
            .max(720);
        let source = field(table, "columnWidths").and_then(Value::as_array);
        let mut widths = (0..columns)
            .map(|index| {
                source
                    .and_then(|values| values.get(index))
                    .map_or(0.0, |value| positive(value, 0.0))
            })
            .map(|width| (width * TWIPS_PER_PIXEL).round().max(0.0) as u32)
            .collect::<Vec<_>>();
        let total = widths.iter().copied().sum::<u32>();
        if total == 0 {
            widths.fill((available / columns as u32).max(1));
        } else {
            for width in &mut widths {
                *width = ((*width as u64 * u64::from(available)) / u64::from(total)).max(1) as u32;
            }
        }
        widths
    }
}

struct TableItem<'a> {
    column: usize,
    span: usize,
    row_span: usize,
    cell: Option<&'a Value>,
    continuation: bool,
}

fn cell_blocks(cell: &Value) -> &[Value] {
    field(cell, "blocks")
        .and_then(Value::as_array)
        .map_or(&[][..], Vec::as_slice)
}

fn table_cell_xml(
    context: &mut WriteContext<'_>,
    blocks: &[Value],
    span: usize,
    row_span: usize,
    merged: bool,
    cell: Option<&Value>,
) -> Result<String, ExportError> {
    let mut cell_properties = String::new();
    if span > 1 {
        cell_properties.push_str(&format!("<w:gridSpan w:val=\"{span}\"/>"));
    }
    if merged {
        let value = if row_span > 1 { "restart" } else { "continue" };
        cell_properties.push_str(&format!("<w:vMerge w:val=\"{value}\"/>"));
    }
    if let Some(cell) = cell {
        let style = get(cell, "style");
        if let Some(fill) = string(get(style, "fill")).and_then(color) {
            cell_properties.push_str(&format!(
                "<w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"{fill}\"/>"
            ));
        }
    }
    let cell_properties = if cell_properties.is_empty() {
        String::new()
    } else {
        format!("<w:tcPr>{cell_properties}</w:tcPr>")
    };
    let mut content = String::new();
    if blocks.is_empty() {
        content.push_str("<w:p><w:r><w:t></w:t></w:r></w:p>");
    } else {
        for block in blocks {
            match string(get(block, "type")).unwrap_or_default().as_str() {
                "table" => {
                    context.warn(
                        "DOCX_NESTED_TABLE",
                        "Nested table content was converted to plain text.",
                    );
                    let text = nested_table_text(block);
                    content.push_str(&format!(
                        "<w:p><w:r><w:t xml:space=\"preserve\">{}</w:t></w:r></w:p>",
                        xml_text(&text)
                    ));
                }
                _ => content.push_str(&context.block(block)?),
            }
        }
    }
    Ok(format!("<w:tc>{cell_properties}{content}</w:tc>"))
}

fn nested_table_text(block: &Value) -> String {
    let rows = field(get(get(block, "metadata"), "table"), "rows").and_then(Value::as_array);
    rows.map_or_else(String::new, |rows| {
        rows.iter()
            .map(|row| {
                field(row, "cells")
                    .and_then(Value::as_array)
                    .map_or_else(String::new, |cells| {
                        cells
                            .iter()
                            .map(|cell| {
                                cell_blocks(cell)
                                    .iter()
                                    .flat_map(|paragraph| {
                                        field(paragraph, "runs")
                                            .and_then(Value::as_array)
                                            .into_iter()
                                            .flatten()
                                            .map(|run| string(get(run, "text")).unwrap_or_default())
                                    })
                                    .collect::<String>()
                            })
                            .collect::<Vec<_>>()
                            .join("\t")
                    })
            })
            .collect::<Vec<_>>()
            .join("\n")
    })
}

fn bounded_span(value: i64, context: &mut WriteContext<'_>) -> usize {
    let value = value.max(1) as usize;
    if value > MAX_TABLE_SPAN {
        context.warn(
            "DOCX_TABLE_SPAN_LIMIT",
            format!("A table cell span larger than {MAX_TABLE_SPAN} was clamped."),
        );
        MAX_TABLE_SPAN
    } else {
        value
    }
}

fn column_intersects(start: usize, other_start: usize, other_span: usize) -> bool {
    start >= other_start && start < other_start.saturating_add(other_span)
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
    let note = get(metadata, "note");
    let kind = match string(get(note, "type")).as_deref() {
        Some("footnote") => NoteKind::Footnote,
        Some("endnote") => NoteKind::Endnote,
        _ => return None,
    };
    string(get(note, "noteId"))
        .or_else(|| get(note, "noteId").as_i64().map(|id| id.to_string()))
        .filter(|id| !id.trim().is_empty())
        .map(|id| (kind, id))
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

fn collect_references(blocks: &[Value], found: &mut BTreeSet<(NoteKind, String)>) {
    for block in blocks {
        if let Some(runs) = field(block, "runs").and_then(Value::as_array) {
            for run in runs {
                if let Some(reference) = note_reference(run) {
                    found.insert(reference);
                }
            }
        }
        let rows = field(get(get(block, "metadata"), "table"), "rows").and_then(Value::as_array);
        if let Some(rows) = rows {
            for cell in rows.iter().flat_map(|row| {
                field(row, "cells")
                    .and_then(Value::as_array)
                    .into_iter()
                    .flatten()
            }) {
                collect_references(cell_blocks(cell), found);
            }
        }
    }
}

fn collect_notes(model: &Value, blocks: &[Value]) -> NoteBundle {
    let mut notes: [Vec<NoteEntry>; 2] = [Vec::new(), Vec::new()];
    let mut ids = HashMap::new();
    let mut next_id = [1i64, 1i64];
    for kind in [NoteKind::Footnote, NoteKind::Endnote] {
        let index = if kind == NoteKind::Endnote { 1 } else { 0 };
        for (source_id, value) in collect_note_entries(model, kind) {
            if ids.contains_key(&(kind, source_id.clone())) {
                continue;
            }
            let id = next_id[index];
            next_id[index] += 1;
            ids.insert((kind, source_id.clone()), id);
            notes[index].push(NoteEntry { id, value });
        }
    }

    let mut references = BTreeSet::new();
    collect_references(blocks, &mut references);
    let mut orphaned = Vec::new();
    for (kind, source_id) in references {
        if ids.contains_key(&(kind, source_id.clone())) {
            continue;
        }
        let index = if kind == NoteKind::Endnote { 1 } else { 0 };
        let id = next_id[index];
        next_id[index] += 1;
        ids.insert((kind, source_id.clone()), id);
        notes[index].push(NoteEntry {
            id,
            value: Value::Null,
        });
        orphaned.push((kind, source_id));
    }
    NoteBundle {
        footnotes: std::mem::take(&mut notes[0]),
        endnotes: std::mem::take(&mut notes[1]),
        ids,
        orphaned,
    }
}

fn collect_content_controls(
    model: &Value,
    blocks: &[Value],
) -> HashMap<String, Vec<ContentControl>> {
    let Some(items) = field(model, "contentControls").and_then(Value::as_array) else {
        return HashMap::new();
    };
    let mut block_ids = BTreeSet::new();
    collect_block_ids(blocks, &mut block_ids);
    let mut controls = HashMap::<String, Vec<ContentControl>>::new();
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
            continue;
        };
        if id.trim().is_empty() || block_id.trim().is_empty() || !block_ids.contains(&block_id) {
            continue;
        }
        let numeric_id = u32::try_from(index.saturating_add(1)).unwrap_or(u32::MAX);
        controls.entry(block_id).or_default().push(ContentControl {
            id,
            kind,
            numeric_id,
            start,
            end,
            value: value.clone(),
        });
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
    controls
}

fn collect_block_ids(blocks: &[Value], found: &mut BTreeSet<String>) {
    for block in blocks {
        if let Some(id) = string(get(block, "id")).filter(|id| !id.trim().is_empty()) {
            found.insert(id);
        }
        let rows = field(get(get(block, "metadata"), "table"), "rows").and_then(Value::as_array);
        if let Some(rows) = rows {
            for cell in rows.iter().flat_map(|row| {
                field(row, "cells")
                    .and_then(Value::as_array)
                    .into_iter()
                    .flatten()
            }) {
                collect_block_ids(cell_blocks(cell), found);
            }
        }
    }
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

fn content_control_xml(control: &ContentControl, content: &str) -> Option<String> {
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
                    "<w:date w:fullDate=\"{}T00:00:00Z\"><w:dateFormat w:val=\"{}\"/><w:lid w:val=\"{}\"/>{}<w:storeMappedDataAs w:val=\"{}\"/></w:date>",
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
                "<w15:repeatingSection{}{} />",
                title
                    .map(|value| format!(" w15:sectionTitle=\"{}\"", xml_attribute(&value)))
                    .unwrap_or_default(),
                if lock_insert {
                    " w15:doNotAllowInsertDelete=\"1\""
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
        _ => return None,
    }
    if let Some(appearance) = string(get(value, "appearance")) {
        properties.push_str(&format!(
            "<w:appearance w:val=\"{}\"/>",
            xml_attribute(&appearance)
        ));
    }
    if let Some(color) = string(get(value, "color")).and_then(color) {
        properties.push_str(&format!("<w:color w:val=\"{color}\"/>"));
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
    Some(format!(
        "<w:sdt><w:sdtPr>{properties}</w:sdtPr><w:sdtContent>{content}</w:sdtContent></w:sdt>"
    ))
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

fn note_runs(value: &Value) -> &[Value] {
    for candidate in [value, get(value, "body")] {
        if let Some(runs) = field(candidate, "runs").and_then(Value::as_array) {
            return runs;
        }
        if let Some(blocks) = field(candidate, "blocks").and_then(Value::as_array) {
            if let Some(runs) = blocks
                .first()
                .and_then(|block| field(block, "runs"))
                .and_then(Value::as_array)
            {
                return runs;
            }
        }
    }
    &[]
}

fn note_body_xml(value: &Value, kind: NoteKind, id: i64) -> String {
    let style = kind.style();
    let reference = kind.body_reference_tag();
    let mut paragraphs = String::new();
    let runs = note_runs(value);
    if !runs.is_empty() {
        let content = runs
            .iter()
            .map(|run| {
                let text = string(get(run, "text")).unwrap_or_default();
                run_xml(run, &text)
            })
            .collect::<String>();
        paragraphs.push_str(&format!(
            "<w:p><w:pPr><w:pStyle w:val=\"{style}\"/></w:pPr><w:r><w:rPr><w:rStyle w:val=\"{style}\"/></w:rPr><w:{reference}/></w:r>{content}</w:p>"
        ));
    } else {
        let text = note_text(value).unwrap_or_default();
        let lines = text.split('\n').collect::<Vec<_>>();
        for (index, line) in lines.iter().enumerate() {
            if index == 0 {
                paragraphs.push_str(&format!(
                    "<w:p><w:pPr><w:pStyle w:val=\"{style}\"/></w:pPr><w:r><w:rPr><w:rStyle w:val=\"{style}\"/></w:rPr><w:{reference}/></w:r><w:r><w:rPr><w:rStyle w:val=\"{style}\"/></w:rPr><w:t xml:space=\"preserve\">{}</w:t></w:r></w:p>",
                    xml_text(line)
                ));
            } else {
                paragraphs.push_str(&format!(
                    "<w:p><w:pPr><w:pStyle w:val=\"{style}\"/></w:pPr><w:r><w:rPr><w:rStyle w:val=\"{style}\"/></w:rPr><w:t xml:space=\"preserve\">{}</w:t></w:r></w:p>",
                    xml_text(line)
                ));
            }
        }
    }
    if paragraphs.is_empty() {
        paragraphs.push_str(&format!(
            "<w:p><w:pPr><w:pStyle w:val=\"{style}\"/></w:pPr><w:r><w:rPr><w:rStyle w:val=\"{style}\"/></w:rPr><w:{reference}/></w:r></w:p>"
        ));
    }
    format!(
        "<w:{} w:id=\"{id}\">{paragraphs}</w:{}>",
        kind.collection().trim_end_matches('s'),
        kind.collection().trim_end_matches('s')
    )
}

fn notes_xml(kind: NoteKind, notes: &[NoteEntry]) -> String {
    let element = kind.collection().trim_end_matches('s');
    let mut body = format!(
        "<w:{element} w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:{element}><w:{element} w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:{element}>"
    );
    for note in notes {
        body.push_str(&note_body_xml(&note.value, kind, note.id));
    }
    format!(
        "{}<w:{} xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">{body}</w:{}>",
        xml::DECL,
        kind.collection(),
        kind.collection()
    )
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
    let notes = collect_notes(&model, blocks);
    let controls_by_block = collect_content_controls(&model, blocks);
    let declared_control_count = field(&model, "contentControls")
        .and_then(Value::as_array)
        .map_or(0, Vec::len);
    let page = page_geometry(get(&model, "page"))?;
    let section_header_footers = field(&model, "sectionHeaderFooters")
        .cloned()
        .unwrap_or(Value::Null);
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
    package.add_xml("word/styles.xml", STYLES_CONTENT_TYPE, styles_xml())?;
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
        note_ids: notes.ids.clone(),
        controls_by_block,
        used_controls: BTreeSet::new(),
        next_relationship: 3,
        next_image: 1,
        page,
        section_header_footer_refs: HashMap::new(),
        current_section: "first".to_owned(),
    };
    let has_different_odd_even =
        prepare_section_header_footers(&mut context, &section_header_footers, blocks)?;
    let mut next_num_id = 1;
    assign_lists(
        blocks,
        &mut context.list_ids,
        &mut context.numbering,
        &mut next_num_id,
    );
    for block in blocks {
        if string(get(block, "type")).as_deref() == Some("table") {
            let rows =
                field(get(get(block, "metadata"), "table"), "rows").and_then(Value::as_array);
            if let Some(rows) = rows {
                for cell in rows.iter().flat_map(|row| {
                    field(row, "cells")
                        .and_then(Value::as_array)
                        .into_iter()
                        .flatten()
                }) {
                    assign_lists(
                        cell_blocks(cell),
                        &mut context.list_ids,
                        &mut context.numbering,
                        &mut next_num_id,
                    );
                }
            }
        }
    }

    let numbering_xml = build_numbering_xml(&context.numbering);
    context
        .package
        .add_xml("word/numbering.xml", NUMBERING_CONTENT_TYPE, numbering_xml)?;
    let mut body = String::new();
    for block in blocks {
        body.push_str(&context.block(block)?);
    }
    for (kind, source_id) in &notes.orphaned {
        context.warn(
            "DOCX_NOTE_MISSING",
            format!(
                "A {0} reference to note {source_id:?} has no matching note body; an empty note was written to keep the reference valid.",
                kind.collection().trim_end_matches('s')
            ),
        );
    }
    let used_control_count = context.used_controls.len();
    if used_control_count < declared_control_count {
        context.warn(
            "DOCX_CONTENT_CONTROL_UNWRITTEN",
            format!(
                "{} content control(s) could not be mapped to a supported inline DOCX control and were preserved as ordinary paragraph text.",
                declared_control_count - used_control_count
            ),
        );
    }
    if !notes.footnotes.is_empty() {
        context.package.add_xml(
            "word/footnotes.xml",
            FOOTNOTES_CONTENT_TYPE,
            notes_xml(NoteKind::Footnote, &notes.footnotes),
        )?;
        context.relationship(REL_FOOTNOTES, "footnotes.xml", false)?;
    }
    if !notes.endnotes.is_empty() {
        context.package.add_xml(
            "word/endnotes.xml",
            ENDNOTES_CONTENT_TYPE,
            notes_xml(NoteKind::Endnote, &notes.endnotes),
        )?;
        context.relationship(REL_ENDNOTES, "endnotes.xml", false)?;
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
    if has_different_odd_even {
        context
            .package
            .add_xml("word/settings.xml", SETTINGS_CONTENT_TYPE, settings_xml())?;
        context.relationship(REL_SETTINGS, "settings.xml", false)?;
    }
    Ok(context.warnings)
}

fn assign_lists(
    blocks: &[Value],
    lookup: &mut HashMap<usize, u32>,
    numbering: &mut Vec<NumberingInstance>,
    next_num_id: &mut u32,
) {
    let mut active_kind = None;
    let mut active_id = None;
    for block in blocks {
        let kind = match string(get(block, "type")).as_deref() {
            Some("bullet") => Some(ListKind::Bullet),
            Some("numbered") => Some(ListKind::Numbered),
            Some("checklist") => Some(
                if get(get(block, "checklist"), "checked").as_bool() == Some(true) {
                    ListKind::ChecklistChecked
                } else {
                    ListKind::ChecklistUnchecked
                },
            ),
            _ => None,
        };
        if let Some(kind) = kind {
            if active_kind != Some(kind) {
                let id = *next_num_id;
                *next_num_id += 1;
                numbering.push(NumberingInstance { id, kind });
                active_id = Some(id);
                active_kind = Some(kind);
            }
            if let Some(id) = active_id {
                lookup.insert(block as *const Value as usize, id);
            }
        } else {
            active_kind = None;
            active_id = None;
        }
    }
}

fn build_numbering_xml(instances: &[NumberingInstance]) -> String {
    let mut abstract_numbering = String::new();
    for (abstract_id, kind) in [
        (0, ListKind::Bullet),
        (1, ListKind::Numbered),
        (2, ListKind::ChecklistUnchecked),
        (3, ListKind::ChecklistChecked),
    ] {
        let mut levels = String::new();
        for level in 0..9 {
            let (format, text, left, font) = match kind {
                ListKind::Bullet => (
                    "bullet",
                    ["&#x2022;", "o", "&#x25AA;"][level % 3],
                    720 + level * 360,
                    "Arial",
                ),
                ListKind::Numbered => (
                    "decimal",
                    match level {
                        0 => "%1.",
                        1 => "%2.",
                        2 => "%3.",
                        _ => "%1.%2.%3.",
                    },
                    720 + level * 360,
                    "Arial",
                ),
                ListKind::ChecklistUnchecked => {
                    ("bullet", "&#x2610;", 720 + level * 360, "Segoe UI Symbol")
                }
                ListKind::ChecklistChecked => {
                    ("bullet", "&#x2611;", 720 + level * 360, "Segoe UI Symbol")
                }
            };
            levels.push_str(&format!(
                "<w:lvl w:ilvl=\"{level}\"><w:start w:val=\"1\"/><w:numFmt w:val=\"{format}\"/><w:lvlText w:val=\"{text}\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:tabs><w:tab w:val=\"num\" w:pos=\"{left}\"/></w:tabs><w:ind w:left=\"{left}\" w:hanging=\"360\"/></w:pPr><w:rPr><w:rFonts w:ascii=\"{font}\" w:hAnsi=\"{font}\"/></w:rPr></w:lvl>"
            ));
        }
        abstract_numbering.push_str(&format!("<w:abstractNum w:abstractNumId=\"{abstract_id}\"><w:multiLevelType w:val=\"multilevel\"/>{levels}</w:abstractNum>"));
    }
    let mut numbers = String::new();
    for instance in instances {
        let abstract_id = match instance.kind {
            ListKind::Bullet => 0,
            ListKind::Numbered => 1,
            ListKind::ChecklistUnchecked => 2,
            ListKind::ChecklistChecked => 3,
        };
        numbers.push_str(&format!(
            "<w:num w:numId=\"{}\"><w:abstractNumId w:val=\"{abstract_id}\"/></w:num>",
            instance.id
        ));
    }
    format!(
        "{}<w:numbering xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">{abstract_numbering}{numbers}</w:numbering>",
        xml::DECL
    )
}

fn styles_xml() -> String {
    let mut headings = String::new();
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
        if section_refs.different_first_page {
            header_footer.push_str("<w:titlePg/>");
        }
    }
    format!(
        "<w:sectPr>{section_type}{header_footer}<w:pgSz w:w=\"{}\" w:h=\"{}\"{orientation}/><w:pgMar w:top=\"{}\" w:right=\"{}\" w:bottom=\"{}\" w:left=\"{}\" w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/></w:sectPr>",
        page.width, page.height, page.top, page.right, page.bottom, page.left
    )
}

fn prepare_section_header_footers(
    context: &mut WriteContext<'_>,
    sections: &Value,
    blocks: &[Value],
) -> Result<bool, ExportError> {
    let mut section_ids = vec!["first".to_owned()];
    section_ids.extend(
        blocks
            .iter()
            .filter(|block| string(get(block, "type")).as_deref() == Some("sectionBreak"))
            .filter_map(|block| string(get(block, "id"))),
    );
    let mut has_different_odd_even = false;
    let mut next_part = 1usize;
    for (section_index, section_id) in section_ids.iter().enumerate() {
        let value = get(sections, section_id);
        let linked = section_index > 0
            && (value.is_null() || get(value, "linkToPrevious").as_bool() == Some(true));
        let different_first_page = get(value, "differentFirstPage").as_bool() == Some(true);
        let different_odd_even = get(value, "differentOddEven").as_bool() == Some(true);
        has_different_odd_even |= different_odd_even;
        let mut refs = SectionHeaderFooterRefs {
            references: Vec::new(),
            different_first_page,
        };
        if !linked {
            for (story, page_type, field_name, relationship_type, content_type) in [
                (
                    "header",
                    "default",
                    "header",
                    REL_HEADER,
                    HEADER_CONTENT_TYPE,
                ),
                (
                    "footer",
                    "default",
                    "footer",
                    REL_FOOTER,
                    FOOTER_CONTENT_TYPE,
                ),
            ] {
                let name = string(get(value, field_name)).unwrap_or_default();
                let page_fields = page_fields_for_story(value, field_name);
                add_header_footer_part(
                    context,
                    &mut refs,
                    &mut next_part,
                    story,
                    page_type,
                    &name,
                    &page_fields,
                    relationship_type,
                    content_type,
                )?;
            }
            if different_first_page {
                for (story, field_name, relationship_type, content_type) in [
                    ("header", "firstPageHeader", REL_HEADER, HEADER_CONTENT_TYPE),
                    ("footer", "firstPageFooter", REL_FOOTER, FOOTER_CONTENT_TYPE),
                ] {
                    let text = string(get(value, field_name)).unwrap_or_default();
                    let page_fields = page_fields_for_story(value, field_name);
                    add_header_footer_part(
                        context,
                        &mut refs,
                        &mut next_part,
                        story,
                        "first",
                        &text,
                        &page_fields,
                        relationship_type,
                        content_type,
                    )?;
                }
            }
            if different_odd_even {
                for (story, field_name, relationship_type, content_type) in [
                    ("header", "evenPageHeader", REL_HEADER, HEADER_CONTENT_TYPE),
                    ("footer", "evenPageFooter", REL_FOOTER, FOOTER_CONTENT_TYPE),
                ] {
                    let text = string(get(value, field_name)).unwrap_or_default();
                    let page_fields = page_fields_for_story(value, field_name);
                    add_header_footer_part(
                        context,
                        &mut refs,
                        &mut next_part,
                        story,
                        "even",
                        &text,
                        &page_fields,
                        relationship_type,
                        content_type,
                    )?;
                }
            }
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
    page_fields: &[String],
    relationship_type: &str,
    content_type: &str,
) -> Result<(), ExportError> {
    let path = format!("word/{story}{}.xml", *next_part);
    *next_part += 1;
    context
        .package
        .add_xml(&path, content_type, header_footer_xml(story, text, page_fields))?;
    let target = path.strip_prefix("word/").unwrap_or(&path);
    let relationship = context.relationship(relationship_type, target, false)?;
    refs.references
        .push((format!("{story}:{page_type}"), relationship));
    Ok(())
}

fn header_footer_xml(story: &str, text: &str, page_fields: &[String]) -> String {
    let root = if story == "header" { "hdr" } else { "ftr" };
    let fields = page_fields.iter().map(|field| {
        format!("<w:fldSimple w:instr=\" {field} \"><w:r><w:t></w:t></w:r></w:fldSimple>")
    }).collect::<String>();
    format!(
        "{}<w:{root} xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:p><w:pPr><w:jc w:val=\"center\"/></w:pPr>{}{fields}</w:p></w:{root}>",
        xml::DECL,
        run_xml(&serde_json::json!({}), text)
    )
}

fn page_fields_for_story(section: &Value, story: &str) -> Vec<String> {
    field(section, "pageFields")
        .and_then(Value::as_array)
        .into_iter()
        .flatten()
        .filter(|entry| string(get(entry, "story")).as_deref() == Some(story))
        .filter_map(|entry| string(get(entry, "field")))
        .filter(|field| matches!(field.as_str(), "PAGE" | "NUMPAGES"))
        .collect()
}

fn settings_xml() -> String {
    format!(
        "{}<w:settings xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:evenAndOddHeaders/><w:updateFields w:val=\"true\"/></w:settings>",
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

fn run_xml(run: &Value, text: &str) -> String {
    let mut properties = String::new();
    let font = string(get(run, "fontFamily")).unwrap_or_else(|| "Aptos".to_owned());
    let font = font
        .split(',')
        .next()
        .unwrap_or("Aptos")
        .trim()
        .trim_matches(['\'', '"']);
    let family = xml_attribute(font);
    properties.push_str(&format!(
        "<w:rFonts w:ascii=\"{family}\" w:hAnsi=\"{family}\" w:eastAsia=\"{family}\"/>"
    ));
    if get(run, "bold").as_bool() == Some(true) {
        properties.push_str("<w:b/><w:bCs/>");
    }
    if get(run, "italic").as_bool() == Some(true) {
        properties.push_str("<w:i/><w:iCs/>");
    }
    if get(run, "strike").as_bool() == Some(true) {
        properties.push_str("<w:strike/>");
    }
    if let Some(color) = string(get(run, "color")).and_then(color) {
        properties.push_str(&format!("<w:color w:val=\"{color}\"/>"));
    }
    let size = (positive(get(run, "fontSize"), 11.0) * 2.0)
        .round()
        .clamp(2.0, 3_276.0) as u32;
    properties.push_str(&format!(
        "<w:sz w:val=\"{size}\"/><w:szCs w:val=\"{size}\"/>"
    ));
    if get(run, "underline").as_bool() == Some(true) {
        properties.push_str("<w:u w:val=\"single\"/>");
    }
    if let Some(highlight) = string(get(get(run, "metadata"), "highlight")).and_then(color) {
        properties.push_str(&format!(
            "<w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"{highlight}\"/>"
        ));
    }
    let vertical_align = string(get(run, "verticalAlign"))
        .or_else(|| string(get(get(run, "metadata"), "verticalAlign")))
        .unwrap_or_default();
    match vertical_align.as_str() {
        "superscript" | "sup" => properties.push_str("<w:vertAlign w:val=\"superscript\"/>"),
        "subscript" | "sub" => properties.push_str("<w:vertAlign w:val=\"subscript\"/>"),
        _ => {}
    }
    let properties = format!("<w:rPr>{properties}</w:rPr>");
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
                { "id": "code", "type": "paragraph", "runs": [{ "text": "const x = 1\\nreturn x" }], "metadata": { "codeBlock": { "source": "const x = 1\\nreturn x", "language": "typescript", "lineNumbers": true } } },
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
                { "id": "section-3", "type": "sectionBreak", "sectionBreak": { "type": "continuous" }, "runs": [] },
                { "id": "body-3", "type": "paragraph", "runs": [{ "text": "Third" }] }
            ],
            "sectionHeaderFooters": {
                "first": { "header": "A&B", "footer": "First footer" },
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
        let section_two_first_header = read_part(&mut archive, "word/header5.xml");
        let even_header = read_part(&mut archive, "word/header7.xml");
        let relationships = read_part(&mut archive, "word/_rels/document.xml.rels");
        let settings = read_part(&mut archive, "word/settings.xml");

        assert!(first_header.contains("A&amp;B"));
        assert!(first_header.starts_with("<?xml"));
        assert!(section_two_first_header.contains("Cover header"));
        assert!(even_header.contains("Even header"));
        assert!(settings.contains("<w:evenAndOddHeaders/>"));
        assert!(relationships.contains("Target=\"header1.xml\""));
        assert_eq!(document.matches("<w:titlePg/>").count(), 1);
        assert!(document.contains("<w:headerReference w:type=\"default\""));
        assert!(document.contains("<w:headerReference w:type=\"first\""));
        assert!(document.contains("<w:headerReference w:type=\"even\""));
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
    fn table_row_span_emits_dynamic_continuation_row() {
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
                }]}}
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
