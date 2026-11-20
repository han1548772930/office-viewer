use base64::{Engine as _, engine::general_purpose::STANDARD};
use serde_json::Value;
use std::collections::{BTreeMap, HashMap};

use crate::{ExportError, ExportWarning, opc::Package, xml};

const REL_OFFICE_DOCUMENT: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const REL_CORE_PROPERTIES: &str =
    "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
const REL_STYLES: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
const REL_NUMBERING: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
const REL_IMAGE: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const REL_HYPERLINK: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
const DOCUMENT_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const STYLES_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
const NUMBERING_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";
const MAX_TABLE_SPAN: usize = 512;
const EMU_PER_PIXEL: f64 = 9_525.0;
const TWIPS_PER_PIXEL: f64 = 15.0;

#[derive(Clone, Copy, PartialEq, Eq)]
enum ListKind {
    Bullet,
    Numbered,
}

struct NumberingInstance {
    id: u32,
    kind: ListKind,
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

struct WriteContext<'a> {
    package: &'a mut Package,
    warnings: Vec<ExportWarning>,
    list_ids: HashMap<usize, u32>,
    numbering: Vec<NumberingInstance>,
    next_relationship: u32,
    next_image: u32,
    page: PageGeometry,
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
            let fragments = run_xml(run, &text);
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
            _ => self.paragraph(block, false),
        }
    }

    fn paragraph(&mut self, block: &Value, divider: bool) -> Result<String, ExportError> {
        let id = block as *const Value as usize;
        let mut properties = String::new();
        let block_type = string(get(block, "type")).unwrap_or_default();
        if block_type == "heading" {
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
        let runs = field(block, "runs")
            .and_then(Value::as_array)
            .map_or(&[][..], Vec::as_slice);
        Ok(format!(
            "<w:p>{paragraph_properties}{}</w:p>",
            self.runs(runs)?
        ))
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
                    row_xml.push_str(&table_cell_xml(&[], 1, 1, false, None, &mut self.warnings)?);
                    column += 1;
                }
                let span = item.span.max(1);
                let merge = !item.continuation && item.row_span > 1;
                let cell = if let Some(cell) = item.cell {
                    table_cell_xml(
                        cell_blocks(cell),
                        span,
                        item.row_span,
                        merge,
                        Some(cell),
                        &mut self.warnings,
                    )?
                } else {
                    table_cell_xml(&[], span, 0, true, None, &mut self.warnings)?
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
    blocks: &[Value],
    span: usize,
    row_span: usize,
    merged: bool,
    cell: Option<&Value>,
    warnings: &mut Vec<ExportWarning>,
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
                    warnings.push(ExportWarning::new(
                        "DOCX_NESTED_TABLE",
                        "Nested table content was converted to plain text.",
                    ));
                    let text = nested_table_text(block);
                    content.push_str(&format!(
                        "<w:p><w:r><w:t xml:space=\"preserve\">{}</w:t></w:r></w:p>",
                        xml_text(&text)
                    ));
                }
                "image" => {
                    warnings.push(ExportWarning::new(
                        "DOCX_TABLE_IMAGE",
                        "Image inside a table cell was exported as alt text.",
                    ));
                    let metadata = get(block, "metadata");
                    let alt = string(get(get(metadata, "image"), "alt"))
                        .unwrap_or_else(|| "Image".to_owned());
                    content.push_str(&format!(
                        "<w:p><w:r><w:t xml:space=\"preserve\">{}</w:t></w:r></w:p>",
                        xml_text(&format!("[Image: {alt}]"))
                    ));
                }
                _ => {
                    let runs = field(block, "runs")
                        .and_then(Value::as_array)
                        .map_or(&[][..], Vec::as_slice);
                    let text = runs
                        .iter()
                        .map(|run| string(get(run, "text")).unwrap_or_default())
                        .collect::<String>();
                    content.push_str(&format!(
                        "<w:p><w:r><w:t xml:space=\"preserve\">{}</w:t></w:r></w:p>",
                        xml_text(&text)
                    ));
                }
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
    let page = page_geometry(get(&model, "page"))?;
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
        next_relationship: 3,
        next_image: 1,
        page,
    };
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
    if blocks.is_empty() {
        body.push_str("<w:p><w:r><w:t></w:t></w:r></w:p>");
    }
    let document_xml = document_xml(&body, page);
    context
        .package
        .add_xml("word/document.xml", DOCUMENT_CONTENT_TYPE, document_xml)?;
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
    for (abstract_id, kind) in [(0, ListKind::Bullet), (1, ListKind::Numbered)] {
        let mut levels = String::new();
        for level in 0..9 {
            let (format, text, left) = match kind {
                ListKind::Bullet => (
                    "bullet",
                    ["&#x2022;", "o", "&#x25AA;"][level % 3],
                    720 + level * 360,
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
                ),
            };
            levels.push_str(&format!(
                "<w:lvl w:ilvl=\"{level}\"><w:start w:val=\"1\"/><w:numFmt w:val=\"{format}\"/><w:lvlText w:val=\"{text}\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:tabs><w:tab w:val=\"num\" w:pos=\"{left}\"/></w:tabs><w:ind w:left=\"{left}\" w:hanging=\"360\"/></w:pPr><w:rPr><w:rFonts w:ascii=\"Arial\" w:hAnsi=\"Arial\"/></w:rPr></w:lvl>"
            ));
        }
        abstract_numbering.push_str(&format!("<w:abstractNum w:abstractNumId=\"{abstract_id}\"><w:multiLevelType w:val=\"multilevel\"/>{levels}</w:abstractNum>"));
    }
    let mut numbers = String::new();
    for instance in instances {
        let abstract_id = if instance.kind == ListKind::Bullet {
            0
        } else {
            1
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
        "{}<w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii=\"Aptos\" w:hAnsi=\"Aptos\" w:eastAsia=\"Aptos\"/><w:sz w:val=\"22\"/><w:szCs w:val=\"22\"/><w:lang w:val=\"zh-CN\"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after=\"120\" w:line=\"360\" w:lineRule=\"auto\"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/><w:qFormat/></w:style><w:style w:type=\"paragraph\" w:styleId=\"Quote\"><w:name w:val=\"Quote\"/><w:basedOn w:val=\"Normal\"/><w:pPr><w:ind w:left=\"360\"/><w:pBdr><w:left w:val=\"single\" w:sz=\"12\" w:space=\"8\" w:color=\"94A3B8\"/></w:pBdr></w:pPr></w:style>{headings}</w:styles>",
        xml::DECL
    )
}

fn document_xml(body: &str, page: PageGeometry) -> String {
    let orientation = if page.width > page.height {
        " w:orient=\"landscape\""
    } else {
        ""
    };
    format!(
        "{}<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><w:body>{body}<w:sectPr><w:pgSz w:w=\"{}\" w:h=\"{}\"{orientation}/><w:pgMar w:top=\"{}\" w:right=\"{}\" w:bottom=\"{}\" w:left=\"{}\" w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/></w:sectPr></w:body></w:document>",
        xml::DECL,
        page.width,
        page.height,
        page.top,
        page.right,
        page.bottom,
        page.left
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
    match string(get(get(run, "metadata"), "verticalAlign"))
        .unwrap_or_default()
        .as_str()
    {
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
        PageGeometry, build_numbering_xml, decode_data_image, document_xml, page_geometry, run_xml,
        styles_xml, write,
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
}
