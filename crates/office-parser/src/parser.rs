use crate::model::{
    AnchorCellPayload, BorderPayload, CellPayload, ChartCachedPointPayload,
    ChartDrawingPayload, ChartPayload, ChartSeriesCachePayload, ChartSeriesPayload,
    ChartSeriesStylePayload, DataValidationPayload, DataValidationRangePayload, DimensionPayload,
    DocumentCommentPayload, DocumentCommentTargetPayload, DocumentModel,
    DocumentPayload, DocumentSectionLayoutPayload, DocumentStoryReferencePayload,
    FontPayload, FreezePanePayload, HyperlinkPayload, LocalDefinedNamePayload,
    ParseEnvelope, ParseFailure, ParserDiagnostics, ProtectionPayload, RangePayload,
    RichTextRunPayload, SCHEMA_VERSION, SheetPayload, StyleLossPayload, StylePayload,
    WorkbookPayload, WorkbookPropertiesPayload,
};
use crate::validation_sidecar::ValidationSidecar;
use crate::sparkline_sidecar::SparklineSidecar;
use crate::pivot_sidecar::PivotSidecar;
use crate::conditional_format_sidecar::ConditionalFormatSidecar;
use duke_pages::{BlockDetails, BlockType, DocNode, DocumentBlock, StructuredDocument, build_doc_model, extract_document_review, parse_doc, parse_document_model};
use quick_xml::events::Event;
use quick_xml::{Reader, XmlVersion};
use rxls::{Cell, Workbook};
use serde_json::{Value, json};

pub fn parse_office_bytes(bytes: &[u8], file_name: &str) -> Result<ParseEnvelope, ParseFailure> {
    validate_input(bytes, file_name)?;
    let format = detect_format(file_name, bytes);
    match format.as_str() {
        "doc" | "docx" | "odt" => parse_document_impl(bytes, file_name, format),
        "xls" | "xlsx" | "xlsb" | "ods" => parse_workbook(bytes, file_name, format),
        _ => Err(ParseFailure::unsupported(
            "Supported formats are .doc, .docx, .odt, .xls, .xlsx, .xlsb and .ods",
        )),
    }
}

pub fn parse_excel_bytes(bytes: &[u8], file_name: &str) -> Result<ParseEnvelope, ParseFailure> {
    validate_input(bytes, file_name)?;
    let format = detect_format(file_name, bytes);
    if !matches!(format.as_str(), "xls" | "xlsx" | "xlsb" | "ods") {
        return Err(ParseFailure::unsupported(
            "Expected an Excel workbook (.xls, .xlsx, .xlsb or .ods)",
        ));
    }
    parse_workbook(bytes, file_name, format)
}

pub fn parse_document_bytes(bytes: &[u8], file_name: &str) -> Result<ParseEnvelope, ParseFailure> {
    validate_input(bytes, file_name)?;
    let format = detect_format(file_name, bytes);
    if !matches!(format.as_str(), "doc" | "docx" | "odt") {
        return Err(ParseFailure::unsupported(
            "Expected a Word document (.doc, .docx or .odt)",
        ));
    }
    parse_document_impl(bytes, file_name, format)
}

fn parse_document_impl(
    bytes: &[u8],
    file_name: &str,
    format: String,
) -> Result<ParseEnvelope, ParseFailure> {
    let is_docx = format == "docx";
    // DOC defaults and content share one complete conversion. Reading defaults
    // must not call parse_doc/parse_document_model a second time.
    let doc_package = if format == "doc" {
        Some(parse_doc(bytes).map_err(ParseFailure::office_file)?)
    } else { None };
    // Keep the one native DOCX package through SDT projection, so nested runs
    // resolve the owning story's links/media without parsing the file twice.
    let docx_package = if is_docx { Some(duke_pages::parse_docx(bytes).map_err(ParseFailure::office_file)?) } else { None };
    let mut source_model = match doc_package.as_ref().or(docx_package.as_ref()) {
        Some(package) => build_doc_model(package),
        None => parse_document_model(bytes).map_err(ParseFailure::office_file)?,
    };
    let mut archive = if matches!(format.as_str(), "docx" | "odt") {
        Some(zip::ZipArchive::new(std::io::Cursor::new(bytes)).map_err(|error| ParseFailure::office_file(error.to_string()))?)
    } else { None };
    let mut format_warnings = Vec::new();
    let custom_xml_parts = archive.as_mut().filter(|_| is_docx).map(crate::document_custom_xml::read).transpose()
        .map_err(ParseFailure::office_file)?.flatten();
    let mut numbering_restoration = if let Some(zip) = archive.as_mut().filter(|_| is_docx) {
        crate::document_format_sidecar::restore_numbering_with_extensions(&mut source_model, zip)
    } else {
        let mut restoration = crate::document_format_sidecar::NumberingRestoration::default();
        restoration.definitions = source_model.metadata.numbering_definitions.as_ref().map(serde_json::to_value).transpose().map_err(|error| ParseFailure::office_file(error.to_string()))?;
        restoration
    };
    format_warnings.append(&mut numbering_restoration.warnings);
    let mut paragraph_wrap = if let Some(zip) = archive.as_mut().filter(|_| is_docx) {
        match crate::document_format_sidecar::read_word_wrap(zip) {
            Ok(wrap) => wrap,
            Err(error) => {
                format_warnings.push(format!("DOCX paragraph style properties could not be restored: {error}"));
                crate::document_format_sidecar::ParagraphWordWrap::default()
            }
        }
    } else { crate::document_format_sidecar::ParagraphWordWrap::default() };
    let mut office_source = archive.as_mut().filter(|_| is_docx).and_then(|zip| {
        let (resources, warnings) = crate::document_format_sidecar::read_office_source(zip);
        format_warnings.extend(warnings);
        resources
    });
    if let Some(zip) = archive.as_mut().filter(|_| is_docx) {
        let legacy_origins = crate::document_format_sidecar::read_legacy_origins(zip)
            .map_err(|error| ParseFailure::office_file(format!("DOCX declared legacy source could not be retained: {error}")))?;
        for origin in &legacy_origins {
            for error in &origin.resource_errors { format_warnings.push(format!("Historical {:?} source resource: {error}", origin.format)); }
        }
        if let Some(source) = &mut office_source { source.legacy_origins = legacy_origins; }
    }
    let mut default_style = if is_docx {
        match crate::document_format_sidecar::read_document_default_style(office_source.as_ref(), &paragraph_wrap) {
            Ok(style) => Some(style),
            Err(error) => { format_warnings.push(format!("DOCX document defaults could not be resolved: {error}")); None },
        }
    } else { None };
    if let Some(package) = doc_package.as_ref() {
        let defaults = crate::document_format_sidecar::read_doc_default_style(bytes, &package);
        office_source = Some(defaults.source);
        default_style = defaults.default_style;
        paragraph_wrap = defaults.paragraph_wrap;
        format_warnings.extend(defaults.warnings);
    } else if format == "odt" {
        if let Some(zip) = archive.as_mut() {
            let defaults = crate::document_format_sidecar::read_odt_default_style(&source_model, zip);
            office_source = Some(defaults.source);
            default_style = defaults.default_style;
            format_warnings.extend(defaults.warnings);
        }
    }
    if let Some(origin) = office_source.as_ref().and_then(|source| source.origin.as_ref()) {
        office_io_common::document_legacy_source::validate_origin(origin)
            .map_err(|error| ParseFailure::office_file(format!("Original Office source exceeds its contract: {error}")))?;
    }
    let story_style_sheet = if is_docx {
        match crate::document_format_sidecar::read_document_style_sheet(office_source.as_ref()) {
            Ok(sheet) => sheet,
            Err(error) => { format_warnings.push(format!("DOCX story styles could not be resolved: {error}")); duke_pages::parse::ParsedStyleSheet::empty() },
        }
    } else { duke_pages::parse::ParsedStyleSheet::empty() };
    let mut native_notes = crate::document_notes::read(docx_package.as_ref().or(doc_package.as_ref()), &mut source_model, &story_style_sheet)
        .map_err(|error| ParseFailure::office_file(format!("Native note settings and special stories could not be retained: {error}")))?;
    let mut document = crate::document_block_extraction::extract_with_resources(&mut source_model, docx_package.as_ref().or(doc_package.as_ref()), &story_style_sheet);
    drop(docx_package);
    drop(doc_package);
    if is_docx || format == "doc" {
        let table_warnings = crate::document_table_sidecar::restore(&mut document, &source_model, archive.as_mut().filter(|_| is_docx))
            .map_err(|error| ParseFailure::office_file(format!("Native document tables could not be retained: {error}")))?;
        document.warnings.extend(table_warnings);
    }
    let mut native_field_source = None;
    if is_docx || format == "doc" {
        let warnings = crate::document_paragraph_source::restore(&mut document, &source_model, archive.as_mut().filter(|_| is_docx))
            .map_err(|error| ParseFailure::office_file(format!("Native paragraph provenance could not be retained: {error}")))?;
        document.warnings.extend(warnings);
        crate::document_format_sidecar::restore_paragraphs(&mut document, &source_model, &paragraph_wrap);
    }
    if is_docx {
        crate::document_format_sidecar::restore_numbering_details(&mut document, &numbering_restoration);
        let field_identities = archive.as_mut().map(crate::document_text_fields::read_identities).transpose()
            .map_err(|error| ParseFailure::office_file(format!("DOCX native field identities could not be retained: {error}")))?.unwrap_or_default();
        let restoration = crate::document_text_fields::restore(&mut document, &source_model, &field_identities);
        document.warnings.extend(restoration.warnings);
        native_field_source = restoration.source;
    }
    native_notes.restore_references(&mut document, &source_model, &story_style_sheet)
        .map_err(|error| ParseFailure::office_file(format!("Native note references could not be retained: {error}")))?;
    let native_content_controls = if let Some(zip) = archive.as_mut().filter(|_| is_docx) {
        let restoration = crate::document_content_controls::restore(&document, &source_model, zip)
            .map_err(|error| ParseFailure::office_file(format!("Native content controls could not be retained: {error}")))?;
        document.warnings.extend(restoration.warnings);
        (!restoration.controls.is_empty()).then_some(restoration.controls)
    } else { None };
    let note_scope = crate::document_paragraph_source::canonical_scope(&source_model).map_err(ParseFailure::office_file)?;
    let body_part = source_model.metadata.source_part_names.as_ref().and_then(|names| names.get("document")).map(String::as_str).unwrap_or("word/document.xml");
    let document_note_columns = native_notes.settings.footnote.as_ref().and_then(|settings| settings.properties.column_count);
    let mut section_notes = source_model.metadata.sections.iter().flatten().map(|section|
        crate::document_notes::read_section(section.section_properties_xml.as_deref(), &note_scope, body_part, document_note_columns)
    ).collect::<Result<Vec<_>, _>>().map_err(ParseFailure::office_file)?;
    for (_, _, warnings) in &mut section_notes { document.warnings.append(warnings); }
    document.warnings.append(&mut native_notes.warnings);
    document.warnings.extend(format_warnings);
    insert_section_breaks(&mut document, &source_model);
    insert_paragraph_breaks(&mut document, &source_model);
    let mut review = extract_document_review(&source_model);
    let comment_warnings = crate::document_format_sidecar::restore_sdt_comment_targets(&source_model, &mut review.comments, &document);
    document.warnings.extend(comment_warnings);
    let mut diagnostics = ParserDiagnostics::document(&document);
    diagnostics.capabilities.extend([
        "pageBreaks".into(),
        "columnBreaks".into(),
        "sectionBreaks".into(),
    ]);
    if default_style.is_some() { diagnostics.capabilities.push("documentDefaultStyle".into()); }
    let mut stories = std::collections::HashMap::new();
    let mut story_reference = |part: &str, kind: &Option<String>, header| {
        let story = stories.entry(part.to_owned()).or_insert_with(|| archive.as_mut().filter(|_| is_docx).and_then(|zip| match crate::story_fields::read_story(zip, part, &story_style_sheet) {
            Ok(story) => story,
            Err(error) => { diagnostics.warnings.push(format!("DOCX story {part} could not be restored: {error}")); None },
        }));
        DocumentStoryReferencePayload {
            part_name: part.into(), reference_type: kind.clone(),
            page_fields: story.as_ref().map(|s| s.fields.iter().map(|f| f.field.clone()).collect()).unwrap_or_else(|| page_fields_for_story_nodes(&source_model, part, header)),
            story_text: story.clone(),
        }
    };
    let section_layout = source_model
        .metadata
        .sections
        .as_deref()
        .unwrap_or_default()
        .iter()
        .zip(section_notes)
        .map(|(section, (note_properties, suppress_endnotes, _))| {
            let (margin_header, margin_footer) =
                story_distances(section.section_properties_xml.as_deref().unwrap_or_default());
            DocumentSectionLayoutPayload {
                margin_header,
                margin_footer,
                start_node_index: section.start_node_index,
                section_properties_xml: section.section_properties_xml.clone(),
                header_sections: section.header_sections.iter().map(|r| story_reference(&r.part_name, &r.reference_type, true)).collect(),
                footer_sections: section.footer_sections.iter().map(|r| story_reference(&r.part_name, &r.reference_type, false)).collect(),
                note_properties,
                suppress_endnotes,
            }
        })
        .collect();
    let even_and_odd_headers = source_model
        .metadata
        .compatibility
        .as_ref()
        .and_then(|settings| settings.even_and_odd_headers)
        .unwrap_or(false);
    let footnotes = source_model.metadata.footnotes.as_ref().map(|notes| notes.iter().filter(|note| !native_notes.is_decoration(office_io_common::document_note_source::DocumentNoteKind::Footnote, note.id))
        .map(|note| native_notes.payload(office_io_common::document_note_source::DocumentNoteKind::Footnote, note)).collect());
    let endnotes = source_model.metadata.endnotes.as_ref().map(|notes| notes.iter().filter(|note| !native_notes.is_decoration(office_io_common::document_note_source::DocumentNoteKind::Endnote, note.id))
        .map(|note| native_notes.payload(office_io_common::document_note_source::DocumentNoteKind::Endnote, note)).collect());
    let note_settings = native_notes.into_settings();
    let mut comments = review.comments.into_iter().map(|comment| {
        let mut targets = comment.targets.into_iter().filter(|target| target.source.story == "body").map(|target| DocumentCommentTargetPayload {
            node_id: target.source.node_id,
            part_name: target.source.part_name,
            story: target.source.story,
            start: target.source.start,
            end: target.source.end,
            start_note_reference_count: None,
            end_note_reference_count: None,
            offset_encoding: target.source.offset_encoding,
        }).collect::<Vec<_>>();
        // A point comment has no text children for Duke to collect, but its
        // reference still identifies the paragraph and exact UTF-16 position.
        if targets.is_empty() {
            if let Some(reference) = comment.reference.filter(|reference| reference.story == "body" && reference.start.is_some() && reference.start == reference.end) {
                targets.push(DocumentCommentTargetPayload { node_id: reference.node_id, part_name: reference.part_name, story: reference.story,
                    start: reference.start, end: reference.end, start_note_reference_count: None, end_note_reference_count: None, offset_encoding: reference.offset_encoding });
            }
        }
        DocumentCommentPayload { id: comment.id.to_string(), author: comment.author, date: comment.date, text: comment.text,
            parent_id: comment.parent_id.map(|id| id.to_string()), resolved: comment.resolved, targets }
    }).collect::<Vec<_>>();
    let comment_note_warnings = crate::document_comment_notes::restore(&source_model, &document, &mut comments, &note_scope)
        .map_err(ParseFailure::office_file)?;
    diagnostics.warnings.extend(comment_note_warnings);
    let comments = (!comments.is_empty()).then_some(comments);
    Ok(ParseEnvelope {
        schema_version: SCHEMA_VERSION,
        parser_version: env!("CARGO_PKG_VERSION").to_owned(),
        format,
        file_name: file_name.to_owned(),
        size: bytes.len(),
        model: "document".into(),
        payload: DocumentModel::Document(DocumentPayload {
            document,
            numbering_definitions: numbering_restoration.definitions,
            section_layout,
            even_and_odd_headers,
            footnotes,
            endnotes,
            note_settings,
            comments,
            office_source,
            default_style,
            native_field_source,
            custom_xml_parts,
            native_content_controls,
        }),
        diagnostics,
    })
}

fn page_fields_for_story_nodes(
    source_model: &duke_pages::DocModel,
    part_name: &str,
    header: bool,
) -> Vec<String> {
    let metadata = serde_json::to_value(&source_model.metadata).unwrap_or(Value::Null);
    let story_key = if header { "headerSections" } else { "footerSections" };
    let mut fields = Vec::new();
    let mut references = Vec::new();
    references.extend(metadata[story_key].as_array().into_iter().flatten());
    references.extend(
        metadata["sections"]
            .as_array()
            .into_iter()
            .flatten()
            .flat_map(|section| section[story_key].as_array().into_iter().flatten()),
    );
    for reference in references {
        if reference["partName"].as_str() != Some(part_name) {
            continue;
        }
        collect_page_fields(&reference["nodes"], &mut fields);
    }
    fields
}

fn collect_page_fields(value: &Value, output: &mut Vec<String>) {
    match value {
        Value::String(xml) if xml.contains("instrText") || xml.contains("fldSimple") => {
            for instruction in page_field_instructions(xml) {
                if !output.contains(&instruction) {
                    output.push(instruction);
                }
            }
        }
        Value::Array(values) => values.iter().for_each(|value| collect_page_fields(value, output)),
        Value::Object(values) => values.values().for_each(|value| collect_page_fields(value, output)),
        _ => {}
    }
}

fn page_field_instructions(xml: &str) -> Vec<String> {
    let mut result = Vec::new();
    let mut remaining = xml;
    while let Some(index) = remaining.find("w:fldSimple") {
        remaining = &remaining[index..];
        let Some(tag_end) = remaining.find('>') else { break };
        let tag = &remaining[..tag_end];
        if let Some(attribute) = tag.find("w:instr=") {
            let value = remaining[attribute + "w:instr=".len()..].trim_start();
            if let Some(quote) = value.chars().next() {
                if let Some(end) = value[1..].find(quote) {
                    push_page_field(&value[1..1 + end], &mut result);
                }
            }
        }
        remaining = &remaining[tag_end + 1..];
    }
    remaining = xml;
    while let Some(index) = remaining.find("<w:instrText") {
        remaining = &remaining[index..];
        let Some(start) = remaining.find('>').map(|offset| offset + 1) else { break };
        let Some(end) = remaining[start..].find("</w:instrText>") else { break };
        push_page_field(&remaining[start..start + end], &mut result);
        remaining = &remaining[start + end + "</w:instrText>".len()..];
    }
    result
}

fn push_page_field(instruction: &str, output: &mut Vec<String>) {
    let name = instruction.split_whitespace().next().unwrap_or_default().to_ascii_uppercase();
    if matches!(name.as_str(), "PAGE" | "NUMPAGES") && !output.contains(&name) {
        output.push(name);
    }
}

fn story_distances(xml: &str) -> (Option<f64>, Option<f64>) {
    let mut reader = Reader::from_str(xml);
    let mut depth = 0usize;
    loop {
        match reader.read_event() {
            Ok(Event::Start(element) | Event::Empty(element)) if depth == 1 && local_name(element.name().as_ref().as_bytes()) == b"pgMar" => {
                let (mut header, mut footer) = (None, None);
                for attr in element.attributes().flatten() {
                    let pixels = attr.normalized_value(XmlVersion::Implicit1_0).ok().and_then(|value| value.parse::<u32>().ok()).map(|twips| f64::from(twips) / 15.0);
                    match local_name(attr.key.as_ref().as_bytes()) {
                        b"header" => header = pixels,
                        b"footer" => footer = pixels,
                        _ => {}
                    }
                }
                return (header, footer);
            }
            Ok(Event::Start(_)) => depth += 1,
            Ok(Event::End(_)) => depth = depth.saturating_sub(1),
            Ok(Event::Eof) | Err(_) => return (None, None),
            _ => {}
        }
    }
}

fn insert_section_breaks(document: &mut StructuredDocument, source_model: &duke_pages::DocModel) {
    let Some(sections) = source_model.metadata.sections.as_ref() else {
        return;
    };
    let mut insertions = Vec::new();
    for (section_index, section) in sections.iter().enumerate().skip(1) {
        let Ok(node_index) = usize::try_from(section.start_node_index) else {
            continue;
        };
        if node_index == 0 || node_index >= source_model.nodes.len() {
            continue;
        }
        let break_type = match section_break_type(
            sections[section_index - 1]
                .section_properties_xml
                .as_deref(),
        ) {
            Ok(break_type) => break_type,
            Err(()) => {
                document.warnings.push(format!(
                    "Section starting at body node {node_index} has an unsupported break type and was imported as nextPage."
                ));
                "nextPage"
            }
        };
        let replace_empty_boundary = matches!(
            source_model.nodes.get(node_index - 1),
            Some(DocNode::Paragraph(_))
        )
        .then(|| body_block_position(document, node_index - 1))
        .flatten()
        .filter(|position| document.blocks[*position].content.trim().is_empty());
        let target_position =
            replace_empty_boundary.or_else(|| body_block_position(document, node_index));
        let Some(position) = target_position else {
            document.warnings.push(format!(
                "Section starting at body node {node_index} could not be attached to an extracted block."
            ));
            continue;
        };
        let marker = DocumentBlock {
            object: "block".into(),
            id: format!("{}-section-break-{node_index}", document.document_id),
            r#type: BlockType::Text,
            content: String::new(),
            details: BlockDetails::Raw(serde_json::json!({
                "type": "section_break_details",
                "sectionBreak": { "type": break_type }
            })),
            note_id: None,
            source_refs: Vec::new(),
            parent_block_id: None,
            children: Vec::new(),
            words: None,
        };
        insertions.push((position, replace_empty_boundary.is_some(), marker));
    }
    insertions.sort_by(|left, right| right.0.cmp(&left.0));
    for (position, replace, marker) in insertions {
        if replace {
            document.blocks[position] = marker;
        } else {
            document.blocks.insert(position, marker);
        }
    }
}

fn insert_paragraph_breaks(document: &mut StructuredDocument, source_model: &duke_pages::DocModel) {
    for (node_index, node) in source_model.nodes.iter().enumerate() {
        let DocNode::Paragraph(paragraph) = node else {
            continue;
        };
        let Some(xml) = paragraph.source_xml.as_deref() else {
            continue;
        };
        let break_type = match word_break_type(xml) {
            Ok(Some(break_type)) => break_type,
            Ok(None) => continue,
            Err(()) => {
                document.warnings.push(format!(
                    "Could not read a manual break in body node {node_index}."
                ));
                continue;
            }
        };
        let details = match break_type {
            "pageBreak" => serde_json::json!({ "pageBreak": true }),
            _ => serde_json::json!({ "columnBreak": true }),
        };
        let marker = DocumentBlock {
            object: "block".into(),
            id: format!("{}-{break_type}-{node_index}", document.document_id),
            r#type: BlockType::Text,
            content: String::new(),
            details: BlockDetails::Raw(details),
            note_id: None,
            source_refs: Vec::new(),
            parent_block_id: None,
            children: Vec::new(),
            words: None,
        };
        if let Some(position) = body_block_position(document, node_index) {
            if !document.blocks[position].content.trim().is_empty() {
                document.warnings.push(format!(
                    "Manual break in body node {node_index} shares a paragraph with text and was retained as a line break."
                ));
                continue;
            }
            document.blocks[position] = marker;
        } else if let Some(position) = body_block_position(document, node_index + 1) {
            document.blocks.insert(position, marker);
        } else {
            document.warnings.push(format!(
                "Manual break in body node {node_index} could not be attached to an extracted block."
            ));
        }
    }
}

fn body_block_position(document: &StructuredDocument, node_index: usize) -> Option<usize> {
    let suffix = format!("/body/{node_index}");
    document.blocks.iter().position(|block| {
        block.parent_block_id.is_none()
            && block
                .source_refs
                .iter()
                .any(|source| source.story == "body" && source.node_id.ends_with(&suffix))
    })
}

fn section_break_type(xml: Option<&str>) -> Result<&'static str, ()> {
    let Some(xml) = xml else {
        return Ok("nextPage");
    };
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    loop {
        match reader.read_event() {
            Ok(Event::Start(element) | Event::Empty(element))
                if local_name(element.name().as_ref().as_bytes()) == b"type" =>
            {
                let value = element
                    .attributes()
                    .with_checks(false)
                    .filter_map(Result::ok)
                    .find(|attribute| local_name(attribute.key.as_ref().as_bytes()) == b"val")
                    .and_then(|attribute| {
                        attribute
                            .normalized_value_with(
                                XmlVersion::Implicit1_0,
                                1,
                                quick_xml::escape::resolve_xml_entity,
                            )
                            .ok()
                            .map(|value| value.into_owned())
                    });
                return match value.as_deref() {
                    None | Some("nextPage") => Ok("nextPage"),
                    Some("continuous") => Ok("continuous"),
                    Some("nextColumn") => Ok("nextColumn"),
                    Some("evenPage") => Ok("evenPage"),
                    Some("oddPage") => Ok("oddPage"),
                    Some(_) => Err(()),
                };
            }
            Ok(Event::Eof) => return Ok("nextPage"),
            Err(_) => return Err(()),
            _ => {}
        }
    }
}

fn word_break_type(xml: &str) -> Result<Option<&'static str>, ()> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    loop {
        match reader.read_event() {
            Ok(Event::Start(element) | Event::Empty(element))
                if local_name(element.name().as_ref().as_bytes()) == b"br" =>
            {
                let value = element
                    .attributes()
                    .with_checks(false)
                    .filter_map(Result::ok)
                    .find(|attribute| local_name(attribute.key.as_ref().as_bytes()) == b"type")
                    .and_then(|attribute| {
                        attribute
                            .normalized_value_with(
                                XmlVersion::Implicit1_0,
                                1,
                                quick_xml::escape::resolve_xml_entity,
                            )
                            .ok()
                            .map(|value| value.into_owned())
                    });
                match value.as_deref() {
                    Some("page") => return Ok(Some("pageBreak")),
                    Some("column") => return Ok(Some("columnBreak")),
                    _ => {}
                }
            }
            Ok(Event::Eof) => return Ok(None),
            Err(_) => return Err(()),
            _ => {}
        }
    }
}

fn local_name(qualified_name: &[u8]) -> &[u8] {
    qualified_name
        .rsplit(|byte| *byte == b':')
        .next()
        .unwrap_or(qualified_name)
}

fn parse_workbook(
    bytes: &[u8],
    file_name: &str,
    format: String,
) -> Result<ParseEnvelope, ParseFailure> {
    let workbook =
        Workbook::open(bytes).map_err(|error| ParseFailure::office_file(error.to_string()))?;
    let metadata = workbook.metadata();
    let mut validation_sidecar = if format == "xlsx" {
        ValidationSidecar::read(bytes)
    } else {
        ValidationSidecar::default()
    };
    if let Some(details) = validation_sidecar.failure.as_deref() {
        let resource_limit = details.starts_with("DV_SOURCE_MODEL_LIMIT") || details.starts_with("DV_SOURCE_FRAGMENT_LIMIT") || details.starts_with("DV_SOURCE_TOTAL_LIMIT");
        return Err(ParseFailure::new(
            if resource_limit { crate::model::ParseErrorCode::ParserFailure } else { crate::model::ParseErrorCode::InvalidOfficeFile },
            if resource_limit { "Data validation exceeds the parser resource limit" } else { "Could not preserve data validation from the XLSX file" },
            Some(details.to_owned()),
        ));
    }
    let mut sparkline_sidecar = if format == "xlsx" {
        SparklineSidecar::read(bytes)
    } else {
        SparklineSidecar::default()
    };
    if let Some(details) = sparkline_sidecar.failure.as_deref() {
        let resource_limit = details.starts_with("SPARKLINE_SOURCE_LIMIT");
        return Err(ParseFailure::new(
            if resource_limit { crate::model::ParseErrorCode::ParserFailure } else { crate::model::ParseErrorCode::InvalidOfficeFile },
            if resource_limit { "Sparkline metadata exceeds the parser resource limit" } else { "Could not preserve sparklines from the XLSX file" },
            Some(details.to_owned()),
        ));
    }
    let mut pivot_sidecar = if format == "xlsx" {
        PivotSidecar::read(bytes)
    } else {
        PivotSidecar::default()
    };
    let table_sidecar = if format == "xlsx" {
        crate::table_sidecar::TableSidecar::read(bytes)
    } else {
        crate::table_sidecar::TableSidecar::default()
    };
    if let Some(details) = table_sidecar.failure.as_deref() {
        return Err(ParseFailure::new(
            if details.contains("LIMIT") || details.contains("limit") { crate::model::ParseErrorCode::ParserFailure } else { crate::model::ParseErrorCode::InvalidOfficeFile },
            "Could not preserve structured tables from the XLSX file", Some(details.to_owned()),
        ));
    }
    let mut conditional_format_sidecar = if format == "xlsx" {
        ConditionalFormatSidecar::read(bytes)
    } else {
        ConditionalFormatSidecar::default()
    };
    if let Some(details) = conditional_format_sidecar.failure.as_deref() {
        let resource_limit = details.starts_with("CF_SOURCE_FRAGMENT_LIMIT")
            || details.starts_with("CF_SOURCE_TOTAL_LIMIT");
        return Err(ParseFailure::new(
            if resource_limit {
                crate::model::ParseErrorCode::ParserFailure
            } else {
                crate::model::ParseErrorCode::InvalidOfficeFile
            },
            if resource_limit {
                "Conditional formatting exceeds the parser resource limit"
            } else {
                "Could not preserve conditional formatting from the XLSX file"
            },
            Some(details.to_owned()),
        ));
    }
    let mut dimension_sidecar = if format == "xlsx" {
        crate::sheet_dimension_sidecar::SheetDimensionSidecar::read(bytes)
    } else {
        crate::sheet_dimension_sidecar::SheetDimensionSidecar::default()
    };
    let mut outline_sidecar = if format == "xlsx" {
        crate::sheet_outline_sidecar::SheetOutlineSidecar::read(bytes)
    } else {
        crate::sheet_outline_sidecar::SheetOutlineSidecar::default()
    };
    if let Some(details) = outline_sidecar.failure.as_deref() {
        return Err(outline_import_failure(details));
    }
    let mut sheets = workbook
        .sheets
        .iter()
        .enumerate()
        .map(|(index, sheet)| sheet_payload(sheet, index, &mut validation_sidecar, &mut sparkline_sidecar, &mut conditional_format_sidecar))
        .collect::<Vec<_>>();
    for sheet in &mut sheets {
        validation_sidecar.restore(sheet);
        dimension_sidecar.restore(sheet);
        outline_sidecar.restore(sheet);
        if let Some(details) = outline_sidecar.failure.as_deref() {
            return Err(outline_import_failure(details));
        }
        if let Some(mut native_charts) = pivot_sidecar.native_charts.remove(&sheet.name) {
            let mut previous = std::mem::take(&mut sheet.charts);
            for native in &mut native_charts {
                if let Some(previous) = previous.iter_mut().find(|previous| previous.index == native.index && previous.series.len() == native.series.len()
                    && previous.series.iter().zip(&native.series).all(|(old, new)| old.values == new.values && old.categories == new.categories)) {
                    if let Some(mut old) = previous.drawing.take() {
                        // rxls contributes cache/style metadata, while the native
                        // drawing reader owns the actual anchor and offsets.
                        if let Some(anchor) = native.drawing.take() {
                            old.from_cell = anchor.from_cell;
                            old.to_cell = anchor.to_cell;
                            old.from_offset_emu = anchor.from_offset_emu;
                            old.to_offset_emu = anchor.to_offset_emu;
                            old.absolute_size_emu = anchor.absolute_size_emu.or(old.absolute_size_emu);
                            old.name = anchor.name.or(old.name);
                            old.alt_text = anchor.alt_text.or(old.alt_text);
                            old.behavior = anchor.behavior;
                        }
                        native.drawing = Some(old);
                    }
                }
            }
            sheet.charts = native_charts;
        }
        for chart in &mut sheet.charts {
            let chart_key = (sheet.name.clone(), chart.index);
            chart.pivot_table_id = pivot_sidecar
                .chart_pivot_ids
                .get(&chart_key)
                .cloned();
        }
    }
    validation_sidecar.warn_unmatched_rules();
    let mut payload = WorkbookPayload {
        formula_calendar_source: None,
        sheets,
        pivot_tables: pivot_sidecar.pivot_tables,
        tables: table_sidecar.tables,
        date1904: workbook.date1904,
        active_sheet: metadata.active_sheet,
        defined_names: workbook
            .defined_names
            .iter()
            .map(|(name, refers_to)| crate::model::DefinedNamePayload {
                name: name.clone(),
                refers_to: refers_to.clone(),
            })
            .collect(),
        local_defined_names: metadata
            .local_defined_names
            .iter()
            .map(|name| LocalDefinedNamePayload {
                sheet: name.sheet.clone(),
                name: name.name.clone(),
                refers_to: name.refers_to.clone(),
            })
            .collect(),
        properties: WorkbookPropertiesPayload {
            title: metadata.properties.title.clone(),
            subject: metadata.properties.subject.clone(),
            creator: metadata.properties.creator.clone(),
            keywords: metadata.properties.keywords.clone(),
            description: metadata.properties.description.clone(),
            last_modified_by: metadata.properties.last_modified_by.clone(),
            company: metadata.properties.company.clone(),
            created: metadata.properties.created.clone(),
        },
        structure_protected: metadata.structure_protected,
        text_truncated: metadata.text_truncated,
    };
    let mut warnings = validation_sidecar.warnings;
    warnings.extend(sparkline_sidecar.warnings);
    warnings.extend(pivot_sidecar.warnings);
    warnings.extend(conditional_format_sidecar.warnings);
    warnings.extend(table_sidecar.warnings);
    warnings.extend(dimension_sidecar.warnings);
    warnings.extend(outline_sidecar.warnings);
    if format == "xlsx" {
        warnings.extend(crate::formula_calendar_sidecar::restore(bytes, &mut payload).map_err(|details| ParseFailure::new(
            if details.contains("LIMIT") || details.contains("limit") { crate::model::ParseErrorCode::ParserFailure } else { crate::model::ParseErrorCode::InvalidOfficeFile },
            "Could not preserve imported formula calendars from the XLSX file", Some(details),
        ))?);
    }
    if workbook.text_truncated {
        warnings.push("Some text cells were omitted after the parser allocation limit".into());
    }
    let diagnostics = ParserDiagnostics::spreadsheet(&payload, warnings);
    Ok(ParseEnvelope {
        schema_version: SCHEMA_VERSION,
        parser_version: env!("CARGO_PKG_VERSION").to_owned(),
        format,
        file_name: file_name.to_owned(),
        size: bytes.len(),
        model: "spreadsheet".into(),
        payload: DocumentModel::Spreadsheet(payload),
        diagnostics,
    })
}

fn outline_import_failure(details: &str) -> ParseFailure {
    let resource_limit = details.to_ascii_lowercase().contains("budget");
    ParseFailure::new(
        if resource_limit { crate::model::ParseErrorCode::ParserFailure } else { crate::model::ParseErrorCode::InvalidOfficeFile },
        if resource_limit { "Worksheet outlines exceed the parser resource limit" } else { "Could not preserve worksheet outlines from the XLSX file" },
        Some(details.to_owned()),
    )
}

fn sheet_payload(sheet: &rxls::Sheet, index: usize, validation_sidecar: &mut ValidationSidecar, sparkline_sidecar: &mut SparklineSidecar, conditional_format_sidecar: &mut ConditionalFormatSidecar) -> SheetPayload {
    let metadata = sheet.metadata();
    let mut cells: Vec<CellPayload> = sheet
        .display_cells()
        .map(|cell| {
            let (value, formula) = cell_value(cell.value);
            CellPayload {
                row: cell.row,
                col: cell.col,
                value,
                data_type: cell_type(cell.value).into(),
                cached_data_type: match cell.value {
                    Cell::Formula { cached, .. } => Some(cell_type(cached).into()),
                    _ => None,
                },
                display: cell.formatted.to_owned(),
                formula,
                style: sheet
                    .resolved_cell_style(cell.row, cell.col)
                    .as_ref()
                    .map(style_payload),
                rich_text: cell
                    .rich_text
                    .map(|runs| {
                        runs.iter()
                            .map(|run| RichTextRunPayload {
                                text: run.text.clone(),
                                font: font_payload(&run.font),
                            })
                            .collect()
                    })
                    .unwrap_or_default(),
                hyperlink: cell.hyperlink.map(str::to_owned),
                sparkline: sparkline_sidecar.take(sheet.name.as_str(), &format!("{}{}", column_label(cell.col), cell.row + 1)),
            }
        })
        .collect();
    let mut known = cells.iter().map(|cell: &CellPayload| (cell.row, cell.col)).collect::<std::collections::HashSet<_>>();
    for (target, sparkline) in sparkline_sidecar.take_sheet(sheet.name.as_str()) {
        let Some((row, col)) = parse_sparkline_target(&target) else { continue };
        if known.insert((row, col)) {
            cells.push(CellPayload {
                row, col, value: Value::Null, data_type: "blank".into(), cached_data_type: None,
                display: String::new(), formula: None, style: None, rich_text: Vec::new(),
                hyperlink: None, sparkline: Some(sparkline),
            });
        }
    }
    let dimensions =
        metadata.dimensions.map(
            |(first_row, first_col, last_row, last_col)| DimensionPayload {
                first_row,
                first_col,
                last_row,
                last_col,
            },
        );
    let merged_ranges = metadata
        .merged_ranges
        .iter()
        .map(|&(first_row, first_col, last_row, last_col)| RangePayload {
            first_row,
            first_col,
            last_row,
            last_col,
        })
        .collect();
    let axis = crate::sheet_dimensions::sheet_dimensions(sheet);
    let view = metadata.sheet_view;
    SheetPayload {
        name: sheet.name.clone(),
        is_worksheet: sheet.is_worksheet,
        visible: format!("{:?}", metadata.visible).to_ascii_lowercase(),
        cells,
        dimensions,
        merged_ranges,
        column_widths: axis.columns,
        row_heights: axis.rows,
        axis_unit: "cssPixel".into(),
        column_width_spans: axis.column_spans,
        row_height_spans: axis.row_spans,
        row_auto_height_spans: axis.row_modes,
        row_auto_heights: axis.automatic_rows,
        default_row_auto_height: Some(axis.default_row_automatic),
        row_count: None,
        column_count: None,
        source_axis_dimensions: axis.source,
        default_column_width: axis.default_column,
        default_row_height: axis.default_row,
        gridlines: Some(!view.hide_gridlines),
        freeze_panes: view
            .freeze
            .map(|(rows, columns)| FreezePanePayload { rows, columns }),
        zoom: view.zoom,
        show_headers: view.show_headers,
        right_to_left: view.right_to_left,
        hidden_rows: sheet.hidden_rows().iter().copied().collect(),
        hidden_columns: sheet.hidden_columns().iter().copied().collect(),
        row_hidden_spans: Vec::new(),
        column_hidden_spans: Vec::new(),
        outline: None,
        source_outline: serde_json::Value::Null,
        auto_filter: None,
        source_auto_filter: None,
        hyperlinks: metadata
            .hyperlinks
            .iter()
            .map(|&(row, col, ref target)| HyperlinkPayload {
                row,
                col,
                target: target.clone(),
            })
            .collect(),
        conditional_formats: conditional_format_sidecar.take(index),
        data_validations: sheet
            .data_validations()
            .iter()
            .enumerate()
            .map(|(index, validation)| {
                data_validation_payload(index, validation, sheet.name.as_str(), validation_sidecar)
            })
            .collect(),
        style_fidelity: format!("{:?}", sheet.style_fidelity()).to_ascii_lowercase(),
        style_losses: sheet
            .style_losses()
            .iter()
            .map(|loss| StyleLossPayload {
                kind: format!("{:?}", loss.kind).to_ascii_lowercase(),
                occurrences: loss.occurrences,
            })
            .collect(),
        charts: sheet
            .charts()
            .iter()
            .enumerate()
            .map(|(index, chart)| chart_payload(index, chart, sheet.drawing_metadata()))
            .collect(),
    }
}

fn parse_sparkline_target(value: &str) -> Option<(u32, u16)> {
    let split = value.find(|character: char| character.is_ascii_digit())?;
    let (letters, digits) = value.split_at(split);
    let mut column = 0u32;
    for character in letters.bytes() {
        if !character.is_ascii_alphabetic() { return None }
        column = column.checked_mul(26)?.checked_add(u32::from(character.to_ascii_uppercase() - b'A' + 1))?;
    }
    let row = digits.parse::<u32>().ok()?.checked_sub(1)?;
    Some((row, u16::try_from(column.checked_sub(1)?).ok()?))
}

fn column_label(index: u16) -> String {
    let mut current = u32::from(index);
    let mut label = String::new();
    loop {
        label.insert(0, char::from_u32(current % 26 + 65).unwrap());
        if current < 26 { break }
        current = current / 26 - 1;
    }
    label
}

fn data_validation_payload(
    index: usize,
    validation: &rxls::DataValidation,
    sheet_name: &str,
    validation_sidecar: &mut ValidationSidecar,
) -> DataValidationPayload {
    let (start_row, start_col, end_row, end_col) = validation.sqref;
    let supplement = validation_sidecar.take(sheet_name, (start_row, start_col, end_row, end_col));
    let (prompt_title, prompt) = validation
        .prompt
        .as_ref()
        .map(|(title, message)| (Some(title.clone()), Some(message.clone())))
        .unwrap_or((None, None));
    let (error_title, error) = validation
        .error
        .as_ref()
        .map(|(title, message)| (Some(title.clone()), Some(message.clone())))
        .unwrap_or((None, None));
    DataValidationPayload {
        id: format!("validation-{}", index + 1),
        validation_type: match validation.kind {
            rxls::DvKind::List => "list",
            rxls::DvKind::Whole => "whole",
            rxls::DvKind::Decimal => "decimal",
            rxls::DvKind::Date => "date",
            rxls::DvKind::Time => "time",
            rxls::DvKind::TextLength => "textLength",
            rxls::DvKind::Custom => "custom",
        }
        .into(),
        ranges: vec![DataValidationRangePayload {
            start_row,
            start_col: u32::from(start_col),
            end_row,
            end_col: u32::from(end_col),
        }],
        formula_anchor: supplement.as_ref().and_then(|value| value.formula_anchor).map(|(row, col)| crate::model::DataValidationFormulaAnchorPayload { row, col: u32::from(col) }),
        operator: match validation.operator {
            rxls::DvOp::Between => "between",
            rxls::DvOp::NotBetween => "notBetween",
            rxls::DvOp::Equal => "equal",
            rxls::DvOp::NotEqual => "notEqual",
            rxls::DvOp::GreaterThan => "greaterThan",
            rxls::DvOp::LessThan => "lessThan",
            rxls::DvOp::GreaterThanOrEqual => "greaterThanOrEqual",
            rxls::DvOp::LessThanOrEqual => "lessThanOrEqual",
        }
        .into(),
        formula1: validation.formula1.clone(),
        formula2: validation.formula2.clone(),
        allow_blank: validation.allow_blank,
        show_drop_down: supplement.as_ref().and_then(|value| value.show_drop_down),
        show_error_message: validation.show_error_message,
        error_style: supplement.and_then(|value| value.error_style),
        error_title,
        error,
        show_input_message: validation.show_input_message,
        prompt_title,
        prompt,
        extensions: serde_json::Map::new(),
    }
}

fn chart_payload(
    index: usize,
    chart: &rxls::Chart,
    drawings: &[rxls::DrawingMetadata],
) -> ChartPayload {
    let drawing = drawings
        .iter()
        .find(|item| item.kind == rxls::DrawingObjectKind::Chart && item.object_index == index)
        .map(chart_drawing_payload);
    ChartPayload {
        index,
        kind: format!("{:?}", chart.kind).to_ascii_lowercase(),
        title: chart.title.clone(),
        subtitle: None,
        series: chart
            .series
            .iter()
            .map(|series| ChartSeriesPayload {
                name: series.name.clone(),
                categories: series.categories.clone(),
                values: series.values.clone(),
                bubble_sizes: series.bubble_sizes.clone(),
                native_kind: None,
                options: None,
                trendlines: Vec::new(),
                error_bars: Vec::new(),
            })
            .collect(),
        legend: chart.legend,
        legend_position: None,
        legend_font: None,
        options: None,
        authored: None,
        unsupported_reasons: Vec::new(),
        office_source: None,
        data_labels: chart.data_labels,
        x_axis_title: chart.x_axis_title.clone(),
        y_axis_title: chart.y_axis_title.clone(),
        x_axis: None,
        y_axis: None,
        from: AnchorCellPayload {
            row: chart.from.0,
            col: chart.from.1,
        },
        to: AnchorCellPayload {
            row: chart.to.0,
            col: chart.to.1,
        },
        drawing,
        pivot_table_id: None,
    }
}

fn chart_drawing_payload(metadata: &rxls::DrawingMetadata) -> ChartDrawingPayload {
    let anchor = |value: Option<(u32, u16)>| value.map(|(row, col)| AnchorCellPayload { row, col });
    let offsets = |value: Option<(i64, i64)>| value.map(|(x, y)| [x, y]);
    let size = |value: Option<(u64, u64)>| value.map(|(width, height)| [width, height]);
    let crop = metadata.crop.map(|value| {
        [
            value.left_ppm,
            value.top_ppm,
            value.right_ppm,
            value.bottom_ppm,
        ]
    });
    let cached_points = |points: &[rxls::ChartCachedPoint]| {
        points
            .iter()
            .map(|point| ChartCachedPointPayload {
                index: point.index,
                value: point.value.clone(),
            })
            .collect()
    };
    let caches = metadata
        .chart_series_caches
        .iter()
        .map(|cache| ChartSeriesCachePayload {
            name: cached_points(&cache.name),
            categories: cached_points(&cache.categories),
            values: cached_points(&cache.values),
            bubble_sizes: cached_points(&cache.bubble_sizes),
        })
        .collect();
    let styles = metadata
        .chart_series_styles
        .iter()
        .map(|style| ChartSeriesStylePayload {
            marker: format!("{:?}", style.marker).to_ascii_lowercase(),
            marker_size: style.marker_size,
            line_visible: style.line_visible,
            line_color: style.line_color.map(color_hex),
            line_width_emu: style.line_width_emu,
            losses: style
                .losses
                .iter()
                .map(|loss| format!("{:?}", loss).to_ascii_lowercase())
                .collect(),
        })
        .collect();
    ChartDrawingPayload {
        object_index: metadata.object_index,
        from_cell: anchor(metadata.from_cell),
        to_cell: anchor(metadata.to_cell),
        from_offset_emu: offsets(metadata.from_offset_emu),
        to_offset_emu: offsets(metadata.to_offset_emu),
        absolute_size_emu: size(metadata.absolute_size_emu),
        crop,
        rotation_mdeg: metadata.rotation_mdeg,
        z_order: metadata.z_order,
        alt_text: metadata.alt_text.clone(),
        name: metadata.name.clone(),
        behavior: format!("{:?}", metadata.behavior).to_ascii_lowercase(),
        chart_palette: metadata
            .chart_palette
            .iter()
            .copied()
            .map(color_hex)
            .collect(),
        chart_default_latin_font_family: metadata.chart_default_latin_font_family.clone(),
        chart_category_axis_visible: metadata.chart_category_axis_visible,
        chart_category_axis_shifted: metadata.chart_category_axis_shifted,
        chart_value_axis_visible: metadata.chart_value_axis_visible,
        chart_category_major_gridlines: metadata.chart_category_major_gridlines,
        chart_value_major_gridlines: metadata.chart_value_major_gridlines,
        chart_bar_direction: format!("{:?}", metadata.chart_bar_direction).to_ascii_lowercase(),
        chart_frame_fill: format!("{:?}", metadata.chart_frame_fill).to_ascii_lowercase(),
        chart_frame_style_losses: metadata
            .chart_frame_style_losses
            .iter()
            .map(|loss| format!("{:?}", loss).to_ascii_lowercase())
            .collect(),
        chart_unsupported_reasons: metadata
            .chart_unsupported_reasons
            .iter()
            .map(|reason| format!("{:?}", reason).to_ascii_lowercase())
            .collect(),
        chart_series_caches: caches,
        chart_series_styles: styles,
        chart_text_styles: json!({
            "chartTitle": chart_text_style(metadata.chart_text_styles.chart_title.as_ref()),
            "categoryAxisTitle": chart_text_style(metadata.chart_text_styles.category_axis_title.as_ref()),
            "valueAxisTitle": chart_text_style(metadata.chart_text_styles.value_axis_title.as_ref()),
            "legend": chart_text_style(metadata.chart_text_styles.legend.as_ref()),
            "categoryAxisLabels": chart_text_style(metadata.chart_text_styles.category_axis_labels.as_ref()),
            "valueAxisLabels": chart_text_style(metadata.chart_text_styles.value_axis_labels.as_ref()),
            "dataLabels": chart_text_style(metadata.chart_text_styles.data_labels.as_ref()),
        }),
    }
}

fn chart_text_style(style: Option<&rxls::ChartTextStyle>) -> Value {
    let Some(style) = style else {
        return Value::Null;
    };
    json!({
        "latinFontFamily": style.latin_font_family,
        "sizeHundredthsOfPoint": style.size_hundredths_of_point,
        "color": color_hex(style.color),
        "bold": style.bold,
        "italic": style.italic,
        "underline": style.underline,
        "strikethrough": style.strikethrough,
        "kerningMinimumHundredthsOfPoint": style.kerning_minimum_hundredths_of_point,
        "rotationDegrees": style.rotation_degrees,
    })
}

fn cell_value(cell: &Cell) -> (Value, Option<String>) {
    match cell {
        Cell::Text(value) => (Value::String(value.clone()), None),
        Cell::Number(value) | Cell::Date(value) => (json!(value), None),
        Cell::Bool(value) => (json!(value), None),
        Cell::Error(value) => (json!({ "error": value }), None),
        Cell::Formula { formula, cached } => {
            let (value, _) = cell_value(cached);
            (value, Some(formula.clone()))
        }
    }
}

fn cell_type(cell: &Cell) -> &'static str {
    match cell {
        Cell::Text(_) => "text",
        Cell::Number(_) => "number",
        Cell::Date(_) => "date",
        Cell::Bool(_) => "boolean",
        Cell::Error(_) => "error",
        Cell::Formula { .. } => "formula",
    }
}

fn style_payload(style: &rxls::CellStyle) -> StylePayload {
    let font = style.font.as_ref().map(font_payload);
    let fill = style
        .pattern_fill
        .as_ref()
        .and_then(|fill| fill.foreground.or(fill.background))
        .map(color_hex)
        .or_else(|| style.fill.map(color_hex));
    let (fill_background, fill_foreground) = style
        .pattern_fill
        .as_ref()
        .map(|fill| {
            (
                fill.background.map(color_hex),
                fill.foreground.map(color_hex),
            )
        })
        .unwrap_or((None, None));
    let fill_pattern = style
        .pattern_fill
        .as_ref()
        .map(|fill| format!("{:?}", fill.pattern).to_ascii_lowercase());
    let border = style.border.as_ref().map(border_payload);
    let protection = style
        .protection
        .as_ref()
        .map(|protection| ProtectionPayload {
            locked: protection.locked,
            hidden: protection.hidden,
        });
    let (rotation, rotation_vertical, indent, shrink_to_fit) = style
        .align
        .as_ref()
        .map(|align| {
            let (rotation, rotation_vertical) = normalize_rotation(align.rotation);
            (
                rotation,
                rotation_vertical,
                align.indent,
                align.shrink_to_fit,
            )
        })
        .unwrap_or((0, false, 0, false));
    StylePayload {
        font,
        fill,
        number_format: style.num_fmt.clone(),
        horizontal: style.align.as_ref().and_then(|align| {
            align
                .horizontal
                .map(|value| format!("{:?}", value).to_ascii_lowercase())
        }),
        vertical: style.align.as_ref().and_then(|align| {
            align
                .vertical
                .map(|value| format!("{:?}", value).to_ascii_lowercase())
        }),
        wrap: style
            .align
            .as_ref()
            .map(|align| align.wrap)
            .unwrap_or(false),
        rotation,
        rotation_vertical,
        indent,
        shrink_to_fit,
        fill_pattern,
        fill_background,
        fill_foreground,
        border,
        protection,
    }
}

fn normalize_rotation(rotation: i16) -> (i16, bool) {
    if rotation == 255 {
        (0, true)
    } else {
        (rotation, false)
    }
}

fn font_payload(font: &rxls::Font) -> FontPayload {
    FontPayload {
        name: font.name.clone(),
        size_pt: font.size_pt,
        bold: font.bold,
        italic: font.italic,
        underline: font.underline,
        strikethrough: font.strikethrough,
        color: font.color.map(color_hex),
        script: match font.script {
            rxls::FormatScript::None => None,
            script => Some(format!("{:?}", script).to_ascii_lowercase()),
        },
    }
}

fn border_payload(border: &rxls::Border) -> BorderPayload {
    BorderPayload {
        left: format!("{:?}", border.left).to_ascii_lowercase(),
        right: format!("{:?}", border.right).to_ascii_lowercase(),
        top: format!("{:?}", border.top).to_ascii_lowercase(),
        bottom: format!("{:?}", border.bottom).to_ascii_lowercase(),
        color: border.color.map(color_hex),
        left_color: border.left_color.map(color_hex),
        right_color: border.right_color.map(color_hex),
        top_color: border.top_color.map(color_hex),
        bottom_color: border.bottom_color.map(color_hex),
    }
}

fn color_hex(color: rxls::Color) -> String {
    let [red, green, blue] = color.as_rgb();
    format!("#{red:02X}{green:02X}{blue:02X}")
}

pub fn detect_format(file_name: &str, bytes: &[u8]) -> String {
    let file_name = file_name.to_ascii_lowercase();
    if file_name.ends_with(".docx") {
        "docx".into()
    } else if file_name.ends_with(".doc") {
        "doc".into()
    } else if file_name.ends_with(".odt") {
        "odt".into()
    } else if file_name.ends_with(".xlsx") {
        "xlsx".into()
    } else if file_name.ends_with(".xlsb") {
        "xlsb".into()
    } else if file_name.ends_with(".xls") {
        "xls".into()
    } else if file_name.ends_with(".ods") {
        "ods".into()
    } else if bytes.starts_with(b"PK") {
        "xlsx".into()
    } else {
        "unknown".into()
    }
}

fn validate_input(bytes: &[u8], file_name: &str) -> Result<(), ParseFailure> {
    if bytes.is_empty() {
        return Err(ParseFailure::invalid_input("The selected file is empty"));
    }
    if file_name.trim().is_empty() {
        return Err(ParseFailure::invalid_input(
            "A file name is required for format detection",
        ));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::{detect_format, normalize_rotation, parse_office_bytes};
    use duke_pages::{create_minimal_docx_package, package_to_bytes};
    use rxls::{Chart, ChartKind, DataValidation, DvKind, DvOp, Series, Workbook};
    use serde_json::{Value, json};
    use std::io::{Read, Write};
    use zip::{ZipArchive, ZipWriter, write::SimpleFileOptions};

    #[test]
    fn reads_current_section_story_distances_not_historical_properties() {
        let xml = r#"<w:sectPr><w:sectPrChange><w:sectPr><w:pgMar w:header="720"/></w:sectPr></w:sectPrChange><w:pgMar w:header="375" w:footer="0"/></w:sectPr>"#;
        assert_eq!(super::story_distances(xml), (Some(25.0), Some(0.0)));
        assert_eq!(super::story_distances("<w:sectPr/>"), (None, None));
    }

    #[test]
    fn detects_supported_extensions() {
        assert_eq!(detect_format("report.docx", &[]), "docx");
        assert_eq!(detect_format("sheet.xlsx", &[]), "xlsx");
        assert_eq!(detect_format("unknown", b"PK\x03\x04"), "xlsx");
    }

    #[test]
    fn maps_the_ooxml_vertical_text_sentinel_to_a_vertical_style() {
        assert_eq!(normalize_rotation(255), (0, true));
        assert_eq!(normalize_rotation(-45), (-45, false));
    }

    #[test]
    fn docx_section_break_is_inserted_before_its_body_start_block() {
        let document_xml = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:sectPr><w:type w:val="continuous"/></w:sectPr></w:pPr><w:r><w:t>First section</w:t></w:r></w:p><w:p><w:r><w:t>Second section</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>"#;
        let package = create_minimal_docx_package(Some(document_xml));
        let bytes = package_to_bytes(&package).expect("the DOCX fixture must serialize");
        let envelope = parse_office_bytes(&bytes, "section-break.docx")
            .expect("the section break fixture must parse");
        let payload = serde_json::to_value(envelope).expect("parse envelope must serialize");
        let blocks = payload["payload"]["data"]["blocks"]
            .as_array()
            .expect("the document blocks must reach the payload");
        let marker_index = blocks
            .iter()
            .position(|block| block["details"]["sectionBreak"]["type"] == "continuous");
        assert!(
            marker_index.is_some(),
            "the section marker must be represented as a body block: {payload:#}"
        );
        let marker_index = marker_index.unwrap();

        assert_eq!(blocks[marker_index - 1]["content"], "First section");
        assert_eq!(blocks[marker_index + 1]["content"], "Second section");
        assert!(
            payload["diagnostics"]["capabilities"]
                .as_array()
                .unwrap()
                .iter()
                .any(|capability| capability == "sectionBreaks")
        );
    }

    #[test]
    fn docx_standalone_page_and_column_breaks_reach_the_document_payload() {
        for (word_type, detail_key) in [("page", "pageBreak"), ("column", "columnBreak")] {
            let document_xml = format!(
                r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Before</w:t></w:r></w:p><w:p><w:r><w:br w:type="{word_type}"/></w:r></w:p><w:p><w:r><w:t>After</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"#
            );
            let package = create_minimal_docx_package(Some(&document_xml));
            let bytes = package_to_bytes(&package).expect("the DOCX fixture must serialize");
            let envelope = parse_office_bytes(&bytes, "manual-break.docx")
                .expect("the manual break fixture must parse");
            let payload = serde_json::to_value(envelope).expect("parse envelope must serialize");
            let blocks = payload["payload"]["data"]["blocks"]
                .as_array()
                .expect("document blocks must reach the payload");

            assert_eq!(blocks[0]["content"], "Before");
            assert_eq!(blocks[1]["type"], "text");
            assert_eq!(blocks[1]["details"][detail_key], true, "{payload:#}");
            assert_eq!(blocks[2]["content"], "After");
        }
    }

    #[test]
    fn docx_roundtrip_tags_for_code_blocks_and_latex_reach_structured_payload() {
        let document_xml = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
            <w:p><w:pPr><w:pStyle w:val="OfficeCodeBlock"/></w:pPr><w:bookmarkStart w:id="1" w:name="office_code_v1_typescript_1"/><w:r><w:t>const x = 1</w:t><w:br/><w:t>return x</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>
            <w:p><w:sdt><w:sdtPr><w:tag w:val="office-latex:v1:inline"/><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>x^2</w:t></w:r></w:sdtContent></w:sdt></w:p>
            <w:sectPr/></w:body></w:document>"#;
        let package = create_minimal_docx_package(Some(document_xml));
        let bytes = package_to_bytes(&package).expect("DOCX fixture should serialize");
        let envelope =
            parse_office_bytes(&bytes, "special-content.docx").expect("DOCX should parse");
        let payload = serde_json::to_value(envelope).expect("parse envelope should serialize");
        let blocks = payload["payload"]["data"]["blocks"].as_array().unwrap();
        assert_eq!(blocks.len(), 2, "{payload:#}");
        assert_eq!(
            blocks[0]["details"]["bookmarks"][0]["name"], "office_code_v1_typescript_1",
            "{payload:#}"
        );
        assert_eq!(blocks[0]["content"], "const x = 1\nreturn x");
        assert_eq!(
            blocks[1]["children"][0]["details"]["field"]["tag"], "office-latex:v1:inline",
            "{payload:#}"
        );
        assert_eq!(blocks[1]["content"], "x^2");
    }

    #[test]
    fn writer_docx_section_headers_and_footers_reopen_with_references_and_settings() {
        let model = serde_json::json!({
            "kind": "document",
            "name": "section headers",
            "sectionSetup": { "marginHeader": 25.5, "marginFooter": 30 },
            "page": { "width": 794, "height": 1123, "marginTop": 72, "marginBottom": 72, "marginLeft": 72, "marginRight": 72 },
            "blocks": [
                { "id": "body-1", "type": "paragraph", "runs": [{ "text": "First section" }] },
                { "id": "section-2", "type": "sectionBreak", "sectionBreak": { "type": "nextPage" }, "runs": [] },
                { "id": "body-2", "type": "paragraph", "runs": [{ "text": "Second section" }] },
                { "id": "section-3", "type": "sectionBreak", "sectionBreak": { "type": "continuous", "setup": { "marginHeader": 0, "marginFooter": 200 } }, "runs": [] },
                { "id": "body-3", "type": "paragraph", "runs": [{ "text": "Third section" }] }
            ],
            "sectionHeaderFooters": {
                "first": {
                    "header": "Main & header",
                    "footer": "Main footer",
                    "firstPageHeader": "Dormant cover header",
                    "firstPageFooter": "Dormant cover footer",
                    "evenPageHeader": "Dormant even header",
                    "evenPageFooter": "Dormant even footer",
                    "pageFields": [
                        { "story": "footer", "field": "PAGE" },
                        { "story": "footer", "field": "NUMPAGES" },
                        { "story": "header", "field": "DATE", "offset": 4, "instruction": "DATE \\@ \"yyyy\"", "cachedResult": "2026" },
                        { "story": "header", "field": "PAGE", "offset": 4 },
                        { "story": "header", "field": "PAGE", "offset": 4 }
                    ]
                },
                "section-2": {
                    "header": "Section header",
                    "footer": "Section footer",
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
        let (bytes, warnings) = office_writer::export_docx_bytes(&model.to_string())
            .expect("writer should create a DOCX with section stories");
        assert!(warnings.is_empty(), "{warnings:?}");

        let envelope = parse_office_bytes(&bytes, "section-headers.docx")
            .expect("writer output should parse through the DOCX parser");
        let value = serde_json::to_value(envelope).expect("parser envelope serializes");
        let data = &value["payload"]["data"];
        let sections = data["sectionLayout"].as_array().expect("section layout is exposed");
        assert_eq!(sections.len(), 3);
        assert_eq!(sections[0]["marginHeader"].as_f64(), Some(383.0 / 15.0));
        assert_eq!(sections[0]["marginFooter"].as_f64(), Some(30.0));
        assert_eq!(sections[1]["marginHeader"].as_f64(), Some(48.0));
        assert_eq!(sections[2]["marginHeader"].as_f64(), Some(0.0));
        assert_eq!(sections[2]["marginFooter"].as_f64(), Some(200.0));
        assert_eq!(sections[0]["headerSections"][0]["referenceType"], "default");
        assert_eq!(sections[0]["headerSections"].as_array().unwrap().len(), 3);
        assert_eq!(sections[0]["footerSections"].as_array().unwrap().len(), 3);
        assert_eq!(sections[0]["headerSections"][1]["referenceType"], "first");
        assert_eq!(sections[0]["headerSections"][1]["storyText"]["text"], "Dormant cover header");
        assert_eq!(sections[0]["headerSections"][2]["referenceType"], "even");
        assert_eq!(sections[0]["headerSections"][2]["storyText"]["text"], "Dormant even header");
        assert!(!sections[0]["sectionPropertiesXml"].as_str().unwrap().contains("titlePg"));
        let header = &sections[0]["headerSections"][0]["storyText"];
        assert_eq!(header["text"], "Main & header");
        let date = &header["fields"][0];
        assert_eq!(date["field"], "DATE"); assert_eq!(date["offset"], 4);
        assert_eq!(date["instruction"], "DATE \\@ \"yyyy\""); assert_eq!(date["cachedResult"], "2026");
        assert_eq!(date["textStyle"]["fontSizePt"].as_f64(), Some(9.0));
        assert_eq!(header["fields"][1]["offset"], 4);
        assert_eq!(header["fields"][2]["field"], "PAGE");
        assert_eq!(sections[0]["footerSections"][0]["pageFields"], serde_json::json!(["PAGE", "NUMPAGES"]));
        assert_eq!(sections[1]["headerSections"][1]["referenceType"], "first");
        assert_eq!(sections[1]["headerSections"][2]["referenceType"], "even");
        assert!(sections[1]["sectionPropertiesXml"].as_str().unwrap().contains("titlePg"));
        assert!(sections[2]["headerSections"].as_array().unwrap().is_empty());
        assert!(sections[2]["footerSections"].as_array().unwrap().is_empty());
        assert_eq!(data["evenAndOddHeaders"], true);

        let stories = data["blocks"].as_array().unwrap();
        assert!(stories.iter().any(|block| {
            block["type"] == "header" && block["content"].as_str().is_some_and(|text| text.contains("Main") && text.contains("& header"))
        }));
        assert!(stories.iter().any(|block| {
            block["type"] == "footer" && block["content"] == "Even footer"
        }));
    }

    #[test]
    fn writer_title_subtitle_and_east_asian_font_reopen_as_independent_properties() {
        let model = json!({ "kind": "document", "name": "font-slots.docx", "blocks": [
            { "id": "p1", "type": "paragraph", "metadata": {"namedStyleType": "title"}, "runs": [{"text": "A中", "fontFamily": "Arial", "eastAsiaFontFamily": "SimSun"}] },
            { "id": "p2", "type": "paragraph", "metadata": {"namedStyleType": "subtitle"}, "runs": [{"text": "Subtitle"}] }
        ] });
        let (bytes, _) = office_writer::export_docx_bytes(&model.to_string()).expect("write DOCX");
        let parsed = parse_office_bytes(&bytes, "font-slots.docx").expect("parse DOCX");
        let payload = serde_json::to_value(parsed).unwrap();
        let blocks = &payload["payload"]["data"]["blocks"];
        assert_eq!(blocks[0]["details"]["styleId"], "Title");
        assert_eq!(blocks[1]["details"]["styleId"], "Subtitle");
        let style = &blocks[0]["details"]["inlines"][0]["style"];
        assert_eq!(style["fontFamilyAscii"], "Arial");
        assert_eq!(style["fontFamilyEastAsia"], "SimSun");
    }

    #[test]
    fn docx_notes_are_exposed_with_inline_reference_metadata() {
        let model = json!({
            "kind": "document",
            "name": "notes.docx",
            "page": { "width": 794, "height": 1123, "marginTop": 72, "marginBottom": 72, "marginLeft": 72, "marginRight": 72 },
            "blocks": [{ "id": "p1", "type": "paragraph", "runs": [
                { "text": "Claim" },
                { "text": "1", "metadata": { "note": { "type": "footnote", "noteId": "12" } } },
                { "text": " and " },
                { "text": "1", "metadata": { "note": { "type": "endnote", "noteId": "14" } } }
            ] }],
            "footnotes": { "12": { "id": "12", "text": "Footnote body" } },
            "endnotes": { "14": { "id": "14", "text": "Endnote body" } }
        });
        let (bytes, warnings) = office_writer::export_docx_bytes(&model.to_string()).expect("writer emits note parts");
        assert!(warnings.is_empty(), "{warnings:?}");

        let envelope = parse_office_bytes(&bytes, "notes.docx").expect("DOCX with notes parses");
        let value = serde_json::to_value(envelope).expect("parser envelope serializes");
        let data = &value["payload"]["data"];
        assert_eq!(data["footnotes"][0]["id"], "1");
        assert_eq!(data["footnotes"][0]["text"], "Footnote body");
        assert_eq!(data["endnotes"][0]["id"], "1");
        let inlines = data["blocks"][0]["details"]["inlines"].as_array().expect("inline runs");
        assert!(inlines.iter().any(|inline| inline["noteReference"]["kind"] == "footnote" && inline["noteReference"]["id"] == 1));
        assert!(inlines.iter().any(|inline| inline["noteReference"]["kind"] == "endnote" && inline["noteReference"]["id"] == 1));
    }

    #[test]
    fn docx_comment_threads_and_range_anchors_roundtrip_through_the_writer() {
        let model = json!({
            "kind": "document",
            "name": "comments.docx",
            "page": { "width": 794, "height": 1123, "marginTop": 72, "marginBottom": 72, "marginLeft": 72, "marginRight": 72 },
            "blocks": [{ "id": "commented-paragraph", "type": "paragraph", "runs": [{ "text": "Hello World" }] }],
            "comments": {
                "thread-1": {
                    "id": "thread-1",
                    "anchor": [{ "blockId": "commented-paragraph", "startOffset": 6, "endOffset": 11, "text": "World" }],
                    "referenceText": "World",
                    "messages": [
                        { "id": "message-1", "authorId": "alice", "authorName": "Alice", "text": "Please review this word.", "createdAt": "2026-01-01T00:00:00Z" },
                        { "id": "message-2", "authorId": "bob", "authorName": "Bob", "text": "Updated.", "createdAt": "2026-01-02T00:00:00Z" }
                    ],
                    "resolved": true
                }
            }
        });
        let (bytes, warnings) = office_writer::export_docx_bytes(&model.to_string()).expect("writer exports DOCX comment parts");
        assert!(warnings.is_empty(), "{warnings:?}");

        let envelope = parse_office_bytes(&bytes, "comments.docx").expect("writer output parses");
        let value = serde_json::to_value(envelope).expect("parser envelope serializes");
        let comments = value["payload"]["data"]["comments"].as_array().expect("comment payload is included");
        assert_eq!(comments.len(), 2);
        assert_eq!(comments[0]["text"], "Please review this word.");
        assert_eq!(comments[0]["resolved"], true);
        assert!(!comments[0]["targets"][0]["nodeId"].as_str().unwrap_or_default().is_empty());
        assert_eq!(comments[0]["targets"][0]["start"], 6);
        assert_eq!(comments[0]["targets"][0]["end"], 11);
        assert_eq!(comments[1]["parentId"], comments[0]["id"]);
        assert_eq!(comments[1]["text"], "Updated.");
    }

    #[test]
    fn xlsx_chart_series_titles_anchors_and_drawing_sidecar_reach_parse_payload() {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_sheet("Data");
        sheet.write(0, 0, "Month");
        sheet.write(0, 1, "Revenue");
        sheet.write(1, 0, "Jan");
        sheet.write(1, 1, 120.0);
        sheet.write(2, 0, "Feb");
        sheet.write(2, 1, 180.0);
        sheet.add_chart(
            Chart::new(ChartKind::Bar, (4, 3), (18, 12))
                .with_title("Monthly revenue")
                .with_x_axis_title("Month")
                .with_y_axis_title("Revenue")
                .with_legend(true)
                .with_data_labels(true)
                .with_series([Series::new("Data!$B$2:$B$3")
                    .with_name("Data!$B$1")
                    .with_categories("Data!$A$2:$A$3")]),
        );

        let bytes = workbook.to_xlsx();
        let envelope = parse_office_bytes(&bytes, "chart-fixture.xlsx")
            .expect("the generated xlsx fixture must parse");
        let payload = serde_json::to_value(envelope).expect("parse envelope must serialize");
        let chart = &payload["payload"]["data"]["sheets"][0]["charts"][0];

        assert_eq!(chart["kind"], "bar");
        assert_eq!(chart["title"], "Monthly revenue");
        assert_eq!(chart["legend"], true);
        assert_eq!(chart["dataLabels"], true);
        assert_eq!(chart["xAxisTitle"], "Month");
        assert_eq!(chart["yAxisTitle"], "Revenue");
        assert_eq!(chart["from"], serde_json::json!({ "row": 4, "col": 3 }));
        assert_eq!(chart["to"], serde_json::json!({ "row": 18, "col": 12 }));
        assert_eq!(chart["series"][0]["name"], "Data!$B$1");
        assert_eq!(chart["series"][0]["categories"], "Data!$A$2:$A$3");
        assert_eq!(chart["series"][0]["values"], "Data!$B$2:$B$3");
        assert!(chart["drawing"].is_object());
        assert_eq!(payload["diagnostics"]["stats"]["chartCount"], 1);
        assert!(
            payload["diagnostics"]["capabilities"]
                .as_array()
                .unwrap()
                .iter()
                .any(|capability| capability == "charts")
        );
    }

    #[test]
    fn xlsx_data_validation_rule_reaches_normalizer_contract_shape() {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_sheet("Data");
        sheet.add_data_validation(
            DataValidation::new((1, 0, 3, 0), DvKind::Whole, DvOp::Between, "2")
                .with_formula2("8")
                .with_allow_blank(false)
                .with_prompt("Enter quantity", "Use a whole number")
                .with_error("Invalid quantity", "Enter a value from 2 to 8"),
        );

        let bytes = workbook.to_xlsx();
        let envelope = parse_office_bytes(&bytes, "validation-fixture.xlsx")
            .expect("the generated xlsx fixture must parse");
        let payload = serde_json::to_value(envelope).expect("parse envelope must serialize");
        let validation = &payload["payload"]["data"]["sheets"][0]["dataValidations"][0];

        assert_eq!(validation["id"], "validation-1");
        assert_eq!(validation["type"], "whole");
        assert_eq!(validation["operator"], "between");
        assert_eq!(validation["formula1"], "2");
        assert_eq!(validation["formula2"], "8");
        assert_eq!(validation["allowBlank"], false);
        assert_eq!(validation["showInputMessage"], true);
        assert_eq!(validation["showErrorMessage"], true);
        assert_eq!(validation["promptTitle"], "Enter quantity");
        assert_eq!(validation["prompt"], "Use a whole number");
        assert_eq!(validation["errorTitle"], "Invalid quantity");
        assert_eq!(validation["error"], "Enter a value from 2 to 8");
        assert_eq!(
            validation["ranges"][0],
            serde_json::json!({ "startRow": 1, "startCol": 0, "endRow": 3, "endCol": 0 })
        );
        assert!(
            payload["diagnostics"]["capabilities"]
                .as_array()
                .unwrap()
                .iter()
                .any(|capability| capability == "dataValidations")
        );
    }

    #[test]
    fn xlsx_multi_range_list_validation_retains_dropdown_and_error_style() {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_sheet("Lists");
        sheet.add_data_validation(
            DataValidation::list((1, 1, 3, 1), "\"Red,Green\"")
                .with_prompt("Choose a color", "Select a listed value")
                .with_error("Invalid color", "Choose Red or Green"),
        );
        let bytes = add_xlsx_validation_attributes(&workbook.to_xlsx());

        let envelope = parse_office_bytes(&bytes, "validation-list-fixture.xlsx")
            .expect("the modified xlsx fixture must parse");
        let payload = serde_json::to_value(envelope).expect("parse envelope must serialize");
        let validations = payload["payload"]["data"]["sheets"][0]["dataValidations"]
            .as_array()
            .expect("the list validation must reach the payload");

        assert_eq!(validations.len(), 2);
        for validation in validations {
            assert_eq!(validation["type"], "list");
            assert_eq!(validation["formula1"], "\"Red,Green\"");
            assert_eq!(validation["showDropDown"], false);
            assert_eq!(validation["errorStyle"], "info");
            assert_eq!(validation["showInputMessage"], true);
            assert_eq!(validation["showErrorMessage"], true);
            assert_eq!(validation["promptTitle"], "Choose a color");
            assert_eq!(validation["errorTitle"], "Invalid color");
        }
        assert_eq!(validations[0]["ranges"][0]["startCol"], 1);
        assert_eq!(validations[0]["ranges"][0]["endRow"], 3);
        assert_eq!(validations[1]["ranges"][0]["startCol"], 3);
        assert_eq!(validations[1]["ranges"][0]["endRow"], 3);
    }

    #[test]
    fn xlsx_sparkline_extension_reaches_the_canonical_cell_payload() {
        let mut workbook = Workbook::new();
        workbook.add_sheet("Data");
        let bytes = add_xlsx_sparkline_extension(&workbook.to_xlsx());
        let envelope = parse_office_bytes(&bytes, "sparkline-fixture.xlsx").expect("sparkline workbook parses");
        let payload = serde_json::to_value(envelope).expect("payload serializes");
        let cell = &payload["payload"]["data"]["sheets"][0]["cells"];
        let sparkline = cell.as_array().unwrap().iter().find(|value| value["row"] == 0 && value["col"] == 5).expect("blank sparkline target is included as a cell");
        assert_eq!(sparkline["sparkline"]["type"], "column");
        assert_eq!(sparkline["sparkline"]["sourceRange"], "B1:E1");
        assert_eq!(sparkline["sparkline"]["color"], "#167A55");
        assert_eq!(sparkline["sparkline"]["negativeColor"], "#C43B3B");
        assert_eq!(sparkline["sparkline"]["highPointColor"], "#F0A000");
        assert_eq!(sparkline["sparkline"]["lowPointColor"], "#2B579A");
        assert_eq!(sparkline["sparkline"]["emptyCells"], "zero");
        assert_eq!(sparkline["sparkline"]["axis"], "custom");
        assert_eq!(sparkline["sparkline"]["min"], -5.0);
        assert_eq!(sparkline["sparkline"]["max"], 20.0);
        assert!(payload["diagnostics"]["capabilities"].as_array().unwrap().iter().any(|value| value == "sparklines"));
    }

    #[test]
    fn writer_conditional_format_xlsx_reopens_with_canonical_rules_and_styles() {
        let model = json!({
            "kind": "workbook", "name": "Conditional", "sheets": [{
                "id": "sheet-1", "name": "Scores", "cells": {},
                "conditionalFormats": [{
                    "id": "high-score", "type": "cellIs",
                    "ranges": [
                        { "startRow": 0, "startCol": 0, "endRow": 4, "endCol": 0 },
                        { "startRow": 0, "startCol": 2, "endRow": 4, "endCol": 2 }
                    ],
                    "operator": "between", "formula1": "80", "formula2": "100",
                    "style": { "fill": "#00AA55", "fontColor": "#FFFFFF", "bold": true },
                    "priority": 1, "stopIfTrue": true
                }]
            }]
        });
        let (bytes, warnings) = office_writer::export_xlsx_bytes(&model.to_string())
            .expect("writer exports conditional formatting");
        assert!(warnings.is_empty(), "{warnings:?}");
        let envelope = parse_office_bytes(&bytes, "conditional-format.xlsx")
            .expect("writer output with conditional formatting parses");
        let payload = serde_json::to_value(envelope).expect("parse envelope serializes");
        let sheet = &payload["payload"]["data"]["sheets"][0];
        let rule = &sheet["conditionalFormats"][0];
        assert_eq!(rule["id"], "high-score");
        assert_eq!(rule["type"], "cellIs");
        assert_eq!(rule["operator"], "between");
        assert_eq!(rule["formula1"], "80");
        assert_eq!(rule["formula2"], "100");
        assert_eq!(rule["priority"], 1);
        assert_eq!(rule["stopIfTrue"], true);
        assert_eq!(rule["style"], json!({ "fill": "#00AA55", "fontColor": "#FFFFFF", "bold": true }));
        assert_eq!(rule["ranges"].as_array().unwrap().len(), 2);
        assert!(payload["diagnostics"]["capabilities"].as_array().unwrap().iter().any(|value| value == "conditionalFormatting"));
    }

    #[test]
    fn writer_sparkline_xlsx_reopens_with_canonical_options() {
        let model = serde_json::json!({
            "kind": "workbook", "name": "Sparklines", "sheets": [{ "id": "sheet-1", "name": "Data", "cells": {
                "F1": { "id": "0:5", "row": 0, "col": 5, "value": null, "formula": null, "display": "", "sparkline": {
                    "type": "line", "sourceRange": "B1:E1", "color": "#167A55", "negativeColor": "#C43B3B",
                    "highPointColor": "#F0A000", "lowPointColor": "#2B579A", "emptyCells": "connect", "axis": "custom", "min": -5.0, "max": 20.0
                } }
            } }]
        });
        let (bytes, warnings) = office_writer::export_xlsx_bytes(&model.to_string()).expect("writer produces XLSX");
        assert!(warnings.is_empty(), "{warnings:?}");
        let envelope = parse_office_bytes(&bytes, "writer-sparkline.xlsx").expect("writer XLSX reopens in parser");
        let payload = serde_json::to_value(envelope).expect("payload serializes");
        let cells = payload["payload"]["data"]["sheets"][0]["cells"].as_array().unwrap();
        let target = cells.iter().find(|cell| cell["row"] == 0 && cell["col"] == 5).expect("target sparkline cell is retained");
        assert_eq!(target["sparkline"]["type"], "line");
        assert_eq!(target["sparkline"]["sourceRange"], "B1:E1");
        assert_eq!(target["sparkline"]["emptyCells"], "connect");
        assert_eq!(target["sparkline"]["axis"], "custom");
        assert_eq!(target["sparkline"]["min"], -5.0);
        assert_eq!(target["sparkline"]["max"], 20.0);
        assert!(payload["diagnostics"]["capabilities"].as_array().unwrap().iter().any(|value| value == "sparklines"));
    }

    #[test]
    fn writer_xlsx_tables_reopen_with_stable_names_ranges_columns_and_options() {
        let model = serde_json::json!({
            "kind":"workbook","name":"Tables","sheets":[{"id":"sheet-1","name":"Data","cells":{}}],
            "tables":[{"id":"table-stable-id","name":"Orders","sheetId":"sheet-1","range":{"startRow":0,"startCol":0,"endRow":4,"endCol":1},"columns":[{"id":"c1","name":"Product"},{"id":"c2","name":"Amount"}],"showHeaderRow":true,"showTotalsRow":true,"styleName":"TableStyleMedium4"}]
        });
        let (bytes, warnings) = office_writer::export_xlsx_bytes(&model.to_string()).expect("writer emits table XLSX");
        assert!(warnings.is_empty(), "{warnings:?}");
        let envelope = parse_office_bytes(&bytes, "tables.xlsx").expect("table XLSX parses");
        let payload = serde_json::to_value(envelope).expect("payload serializes");
        let tables = payload["payload"]["data"]["tables"].as_array().expect("table definitions are present");
        assert_eq!(tables.len(), 1);
        assert_eq!(tables[0]["name"], "Orders");
        assert_eq!(tables[0]["sheetName"], "Data");
        assert_eq!(tables[0]["range"], json!({"startRow":0,"startCol":0,"endRow":4,"endCol":1}));
        assert_eq!(tables[0]["columns"][1]["name"], "Amount");
        assert_eq!(tables[0]["showTotalsRow"], true);
        assert_eq!(tables[0]["styleName"], "TableStyleMedium4");
    }

    #[test]
    fn writer_pivot_chart_xlsx_reopens_with_canonical_pivot_model_and_link() {
        let model = serde_json::json!({
            "kind":"workbook", "name":"Pivot", "sheets":[
                {"id":"data","name":"Data","cells":{
                    "A1":{"value":"Region"},"B1":{"value":"Sales"},
                    "A2":{"value":"East"},"B2":{"value":10},
                    "A3":{"value":"West"},"B3":{"value":20}
                }},
                {"id":"pivot","name":"Pivot View","cells":{
                    "A1":{"value":"Region"},"B1":{"value":"Sales"},
                    "A2":{"value":"East"},"B2":{"value":10},
                    "A3":{"value":"West"},"B3":{"value":20}
                }}
            ],
            "pivotTables":[{"id":"pivot-source-id",
                "source":{"unitId":"workbook","subUnitId":"data","sheetName":"Data","range":{"startRow":0,"startCol":0,"endRow":2,"endCol":1}},
                "anchor":{"unitId":"workbook","subUnitId":"pivot","row":0,"col":0},
                "outputRange":{"startRow":0,"startCol":0,"endRow":2,"endCol":1},
                "fields":[{"id":"0","sourceColumn":0,"name":"Region","area":"row","position":0},{"id":"1","sourceColumn":1,"name":"Sales","area":"value","position":0,"aggregation":"sum","displayName":"Revenue"}],
                "options":{"repeatRowLabels":true,"showRowGrandTotal":true,"showColGrandTotal":true,"showRowSubtotals":true,"showColSubtotals":true,"valuePosition":"column"},"revision":1
            }],
            "charts":[{"id":"chart-1","sheetId":"pivot","type":"bar","source":{"range":{"startRow":0,"startCol":0,"endRow":2,"endCol":1},"orientation":"columns"},"position":{"x":20,"y":20},"size":{"width":400,"height":240},"categoryField":0,"valueFields":[1],"pivotTableId":"pivot-source-id","legend":{"visible":true,"position":"bottom"},"xAxis":{"scaleType":"linear"},"yAxis":{"scaleType":"linear"}}]
        });
        let (bytes, warnings) = office_writer::export_xlsx_bytes(&model.to_string()).expect("writer exports pivot parts");
        assert!(warnings.is_empty(), "{warnings:?}");
        let envelope = parse_office_bytes(&bytes, "pivot.xlsx").expect("writer pivot XLSX reopens");
        let payload = serde_json::to_value(envelope).expect("parse payload serializes");
        let pivots = payload["payload"]["data"]["pivotTables"].as_array().expect("pivot models are restored");
        assert_eq!(pivots.len(), 1, "pivot warnings: {}", payload["diagnostics"]["warnings"]);
        let pivot = &pivots[0];
        assert_eq!(pivot["source"]["sheetName"], "Data");
        assert_eq!(pivot["source"]["range"], serde_json::json!({"startRow":0,"startCol":0,"endRow":2,"endCol":1}));
        assert_eq!(pivot["anchor"]["subUnitId"], "Pivot View");
        assert_eq!(pivot["outputRange"]["endCol"], 1);
        assert_eq!(pivot["fields"][0]["area"], "row");
        assert_eq!(pivot["fields"][1]["aggregation"], "sum");
        let chart = &payload["payload"]["data"]["sheets"][1]["charts"][0];
        assert_eq!(chart["pivotTableId"], pivot["id"], "chart payload: {chart}");
    }

    #[test]
    fn writer_scatter_chart_xlsx_reopens_with_independent_x_fields_and_log_axes() {
        let model = serde_json::json!({
            "kind": "workbook", "name": "Scatter", "sheets": [{ "id": "data", "name": "Data", "cells": {
                "A1": { "value": "Shared X" }, "B1": { "value": "First Y" }, "C1": { "value": "Other X" }, "D1": { "value": "Second Y" },
                "A2": { "value": 1 }, "B2": { "value": 10 }, "C2": { "value": 5 }, "D2": { "value": 30 },
                "A3": { "value": 2 }, "B3": { "value": 20 }, "C3": { "value": 9 }, "D3": { "value": 40 }
            }}], "charts": [{
                "id": "scatter-1", "sheetId": "data", "type": "scatter",
                "source": { "range": { "startRow": 0, "startCol": 0, "endRow": 2, "endCol": 3 }, "orientation": "columns" },
                "position": { "x": 20, "y": 20 }, "size": { "width": 400, "height": 240 },
                "categoryField": 0, "valueFields": [1, 3], "seriesToCategory": { "3": [2] },
                "xAxis": { "scaleType": "log", "logBase": 2 }, "yAxis": { "scaleType": "linear" }
            }]
        });
        let (bytes, warnings) = office_writer::export_xlsx_bytes(&model.to_string()).expect("scatter chart exports");
        assert!(warnings.is_empty(), "{warnings:?}");
        let envelope = parse_office_bytes(&bytes, "scatter.xlsx").expect("scatter chart reopens");
        let payload = serde_json::to_value(envelope).expect("parse payload serializes");
        let chart = &payload["payload"]["data"]["sheets"][0]["charts"][0];
        assert_eq!(chart["kind"], "scatter");
        assert_eq!(chart["series"][0]["categories"], "'Data'!$A$2:$A$3");
        assert_eq!(chart["series"][1]["categories"], "'Data'!$C$2:$C$3");
        assert_eq!(chart["series"][1]["values"], "'Data'!$D$2:$D$3");
        assert_eq!(chart["xAxis"]["scaleType"], "log");
        assert_eq!(chart["xAxis"]["logBase"], 2.0);
        assert_eq!(chart["yAxis"]["scaleType"], "linear");
    }

    #[test]
    fn writer_radar_chart_xlsx_reopens_with_series_layout_axes_and_legend() {
        let model = serde_json::json!({
            "kind": "workbook", "name": "Radar", "sheets": [{"id": "data", "name": "Team's Survey", "cells": {
                "A1": {"value": "Metric"}, "B1": {"value": "Team A"}, "C1": {"value": "Team B"},
                "A2": {"value": "Quality"}, "B2": {"value": 8}, "C2": {"value": 7},
                "A3": {"value": "Speed"}, "B3": {"value": 9}, "C3": {"value": 6},
                "A4": {"value": "Cost"}, "B4": {"value": 3}, "C4": {"value": 5}
            }}], "charts": [{
                "id": "radar-1", "sheetId": "data", "type": "radar",
                "source": {"range": {"startRow": 0, "startCol": 0, "endRow": 3, "endCol": 2}, "orientation": "columns"},
                "position": {"x": 24, "y": 48}, "size": {"width": 480, "height": 320},
                "title": "Team scores", "legend": {"visible": true, "position": "right"},
                "categoryField": 0, "valueFields": [1, 2],
                "xAxis": {"title": "Metrics", "scaleType": "linear"},
                "yAxis": {"title": "Score", "scaleType": "linear"}
            }]
        });
        let (bytes, warnings) =
            office_writer::export_xlsx_bytes(&model.to_string()).expect("radar chart exports");
        assert!(warnings.is_empty(), "{warnings:?}");
        let envelope = parse_office_bytes(&bytes, "radar.xlsx").expect("radar chart reopens");
        let payload = serde_json::to_value(envelope).expect("parse payload serializes");
        let chart = &payload["payload"]["data"]["sheets"][0]["charts"][0];
        assert_eq!(chart["kind"], "radar");
        assert_eq!(chart["title"], "Team scores");
        assert_eq!(chart["series"].as_array().unwrap().len(), 2);
        assert_eq!(chart["series"][0]["categories"], "'Team''s Survey'!$A$2:$A$4");
        assert_eq!(chart["series"][0]["values"], "'Team''s Survey'!$B$2:$B$4");
        assert_eq!(chart["series"][1]["categories"], "'Team''s Survey'!$A$2:$A$4");
        assert_eq!(chart["series"][1]["values"], "'Team''s Survey'!$C$2:$C$4");
        assert_eq!(chart["drawing"]["behavior"], "absolute");
        assert_eq!(chart["drawing"]["fromOffsetEmu"], json!([228_600, 457_200]));
        assert_eq!(
            chart["drawing"]["absoluteSizeEmu"],
            json!([4_572_000, 3_048_000])
        );
        assert_eq!(chart["xAxisTitle"], "Metrics");
        assert_eq!(chart["yAxisTitle"], "Score");
        assert_eq!(chart["xAxis"]["scaleType"], "linear");
        assert_eq!(chart["yAxis"]["scaleType"], "linear");
        assert_eq!(chart["legend"], true);
        assert_eq!(chart["legendPosition"], "right");
    }

    fn add_xlsx_validation_attributes(bytes: &[u8]) -> Vec<u8> {
        let mut archive = ZipArchive::new(std::io::Cursor::new(bytes)).expect("valid XLSX ZIP");
        let mut output = ZipWriter::new(std::io::Cursor::new(Vec::new()));
        for index in 0..archive.len() {
            let mut entry = archive.by_index(index).expect("XLSX package entry");
            let name = entry.name().unwrap().into_owned();
            let mut contents = Vec::new();
            entry
                .read_to_end(&mut contents)
                .expect("XLSX package entry contents");
            if name == "xl/worksheets/sheet1.xml" {
                let xml = String::from_utf8(contents).expect("worksheet XML is UTF-8");
                let start = xml
                    .find("<dataValidation ")
                    .expect("worksheet contains a validation element");
                let end = start + xml[start..].find('>').expect("validation element tag") + 1;
                let mut tag = xml[start..end].to_owned();
                tag = tag.replace("B2:B4", "B2:B4 D2:D4");
                tag.insert_str(
                    tag.len() - 1,
                    " showDropDown=\"1\" errorStyle=\"information\"",
                );
                let mut modified = xml;
                modified.replace_range(start..end, &tag);
                contents = modified.into_bytes();
            }
            output
                .start_file(name, SimpleFileOptions::default())
                .expect("write XLSX package entry");
            output
                .write_all(&contents)
                .expect("write XLSX package contents");
        }
        output.finish().expect("finish XLSX ZIP").into_inner()
    }

    fn add_xlsx_sparkline_extension(bytes: &[u8]) -> Vec<u8> {
        let mut archive = ZipArchive::new(std::io::Cursor::new(bytes)).expect("valid XLSX ZIP");
        let mut output = ZipWriter::new(std::io::Cursor::new(Vec::new()));
        for index in 0..archive.len() {
            let mut entry = archive.by_index(index).expect("XLSX package entry");
            let name = entry.name().unwrap().into_owned();
            let mut contents = Vec::new();
            entry.read_to_end(&mut contents).expect("XLSX package entry contents");
            if name == "xl/worksheets/sheet1.xml" {
                let mut xml = String::from_utf8(contents).expect("worksheet XML is UTF-8");
                let ext = "<extLst><ext uri=\"{05C60535-1F16-4fd2-B633-F4F36F0B64E0}\"><x14:sparklineGroups xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\"><x14:sparklineGroup type=\"column\" displayEmptyCellsAs=\"zero\" minAxisType=\"custom\" maxAxisType=\"custom\" manualMin=\"-5\" manualMax=\"20\" negative=\"1\" high=\"1\" low=\"1\"><x14:colorSeries rgb=\"FF167A55\"/><x14:colorNegative rgb=\"FFC43B3B\"/><x14:colorHigh rgb=\"FFF0A000\"/><x14:colorLow rgb=\"FF2B579A\"/><x14:sparklines><x14:sparkline><xm:f>'Data'!B1:E1</xm:f><xm:sqref>F1</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext></extLst>";
                xml.insert_str(xml.rfind("</worksheet>").expect("worksheet close"), ext);
                contents = xml.into_bytes();
            }
            output.start_file(name, SimpleFileOptions::default()).expect("write XLSX package entry");
            output.write_all(&contents).expect("write XLSX package contents");
        }
        output.finish().expect("finish XLSX ZIP").into_inner()
    }

    #[test]
    fn chart_payload_retains_unsupported_constructs_as_structured_diagnostics() {
        let mut metadata = rxls::DrawingMetadata::default();
        metadata.kind = rxls::DrawingObjectKind::Chart;
        metadata.chart_unsupported_reasons = vec![rxls::ChartUnsupportedReason::Combo];
        metadata.from_cell = Some((1, 2));
        metadata.absolute_size_emu = Some((914_400, 457_200));
        let chart = Chart::new(ChartKind::Bar, (1, 2), (8, 8));
        let payload = super::chart_payload(0, &chart, &[metadata]);
        let json: Value = serde_json::to_value(payload).expect("chart payload must serialize");

        assert_eq!(json["drawing"]["chartUnsupportedReasons"][0], "combo");
        assert_eq!(
            json["drawing"]["fromCell"],
            serde_json::json!({ "row": 1, "col": 2 })
        );
        assert_eq!(
            json["drawing"]["absoluteSizeEmu"],
            serde_json::json!([914_400, 457_200])
        );
    }
}
