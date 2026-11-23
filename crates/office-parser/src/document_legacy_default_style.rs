use std::io::Cursor;

use duke_pages::{
    DocModel, OoxmlPackage, ParagraphStyle,
    doc::{cfb::CompoundFile, fib::{Fib, fcidx}, fonts::FontTable, stsh::StyleSheet},
};
use zip::ZipArchive;
use office_io_common::document_legacy_source::BINARY_BYTE_LIMIT as DOC_BINARY_LIMIT;

use crate::model::{
    DocumentDefaultStylePayload, DocumentOfficeSourceOriginFormat,
    DocumentOfficeSourceOriginPayload, DocumentOfficeSourcePayload,
    DocumentOfficeSourceResourcePayload, DocumentOfficeXmlPartPayload,
};
use super::{ParagraphWordWrap, XML_LIMIT, parse_xml, read_document_default_style, read_word_wrap_xml, read_xml};

const ODF_OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";

pub(crate) struct LegacyDocumentDefaults {
    pub source: DocumentOfficeSourcePayload,
    pub default_style: Option<DocumentDefaultStylePayload>,
    pub paragraph_wrap: ParagraphWordWrap,
    pub warnings: Vec<String>,
}

/// Reuse the same synthesized package that built the DOC model. The original
/// binary style/font records are separate resources, never labelled source XML.
pub(crate) fn read_doc_default_style(bytes: &[u8], package: &OoxmlPackage) -> LegacyDocumentDefaults {
    let mut resources = Vec::new();
    let mut errors = Vec::new();
    let mut warnings = Vec::new();
    if let Err(error) = capture_doc_resources(bytes, &mut resources, &mut errors, &mut warnings) {
        errors.push(format!("DOC original style resources could not be retained: {error}"));
    }
    let styles = match package.parts.get("word/styles.xml") {
        Some(part) if part.content.len() as u64 <= XML_LIMIT => Some(DocumentOfficeXmlPartPayload {
            part_name: "word/styles.xml".into(), xml: part.content.clone(),
        }),
        Some(_) => {
            errors.push("DOC synthesized styles exceed the 16 MiB limit.".into());
            None
        },
        None => {
            errors.push("DOC synthesized word/styles.xml is missing.".into());
            None
        },
    };
    let (numbering, numbering_errors) = super::numbering_source::synthesized(package);
    errors.extend(numbering_errors.into_iter().map(|error| format!("DOC synthesized numbering resources could not be retained: {error}")));
    let mut source = DocumentOfficeSourcePayload {
        version: 1, styles, theme: None, numbering, resource_errors: errors,
        origin: Some(DocumentOfficeSourceOriginPayload {
            format: DocumentOfficeSourceOriginFormat::Doc, synthesized: true, resources,
            resource_errors: Vec::new(),
        }),
        legacy_origins: Vec::new(),
    };
    let paragraph_wrap = match source.styles.as_ref().map(|part| read_word_wrap_xml(&part.xml)).transpose() {
        Ok(wrap) => wrap.unwrap_or_default(),
        Err(error) => {
            source.resource_errors.push(format!("DOC synthesized paragraph defaults could not be read: {error}"));
            ParagraphWordWrap::default()
        },
    };
    let default_style = if source.resource_errors.is_empty() {
        match read_document_default_style(Some(&source), &paragraph_wrap) {
            Ok(style) => Some(style),
            Err(error) => {
                source.resource_errors.push(format!("DOC document defaults could not be resolved: {error}"));
                None
            },
        }
    } else { None };
    if let Some(origin) = &mut source.origin { origin.resource_errors = source.resource_errors.clone(); }
    warnings.extend(source.resource_errors.iter().cloned());
    LegacyDocumentDefaults { source, default_style, paragraph_wrap, warnings }
}

fn capture_doc_resources(
    bytes: &[u8],
    resources: &mut Vec<DocumentOfficeSourceResourcePayload>,
    errors: &mut Vec<String>,
    warnings: &mut Vec<String>,
) -> Result<(), String> {
    // Duke exposes only the complete CFB reader publicly. This extra resource
    // read copies streams; its time/peak-memory/hex serialization cost is unmeasured.
    let compound = CompoundFile::parse(bytes)?;
    let word = compound.try_stream("WordDocument")?.ok_or("DOC WordDocument stream is missing.")?;
    let fib = Fib::parse(word)?;
    let preferred = fib.table_stream_name();
    let (stream_name, table) = match compound.try_stream(preferred)? {
        Some(table) => (preferred, table),
        None => {
            let fallback = if preferred == "1Table" { "0Table" } else { "1Table" };
            let table = compound.try_stream(fallback)?.ok_or("DOC Table stream is missing.")?;
            warnings.push(format!("DOC original styles use {fallback}; the FIB-selected {preferred} stream is absent."));
            (fallback, table)
        },
    };
    let mut retained = 0usize;
    let mut style_table = None;
    let mut font_table = None;
    for (index, label) in [(fcidx::STSHF, "STSHF"), (fcidx::STTBF_FFN, "STTBF_FFN")] {
        let result = (|| {
            let (offset, length) = fib.fc_lcb(index).ok_or_else(|| format!("FIB has no {label} pointer pair."))?;
            let data = checked_doc_table_range(table, offset, length)?;
            let total = retained.checked_add(data.len()).ok_or("DOC style resource size overflows.")?;
            if total > DOC_BINARY_LIMIT { return Err("DOC original style resources exceed the 16 MiB limit.".into()); }
            retained = total;
            // Even a structurally invalid record remains available for a later
            // lossless adapter. A truncated FIB range is never stored as complete.
            resources.push(DocumentOfficeSourceResourcePayload::Binary {
                stream_name: stream_name.into(), offset, data_hex: encode_hex(data),
            });
            if index == fcidx::STSHF {
                validate_doc_styles(data)?;
                style_table = Some(data);
            } else {
                validate_doc_fonts(data)?;
                font_table = Some(data);
            }
            Ok::<(), String>(())
        })();
        if let Err(error) = result { errors.push(format!("DOC {label} resource could not be resolved: {error}")); }
    }
    if let (Some(styles), Some(fonts)) = (style_table, font_table) {
        let styles = StyleSheet::parse(styles);
        let fonts = FontTable::parse(fonts);
        if styles.ftc_standard.iter().flatten().any(|index| fonts.name(*index).is_none()) {
            errors.push("DOC STSHI standard font slot references an absent or empty FFN font.".into());
        }
    }
    Ok(())
}

fn checked_doc_table_range(table: &[u8], offset: u32, length: u32) -> Result<&[u8], String> {
    if length == 0 { return Err("FIB style resource is absent or empty.".into()); }
    let end = offset.checked_add(length).ok_or("FIB style resource offset/length overflows u32.")?;
    let start = usize::try_from(offset).map_err(|_| "FIB style resource offset is not addressable.")?;
    let end = usize::try_from(end).map_err(|_| "FIB style resource end is not addressable.")?;
    table.get(start..end).ok_or_else(|| "FIB style resource extends beyond the Table stream.".into())
}

fn encode_hex(data: &[u8]) -> String {
    const HEX: &[u8; 16] = b"0123456789abcdef";
    let mut result = String::with_capacity(data.len() * 2);
    for byte in data {
        result.push(char::from(HEX[(byte >> 4) as usize]));
        result.push(char::from(HEX[(byte & 15) as usize]));
    }
    result
}

fn u16_at(data: &[u8], offset: usize) -> Result<u16, String> {
    let end = offset.checked_add(2).ok_or("DOC style record offset overflows.")?;
    let bytes = data.get(offset..end).ok_or("DOC style record is truncated.")?;
    Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
}

fn validate_doc_styles(data: &[u8]) -> Result<(), String> {
    let header_length = usize::from(u16_at(data, 0)?);
    let header_end = 2 + header_length;
    let header = data.get(2..header_end).ok_or("DOC STSHI header is truncated.")?;
    if header.len() < 18 { return Err("DOC STSHI does not contain its standard font slots.".into()); }
    let count = usize::from(u16_at(header, 0)?);
    let base_length = usize::from(u16_at(header, 2)?);
    if count > 4095 || base_length < 10 { return Err("DOC STSHI has invalid style count or STD base length.".into()); }
    let mut position = header_end;
    let mut nonempty = Vec::new();
    for index in 0..count {
        let length = usize::from(u16_at(data, position)?);
        position += 2;
        let end = position.checked_add(length).ok_or("DOC STD length overflows.")?;
        let record = data.get(position..end).ok_or("DOC STD record is truncated.")?;
        if length != 0 {
            if record.len() < base_length { return Err("DOC STD is smaller than its declared base.".into()); }
            nonempty.push(index);
        }
        position = end + length % 2;
        if position > data.len() { return Err("DOC STD alignment padding is missing.".into()); }
    }
    let parsed = StyleSheet::parse(data);
    if parsed.styles.len() != count || nonempty.iter().any(|index| parsed.styles[*index].is_none()) {
        return Err("DOC STD name or UPX properties could not be read completely.".into());
    }
    Ok(())
}

fn validate_doc_fonts(data: &[u8]) -> Result<(), String> {
    let first = u16_at(data, 0)?;
    let (count, extra, mut position) = if first == 0xffff {
        (usize::from(u16_at(data, 2)?), u16_at(data, 4)?, 6usize)
    } else { (usize::from(first), u16_at(data, 2)?, 4usize) };
    if extra != 0 { return Err("DOC STTBF_FFN extra-record data is not yet resolved.".into()); }
    for _ in 0..count {
        let length = usize::from(*data.get(position).ok_or("DOC FFN record is missing.")?);
        let end = position.checked_add(1 + length).ok_or("DOC FFN length overflows.")?;
        let record = data.get(position + 1..end).ok_or("DOC FFN record is truncated.")?;
        let name = record.get(39..).ok_or("DOC FFN fixed header is truncated.")?;
        if !name.chunks_exact(2).any(|unit| unit == [0, 0]) { return Err("DOC FFN font name has no UTF-16 terminator.".into()); }
        position = end;
    }
    Ok(())
}

/// ODF has no Word style part. Keep its package XML only under origin and read
/// the adapter's explicit global definition; never sample a body paragraph.
pub(crate) fn read_odt_default_style(
    model: &DocModel,
    archive: &mut ZipArchive<Cursor<&[u8]>>,
) -> LegacyDocumentDefaults {
    let mut resources = Vec::new();
    let mut errors = Vec::new();
    for (part, root_name, required) in [
        ("styles.xml", "document-styles", false),
        ("content.xml", "document-content", true),
    ] {
        match read_xml(archive, part) {
            Ok(Some(xml)) => {
                if let Err(error) = validate_odt_resource(&xml, root_name) {
                    errors.push(format!("ODT {part} style resource is invalid: {error}"));
                }
                resources.push(DocumentOfficeSourceResourcePayload::Xml { part_name: part.into(), xml });
            },
            Ok(None) if required => errors.push(format!("ODT required {part} is missing.")),
            Ok(None) => {},
            Err(error) => errors.push(format!("ODT {part} style resource could not be retained: {error}")),
        }
    }
    let mut source = DocumentOfficeSourcePayload {
        version: 1, styles: None, theme: None, numbering: None, resource_errors: errors,
        origin: Some(DocumentOfficeSourceOriginPayload {
            format: DocumentOfficeSourceOriginFormat::Odt, synthesized: false, resources,
            resource_errors: Vec::new(),
        }),
        legacy_origins: Vec::new(),
    };
    let default_style = if source.resource_errors.is_empty() {
        match odt_defaults_from_model(model) {
            Ok(style) => Some(style),
            Err(error) => {
                source.resource_errors.push(format!("ODT global defaults could not be resolved: {error}"));
                None
            },
        }
    } else { None };
    if let Some(origin) = &mut source.origin { origin.resource_errors = source.resource_errors.clone(); }
    let warnings = source.resource_errors.clone();
    LegacyDocumentDefaults { source, default_style, paragraph_wrap: ParagraphWordWrap::default(), warnings }
}

fn validate_odt_resource(xml: &str, expected: &str) -> Result<(), String> {
    let root = parse_xml(xml)?;
    let prefix = root.qualified_name.split_once(':').map_or("", |(prefix, _)| prefix);
    if root.name != expected || root.namespaces.get(prefix).map(String::as_str) != Some(ODF_OFFICE) {
        return Err(format!("Expected the ODF office:{expected} root."));
    }
    Ok(())
}

fn odt_defaults_from_model(model: &DocModel) -> Result<DocumentDefaultStylePayload, String> {
    let id = model.metadata.default_paragraph_style_id.as_deref().filter(|id| !id.is_empty())
        .ok_or("ODT metadata has no default paragraph style ID.")?;
    let mut matching = model.metadata.paragraph_styles.iter().filter(|style| style.id == id);
    let definition = matching.next().ok_or("ODT default paragraph style definition is missing.")?;
    if matching.next().is_some() || definition.is_default != Some(true)
        || model.metadata.paragraph_styles.iter().filter(|style| style.is_default == Some(true)).count() != 1 {
        return Err("ODT default paragraph style ID and unique default marker disagree.".into());
    }
    let paragraph = ParagraphStyle {
        align: definition.align, snap_to_grid: definition.snap_to_grid,
        text_alignment: definition.text_alignment,
        spacing: definition.spacing.clone(), indent: definition.indent.clone(),
        background_color: definition.background_color.clone(), borders: definition.borders.clone(),
        tab_stops: definition.tab_stops.clone(), contextual_spacing: definition.contextual_spacing,
        keep_next: definition.keep_next, keep_lines: definition.keep_lines,
        widow_control: definition.widow_control, page_break_before: definition.page_break_before,
        ..Default::default()
    };
    Ok(DocumentDefaultStylePayload {
        version: 1, text: definition.run_style.clone(), paragraph: Some(paragraph), word_wrap: None,
    })
}

#[cfg(test)]
#[path = "document_legacy_default_style_tests.rs"]
mod tests;
