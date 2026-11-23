// Drafts only: not compiled or executed in the 2026-11-22 continuation.
use super::*;
use std::io::Write;

use crate::model::{DocumentModel, DocumentPayload};
use duke_pages::ParagraphAlignment;
use zip::{ZipWriter, write::SimpleFileOptions};

#[path = "document_legacy_source_roundtrip_tests.rs"]
mod source_roundtrip;

const STYLE_OFFSET: usize = 32;
const FONT_OFFSET: usize = 512;

fn put16(data: &mut [u8], offset: usize, value: u16) {
    data[offset..offset + 2].copy_from_slice(&value.to_le_bytes());
}

fn put32(data: &mut [u8], offset: usize, value: u32) {
    data[offset..offset + 4].copy_from_slice(&value.to_le_bytes());
}

fn fib_pair(word: &mut [u8], index: usize, offset: u32, length: u32) {
    put32(word, 154 + index * 8, offset);
    put32(word, 158 + index * 8, length);
}

fn std_record(sti: u16, parent: u16, next: u16, name: &str, papx: &[u8], chpx: &[u8]) -> Vec<u8> {
    let mut record = vec![0u8; 10];
    put16(&mut record, 0, sti);
    put16(&mut record, 2, (parent << 4) | 1);
    put16(&mut record, 4, (next << 4) | 2);
    let units: Vec<u16> = name.encode_utf16().collect();
    record.extend_from_slice(&(units.len() as u16).to_le_bytes());
    for unit in units { record.extend_from_slice(&unit.to_le_bytes()); }
    record.extend_from_slice(&0u16.to_le_bytes());
    record.extend_from_slice(&((papx.len() + 2) as u16).to_le_bytes());
    record.extend_from_slice(&0u16.to_le_bytes());
    record.extend_from_slice(papx);
    if record.len() % 2 != 0 { record.push(0); }
    record.extend_from_slice(&(chpx.len() as u16).to_le_bytes());
    record.extend_from_slice(chpx);
    if record.len() % 2 != 0 { record.push(0); }
    record
}

fn stsh() -> Vec<u8> {
    // Normal inherits Base's italic/right/after/line values, but overrides size
    // and indent. The first body paragraph uses a separate 42pt Leading style.
    let records = [
        std_record(0, 1, 0, "Normal", &[0x0f, 0x84, 0xd0, 0x02], &[0x43, 0x4a, 32, 0]),
        std_record(20, 0x0fff, 1, "Base", &[0x03, 0x24, 2, 0x14, 0xa4, 180, 0, 0x12, 0x64, 0x20, 0x01, 1, 0], &[0x36, 0x08, 1]),
        std_record(21, 0x0fff, 2, "Leading", &[0x03, 0x24, 1], &[0x43, 0x4a, 84, 0, 0x4f, 0x4a, 3, 0, 0x35, 0x08, 1]),
    ];
    let mut header = vec![0u8; 18];
    put16(&mut header, 0, records.len() as u16);
    put16(&mut header, 2, 10);
    put16(&mut header, 12, 0);
    put16(&mut header, 14, 1);
    put16(&mut header, 16, 2);
    let mut result = Vec::new();
    result.extend_from_slice(&(header.len() as u16).to_le_bytes());
    result.extend_from_slice(&header);
    for record in records {
        result.extend_from_slice(&(record.len() as u16).to_le_bytes());
        result.extend_from_slice(&record);
    }
    result
}

fn fonts() -> Vec<u8> {
    let mut table = Vec::new();
    table.extend_from_slice(&0xffffu16.to_le_bytes());
    table.extend_from_slice(&4u16.to_le_bytes());
    table.extend_from_slice(&0u16.to_le_bytes());
    for name in ["Georgia", "宋体", "Amiri", "Courier New"] {
        let mut body = vec![0u8; 39];
        for unit in name.encode_utf16() { body.extend_from_slice(&unit.to_le_bytes()); }
        body.extend_from_slice(&0u16.to_le_bytes());
        table.push(u8::try_from(body.len()).unwrap());
        table.extend_from_slice(&body);
    }
    table
}

fn doc_streams() -> (Vec<u8>, Vec<u8>) {
    let text = "First styled paragraph\rSecond paragraph\r";
    let units: Vec<u16> = text.encode_utf16().collect();
    let mut word = vec![0u8; 4096];
    put16(&mut word, 0, 0xa5ec);
    put16(&mut word, 2, 0x00c1);
    put16(&mut word, 10, 0x0204);
    put16(&mut word, 32, 14);
    put16(&mut word, 62, 22);
    put32(&mut word, 76, units.len() as u32);
    put16(&mut word, 152, 34);
    for (index, unit) in units.iter().enumerate() { put16(&mut word, 512 + index * 2, *unit); }
    let first_end = 512 + "First styled paragraph\r".encode_utf16().count() * 2;
    let all_end = 512 + units.len() * 2;
    let page = &mut word[1024..1536];
    page[511] = 2;
    put32(page, 0, 512);
    put32(page, 4, first_end as u32);
    put32(page, 8, all_end as u32);
    page[12] = 60;
    // cb=0 escape: two-byte payload containing the Leading style index.
    page[120] = 0; page[121] = 1; put16(page, 122, 2);
    // The second paragraph inherits Normal through an empty BX record.
    let styles = stsh();
    let fonts = fonts();
    let mut table = vec![0u8; 4096];
    table[0] = 2;
    put32(&mut table, 1, 16);
    put32(&mut table, 9, units.len() as u32);
    put32(&mut table, 15, 512);
    table[STYLE_OFFSET..STYLE_OFFSET + styles.len()].copy_from_slice(&styles);
    table[FONT_OFFSET..FONT_OFFSET + fonts.len()].copy_from_slice(&fonts);
    put32(&mut table, 1024, 512);
    put32(&mut table, 1028, all_end as u32);
    put32(&mut table, 1032, 2);
    fib_pair(&mut word, fcidx::CLX, 0, 21);
    fib_pair(&mut word, fcidx::STSHF, STYLE_OFFSET as u32, styles.len() as u32);
    fib_pair(&mut word, fcidx::STTBF_FFN, FONT_OFFSET as u32, fonts.len() as u32);
    fib_pair(&mut word, fcidx::PLCF_BTE_PAPX, 1024, 12);
    (word, table)
}

/// Small project-owned CFB fixture with two regular 4096-byte streams.
/// Uses the published CFB/FIB/STD layout, not Duke's private test builder.
fn compound(word: &[u8], table: &[u8], table_name: &str) -> Vec<u8> {
    const FREE: u32 = 0xffffffff;
    const END: u32 = 0xfffffffe;
    assert_eq!(word.len(), 4096);
    assert_eq!(table.len(), 4096);
    let mut header = vec![0u8; 512];
    header[..8].copy_from_slice(&[0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
    put16(&mut header, 24, 0x003e);
    put16(&mut header, 26, 3);
    put16(&mut header, 28, 0xfffe);
    put16(&mut header, 30, 9);
    put16(&mut header, 32, 6);
    put32(&mut header, 44, 1);
    put32(&mut header, 48, 16);
    put32(&mut header, 56, 4096);
    put32(&mut header, 60, END);
    put32(&mut header, 68, END);
    for index in 0..109 { put32(&mut header, 76 + index * 4, FREE); }
    put32(&mut header, 76, 17);
    let mut directory = vec![0u8; 512];
    for (index, name, kind, right, child, start, size) in [
        (0usize, "Root Entry", 5u8, FREE, 1u32, END, 0u64),
        (1, "WordDocument", 2, 2, FREE, 0, 4096),
        (2, table_name, 2, FREE, FREE, 8, 4096),
    ] {
        let entry = &mut directory[index * 128..(index + 1) * 128];
        let units: Vec<u16> = name.encode_utf16().collect();
        for (offset, unit) in units.iter().enumerate() { put16(entry, offset * 2, *unit); }
        put16(entry, 64, ((units.len() + 1) * 2) as u16);
        entry[66] = kind; entry[67] = 1;
        put32(entry, 68, FREE); put32(entry, 72, right); put32(entry, 76, child);
        put32(entry, 116, start);
        entry[120..128].copy_from_slice(&size.to_le_bytes());
    }
    let mut fat = vec![0xff; 512];
    for start in [0usize, 8] {
        for index in start..start + 8 { put32(&mut fat, index * 4, if index == start + 7 { END } else { index as u32 + 1 }); }
    }
    put32(&mut fat, 16 * 4, END);
    put32(&mut fat, 17 * 4, 0xfffffffd);
    header.extend_from_slice(word);
    header.extend_from_slice(table);
    header.extend_from_slice(&directory);
    header.extend_from_slice(&fat);
    header
}

fn document_payload(bytes: &[u8], name: &str) -> (DocumentPayload, Vec<String>) {
    let envelope = crate::parser::parse_document_bytes(bytes, name).unwrap();
    match envelope.payload {
        DocumentModel::Document(payload) => (payload, envelope.diagnostics.capabilities),
        _ => panic!("expected document payload"),
    }
}

#[test]
fn doc_binary_globals_and_normal_inheritance_are_independent_of_the_first_paragraph() {
    let (word, table) = doc_streams();
    let bytes = compound(&word, &table, "1Table");
    let (payload, capabilities) = document_payload(&bytes, "globals.doc");
    assert!(capabilities.iter().any(|name| name == "documentDefaultStyle"));
    let defaults = payload.default_style.unwrap();
    let text = defaults.text.unwrap();
    assert_eq!(text.font_size_pt, Some(16.0));
    assert_eq!(text.italic, Some(true));
    assert_eq!(text.font_family_ascii.as_deref(), Some("Georgia"));
    assert_eq!(text.font_family_east_asia.as_deref(), Some("宋体"));
    assert_eq!(text.font_family_cs.as_deref(), Some("Amiri"));
    let paragraph = defaults.paragraph.unwrap();
    assert_eq!(paragraph.align, Some(ParagraphAlignment::Right));
    assert_eq!(paragraph.indent.unwrap().left_twips, Some(720));
    assert_eq!(paragraph.spacing.unwrap().after_twips, Some(180));
    let first = payload.document.blocks.iter().find(|block| block.content == "First styled paragraph").unwrap();
    let first = first.details.to_json_value();
    assert_eq!(first["inlines"][0]["style"]["fontSizePt"], 42.0);
    assert_eq!(first["style"]["align"], "center");
    let source = serde_json::to_value(payload.office_source.unwrap()).unwrap();
    assert_eq!(source["origin"]["format"], "doc");
    assert_eq!(source["origin"]["synthesized"], true);
    assert_eq!(source["origin"]["resources"][0]["streamName"], "1Table");
    assert_eq!(source["origin"]["resources"][0]["offset"], STYLE_OFFSET as u32);
    assert_eq!(source["origin"]["resources"][0]["dataHex"], encode_hex(&stsh()));
    assert_eq!(source["origin"]["resources"][1]["dataHex"], encode_hex(&fonts()));
    assert!(source["styles"]["xml"].as_str().unwrap().contains("w:basedOn"));
}

#[test]
fn doc_fib_outside_or_overflowing_ranges_never_become_truncated_successes() {
    for (offset, length) in [(4090, 20), (0xfffffff0, 32), (32, 0)] {
        let (mut word, table) = doc_streams();
        fib_pair(&mut word, fcidx::STSHF, offset, length);
        let bytes = compound(&word, &table, "1Table");
        let (payload, capabilities) = document_payload(&bytes, "bad-range.doc");
        assert!(payload.default_style.is_none());
        assert!(!capabilities.iter().any(|name| name == "documentDefaultStyle"));
        let source = payload.office_source.unwrap();
        assert!(source.resource_errors.iter().any(|error| error.contains("STSHF")));
        let origin = source.origin.unwrap();
        assert_eq!(origin.resources.len(), 1, "only the complete font record is retained");
        let serialized = serde_json::to_value(&origin.resources[0]).unwrap();
        assert_eq!(serialized["offset"], FONT_OFFSET as u32);
        assert_eq!(serialized["dataHex"], encode_hex(&fonts()));
    }
}

#[test]
fn doc_declared_internal_record_truncation_and_bad_font_indices_keep_original_bytes_and_errors() {
    let (word, mut table) = doc_streams();
    put16(&mut table, STYLE_OFFSET, 500);
    let bytes = compound(&word, &table, "1Table");
    let package = duke_pages::parse_doc(&bytes).unwrap();
    let defaults = read_doc_default_style(&bytes, &package);
    assert!(defaults.default_style.is_none());
    assert!(defaults.source.resource_errors.iter().any(|error| error.contains("STSHI header")));
    assert_eq!(defaults.source.origin.unwrap().resources.len(), 2);
    let (word, mut table) = doc_streams();
    put16(&mut table, STYLE_OFFSET + 2 + 14, 99);
    let bytes = compound(&word, &table, "1Table");
    let package = duke_pages::parse_doc(&bytes).unwrap();
    let defaults = read_doc_default_style(&bytes, &package);
    assert!(defaults.default_style.is_none());
    assert!(defaults.source.resource_errors.iter().any(|error| error.contains("standard font slot")));
    assert_eq!(defaults.source.origin.unwrap().resources.len(), 2);
    let mut font_record = fonts();
    font_record.pop();
    assert!(validate_doc_fonts(&font_record).is_err());
    let mut style_record = stsh();
    style_record.pop();
    assert!(validate_doc_styles(&style_record).is_err());
}

#[test]
fn doc_table_fallback_records_the_actual_stream_without_hiding_it() {
    let (word, table) = doc_streams();
    let bytes = compound(&word, &table, "0Table");
    let package = duke_pages::parse_doc(&bytes).unwrap();
    let defaults = read_doc_default_style(&bytes, &package);
    assert!(defaults.default_style.is_some());
    assert!(defaults.source.resource_errors.is_empty());
    assert!(defaults.warnings.iter().any(|error| error.contains("FIB-selected 1Table")));
    for resource in defaults.source.origin.unwrap().resources {
        let resource = serde_json::to_value(resource).unwrap();
        assert_eq!(resource["streamName"], "0Table");
    }
}

fn odt_styles(standard: bool) -> String {
    let standard = if standard {
        r#"<style:style style:name="Standard" style:family="paragraph" style:parent-style-name="Base"><style:text-properties fo:font-size="18pt"/></style:style>"#
    } else { "" };
    format!(r#"<office:document-styles xmlns:office="{ODF_OFFICE}" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:extra="urn:office-viewer:test" office:version="1.2">
<office:font-face-decls><style:font-face style:name="LatinDefault" svg:font-family="Georgia"/><style:font-face style:name="EastDefault" svg:font-family="宋体"/><style:font-face style:name="ComplexDefault" svg:font-family="Amiri"/></office:font-face-decls>
<office:styles>
<style:default-style style:family="paragraph"><style:text-properties style:font-name="LatinDefault" style:font-name-asian="EastDefault" style:font-name-complex="ComplexDefault" fo:font-size="15pt" fo:font-weight="normal" fo:font-style="normal"/><style:paragraph-properties fo:text-align="left" fo:margin-left="0.5in" fo:margin-right="0.25in" fo:text-indent="-0.1in" fo:margin-top="4pt" fo:margin-bottom="6pt" fo:line-height="125%"/><extra:unmodelled extra:value="retain"/></style:default-style>
<style:style style:name="Base" style:family="paragraph"><style:text-properties fo:font-style="italic"/><style:paragraph-properties fo:text-align="right"/></style:style>
{standard}
<style:style style:name="Leading" style:family="paragraph"><style:text-properties fo:font-family="Courier New" fo:font-size="42pt" fo:font-weight="bold"/><style:paragraph-properties fo:text-align="center"/></style:style>
</office:styles><office:automatic-styles/><office:master-styles/></office:document-styles>"#)
}

fn odt_content() -> String {
    format!(r#"<office:document-content xmlns:office="{ODF_OFFICE}" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.2"><office:automatic-styles/><office:body><office:text><text:p text:style-name="Leading">First styled paragraph</text:p><text:p>Second paragraph</text:p></office:text></office:body></office:document-content>"#)
}

fn odt_bytes(styles: Option<&str>, content: Option<&str>) -> Vec<u8> {
    let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
    let options = SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    writer.start_file("mimetype", options).unwrap();
    writer.write_all(b"application/vnd.oasis.opendocument.text").unwrap();
    for (name, xml) in [("styles.xml", styles), ("content.xml", content)] {
        if let Some(xml) = xml {
            writer.start_file(name, options).unwrap();
            writer.write_all(xml.as_bytes()).unwrap();
        }
    }
    writer.finish().unwrap().into_inner()
}

#[test]
fn odt_standard_inheritance_and_global_fallback_never_use_the_first_paragraph() {
    for standard in [true, false] {
        let styles = odt_styles(standard);
        let content = odt_content();
        let bytes = odt_bytes(Some(&styles), Some(&content));
        let (payload, capabilities) = document_payload(&bytes, "global.odt");
        assert!(capabilities.iter().any(|name| name == "documentDefaultStyle"));
        let defaults = payload.default_style.unwrap();
        let text = defaults.text.unwrap();
        assert_eq!(text.font_size_pt, Some(if standard { 18.0 } else { 15.0 }));
        assert_eq!(text.italic, Some(standard));
        assert_eq!(text.bold, Some(false));
        assert_eq!(text.font_family_ascii.as_deref(), Some("Georgia"));
        assert_eq!(text.font_family_east_asia.as_deref(), Some("宋体"));
        assert_eq!(text.font_family_cs.as_deref(), Some("Amiri"));
        let paragraph = defaults.paragraph.unwrap();
        assert_eq!(paragraph.align, Some(if standard { ParagraphAlignment::Right } else { ParagraphAlignment::Left }));
        let indent = paragraph.indent.unwrap();
        assert_eq!(indent.left_twips, Some(720));
        assert_eq!(indent.right_twips, Some(360));
        assert_eq!(indent.hanging_twips, Some(144));
        let spacing = paragraph.spacing.unwrap();
        assert_eq!(spacing.before_twips, Some(80));
        assert_eq!(spacing.after_twips, Some(120));
        assert_eq!(spacing.line_twips, Some(300));
        let first = payload.document.blocks.iter().find(|block| block.content == "First styled paragraph").unwrap();
        assert_eq!(first.details.to_json_value()["inlines"][0]["style"]["fontSizePt"], 42.0);
        let source = payload.office_source.unwrap();
        assert!(source.styles.is_none() && source.theme.is_none());
        let origin = serde_json::to_value(source.origin.unwrap()).unwrap();
        assert_eq!(origin["format"], "odt");
        assert_eq!(origin["synthesized"], false);
        assert_eq!(origin["resources"][0]["xml"], styles);
        assert_eq!(origin["resources"][1]["xml"], content);
    }
}

#[test]
fn odt_empty_global_definition_is_explicit_and_does_not_copy_body_formatting() {
    let content = odt_content();
    let bytes = odt_bytes(None, Some(&content));
    let (payload, _) = document_payload(&bytes, "empty-default.odt");
    let defaults = payload.default_style.unwrap();
    assert!(defaults.text.is_none());
    let paragraph = defaults.paragraph.unwrap();
    assert!(paragraph.align.is_none() && paragraph.spacing.is_none() && paragraph.indent.is_none());
    let source = payload.office_source.unwrap();
    assert!(source.resource_errors.is_empty());
    assert_eq!(source.origin.unwrap().resources.len(), 1);
}

#[test]
fn odt_missing_duplicate_and_inconsistent_global_markers_are_rejected_without_mutation() {
    let bytes = odt_bytes(Some(&odt_styles(true)), Some(&odt_content()));
    let source = duke_pages::parse_odt(&bytes).unwrap();
    for case in 0..5 {
        let mut model = source.clone();
        match case {
            0 => model.metadata.default_paragraph_style_id = None,
            1 => model.metadata.paragraph_styles.retain(|style| style.id != "Normal"),
            2 => model.metadata.paragraph_styles.iter_mut().find(|style| style.id == "Normal").unwrap().is_default = None,
            3 => {
                let duplicate = model.metadata.paragraph_styles.iter().find(|style| style.id == "Normal").unwrap().clone();
                model.metadata.paragraph_styles.push(duplicate);
            },
            _ => model.metadata.paragraph_styles.iter_mut().find(|style| style.id != "Normal").unwrap().is_default = Some(true),
        }
        let before = model.clone();
        assert!(odt_defaults_from_model(&model).is_err());
        assert_eq!(model, before);
    }
}

#[test]
fn odt_bad_xml_and_missing_required_content_preserve_available_resources_and_errors() {
    let valid = odt_bytes(Some(&odt_styles(true)), Some(&odt_content()));
    let model = duke_pages::parse_odt(&valid).unwrap();
    for (styles, content) in [
        ("<!DOCTYPE styles><styles/>".to_owned(), Some(odt_content())),
        (odt_styles(true), None),
    ] {
        let bytes = odt_bytes(Some(&styles), content.as_deref());
        let mut archive = ZipArchive::new(Cursor::new(bytes.as_slice())).unwrap();
        let defaults = read_odt_default_style(&model, &mut archive);
        assert!(defaults.default_style.is_none());
        assert!(!defaults.source.resource_errors.is_empty());
        let origin = serde_json::to_value(defaults.source.origin.unwrap()).unwrap();
        assert_eq!(origin["resources"][0]["xml"], styles);
    }
}
