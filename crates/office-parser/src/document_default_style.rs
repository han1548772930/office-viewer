use std::collections::HashMap;

use duke_pages::{OoxmlPackage, OoxmlPart, parse::{ParsedStyleSheet, parse_style_sheet}};
use quick_xml::{NsReader, Writer, events::{BytesEnd, BytesStart, Event}, name::{QName, ResolveResult}};

use crate::model::{DocumentDefaultStylePayload, DocumentOfficeSourcePayload};
use super::{ParagraphWordWrap, parse_xml};

const WORD: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const DRAWING: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";

fn canonical_name(reader: &NsReader<&[u8]>, name: QName<'_>, attribute: bool) -> Result<String, String> {
    let (namespace, local) = if attribute { reader.resolver().resolve_attribute(name) } else { reader.resolver().resolve_element(name) };
    let original = std::str::from_utf8(name.as_ref()).map_err(|error| error.to_string())?;
    let local = std::str::from_utf8(local.as_ref()).map_err(|error| error.to_string())?;
    let prefix = match namespace {
        ResolveResult::Bound(namespace) => match namespace.as_ref() {
            value if value == WORD.as_bytes() || value == b"http://purl.oclc.org/ooxml/wordprocessingml/main" => Some("w"),
            value if value == DRAWING.as_bytes() || value == b"http://purl.oclc.org/ooxml/drawingml/main" => Some("a"),
            _ => None,
        },
        ResolveResult::Unbound => None,
        ResolveResult::Unknown(prefix) => return Err(format!("Default-style XML has an unbound prefix {:?}.", String::from_utf8_lossy(&prefix))),
    };
    if let Some(prefix) = prefix { return Ok(format!("{prefix}:{local}")); }
    if original.starts_with("w:") || original.starts_with("a:") { return Err("Default-style XML rebinds a canonical namespace prefix.".into()); }
    Ok(original.to_owned())
}

/// The style engine accepts canonical w/a prefixes. Normalize only XML names;
/// retain the original resource verbatim in officeSource for source fidelity.
pub(crate) fn canonical_resource(xml: &str, expected: &str) -> Result<String, String> {
    // Reuse the bounded project XML reader for well-formedness, depth and DTD checks.
    parse_xml(xml)?;
    let mut reader = NsReader::from_str(xml);
    let mut writer = Writer::new(Vec::new());
    let mut root_seen = false;
    loop {
        let event = reader.read_event().map_err(|error| error.to_string())?;
        let empty = matches!(&event, Event::Empty(_));
        let event = match event {
            Event::Start(tag) | Event::Empty(tag) => {
                let name = canonical_name(&reader, tag.name(), false)?;
                let root = !root_seen;
                if root && name != expected { return Err(format!("Default-style resource has root {name:?}, expected {expected:?}.")); }
                root_seen = true;
                let mut result = BytesStart::new(name);
                for attribute in tag.attributes() {
                    let attribute = attribute.map_err(|error| error.to_string())?;
                    let key = attribute.key.as_ref();
                    if key == b"xmlns:w" || key == b"xmlns:a" { continue; }
                    let name = if key == b"xmlns" || key.starts_with(b"xmlns:") {
                        std::str::from_utf8(key).map_err(|error| error.to_string())?.to_owned()
                    } else { canonical_name(&reader, attribute.key, true)? };
                    result.push_attribute((name.as_bytes(), attribute.value.as_ref()));
                }
                if root {
                    result.push_attribute(("xmlns:w", WORD));
                    result.push_attribute(("xmlns:a", DRAWING));
                }
                if empty { Event::Empty(result.into_owned()) } else { Event::Start(result.into_owned()) }
            },
            Event::End(tag) => Event::End(BytesEnd::new(canonical_name(&reader, tag.name(), false)?)),
            Event::DocType(_) => return Err("Default-style XML must not declare a document type.".into()),
            Event::Eof => break,
            event => event.into_owned(),
        };
        writer.write_event(event).map_err(|error| error.to_string())?;
    }
    String::from_utf8(writer.into_inner()).map_err(|error| error.to_string())
}

pub(crate) fn read_document_style_sheet(source: Option<&DocumentOfficeSourcePayload>) -> Result<ParsedStyleSheet, String> {
    if let Some(error) = source.and_then(|source| source.resource_errors.first()) { return Err(error.clone()); }
    let styles = source.and_then(|source| source.styles.as_ref()).map(|part| canonical_resource(&part.xml, "w:styles")).transpose()?
        .unwrap_or_else(|| format!("<w:styles xmlns:w=\"{WORD}\"/>"));
    let mut parts = HashMap::new();
    parts.insert("word/styles.xml".into(), OoxmlPart { name: "word/styles.xml".into(), content: styles });
    if let Some(theme) = source.and_then(|source| source.theme.as_ref()) {
        parts.insert("word/theme/theme1.xml".into(), OoxmlPart { name: "word/theme/theme1.xml".into(), content: canonical_resource(&theme.xml, "a:theme")? });
    }
    // Only styles and theme are needed; do not reopen or clone the entire Office package.
    Ok(parse_style_sheet(&OoxmlPackage { parts, binary_assets: HashMap::new(), warnings: Vec::new() }))
}

pub(crate) fn read_document_default_style(source: Option<&DocumentOfficeSourcePayload>, wrap: &ParagraphWordWrap) -> Result<DocumentDefaultStylePayload, String> {
    let sheet = read_document_style_sheet(source)?;
    Ok(DocumentDefaultStylePayload { version: 1, text: sheet.default_run_style, paragraph: sheet.default_paragraph_style, word_wrap: wrap.resolve("") })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::DocumentOfficeXmlPartPayload;

    #[test]
    fn defaults_resolve_based_on_and_theme_fonts_without_changing_source_xml() {
        let styles = format!(r#"<s:styles xmlns:s="{WORD}"><s:docDefaults><s:rPrDefault><s:rPr><s:rFonts s:asciiTheme="minorHAnsi" s:eastAsiaTheme="minorEastAsia"/><s:sz s:val="30"/></s:rPr></s:rPrDefault><s:pPrDefault><s:pPr><s:spacing s:line="288" s:after="180"/></s:pPr></s:pPrDefault></s:docDefaults><s:style s:type="paragraph" s:styleId="Base"><s:rPr><s:i/></s:rPr><s:pPr><s:jc s:val="right"/></s:pPr></s:style><s:style s:type="paragraph" s:default="1" s:styleId="Normal"><s:basedOn s:val="Base"/><s:rPr><s:sz s:val="32"/></s:rPr></s:style></s:styles>"#);
        let theme = format!(r#"<t:theme xmlns:t="{DRAWING}"><t:themeElements><t:fontScheme><t:minorFont><t:latin typeface="Georgia"/><t:ea typeface="宋体"/></t:minorFont></t:fontScheme></t:themeElements></t:theme>"#);
        let source = DocumentOfficeSourcePayload { version: 1, styles: Some(DocumentOfficeXmlPartPayload { part_name: "custom/styles.xml".into(), xml: styles.clone() }), theme: Some(DocumentOfficeXmlPartPayload { part_name: "custom/theme.xml".into(), xml: theme }), resource_errors: Vec::new() };
        let defaults = read_document_default_style(Some(&source), &ParagraphWordWrap::default()).unwrap();
        let text = defaults.text.unwrap();
        assert_eq!(text.font_size_pt, Some(16.0));
        assert_eq!(text.italic, Some(true));
        assert_eq!(text.resolved_font_family_ascii.as_deref(), Some("Georgia"));
        assert_eq!(text.resolved_font_family_east_asia.as_deref(), Some("宋体"));
        assert_eq!(defaults.paragraph.unwrap().align, Some(duke_pages::ParagraphAlignment::Right));
        assert_eq!(source.styles.unwrap().xml, styles);
    }

    #[test]
    fn absent_and_declared_empty_defaults_keep_the_readers_word_defaults() {
        let absent = read_document_default_style(None, &ParagraphWordWrap::default()).unwrap();
        assert_eq!(absent.text.unwrap().font_size_pt, Some(12.0));
        let source = DocumentOfficeSourcePayload { version: 1, styles: Some(DocumentOfficeXmlPartPayload { part_name: "word/styles.xml".into(), xml: format!(r#"<w:styles xmlns:w="{WORD}"><w:docDefaults><w:rPrDefault><w:rPr/></w:rPrDefault><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults></w:styles>"#) }), theme: None, resource_errors: Vec::new() };
        let declared = read_document_default_style(Some(&source), &ParagraphWordWrap::default()).unwrap();
        assert_eq!(declared.text.unwrap().font_size_pt, Some(10.0));
        assert!(declared.paragraph.is_none());
    }

    #[test]
    fn malformed_or_rebound_style_names_do_not_produce_silent_defaults() {
        assert!(canonical_resource("<!DOCTYPE styles><styles/>", "w:styles").is_err());
        assert!(canonical_resource("<w:styles xmlns:w=\"urn:other\"/>", "w:styles").is_err());
        assert!(canonical_resource("<w:styles>", "w:styles").is_err());
    }
}
