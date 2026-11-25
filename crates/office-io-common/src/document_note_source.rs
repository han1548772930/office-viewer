//! Shared Word note properties. The authored values and unmodeled source XML
//! remain separate so resetting an explicit value cannot resurrect a stale one.
use serde::{Deserialize, Serialize};
use serde_json::{Value, json};

use crate::{chart_xml, document_content_controls::{word, word_attribute}};

pub const WORD_2013_NAMESPACE: &str = "http://schemas.microsoft.com/office/word/2012/wordml";
pub const NOTE_EXTENSION_NAMESPACE: &str = "urn:office-studio:document-notes:1";
pub const NOTE_SAFE_INTEGER: u64 = 9_007_199_254_740_991;
pub const NOTE_DECORATION_TYPES: [&str; 3] = ["separator", "continuationSeparator", "continuationNotice"];

/// Public note IDs are UTF-16 strings, independent of the native numeric ID.
pub fn validate_public_note_source_id(value: &str) -> Result<(), String> {
    if value.trim().is_empty() || value.encode_utf16().count() > 4096 || value.contains(['\r', '\n', '\0']) {
        return Err("DOCX_NOTE_SOURCE_INVALID: public source ID must be a nonempty bounded string without CR, LF or NUL".into());
    }
    Ok(())
}

pub fn public_note_decoration_source_id<'a>(value: &'a Value, kind: DocumentNoteKind, note_type: &str) -> Result<&'a str, String> {
    if !value.is_object() || value.get("version").and_then(Value::as_u64) != Some(1)
        || value.get("kind").and_then(Value::as_str) != Some(kind.name())
        || !NOTE_DECORATION_TYPES.contains(&note_type)
        || value.get("type").and_then(Value::as_str) != Some(note_type) {
        return Err("DOCX_NOTE_SOURCE_INVALID: public decoration source version, kind or type is invalid".into());
    }
    let id = value.get("noteId").and_then(Value::as_str).ok_or("DOCX_NOTE_SOURCE_INVALID: public decoration source ID must be a string")?;
    validate_public_note_source_id(id)?;
    Ok(id)
}

/// XML attributes normalize literal whitespace and cannot carry every valid
/// public ID character. Hex encodes the exact UTF-8 ID without normalizing it.
pub fn encode_public_note_source_id(value: &str) -> String {
    const HEX: &[u8; 16] = b"0123456789abcdef";
    value.as_bytes().iter().flat_map(|byte| [HEX[(byte >> 4) as usize] as char, HEX[(byte & 15) as usize] as char]).collect()
}

pub fn decode_public_note_source_id(value: &str) -> Result<String, String> {
    if value.is_empty() || value.len() > 4096 * 4 * 2 || value.len() % 2 != 0 {
        return Err("DOCX_NOTE_SOURCE_INVALID: public source ID encoding exceeds its contract".into());
    }
    let bytes = value.as_bytes().chunks_exact(2).map(|pair| {
        let high = (pair[0] as char).to_digit(16).ok_or("DOCX_NOTE_SOURCE_INVALID: public source ID encoding is not hexadecimal")?;
        let low = (pair[1] as char).to_digit(16).ok_or("DOCX_NOTE_SOURCE_INVALID: public source ID encoding is not hexadecimal")?;
        Ok((high * 16 + low) as u8)
    }).collect::<Result<Vec<_>, String>>()?;
    let id = String::from_utf8(bytes).map_err(|_| "DOCX_NOTE_SOURCE_INVALID: public source ID encoding is not UTF-8")?;
    validate_public_note_source_id(&id)?;
    Ok(id)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum DocumentNoteKind { Footnote, Endnote }

impl DocumentNoteKind {
    pub fn name(self) -> &'static str { if self == Self::Footnote { "footnote" } else { "endnote" } }
    pub fn collection(self) -> &'static str { if self == Self::Footnote { "footnotes" } else { "endnotes" } }
    pub fn properties_tag(self) -> &'static str { if self == Self::Footnote { "footnotePr" } else { "endnotePr" } }
    pub fn reference_tag(self) -> &'static str { if self == Self::Footnote { "footnoteReference" } else { "endnoteReference" } }
    pub fn body_reference_tag(self) -> &'static str { if self == Self::Footnote { "footnoteRef" } else { "endnoteRef" } }
}

#[derive(Debug, Default, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DocumentNoteProperties {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub position: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub restart: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub start_number: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub number_format: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub column_count: Option<u8>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub metadata: Option<Value>,
}

#[derive(Debug, Default, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct DocumentSectionNoteProperties {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub footnote: Option<DocumentNoteProperties>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub endnote: Option<DocumentNoteProperties>,
}

pub struct NotePropertyRead { pub properties: DocumentNoteProperties, pub warnings: Vec<String> }

pub fn read_properties(xml: &str, kind: DocumentNoteKind, part_name: &str) -> Result<NotePropertyRead, String> {
    let root = chart_xml::parse(xml)?;
    if !word(&root, kind.properties_tag()) { return Err("DOCX_NOTE_SETTINGS_INVALID: note property root has the wrong namespace or kind".into()); }
    let mut properties = DocumentNoteProperties::default();
    let mut warnings = Vec::new();
    if let Some(value) = root.attribute_ns(NOTE_EXTENSION_NAMESPACE, "columnCount") {
        let count = value.parse::<u8>().ok().filter(|value| kind == DocumentNoteKind::Footnote && *value <= 4)
            .ok_or("DOCX_NOTE_SETTINGS_INVALID: global footnote columns must be between zero and four")?;
        properties.column_count = Some(count);
    }
    for name in ["pos", "numFmt", "numStart", "numRestart"] {
        let mut nodes = root.children.iter().filter(|node| word(node, name));
        let Some(node) = nodes.next() else { continue; };
        if nodes.next().is_some() { return Err(format!("DOCX_NOTE_SETTINGS_INVALID: duplicate {name}")); }
        let value = word_attribute(node, "val").ok_or_else(|| format!("DOCX_NOTE_SETTINGS_INVALID: {name} has no Word value"))?;
        match name {
            "pos" => {
                let allowed = if kind == DocumentNoteKind::Footnote { ["pageBottom", "beneathText"] } else { ["docEnd", "sectEnd"] };
                if allowed.contains(&value) { properties.position = Some(value.into()); }
                else { warnings.push(format!("Unrecognized {} placement {value:?} remains in native property provenance.", kind.name())); }
            }
            "numRestart" => {
                if matches!(value, "continuous" | "eachSect") || kind == DocumentNoteKind::Footnote && value == "eachPage" {
                    properties.restart = Some(value.into());
                } else { warnings.push(format!("Unrecognized {} numbering restart {value:?} remains in native property provenance.", kind.name())); }
            }
            "numStart" => {
                let number = value.parse::<u64>().ok().filter(|value| *value > 0 && *value <= NOTE_SAFE_INTEGER)
                    .ok_or("DOCX_NOTE_SETTINGS_INVALID: starting number must be a positive JavaScript safe integer")?;
                properties.start_number = Some(number);
            }
            "numFmt" => {
                if value.is_empty() || value.len() > 4096 { return Err("DOCX_NOTE_SETTINGS_INVALID: numbering format is empty or exceeds its contract".into()); }
                // Word has a larger format vocabulary than the current picker.
                properties.number_format = Some(value.into());
            }
            _ => unreachable!(),
        }
    }
    properties.metadata = Some(json!({ "officeNoteProperties": { "version": 1, "kind": kind,
        "partName": part_name, "xml": xml, "baseline": {
            "position": properties.position, "restart": properties.restart,
            "startNumber": properties.start_number, "numberFormat": properties.number_format, "columnCount": properties.column_count
        } } }));
    Ok(NotePropertyRead { properties, warnings })
}

pub fn validate_properties(properties: &DocumentNoteProperties, kind: DocumentNoteKind) -> Result<(), String> {
    if properties.position.as_deref().is_some_and(|value| !(if kind == DocumentNoteKind::Footnote { ["pageBottom", "beneathText"] } else { ["docEnd", "sectEnd"] }).contains(&value)) {
        return Err("DOCX_NOTE_SETTINGS_INVALID: placement is invalid for this note kind".into());
    }
    if properties.restart.as_deref().is_some_and(|value| !matches!(value, "continuous" | "eachSect") && !(kind == DocumentNoteKind::Footnote && value == "eachPage")) {
        return Err("DOCX_NOTE_SETTINGS_INVALID: numbering restart is invalid for this note kind".into());
    }
    if properties.start_number.is_some_and(|value| value == 0 || value > NOTE_SAFE_INTEGER) {
        return Err("DOCX_NOTE_SETTINGS_INVALID: starting number must be a positive JavaScript safe integer".into());
    }
    if properties.number_format.as_deref().is_some_and(|value| value.is_empty() || value.len() > 4096) {
        return Err("DOCX_NOTE_SETTINGS_INVALID: numbering format is empty or exceeds its contract".into());
    }
    if properties.column_count.is_some_and(|value| kind != DocumentNoteKind::Footnote || value > 4) {
        return Err("DOCX_NOTE_SETTINGS_INVALID: footnote columns must be between zero and four".into());
    }
    if properties.metadata.as_ref().is_some_and(|value| !value.is_object()) {
        return Err("DOCX_NOTE_SETTINGS_INVALID: provenance must be an object".into());
    }
    Ok(())
}

pub fn on_off(value: Option<&str>) -> Result<bool, String> {
    match value {
        None | Some("true" | "1" | "on") => Ok(true),
        Some("false" | "0" | "off") => Ok(false),
        _ => Err("DOCX_NOTE_SETTINGS_INVALID: on/off property is invalid".into()),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    // Drafts for the unified verification stage.
    #[test]
    fn public_source_ids_keep_unicode_and_attribute_whitespace_without_native_id_coercion() {
        for id in ["-1", "0007", "引用 😀", "\t source\tid "] {
            let encoded = encode_public_note_source_id(id);
            assert_eq!(decode_public_note_source_id(&encoded).unwrap(), id);
            let source = json!({ "version": 1, "kind": "footnote", "type": "separator", "noteId": id });
            assert_eq!(public_note_decoration_source_id(&source, DocumentNoteKind::Footnote, "separator").unwrap(), id);
        }
        for id in ["", " \t", "a\nb", "a\rb", "a\0b"] { assert!(validate_public_note_source_id(id).is_err()); }
        assert!(validate_public_note_source_id(&"😀".repeat(2048)).is_ok());
        assert!(validate_public_note_source_id(&"😀".repeat(2049)).is_err());
        for encoded in ["0", "zz", "ff", "c080", "00"] { assert!(decode_public_note_source_id(encoded).is_err()); }
        let source = json!({ "version": 1, "kind": "footnote", "type": "separator", "noteId": "-1" });
        assert!(public_note_decoration_source_id(&source, DocumentNoteKind::Endnote, "separator").is_err());
        assert!(public_note_decoration_source_id(&source, DocumentNoteKind::Footnote, "continuationSeparator").is_err());
    }

    #[test]
    fn properties_use_expanded_names_and_preserve_unknown_formats_and_extensions() {
        let xml = "<x:footnotePr xmlns:x='http://purl.oclc.org/ooxml/wordprocessingml/main' xmlns:q='urn:extension'><q:numStart q:val='900'/><x:numStart x:val='100000'/><x:numFmt x:val='hebrew1'/><x:pos x:val='beneathText'/><x:numRestart x:val='eachPage'/></x:footnotePr>";
        let parsed = read_properties(xml, DocumentNoteKind::Footnote, "parts/notes.xml").unwrap();
        assert_eq!(parsed.properties.start_number, Some(100000));
        assert_eq!(parsed.properties.number_format.as_deref(), Some("hebrew1"));
        assert_eq!(parsed.properties.position.as_deref(), Some("beneathText"));
        assert_eq!(parsed.properties.restart.as_deref(), Some("eachPage"));
        assert_eq!(parsed.properties.metadata.unwrap()["officeNoteProperties"]["xml"], xml);
    }

    #[test]
    fn unsupported_placement_is_retained_without_activating_another_notes_placement() {
        let xml = "<w:endnotePr xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'><w:pos w:val='pageBottom'/><w:numRestart w:val='eachPage'/></w:endnotePr>";
        let parsed = read_properties(xml, DocumentNoteKind::Endnote, "word/settings.xml").unwrap();
        assert!(parsed.properties.position.is_none());
        assert!(parsed.properties.restart.is_none());
        assert_eq!(parsed.warnings.len(), 2);
    }

    #[test]
    fn source_and_authored_start_numbers_share_the_safe_integer_contract() {
        for value in ["0", "-1", "9007199254740992", "1.5"] {
            let xml = format!("<w:footnotePr xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'><w:numStart w:val='{value}'/></w:footnotePr>");
            assert!(read_properties(&xml, DocumentNoteKind::Footnote, "word/settings.xml").is_err());
        }
        let properties = DocumentNoteProperties { start_number: Some(NOTE_SAFE_INTEGER), ..Default::default() };
        assert!(validate_properties(&properties, DocumentNoteKind::Footnote).is_ok());
    }
}
