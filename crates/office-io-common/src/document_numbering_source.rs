//! Native numbering resources are provenance, not a second list evaluator.
//! Both Office readers and writers use this bounded control DTO and OPC paths.

use std::collections::BTreeSet;
use serde::{Deserialize, Serialize};
use base64::{Engine as _, engine::general_purpose::STANDARD};

pub const XML_BYTE_LIMIT: usize = 16 * 1024 * 1024;
pub const MEDIA_BYTE_LIMIT: usize = 16 * 1024 * 1024;
pub const TOTAL_BYTE_LIMIT: usize = 32 * 1024 * 1024;
pub const RESOURCE_COUNT_LIMIT: usize = 1024;
pub const NAME_BYTE_LIMIT: usize = 1024;

#[derive(Debug, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct DocumentNumberingSource {
    pub part_name: String,
    pub xml: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub relationships_xml: Option<String>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub resources: Vec<DocumentNumberingMedia>,
}

#[derive(Debug, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct DocumentNumberingMedia {
    pub part_name: String,
    pub content_type: String,
    pub data_base64: String,
}

pub fn package_part_name(name: &str) -> Result<String, String> {
    if name.is_empty() || name.len() > NAME_BYTE_LIMIT || name.contains(['\\', '\0', '?', '#', ':']) {
        return Err("DOCX_NUMBERING_RESOURCE_PATH: invalid Office package part name".into());
    }
    let name = name.trim_start_matches('/');
    if name.is_empty() || name.split('/').any(|part| matches!(part, "" | "." | "..")) {
        return Err("DOCX_NUMBERING_RESOURCE_PATH: Office part names must be canonical package paths".into());
    }
    Ok(name.to_owned())
}

pub fn resolve_relationship(owner: &str, target: &str) -> Result<String, String> {
    let owner = package_part_name(owner)?;
    if target.is_empty() || target.len() > NAME_BYTE_LIMIT || target.contains(['\\', '\0', '?', '#', ':']) {
        return Err("DOCX_NUMBERING_RESOURCE_PATH: invalid Office relationship target".into());
    }
    let directory = owner.rsplit_once('/').map_or("", |(directory, _)| directory);
    let path = if target.starts_with('/') || directory.is_empty() { target.to_owned() } else { format!("{directory}/{target}") };
    let mut parts = Vec::new();
    for part in path.split('/') {
        match part {
            "" | "." => {},
            ".." => { if parts.pop().is_none() { return Err("DOCX_NUMBERING_RESOURCE_PATH: relationship escapes the package root".into()); } },
            _ => parts.push(part),
        }
    }
    package_part_name(&parts.join("/"))
}

fn base64_digit(byte: u8) -> Option<u8> {
    match byte { b'A'..=b'Z' => Some(byte - b'A'), b'a'..=b'z' => Some(byte - b'a' + 26),
        b'0'..=b'9' => Some(byte - b'0' + 52), b'+' => Some(62), b'/' => Some(63), _ => None }
}

/// Validate canonical base64 and measure it without allocating decoded media.
pub fn decoded_media_bytes(value: &str) -> Result<usize, String> {
    let bytes = value.as_bytes();
    if bytes.is_empty() || bytes.len() % 4 != 0 || bytes.len() > MEDIA_BYTE_LIMIT.div_ceil(3) * 4 {
        return Err("DOCX_NUMBERING_MEDIA_DATA: invalid or oversized base64 media".into());
    }
    let padding = if bytes.ends_with(b"==") { 2 } else { usize::from(bytes.ends_with(b"=")) };
    let end = bytes.len() - padding;
    if bytes[..end].iter().any(|byte| base64_digit(*byte).is_none())
        || padding == 2 && base64_digit(bytes[end - 1]).is_none_or(|digit| digit & 15 != 0)
        || padding == 1 && base64_digit(bytes[end - 1]).is_none_or(|digit| digit & 3 != 0) {
        return Err("DOCX_NUMBERING_MEDIA_DATA: media must use canonical base64 padding".into());
    }
    let size = bytes.len() / 4 * 3 - padding;
    if size > MEDIA_BYTE_LIMIT { return Err("DOCX_NUMBERING_MEDIA_LIMIT: one numbering image exceeds 16 MiB".into()); }
    Ok(size)
}

pub fn encode_media(bytes: &[u8]) -> Result<String, String> {
    if bytes.is_empty() || bytes.len() > MEDIA_BYTE_LIMIT { return Err("DOCX_NUMBERING_MEDIA_LIMIT: numbering image is empty or exceeds 16 MiB".into()); }
    let mut encoded = String::new();
    encoded.try_reserve_exact(bytes.len().div_ceil(3) * 4).map_err(|_| "DOCX_NUMBERING_MEDIA_LIMIT: unable to allocate encoded numbering media")?;
    STANDARD.encode_string(bytes, &mut encoded);
    Ok(encoded)
}

pub fn decode_media(value: &str) -> Result<Vec<u8>, String> {
    let length = decoded_media_bytes(value)?;
    let mut bytes = Vec::new();
    bytes.try_reserve_exact(length).map_err(|_| "DOCX_NUMBERING_MEDIA_LIMIT: unable to allocate decoded numbering media")?;
    bytes.resize(length, 0);
    let written = STANDARD.decode_slice(value, &mut bytes).map_err(|error| format!("DOCX_NUMBERING_MEDIA_DATA: {error}"))?;
    if written != length { return Err("DOCX_NUMBERING_MEDIA_DATA: decoded numbering media length changed".into()); }
    Ok(bytes)
}

impl DocumentNumberingSource {
    pub fn validate(&self) -> Result<(), String> {
        package_part_name(&self.part_name)?;
        if self.xml.len() > XML_BYTE_LIMIT || self.relationships_xml.as_ref().is_some_and(|xml| xml.len() > XML_BYTE_LIMIT)
            || self.resources.len() > RESOURCE_COUNT_LIMIT {
            return Err("DOCX_NUMBERING_RESOURCE_LIMIT: numbering XML or media count exceeds its budget".into());
        }
        let mut total = self.xml.len().checked_add(self.relationships_xml.as_ref().map_or(0, String::len))
            .filter(|total| *total <= TOTAL_BYTE_LIMIT).ok_or("DOCX_NUMBERING_RESOURCE_LIMIT: numbering source exceeds 32 MiB")?;
        let mut names = BTreeSet::new();
        for resource in &self.resources {
            let path = package_part_name(&resource.part_name)?;
            if !names.insert(path) { return Err("DOCX_NUMBERING_MEDIA_DUPLICATE: a source media part occurs more than once".into()); }
            if resource.content_type.is_empty() || resource.content_type.len() > NAME_BYTE_LIMIT || resource.content_type.contains(['\0', '\r', '\n']) {
                return Err("DOCX_NUMBERING_MEDIA_TYPE: numbering media has an invalid content type".into());
            }
            total = total.checked_add(decoded_media_bytes(&resource.data_base64)?)
                .filter(|total| *total <= TOTAL_BYTE_LIMIT).ok_or("DOCX_NUMBERING_RESOURCE_LIMIT: numbering source exceeds 32 MiB")?;
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn relationship_paths_resolve_inside_the_package() {
        assert_eq!(resolve_relationship("parts/lists/numbering.xml", "../../media/bullet.png").unwrap(), "media/bullet.png");
        assert_eq!(resolve_relationship("word/numbering.xml", "/media/bullet.png").unwrap(), "media/bullet.png");
        for target in ["../../outside.png", "media\\bullet.png", "https://example.test/bullet.png", "media/bullet.png#fragment"] {
            assert!(resolve_relationship("word/numbering.xml", target).is_err());
        }
    }

    #[test]
    fn base64_budget_has_canonical_padding_and_no_decoded_allocation() {
        for (encoded, bytes) in [("AQ==", 1), ("AQI=", 2), ("AQID", 3)] {
            assert_eq!(decoded_media_bytes(encoded).unwrap(), bytes);
            assert_eq!(encode_media(&decode_media(encoded).unwrap()).unwrap(), encoded);
        }
        for invalid in ["", "AQ", "AR==", "AQJ=", "AQ==\n", "=Q==", "A==="] { assert!(decoded_media_bytes(invalid).is_err(), "{invalid:?}"); }
    }

    #[test]
    fn duplicate_media_and_unknown_source_fields_are_not_silently_retained() {
        let source = r#"{"partName":"word/numbering.xml","xml":"<numbering/>","resources":[{"partName":"word/media/a.png","contentType":"image/png","dataBase64":"AQ=="},{"partName":"/word/media/a.png","contentType":"image/png","dataBase64":"AQ=="}]}"#;
        assert!(serde_json::from_str::<DocumentNumberingSource>(source).unwrap().validate().unwrap_err().contains("DUPLICATE"));
        assert!(serde_json::from_str::<DocumentNumberingSource>(r#"{"partName":"word/numbering.xml","xml":"<numbering/>","oldEngine":{}}"#).is_err());
    }
}
