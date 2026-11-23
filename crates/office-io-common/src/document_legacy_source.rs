//! The application's lossless provenance part. These are historical resources,
//! not Word styles/theme parts and not a claim that their semantics are rendered.

use std::{fmt, io::{self, Write}, marker::PhantomData};

use serde::{Deserialize, Deserializer, Serialize, de::{self, IgnoredAny, SeqAccess, Visitor}};
use serde_json::Value;

pub const LEGACY_SOURCE_PART: &str = "word/office-studio/legacy-sources.json";
pub const LEGACY_SOURCE_TARGET: &str = "office-studio/legacy-sources.json";
pub const LEGACY_SOURCE_CONTENT_TYPE: &str = "application/vnd.office-studio.document-legacy-sources+json";
pub const LEGACY_SOURCE_RELATIONSHIP: &str = "urn:office-studio:relationships:document-legacy-sources:1";

pub const XML_BYTE_LIMIT: usize = 16 * 1024 * 1024;
pub const BINARY_BYTE_LIMIT: usize = 16 * 1024 * 1024;
pub const TOTAL_RESOURCE_BYTE_LIMIT: usize = 32 * 1024 * 1024;
pub const JSON_BYTE_LIMIT: usize = 64 * 1024 * 1024;
pub const ORIGIN_COUNT_LIMIT: usize = 32;
pub const RESOURCE_COUNT_LIMIT: usize = 1_024;
pub const NAME_BYTE_LIMIT: usize = 1_024;
pub const ERROR_COUNT_LIMIT: usize = 1_024;
pub const ERROR_BYTE_LIMIT: usize = 4_096;
pub const TOTAL_ERROR_BYTE_LIMIT: usize = 1024 * 1024;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum DocumentOfficeSourceOriginFormat { Doc, Odt }

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(tag = "kind", rename_all = "lowercase", rename_all_fields = "camelCase", deny_unknown_fields)]
pub enum DocumentOfficeSourceResource {
    Xml {
        #[serde(deserialize_with = "deserialize_name")]
        part_name: String,
        #[serde(deserialize_with = "deserialize_xml")]
        xml: String,
    },
    Binary {
        #[serde(deserialize_with = "deserialize_name")]
        stream_name: String,
        offset: u32,
        #[serde(deserialize_with = "deserialize_hex")]
        data_hex: String,
    },
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct DocumentOfficeSourceOrigin {
    pub format: DocumentOfficeSourceOriginFormat,
    pub synthesized: bool,
    #[serde(deserialize_with = "deserialize_resources")]
    pub resources: Vec<DocumentOfficeSourceResource>,
    /// Diagnostics belong to this historical source. They must not prevent the
    /// new DOCX's independent Word default-style chain from being resolved.
    #[serde(default, skip_serializing_if = "Vec::is_empty", deserialize_with = "deserialize_errors")]
    pub resource_errors: Vec<String>,
}

#[derive(Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct LegacySourcePart {
    version: u32,
    #[serde(deserialize_with = "deserialize_origins")]
    legacy_origins: Vec<DocumentOfficeSourceOrigin>,
}

struct LimitedString<const LIMIT: usize>;

impl<'de, const LIMIT: usize> Visitor<'de> for LimitedString<LIMIT> {
    type Value = String;
    fn expecting(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result { write!(formatter, "a UTF-8 string of at most {LIMIT} bytes") }
    fn visit_str<E: de::Error>(self, value: &str) -> Result<String, E> {
        if value.len() > LIMIT { return Err(E::custom("Office source string exceeds its byte limit.")); }
        Ok(value.to_owned())
    }
    fn visit_string<E: de::Error>(self, value: String) -> Result<String, E> {
        if value.len() > LIMIT { return Err(E::custom("Office source string exceeds its byte limit.")); }
        Ok(value)
    }
}

fn deserialize_name<'de, D: Deserializer<'de>>(deserializer: D) -> Result<String, D::Error> { deserializer.deserialize_string(LimitedString::<NAME_BYTE_LIMIT>) }
fn deserialize_xml<'de, D: Deserializer<'de>>(deserializer: D) -> Result<String, D::Error> { deserializer.deserialize_string(LimitedString::<XML_BYTE_LIMIT>) }
fn deserialize_hex<'de, D: Deserializer<'de>>(deserializer: D) -> Result<String, D::Error> { deserializer.deserialize_string(LimitedString::<{ BINARY_BYTE_LIMIT * 2 }>) }

struct LimitedVec<T, const LIMIT: usize>(PhantomData<T>);

impl<'de, T: Deserialize<'de>, const LIMIT: usize> Visitor<'de> for LimitedVec<T, LIMIT> {
    type Value = Vec<T>;
    fn expecting(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result { write!(formatter, "an array of at most {LIMIT} entries") }
    fn visit_seq<A: SeqAccess<'de>>(self, mut access: A) -> Result<Vec<T>, A::Error> {
        if access.size_hint().is_some_and(|length| length > LIMIT) { return Err(de::Error::custom("Office source array exceeds its entry limit.")); }
        let mut values = Vec::new();
        for _ in 0..LIMIT {
            match access.next_element()? { Some(value) => values.push(value), None => return Ok(values) }
        }
        if access.next_element::<IgnoredAny>()?.is_some() { return Err(de::Error::custom("Office source array exceeds its entry limit.")); }
        Ok(values)
    }
}

fn deserialize_resources<'de, D: Deserializer<'de>>(deserializer: D) -> Result<Vec<DocumentOfficeSourceResource>, D::Error> {
    deserializer.deserialize_seq(LimitedVec::<DocumentOfficeSourceResource, RESOURCE_COUNT_LIMIT>(PhantomData))
}
fn deserialize_origins<'de, D: Deserializer<'de>>(deserializer: D) -> Result<Vec<DocumentOfficeSourceOrigin>, D::Error> {
    deserializer.deserialize_seq(LimitedVec::<DocumentOfficeSourceOrigin, ORIGIN_COUNT_LIMIT>(PhantomData))
}

#[derive(Deserialize)]
#[serde(transparent)]
struct ResourceError(#[serde(deserialize_with = "deserialize_error")] String);
fn deserialize_error<'de, D: Deserializer<'de>>(deserializer: D) -> Result<String, D::Error> { deserializer.deserialize_string(LimitedString::<ERROR_BYTE_LIMIT>) }
fn deserialize_errors<'de, D: Deserializer<'de>>(deserializer: D) -> Result<Vec<String>, D::Error> {
    let errors = deserializer.deserialize_seq(LimitedVec::<ResourceError, ERROR_COUNT_LIMIT>(PhantomData))?;
    Ok(errors.into_iter().map(|error| error.0).collect())
}

pub fn read_origin_value(value: &Value) -> Result<DocumentOfficeSourceOrigin, String> {
    let origin = DocumentOfficeSourceOrigin::deserialize(value).map_err(|error| error.to_string())?;
    validate_origin(&origin)?;
    Ok(origin)
}

pub fn read_origin_list_value(value: &Value) -> Result<Vec<DocumentOfficeSourceOrigin>, String> {
    #[derive(Deserialize)]
    #[serde(transparent)]
    struct Origins(#[serde(deserialize_with = "deserialize_origins")] Vec<DocumentOfficeSourceOrigin>);
    Origins::deserialize(value).map(|origins| origins.0).map_err(|error| error.to_string())
}

pub fn read_resource_errors_value(value: &Value) -> Result<Vec<String>, String> {
    #[derive(Deserialize)]
    #[serde(transparent)]
    struct Errors(#[serde(deserialize_with = "deserialize_errors")] Vec<String>);
    let errors = Errors::deserialize(value).map_err(|error| error.to_string())?.0;
    validate_errors(&errors)?;
    Ok(errors)
}

fn validate_name(name: &str) -> Result<(), String> {
    if name.is_empty() || name.len() > NAME_BYTE_LIMIT || name.contains('\0') { return Err("Office source resource name is invalid.".into()); }
    Ok(())
}

fn validate_errors(errors: &[String]) -> Result<(), String> {
    if errors.len() > ERROR_COUNT_LIMIT || errors.iter().any(|error| error.len() > ERROR_BYTE_LIMIT)
        || errors.iter().map(String::len).sum::<usize>() > TOTAL_ERROR_BYTE_LIMIT {
        return Err("Office source diagnostics exceed their entry or byte limit.".into());
    }
    Ok(())
}

#[derive(Default)]
struct Budget { resources: usize, binary_bytes: usize, raw_bytes: usize, errors: usize, error_bytes: usize }

impl Budget {
    fn observe(&mut self, origin: &DocumentOfficeSourceOrigin) -> Result<(), String> {
        if origin.synthesized != (origin.format == DocumentOfficeSourceOriginFormat::Doc) { return Err("Office source format and synthesized marker disagree.".into()); }
        validate_errors(&origin.resource_errors)?;
        self.errors = self.errors.checked_add(origin.resource_errors.len()).ok_or("Office source diagnostics count overflows.")?;
        self.error_bytes = self.error_bytes.checked_add(origin.resource_errors.iter().map(String::len).sum()).ok_or("Office source diagnostics size overflows.")?;
        if self.errors > ERROR_COUNT_LIMIT || self.error_bytes > TOTAL_ERROR_BYTE_LIMIT { return Err("Combined Office source diagnostics exceed their budget.".into()); }
        for resource in &origin.resources {
            self.resources += 1;
            if self.resources > RESOURCE_COUNT_LIMIT { return Err("Combined Office source resources exceed 1024 entries.".into()); }
            let bytes = match (origin.format, resource) {
                (DocumentOfficeSourceOriginFormat::Odt, DocumentOfficeSourceResource::Xml { part_name, xml }) => {
                    validate_name(part_name)?;
                    if xml.len() > XML_BYTE_LIMIT { return Err("Office source XML exceeds the 16 MiB part limit.".into()); }
                    xml.len()
                },
                (DocumentOfficeSourceOriginFormat::Doc, DocumentOfficeSourceResource::Binary { stream_name, offset, data_hex }) => {
                    validate_name(stream_name)?;
                    if data_hex.is_empty() || data_hex.len() % 2 != 0 || data_hex.len() > BINARY_BYTE_LIMIT * 2
                        || !data_hex.bytes().all(|byte| byte.is_ascii_hexdigit()) { return Err("Office source binary hex is invalid.".into()); }
                    let bytes = data_hex.len() / 2;
                    let length = u32::try_from(bytes).map_err(|_| "Office source binary length exceeds u32.")?;
                    offset.checked_add(length).ok_or("Office source binary offset and length overflow u32.")?;
                    self.binary_bytes = self.binary_bytes.checked_add(bytes).ok_or("Office source binary size overflows.")?;
                    if self.binary_bytes > BINARY_BYTE_LIMIT { return Err("Combined Office source binary resources exceed 16 MiB.".into()); }
                    bytes
                },
                _ => return Err("Office source resource kind does not match its original format.".into()),
            };
            self.raw_bytes = self.raw_bytes.checked_add(bytes).ok_or("Office source resource size overflows.")?;
            if self.raw_bytes > TOTAL_RESOURCE_BYTE_LIMIT { return Err("Combined Office source resources exceed 32 MiB.".into()); }
        }
        Ok(())
    }
}

pub fn validate_origin(origin: &DocumentOfficeSourceOrigin) -> Result<(), String> { Budget::default().observe(origin) }

pub fn validate_origins(origins: &[DocumentOfficeSourceOrigin]) -> Result<(), String> {
    if origins.len() > ORIGIN_COUNT_LIMIT { return Err("Combined Office sources exceed 32 origins.".into()); }
    let mut budget = Budget::default();
    for origin in origins { budget.observe(origin)?; }
    Ok(())
}

fn same_resource(left: &DocumentOfficeSourceResource, right: &DocumentOfficeSourceResource) -> bool {
    match (left, right) {
        (DocumentOfficeSourceResource::Xml { part_name: a, xml: x }, DocumentOfficeSourceResource::Xml { part_name: b, xml: y }) => a == b && x == y,
        (DocumentOfficeSourceResource::Binary { stream_name: a, offset: x, data_hex: p }, DocumentOfficeSourceResource::Binary { stream_name: b, offset: y, data_hex: q }) => a == b && x == y && p.eq_ignore_ascii_case(q),
        _ => false,
    }
}

fn same_origin(left: &DocumentOfficeSourceOrigin, right: &DocumentOfficeSourceOrigin) -> bool {
    left.format == right.format && left.synthesized == right.synthesized
        && left.resources.len() == right.resources.len()
        && left.resources.iter().zip(&right.resources).all(|(a, b)| same_resource(a, b))
}

/// Preserve the first source's resource order/text. Hex case has no effect on
/// byte identity; diagnostics are unioned rather than silently discarded.
pub fn merge_origins(origins: Vec<DocumentOfficeSourceOrigin>) -> Result<Vec<DocumentOfficeSourceOrigin>, String> {
    // The writer accepts one current origin plus the bounded historical array.
    if origins.len() > ORIGIN_COUNT_LIMIT + 1 { return Err("Office source input exceeds the bounded origin count.".into()); }
    let mut merged: Vec<DocumentOfficeSourceOrigin> = Vec::new();
    for origin in origins {
        validate_origin(&origin)?;
        if let Some(existing) = merged.iter_mut().find(|existing| same_origin(existing, &origin)) {
            for error in origin.resource_errors { if !existing.resource_errors.contains(&error) { existing.resource_errors.push(error); } }
        } else { merged.push(origin); }
    }
    validate_origins(&merged)?;
    Ok(merged)
}

struct LimitedOutput { bytes: Vec<u8> }
impl Write for LimitedOutput {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.bytes.len().checked_add(bytes.len()).is_none_or(|size| size > JSON_BYTE_LIMIT) {
            return Err(io::Error::new(io::ErrorKind::InvalidData, "Office source JSON exceeds the 64 MiB limit."));
        }
        self.bytes.extend_from_slice(bytes);
        Ok(bytes.len())
    }
    fn flush(&mut self) -> io::Result<()> { Ok(()) }
}

pub fn encode_origins(origins: Vec<DocumentOfficeSourceOrigin>) -> Result<Vec<u8>, String> {
    let legacy_origins = merge_origins(origins)?;
    let part = LegacySourcePart { version: 1, legacy_origins };
    let mut output = LimitedOutput { bytes: Vec::new() };
    serde_json::to_writer(&mut output, &part).map_err(|error| error.to_string())?;
    Ok(output.bytes)
}

pub fn decode_origins(bytes: &[u8]) -> Result<Vec<DocumentOfficeSourceOrigin>, String> {
    if bytes.len() > JSON_BYTE_LIMIT { return Err("Office source JSON exceeds the 64 MiB limit.".into()); }
    let part: LegacySourcePart = serde_json::from_slice(bytes).map_err(|error| error.to_string())?;
    if part.version != 1 { return Err("Office source JSON has an unsupported version.".into()); }
    merge_origins(part.legacy_origins)
}

#[cfg(test)]
#[path = "document_legacy_source_tests.rs"]
mod tests;
