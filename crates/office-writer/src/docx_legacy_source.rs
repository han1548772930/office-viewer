use office_io_common::document_legacy_source::{
    DocumentOfficeSourceOriginFormat, NAME_BYTE_LIMIT, XML_BYTE_LIMIT,
    encode_origins, read_origin_list_value, read_origin_value, read_resource_errors_value,
};
use serde_json::Value;
use serde::Deserialize;

use super::get;
use crate::ExportError;

fn xml_part(value: &Value) -> Result<(), ExportError> {
    if value.is_null() { return Ok(()); }
    let part = value.as_object().ok_or_else(|| ExportError::invalid("Office source XML part must be an object."))?;
    if part.keys().any(|key| !matches!(key.as_str(), "partName" | "xml")) {
        return Err(ExportError::invalid("Office source XML part contains unsupported fields; refusing to discard them."));
    }
    let name = get(value, "partName").as_str().ok_or_else(|| ExportError::invalid("Office source XML part has no name."))?;
    if name.is_empty() || name.len() > NAME_BYTE_LIMIT || name.contains('\0') { return Err(ExportError::invalid("Office source XML part name is invalid.")); }
    let xml = get(value, "xml").as_str().ok_or_else(|| ExportError::invalid("Office source XML part has no XML."))?;
    if xml.len() > XML_BYTE_LIMIT { return Err(ExportError::invalid("Office source XML exceeds the 16 MiB part limit.")); }
    Ok(())
}

/// Complete validation and serialization before the writer mutates its package.
/// The current DOC/ODT origin becomes history. DOCX's native Word resources do
/// not become historical DOC/ODT resources merely because history exists.
pub(super) fn prepare(model: &Value) -> Result<Option<Vec<u8>>, ExportError> {
    let source = get(get(get(model, "metadata"), "source"), "officeSource");
    if source.is_null() { return Ok(None); }
    let object = source.as_object().ok_or_else(|| ExportError::invalid("Office source must be an object."))?;
    if get(source, "version").as_u64() != Some(1) { return Err(ExportError::invalid("Office source has an unsupported version.")); }
    if object.keys().any(|key| !matches!(key.as_str(), "version" | "styles" | "theme" | "numbering" | "resourceErrors" | "origin" | "legacyOrigins")) {
        return Err(ExportError::invalid("Office source contains unsupported fields; refusing to discard them."));
    }
    xml_part(get(source, "styles"))?;
    xml_part(get(source, "theme"))?;
    let numbering = get(source, "numbering");
    if !numbering.is_null() {
        let numbering = office_io_common::document_numbering_source::DocumentNumberingSource::deserialize(numbering)
            .map_err(|error| ExportError::invalid(format!("DOCX_NUMBERING_SOURCE: {error}")))?;
        numbering.validate().map_err(ExportError::invalid)?;
    }
    let errors = source.get("resourceErrors").map(read_resource_errors_value).transpose().map_err(ExportError::invalid)?.unwrap_or_default();
    let mut origins = Vec::new();
    if let Some(value) = source.get("origin") {
        let mut origin = read_origin_value(value).map_err(ExportError::invalid)?;
        if origin.format == DocumentOfficeSourceOriginFormat::Odt && (!get(source, "styles").is_null() || !get(source, "theme").is_null() || !numbering.is_null()) {
            return Err(ExportError::invalid("ODT resources cannot be interpreted as native Word styles, theme, or numbering."));
        }
        for error in errors { if !origin.resource_errors.contains(&error) { origin.resource_errors.push(error); } }
        origins.push(origin);
    }
    if let Some(value) = source.get("legacyOrigins") {
        origins.extend(read_origin_list_value(value).map_err(ExportError::invalid)?);
    }
    if origins.is_empty() { return Ok(None); }
    encode_origins(origins).map(Some).map_err(|error| ExportError::invalid(format!("Office legacy source could not be retained: {error}")))
}

#[cfg(test)]
#[path = "docx_legacy_source_tests.rs"]
mod tests;
