mod docx;
mod error;
mod opc;
mod wasm_api;
mod xlsx;
pub(crate) mod xml;

pub use error::{ExportError, ExportWarning};
pub use wasm_api::export_office;

/// Export a workbook model to a native XLSX package and return export diagnostics.
pub fn export_xlsx_bytes(data: &str) -> Result<(Vec<u8>, Vec<ExportWarning>), ExportError> {
    let mut package = opc::Package::new();
    let warnings = xlsx::write(data, &mut package)?;
    Ok((package.finish()?, warnings))
}

/// Export a document model to a native DOCX package and return export diagnostics.
pub fn export_docx_bytes(data: &str) -> Result<(Vec<u8>, Vec<ExportWarning>), ExportError> {
    let mut package = opc::Package::new();
    let warnings = docx::write(data, &mut package)?;
    Ok((package.finish()?, warnings))
}
