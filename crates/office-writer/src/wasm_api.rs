use crate::docx;
use crate::error::{ExportError, ExportWarning};
use crate::opc::Package;
use crate::xlsx;
use wasm_bindgen::prelude::*;

#[derive(serde::Serialize)]
struct ErrorPayload<'a> {
    ok: bool,
    error: &'a ExportError,
}

fn run(kind: &str, data: &str) -> Result<(Vec<u8>, Vec<ExportWarning>), ExportError> {
    let mut package = Package::new();
    let warnings = match kind {
        "workbook" | "xlsx" => xlsx::write(data, &mut package)?,
        "document" | "docx" => docx::write(data, &mut package)?,
        other => {
            return Err(ExportError::unsupported(format!(
                "unsupported export kind: {other}"
            )));
        }
    };
    Ok((package.finish()?, warnings))
}

#[wasm_bindgen]
pub struct ExportResult {
    bytes: Vec<u8>,
    warnings: String,
}

#[wasm_bindgen]
impl ExportResult {
    #[wasm_bindgen(getter)]
    pub fn bytes(&self) -> Vec<u8> {
        self.bytes.clone()
    }

    #[wasm_bindgen(getter)]
    pub fn warnings(&self) -> String {
        self.warnings.clone()
    }
}

fn js_error(error: ExportError) -> JsValue {
    let payload = ErrorPayload {
        ok: false,
        error: &error,
    };
    serde_json::to_string(&payload).map_or_else(|_| JsValue::from_str("{\"ok\":false,\"error\":{\"code\":\"EXPORT_FAILURE\",\"message\":\"Unable to encode export error\"}}"), |value| JsValue::from_str(&value))
}

#[wasm_bindgen]
pub fn export_office(kind: &str, data: &str) -> Result<Vec<u8>, JsValue> {
    run(kind, data).map(|(bytes, _)| bytes).map_err(js_error)
}

#[wasm_bindgen]
pub fn export_office_report(kind: &str, data: &str) -> Result<String, JsValue> {
    let (bytes, warnings) = run(kind, data).map_err(js_error)?;
    let warnings = serde_json::to_string(&warnings)
        .map_err(|error| js_error(ExportError::package(error.to_string())))?;
    Ok(
        serde_json::json!({ "schemaVersion": 1, "size": bytes.len(), "warnings": warnings })
            .to_string(),
    )
}

#[wasm_bindgen]
pub fn write_office(kind: &str, data: &str) -> Result<ExportResult, JsValue> {
    let (bytes, warnings) = run(kind, data).map_err(js_error)?;
    let warnings = serde_json::to_string(&warnings)
        .map_err(|error| js_error(ExportError::package(error.to_string())))?;
    Ok(ExportResult { bytes, warnings })
}
