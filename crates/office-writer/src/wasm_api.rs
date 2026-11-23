use crate::{docx, xlsx, opc::Package, error::{ExportError, ExportWarning}};
use office_wasm_bridge::{
    BoundaryCopyStats, TaskControl, TaskKind, CONTROL_VERSION,
    copy_bytes_out, error_to_js, from_control, to_control, yield_to_browser,
};
use serde::Serialize;
use wasm_bindgen::prelude::*;

fn run(kind: &str, data: &str) -> Result<(Vec<u8>, Vec<ExportWarning>), ExportError> {
    let mut package = Package::new();
    let warnings = match kind {
        "workbook" => xlsx::write(data, &mut package)?,
        "document" => docx::write(data, &mut package)?,
        other => return Err(ExportError::unsupported(format!("unsupported export kind: {other}"))),
    };
    Ok((package.finish()?, warnings))
}

fn js_error(error: ExportError) -> JsValue {
    #[derive(Serialize)]
    struct Failure { ok: bool, error: ExportError }
    to_control(&Failure { ok: false, error }).unwrap_or_else(|error| error)
}

#[wasm_bindgen]
pub struct ExportResult {
    bytes: Vec<u8>,
    warnings: Vec<ExportWarning>,
    context: TaskControl,
    transfer: BoundaryCopyStats,
}

#[wasm_bindgen]
impl ExportResult {
    /// Safe default: one copy of the valid result region into an independent
    /// transferable buffer. The retained Rust Vec is never cloned or transferred.
    pub fn copy_bytes(&mut self) -> Result<js_sys::Uint8Array, JsValue> {
        self.context.budget.check_output(self.bytes.len()).map_err(error_to_js)?;
        copy_bytes_out(&self.bytes, &mut self.transfer)
    }

    pub fn metadata(&mut self) -> Result<JsValue, JsValue> {
        #[derive(Serialize)]
        #[serde(rename_all = "camelCase")]
        struct Metadata<'a> {
            context: &'a TaskControl,
            byte_length: usize,
            warnings: &'a [ExportWarning],
            transfer: &'a BoundaryCopyStats,
        }
        self.transfer.control_dto_conversions += 1;
        to_control(&Metadata {
            context: &self.context, byte_length: self.bytes.len(),
            warnings: &self.warnings, transfer: &self.transfer,
        })
    }
}

#[wasm_bindgen]
pub fn writer_protocol_version() -> u32 { CONTROL_VERSION }

#[wasm_bindgen]
pub async fn write_office(kind: String, data: String, control: JsValue) -> Result<ExportResult, JsValue> {
    let context: TaskControl = from_control(control)?;
    context.validate(TaskKind::Export).map_err(error_to_js)?;
    context.budget.check_input(data.len()).map_err(error_to_js)?;
    context.budget.check_working(data.len()).map_err(error_to_js)?;
    yield_to_browser().await?;
    let mut transfer = BoundaryCopyStats {
        control_dto_conversions: 1,
        // The persisted Office snapshot is still a JSON input to the writer.
        // This and its wasm-bindgen UTF-8 String copy are explicitly counted.
        json_conversions: 1,
        ..Default::default()
    };
    transfer.input(data.len());
    let (bytes, warnings) = run(&kind, &data).map_err(js_error)?;
    context.budget.check_output(bytes.len()).map_err(error_to_js)?;
    context.budget.check_working(bytes.len().saturating_add(data.len())).map_err(error_to_js)?;
    Ok(ExportResult { bytes, warnings, context, transfer })
}
