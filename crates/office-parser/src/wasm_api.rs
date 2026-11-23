use crate::model::{ParseEnvelope, ParseFailure};
use crate::parser::{parse_document_bytes, parse_excel_bytes, parse_office_bytes};
use office_wasm_bridge::{
    BoundaryCopyStats, TaskControl, TaskKind, CONTROL_VERSION,
    copy_bytes_in, error_to_js, from_control, reject_linear_memory_input, to_control, yield_to_browser,
};
use serde::Serialize;
use wasm_bindgen::prelude::*;

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct ParseResult {
    context: TaskControl,
    envelope: ParseEnvelope,
    transfer: BoundaryCopyStats,
}

fn to_js_error(error: ParseFailure) -> JsValue {
    #[derive(Serialize)]
    struct Failure { ok: bool, error: crate::model::ParseError }
    to_control(&Failure { ok: false, error: error.error })
        .unwrap_or_else(|error| error)
}

type Parser = fn(&[u8], &str) -> Result<ParseEnvelope, ParseFailure>;

async fn parse(bytes: js_sys::Uint8Array, file_name: String, control: JsValue, parser: Parser) -> Result<JsValue, JsValue> {
    reject_linear_memory_input(&bytes.buffer())?;
    let context: TaskControl = from_control(control)?;
    context.validate(TaskKind::Parse).map_err(error_to_js)?;
    context.budget.check_input(bytes.length() as usize).map_err(error_to_js)?;
    context.budget.check_working(bytes.length() as usize).map_err(error_to_js)?;
    // This is a real browser/Worker macro task. The parser itself is synchronous:
    // its Worker is terminated on cancel/timeout, never reported as interruptible.
    yield_to_browser().await?;
    let mut transfer = BoundaryCopyStats { control_dto_conversions: 1, ..Default::default() };
    let input = copy_bytes_in(&bytes, context.budget.max_input_bytes, &mut transfer)?;
    let envelope = parser(&input, &file_name).map_err(to_js_error)?;
    drop(input);
    transfer.control_dto_conversions += 1;
    to_control(&ParseResult { context, envelope, transfer })
}

#[wasm_bindgen]
pub fn parser_protocol_version() -> u32 { CONTROL_VERSION }

#[wasm_bindgen]
pub async fn parse_office(bytes: js_sys::Uint8Array, file_name: String, control: JsValue) -> Result<JsValue, JsValue> {
    parse(bytes, file_name, control, parse_office_bytes).await
}

#[wasm_bindgen]
pub async fn parse_excel(bytes: js_sys::Uint8Array, file_name: String, control: JsValue) -> Result<JsValue, JsValue> {
    parse(bytes, file_name, control, parse_excel_bytes).await
}

#[wasm_bindgen]
pub async fn parse_document(bytes: js_sys::Uint8Array, file_name: String, control: JsValue) -> Result<JsValue, JsValue> {
    parse(bytes, file_name, control, parse_document_bytes).await
}
