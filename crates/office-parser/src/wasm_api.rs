use crate::model::{ParseError, ParseFailure};
use crate::parser::{parse_document_bytes, parse_excel_bytes, parse_office_bytes};
use serde::Serialize;
use wasm_bindgen::prelude::*;

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
struct WasmError<'a> {
    ok: bool,
    error: &'a ParseError,
}

fn encode(result: Result<crate::model::ParseEnvelope, ParseFailure>) -> Result<String, JsValue> {
    match result {
        Ok(envelope) => serde_json::to_string(&envelope)
            .map_err(|error| to_js_error(ParseFailure::serialization(error.to_string()))),
        Err(error) => Err(to_js_error(error)),
    }
}

fn to_js_error(error: ParseFailure) -> JsValue {
    let payload = WasmError {
        ok: false,
        error: &error.error,
    };
    match serde_json::to_string(&payload) {
        Ok(value) => JsValue::from_str(&value),
        Err(_) => JsValue::from_str(
            r#"{"ok":false,"error":{"code":"PARSER_FAILURE","message":"Unable to encode parser error","details":null}}"#,
        ),
    }
}

#[wasm_bindgen]
pub fn parse_office(bytes: &[u8], file_name: &str) -> Result<String, JsValue> {
    encode(parse_office_bytes(bytes, file_name))
}

#[wasm_bindgen]
pub fn parse_excel(bytes: &[u8], file_name: &str) -> Result<String, JsValue> {
    encode(parse_excel_bytes(bytes, file_name))
}

#[wasm_bindgen]
pub fn parse_document(bytes: &[u8], file_name: &str) -> Result<String, JsValue> {
    encode(parse_document_bytes(bytes, file_name))
}
