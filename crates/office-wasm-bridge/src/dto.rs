use crate::BridgeError;
use serde::{Serialize, de::DeserializeOwned};
use wasm_bindgen::prelude::*;

pub fn to_control<T: Serialize + ?Sized>(value: &T) -> Result<JsValue, JsValue> {
    // Match the existing JSON contract for objects and absent Option fields,
    // retaining serde's rejection of unsafe integer conversion.
    value.serialize(&serde_wasm_bindgen::Serializer::new()
        .serialize_maps_as_objects(true)
        .serialize_missing_as_null(true))
        .map_err(|error| error_to_js(BridgeError::new(
            "WASM_CONTROL_SERIALIZATION", "Could not encode a WASM control DTO",
        ).detail(error.to_string())))
}

pub fn from_control<T: DeserializeOwned>(value: JsValue) -> Result<T, JsValue> {
    serde_wasm_bindgen::from_value(value).map_err(|error| error_to_js(
        BridgeError::new("WASM_CONTROL_INVALID", "Invalid WASM control DTO")
            .detail(error.to_string()),
    ))
}

pub fn error_to_js(error: BridgeError) -> JsValue {
    #[derive(Serialize)]
    struct Failure<'a> { ok: bool, error: &'a BridgeError }
    Failure { ok: false, error: &error }
        .serialize(&serde_wasm_bindgen::Serializer::new()
            .serialize_maps_as_objects(true)
            .serialize_missing_as_null(true))
        .unwrap_or_else(|_| js_sys::Error::new(&error.to_string()).into())
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct Capabilities {
    protocol_version: u32,
    structured_control: bool,
    retained_buffer_leases: bool,
    macro_task_yield: bool,
    dom_hosts: bool,
}

#[wasm_bindgen]
pub fn office_wasm_bridge_capabilities() -> Result<JsValue, JsValue> {
    to_control(&Capabilities {
        protocol_version: crate::CONTROL_VERSION, structured_control: true,
        retained_buffer_leases: true, macro_task_yield: true,
        dom_hosts: web_sys::window().is_some(),
    })
}
