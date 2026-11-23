use crate::{BridgeError, error_to_js};
use js_sys::{Function, Promise};
use wasm_bindgen::{JsCast, prelude::*};
use wasm_bindgen_futures::JsFuture;

#[derive(Clone)]
enum TimerTarget {
    Window(web_sys::Window),
    Worker(web_sys::WorkerGlobalScope),
}

impl TimerTarget {
    fn current() -> Result<Self, JsValue> {
        if let Some(window) = web_sys::window() {
            return Ok(Self::Window(window));
        }
        js_sys::global().dyn_into::<web_sys::WorkerGlobalScope>()
            .map(Self::Worker)
            .map_err(|_| error_to_js(BridgeError::new(
                "WASM_PLATFORM_UNAVAILABLE", "This task requires a browser or Worker timer",
            )))
    }

    fn schedule(&self, callback: &Function) -> Result<i32, JsValue> {
        match self {
            Self::Window(target) => target.set_timeout_with_callback_and_timeout_and_arguments_0(callback, 0),
            Self::Worker(target) => target.set_timeout_with_callback_and_timeout_and_arguments_0(callback, 0),
        }
    }

    fn clear(&self, handle: i32) {
        match self {
            Self::Window(target) => target.clear_timeout_with_handle(handle),
            Self::Worker(target) => target.clear_timeout_with_handle(handle),
        }
    }
}

struct ScheduledTimer {
    target: TimerTarget,
    handle: i32,
    // The callback remains valid until the future completes or is dropped.
    _callback: Closure<dyn FnMut()>,
}

impl Drop for ScheduledTimer {
    fn drop(&mut self) {
        self.target.clear(self.handle);
    }
}

/// A macro task, allowing ordinary Worker cancel messages to run. This does
/// not interrupt a synchronous parser/writer call; those are cancelled by
/// terminating their Worker. No Closure::forget or permanently owned timer.
pub async fn yield_to_browser() -> Result<(), JsValue> {
    let target = TimerTarget::current()?;
    let mut scheduled = None;
    let promise = Promise::new(&mut |resolve, reject| {
        let callback = Closure::new(move || {
            let _ = resolve.call0(&JsValue::UNDEFINED);
        });
        match target.schedule(callback.as_ref().unchecked_ref()) {
            Ok(handle) => {
                scheduled = Some(ScheduledTimer { target: target.clone(), handle, _callback: callback });
            }
            Err(error) => { let _ = reject.call1(&JsValue::UNDEFINED, &error); }
        }
    });
    let outcome = JsFuture::from(promise).await;
    drop(scheduled);
    outcome.map(|_| ())
}

#[wasm_bindgen]
pub async fn office_wasm_yield() -> Result<(), JsValue> {
    yield_to_browser().await
}
