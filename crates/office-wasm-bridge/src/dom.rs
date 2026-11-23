use crate::{BridgeError, error_to_js, from_control, to_control};
use serde::{Deserialize, Serialize};
use std::{cell::{Cell, RefCell}, rc::Rc};
use wasm_bindgen::{JsCast, prelude::*};
use web_sys::{Event, EventTarget, HtmlElement, HtmlInputElement, HtmlTextAreaElement};

#[derive(Clone, PartialEq, Eq, Deserialize, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct InputHostContext {
    pub unit_id: String,
    pub generation: u32,
    pub model_revision: u64,
}

impl InputHostContext {
    fn validate(&self) -> Result<(), JsValue> {
        if self.unit_id.trim().is_empty() || self.model_revision > crate::JS_SAFE_INTEGER {
            Err(error_to_js(BridgeError::new("WASM_INPUT_CONTEXT_INVALID", "Invalid editor host context")))
        } else { Ok(()) }
    }
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct InputSelection {
    context: InputHostContext,
    start: u32,
    end: u32,
    direction: String,
}

enum TextInput { Input(HtmlInputElement), TextArea(HtmlTextAreaElement) }

impl TextInput {
    fn value(&self) -> String {
        match self { Self::Input(input) => input.value(), Self::TextArea(input) => input.value() }
    }
    fn target(&self) -> &EventTarget {
        match self { Self::Input(input) => input.as_ref(), Self::TextArea(input) => input.as_ref() }
    }
    fn element(&self) -> &HtmlElement {
        match self { Self::Input(input) => input.as_ref(), Self::TextArea(input) => input.as_ref() }
    }
    fn select(&self, start: u32, end: u32, direction: &str) -> Result<(), JsValue> {
        match self {
            Self::Input(input) => input.set_selection_range_with_direction(start, end, direction),
            Self::TextArea(input) => input.set_selection_range_with_direction(start, end, direction),
        }
    }
}

struct Listener {
    target: EventTarget,
    name: &'static str,
    callback: Closure<dyn FnMut(Event)>,
    attached: bool,
}

impl Listener {
    fn new(target: &EventTarget, name: &'static str, composing: Rc<Cell<bool>>, value: bool) -> Result<Self, JsValue> {
        let callback = Closure::new(move |_: Event| composing.set(value));
        target.add_event_listener_with_callback(name, callback.as_ref().unchecked_ref())?;
        Ok(Self { target: target.clone(), name, callback, attached: true })
    }
    fn detach(&mut self) -> Result<(), JsValue> {
        if self.attached {
            self.target.remove_event_listener_with_callback(self.name, self.callback.as_ref().unchecked_ref())?;
            self.attached = false;
        }
        Ok(())
    }
}

impl Drop for Listener {
    fn drop(&mut self) { let _ = self.detach(); }
}

#[wasm_bindgen]
pub struct EditorInputHost {
    input: TextInput,
    context: RefCell<InputHostContext>,
    composing: Rc<Cell<bool>>,
    disposed: Cell<bool>,
    listeners: Vec<Listener>,
}

#[wasm_bindgen]
impl EditorInputHost {
    #[wasm_bindgen(constructor)]
    pub fn new(element: JsValue, context: JsValue) -> Result<EditorInputHost, JsValue> {
        if web_sys::window().is_none() {
            return Err(error_to_js(BridgeError::new("WASM_DOM_UNAVAILABLE", "DOM hosts require the main browser thread")));
        }
        let context: InputHostContext = from_control(context)?;
        context.validate()?;
        let input = if let Some(input) = element.dyn_ref::<HtmlInputElement>() {
            TextInput::Input(input.clone())
        } else if let Some(input) = element.dyn_ref::<HtmlTextAreaElement>() {
            TextInput::TextArea(input.clone())
        } else {
            return Err(error_to_js(BridgeError::new("WASM_INPUT_HOST_INVALID", "Expected an existing Input or TextArea element")));
        };
        let composing = Rc::new(Cell::new(false));
        let listeners = vec![
            Listener::new(input.target(), "compositionstart", composing.clone(), true)?,
            Listener::new(input.target(), "compositionend", composing.clone(), false)?,
        ];
        Ok(Self { input, context: RefCell::new(context), composing, disposed: Cell::new(false), listeners })
    }

    pub fn update_context(&self, context: JsValue) -> Result<(), JsValue> {
        if self.disposed.get() {
            return Err(error_to_js(BridgeError::new("WASM_DOM_DISPOSED", "Input host was disposed")));
        }
        let next: InputHostContext = from_control(context)?;
        next.validate()?;
        *self.context.borrow_mut() = next;
        Ok(())
    }

    /// Restores a logical session range, without making DOM value authoritative.
    /// RefCell borrows end before focus because focus can synchronously re-enter.
    pub fn focus_and_select(&self, selection: JsValue) -> Result<bool, JsValue> {
        let selection: InputSelection = from_control(selection)?;
        selection.context.validate()?;
        if self.disposed.get() || self.composing.get() || *self.context.borrow() != selection.context {
            return Ok(false);
        }
        if selection.start > selection.end
            || selection.end as usize > self.input.value().encode_utf16().count()
            || !matches!(selection.direction.as_str(), "forward" | "backward" | "none")
        {
            return Err(error_to_js(BridgeError::new("WASM_INPUT_RANGE_INVALID", "Invalid UTF-16 input selection")));
        }
        let options = web_sys::FocusOptions::new();
        options.set_prevent_scroll(true);
        self.input.element().focus_with_options(&options)?;
        if self.disposed.get() || self.composing.get() || *self.context.borrow() != selection.context {
            return Ok(false);
        }
        self.input.select(selection.start, selection.end, &selection.direction)?;
        Ok(true)
    }

    pub fn status(&self) -> Result<JsValue, JsValue> {
        #[derive(Serialize)]
        struct Status { composing: bool, disposed: bool }
        to_control(&Status { composing: self.composing.get(), disposed: self.disposed.get() })
    }

    pub fn dispose(&mut self) -> Result<(), JsValue> {
        self.disposed.set(true);
        for listener in &mut self.listeners { listener.detach()?; }
        self.listeners.clear();
        Ok(())
    }
}

#[wasm_bindgen]
pub struct CanvasDomHost { canvas: Option<web_sys::HtmlCanvasElement> }

#[wasm_bindgen]
impl CanvasDomHost {
    #[wasm_bindgen(constructor)]
    pub fn new(canvas: web_sys::HtmlCanvasElement) -> Result<CanvasDomHost, JsValue> {
        if web_sys::window().is_none() {
            return Err(error_to_js(BridgeError::new("WASM_DOM_UNAVAILABLE", "Canvas DOM hosts require the main browser thread")));
        }
        Ok(Self { canvas: Some(canvas) })
    }

    /// The renderer still owns the backing store, GPU context and DPR. CSS
    /// dimensions are independent of page geometry and editor content.
    pub fn set_css_size(&self, width: f64, height: f64) -> Result<(), JsValue> {
        if !width.is_finite() || !height.is_finite() || width < 0.0 || height < 0.0 {
            return Err(error_to_js(BridgeError::new("WASM_CANVAS_SIZE_INVALID", "Canvas CSS size must be finite and nonnegative")));
        }
        let canvas = self.canvas.as_ref().ok_or_else(|| error_to_js(
            BridgeError::new("WASM_DOM_DISPOSED", "Canvas host was disposed"),
        ))?;
        canvas.style().set_property("width", &format!("{width}px"))?;
        canvas.style().set_property("height", &format!("{height}px"))?;
        Ok(())
    }

    pub fn dispose(&mut self) { self.canvas = None; }
}
