use crate::{BoundaryCopyStats, BridgeError, error_to_js, to_control};
use js_sys::{Float64Array, Object, Uint8Array, Uint32Array, WebAssembly};
use serde::Serialize;
use std::cell::{Cell, RefCell};
use wasm_bindgen::{JsCast, prelude::*};

thread_local! {
    static NEXT_OWNER: Cell<u32> = const { Cell::new(1) };
}

fn owner_id() -> u32 {
    NEXT_OWNER.with(|next| {
        let value = next.get();
        next.set(value.wrapping_add(1).max(1));
        value
    })
}

fn current_memory_buffer() -> JsValue {
    let memory: WebAssembly::Memory = wasm_bindgen::memory().unchecked_into();
    memory.buffer()
}

pub fn is_linear_memory_buffer(buffer: &JsValue) -> bool {
    Object::is(buffer, &current_memory_buffer())
}

pub fn reject_linear_memory_input(buffer: &JsValue) -> Result<(), JsValue> {
    // Allocating the destination can grow memory and detach the source view.
    // Internal kernels borrow Rust slices instead of re-entering this API.
    if is_linear_memory_buffer(buffer) {
        Err(error_to_js(BridgeError::new(
            "WASM_INPUT_ALIAS", "Input must be an independent buffer, not this WASM memory",
        )))
    } else { Ok(()) }
}

pub fn copy_bytes_in(input: &Uint8Array, maximum: u32, stats: &mut BoundaryCopyStats) -> Result<Vec<u8>, JsValue> {
    reject_linear_memory_input(&input.buffer())?;
    let length = input.length() as usize;
    if length > maximum as usize {
        return Err(error_to_js(BridgeError::new("WASM_MEMORY_LIMIT", "Input buffer exceeds its budget")));
    }
    let mut output = Vec::new();
    output.try_reserve_exact(length).map_err(|_| error_to_js(
        BridgeError::new("WASM_MEMORY_LIMIT", "Could not reserve the input buffer"),
    ))?;
    output.resize(length, 0);
    input.copy_to(&mut output);
    stats.input(length);
    Ok(output)
}

pub fn copy_numbers_in(input: &Float64Array, maximum: u32, stats: &mut BoundaryCopyStats) -> Result<Vec<f64>, JsValue> {
    reject_linear_memory_input(&input.buffer())?;
    let length = input.length() as usize;
    let bytes = length.checked_mul(8).ok_or_else(|| error_to_js(
        BridgeError::new("WASM_MEMORY_LIMIT", "Numeric buffer size overflow"),
    ))?;
    if bytes > maximum as usize {
        return Err(error_to_js(BridgeError::new("WASM_MEMORY_LIMIT", "Numeric input exceeds its budget")));
    }
    let mut output = Vec::new();
    output.try_reserve_exact(length).map_err(|_| error_to_js(
        BridgeError::new("WASM_MEMORY_LIMIT", "Could not reserve numeric input"),
    ))?;
    output.resize(length, 0.0);
    input.copy_to(&mut output);
    stats.input(bytes);
    Ok(output)
}

/// The independent JS destination is allocated before the temporary WASM
/// view. The view never escapes this function and no allocating WASM call,
/// arbitrary callback or await occurs while it exists.
pub fn copy_bytes_out(bytes: &[u8], stats: &mut BoundaryCopyStats) -> Result<Uint8Array, JsValue> {
    let length = u32::try_from(bytes.len()).map_err(|_| error_to_js(
        BridgeError::new("WASM_MEMORY_LIMIT", "Output does not fit a TypedArray"),
    ))?;
    let output = Uint8Array::new_with_length(length);
    // SAFETY: bytes remain borrowed and stable until set completes. All Rust
    // allocation and the destination allocation happened before acquiring view.
    let view = unsafe { Uint8Array::view(bytes) };
    output.set(&view, 0);
    drop(view);
    stats.output(bytes.len());
    Ok(output)
}

pub fn copy_indices_out(indices: &[u32], stats: &mut BoundaryCopyStats) -> Result<Uint32Array, JsValue> {
    let length = u32::try_from(indices.len()).map_err(|_| error_to_js(
        BridgeError::new("WASM_MEMORY_LIMIT", "Index output does not fit a TypedArray"),
    ))?;
    let output = Uint32Array::new_with_length(length);
    // SAFETY: same synchronous, immutable lease as copy_bytes_out.
    let view = unsafe { Uint32Array::view(indices) };
    output.set(&view, 0);
    drop(view);
    stats.output(indices.len() * 4);
    Ok(output)
}

#[wasm_bindgen]
pub struct ByteLeaseBuffer {
    bytes: Option<Vec<u8>>,
    owner: u32,
    generation: u32,
    leased: bool,
    lease_epoch: u32,
    memory_epoch: Cell<u32>,
    memory_buffer: RefCell<JsValue>,
    stats: BoundaryCopyStats,
}

impl ByteLeaseBuffer {
    pub fn from_vec(bytes: Vec<u8>, stats: BoundaryCopyStats) -> Self {
        Self { bytes: Some(bytes), owner: owner_id(), generation: 1, leased: false,
            lease_epoch: 0, memory_epoch: Cell::new(0),
            memory_buffer: RefCell::new(current_memory_buffer()), stats }
    }

    fn observe_memory_epoch(&self) -> u32 {
        let current = current_memory_buffer();
        if !Object::is(&current, &self.memory_buffer.borrow()) {
            *self.memory_buffer.borrow_mut() = current;
            self.memory_epoch.set(self.memory_epoch.get().wrapping_add(1));
        }
        self.memory_epoch.get()
    }

    pub fn as_slice(&self) -> Result<&[u8], JsValue> {
        self.bytes.as_deref().ok_or_else(|| error_to_js(
            BridgeError::new("WASM_BUFFER_DISPOSED", "Buffer owner was disposed"),
        ))
    }

    pub fn into_vec(mut self) -> Result<Vec<u8>, JsValue> {
        if self.leased {
            return Err(error_to_js(BridgeError::new("WASM_LEASE_ACTIVE", "Cannot consume an active buffer lease")));
        }
        self.bytes.take().ok_or_else(|| error_to_js(
            BridgeError::new("WASM_BUFFER_DISPOSED", "Buffer owner was disposed"),
        ))
    }
}

#[wasm_bindgen]
impl ByteLeaseBuffer {
    #[wasm_bindgen(constructor)]
    pub fn new(input: Uint8Array, maximum_bytes: u32) -> Result<ByteLeaseBuffer, JsValue> {
        let mut stats = BoundaryCopyStats::default();
        let bytes = copy_bytes_in(&input, maximum_bytes, &mut stats)?;
        Ok(Self::from_vec(bytes, stats))
    }

    #[wasm_bindgen(getter)]
    pub fn byte_length(&self) -> Result<u32, JsValue> {
        Ok(self.as_slice()?.len() as u32)
    }

    pub fn begin_lease(&mut self) -> Result<u32, JsValue> {
        self.as_slice()?;
        if self.leased {
            return Err(error_to_js(BridgeError::new("WASM_LEASE_ACTIVE", "Buffer already has an active lease")));
        }
        self.lease_epoch = self.observe_memory_epoch();
        self.leased = true;
        Ok(self.generation)
    }

    /// Low-level synchronous lease. Keep this owner alive, never free/consume it
    /// or re-enter allocating WASM while reading. After ANY memory.grow obtain
    /// a fresh view. Public callers should prefer copy_bytes instead.
    pub fn lease_view(&self, generation: u32) -> Result<Uint8Array, JsValue> {
        if !self.leased || generation != self.generation {
            return Err(error_to_js(BridgeError::new("WASM_LEASE_STALE", "Invalid buffer lease generation")));
        }
        if self.observe_memory_epoch() != self.lease_epoch {
            return Err(error_to_js(BridgeError::new(
                "WASM_MEMORY_CHANGED", "Memory changed during the lease; end it and acquire a new lease",
            )));
        }
        // SAFETY: storage is immutable while leased. The caller is responsible
        // for synchronous lifetime and memory-epoch rules, as documented above.
        Ok(unsafe { Uint8Array::view(self.as_slice()?) })
    }

    pub fn end_lease(&mut self, generation: u32) -> Result<(), JsValue> {
        if !self.leased || generation != self.generation {
            return Err(error_to_js(BridgeError::new("WASM_LEASE_STALE", "Invalid buffer lease generation")));
        }
        self.leased = false;
        self.generation = self.generation.wrapping_add(1).max(1);
        Ok(())
    }

    pub fn copy_bytes(&mut self) -> Result<Uint8Array, JsValue> {
        if self.leased {
            return Err(error_to_js(BridgeError::new("WASM_LEASE_ACTIVE", "Release the low-level lease before copying")));
        }
        let bytes = self.bytes.as_deref().ok_or_else(|| error_to_js(
            BridgeError::new("WASM_BUFFER_DISPOSED", "Buffer owner was disposed"),
        ))?;
        copy_bytes_out(bytes, &mut self.stats)
    }

    pub fn metadata(&self) -> Result<JsValue, JsValue> {
        if self.leased {
            return Err(error_to_js(BridgeError::new(
                "WASM_LEASE_ACTIVE", "Read allocating metadata before beginning a lease",
            )));
        }
        #[derive(Serialize)]
        #[serde(rename_all = "camelCase")]
        struct Metadata<'a> {
            owner_id: u32, generation: u32, byte_length: usize,
            element_type: &'static str, byte_offset: usize, length: usize,
            memory_epoch: u32, access: &'static str, transfer: &'a BoundaryCopyStats,
        }
        let bytes = self.as_slice()?;
        to_control(&Metadata {
            owner_id: self.owner, generation: self.generation,
            byte_length: bytes.len(), element_type: "u8", byte_offset: bytes.as_ptr() as usize,
            length: bytes.len(), memory_epoch: self.observe_memory_epoch(),
            access: "read", transfer: &self.stats,
        })
    }

    pub fn dispose(&mut self) -> Result<(), JsValue> {
        if self.leased {
            return Err(error_to_js(BridgeError::new("WASM_LEASE_ACTIVE", "Cannot dispose an active buffer lease")));
        }
        self.as_slice()?;
        self.bytes = None;
        self.generation = self.generation.wrapping_add(1).max(1);
        Ok(())
    }
}
