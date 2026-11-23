use serde::{Deserialize, Serialize};
use std::fmt;

pub const CONTROL_VERSION: u32 = 2;
pub const JS_SAFE_INTEGER: u64 = 9_007_199_254_740_991;

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct TaskBudget {
    pub max_input_bytes: u32,
    pub max_output_bytes: u32,
    pub max_working_bytes: u32,
}

impl Default for TaskBudget {
    fn default() -> Self {
        Self {
            max_input_bytes: 256 * 1024 * 1024,
            max_output_bytes: 256 * 1024 * 1024,
            max_working_bytes: 768 * 1024 * 1024,
        }
    }
}

impl TaskBudget {
    pub fn check_input(&self, bytes: usize) -> Result<(), BridgeError> {
        self.check(bytes, self.max_input_bytes, "Input")
    }

    pub fn check_output(&self, bytes: usize) -> Result<(), BridgeError> {
        self.check(bytes, self.max_output_bytes, "Output")
    }

    pub fn check_working(&self, bytes: usize) -> Result<(), BridgeError> {
        self.check(bytes, self.max_working_bytes, "Retained working buffers")
    }

    fn check(&self, bytes: usize, limit: u32, label: &str) -> Result<(), BridgeError> {
        if bytes > limit as usize {
            Err(BridgeError::new(
                "WASM_MEMORY_LIMIT",
                format!("{label} exceed the configured boundary memory budget"),
            ))
        } else {
            Ok(())
        }
    }
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "camelCase")]
pub enum TaskKind {
    Parse,
    Export,
    Aggregate,
    Sort,
    Filter,
    ChartStatistics,
    Layout,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct TaskControl {
    pub protocol_version: u32,
    pub request_id: String,
    pub unit_id: String,
    pub base_revision: u64,
    pub kind: TaskKind,
    #[serde(default)]
    pub budget: TaskBudget,
}

impl TaskControl {
    pub fn validate(&self, expected: TaskKind) -> Result<(), BridgeError> {
        if self.protocol_version != CONTROL_VERSION {
            return Err(BridgeError::new("WASM_PROTOCOL_MISMATCH", "Unsupported WASM control protocol"));
        }
        if self.kind != expected {
            return Err(BridgeError::new("WASM_CONTROL_INVALID", "Task kind does not match the entry point"));
        }
        if self.request_id.trim().is_empty() || self.unit_id.trim().is_empty() {
            return Err(BridgeError::new("WASM_CONTROL_INVALID", "Request and unit IDs must be nonempty"));
        }
        if self.base_revision > JS_SAFE_INTEGER {
            return Err(BridgeError::new("WASM_CONTROL_INVALID", "Revision exceeds the JavaScript safe integer range"));
        }
        if self.budget.max_working_bytes == 0 {
            return Err(BridgeError::new("WASM_MEMORY_LIMIT", "Working memory budget must be positive"));
        }
        Ok(())
    }
}

/// Only measured JS/WASM boundary copies. Parser internals, object cloning,
/// transfer, GPU upload and browser heap peaks are separate measurements.
#[derive(Debug, Clone, Default, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct BoundaryCopyStats {
    pub input_copies: u32,
    pub input_bytes: u64,
    pub output_copies: u32,
    pub output_bytes: u64,
    pub rust_buffer_clones: u32,
    pub control_dto_conversions: u32,
    pub json_conversions: u32,
}

impl BoundaryCopyStats {
    pub fn input(&mut self, bytes: usize) {
        if bytes > 0 {
            self.input_copies += 1;
            self.input_bytes += bytes as u64;
        }
    }

    pub fn output(&mut self, bytes: usize) {
        if bytes > 0 {
            self.output_copies += 1;
            self.output_bytes += bytes as u64;
        }
    }
}

#[derive(Debug, Clone, Serialize, PartialEq, Eq)]
#[serde(rename_all = "camelCase")]
pub struct BridgeError {
    pub code: String,
    pub message: String,
    pub details: Option<String>,
}

impl BridgeError {
    pub fn new(code: impl Into<String>, message: impl Into<String>) -> Self {
        Self { code: code.into(), message: message.into(), details: None }
    }

    pub fn detail(mut self, detail: impl Into<String>) -> Self {
        self.details = Some(detail.into());
        self
    }
}

impl fmt::Display for BridgeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(formatter, "{}: {}", self.code, self.message)
    }
}

impl std::error::Error for BridgeError {}

#[cfg(test)]
mod tests {
    use super::*;

    fn context() -> TaskControl {
        TaskControl {
            protocol_version: CONTROL_VERSION, request_id: "request".into(),
            unit_id: "workbook".into(), base_revision: 7, kind: TaskKind::Sort,
            budget: TaskBudget::default(),
        }
    }

    #[test]
    fn rejects_wrong_task_revision_and_protocol_before_mutation() {
        let mut control = context();
        assert!(control.validate(TaskKind::Filter).is_err());
        control.base_revision = JS_SAFE_INTEGER + 1;
        assert!(control.validate(TaskKind::Sort).is_err());
        control.base_revision = 7;
        control.protocol_version = CONTROL_VERSION + 1;
        assert!(control.validate(TaskKind::Sort).is_err());
    }

    #[test]
    fn control_rejects_unknown_fields_and_budget_checks_exact_boundary() {
        let mut value = serde_json::to_value(context()).unwrap();
        value["unexpected"] = serde_json::json!(true);
        assert!(serde_json::from_value::<TaskControl>(value).is_err());
        let budget = TaskBudget { max_input_bytes: 4, max_output_bytes: 8, max_working_bytes: 12 };
        assert!(budget.check_input(4).is_ok());
        assert!(budget.check_input(5).is_err());
        assert!(budget.check_output(9).is_err());
        assert!(budget.check_working(13).is_err());
    }
}
