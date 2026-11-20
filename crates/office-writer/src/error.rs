use serde::Serialize;
use std::fmt::{Display, Formatter};

#[derive(Debug, Clone, Serialize)]
pub struct ExportWarning {
    pub code: &'static str,
    pub message: String,
}

impl ExportWarning {
    pub fn new(code: &'static str, message: impl Into<String>) -> Self {
        Self {
            code,
            message: message.into(),
        }
    }
}

#[derive(Debug, Clone, Serialize)]
pub struct ExportError {
    pub code: &'static str,
    pub message: String,
}

impl ExportError {
    pub fn invalid(detail: impl Into<String>) -> Self {
        Self {
            code: "INVALID_MODEL",
            message: detail.into(),
        }
    }

    pub fn unsupported(detail: impl Into<String>) -> Self {
        Self {
            code: "UNSUPPORTED_EXPORT",
            message: detail.into(),
        }
    }

    pub fn package(detail: impl Into<String>) -> Self {
        Self {
            code: "PACKAGE_ERROR",
            message: detail.into(),
        }
    }
}

impl Display for ExportError {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        write!(formatter, "{}: {}", self.code, self.message)
    }
}

impl std::error::Error for ExportError {}
