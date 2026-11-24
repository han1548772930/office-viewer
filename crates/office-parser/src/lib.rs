mod model;
mod parser;
mod sheet_dimensions;
mod sheet_dimension_sidecar;
mod sheet_filter_import;
mod xlsx_parts;
#[cfg(test)]
mod sheet_dimensions_roundtrip_tests;
#[cfg(test)]
mod sheet_filter_roundtrip_tests;
mod document_format_sidecar;
#[cfg(test)]
mod document_format_roundtrip_tests;
#[cfg(test)]
mod document_story_roundtrip_tests;
mod story_fields;
mod document_text_fields;
mod document_block_extraction;
mod document_story_source;
mod document_paragraph_source;
mod document_table_sidecar;
mod document_table_merges;
#[cfg(test)]
mod document_table_roundtrip_tests;
#[cfg(test)]
mod document_numbering_roundtrip_tests;
#[cfg(test)]
mod document_text_fields_roundtrip_tests;
mod validation_sidecar;
#[cfg(test)]
mod sheet_validation_roundtrip_tests;
mod sparkline_sidecar;
mod pivot_sidecar;
mod native_chart;
#[cfg(test)]
mod native_chart_roundtrip_tests;
mod conditional_format_sidecar;
mod table_sidecar;
#[cfg(target_arch = "wasm32")]
mod wasm_api;

pub use model::{
    DocumentModel, ParseEnvelope, ParseError, ParseErrorCode, ParseStats, ParserDiagnostics,
    SCHEMA_VERSION,
};
#[cfg(target_arch = "wasm32")]
pub use wasm_api::{parse_document, parse_excel, parse_office};
