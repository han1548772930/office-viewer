// Drafts only: not compiled or executed in the 2026-11-22 continuation.
use super::*;
use serde_json::json;

fn doc(offset: u32, hex: &str) -> DocumentOfficeSourceOrigin {
    DocumentOfficeSourceOrigin {
        format: DocumentOfficeSourceOriginFormat::Doc, synthesized: true,
        resources: vec![DocumentOfficeSourceResource::Binary { stream_name: "1Table".into(), offset, data_hex: hex.into() }],
        resource_errors: Vec::new(),
    }
}

fn odt(xml: &str) -> DocumentOfficeSourceOrigin {
    DocumentOfficeSourceOrigin {
        format: DocumentOfficeSourceOriginFormat::Odt, synthesized: false,
        resources: vec![DocumentOfficeSourceResource::Xml { part_name: "styles.xml".into(), xml: xml.into() }],
        resource_errors: Vec::new(),
    }
}

#[test]
fn source_codec_preserves_literal_xml_binary_text_order_and_historical_diagnostics() {
    let mut first = doc(32, "00AbFF");
    first.resource_errors = vec!["An original record is unresolved.".into()];
    let mut duplicate = doc(32, "00abff");
    duplicate.resource_errors = vec!["An original record is unresolved.".into(), "A second diagnostic.".into()];
    let xml = "<?xml version=\"1.0\"?><office:styles xmlns:office=\"odf\">\n<!-- 原始 XML 😀 -->\n</office:styles>";
    let bytes = encode_origins(vec![first.clone(), odt(xml), duplicate]).unwrap();
    let decoded = decode_origins(&bytes).unwrap();
    assert_eq!(decoded.len(), 2);
    assert_eq!(decoded[0].resources, first.resources);
    assert_eq!(decoded[0].resource_errors, ["An original record is unresolved.", "A second diagnostic."]);
    assert_eq!(decoded[1], odt(xml));
    assert_eq!(encode_origins(decoded.clone()).unwrap(), bytes);
    assert_eq!(decode_origins(&encode_origins(decoded).unwrap()).unwrap().len(), 2);
}

#[test]
fn changing_resource_order_bytes_or_binary_offsets_keeps_distinct_origins() {
    let first = doc(0, "00ff");
    let changed_bytes = doc(0, "00fe");
    let changed_offset = doc(1, "00ff");
    let mut ordered = first.clone();
    ordered.resources.push(DocumentOfficeSourceResource::Binary { stream_name: "0Table".into(), offset: 5, data_hex: "02".into() });
    let mut reversed = ordered.clone(); reversed.resources.reverse();
    assert_eq!(merge_origins(vec![first, changed_bytes, changed_offset, ordered, reversed]).unwrap().len(), 5);
}

#[test]
fn source_json_rejects_unknown_fields_versions_and_duplicate_json_keys() {
    for bytes in [
        br#"{"version":2,"legacyOrigins":[]}"#.as_slice(),
        br#"{"version":1,"legacyOrigins":[],"unretainedExtension":{"x":1}}"#.as_slice(),
        br#"{"version":1,"version":1,"legacyOrigins":[]}"#.as_slice(),
        br#"{"version":1,"legacyOrigins":[{"format":"doc","synthesized":true,"resources":[],"extra":1}]}"#.as_slice(),
        br#"{"version":1,"legacyOrigins":[{"format":"odt","synthesized":false,"resources":[{"kind":"xml","partName":"styles.xml","xml":"<styles/>","extra":1}]}]}"#.as_slice(),
    ] { assert!(decode_origins(bytes).is_err()); }
}

#[test]
fn source_contract_rejects_wrong_kind_marker_hex_offsets_and_utf8_name_budget() {
    for value in [
        json!({"format":"doc","synthesized":false,"resources":[]}),
        json!({"format":"doc","synthesized":true,"resources":[{"kind":"xml","partName":"styles.xml","xml":"x"}]}),
        json!({"format":"doc","synthesized":true,"resources":[{"kind":"binary","streamName":"1Table","offset":0,"dataHex":"a"}]}),
        json!({"format":"doc","synthesized":true,"resources":[{"kind":"binary","streamName":"1Table","offset":0,"dataHex":"gg"}]}),
        json!({"format":"doc","synthesized":true,"resources":[{"kind":"binary","streamName":"1Table","offset":u32::MAX,"dataHex":"00"}]}),
        json!({"format":"doc","synthesized":true,"resources":[{"kind":"binary","streamName":"1Table","offset":0.5,"dataHex":"00"}]}),
        json!({"format":"odt","synthesized":false,"resources":[{"kind":"xml","partName":"中".repeat(342),"xml":"x"}]}),
    ] { assert!(read_origin_value(&value).is_err()); }
}

#[test]
fn origin_resource_and_diagnostic_limits_apply_to_the_combined_deduplicated_history() {
    let many = (0..ORIGIN_COUNT_LIMIT + 1).map(|offset| doc(offset as u32, "00")).collect::<Vec<_>>();
    assert!(merge_origins(many).is_err());
    let duplicate = (0..ORIGIN_COUNT_LIMIT + 1).map(|_| doc(0, "00")).collect::<Vec<_>>();
    assert_eq!(merge_origins(duplicate).unwrap().len(), 1);
    let (mut a, mut b) = (doc(0, "00"), doc(1, "00"));
    a.resources = (0..600).map(|offset| DocumentOfficeSourceResource::Binary { stream_name: "1Table".into(), offset, data_hex: "00".into() }).collect();
    b.resources = (600..1200).map(|offset| DocumentOfficeSourceResource::Binary { stream_name: "1Table".into(), offset, data_hex: "00".into() }).collect();
    assert!(merge_origins(vec![a, b]).is_err());
    let (mut a, mut b) = (doc(0, "00"), doc(1, "00"));
    a.resource_errors = (0..600).map(|index| format!("First {index}")).collect();
    b.resource_errors = (0..600).map(|index| format!("Second {index}")).collect();
    assert!(merge_origins(vec![a, b]).is_err());
    let value = json!((0..ORIGIN_COUNT_LIMIT + 1).map(|_| serde_json::to_value(doc(0, "00")).unwrap()).collect::<Vec<_>>());
    assert!(read_origin_list_value(&value).is_err());
}

#[test]
fn original_byte_budgets_reject_binary_and_mixed_history_overflow_before_serialization() {
    let a = doc(0, &"00".repeat(BINARY_BYTE_LIMIT / 2 + 1));
    let b = doc(1, &"00".repeat(BINARY_BYTE_LIMIT / 2));
    assert!(merge_origins(vec![a, b]).is_err());
    let mut text = odt(&"x".repeat(XML_BYTE_LIMIT));
    text.resources.push(DocumentOfficeSourceResource::Xml { part_name: "content.xml".into(), xml: "x".into() });
    assert!(merge_origins(vec![text, doc(0, &"00".repeat(BINARY_BYTE_LIMIT))]).is_err());
}

#[test]
fn serialized_json_budget_counts_escape_expansion_and_refuses_truncated_output() {
    // 11 MiB of control characters fit the raw source budget but expand to more
    // than 64 MiB in JSON. The codec must fail, never write a shortened resource.
    let source = odt(&"\u{0001}".repeat(11 * 1024 * 1024));
    let error = encode_origins(vec![source]).unwrap_err();
    assert!(error.contains("64 MiB"));
}
