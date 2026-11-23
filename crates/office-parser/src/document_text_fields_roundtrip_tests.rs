use serde_json::{Value, json};

use crate::parser::parse_office_bytes;

fn reopen(model: Value) -> Value {
    let (bytes, warnings) = office_writer::export_docx_bytes(&model.to_string()).unwrap();
    assert!(warnings.is_empty(), "{warnings:?}");
    let result = parse_office_bytes(&bytes, "native-fields.docx").unwrap();
    serde_json::to_value(result).unwrap()["payload"]["data"].clone()
}

fn paragraphs(data: &Value) -> Vec<&Value> {
    data["blocks"].as_array().unwrap().iter().filter(|block| block["details"]["kind"] == "paragraph").collect()
}

#[test]
fn edited_cache_uses_current_runs_and_preserves_lock_instruction_and_identity() {
    let result = reopen(json!({ "kind": "document", "name": "fields", "blocks": [{ "id": "p", "type": "paragraph",
        "runs": [{ "text": "A😀" }, { "text": "2026/11/23", "metadata": { "officeTextFieldStyleOutsideResult": true } }, { "text": " tail" }],
        "metadata": { "documentTextFields": { "version": 1, "fields": [{ "id": "日期字段-稳定标识", "field": "DATE", "instruction": " DATE \\@ \"yyyy/MM/dd\" ",
            "cachedResult": "obsolete", "startOffset": 3, "endOffset": 13, "locked": true, "language": "fr-FR" }] } } }] }));
    let blocks = paragraphs(&result);
    let field = &blocks[0]["details"]["documentTextFields"]["fields"][0];
    assert_eq!(blocks[0]["content"], "A😀2026/11/23 tail");
    assert_eq!(field["id"], "日期字段-稳定标识");
    assert_eq!(field["cachedResult"], "2026/11/23");
    assert_eq!(field["instruction"], "DATE \\@ \"yyyy/MM/dd\"");
    assert_eq!(field["startOffset"], 3);
    assert_eq!(field["endOffset"], 13);
    assert_eq!(field["locked"], true);
    assert_eq!(field["language"], "fr-FR");
}

#[test]
fn cross_paragraph_native_field_end_uses_the_reopened_paragraph_identity() {
    let result = reopen(json!({ "kind": "document", "name": "cross-field", "blocks": [
        { "id": "first", "type": "paragraph", "runs": [{ "text": "before 20" }], "metadata": { "documentTextFields": { "version": 1, "fields": [
            { "id": "cross-field", "field": "DATE", "instruction": "DATE", "startOffset": 7, "endOffset": 2, "endBlockId": "last" }
        ] } } },
        { "id": "last", "type": "paragraph", "runs": [{ "text": "26 after" }] }
    ] }));
    let blocks = paragraphs(&result);
    assert_eq!(blocks.len(), 2);
    let field = &blocks[0]["details"]["documentTextFields"]["fields"][0];
    assert_eq!(field["cachedResult"], "20\n26");
    assert_eq!(field["startOffset"], 7);
    assert_eq!(field["endOffset"], 2);
    assert_eq!(field["endBlockId"], blocks[1]["id"]);
}

#[test]
fn nested_result_fields_remain_definitions_and_block_automatic_outer_refresh() {
    let result = reopen(json!({ "kind": "document", "name": "nested", "blocks": [{ "id": "p", "type": "paragraph", "runs": [{ "text": "abcdef" }],
        "metadata": { "documentTextFields": { "version": 1, "fields": [
            { "id": "outer", "field": "REF", "instruction": "REF Example", "startOffset": 0, "endOffset": 6, "containsNestedFields": true },
            { "id": "inner", "field": "DATE", "instruction": "DATE", "startOffset": 1, "endOffset": 4, "locked": true }
        ] } } }] }));
    let blocks = paragraphs(&result);
    let fields = blocks[0]["details"]["documentTextFields"]["fields"].as_array().unwrap();
    let outer = fields.iter().find(|field| field["id"] == "outer").unwrap();
    let inner = fields.iter().find(|field| field["id"] == "inner").unwrap();
    assert_eq!(outer["containsNestedFields"], true);
    assert_eq!(outer["cachedResult"], "abcdef");
    assert_eq!(inner["cachedResult"], "bcd");
    assert_eq!(inner["locked"], true);
    assert_eq!(inner["metadata"]["source"]["parentFieldId"], "outer");
}

#[test]
fn field_results_cannot_cross_surrogate_pairs_or_partially_overlap() {
    for fields in [
        json!([{ "id": "half", "field": "DATE", "startOffset": 2, "endOffset": 3 }]),
        json!([{ "id": "a", "field": "DATE", "startOffset": 0, "endOffset": 3 }, { "id": "b", "field": "TIME", "startOffset": 1, "endOffset": 4 }]),
    ] {
        let model = json!({ "kind": "document", "blocks": [{ "id": "p", "type": "paragraph", "runs": [{ "text": "a😀b" }], "metadata": { "documentTextFields": { "version": 1, "fields": fields } } }] });
        assert!(office_writer::export_docx_bytes(&model.to_string()).is_err());
    }
}
