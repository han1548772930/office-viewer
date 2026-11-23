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

#[test]
fn an_unbookmarked_empty_native_field_remains_an_addressable_empty_paragraph() {
    let xml = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
        <w:r><w:fldChar w:fldCharType="begin"/><w:instrText>DATE</w:instrText><w:fldChar w:fldCharType="separate"/><w:fldChar w:fldCharType="end"/></w:r>
        </w:p><w:sectPr/></w:body></w:document>"#;
    let package = duke_pages::create_minimal_docx_package(Some(xml));
    let bytes = duke_pages::package_to_bytes(&package).unwrap();
    let envelope = serde_json::to_value(parse_office_bytes(&bytes, "empty-native-field.docx").unwrap()).unwrap();
    let data = &envelope["payload"]["data"]; let blocks = paragraphs(data);
    assert_eq!(blocks.len(), 1);
    assert_eq!(blocks[0]["content"], "");
    let field = &blocks[0]["details"]["documentTextFields"]["fields"][0];
    assert_eq!(field["field"], "DATE");
    assert_eq!(field["cachedResult"], "");
    assert_eq!(field["startOffset"], 0);
    assert_eq!(field["endOffset"], 0);
    assert!(data.get("nativeFieldSource").is_none());
    assert!(!envelope.to_string().contains("office_studio_internal_keep_"));
}

#[test]
fn a_native_result_crossing_an_empty_paragraph_keeps_both_newline_boundaries_after_reopen() {
    let result = reopen(json!({ "kind": "document", "name": "empty-cross-field", "blocks": [
        { "id": "first", "type": "paragraph", "runs": [{ "text": "A" }], "metadata": { "documentTextFields": { "version": 1, "fields": [
            { "id": "cross-empty", "field": "REF", "instruction": "REF Example", "startOffset": 0, "endOffset": 1, "endBlockId": "last" }
        ] } } },
        { "id": "empty", "type": "paragraph", "runs": [] },
        { "id": "last", "type": "paragraph", "runs": [{ "text": "B after" }] }
    ] }));
    let blocks = paragraphs(&result);
    assert_eq!(blocks.len(), 3);
    assert_eq!(blocks[1]["content"], "");
    let field = &blocks[0]["details"]["documentTextFields"]["fields"][0];
    assert_eq!(field["cachedResult"], "A\n\nB");
    assert_eq!(field["endBlockId"], blocks[2]["id"]);
    assert!(result.get("nativeFieldSource").is_none());
}

#[test]
fn native_empty_and_equal_range_children_keep_their_explicit_parent_after_export_and_reopen() {
    for (text, start, end) in [("", 0, 0), ("a😀b", 1, 3)] {
        let result = reopen(json!({ "kind": "document", "blocks": [{ "id": "p", "type": "paragraph", "runs": [{ "text": text }],
            "metadata": { "documentTextFields": { "version": 1, "fields": [
                { "id": "inner", "field": "DATE", "instruction": "DATE", "startOffset": start, "endOffset": end,
                    "metadata": { "source": { "parentFieldId": "outer" } } },
                { "id": "outer", "field": "REF", "instruction": "REF Example", "startOffset": start, "endOffset": end,
                    "metadata": { "source": { "parentFieldId": null } } }
            ] } } }] }));
        let blocks = paragraphs(&result);
        assert_eq!(blocks.len(), 1);
        assert_eq!(blocks[0]["content"], text);
        let fields = blocks[0]["details"]["documentTextFields"]["fields"].as_array().unwrap();
        let inner = fields.iter().find(|field| field["id"] == "inner").unwrap();
        let outer = fields.iter().find(|field| field["id"] == "outer").unwrap();
        assert_eq!(inner["metadata"]["source"]["parentFieldId"], "outer");
        assert!(outer["metadata"]["source"]["parentFieldId"].is_null());
        assert_eq!(outer["containsNestedFields"], true);
        assert!(outer["metadata"]["source"]["beginSequence"].as_u64().unwrap() < inner["metadata"]["source"]["beginSequence"].as_u64().unwrap());
        assert_eq!(inner["startOffset"], start);
        assert_eq!(inner["endOffset"], end);
    }
}

#[test]
fn native_empty_children_at_result_edges_and_adjacent_empty_roots_preserve_nesting() {
    let result = reopen(json!({ "kind": "document", "blocks": [{ "id": "p", "type": "paragraph", "runs": [{ "text": "AB" }],
        "metadata": { "documentTextFields": { "version": 1, "fields": [
            { "id": "outer", "field": "REF", "instruction": "REF Example", "startOffset": 0, "endOffset": 2,
                "metadata": { "source": { "parentFieldId": null } } },
            { "id": "first", "field": "DATE", "startOffset": 0, "endOffset": 0,
                "metadata": { "source": { "parentFieldId": "outer" } } },
            { "id": "last", "field": "TIME", "startOffset": 2, "endOffset": 2,
                "metadata": { "source": { "parentFieldId": "outer" } } },
            { "id": "before", "field": "PAGE", "startOffset": 0, "endOffset": 0,
                "metadata": { "source": { "parentFieldId": null } } },
            { "id": "after", "field": "NUMPAGES", "startOffset": 2, "endOffset": 2,
                "metadata": { "source": { "parentFieldId": null } } }
        ] } } }] }));
    let blocks = paragraphs(&result);
    assert_eq!(blocks[0]["content"], "AB");
    let fields = blocks[0]["details"]["documentTextFields"]["fields"].as_array().unwrap();
    assert_eq!(fields.len(), 5);
    for id in ["first", "last"] {
        let field = fields.iter().find(|field| field["id"] == id).unwrap();
        assert_eq!(field["metadata"]["source"]["parentFieldId"], "outer");
        assert_eq!(field["cachedResult"], "");
    }
    for id in ["outer", "before", "after"] {
        assert!(fields.iter().find(|field| field["id"] == id).unwrap()["metadata"]["source"]["parentFieldId"].is_null());
    }
}

#[test]
fn invalid_native_parent_graphs_fail_export_instead_of_becoming_different_fields() {
    for fields in [
        json!([{ "id": "a", "field": "DATE", "startOffset": 0, "endOffset": 1, "metadata": { "source": { "parentFieldId": "absent" } } }]),
        json!([{ "id": "a", "field": "DATE", "startOffset": 0, "endOffset": 1, "metadata": { "source": { "parentFieldId": "a" } } }]),
        json!([
            { "id": "a", "field": "DATE", "startOffset": 0, "endOffset": 1, "metadata": { "source": { "parentFieldId": "b" } } },
            { "id": "b", "field": "TIME", "startOffset": 0, "endOffset": 1, "metadata": { "source": { "parentFieldId": "a" } } }
        ]),
        json!([
            { "id": "a", "field": "DATE", "startOffset": 0, "endOffset": 1, "metadata": { "source": { "parentFieldId": null } } },
            { "id": "b", "field": "TIME", "startOffset": 0, "endOffset": 2, "metadata": { "source": { "parentFieldId": "a" } } }
        ]),
    ] {
        let model = json!({ "kind": "document", "blocks": [{ "id": "p", "type": "paragraph", "runs": [{ "text": "AB" }],
            "metadata": { "documentTextFields": { "version": 1, "fields": fields } } }] });
        assert!(office_writer::export_docx_bytes(&model.to_string()).is_err());
    }
}
