//! Deferred XLSX fixtures for validation origins, editable attributes and outside edits.

use std::io::{Cursor, Read, Write};

use office_io_common::sheet_validation_source::{MODEL_EXTENSION_URI, MODEL_NAMESPACE};
use serde_json::{Value, json};

use crate::parser::parse_office_bytes;

fn export(rules: Value) -> Vec<u8> {
    office_writer::export_xlsx_bytes(&json!({"kind":"workbook","sheets":[{"id":"data","name":"Data","cells":{},"dataValidations":rules}]}).to_string()).unwrap().0
}
fn parsed(bytes: &[u8]) -> Value { serde_json::to_value(parse_office_bytes(bytes, "validation-roundtrip.xlsx").unwrap()).unwrap() }
fn worksheet(bytes: &[u8]) -> String {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut xml = String::new(); archive.by_name("xl/worksheets/sheet1.xml").unwrap().read_to_string(&mut xml).unwrap(); xml
}
fn rewrite(bytes: &[u8], edit: impl Fn(String) -> String) -> Vec<u8> {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options = zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);
    for index in 0..archive.len() {
        let mut entry = archive.by_index(index).unwrap(); let name = entry.name().unwrap().into_owned();
        let mut bytes = Vec::new(); entry.read_to_end(&mut bytes).unwrap();
        if name == "xl/worksheets/sheet1.xml" { bytes = edit(String::from_utf8(bytes).unwrap()).into_bytes(); }
        writer.start_file(name, options).unwrap(); writer.write_all(&bytes).unwrap();
    }
    writer.finish().unwrap().into_inner()
}
fn remove_model_extension(xml: String) -> String {
    let start = xml.find(&format!("<ext uri=\"{MODEL_EXTENSION_URI}\"")).unwrap();
    let end = start + xml[start..].find("</ext>").unwrap() + "</ext>".len();
    format!("{}{}", &xml[..start], &xml[end..])
}
fn anchored() -> Value {
    json!({"id":"keep-origin","type":"custom","formulaAnchor":{"row":2,"col":2},"formula1":"=IFERROR(A1,0)>0",
        "ranges":[{"startRow":0,"startCol":0,"endRow":0,"endCol":0},{"startRow":3,"startCol":3,"endRow":4,"endCol":3}],
        "allowBlank":false,"showDropDown":false,"showInputMessage":true,"promptTitle":"Input","prompt":"Enter a positive value",
        "showErrorMessage":true,"errorStyle":"warning","errorTitle":"Invalid","error":"Check the referenced value","renderMode":"text",
        "metadata":{"palette":["#112233","#445566"],"future":{"preserve":true}},"vendor":{"key":"retained"}})
}

#[test]
fn original_validation_origin_ranges_and_editable_settings_survive_binary_reopen() {
    let original = anchored(); let bytes = export(json!([original.clone()]));
    let xml = worksheet(&bytes);
    assert!(xml.contains("sqref=\"A1 D4:D5\""));
    assert!(xml.contains("IF(ROW()-2&lt;1,1048577,ROW()-2)"));
    assert_eq!(xml.matches("<dataValidation ").count(), 1);
    let reopened = parsed(&bytes); let restored = &reopened["payload"]["data"]["sheets"][0]["dataValidations"][0];
    for field in ["id","type","ranges","formulaAnchor","formula1","allowBlank","showDropDown","showInputMessage","promptTitle","prompt",
        "showErrorMessage","errorStyle","errorTitle","error","renderMode","metadata","vendor"] { assert_eq!(restored[field], original[field], "{field}"); }
}

#[test]
fn checkbox_multi_selection_and_prompt_only_rules_keep_their_product_types() {
    let rules = json!([
        {"id":"check","type":"checkbox","formula1":"yes","formula2":"no","ranges":[{"startRow":0,"startCol":1,"endRow":3,"endCol":1}],"metadata":{"checkColor":"#123456"}},
        {"id":"multi","type":"listMultiple","formula1":"\"Red,Green\"","formula2":"#ff0000,#00ff00","renderMode":"custom","ranges":[{"startRow":0,"startCol":2,"endRow":3,"endCol":2}]},
        {"id":"any","type":"any","allowBlank":false,"showInputMessage":true,"prompt":"Any value is accepted","ranges":[{"startRow":0,"startCol":3,"endRow":3,"endCol":3}]}
    ]);
    let bytes = export(rules.clone()); let reopened = parsed(&bytes);
    let restored = reopened["payload"]["data"]["sheets"][0]["dataValidations"].as_array().unwrap();
    assert_eq!(restored.len(), 3);
    for (index, original) in rules.as_array().unwrap().iter().enumerate() {
        for field in ["id","type","ranges","formula1","formula2","renderMode","metadata"] {
            if !original[field].is_null() { assert_eq!(restored[index][field], original[field], "{index}: {field}"); }
        }
    }
    assert_eq!(restored[2]["prompt"], "Any value is accepted");
    assert!(worksheet(&bytes).contains("type=\"none\""));
}

#[test]
fn externally_changed_standard_formula_invalidates_the_original_model_snapshot() {
    let bytes = export(json!([anchored()]));
    let changed = rewrite(&bytes, |xml| {
        let start = xml.find("<dataValidations ").unwrap(); let end = start + xml[start..].find("</dataValidations>").unwrap();
        format!("{}{}{}", &xml[..start], xml[start..end].replace("&gt;0</formula1>", "&gt;10</formula1>"), &xml[end..])
    });
    let reopened = parsed(&changed); let rules = reopened["payload"]["data"]["sheets"][0]["dataValidations"].as_array().unwrap();
    assert_eq!(rules.len(), 2, "standard disjoint ranges are supplied by the imported reader");
    for rule in rules {
        assert!(rule["formula1"].as_str().unwrap().ends_with(">10"));
        assert_eq!(rule["formulaAnchor"], json!({"row":0,"col":0}));
        assert_ne!(rule["id"], "keep-origin");
    }
    assert!(reopened["diagnostics"]["warnings"].as_array().unwrap().iter().any(|warning| warning.as_str().is_some_and(|warning| warning.contains("stale validation"))));
}

#[test]
fn removing_project_metadata_keeps_the_real_standard_formula_and_shared_origin() {
    let bytes = rewrite(&export(json!([anchored()])), remove_model_extension);
    let reopened = parsed(&bytes); let rules = reopened["payload"]["data"]["sheets"][0]["dataValidations"].as_array().unwrap();
    assert_eq!(rules.len(), 2);
    assert_eq!(rules[0]["formula1"], rules[1]["formula1"]);
    assert!(rules[1]["formula1"].as_str().unwrap().contains("INDEX("));
    assert_eq!(rules[1]["formulaAnchor"], json!({"row":0,"col":0}));
}

#[test]
fn ordinary_office_none_type_reopens_as_an_any_value_prompt_without_the_project_extension() {
    let rules = json!([{"id":"any","type":"any","showInputMessage":true,"promptTitle":"Guidance","prompt":"Use any content",
        "ranges":[{"startRow":1,"startCol":1,"endRow":2,"endCol":1}]}]);
    let bytes = rewrite(&export(rules), remove_model_extension);
    let reopened = parsed(&bytes); let restored = &reopened["payload"]["data"]["sheets"][0]["dataValidations"];
    assert_eq!(restored.as_array().unwrap().len(), 1); assert_eq!(restored[0]["type"], "any");
    assert_eq!(restored[0]["promptTitle"], "Guidance"); assert_eq!(restored[0]["prompt"], "Use any content");
    assert_eq!(restored[0]["showInputMessage"], true);
}

#[test]
fn invalid_trusted_validation_metadata_rejects_import_instead_of_dropping_model_settings() {
    let bytes = rewrite(&export(json!([anchored()])), |xml| {
        let mut root = office_io_common::conditional_format_xml::parse(&xml).unwrap();
        let extension = root.child_mut("extLst").unwrap().children.iter_mut()
            .find(|node| node.attribute_ns("", "uri") == Some(MODEL_EXTENSION_URI)).unwrap();
        let state = extension.children.iter_mut().find(|node| node.is(MODEL_NAMESPACE, "validationState")).unwrap();
        let mut model: Value = serde_json::from_str(&state.text).unwrap();
        model[0]["formulaAnchor"]["row"] = json!(1_048_576); state.set_text(model.to_string());
        root.render()
    });
    let error = parse_office_bytes(&bytes, "invalid-validation-model.xlsx").unwrap_err();
    assert_eq!(error.code, crate::model::ParseErrorCode::InvalidOfficeFile);
    assert!(error.details.as_deref().unwrap().contains("DV_SOURCE_MODEL_INVALID"));
}
