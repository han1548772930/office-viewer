//! A picture control owns its real object. The editor's empty input anchor and
//! image block share one native paragraph when they represent one picture.
use std::collections::HashMap;

use serde_json::Value;

use super::{ContentControl, decode_data_image, get};
use crate::ExportError;

pub(super) fn linked_image(src: &str) -> bool {
    let Some(target) = src.strip_prefix("https://").or_else(|| src.strip_prefix("http://")) else { return false; };
    !src.chars().any(char::is_whitespace) && !target.split(['/', '?', '#']).next().unwrap_or_default().is_empty()
}

pub(super) fn validate(control: &ContentControl, owned: &[&Value]) -> Result<Option<String>, ExportError> {
    let images = owned.iter().copied().filter(|block| get(block, "type").as_str() == Some("image")).collect::<Vec<_>>();
    let metadata = get(&control.value, "metadata");
    if images.is_empty() && get(&control.value, "showingPlaceholder").as_bool() == Some(true)
        && get(metadata, "pictureBlockId").is_null() { return Ok(None); }
    if images.len() != 1 { return Err(ExportError::invalid("DOCX_PICTURE_CONTROL_OWNERSHIP: a picture control requires exactly one owned image")); }
    let image = images[0];
    let id = get(image, "id").as_str().ok_or_else(|| ExportError::invalid("DOCX picture control image has no identity"))?;
    if !get(metadata, "pictureBlockId").is_null() && get(metadata, "pictureBlockId").as_str() != Some(id)
        || !get(metadata, "pictureAnchorBlockId").is_null() && get(metadata, "pictureAnchorBlockId").as_str() != get(&control.value, "blockId").as_str() {
        return Err(ExportError::invalid("DOCX_PICTURE_CONTROL_OWNERSHIP: picture metadata no longer identifies its real image and input anchor"));
    }
    let source = get(get(get(image, "metadata"), "image"), "src").as_str()
        .or_else(|| get(get(image, "metadata"), "src").as_str()).unwrap_or_default();
    if decode_data_image(source).is_err() && !linked_image(source) {
        return Err(ExportError::invalid("DOCX_PICTURE_CONTROL_SOURCE: the owned picture cannot be written as a real embedded or linked image"));
    }
    Ok(Some(id.to_owned()))
}

pub(super) fn merge_contents(blocks: &[Value], contents: &mut [String], controls: &HashMap<String, Vec<ContentControl>>) -> Result<(), ExportError> {
    for index in 0..blocks.len().saturating_sub(1) {
        let anchor = &blocks[index];
        let Some(anchor_id) = get(anchor, "id").as_str() else { continue; };
        let image = &blocks[index + 1];
        let Some(image_id) = get(image, "id").as_str() else { continue; };
        if get(image, "type").as_str() != Some("image") || !matches!(get(anchor, "type").as_str(), Some("paragraph" | "heading" | "bullet" | "numbered" | "checklist" | "quote"))
            || !get(get(anchor, "metadata"), "codeBlock").is_null()
            || !get(get(anchor, "metadata"), "documentTextFields").is_null()
                && get(get(get(anchor, "metadata"), "documentTextFields"), "fields").as_array().is_none_or(|fields| !fields.is_empty())
            || get(anchor, "runs").as_array().into_iter().flatten().any(|run| {
                get(run, "text").as_str().is_some_and(|text| !text.is_empty())
                    || ["note", "noteReference", "latexFormula"].iter().any(|key| !get(get(run, "metadata"), key).is_null())
            }) { continue; }
        let owns_pair = controls.get(anchor_id).into_iter().flatten().any(|control| {
            let ids = get(get(&control.value, "blockRange"), "blockIds").as_array();
            control.kind == "picture" && get(&control.value, "placement").as_str() == Some("block")
                && get(&control.value, "blockId").as_str() == Some(anchor_id)
                && ids.is_some_and(|ids| ids.len() == 2 && ids[0].as_str() == Some(anchor_id) && ids[1].as_str() == Some(image_id))
        });
        let source = get(image, "metadata");
        let input_bridge = get(source, "pictureInputAnchor").as_bool() == Some(true)
            && get(source, "anchorBlockId").as_str() == Some(anchor_id) && get(source, "anchorOffset").as_u64() == Some(0);
        if !owns_pair && !input_bridge { continue; }
        let paragraph = contents[index].strip_prefix("<w:p>").and_then(|value| value.strip_suffix("</w:p>"))
            .ok_or_else(|| ExportError::invalid("DOCX picture input anchor is not an ordinary native paragraph"))?;
        let object = contents[index + 1].strip_prefix("<w:p>").and_then(|value| value.strip_suffix("</w:p>"))
            .ok_or_else(|| ExportError::invalid("DOCX picture object is not an ordinary native paragraph"))?;
        contents[index] = format!("<w:p>{paragraph}{object}</w:p>");
        contents[index + 1].clear();
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use std::io::{Cursor, Read};
    use serde_json::json;
    use super::*;

    fn model(src: &str) -> Value {
        json!({ "kind": "document", "name": "picture-pair", "blocks": [
            { "id": "anchor", "type": "paragraph", "runs": [{ "text": "" }] },
            { "id": "image", "type": "image", "runs": [], "metadata": { "image": { "src": src, "width": 80, "height": 60 } } }
        ], "contentControls": [{ "id": "picture", "type": "picture", "placement": "block", "blockId": "anchor", "startOffset": 0, "endOffset": 0,
            "blockRange": { "blockIds": ["anchor", "image"] }, "metadata": { "pictureBlockId": "image", "pictureAnchorBlockId": "anchor" } }] })
    }

    #[test]
    fn a_real_picture_and_its_empty_input_anchor_write_one_native_paragraph() {
        let document = model("data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aXuoAAAAASUVORK5CYII=");
        let (bytes, warnings) = crate::export_docx_bytes(&document.to_string()).unwrap();
        assert!(warnings.is_empty(), "{warnings:?}");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
        let mut xml = String::new(); archive.by_name("word/document.xml").unwrap().read_to_string(&mut xml).unwrap();
        let root = office_io_common::chart_xml::parse(&xml).unwrap();
        let body = root.child("body").unwrap();
        let paragraph = office_io_common::document_content_controls::word_children(body, "p");
        assert_eq!(paragraph.len(), 1, "The input bridge must not add a physical blank paragraph");
        assert_eq!(paragraph[0].descendants("drawing").len(), 1);
        assert_eq!(body.descendants("sdt").len(), 1);
    }

    #[test]
    fn linked_picture_writes_an_external_image_relationship_and_real_drawing() {
        let (bytes, warnings) = crate::export_docx_bytes(&model("https://example.test/picture.png").to_string()).unwrap();
        assert!(warnings.is_empty(), "{warnings:?}");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
        let mut xml = String::new(); archive.by_name("word/document.xml").unwrap().read_to_string(&mut xml).unwrap();
        assert!(xml.contains("r:link="));
        let mut relationships = String::new(); archive.by_name("word/_rels/document.xml.rels").unwrap().read_to_string(&mut relationships).unwrap();
        assert!(office_io_common::document_custom_xml::relationships(&relationships).unwrap().iter()
            .any(|entry| entry.external && entry.target == "https://example.test/picture.png" && entry.kind.ends_with("/image")));
    }

    #[test]
    fn temporary_picture_keeps_its_physical_input_bridge_with_an_empty_field_collection() {
        let mut document = model("https://example.test/picture.png");
        document["contentControls"] = json!([]);
        document["blocks"][0]["metadata"] = json!({ "documentTextFields": { "version": 1, "fields": [] } });
        document["blocks"][1]["metadata"]["pictureInputAnchor"] = json!(true);
        document["blocks"][1]["metadata"]["anchorBlockId"] = json!("anchor");
        document["blocks"][1]["metadata"]["anchorOffset"] = json!(0);
        let (bytes, warnings) = crate::export_docx_bytes(&document.to_string()).unwrap();
        assert!(warnings.is_empty(), "{warnings:?}");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
        let mut xml = String::new(); archive.by_name("word/document.xml").unwrap().read_to_string(&mut xml).unwrap();
        let root = office_io_common::chart_xml::parse(&xml).unwrap();
        let body = root.child("body").unwrap();
        assert_eq!(body.children.iter().filter(|node| office_io_common::document_content_controls::word(node, "p")).count(), 1);
        assert_eq!(body.descendants("drawing").len(), 1); assert!(body.descendants("sdt").is_empty());
    }

    #[test]
    fn empty_text_with_a_note_formula_or_field_remains_a_distinct_paragraph() {
        for key in ["note", "noteReference", "latexFormula"] {
            let mut document = model("https://example.test/picture.png");
            document["blocks"][0]["runs"][0]["metadata"] = json!({});
            document["blocks"][0]["runs"][0]["metadata"][key] = json!({ "id": "entity" });
            let blocks = document["blocks"].as_array().unwrap();
            let mut contents = vec!["<w:p/>".into(), "<w:p/>".into()];
            let mut controls = HashMap::new();
            let control = ContentControl { value: document["contentControls"][0].clone(), id: "picture".into(),
                kind: "picture".into(), start: 0, end: 0, start_note_reference_count: None, end_note_reference_count: None, numeric_id: 1 };
            controls.insert("anchor".into(), vec![control]);
            merge_contents(blocks, &mut contents, &controls).unwrap();
            assert_eq!(contents, vec!["<w:p/>", "<w:p/>"]);
        }
    }

    #[test]
    fn invalid_owned_images_and_stale_picture_identities_reject_the_native_export() {
        for invalid in ["", "blob:temporary", "data:image/png;base64,!", "https://"] {
            assert!(crate::export_docx_bytes(&model(invalid).to_string()).is_err());
        }
        let mut stale = model("https://example.test/picture.png");
        stale["contentControls"][0]["metadata"]["pictureBlockId"] = json!("other");
        assert!(crate::export_docx_bytes(&stale.to_string()).is_err());
    }
}
