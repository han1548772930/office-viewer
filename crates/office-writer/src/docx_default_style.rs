use serde_json::Value;

use super::{color, get, run_properties_xml, signed_twips, text_color, twips, TWIPS_PER_PIXEL};
use crate::ExportError;

#[derive(Clone, Copy, Default)]
pub(super) struct RunDefaultEffects {
    pub vertical_alignment: bool,
    pub highlight: bool,
}

impl RunDefaultEffects {
    pub(super) fn for_model(model: &Value) -> Self {
        let text = get(get(model, "defaultStyle"), "text");
        Self {
            vertical_alignment: matches!(get(text, "verticalAlign").as_str(), Some("superscript" | "subscript")),
            highlight: get(text, "highlight").as_str().and_then(|value| highlight_properties(value.to_owned())).is_some(),
        }
    }
}

// ST_HighlightColor uses Office names; arbitrary RGB highlighting is run shading.
const HIGHLIGHTS: [(&str, &str); 16] = [
    ("black", "000000"), ("blue", "0000FF"), ("cyan", "00FFFF"), ("green", "00FF00"),
    ("magenta", "FF00FF"), ("red", "FF0000"), ("yellow", "FFFF00"), ("white", "FFFFFF"),
    ("darkBlue", "000080"), ("darkCyan", "008080"), ("darkGreen", "008000"), ("darkMagenta", "800080"),
    ("darkRed", "800000"), ("darkYellow", "808000"), ("darkGray", "808080"), ("lightGray", "C0C0C0"),
];

pub(super) fn highlight_properties(value: String) -> Option<String> {
    let value = value.trim();
    if let Some((name, _)) = HIGHLIGHTS.iter().find(|(name, _)| name.eq_ignore_ascii_case(value)) {
        return Some(format!("<w:highlight w:val=\"{name}\"/>"));
    }
    let rgb = color(value.to_owned())?;
    if let Some((name, _)) = HIGHLIGHTS.iter().find(|(_, candidate)| *candidate == rgb) {
        Some(format!("<w:highlight w:val=\"{name}\"/>"))
    } else {
        Some(format!("<w:highlight w:val=\"none\"/><w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"{rgb}\"/>"))
    }
}

fn finite(value: &Value) -> bool { value.as_f64().is_some_and(f64::is_finite) }
fn positive(value: &Value) -> bool { value.as_f64().is_some_and(|value| value.is_finite() && value > 0.0) }
fn nonnegative(value: &Value) -> bool { value.as_f64().is_some_and(|value| value.is_finite() && value >= 0.0) }
fn nonempty(value: &Value) -> bool { value.as_str().is_some_and(|value| !value.trim().is_empty()) }

pub(super) fn validate(defaults: &Value) -> Result<(), ExportError> {
    if get(defaults, "version").as_u64() != Some(1) || !get(defaults, "text").is_object() || !get(defaults, "paragraph").is_object() {
        return Err(ExportError::invalid("文档默认格式版本无效。"));
    }
    let text = get(defaults, "text");
    if !nonempty(get(text, "fontFamily")) || !positive(get(text, "fontSize"))
        || get(text, "color").as_str().and_then(|value| text_color(value.to_owned())).is_none()
        || ["bold", "italic", "underline", "strike"].iter().any(|field| !get(text, field).is_boolean())
        || text.get("eastAsiaFontFamily").is_some_and(|value| !nonempty(value))
        || text.get("highlight").is_some_and(|value| !value.is_string()
            || value.as_str().is_some_and(|value| !value.is_empty() && !value.eq_ignore_ascii_case("none") && highlight_properties(value.to_owned()).is_none()))
        || text.get("verticalAlign").is_some_and(|value| !matches!(value.as_str(), Some("baseline" | "superscript" | "subscript")))
    {
        return Err(ExportError::invalid("文档默认字符格式无效。"));
    }
    let paragraph = get(defaults, "paragraph");
    if !matches!(get(paragraph, "align").as_str(), Some("left" | "center" | "right" | "justify"))
        || !finite(get(paragraph, "indent")) || !positive(get(paragraph, "lineHeight"))
        || !nonnegative(get(paragraph, "spaceBefore")) || !nonnegative(get(paragraph, "spaceAfter"))
        || ["indentStart", "indentEnd", "indentFirstLine", "hanging"].iter().any(|field| paragraph.get(*field).is_some_and(|value| !finite(value)))
        || paragraph.get("spacingRule").is_some_and(|value| !matches!(value.as_str(), Some("auto" | "exact" | "atLeast")))
        || ["wordWrap", "snapToGrid"].iter().any(|field| paragraph.get(*field).is_some_and(|value| !value.is_boolean()))
    {
        return Err(ExportError::invalid("文档默认段落格式无效。"));
    }
    Ok(())
}

pub(super) fn paragraph_properties(paragraph: &Value) -> String {
    // Validation above guarantees the required numeric values.
    let before = twips(get(paragraph, "spaceBefore").as_f64().unwrap_or_default());
    let after = twips(get(paragraph, "spaceAfter").as_f64().unwrap_or_default());
    let rule = get(paragraph, "spacingRule").as_str().unwrap_or("auto");
    let line = (get(paragraph, "lineHeight").as_f64().unwrap_or(1.0) * if rule == "auto" { 240.0 } else { TWIPS_PER_PIXEL })
        .round().clamp(1.0, 31_680.0) as u32;
    let left = get(paragraph, "indentStart").as_f64().unwrap_or_else(|| get(paragraph, "indent").as_f64().unwrap_or_default() * 24.0);
    let right = get(paragraph, "indentEnd").as_f64().unwrap_or_default();
    let first = signed_twips(get(paragraph, "indentFirstLine").as_f64().unwrap_or_default() - get(paragraph, "hanging").as_f64().unwrap_or_default());
    let (first_property, first_value) = if first < 0 { ("hanging", -first) } else { ("firstLine", first) };
    let align = match get(paragraph, "align").as_str() { Some("justify") => "both", Some("center") => "center", Some("right") => "right", _ => "left" };
    let mut result = format!("<w:spacing w:before=\"{before}\" w:after=\"{after}\" w:line=\"{line}\" w:lineRule=\"{rule}\"/><w:ind w:left=\"{}\" w:right=\"{}\" w:{first_property}=\"{first_value}\"/><w:jc w:val=\"{align}\"/>", signed_twips(left), signed_twips(right));
    for field in ["wordWrap", "snapToGrid"] {
        if let Some(value) = get(paragraph, field).as_bool() { result.push_str(&format!("<w:{field} w:val=\"{}\"/>", u8::from(value))); }
    }
    result
}

pub(super) fn styles_xml(model: &Value, mut generated: String) -> Result<String, ExportError> {
    let Some(defaults) = model.get("defaultStyle") else { return Ok(generated); };
    validate(defaults)?;
    let run = run_properties_xml(get(defaults, "text"));
    let paragraph = paragraph_properties(get(defaults, "paragraph"));
    let replacement = format!("<w:docDefaults><w:rPrDefault><w:rPr>{run}<w:lang w:val=\"zh-CN\"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>{paragraph}</w:pPr></w:pPrDefault></w:docDefaults>");
    // This is only the writer's controlled template, never untrusted source XML.
    let start = generated.find("<w:docDefaults>").ok_or_else(|| ExportError::invalid("DOCX 默认格式模板缺失。"))?;
    let end = generated[start..].find("</w:docDefaults>").map(|end| start + end + "</w:docDefaults>".len())
        .ok_or_else(|| ExportError::invalid("DOCX 默认格式模板未闭合。"))?;
    generated.replace_range(start..end, &replacement);
    Ok(generated)
}

#[cfg(test)]
mod tests {
    use serde_json::json;
    use super::*;

    fn defaults() -> Value { json!({ "version": 1, "text": { "fontFamily": "Georgia", "eastAsiaFontFamily": "宋体", "fontSize": 15, "color": "auto", "bold": true, "italic": false, "underline": true, "strike": false, "highlight": "darkYellow", "verticalAlign": "superscript" }, "paragraph": { "align": "right", "indent": 0, "indentStart": -12, "indentEnd": 9, "indentFirstLine": 4, "hanging": 18, "spacingRule": "exact", "lineHeight": 24, "spaceBefore": 0, "spaceAfter": 12, "wordWrap": false, "snapToGrid": true } }) }

    #[test]
    fn document_defaults_replace_only_generated_defaults_and_keep_source_style_support() {
        let xml = styles_xml(&json!({ "defaultStyle": defaults() }), super::super::styles_xml()).unwrap();
        let declaration = xml.split("</w:docDefaults>").next().unwrap();
        assert!(declaration.contains("w:ascii=\"Georgia\" w:hAnsi=\"Georgia\" w:eastAsia=\"宋体\""));
        assert!(declaration.contains("<w:sz w:val=\"30\"/>") && declaration.contains("<w:color w:val=\"auto\"/>"));
        assert!(declaration.contains("<w:i w:val=\"0\"/>") && declaration.contains("<w:u w:val=\"single\"/>"));
        assert!(declaration.contains("<w:highlight w:val=\"darkYellow\"/>") && declaration.contains("<w:vertAlign w:val=\"superscript\"/>"));
        assert!(declaration.contains("w:before=\"0\" w:after=\"180\" w:line=\"360\" w:lineRule=\"exact\""));
        assert!(declaration.contains("w:left=\"-180\" w:right=\"135\" w:hanging=\"210\""));
        assert!(declaration.contains("<w:wordWrap w:val=\"0\"/>") && declaration.contains("<w:snapToGrid w:val=\"1\"/>"));
        assert!(xml.contains("w:styleId=\"Heading1\"") && xml.contains("w:styleId=\"FootnoteReference\""));
        assert_eq!(xml.matches("<w:docDefaults>").count(), 1);
    }

    #[test]
    fn old_models_keep_the_generated_template_and_invalid_defaults_fail_export() {
        let template = super::super::styles_xml();
        assert_eq!(styles_xml(&json!({}), template.clone()).unwrap(), template);
        for field in ["version", "text", "paragraph"] {
            let mut invalid = defaults(); invalid[field] = Value::Null;
            assert!(styles_xml(&json!({ "defaultStyle": invalid }), template.clone()).is_err());
        }
        let mut invalid = defaults(); invalid["paragraph"]["lineHeight"] = json!(0);
        assert!(styles_xml(&json!({ "defaultStyle": invalid }), template.clone()).is_err());
        for (section, field, value) in [
            ("text", "color", json!("red")), ("text", "highlight", json!("tomato")),
            ("paragraph", "wordWrap", json!("yes")), ("paragraph", "snapToGrid", json!(1)),
        ] {
            let mut invalid = defaults(); invalid[section][field] = value;
            assert!(styles_xml(&json!({ "defaultStyle": invalid }), template.clone()).is_err());
        }
    }

    #[test]
    fn rgb_highlights_use_office_names_when_available_and_shading_for_custom_colors() {
        assert_eq!(highlight_properties("#808000".into()), Some("<w:highlight w:val=\"darkYellow\"/>".into()));
        assert_eq!(highlight_properties("#abcdef".into()), Some("<w:highlight w:val=\"none\"/><w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"ABCDEF\"/>".into()));
        assert_eq!(highlight_properties("none".into()), None);
    }
}
