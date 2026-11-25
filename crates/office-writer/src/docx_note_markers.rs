//! Reference and body markers retain source rPr and custom symbol runs.
use office_io_common::document_note_source::NOTE_SAFE_INTEGER;
use serde_json::{Value, json};

use super::super::{CONTENT_NAMESPACES, NoteKind, WriteContext, get, source_xml::{self, Content, Element}, xml_text};
use crate::ExportError;

#[derive(Clone, Default)]
pub(in super::super) struct ReferenceProperties {
    authored: bool,
    custom_mark: Option<String>,
    text_style: Option<Value>,
    baseline: Option<Value>,
    body_source: Option<Value>,
}

impl ReferenceProperties {
    pub fn read(value: &Value) -> Result<Self, ExportError> {
        let custom_mark = value.get("customMark").map(|value| value.as_str().map(str::to_owned)
            .ok_or_else(|| ExportError::invalid("注释自定义标记必须是字符串。"))).transpose()?;
        if custom_mark.as_ref().is_some_and(|mark| mark.encode_utf16().count() > 4096) { return Err(ExportError::invalid("注释自定义标记超过长度预算。")); }
        let text_style = value.get("referenceTextStyle").map(|value| if value.is_object() { Ok(value.clone()) }
            else { Err(ExportError::invalid("注释引用文字样式必须是对象。")) }).transpose()?;
        if value.get("showBodyReference").is_some_and(|value| !value.is_boolean()) { return Err(ExportError::invalid("注释正文标记显隐必须是布尔值。")); }
        let source = get(get(value, "metadata"), "officeNoteSource");
        let baseline = source.get("referenceTextStyleBaseline").map(|value| if value.is_object() { Ok(value.clone()) }
            else { Err(ExportError::invalid("注释引用样式来源基线必须是对象。")) }).transpose()?;
        let body_source = source.get("bodyMarkerSource").map(|value| {
            if value.is_object() && get(value, "version").as_u64() == Some(1) { Ok(value.clone()) }
            else { Err(ExportError::invalid("注释正文标记来源版本或结构无效。")) }
        }).transpose()?;
        Ok(Self { authored: true, custom_mark, text_style, baseline, body_source })
    }
}

fn reference_style(kind: NoteKind) -> &'static str {
    if kind == NoteKind::Footnote { "FootnoteReference" } else { "EndnoteReference" }
}

fn marker_properties(context: &mut WriteContext<'_>, run: &Value, kind: NoteKind, style: Option<&Value>) -> Result<String, ExportError> {
    let mut author = run.as_object().cloned().unwrap_or_default();
    if let Some(style) = style.and_then(Value::as_object) { for (key, value) in style { author.insert(key.clone(), value.clone()); } }
    let author = Value::Object(author);
    let mut properties = context.run_properties(&author)?;
    let native = get(get(&author, "metadata"), "source");
    let native_format = native.get("sourceFormat").unwrap_or_else(|| get(get(native, "details"), "sourceFormat"));
    let cleared = get(get(&author, "metadata"), "officeStoryTextFormatCleared").as_bool() == Some(true);
    // A missing rStyle in a native reference is an observable source fact.
    // New markers and explicitly cleared markers use the destination default.
    if properties.child("rStyle").is_none() && (native_format.is_null() || cleared) {
        let mut style = Element::word("rStyle"); style.set_word_attribute("val", reference_style(kind));
        properties.content.insert(0, Content::Element(style));
    }
    Ok(properties.render(None))
}

fn with_source_format(run: &Value, format: &Value, author: &ReferenceProperties) -> Result<Value, ExportError> {
    if !format.is_object() || get(format, "version").as_u64() != Some(1) {
        return Err(ExportError::invalid("注释标记源文字格式版本或结构无效。"));
    }
    let mut result = run.as_object().cloned().unwrap_or_default();
    if let Some(style) = author.text_style.as_ref().and_then(Value::as_object) {
        for (key, value) in style { result.insert(key.clone(), value.clone()); }
    }
    let mut metadata = get(run, "metadata").as_object().cloned().unwrap_or_default();
    metadata.insert("source".into(), json!({ "sourceFormat": format }));
    if let Some(baseline) = &author.baseline { metadata.insert("officeTextStyleBaseline".into(), baseline.clone()); }
    result.insert("metadata".into(), Value::Object(metadata));
    Ok(Value::Object(result))
}

fn preserved_marks(context: &mut WriteContext<'_>, source: &Value, mark: &str, kind: NoteKind, author: &ReferenceProperties, run: &Value) -> Result<Option<String>, ExportError> {
    if !source.is_object() || get(source, "version").as_u64() != Some(1) {
        return Err(ExportError::invalid("注释自定义标记来源版本或结构无效。"));
    }
    let baseline = get(source, "text").as_str().filter(|text| text.encode_utf16().count() <= 4096)
        .ok_or_else(|| ExportError::invalid("注释自定义标记来源文字无效。"))?;
    let sources = get(source, "runs").as_array().filter(|runs| !runs.is_empty() && runs.len() <= 4096)
        .ok_or_else(|| ExportError::invalid("注释自定义标记来源 run 列表无效。"))?;
    if baseline != mark { return Ok(None); }
    let mut output = String::new();
    for source in sources {
        let mut native = source_xml::parse_xml(get(source, "sourceXml").as_str().ok_or_else(|| ExportError::invalid("注释自定义标记来源 XML 缺失。"))?, false).map_err(ExportError::invalid)?;
        source_xml::canonicalize_office_namespaces(&mut native).map_err(ExportError::invalid)?;
        if !native.is_word("r") || !native.text().trim().is_empty()
            || !native.children().any(|node| node.is_word("t") || node.is_word("sym"))
            || native.children().any(|node| !node.is_word("rPr") && !node.is_word("t") && !node.is_word("sym")) {
            return Err(ExportError::invalid("注释自定义标记来源只能包含真实文字或符号 run。"));
        }
        let current = with_source_format(run, get(source, "sourceFormat"), author)?;
        let properties = marker_properties(context, &current, kind, None)?;
        native.content.retain(|content| !matches!(content, Content::Element(node) if node.is_word("rPr")));
        let contents = native.content.iter().map(|content| match content {
            Content::Element(element) => element.render(None), Content::Text(text) => xml_text(text),
        }).collect::<String>();
        output.push_str(&format!("{}{properties}{contents}</{}>", native.opening(None, false), native.name));
    }
    Ok(Some(output))
}

pub(in super::super) fn reference_run(context: &mut WriteContext<'_>, run: &Value, kind: NoteKind, id: i64, author: &ReferenceProperties) -> Result<String, ExportError> {
    if id.unsigned_abs() > NOTE_SAFE_INTEGER { return Err(ExportError::invalid("注释引用 ID 超出安全整数范围。")); }
    let source = get(get(run, "metadata"), "noteReference");
    let custom_mark = if author.authored { author.custom_mark.as_deref() } else { get(source, "customMark").as_str() };
    let properties = marker_properties(context, run, kind, author.text_style.as_ref())?;
    let mut reference = if let Some(xml) = get(source, "sourceXml").as_str() {
        source_xml::parse_xml(xml, false).map_err(ExportError::invalid)?
    } else { Element::word(kind.reference_tag()) };
    source_xml::canonicalize_office_namespaces(&mut reference).map_err(ExportError::invalid)?;
    if !reference.is_word("footnoteReference") && !reference.is_word("endnoteReference") { return Err(ExportError::invalid("注释引用来源类型无效。")); }
    reference.name = format!("w:{}", kind.reference_tag());
    reference.local = kind.reference_tag().into();
    reference.set_word_attribute("id", id.to_string());
    reference.attributes.retain(|attribute| attribute.namespace != source_xml::WORD_NAMESPACE || attribute.local != "customMarkFollows");
    if custom_mark.is_some() { reference.set_word_attribute("customMarkFollows", "1"); }
    if let Some(mark) = custom_mark {
        if let Some(source) = source.get("customMarkSource") {
            if let Some(custom) = preserved_marks(context, source, mark, kind, author, run)? {
                return Ok(format!("<w:r>{properties}{}</w:r>{custom}", reference.render(None)));
            }
        }
    }
    let custom = custom_mark.map_or_else(String::new, |mark| format!("<w:t xml:space=\"preserve\">{}</w:t>", xml_text(mark)));
    Ok(format!("<w:r>{properties}{}{custom}</w:r>", reference.render(None)))
}

fn body_marker_xml(context: &mut WriteContext<'_>, author: &ReferenceProperties, kind: NoteKind) -> Result<String, ExportError> {
    let mut run = author.text_style.clone().unwrap_or_else(|| json!({}));
    if let Some(source) = &author.body_source {
        if let Some(mark) = &author.custom_mark {
            if source.get("runs").is_some() {
                if let Some(custom) = preserved_marks(context, source, mark, kind, author, &run)? { return Ok(custom); }
            }
        }
        let format = source.get("sourceFormat").or_else(|| get(source, "runs").as_array()
            .and_then(|runs| runs.first()).and_then(|run| run.get("sourceFormat")))
            .ok_or_else(|| ExportError::invalid("注释正文标记来源缺少真实文字格式。"))?;
        run = with_source_format(&run, format, author)?;
    }
    let properties = marker_properties(context, &run, kind, author.text_style.as_ref())?;
    let marker = if let Some(mark) = &author.custom_mark {
        format!("<w:t xml:space=\"preserve\">{}</w:t>", xml_text(mark))
    } else {
        let mut marker = author.body_source.as_ref().and_then(|source| source.get("sourceXml")).map(|value|
            value.as_str().ok_or_else(|| ExportError::invalid("注释正文引用来源 XML 必须为字符串。"))
                .and_then(|xml| source_xml::parse_xml(xml, false).map_err(ExportError::invalid))).transpose()?
            .unwrap_or_else(|| Element::word(kind.body_reference_tag()));
        source_xml::canonicalize_office_namespaces(&mut marker).map_err(ExportError::invalid)?;
        if !marker.is_word("footnoteRef") && !marker.is_word("endnoteRef") { return Err(ExportError::invalid("注释正文引用来源类型无效。")); }
        marker.name = format!("w:{}", kind.body_reference_tag()); marker.local = kind.body_reference_tag().into();
        marker.render(None)
    };
    Ok(format!("<w:r>{properties}{marker}</w:r>"))
}

fn tag_body_marker(xml: &str, custom: bool) -> Result<String, ExportError> {
    let mut paragraph = source_xml::parse_xml(&format!("<w:p {CONTENT_NAMESPACES}>{xml}</w:p>"), false)
        .map_err(ExportError::invalid)?;
    source_xml::canonicalize_office_namespaces(&mut paragraph).map_err(ExportError::invalid)?;
    let mut output = String::new();
    for content in &mut paragraph.content {
        let Content::Element(run) = content else { return Err(ExportError::invalid("注释正文标记必须由真实 run 组成。")); };
        if !run.is_word("r") { return Err(ExportError::invalid("注释正文标记必须由真实 run 组成。")); }
        super::extension_attribute(run, "bodyMarker", if custom { "custom" } else { "automatic" });
        output.push_str(&run.render(None));
    }
    Ok(output)
}

pub(super) fn body_marker(context: &mut WriteContext<'_>, value: &Value, kind: NoteKind) -> Result<String, ExportError> {
    let author = ReferenceProperties::read(value)?;
    if get(value, "showBodyReference").as_bool() == Some(false) { return Ok(String::new()); }
    tag_body_marker(&body_marker_xml(context, &author, kind)?, author.custom_mark.is_some())
}
