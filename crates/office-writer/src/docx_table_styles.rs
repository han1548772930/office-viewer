//! CSS-pixel table styles and native property preservation at the DOCX boundary.
use std::collections::BTreeMap;

use serde_json::{Value, json};

use super::super::{WriteContext, get, source_xml};
use crate::ExportError;

const PX_TO_TWIPS: f64 = 15.0;
const MAX_TWIPS: f64 = 2_147_483_647.0;
const SIDES: [&str; 4] = ["top", "right", "bottom", "left"];
const TABLE_BORDERS: [(&str, &str); 6] = [("top", "top"), ("left", "left"), ("bottom", "bottom"), ("right", "right"),
    ("insideHorizontal", "insideH"), ("insideVertical", "insideV")];

pub(super) fn number(value: &Value, name: &str, positive: bool) -> Result<Option<f64>, ExportError> {
    let value = get(value, name);
    if value.is_null() { return Ok(None); }
    let number = value.as_f64().filter(|number| number.is_finite() && (!positive || *number > 0.0))
        .ok_or_else(|| ExportError::invalid(format!("DOCX table {name} must be a finite {}CSS-pixel number", if positive { "positive " } else { "" })))?;
    if (number * PX_TO_TWIPS).abs() > MAX_TWIPS { return Err(ExportError::invalid(format!("DOCX table {name} exceeds the native measurement range"))); }
    Ok(Some(number))
}

fn pixels(value: f64) -> i64 { (value * PX_TO_TWIPS).round() as i64 }

fn enum_value<'a>(value: &'a Value, field: &str, allowed: &[&str], fallback: &'a str) -> Result<&'a str, ExportError> {
    match value.get(field).filter(|value| !value.is_null()) {
        None => Ok(fallback),
        Some(Value::String(value)) if allowed.contains(&value.as_str()) => Ok(value),
        Some(_) => Err(ExportError::invalid(format!("DOCX table {field} has an invalid value"))),
    }
}

fn boolean(value: &Value, field: &str, fallback: bool) -> Result<bool, ExportError> {
    match value.get(field).filter(|value| !value.is_null()) {
        None => Ok(fallback), Some(Value::Bool(value)) => Ok(*value),
        Some(_) => Err(ExportError::invalid(format!("DOCX table {field} must be boolean"))),
    }
}

fn css_color(context: &mut WriteContext<'_>, value: &str) -> Result<String, ExportError> {
    if value == "auto" { return Ok("auto".into()); }
    let (rgb, alpha) = office_io_common::chart_color::rgba(value).map_err(ExportError::invalid)?;
    if alpha < 1.0 {
        context.warn("DOCX_TABLE_COLOR_ALPHA", "Word table borders and shading use an opaque color; the authored CSS alpha is retained in the table source extension.");
    }
    let composed = rgb.map(|channel| (f64::from(channel) * alpha + 255.0 * (1.0 - alpha)).round() as u8);
    Ok(format!("{:02X}{:02X}{:02X}", composed[0], composed[1], composed[2]))
}

fn color_value<'a>(style: &'a Value, field: &str) -> Result<Option<&'a str>, ExportError> {
    match style.get(field).filter(|value| !value.is_null()) {
        None => Ok(None), Some(Value::String(value)) => Ok(Some(value)),
        Some(_) => Err(ExportError::invalid(format!("DOCX table {field} must be a color string"))),
    }
}

fn fill<'a>(style: &'a Value) -> Result<Option<&'a str>, ExportError> {
    Ok(color_value(style, "fill")?.or(color_value(style, "background")?))
}

pub(super) fn margins(table: &Value, cell: Option<&Value>) -> Result<[f64; 4], ExportError> {
    let table_style = get(table, "style");
    let margin = cell.map(|cell| get(get(cell, "style"), "margin")).filter(|value| !value.is_null())
        .unwrap_or_else(|| get(table_style, "cellMargin"));
    if margin.is_null() { return Ok([8.0; 4]); }
    if !margin.is_object() { return Err(ExportError::invalid("DOCX table cell margin must be an object")); }
    let mut result = [0.0; 4];
    for (index, side) in SIDES.iter().enumerate() {
        result[index] = number(margin, side, false)?.filter(|value| *value >= 0.0)
            .ok_or_else(|| ExportError::invalid("DOCX table cell margins require four nonnegative CSS-pixel sides"))?;
    }
    Ok(result)
}

fn margin_xml(tag: &str, margins: [f64; 4]) -> String {
    let contents = SIDES.into_iter().zip(margins).map(|(side, value)| format!("<w:{side} w:w=\"{}\" w:type=\"dxa\"/>", pixels(value))).collect::<String>();
    format!("<w:{tag}>{contents}</w:{tag}>")
}

fn border(context: &mut WriteContext<'_>, name: &str, value: Option<&Value>) -> Result<(String, Value), ExportError> {
    let fallback = json!({"width":1,"color":"#aab3bf","style":"solid"});
    let value = value.filter(|value| !value.is_null()).unwrap_or(&fallback);
    if !value.is_object() { return Err(ExportError::invalid("DOCX table border must be an object")); }
    let width = number(value, "width", false)?.filter(|width| *width >= 0.0).ok_or_else(|| ExportError::invalid("DOCX table border width must be nonnegative"))?;
    let style = enum_value(value, "style", &["solid", "dot", "dash"], "solid")?;
    let color = color_value(value, "color")?.ok_or_else(|| ExportError::invalid("DOCX table border color is missing"))?;
    let native_color = css_color(context, color)?;
    let native_style = if width == 0.0 { "nil" } else { match style { "dot" => "dotted", "dash" => "dashed", _ => "single" } };
    // Word's border size is an eighth of a point, or one sixth of a CSS pixel.
    let size = (width * 6.0).round();
    if width > 0.0 && !(2.0..=96.0).contains(&size) {
        context.warn("DOCX_TABLE_BORDER_WIDTH_NATIVE_LIMIT", "Word clamps a visible table border to 0.25–12 points; the exact authored width is retained in the table source extension.");
    }
    let size = if width == 0.0 { 0 } else { size.clamp(2.0, 96.0) as u32 };
    Ok((format!("<w:{name} w:val=\"{native_style}\" w:sz=\"{size}\" w:space=\"0\" w:color=\"{native_color}\"/>"), json!({"width":width,"style":style,"color":color})))
}

fn source_property(style: &Value) -> Option<&Value> { style.get("nativeProperties").filter(|value| !value.is_null()) }

/// Replace managed properties only when their domain value changed. Unsupported
/// native properties keep their resolved namespaces and original XML meaning.
fn properties(context: &mut WriteContext<'_>, tag: &str, style: &Value, generated: Vec<(&str, String, Value, bool)>) -> Result<String, ExportError> {
    let word_namespace = source_xml::WORD_NAMESPACE;
    let mut root = if let Some(source) = source_property(style) {
        if get(source, "version").as_u64() != Some(1) { return Err(ExportError::invalid("DOCX table native property source version is invalid")); }
        let xml = get(source, "propertiesXml").as_str().ok_or_else(|| ExportError::invalid("DOCX table native properties require XML"))?;
        let namespaces = get(source, "namespaces").as_object().ok_or_else(|| ExportError::invalid("DOCX table native properties require namespace bindings"))?;
        let mut declarations = String::new();
        for (prefix, namespace) in namespaces {
            if !prefix.is_empty() && !source_xml::valid_name(prefix) { return Err(ExportError::invalid("DOCX table native namespace prefix is invalid")); }
            let namespace = namespace.as_str().ok_or_else(|| ExportError::invalid("DOCX table native namespace must be a string"))?;
            let name = if prefix.is_empty() { "xmlns".to_owned() } else { format!("xmlns:{prefix}") };
            declarations.push_str(&crate::xml::attr(&name, namespace));
        }
        let wrapper = source_xml::parse_xml(&format!("<office{declarations}>{xml}</office>"), true).map_err(ExportError::invalid)?;
        let mut root = wrapper.children().next().cloned().ok_or_else(|| ExportError::invalid("DOCX table native properties have no root"))?;
        if wrapper.children().count() != 1 { return Err(ExportError::invalid("DOCX table native properties have multiple roots")); }
        for (prefix, namespace) in namespaces {
            let name = if prefix.is_empty() { "xmlns".to_owned() } else { format!("xmlns:{prefix}") };
            if !root.attributes.iter().any(|attribute| attribute.name == name) {
                root.attributes.push(source_xml::Attribute { name, namespace: source_xml::XMLNS_NAMESPACE.into(), local: prefix.clone(), value: namespace.as_str().unwrap_or_default().to_owned() });
            }
        }
        source_xml::canonicalize_office_namespaces(&mut root).map_err(ExportError::invalid)?;
        if !root.is_word(tag) { return Err(ExportError::invalid("DOCX table native property root has the wrong type")); }
        root
    } else {
        source_xml::parse_xml(&format!("<w:{tag} xmlns:w=\"{word_namespace}\"/>"), false).map_err(ExportError::invalid)?
    };
    let baseline = source_property(style).map(|source| get(source, "baseline"));
    let rewrites = source_property(style).and_then(|source| source.get("rewriteProperties"));
    let rewrites = match rewrites {
        None => &[][..],
        Some(Value::Array(values)) if tag == "tcPr" && values.iter().all(|value| matches!(value.as_str(), Some("tcW" | "tcBorders"))) => values.as_slice(),
        Some(_) => return Err(ExportError::invalid("DOCX table native property rewrite list is invalid")),
    };
    for (name, xml, current, always_replace) in generated {
        if !always_replace && !rewrites.iter().any(|value| value.as_str() == Some(name))
            && baseline.is_some_and(|baseline| baseline.get(name) == Some(&current)) { continue; }
        root.content.retain(|content| !matches!(content, source_xml::Content::Element(child) if child.is_word(name)));
        if !xml.is_empty() {
            let wrapper = source_xml::parse_xml(&format!("<w:{tag} xmlns:w=\"{word_namespace}\">{xml}</w:{tag}>"), false).map_err(ExportError::invalid)?;
            root.content.extend(wrapper.content);
        }
    }
    let order = match tag {
        "tblPr" => &["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook", "tblCaption", "tblDescription", "tblPrChange"][..],
        "trPr" => &["cnfStyle", "divId", "gridBefore", "gridAfter", "wBefore", "wAfter", "cantSplit", "trHeight", "tblHeader", "tblCellSpacing", "jc", "hidden", "ins", "del", "trPrChange"][..],
        _ => &["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark", "headers", "cellIns", "cellDel", "cellMerge", "tcPrChange"][..],
    };
    root.content.sort_by_key(|content| match content {
        source_xml::Content::Element(child) if child.namespace == word_namespace => order.iter().position(|name| *name == child.local).unwrap_or(order.len()),
        _ => order.len(),
    });
    root.empty = root.content.is_empty();
    context.source_styles.restore_table_properties(&mut root).map_err(ExportError::invalid)?;
    source_xml::validate_namespace_bindings(&root, &BTreeMap::new()).map_err(ExportError::invalid)?;
    Ok(root.render(None))
}

pub(super) fn table(context: &mut WriteContext<'_>, table: &Value, width: i64, column_widths: &[i64]) -> Result<String, ExportError> {
    let style = get(table, "style");
    if !style.is_null() && !style.is_object() { return Err(ExportError::invalid("DOCX table style must be an object")); }
    let align = enum_value(style, "align", &["left", "center", "right"], "left")?;
    let layout = enum_value(style, "layout", &["auto", "fixed"], "fixed")?;
    let indent = number(style, "indent", false)?.unwrap_or(0.0);
    let offset_y = number(style, "offsetY", false)?;
    let cell_margin = margins(table, None)?;
    let borders = get(style, "borders");
    if !borders.is_null() && !borders.is_object() { return Err(ExportError::invalid("DOCX table borders must be an object")); }
    let mut border_xml = String::new();
    let mut border_values = serde_json::Map::new();
    for (field, tag) in TABLE_BORDERS {
        let (xml, value) = border(context, tag, borders.get(field))?;
        border_xml.push_str(&xml); border_values.insert(field.to_owned(), value);
    }
    let fill = fill(style)?;
    let shading = fill.map(|value| css_color(context, value)).transpose()?.map(|color| format!("<w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"{color}\"/>")).unwrap_or_default();
    let floating = offset_y.filter(|value| *value != 0.0).map(|value| format!("<w:tblpPr w:leftFromText=\"0\" w:rightFromText=\"0\" w:topFromText=\"0\" w:bottomFromText=\"0\" w:vertAnchor=\"text\" w:horzAnchor=\"text\" w:tblpX=\"{}\" w:tblpY=\"{}\"/>", pixels(indent), pixels(value))).unwrap_or_default();
    properties(context, "tblPr", style, vec![
        ("tblW", format!("<w:tblW w:w=\"{width}\" w:type=\"dxa\"/>"), json!([width, column_widths]), false),
        ("jc", format!("<w:jc w:val=\"{align}\"/>"), json!(align), false),
        ("tblInd", format!("<w:tblInd w:w=\"{}\" w:type=\"dxa\"/>", pixels(indent)), json!(indent), false),
        ("tblLayout", format!("<w:tblLayout w:type=\"{}\"/>", if layout == "auto" { "autofit" } else { "fixed" }), json!(layout), false),
        ("tblCellMar", margin_xml("tblCellMar", cell_margin), json!(cell_margin), false),
        ("tblBorders", format!("<w:tblBorders>{border_xml}</w:tblBorders>"), Value::Object(border_values), false),
        ("shd", shading, json!(fill), false),
        ("tblpPr", floating, json!([offset_y, indent]), false),
    ])
}

pub(super) fn row(context: &mut WriteContext<'_>, table: &Value, row: &Value, index: usize, header_rows: usize) -> Result<String, ExportError> {
    let legacy = get(table, "rowStyles").as_array().and_then(|rows| rows.get(index)).unwrap_or(&Value::Null);
    let height = match number(row, "height", true)? {
        Some(height) => Some(height),
        None => match number(legacy, "height", true)? { Some(height) => Some(height), None => number(legacy, "heightPx", true)? },
    };
    let rule = if row.get("heightRule").is_some() { enum_value(row, "heightRule", &["auto", "atLeast", "exact"], if height.is_some() { "atLeast" } else { "auto" })? }
        else { enum_value(legacy, "heightRule", &["auto", "atLeast", "exact"], if height.is_some() { "atLeast" } else { "auto" })? };
    if height.is_none() && rule != "auto" { return Err(ExportError::invalid("DOCX table exact/minimum row height requires a height")); }
    let allow_split = boolean(row, "allowSplit", boolean(legacy, "allowSplit", true)?)?;
    let height_xml = height.map(|height| format!("<w:trHeight w:val=\"{}\" w:hRule=\"{rule}\"/>", pixels(height))).unwrap_or_default();
    properties(context, "trPr", get(row, "style"), vec![
        ("trHeight", height_xml, json!([height, rule]), false),
        ("cantSplit", format!("<w:cantSplit w:val=\"{}\"/>", u8::from(!allow_split)), json!(allow_split), false),
        ("tblHeader", format!("<w:tblHeader w:val=\"{}\"/>", u8::from(index < header_rows)), json!(index < header_rows), false),
        // This domain's grid is completely covered; legacy omissions cannot be
        // applied again after anchors have been explicitly materialized.
        ("gridBefore", String::new(), Value::Null, true),
        ("gridAfter", String::new(), Value::Null, true),
        ("wBefore", String::new(), Value::Null, true),
        ("wAfter", String::new(), Value::Null, true),
    ])
}

pub(super) fn cell(context: &mut WriteContext<'_>, table: &Value, cell: &Value, width: i64, span: usize,
    row_span: usize, continuation: bool, row_index: usize) -> Result<String, ExportError> {
    let style = get(cell, "style");
    if !style.is_null() && !style.is_object() { return Err(ExportError::invalid("DOCX table cell style must be an object")); }
    let align = enum_value(style, "verticalAlign", &["top", "center", "bottom"], "top")?;
    let cell_margin = margins(table, Some(cell))?;
    let margin = style.get("margin").filter(|value| !value.is_null()).map(|_| margin_xml("tcMar", cell_margin)).unwrap_or_default();
    let borders = get(style, "borders");
    if !borders.is_null() && !borders.is_object() { return Err(ExportError::invalid("DOCX table cell borders must be an object")); }
    let mut border_xml = String::new();
    let mut border_values = serde_json::Map::new();
    for side in SIDES {
        if let Some(value) = borders.get(side).filter(|value| !value.is_null()) {
            let (xml, value) = border(context, side, Some(value))?;
            border_xml.push_str(&xml); border_values.insert(side.to_owned(), value);
        }
    }
    let fill = fill(style)?;
    let shading = fill.map(|value| css_color(context, value)).transpose()?.map(|color| format!("<w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"{color}\"/>")).unwrap_or_default();
    let mut continuation_style;
    let source_style = if continuation {
        continuation_style = style.clone();
        let mapping = get(style, "nativeContinuations");
        if !mapping.is_null() && !mapping.is_object() { return Err(ExportError::invalid("DOCX table continuation properties require a source mapping")); }
        let row_id = get(table, "rows").as_array().and_then(|rows| rows.get(row_index)).and_then(|row| row.get("id")).and_then(Value::as_str);
        let owner_id = get(cell, "id").as_str();
        let matches = mapping.as_object().into_iter().flat_map(|mapping| mapping.values()).filter(|source| {
            row_id.is_some() && get(source, "rowId").as_str() == row_id && get(source, "ownerCellId").as_str() == owner_id
                && get(source, "colSpan").as_u64() == Some(span as u64)
        }).collect::<Vec<_>>();
        if matches.len() > 1 { return Err(ExportError::invalid("DOCX table continuation source has duplicate physical row identities")); }
        if let Some(source) = matches.first() {
            if let Some(object) = continuation_style.as_object_mut() { object.insert("nativeProperties".into(), (*source).clone()); }
        } else if let Some(object) = continuation_style.as_object_mut() { object.remove("nativeProperties"); }
        &continuation_style
    } else { style };
    properties(context, "tcPr", source_style, vec![
        ("tcW", format!("<w:tcW w:w=\"{width}\" w:type=\"dxa\"/>"), json!(width), false),
        ("gridSpan", if span > 1 { format!("<w:gridSpan w:val=\"{span}\"/>") } else { String::new() }, json!(span), true),
        ("vMerge", if row_span > 1 { format!("<w:vMerge w:val=\"{}\"/>", if continuation { "continue" } else { "restart" }) } else { String::new() }, json!([row_span, continuation]), true),
        ("tcBorders", if border_xml.is_empty() { String::new() } else { format!("<w:tcBorders>{border_xml}</w:tcBorders>") }, Value::Object(border_values), false),
        ("shd", shading, json!(fill), false),
        ("tcMar", margin, json!(style.get("margin").filter(|value| !value.is_null()).map(|_| cell_margin)), false),
        ("vAlign", format!("<w:vAlign w:val=\"{align}\"/>"), json!(align), false),
        ("hMerge", String::new(), Value::Null, true),
    ])
}

pub(super) fn validate_column_types(table: &Value, columns: usize) -> Result<(), ExportError> {
    if let Some(value) = table.get("titleRow").filter(|value| !value.is_null()) {
        if !value.is_boolean() { return Err(ExportError::invalid("DOCX table titleRow must be boolean")); }
    }
    let Some(types) = table.get("columnTypes").filter(|value| !value.is_null()) else { return Ok(()); };
    let types = types.as_object().ok_or_else(|| ExportError::invalid("DOCX table columnTypes must be an object"))?;
    for (column, config) in types {
        if column.parse::<usize>().ok().filter(|column| *column < columns).is_none() { return Err(ExportError::invalid("DOCX table column type is outside the logical grid")); }
        enum_value(config, "type", &["text", "number", "date", "checkbox", "dropdown"], "text")?;
        if let Some(options) = config.get("dropdownOptions").filter(|value| !value.is_null()) {
            let options = options.as_array().ok_or_else(|| ExportError::invalid("DOCX table dropdown options must be an array"))?;
            let mut ids = std::collections::BTreeSet::new();
            for option in options {
                let id = get(option, "id").as_str().filter(|id| !id.is_empty()).ok_or_else(|| ExportError::invalid("DOCX table dropdown option requires an ID"))?;
                if !ids.insert(id) || !get(option, "label").is_string() { return Err(ExportError::invalid("DOCX table dropdown option has a duplicate ID or invalid label")); }
                if let Some(color) = color_value(option, "color")? { office_io_common::chart_color::rgba(color).map_err(ExportError::invalid)?; }
            }
        }
    }
    Ok(())
}
