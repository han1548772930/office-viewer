//! Standard Office color predicates use one actual differential-style record.
use super::{ExportError, ExportWarning, SheetInput, StyleRegistry};
use office_io_common::{OfficeFilterNode, conditional_format_xml::{self as xml, Element}, filter_color};

pub(super) fn validate(fill: Option<&[Option<String>]>, text: Option<&[String]>) -> Result<(), ExportError> {
    let invalid = |value: &String| value.trim().is_empty() || value.len() > 256 || office_io_common::sparkline_io::css_argb(value).is_err();
    if fill.is_none() && text.is_none() || fill.is_some_and(|colors| colors.len() > 4096 || colors.iter().flatten().any(invalid))
        || text.is_some_and(|colors| colors.len() > 4096 || colors.iter().any(invalid)) {
        return Err(ExportError::invalid("autoFilter color predicate requires bounded CSS fill or text colors"));
    }
    Ok(())
}

pub(super) fn write_xml(sheet: &SheetInput, column: u32, fill: Option<&[Option<String>]>, text: Option<&[String]>,
    registry: &mut StyleRegistry, warnings: &mut Vec<ExportWarning>) -> Result<String, ExportError> {
    validate(fill, text)?;
    let (is_fill, colors) = if let Some(fill) = fill { (true, fill.to_vec()) }
        else { (false, text.unwrap_or_default().iter().cloned().map(Some).collect()) };
    if colors.is_empty() { return Ok(super::none_match()); }
    if colors.len() > 1 {
        super::warning(warnings, "AUTOFILTER_MULTIPLE_COLORS", sheet, column,
            "standard Office colorFilter accepts one color; the full predicate and current visibility are preserved in the project extension, while Office reapply uses its first color");
    }
    let mut dxf = Element::new(xml::MAIN_NAMESPACE, "dxf");
    if is_fill {
        let mut fill = Element::new(xml::MAIN_NAMESPACE, "fill");
        let mut pattern = Element::new(xml::MAIN_NAMESPACE, "patternFill");
        pattern.set_attr("patternType", if colors[0].is_some() { "solid" } else { "none" });
        if let Some(color) = &colors[0] {
            let mut foreground = Element::new(xml::MAIN_NAMESPACE, "fgColor");
            foreground.set_attr("rgb", office_io_common::sparkline_io::css_argb(color).map_err(ExportError::invalid)?); pattern.push(foreground);
        }
        fill.push(pattern); dxf.push(fill);
    } else {
        let mut font = Element::new(xml::MAIN_NAMESPACE, "font");
        let mut color = Element::new(xml::MAIN_NAMESPACE, "color");
        color.set_attr("rgb", office_io_common::sparkline_io::css_argb(colors[0].as_deref().unwrap_or("#000000")).map_err(ExportError::invalid)?);
        font.push(color); dxf.push(font);
    }
    let index = registry.register_conditional_style_xml(&dxf);
    Ok(format!("<colorFilter dxfId=\"{index}\" cellColor=\"{}\"/>", u8::from(is_fill)))
}

pub(in crate::xlsx) fn color_signature(node: &OfficeFilterNode, registry: &StyleRegistry) -> Result<Option<String>, String> {
    let dxf = filter_color::dxf_index(node)?.map(|index| {
        let xml = registry.conditional_styles.get(index).ok_or("AUTOFILTER_COLOR_REQUIRED_RESOURCE: unregistered differential style")?;
        xml::parse(xml)
    }).transpose()?;
    Ok(Some(filter_color::decode(node, dxf.as_ref(), None, None)?.signature()))
}
