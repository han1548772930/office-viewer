//! Color-filter dependencies shared by native import, export and stale-cache checks.
use crate::{OfficeFilterNode, conditional_format_color::color_value, conditional_format_xml::Element};

pub const TABLE_FILTER_NAMESPACE: &str = "urn:office-studio:table-filter:1";
pub const TABLE_FILTER_EXTENSION_URI: &str = "{E3D8D1E4-224F-488F-A6D4-56AD07F3B728}";

#[derive(Debug, PartialEq, Eq)]
pub enum FilterColor { Fill(Option<String>), Text(String) }

impl FilterColor {
    pub fn signature(&self) -> String {
        match self { Self::Fill(None) => "fill:none".into(), Self::Fill(Some(color)) => format!("fill:{color}"), Self::Text(color) => format!("text:{color}") }
    }
}

pub fn dxf_index(node: &OfficeFilterNode) -> Result<Option<usize>, String> {
    node.attributes.get("dxfId").map(|id| id.parse::<usize>().map_err(|_| "AUTOFILTER_COLOR_INVALID: invalid dxfId".into())).transpose()
}

pub fn decode(node: &OfficeFilterNode, dxf: Option<&Element>, theme: Option<&Element>, palette: Option<&[String]>) -> Result<FilterColor, String> {
    if !node.is_standard("colorFilter") || !node.children.is_empty() || !node.text.trim().is_empty() { return Err("AUTOFILTER_COLOR_INVALID: invalid colorFilter".into()); }
    if dxf_index(node)?.is_some() && dxf.is_none() { return Err("AUTOFILTER_COLOR_REQUIRED_RESOURCE: unavailable differential style".into()); }
    let fill = match node.attributes.get("cellColor").map(String::as_str) {
        None | Some("1" | "true") => true, Some("0" | "false") => false, _ => return Err("AUTOFILTER_COLOR_INVALID: invalid cellColor flag".into()),
    };
    if dxf.is_some_and(|node| !node.is_main("dxf")) { return Err("AUTOFILTER_COLOR_REQUIRED_RESOURCE: invalid differential style root".into()); }
    let color = if fill {
        let pattern = dxf.and_then(|node| node.child("fill")).and_then(|node| node.child("patternFill"));
        match pattern {
            None => return Ok(FilterColor::Fill(None)),
            Some(pattern) if matches!(pattern.attribute_ns("", "patternType"), None | Some("none")) => return Ok(FilterColor::Fill(None)),
            Some(pattern) if pattern.attribute_ns("", "patternType") == Some("solid") => pattern.child("fgColor"),
            _ => return Err("AUTOFILTER_COLOR_UNPROJECTED: patterned/gradient fill requires its own matching adapter".into()),
        }
    } else { dxf.and_then(|node| node.child("font")).and_then(|node| node.child("color")) };
    let Some(color) = color else {
        return if fill { Err("AUTOFILTER_COLOR_REQUIRED_RESOURCE: solid fill has no color".into()) } else { Ok(FilterColor::Text("#000000".into())) };
    };
    let mut resolved = color.clone();
    if let Some(index) = color.attribute_ns("", "indexed") {
        let index = index.parse::<usize>().map_err(|_| "AUTOFILTER_COLOR_INVALID: invalid palette index")?;
        if index == 64 || index == 65 {
            let color = if index == 64 { "#000000" } else { "#FFFFFF" }.to_owned();
            return Ok(if fill { FilterColor::Fill(Some(color)) } else { FilterColor::Text(color) });
        }
        if let Some(palette) = palette {
            let rgb = palette.get(index).ok_or("AUTOFILTER_COLOR_REQUIRED_RESOURCE: missing custom palette entry")?;
            resolved.remove_attr("indexed"); resolved.set_attr("rgb", rgb);
        }
    }
    if matches!(color.attribute_ns("", "auto"), Some("1" | "true")) && !fill { return Ok(FilterColor::Text("#000000".into())); }
    let color = color_value(&resolved, theme).ok_or("AUTOFILTER_COLOR_REQUIRED_RESOURCE: unresolved RGB/theme/tint color")?;
    Ok(if fill { FilterColor::Fill(Some(color)) } else { FilterColor::Text(color) })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{conditional_format_xml::{self as xml, MAIN_NAMESPACE}, office_filter_fingerprint};
    fn node(attributes: &str) -> OfficeFilterNode {
        OfficeFilterNode::from_element(&xml::parse(&format!("<colorFilter xmlns=\"{MAIN_NAMESPACE}\" {attributes}/>")).unwrap())
    }
    fn style(contents: &str) -> Element { xml::parse(&format!("<dxf xmlns=\"{MAIN_NAMESPACE}\">{contents}</dxf>")).unwrap() }

    #[test]
    fn fill_none_and_font_color_read_actual_differential_style_components() {
        let fill = style(r#"<fill><patternFill patternType="solid"><fgColor rgb="FF123456"/></patternFill></fill>"#);
        assert_eq!(decode(&node(r#"dxfId="0""#), Some(&fill), None, None).unwrap(), FilterColor::Fill(Some("#123456".into())));
        let none = style(r#"<fill><patternFill patternType="none"/></fill>"#);
        assert_eq!(decode(&node(r#"dxfId="0" cellColor="true""#), Some(&none), None, None).unwrap(), FilterColor::Fill(None));
        let font = style(r#"<font><color rgb="FFCC0000"/></font>"#);
        assert_eq!(decode(&node(r#"dxfId="0" cellColor="false""#), Some(&font), None, None).unwrap(), FilterColor::Text("#CC0000".into()));
    }

    #[test]
    fn custom_palette_and_theme_tint_are_resolved_without_flattening_native_source() {
        let theme = xml::parse(r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme><a:accent1><a:srgbClr val="336699"/></a:accent1></a:clrScheme></a:themeElements></a:theme>"#).unwrap();
        let font = style(r#"<font><color theme="4" tint="0.5"/></font>"#);
        let source = font.render();
        assert!(decode(&node(r#"dxfId="0" cellColor="0""#), Some(&font), Some(&theme), None).unwrap().signature().starts_with("text:#"));
        assert_eq!(font.render(), source);
        let indexed = style(r#"<font><color indexed="0"/></font>"#);
        assert_eq!(decode(&node(r#"dxfId="0" cellColor="0""#), Some(&indexed), None, Some(&["FF102030".into()])).unwrap(), FilterColor::Text("#102030".into()));
        assert!(decode(&node(r#"dxfId="0" cellColor="0""#), Some(&font), None, None).is_err());
    }

    #[test]
    fn cache_signature_tracks_effective_colors_and_boolean_defaults() {
        let first = node(r#"dxfId="0" cellColor="1""#);
        let same = node(r#"dxfId="0" cellColor="true""#);
        let red = style(r#"<fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill></fill>"#);
        let blue = style(r#"<fill><patternFill patternType="solid"><fgColor rgb="FF0000FF"/></patternFill></fill>"#);
        let signature = |node: &OfficeFilterNode, dxf: &Element| office_filter_fingerprint(Some(node), |node| Ok(Some(decode(node, Some(dxf), None, None)?.signature()))).unwrap();
        assert_eq!(signature(&first, &red), signature(&same, &red));
        assert_ne!(signature(&first, &red), signature(&first, &blue));
        assert!(decode(&first, None, None, None).is_err());
        assert!(decode(&node(r#"dxfId="0" cellColor="yes""#), Some(&red), None, None).is_err());
    }
}
