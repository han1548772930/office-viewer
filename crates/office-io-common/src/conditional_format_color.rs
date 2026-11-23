//! Inspectable SpreadsheetML colors; authored color XML remains the authority.
use crate::conditional_format_xml::{Element, MAIN_NAMESPACE, STRICT_MAIN_NAMESPACE, X14_NAMESPACE};

const DRAWING: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const STRICT_DRAWING: &str = "http://purl.oclc.org/ooxml/drawingml/main";

fn spreadsheet(node: &Element) -> bool {
    matches!(node.namespace.as_str(), MAIN_NAMESPACE | STRICT_MAIN_NAMESPACE | X14_NAMESPACE | "")
}

fn theme_color_reference(node: &Element) -> bool {
    spreadsheet(node) && node.attrs.contains_key("theme") && matches!(node.name.as_str(), "color" | "fgColor" | "bgColor"
        | "fillColor" | "negativeFillColor" | "borderColor" | "negativeBorderColor" | "axisColor")
}

fn scheme_font(node: &Element) -> Option<&'static str> {
    if !node.is_main("scheme") { return None }
    match node.attrs.get("val").map(String::as_str) { Some("major") => Some("majorFont"), Some("minor") => Some("minorFont"), _ => None }
}

pub fn uses_theme(node: &Element) -> bool {
    theme_color_reference(node) || scheme_font(node).is_some() || node.children.iter().any(uses_theme)
}

/// Import and export reject the same unavailable declared Office dependencies.
pub fn validate_theme_references(node: &Element, theme: Option<&Element>) -> Result<(), String> {
    if theme_color_reference(node) && color_value(node, theme).is_none() {
        return Err("CF_SOURCE_REQUIRED_RESOURCE: conditional-format color cannot resolve its declared theme index or tint".into())
    }
    if let Some(font) = scheme_font(node) {
        if theme.is_none_or(|theme| !theme.descendants(font).iter().any(|node| matches!(node.namespace.as_str(), DRAWING | STRICT_DRAWING))) {
            return Err("CF_SOURCE_REQUIRED_RESOURCE: conditional-format font scheme has no preserved theme font".into())
        }
    }
    for child in &node.children { validate_theme_references(child, theme)?; }
    Ok(())
}

fn rgb(value: &str) -> Option<[u8; 3]> {
    let value = value.strip_prefix('#').unwrap_or(value);
    if (value.len() != 6 && value.len() != 8) || !value.bytes().all(|value| value.is_ascii_hexdigit()) { return None }
    let value = if value.len() == 8 { &value[2..] } else { value };
    Some([u8::from_str_radix(&value[0..2], 16).ok()?, u8::from_str_radix(&value[2..4], 16).ok()?, u8::from_str_radix(&value[4..6], 16).ok()?])
}

fn theme_color(theme: &Element, index: usize) -> Option<[u8; 3]> {
    let drawing = |node: &Element| matches!(node.namespace.as_str(), DRAWING | STRICT_DRAWING);
    let scheme = theme.descendants("clrScheme").into_iter().find(&drawing)?;
    let key = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"].get(index)?;
    let color = scheme.children.iter().find(|node| drawing(node) && node.name == *key)?.children.first()?;
    if !drawing(color) || !matches!(color.name.as_str(), "sysClr" | "srgbClr") { return None }
    rgb(color.attrs.get(if color.name == "sysClr" { "lastClr" } else { "val" })?)
}

fn tinted(color: [u8; 3], tint: f64) -> [u8; 3] {
    if tint == 0.0 { return color }
    let [r, g, b] = color.map(|value| f64::from(value) / 255.0);
    let max = r.max(g).max(b); let min = r.min(g).min(b); let delta = max - min;
    let light = (max + min) / 2.0;
    let light_after = if tint < 0.0 { light * (1.0 + tint) } else { light * (1.0 - tint) + tint };
    let saturation = if delta == 0.0 { 0.0 } else { delta / (1.0 - (2.0 * light - 1.0).abs()) };
    let hue = if delta == 0.0 { 0.0 } else if max == r { ((g - b) / delta).rem_euclid(6.0) }
        else if max == g { (b - r) / delta + 2.0 } else { (r - g) / delta + 4.0 };
    let chroma = (1.0 - (2.0 * light_after - 1.0).abs()) * saturation;
    let x = chroma * (1.0 - (hue.rem_euclid(2.0) - 1.0).abs()); let offset = light_after - chroma / 2.0;
    let channels = if hue < 1.0 { [chroma, x, 0.0] } else if hue < 2.0 { [x, chroma, 0.0] }
        else if hue < 3.0 { [0.0, chroma, x] } else if hue < 4.0 { [0.0, x, chroma] }
        else if hue < 5.0 { [x, 0.0, chroma] } else { [chroma, 0.0, x] };
    channels.map(|value| ((value + offset) * 255.0).round().clamp(0.0, 255.0) as u8)
}

pub fn color_value(node: &Element, theme: Option<&Element>) -> Option<String> {
    const INDEXED: &[&str] = &["000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
        "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF", "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080",
        "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF", "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF",
        "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99", "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696",
        "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333"];
    let color = if let Some(value) = node.attrs.get("rgb") { rgb(value)? }
        else if let Some(value) = node.attrs.get("theme") { theme_color(theme?, value.parse().ok()?)? }
        else if let Some(value) = node.attrs.get("indexed") { rgb(INDEXED.get(value.parse::<usize>().ok()?)?)? }
        else { return None };
    let tint = node.attrs.get("tint").map(|value| value.parse::<f64>().ok()).unwrap_or(Some(0.0))?;
    if !tint.is_finite() || !(-1.0..=1.0).contains(&tint) { return None }
    let [r, g, b] = tinted(color, tint); Some(format!("#{r:02X}{g:02X}{b:02X}"))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::conditional_format_xml::parse;

    #[test]
    fn theme_indexes_follow_spreadsheet_order_and_tint_adjusts_hsl_luminance() {
        let theme = parse(r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:accent1><a:srgbClr val="FF0000"/></a:accent1></a:clrScheme></a:themeElements></a:theme>"#).unwrap();
        assert_eq!(color_value(&parse(r#"<color theme="0"/>"#).unwrap(), Some(&theme)).as_deref(), Some("#FFFFFF"));
        assert_eq!(color_value(&parse(r#"<color theme="1"/>"#).unwrap(), Some(&theme)).as_deref(), Some("#000000"));
        assert_eq!(color_value(&parse(r#"<color theme="4" tint="0.5"/>"#).unwrap(), Some(&theme)).as_deref(), Some("#FF8080"));
        assert_eq!(color_value(&parse(r#"<color indexed="10"/>"#).unwrap(), None).as_deref(), Some("#FF0000"));
    }

    #[test]
    fn strict_theme_colors_are_inspectable_and_malformed_rgb_never_uses_utf8_byte_slices() {
        let theme = parse(r#"<a:theme xmlns:a="http://purl.oclc.org/ooxml/drawingml/main"><a:themeElements><a:clrScheme><a:accent1><a:srgbClr val="336699"/></a:accent1></a:clrScheme></a:themeElements></a:theme>"#).unwrap();
        assert_eq!(color_value(&parse(r#"<color theme="4"/>"#).unwrap(), Some(&theme)).as_deref(), Some("#336699"));
        for value in ["颜色", "ééé", "FFGG0000", "#12", "000000000"] {
            let mut color = Element::new("", "color"); color.set_attr("rgb", value); assert!(color_value(&color, None).is_none());
        }
        assert!(color_value(&parse(r#"<color theme="4"/>"#).unwrap(), None).is_none());
    }

    #[test]
    fn declared_theme_resources_validate_in_both_dialects_while_vendor_attributes_and_none_scheme_stay_independent() {
        let dependency = parse(&format!(r#"<dxf xmlns="{MAIN_NAMESPACE}"><font><color theme="4" tint="0.5"/><scheme val="major"/></font></dxf>"#)).unwrap();
        let vendor = parse(r#"<v:payload xmlns:v="urn:vendor" theme="brand-only"><v:scheme val="major"/></v:payload>"#).unwrap();
        let empty_vendor = parse(r#"<payload xmlns="" theme="brand-only"/>"#).unwrap();
        let none = parse(&format!(r#"<scheme xmlns="{MAIN_NAMESPACE}" val="none"/>"#)).unwrap();
        assert!(uses_theme(&dependency)); assert!(!uses_theme(&vendor)); assert!(!uses_theme(&empty_vendor)); assert!(!uses_theme(&none));
        validate_theme_references(&vendor, None).unwrap(); validate_theme_references(&empty_vendor, None).unwrap(); validate_theme_references(&none, None).unwrap();
        assert!(validate_theme_references(&dependency, None).unwrap_err().starts_with("CF_SOURCE_REQUIRED_RESOURCE"));
        for namespace in [DRAWING, STRICT_DRAWING] {
            let theme = parse(&format!(r#"<a:theme xmlns:a="{namespace}"><a:themeElements><a:clrScheme><a:accent1><a:srgbClr val="336699"/></a:accent1></a:clrScheme><a:fontScheme><a:majorFont/><a:minorFont/></a:fontScheme></a:themeElements></a:theme>"#)).unwrap();
            validate_theme_references(&dependency, Some(&theme)).unwrap();
            let x14 = parse(&format!(r#"<x:fillColor xmlns:x="{X14_NAMESPACE}" theme="4"/>"#)).unwrap();
            validate_theme_references(&x14, Some(&theme)).unwrap();
            for xml in [r#"<color theme="9"/>"#, r#"<color theme="4" tint="2"/>"#] {
                assert!(validate_theme_references(&parse(xml).unwrap(), Some(&theme)).unwrap_err().starts_with("CF_SOURCE_REQUIRED_RESOURCE"));
            }
            let mut missing_font = theme.clone();
            let font_scheme = missing_font.children.iter_mut().find(|node| node.is(namespace, "themeElements")).unwrap()
                .children.iter_mut().find(|node| node.is(namespace, "fontScheme")).unwrap();
            font_scheme.replace_children(|node| node.is(namespace, "majorFont"), Vec::new());
            let error = validate_theme_references(&dependency, Some(&missing_font)).unwrap_err();
            assert!(error.starts_with("CF_SOURCE_REQUIRED_RESOURCE")); assert!(error.contains("font scheme"));
        }
    }
}
