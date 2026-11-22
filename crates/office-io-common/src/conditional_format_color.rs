//! Inspectable SpreadsheetML colors; authored color XML remains the authority.
use crate::conditional_format_xml::Element;

pub fn uses_theme(node: &Element) -> bool {
    node.attrs.contains_key("theme") || node.is_main("scheme") || node.children.iter().any(uses_theme)
}

fn rgb(value: &str) -> Option<[u8; 3]> {
    let value = value.strip_prefix('#').unwrap_or(value);
    if (value.len() != 6 && value.len() != 8) || !value.bytes().all(|value| value.is_ascii_hexdigit()) { return None }
    let value = if value.len() == 8 { &value[2..] } else { value };
    Some([u8::from_str_radix(&value[0..2], 16).ok()?, u8::from_str_radix(&value[2..4], 16).ok()?, u8::from_str_radix(&value[4..6], 16).ok()?])
}

fn theme_color(theme: &Element, index: usize) -> Option<[u8; 3]> {
    const DRAWING: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
    let scheme = theme.descendants("clrScheme").into_iter().find(|node| node.namespace == DRAWING)?;
    let key = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"].get(index)?;
    let color = scheme.children.iter().find(|node| node.namespace == DRAWING && node.name == *key)?.children.first()?;
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
}
