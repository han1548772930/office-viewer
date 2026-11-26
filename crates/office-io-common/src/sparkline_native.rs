//! Native x14 projection. The local extension preserves features without an x14 equivalent.

use std::collections::{BTreeMap, BTreeSet};
use crate::conditional_format_xml::{self as xml, Element};
use super::{OfficeSparklineSource, SparklineAxis, SparklineExtremum, SparklinePoint, SparklineSpec,
    COLOR_NAMES, MAX_GROUP_SOURCE_BYTES, MAX_MEMBERS, argb, css_argb, normalize_source_reference, normalize_target};

pub struct NativeSparklineGroup {
    pub group: Element,
    pub settings: SparklineSpec,
    pub members: Vec<(String, String)>,
}

fn flag(group: &Element, name: &str, default: bool) -> Result<bool, String> {
    match group.attribute_ns("", name) {
        None => Ok(default), Some("1" | "true") => Ok(true), Some("0" | "false") => Ok(false),
        _ => Err(format!("SPARKLINE_NATIVE_INVALID: invalid {name} flag")),
    }
}
fn number(group: &Element, name: &str) -> Result<Option<f64>, String> {
    group.attribute_ns("", name).map(|value| value.parse::<f64>().ok().filter(|value| value.is_finite())
        .ok_or_else(|| format!("SPARKLINE_NATIVE_INVALID: invalid {name} number"))).transpose()
}
fn unique_child<'a>(group: &'a Element, namespace: &str, name: &str) -> Result<Option<&'a Element>, String> {
    let mut children = group.children.iter().filter(|child| child.is(namespace, name));
    let first = children.next();
    if children.next().is_some() { return Err(format!("SPARKLINE_NATIVE_INVALID: duplicate {name}")); }
    Ok(first)
}
fn color_css(node: &Element, theme: Option<&Element>, palette: Option<&[String]>) -> Result<String, String> {
    let mut resolved = node.clone();
    if let Some(index) = node.attribute_ns("", "indexed") {
        let index = index.parse::<usize>().map_err(|_| "SPARKLINE_NATIVE_INVALID: invalid indexed color")?;
        if let Some(palette) = palette {
            let color = palette.get(index).ok_or("SPARKLINE_COLOR_UNRESOLVED: custom palette index is unavailable")?;
            resolved.remove_attr("indexed"); resolved.set_attr("rgb", color);
        }
    }
    if flag(node, "auto", false)? && !["rgb", "theme", "indexed"].iter().any(|key| node.attribute_ns("", key).is_some()) {
        resolved.set_attr("rgb", "FF000000");
    }
    let color = crate::conditional_format_color::color_value(&resolved, theme)
        .ok_or("SPARKLINE_COLOR_UNRESOLVED: native color/theme/tint cannot resolve")?;
    let (rgb, _) = argb(&color).ok_or("SPARKLINE_COLOR_UNRESOLVED: invalid resolved sRGB")?;
    let alpha = resolved.attribute_ns("", "rgb").and_then(argb).map(|(_, alpha)| alpha).unwrap_or(1.0);
    if alpha == 1.0 { Ok(format!("#{:02X}{:02X}{:02X}", rgb[0], rgb[1], rgb[2])) }
    else { Ok(crate::chart_color::css_rgba(rgb, alpha)) }
}

pub fn decode_native_group(
    group: &Element, group_id: &str, theme: Option<&Element>, theme_xml: Option<&str>,
    indexed_colors: Option<&[String]>, warnings: &mut Vec<String>,
) -> Result<NativeSparklineGroup, String> {
    if !group.is(xml::X14_NAMESPACE, "sparklineGroup") { return Err("SPARKLINE_NATIVE_INVALID: wrong group namespace".into()); }
    let kind = match group.attribute_ns("", "type").unwrap_or("line") {
        "line" => "line", "column" => "column", "stacked" => "winloss", _ => return Err("SPARKLINE_NATIVE_INVALID: unknown native type".into()),
    };
    let empty = match group.attribute_ns("", "displayEmptyCellsAs").unwrap_or("gap") {
        "gap" => "gap", "zero" => "zero", "span" => "connect", _ => return Err("SPARKLINE_NATIVE_INVALID: unknown displayEmptyCellsAs".into()),
    };
    let extremum = |maximum: bool| -> Result<SparklineExtremum, String> {
        let mode = group.attribute_ns("", if maximum { "maxAxisType" } else { "minAxisType" }).unwrap_or("individual");
        if !matches!(mode, "individual" | "group" | "custom") { return Err("SPARKLINE_NATIVE_INVALID: unknown min/maxAxisType".into()); }
        let value = number(group, if maximum { "manualMax" } else { "manualMin" })?;
        if mode == "custom" && value.is_none() { return Err("SPARKLINE_NATIVE_INVALID: missing custom axis value".into()); }
        Ok(SparklineExtremum { mode: mode.into(), value })
    };
    let minimum = extremum(false)?; let maximum = extremum(true)?;
    let width = number(group, "lineWeight")?.unwrap_or(0.75);
    if !(0.0..=1584.0).contains(&width) { return Err("SPARKLINE_NATIVE_INVALID: line weight exceeds the Office admission budget".into()); }
    let mut colors = BTreeMap::new();
    for name in COLOR_NAMES {
        let color = match unique_child(group, xml::X14_NAMESPACE, name)? {
            Some(color) => {
                if flag(color, "auto", false)? { warnings.push("SPARKLINE_AUTO_COLOR_RETAINED: automatic Office color is retained; editor theme behavior remains unverified".into()); }
                color_css(color, theme, indexed_colors)?
            }
            None => if name == "colorAxis" { "#000000".into() } else if name == "colorNegative" || name == "colorLow" { "#D00000".into() } else { "#4472C4".into() },
        };
        colors.insert(name.to_owned(), color);
    }
    let mut points = BTreeMap::new();
    for (name, attribute, color) in [
        ("highPoint", "high", "colorHigh"), ("firstPoint", "first", "colorFirst"), ("lowPoint", "low", "colorLow"),
        ("lastPoint", "last", "colorLast"), ("negativePoint", "negative", "colorNegative"), ("markersPoint", "markers", "colorMarkers"),
    ] { points.insert(name.into(), SparklinePoint { visible: Some(flag(group, attribute, false)?), color: colors.get(color).cloned() }); }
    let members_node = unique_child(group, xml::X14_NAMESPACE, "sparklines")?.ok_or("SPARKLINE_NATIVE_INVALID: missing native group members")?;
    let mut members = Vec::new(); let mut targets = BTreeSet::new();
    for member in &members_node.children {
        if !member.is(xml::X14_NAMESPACE, "sparkline") { return Err("SPARKLINE_NATIVE_INVALID: unknown native member element".into()); }
        let formula = unique_child(member, xml::XM_NAMESPACE, "f")?.ok_or("SPARKLINE_REFERENCE_UNRESOLVED: missing native source formula")?;
        let target = unique_child(member, xml::XM_NAMESPACE, "sqref")?.ok_or("SPARKLINE_NATIVE_INVALID: missing target reference")?;
        let target = normalize_target(&target.text).ok_or("SPARKLINE_NATIVE_INVALID: target must be one in-bounds cell")?;
        if !targets.insert(target.clone()) { return Err("SPARKLINE_NATIVE_INVALID: repeated target in native group".into()); }
        if !member.attrs.iter().all(|(name, _)| name == "xmlns" || name.starts_with("xmlns:") || member.attribute_namespaces.get(name).is_some_and(|ns| ns == "http://schemas.openxmlformats.org/markup-compatibility/2006"))
            || member.children.iter().any(|child| !child.is(xml::XM_NAMESPACE, "f") && !child.is(xml::XM_NAMESPACE, "sqref")) {
            return Err("SPARKLINE_NATIVE_MEMBER_UNRESOLVED: unsupported member metadata cannot be silently discarded".into());
        }
        members.push((target, normalize_source_reference(&formula.text)?));
        if members.len() > MAX_MEMBERS { return Err("SPARKLINE_MODEL_LIMIT: native member count exceeds admission budget".into()); }
    }
    if members.is_empty() { return Err("SPARKLINE_NATIVE_INVALID: empty native group".into()); }
    if flag(group, "dateAxis", false)? || unique_child(group, xml::XM_NAMESPACE, "f")?.is_some() {
        warnings.push("SPARKLINE_DATE_AXIS_RETAINED: native date axis and formula are preserved; Canvas date-axis behavior remains unverified".into());
    }
    let mut template = group.clone();
    template.replace_children(|child| child.is(xml::X14_NAMESPACE, "sparklines"), Vec::new());
    if template.rendered_size() > MAX_GROUP_SOURCE_BYTES { return Err("SPARKLINE_SOURCE_LIMIT: group template exceeds 128 KiB".into()); }
    let uses_theme = template.children.iter().any(|child| child.is(xml::X14_NAMESPACE, &child.name) && child.attribute_ns("", "theme").is_some());
    let uses_palette = template.children.iter().any(|child| child.is(xml::X14_NAMESPACE, &child.name) && child.attribute_ns("", "indexed").is_some());
    let known_attributes = ["manualMax", "manualMin", "lineWeight", "type", "dateAxis", "displayEmptyCellsAs", "markers", "high", "low", "first", "last", "negative", "displayXAxis", "displayHidden", "minAxisType", "maxAxisType", "rightToLeft"];
    let unknown = template.attrs.iter().any(|(key, _)| !key.starts_with("xmlns") && !known_attributes.contains(&key.as_str()))
        || template.children.iter().any(|child| !(child.namespace == xml::X14_NAMESPACE && COLOR_NAMES.contains(&child.name.as_str()) || child.is(xml::XM_NAMESPACE, "f")));
    if unknown { warnings.push("SPARKLINE_NATIVE_SOURCE_RETAINED: extra native group metadata is preserved in the Office source".into()); }
    let axis_custom = minimum.mode == "custom" && maximum.mode == "custom";
    let settings = SparklineSpec {
        sparkline_type: kind.into(), source_range: members[0].1.clone(), color: colors["colorSeries"].clone(),
        negative_color: points["negativePoint"].visible.filter(|flag| *flag).and_then(|_| colors.get("colorNegative").cloned()),
        high_point_color: points["highPoint"].visible.filter(|flag| *flag).and_then(|_| colors.get("colorHigh").cloned()),
        low_point_color: points["lowPoint"].visible.filter(|flag| *flag).and_then(|_| colors.get("colorLow").cloned()),
        empty_cells: empty.into(), axis: if axis_custom { "custom".into() } else { "individual".into() },
        min: minimum.value, max: maximum.value, group_id: Some(group_id.into()), theme_type: Some("custom".into()),
        points: Some(points), axis_options: Some(SparklineAxis { visible: Some(flag(group, "displayXAxis", false)?),
            reverse: Some(flag(group, "rightToLeft", false)?), color: colors.get("colorAxis").cloned() }),
        extremum_min: Some(minimum), extremum_max: Some(maximum), non_numeric_cells: Some("gap".into()),
        contain_hidden_cells: Some(flag(group, "displayHidden", false)?), line_width: Some(crate::points_to_pixels(width)),
        show_gradient: Some(false), show_radius: Some(false),
        office_source: Some(OfficeSparklineSource { version: 1, group_xml: template.render(), resolved_colors: colors,
            theme_xml: if uses_theme || unknown { theme_xml.map(str::to_owned) } else { None },
            indexed_colors: if uses_palette || unknown { indexed_colors.map(<[String]>::to_vec) } else { None } }),
    };
    settings.validate()?;
    Ok(NativeSparklineGroup { group: group.clone(), settings, members })
}

/// Merge authored settings into their retained native template; unused native properties survive.
pub fn native_group_for_spec(spec: &SparklineSpec) -> Result<Option<Element>, String> {
    spec.validate()?;
    if spec.sparkline_type == "pie" { return Ok(None); }
    let mut group = spec.office_source.as_ref().map(|source| xml::parse(&source.group_xml)).transpose()?
        .unwrap_or_else(|| Element::prefixed(xml::X14_NAMESPACE, "x14", "sparklineGroup"));
    group.set_attr("type", if spec.sparkline_type == "winloss" { "stacked" } else { &spec.sparkline_type });
    group.set_attr("displayEmptyCellsAs", if spec.empty_cells == "connect" { "span" } else { &spec.empty_cells });
    let (minimum_mode, minimum) = spec.extremum(false); let (maximum_mode, maximum) = spec.extremum(true);
    group.set_attr("minAxisType", minimum_mode); group.set_attr("maxAxisType", maximum_mode);
    if let Some(value) = minimum { group.set_attr("manualMin", value.to_string()); }
    if let Some(value) = maximum { group.set_attr("manualMax", value.to_string()); }
    group.set_attr("lineWeight", crate::pixels_to_points(spec.line_width.unwrap_or(2.0)).to_string());
    let axis = spec.axis_options.as_ref();
    for (name, value) in [("displayXAxis", axis.and_then(|axis| axis.visible).unwrap_or(false)),
        ("rightToLeft", axis.and_then(|axis| axis.reverse).unwrap_or(false)), ("displayHidden", spec.contain_hidden_cells.unwrap_or(false))] {
        group.set_attr(name, if value { "1" } else { "0" });
    }
    let mut authored_colors = BTreeMap::new();
    authored_colors.insert("colorSeries", spec.color.clone());
    authored_colors.insert("colorAxis", axis.and_then(|axis| axis.color.as_ref()).cloned().unwrap_or_else(|| "#000000".into()));
    for (name, attribute, color) in [
        ("highPoint", "high", "colorHigh"), ("firstPoint", "first", "colorFirst"), ("lowPoint", "low", "colorLow"),
        ("lastPoint", "last", "colorLast"), ("negativePoint", "negative", "colorNegative"), ("markersPoint", "markers", "colorMarkers"),
    ] {
        let (visible, value) = spec.point(name);
        group.set_attr(attribute, if visible { "1" } else { "0" }); authored_colors.insert(color, value);
    }
    let mut color_nodes = Vec::new();
    for name in COLOR_NAMES {
        let value = &authored_colors[name];
        let unchanged = spec.office_source.as_ref().and_then(|source| source.resolved_colors.get(name))
            .is_some_and(|before| css_argb(before).ok() == css_argb(value).ok());
        if unchanged {
            // Omitted native colors remain omitted; their imported editor fallback does not become authored RGB.
            if let Some(existing) = unique_child(&group, xml::X14_NAMESPACE, name)? { color_nodes.push(existing.clone()); }
        } else {
            let mut generated = Element::prefixed(xml::X14_NAMESPACE, "x14", name);
            if let Some(existing) = unique_child(&group, xml::X14_NAMESPACE, name)? { generated = existing.clone(); }
            for key in ["rgb", "theme", "indexed", "auto", "tint"] { generated.remove_attr(key); }
            generated.set_attr("rgb", css_argb(value)?); color_nodes.push(generated);
        }
    }
    group.replace_children(|child| child.namespace == xml::X14_NAMESPACE && COLOR_NAMES.contains(&child.name.as_str()), color_nodes);
    group.order_children(&["colorSeries", "colorNegative", "colorAxis", "colorMarkers", "colorFirst", "colorLast", "colorHigh", "colorLow", "f", "sparklines", "extLst"]);
    Ok(Some(group))
}

/// Normalize Office defaults, numeric spelling, source quoting and member order for stale-source checks.
pub fn normalize_native_group(group: &Element) -> Result<Element, String> {
    if !group.is(xml::X14_NAMESPACE, "sparklineGroup") { return Err("SPARKLINE_NATIVE_INVALID: wrong fingerprint group namespace".into()); }
    let mut normalized = group.clone();
    for (name, default) in [("type", "line"), ("displayEmptyCellsAs", "gap"), ("minAxisType", "individual"), ("maxAxisType", "individual")] {
        if normalized.attribute_ns("", name).is_none() { normalized.set_attr(name, default); }
    }
    for name in ["dateAxis", "markers", "high", "low", "first", "last", "negative", "displayXAxis", "displayHidden", "rightToLeft"] {
        normalized.set_attr(name, if flag(group, name, false)? { "1" } else { "0" });
    }
    normalized.set_attr("lineWeight", number(group, "lineWeight")?.unwrap_or(0.75).to_string());
    for name in ["manualMin", "manualMax"] { if let Some(value) = number(group, name)? { normalized.set_attr(name, value.to_string()); } }
    for color in normalized.children.iter_mut().filter(|child| child.namespace == xml::X14_NAMESPACE && COLOR_NAMES.contains(&child.name.as_str())) {
        if let Some(rgb) = color.attribute_ns("", "rgb") {
            if argb(rgb).is_none() { return Err("SPARKLINE_NATIVE_INVALID: invalid native ARGB".into()); }
            color.set_attr("rgb", if rgb.len() == 6 { format!("FF{}", rgb.to_ascii_uppercase()) } else { rgb.to_ascii_uppercase() });
        }
        for name in ["theme", "indexed"] {
            if let Some(value) = color.attribute_ns("", name) {
                let value = value.parse::<u32>().map_err(|_| "SPARKLINE_NATIVE_INVALID: invalid color index")?;
                color.set_attr(name, value.to_string());
            }
        }
        if color.attribute_ns("", "auto").is_some() { color.set_attr("auto", if flag(color, "auto", false)? { "1" } else { "0" }); }
        if let Some(value) = number(color, "tint")? { color.set_attr("tint", value.to_string()); }
    }
    if let Some(members) = normalized.children.iter_mut().find(|child| child.is(xml::X14_NAMESPACE, "sparklines")) {
        for member in &mut members.children {
            for child in &mut member.children {
                if child.is(xml::XM_NAMESPACE, "f") { child.set_text(normalize_source_reference(&child.text)?); }
                if child.is(xml::XM_NAMESPACE, "sqref") { child.set_text(normalize_target(&child.text).ok_or("SPARKLINE_NATIVE_INVALID: invalid signature target")?); }
            }
        }
        members.children.sort_by_key(|member| member.children.iter().find(|child| child.is(xml::XM_NAMESPACE, "sqref")).map(|child| child.text.clone()).unwrap_or_default());
        members.content.clear(); members.text.clear();
        members.content.extend((0..members.children.len()).map(xml::Content::Child));
    }
    if let Some(date_formula) = normalized.children.iter_mut().find(|child| child.is(xml::XM_NAMESPACE, "f")) {
        date_formula.set_text(normalize_source_reference(&date_formula.text).unwrap_or_else(|_| date_formula.text.trim().to_owned()));
    }
    normalized.order_children(&["colorSeries", "colorNegative", "colorAxis", "colorMarkers", "colorFirst", "colorLast", "colorHigh", "colorLow", "f", "sparklines", "extLst"]);
    Ok(normalized)
}

pub fn native_group_signature(group: &Element, target: &str) -> Result<String, String> {
    native_projection_signature(&native_group_fingerprint(group, None, None)?, target)
}

/// Include resolved color bindings so an external theme/palette edit cannot restore stale editor colors.
pub fn native_group_fingerprint(group: &Element, theme: Option<&Element>, indexed_colors: Option<&[String]>) -> Result<String, String> {
    let mut projection = Element::new(super::MODEL_NAMESPACE, "nativeGroup");
    projection.push(normalize_native_group(group)?);
    for name in COLOR_NAMES {
        if let Some(color) = unique_child(group, xml::X14_NAMESPACE, name)? {
            projection.set_attr(name, color_css(color, theme, indexed_colors)?);
        }
    }
    Ok(projection.fingerprint())
}

/// Compute the shared group fingerprint once, then derive constant-size signatures per member.
pub fn native_projection_signature(group_fingerprint: &str, target: &str) -> Result<String, String> {
    let target = normalize_target(target).ok_or("SPARKLINE_NATIVE_INVALID: invalid signature target")?;
    let mut signature = Element::new(super::MODEL_NAMESPACE, "projection"); signature.set_attr("target", target);
    signature.set_attr("groupFingerprint", group_fingerprint); Ok(signature.fingerprint())
}
