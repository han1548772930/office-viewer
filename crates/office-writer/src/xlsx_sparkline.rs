//! XLSX x14 groups and editable local sparkline metadata from the canonical cell model.
use std::collections::{BTreeMap, BTreeSet};
use serde_json::Value;
use office_io_common::conditional_format_xml::{self as source_xml, Element};
use office_io_common::sparkline_io::{self as io, SparklineModelBudget, SparklineModelRecord, SparklineSpec};

use crate::{ExportError, ExportWarning};

struct Group {
    template: Element,
    members: Vec<(String, String)>,
    theme: Option<Element>,
    palette: Option<Vec<String>>,
}

fn invalid(message: impl Into<String>) -> ExportError { ExportError::invalid(message) }
fn parsed_spec(value: &Value) -> Result<SparklineSpec, ExportError> {
    let spec: SparklineSpec = serde_json::from_value(value.clone()).map_err(|error| invalid(format!("SPARKLINE_MODEL_INVALID: {error}")))?;
    spec.validate().map_err(invalid)?; Ok(spec)
}

pub(crate) fn write_xml(
    sheet_name: &str, cells: &serde_json::Map<String, Value>, warnings: &mut Vec<ExportWarning>,
) -> Result<String, ExportError> {
    let mut groups = BTreeMap::<(String, String), Group>::new();
    let mut authored = BTreeMap::<String, SparklineSpec>::new();
    let mut group_styles = BTreeMap::<String, BTreeSet<String>>::new();
    let mut budget = SparklineModelBudget::default();
    let mut local_features = false;
    for (raw_target, cell) in cells {
        let Some(raw) = cell.get("sparkline").filter(|value| !value.is_null()) else { continue; };
        let target = io::normalize_target(raw_target).ok_or_else(|| invalid(format!("SPARKLINE_REFERENCE_INVALID: invalid target {raw_target:?}")))?;
        let spec = parsed_spec(raw)?; budget.include(&spec).map_err(invalid)?;
        if authored.contains_key(&target) { return Err(invalid("SPARKLINE_MODEL_INVALID: repeated canonical target")); }
        local_features |= spec.show_gradient == Some(true) || spec.show_radius == Some(true)
            || spec.non_numeric_cells.as_deref().is_some_and(|value| value != "gap");
        if let Some(template) = io::native_group_for_spec(&spec).map_err(invalid)? {
            let source = io::source_formula(sheet_name, &spec.source_range).map_err(invalid)?;
            let theme = spec.office_source.as_ref().and_then(|source| source.theme_xml.as_deref()).map(source_xml::parse).transpose().map_err(invalid)?;
            let palette = spec.office_source.as_ref().and_then(|source| source.indexed_colors.clone());
            let style = io::native_group_fingerprint(&template, theme.as_ref(), palette.as_deref()).map_err(invalid)?;
            let identity = spec.group_id.clone().unwrap_or_else(|| format!("ungrouped:{target}"));
            group_styles.entry(identity.clone()).or_default().insert(style.clone());
            groups.entry((identity, style)).or_insert_with(|| Group { template, members: Vec::new(), theme, palette }).members.push((target.clone(), source));
        } else {
            warnings.push(ExportWarning::new("SPARKLINE_LOCAL_ONLY",
                format!("Pie sparkline at {target} is saved as editable Office Studio metadata; XLSX has no native Pie sparkline type")));
        }
        authored.insert(target, spec);
    }
    if authored.is_empty() { return Ok(String::new()); }
    for (identity, styles) in group_styles {
        if styles.len() > 1 {
            warnings.push(ExportWarning::new("SPARKLINE_NATIVE_GROUP_SPLIT",
                format!("Sparkline group {identity:?} has incompatible retained native settings; native groups were split and the editable group ID was retained")));
        }
    }
    if local_features {
        warnings.push(ExportWarning::new("SPARKLINE_LOCAL_FEATURES",
            "Gradient, rounded corners and non-numeric display policies are retained in the editable local extension; native XLSX rendering differs"));
    }
    let mut signatures = BTreeMap::<String, String>::new();
    let mut native_groups = Element::prefixed(source_xml::X14_NAMESPACE, "x14", "sparklineGroups");
    for (_, mut group) in groups {
        let mut members = Element::prefixed(source_xml::X14_NAMESPACE, "x14", "sparklines");
        for (target, formula) in &group.members {
            let mut member = Element::prefixed(source_xml::X14_NAMESPACE, "x14", "sparkline");
            let mut f = Element::prefixed(source_xml::XM_NAMESPACE, "xm", "f"); f.set_text(formula);
            let mut sqref = Element::prefixed(source_xml::XM_NAMESPACE, "xm", "sqref"); sqref.set_text(target);
            member.push(f); member.push(sqref); members.push(member);
        }
        group.template.push(members);
        group.template.order_children(&["colorSeries", "colorNegative", "colorAxis", "colorMarkers", "colorFirst", "colorLast", "colorHigh", "colorLow", "f", "sparklines", "extLst"]);
        group.template.validate_fragment().map_err(invalid)?;
        let fingerprint = io::native_group_fingerprint(&group.template, group.theme.as_ref(), group.palette.as_deref()).map_err(invalid)?;
        for (target, _) in &group.members { signatures.insert(target.clone(), io::native_projection_signature(&fingerprint, target).map_err(invalid)?); }
        native_groups.push(group.template);
    }
    let mut ext_list = Element::new(source_xml::MAIN_NAMESPACE, "extLst");
    if !native_groups.children.is_empty() {
        let mut extension = Element::new(source_xml::MAIN_NAMESPACE, "ext"); extension.set_attr("uri", io::NATIVE_EXTENSION_URI);
        extension.push(native_groups); ext_list.push(extension);
    }
    let mut local = Element::prefixed(io::MODEL_NAMESPACE, "os", "sparklines"); local.set_attr("version", "1");
    for (target, spec) in authored {
        let record = SparklineModelRecord { native_signature: signatures.remove(&target), target: target.clone(), spec };
        let mut node = Element::prefixed(io::MODEL_NAMESPACE, "os", "sparkline"); node.set_attr("target", target);
        node.set_text(serde_json::to_string(&record).map_err(|error| invalid(error.to_string()))?); local.push(node);
    }
    let mut extension = Element::new(source_xml::MAIN_NAMESPACE, "ext"); extension.set_attr("uri", io::MODEL_EXTENSION_URI);
    extension.push(local); ext_list.push(extension);
    // The full extension is admitted before a String allocation. No partial output is returned.
    ext_list.validate_fragment().map_err(invalid)?;
    Ok(ext_list.render())
}

fn all_specs(workbook: &super::WorkbookInput, mut visit: impl FnMut(SparklineSpec) -> Result<(), ExportError>) -> Result<(), ExportError> {
    for sheet in &workbook.sheets {
        for cell in sheet.cells.values() {
            if let Some(raw) = cell.get("sparkline").filter(|value| !value.is_null()) { visit(parsed_spec(raw)?)?; }
        }
    }
    Ok(())
}

/// The parent emits this one theme part and relationship for both conditional formats and sparklines.
pub(super) fn merge_theme(workbook: &super::WorkbookInput, conditional_theme: Option<String>) -> Result<Option<String>, ExportError> {
    let mut selected = conditional_theme.map(|xml| source_xml::parse(&xml).map(|node| (node.fingerprint(), xml))).transpose().map_err(invalid)?;
    let mut budget = SparklineModelBudget::default();
    all_specs(workbook, |spec| {
        budget.include(&spec).map_err(invalid)?;
        if spec.sparkline_type == "pie" { return Ok(()); }
        if let Some(theme) = spec.office_source.and_then(|source| source.theme_xml) {
            let fingerprint = source_xml::parse(&theme).map_err(invalid)?.fingerprint();
            if selected.as_ref().is_some_and(|(existing, _)| existing != &fingerprint) {
                return Err(invalid("SPARKLINE_THEME_CONFLICT: a single workbook theme cannot preserve conflicting native source themes"));
            }
            if selected.is_none() { selected = Some((fingerprint, theme)); }
        }
        Ok(())
    })?;
    Ok(selected.map(|(_, xml)| xml))
}

/// Append the authored indexed palette to generated styles, without building a tree of the full styles part.
pub(super) fn apply_indexed_colors(workbook: &super::WorkbookInput, mut styles_xml: String) -> Result<String, ExportError> {
    let mut selected = None::<Vec<String>>;
    all_specs(workbook, |spec| {
        if spec.sparkline_type == "pie" { return Ok(()); }
        if let Some(palette) = spec.office_source.and_then(|source| source.indexed_colors) {
            if selected.as_ref().is_some_and(|existing| existing != &palette) {
                return Err(invalid("SPARKLINE_PALETTE_CONFLICT: a single workbook palette cannot preserve conflicting indexed source colors"));
            }
            if selected.is_none() { selected = Some(palette); }
        }
        Ok(())
    })?;
    let Some(palette) = selected else { return Ok(styles_xml); };
    // Mixing a copied default-palette source with a different custom palette must never change its colors silently.
    all_specs(workbook, |spec| {
        if spec.sparkline_type == "pie" { return Ok(()); }
        if let Some(source) = spec.office_source.filter(|source| source.indexed_colors.is_none()) {
            let group = source_xml::parse(&source.group_xml).map_err(invalid)?;
            let theme = source.theme_xml.as_deref().map(source_xml::parse).transpose().map_err(invalid)?;
            let before = io::native_group_fingerprint(&group, theme.as_ref(), None).map_err(invalid)?;
            let after = io::native_group_fingerprint(&group, theme.as_ref(), Some(&palette)).map_err(invalid)?;
            if before != after { return Err(invalid("SPARKLINE_PALETTE_CONFLICT: the selected palette changes a retained default-palette sparkline")); }
        }
        Ok(())
    })?;
    let mut colors = Element::new(source_xml::MAIN_NAMESPACE, "colors");
    let mut indexed = Element::new(source_xml::MAIN_NAMESPACE, "indexedColors");
    for value in palette {
        let mut entry = Element::new(source_xml::MAIN_NAMESPACE, "rgbColor"); entry.set_attr("rgb", value); indexed.push(entry);
    }
    colors.push(indexed);
    let position = styles_xml.rfind("</styleSheet>").ok_or_else(|| invalid("SPARKLINE_PALETTE_INVALID: generated styles closing tag is unavailable"))?;
    if styles_xml.contains("<colors") { return Err(invalid("SPARKLINE_PALETTE_CONFLICT: generated styles already contain a color palette")); }
    styles_xml.insert_str(position, &colors.render()); Ok(styles_xml)
}

#[cfg(test)]
#[path = "xlsx_sparkline_tests.rs"]
mod tests;
