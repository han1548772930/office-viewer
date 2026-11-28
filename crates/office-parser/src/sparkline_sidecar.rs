//! Namespace-aware native and local sparkline metadata; publish only a complete model.
use std::collections::HashMap;
use std::io::{BufReader, Cursor};

use zip::ZipArchive;
use office_io_common::conditional_format_xml::{self as xml, Element, XmlSelection, XmlSelectionLimits};
use office_io_common::sparkline_io::{self as io, SparklineModelBudget, SparklineModelRecord};
use crate::model::SparklinePayload;

#[path = "sparkline_parts.rs"]
mod parts;

#[derive(Debug, Default)]
pub(crate) struct SparklineSidecar {
    by_sheet: HashMap<String, HashMap<String, SparklinePayload>>,
    pub(crate) warnings: Vec<String>,
    pub(crate) failure: Option<String>,
}

impl SparklineSidecar {
    pub(crate) fn read(bytes: &[u8]) -> Self {
        match read_sidecar(bytes) {
            Ok(sidecar) => sidecar,
            Err(error) => Self { warnings: vec![error.clone()], failure: Some(error), ..Self::default() },
        }
    }

    pub(crate) fn take(&mut self, sheet: &str, target: &str) -> Option<SparklinePayload> {
        self.by_sheet.get_mut(sheet)?.remove(&io::normalize_target(target)?)
    }

    pub(crate) fn take_sheet(&mut self, sheet: &str) -> HashMap<String, SparklinePayload> {
        self.by_sheet.remove(sheet).unwrap_or_default()
    }
}

fn main(namespace: &str) -> bool { matches!(namespace, xml::MAIN_NAMESPACE | xml::STRICT_MAIN_NAMESPACE) }

fn select_sheet<R: std::io::BufRead>(reader: R, limits: XmlSelectionLimits) -> Result<XmlSelection, String> {
    let mut extension = None::<String>;
    let mut local_collections = 0usize;
    let selected = xml::extract(reader, limits, |path, node| {
        if path.len() == 3 {
            extension = if main(&path[0].namespace) && path[0].name == "worksheet"
                && main(&path[1].namespace) && path[1].name == "extLst"
                && main(&path[2].namespace) && path[2].name == "ext" {
                node.attribute_ns("", "uri").map(str::to_owned)
            } else { None };
        }
        if path.len() == 4 && extension.as_deref().is_some_and(|uri| uri.eq_ignore_ascii_case(io::MODEL_EXTENSION_URI))
            && path[3].is(io::MODEL_NAMESPACE, "sparklines") {
            if node.attribute_ns("", "version") != Some("1") { return Err("SPARKLINE_MODEL_INVALID: unsupported local extension version".into()); }
            local_collections += 1;
            if local_collections > 1 { return Err("SPARKLINE_MODEL_INVALID: duplicate local collections".into()); }
        }
        if path.len() != 5 { return Ok(false); }
        let native = extension.as_deref().is_some_and(|uri| uri.eq_ignore_ascii_case(io::NATIVE_EXTENSION_URI))
            && path[3].is(xml::X14_NAMESPACE, "sparklineGroups") && path[4].is(xml::X14_NAMESPACE, "sparklineGroup");
        let local = extension.as_deref().is_some_and(|uri| uri.eq_ignore_ascii_case(io::MODEL_EXTENSION_URI))
            && path[3].is(io::MODEL_NAMESPACE, "sparklines") && path[4].is(io::MODEL_NAMESPACE, "sparkline");
        Ok(native || local)
    })?;
    if selected.root.name != "worksheet" || !main(&selected.root.namespace) { return Err("SPARKLINE_NATIVE_INVALID: invalid worksheet root".into()); }
    Ok(selected)
}

fn read_sidecar(bytes: &[u8]) -> Result<SparklineSidecar, String> {
    let mut archive = ZipArchive::new(Cursor::new(bytes)).map_err(|error| format!("SPARKLINE_OPC_INVALID: invalid XLSX ZIP: {error}"))?;
    let sheets = crate::xlsx_parts::worksheet_parts(&mut archive)?;
    let mut selected_sheets = Vec::new();
    let mut selected_bytes = 0usize; let mut selected_elements = 0usize;
    for sheet in sheets {
        let entry = archive.by_name(&sheet.path).map_err(|error| format!("SPARKLINE_OPC_INVALID: missing worksheet {}: {error}", sheet.path))?;
        let mut limits = XmlSelectionLimits::default();
        limits.total_bytes = limits.total_bytes.saturating_sub(selected_bytes);
        limits.total_elements = limits.total_elements.saturating_sub(selected_elements);
        let selected = select_sheet(BufReader::new(entry), limits).map_err(|error| format!("{error} [worksheet {}]", sheet.name))?;
        selected_bytes = selected_bytes.checked_add(selected.selected_bytes).ok_or("SPARKLINE_SOURCE_LIMIT: selected byte counter overflow")?;
        selected_elements = selected_elements.checked_add(selected.selected_elements).ok_or("SPARKLINE_SOURCE_LIMIT: selected element counter overflow")?;
        if !selected.fragments.is_empty() { selected_sheets.push((sheet.name, selected.fragments)); }
    }
    let has_native = selected_sheets.iter().any(|(_, fragments)| fragments.iter().any(|node| node.is(xml::X14_NAMESPACE, "sparklineGroup")));
    let resources = if has_native { parts::read_color_resources(&mut archive)? } else { parts::ColorResources::default() };
    let mut sidecar = SparklineSidecar::default();
    let mut budget = SparklineModelBudget::default();
    for (sheet_name, fragments) in selected_sheets {
        let specs = parse_sheet_fragments(&fragments, &resources, &mut budget, &mut sidecar.warnings)?;
        if !specs.is_empty() { sidecar.by_sheet.insert(sheet_name, specs); }
    }
    Ok(sidecar)
}

fn parse_sheet_fragments(
    fragments: &[Element], resources: &parts::ColorResources, budget: &mut SparklineModelBudget,
    warnings: &mut Vec<String>,
) -> Result<HashMap<String, SparklinePayload>, String> {
    let mut specs = HashMap::new();
    let mut native_signatures = HashMap::new();
    let mut rank = 0usize;
    for group in fragments.iter().filter(|node| node.is(xml::X14_NAMESPACE, "sparklineGroup")) {
        let fingerprint = io::native_group_fingerprint(group, resources.theme.as_ref(), resources.indexed_colors.as_deref())?;
        let group_id = format!("office-sparkline-{rank}-{}", &fingerprint[..12]); rank += 1;
        let decoded = io::decode_native_group(group, &group_id, resources.theme.as_ref(), resources.theme_xml.as_deref(),
            resources.indexed_colors.as_deref(), warnings)?;
        for (target, source) in decoded.members {
            if specs.contains_key(&target) { return Err("SPARKLINE_NATIVE_INVALID: target belongs to multiple native groups".into()); }
            let mut spec = decoded.settings.clone(); spec.source_range = source;
            budget.include(&spec)?;
            native_signatures.insert(target.clone(), io::native_projection_signature(&fingerprint, &target)?);
            specs.insert(target, spec);
        }
    }
    let mut local_targets = std::collections::HashSet::new();
    for local in fragments.iter().filter(|node| node.is(io::MODEL_NAMESPACE, "sparkline")) {
        if !local.children.is_empty() { return Err("SPARKLINE_MODEL_INVALID: local record must contain JSON text".into()); }
        let mut record: SparklineModelRecord = serde_json::from_str(&local.text).map_err(|error| format!("SPARKLINE_MODEL_INVALID: invalid local record: {error}"))?;
        let target = io::normalize_target(&record.target).ok_or("SPARKLINE_MODEL_INVALID: invalid local target")?;
        if local.attribute_ns("", "target") != Some(target.as_str()) || !local_targets.insert(target.clone()) {
            return Err("SPARKLINE_MODEL_INVALID: conflicting/duplicate local target".into());
        }
        record.spec.validate()?;
        if record.spec.sparkline_type == "pie" {
            if record.native_signature.is_some() { return Err("SPARKLINE_MODEL_INVALID: Pie cannot declare a native projection".into()); }
            if native_signatures.contains_key(&target) {
                warnings.push(format!("SPARKLINE_LOCAL_STALE: native content at {target} takes precedence over the local Pie record"));
                continue;
            }
            budget.include(&record.spec)?;
            warnings.push(format!("SPARKLINE_LOCAL_ONLY: Pie at {target} was restored from the editable Office Studio extension"));
            specs.insert(target, record.spec);
        } else {
            let signature = record.native_signature.as_deref().filter(|value| value.len() == 16 && value.bytes().all(|byte| byte.is_ascii_hexdigit()))
                .ok_or("SPARKLINE_MODEL_INVALID: native record needs a valid projection signature")?;
            if native_signatures.get(&target).map(String::as_str) != Some(signature) {
                warnings.push(format!("SPARKLINE_LOCAL_STALE: native sparkline at {target} was changed/deleted; its old local configuration was not applied"));
                continue;
            }
            let previous = specs.get(&target).ok_or("SPARKLINE_MODEL_INVALID: matching native target is unavailable")?;
            // Refresh inert source resources from the actual package. Unrelated external theme edits
            // must survive even when the group's resolved colors still match its native signature.
            record.spec.office_source = previous.office_source.clone();
            budget.replace(previous, &record.spec)?;
            specs.insert(target, record.spec);
        }
    }
    Ok(specs)
}

#[cfg(test)]
#[path = "sparkline_sidecar_tests.rs"]
mod tests;
